#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxcell.h"
#include "foo.h"
#include "utils.h"

using namespace Rcpp;

xlsxcell::xlsxcell(
    xlsxbook& book,
    foo& sheet,
    rapidxml::xml_node<>* cell,
    unsigned long long int& i
    ):
  book_(book),
  sheet_(sheet),
  cell_(cell),
  i_(i)
{
    cacheAddress();
    cacheComment();
    cacheValue();
    cacheFormula();
}

// Based on hadley/readxl
// Get the A1-style address, and parse it for the row and column numbers.
// Simple parser: does not check that order of numbers and letters is correct
// row_ and column_ are one-based
void xlsxcell::cacheAddress() {
  rapidxml::xml_attribute<>* r = cell_->first_attribute("r");
  if (r == NULL)
    stop("Invalid cell: lacks 'r' attribute");
  address_.assign(r->value(), r->value_size());
  col_ = 0;
  row_ = 0;
  // Iterate though the A1-style address string character by character
  for(std::string::const_iterator iter = address_.begin();
      iter != address_.end(); ++iter) {
    if (*iter >= '0' && *iter <= '9') { // If it's a number
      row_ = row_ * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number
    } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character
      col_ = 26 * col_ + (*iter - 'A' + 1); // Then do similarly with columns
    }
  }
  book_.address_[i_] = address_;
  book_.row_[i_] = row_;
  book_.col_[i_] = col_;
}

void xlsxcell::cacheComment() {
  // Look up any comment using the address, and delete it if found
  std::map<std::string, std::string>& comments = sheet_.comments_;
  std::map<std::string, std::string>::iterator it = comments.find(address_);
  if(it != comments.end()) {
    book_.comment_[i_] = it->second;
    comments.erase(it);
  }
}

void xlsxcell::cacheValue() {
  // 'v' for 'value' is either literal (numeric) or an index into a string table
  rapidxml::xml_node<>* v = cell_->first_node("v");
  std::string vvalue;
  if (v != NULL) {
    vvalue.assign(v->value(), v->value_size());
  }

  // 't' for 'type' defines the meaning of 'v' for value
  rapidxml::xml_attribute<>* t = cell_->first_attribute("t");
  std::string tvalue;
  if (t != NULL) {
    tvalue.assign(t->value(), t->value_size());
  }

  // 's' for 'style' indexes into data structures of formatting
  rapidxml::xml_attribute<>* s = cell_->first_attribute("s");
  // Default the local format id to '1' if not present
  int svalue;
  if (s != NULL) {
    svalue = std::stoi(std::string(s->value(), s->value_size()));
  } else {
    svalue = 0;
  }
  book_.local_format_id_[i_] = svalue + 1;
  book_.style_format_[i_] =
    book_.styles_.cellStyles_map_[book_.styles_.cellXfs_[svalue].xfId_[0]];

  if (t != NULL && tvalue == "inlineStr") {
    book_.data_type_[i_] = "character";
    rapidxml::xml_node<>* is = cell_->first_node("is");
    if (is != NULL) { // Get the inline string if it's really there
      std::string inlineString;
      parseString(is, inlineString); // value is modified in place
      book_.character_[i_] = inlineString;
    }
    return;
  } else if (v == NULL) {
    // Can't now be an inline string (tested above)
    book_.data_type_[i_] = "blank";
    return;
  } else if (t == NULL || tvalue == "n") {
    if (book_.styles_.cellXfs_[svalue].applyNumberFormat_[0] == 1) {
      // local number format applies
      if (book_.styles_.isDate_[book_.styles_.cellXfs_[svalue].numFmtId_[0]]) {
        // local number format is a date format
        book_.data_type_[i_] = "date";
        double date = strtod(vvalue.c_str(), NULL);
        book_.date_[i_] = checkDate(date,
                                    book_.dateSystem_,
                                    book_.dateOffset_,
                                    ref(sheet_.name_, address_));
        return;
      } else {
        book_.data_type_[i_] = "numeric";
        book_.numeric_[i_] = strtod(vvalue.c_str(), NULL);
      }
    } else if (
          book_.styles_.isDate_[
            book_.styles_.cellStyleXfs_[
              book_.styles_.cellXfs_[svalue].xfId_[0]
            ].numFmtId_[0]
          ]
        ) {
      // style number format is a date format
      book_.data_type_[i_] = "date";
      double date = strtod(vvalue.c_str(), NULL);
      book_.date_[i_] = checkDate(date,
                                  book_.dateSystem_,
                                  book_.dateOffset_,
                                  ref(sheet_.name_, address_));
      return;
    } else {
      book_.data_type_[i_] = "numeric";
      book_.numeric_[i_] = strtod(vvalue.c_str(), NULL);
    }
  } else if (tvalue == "s") {
    // the t attribute exists and its value is exactly "s", so v is an index
    // into the string table.
      book_.data_type_[i_] = "character";
    book_.character_[i_] = book_.strings_[strtol(vvalue.c_str(), NULL, 10)];
    return;
  } else if (tvalue == "str") {
    // Formula, which could have evaluated to anything, so only a string is safe
    book_.data_type_[i_] = "character";
    book_.character_[i_] = vvalue;
    return;
  } else if (tvalue == "b"){
    book_.data_type_[i_] = "logical";
    book_.logical_[i_] = strtod(vvalue.c_str(), NULL);
    return;
  } else if (tvalue == "e") {
    book_.data_type_[i_] = "error";
    book_.error_[i_] = vvalue;
    return;
  } else if (tvalue == "d") {
    // Does excel use this date type? Regardless, don't have cross-platform
    // ISO8601 parser (yet) so need to return as text.
    book_.data_type_[i_] = "date (ISO8601)";
    return;
  } else {
    book_.data_type_[i_] = "unknown";
    warning("Unknown data type detected");
    return;
  }
}

void xlsxcell::cacheFormula() {
  // TODO: Array formulas use the ref attribute for their range, and t to
  // state that they're 'array'.
  rapidxml::xml_node<>* f = cell_->first_node("f");
  std::string formula;
  int si_number;
  std::map<int, shared_formula>::iterator it;
  if (f != NULL) {
    formula = f->value();
    rapidxml::xml_attribute<>* f_t = f->first_attribute("t");
    if (f_t != NULL) {
      book_.formula_type_[i_] = f_t->value();
    }
    rapidxml::xml_attribute<>* ref = f->first_attribute("ref");
    if (ref != NULL) {
      book_.formula_ref_[i_] = ref->value();
    }
    rapidxml::xml_attribute<>* si = f->first_attribute("si");
    if (si != NULL) {
      si_number = std::stoi(std::string(si->value()));
      book_.formula_group_[i_] = si_number;
      if (formula.length() == 0) { // inherits definition
        it = sheet_.shared_formulas_.find(si_number);
        formula = it->second.offset(row_, col_);
      } else { // defines shared formula
        shared_formula new_shared_formula(formula, row_, col_);
        sheet_.shared_formulas_.insert({si_number, new_shared_formula});
      }
    }
    book_.formula_[i_] = formula;
  }
}
