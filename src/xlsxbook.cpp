#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxsheet.h"
#include "styles.h"
#include "utils.h"
#include "foo.h"

using namespace Rcpp;

// Constructor for cell data
xlsxbook::xlsxbook(
    const std::string& path,
    CharacterVector& sheet_paths,
    CharacterVector& sheet_names,
    CharacterVector& comments_paths
    ):
  path_(path),
  styles_(path_),
  sheet_paths_(sheet_paths),
  sheet_names_(sheet_names),
  comments_paths_(comments_paths) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  Rcpp::Rcout << "cacheDateOffset\n";
  cacheDateOffset(workbook);
  Rcpp::Rcout << "cacheStrings\n";
  cacheStrings();
  Rcpp::Rcout << "cacheSheets\n";
  cacheSheets();
  Rcpp::Rcout << "countCells\n";
  countCells();
  Rcpp::Rcout << "initializeColumns\n";
  initializeColumns();
  Rcpp::Rcout << "collectCellValues\n";
  collectCellValues();
}

// Construct for formats
xlsxbook::xlsxbook(const std::string& path): path_(path), styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook);
  cacheStrings();
}


// Based on hadley/readxl
void xlsxbook::cacheStrings() {
  if (!zip_has_file(path_, "xl/sharedStrings.xml"))
    return;

  std::string xml = zip_buffer(path_, "xl/sharedStrings.xml");
  rapidxml::xml_document<> sharedStrings;
  sharedStrings.parse<0>(&xml[0]);

  rapidxml::xml_node<>* sst = sharedStrings.first_node("sst");
  rapidxml::xml_attribute<>* uniqueCount = sst->first_attribute("uniqueCount");
  if (uniqueCount != NULL) {
    unsigned long int n = strtol(uniqueCount->value(), NULL, 10);
    strings_.reserve(n);
  }

  // 18.4.8 si (String Item) [p1725]
  for (rapidxml::xml_node<>* string = sst->first_node();
      string; string = string->next_sibling()) {
    std::string out;
    parseString(string, out);    // missing strings are treated as empty ""
    strings_.push_back(out);
  }
}

void xlsxbook::cacheDateOffset(rapidxml::xml_node<>* workbook) {
  rapidxml::xml_node<>* workbookPr = workbook->first_node("workbookPr");
  if (workbookPr != NULL) {
    rapidxml::xml_attribute<>* date1904 = workbookPr->first_attribute("date1904");
    if (date1904 != NULL) {
      std::string is1904 = date1904->value();
      if ((is1904 == "1") || (is1904 == "true")) {
        dateSystem_ = 1904;
        dateOffset_ = 24107;
        return;
      }
    }
  }

  dateSystem_ = 1900;
  dateOffset_ = 25569;
}

void xlsxbook::cacheSheets() {
  // Loop through sheets
  sheets_.reserve(sheet_paths_.size());
  CharacterVector::iterator in_it;
  int i = 0;
  for(in_it = sheet_paths_.begin();
      in_it != sheet_paths_.end();
      ++in_it) {
    String sheet_path(sheet_paths_[i]);
    String sheet_name(sheet_names_[i]);
    String comments_path(comments_paths_[i]);
    sheets_.emplace_back(foo(sheet_name, sheet_path, *this, comments_path));
    ++i;
    break;
  }
  for(std::vector<foo>::iterator it = sheets_.begin();
      it != sheets_.end();
      ++it) {
    Rcpp::Rcout << &(it->sheet_) << "\n";
    break;
  }
}

void xlsxbook::countCells() {
  // Count the number of cells in all sheets together
  cellcount_ = 0;
  for(std::vector<foo>::iterator it = sheets_.begin();
      it != sheets_.end();
      ++it) {
    cellcount_ += it->cellcount_;
  }
}

void xlsxbook::initializeColumns() {
  // Having done cacheCellcount(), make columns of that length
  sheet_name_      = CharacterVector(cellcount_, NA_STRING);
  address_         = CharacterVector(cellcount_, NA_STRING);
  row_             = IntegerVector(cellcount_,   NA_INTEGER);
  col_             = IntegerVector(cellcount_,   NA_INTEGER);
  formula_         = CharacterVector(cellcount_, NA_STRING);
  formula_type_    = CharacterVector(cellcount_, NA_STRING);
  formula_ref_     = CharacterVector(cellcount_, NA_STRING);
  formula_group_   = IntegerVector(cellcount_,   NA_INTEGER);
  data_type_       = CharacterVector(cellcount_, NA_STRING);
  error_           = CharacterVector(cellcount_, NA_STRING);
  logical_         = LogicalVector(cellcount_,   NA_LOGICAL);
  numeric_         = NumericVector(cellcount_,   NA_REAL);
  date_            = NumericVector(cellcount_,   NA_REAL);
  date_.attr("class") = CharacterVector::create("POSIXct", "POSIXt");
  date_.attr("tzone") = "UTC";
  character_       = CharacterVector(cellcount_, NA_STRING);
  comment_         = CharacterVector(cellcount_, NA_STRING);
  height_          = NumericVector(cellcount_,   NA_REAL);
  width_           = NumericVector(cellcount_,   NA_REAL);
  style_format_    = CharacterVector(cellcount_, NA_STRING);
  local_format_id_ = IntegerVector(cellcount_,   NA_INTEGER);
}

void xlsxbook::collectCellValues() {
  unsigned long long int i(0); // cell position within output dataframe
  for(std::vector<foo>::iterator it = sheets_.begin();
      it != sheets_.end();
      ++it) {
    it->collectCellValues(i); // i is modified in place ready for the next sheet
  }
}

List xlsxbook::information() {
  // Return a nested data frame of cells and their contents.
  // This maxes out the list constructor at 20 vectors, so to increase from here
  // you must construct a list
  //   List out(23);
  // Then assign the vectors
  //   out[0] = blah;
  // Then name the vectors with a character vector
  //   static std::vector<std::string> names;
  //   names = {"sheet", "address", "etc."};
  //   out.attr("names) = names;

  List out = List::create(
      _["sheet"] = sheet_name_,
      _["address"] = address_,
      _["row"] = row_,
      _["col"] = col_,
      _["formula"] = formula_,
      _["formula_type"] = formula_type_,
      _["formula_ref"] = formula_ref_,
      _["formula_group"] = formula_group_,
      _["data_type"] = data_type_,
      _["error"] = error_,
      _["logical"] = logical_,
      _["numeric"] = numeric_,
      _["date"] = date_,
      _["character"] = character_,
      _["comment"] = comment_,
      _["height"] = height_,
      _["width"] = width_,
      _["style_format"] = style_format_,
      _["local_format_id"] = local_format_id_);

  // Turn list of vectors into a data frame without checking anything
  int n = Rf_length(out[0]);
  out.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  out.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)
  out.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -cellcount_); // Dunno how this works (the -cellcount_ part)

  return out;
}
