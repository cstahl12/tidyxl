#include <Rcpp.h>
/* #include "zip.h" */
#include "rapidxml.h"
#include "xlsxcell.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"
#include "utils.h"

using namespace Rcpp;

inline Rcpp::Function tidyxl(const std::string& fun){
  Rcpp::Environment env = Rcpp::Environment::namespace_env("tidyxl");
  return env[fun];
}

inline std::string zip_buffer(
    const std::string& zip_path,
    const std::string& file_path) {
  Rcpp::Function zip_buffer = tidyxl("zip_buffer");
  Rcpp::RawVector xml = Rcpp::as<Rcpp::RawVector>(zip_buffer(zip_path, file_path));
  std::string buffer(RAW(xml), RAW(xml) + xml.size());
  buffer.push_back('\0');
  return buffer;
}

xlsxsheet::xlsxsheet(
    const std::string& name,
    const std::string& sheet_path,
    xlsxbook& book,
    Rcpp::String comments_path):
  name_(name),
  book_(book) {

/*   // Doesn't work */
/*   sheet_ = zip_buffer(book_.path_, sheet_path); */

  // might work
  std::string buffer = zip_buffer("./tests/testthat/examples.xlsx", "xl/worksheets/sheet2.xml");
  sheet_ = buffer;

  Rcpp::Rcout << "constructor:\n" << sheet_ << "\n\n";

/*   rapidxml::xml_document<> xml; */
/*   xml.parse<0>(&sheet_[0]); */

/*   worksheet_ = xml.first_node("worksheet"); */
/*   sheetData_ = worksheet_->first_node("sheetData"); */

  /* defaultRowHeight_ = 15; */
  /* defaultColWidth_ = 8.47; */

  cacheCellcount();
  Rcpp::Rcout << "constructed\n";
  /* cacheDefaultRowColDims(worksheet_); */
  /* cacheColWidths(worksheet_); */
  /* initializeColumns(); */
  /* cacheComments(comments_path); */
  /* /1* appendComments(); *1/ */
}

List& xlsxsheet::information() {
  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.
  information_ = List::create(
      _["address"] = address_,
      _["row"] = row_,
      _["col"] = col_,
      _["content"] = content_,
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
  int n = Rf_length(information_[0]);
  information_.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  information_.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return information_;
}

void xlsxsheet::cacheDefaultRowColDims(rapidxml::xml_node<>* worksheet) {
  rapidxml::xml_node<>* sheetFormatPr_ = worksheet->first_node("sheetFormatPr");

  if (sheetFormatPr_ != NULL) {
    // Don't use utils::getAttributeValueDouble because it might overwrite the
    // default value with NA_REAL.
    rapidxml::xml_attribute<>* defaultRowHeight =
      sheetFormatPr_->first_attribute("defaultRowHeight");
    if (defaultRowHeight != NULL)
      defaultRowHeight_ = strtod(defaultRowHeight->value(), NULL);

    rapidxml::xml_attribute<>*defaultColWidth =
      sheetFormatPr_->first_attribute("defaultColWidth");
    if (defaultColWidth != NULL) {
      defaultColWidth_ = strtod(defaultColWidth->value(), NULL);
    } else {
      // If defaultColWidth not given, ECMA says you can work it out based on
      // baseColWidth, but that isn't necessarily given either, and the formula
      // is wrong because the reality is so complicated, see
      // https://support.microsoft.com/en-gb/kb/214123.
      defaultColWidth_ = 8.38;
    }
  }
}

void xlsxsheet::cacheColWidths(rapidxml::xml_node<>* worksheet) {
  // Having done cacheDefaultRowColDims(), initilize vector to default width,
  // then update with custom widths.  The number of columns might be available
  // by parsing <dimension><ref>, but if not then only by parsing the address of
  // all the cells.  I think it's better just to use the maximum possible number
  // of columns, 16834.

  colWidths_.assign(16384, defaultColWidth_);

  rapidxml::xml_node<>* cols = worksheet->first_node("cols");
  if (cols == NULL)
    return; // No custom widths

  for (rapidxml::xml_node<>* col = cols->first_node("col");
      col; col = col->next_sibling("col")) {

    // <col> applies to columns from a min to a max, which must be iterated over
    unsigned int min  = strtol(col->first_attribute("min")->value(), NULL, 10);
    unsigned int max  = strtol(col->first_attribute("max")->value(), NULL, 10);
    double width = strtod(col->first_attribute("width")->value(), NULL);

    for (unsigned int column = min; column <= max; ++column)
      colWidths_[column - 1] = width;
  }
}

void xlsxsheet::cacheCellcount() {
  // Iterate over all rows and cells to count.  The 'dimension' tag is no use
  // here because it describes a rectangle of cells, many of which may be blank.
  Rcpp::Rcout << "method:\n" << sheet_ << "\n\n";
  rapidxml::xml_document<> xml;
  xml.parse<0>(&sheet_[0]);
  worksheet_ = xml.first_node("worksheet");
  sheetData_ = worksheet_->first_node("sheetData");

  cellcount_ = 0;
  for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
      row; row = row->next_sibling("row")) {
    Rcpp::Rcout << "row\n";
    for (rapidxml::xml_node<>* c = row->first_node("c");
        c; c = c->next_sibling("c")) {
      Rcpp::Rcout << "cell\n";
      ++cellcount_;
    }
  }
}

void xlsxsheet::initializeColumns() {
  // Having done cacheCellcount(), make columns of that length
  address_         = CharacterVector(cellcount_, NA_STRING);
  row_             = IntegerVector(cellcount_,   NA_INTEGER);
  col_             = IntegerVector(cellcount_,   NA_INTEGER);
  content_         = CharacterVector(cellcount_, NA_STRING);
  formula_         = CharacterVector(cellcount_, NA_STRING);
  formula_type_    = CharacterVector(cellcount_, NA_STRING);
  formula_ref_     = CharacterVector(cellcount_, NA_STRING);
  formula_group_   = IntegerVector(cellcount_,   NA_INTEGER);
  value_           = List(cellcount_);
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

void xlsxsheet::cacheComments(Rcpp::String comments_path) {
  // Having constructed the map, they will each be deleted when they are matched
  // to a cell.  That will leave only those comments that are on empty cells.
  // Those are then appended as empty cells with comments.
  if (comments_path != NA_STRING) {
    std::string comments_file = zip_buffer(book_.path_, comments_path);
    rapidxml::xml_document<> xml;
    xml.parse<0>(&comments_file[0]);

    // Iterate over the comments to store the ref and text
    rapidxml::xml_node<>* comments = xml.first_node("comments");
    rapidxml::xml_node<>* commentList = comments->first_node("commentList");
    for (rapidxml::xml_node<>* comment = commentList->first_node();
        comment; comment = comment->next_sibling()) {
      std::string ref(comment->first_attribute("ref")->value());
      rapidxml::xml_node<>* r = comment->first_node();
      // Get the inline string
      std::string inlineString;
      parseString(r, inlineString); // value is modified in place
      comments_[ref] = inlineString;
    }
  }
}

void xlsxsheet::parseSheetData() {
  // Iterate through rows and cells in sheetData.  Cell elements are children
  // of row elements.  Columns are described elswhere in cols->col.
  rowHeights_.assign(1048576, defaultRowHeight_); // cache rowHeight while here
  unsigned long int rowNumber;
  i_ = 0; // counter for checkUserInterrupt
  for (rapidxml::xml_node<>* row = sheetData_->first_node();
      row; row = row->next_sibling()) {
    Rcpp::Rcout << i_ << "\n";
    ++i_;
    /* rapidxml::xml_attribute<>* r = row->first_attribute("r"); */
    /* if (r == NULL) */
    /*   stop("Invalid row: lacks 'r' attribute"); */
    /* rowNumber = strtod(r->value(), NULL); */
    // Check for custom row height
    /* double rowHeight = defaultRowHeight_; */
    /* rapidxml::xml_attribute<>* ht = row->first_attribute("ht"); */
    /* if (ht != NULL) { */
    /*   rowHeight = strtod(ht->value(), NULL); */
    /*   rowHeights_[rowNumber - 1] = rowHeight; */
    /* } */

    /* for (rapidxml::xml_node<>* c = row->first_node(); */
    /*     c; c = c->next_sibling()) { */
    /*   /1* xlsxcell cell(c, this, book_, i_); *1/ */

    /*   // Height and width aren't really determined by the cell, so they're done */
    /*   // in this sheet instance */
    /*   book_.height_[i_] = rowHeight; */
    /*   book_.width_[i_] = colWidths_[col_[i_] - 1]; */

    /*   ++i_; */
    /*   if ((i_ + 1) % 1000 == 0) */
    /*     checkUserInterrupt(); */
    /* } */
  }
}

/* void xlsxsheet::appendComments() { */
/*   // Having constructed the comments_ map, they are each be deleted when they */
/*   // are matched to a cell.  That leaves only those comments that are on empty */
/*   // cells.  This code appends those remaining comments as empty cells. */
/*   int col; */
/*   int row; */
/*   for(std::map<std::string, std::string>::iterator it = comments_.begin(); */
/*       it != comments_.end(); ++it) { */
/*     ++i_; */
/*     // TODO: move address parsing to utils */
/*     std::string address = it->first.c_str(); // we need this std::string in a moment */
/*     // Iterate though the A1-style address string character by character */
/*     col = 0; */
/*     row = 0; */
/*     for(std::string::const_iterator iter = address.begin(); */
/*         iter != address.end(); ++iter) { */
/*       if (*iter >= '0' && *iter <= '9') { // If it's a number */
/*         row = row * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number */
/*       } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character */
/*         col = 26 * col + (*iter - 'A' + 1); // Then do similarly with columns */
/*       } */
/*     } */
/*     address_.push_back(address); */
/*     row_.push_back(row); */
/*     col_.push_back(col); */
/*     content_.push_back(NA_STRING); */
/*     formula_.push_back(NA_STRING); */
/*     formula_type_.push_back(NA_STRING); */
/*     formula_ref_.push_back(NA_STRING); */
/*     formula_group_.push_back(NA_INTEGER); */
/*     type_.push_back(NA_STRING); */
/*     data_type_.push_back("blank"); */
/*     error_.push_back(NA_STRING); */
/*     logical_.push_back(NA_LOGICAL); */
/*     numeric_.push_back(NA_REAL); */
/*     date_.push_back(NA_REAL); */
/*     character_.push_back(NA_STRING); */
/*     comment_.push_back(it->second.c_str()); */
/*     height_.push_back(rowHeights_[row - 1]); */
/*     width_.push_back(colWidths_[col - 1]); */
/*     style_format_.push_back("Normal"); */
/*     local_format_id_.push_back(1); */
/*   } */
/*   // Iterate though the A1-style address string character by character */

/* } */

