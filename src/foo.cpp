#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "rapidxml_print.h"
#include "xlsxcell.h"
#include "xlsxbook.h"
#include "utils.h"
#include "foo.h"

using namespace Rcpp;

foo::foo(
    const std::string& name,
    const std::string& sheet_path,
    xlsxbook& book,
    Rcpp::String comments_path):
  name_(name),
  book_(book),
  comments_path_(comments_path)
{
  // Load the xml file -- the two-step assignment is necessary to evade pointer
  // problems when calling methods from outside the class.
  std::string buffer = zip_buffer(book_.path_, sheet_path);
  sheet_ = buffer;
  Rcpp::Rcout << &sheet_ << "\n";

  rapidxml::xml_document<> xml;
  xml.parse<0>(&sheet_[0]);

  worksheet_ = xml.first_node("worksheet");
  sheetData_ = worksheet_->first_node("sheetData");

  Rcpp::Rcout << "cacheDefaultRowColDims\n";
  cacheDefaultRowColDims();
  Rcpp::Rcout << "cacheColWidths\n";
  cacheColWidths();
  Rcpp::Rcout << "cacheComments\n";
  cacheComments();
  Rcpp::Rcout << "cacheCellcount\n";
  cacheCellcount(); // Also caches custom row heights
};

void foo::cacheDefaultRowColDims() {
  rapidxml::xml_node<>* sheetFormatPr_ = worksheet_->first_node("sheetFormatPr");

  if (sheetFormatPr_ != NULL) {
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

void foo::cacheColWidths() {
  // Having done cacheDefaultRowColDims(), initilize vector to default width,
  // then update with custom widths.  The number of columns might be available
  // by parsing <dimension><ref>, but if not then only by parsing the address of
  // all the cells.  I think it's better just to use the maximum possible number
  // of columns, 16834.

  colWidths_.assign(16384, defaultColWidth_);

  rapidxml::xml_node<>* cols = worksheet_->first_node("cols");
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

void foo::cacheComments() {
  // Having constructed the map, they will each be deleted when they are matched
  // to a cell.  That will leave only those comments that are on empty cells.
  // Those are then appended as empty cells with comments.
  if (comments_path_ != NA_STRING) {
    std::string comments_file = zip_buffer(book_.path_, comments_path_);
    rapidxml::xml_document<> xml;
    xml.parse<0>(&comments_file[0]);

    // Iterate over the comments to store the ref and text
    rapidxml::xml_node<>* comments = xml.first_node("comments");
    rapidxml::xml_node<>* commentList = comments->first_node("commentList");
    for (rapidxml::xml_node<>* comment = commentList->first_node();
        comment; comment = comment->next_sibling()) {
      rapidxml::xml_attribute<>* ref = comment->first_attribute("ref");
      std::string reference(ref->value(), ref->value_size());
      rapidxml::xml_node<>* r = comment->first_node();
      // Get the inline string
      std::string inlineString;
      parseString(r, inlineString); // value is modified in place
      comments_[reference] = inlineString;
    }
  }
}

void foo::cacheCellcount() {
  // Iterate over all rows and cells to count (first pass).  The 'dimension' tag
  // is no use here because it describes a rectangle of cells, many of which may
  // be blank.

  cellcount_ = 0;

  // The cell addresses are inspected (but not parsed) to count the number of
  // comments that have matching cells.  This is so that the number of comments
  // that don't have matching cells (because they're blank) can be added to the
  // number of cells for initialising columns at their maximum length.

  // The matched pairs have to be found again to assign the correct comments to
  // the correct cell -- previous attempts avoided this by creating the cells in
  // the first pass, having them store the matching comment in a member
  // variable, but this was slow and filled up the stack.

  // variables for tallying matching comments
  unsigned long int rowNumber;
  std::string address;
  int comment_cell_tally = 0; //

  // While here, also cache custom rowHeight.
  rowHeights_.assign(1048576, defaultRowHeight_);

  for (rapidxml::xml_node<>* row = sheetData_->first_node("row");
      row; row = row->next_sibling("row")) {

    // Check for custom row height
    rapidxml::xml_attribute<>* r = row->first_attribute("r");
    if (r == NULL)
      stop("Invalid row: lacks 'r' attribute");
    rowNumber = strtod(r->value(), NULL);
    double rowHeight = defaultRowHeight_;
    rapidxml::xml_attribute<>* ht = row->first_attribute("ht");
    if (ht != NULL) {
      rowHeight = strtod(ht->value(), NULL);
      rowHeights_[rowNumber - 1] = rowHeight;
    }

    for (rapidxml::xml_node<>* c = row->first_node("c");
        c;
        c = c->next_sibling("c")) {

      // Tally any cell<-->comment pairs
      rapidxml::xml_attribute<>* r = c->first_attribute("r");
      if (r == NULL)
        stop("Invalid cell: lacks 'r' attribute");
      address.assign(r->value(), r->value_size());
      std::map<std::string, std::string>::iterator match = comments_.find(address);
      if(match != comments_.end())
        ++comment_cell_tally;

      ++cellcount_;
      if ((cellcount_ + 1) % 1000 == 0) {
        checkUserInterrupt();
      }
    }
  }

  // Add any remaining comment-only cells to the overall count
  cellcount_ += (comments_.size() - comment_cell_tally);
}

void foo::collectCellValues(unsigned long long int& i) {
  // Second pass through the cells, this time storing all the values directly in
  // the vectors in the book object.

  Rcpp::Rcout << "collecting cell values\n";
  Rcpp::Rcout << &sheet_ << "\n";

  // Obtain the sheetData node again -- this is because of pointer problems in
  // methods called from outside the class.
  rapidxml::xml_document<> xml;
  xml.parse<0>(&sheet_[0]);
  rapidxml::xml_node<>* worksheet = xml.first_node("worksheet");
  rapidxml::xml_node<>* sheetData = worksheet_->first_node("sheetData");
  Rcpp::Rcout << "obtained sheetData node\n";

  for (rapidxml::xml_node<>* row = sheetData->first_node("row");
      row; row = row->next_sibling("row")) {

    for (rapidxml::xml_node<>* c = row->first_node("c");
        c;
        c = c->next_sibling("c")) {

      Rcpp::Rcout << "create cell\n";
      // Create the cell, which will obtain its own properties and put them into
      // the book's vectors.
      xlsxcell cell = xlsxcell(book_, *this, c, i);

      Rcpp::Rcout << "name, height, width\n";
      // Height, width and sheet_name aren't really determined by the cell, so
      // they're done in this sheet instance
      book_.sheet_name_[i] = name_;
      book_.height_[i] = rowHeights_[cell.row_ - 1];
      book_.width_[i] = colWidths_[cell.col_ - 1];

    ++i;
    if ((i + 1) % 1000 == 0)
      checkUserInterrupt();
    }
  }

  Rcpp::Rcout << "Append comments\n";
  // Append comments from otherwise blank cells
  for(std::map<std::string, std::string>::iterator comment = comments_.begin();
      comment != comments_.end(); ++comment) {
    // TODO: move address parsing to utils
    std::string address = comment->first.c_str(); // we need this std::string in a moment
    // Iterate though the A1-style address string character by character
    int col = 0;
    int row = 0;
    for(std::string::const_iterator it = address.begin();
        it != address.end(); ++it) {
      if (*it >= '0' && *it <= '9') { // If it's a number
        row = row * 10 + (*it - '0'); // Then multiply existing row by 10 and add new number
      } else if (*it >= 'A' && *it <= 'Z') { // If it's a character
        col = 26 * col + (*it - 'A' + 1); // Then do similarly with columns
      }
    }
    book_.sheet_name_[i] = name_;
    book_.address_[i] = address;
    book_.col_[i] = col;
    book_.row_[i] = row;
    book_.data_type_[i] = "blank";
    book_.height_[i] = rowHeights_[row - 1];
    book_.width_[i] = colWidths_[col - 1];
    book_.comment_[i] = comment->second;
    book_.style_format_[i] = "Normal";
    book_.local_format_id_[i] = 1;
    ++i;
  }
}
