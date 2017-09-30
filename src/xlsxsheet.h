#ifndef XLSXSHEET_
#define XLSXSHEET_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "shared_formula.h"

class xlsxbook; // forward declaration because included in xlsxbook.h

class xlsxsheet {

  public:

    std::string name_;
    std::string sheet_path_;
    xlsxbook& book_; // reference to parent workbook
    std::string sheet_; // xml string

    rapidxml::xml_node<>* worksheet_;
    rapidxml::xml_node<>* sheetData_;

    unsigned long long int i_;         // cell counter
    unsigned long long int cellcount_; // count of cells in sheet

    double defaultRowHeight_;
    double defaultColWidth_;

    std::vector<double> colWidths_;
    std::vector<double> rowHeights_;

    // lookup tables
    std::map<int, shared_formula> shared_formulas_;
    std::map<std::string, std::string> comments_;

    Rcpp::List information_; // Wrapper for vectors returned to R

    // These vectors go to R
    Rcpp::CharacterVector address_;   // Value of cell node r
    Rcpp::IntegerVector   row_;       // Parsed address_ (one-based)
    Rcpp::IntegerVector   col_;       // Parsed address_ (one-based)
    Rcpp::CharacterVector content_;   // Unparsed value of cell node v
    Rcpp::CharacterVector formula_;   // If present
    Rcpp::CharacterVector formula_type_; // If present
    Rcpp::CharacterVector formula_ref_;  // If present
    Rcpp::IntegerVector   formula_group_; // If present
    Rcpp::List  value_;               // Parsed values wrapped in unnamed lists
    Rcpp::CharacterVector data_type_; // Type of the parsed value
    Rcpp::CharacterVector error_;     // Parsed value
    Rcpp::LogicalVector   logical_;   // Parsed value
    Rcpp::NumericVector   numeric_;   // Parsed value
    Rcpp::NumericVector   date_;      // Parsed value
    Rcpp::CharacterVector character_; // Parsed value
    Rcpp::CharacterVector comment_;   // Looked up in the lookup table
    Rcpp::NumericVector   height_;          // Provided to cell constructor
    Rcpp::NumericVector   width_;           // Provided to cell constructor
    Rcpp::CharacterVector style_format_;    // cellXfs xfId links to cellStyleXfs entry
    Rcpp::IntegerVector   local_format_id_; // cell 'c' links to cellXfs entry

    xlsxsheet(xlsxbook& book);

    xlsxsheet(
        const std::string& name,
        const std::string& sheet_path,
        xlsxbook& book,
        Rcpp::String comments_path);
    Rcpp::List& information();       // Cells contents and styles DF wrapped in list

    void cacheDefaultRowColDims(rapidxml::xml_node<>* worksheet);
    void cacheColWidths(rapidxml::xml_node<>* worksheet);
    void cacheCellcount();
    void cacheComments(Rcpp::String comments_path);
    void initializeColumns();
    void parseSheetData();
    void appendComments();

};

#endif
