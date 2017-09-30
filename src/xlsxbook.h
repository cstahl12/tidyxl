#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"
#include "styles.h"
#include "xlsxsheet.h"
#include "foo.h"

class xlsxbook {

  public:

    const std::string& path_;                     // workbook path
    Rcpp::CharacterVector sheet_paths_;
    Rcpp::CharacterVector sheet_names_;
    Rcpp::CharacterVector comments_paths_;

    int dateSystem_; // 1900 or 1904
    int dateOffset_; // for converting 1900 or 1904 Excel datetimes to R

    styles styles_;
    std::vector<std::string> strings_; // string lookup table
    std::vector<foo> sheets_;

    unsigned long long int cellcount_; // count of cells in all sheets together

    // These variables go to R
    Rcpp::CharacterVector sheet_name_;     // Sheet name
    Rcpp::CharacterVector address_;   // Value of cell node r
    Rcpp::IntegerVector   row_;       // Parsed address_ (one-based)
    Rcpp::IntegerVector   col_;       // Parsed address_ (one-based)
    Rcpp::CharacterVector formula_;   // If present
    Rcpp::CharacterVector formula_type_; // If present
    Rcpp::CharacterVector formula_ref_;  // If present
    Rcpp::IntegerVector   formula_group_; // If present
    Rcpp::CharacterVector data_type_; // Type of the parsed value
    Rcpp::CharacterVector error_;     // Parsed value
    Rcpp::LogicalVector   logical_;   // Parsed value
    Rcpp::NumericVector   numeric_;   // Parsed value
    Rcpp::NumericVector   date_;      // Parsed value
    Rcpp::CharacterVector character_; // Parsed value
    Rcpp::CharacterVector comment_;   // Looked up in comments_
    Rcpp::NumericVector   height_;          // Provided to cell constructor
    Rcpp::NumericVector   width_;           // Provided to cell constructor
    Rcpp::CharacterVector style_format_;    // cellXfs xfId links to cellStyleXfs entry
    Rcpp::IntegerVector   local_format_id_; // cell 'c' links to cellXfs entry

    xlsxbook(                          // constructor for cell data
        const std::string& path,
        Rcpp::CharacterVector& sheet_paths,
        Rcpp::CharacterVector& sheet_names,
        Rcpp::CharacterVector& comments_paths
        );

    xlsxbook(const std::string& path); // constructor for formats

    void cacheStrings();
    void cacheStyles();
    void cacheDateOffset(rapidxml::xml_node<>* workbook);
    void cacheSheets();
    void countCells();
    void initializeColumns();
    void collectCellValues();
    Rcpp::List information(); // Dataframe of cells and their contents
};

#endif
