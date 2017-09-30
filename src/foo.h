#ifndef FOO_
#define FOO_

#include <Rcpp.h>
#include "rapidxml.h"
#include "zip.h"
#include "xlsxbook.h"
#include "xlsxcell.h"

class xlsxcell; // Forward declaration because foo is included in xlsxcell

class foo {
  public:

    std::string name_;
    std::string sheet_path_;
    xlsxbook& book_; // reference to parent workbook
    Rcpp::String comments_path_;

    std::string sheet_; // xml string
    rapidxml::xml_node<>* worksheet_;
    rapidxml::xml_node<>* sheetData_;

    unsigned long long int cellcount_; // count of cells in sheet
    std::vector<xlsxcell> cells_;      // cells in sheet

    double defaultRowHeight_;
    double defaultColWidth_;

    std::vector<double> colWidths_;
    std::vector<double> rowHeights_;

    // lookup tables
    std::map<int, shared_formula> shared_formulas_;
    std::map<std::string, std::string> comments_;

    foo(const std::string& name,
        const std::string& sheet_path,
        xlsxbook& book,
        Rcpp::String comments_path);

    void cacheDefaultRowColDims();
    void cacheColWidths();
    void cacheComments();
    void cacheCellcount();
    void collectCellValues(unsigned long long int& i);

};

#endif
