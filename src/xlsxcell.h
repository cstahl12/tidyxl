#ifndef XLSXCELL_
#define XLSXCELL_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class foo; // Forward declaration because xlsxcell is included in foo

class xlsxcell {

  public:

    xlsxbook&   book_;  // parent workbook
    foo&        sheet_; // parent worksheet
    rapidxml::xml_node<>* cell_;
    unsigned long long int& i_; // position of cell in output vectors

/*     // These variables go to R */
    std::string  sheet_name_; // won't ever be NA
    std::string  address_;    // won't ever be NA
    int          col_;
    int          row_;
/*     Rcpp::String formula_; */
/*     Rcpp::String formula_type_; */
/*     Rcpp::String formula_ref_; */
/*     int          formula_group_; */
/*     Rcpp::String data_type_; */
/*     Rcpp::String error_; */
/*     int          logical_; */
/*     double       numeric_; */
/*     double       date_; */
/*     Rcpp::String character_; */
/*     Rcpp::String comment_; */
/*     double       height_; */
/*     double       width_; */
/*     Rcpp::String style_format_; */
/*     int          local_format_id_; */

    xlsxcell(
      xlsxbook& book,
      foo& sheet,
      rapidxml::xml_node<>* cell,
      unsigned long long int& i
      );

    void initializeValues();
    void cacheAddress();
    void cacheComment();
    void cacheValue();
    void cacheFormula();
};

#endif
