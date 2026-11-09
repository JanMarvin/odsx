#ifndef ODSX_TYPES_H
#define ODSX_TYPES_H

#include <string>
#include <vector>
#include <RcppCommon.h>

// Strings crossing between R and C++. Text is translated to UTF-8 when it
// comes from R and marked as UTF-8 when it goes to R. File and directory names
// are translated to the native encoding, because the file system expects that.
struct utf8_string : std::string {
  utf8_string() = default;
  utf8_string(std::string x) : std::string(std::move(x)) {}
  utf8_string(const char* x) : std::string(x) {}
};

struct native_path : std::string {
  native_path() = default;
  native_path(std::string x) : std::string(std::move(x)) {}
};

namespace Rcpp {
  template <> utf8_string as(SEXP x);
  template <> std::vector<utf8_string> as(SEXP x);
  template <> native_path as(SEXP x);
  template <> SEXP wrap(const utf8_string& x);
  template <> SEXP wrap(const std::vector<utf8_string>& x);
}

#include <Rcpp.h>

inline SEXP utf8_charsxp(const std::string& x) {
  return Rf_mkCharLenCE(x.data(), static_cast<int>(x.size()), CE_UTF8);
}

namespace Rcpp {
  template <> inline utf8_string as(SEXP x) {
    if (TYPEOF(x) != STRSXP || Rf_xlength(x) != 1) Rcpp::stop("Expecting a single string value");
    return utf8_string(Rf_translateCharUTF8(STRING_ELT(x, 0)));
  }

  template <> inline std::vector<utf8_string> as(SEXP x) {
    Rcpp::Shield<SEXP> strs(Rcpp::r_cast<STRSXP>(x));
    std::vector<utf8_string> res;
    res.reserve(static_cast<size_t>(Rf_xlength(strs)));
    for (R_xlen_t i = 0; i < Rf_xlength(strs); ++i) {
      res.push_back(utf8_string(Rf_translateCharUTF8(STRING_ELT(strs, i))));
    }
    return res;
  }

  template <> inline native_path as(SEXP x) {
    if (TYPEOF(x) != STRSXP || Rf_xlength(x) != 1) Rcpp::stop("Expecting a single string value");
    return native_path(Rf_translateChar(STRING_ELT(x, 0)));
  }

  template <> inline SEXP wrap(const utf8_string& x) {
    return Rf_ScalarString(utf8_charsxp(x));
  }

  template <> inline SEXP wrap(const std::vector<utf8_string>& x) {
    Rcpp::Shield<SEXP> res(Rf_allocVector(STRSXP, static_cast<R_xlen_t>(x.size())));
    for (size_t i = 0; i < x.size(); ++i) {
      SET_STRING_ELT(res, static_cast<R_xlen_t>(i), utf8_charsxp(x[i]));
    }
    return res;
  }
}

#include <pugixml/pugixml.hpp>

#endif
