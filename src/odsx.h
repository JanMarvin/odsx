#ifndef ODSX_H
#define ODSX_H

#include <cstdint>
#include <string>

#include "odsx_types.h"

#define MAX_OOXML_COL_INT 16384
#define MAX_OOXML_ROW_INT 1048576

struct xml_string_writer : pugi::xml_writer {
  std::string result;
  void write(const void* data, size_t size) override {
    result.append(static_cast<const char*>(data), size);
  }
};

inline void checkInterrupt(R_xlen_t& iteration, R_xlen_t frequency = 10000) {
  if (iteration % frequency == 0) Rcpp::checkUserInterrupt();
  ++iteration;
}

// 1 -> A, 27 -> AA
template <typename T>
inline std::string int_to_col(T cell) {
  if (cell <= 0 || cell > MAX_OOXML_COL_INT) Rcpp::stop("Column exceeds valid range");
  int64_t value = static_cast<int64_t>(cell);
  std::string col;
  while (value > 0) {
    int64_t modulo = (value - 1) % 26;
    col.insert(col.begin(), static_cast<char>('A' + modulo));
    value = (value - modulo) / 26;
  }
  return col;
}

// the row of a cell reference such as A1 or $A$1
inline int32_t cell_to_rowint(const std::string& ref) {
  int64_t row = 0;
  for (char c : ref) {
    if (c >= '0' && c <= '9') row = row * 10 + (c - '0');
    if (row > MAX_OOXML_ROW_INT) break;
  }
  if (row < 1 || row > MAX_OOXML_ROW_INT) Rcpp::stop("Row exceeds valid range");
  return static_cast<int32_t>(row);
}

// the column of a cell reference such as A1 or $A$1
inline int32_t cell_to_colint(const std::string& ref) {
  int64_t col = 0;
  for (char c : ref) {
    if (c >= 'a' && c <= 'z') c = static_cast<char>(c - 'a' + 'A');
    if (c >= 'A' && c <= 'Z') col = col * 26 + (c - 'A' + 1);
    if (col > MAX_OOXML_COL_INT) break;
  }
  if (col < 1 || col > MAX_OOXML_COL_INT) Rcpp::stop("Column exceeds valid range");
  return static_cast<int32_t>(col);
}

// colors of the indexed palette 8 to 63 as 0xRRGGBB, the indexes 0 to 7 repeat the first eight
static const uint32_t XlsDefaultPalette[56] = {
  0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
  0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
  0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
  0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
  0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
  0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
  0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333};

// number formats that are not stored in the file (ECMA-376 18.8.30)
inline const char* xls_builtin_format(uint32_t id) {
  switch (id) {
    case 1: return "0";
    case 2: return "0.00";
    case 3: return "#,##0";
    case 4: return "#,##0.00";
    case 9: return "0%";
    case 10: return "0.00%";
    case 11: return "0.00E+00";
    case 12: return "# ?/?";
    case 13: return "# ?\?/?\?";
    case 14: return "mm-dd-yy";
    case 15: return "d-mmm-yy";
    case 16: return "d-mmm";
    case 17: return "mmm-yy";
    case 18: return "h:mm AM/PM";
    case 19: return "h:mm:ss AM/PM";
    case 20: return "h:mm";
    case 21: return "h:mm:ss";
    case 22: return "m/d/yy h:mm";
    case 37: return "#,##0 ;(#,##0)";
    case 38: return "#,##0 ;[Red](#,##0)";
    case 39: return "#,##0.00;(#,##0.00)";
    case 40: return "#,##0.00;[Red](#,##0.00)";
    case 45: return "mm:ss";
    case 46: return "[h]:mm:ss";
    case 47: return "mmss.0";
    case 48: return "##0.0E+0";
    case 49: return "@";
    default: return nullptr;
  }
}

#endif
