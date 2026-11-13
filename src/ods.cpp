#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <functional>
#include <iterator>
#include <map>
#include <set>
#include <string>
#include <vector>

#include "odsx.h"

// OpenDocument Spreadsheet (ODF 1.3) conversion. Writing builds content.xml,
// styles.xml, meta.xml, settings.xml, the chart objects and the manifest from
// the workbook parts. Reading converts content.xml, styles.xml, settings.xml
// and meta.xml into pseudo xlsx parts that are loaded like any other xlsx
// file. Comments, images and charts are returned and added to the loaded
// workbook. Elements are matched by their usual prefixes (office:, table:,
// style:, ...), namespaces are not resolved.
//
// Formulas are translated between the xlsx syntax and OpenFormula (of:).
// Formulas of ODF 1.0 and 1.1 (oooc:) are read like OpenFormula, they use the
// same references and separators. Formulas with the msoxl: prefix are read in
// the xlsx syntax. Formulas with other prefixes are dropped on reading and only
// their values are kept.

static const char* const ods_namespaces[][2] = {
  {"xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"},
  {"xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"},
  {"xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"},
  {"xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"},
  {"xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"},
  {"xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"},
  {"xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"},
  {"xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"},
  {"xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"},
  {"xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"},
  {"xmlns:xlink", "http://www.w3.org/1999/xlink"},
  {"xmlns:dc", "http://purl.org/dc/elements/1.1/"},
  {"xmlns:calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"},
  {"xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"},
  {"xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"},
  {"xmlns:ooo", "http://openoffice.org/2004/office"}
};

// Producers may bind the namespaces to other prefixes or use them as default
// namespace. Loaded documents are renamed to these prefixes.
static const char* const ods_prefixes[][2] = {
  {"office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"},
  {"style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"},
  {"text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"},
  {"table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"},
  {"draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"},
  {"fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"},
  {"xlink", "http://www.w3.org/1999/xlink"},
  {"dc", "http://purl.org/dc/elements/1.1/"},
  {"meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"},
  {"number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"},
  {"svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"},
  {"chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"},
  {"dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"},
  {"form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"},
  {"script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"},
  {"config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0"},
  {"presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"},
  {"manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"},
  {"of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"},
  {"ooo", "http://openoffice.org/2004/office"},
  {"oooc", "http://openoffice.org/2004/calc"},
  {"tableooo", "http://openoffice.org/2009/table"},
  {"drawooo", "http://openoffice.org/2010/draw"},
  {"calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"},
  {"loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"}
};

static void ods_normalize_prefixes(pugi::xml_node node, std::map<std::string, std::string> scope) {
  bool declares = false;
  for (pugi::xml_attribute attr : node.attributes()) {
    if (std::strncmp(attr.name(), "xmlns", 5) == 0) declares = true;
  }
  // the prefixes declared here: the default namespace has the empty prefix
  if (declares) {
    for (pugi::xml_attribute attr : node.attributes()) {
      std::string name = attr.name();
      if (name == "xmlns") scope[""] = attr.value();
      else if (name.compare(0, 6, "xmlns:") == 0) scope[name.substr(6)] = attr.value();
    }
  }
  auto rename = [&](const std::string& name, bool element) {
    size_t colon = name.find(':');
    if (colon == std::string::npos && !element) return name;
    std::string prefix = colon == std::string::npos ? "" : name.substr(0, colon);
    if (prefix == "xmlns" || (prefix.empty() && name == "xmlns")) return name;
    std::map<std::string, std::string>::const_iterator uri = scope.find(prefix);
    if (uri == scope.end()) return name;
    for (const auto& known : ods_prefixes) {
      if (uri->second == known[1]) return std::string(known[0]) + ":" + name.substr(colon == std::string::npos ? 0 : colon + 1);
    }
    return name;
  };
  std::string element = rename(node.name(), true);
  if (element != node.name()) node.set_name(element.c_str());
  for (pugi::xml_attribute attr : node.attributes()) {
    std::string name = rename(attr.name(), false);
    if (name != attr.name()) attr.set_name(name.c_str());
  }
  for (pugi::xml_node child : node.children()) {
    if (child.type() == pugi::node_element) ods_normalize_prefixes(child, scope);
  }
}

// renames the elements and attributes of a document only if one of its namespaces has another prefix
static void ods_normalize_document(pugi::xml_document& doc) {
  bool other = false;
  std::function<void(pugi::xml_node)> check = [&](pugi::xml_node node) {
    for (pugi::xml_attribute attr : node.attributes()) {
      std::string name = attr.name();
      if (name != "xmlns" && name.compare(0, 6, "xmlns:") != 0) continue;
      std::string prefix = name == "xmlns" ? "" : name.substr(6);
      for (const auto& known : ods_prefixes) {
        if (attr.value() == std::string(known[1]) && prefix != known[0]) other = true;
      }
    }
    for (pugi::xml_node child : node.children()) {
      if (!other && child.type() == pugi::node_element) check(child);
    }
  };
  check(doc.document_element());
  if (!other) return;
  ods_normalize_prefixes(doc.document_element(), std::map<std::string, std::string>());
  for (const auto& known : ods_prefixes) {
    std::string decl = std::string("xmlns:") + known[0];
    if (!doc.document_element().attribute(decl.c_str())) doc.document_element().append_attribute(decl.c_str()) = known[1];
  }
}

// xlsx column widths are multiples of the maximum digit width of the default
// font in pixels at 96 dpi, ODF widths are lengths. The digit widths are the
// advance widths of the digits relative to the font size.
static double ods_digit_width(const std::string& font, double size) {
  static const std::pair<const char*, double> digits[] = {
    {"Calibri", 0.507}, {"Carlito", 0.507}, {"Arial", 0.556}, {"Helvetica", 0.556}, {"Liberation Sans", 0.556},
    {"Times New Roman", 0.5}, {"Liberation Serif", 0.5}, {"Courier New", 0.6}, {"Liberation Mono", 0.6},
    {"Verdana", 0.636}, {"DejaVu Sans", 0.636}, {"Tahoma", 0.546}
  };
  double em = 0.507;
  for (const std::pair<const char*, double>& digit : digits) {
    if (font == digit.first) em = digit.second;
  }
  if (!(size > 0.0)) size = 11.0;
  return std::max(1.0, std::floor(em * size * 96.0 / 72.0 + 0.5));
}

enum ods_kind {
  ods_kind_number,
  ods_kind_percent,
  ods_kind_currency,
  ods_kind_date,
  ods_kind_time,
  ods_kind_text
};

static std::string ods_num(double x) {
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%.15g", x);
  return buf;
}

static double ods_length_pt(const char* x) {
  char* end = nullptr;
  double val = std::strtod(x, &end);
  if (end == x) return -1.0;
  std::string unit(end);
  if (unit == "%") return -1.0;
  if (unit == "in") return val * 72.0;
  if (unit == "cm") return val * 72.0 / 2.54;
  if (unit == "mm") return val * 72.0 / 25.4;
  if (unit == "pc") return val * 12.0;
  if (unit == "px") return val * 0.75;
  return val;
}

static void ods_declaration(pugi::xml_document& doc) {
  pugi::xml_node decl = doc.prepend_child(pugi::node_declaration);
  decl.append_attribute("version") = "1.0";
  decl.append_attribute("encoding") = "UTF-8";
}

static void ods_save(pugi::xml_document& doc, const std::string& path) {
  if (!doc.save_file(path.c_str(), "", pugi::format_raw, pugi::encoding_utf8))
    Rcpp::stop("could not write %s", path);
}

// a vector of xml fragments, all parsed into one document
// the descendants of node named name in document order, all nodes for an empty name.
// Matching elements and elements named skip are not searched further.
static std::vector<pugi::xml_node> ods_descendants(pugi::xml_node node, const char* name, const char* skip = "") {
  std::vector<pugi::xml_node> out;
  pugi::xml_node n = node.first_child();
  while (n) {
    bool skipped = *skip != '\0' && std::strcmp(n.name(), skip) == 0;
    bool matched = !skipped && (*name == '\0' || std::strcmp(n.name(), name) == 0);
    if (matched) out.push_back(n);
    if (!skipped && (!matched || *name == '\0') && n.first_child()) {
      n = n.first_child();
      continue;
    }
    while (n != node && !n.next_sibling()) n = n.parent();
    if (n == node) break;
    n = n.next_sibling();
  }
  return out;
}

// the first descendant of node named name
static pugi::xml_node ods_descendant(pugi::xml_node node, const char* name) {
  return node.find_node([name](pugi::xml_node n) { return std::strcmp(n.name(), name) == 0; });
}

static std::string ods_xml_string(pugi::xml_node node) {
  xml_string_writer writer;
  node.print(writer, "", pugi::format_raw);
  return writer.result;
}

static std::vector<pugi::xml_node> ods_load_list(pugi::xml_document& doc, const std::vector<utf8_string>& x) {
  std::string buf = "<x>";
  for (const std::string& s : x) buf += s;
  buf += "</x>";
  std::vector<pugi::xml_node> nodes;
  if (!doc.load_string(buf.c_str(), pugi::parse_default | pugi::parse_ws_pcdata_single)) return nodes;
  for (pugi::xml_node node : doc.child("x").children()) nodes.push_back(node);
  return nodes;
}

static size_t ods_utf8_len(const std::string& s, size_t pos) {
  size_t len = 1;
  while (pos + len < s.size() && (static_cast<unsigned char>(s[pos + len]) & 0xC0u) == 0x80u) ++len;
  return len;
}

static std::string ods_lower(std::string x) {
  for (char& c : x) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
  return x;
}

// dates. Serial numbers count days since 1899-12-30, or 1904-01-01.

static int64_t ods_days_from_civil(int64_t y, int64_t m, int64_t d) {
  if (m <= 2) y -= 1;
  const int64_t era = (y >= 0 ? y : y - 399) / 400;
  const int64_t yoe = y - era * 400;
  const int64_t doy = (153 * (m > 2 ? m - 3 : m + 9) + 2) / 5 + d - 1;
  const int64_t doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
  return era * 146097 + doe - 719468;
}

static void ods_civil_from_days(int64_t z, int64_t& y, int64_t& m, int64_t& d) {
  z += 719468;
  const int64_t era = (z >= 0 ? z : z - 146096) / 146097;
  const int64_t doe = z - era * 146097;
  const int64_t yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
  const int64_t doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
  const int64_t mp = (5 * doy + 2) / 153;
  d = doy - (153 * mp + 2) / 5 + 1;
  m = mp < 10 ? mp + 3 : mp - 9;
  y = yoe + era * 400 + (m <= 2 ? 1 : 0);
}

static int64_t ods_epoch(bool date1904) {
  return date1904 ? ods_days_from_civil(1904, 1, 1) : ods_days_from_civil(1899, 12, 30);
}

static std::string ods_date_value(double serial, bool date1904) {
  int64_t ms = std::llround(serial * 86400000.0);
  int64_t days = ms / 86400000;
  int64_t rest = ms % 86400000;
  if (rest < 0) {
    rest += 86400000;
    days -= 1;
  }
  int64_t y = 0, m = 0, d = 0;
  ods_civil_from_days(ods_epoch(date1904) + days, y, m, d);
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%04d-%02d-%02d", static_cast<int>(y), static_cast<int>(m), static_cast<int>(d));
  std::string out = buf;
  if (rest > 0) {
    int32_t secs = static_cast<int32_t>(rest / 1000);
    int32_t frac = static_cast<int32_t>(rest % 1000);
    std::snprintf(buf, sizeof(buf), "T%02d:%02d:%02d", secs / 3600, (secs / 60) % 60, secs % 60);
    out += buf;
    if (frac > 0) {
      std::snprintf(buf, sizeof(buf), ".%03d", frac);
      out += buf;
    }
  }
  return out;
}

static std::string ods_time_value(double serial) {
  int64_t ms = std::llround(serial * 86400000.0);
  int64_t secs = ms / 1000;
  int64_t frac = ms % 1000;
  std::string out = "PT" + std::to_string(secs / 3600) + "H" + std::to_string((secs / 60) % 60) + "M" + std::to_string(secs % 60);
  if (frac > 0) {
    char buf[16];
    std::snprintf(buf, sizeof(buf), ".%03d", static_cast<int>(frac));
    out += buf;
  }
  return out + "S";
}

static bool ods_parse_date(const char* x, bool date1904, double& serial) {
  char* end = nullptr;
  int64_t y = std::strtoll(x, &end, 10);
  if (end == x || *end != '-') return false;
  const char* p = end + 1;
  int64_t mo = std::strtoll(p, &end, 10);
  if (end == p || *end != '-') return false;
  p = end + 1;
  int64_t d = std::strtoll(p, &end, 10);
  if (end == p) return false;
  double secs = 0.0;
  if (*end == 'T') {
    p = end + 1;
    int64_t hh = std::strtoll(p, &end, 10);
    int64_t mi = 0;
    double ss = 0.0;
    if (*end == ':') {
      p = end + 1;
      mi = std::strtoll(p, &end, 10);
    }
    if (*end == ':') {
      p = end + 1;
      ss = std::strtod(p, &end);
    }
    secs = static_cast<double>(hh * 3600 + mi * 60) + ss;
  }
  serial = static_cast<double>(ods_days_from_civil(y, mo, d) - ods_epoch(date1904)) + secs / 86400.0;
  return true;
}

static bool ods_parse_time(const char* x, double& serial) {
  const char* p = x;
  bool neg = false;
  if (*p == '-') {
    neg = true;
    ++p;
  }
  if (*p != 'P') return false;
  ++p;
  bool in_time = false;
  double days = 0.0;
  while (*p) {
    if (*p == 'T') {
      in_time = true;
      ++p;
      continue;
    }
    char* end = nullptr;
    double val = std::strtod(p, &end);
    if (end == p) return false;
    char unit = *end;
    if (unit == 'D') days += val;
    else if (unit == 'H' && in_time) days += val / 24.0;
    else if (unit == 'M' && in_time) days += val / 1440.0;
    else if (unit == 'S' && in_time) days += val / 86400.0;
    else if (unit == '\0') break;
    p = end + 1;
  }
  serial = neg ? -days : days;
  return true;
}

static std::string ods_hex_tint(const std::string& hex, double tint) {
  uint32_t rgb = static_cast<uint32_t>(std::strtoul(hex.c_str(), nullptr, 16));
  double r = static_cast<double>((rgb >> 16) & 0xFFu) / 255.0;
  double g = static_cast<double>((rgb >> 8) & 0xFFu) / 255.0;
  double b = static_cast<double>(rgb & 0xFFu) / 255.0;
  double mx = std::max(r, std::max(g, b));
  double mn = std::min(r, std::min(g, b));
  double l = (mx + mn) / 2.0;
  double h = 0.0, s = 0.0, delta = mx - mn;
  if (delta > 0.0) {
    s = l > 0.5 ? delta / (2.0 - mx - mn) : delta / (mx + mn);
    if (mx == r) h = (g - b) / delta + (g < b ? 6.0 : 0.0);
    else if (mx == g) h = (b - r) / delta + 2.0;
    else h = (r - g) / delta + 4.0;
    h /= 6.0;
  }
  l = tint < 0.0 ? l * (1.0 + tint) : l * (1.0 - tint) + tint;

  double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
  double p = 2.0 * l - q;
  double rgb_out[3];
  for (int32_t k = 0; k < 3; ++k) {
    double t = h + (1.0 - static_cast<double>(k)) / 3.0;
    if (t < 0.0) t += 1.0;
    if (t > 1.0) t -= 1.0;
    double v = p;
    if (t < 1.0 / 6.0) v = p + (q - p) * 6.0 * t;
    else if (t < 0.5) v = q;
    else if (t < 2.0 / 3.0) v = p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    rgb_out[k] = s > 0.0 ? v : l;
  }
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%02x%02x%02x",
                static_cast<unsigned int>(std::lround(rgb_out[0] * 255.0)),
                static_cast<unsigned int>(std::lround(rgb_out[1] * 255.0)),
                static_cast<unsigned int>(std::lround(rgb_out[2] * 255.0)));
  return buf;
}

// "#rrggbb" for a color element, empty for automatic colors
static std::string ods_color(pugi::xml_node color, const std::vector<std::string>& theme, const std::vector<std::string>& palette) {
  std::string hex;
  if (pugi::xml_attribute rgb = color.attribute("rgb")) {
    hex = rgb.as_string();
    if (hex.size() == 8) hex = hex.substr(2);
  } else if (pugi::xml_attribute th = color.attribute("theme")) {
    // the first four theme colors are dk1, lt1, dk2, lt2, but referenced as lt1, dk1, lt2, dk2
    size_t idx = th.as_uint();
    if (idx < 4) idx ^= 1u;
    if (idx < theme.size()) hex = theme[idx];
  } else if (pugi::xml_attribute ind = color.attribute("indexed")) {
    size_t idx = ind.as_uint();
    if (idx < palette.size()) hex = palette[idx];
  }
  if (hex.size() != 6) return "";
  double tint = color.attribute("tint").as_double();
  if (tint != 0.0) hex = ods_hex_tint(hex, tint);
  return "#" + ods_lower(hex);
}

// number formats: xlsx format code -> ODF data style

enum ods_tok_type {
  ods_tok_lit,
  ods_tok_digits,
  ods_tok_exp,
  ods_tok_date,
  ods_tok_elapsed,
  ods_tok_ampm,
  ods_tok_percent,
  ods_tok_at,
  ods_tok_general,
  ods_tok_currency,
  ods_tok_color
};

struct ods_tok {
  ods_tok_type type;
  std::string text;
  size_t len;
};

static std::vector<std::string> ods_sections(const std::string& code) {
  std::vector<std::string> out;
  std::string cur;
  bool quoted = false, bracket = false;
  for (size_t i = 0; i < code.size(); ++i) {
    char c = code[i];
    if (!quoted && !bracket && c == '\\' && i + 1 < code.size()) {
      cur += c;
      cur += code[++i];
      continue;
    }
    if (c == '"' && !bracket) quoted = !quoted;
    if (c == '[' && !quoted) bracket = true;
    if (c == ']' && !quoted) bracket = false;
    if (c == ';' && !quoted && !bracket) {
      out.push_back(cur);
      cur.clear();
      continue;
    }
    cur += c;
  }
  out.push_back(cur);
  return out;
}

static std::vector<ods_tok> ods_tokenize(const std::string& s) {
  static const char* const colors[] = {"black", "blue", "cyan", "green", "magenta", "red", "white", "yellow"};
  std::vector<ods_tok> toks;
  size_t n = s.size();
  size_t i = 0;
  while (i < n) {
    char c = s[i];
    char lc = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    ods_tok tok = {ods_tok_lit, "", 0};

    if (c == '"') {
      size_t end = s.find('"', i + 1);
      if (end == std::string::npos) end = n;
      tok.text = s.substr(i + 1, end - i - 1);
      i = end + 1;
    } else if (c == '\\' && i + 1 < n) {
      size_t len = ods_utf8_len(s, i + 1);
      tok.text = s.substr(i + 1, len);
      i += 1 + len;
    } else if (c == '_' && i + 1 < n) {
      tok.text = " ";
      i += 1 + ods_utf8_len(s, i + 1);
    } else if (c == '*' && i + 1 < n) {
      i += 1 + ods_utf8_len(s, i + 1);
      continue;
    } else if (c == '[') {
      size_t end = s.find(']', i);
      if (end == std::string::npos) end = n;
      std::string inner = s.substr(i + 1, end - i - 1);
      i = end + 1;
      std::string lower = ods_lower(inner);
      if (!inner.empty() && inner[0] == '$') {
        size_t dash = inner.find('-');
        tok.type = ods_tok_currency;
        tok.text = inner.substr(1, dash == std::string::npos ? std::string::npos : dash - 1);
        if (tok.text.empty()) continue;
      } else if (!lower.empty() && (lower[0] == 'h' || lower[0] == 'm' || lower[0] == 's') &&
                 lower.find_first_not_of(lower[0]) == std::string::npos) {
        tok.type = ods_tok_elapsed;
        tok.text = lower.substr(0, 1);
        tok.len = lower.size();
      } else if (std::find(std::begin(colors), std::end(colors), lower) != std::end(colors)) {
        tok.type = ods_tok_color;
        tok.text = lower;
      } else {
        continue;
      }
    } else if (c == '0' || c == '#' || c == '?' ||
               ((c == '.' || c == ',') && i + 1 < n && std::strchr("0#?", s[i + 1]) != nullptr && s[i + 1] != '\0')) {
      size_t j = i;
      while (j < n && std::strchr("0#?.,", s[j]) != nullptr) ++j;
      if (j + 1 < n && s[j] == '/' && std::strchr("0#?123456789", s[j + 1]) != nullptr) {
        ++j;
        while (j < n && std::strchr("0#?0123456789", s[j]) != nullptr) ++j;
      }
      tok.type = ods_tok_digits;
      tok.text = s.substr(i, j - i);
      i = j;
    } else if ((c == 'E' || c == 'e') && i + 1 < n && (s[i + 1] == '+' || s[i + 1] == '-')) {
      size_t j = i + 2;
      while (j < n && std::strchr("0#?", s[j]) != nullptr) ++j;
      tok.type = ods_tok_exp;
      tok.len = j - i - 2;
      i = j;
    } else if (ods_lower(s.substr(i, 5)) == "am/pm") {
      tok.type = ods_tok_ampm;
      i += 5;
    } else if (ods_lower(s.substr(i, 3)) == "a/p") {
      tok.type = ods_tok_ampm;
      i += 3;
    } else if (lc == 'y' || lc == 'm' || lc == 'd' || lc == 'h' || lc == 's') {
      size_t j = i;
      while (j < n && std::tolower(static_cast<unsigned char>(s[j])) == lc) ++j;
      tok.type = ods_tok_date;
      tok.text = std::string(1, lc);
      tok.len = j - i;
      i = j;
    } else if (c == '%') {
      tok.type = ods_tok_percent;
      ++i;
    } else if (c == '@') {
      tok.type = ods_tok_at;
      ++i;
    } else if (ods_lower(s.substr(i, 7)) == "general") {
      tok.type = ods_tok_general;
      i += 7;
    } else {
      size_t len = ods_utf8_len(s, i);
      tok.text = s.substr(i, len);
      i += len;
    }

    if (tok.type == ods_tok_lit && !toks.empty() && toks.back().type == ods_tok_lit) {
      toks.back().text += tok.text;
    } else {
      toks.push_back(tok);
    }
  }
  return toks;
}

static void ods_append_text(pugi::xml_node style, const std::string& text) {
  if (text.empty()) return;
  pugi::xml_node last = style.last_child();
  if (std::strcmp(last.name(), "number:text") == 0) {
    std::string merged = std::string(last.text().get()) + text;
    last.text().set(merged.c_str());
  } else {
    style.append_child("number:text").text().set(text.c_str());
  }
}

static ods_kind ods_data_style(pugi::xml_node parent, const std::string& name, const std::string& section, pugi::xml_node& style) {
  std::vector<ods_tok> toks = ods_tokenize(section);

  // m is a minute after an hour or before a second, otherwise a month
  std::vector<bool> minute(toks.size(), false);
  bool has_date = false, has_ymd = false, has_digits = false, has_pct = false;
  bool has_exp = false, has_frac = false, has_cur = false, has_at = false, elapsed = false;
  std::string color;
  for (size_t k = 0; k < toks.size(); ++k) {
    const ods_tok& tok = toks[k];
    if (tok.type == ods_tok_date || tok.type == ods_tok_ampm || tok.type == ods_tok_elapsed) has_date = true;
    if (tok.type == ods_tok_elapsed) elapsed = true;
    if (tok.type == ods_tok_digits) has_digits = true;
    if (tok.type == ods_tok_digits && tok.text.find('/') != std::string::npos) has_frac = true;
    if (tok.type == ods_tok_percent) has_pct = true;
    if (tok.type == ods_tok_exp) has_exp = true;
    if (tok.type == ods_tok_currency) has_cur = true;
    if (tok.type == ods_tok_at) has_at = true;
    if (tok.type == ods_tok_color) color = tok.text;
    if (tok.type != ods_tok_date) continue;
    if (tok.text == "y" || tok.text == "d") has_ymd = true;
    if (tok.text != "m") continue;

    bool prev_hour = false, next_second = false;
    for (size_t p = k; p-- > 0;) {
      if (toks[p].type == ods_tok_date || toks[p].type == ods_tok_elapsed) {
        prev_hour = toks[p].text == "h";
        break;
      }
    }
    for (size_t q = k + 1; q < toks.size(); ++q) {
      if (toks[q].type == ods_tok_date || toks[q].type == ods_tok_elapsed) {
        next_second = toks[q].text == "s";
        break;
      }
    }
    minute[k] = prev_hour || next_second;
    if (!minute[k]) has_ymd = true;
  }

  ods_kind kind = ods_kind_number;
  const char* element = "number:number-style";
  if (has_date) {
    kind = has_ymd ? ods_kind_date : ods_kind_time;
    element = has_ymd ? "number:date-style" : "number:time-style";
  } else if (has_pct) {
    kind = ods_kind_percent;
    element = "number:percentage-style";
  } else if (has_cur && !has_exp && !has_frac) {
    kind = ods_kind_currency;
    element = "number:currency-style";
  } else if (has_at && !has_digits) {
    kind = ods_kind_text;
    element = "number:text-style";
  }

  style = parent.append_child(element);
  style.append_attribute("style:name") = name.c_str();
  if (elapsed) style.append_attribute("number:truncate-on-overflow") = "false";

  if (!color.empty()) {
    const char* hex = "#000000";
    if (color == "blue") hex = "#0000ff";
    else if (color == "cyan") hex = "#00ffff";
    else if (color == "green") hex = "#00ff00";
    else if (color == "magenta") hex = "#ff00ff";
    else if (color == "red") hex = "#ff0000";
    else if (color == "white") hex = "#ffffff";
    else if (color == "yellow") hex = "#ffff00";
    style.append_child("style:text-properties").append_attribute("fo:color") = hex;
  }

  pugi::xml_node last_number, last_seconds;
  for (size_t k = 0; k < toks.size(); ++k) {
    const ods_tok& tok = toks[k];
    switch (tok.type) {
      case ods_tok_lit:
        ods_append_text(style, tok.text);
        break;
      case ods_tok_percent:
        ods_append_text(style, "%");
        break;
      case ods_tok_at:
        if (kind == ods_kind_text) style.append_child("number:text-content");
        break;
      case ods_tok_currency:
        if (kind == ods_kind_currency) style.append_child("number:currency-symbol").text().set(tok.text.c_str());
        else ods_append_text(style, tok.text);
        break;
      case ods_tok_general:
        if (kind == ods_kind_date || kind == ods_kind_time || kind == ods_kind_text) break;
        last_number = style.append_child("number:number");
        last_number.append_attribute("number:min-integer-digits") = 1;
        break;
      case ods_tok_digits: {
        std::string d = tok.text;
        if (kind == ods_kind_date || kind == ods_kind_time) {
          if (last_seconds && d[0] == '.') {
            last_seconds.append_attribute("number:decimal-places") = static_cast<uint32_t>(d.size() - 1);
          } else {
            ods_append_text(style, d);
          }
          break;
        }

        size_t slash = d.find('/');
        if (slash != std::string::npos) {
          pugi::xml_node frac = style.append_child("number:fraction");
          std::string num = d.substr(0, slash), den = d.substr(slash + 1);
          frac.append_attribute("number:min-numerator-digits") = static_cast<uint32_t>(num.size());
          if (den.find_first_of("123456789") != std::string::npos) {
            frac.append_attribute("number:denominator-value") = std::atoi(den.c_str());
          } else {
            frac.append_attribute("number:min-denominator-digits") = static_cast<uint32_t>(den.size());
          }
          // "# ?/?": the preceding number and the space are the integer part
          pugi::xml_node space = last_number.next_sibling();
          if (last_number && space && std::strcmp(space.name(), "number:text") == 0 &&
              std::string(space.text().get()) == " " && space.next_sibling() == frac) {
            frac.append_attribute("number:min-integer-digits") = last_number.attribute("number:min-integer-digits").as_uint();
            style.remove_child(space);
            style.remove_child(last_number);
          }
          last_number = pugi::xml_node();
          break;
        }

        size_t factor = 0;
        while (!d.empty() && d.back() == ',') {
          d.pop_back();
          ++factor;
        }
        size_t dot = d.find('.');
        std::string ip = d.substr(0, dot);
        std::string dp = dot == std::string::npos ? "" : d.substr(dot + 1);
        uint32_t min_int = static_cast<uint32_t>(std::count(ip.begin(), ip.end(), '0'));
        uint32_t decimals = static_cast<uint32_t>(dp.size() - static_cast<size_t>(std::count(dp.begin(), dp.end(), ',')));
        uint32_t min_dec = static_cast<uint32_t>(std::count(dp.begin(), dp.end(), '0'));
        bool grouping = ip.find(',') != std::string::npos;

        if (k + 1 < toks.size() && toks[k + 1].type == ods_tok_exp) {
          pugi::xml_node sci = style.append_child("number:scientific-number");
          sci.append_attribute("number:decimal-places") = decimals;
          sci.append_attribute("number:min-decimal-places") = min_dec;
          sci.append_attribute("number:min-integer-digits") = min_int;
          sci.append_attribute("number:min-exponent-digits") = static_cast<uint32_t>(toks[k + 1].len);
          ++k;
          break;
        }

        last_number = style.append_child("number:number");
        last_number.append_attribute("number:decimal-places") = decimals;
        last_number.append_attribute("number:min-decimal-places") = min_dec;
        last_number.append_attribute("number:min-integer-digits") = min_int;
        if (grouping) last_number.append_attribute("number:grouping") = "true";
        if (factor > 0) {
          double display_factor = std::pow(1000.0, static_cast<double>(factor));
          last_number.append_attribute("number:display-factor") = ods_num(display_factor).c_str();
        }
        break;
      }
      case ods_tok_date:
      case ods_tok_elapsed: {
        bool long_style = tok.len >= 2;
        pugi::xml_node part;
        if (tok.text == "y") {
          part = style.append_child("number:year");
          long_style = tok.len > 2;
        } else if (tok.text == "m" && !minute[k] && tok.type == ods_tok_date) {
          part = style.append_child("number:month");
          if (tok.len >= 3) part.append_attribute("number:textual") = "true";
          long_style = tok.len == 2 || tok.len == 4;
        } else if (tok.text == "m") {
          part = style.append_child("number:minutes");
        } else if (tok.text == "d") {
          part = style.append_child(tok.len >= 3 ? "number:day-of-week" : "number:day");
          long_style = tok.len == 2 || tok.len >= 4;
        } else if (tok.text == "h") {
          part = style.append_child("number:hours");
        } else {
          part = style.append_child("number:seconds");
          last_seconds = part;
        }
        if (long_style) part.append_attribute("number:style") = "long";
        break;
      }
      case ods_tok_ampm:
        style.append_child("number:am-pm");
        break;
      case ods_tok_exp:
      case ods_tok_color:
        break;
    }
  }
  return kind;
}

// the condition of a section, e.g. "<100" for "[<100]0.0"
static std::string ods_condition(const std::string& section) {
  bool quoted = false;
  for (size_t i = 0; i < section.size(); ++i) {
    char c = section[i];
    if (c == '"') quoted = !quoted;
    if (quoted || c != '[' || i + 1 >= section.size()) continue;
    if (section[i + 1] != '<' && section[i + 1] != '>' && section[i + 1] != '=') continue;
    size_t end = section.find(']', i);
    if (end == std::string::npos) return "";
    return section.substr(i + 1, end - i - 1);
  }
  return "";
}

static ods_kind ods_number_format(pugi::xml_node parent, uint32_t id, const std::string& code) {
  std::vector<std::string> sections = ods_sections(code);
  std::string base = "N" + std::to_string(id);
  pugi::xml_node style, unused;

  if (sections.size() == 1) return ods_data_style(parent, base, sections[0], style);

  std::string cond0 = ods_condition(sections[0]);
  std::string cond1 = sections.size() >= 3 ? ods_condition(sections[1]) : "";
  if (cond0.empty()) cond0 = sections.size() == 2 ? ">=0" : ">0";
  if (cond1.empty()) cond1 = "<0";

  ods_kind kind = ods_data_style(parent, base + "P0", sections[0], unused);
  if (sections.size() >= 3) ods_data_style(parent, base + "P1", sections[1], unused);
  ods_data_style(parent, base, sections.size() == 2 ? sections[1] : sections[2], style);

  pugi::xml_node map = style.append_child("style:map");
  map.append_attribute("style:condition") = ("value()" + cond0).c_str();
  map.append_attribute("style:apply-style-name") = (base + "P0").c_str();
  if (sections.size() >= 3) {
    map = style.append_child("style:map");
    map.append_attribute("style:condition") = ("value()" + cond1).c_str();
    map.append_attribute("style:apply-style-name") = (base + "P1").c_str();
  }
  return kind;
}

// cell styles: xf -> style:style

static std::string ods_border(pugi::xml_node side, const std::vector<std::string>& theme, const std::vector<std::string>& palette) {
  std::string kind = side.attribute("style").as_string();
  if (kind.empty() || kind == "none") return "";
  std::string color = ods_color(side.child("color"), theme, palette);
  if (color.empty()) color = "#000000";

  const char* line = "0.74pt solid";
  if (kind == "hair") line = "0.06pt solid";
  else if (kind == "medium") line = "2pt solid";
  else if (kind == "thick") line = "thick solid";
  else if (kind == "double") line = "2.6pt double";
  else if (kind == "dotted") line = "0.74pt dotted";
  else if (kind == "dashed" || kind == "dashDot" || kind == "dashDotDot") line = "0.74pt dashed";
  else if (kind == "mediumDashed" || kind == "mediumDashDot" || kind == "mediumDashDotDot" || kind == "slantDashDot") line = "2pt dashed";
  return std::string(line) + " " + color;
}

// For differential formats (dxf) a pattern fill without type is solid and the
// color of a solid fill is the background color.
// the text properties of a font or of the run properties of rich text
static void ods_font_properties(pugi::xml_node tp, pugi::xml_node font,
                                const std::vector<std::string>& theme, const std::vector<std::string>& palette) {
  pugi::xml_node name = font.child("name");
  if (!name) name = font.child("rFont");
  if (name) {
    std::string family = name.attribute("val").as_string();
    if (family.find(' ') != std::string::npos) family = "'" + family + "'";
    tp.append_attribute("fo:font-family") = family.c_str();
  }
  if (pugi::xml_node sz = font.child("sz")) {
    tp.append_attribute("fo:font-size") = (std::string(sz.attribute("val").as_string()) + "pt").c_str();
  }
  if (font.child("b") && font.child("b").attribute("val").as_bool(true)) tp.append_attribute("fo:font-weight") = "bold";
  if (font.child("i") && font.child("i").attribute("val").as_bool(true)) tp.append_attribute("fo:font-style") = "italic";
  if (pugi::xml_node u = font.child("u")) {
    std::string val = u.attribute("val").as_string();
    if (val != "none") {
      tp.append_attribute("style:text-underline-style") = "solid";
      tp.append_attribute("style:text-underline-width") = "auto";
      tp.append_attribute("style:text-underline-color") = "font-color";
      if (val == "double" || val == "doubleAccounting") tp.append_attribute("style:text-underline-type") = "double";
    }
  }
  if (font.child("strike") && font.child("strike").attribute("val").as_bool(true)) {
    tp.append_attribute("style:text-line-through-style") = "solid";
  }
  if (font.child("outline") && font.child("outline").attribute("val").as_bool(true)) tp.append_attribute("style:text-outline") = "true";
  if (font.child("shadow") && font.child("shadow").attribute("val").as_bool(true)) tp.append_attribute("fo:text-shadow") = "1pt 1pt";
  std::string vert_align = font.child("vertAlign").attribute("val").as_string();
  if (vert_align == "superscript") tp.append_attribute("style:text-position") = "super 58%";
  if (vert_align == "subscript") tp.append_attribute("style:text-position") = "sub 58%";
  std::string color = ods_color(font.child("color"), theme, palette);
  if (!color.empty()) tp.append_attribute("fo:color") = color.c_str();
}

static void ods_cell_style(pugi::xml_node style, pugi::xml_node xf, pugi::xml_node font, pugi::xml_node fill, pugi::xml_node border,
                           const std::vector<std::string>& theme, const std::vector<std::string>& palette, bool dxf = false) {
  pugi::xml_node tcp = style.append_child("style:table-cell-properties");
  pugi::xml_node pp = style.append_child("style:paragraph-properties");
  pugi::xml_node tp = style.append_child("style:text-properties");

  pugi::xml_node pattern = fill.child("patternFill");
  std::string pattern_type = pattern.attribute("patternType").as_string();
  std::string background;
  if (dxf && pattern && pattern_type.empty()) pattern_type = "solid";
  if (pattern && !pattern_type.empty() && pattern_type != "none") {
    background = ods_color(pattern.child("fgColor"), theme, palette);
    if (background.empty() && (pattern_type != "solid" || dxf)) background = ods_color(pattern.child("bgColor"), theme, palette);
  } else if (pugi::xml_node gradient = fill.child("gradientFill")) {
    background = ods_color(gradient.child("stop").child("color"), theme, palette);
  }
  if (!background.empty()) tcp.append_attribute("fo:background-color") = background.c_str();

  static const char* const sides[4][2] = {
    {"left", "fo:border-left"}, {"right", "fo:border-right"}, {"top", "fo:border-top"}, {"bottom", "fo:border-bottom"}
  };
  for (const auto& side : sides) {
    std::string line = ods_border(border.child(side[0]), theme, palette);
    if (!line.empty()) tcp.append_attribute(side[1]) = line.c_str();
  }
  std::string diagonal = ods_border(border.child("diagonal"), theme, palette);
  if (!diagonal.empty()) {
    if (border.attribute("diagonalDown").as_bool()) tcp.append_attribute("style:diagonal-tl-br") = diagonal.c_str();
    if (border.attribute("diagonalUp").as_bool()) tcp.append_attribute("style:diagonal-bl-tr") = diagonal.c_str();
  }

  if (pugi::xml_node align = xf.child("alignment")) {
    std::string horizontal = align.attribute("horizontal").as_string();
    const char* text_align = nullptr;
    if (horizontal == "left" || horizontal == "fill") text_align = "start";
    else if (horizontal == "center" || horizontal == "centerContinuous") text_align = "center";
    else if (horizontal == "right") text_align = "end";
    else if (horizontal == "justify" || horizontal == "distributed") text_align = "justify";
    if (text_align) {
      tcp.append_attribute("style:text-align-source") = "fix";
      pp.append_attribute("fo:text-align") = text_align;
    }
    if (horizontal == "fill") tcp.append_attribute("style:repeat-content") = "true";

    std::string vertical = align.attribute("vertical").as_string();
    if (vertical == "top") tcp.append_attribute("style:vertical-align") = "top";
    else if (vertical == "center" || vertical == "justify" || vertical == "distributed") tcp.append_attribute("style:vertical-align") = "middle";
    else if (vertical == "bottom") tcp.append_attribute("style:vertical-align") = "bottom";

    if (align.attribute("wrapText").as_bool()) tcp.append_attribute("fo:wrap-option") = "wrap";
    if (align.attribute("shrinkToFit").as_bool()) tcp.append_attribute("style:shrink-to-fit") = "true";

    uint32_t rotation = align.attribute("textRotation").as_uint();
    if (rotation == 255) {
      tcp.append_attribute("style:direction") = "ttb";
    } else if (rotation > 0 && rotation <= 90) {
      tcp.append_attribute("style:rotation-angle") = rotation;
    } else if (rotation > 90 && rotation <= 180) {
      tcp.append_attribute("style:rotation-angle") = 450 - rotation;
    }

    uint32_t indent = align.attribute("indent").as_uint();
    if (indent > 0) pp.append_attribute("fo:margin-left") = (std::to_string(indent * 9) + "pt").c_str();
  }

  if (pugi::xml_node protection = xf.child("protection")) {
    bool locked = protection.attribute("locked").as_bool(true);
    bool hidden = protection.attribute("hidden").as_bool(false);
    const char* protect = "protected";
    if (!locked && !hidden) protect = "none";
    else if (!locked && hidden) protect = "formula-hidden";
    else if (locked && hidden) protect = "hidden-and-protected";
    tcp.append_attribute("style:cell-protect") = protect;
  }

  ods_font_properties(tp, font, theme, palette);

  if (!tcp.first_attribute()) style.remove_child(tcp);
  if (!pp.first_attribute()) style.remove_child(pp);
  if (!tp.first_attribute()) style.remove_child(tp);
}

// cell text: line breaks become paragraphs, spaces that ODF consumers would
// collapse become text:s. Text with a style is written as span. A run
// continues the last paragraph of the cell.
static void ods_cell_text(pugi::xml_node cell, const std::string& txt, const std::string& style = "", bool run = false) {
  pugi::xml_node p = run && std::strcmp(cell.last_child().name(), "text:p") == 0 ? cell.last_child() : cell.append_child("text:p");
  pugi::xml_node out = p;
  if (!style.empty()) {
    out = p.append_child("text:span");
    out.append_attribute("text:style-name") = style.c_str();
  }
  std::string chars;
  size_t n = txt.size();
  for (size_t i = 0; i < n; ++i) {
    char c = txt[i];
    if (c == '\r') continue;
    if (c != '\n' && c != '\t' && c != ' ') {
      chars += c;
      continue;
    }
    if (c == ' ' && !chars.empty() && i + 1 < n && txt[i + 1] != ' ' && txt[i + 1] != '\n') {
      chars += c;
      continue;
    }
    if (!chars.empty()) {
      out.append_child(pugi::node_pcdata).set_value(chars.c_str());
      chars.clear();
    }
    if (c == '\n') {
      p = cell.append_child("text:p");
      out = p;
      if (!style.empty()) {
        out = p.append_child("text:span");
        out.append_attribute("text:style-name") = style.c_str();
      }
    } else if (c == '\t') {
      out.append_child("text:tab");
    } else {
      size_t count = 1;
      while (i + count < n && txt[i + count] == ' ') ++count;
      bool inner = out.first_child() && i + count < n && txt[i + count] != '\n';
      if (inner) {
        out.append_child(pugi::node_pcdata).set_value(" ");
        --count;
      }
      if (count > 0) {
        pugi::xml_node sp = out.append_child("text:s");
        if (count > 1) sp.append_attribute("text:c") = static_cast<uint32_t>(count);
      }
      i += count + (inner ? 1 : 0) - 1;
    }
  }
  if (!chars.empty()) out.append_child(pugi::node_pcdata).set_value(chars.c_str());
  if (!style.empty() && !out.first_child()) p.remove_child(out);
}

// the runs of rich text: the text and the run properties of each run
static std::vector<std::pair<std::string, pugi::xml_node>> ods_rich_runs(pugi::xml_node parent) {
  std::vector<std::pair<std::string, pugi::xml_node>> runs;
  for (pugi::xml_node r : parent.children("r")) {
    std::string text;
    for (pugi::xml_node t : r.children("t")) text += t.text().get();
    runs.push_back(std::make_pair(text, r.child("rPr")));
  }
  return runs;
}

static std::string ods_plain_text(const std::string& xml) {
  pugi::xml_document doc;
  if (!doc.load_string(xml.c_str(), pugi::parse_default | pugi::parse_ws_pcdata_single)) return "";
  std::string out;
  for (pugi::xml_node t : ods_descendants(doc, "t", "rPh")) out += t.text().get();
  return out;
}

// formulas and values in the cell data are stored with xml escapes
static std::string ods_unescape(const std::string& x) {
  if (x.find('&') == std::string::npos) return x;
  pugi::xml_document doc;
  if (!doc.load_string(("<x>" + x + "</x>").c_str(), pugi::parse_default | pugi::parse_ws_pcdata_single)) return x;
  return doc.child("x").text().get();
}

// formulas: xlsx (A1 syntax) <-> OpenFormula
//
// Function names that differ between both. Other functions with the _xlfn.
// prefix use the COM.MICROSOFT. namespace in OpenFormula.
static const char* const ods_functions[][2] = {
  {"_XLFN.ACOT", "ACOT"}, {"_XLFN.ACOTH", "ACOTH"}, {"_XLFN.ARABIC", "ARABIC"}, {"_XLFN.BASE", "BASE"},
  {"_XLFN.BITAND", "BITAND"}, {"_XLFN.BITLSHIFT", "BITLSHIFT"}, {"_XLFN.BITOR", "BITOR"},
  {"_XLFN.BITRSHIFT", "BITRSHIFT"}, {"_XLFN.BITXOR", "BITXOR"}, {"_XLFN.COMBINA", "COMBINA"},
  {"_XLFN.COT", "COT"}, {"_XLFN.COTH", "COTH"}, {"_XLFN.CSC", "CSC"}, {"_XLFN.CSCH", "CSCH"},
  {"_XLFN.DAYS", "DAYS"}, {"_XLFN.DECIMAL", "DECIMAL"}, {"_XLFN.GAMMA", "GAMMA"}, {"_XLFN.GAUSS", "GAUSS"},
  {"_XLFN.IFNA", "IFNA"}, {"_XLFN.ISFORMULA", "ISFORMULA"}, {"_XLFN.ISOWEEKNUM", "ISOWEEKNUM"},
  {"_XLFN.MUNIT", "MUNIT"}, {"_XLFN.NUMBERVALUE", "NUMBERVALUE"}, {"_XLFN.PDURATION", "PDURATION"},
  {"_XLFN.PERMUTATIONA", "PERMUTATIONA"}, {"_XLFN.PHI", "PHI"}, {"_XLFN.RRI", "RRI"}, {"_XLFN.SEC", "SEC"},
  {"_XLFN.SECH", "SECH"}, {"_XLFN.SHEET", "SHEET"}, {"_XLFN.SHEETS", "SHEETS"}, {"_XLFN.UNICHAR", "UNICHAR"},
  {"_XLFN.UNICODE", "UNICODE"}, {"_XLFN.XOR", "XOR"},
  {"_XLFN.F.DIST", "FDIST"}, {"_XLFN.F.INV", "FINV"}, {"_XLFN.FORMULATEXT", "FORMULA"}, {"_XLFN.SKEW.P", "SKEWP"},
  {"CEILING", "COM.MICROSOFT.CEILING"}, {"FLOOR", "COM.MICROSOFT.FLOOR"},
  {"CHIDIST", "LEGACY.CHIDIST"}, {"CHIINV", "LEGACY.CHIINV"}, {"CHITEST", "LEGACY.CHITEST"},
  {"FDIST", "LEGACY.FDIST"}, {"FINV", "LEGACY.FINV"}, {"NORMSDIST", "LEGACY.NORMSDIST"},
  {"NORMSINV", "LEGACY.NORMSINV"}, {"TDIST", "LEGACY.TDIST"}
};

static std::string ods_upper(std::string x) {
  for (char& c : x) c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
  return x;
}

static bool ods_ident_char(char c) {
  unsigned char u = static_cast<unsigned char>(c);
  return std::isalnum(u) || c == '_' || c == '.' || c == '\\' || u >= 0x80;
}

struct ods_ref_part {
  bool col_abs = false;
  bool row_abs = false;
  int64_t col = 0;
  int64_t row = 0;
};

// $A$1, A1, $A or $1 at pos. Returns the end of the part, or pos if there is none.
static size_t ods_parse_ref_part(const std::string& s, size_t pos, ods_ref_part& part) {
  part = ods_ref_part();
  size_t n = s.size(), i = pos;
  bool abs1 = i < n && s[i] == '$';
  if (abs1) ++i;
  size_t letters = i;
  int64_t col = 0;
  while (i < n && i - letters < 3 && std::isalpha(static_cast<unsigned char>(s[i]))) {
    col = col * 26 + (std::toupper(static_cast<unsigned char>(s[i])) - 'A' + 1);
    ++i;
  }
  bool has_col = i > letters;
  bool abs2 = has_col && i < n && s[i] == '$';
  if (abs2) ++i;
  size_t digits = i;
  int64_t row = 0;
  while (i < n && i - digits < 8 && std::isdigit(static_cast<unsigned char>(s[i]))) {
    row = row * 10 + (s[i] - '0');
    ++i;
  }
  bool has_row = i > digits;
  if (!has_row && (!has_col || abs2)) return pos;
  if (i < n && (ods_ident_char(s[i]) || s[i] == '(')) return pos;
  if (has_col && col > MAX_OOXML_COL_INT) return pos;
  if (has_row && (row < 1 || row > MAX_OOXML_ROW_INT)) return pos;
  if (has_col) {
    part.col = col;
    part.col_abs = abs1;
    part.row_abs = abs2;
  } else {
    part.row_abs = abs1;
  }
  if (has_row) part.row = row;
  return i;
}

static bool ods_write_ref_part(const ods_ref_part& part, int64_t drow, int64_t dcol, std::string& out) {
  if (part.col > 0) {
    int64_t col = part.col + (part.col_abs ? 0 : dcol);
    if (col < 1 || col > MAX_OOXML_COL_INT) return false;
    if (part.col_abs) out += '$';
    out += int_to_col(col);
  }
  if (part.row > 0) {
    int64_t row = part.row + (part.row_abs ? 0 : drow);
    if (row < 1 || row > MAX_OOXML_ROW_INT) return false;
    if (part.row_abs) out += '$';
    out += std::to_string(row);
  }
  return true;
}

// sheet names are quoted unless they are plain words, the quotes inside the
// name are already doubled
static std::string ods_sheet_name(const std::string& name) {
  bool plain = !name.empty() && !std::isdigit(static_cast<unsigned char>(name[0]));
  for (char c : name) {
    if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') plain = false;
  }
  ods_ref_part part;
  if (plain && ods_parse_ref_part(name, 0, part) == name.size()) plain = false;
  return plain ? name : "'" + name + "'";
}

// Sheet!, 'Sheet 1'!, Sheet1:Sheet3! or 'Sheet 1:Sheet 3'! at pos. Returns the
// position after the "!", or pos if there is none.
static size_t ods_parse_sheet(const std::string& s, size_t pos, std::string& first, std::string& last) {
  size_t n = s.size(), i = pos;
  first.clear();
  last.clear();
  if (i < n && s[i] == '\'') {
    std::string name;
    ++i;
    while (i < n) {
      if (s[i] == '\'') {
        if (i + 1 < n && s[i + 1] == '\'') {
          name += "''";
          i += 2;
          continue;
        }
        break;
      }
      name += s[i++];
    }
    if (i + 1 >= n || s[i + 1] != '!') return pos;
    size_t colon = name.find(':');
    first = name.substr(0, colon);
    if (colon != std::string::npos) last = name.substr(colon + 1);
    return i + 2;
  }
  while (i < n && ods_ident_char(s[i])) ++i;
  if (i == pos || i >= n) return pos;
  if (s[i] == '!') {
    first = s.substr(pos, i - pos);
    return i + 1;
  }
  if (s[i] != ':') return pos;
  size_t j = i + 1;
  while (j < n && ods_ident_char(s[j])) ++j;
  if (j == i + 1 || j >= n || s[j] != '!') return pos;
  first = s.substr(pos, i - pos);
  last = s.substr(i + 1, j - i - 1);
  return j + 1;
}

struct ods_table {
  std::string name, sheet;
  int64_t r1 = 0, c1 = 0, r2 = 0, c2 = 0, headers = 1, totals = 0;
  bool filter = false;
  std::vector<std::string> columns;
};

// the cell a formula is written to, used by structured references
struct ods_formula_ctx {
  const std::vector<ods_table>* tables = nullptr;
  int64_t row = 0;
  int64_t col = 0;
  std::string sheet;
};

// the end of the bracket at pos, brackets inside are nested and ' escapes the next character
static size_t ods_bracket_end(const std::string& f, size_t pos) {
  int32_t depth = 0;
  size_t i = pos;
  for (; i < f.size(); ++i) {
    if (f[i] == '\'' && i + 1 < f.size()) {
      ++i;
      continue;
    }
    if (f[i] == '[') ++depth;
    if (f[i] == ']' && --depth == 0) break;
  }
  return i;
}

// The structured reference in the brackets at pos, e.g. [[#This Row],[a]:[b]],
// as OpenFormula reference to the cells of the table.
static std::string ods_structured_ref(const std::string& content, const ods_table& table, const ods_formula_ctx& ctx) {
  std::string spec = content;
  bool this_row = false;
  if (!spec.empty() && spec[0] == '@') {
    this_row = true;
    spec = spec.substr(1);
  }

  std::vector<std::string> items;
  bool col_range = false;
  if (!spec.empty() && spec[0] == '[') {
    size_t i = 0;
    while (i < spec.size()) {
      if (spec[i] == '[') {
        size_t end = ods_bracket_end(spec, i);
        items.push_back(spec.substr(i + 1, end - i - 1));
        i = end + 1;
      } else {
        if (spec[i] == ':') col_range = true;
        ++i;
      }
    }
  } else if (!spec.empty()) {
    items.push_back(spec);
  }

  bool all = false, data = false, headers = false, totals = false;
  std::vector<int64_t> cols;
  for (const std::string& raw : items) {
    std::string item;
    for (size_t i = 0; i < raw.size(); ++i) {
      if (raw[i] == '\'' && i + 1 < raw.size()) ++i;
      item += raw[i];
    }
    std::string lower = ods_lower(item);
    if (lower == "#all") all = true;
    else if (lower == "#data") data = true;
    else if (lower == "#headers") headers = true;
    else if (lower == "#totals") totals = true;
    else if (lower == "#this row") this_row = true;
    else {
      size_t k = 0;
      while (k < table.columns.size() && ods_lower(table.columns[k]) != lower) ++k;
      if (k == table.columns.size()) return "#REF!";
      cols.push_back(table.c1 + static_cast<int64_t>(k));
    }
  }

  int64_t data_r1 = table.r1 + table.headers, data_r2 = table.r2 - table.totals;
  int64_t row1 = MAX_OOXML_ROW_INT + 1, row2 = 0;
  if (!all && !headers && !totals && !this_row) data = true;
  if (all) {
    row1 = table.r1;
    row2 = table.r2;
  }
  // a missing header or totals row is only an error if nothing else is selected
  if (headers && table.headers > 0) {
    row1 = std::min(row1, table.r1);
    row2 = std::max(row2, table.r1 + table.headers - 1);
  }
  if (data) {
    row1 = std::min(row1, data_r1);
    row2 = std::max(row2, data_r2);
  }
  if (totals && table.totals > 0) {
    row1 = std::min(row1, data_r2 + 1);
    row2 = std::max(row2, table.r2);
  }
  if (this_row) {
    if (ctx.row < table.r1 || ctx.row > table.r2) return "#REF!";
    row1 = row2 = ctx.row;
  }
  if (row1 > row2) return "#REF!";

  int64_t col1 = table.c1, col2 = table.c2;
  if (!cols.empty()) {
    col1 = col_range ? std::min(cols.front(), cols.back()) : cols.front();
    col2 = col_range ? std::max(cols.front(), cols.back()) : cols.front();
  }

  std::string row1_s = this_row ? std::to_string(row1) : "$" + std::to_string(row1);
  std::string row2_s = this_row ? std::to_string(row2) : "$" + std::to_string(row2);
  std::string out = "[$" + ods_sheet_name(table.sheet) + ".$" + int_to_col(col1) + row1_s;
  if (col1 != col2 || row1 != row2) out += ":.$" + int_to_col(col2) + row2_s;
  return out + "]";
}

// an xlsx reference at pos as OpenFormula reference. Returns the end of the
// reference, or pos if there is none. Relative parts are moved by drow and dcol.
static size_t ods_xlsx_ref(const std::string& f, size_t pos, int64_t drow, int64_t dcol, std::string& out) {
  std::string sheet1, sheet2;
  size_t j = ods_parse_sheet(f, pos, sheet1, sheet2);
  if (j > pos && f.compare(j, 5, "#REF!") == 0) {
    out = "#REF!";
    return j + 5;
  }
  ods_ref_part p1, p2;
  size_t k = ods_parse_ref_part(f, j, p1);
  if (k == j) return pos;
  // references to other files
  bool external = !sheet1.empty() && sheet1[0] == '[';
  size_t end = k;
  bool range = false;
  if (k < f.size() && f[k] == ':') {
    size_t m = ods_parse_ref_part(f, k + 1, p2);
    if (m > k + 1 && (p1.col > 0) == (p2.col > 0) && (p1.row > 0) == (p2.row > 0)) {
      range = true;
      end = m;
    }
  }
  if ((p1.col == 0 || p1.row == 0) && !range) return pos;

  std::string ref = "[";
  if (!sheet1.empty()) ref += "$" + ods_sheet_name(sheet1);
  ref += ".";
  bool ok = ods_write_ref_part(p1, drow, dcol, ref);
  if (range || !sheet2.empty()) {
    ref += ":";
    if (!sheet2.empty()) ref += "$" + ods_sheet_name(sheet2);
    ref += ".";
    ok = ods_write_ref_part(range ? p2 : p1, drow, dcol, ref) && ok;
  }
  out = ok && !external ? ref + "]" : "#REF!";
  return end;
}

static std::string ods_formula_to_of(const std::string& f, int64_t drow, int64_t dcol, const ods_formula_ctx& ctx) {
  std::string out = "of:=";
  std::vector<bool> function_parens;
  bool in_array = false, last_ref = false;
  size_t n = f.size(), i = std::min(f.find_first_not_of('='), n);
  while (i < n) {
    char c = f[i];
    std::string ref;

    if (c == '\'' || c == '$' || ods_ident_char(c)) {
      size_t end = ods_xlsx_ref(f, i, drow, dcol, ref);
      if (end > i) {
        out += ref;
        i = end;
        last_ref = true;
        continue;
      }
    }

    if (c == ' ') {
      size_t j = i;
      while (j < n && f[j] == ' ') ++j;
      // a space between two references is the intersection
      if (last_ref && j < n && ods_xlsx_ref(f, j, drow, dcol, ref) > j) {
        out += "!";
      } else {
        out += f.substr(i, j - i);
      }
      i = j;
      continue;
    }

    last_ref = false;
    if (c == '"') {
      size_t j = i + 1;
      while (j < n) {
        if (f[j] == '"') {
          if (j + 1 < n && f[j + 1] == '"') {
            j += 2;
            continue;
          }
          break;
        }
        ++j;
      }
      out += f.substr(i, j - i + 1);
      i = j + 1;
    } else if (c == '#') {
      size_t j = i;
      if (f.compare(i, 4, "#N/A") == 0) {
        j = i + 4;
      } else {
        while (j < n && f[j] != '!' && f[j] != '?') ++j;
        j = std::min(j + 1, n);
      }
      out += f.substr(i, j - i);
      i = j;
    } else if (c == '[') {
      size_t end = ods_bracket_end(f, i);
      std::string content = f.substr(i + 1, end - i - 1);
      i = std::min(end + 1, n);
      // [@Col] or [Col] inside a table
      const ods_table* table = nullptr;
      if (ctx.tables) {
        for (const ods_table& t : *ctx.tables) {
          if (t.sheet == ctx.sheet && ctx.row >= t.r1 && ctx.row <= t.r2 && ctx.col >= t.c1 && ctx.col <= t.c2) table = &t;
        }
      }
      // [1]Sheet1!A1 or [file.xlsx]Sheet1!A1: references to other files are not translated
      size_t j = ods_xlsx_ref(f, i, 0, 0, ref);
      if (j > i || !table) {
        out += "#REF!";
        i = j;
        continue;
      }
      out += ods_structured_ref(content, *table, ctx);
    } else if (c == '{' || c == '}') {
      in_array = c == '{';
      out += c;
      ++i;
    } else if (c == ',') {
      if (in_array || (!function_parens.empty() && function_parens.back())) out += ';';
      else out += '~';
      ++i;
    } else if (c == ';' && in_array) {
      out += '|';
      ++i;
    } else if (c == '(') {
      function_parens.push_back(false);
      out += c;
      ++i;
    } else if (c == ')') {
      if (!function_parens.empty()) function_parens.pop_back();
      out += c;
      ++i;
    } else if (std::isdigit(static_cast<unsigned char>(c)) || (c == '.' && i + 1 < n && std::isdigit(static_cast<unsigned char>(f[i + 1])))) {
      size_t j = i;
      while (j < n && (std::isdigit(static_cast<unsigned char>(f[j])) || f[j] == '.')) ++j;
      if (j < n && (f[j] == 'E' || f[j] == 'e')) {
        size_t k = j + 1;
        if (k < n && (f[k] == '+' || f[k] == '-')) ++k;
        if (k < n && std::isdigit(static_cast<unsigned char>(f[k]))) {
          j = k;
          while (j < n && std::isdigit(static_cast<unsigned char>(f[j]))) ++j;
        }
      }
      out += f.substr(i, j - i);
      i = j;
    } else if (ods_ident_char(c)) {
      size_t j = i;
      while (j < n && ods_ident_char(f[j])) ++j;
      std::string name = f.substr(i, j - i);
      std::string upper = ods_upper(name);
      const ods_table* table = nullptr;
      if (j < n && f[j] == '[' && ctx.tables) {
        for (const ods_table& t : *ctx.tables) {
          if (ods_upper(t.name) == upper) table = &t;
        }
      }
      if (table) {
        size_t end = ods_bracket_end(f, j);
        out += ods_structured_ref(f.substr(j + 1, end - j - 1), *table, ctx);
        i = std::min(end + 1, n);
        continue;
      }
      if (j < n && f[j] == '(') {
        std::string mapped;
        for (const auto& fn : ods_functions) {
          if (upper == fn[0]) mapped = fn[1];
        }
        if (mapped.empty()) mapped = upper.compare(0, 6, "_XLFN.") == 0 ? "COM.MICROSOFT." + name.substr(6) : name;
        out += mapped + "(";
        function_parens.push_back(true);
        j += 1;
      } else if (upper == "TRUE" || upper == "FALSE") {
        out += upper + "()";
      } else {
        out += name;
      }
      i = j;
    } else {
      out += c;
      ++i;
    }
  }
  return out;
}

// the inside of an OpenFormula reference ([...]) as xlsx reference
static std::string ods_of_ref(const std::string& inner, const std::map<std::string, std::string>& renames) {
  if (inner.find('#') != std::string::npos) return "#REF!";

  std::vector<std::string> sheets, refs;
  size_t i = 0, n = inner.size();
  while (i <= n) {
    std::string sheet;
    if (i < n && inner[i] == '$') ++i;
    if (i < n && inner[i] == '\'') {
      ++i;
      while (i < n) {
        if (inner[i] == '\'') {
          if (i + 1 < n && inner[i + 1] == '\'') {
            sheet += "''";
            i += 2;
            continue;
          }
          ++i;
          break;
        }
        sheet += inner[i++];
      }
    } else {
      size_t colon = inner.find(':', i);
      size_t dot = inner.rfind('.', colon == std::string::npos ? n : colon);
      if (dot != std::string::npos && dot >= i) {
        sheet = inner.substr(i, dot - i);
        i = dot;
      }
    }
    if (i < n && inner[i] == '.') ++i;
    size_t colon = inner.find(':', i);
    size_t end = colon == std::string::npos ? n : colon;
    sheets.push_back(sheet);
    refs.push_back(inner.substr(i, end - i));
    i = end + 1;
  }

  // sheets with names that are not valid in xlsx are renamed
  for (std::string& sheet : sheets) {
    if (renames.empty() || sheet.empty()) break;
    std::string plain;
    for (size_t k = 0; k < sheet.size(); ++k) {
      plain += sheet[k];
      if (sheet[k] == '\'' && k + 1 < sheet.size() && sheet[k + 1] == '\'') ++k;
    }
    std::map<std::string, std::string>::const_iterator renamed = renames.find(plain);
    if (renamed == renames.end()) continue;
    sheet.clear();
    for (char ch : renamed->second) {
      sheet += ch;
      if (ch == '\'') sheet += ch;
    }
  }

  std::string out;
  bool three_d = sheets.size() > 1 && !sheets[1].empty() && sheets[1] != sheets[0];
  if (three_d) {
    std::string first = ods_sheet_name(sheets[0]), last = ods_sheet_name(sheets[1]);
    if (first == sheets[0] && last == sheets[1]) out += first + ":" + last + "!";
    else out += "'" + sheets[0] + ":" + sheets[1] + "'!";
  } else if (!sheets[0].empty()) {
    out += ods_sheet_name(sheets[0]) + "!";
  }
  out += refs[0];
  if (refs.size() > 1 && !(three_d && refs[1] == refs[0])) out += ":" + refs[1];
  return out;
}

static std::string ods_formula_to_xlsx(const std::string& f, const std::map<std::string, std::string>& renames) {
  std::string out;
  size_t n = f.size(), i = 0;
  if (i < n && f[i] == '=') ++i;
  bool in_array = false;
  while (i < n) {
    char c = f[i];
    if (c == '"') {
      size_t j = i + 1;
      while (j < n) {
        if (f[j] == '"') {
          if (j + 1 < n && f[j + 1] == '"') {
            j += 2;
            continue;
          }
          break;
        }
        ++j;
      }
      out += f.substr(i, j - i + 1);
      i = j + 1;
    } else if (c == '[') {
      size_t j = i + 1;
      bool quoted = false;
      while (j < n && (quoted || f[j] != ']')) {
        if (f[j] == '\'') quoted = !quoted;
        ++j;
      }
      out += ods_of_ref(f.substr(i + 1, j - i - 1), renames);
      i = j + 1;
    } else if (c == '{' || c == '}') {
      in_array = c == '{';
      out += c;
      ++i;
    } else if (c == ';' || c == '~') {
      out += ',';
      ++i;
    } else if (c == '|' && in_array) {
      out += ';';
      ++i;
    } else if (c == '!') {
      out += ' ';
      ++i;
    } else if (c == '#') {
      size_t j = i;
      if (f.compare(i, 4, "#N/A") == 0) {
        j = i + 4;
      } else {
        while (j < n && f[j] != '!' && f[j] != '?') ++j;
        j = std::min(j + 1, n);
      }
      out += f.substr(i, j - i);
      i = j;
    } else if (std::isdigit(static_cast<unsigned char>(c))) {
      size_t j = i;
      while (j < n && (std::isalnum(static_cast<unsigned char>(f[j])) || f[j] == '.' ||
                       ((f[j] == '+' || f[j] == '-') && (f[j - 1] == 'E' || f[j - 1] == 'e')))) ++j;
      out += f.substr(i, j - i);
      i = j;
    } else if (ods_ident_char(c)) {
      size_t j = i;
      while (j < n && ods_ident_char(f[j])) ++j;
      std::string name = f.substr(i, j - i);
      std::string upper = ods_upper(name);
      if (j < n && f[j] == '(') {
        if ((upper == "TRUE" || upper == "FALSE") && j + 1 < n && f[j + 1] == ')') {
          out += upper;
          i = j + 2;
          continue;
        }
        std::string mapped;
        for (const auto& fn : ods_functions) {
          if (upper == fn[1]) mapped = fn[0][0] == '_' ? "_xlfn." + std::string(fn[0] + 6) : fn[0];
        }
        if (mapped.empty()) mapped = upper.compare(0, 14, "COM.MICROSOFT.") == 0 ? "_xlfn." + name.substr(14) : name;
        out += mapped + "(";
        i = j + 1;
      } else {
        out += name;
        i = j;
      }
    } else {
      out += c;
      ++i;
    }
  }
  return out;
}

// drawings, charts, comments, page setup, view settings and document properties

static const double ods_emu_per_pt = 12700.0;

// The size in points of a PNG, JPEG, GIF or BMP image from its header, at the
// resolution stored in the image or at 96 dpi
static bool ods_image_size(const std::string& file, double& width_pt, double& height_pt) {
  std::ifstream in(file.c_str(), std::ios::binary);
  if (!in) return false;
  std::vector<unsigned char> buf(65536);
  in.read(reinterpret_cast<char*>(buf.data()), static_cast<std::streamsize>(buf.size()));
  size_t n = static_cast<size_t>(in.gcount());
  auto be16 = [&](size_t i) { return i + 1 < n ? static_cast<uint32_t>(buf[i]) << 8 | buf[i + 1] : 0u; };
  auto be32 = [&](size_t i) { return i + 3 < n ? be16(i) << 16 | be16(i + 2) : 0u; };
  auto le16 = [&](size_t i) { return i + 1 < n ? static_cast<uint32_t>(buf[i + 1]) << 8 | buf[i] : 0u; };
  auto le32 = [&](size_t i) { return i + 3 < n ? le16(i + 2) << 16 | le16(i) : 0u; };
  double width = 0.0, height = 0.0, dpi_x = 96.0, dpi_y = 96.0;
  if (n > 24 && buf[0] == 0x89 && buf[1] == 'P' && buf[2] == 'N' && buf[3] == 'G') {
    width = be32(16);
    height = be32(20);
    // the pHYs chunk with pixels per meter
    for (size_t i = 8; i + 8 < n;) {
      uint32_t len = be32(i);
      if (std::memcmp(&buf[i + 4], "pHYs", 4) == 0 && i + 16 < n && buf[i + 16] == 1) {
        dpi_x = be32(i + 8) * 0.0254;
        dpi_y = be32(i + 12) * 0.0254;
        break;
      }
      if (std::memcmp(&buf[i + 4], "IDAT", 4) == 0) break;
      i += static_cast<size_t>(len) + 12;
    }
  } else if (n > 4 && buf[0] == 0xFF && buf[1] == 0xD8) {
    // the JFIF density and the frame size of the start of frame marker
    for (size_t i = 2; i + 9 < n;) {
      if (buf[i] != 0xFF) break;
      uint32_t marker = buf[i + 1], len = be16(i + 2);
      if (marker == 0xE0 && std::memcmp(&buf[i + 4], "JFIF", 4) == 0 && i + 15 < n) {
        uint32_t unit = buf[i + 11];
        double dx = be16(i + 12), dy = be16(i + 14);
        if (unit == 1 && dx > 0 && dy > 0) {
          dpi_x = dx;
          dpi_y = dy;
        } else if (unit == 2 && dx > 0 && dy > 0) {
          dpi_x = dx * 2.54;
          dpi_y = dy * 2.54;
        }
      }
      if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC) {
        height = be16(i + 5);
        width = be16(i + 7);
        break;
      }
      i += static_cast<size_t>(len) + 2;
    }
  } else if (n > 10 && buf[0] == 'G' && buf[1] == 'I' && buf[2] == 'F') {
    width = le16(6);
    height = le16(8);
  } else if (n > 46 && buf[0] == 'B' && buf[1] == 'M') {
    width = le32(18);
    height = static_cast<double>(static_cast<int32_t>(le32(22)));
    if (height < 0) height = -height;
    if (le32(38) > 0 && le32(42) > 0) {
      dpi_x = le32(38) * 0.0254;
      dpi_y = le32(42) * 0.0254;
    }
  }
  if (!(width > 0.0) || !(height > 0.0) || !(dpi_x > 0.0) || !(dpi_y > 0.0)) return false;
  width_pt = width / dpi_x * 72.0;
  height_pt = height / dpi_y * 72.0;
  return true;
}

static std::string ods_cm(double pt) {
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%.4fcm", pt / 72.0 * 2.54);
  return buf;
}

// a chart range in xlsx syntax (Sheet1!$A$1:$A$3) as ODF cell range address (Sheet1.$A$1:Sheet1.$A$3);
// a union of ranges in parentheses is a list of addresses
static std::string ods_chart_range(const std::string& f) {
  std::string ref = f;
  if (ref.size() > 2 && ref.front() == '(' && ref.back() == ')') ref = ref.substr(1, ref.size() - 2);
  std::string out;
  size_t j = 0;
  while (j < ref.size()) {
    std::string sheet1, sheet2;
    size_t start = j;
    j = ods_parse_sheet(ref, start, sheet1, sheet2);
    if (j == start || !sheet2.empty()) return "";
    ods_ref_part p1, p2;
    size_t k = ods_parse_ref_part(ref, j, p1);
    if (k == j) return "";
    std::string sheet = ods_sheet_name(sheet1);
    out += (out.empty() ? "" : " ") + sheet + ".";
    ods_write_ref_part(p1, 0, 0, out);
    if (k < ref.size() && ref[k] == ':') {
      size_t m = ods_parse_ref_part(ref, k + 1, p2);
      if (m == k + 1) return "";
      out += ":" + sheet + ".";
      ods_write_ref_part(p2, 0, 0, out);
      k = m;
    }
    if (k < ref.size() && ref[k] != ',') return "";
    j = k + 1;
  }
  return out;
}

// the preset dashes of lines as dashes and gaps relative to the line width in percent
struct ods_dash {
  const char* name;
  int32_t dots1, length1, dots2, length2, distance;
};
static const ods_dash ods_dashes[] = {
  {"sysDot", 1, 100, 0, 0, 100}, {"sysDash", 1, 300, 0, 0, 100}, {"dash", 1, 400, 0, 0, 300},
  {"dashDot", 1, 400, 1, 100, 300}, {"lgDash", 1, 800, 0, 0, 300}, {"lgDashDot", 1, 800, 1, 100, 300},
  {"lgDashDotDot", 1, 800, 2, 100, 300}, {"sysDashDot", 1, 300, 1, 100, 100}, {"sysDashDotDot", 1, 300, 2, 100, 100},
  {"dot", 1, 100, 0, 0, 300}
};

// the line ends as markers, oval and diamond markers are centered on the end of the line
struct ods_marker {
  const char* type;
  const char* name;
  const char* view;
  const char* path;
  bool center;
};
static const ods_marker ods_markers[] = {
  {"triangle", "msArrowEnd", "0 0 500 500", "M250 0l250 500h-500z", false},
  {"stealth", "msArrowStealthEnd", "0 0 300 300", "M150 0l150 300-150-120-150 120z", false},
  {"diamond", "msArrowDiamondEnd", "0 0 300 300", "M150 0l150 150-150 150-150-150z", true},
  {"oval", "msArrowOvalEnd", "0 0 300 300", "M150 0c83 0 150 67 150 150 0 83-67 150-150 150-83 0-150-67-150-150 0-83 67-150 150-150z", true},
  {"arrow", "msArrowOpenEnd", "0 0 350 350", "M175 0l175 293-57 57-81-142v142h-75v-142l-81 142-56-57z", false}
};

static std::string ods_drawing_text(pugi::xml_node node) {
  std::string out;
  for (pugi::xml_node t : ods_descendants(node, "a:t")) out += t.text().get();
  if (out.empty()) out = ods_descendant(node, "c:v").text().get();
  return out;
}

// the opacity of the color of a fill, 1 without alpha
static double ods_drawing_alpha(pugi::xml_node fill) {
  pugi::xml_node alpha = fill.first_child().child("a:alpha");
  return alpha ? alpha.attribute("val").as_double() / 100000.0 : 1.0;
}

static std::string ods_drawing_color(pugi::xml_node sppr, const std::vector<std::string>& theme) {
  static const char* const names[] = {"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                                      "accent4", "accent5", "accent6", "hlink", "folHlink"};
  pugi::xml_node fill = sppr.child("a:solidFill");
  if (!fill) fill = sppr.child("a:ln").child("a:solidFill");
  pugi::xml_node color = fill.child("a:srgbClr");
  std::string hex = color.attribute("val").as_string();
  if (!color) {
    color = fill.child("a:schemeClr");
    std::string val = color.attribute("val").as_string();
    if (val == "tx1") val = "dk1";
    if (val == "bg1") val = "lt1";
    if (val == "tx2") val = "dk2";
    if (val == "bg2") val = "lt2";
    for (size_t k = 0; k < 12 && k < theme.size(); ++k) {
      if (val == names[k]) hex = theme[k];
    }
  }
  if (hex.size() != 6) return "";

  // shades and tints are applied to the linear components, luminance changes in HSL
  double rgb[3];
  for (size_t k = 0; k < 3; ++k) rgb[k] = static_cast<double>(std::strtol(hex.substr(2 * k, 2).c_str(), nullptr, 16)) / 255.0;
  auto linear = [](double c) { return c <= 0.04045 ? c / 12.92 : std::pow((c + 0.055) / 1.055, 2.4); };
  auto gamma = [](double c) { return c <= 0.0031308 ? c * 12.92 : 1.055 * std::pow(c, 1.0 / 2.4) - 0.055; };
  for (pugi::xml_node mod : color.children()) {
    std::string name = mod.name();
    double v = mod.attribute("val").as_double() / 100000.0;
    if (name == "a:shade" || name == "a:tint") {
      for (double& c : rgb) c = gamma(name == "a:shade" ? linear(c) * v : 1.0 - (1.0 - linear(c)) * v);
    } else if (name == "a:lumMod" || name == "a:lumOff") {
      double mx = std::max(rgb[0], std::max(rgb[1], rgb[2])), mn = std::min(rgb[0], std::min(rgb[1], rgb[2]));
      double l = (mx + mn) / 2.0, d = mx - mn, h = 0.0, sat = 0.0;
      if (d > 0.0) {
        sat = l > 0.5 ? d / (2.0 - mx - mn) : d / (mx + mn);
        if (mx == rgb[0]) h = std::fmod((rgb[1] - rgb[2]) / d + 6.0, 6.0);
        else if (mx == rgb[1]) h = (rgb[2] - rgb[0]) / d + 2.0;
        else h = (rgb[0] - rgb[1]) / d + 4.0;
      }
      l = name == "a:lumMod" ? l * v : l + v;
      l = std::min(1.0, std::max(0.0, l));
      double ch = (1.0 - std::fabs(2.0 * l - 1.0)) * sat, x = ch * (1.0 - std::fabs(std::fmod(h, 2.0) - 1.0)), m = l - ch / 2.0;
      int32_t sector = static_cast<int32_t>(h);
      double r = sector == 0 || sector == 5 ? ch : (sector == 1 || sector == 4 ? x : 0.0);
      double g = sector == 1 || sector == 2 ? ch : (sector == 0 || sector == 3 ? x : 0.0);
      double b = sector == 3 || sector == 4 ? ch : (sector == 2 || sector == 5 ? x : 0.0);
      rgb[0] = r + m;
      rgb[1] = g + m;
      rgb[2] = b + m;
    }
  }
  char buf[8];
  std::snprintf(buf, sizeof(buf), "#%02x%02x%02x", static_cast<uint32_t>(std::lround(std::min(1.0, std::max(0.0, rgb[0])) * 255.0)),
                static_cast<uint32_t>(std::lround(std::min(1.0, std::max(0.0, rgb[1])) * 255.0)),
                static_cast<uint32_t>(std::lround(std::min(1.0, std::max(0.0, rgb[2])) * 255.0)));
  return buf;
}

// the values of a range of cells: numeric or not and the value
typedef std::vector<std::pair<bool, std::string>> ods_values;

// the cached or literal values of a chart range, or the values of its cells
static ods_values ods_chart_values(pugi::xml_node ref, const std::function<ods_values(const std::string&)>& lookup) {
  ods_values values;
  pugi::xml_node cache = ods_descendant(ref, "c:numCache");
  if (!cache) cache = ref.child("c:numLit");
  bool numeric = static_cast<bool>(cache);
  if (!cache) cache = ods_descendant(ref, "c:strCache");
  if (!cache) cache = ref.child("c:strLit");
  int64_t count = cache.child("c:ptCount").attribute("val").as_llong(0);
  if (cache && count > 0 && count <= 65536) {
    values.assign(static_cast<size_t>(count), std::make_pair(numeric, std::string()));
    for (pugi::xml_node pt : cache.children("c:pt")) {
      int64_t idx = pt.attribute("idx").as_llong(-1);
      if (idx >= 0 && idx < count) values[static_cast<size_t>(idx)].second = pt.child_value("c:v");
    }
    return values;
  }
  std::string f = ods_descendant(ref, "c:f").text().get();
  return f.empty() ? values : lookup(f);
}

// Converts a DrawingML chart into the content of an ODF chart object. Bar,
// line, area, pie, doughnut, scatter, radar, bubble and stock plots are
// converted with their series, categories, secondary axes, titles, legend,
// colors, markers, data labels, data point colors, trendlines and axis scaling.
static bool ods_chart_object(const std::string& xml, const std::vector<std::string>& theme, double width_pt, double height_pt,
                             const std::string& path, const std::string& version, const std::function<ods_values(const std::string&)>& lookup,
                             std::string& ranges, const std::set<std::string>& pivot_names) {
  pugi::xml_document src;
  if (!src.load_string(xml.c_str())) return false;
  pugi::xml_node chart = src.child("c:chartSpace").child("c:chart");
  pugi::xml_node plot = chart.child("c:plotArea");
  if (!plot) return false;

  pugi::xml_document doc;
  ods_declaration(doc);
  pugi::xml_node root = doc.append_child("office:document-content");
  root.append_attribute("xmlns:office") = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
  root.append_attribute("xmlns:style") = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
  root.append_attribute("xmlns:chart") = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
  root.append_attribute("xmlns:table") = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
  root.append_attribute("xmlns:text") = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
  root.append_attribute("xmlns:draw") = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
  root.append_attribute("xmlns:svg") = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
  root.append_attribute("xmlns:xlink") = "http://www.w3.org/1999/xlink";
  root.append_attribute("xmlns:fo") = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
  root.append_attribute("xmlns:number") = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
  root.append_attribute("office:version") = version.c_str();
  pugi::xml_node autos = root.append_child("office:automatic-styles");
  pugi::xml_node node = root.append_child("office:body").append_child("office:chart").append_child("chart:chart");
  node.append_attribute("svg:width") = ods_cm(width_pt).c_str();
  node.append_attribute("svg:height") = ods_cm(height_pt).c_str();

  // a new automatic chart style
  int32_t style_count = 0;
  auto new_style = [&]() {
    pugi::xml_node style = autos.append_child("style:style");
    style.append_attribute("style:name") = ("ch" + std::to_string(++style_count)).c_str();
    style.append_attribute("style:family") = "chart";
    return style;
  };
  auto set_attr = [](pugi::xml_node target, const char* name, const std::string& value) {
    pugi::xml_attribute attr = target.attribute(name);
    if (!attr) attr = target.append_attribute(name);
    attr = value.c_str();
  };
  auto percent_of = [](double x) { return ods_num(std::round(x * 1000.0) / 10.0) + "%"; };
  auto pt_length = [](double x) { return ods_num(std::round(x * 1e6) / 1e6) + "pt"; };

  // the fill and line of an element of the chart; the dashes are written to the styles of the object
  std::set<std::string> dashes;
  auto graphic_props = [&](pugi::xml_node sppr, pugi::xml_node style) {
    if (!sppr) return;
    // the graphic properties follow the chart properties
    pugi::xml_node gp = style.child("style:graphic-properties");
    if (!gp && style.child("style:chart-properties")) gp = style.insert_child_after("style:graphic-properties", style.child("style:chart-properties"));
    if (!gp) gp = style.prepend_child("style:graphic-properties");
    if (sppr.child("a:noFill")) {
      set_attr(gp, "draw:fill", "none");
    } else if (pugi::xml_node fill = sppr.child("a:solidFill")) {
      std::string color = ods_drawing_color(sppr, theme);
      if (!color.empty()) {
        set_attr(gp, "draw:fill", "solid");
        set_attr(gp, "draw:fill-color", color);
        double alpha = ods_drawing_alpha(fill);
        if (alpha < 1.0) set_attr(gp, "draw:opacity", percent_of(alpha));
      }
    }
    pugi::xml_node ln = sppr.child("a:ln");
    if (ln.child("a:noFill")) {
      set_attr(gp, "draw:stroke", "none");
    } else if (ln) {
      const ods_dash* dash = nullptr;
      for (const ods_dash& d : ods_dashes) {
        if (std::strcmp(ln.child("a:prstDash").attribute("val").as_string(), d.name) == 0) dash = &d;
      }
      set_attr(gp, "draw:stroke", dash ? "dash" : "solid");
      if (dash) {
        set_attr(gp, "draw:stroke-dash", dash->name);
        dashes.insert(dash->name);
      }
      std::string color = ods_drawing_color(ln, theme);
      if (!color.empty()) set_attr(gp, "svg:stroke-color", color);
      if (ln.attribute("w")) set_attr(gp, "svg:stroke-width", pt_length(ln.attribute("w").as_double() / ods_emu_per_pt));
      double alpha = ods_drawing_alpha(ln.child("a:solidFill"));
      if (alpha < 1.0) set_attr(gp, "svg:stroke-opacity", percent_of(alpha));
    }
    if (!gp.first_attribute()) style.remove_child(gp);
  };

  // the text properties of an element: those of the chart, then of the element, then of its first run
  pugi::xml_node chart_text = ods_descendant(src.child("c:chartSpace").child("c:txPr"), "a:defRPr");
  auto text_props = [&](pugi::xml_node element, pugi::xml_node style) {
    pugi::xml_node rich = element.child("c:tx").child("c:rich");
    pugi::xml_node layers[] = {
      chart_text, ods_descendant(element.child("c:txPr"), "a:defRPr"), ods_descendant(rich, "a:defRPr"), ods_descendant(rich, "a:rPr")
    };
    pugi::xml_document rpr_doc;
    pugi::xml_node rpr = rpr_doc.append_child("rPr");
    for (pugi::xml_node layer : layers) {
      for (pugi::xml_attribute attr : layer.attributes()) set_attr(rpr, attr.name(), attr.value());
      for (const char* name : {"a:solidFill", "a:latin"}) {
        if (!layer.child(name)) continue;
        rpr.remove_child(name);
        rpr.append_copy(layer.child(name));
      }
    }
    pugi::xml_node tp = style.append_child("style:text-properties");
    if (rpr.attribute("sz")) set_attr(tp, "fo:font-size", ods_num(rpr.attribute("sz").as_double() / 100.0) + "pt");
    if (rpr.attribute("b")) set_attr(tp, "fo:font-weight", rpr.attribute("b").as_bool() ? "bold" : "normal");
    if (rpr.attribute("i")) set_attr(tp, "fo:font-style", rpr.attribute("i").as_bool() ? "italic" : "normal");
    std::string color = ods_drawing_color(rpr, theme);
    if (!color.empty()) set_attr(tp, "fo:color", color);
    std::string typeface = rpr.child("a:latin").attribute("typeface").as_string();
    if (!typeface.empty() && typeface[0] != '+') set_attr(tp, "fo:font-family", typeface);
    if (!tp.first_attribute()) style.remove_child(tp);
  };

  // a number format that is not that of the source as data style of the element
  int32_t number_styles = 0;
  auto number_format = [&](pugi::xml_node num_fmt, pugi::xml_node style, pugi::xml_node props) {
    std::string code = num_fmt.attribute("formatCode").as_string();
    if (!num_fmt || num_fmt.attribute("sourceLinked").as_bool(true) || code.empty() || code == "General") return;
    ods_number_format(autos, static_cast<uint32_t>(++number_styles), code);
    style.append_attribute("style:data-style-name") = ("N" + std::to_string(number_styles)).c_str();
    set_attr(props, "chart:link-data-style-to-source", "false");
  };

  struct ods_series {
    std::string cls, values, label, domain, domain2, color;
    pugi::xml_node ser, group;
    bool secondary = false;
    // the cached values of the series, its domains and its label for the local table
    ods_values cache, domain_cache, domain2_cache, label_cache;
  };
  ods_values category_cache;
  std::vector<ods_series> series;
  std::string main_class, categories;
  bool vertical = false, stacked = false, percent = false, three_d = false, smooth = false;
  pugi::xml_node bar_group, stock_group;
  std::vector<std::string> primary_ids;
  std::string secondary_val;

  // the data of a pivot chart are the points of its cache, the ranges of its pivot table also hold the totals
  bool pivot_chart = src.child("c:chartSpace").child("c:pivotSource");
  auto data_range = [&](pugi::xml_node ref) {
    std::string f = ods_descendant(ref, "c:f").text().get();
    pugi::xml_node count = ods_descendant(ref, "c:ptCount");
    size_t bang = f.rfind('!');
    size_t colon = f.find(':', bang == std::string::npos ? 0 : bang);
    if (!pivot_chart || !count || bang == std::string::npos || colon == std::string::npos || f.find(',') != std::string::npos) return ods_chart_range(f);
    std::string first, last;
    for (char ch : f.substr(bang + 1, colon - bang - 1)) {
      if (ch != '$') first += ch;
    }
    for (char ch : f.substr(colon + 1)) {
      if (ch != '$') last += ch;
    }
    int32_t r1 = cell_to_rowint(first), c1 = cell_to_colint(first), r2 = cell_to_rowint(last), c2 = cell_to_colint(last);
    int32_t n = count.attribute("val").as_int();
    if (n <= 0) return ods_chart_range(f);
    if (c1 == c2 && r2 - r1 + 1 > n) r2 = r1 + n - 1;
    else if (r1 == r2 && c2 - c1 + 1 > n) c2 = c1 + n - 1;
    return ods_chart_range(f.substr(0, bang + 1) + "$" + int_to_col(static_cast<uint32_t>(c1)) + "$" + std::to_string(r1) + ":$" +
                           int_to_col(static_cast<uint32_t>(c2)) + "$" + std::to_string(r2));
  };

  for (pugi::xml_node p : plot.children()) {
    std::string name = p.name();
    std::string cls;
    if (name == "c:barChart" || name == "c:bar3DChart") cls = "chart:bar";
    else if (name == "c:lineChart" || name == "c:line3DChart") cls = "chart:line";
    else if (name == "c:areaChart" || name == "c:area3DChart") cls = "chart:area";
    else if (name == "c:pieChart" || name == "c:pie3DChart" || name == "c:ofPieChart") cls = "chart:circle";
    else if (name == "c:doughnutChart") cls = "chart:ring";
    else if (name == "c:scatterChart") cls = "chart:scatter";
    else if (name == "c:bubbleChart") cls = "chart:bubble";
    else if (name == "c:stockChart") cls = "chart:stock";
    else if (name == "c:radarChart") cls = std::string(p.child("c:radarStyle").attribute("val").as_string()) == "filled" ? "chart:filled-radar" : "chart:radar";
    if (cls.empty()) continue;
    if (name.find("3D") != std::string::npos) three_d = true;
    if (cls == "chart:bar" && !bar_group) bar_group = p;
    if (cls == "chart:stock") stock_group = p;

    // the plot types on other axes than the first plot type are on the secondary axes
    std::vector<std::string> ids;
    for (pugi::xml_node id : p.children("c:axId")) ids.push_back(id.attribute("val").as_string());
    bool secondary = false;
    if (main_class.empty()) {
      main_class = cls;
      primary_ids = ids;
      vertical = std::string(p.child("c:barDir").attribute("val").as_string()) == "bar";
      std::string grouping = p.child("c:grouping").attribute("val").as_string();
      stacked = grouping == "stacked" || grouping == "percentStacked";
      percent = grouping == "percentStacked";
    } else if (!ids.empty() && ids != primary_ids) {
      secondary = true;
      for (const std::string& id : ids) {
        for (pugi::xml_node ax : plot.children("c:valAx")) {
          if (id == ax.child("c:axId").attribute("val").as_string()) secondary_val = id;
        }
      }
    }
    if (p.child("c:smooth").attribute("val").as_bool()) smooth = true;

    for (pugi::xml_node ser : p.children("c:ser")) {
      ods_series s;
      s.cls = cls;
      s.ser = ser;
      s.group = p;
      s.secondary = secondary;
      bool xy = cls == "chart:scatter" || cls == "chart:bubble";
      s.values = data_range(ser.child(xy ? "c:yVal" : "c:val"));
      if (xy) s.domain2 = data_range(ser.child("c:xVal"));
      // bubble series: the sizes are the values, the y values the first and the x values the second domain
      if (cls == "chart:bubble") {
        s.domain = s.values;
        s.values = data_range(ser.child("c:bubbleSize"));
        s.domain_cache = ods_chart_values(ser.child("c:yVal"), lookup);
        s.domain2_cache = ods_chart_values(ser.child("c:xVal"), lookup);
      } else if (cls == "chart:scatter") {
        s.domain = s.domain2;
        s.domain2.clear();
        s.domain_cache = ods_chart_values(ser.child("c:xVal"), lookup);
      }
      s.label = ods_chart_range(ods_descendant(ser.child("c:tx"), "c:f").text().get());
      if (categories.empty() && category_cache.empty() && !xy) {
        categories = data_range(ser.child("c:cat"));
        category_cache = ods_chart_values(ser.child("c:cat"), lookup);
      }
      s.cache = ods_chart_values(ser.child(cls == "chart:bubble" ? "c:bubbleSize" : xy ? "c:yVal" : "c:val"), lookup);
      s.label_cache = ods_chart_values(ser.child("c:tx"), lookup);
      if (s.label_cache.empty() && ser.child("c:tx").child("c:v")) s.label_cache.push_back(std::make_pair(false, std::string(ser.child("c:tx").child_value("c:v"))));
      s.color = ods_drawing_color(ser.child("c:spPr"), theme);
      if (ser.child("c:smooth").attribute("val").as_bool()) smooth = true;
      if (!s.values.empty() || !s.cache.empty()) series.push_back(s);
    }
  }
  if (main_class.empty() || series.empty()) return false;

  // stock series are open, high, low and close in xlsx, but open, low, high and close in ODF
  if (main_class == "chart:stock" && series.size() >= 3) {
    size_t high = series.size() == 4 ? 1 : 0;
    std::swap(series[high], series[high + 1]);
  }

  // the local table holds the cached values: the categories in the first column, then the domains and the
  // values of each series. Series without ranges of cells refer to it
  std::vector<std::pair<std::pair<bool, std::string>, ods_values>> local_columns;
  size_t local_rows = category_cache.size();
  for (const ods_series& sr : series) {
    local_rows = std::max(local_rows, std::max(sr.cache.size(), std::max(sr.domain_cache.size(), sr.domain2_cache.size())));
  }
  auto local_column = [&](const std::pair<bool, std::string>& header, const ods_values& values) {
    local_columns.push_back(std::make_pair(header, values));
    std::string col = "$" + int_to_col(static_cast<uint32_t>(local_columns.size()));
    return "local-table." + col + "$2:local-table." + col + "$" + std::to_string(std::max<size_t>(local_rows, 1) + 1);
  };
  std::pair<bool, std::string> no_header(false, std::string());
  std::string local_categories = local_column(no_header, category_cache);
  if (categories.empty() && !category_cache.empty()) categories = local_categories;
  for (ods_series& sr : series) {
    if (!sr.domain2_cache.empty()) {
      std::string address = local_column(no_header, sr.domain2_cache);
      if (sr.domain2.empty()) sr.domain2 = address;
    }
    if (!sr.domain_cache.empty()) {
      std::string address = local_column(no_header, sr.domain_cache);
      if (sr.domain.empty()) sr.domain = address;
    }
    std::string address = local_column(sr.label_cache.empty() ? no_header : sr.label_cache[0], sr.cache);
    if (sr.values.empty()) sr.values = address;
    if (sr.label.empty() && !sr.label_cache.empty()) sr.label = "local-table.$" + int_to_col(static_cast<uint32_t>(local_columns.size())) + "$1";
  }

  node.append_attribute("chart:class") = main_class.c_str();
  node.append_attribute("xlink:href") = "..";
  node.append_attribute("xlink:type") = "simple";
  // the pivot table of a pivot chart, the name of its source follows the sheet
  if (pivot_chart) {
    std::string source = src.child("c:chartSpace").child("c:pivotSource").child_value("c:name");
    std::string pivot = source.substr(source.rfind('!') == std::string::npos ? 0 : source.rfind('!') + 1);
    if (pivot_names.count(pivot)) {
      root.append_attribute("xmlns:loext") = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
      node.append_attribute("loext:data-pilot-source") = pivot.c_str();
    }
  }
  // the fill and border of the chart
  // the fill and border of elements without their own are those of the default chart style: a white chart with a
  // light gray border, the legend and plot area without fill and border
  auto default_props = [&](pugi::xml_node style, const char* fill, const char* stroke) {
    pugi::xml_node gp = style.child("style:graphic-properties");
    if (!gp && style.child("style:chart-properties")) gp = style.insert_child_after("style:graphic-properties", style.child("style:chart-properties"));
    if (!gp) gp = style.prepend_child("style:graphic-properties");
    if (fill) set_attr(gp, "draw:fill", fill[0] == '#' ? "solid" : fill);
    if (fill && fill[0] == '#') set_attr(gp, "draw:fill-color", fill);
    if (stroke) set_attr(gp, "draw:stroke", stroke[0] == '#' ? "solid" : stroke);
    if (stroke && stroke[0] == '#') {
      set_attr(gp, "svg:stroke-color", stroke);
      set_attr(gp, "svg:stroke-width", "0.75pt");
    }
  };
  pugi::xml_node chart_style = new_style();
  if (src.child("c:chartSpace").child("c:spPr")) graphic_props(src.child("c:chartSpace").child("c:spPr"), chart_style);
  else default_props(chart_style, "#ffffff", "#d9d9d9");
  node.append_attribute("chart:style-name") = chart_style.attribute("style:name").value();

  pugi::xml_node title = chart.child("c:title");
  if (title && !chart.child("c:autoTitleDeleted").attribute("val").as_bool()) {
    std::string text = ods_drawing_text(title.child("c:tx"));
    if (!text.empty()) {
      pugi::xml_node title_node = node.append_child("chart:title");
      pugi::xml_node title_style = new_style();
      title_node.append_attribute("chart:style-name") = title_style.attribute("style:name").value();
      text_props(title, title_style);
      graphic_props(title.child("c:spPr"), title_style);
      title_node.append_child("text:p").text().set(text.c_str());
    }
  }
  if (pugi::xml_node legend = chart.child("c:legend")) {
    std::string pos = legend.child("c:legendPos").attribute("val").as_string("r");
    const char* position = "end";
    if (pos == "l") position = "start";
    else if (pos == "t") position = "top";
    else if (pos == "b") position = "bottom";
    else if (pos == "tr") position = "top-end";
    pugi::xml_node legend_node = node.append_child("chart:legend");
    legend_node.append_attribute("chart:legend-position") = position;
    pugi::xml_node legend_style = new_style();
    legend_node.append_attribute("chart:style-name") = legend_style.attribute("style:name").value();
    text_props(legend, legend_style);
    if (legend.child("c:spPr")) graphic_props(legend.child("c:spPr"), legend_style);
    else default_props(legend_style, "none", "none");
  }

  pugi::xml_node area = node.append_child("chart:plot-area");
  pugi::xml_node pa = new_style();
  area.append_attribute("chart:style-name") = pa.attribute("style:name").value();
  pugi::xml_node pa_props = pa.append_child("style:chart-properties");
  pa_props.append_attribute("chart:vertical") = vertical ? "true" : "false";
  if (stacked) pa_props.append_attribute("chart:stacked") = "true";
  if (percent) pa_props.append_attribute("chart:percentage") = "true";
  if (three_d) pa_props.append_attribute("chart:three-dimensional") = "true";
  if (smooth) pa_props.append_attribute("chart:interpolation") = "cubic-spline";
  if (main_class == "chart:stock" && series.size() == 4) pa_props.append_attribute("chart:japanese-candle-stick") = "true";
  // empty cells and the angle of the first slice, counterclockwise from three o'clock
  std::string blanks = chart.child("c:dispBlanksAs").attribute("val").as_string();
  if (!blanks.empty()) pa_props.append_attribute("chart:treat-empty-cells") = blanks == "zero" ? "use-zero" : blanks == "span" ? "ignore" : "leave-gap";
  for (const ods_series& sr : series) {
    if (pugi::xml_node first = sr.group.child("c:firstSliceAng")) {
      int32_t angle = ((90 - first.attribute("val").as_int()) % 360 + 360) % 360;
      if (!pa_props.attribute("chart:angle-offset")) pa_props.append_attribute("chart:angle-offset") = angle;
    }
  }

  // axes: the x axis is the category axis, or for scatter and bubble plots the horizontal value axis
  bool xy_plot = main_class == "chart:scatter" || main_class == "chart:bubble";
  pugi::xml_node x_axis, y_axis, y2_axis;
  for (pugi::xml_node ax : plot.children()) {
    std::string name = ax.name();
    std::string id = ax.child("c:axId").attribute("val").as_string();
    if (std::find(primary_ids.begin(), primary_ids.end(), id) == primary_ids.end()) {
      if (name == "c:valAx" && id == secondary_val) y2_axis = ax;
      continue;
    }
    std::string pos = ax.child("c:axPos").attribute("val").as_string();
    bool horizontal = pos == "b" || pos == "t";
    if (!x_axis && (name == "c:catAx" || name == "c:dateAx" || (xy_plot && name == "c:valAx" && horizontal != vertical))) x_axis = ax;
    else if (!y_axis && name == "c:valAx") y_axis = ax;
  }
  auto add_axis = [&](pugi::xml_node source, const char* dimension, const char* name, bool primary_y) {
    pugi::xml_node axis = area.append_child("chart:axis");
    axis.append_attribute("chart:dimension") = dimension;
    axis.append_attribute("chart:name") = name;
    pugi::xml_node style = new_style();
    axis.append_attribute("chart:style-name") = style.attribute("style:name").value();
    pugi::xml_node props = style.append_child("style:chart-properties");
    props.append_attribute("chart:display-label") = source.child("c:delete").attribute("val").as_bool() ? "false" : "true";
    pugi::xml_node scaling = source.child("c:scaling");
    if (std::string(scaling.child("c:orientation").attribute("val").as_string()) == "maxMin") props.append_attribute("chart:reverse-direction") = "true";
    if (scaling.child("c:logBase")) props.append_attribute("chart:logarithmic") = "true";
    if (pugi::xml_node min = scaling.child("c:min")) props.append_attribute("chart:minimum") = min.attribute("val").as_string();
    if (pugi::xml_node max = scaling.child("c:max")) props.append_attribute("chart:maximum") = max.attribute("val").as_string();
    if (pugi::xml_node unit = source.child("c:majorUnit")) props.append_attribute("chart:interval-major") = unit.attribute("val").as_string();
    double major = source.child("c:majorUnit").attribute("val").as_double(), minor = source.child("c:minorUnit").attribute("val").as_double();
    if (major > 0.0 && minor > 0.0) props.append_attribute("chart:interval-minor-divisor") = static_cast<int32_t>(std::max(std::lround(major / minor), 1L));
    if (std::strcmp(name, "secondary-y") == 0) props.append_attribute("chart:axis-position") = "end";
    else if (pugi::xml_node at = source.child("c:crossesAt")) props.append_attribute("chart:axis-position") = at.attribute("val").as_string();
    else if (std::string(source.child("c:crosses").attribute("val").as_string()) == "max") props.append_attribute("chart:axis-position") = "end";
    else if (std::string(source.child("c:crosses").attribute("val").as_string()) == "min") props.append_attribute("chart:axis-position") = "start";
    if (primary_y && bar_group) {
      if (pugi::xml_node gap = bar_group.child("c:gapWidth")) props.append_attribute("chart:gap-width") = gap.attribute("val").as_string();
      if (pugi::xml_node overlap = bar_group.child("c:overlap")) props.append_attribute("chart:overlap") = overlap.attribute("val").as_string();
    }
    // tick marks inside and outside of the axis, the position of the labels
    for (const char* tick : {"major", "minor"}) {
      pugi::xml_node mark = source.child(tick[1] == 'a' ? "c:majorTickMark" : "c:minorTickMark");
      if (!mark) continue;
      std::string val = mark.attribute("val").as_string("cross");
      props.append_attribute(("chart:tick-marks-" + std::string(tick) + "-inner").c_str()) = val == "in" || val == "cross" ? "true" : "false";
      props.append_attribute(("chart:tick-marks-" + std::string(tick) + "-outer").c_str()) = val == "out" || val == "cross" ? "true" : "false";
    }
    std::string label_pos = source.child("c:tickLblPos").attribute("val").as_string();
    if (label_pos == "none") set_attr(props, "chart:display-label", "false");
    else if (label_pos == "low") props.append_attribute("chart:axis-label-position") = "outside-start";
    else if (label_pos == "high") props.append_attribute("chart:axis-label-position") = "outside-end";
    number_format(source.child("c:numFmt"), style, props);
    text_props(source, style);
    graphic_props(source.child("c:spPr"), style);
    if (source.child("c:delete").attribute("val").as_bool()) {
      pugi::xml_node gp = style.child("style:graphic-properties");
      if (!gp) gp = style.insert_child_after("style:graphic-properties", props);
      set_attr(gp, "draw:stroke", "none");
    }
    pugi::xml_node axis_title = source.child("c:title");
    std::string text = ods_drawing_text(axis_title.child("c:tx"));
    if (!text.empty()) {
      pugi::xml_node title_node = axis.append_child("chart:title");
      pugi::xml_node title_style = new_style();
      title_node.append_attribute("chart:style-name") = title_style.attribute("style:name").value();
      text_props(axis_title, title_style);
      title_node.append_child("text:p").text().set(text.c_str());
    }
    return axis;
  };
  // a grid with the line of the gridlines
  auto add_grid = [&](pugi::xml_node axis, pugi::xml_node lines, const char* cls) {
    pugi::xml_node grid = axis.append_child("chart:grid");
    pugi::xml_node grid_style = new_style();
    graphic_props(lines.child("c:spPr"), grid_style);
    if (grid_style.first_child()) grid.append_attribute("chart:style-name") = grid_style.attribute("style:name").value();
    else autos.remove_child(grid_style);
    grid.append_attribute("chart:class") = cls;
  };
  bool circle = main_class == "chart:circle" || main_class == "chart:ring";
  if (!circle || !categories.empty()) {
    pugi::xml_node axis = add_axis(x_axis, "x", "primary-x", false);
    if (!categories.empty()) axis.append_child("chart:categories").append_attribute("table:cell-range-address") = categories.c_str();
    if (x_axis.child("c:majorGridlines")) add_grid(axis, x_axis.child("c:majorGridlines"), "major");
    if (x_axis.child("c:minorGridlines")) add_grid(axis, x_axis.child("c:minorGridlines"), "minor");
  }
  if (circle) {
    // the slices of pies go clockwise
    pugi::xml_node axis = area.append_child("chart:axis");
    axis.append_attribute("chart:dimension") = "y";
    axis.append_attribute("chart:name") = "primary-y";
    pugi::xml_node style = new_style();
    axis.append_attribute("chart:style-name") = style.attribute("style:name").value();
    pugi::xml_node props = style.append_child("style:chart-properties");
    props.append_attribute("chart:display-label") = "false";
    props.append_attribute("chart:reverse-direction") = "true";
    props.append_attribute("chart:visible") = "false";
  }
  if (!circle) {
    pugi::xml_node axis = add_axis(y_axis, "y", "primary-y", true);
    if (!y_axis || y_axis.child("c:majorGridlines")) add_grid(axis, y_axis.child("c:majorGridlines"), "major");
    if (y_axis.child("c:minorGridlines")) add_grid(axis, y_axis.child("c:minorGridlines"), "minor");
  }
  bool has_secondary = false;
  for (const ods_series& s : series) {
    if (s.secondary) has_secondary = true;
  }
  if (has_secondary && !circle) add_axis(y2_axis, "y", "secondary-y", false);

  static const char* const trend_types[][2] = {
    {"linear", "linear"}, {"exp", "exponential"}, {"log", "logarithmic"}, {"poly", "polynomial"},
    {"power", "power"}, {"movingAvg", "moving-average"}
  };
  // the style of the chart, also of the fallback of an alternate content
  pugi::xml_node style_id = src.child("c:chartSpace").child("c:style");
  if (!style_id) style_id = src.child("c:chartSpace").child("mc:AlternateContent").child("mc:Fallback").child("c:style");
  int32_t chart_style_id = style_id.attribute("val").as_int();
  // the accents of the theme, then those with a luminance modulation and offset
  auto accent_color = [&](size_t k) {
    static const int32_t mods[][2] = {{100000, 0}, {60000, 0}, {80000, 20000}, {80000, 0}, {60000, 40000}, {50000, 0}};
    pugi::xml_document color_doc;
    pugi::xml_node clr = color_doc.append_child("c").append_child("a:solidFill").append_child("a:schemeClr");
    clr.append_attribute("val") = ("accent" + std::to_string(k % 6 + 1)).c_str();
    const int32_t* mod = mods[(k / 6) % 6];
    if (mod[0] != 100000) clr.append_child("a:lumMod").append_attribute("val") = mod[0];
    if (mod[1] != 0) clr.append_child("a:lumOff").append_attribute("val") = mod[1];
    return ods_drawing_color(color_doc.child("c"), theme);
  };
  size_t series_index = 0;
  for (const ods_series& s : series) {
    pugi::xml_node ser = area.append_child("chart:series");
    ser.append_attribute("chart:values-cell-range-address") = s.values.c_str();
    if (!s.label.empty()) ser.append_attribute("chart:label-cell-address") = s.label.c_str();
    if (main_class != "chart:stock") ser.append_attribute("chart:class") = s.cls.c_str();
    if (s.secondary) ser.append_attribute("chart:attached-axis") = "secondary-y";

    // the colors, markers and data labels of the series; series without colors have those of the
    // accents of the theme, darker or lighter after the first six series
    pugi::xml_node style = new_style();
    pugi::xml_node props = style.append_child("style:chart-properties");
    pugi::xml_node graphic = style.append_child("style:graphic-properties");
    std::string color = s.color;
    bool accents = !chart_style_id || chart_style_id == 2;
    bool vary = s.group.child("c:varyColors").attribute("val").as_bool(true) && s.group.child("c:varyColors");
    if (color.empty() && accents && !(circle && vary) && main_class != "chart:stock") color = accent_color(series_index);
    ++series_index;
    if (!color.empty()) {
      graphic.append_attribute("draw:fill-color") = color.c_str();
      graphic.append_attribute("svg:stroke-color") = color.c_str();
    }
    pugi::xml_node ln = s.ser.child("c:spPr").child("a:ln");
    if (ln.child("a:noFill")) graphic.append_attribute("draw:stroke") = "none";
    // the lines of the default chart style are 2.25 points wide
    std::string group_name = s.group.name();
    bool line_group = group_name == "c:lineChart" || group_name == "c:scatterChart" || group_name == "c:radarChart" || group_name == "c:line3DChart";
    if (line_group && !ln.attribute("w") && !ln.child("a:noFill") && accents) graphic.append_attribute("svg:stroke-width") = "2.25pt";
    graphic_props(s.ser.child("c:spPr"), style);
    graphic = style.child("style:graphic-properties");
    std::string group = s.group.name();
    bool markers = group == "c:scatterChart" || (group == "c:lineChart" && s.group.child("c:marker").attribute("val").as_bool(true)) ||
                   (group == "c:radarChart" && std::string(s.group.child("c:radarStyle").attribute("val").as_string()) == "marker");
    std::string symbol = s.ser.child("c:marker").child("c:symbol").attribute("val").as_string();
    if (symbol == "none") markers = false;
    else if (!symbol.empty()) markers = true;
    if (group == "c:lineChart" || group == "c:scatterChart" || group == "c:radarChart") {
      // the symbols of spreadsheet software as named symbols
      static const char* const symbols[][2] = {
        {"circle", "circle"}, {"dot", "circle"}, {"square", "square"}, {"diamond", "diamond"}, {"triangle", "arrow-up"},
        {"x", "x"}, {"star", "asterisk"}, {"dash", "horizontal-bar"}, {"plus", "plus"}
      };
      const char* named = nullptr;
      for (const auto& sym : symbols) {
        if (symbol == sym[0]) named = sym[1];
      }
      props.append_attribute("chart:symbol-type") = markers ? (named ? "named-symbol" : "automatic") : "none";
      if (markers && named) props.append_attribute("chart:symbol-name") = named;
      if (pugi::xml_node size = s.ser.child("c:marker").child("c:size")) {
        if (markers) {
          props.append_attribute("chart:symbol-width") = pt_length(size.attribute("val").as_double());
          props.append_attribute("chart:symbol-height") = pt_length(size.attribute("val").as_double());
        }
      }
    }
    pugi::xml_node labels = s.ser.child("c:dLbls");
    if (!labels) labels = s.group.child("c:dLbls");
    if (labels && !labels.child("c:delete").attribute("val").as_bool()) {
      bool value = labels.child("c:showVal").attribute("val").as_bool();
      bool pct = labels.child("c:showPercent").attribute("val").as_bool();
      if (value || pct) props.append_attribute("chart:data-label-number") = value && pct ? "value-and-percentage" : value ? "value" : "percentage";
      if (labels.child("c:showCatName").attribute("val").as_bool()) props.append_attribute("chart:data-label-text") = "true";
      if (labels.child("c:showSerName").attribute("val").as_bool()) props.append_attribute("chart:data-label-series") = "true";
      static const char* const placements[][2] = {
        {"ctr", "center"}, {"inEnd", "inside"}, {"inBase", "near-origin"}, {"outEnd", "outside"}, {"l", "left"},
        {"r", "right"}, {"t", "top"}, {"b", "bottom"}, {"bestFit", "avoid-overlap"}
      };
      std::string pos = labels.child("c:dLblPos").attribute("val").as_string();
      for (const auto& placement : placements) {
        if (pos == placement[0]) props.append_attribute("chart:label-position") = placement[1];
      }
      if (value) number_format(labels.child("c:numFmt"), style, props);
      if (labels.child("c:txPr")) text_props(labels, style);
    }
    if (pugi::xml_node explosion = s.ser.child("c:explosion")) props.append_attribute("chart:pie-offset") = explosion.attribute("val").as_int();
    if (!props.first_attribute()) style.remove_child(props);
    if (graphic && !graphic.first_attribute()) style.remove_child(graphic);
    if (style.first_child()) {
      ser.append_attribute("chart:style-name") = style.attribute("style:name").value();
    } else {
      autos.remove_child(style);
      --style_count;
    }

    if (!s.domain.empty()) ser.append_child("chart:domain").append_attribute("table:cell-range-address") = s.domain.c_str();
    if (!s.domain2.empty()) ser.append_child("chart:domain").append_attribute("table:cell-range-address") = s.domain2.c_str();

    // trendlines as regression curves
    for (pugi::xml_node tl : s.ser.children("c:trendline")) {
      std::string type = tl.child("c:trendlineType").attribute("val").as_string();
      const char* regression = nullptr;
      for (const auto& tt : trend_types) {
        if (type == tt[0]) regression = tt[1];
      }
      if (!regression) continue;
      pugi::xml_node curve = ser.append_child("chart:regression-curve");
      pugi::xml_node cs = new_style();
      curve.append_attribute("chart:style-name") = cs.attribute("style:name").value();
      pugi::xml_node cp = cs.append_child("style:chart-properties");
      cp.append_attribute("chart:regression-type") = regression;
      std::string name = tl.child_value("c:name");
      if (!name.empty()) cp.append_attribute("chart:regression-name") = name.c_str();
      if (pugi::xml_node order = tl.child("c:order")) cp.append_attribute("chart:regression-max-degree") = order.attribute("val").as_string();
      if (pugi::xml_node period = tl.child("c:period")) cp.append_attribute("chart:regression-period") = period.attribute("val").as_string();
      bool eq = tl.child("c:dispEq").attribute("val").as_bool(), r2 = tl.child("c:dispRSqr").attribute("val").as_bool();
      if (eq || r2) {
        pugi::xml_node equation = curve.append_child("chart:equation");
        equation.append_attribute("chart:display-equation") = eq ? "true" : "false";
        equation.append_attribute("chart:display-r-square") = r2 ? "true" : "false";
      }
    }

    // data points with their own color or offset, the points in between are repeated; the points of
    // plots with varied colors without their own color have the colors of the accents
    std::map<int64_t, pugi::xml_node> own_points;
    for (pugi::xml_node pt : s.ser.children("c:dPt")) own_points[pt.child("c:idx").attribute("val").as_llong(-1)] = pt;
    int64_t points = vary && accents && circle ? static_cast<int64_t>(std::max(s.cache.size(), category_cache.size())) : 0;
    if (!own_points.empty()) points = std::max(points, own_points.rbegin()->first + 1);
    int64_t next = 0;
    for (int64_t idx = 0; idx < points && idx < 1024; ++idx) {
      auto own = own_points.find(idx);
      pugi::xml_node pt = own != own_points.end() ? own->second : pugi::xml_node();
      std::string point_color = ods_drawing_color(pt.child("c:spPr"), theme);
      if (point_color.empty() && vary && accents && circle) point_color = accent_color(static_cast<size_t>(idx));
      pugi::xml_node explosion = pt.child("c:explosion");
      if (point_color.empty() && !explosion) continue;
      if (idx > next) {
        pugi::xml_node gap = ser.append_child("chart:data-point");
        if (idx - next > 1) gap.append_attribute("chart:repeated") = idx - next;
      }
      pugi::xml_node point = ser.append_child("chart:data-point");
      pugi::xml_node ps = new_style();
      point.append_attribute("chart:style-name") = ps.attribute("style:name").value();
      if (explosion) ps.append_child("style:chart-properties").append_attribute("chart:pie-offset") = explosion.attribute("val").as_int();
      pugi::xml_node gp = ps.append_child("style:graphic-properties");
      if (!point_color.empty()) {
        gp.append_attribute("draw:fill-color") = point_color.c_str();
        gp.append_attribute("svg:stroke-color") = point_color.c_str();
      }
      graphic_props(pt.child("c:spPr"), ps);
      if (ps.child("style:graphic-properties") && !ps.child("style:graphic-properties").first_attribute()) ps.remove_child("style:graphic-properties");
      next = idx + 1;
    }
  }

  if (stock_group) {
    if (stock_group.child("c:upDownBars")) {
      area.append_child("chart:stock-gain-marker");
      area.append_child("chart:stock-loss-marker");
    }
    if (stock_group.child("c:hiLowLines")) area.append_child("chart:stock-range-line");
  }
  // the fill and border of the plot area are those of the wall
  pugi::xml_node wall_style = new_style();
  if (plot.child("c:spPr")) graphic_props(plot.child("c:spPr"), wall_style);
  else default_props(wall_style, "none", "none");
  area.append_child("chart:wall").append_attribute("chart:style-name") = wall_style.attribute("style:name").value();

  // the ranges of the chart, the consumer updates the chart when they change
  ranges.clear();
  for (const ods_series& sr : series) {
    for (const std::string& range : {categories, sr.label, sr.domain, sr.domain2, sr.values}) {
      if (range.empty() || range.compare(0, 12, "local-table.") == 0 || ranges.find(range) != std::string::npos) continue;
      if (!ranges.empty()) ranges += " ";
      ranges += range;
    }
  }

  pugi::xml_node table = node.append_child("table:table");
  table.append_attribute("table:name") = "local-table";
  table.append_child("table:table-header-columns").append_child("table:table-column");
  pugi::xml_node columns = table.append_child("table:table-columns").append_child("table:table-column");
  columns.append_attribute("table:number-columns-repeated") = static_cast<uint32_t>(local_columns.size() - 1);
  auto local_cell = [](pugi::xml_node row, const std::pair<bool, std::string>& value) {
    pugi::xml_node cell = row.append_child("table:table-cell");
    if (value.second.empty()) return;
    char* end = nullptr;
    std::strtod(value.second.c_str(), &end);
    bool number = value.first && end && *end == '\0';
    cell.append_attribute("office:value-type") = number ? "float" : "string";
    if (number) cell.append_attribute("office:value") = value.second.c_str();
    cell.append_child("text:p").text().set(value.second.c_str());
  };
  pugi::xml_node header = table.append_child("table:table-header-rows").append_child("table:table-row");
  for (const auto& column : local_columns) local_cell(header, column.first);
  pugi::xml_node body = table.append_child("table:table-rows");
  for (size_t k = 0; k < local_rows; ++k) {
    pugi::xml_node row = body.append_child("table:table-row");
    for (const auto& column : local_columns) local_cell(row, k < column.second.size() ? column.second[k] : no_header);
  }

  ods_save(doc, path);

  // the styles of the object
  pugi::xml_document styles;
  ods_declaration(styles);
  pugi::xml_node sroot = styles.append_child("office:document-styles");
  sroot.append_attribute("xmlns:office") = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
  sroot.append_attribute("xmlns:style") = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
  if (!dashes.empty()) sroot.append_attribute("xmlns:draw") = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
  sroot.append_attribute("office:version") = version.c_str();
  pugi::xml_node office_styles = sroot.append_child("office:styles");
  for (const ods_dash& dash : ods_dashes) {
    if (dashes.find(dash.name) == dashes.end()) continue;
    pugi::xml_node sd = office_styles.append_child("draw:stroke-dash");
    sd.append_attribute("draw:name") = dash.name;
    sd.append_attribute("draw:style") = "rect";
    sd.append_attribute("draw:dots1") = dash.dots1;
    sd.append_attribute("draw:dots1-length") = (std::to_string(dash.length1) + "%").c_str();
    if (dash.dots2 > 0) {
      sd.append_attribute("draw:dots2") = dash.dots2;
      sd.append_attribute("draw:dots2-length") = (std::to_string(dash.length2) + "%").c_str();
    }
    sd.append_attribute("draw:distance") = (std::to_string(dash.distance) + "%").c_str();
  }
  ods_save(styles, path.substr(0, path.rfind('/') + 1) + "styles.xml");
  return true;
}

// header and footer codes (&P, &N, &D, &T, &A, &F, &Z) as ODF text fields; the font, size, color and
// text decoration codes as spans with text styles of the automatic styles, shared by their properties
static void ods_header_region(pugi::xml_node region, const std::string& code, pugi::xml_node autos, std::map<std::string, std::string>& text_styles) {
  pugi::xml_node p = region.append_child("text:p");
  std::string family, color;
  double size = 0.0;
  bool bold = false, italic = false, underline = false, double_underline = false, strike = false, super = false, sub = false;
  // the paragraph or a span with the text style of the current properties
  pugi::xml_node target;
  auto current = [&]() {
    if (target) return target;
    pugi::xml_document style_doc;
    pugi::xml_node tp = style_doc.append_child("style:text-properties");
    if (!family.empty()) tp.append_attribute("fo:font-family") = family.c_str();
    if (size > 0.0) tp.append_attribute("fo:font-size") = (ods_num(size) + "pt").c_str();
    if (bold) tp.append_attribute("fo:font-weight") = "bold";
    if (italic) tp.append_attribute("fo:font-style") = "italic";
    if (underline || double_underline) {
      tp.append_attribute("style:text-underline-style") = "solid";
      tp.append_attribute("style:text-underline-width") = "auto";
      tp.append_attribute("style:text-underline-color") = "font-color";
      if (double_underline) tp.append_attribute("style:text-underline-type") = "double";
    }
    if (strike) tp.append_attribute("style:text-line-through-style") = "solid";
    if (super || sub) tp.append_attribute("style:text-position") = super ? "super 58%" : "sub 58%";
    if (!color.empty()) tp.append_attribute("fo:color") = color.c_str();
    if (!tp.first_attribute()) {
      target = p;
      return target;
    }
    std::string key = ods_xml_string(tp);
    auto known = text_styles.find(key);
    if (known == text_styles.end()) {
      std::string name = "HT" + std::to_string(text_styles.size() + 1);
      pugi::xml_node style = autos.append_child("style:style");
      style.append_attribute("style:name") = name.c_str();
      style.append_attribute("style:family") = "text";
      style.append_copy(tp);
      known = text_styles.insert(std::make_pair(key, name)).first;
    }
    target = p.append_child("text:span");
    target.append_attribute("text:style-name") = known->second.c_str();
    return target;
  };
  std::string run;
  auto flush = [&]() {
    if (!run.empty()) current().append_child(pugi::node_pcdata).set_value(run.c_str());
    run.clear();
  };
  size_t n = code.size();
  for (size_t i = 0; i < n; ++i) {
    char c = code[i];
    if (c != '&' && c != '\n') {
      run += c;
      continue;
    }
    if (c == '&' && i + 1 < n && code[i + 1] == '&') {
      run += '&';
      ++i;
      continue;
    }
    flush();
    if (c == '\n') {
      p = region.append_child("text:p");
      target = pugi::xml_node();
      continue;
    }
    if (i + 1 >= n) break;
    char next = code[++i];
    if (next == 'P') current().append_child("text:page-number").text().set("1");
    else if (next == 'N') current().append_child("text:page-count").text().set("1");
    else if (next == 'D') current().append_child("text:date");
    else if (next == 'T') current().append_child("text:time");
    else if (next == 'A') current().append_child("text:sheet-name");
    else if (next == 'F') current().append_child("text:file-name").append_attribute("text:display") = "name-and-extension";
    else if (next == 'Z') current().append_child("text:file-name").append_attribute("text:display") = "path";
    else if (next == '"') {
      // the font and its style, a font name of - keeps the font
      size_t end = code.find('"', i + 1);
      std::string font = code.substr(i + 1, end == std::string::npos ? std::string::npos : end - i - 1);
      i = end == std::string::npos ? n : end;
      size_t comma = font.find(',');
      std::string name = font.substr(0, comma), style = comma == std::string::npos ? "" : ods_lower(font.substr(comma + 1));
      if (name != "-") family = name;
      if (comma != std::string::npos) {
        bold = style.find("bold") != std::string::npos;
        italic = style.find("italic") != std::string::npos;
      }
      target = pugi::xml_node();
    } else if (next == 'K') {
      std::string hex = code.substr(i + 1, 6);
      if (hex.size() == 6 && hex.find_first_not_of("0123456789abcdefABCDEF") == std::string::npos) color = "#" + ods_lower(hex);
      i = std::min(i + 6, n - 1);
      target = pugi::xml_node();
    } else if (std::isdigit(static_cast<unsigned char>(next))) {
      std::string digits(1, next);
      while (i + 1 < n && std::isdigit(static_cast<unsigned char>(code[i + 1]))) digits += code[++i];
      size = std::strtod(digits.c_str(), nullptr);
      target = pugi::xml_node();
    } else {
      if (next == 'B') bold = !bold;
      else if (next == 'I') italic = !italic;
      else if (next == 'U') underline = !underline;
      else if (next == 'E') double_underline = !double_underline;
      else if (next == 'S') strike = !strike;
      else if (next == 'X') super = !super;
      else if (next == 'Y') sub = !sub;
      target = pugi::xml_node();
    }
  }
  flush();
}

// paper sizes in mm of the paperSize values of the page setup
static const struct {
  uint32_t id;
  double width, height;
} ods_papers[] = {
  {1, 215.9, 279.4}, {3, 279.4, 431.8}, {5, 215.9, 355.6}, {7, 184.15, 266.7}, {8, 297, 420},
  {9, 210, 297}, {11, 148, 210}, {12, 257, 364}, {13, 182, 257}, {70, 105, 148}
};

static std::string ods_media_type(const std::string& file) {
  std::string ext = ods_lower(file.substr(file.rfind('.') + 1));
  if (ext == "png") return "image/png";
  if (ext == "jpg" || ext == "jpeg") return "image/jpeg";
  if (ext == "gif") return "image/gif";
  if (ext == "bmp") return "image/bmp";
  if (ext == "svg") return "image/svg+xml";
  if (ext == "tif" || ext == "tiff") return "image/tiff";
  if (ext == "emf") return "image/x-emf";
  if (ext == "wmf") return "image/x-wmf";
  return "application/octet-stream";
}

// data validations, conditional formats and shapes

// the parts of a sqref (A1:B2 C3 D:D) as cell ranges
struct ods_range {
  int64_t r1, c1, r2, c2;
};

static std::vector<ods_range> ods_sqref(const std::string& sqref) {
  std::vector<ods_range> out;
  size_t pos = 0;
  while (pos < sqref.size()) {
    size_t end = sqref.find(' ', pos);
    if (end == std::string::npos) end = sqref.size();
    std::string part = sqref.substr(pos, end - pos);
    pos = end + 1;
    if (part.empty()) continue;
    ods_ref_part p1, p2;
    size_t k = ods_parse_ref_part(part, 0, p1);
    if (k == 0) continue;
    p2 = p1;
    if (k < part.size() && part[k] == ':' && ods_parse_ref_part(part, k + 1, p2) != part.size()) continue;
    ods_range r = {p1.row > 0 ? p1.row : 1, p1.col > 0 ? p1.col : 1,
                   p2.row > 0 ? p2.row : MAX_OOXML_ROW_INT, p2.col > 0 ? p2.col : MAX_OOXML_COL_INT};
    if (r.r2 < r.r1) std::swap(r.r1, r.r2);
    if (r.c2 < r.c1) std::swap(r.c1, r.c2);
    out.push_back(r);
  }
  return out;
}

static std::string ods_range_address(const std::string& sheet, const ods_range& r) {
  std::string name = ods_sheet_name(sheet);
  std::string out = name + "." + int_to_col(r.c1) + std::to_string(r.r1);
  if (r.r1 != r.r2 || r.c1 != r.c2) out += ":" + name + "." + int_to_col(r.c2) + std::to_string(r.r2);
  return out;
}

// the name of a shared item of a pivot cache field, as data pilot member name
static std::string ods_pivot_item(pugi::xml_node cache_field, int64_t x) {
  int64_t k = 0;
  for (pugi::xml_node item : cache_field.child("sharedItems").children()) {
    if (k++ != x) continue;
    std::string name = item.name();
    if (name == "b") return item.attribute("v").as_bool() ? "TRUE" : "FALSE";
    if (name == "m") return "";
    return item.attribute("v").as_string();
  }
  return "";
}

// a pivot table and its cache as data pilot table. The target range starts with the page fields.
// first_values are the values of the first data row of the source, one for each cache field.
static bool ods_data_pilot(pugi::xml_node pilots, pugi::xml_node pt, pugi::xml_node cache,
                           const std::string& sheet_quoted, const std::vector<ods_table>& tables,
                           const std::vector<utf8_string>& first_values) {
  pugi::xml_node source = cache.child("cacheSource").child("worksheetSource");
  std::string source_address;
  if (source.attribute("sheet") && source.attribute("ref")) {
    std::vector<ods_range> src = ods_sqref(source.attribute("ref").as_string());
    if (src.empty()) return false;
    std::string quoted;
    for (char c : std::string(source.attribute("sheet").as_string())) {
      quoted += c;
      if (c == '\'') quoted += c;
    }
    source_address = ods_range_address(quoted, src[0]);
  } else if (pugi::xml_attribute name = source.attribute("name")) {
    for (const ods_table& t : tables) {
      if (t.name != name.as_string()) continue;
      ods_range range = {t.r1, t.c1, t.r2, t.c2};
      source_address = ods_range_address(t.sheet, range);
    }
  }
  if (source_address.empty()) return false;

  std::vector<ods_range> location = ods_sqref(pt.child("location").attribute("ref").as_string());
  if (location.empty()) return false;
  std::vector<pugi::xml_node> cache_fields, pivot_fields;
  for (pugi::xml_node cf : cache.child("cacheFields").children("cacheField")) cache_fields.push_back(cf);
  for (pugi::xml_node pf : pt.child("pivotFields").children("pivotField")) pivot_fields.push_back(pf);

  pugi::xml_object_range<pugi::xml_named_node_iterator> page_fields = pt.child("pageFields").children("pageField");
  int64_t pages = std::distance(page_fields.begin(), page_fields.end());
  ods_range target = location[0];
  if (pages > 0) target.r1 = std::max<int64_t>(1, target.r1 - pages - 1);

  pugi::xml_node dp = pilots.append_child("table:data-pilot-table");
  dp.append_attribute("table:name") = pt.attribute("name").as_string("DataPilot");
  dp.append_attribute("table:target-range-address") = ods_range_address(sheet_quoted, target).c_str();
  bool row_grand = pt.attribute("rowGrandTotals").as_bool(true), col_grand = pt.attribute("colGrandTotals").as_bool(true);
  if (!row_grand || !col_grand) {
    dp.append_attribute("table:grand-total") = row_grand ? "row" : (col_grand ? "column" : "none");
  }

  // the field buttons: the page fields, the row header and the column header
  pugi::xml_node loc = pt.child("location");
  int64_t first_header = loc.attribute("firstHeaderRow").as_llong(0), first_data = loc.attribute("firstDataRow").as_llong(1);
  int64_t first_col = loc.attribute("firstDataCol").as_llong(1);
  std::string buttons;
  auto add_button = [&](int64_t row, int64_t col) {
    if (row < 1 || col < 1) return;
    if (!buttons.empty()) buttons += " ";
    buttons += ods_sheet_name(sheet_quoted) + "." + int_to_col(col) + std::to_string(row);
  };
  for (int64_t i = 0; i < pages; ++i) add_button(target.r1 + i, location[0].c1);
  if (pt.child("rowFields").child("field")) add_button(location[0].r1 + first_data - 1, location[0].c1);
  for (pugi::xml_node field : pt.child("colFields").children("field")) {
    if (field.attribute("x").as_llong(-2) < 0) continue;
    add_button(location[0].r1 + first_header - 1, location[0].c1 + first_col);
    break;
  }
  if (!buttons.empty()) dp.append_attribute("table:buttons") = buttons.c_str();
  dp.append_attribute("table:show-filter-button") = "false";
  dp.append_child("table:source-cell-range").append_attribute("table:cell-range-address") = source_address.c_str();

  // a source field with its hidden members, a selected page or the function of a data field
  // calculated fields of the cache are not part of the source range and are left out
  auto add_field = [&](int64_t fld, const char* orientation, const char* function, const std::string& selected, pugi::xml_node data_field) {
    if (fld >= 0 && static_cast<size_t>(fld) < cache_fields.size() && cache_fields[static_cast<size_t>(fld)].attribute("formula")) return;
    pugi::xml_node field = dp.append_child("table:data-pilot-field");
    if (fld < 0) {
      field.append_attribute("table:source-field-name") = "";
      field.append_attribute("table:is-data-layout-field") = "true";
      field.append_attribute("table:orientation") = orientation;
      return;
    }
    size_t f = static_cast<size_t>(fld);
    field.append_attribute("table:source-field-name") = f < cache_fields.size() ? cache_fields[f].attribute("name").as_string() : "";
    field.append_attribute("table:orientation") = orientation;
    if (function) field.append_attribute("table:function") = function;
    if (std::strcmp(orientation, "page") == 0) field.append_attribute("table:selected-page") = selected.c_str();

    // a data field shown relative to a field, an item or the totals
    if (data_field) {
      static const char* const shown[][2] = {
        {"difference", "member-difference"}, {"percent", "member-percentage"}, {"percentDiff", "member-percentage-difference"},
        {"runTotal", "running-total"}, {"percentOfRow", "row-percentage"}, {"percentOfCol", "column-percentage"},
        {"percentOfTotal", "total-percentage"}, {"index", "index"}
      };
      std::string show = data_field.attribute("showDataAs").as_string();
      const char* type = nullptr;
      for (const auto& sh : shown) {
        if (show == sh[0]) type = sh[1];
      }
      if (!type) return;
      pugi::xml_node reference = field.append_child("table:data-pilot-field-reference");
      size_t base = static_cast<size_t>(data_field.attribute("baseField").as_llong(0));
      reference.append_attribute("table:field-name") = base < cache_fields.size() ? cache_fields[base].attribute("name").as_string() : "";
      // a named base item requires its name, without a name the previous item is used
      int64_t base_item = data_field.attribute("baseItem").as_llong(0);
      std::string member_type = base_item == 1048828 ? "previous" : base_item == 1048829 ? "next" : "named", member_name;
      if (member_type == "named" && base < pivot_fields.size() && base < cache_fields.size()) {
        int64_t k = 0;
        for (pugi::xml_node it : pivot_fields[base].child("items").children("item")) {
          if (k++ == base_item && it.attribute("x")) member_name = ods_pivot_item(cache_fields[base], it.attribute("x").as_llong());
        }
        if (member_name.empty()) member_name = ods_pivot_item(cache_fields[base], 0);
      }
      if (member_type == "named" && member_name.empty()) member_type = "previous";
      reference.append_attribute("table:member-type") = member_type.c_str();
      if (member_type == "named") reference.append_attribute("table:member-name") = member_name.c_str();
      reference.append_attribute("table:type") = type;
      return;
    }
    if (function || f >= pivot_fields.size() || f >= cache_fields.size()) return;

    // hidden members and the sort order of the items
    pugi::xml_node level, members;
    for (pugi::xml_node item : pivot_fields[f].child("items").children("item")) {
      if (!item.attribute("x") || !item.attribute("h").as_bool()) continue;
      if (!level) level = field.append_child("table:data-pilot-level");
      if (!members) members = level.append_child("table:data-pilot-members");
      pugi::xml_node member = members.append_child("table:data-pilot-member");
      member.append_attribute("table:name") = ods_pivot_item(cache_fields[f], item.attribute("x").as_llong()).c_str();
      member.append_attribute("table:display") = "false";
    }
    std::string sort = pivot_fields[f].attribute("sortType").as_string();
    if (sort == "ascending" || sort == "descending") {
      if (!level) level = field.append_child("table:data-pilot-level");
      pugi::xml_node sort_info = level.append_child("table:data-pilot-sort-info");
      sort_info.append_attribute("table:sort-mode") = "name";
      sort_info.append_attribute("table:order") = sort.c_str();
    }

    // the layout of a row field: tabular, outline or compact, which is an outline with all fields in one column
    if (std::strcmp(orientation, "row") != 0) return;
    pugi::xml_node pf = pivot_fields[f];
    bool outline = pt.attribute("outline").as_bool(false) && pf.attribute("outline").as_bool(true);
    bool compact = outline && pt.attribute("compact").as_bool(true) && pf.attribute("compact").as_bool(true);
    if (!level) level = field.append_child("table:data-pilot-level");
    pugi::xml_node layout_info = level.append_child("table:data-pilot-layout-info");
    layout_info.append_attribute("table:add-empty-lines") = pf.attribute("insertBlankRow").as_bool(false) ? "true" : "false";
    layout_info.append_attribute("table:layout-mode") = !outline ? "tabular-layout" : pf.attribute("subtotalTop").as_bool(true) ? "outline-subtotals-top" : "outline-subtotals-bottom";
    if (compact) layout_info.append_attribute("loext:layout-mode") = "compact-layout";
  };

  pugi::xml_object_range<pugi::xml_named_node_iterator> data_list = pt.child("dataFields").children("dataField");
  int64_t data_fields = std::distance(data_list.begin(), data_list.end());
  bool layout = false;
  for (pugi::xml_node fl : pt.child("rowFields").children("field")) {
    if (fl.attribute("x").as_llong() < 0) layout = true;
    add_field(fl.attribute("x").as_llong(), "row", nullptr, "", pugi::xml_node());
  }
  for (pugi::xml_node fl : pt.child("colFields").children("field")) {
    if (fl.attribute("x").as_llong() < 0) layout = true;
    add_field(fl.attribute("x").as_llong(), "column", nullptr, "", pugi::xml_node());
  }
  if (!layout && data_fields > 1) add_field(-2, pt.attribute("dataOnRows").as_bool() ? "row" : "column", nullptr, "", pugi::xml_node());
  // a page field requires a selected member. Without a selection in the pivot table
  // the first member is written and the spreadsheet software is told to ignore it.
  for (pugi::xml_node pg : page_fields) {
    int64_t fld = pg.attribute("fld").as_llong();
    size_t f = static_cast<size_t>(fld);
    std::string selected;
    pugi::xml_attribute item = pg.attribute("item");
    if (item && f < pivot_fields.size() && f < cache_fields.size()) {
      int64_t k = 0;
      for (pugi::xml_node it : pivot_fields[f].child("items").children("item")) {
        if (k++ == item.as_llong()) selected = ods_pivot_item(cache_fields[f], it.attribute("x").as_llong());
      }
    }
    if (!item) {
      if (f < cache_fields.size() && cache_fields[f].child("sharedItems").first_child()) selected = ods_pivot_item(cache_fields[f], 0);
      else if (f < first_values.size()) selected = first_values[f];
    }
    pugi::xml_node last = dp.last_child();
    add_field(fld, "page", nullptr, selected, pugi::xml_node());
    if (!item && dp.last_child() != last) dp.last_child().append_attribute("loext:ignore-selected-page") = "true";
  }

  static const char* const functions[][2] = {
    {"sum", "sum"}, {"count", "count"}, {"average", "average"}, {"max", "max"}, {"min", "min"},
    {"product", "product"}, {"countNums", "countnums"}, {"stdDev", "stdev"}, {"stdDevp", "stdevp"},
    {"var", "var"}, {"varp", "varp"}
  };
  for (pugi::xml_node df : data_list) {
    std::string subtotal = df.attribute("subtotal").as_string("sum");
    const char* function = "sum";
    for (const auto& fn : functions) {
      if (subtotal == fn[0]) function = fn[1];
    }
    add_field(df.attribute("fld").as_llong(), "data", function, "", df);
  }
  return true;
}

// a formula without the of:= prefix
static std::string ods_of_value(const std::string& f, const ods_formula_ctx& ctx) {
  return ods_formula_to_of(f, 0, 0, ctx).substr(4);
}

static const char* const ods_operators[][2] = {
  {"equal", "="}, {"notEqual", "!="}, {"greaterThan", ">"}, {"lessThan", "<"},
  {"greaterThanOrEqual", ">="}, {"lessThanOrEqual", "<="}
};

// the condition of a data validation
static std::string ods_validation_condition(pugi::xml_node dv, const ods_formula_ctx& ctx) {
  std::string type = dv.attribute("type").as_string();
  std::string op = dv.attribute("operator").as_string("between");
  std::string f1 = dv.child_value("formula1"), f2 = dv.child_value("formula2");

  if (type == "list") {
    if (!f1.empty() && f1[0] == '"') {
      std::string items, item;
      std::string inner = f1.substr(1, f1.size() >= 2 ? f1.size() - 2 : 0);
      for (size_t i = 0; i <= inner.size(); ++i) {
        if (i == inner.size() || inner[i] == ',') {
          items += (items.empty() ? "\"" : ";\"") + item + "\"";
          item.clear();
        } else {
          item += inner[i];
        }
      }
      return "of:cell-content-is-in-list(" + items + ")";
    }
    return "of:cell-content-is-in-list(" + ods_of_value(f1, ctx) + ")";
  }
  if (type == "custom") return "of:is-true-formula(" + ods_of_value(f1, ctx) + ")";

  std::string prefix, content = "cell-content()", between = "cell-content-is-between(", not_between = "cell-content-is-not-between(";
  if (type == "whole") prefix = "cell-content-is-whole-number() and ";
  else if (type == "decimal") prefix = "cell-content-is-decimal-number() and ";
  else if (type == "date") prefix = "cell-content-is-date() and ";
  else if (type == "time") prefix = "cell-content-is-time() and ";
  else if (type == "textLength") {
    content = "cell-content-text-length()";
    between = "cell-content-text-length-is-between(";
    not_between = "cell-content-text-length-is-not-between(";
  } else {
    return "";
  }

  if (op == "between") return "of:" + prefix + between + ods_of_value(f1, ctx) + "," + ods_of_value(f2, ctx) + ")";
  if (op == "notBetween") return "of:" + prefix + not_between + ods_of_value(f1, ctx) + "," + ods_of_value(f2, ctx) + ")";
  for (const auto& o : ods_operators) {
    if (op == o[0]) return "of:" + prefix + content + o[1] + ods_of_value(f1, ctx);
  }
  return "";
}

// the arguments of a function call starting at pos (after the opening bracket),
// split at commas or semicolons outside of strings and brackets
static std::vector<std::string> ods_arguments(const std::string& s, size_t pos, size_t& end) {
  std::vector<std::string> args;
  std::string cur;
  int32_t depth = 0;
  bool quoted = false;
  size_t i = pos;
  for (; i < s.size(); ++i) {
    char c = s[i];
    if (c == '"') quoted = !quoted;
    if (!quoted) {
      if (c == '(' || c == '[' || c == '{') ++depth;
      if (c == ')' || c == ']' || c == '}') {
        if (depth == 0) break;
        --depth;
      }
      if ((c == ',' || c == ';') && depth == 0) {
        args.push_back(cur);
        cur.clear();
        continue;
      }
    }
    cur += c;
  }
  args.push_back(cur);
  end = i;
  return args;
}

// an ODF validation condition as type, operator and formulas of a data validation
static bool ods_validation_xlsx(std::string cond, pugi::xml_node dv, const std::map<std::string, std::string>& renames) {
  if (cond.compare(0, 3, "of:") == 0) cond = cond.substr(3);
  if (cond.compare(0, 5, "oooc:") == 0) cond = cond.substr(5);
  size_t end = 0;
  auto set_formulas = [&](const std::vector<std::string>& args) {
    if (!args.empty()) dv.append_child("formula1").text().set(ods_formula_to_xlsx(args[0], renames).c_str());
    if (args.size() > 1) dv.append_child("formula2").text().set(ods_formula_to_xlsx(args[1], renames).c_str());
  };

  const std::string list = "cell-content-is-in-list(";
  if (cond.compare(0, list.size(), list) == 0) {
    dv.append_attribute("type") = "list";
    std::vector<std::string> args = ods_arguments(cond, list.size(), end);
    bool literal = true;
    for (const std::string& a : args) {
      if (a.size() < 2 || a.front() != '"' || a.back() != '"') literal = false;
    }
    if (literal) {
      std::string items;
      for (const std::string& a : args) items += (items.empty() ? "" : ",") + a.substr(1, a.size() - 2);
      dv.append_child("formula1").text().set(("\"" + items + "\"").c_str());
    } else {
      dv.append_child("formula1").text().set(ods_formula_to_xlsx(args[0], renames).c_str());
    }
    return true;
  }
  const std::string custom = "is-true-formula(";
  if (cond.compare(0, custom.size(), custom) == 0) {
    dv.append_attribute("type") = "custom";
    set_formulas(ods_arguments(cond, custom.size(), end));
    return true;
  }

  static const char* const types[][2] = {
    {"cell-content-is-whole-number() and ", "whole"}, {"cell-content-is-decimal-number() and ", "decimal"},
    {"cell-content-is-date() and ", "date"}, {"cell-content-is-time() and ", "time"}
  };
  std::string type;
  for (const auto& t : types) {
    size_t len = std::strlen(t[0]);
    if (cond.compare(0, len, t[0]) == 0) {
      type = t[1];
      cond = cond.substr(len);
    }
  }
  const std::string text_length = "cell-content-text-length";
  if (cond.compare(0, text_length.size(), text_length) == 0) type = "textLength";
  if (type.empty()) return false;
  dv.append_attribute("type") = type.c_str();

  static const char* const ranges[][2] = {
    {"cell-content-is-between(", "between"}, {"cell-content-is-not-between(", "notBetween"},
    {"cell-content-text-length-is-between(", "between"}, {"cell-content-text-length-is-not-between(", "notBetween"}
  };
  for (const auto& rg : ranges) {
    size_t len = std::strlen(rg[0]);
    if (cond.compare(0, len, rg[0]) != 0) continue;
    dv.append_attribute("operator") = rg[1];
    set_formulas(ods_arguments(cond, len, end));
    return true;
  }
  for (const char* content : {"cell-content()", "cell-content-text-length()"}) {
    size_t len = std::strlen(content);
    if (cond.compare(0, len, content) != 0) continue;
    std::string rest = cond.substr(len);
    // the longest operators first
    static const char* const ops[] = {"!=", ">=", "<=", "=", ">", "<"};
    for (const char* o : ops) {
      size_t olen = std::strlen(o);
      if (rest.compare(0, olen, o) != 0) continue;
      for (const auto& op : ods_operators) {
        if (std::strcmp(op[1], o) == 0) dv.append_attribute("operator") = op[0];
      }
      dv.append_child("formula1").text().set(ods_formula_to_xlsx(rest.substr(olen), renames).c_str());
      return true;
    }
  }
  return false;
}

// conditional formats in the calcext: namespace of the LibreOffice extensions
static const char* const ods_cfvo_types[][2] = {
  {"min", "minimum"}, {"max", "maximum"}, {"num", "number"}, {"percent", "percent"},
  {"percentile", "percentile"}, {"formula", "formula"}
};

static void ods_cf_entries(pugi::xml_node rule, pugi::xml_node out, const char* element, bool bar) {
  for (pugi::xml_node cfvo : rule.children("cfvo")) {
    std::string type = cfvo.attribute("type").as_string();
    std::string odf_type = "number";
    for (const auto& t : ods_cfvo_types) {
      if (type == t[0]) odf_type = t[1];
    }
    if (bar && type == "min") odf_type = "auto-minimum";
    if (bar && type == "max") odf_type = "auto-maximum";
    pugi::xml_node entry = out.append_child(element);
    entry.append_attribute("calcext:value") = cfvo.attribute("val").as_string("0");
    entry.append_attribute("calcext:type") = odf_type.c_str();
  }
}

static void ods_conditional_format(pugi::xml_node format, pugi::xml_node rule, const ods_range& base, const std::string& sheet,
                                   const ods_formula_ctx& ctx, const std::vector<std::string>& theme, const std::vector<std::string>& palette,
                                   const std::map<std::string, std::string>& bar_extensions) {
  std::string type = rule.attribute("type").as_string();
  if (type == "colorScale") {
    pugi::xml_node scale = rule.child("colorScale");
    pugi::xml_node out = format.append_child("calcext:color-scale");
    std::vector<pugi::xml_node> colors;
    for (pugi::xml_node color : scale.children("color")) colors.push_back(color);
    size_t k = 0;
    for (pugi::xml_node cfvo : scale.children("cfvo")) {
      std::string cfvo_type = cfvo.attribute("type").as_string();
      std::string odf_type = "number";
      for (const auto& t : ods_cfvo_types) {
        if (cfvo_type == t[0]) odf_type = t[1];
      }
      pugi::xml_node entry = out.append_child("calcext:color-scale-entry");
      entry.append_attribute("calcext:value") = cfvo.attribute("val").as_string("0");
      entry.append_attribute("calcext:type") = odf_type.c_str();
      std::string color = k < colors.size() ? ods_color(colors[k], theme, palette) : "";
      entry.append_attribute("calcext:color") = color.empty() ? "#000000" : color.c_str();
      ++k;
    }
    return;
  }
  if (type == "dataBar") {
    pugi::xml_node bar = rule.child("dataBar");
    pugi::xml_node out = format.append_child("calcext:data-bar");
    std::string color = ods_color(bar.child("color"), theme, palette);
    // the extension of the data bar has the negative and axis colors, the fill and the lengths
    pugi::xml_document ext_doc;
    pugi::xml_node ext;
    std::map<std::string, std::string>::const_iterator found = bar_extensions.find(ods_descendant(rule, "x14:id").text().get());
    if (found != bar_extensions.end() && ext_doc.load_string(found->second.c_str())) ext = ext_doc.document_element();
    std::string negative, axis;
    for (pugi::xml_node part : ext.children()) {
      std::string name = part.name();
      if (name.size() > 17 && name.compare(name.size() - 17, 17, "negativeFillColor") == 0) negative = ods_color(part, theme, palette);
      if (name.size() > 9 && name.compare(name.size() - 9, 9, "axisColor") == 0) axis = ods_color(part, theme, palette);
    }
    out.append_attribute("calcext:positive-color") = color.empty() ? "#638ec6" : color.c_str();
    out.append_attribute("calcext:negative-color") = negative.empty() ? "#ff0000" : negative.c_str();
    if (!axis.empty()) out.append_attribute("calcext:axis-color") = axis.c_str();
    std::string position = ext.attribute("axisPosition").as_string("automatic");
    if (position == "middle" || position == "none") out.append_attribute("calcext:axis-position") = position.c_str();
    if (ext && !ext.attribute("gradient").as_bool(true)) out.append_attribute("calcext:gradient") = "false";
    if (!bar.attribute("showValue").as_bool(true)) out.append_attribute("calcext:show-value") = "false";
    if (pugi::xml_attribute min_length = ext.attribute("minLength")) out.append_attribute("calcext:min-length") = min_length.as_string();
    if (pugi::xml_attribute max_length = ext.attribute("maxLength")) out.append_attribute("calcext:max-length") = max_length.as_string();
    ods_cf_entries(bar, out, "calcext:formatting-entry", true);
    return;
  }
  if (type == "iconSet") {
    pugi::xml_node set = rule.child("iconSet");
    pugi::xml_node out = format.append_child("calcext:icon-set");
    out.append_attribute("calcext:icon-set-type") = set.attribute("iconSet").as_string("3TrafficLights1");
    if (!set.attribute("showValue").as_bool(true)) out.append_attribute("calcext:show-value") = "false";
    ods_cf_entries(set, out, "calcext:formatting-entry", false);
    return;
  }

  // dates in a period relative to today
  static const char* const periods[][2] = {
    {"today", "today"}, {"yesterday", "yesterday"}, {"tomorrow", "tomorrow"}, {"last7Days", "last-7-days"}, {"thisWeek", "this-week"},
    {"lastWeek", "last-week"}, {"nextWeek", "next-week"}, {"thisMonth", "this-month"}, {"lastMonth", "last-month"}, {"nextMonth", "next-month"}
  };
  if (type == "timePeriod") {
    for (const auto& period : periods) {
      if (std::strcmp(rule.attribute("timePeriod").as_string(), period[0]) != 0) continue;
      pugi::xml_node date = format.append_child("calcext:date-is");
      date.append_attribute("calcext:date") = period[1];
      pugi::xml_attribute dxf = rule.attribute("dxfId");
      date.append_attribute("calcext:style") = dxf ? ("ConditionalStyle_" + std::to_string(dxf.as_uint() + 1)).c_str() : "Default";
    }
    return;
  }

  std::string value;
  std::string f1 = rule.child_value("formula");
  std::string text = rule.attribute("text").as_string();
  if (type == "cellIs") {
    std::vector<std::string> formulas;
    for (pugi::xml_node f : rule.children("formula")) formulas.push_back(ods_of_value(f.text().get(), ctx));
    std::string op = rule.attribute("operator").as_string();
    if ((op == "between" || op == "notBetween") && formulas.size() == 2) {
      value = (op == "between" ? "between(" : "not-between(") + formulas[0] + "," + formulas[1] + ")";
    }
    for (const auto& o : ods_operators) {
      if (op == o[0] && !formulas.empty()) value = std::string(o[1]) + formulas[0];
    }
  } else if (type == "containsText" || type == "notContainsText" || type == "beginsWith" || type == "endsWith") {
    std::string fn = type == "containsText" ? "contains-text" : type == "notContainsText" ? "not-contains-text" : type == "beginsWith" ? "begins-with" : "ends-with";
    std::string quoted;
    for (char c : text) quoted += c == '"' ? std::string("\"\"") : std::string(1, c);
    value = fn + "(\"" + quoted + "\")";
  } else if (type == "containsErrors" || type == "notContainsErrors") {
    value = type == "containsErrors" ? "is-error" : "is-no-error";
  } else if (type == "duplicateValues") {
    value = "duplicate";
  } else if (type == "uniqueValues") {
    value = "unique";
  } else if (type == "top10") {
    bool bottom = rule.attribute("bottom").as_bool(), percent = rule.attribute("percent").as_bool();
    value = std::string(bottom ? "bottom-" : "top-") + (percent ? "percent(" : "elements(") + rule.attribute("rank").as_string("10") + ")";
  } else if (type == "aboveAverage") {
    bool above = rule.attribute("aboveAverage").as_bool(true), equal = rule.attribute("equalAverage").as_bool();
    value = std::string(above ? "above" : "below") + (equal ? "-equal" : "") + "-average";
  } else if (!f1.empty()) {
    value = "formula-is(" + ods_of_value(f1, ctx) + ")";
  }
  if (value.empty()) return;

  pugi::xml_node condition = format.append_child("calcext:condition");
  pugi::xml_attribute dxf = rule.attribute("dxfId");
  if (dxf) condition.append_attribute("calcext:apply-style-name") = ("ConditionalStyle_" + std::to_string(dxf.as_uint() + 1)).c_str();
  else condition.append_attribute("calcext:apply-style-name") = "Default";
  condition.append_attribute("calcext:value") = value.c_str();
  ods_range first = {base.r1, base.c1, base.r1, base.c1};
  condition.append_attribute("calcext:base-cell-address") = ods_range_address(sheet, first).c_str();
}

// preset shapes, the shape types of ODF consumers and the outlines of the presets in a 21600 x 21600 box
static const char* const ods_shapes[][3] = {
  {"rect", "rectangle",
   "M 0 0 L 21600 0 21600 21600 0 21600 Z N"},
  {"roundRect", "round-rectangle",
   "M 3600 0 L 18000 0 X 21600 3600 L 21600 18000 Y 18000 21600 L 3600 21600 X 0 18000 L 0 3600 Y 3600 0 Z N"},
  {"ellipse", "ellipse",
   "M 10800 0 X 21600 10800 Y 10800 21600 X 0 10800 Y 10800 0 Z N"},
  {"triangle", "isosceles-triangle",
   "M 10800 0 L 21600 21600 0 21600 Z N"},
  {"rtTriangle", "right-triangle",
   "M 0 0 L 21600 21600 0 21600 Z N"},
  {"diamond", "diamond",
   "M 10800 0 L 21600 10800 10800 21600 0 10800 Z N"},
  {"parallelogram", "parallelogram",
   "M 5400 0 L 21600 0 16200 21600 0 21600 Z N"},
  {"trapezoid", "trapezoid",
   "M 0 21600 L 5400 0 16200 0 21600 21600 Z N"},
  {"pentagon", "pentagon",
   "M 10800 0 L 21600 8250 17475 21600 4125 21600 0 8250 Z N"},
  {"hexagon", "hexagon",
   "M 5400 0 L 16200 0 21600 10800 16200 21600 5400 21600 0 10800 Z N"},
  {"octagon", "octagon",
   "M 6326 0 L 15274 0 21600 6326 21600 15274 15274 21600 6326 21600 0 15274 0 6326 Z N"},
  {"star4", "star4",
   "M 10800 0 L 12709 8891 21600 10800 12709 12709 10800 21600 8891 12709 0 10800 8891 8891 Z N"},
  {"star5", "star5",
   "M 10800 0 L 13225 8250 21071 8250 14724 13350 17148 21600 10800 16501 4452 21600 6876 13350 529 8250 "
   "8375 8250 Z N"},
  {"star8", "star8",
   "M 10800 0 L 13962 3167 18437 3163 18433 7638 21600 10800 18433 13962 18437 18437 13962 18433 10800 "
   "21600 7638 18433 3163 18437 3167 13962 0 10800 3167 7638 3163 3163 7638 3167 Z N"},
  {"star24", "star24",
   "M 10800 0 L 11857 2769 13595 368 13900 3317 16200 1447 15731 4374 18437 3163 17226 5869 20153 5400 "
   "18283 7700 21232 8005 18831 9743 21600 10800 18831 11857 21232 13595 18283 13900 20153 16200 17226 "
   "15731 18437 18437 15731 17226 16200 20153 13900 18283 13595 21232 11857 18831 10800 21600 9743 18831 "
   "8005 21232 7700 18283 5400 20153 5869 17226 3163 18437 4374 15731 1447 16200 3317 13900 368 13595 "
   "2769 11857 0 10800 2769 9743 368 8005 3317 7700 1447 5400 4374 5869 3163 3163 5869 4374 5400 1447 "
   "7700 3317 8005 368 9743 2769 Z N"},
  {"rightArrow", "right-arrow",
   "M 0 5400 L 10800 5400 10800 0 21600 10800 10800 21600 10800 16200 0 16200 Z N"},
  {"leftArrow", "left-arrow",
   "M 21600 5400 L 10800 5400 10800 0 0 10800 10800 21600 10800 16200 21600 16200 Z N"},
  {"upArrow", "up-arrow",
   "M 5400 21600 L 5400 10800 0 10800 10800 0 21600 10800 16200 10800 16200 21600 Z N"},
  {"downArrow", "down-arrow",
   "M 5400 0 L 5400 10800 0 10800 10800 21600 21600 10800 16200 10800 16200 0 Z N"},
  {"leftRightArrow", "left-right-arrow",
   "M 0 10800 L 5400 0 5400 5400 16200 5400 16200 0 21600 10800 16200 21600 16200 16200 5400 16200 5400 21600 Z N"},
  {"upDownArrow", "up-down-arrow",
   "M 10800 0 L 21600 5400 16200 5400 16200 16200 21600 16200 10800 21600 0 16200 5400 16200 5400 5400 0 5400 Z N"},
  {"can", "can",
   "M 0 2700 L 0 18900 Y 10800 21600 X 21600 18900 L 21600 2700 Y 10800 0 X 0 2700 Z N M 0 2700 Y 10800 "
   "5400 X 21600 2700 F N"},
  {"cube", "cube",
   "M 0 5400 L 5400 0 21600 0 21600 16200 16200 21600 0 21600 Z N M 0 5400 L 16200 5400 21600 0 F N M "
   "16200 5400 L 16200 21600 F N"},
  {"heart", "heart",
   "M 10800 5400 C 12420 0 21600 0 21600 6480 C 21600 12960 13500 16200 10800 21600 C 8100 16200 0 12960 "
   "0 6480 C 0 0 9180 0 10800 5400 Z N"},
  {"smileyFace", "smiley",
   "M 10800 0 X 21600 10800 Y 10800 21600 X 0 10800 Y 10800 0 Z N M 7290 6390 X 8460 7560 Y 7290 8730 X "
   "6120 7560 Y 7290 6390 Z N M 14310 6390 X 15480 7560 Y 14310 8730 X 13140 7560 Y 14310 6390 Z N M "
   "4860 15480 C 8640 17520 12960 17520 16740 15480 F N"},
  {"sun", "sun",
   "M 10800 5400 X 16200 10800 Y 10800 16200 X 5400 10800 Y 10800 5400 Z N M 17911 9674 L 21600 10800 "
   "17911 11926 Z N M 16625 15032 L 18437 18437 15032 16625 Z N M 11926 17911 L 10800 21600 9674 17911 Z "
   "N M 6568 16625 L 3163 18437 4975 15032 Z N M 3689 11926 L 0 10800 3689 9674 Z N M 4975 6568 L 3163 "
   "3163 6568 4975 Z N M 9674 3689 L 10800 0 11926 3689 Z N M 15032 4975 L 18437 3163 16625 6568 Z N"},
  {"moon", "moon",
   "M 21600 0 X 0 10800 Y 21600 21600 X 10800 10800 Y 21600 0 Z N"},
  {"lightningBolt", "lightning",
   "M 8640 0 L 12960 6480 11340 7020 16740 12150 15390 12960 21600 21600 10800 14580 12420 13770 5400 "
   "9450 7020 8640 0 3240 Z N"},
  {"donut", "ring",
   "M 10800 0 X 21600 10800 Y 10800 21600 X 0 10800 Y 10800 0 Z M 10800 5400 X 5400 10800 Y 10800 16200 "
   "X 16200 10800 Y 10800 5400 Z N"},
  {"noSmoking", "forbidden",
   "M 10800 0 X 21600 10800 Y 10800 21600 X 0 10800 Y 10800 0 Z M 10800 4050 X 4050 10800 Y 10800 17550 "
   "X 17550 10800 Y 10800 4050 Z N M 4815 7679 L 13921 16785 16785 13921 7679 4815 Z N"},
  {"plus", "cross",
   "M 5400 0 L 16200 0 16200 5400 21600 5400 21600 16200 16200 16200 16200 21600 5400 21600 5400 16200 0 "
   "16200 0 5400 5400 5400 Z N"},
  {"chevron", "chevron",
   "M 0 0 L 10800 0 21600 10800 10800 21600 0 21600 10800 10800 Z N"},
  {"homePlate", "pentagon-right",
   "M 0 0 L 10800 0 21600 10800 10800 21600 0 21600 Z N"},
  {"blockArc", "block-arc",
   "M 0 10800 Y 10800 0 X 21600 10800 L 16200 10800 Y 10800 5400 X 5400 10800 Z N"},
  {"flowChartProcess", "flowchart-process",
   "M 0 0 L 21600 0 21600 21600 0 21600 Z N"},
  {"flowChartDecision", "flowchart-decision",
   "M 10800 0 L 21600 10800 10800 21600 0 10800 Z N"},
  {"flowChartTerminator", "flowchart-terminator",
   "M 3475 0 L 18125 0 X 21600 10800 Y 18125 21600 L 3475 21600 X 0 10800 Y 3475 0 Z N"},
  {"wedgeRectCallout", "rectangular-callout",
   "M 0 0 L 21600 0 21600 21600 9000 21600 6300 24300 3600 21600 0 21600 Z N"},
  {"wedgeRoundRectCallout", "round-rectangular-callout",
   "M 3600 0 L 18000 0 X 21600 3600 L 21600 18000 Y 18000 21600 L 9000 21600 6300 24300 3600 21600 X 0 "
   "18000 L 0 3600 Y 3600 0 Z N"},
  {"wedgeEllipseCallout", "round-callout",
   "M 10800 0 X 21600 10800 Y 10800 21600 C 9309 21600 7888 21298 6596 20751 L 6300 24300 1618 16489 C "
   "592 14837 0 12888 0 10800 Y 10800 0 Z N"},
  {"cloudCallout", "cloud-callout",
   "M 10800 3400 C 14290 1385 13831 1244 15385 4813 C 19478 5123 19195 4753 18218 8513 C 21352 11029 "
   "21352 10571 18218 13087 C 19195 16847 19478 16477 15385 16787 C 13831 20356 14290 20215 10800 18200 "
   "C 7310 20215 7769 20356 6215 16787 C 2122 16477 2405 16847 3382 13087 C 248 10571 248 11029 3382 "
   "8513 C 2405 4753 2122 5123 6215 4813 C 7769 1244 7310 1385 10800 3400 Z N M 7600 20000 X 8500 20900 "
   "Y 7600 21800 X 6700 20900 Y 7600 20000 Z N M 6900 22000 X 7500 22600 Y 6900 23200 X 6300 22600 Y "
   "6900 22000 Z N M 6300 23550 X 6650 23900 Y 6300 24250 X 5950 23900 Y 6300 23550 Z N"},
  {"line", "mso-spt20", "M 0 0 L 21600 21600 F N"},
  {"straightConnector1", "mso-spt32", "M 0 0 L 21600 21600 F N"},
  {"bentConnector2", "mso-spt33", "M 0 0 L 21600 0 21600 21600 F N"},
  {"bentConnector3", "mso-spt34", "M 0 0 L 10800 0 10800 21600 21600 21600 F N"},
  {"curvedConnector3", "mso-spt38", "M 0 0 C 5400 0 10800 5400 10800 10800 C 10800 16200 16200 21600 21600 21600 F N"}
};

// The custom geometry of a DrawingML shape as enhanced path. The guides are evaluated for the extent of
// the shape (cx, cy in EMU), the path is in hundredths of a point and arcs are cubic Bezier curves.
static bool ods_custom_path(pugi::xml_node geom, double cx, double cy, std::string& view, std::string& path) {
  const double pi = 3.14159265358979323846, deg = 60000.0;
  if (!(cx > 0.0) || !(cy > 0.0)) return false;
  std::map<std::string, double> guides = {
    {"w", cx}, {"h", cy}, {"l", 0.0}, {"t", 0.0}, {"r", cx}, {"b", cy}, {"hc", cx / 2.0}, {"vc", cy / 2.0},
    {"ss", std::min(cx, cy)}, {"ls", std::max(cx, cy)}, {"cd2", 10800000.0}, {"cd4", 5400000.0}, {"cd8", 2700000.0},
    {"3cd4", 16200000.0}, {"3cd8", 8100000.0}, {"5cd8", 13500000.0}, {"7cd8", 18900000.0}
  };
  for (int32_t d : {2, 3, 4, 5, 6, 8, 10, 12, 16, 32}) {
    guides["wd" + std::to_string(d)] = cx / d;
    guides["hd" + std::to_string(d)] = cy / d;
    guides["ssd" + std::to_string(d)] = std::min(cx, cy) / d;
  }
  guides["3wd4"] = cx * 3.0 / 4.0;
  guides["3hd4"] = cy * 3.0 / 4.0;
  guides["wd3x2"] = cx * 2.0 / 3.0;
  auto value = [&](const std::string& x) {
    std::map<std::string, double>::const_iterator g = guides.find(x);
    return g != guides.end() ? g->second : std::strtod(x.c_str(), nullptr);
  };
  // the formulas of the adjust values and guides: an operator with up to three arguments
  for (pugi::xml_node list : {geom.child("a:avLst"), geom.child("a:gdLst")}) {
    for (pugi::xml_node gd : list.children("a:gd")) {
      std::vector<std::string> parts;
      std::string fmla = gd.attribute("fmla").as_string();
      for (size_t pos = 0; pos < fmla.size();) {
        size_t end = fmla.find(' ', pos);
        if (end == std::string::npos) end = fmla.size();
        if (end > pos) parts.push_back(fmla.substr(pos, end - pos));
        pos = end + 1;
      }
      if (parts.empty()) continue;
      double x = parts.size() > 1 ? value(parts[1]) : 0.0, y = parts.size() > 2 ? value(parts[2]) : 0.0, z = parts.size() > 3 ? value(parts[3]) : 0.0;
      const std::string& op = parts[0];
      double v = 0.0;
      if (op == "val") v = x;
      else if (op == "*/") v = z != 0.0 ? x * y / z : 0.0;
      else if (op == "+-") v = x + y - z;
      else if (op == "+/") v = z != 0.0 ? (x + y) / z : 0.0;
      else if (op == "?:") v = x > 0.0 ? y : z;
      else if (op == "abs") v = std::fabs(x);
      else if (op == "at2") v = std::atan2(y, x) / pi * 180.0 * deg;
      else if (op == "cat2") v = x * std::cos(std::atan2(z, y));
      else if (op == "sat2") v = x * std::sin(std::atan2(z, y));
      else if (op == "cos") v = x * std::cos(y / deg / 180.0 * pi);
      else if (op == "sin") v = x * std::sin(y / deg / 180.0 * pi);
      else if (op == "tan") v = x * std::tan(y / deg / 180.0 * pi);
      else if (op == "max") v = std::max(x, y);
      else if (op == "min") v = std::min(x, y);
      else if (op == "mod") v = std::sqrt(x * x + y * y + z * z);
      else if (op == "pin") v = y < x ? x : (y > z ? z : y);
      else if (op == "sqrt") v = std::sqrt(std::max(x, 0.0));
      guides[gd.attribute("name").as_string()] = v;
    }
  }

  // coordinates in hundredths of a point
  auto num = [](double v) { return std::to_string(static_cast<int64_t>(std::llround(v / 127.0))); };
  std::string out;
  for (pugi::xml_node p : geom.child("a:pathLst").children("a:path")) {
    double pw = p.attribute("w").as_double(0.0), ph = p.attribute("h").as_double(0.0);
    double sx = pw > 0.0 ? cx / pw : 1.0, sy = ph > 0.0 ? cy / ph : 1.0;
    double px = 0.0, py = 0.0;
    auto point = [&](pugi::xml_node pt, double& x, double& y) {
      x = value(pt.attribute("x").as_string("0")) * sx;
      y = value(pt.attribute("y").as_string("0")) * sy;
    };
    for (pugi::xml_node cmd : p.children()) {
      std::string name = cmd.name();
      if (name == "a:moveTo" || name == "a:lnTo") {
        point(cmd.child("a:pt"), px, py);
        out += (name == "a:moveTo" ? " M " : " L ") + num(px) + " " + num(py);
      } else if (name == "a:cubicBezTo" || name == "a:quadBezTo") {
        std::vector<double> xs, ys;
        for (pugi::xml_node pt : cmd.children("a:pt")) {
          double x = 0.0, y = 0.0;
          point(pt, x, y);
          xs.push_back(x);
          ys.push_back(y);
        }
        // a quadratic curve as cubic curve
        if (name == "a:quadBezTo" && xs.size() == 2) {
          xs = {px + 2.0 / 3.0 * (xs[0] - px), xs[1] + 2.0 / 3.0 * (xs[0] - xs[1]), xs[1]};
          ys = {py + 2.0 / 3.0 * (ys[0] - py), ys[1] + 2.0 / 3.0 * (ys[0] - ys[1]), ys[1]};
        }
        if (xs.size() != 3) continue;
        out += " C";
        for (size_t k = 0; k < 3; ++k) out += " " + num(xs[k]) + " " + num(ys[k]);
        px = xs[2];
        py = ys[2];
      } else if (name == "a:arcTo") {
        // the angles are visual angles of the ellipse, the curve uses the angles of its parametric form
        double rx = value(cmd.attribute("wR").as_string("0")) * sx, ry = value(cmd.attribute("hR").as_string("0")) * sy;
        double start = value(cmd.attribute("stAng").as_string("0")) / deg / 180.0 * pi;
        double sweep = value(cmd.attribute("swAng").as_string("0")) / deg / 180.0 * pi;
        if (!(rx > 0.0) || !(ry > 0.0) || sweep == 0.0) continue;
        auto param = [&](double a) { return std::atan2(rx * std::sin(a), ry * std::cos(a)); };
        double t0 = param(start);
        double ccx = px - rx * std::cos(t0), ccy = py - ry * std::sin(t0);
        int32_t parts = static_cast<int32_t>(std::ceil(std::fabs(sweep) / (pi / 2.0) - 1e-9));
        for (int32_t k = 0; k < parts; ++k) {
          double a1 = start + sweep * (k + 1) / parts;
          double t1 = param(a1);
          double diff = sweep * 1.0 / parts;
          while (t1 - t0 > pi) t1 -= 2.0 * pi;
          while (t1 - t0 < -pi) t1 += 2.0 * pi;
          if (std::fabs(t1 - t0) < 1e-12 && std::fabs(diff) > 1e-9) t1 += diff > 0.0 ? 2.0 * pi : -2.0 * pi;
          double kk = 4.0 / 3.0 * std::tan((t1 - t0) / 4.0);
          double x0 = ccx + rx * std::cos(t0), y0 = ccy + ry * std::sin(t0);
          double x3 = ccx + rx * std::cos(t1), y3 = ccy + ry * std::sin(t1);
          double x1 = x0 - kk * rx * std::sin(t0), y1 = y0 + kk * ry * std::cos(t0);
          double x2 = x3 + kk * rx * std::sin(t1), y2 = y3 - kk * ry * std::cos(t1);
          out += " C " + num(x1) + " " + num(y1) + " " + num(x2) + " " + num(y2) + " " + num(x3) + " " + num(y3);
          px = x3;
          py = y3;
          t0 = t1;
        }
      } else if (name == "a:close") {
        out += " Z";
      }
    }
    // paths without fill or outline
    if (std::strcmp(p.attribute("fill").as_string(), "none") == 0) out += " F";
    if (!p.attribute("stroke").as_bool(true)) out += " S";
    out += " N";
  }
  if (out.empty()) return false;
  path = out.substr(1);
  view = "0 0 " + num(cx) + " " + num(cy);
  return true;
}

struct ods_frame {
  int32_t row = 0, col = 0, group = 0, end_row = 0;
  double x = 0, y = 0, width = 0, height = 0, rotation = 0;
  // the cropped part of a picture on each side (left, top, right, bottom) in 1/100000
  double crop[4] = {0, 0, 0, 0};
  std::string end_cell, name, href, target, ranges;
  double end_x = 0, end_y = 0;
  // the stacking order on the sheet
  int32_t z_index = 0;
  bool chart = false;
  // a straight line from the top left to the bottom right corner, mirrored by the flips
  bool line = false, flip_h = false, flip_v = false;
  std::string shape, style, paragraph_style;
  // the custom geometry of a shape
  std::string geometry;
  // the paragraphs of the shape text as runs of text and text style
  std::vector<std::vector<std::pair<std::string, std::string>>> paragraphs;
};

static void ods_config_item(pugi::xml_node parent, const char* name, const char* type, const std::string& value) {
  pugi::xml_node item = parent.append_child("config:config-item");
  item.append_attribute("config:name") = name;
  item.append_attribute("config:type") = type;
  item.text().set(value.c_str());
}

// a frame or shape. A rotated frame has no position, its transformation
// rotates it around its center and moves it to its position.
static void ods_append_frame(pugi::xml_node parent, const ods_frame& fr) {
  if (fr.line) {
    // the end points rotated around the center of the line
    pugi::xml_node line = parent.append_child("draw:line");
    line.append_attribute("draw:z-index") = fr.z_index;
    if (!fr.name.empty()) line.append_attribute("draw:name") = fr.name.c_str();
    if (!fr.style.empty()) line.append_attribute("draw:style-name") = fr.style.c_str();
    double a = fr.rotation * 3.14159265358979323846 / 180.0;
    double cx = fr.x + fr.width / 2.0, cy = fr.y + fr.height / 2.0;
    double dx = (fr.flip_h ? 1.0 : -1.0) * fr.width / 2.0, dy = (fr.flip_v ? 1.0 : -1.0) * fr.height / 2.0;
    line.append_attribute("svg:x1") = ods_cm(cx + dx * std::cos(a) - dy * std::sin(a)).c_str();
    line.append_attribute("svg:y1") = ods_cm(cy + dx * std::sin(a) + dy * std::cos(a)).c_str();
    line.append_attribute("svg:x2") = ods_cm(cx - dx * std::cos(a) + dy * std::sin(a)).c_str();
    line.append_attribute("svg:y2") = ods_cm(cy - dx * std::sin(a) - dy * std::cos(a)).c_str();
    if (!fr.end_cell.empty()) {
      line.append_attribute("table:end-cell-address") = fr.end_cell.c_str();
      line.append_attribute("table:end-x") = ods_cm(fr.end_x).c_str();
      line.append_attribute("table:end-y") = ods_cm(fr.end_y).c_str();
    }
    return;
  }
  pugi::xml_node frame = parent.append_child("draw:frame");
  frame.append_attribute("draw:z-index") = fr.z_index;
  if (!fr.name.empty()) frame.append_attribute("draw:name") = fr.name.c_str();
  frame.append_attribute("svg:width") = ods_cm(fr.width).c_str();
  frame.append_attribute("svg:height") = ods_cm(fr.height).c_str();
  if (fr.rotation != 0.0) {
    double a = -fr.rotation * 3.14159265358979323846 / 180.0;
    double tx = fr.x + fr.width / 2.0 - (fr.width / 2.0 * std::cos(a) + fr.height / 2.0 * std::sin(a));
    double ty = fr.y + fr.height / 2.0 - (-fr.width / 2.0 * std::sin(a) + fr.height / 2.0 * std::cos(a));
    char buf[160];
    std::snprintf(buf, sizeof(buf), "rotate (%.10g) translate (%s %s)", a, ods_cm(tx).c_str(), ods_cm(ty).c_str());
    frame.append_attribute("draw:transform") = buf;
  } else {
    frame.append_attribute("svg:x") = ods_cm(fr.x).c_str();
    frame.append_attribute("svg:y") = ods_cm(fr.y).c_str();
    if (!fr.end_cell.empty()) {
      frame.append_attribute("table:end-cell-address") = fr.end_cell.c_str();
      frame.append_attribute("table:end-x") = ods_cm(fr.end_x).c_str();
      frame.append_attribute("table:end-y") = ods_cm(fr.end_y).c_str();
    }
  }
  if (!fr.style.empty()) frame.append_attribute("draw:style-name") = fr.style.c_str();
  if (!fr.shape.empty()) {
    frame.set_name("draw:custom-shape");
    for (const auto& paragraph : fr.paragraphs) {
      pugi::xml_node last = frame.last_child();
      for (size_t k = 0; k < paragraph.size(); ++k) ods_cell_text(frame, paragraph[k].first, paragraph[k].second, k > 0);
      if (paragraph.empty()) frame.append_child("text:p");
      for (pugi::xml_node p = last ? last.next_sibling() : frame.first_child(); p; p = p.next_sibling()) {
        if (!fr.paragraph_style.empty() && std::strcmp(p.name(), "text:p") == 0) p.append_attribute("text:style-name") = fr.paragraph_style.c_str();
      }
    }
    pugi::xml_node geometry = frame.append_child("draw:enhanced-geometry");
    if (fr.flip_h) geometry.append_attribute("draw:mirror-horizontal") = "true";
    if (fr.flip_v) geometry.append_attribute("draw:mirror-vertical") = "true";
    // the outline of a custom geometry for the size of the shape
    pugi::xml_document geometry_doc;
    std::string view, path;
    if (!fr.geometry.empty() && geometry_doc.load_string(fr.geometry.c_str()) &&
        ods_custom_path(geometry_doc.first_child(), fr.width * ods_emu_per_pt, fr.height * ods_emu_per_pt, view, path)) {
      geometry.append_attribute("svg:viewBox") = view.c_str();
      geometry.append_attribute("draw:type") = "non-primitive";
      geometry.append_attribute("draw:enhanced-path") = path.c_str();
      return;
    }
    geometry.append_attribute("svg:viewBox") = "0 0 21600 21600";
    // the preset of the shape with its outline, or the outline of a rectangle
    geometry.append_attribute("draw:type") = ("ooxml-" + fr.shape).c_str();
    const char* outline = ods_shapes[0][2];
    for (const auto& shape : ods_shapes) {
      if (fr.shape == shape[0]) outline = shape[2];
    }
    geometry.append_attribute("draw:enhanced-path") = outline;
    return;
  }
  pugi::xml_node content = frame.append_child(fr.chart ? "draw:object" : "draw:image");
  if (!fr.ranges.empty()) content.append_attribute("draw:notify-on-update-of-ranges") = fr.ranges.c_str();
  content.append_attribute("xlink:href") = fr.chart ? ("./" + fr.href).c_str() : fr.href.c_str();
  content.append_attribute("xlink:type") = "simple";
  content.append_attribute("xlink:show") = "embed";
  content.append_attribute("xlink:actuate") = "onLoad";
}

struct ods_out_cell {
  int32_t row;
  int32_t col;
  int32_t style;
  size_t idx;
};

struct ods_out_col {
  int32_t min;
  int32_t max;
  std::string width;
  bool hidden;
  int32_t style;
  int32_t level;
  bool collapsed = false;
};

struct ods_out_row {
  std::string height;
  bool custom;
  bool hidden;
  int32_t style;
  int32_t level;
  bool collapsed = false;
};

struct ods_merge {
  int32_t r1, c1, r2, c2;
};

// [[Rcpp::export]]
std::vector<utf8_string> write_ods_parts(
    native_path dir,
    std::vector<utf8_string> sheet_names,
    Rcpp::LogicalVector hidden,
    Rcpp::List sheets,
    std::vector<utf8_string> shared_strings,
    std::vector<utf8_string> xfs,
    std::vector<utf8_string> fonts,
    std::vector<utf8_string> fills,
    std::vector<utf8_string> borders,
    std::vector<utf8_string> numfmts,
    std::vector<utf8_string> indexed_colors,
    utf8_string theme,
    utf8_string workbook_pr,
    std::vector<utf8_string> defined_names,
    std::vector<int32_t> sheet_index,
    std::vector<utf8_string> table_xml,
    std::vector<utf8_string> table_sheets,
    Rcpp::List book
) {
  std::string path = std::string(dir) + "/";
  std::vector<utf8_string> charts = Rcpp::as<std::vector<utf8_string>>(book["charts"]);
  // the differential formats of the workbook and those of conditional formats of the extensions
  size_t n_dxfs = Rcpp::as<std::vector<utf8_string>>(book["dxfs"]).size();
  std::vector<std::string> extra_dxfs;

  // the values of cells for the cached data of charts: a range in xlsx syntax with its sheet
  auto cell_values = [&](const std::string& f) {
    ods_values values;
    std::string ref = f;
    if (ref.size() > 2 && ref.front() == '(' && ref.back() == ')') ref = ref.substr(1, ref.size() - 2);
    std::string sheet1, sheet2;
    size_t j = ods_parse_sheet(ref, 0, sheet1, sheet2);
    ods_ref_part p1, p2;
    size_t k = ods_parse_ref_part(ref, j, p1);
    if (j == 0 || k == j || p1.row <= 0 || p1.col <= 0) return values;
    p2 = p1;
    if (k < ref.size() && ref[k] == ':') ods_parse_ref_part(ref, k + 1, p2);
    std::vector<utf8_string>::const_iterator pos = std::find(sheet_names.begin(), sheet_names.end(), sheet1);
    if (pos == sheet_names.end() || p2.row < p1.row || p2.col < p1.col || (p2.row - p1.row + 1) * (p2.col - p1.col + 1) > 65536) return values;
    SEXP one_sexp = sheets[pos - sheet_names.begin()];
    Rcpp::List one(one_sexp);
    SEXP cells_sexp = one["cc"];
    if (Rf_isNull(cells_sexp)) return values;
    Rcpp::List cells(cells_sexp);
    std::vector<std::string> cr = Rcpp::as<std::vector<std::string>>(cells["r"]), ct = Rcpp::as<std::vector<std::string>>(cells["c_t"]);
    std::vector<utf8_string> cv = Rcpp::as<std::vector<utf8_string>>(cells["v"]), cis = Rcpp::as<std::vector<utf8_string>>(cells["is"]);
    std::map<std::pair<int64_t, int64_t>, std::pair<bool, std::string>> found;
    for (size_t i = 0; i < cr.size(); ++i) {
      int64_t r = cell_to_rowint(cr[i]), c = cell_to_colint(cr[i]);
      if (r < p1.row || r > p2.row || c < p1.col || c > p2.col) continue;
      std::pair<bool, std::string> value(true, cv[i]);
      if (ct[i] == "s") {
        size_t si = static_cast<size_t>(std::strtoul(cv[i].c_str(), nullptr, 10));
        value = std::make_pair(false, si < shared_strings.size() ? ods_plain_text(shared_strings[si]) : std::string());
      } else if (ct[i] == "inlineStr") {
        value = std::make_pair(false, ods_plain_text(cis[i]));
      } else if (ct[i] == "str" || ct[i] == "e" || ct[i] == "b") {
        value.first = false;
      }
      found[std::make_pair(r, c)] = value;
    }
    for (int64_t r = p1.row; r <= p2.row; ++r) {
      for (int64_t c = p1.col; c <= p2.col; ++c) {
        auto it = found.find(std::make_pair(r, c));
        values.push_back(it == found.end() ? std::make_pair(true, std::string()) : it->second);
      }
    }
    return values;
  };
  // the ODF version of all parts
  std::string version = Rcpp::as<utf8_string>(book["version"]);

  std::vector<ods_table> tables_info;
  for (size_t i = 0; i < table_xml.size() && i < table_sheets.size(); ++i) {
    pugi::xml_document tdoc;
    if (!tdoc.load_string(table_xml[i].c_str())) continue;
    pugi::xml_node tn = tdoc.child("table");
    std::string ref = tn.attribute("ref").as_string();
    size_t colon = ref.find(':');
    if (ref.empty() || colon == std::string::npos) continue;
    ods_table t;
    t.name = tn.attribute("displayName").as_string(tn.attribute("name").as_string());
    // the quotes of a sheet name are doubled inside a quoted reference
    for (char c : std::string(table_sheets[i])) {
      t.sheet += c;
      if (c == '\'') t.sheet += c;
    }
    t.r1 = cell_to_rowint(ref.substr(0, colon));
    t.c1 = cell_to_colint(ref.substr(0, colon));
    t.r2 = cell_to_rowint(ref.substr(colon + 1));
    t.c2 = cell_to_colint(ref.substr(colon + 1));
    t.headers = tn.attribute("headerRowCount").as_llong(1);
    t.totals = tn.attribute("totalsRowCount").as_llong(0);
    t.filter = tn.child("autoFilter");
    for (pugi::xml_node tc : tn.child("tableColumns").children("tableColumn")) t.columns.push_back(tc.attribute("name").as_string());
    tables_info.push_back(t);
  }

  pugi::xml_document xf_doc, font_doc, fill_doc, border_doc, numfmt_doc, palette_doc, theme_doc, pr_doc;
  std::vector<pugi::xml_node> xf_nodes = ods_load_list(xf_doc, xfs);
  std::vector<pugi::xml_node> font_nodes = ods_load_list(font_doc, fonts);
  std::vector<pugi::xml_node> fill_nodes = ods_load_list(fill_doc, fills);
  std::vector<pugi::xml_node> border_nodes = ods_load_list(border_doc, borders);
  std::vector<pugi::xml_node> numfmt_nodes = ods_load_list(numfmt_doc, numfmts);

  // the width of a column unit in points from the font of the default style
  pugi::xml_node default_font;
  if (!xf_nodes.empty() && xf_nodes[0].attribute("fontId").as_uint() < font_nodes.size()) default_font = font_nodes[xf_nodes[0].attribute("fontId").as_uint()];
  double digit_px = ods_digit_width(default_font.child("name").attribute("val").as_string("Calibri"), default_font.child("sz").attribute("val").as_double(11.0));
  double unit_pt = digit_px * 0.75;

  pr_doc.load_string(workbook_pr.c_str());
  bool date1904 = pr_doc.child("workbookPr").attribute("date1904").as_bool();

  std::vector<std::string> theme_colors;
  if (theme_doc.load_string(theme.c_str())) {
    for (pugi::xml_node clr : theme_doc.child("a:theme").child("a:themeElements").child("a:clrScheme").children()) {
      pugi::xml_node val = clr.first_child();
      if (std::strcmp(val.name(), "a:srgbClr") == 0) theme_colors.push_back(val.attribute("val").as_string());
      else if (std::strcmp(val.name(), "a:sysClr") == 0) theme_colors.push_back(val.attribute("lastClr").as_string());
      else theme_colors.push_back("");
    }
  }
  // the colors of the default theme of new workbooks
  if (theme_colors.empty()) {
    theme_colors = {"000000", "FFFFFF", "0E2841", "E8E8E8", "156082", "E97132",
                    "196B24", "0F9ED5", "A02B93", "4EA72E", "467886", "96607D"};
  }
  // the line widths of the line styles of the theme in EMU
  std::vector<double> theme_lines;
  for (pugi::xml_node ln : theme_doc.child("a:theme").child("a:themeElements").child("a:fmtScheme").child("a:lnStyleLst").children("a:ln")) {
    theme_lines.push_back(ln.attribute("w").as_double());
  }
  if (theme_lines.empty()) theme_lines = {12700, 19050, 25400};

  std::vector<std::string> palette;
  for (pugi::xml_node rgb : ods_load_list(palette_doc, indexed_colors)) {
    if (std::strcmp(rgb.name(), "indexedColors") != 0) continue;
    for (pugi::xml_node c : rgb.children("rgbColor")) {
      std::string hex = c.attribute("rgb").as_string();
      palette.push_back(hex.size() == 8 ? hex.substr(2) : hex);
    }
  }
  if (palette.empty()) {
    char buf[16];
    for (uint32_t k = 0; k < 64; ++k) {
      std::snprintf(buf, sizeof(buf), "%06X", static_cast<unsigned int>(XlsDefaultPalette[k < 8 ? k : k - 8]));
      palette.push_back(buf);
    }
  }
  palette.resize(64);
  palette.push_back("000000");
  palette.push_back("ffffff");

  pugi::xml_document empty_doc;
  pugi::xml_node empty = empty_doc.append_child("x");
  size_t n_xfs = xf_nodes.size();

  pugi::xml_document content;
  ods_declaration(content);
  pugi::xml_node root = content.append_child("office:document-content");
  for (const auto& ns : ods_namespaces) root.append_attribute(ns[0]) = ns[1];
  root.append_attribute("office:version") = version.c_str();
  pugi::xml_node font_faces = root.append_child("office:font-face-decls");
  pugi::xml_node autos = root.append_child("office:automatic-styles");
  pugi::xml_node spreadsheet = root.append_child("office:body").append_child("office:spreadsheet");
  // formulas of xlsx compare text without case and with wildcards, not regular expressions, and find no labels
  pugi::xml_node calculation = spreadsheet.append_child("table:calculation-settings");
  calculation.append_attribute("table:case-sensitive") = "false";
  calculation.append_attribute("table:automatic-find-labels") = "false";
  calculation.append_attribute("table:use-regular-expressions") = "false";
  calculation.append_attribute("table:use-wildcards") = "true";
  if (date1904) calculation.append_child("table:null-date").append_attribute("table:date-value") = "1904-01-01";

  pugi::xml_document wbp_doc;
  wbp_doc.load_string(Rcpp::as<utf8_string>(book["workbook_protection"]).c_str());
  if (wbp_doc.child("workbookProtection").attribute("lockStructure").as_bool()) {
    spreadsheet.append_attribute("table:structure-protected") = "true";
  }

  // the view of a sheet: frozen rows and columns, the top left cell of the view and of the frozen pane,
  // the cursor, the zoom of the normal view and of the page break preview
  struct ods_view {
    int64_t rows = 0, cols = 0, zoom = 100, page_zoom = 60;
    int64_t left = 0, top = 0, pane_left = 0, pane_top = 0, cursor_col = 0, cursor_row = 0;
    bool grid = true, zeros = true, headers = true, outline = true, page_break = false;
    // a split view that is not frozen has its split in pixels instead of rows and columns
    bool split = false;
    std::string grid_color;
  };
  struct ods_page {
    double paper_w = 0, paper_h = 0;
    bool landscape = false, fit = false, grid = false, headings = false, comments = false, across = false;
    bool center_h = false, center_v = false;
    double top = 0.75, bottom = 0.75, left = 0.7, right = 0.7, header = 0.3, footer = 0.3;
    int64_t scale = 100, fit_x = 1, fit_y = 1, first_page = 0;
    // the left, center and right parts of the header and footer of odd, even and first pages
    std::vector<std::string> hf;
  };
  std::vector<ods_view> views;
  std::vector<ods_page> pages;
  std::set<std::string> pictures;
  std::vector<std::string> objects;
  int32_t graphic_styles = 0;
  std::string frame_style;

  std::map<uint32_t, std::string> custom_codes;
  for (pugi::xml_node nf : numfmt_nodes) custom_codes[nf.attribute("numFmtId").as_uint()] = nf.attribute("formatCode").as_string();

  std::map<uint32_t, ods_kind> numfmt_kind;
  std::vector<ods_kind> xf_kind(n_xfs, ods_kind_number);
  for (size_t i = 0; i < n_xfs; ++i) {
    uint32_t id = xf_nodes[i].attribute("numFmtId").as_uint();
    if (id == 0) continue;
    if (numfmt_kind.find(id) == numfmt_kind.end()) {
      std::string code;
      if (custom_codes.count(id)) {
        code = custom_codes[id];
      } else if (const char* builtin = xls_builtin_format(id)) {
        code = builtin;
      }
      numfmt_kind[id] = code.empty() ? ods_kind_number : ods_number_format(autos, id, code);
      if (code.empty()) continue;
    }
    xf_kind[i] = numfmt_kind[id];
  }

  for (size_t i = 1; i < n_xfs; ++i) {
    pugi::xml_node xf = xf_nodes[i];
    size_t font_id = xf.attribute("fontId").as_uint(), fill_id = xf.attribute("fillId").as_uint(), border_id = xf.attribute("borderId").as_uint();
    uint32_t numfmt_id = xf.attribute("numFmtId").as_uint();
    pugi::xml_node style = autos.append_child("style:style");
    style.append_attribute("style:name") = ("ce" + std::to_string(i)).c_str();
    style.append_attribute("style:family") = "table-cell";
    style.append_attribute("style:parent-style-name") = "Default";
    if (numfmt_id != 0 && (custom_codes.count(numfmt_id) || xls_builtin_format(numfmt_id))) {
      style.append_attribute("style:data-style-name") = ("N" + std::to_string(numfmt_id)).c_str();
    }
    ods_cell_style(style, xf,
                   font_id < font_nodes.size() ? font_nodes[font_id] : empty,
                   fill_id < fill_nodes.size() ? fill_nodes[fill_id] : empty,
                   border_id < border_nodes.size() ? border_nodes[border_id] : empty,
                   theme_colors, palette);
  }

  std::map<std::string, std::string> col_styles, row_styles;

  // the run properties of rich text as automatic text styles
  std::map<std::string, std::string> text_styles;
  auto text_style = [&](pugi::xml_node rpr) {
    pugi::xml_document doc;
    pugi::xml_node tp = doc.append_child("style:text-properties");
    ods_font_properties(tp, rpr, theme_colors, palette);
    if (!tp.first_attribute()) return std::string();
    // spans refer to their font by a font face
    if (pugi::xml_attribute family = tp.attribute("fo:font-family")) {
      std::string face = family.as_string();
      face.erase(std::remove(face.begin(), face.end(), '\''), face.end());
      bool declared = false;
      for (pugi::xml_node ff : font_faces.children("style:font-face")) {
        if (face == ff.attribute("style:name").as_string()) declared = true;
      }
      if (!declared) {
        pugi::xml_node ff = font_faces.append_child("style:font-face");
        ff.append_attribute("style:name") = face.c_str();
        ff.append_attribute("svg:font-family") = family.as_string();
      }
      tp.remove_attribute(family);
      tp.prepend_attribute("style:font-name") = face.c_str();
    }
    std::string key = ods_xml_string(tp);
    auto known = text_styles.find(key);
    if (known != text_styles.end()) return known->second;
    std::string name = "T" + std::to_string(text_styles.size() + 1);
    pugi::xml_node style = autos.append_child("style:style");
    style.append_attribute("style:name") = name.c_str();
    style.append_attribute("style:family") = "text";
    style.append_copy(tp);
    text_styles[key] = name;
    return name;
  };
  // the alignment of shape text as automatic paragraph styles
  std::map<std::string, std::string> paragraph_styles;
  auto paragraph_style = [&](const char* align) {
    auto known = paragraph_styles.find(align);
    if (known != paragraph_styles.end()) return known->second;
    std::string name = "P" + std::to_string(paragraph_styles.size() + 1);
    pugi::xml_node style = autos.append_child("style:style");
    style.append_attribute("style:name") = name.c_str();
    style.append_attribute("style:family") = "paragraph";
    style.append_child("style:paragraph-properties").append_attribute("fo:text-align") = align;
    paragraph_styles[align] = name;
    return name;
  };
  // a shared or inline string, rich text as spans
  auto string_text = [&](pugi::xml_node cell, const std::string& xml) {
    pugi::xml_document doc;
    if (!doc.load_string(xml.c_str(), pugi::parse_default | pugi::parse_ws_pcdata_single)) return;
    std::vector<std::pair<std::string, pugi::xml_node>> runs = ods_rich_runs(doc.first_child());
    if (runs.empty()) {
      ods_cell_text(cell, ods_plain_text(xml));
      return;
    }
    for (size_t k = 0; k < runs.size(); ++k) ods_cell_text(cell, runs[k].first, text_style(runs[k].second), k > 0);
  };
  std::vector<pugi::xml_node> tables;
  pugi::xml_node validations = spreadsheet.append_child("table:content-validations");
  std::vector<std::pair<std::string, std::string>> autofilters;
  int32_t validation_count = 0;
  int32_t control_count = 0;
  int32_t map_style_count = 0;
  // the gradients of shapes, written to the styles
  std::vector<std::string> gradients;
  pugi::xml_document gradient_doc;
  // the dashes and line ends of shapes written to the styles
  std::set<std::string> drawing_defs;

  // pivot tables are found by the relationships of their sheet, their caches by their own relationships
  pugi::xml_document pivot_doc, pivot_rels_doc, cache_doc, pilots_doc;
  std::vector<pugi::xml_node> pivot_nodes = ods_load_list(pivot_doc, Rcpp::as<std::vector<utf8_string>>(book["pivot_tables"]));
  std::set<std::string> pivot_names;
  for (pugi::xml_node pt : pivot_nodes) pivot_names.insert(pt.attribute("name").as_string());
  std::vector<pugi::xml_node> pivot_rels = ods_load_list(pivot_rels_doc, Rcpp::as<std::vector<utf8_string>>(book["pivot_tables_rels"]));
  std::vector<pugi::xml_node> cache_nodes = ods_load_list(cache_doc, Rcpp::as<std::vector<utf8_string>>(book["pivot_caches"]));
  Rcpp::List pivot_first = book["pivot_first"];
  pugi::xml_node pilots = pilots_doc.append_child("table:data-pilot-tables");

  // print ranges and print titles of the sheets, the parts of a name are separated by commas outside of quotes
  struct ods_print {
    std::vector<ods_range> ranges;
    int64_t r1 = 0, r2 = 0, c1 = 0, c2 = 0;
  };
  std::vector<ods_print> prints(sheet_names.size());
  pugi::xml_document print_names_doc;
  for (pugi::xml_node dn : ods_load_list(print_names_doc, defined_names)) {
    std::string name = dn.attribute("name").as_string();
    if (name != "_xlnm.Print_Area" && name != "_xlnm.Print_Titles") continue;
    auto pos = std::find(sheet_index.begin(), sheet_index.end(), dn.attribute("localSheetId").as_int(-1));
    if (pos == sheet_index.end() || static_cast<size_t>(pos - sheet_index.begin()) >= prints.size()) continue;
    ods_print& pr = prints[static_cast<size_t>(pos - sheet_index.begin())];
    std::string value = dn.text().get();
    size_t start = 0;
    bool quoted = false;
    for (size_t i = 0; i <= value.size(); ++i) {
      if (i < value.size() && value[i] == '\'') quoted = !quoted;
      if (i < value.size() && (quoted || value[i] != ',')) continue;
      std::string part = value.substr(start, i - start), first, last;
      start = i + 1;
      size_t ref_pos = ods_parse_sheet(part, 0, first, last);
      for (const ods_range& rg : ods_sqref(part.substr(ref_pos))) {
        if (name == "_xlnm.Print_Area") {
          pr.ranges.push_back(rg);
        } else if (rg.c1 == 1 && rg.c2 == MAX_OOXML_COL_INT) {
          pr.r1 = rg.r1;
          pr.r2 = rg.r2;
        } else if (rg.r1 == 1 && rg.r2 == MAX_OOXML_ROW_INT) {
          pr.c1 = rg.c1;
          pr.c2 = rg.c2;
        }
      }
    }
  }

  for (R_xlen_t s = 0; s < sheets.size(); ++s) {
    SEXP sheet_sexp = sheets[s];
    Rcpp::List sheet(sheet_sexp);
    pugi::xml_node table = spreadsheet.append_child("table:table");
    tables.push_back(table);
    std::string sheet_quoted;
    for (char c : std::string(sheet_names[static_cast<size_t>(s)])) {
      sheet_quoted += c;
      if (c == '\'') sheet_quoted += c;
    }
    table.append_attribute("table:name") = sheet_names[static_cast<size_t>(s)].c_str();
    const ods_print& pr = prints[static_cast<size_t>(s)];
    if (!pr.ranges.empty()) {
      std::string print_ranges;
      for (const ods_range& rg : pr.ranges) print_ranges += (print_ranges.empty() ? "" : " ") + ods_range_address(sheet_quoted, rg);
      table.append_attribute("table:print-ranges") = print_ranges.c_str();
    }
    std::string ta = "ta" + std::to_string(s + 1);
    pugi::xml_node table_style = autos.append_child("style:style");
    table_style.append_attribute("style:name") = ta.c_str();
    table_style.append_attribute("style:family") = "table";
    table_style.append_attribute("style:master-page-name") = ("mp" + std::to_string(s + 1)).c_str();
    pugi::xml_node table_props = table_style.append_child("style:table-properties");
    table_props.append_attribute("table:display") = hidden[s] ? "false" : "true";
    table.append_attribute("table:style-name") = ta.c_str();

    // the background picture of the sheet is the tiled background of the table
    pugi::xml_document picture_doc, picture_rels_doc;
    picture_doc.load_string(Rcpp::as<utf8_string>(sheet["picture"]).c_str());
    std::string picture_id = picture_doc.child("picture").attribute("r:id").as_string();
    if (!picture_id.empty()) {
      for (pugi::xml_node rel : ods_load_list(picture_rels_doc, Rcpp::as<std::vector<utf8_string>>(sheet["rels"]))) {
        if (picture_id != rel.attribute("Id").as_string()) continue;
        std::string target = rel.attribute("Target").as_string();
        std::string href = "Pictures/" + target.substr(target.rfind('/') + 1);
        std::ifstream exists((path + href).c_str(), std::ios::binary);
        if (!exists) continue;
        pictures.insert(href);
        pugi::xml_node background = table_props.append_child("style:background-image");
        background.append_attribute("xlink:href") = href.c_str();
        background.append_attribute("xlink:type") = "simple";
        background.append_attribute("xlink:actuate") = "onLoad";
        background.append_attribute("style:repeat") = "repeat";
      }
    }

    pugi::xml_document prot_doc;
    prot_doc.load_string(Rcpp::as<utf8_string>(sheet["protection"]).c_str());
    if (prot_doc.child("sheetProtection").attribute("sheet").as_bool()) table.append_attribute("table:protected") = "true";

    // freeze panes, grid lines and zoom
    ods_view view;
    pugi::xml_document pane_doc, view_doc;
    view_doc.load_string(Rcpp::as<utf8_string>(sheet["views"]).c_str());
    pugi::xml_node sheet_view = view_doc.child("sheetViews").child("sheetView");
    std::vector<pugi::xml_node> view_parts = ods_load_list(pane_doc, Rcpp::as<std::vector<utf8_string>>(sheet["pane"]));
    for (pugi::xml_node part : sheet_view.children()) view_parts.push_back(part);
    std::string active_pane;
    for (pugi::xml_node pane : view_parts) {
      if (std::strcmp(pane.name(), "pane") != 0) continue;
      std::string state = pane.attribute("state").as_string("split");
      view.split = state == "split";
      // the split of a view that is not frozen is in twentieths of a point
      double scale = view.split ? 1.0 / 15.0 : 1.0;
      view.cols = static_cast<int64_t>(std::lround(pane.attribute("xSplit").as_double() * scale));
      view.rows = static_cast<int64_t>(std::lround(pane.attribute("ySplit").as_double() * scale));
      active_pane = pane.attribute("activePane").as_string();
      std::string pane_cell = pane.attribute("topLeftCell").as_string();
      view.pane_left = pane_cell.empty() ? (view.split ? 0 : view.cols) : cell_to_colint(pane_cell) - 1;
      view.pane_top = pane_cell.empty() ? (view.split ? 0 : view.rows) : cell_to_rowint(pane_cell) - 1;
    }
    // the active cell of the active pane
    for (pugi::xml_node selection : view_parts) {
      if (std::strcmp(selection.name(), "selection") != 0 || !selection.attribute("activeCell")) continue;
      if (selection.attribute("pane") && active_pane != selection.attribute("pane").as_string()) continue;
      view.cursor_col = cell_to_colint(selection.attribute("activeCell").as_string()) - 1;
      view.cursor_row = cell_to_rowint(selection.attribute("activeCell").as_string()) - 1;
    }
    if (pugi::xml_attribute top_left = sheet_view.attribute("topLeftCell")) {
      view.left = cell_to_colint(top_left.as_string()) - 1;
      view.top = cell_to_rowint(top_left.as_string()) - 1;
    }
    view.grid = sheet_view.attribute("showGridLines").as_bool(true);
    view.zeros = sheet_view.attribute("showZeros").as_bool(true);
    view.headers = sheet_view.attribute("showRowColHeaders").as_bool(true);
    view.outline = sheet_view.attribute("showOutlineSymbols").as_bool(true);
    if (!sheet_view.attribute("defaultGridColor").as_bool(true)) {
      pugi::xml_document grid_doc;
      grid_doc.append_child("color").append_attribute("indexed") = sheet_view.attribute("colorId").as_uint(64);
      view.grid_color = ods_color(grid_doc.child("color"), theme_colors, palette);
    }
    // zoom 0 is the default zoom, page layout views have no counterpart and use the zoom of the normal view
    std::string mode = sheet_view.attribute("view").as_string("normal");
    int64_t zoom = sheet_view.attribute("zoomScale").as_llong(100);
    int64_t zoom_normal = sheet_view.attribute("zoomScaleNormal").as_llong(0);
    int64_t zoom_break = sheet_view.attribute("zoomScaleSheetLayoutView").as_llong(0);
    view.page_break = mode == "pageBreakPreview";
    view.zoom = mode == "normal" ? zoom : zoom_normal;
    view.page_zoom = view.page_break ? zoom : zoom_break;
    if (view.zoom <= 0) view.zoom = 100;
    if (view.page_zoom <= 0) view.page_zoom = 60;
    view.zoom = std::min<int64_t>(std::max<int64_t>(view.zoom, 10), 400);
    view.page_zoom = std::min<int64_t>(std::max<int64_t>(view.page_zoom, 10), 400);
    views.push_back(view);

    // the direction of the sheet and the color of its tab
    if (sheet_view.attribute("rightToLeft").as_bool()) table_props.append_attribute("style:writing-mode") = "rl-tb";
    pugi::xml_document tab_doc;
    tab_doc.load_string(Rcpp::as<utf8_string>(sheet["sheet_pr"]).c_str());
    std::string tab_color = ods_color(tab_doc.child("sheetPr").child("tabColor"), theme_colors, palette);
    if (!tab_color.empty()) table_props.append_attribute("table:tab-color") = tab_color.c_str();

    // page setup
    ods_page page;
    pugi::xml_document setup_doc, margins_doc, sheet_pr_doc, print_doc;
    setup_doc.load_string(Rcpp::as<utf8_string>(sheet["page_setup"]).c_str());
    margins_doc.load_string(Rcpp::as<utf8_string>(sheet["page_margins"]).c_str());
    sheet_pr_doc.load_string(Rcpp::as<utf8_string>(sheet["sheet_pr"]).c_str());
    print_doc.load_string(Rcpp::as<utf8_string>(sheet["print_options"]).c_str());
    pugi::xml_node setup = setup_doc.child("pageSetup");
    uint32_t paper = setup.attribute("paperSize").as_uint(0);
    for (const auto& ps : ods_papers) {
      if (ps.id != paper) continue;
      page.paper_w = ps.width;
      page.paper_h = ps.height;
    }
    page.landscape = std::string(setup.attribute("orientation").as_string()) == "landscape";
    page.scale = setup.attribute("scale").as_llong(100);
    page.fit = sheet_pr_doc.child("sheetPr").child("pageSetUpPr").attribute("fitToPage").as_bool();
    page.fit_x = setup.attribute("fitToWidth").as_llong(1);
    page.fit_y = setup.attribute("fitToHeight").as_llong(1);
    pugi::xml_node margins = margins_doc.child("pageMargins");
    page.top = margins.attribute("top").as_double(0.75);
    page.bottom = margins.attribute("bottom").as_double(0.75);
    page.left = margins.attribute("left").as_double(0.7);
    page.right = margins.attribute("right").as_double(0.7);
    page.header = margins.attribute("header").as_double(0.3);
    page.footer = margins.attribute("footer").as_double(0.3);
    pugi::xml_node print_options = print_doc.child("printOptions");
    page.grid = print_options.attribute("gridLines").as_bool();
    page.headings = print_options.attribute("headings").as_bool();
    page.center_h = print_options.attribute("horizontalCentered").as_bool();
    page.center_v = print_options.attribute("verticalCentered").as_bool();
    std::string cell_comments = setup.attribute("cellComments").as_string("none");
    page.comments = cell_comments == "atEnd" || cell_comments == "asDisplayed";
    page.across = std::string(setup.attribute("pageOrder").as_string()) == "overThenDown";
    if (setup.attribute("useFirstPageNumber").as_bool()) page.first_page = setup.attribute("firstPageNumber").as_llong(1);
    for (const utf8_string& x : Rcpp::as<std::vector<utf8_string>>(sheet["header_footer"])) page.hf.push_back(ods_unescape(x));
    page.hf.resize(18);
    pages.push_back(page);

    // manual page breaks are before the row or column after the break
    std::set<int64_t> row_breaks, col_breaks;
    pugi::xml_document breaks_doc;
    for (pugi::xml_node brk : ods_load_list(breaks_doc, Rcpp::as<std::vector<utf8_string>>(sheet["row_breaks"]))) {
      if (brk.attribute("man").as_bool(true)) row_breaks.insert(brk.attribute("id").as_llong(0) + 1);
    }
    for (pugi::xml_node brk : ods_load_list(breaks_doc, Rcpp::as<std::vector<utf8_string>>(sheet["col_breaks"]))) {
      if (brk.attribute("man").as_bool(true)) col_breaks.insert(brk.attribute("id").as_llong(0) + 1);
    }

    std::vector<utf8_string> cc_r, cc_s, cc_t, cc_v, cc_f, cc_fattr, cc_is;
    SEXP cc_sexp = sheet["cc"];
    if (!Rf_isNull(cc_sexp)) {
      Rcpp::List cc(cc_sexp);
      cc_r = Rcpp::as<std::vector<utf8_string>>(cc["r"]);
      cc_s = Rcpp::as<std::vector<utf8_string>>(cc["c_s"]);
      cc_t = Rcpp::as<std::vector<utf8_string>>(cc["c_t"]);
      cc_v = Rcpp::as<std::vector<utf8_string>>(cc["v"]);
      cc_f = Rcpp::as<std::vector<utf8_string>>(cc["f"]);
      cc_fattr = Rcpp::as<std::vector<utf8_string>>(cc["f_attr"]);
      cc_is = Rcpp::as<std::vector<utf8_string>>(cc["is"]);
    }

    int32_t max_row = 0, max_col = 0;
    std::vector<ods_out_cell> cells;
    for (size_t i = 0; i < cc_r.size(); ++i) {
      ods_out_cell cell;
      cell.row = cell_to_rowint(cc_r[i]);
      cell.col = cell_to_colint(cc_r[i]);
      cell.style = cc_s[i].empty() ? 0 : static_cast<int32_t>(std::strtol(cc_s[i].c_str(), nullptr, 10));
      if (cell.style < 0 || static_cast<size_t>(cell.style) >= n_xfs) cell.style = 0;
      cell.idx = i;
      max_row = std::max(max_row, cell.row);
      max_col = std::max(max_col, cell.col);
      cells.push_back(cell);
    }
    std::sort(cells.begin(), cells.end(), [](const ods_out_cell& a, const ods_out_cell& b) {
      return a.row != b.row ? a.row < b.row : a.col < b.col;
    });

    struct ods_shared {
      int32_t row;
      int32_t col;
      std::string f;
    };
    std::map<std::string, ods_shared> shared;
    for (const ods_out_cell& cell : cells) {
      if (cc_f[cell.idx].empty() || cc_fattr[cell.idx].find("shared") == std::string::npos) continue;
      pugi::xml_document fattr_doc;
      if (!fattr_doc.load_string(("<f " + cc_fattr[cell.idx] + "/>").c_str())) continue;
      pugi::xml_node fa = fattr_doc.child("f");
      if (std::string(fa.attribute("t").as_string()) != "shared") continue;
      ods_shared anchor = {cell.row, cell.col, ods_unescape(cc_f[cell.idx])};
      shared[fa.attribute("si").as_string()] = anchor;
    }

    std::vector<ods_merge> merges;
    pugi::xml_document merge_doc;
    for (pugi::xml_node mc : ods_load_list(merge_doc, Rcpp::as<std::vector<utf8_string>>(sheet["merges"]))) {
      std::string ref = mc.attribute("ref").as_string();
      if (ref.empty()) continue;
      size_t colon = ref.find(':');
      std::string from = ref.substr(0, colon), to = colon == std::string::npos ? from : ref.substr(colon + 1);
      ods_merge m = {cell_to_rowint(from), cell_to_colint(from), cell_to_rowint(to), cell_to_colint(to)};
      merges.push_back(m);
      max_row = std::max(max_row, m.r2);
      max_col = std::max(max_col, m.c2);
    }

    std::map<int32_t, ods_out_row> rows;
    SEXP row_sexp = sheet["row_attr"];
    if (!Rf_isNull(row_sexp)) {
      Rcpp::List ra(row_sexp);
      std::vector<utf8_string> ra_r = Rcpp::as<std::vector<utf8_string>>(ra["r"]);
      std::vector<utf8_string> ra_ht = Rcpp::as<std::vector<utf8_string>>(ra["ht"]);
      std::vector<utf8_string> ra_custom = Rcpp::as<std::vector<utf8_string>>(ra["customHeight"]);
      std::vector<utf8_string> ra_hidden = Rcpp::as<std::vector<utf8_string>>(ra["hidden"]);
      std::vector<utf8_string> ra_s = Rcpp::as<std::vector<utf8_string>>(ra["s"]);
      std::vector<utf8_string> ra_format = Rcpp::as<std::vector<utf8_string>>(ra["customFormat"]);
      std::vector<utf8_string> ra_level = Rcpp::as<std::vector<utf8_string>>(ra["outlineLevel"]);
      std::vector<utf8_string> ra_collapsed = Rcpp::as<std::vector<utf8_string>>(ra["collapsed"]);
      for (size_t i = 0; i < ra_r.size(); ++i) {
        bool row_hidden = ra_hidden[i] == "1" || ra_hidden[i] == "true";
        bool row_format = ra_format[i] == "1" || ra_format[i] == "true";
        int32_t row_style = row_format && !ra_s[i].empty() ? static_cast<int32_t>(std::strtol(ra_s[i].c_str(), nullptr, 10)) : 0;
        if (row_style < 0 || static_cast<size_t>(row_style) >= n_xfs) row_style = 0;
        int32_t row_level = static_cast<int32_t>(std::strtol(ra_level[i].c_str(), nullptr, 10));
        bool row_collapsed = ra_collapsed[i] == "1" || ra_collapsed[i] == "true";
        if (ra_ht[i].empty() && !row_hidden && row_style == 0 && row_level <= 0 && !row_collapsed) continue;
        int32_t r = static_cast<int32_t>(std::strtol(ra_r[i].c_str(), nullptr, 10));
        if (r < 1 || r > MAX_OOXML_ROW_INT) continue;
        // ODF heights are positive, rows without height are hidden
        std::string height = ra_ht[i];
        if (!height.empty() && !(std::strtod(height.c_str(), nullptr) > 0.0)) {
          height.clear();
          row_hidden = true;
        }
        ods_out_row row = {height, ra_custom[i] == "1" || ra_custom[i] == "true", row_hidden, row_style, std::max(row_level, 0), row_collapsed};
        rows[r] = row;
        max_row = std::max(max_row, r);
      }
    }

    pugi::xml_document format_doc;
    format_doc.load_string(Rcpp::as<utf8_string>(sheet["format"]).c_str());
    // without a default width, the base width in digits with padding is rounded up to a multiple of 8 pixels
    double base_width = format_doc.child("sheetFormatPr").attribute("baseColWidth").as_double(8.0);
    if (!(base_width > 0.0)) base_width = 8.0;
    double default_width = format_doc.child("sheetFormatPr").attribute("defaultColWidth").as_double(0.0);
    if (!(default_width > 0.0)) default_width = std::ceil((base_width * digit_px + 5.0) / 8.0) * 8.0 / digit_px;

    std::vector<ods_out_col> cols;
    pugi::xml_document col_doc;
    for (pugi::xml_node col : ods_load_list(col_doc, Rcpp::as<std::vector<utf8_string>>(sheet["cols"]))) {
      ods_out_col oc;
      oc.min = static_cast<int32_t>(col.attribute("min").as_int());
      oc.max = std::min(static_cast<int32_t>(col.attribute("max").as_int()), MAX_OOXML_COL_INT);
      if (oc.min < 1 || oc.max < oc.min) continue;
      oc.width = col.attribute("width").as_string();
      std::string col_hidden = col.attribute("hidden").as_string();
      oc.hidden = col_hidden == "1" || col_hidden == "true";
      oc.style = static_cast<int32_t>(col.attribute("style").as_int());
      oc.level = std::max(static_cast<int32_t>(col.attribute("outlineLevel").as_int()), 0);
      oc.collapsed = col.attribute("collapsed").as_bool();
      if (oc.style < 0 || static_cast<size_t>(oc.style) >= n_xfs) oc.style = 0;
      cols.push_back(oc);
    }
    std::sort(cols.begin(), cols.end(), [](const ods_out_col& a, const ods_out_col& b) { return a.min < b.min; });

    double default_height = format_doc.child("sheetFormatPr").attribute("defaultRowHeight").as_double(15.0);
    // the position in points of the left edge of a column and the top edge of a row (0-based)
    auto col_x = [&](int64_t col0) {
      double x = 0.0;
      for (int64_t k = 1; k <= col0 && k <= MAX_OOXML_COL_INT; ++k) {
        double w = default_width;
        bool col_hidden = false;
        for (const ods_out_col& oc : cols) {
          if (k < oc.min || k > oc.max) continue;
          if (!oc.width.empty()) w = std::strtod(oc.width.c_str(), nullptr);
          col_hidden = oc.hidden;
        }
        if (!col_hidden) x += w * unit_pt;
      }
      return x;
    };
    auto row_y = [&](int64_t row0) {
      double y = 0.0;
      for (int64_t k = 1; k <= row0 && k <= MAX_OOXML_ROW_INT; ++k) {
        auto ra = rows.find(static_cast<int32_t>(k));
        if (ra == rows.end()) y += default_height;
        else if (!ra->second.hidden) y += ra->second.height.empty() ? default_height : std::strtod(ra->second.height.c_str(), nullptr);
      }
      return y;
    };

    // comments: the author is a separate element in ODF, not the first line of the text.
    // The text is a vector of strings with their run properties or of rich text runs.
    struct ods_comment {
      std::string author;
      std::vector<std::pair<std::string, std::string>> runs;
    };
    std::map<std::pair<int32_t, int32_t>, ods_comment> comments;
    SEXP comments_sexp = sheet["comments"];
    if (!Rf_isNull(comments_sexp)) {
      Rcpp::List comment_list(comments_sexp);
      for (R_xlen_t k = 0; k < comment_list.size(); ++k) {
        SEXP one_sexp = comment_list[k];
        Rcpp::List one(one_sexp);
        ods_comment comment;
        std::string ref = Rcpp::as<utf8_string>(one["ref"]);
        comment.author = Rcpp::as<utf8_string>(one["author"]);
        std::vector<utf8_string> texts = Rcpp::as<std::vector<utf8_string>>(one["comment"]);
        std::vector<utf8_string> styles;
        if (one.containsElementNamed("style") && !Rf_isNull(one["style"])) styles = Rcpp::as<std::vector<utf8_string>>(one["style"]);
        size_t first = texts.size() >= 2 && texts[0] == comment.author + ":" && texts[1] == "\n" ? 2 : 0;
        for (size_t t = first; t < texts.size(); ++t) {
          const std::string& text = texts[t];
          pugi::xml_document run_doc;
          bool xml = (text.compare(0, 2, "<r") == 0 || text.compare(0, 2, "<t") == 0) &&
                     run_doc.load_string(("<x>" + text + "</x>").c_str(), pugi::parse_default | pugi::parse_ws_pcdata_single);
          if (!xml) {
            comment.runs.push_back(std::make_pair(text, t < styles.size() ? std::string(styles[t]) : std::string()));
            continue;
          }
          for (pugi::xml_node node : run_doc.child("x").children()) {
            if (std::strcmp(node.name(), "t") == 0) comment.runs.push_back(std::make_pair(std::string(node.text().get()), std::string()));
          }
          for (const auto& run : ods_rich_runs(run_doc.child("x"))) {
            comment.runs.push_back(std::make_pair(run.first, run.second ? ods_xml_string(run.second) : std::string()));
          }
        }
        int32_t cr = cell_to_rowint(ref), cc = cell_to_colint(ref);
        comments[std::make_pair(cr, cc)] = comment;
        max_row = std::max(max_row, cr);
        max_col = std::max(max_col, cc);
      }
    }

    // hyperlinks
    std::map<std::string, std::string> rel_targets;
    pugi::xml_document rels_doc, links_doc;
    for (pugi::xml_node rel : ods_load_list(rels_doc, Rcpp::as<std::vector<utf8_string>>(sheet["rels"]))) {
      rel_targets[rel.attribute("Id").as_string()] = rel.attribute("Target").as_string();
    }
    std::map<std::pair<int32_t, int32_t>, std::pair<std::string, std::string>> links;
    for (pugi::xml_node hl : ods_load_list(links_doc, Rcpp::as<std::vector<utf8_string>>(sheet["hyperlinks"]))) {
      std::string ref = hl.attribute("ref").as_string();
      if (ref.empty()) continue;
      std::string href = rel_targets[hl.attribute("r:id").as_string()];
      std::string location = hl.attribute("location").as_string();
      if (!location.empty()) {
        std::string loc;
        size_t end = ods_xlsx_ref(location, 0, 0, 0, loc);
        if (end == location.size() && loc.size() > 2 && loc[0] == '[') {
          loc = loc.substr(1, loc.size() - 2);
          if (loc[0] == '$') loc = loc.substr(1);
        } else {
          loc = location;
        }
        href += "#" + loc;
      }
      if (href.empty()) continue;
      size_t colon = ref.find(':');
      std::string from = ref.substr(0, colon), to = colon == std::string::npos ? from : ref.substr(colon + 1);
      for (int32_t lr = cell_to_rowint(from); lr <= cell_to_rowint(to); ++lr) {
        for (int32_t lc = cell_to_colint(from); lc <= cell_to_colint(to); ++lc) {
          links[std::make_pair(lr, lc)] = std::make_pair(href, std::string(hl.attribute("tooltip").as_string()));
        }
      }
    }

    // images and charts of the drawing
    std::vector<ods_frame> frames, sheet_frames;
    pugi::xml_document drawing_doc, drawing_rels_doc;
    if (drawing_doc.load_string(Rcpp::as<utf8_string>(sheet["drawing"]).c_str())) {
      std::map<std::string, std::string> targets;
      for (pugi::xml_node rel : ods_load_list(drawing_rels_doc, Rcpp::as<std::vector<utf8_string>>(sheet["drawing_rels"]))) {
        targets[rel.attribute("Id").as_string()] = rel.attribute("Target").as_string();
      }
      // a picture, chart or shape of the drawing; the position is set by the anchor or the group
      auto element_frame = [&](pugi::xml_node el, ods_frame& fr) {
        std::string el_name = el.name();
        if (el_name == "xdr:pic") {
          fr.target = targets[el.child("xdr:blipFill").child("a:blip").attribute("r:embed").as_string()];
          fr.name = el.child("xdr:nvPicPr").child("xdr:cNvPr").attribute("name").as_string();
          fr.href = "Pictures/" + fr.target.substr(fr.target.rfind('/') + 1);
          fr.rotation = el.child("xdr:spPr").child("a:xfrm").attribute("rot").as_double() / 60000.0;
          pugi::xml_node src = el.child("xdr:blipFill").child("a:srcRect");
          static const char* const sides[] = {"l", "t", "r", "b"};
          for (size_t k = 0; k < 4; ++k) fr.crop[k] = src.attribute(sides[k]).as_double();
          return !fr.target.empty();
        }
        if (el_name == "xdr:graphicFrame") {
          fr.target = targets[el.child("a:graphic").child("a:graphicData").child("c:chart").attribute("r:id").as_string()];
          fr.name = el.child("xdr:nvGraphicFramePr").child("xdr:cNvPr").attribute("name").as_string();
          fr.chart = true;
          return !fr.target.empty();
        }
        if (el_name != "xdr:sp" && el_name != "xdr:cxnSp") return false;
        pugi::xml_node sppr = el.child("xdr:spPr"), sp_style = el.child("xdr:style");
        std::string prst = sppr.child("a:prstGeom").attribute("prst").as_string("rect");
        fr.shape = prst;
        if (pugi::xml_node geom = sppr.child("a:custGeom")) fr.geometry = ods_xml_string(geom);
        fr.line = fr.geometry.empty() && (prst == "line" || prst == "straightConnector1");
        fr.name = el.child(el_name == "xdr:sp" ? "xdr:nvSpPr" : "xdr:nvCxnSpPr").child("xdr:cNvPr").attribute("name").as_string();
        fr.rotation = sppr.child("a:xfrm").attribute("rot").as_double() / 60000.0;
        fr.flip_h = sppr.child("a:xfrm").attribute("flipH").as_bool();
        fr.flip_v = sppr.child("a:xfrm").attribute("flipV").as_bool();

        // the text: runs with their text properties, the alignment of the first paragraph; the color
        // of the font of the shape style is the default color of the runs
        pugi::xml_document font_ref_doc;
        font_ref_doc.append_child("f").append_child("a:solidFill").append_copy(el.child("xdr:style").child("a:fontRef").first_child());
        std::string font_color = ods_drawing_color(font_ref_doc.child("f"), theme_colors);
        bool has_text = false;
        for (pugi::xml_node para : el.child("xdr:txBody").children("a:p")) {
          std::vector<std::pair<std::string, std::string>> runs;
          for (pugi::xml_node run : para.children()) {
            std::string run_name = run.name();
            if (run_name == "a:br") {
              runs.push_back(std::make_pair(std::string("\n"), std::string()));
              continue;
            }
            if (run_name != "a:r" && run_name != "a:fld") continue;
            pugi::xml_node rpr = run.child("a:rPr");
            pugi::xml_document run_font_doc;
            pugi::xml_node font = run_font_doc.append_child("rPr");
            if (rpr.attribute("b").as_bool()) font.append_child("b");
            if (rpr.attribute("i").as_bool()) font.append_child("i");
            std::string strike = rpr.attribute("strike").as_string("noStrike");
            if (strike != "noStrike") font.append_child("strike");
            std::string underline = rpr.attribute("u").as_string("none");
            if (underline != "none") font.append_child("u").append_attribute("val") = underline == "dbl" ? "double" : "single";
            std::string baseline = rpr.attribute("baseline").as_string();
            if (!baseline.empty()) font.append_child("vertAlign").append_attribute("val") = baseline[0] == '-' ? "subscript" : "superscript";
            if (pugi::xml_attribute sz = rpr.attribute("sz")) font.append_child("sz").append_attribute("val") = ods_num(sz.as_double() / 100.0).c_str();
            std::string color = ods_drawing_color(rpr, theme_colors);
            if (color.empty()) color = font_color;
            if (color.size() == 7) font.append_child("color").append_attribute("rgb") = ("FF" + ods_upper(color.substr(1))).c_str();
            std::string typeface = rpr.child("a:latin").attribute("typeface").as_string();
            if (!typeface.empty() && typeface[0] != '+') font.append_child("rFont").append_attribute("val") = typeface.c_str();
            std::string text = run.child_value("a:t");
            if (!text.empty()) has_text = true;
            runs.push_back(std::make_pair(text, text_style(font)));
          }
          fr.paragraphs.push_back(runs);
          if (fr.paragraphs.size() == 1) {
            std::string algn = para.child("a:pPr").attribute("algn").as_string();
            const char* align = nullptr;
            if (algn == "ctr") align = "center";
            else if (algn == "r") align = "end";
            else if (algn == "just" || algn == "justLow" || algn == "dist" || algn == "thaiDist") align = "justify";
            if (align) fr.paragraph_style = paragraph_style(align);
          }
        }
        if (!has_text) fr.paragraphs.clear();

        // explicit colors, or the colors of the shape style
        pugi::xml_document ref_doc;
        pugi::xml_node fill_ref = ref_doc.append_child("fill"), line_ref = ref_doc.append_child("line");
        // the references of index 0 are without fill or line
        int32_t fill_idx = sp_style.child("a:fillRef").attribute("idx").as_int(), line_idx = sp_style.child("a:lnRef").attribute("idx").as_int();
        if (fill_idx > 0) fill_ref.append_child("a:solidFill").append_copy(sp_style.child("a:fillRef").first_child());
        if (line_idx > 0) line_ref.append_child("a:solidFill").append_copy(sp_style.child("a:lnRef").first_child());
        pugi::xml_node fill_src = sppr.child("a:solidFill") ? sppr : sppr.child("a:gradFill").child("a:gsLst").child("a:gs");
        if (!fill_src.child("a:solidFill") && !fill_src.child("a:srgbClr") && !fill_src.child("a:schemeClr")) fill_src = fill_ref;
        std::string fill;
        if (fill_src.child("a:solidFill")) {
          fill = ods_drawing_color(fill_src, theme_colors);
        } else {
          // the first stop of a gradient
          pugi::xml_document stop_doc;
          pugi::xml_node stop = stop_doc.append_child("stop");
          pugi::xml_node solid = stop.append_child("a:solidFill");
          for (pugi::xml_node child : fill_src.children()) solid.append_copy(child);
          fill = ods_drawing_color(stop, theme_colors);
        }
        std::string line = ods_drawing_color(sppr.child("a:ln").child("a:solidFill") ? sppr.child("a:ln") : line_ref, theme_colors);
        fr.style = "gr" + std::to_string(++graphic_styles);
        pugi::xml_node gr = autos.append_child("style:style");
        gr.append_attribute("style:name") = fr.style.c_str();
        gr.append_attribute("style:family") = "graphic";
        pugi::xml_node gp = gr.append_child("style:graphic-properties");
        pugi::xml_node ln = sppr.child("a:ln");
        bool no_fill = fr.line || sppr.child("a:noFill") || fill.empty();
        bool no_line = ln.child("a:noFill") || line.empty();
        // a gradient from its first to its last stop, linear or radial
        pugi::xml_node grad = sppr.child("a:gradFill");
        std::vector<std::string> stops;
        for (pugi::xml_node gs : grad.child("a:gsLst").children("a:gs")) {
          pugi::xml_document stop_doc;
          pugi::xml_node solid = stop_doc.append_child("stop").append_child("a:solidFill");
          for (pugi::xml_node child : gs.children()) solid.append_copy(child);
          std::string color = ods_drawing_color(stop_doc.child("stop"), theme_colors);
          if (!color.empty()) stops.push_back(color);
        }
        if (grad && stops.size() >= 2) {
          std::string gradient = "Gradient_" + std::to_string(gradients.size() + 1);
          pugi::xml_node gn = gradient_doc.append_child("draw:gradient");
          gn.append_attribute("draw:name") = gradient.c_str();
          std::string path_type = grad.child("a:path").attribute("path").as_string();
          gn.append_attribute("draw:style") = grad.child("a:path") ? (path_type == "circle" ? "radial" : "rectangular") : "linear";
          // a path gradient goes from its center to the border, a radial gradient from the border to its center
          bool path_gradient = grad.child("a:path");
          gn.append_attribute("draw:start-color") = (path_gradient ? stops.back() : stops.front()).c_str();
          gn.append_attribute("draw:end-color") = (path_gradient ? stops.front() : stops.back()).c_str();
          if (path_gradient) {
            pugi::xml_node rect = grad.child("a:path").child("a:fillToRect");
            double cx = (rect.attribute("l").as_double() + 100000.0 - rect.attribute("r").as_double()) / 2000.0;
            double cy = (rect.attribute("t").as_double() + 100000.0 - rect.attribute("b").as_double()) / 2000.0;
            gn.append_attribute("draw:cx") = (ods_num(cx) + "%").c_str();
            gn.append_attribute("draw:cy") = (ods_num(cy) + "%").c_str();
          }
          double angle = 90.0 - grad.child("a:lin").attribute("ang").as_double() / 60000.0;
          while (angle < 0.0) angle += 360.0;
          gn.append_attribute("draw:angle") = (ods_num(std::fmod(angle, 360.0)) + "deg").c_str();
          gn.append_attribute("draw:border") = "0%";
          gradients.push_back(gradient);
          gp.append_attribute("draw:fill") = "gradient";
          gp.append_attribute("draw:fill-gradient-name") = gradient.c_str();
        } else {
          gp.append_attribute("draw:fill") = no_fill ? "none" : "solid";
          if (!no_fill) gp.append_attribute("draw:fill-color") = fill.c_str();
          double alpha = ods_drawing_alpha(fill_src.child("a:solidFill"));
          if (!no_fill && alpha < 1.0) gp.append_attribute("draw:opacity") = (ods_num(std::round(alpha * 1000.0) / 10.0) + "%").c_str();
        }

        // lengths in points, exact for those of EMU
        auto pt_length = [](double x) { return ods_num(std::round(x * 1e6) / 1e6) + "pt"; };
        // the line with its width, dash and ends; the width of a line of the shape style is that of the theme
        double line_w = ln.attribute("w") ? ln.attribute("w").as_double() :
                        line_idx > 0 && static_cast<size_t>(line_idx) <= theme_lines.size() ? theme_lines[static_cast<size_t>(line_idx) - 1] : 0.0;
        line_w /= ods_emu_per_pt;
        const ods_dash* dash = nullptr;
        for (const ods_dash& d : ods_dashes) {
          if (std::strcmp(ln.child("a:prstDash").attribute("val").as_string(), d.name) == 0) dash = &d;
        }
        gp.append_attribute("draw:stroke") = no_line ? "none" : dash ? "dash" : "solid";
        if (!no_line && dash) {
          gp.append_attribute("draw:stroke-dash") = dash->name;
          if (drawing_defs.insert(std::string("dash ") + dash->name).second) {
            pugi::xml_node sd = gradient_doc.append_child("draw:stroke-dash");
            sd.append_attribute("draw:name") = dash->name;
            sd.append_attribute("draw:style") = "rect";
            sd.append_attribute("draw:dots1") = dash->dots1;
            sd.append_attribute("draw:dots1-length") = (std::to_string(dash->length1) + "%").c_str();
            if (dash->dots2 > 0) {
              sd.append_attribute("draw:dots2") = dash->dots2;
              sd.append_attribute("draw:dots2-length") = (std::to_string(dash->length2) + "%").c_str();
            }
            sd.append_attribute("draw:distance") = (std::to_string(dash->distance) + "%").c_str();
          }
        }
        if (!no_line) {
          gp.append_attribute("svg:stroke-color") = line.c_str();
          if (line_w > 0.0) gp.append_attribute("svg:stroke-width") = pt_length(line_w).c_str();
          double alpha = ods_drawing_alpha(ln.child("a:solidFill") ? ln.child("a:solidFill") : line_ref.child("a:solidFill"));
          if (alpha < 1.0) gp.append_attribute("svg:stroke-opacity") = (ods_num(std::round(alpha * 1000.0) / 10.0) + "%").c_str();
          // the size of a line end is a multiple of the line width, at least of 2 points
          static const char* const ends[2][2] = {{"a:headEnd", "draw:marker-start"}, {"a:tailEnd", "draw:marker-end"}};
          for (const auto& end : ends) {
            pugi::xml_node node = ln.child(end[0]);
            const ods_marker* marker = nullptr;
            for (const ods_marker& m : ods_markers) {
              if (std::strcmp(node.attribute("type").as_string(), m.type) == 0) marker = &m;
            }
            if (!marker) continue;
            std::string w = node.attribute("w").as_string("med"), len = node.attribute("len").as_string("med");
            int32_t wi = w == "sm" ? 0 : w == "lg" ? 2 : 1, li = len == "sm" ? 0 : len == "lg" ? 2 : 1;
            std::string size = std::to_string(3 * wi + li + 1);
            std::string name = std::string(marker->name) + "_20_" + size;
            if (drawing_defs.insert("marker " + name).second) {
              pugi::xml_node mk = gradient_doc.append_child("draw:marker");
              mk.append_attribute("draw:name") = name.c_str();
              mk.append_attribute("draw:display-name") = (std::string(marker->name) + " " + size).c_str();
              mk.append_attribute("svg:viewBox") = marker->view;
              mk.append_attribute("svg:d") = marker->path;
            }
            gp.append_attribute(end[1]) = name.c_str();
            gp.append_attribute((std::string(end[1]) + "-width").c_str()) = pt_length((wi == 0 ? 2.0 : wi == 1 ? 3.0 : 5.0) * std::max(line_w, 2.0)).c_str();
            gp.append_attribute((std::string(end[1]) + "-center").c_str()) = marker->center ? "true" : "false";
          }
        }

        // the shadow, glow and soft edges
        pugi::xml_node effects = sppr.child("a:effectLst");
        if (pugi::xml_node shadow = effects.child("a:outerShdw")) {
          pugi::xml_document color_doc;
          pugi::xml_node solid = color_doc.append_child("c").append_child("a:solidFill");
          for (pugi::xml_node child : shadow.children()) solid.append_copy(child);
          double dist = shadow.attribute("dist").as_double() / ods_emu_per_pt, angle = shadow.attribute("dir").as_double() / 60000.0 * 3.14159265358979323846 / 180.0;
          std::string color = ods_drawing_color(color_doc.child("c"), theme_colors);
          gp.append_attribute("draw:shadow") = "visible";
          gp.append_attribute("draw:shadow-offset-x") = pt_length(dist * std::cos(angle)).c_str();
          gp.append_attribute("draw:shadow-offset-y") = pt_length(dist * std::sin(angle)).c_str();
          if (!color.empty()) gp.append_attribute("draw:shadow-color") = color.c_str();
          gp.append_attribute("draw:shadow-opacity") = (ods_num(std::round(ods_drawing_alpha(solid) * 1000.0) / 10.0) + "%").c_str();
          gp.append_attribute("loext:shadow-blur") = pt_length(shadow.attribute("blurRad").as_double() / ods_emu_per_pt).c_str();
        }
        if (pugi::xml_node glow = effects.child("a:glow")) {
          pugi::xml_document color_doc;
          pugi::xml_node solid = color_doc.append_child("c").append_child("a:solidFill");
          for (pugi::xml_node child : glow.children()) solid.append_copy(child);
          std::string color = ods_drawing_color(color_doc.child("c"), theme_colors);
          gp.append_attribute("loext:glow-radius") = pt_length(glow.attribute("rad").as_double() / ods_emu_per_pt).c_str();
          if (!color.empty()) gp.append_attribute("loext:glow-color") = color.c_str();
          gp.append_attribute("loext:glow-transparency") = (ods_num(std::round((1.0 - ods_drawing_alpha(solid)) * 1000.0) / 10.0) + "%").c_str();
        }
        if (pugi::xml_node soft = effects.child("a:softEdge")) {
          gp.append_attribute("loext:softedge-radius") = pt_length(soft.attribute("rad").as_double() / ods_emu_per_pt).c_str();
        }

        // the text box: insets, vertical anchor, wrapping and growing with the text
        pugi::xml_node body = el.child("xdr:txBody").child("a:bodyPr");
        if (!fr.line) {
          std::string anchor = body.attribute("anchor").as_string("t");
          gp.append_attribute("draw:textarea-vertical-align") = anchor == "ctr" ? "middle" : anchor == "b" ? "bottom" : "top";
          static const char* const insets[4][3] = {
            {"lIns", "fo:padding-left", "91440"}, {"tIns", "fo:padding-top", "45720"}, {"rIns", "fo:padding-right", "91440"}, {"bIns", "fo:padding-bottom", "45720"}
          };
          for (const auto& inset : insets) {
            gp.append_attribute(inset[1]) = pt_length(body.attribute(inset[0]).as_double(std::atof(inset[2])) / ods_emu_per_pt).c_str();
          }
          gp.append_attribute("fo:wrap-option") = std::strcmp(body.attribute("wrap").as_string(), "none") == 0 ? "no-wrap" : "wrap";
        }
        gp.append_attribute("draw:auto-grow-height") = body.child("a:spAutoFit") ? "true" : "false";
        gp.append_attribute("draw:auto-grow-width") = "false";
        return true;
      };

      // charts are written as objects once their size is known
      int32_t z_index = 0;
      auto finish_frame = [&](ods_frame& fr) {
        fr.z_index = z_index++;
        if (!fr.shape.empty()) return true;
        // pictures and charts without fill and border
        if (frame_style.empty()) {
          frame_style = "gr" + std::to_string(++graphic_styles);
          pugi::xml_node gr = autos.append_child("style:style");
          gr.append_attribute("style:name") = frame_style.c_str();
          gr.append_attribute("style:family") = "graphic";
          pugi::xml_node gp = gr.append_child("style:graphic-properties");
          gp.append_attribute("draw:stroke") = "none";
          gp.append_attribute("draw:fill") = "none";
        }
        fr.style = frame_style;
        if (!fr.chart) {
          pictures.insert(fr.href);
          // a cropped picture is clipped relative to its own size
          double image_w = 0.0, image_h = 0.0;
          bool cropped = fr.crop[0] != 0.0 || fr.crop[1] != 0.0 || fr.crop[2] != 0.0 || fr.crop[3] != 0.0;
          if (cropped && ods_image_size(path + fr.href, image_w, image_h)) {
            fr.style = "gr" + std::to_string(++graphic_styles);
            pugi::xml_node gr = autos.append_child("style:style");
            gr.append_attribute("style:name") = fr.style.c_str();
            gr.append_attribute("style:family") = "graphic";
            char clip[160];
            std::snprintf(clip, sizeof(clip), "rect(%.6fcm, %.6fcm, %.6fcm, %.6fcm)",
                          fr.crop[1] / 100000.0 * image_h / 72.0 * 2.54, fr.crop[2] / 100000.0 * image_w / 72.0 * 2.54,
                          fr.crop[3] / 100000.0 * image_h / 72.0 * 2.54, fr.crop[0] / 100000.0 * image_w / 72.0 * 2.54);
            pugi::xml_node gp = gr.append_child("style:graphic-properties");
            gp.append_attribute("draw:stroke") = "none";
            gp.append_attribute("draw:fill") = "none";
            gp.append_attribute("fo:clip") = clip;
          }
          return true;
        }
        // the chart is the part chartN.xml of the workbook charts
        std::string file = fr.target.substr(fr.target.rfind('/') + 1);
        size_t id = file.compare(0, 5, "chart") == 0 ? static_cast<size_t>(std::strtoul(file.c_str() + 5, nullptr, 10)) : 0;
        fr.href = "Object " + std::to_string(id);
        if (id == 0 || id > charts.size()) return false;
        if (!ods_chart_object(charts[id - 1], theme_colors, fr.width, fr.height, path + fr.href + "/content.xml", version, cell_values, fr.ranges, pivot_names)) {
          return false;
        }
        objects.push_back(fr.href);
        return true;
      };

      // the elements of a group are placed by the child offsets and extents of the group
      int32_t groups = 0;
      std::function<void(pugi::xml_node, double, double, double, double, int32_t)> group_frames;
      group_frames = [&](pugi::xml_node grp, double gx, double gy, double gw, double gh, int32_t group) {
        pugi::xml_node xfrm = grp.child("xdr:grpSpPr").child("a:xfrm");
        double ch_x = xfrm.child("a:chOff").attribute("x").as_double(), ch_y = xfrm.child("a:chOff").attribute("y").as_double();
        double ch_w = xfrm.child("a:chExt").attribute("cx").as_double(), ch_h = xfrm.child("a:chExt").attribute("cy").as_double();
        double sx = ch_w > 0.0 ? gw / ch_w : 1.0 / ods_emu_per_pt, sy = ch_h > 0.0 ? gh / ch_h : 1.0 / ods_emu_per_pt;
        for (pugi::xml_node el : grp.children()) {
          std::string el_name = el.name();
          pugi::xml_node el_xfrm = el_name == "xdr:graphicFrame" ? el.child("xdr:xfrm") :
                                   el_name == "xdr:grpSp" ? el.child("xdr:grpSpPr").child("a:xfrm") : el.child("xdr:spPr").child("a:xfrm");
          double x = gx + (el_xfrm.child("a:off").attribute("x").as_double() - ch_x) * sx;
          double y = gy + (el_xfrm.child("a:off").attribute("y").as_double() - ch_y) * sy;
          double w = el_xfrm.child("a:ext").attribute("cx").as_double() * sx, h = el_xfrm.child("a:ext").attribute("cy").as_double() * sy;
          if (el_name == "xdr:grpSp") {
            group_frames(el, x, y, w, h, group);
            continue;
          }
          ods_frame fr;
          if (!element_frame(el, fr)) continue;
          fr.x = x;
          fr.y = y;
          fr.width = w;
          fr.height = h;
          fr.group = group;
          if (finish_frame(fr)) sheet_frames.push_back(fr);
        }
      };

      for (pugi::xml_node anchor : drawing_doc.child("xdr:wsDr").children()) {
        std::string kind = anchor.name();
        if (kind != "xdr:twoCellAnchor" && kind != "xdr:oneCellAnchor" && kind != "xdr:absoluteAnchor") continue;
        ods_frame fr;
        pugi::xml_node grp = anchor.child("xdr:grpSp");
        pugi::xml_node el = anchor.child("xdr:pic");
        if (!el) el = anchor.child("xdr:graphicFrame");
        if (!el) el = anchor.child("xdr:sp");
        if (!el) el = anchor.child("xdr:cxnSp");
        if (!grp && !element_frame(el, fr)) continue;

        pugi::xml_node from = anchor.child("xdr:from"), to = anchor.child("xdr:to");
        pugi::xml_node ext = anchor.child("xdr:ext");
        int64_t c0 = 0, r0 = 0;
        if (kind == "xdr:absoluteAnchor") {
          fr.x = anchor.child("xdr:pos").attribute("x").as_double() / ods_emu_per_pt;
          fr.y = anchor.child("xdr:pos").attribute("y").as_double() / ods_emu_per_pt;
        } else {
          c0 = std::atoll(from.child_value("xdr:col"));
          r0 = std::atoll(from.child_value("xdr:row"));
          fr.col = static_cast<int32_t>(c0 + 1);
          fr.row = static_cast<int32_t>(r0 + 1);
          fr.x = std::atof(from.child_value("xdr:colOff")) / ods_emu_per_pt;
          fr.y = std::atof(from.child_value("xdr:rowOff")) / ods_emu_per_pt;
          if (kind == "xdr:twoCellAnchor") {
            int64_t c1 = std::atoll(to.child_value("xdr:col")), r1 = std::atoll(to.child_value("xdr:row"));
            fr.end_x = std::atof(to.child_value("xdr:colOff")) / ods_emu_per_pt;
            fr.end_y = std::atof(to.child_value("xdr:rowOff")) / ods_emu_per_pt;
            fr.width = col_x(c1) + fr.end_x - col_x(c0) - fr.x;
            fr.height = row_y(r1) + fr.end_y - row_y(r0) - fr.y;
            if (c1 < MAX_OOXML_COL_INT && r1 < MAX_OOXML_ROW_INT) {
              fr.end_cell = ods_sheet_name(sheet_quoted) + "." + int_to_col(c1 + 1) + std::to_string(r1 + 1);
              fr.end_row = static_cast<int32_t>(r1 + 1);
            }
          }
        }
        if (kind != "xdr:twoCellAnchor") {
          fr.width = ext.attribute("cx").as_double() / ods_emu_per_pt;
          fr.height = ext.attribute("cy").as_double() / ods_emu_per_pt;
        }

        // groups are placed on the sheet
        if (grp) {
          double gx = fr.x, gy = fr.y;
          if (kind != "xdr:absoluteAnchor") {
            gx += col_x(c0);
            gy += row_y(r0);
          }
          group_frames(grp, gx, gy, fr.width, fr.height, ++groups);
          continue;
        }
        if (!finish_frame(fr)) continue;
        if (fr.row > 0) {
          frames.push_back(fr);
          max_row = std::max(max_row, fr.row);
          max_col = std::max(max_col, fr.col);
        } else {
          sheet_frames.push_back(fr);
        }
      }
    }
    // form controls of the VML drawing as controls of a form of the sheet, placed on the sheet
    pugi::xml_document vml_doc, controls_doc;
    pugi::xml_node form = controls_doc.append_child("form:form");
    std::vector<std::pair<std::string, std::vector<double>>> control_frames;
    if (vml_doc.load_string(Rcpp::as<utf8_string>(sheet["vml"]).c_str())) {
      for (pugi::xml_node shape : ods_descendants(vml_doc, "v:shape")) {
        pugi::xml_node cd = shape.child("x:ClientData");
        std::string type = cd.attribute("ObjectType").as_string();
        // the control and its implementation in the office application
        const char* element = nullptr;
        const char* implementation = nullptr;
        if (type == "Checkbox") {
          element = "form:checkbox";
          implementation = "ooo:com.sun.star.form.component.CheckBox";
        } else if (type == "Radio") {
          element = "form:radio";
          implementation = "ooo:com.sun.star.form.component.RadioButton";
        } else if (type == "Drop" || type == "List") {
          element = "form:listbox";
          implementation = "ooo:com.sun.star.form.component.ListBox";
        } else if (type == "Button") {
          element = "form:button";
          implementation = "ooo:com.sun.star.form.component.CommandButton";
        } else if (type == "Label") {
          element = "form:fixed-text";
          implementation = "ooo:com.sun.star.form.component.FixedText";
        }
        if (!element) continue;

        // the anchor: left column, offset, top row, offset, right column, offset, bottom row, offset; offsets in pixels
        std::vector<double> anchor;
        const char* a = cd.child_value("x:Anchor");
        while (*a && anchor.size() < 8) {
          char* end = nullptr;
          double v = std::strtod(a, &end);
          if (end == a) break;
          anchor.push_back(v);
          a = end;
          while (*a == ',' || *a == ' ') ++a;
        }
        if (anchor.size() < 8) continue;
        double x = col_x(static_cast<int64_t>(anchor[0])) + anchor[1] * 0.75, y = row_y(static_cast<int64_t>(anchor[2])) + anchor[3] * 0.75;
        double x2 = col_x(static_cast<int64_t>(anchor[4])) + anchor[5] * 0.75, y2 = row_y(static_cast<int64_t>(anchor[6])) + anchor[7] * 0.75;

        std::string id = "control" + std::to_string(++control_count), label;
        for (pugi::xml_node t : ods_descendants(shape.child("v:textbox"), "")) {
          if (t.type() == pugi::node_pcdata) label += t.value();
        }
        pugi::xml_node control = form.append_child(element);
        control.append_attribute("form:name") = (type + " " + std::to_string(control_count)).c_str();
        control.append_attribute("form:control-implementation") = implementation;
        control.append_attribute("xml:id") = id.c_str();
        control.append_attribute("form:id") = id.c_str();
        if (!label.empty() && type != "Drop" && type != "List") control.append_attribute("form:label") = label.c_str();
        std::string link = cd.child_value("x:FmlaLink");
        std::string xlsx_sheet = "'" + sheet_quoted + "'!";
        if (!link.empty()) {
          std::string address = ods_chart_range(link.find('!') == std::string::npos ? xlsx_sheet + link : link);
          if (!address.empty() && type != "Button" && type != "Label") control.append_attribute("form:linked-cell") = address.c_str();
        }
        bool checked = std::atoi(cd.child_value("x:Checked")) == 1;
        if (type == "Checkbox") control.append_attribute("form:current-state") = checked ? "checked" : "unchecked";
        if (type == "Radio" && checked) control.append_attribute("form:current-selected") = "true";
        if (type == "Drop" || type == "List") {
          control.append_attribute("form:dropdown") = type == "Drop" ? "true" : "false";
          std::string range = cd.child_value("x:FmlaRange");
          std::string address = range.empty() ? "" : ods_chart_range(range.find('!') == std::string::npos ? xlsx_sheet + range : range);
          if (!address.empty()) control.append_attribute("form:source-cell-range") = address.c_str();
          if (!link.empty()) control.append_attribute("form:list-linkage-type") = "selection-indices";
        }
        control_frames.push_back(std::make_pair(id, std::vector<double>{x, y, std::max(x2 - x, 1.0), std::max(y2 - y, 1.0)}));
      }
    }
    if (!sheet_frames.empty() || !control_frames.empty()) {
      pugi::xml_node shapes = table.prepend_child("table:shapes");
      for (const auto& cf : control_frames) {
        pugi::xml_node control = shapes.append_child("draw:control");
        control.append_attribute("svg:width") = ods_cm(cf.second[2]).c_str();
        control.append_attribute("svg:height") = ods_cm(cf.second[3]).c_str();
        control.append_attribute("svg:x") = ods_cm(cf.second[0]).c_str();
        control.append_attribute("svg:y") = ods_cm(cf.second[1]).c_str();
        control.append_attribute("draw:control") = cf.first.c_str();
      }
      if (!control_frames.empty()) {
        pugi::xml_node forms = table.prepend_child("office:forms");
        forms.append_attribute("form:automatic-focus") = "false";
        forms.append_attribute("form:apply-design-mode") = "false";
        pugi::xml_node sheet_form = forms.append_copy(form);
        sheet_form.prepend_attribute("form:name") = "Form";
      }
      pugi::xml_node group_node;
      int32_t group = 0;
      for (const ods_frame& fr : sheet_frames) {
        if (fr.group == 0) {
          ods_append_frame(shapes, fr);
          continue;
        }
        if (fr.group != group) {
          group_node = shapes.append_child("draw:g");
          group_node.append_attribute("draw:z-index") = fr.z_index;
          group = fr.group;
        }
        ods_append_frame(group_node, fr);
      }
    }

    auto col_style = [&](double width, bool page_break) {
      char buf[32];
      std::snprintf(buf, sizeof(buf), "%.4fin", width * unit_pt / 72.0);
      std::string key = std::string(buf) + (page_break ? "b" : "");
      if (col_styles.find(key) == col_styles.end()) {
        std::string name = "co" + std::to_string(col_styles.size() + 1);
        pugi::xml_node co = autos.append_child("style:style");
        co.append_attribute("style:name") = name.c_str();
        co.append_attribute("style:family") = "table-column";
        pugi::xml_node props = co.append_child("style:table-column-properties");
        if (page_break) props.append_attribute("fo:break-before") = "page";
        props.append_attribute("style:column-width") = buf;
        col_styles[key] = name;
      }
      return col_styles[key];
    };

    // data validations and conditional formats of the extensions of the worksheet. Validations and
    // formats of their own are added to the others, extensions of data bars are kept by their id.
    std::vector<utf8_string> validation_list = Rcpp::as<std::vector<utf8_string>>(sheet["validations"]);
    std::vector<utf8_string> cf_sqref_list = Rcpp::as<std::vector<utf8_string>>(sheet["cf_sqref"]);
    std::vector<utf8_string> cf_rule_list = Rcpp::as<std::vector<utf8_string>>(sheet["cf_rule"]);
    std::map<std::string, std::string> bar_extensions;
    {
      auto local = [](pugi::xml_node node) {
        const char* colon = std::strchr(node.name(), ':');
        return std::string(colon ? colon + 1 : node.name());
      };
      auto child = [&](pugi::xml_node node, const char* name) {
        for (pugi::xml_node c : node.children()) {
          if (local(c) == name) return c;
        }
        return pugi::xml_node();
      };
      // the rules of the base list that have an extension
      std::set<std::string> extended;
      for (const utf8_string& rule : cf_rule_list) {
        pugi::xml_document rd;
        rd.load_string(rule.c_str());
        for (pugi::xml_node id : ods_descendants(rd.document_element(), "x14:id")) extended.insert(id.text().get());
      }
      pugi::xml_document ext_doc;
      for (pugi::xml_node ext : ods_load_list(ext_doc, Rcpp::as<std::vector<utf8_string>>(sheet["ext"]))) {
        std::string uri = ext.attribute("uri").as_string();
        if (uri == "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}") {
          for (pugi::xml_node dv : ext.first_child().children()) {
            if (local(dv) != "dataValidation") continue;
            pugi::xml_document out_doc;
            pugi::xml_node out = out_doc.append_child("dataValidation");
            for (pugi::xml_attribute attr : dv.attributes()) {
              if (std::strchr(attr.name(), ':') == nullptr) out.append_attribute(attr.name()) = attr.value();
            }
            out.append_attribute("sqref") = child(dv, "sqref").text().get();
            for (const char* formula : {"formula1", "formula2"}) {
              pugi::xml_node f = child(child(dv, formula), "f");
              if (f) out.append_child(formula).text().set(f.text().get());
            }
            validation_list.push_back(ods_xml_string(out));
          }
        } else if (uri == "{78C0D931-6437-407d-A8EE-F0AAD7539E65}") {
          for (pugi::xml_node cf : ext.first_child().children()) {
            if (local(cf) != "conditionalFormatting") continue;
            std::string sqref = child(cf, "sqref").text().get();
            for (pugi::xml_node rule : cf.children()) {
              if (local(rule) != "cfRule") continue;
              std::string id = rule.attribute("id").as_string();
              if (extended.count(id)) {
                if (pugi::xml_node bar = child(rule, "dataBar")) bar_extensions[id] = ods_xml_string(bar);
                continue;
              }
              // a rule of its own with its formulas, the values of its thresholds and its differential format
              pugi::xml_document out_doc;
              pugi::xml_node out = out_doc.append_child("cfRule");
              for (pugi::xml_attribute attr : rule.attributes()) {
                if (std::strcmp(attr.name(), "id") != 0 && std::strchr(attr.name(), ':') == nullptr) out.append_attribute(attr.name()) = attr.value();
              }
              if (pugi::xml_node dxf = child(rule, "dxf")) {
                pugi::xml_document dxf_out;
                pugi::xml_node copy = dxf_out.append_child("dxf");
                for (pugi::xml_node part : dxf.children()) copy.append_copy(part);
                out.append_attribute("dxfId") = static_cast<uint32_t>(n_dxfs + extra_dxfs.size());
                extra_dxfs.push_back(ods_xml_string(copy));
              }
              for (pugi::xml_node part : rule.children()) {
                std::string kind = local(part);
                if (kind == "f") out.append_child("formula").text().set(part.text().get());
                if (kind != "colorScale" && kind != "iconSet" && kind != "dataBar") continue;
                pugi::xml_node inner = out.append_child(kind.c_str());
                for (pugi::xml_attribute attr : part.attributes()) inner.append_attribute(attr.name()) = attr.value();
                for (pugi::xml_node item : part.children()) {
                  std::string item_kind = local(item);
                  if (item_kind == "cfvo") {
                    pugi::xml_node cfvo = inner.append_child("cfvo");
                    for (pugi::xml_attribute attr : item.attributes()) cfvo.append_attribute(attr.name()) = attr.value();
                    if (pugi::xml_node f = child(item, "f")) cfvo.append_attribute("val") = f.text().get();
                  } else if (item_kind == "color" || item_kind == "fillColor") {
                    pugi::xml_node color = inner.append_child("color");
                    for (pugi::xml_attribute attr : item.attributes()) color.append_attribute(attr.name()) = attr.value();
                  }
                }
              }
              cf_sqref_list.push_back(sqref);
              cf_rule_list.push_back(ods_xml_string(out));
            }
          }
        }
      }
    }

    // conditional formats with a condition of a cell style are also written as style maps of
    // the cell styles, for consumers without the calc extension
    struct ods_cf_map {
      std::vector<ods_range> ranges;
      std::string condition, style, base;
      int64_t priority;
    };
    std::vector<ods_cf_map> cf_maps;
    {
      const std::vector<utf8_string>& map_sqref = cf_sqref_list;
      const std::vector<utf8_string>& map_rule = cf_rule_list;
      static const char* const cell_ops[][2] = {
        {"equal", "="}, {"notEqual", "!="}, {"greaterThan", ">"}, {"lessThan", "<"},
        {"greaterThanOrEqual", ">="}, {"lessThanOrEqual", "<="}
      };
      for (size_t k = 0; k < map_sqref.size() && k < map_rule.size(); ++k) {
        pugi::xml_document rd;
        if (!rd.load_string(map_rule[k].c_str())) continue;
        pugi::xml_node rule = rd.child("cfRule");
        if (!rule.attribute("dxfId")) continue;
        std::vector<ods_range> ranges = ods_sqref(map_sqref[k]);
        if (ranges.empty()) continue;
        ods_formula_ctx mctx;
        mctx.tables = &tables_info;
        mctx.sheet = sheet_quoted;
        std::vector<std::string> formulas;
        for (pugi::xml_node f : rule.children("formula")) formulas.push_back(ods_of_value(f.text().get(), mctx));
        std::string type = rule.attribute("type").as_string(), op = rule.attribute("operator").as_string(), condition;
        if (type == "cellIs" && !formulas.empty()) {
          if ((op == "between" || op == "notBetween") && formulas.size() == 2) {
            condition = std::string(op == "between" ? "cell-content-is-between(" : "cell-content-is-not-between(") + formulas[0] + "," + formulas[1] + ")";
          }
          for (const auto& co : cell_ops) {
            if (op == co[0]) condition = std::string("cell-content()") + co[1] + formulas[0];
          }
        } else if (type != "cellIs" && type != "colorScale" && type != "dataBar" && type != "iconSet" && !formulas.empty()) {
          condition = "is-true-formula(" + formulas[0] + ")";
        }
        if (condition.empty()) continue;
        ods_cf_map map;
        map.ranges = ranges;
        map.condition = condition;
        map.style = "ConditionalStyle_" + std::to_string(rule.attribute("dxfId").as_uint() + 1);
        map.base = ods_range_address(sheet_quoted, ods_range{ranges[0].r1, ranges[0].c1, ranges[0].r1, ranges[0].c1});
        map.priority = rule.attribute("priority").as_llong(0);
        cf_maps.push_back(map);
      }
      std::stable_sort(cf_maps.begin(), cf_maps.end(), [](const ods_cf_map& a, const ods_cf_map& b) { return a.priority < b.priority; });
    }
    // the cell style of a cell in conditional formats: a copy of its style with the maps
    std::map<std::string, std::string> map_styles;
    auto cell_style_name = [&](int32_t style, int64_t row, int64_t col) {
      std::string base = style > 0 ? "ce" + std::to_string(style) : "";
      std::string key = base;
      std::vector<const ods_cf_map*> maps;
      for (const ods_cf_map& map : cf_maps) {
        for (const ods_range& range : map.ranges) {
          if (row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) {
            maps.push_back(&map);
            key += "|" + std::to_string(&map - cf_maps.data());
            break;
          }
        }
      }
      if (maps.empty()) return base;
      auto known = map_styles.find(key);
      if (known != map_styles.end()) return known->second;
      std::string name = "cm" + std::to_string(++map_style_count);
      pugi::xml_node source, copy;
      for (pugi::xml_node st : autos.children("style:style")) {
        if (!source && !base.empty() && base == st.attribute("style:name").as_string()) source = st;
      }
      if (source) copy = autos.append_copy(source);
      if (!copy) {
        copy = autos.append_child("style:style");
        copy.append_attribute("style:name") = name.c_str();
        copy.append_attribute("style:family") = "table-cell";
        copy.append_attribute("style:parent-style-name") = "Default";
      }
      copy.attribute("style:name") = name.c_str();
      for (const ods_cf_map* map : maps) {
        pugi::xml_node style_map = copy.append_child("style:map");
        style_map.append_attribute("style:condition") = map->condition.c_str();
        style_map.append_attribute("style:apply-style-name") = map->style.c_str();
        style_map.append_attribute("style:base-cell-address") = map->base.c_str();
      }
      map_styles[key] = name;
      return name;
    };

    // data validations; the cells of a validation carry its name
    struct ods_valid {
      ods_range range;
      std::string name;
    };
    std::vector<ods_valid> valids;
    pugi::xml_document dv_doc;
    for (pugi::xml_node dv : ods_load_list(dv_doc, validation_list)) {
      std::vector<ods_range> ranges = ods_sqref(dv.attribute("sqref").as_string());
      if (ranges.empty()) continue;
      ods_formula_ctx vctx;
      vctx.tables = &tables_info;
      vctx.sheet = sheet_quoted;
      std::string condition = ods_validation_condition(dv, vctx);
      if (condition.empty()) continue;
      std::string vname = "val" + std::to_string(++validation_count);
      pugi::xml_node cv = validations.append_child("table:content-validation");
      cv.append_attribute("table:name") = vname.c_str();
      cv.append_attribute("table:condition") = condition.c_str();
      cv.append_attribute("table:allow-empty-cell") = dv.attribute("allowBlank").as_bool() ? "true" : "false";
      if (std::string(dv.attribute("type").as_string()) == "list") {
        cv.append_attribute("table:display-list") = dv.attribute("showDropDown").as_bool() ? "none" : "unsorted";
      }
      ods_range first = {ranges[0].r1, ranges[0].c1, ranges[0].r1, ranges[0].c1};
      cv.append_attribute("table:base-cell-address") = ods_range_address(sheet_quoted, first).c_str();
      pugi::xml_node help = cv.append_child("table:help-message");
      if (*dv.attribute("promptTitle").as_string()) help.append_attribute("table:title") = dv.attribute("promptTitle").as_string();
      help.append_attribute("table:display") = dv.attribute("showInputMessage").as_bool() ? "true" : "false";
      if (*dv.attribute("prompt").as_string()) ods_cell_text(help, dv.attribute("prompt").as_string());
      pugi::xml_node error = cv.append_child("table:error-message");
      if (*dv.attribute("errorTitle").as_string()) error.append_attribute("table:title") = dv.attribute("errorTitle").as_string();
      error.append_attribute("table:display") = dv.attribute("showErrorMessage").as_bool() ? "true" : "false";
      std::string style = dv.attribute("errorStyle").as_string("stop");
      error.append_attribute("table:message-type") = style == "warning" ? "warning" : style == "information" ? "information" : "stop";
      if (*dv.attribute("error").as_string()) ods_cell_text(error, dv.attribute("error").as_string());
      for (const ods_range& range : ranges) {
        ods_valid v = {range, vname};
        valids.push_back(v);
      }
    }

    // groups of the outline are collapsed with a collapsed summary row or column after them (or before
    // them with the summary above or left), or when all of their rows or columns are hidden
    pugi::xml_document outline_doc;
    outline_doc.load_string(Rcpp::as<utf8_string>(sheet["sheet_pr"]).c_str());
    pugi::xml_node outline_pr = outline_doc.child("sheetPr").child("outlinePr");
    bool summary_below = outline_pr.attribute("summaryBelow").as_bool(true), summary_right = outline_pr.attribute("summaryRight").as_bool(true);
    auto col_at = [&](int64_t col) {
      const ods_out_col* found = nullptr;
      for (const ods_out_col& oc : cols) {
        if (col >= oc.min && col <= oc.max) found = &oc;
      }
      return found;
    };
    auto row_at = [&](int64_t row) {
      auto it = rows.find(static_cast<int32_t>(row));
      return it == rows.end() ? nullptr : &it->second;
    };
    std::vector<pugi::xml_node> col_groups = {table};
    auto col_parent = [&](int64_t start, int32_t level) {
      while (static_cast<int32_t>(col_groups.size()) - 1 > level) col_groups.pop_back();
      while (static_cast<int32_t>(col_groups.size()) - 1 < level) {
        int32_t group_level = static_cast<int32_t>(col_groups.size());
        int64_t end = start;
        bool all_hidden = true;
        for (const ods_out_col* oc = col_at(end); oc && oc->level >= group_level && end <= MAX_OOXML_COL_INT; oc = col_at(++end)) {
          if (!oc->hidden) all_hidden = false;
        }
        const ods_out_col* summary = col_at(summary_right ? end : start - 1);
        pugi::xml_node group = col_groups.back().append_child("table:table-column-group");
        if (all_hidden || (summary && summary->collapsed)) group.append_attribute("table:display") = "false";
        col_groups.push_back(group);
      }
      return col_groups.back();
    };

    // the columns as segments without gaps, split at the print title columns
    std::vector<ods_out_col> segments;
    int32_t next_col = 1;
    for (const ods_out_col& oc : cols) {
      if (oc.min < next_col) continue;
      if (oc.min > next_col) segments.push_back(ods_out_col{next_col, oc.min - 1, "", false, 0, 0});
      segments.push_back(oc);
      next_col = oc.max + 1;
    }
    // the columns after the last one have the default width of the sheet
    if (next_col <= MAX_OOXML_COL_INT) segments.push_back(ods_out_col{next_col, MAX_OOXML_COL_INT, "", false, 0, 0});
    pugi::xml_node header_cols;
    for (const ods_out_col& seg : segments) {
      std::vector<int32_t> cuts = {seg.min};
      for (int64_t bound : {pr.c1, pr.c2 + 1}) {
        if (pr.c1 > 0 && bound > seg.min && bound <= seg.max) cuts.push_back(static_cast<int32_t>(bound));
      }
      // a column with a page break before it is a column of its own
      for (int64_t brk : col_breaks) {
        if (brk < seg.min || brk > seg.max) continue;
        cuts.push_back(static_cast<int32_t>(brk));
        cuts.push_back(static_cast<int32_t>(brk) + 1);
      }
      cuts.push_back(seg.max + 1);
      std::sort(cuts.begin(), cuts.end());
      cuts.erase(std::unique(cuts.begin(), cuts.end()), cuts.end());
      for (size_t k = 0; k + 1 < cuts.size(); ++k) {
        ods_out_col oc = seg;
        oc.min = cuts[k];
        oc.max = cuts[k + 1] - 1;
        pugi::xml_node parent = col_parent(oc.min, oc.level);
        if (pr.c1 > 0 && oc.min >= pr.c1 && oc.max <= pr.c2) {
          if (!header_cols) header_cols = parent.append_child("table:table-header-columns");
          if (header_cols.parent() == parent) parent = header_cols;
        }
        pugi::xml_node column = parent.append_child("table:table-column");
        // ODF widths are positive, columns without width are hidden
        double width = oc.width.empty() ? default_width : std::strtod(oc.width.c_str(), nullptr);
        bool col_hidden = oc.hidden || !(width > 0.0);
        column.append_attribute("table:style-name") = col_style(width > 0.0 ? width : default_width, col_breaks.count(oc.min) > 0).c_str();
        if (oc.max > oc.min) column.append_attribute("table:number-columns-repeated") = oc.max - oc.min + 1;
        if (col_hidden) column.append_attribute("table:visibility") = "collapse";
        if (oc.style > 0) column.append_attribute("table:default-cell-style-name") = ("ce" + std::to_string(oc.style)).c_str();
      }
    }

    std::set<int32_t> row_numbers;
    for (const ods_out_cell& cell : cells) row_numbers.insert(cell.row);
    for (const auto& row : rows) row_numbers.insert(row.first);
    for (const ods_merge& m : merges) {
      for (int32_t r = m.r1; r <= m.r2; ++r) row_numbers.insert(r);
    }
    for (const auto& cm : comments) row_numbers.insert(cm.first.first);
    // the rows covered by frames are written with their height
    for (const ods_frame& fr : frames) {
      row_numbers.insert(fr.row);
      if (fr.end_row > 0) row_numbers.insert(fr.end_row);
    }
    std::vector<double> bottoms;
    for (const ods_frame& fr : sheet_frames) bottoms.push_back(fr.y + fr.height);
    for (const auto& cf : control_frames) bottoms.push_back(cf.second[1] + cf.second[3]);
    for (double bottom : bottoms) {
      double top = 0.0;
      int32_t r = 1;
      for (; r < MAX_OOXML_ROW_INT && r < 65536; ++r) {
        auto ra = rows.find(r);
        top += ra == rows.end() || ra->second.height.empty() ? default_height : std::strtod(ra->second.height.c_str(), nullptr);
        if (top >= bottom) break;
      }
      row_numbers.insert(r);
    }
    for (int64_t r = pr.r1; r > 0 && r <= pr.r2; ++r) row_numbers.insert(static_cast<int32_t>(r));
    for (int64_t r : row_breaks) {
      if (r > 1 && r <= MAX_OOXML_ROW_INT) row_numbers.insert(static_cast<int32_t>(r));
    }

    // row groups of the outline
    std::vector<pugi::xml_node> row_groups = {table};
    auto row_parent = [&](int64_t start, int32_t level) {
      while (static_cast<int32_t>(row_groups.size()) - 1 > level) row_groups.pop_back();
      while (static_cast<int32_t>(row_groups.size()) - 1 < level) {
        int32_t group_level = static_cast<int32_t>(row_groups.size());
        int64_t end = start;
        bool all_hidden = true;
        for (const ods_out_row* ro = row_at(end); ro && ro->level >= group_level && end <= MAX_OOXML_ROW_INT; ro = row_at(++end)) {
          if (!ro->hidden) all_hidden = false;
        }
        const ods_out_row* summary = row_at(summary_below ? end : start - 1);
        pugi::xml_node group = row_groups.back().append_child("table:table-row-group");
        if (all_hidden || (summary && summary->collapsed)) group.append_attribute("table:display") = "false";
        row_groups.push_back(group);
      }
      return row_groups.back();
    };
    // the print title rows are header rows
    pugi::xml_node header_rows;
    auto row_container = [&](int64_t r, int32_t level) {
      pugi::xml_node parent = row_parent(r, level);
      if (pr.r1 == 0 || r < pr.r1 || r > pr.r2) return parent;
      if (!header_rows) header_rows = parent.append_child("table:table-header-rows");
      return header_rows.parent() == parent ? header_rows : parent;
    };

    // the height of rows; rows without height have the default height of the sheet
    auto row_style = [&](const std::string& height, bool custom, bool page_break) {
      std::string key = height + (custom ? "c" : "") + (page_break ? "b" : "");
      if (row_styles.find(key) == row_styles.end()) {
        std::string name = "ro" + std::to_string(row_styles.size() + 1);
        pugi::xml_node ro = autos.append_child("style:style");
        ro.append_attribute("style:name") = name.c_str();
        ro.append_attribute("style:family") = "table-row";
        pugi::xml_node props = ro.append_child("style:table-row-properties");
        props.append_attribute("style:row-height") = (height + "pt").c_str();
        if (page_break) props.append_attribute("fo:break-before") = "page";
        props.append_attribute("style:use-optimal-row-height") = custom ? "false" : "true";
        row_styles[key] = name;
      }
      return row_styles[key];
    };
    std::string default_height_pt = ods_num(default_height > 0.0 ? default_height : 15.0);
    std::string default_row = row_style(default_height_pt, false, false);

    // rows without cells; validations split them into blocks of equal rows
    auto empty_rows = [&](int64_t from, int64_t to) {
      std::set<int64_t> bounds = {from, to + 1};
      for (int64_t bound : {pr.r1, pr.r2 + 1}) {
        if (pr.r1 > 0 && bound > from && bound <= to) bounds.insert(bound);
      }
      for (const ods_valid& v : valids) {
        if (v.range.r1 > from && v.range.r1 <= to) bounds.insert(v.range.r1);
        if (v.range.r2 + 1 > from && v.range.r2 + 1 <= to) bounds.insert(v.range.r2 + 1);
      }
      for (auto b = bounds.begin(); std::next(b) != bounds.end(); ++b) {
        int64_t start = *b, count = *std::next(b) - *b;
        pugi::xml_node gap = row_container(start, 0).append_child("table:table-row");
        gap.append_attribute("table:style-name") = default_row.c_str();
        if (count > 1) gap.append_attribute("table:number-rows-repeated") = count;
        std::map<int64_t, const ods_valid*> covered;
        for (const ods_valid& v : valids) {
          if (v.range.r1 <= start && v.range.r2 >= start) covered[v.range.c1] = &v;
        }
        int64_t col = 1;
        for (const auto& cv : covered) {
          if (cv.first < col) continue;
          if (cv.first > col) {
            pugi::xml_node cell = gap.append_child("table:table-cell");
            if (cv.first - col > 1) cell.append_attribute("table:number-columns-repeated") = cv.first - col;
          }
          pugi::xml_node cell = gap.append_child("table:table-cell");
          cell.append_attribute("table:content-validation-name") = cv.second->name.c_str();
          if (cv.second->range.c2 > cv.first) cell.append_attribute("table:number-columns-repeated") = cv.second->range.c2 - cv.first + 1;
          col = cv.second->range.c2 + 1;
        }
        if (!gap.first_child()) gap.append_child("table:table-cell");
      }
    };


    size_t ci = 0;
    int32_t next_row = 1;
    R_xlen_t iter = 0;
    for (int32_t r : row_numbers) {
      checkInterrupt(iter);
      if (r > next_row) empty_rows(next_row, r - 1);
      next_row = r + 1;

      auto attr = rows.find(r);
      int32_t row_level = attr != rows.end() ? attr->second.level : 0;
      pugi::xml_node row = row_container(r, row_level).append_child("table:table-row");
      bool page_break = row_breaks.count(r) > 0;
      if (attr != rows.end() && !attr->second.height.empty()) {
        row.append_attribute("table:style-name") = row_style(attr->second.height, attr->second.custom, page_break).c_str();
      } else {
        row.append_attribute("table:style-name") = page_break ? row_style(default_height_pt, false, true).c_str() : default_row.c_str();
      }
      if (attr != rows.end()) {
        if (attr->second.hidden) row.append_attribute("table:visibility") = "collapse";
        if (attr->second.style > 0) row.append_attribute("table:default-cell-style-name") = ("ce" + std::to_string(attr->second.style)).c_str();
      }

      // slots of this row: cells, merge anchors and covered cells
      struct ods_slot {
        int64_t cell = -1;
        bool covered = false;
        int32_t col_span = 1;
        int32_t row_span = 1;
        const std::string* valid = nullptr;
      };
      std::map<int32_t, ods_slot> slots;
      while (ci < cells.size() && cells[ci].row < r) ++ci;
      for (size_t cj = ci; cj < cells.size() && cells[cj].row == r; ++cj) slots[cells[cj].col].cell = static_cast<int64_t>(cj);
      for (const auto& cm : comments) {
        if (cm.first.first == r) slots[cm.first.second];
      }
      for (const ods_frame& fr : frames) {
        if (fr.row == r) slots[fr.col];
      }
      for (const ods_valid& v : valids) {
        if (r < v.range.r1 || r > v.range.r2) continue;
        for (int64_t c = v.range.c1; c <= v.range.c2; ++c) slots[static_cast<int32_t>(c)].valid = &v.name;
      }
      for (const ods_merge& m : merges) {
        if (r < m.r1 || r > m.r2) continue;
        for (int32_t c = m.c1; c <= m.c2; ++c) {
          ods_slot& slot = slots[c];
          if (r == m.r1 && c == m.c1) {
            slot.col_span = m.c2 - m.c1 + 1;
            slot.row_span = m.r2 - m.r1 + 1;
          } else {
            slot.covered = true;
          }
        }
      }
      // the cells of a formatted row without own cell have the style of the row up to the end of the sheet
      std::string row_cell_style = attr != rows.end() && attr->second.style > 0 ? "ce" + std::to_string(attr->second.style) : "";
      if (slots.empty() && row_cell_style.empty()) row.append_child("table:table-cell");

      int32_t col = 1;
      for (auto it = slots.begin(); it != slots.end(); ++it) {
        const auto& entry = *it;
        if (entry.first < col) continue;
        if (entry.first > col) {
          pugi::xml_node gap = row.append_child("table:table-cell");
          if (!row_cell_style.empty()) gap.append_attribute("table:style-name") = row_cell_style.c_str();
          if (entry.first - col > 1) gap.append_attribute("table:number-columns-repeated") = entry.first - col;
        }
        col = entry.first + 1;

        const ods_slot& slot = entry.second;
        auto is_empty = [&](const ods_slot& sl, int32_t c) {
          return sl.cell < 0 && !sl.covered && sl.col_span == 1 && sl.row_span == 1 &&
                 comments.find(std::make_pair(r, c)) == comments.end() &&
                 std::none_of(frames.begin(), frames.end(), [&](const ods_frame& fr) { return fr.row == r && fr.col == c; });
        };
        // consecutive empty cells of a validation are one repeated cell
        if (slot.valid && is_empty(slot, entry.first)) {
          int32_t count = 1;
          for (auto next = std::next(it); next != slots.end() && next->first == entry.first + count &&
               next->second.valid == slot.valid && is_empty(next->second, next->first); ++next) ++count;
          pugi::xml_node cell = row.append_child("table:table-cell");
          cell.append_attribute("table:content-validation-name") = slot.valid->c_str();
          if (count > 1) cell.append_attribute("table:number-columns-repeated") = count;
          col = entry.first + count;
          continue;
        }
        pugi::xml_node cell = row.append_child(slot.covered ? "table:covered-table-cell" : "table:table-cell");
        if (slot.valid) cell.append_attribute("table:content-validation-name") = slot.valid->c_str();
        if (!slot.covered && slot.col_span > 1) cell.append_attribute("table:number-columns-spanned") = slot.col_span;
        if (!slot.covered && slot.row_span > 1) cell.append_attribute("table:number-rows-spanned") = slot.row_span;

        auto comment = comments.find(std::make_pair(r, entry.first));
        if (comment != comments.end()) {
          pugi::xml_node annotation = cell.append_child("office:annotation");
          if (!comment->second.author.empty()) annotation.append_child("dc:creator").text().set(comment->second.author.c_str());
          const std::vector<std::pair<std::string, std::string>>& runs = comment->second.runs;
          for (size_t k = 0; k < runs.size(); ++k) {
            pugi::xml_document rpr_doc;
            rpr_doc.load_string(runs[k].second.c_str());
            ods_cell_text(annotation, runs[k].first, text_style(rpr_doc.child("rPr")), k > 0);
          }
          if (!annotation.child("text:p")) annotation.append_child("text:p");
        }
        for (const ods_frame& fr : frames) {
          if (fr.row == r && fr.col == entry.first) ods_append_frame(cell, fr);
        }
        if (slot.cell < 0) continue;

        const ods_out_cell& oc = cells[static_cast<size_t>(slot.cell)];
        size_t i = oc.idx;
        std::string cell_style = cf_maps.empty() ? (oc.style > 0 ? "ce" + std::to_string(oc.style) : "") : cell_style_name(oc.style, r, entry.first);

        // a merged cell has the right border of its top right cell and the bottom border of its bottom left cell
        if (slot.col_span > 1 || slot.row_span > 1) {
          auto style_at = [&](int32_t edge_row, int32_t edge_col) {
            auto found = std::lower_bound(cells.begin(), cells.end(), std::make_pair(edge_row, edge_col), [](const ods_out_cell& a, const std::pair<int32_t, int32_t>& b) {
              return a.row < b.first || (a.row == b.first && a.col < b.second);
            });
            return found != cells.end() && found->row == edge_row && found->col == edge_col ? found->style : 0;
          };
          int32_t right_style = style_at(r, entry.first + slot.col_span - 1);
          int32_t bottom_style = style_at(r + slot.row_span - 1, entry.first);
          pugi::xml_node source, right_source, bottom_source;
          for (pugi::xml_node st : autos.children("style:style")) {
            std::string name = st.attribute("style:name").as_string();
            if (name == cell_style) source = st;
            if (right_style > 0 && name == "ce" + std::to_string(right_style)) right_source = st;
            if (bottom_style > 0 && name == "ce" + std::to_string(bottom_style)) bottom_source = st;
          }
          pugi::xml_node own = source.child("style:table-cell-properties");
          std::string right = slot.col_span > 1 ? right_source.child("style:table-cell-properties").attribute("fo:border-right").as_string() : "";
          std::string bottom = slot.row_span > 1 ? bottom_source.child("style:table-cell-properties").attribute("fo:border-bottom").as_string() : "";
          if (own.attribute("fo:border-right") || own.attribute("fo:border")) right.clear();
          if (own.attribute("fo:border-bottom") || own.attribute("fo:border")) bottom.clear();
          if (!right.empty() || !bottom.empty()) {
            std::string key = cell_style + "|" + right + "|" + bottom;
            auto known = map_styles.find(key);
            if (known != map_styles.end()) {
              cell_style = known->second;
            } else {
              pugi::xml_node copy = source ? autos.append_copy(source) : autos.append_child("style:style");
              std::string name = "cm" + std::to_string(++map_style_count);
              if (!source) {
                copy.append_attribute("style:name");
                copy.append_attribute("style:family") = "table-cell";
                copy.append_attribute("style:parent-style-name") = "Default";
              }
              copy.attribute("style:name") = name.c_str();
              pugi::xml_node tcp = copy.child("style:table-cell-properties");
              if (!tcp) tcp = copy.prepend_child("style:table-cell-properties");
              if (!right.empty()) tcp.append_attribute("fo:border-right") = right.c_str();
              if (!bottom.empty()) tcp.append_attribute("fo:border-bottom") = bottom.c_str();
              map_styles[key] = name;
              cell_style = name;
            }
          }
        }
        if (!cell_style.empty()) cell.append_attribute("table:style-name") = cell_style.c_str();

        const std::string& type = cc_t[i];
        std::string v = ods_unescape(cc_v[i]);
        ods_formula_ctx ctx;
        ctx.tables = &tables_info;
        ctx.row = oc.row;
        ctx.col = oc.col;
        ctx.sheet = sheet_quoted;
        std::string formula;
        pugi::xml_document fattr_doc;
        if (!cc_fattr[i].empty() && fattr_doc.load_string(("<f " + cc_fattr[i] + "/>").c_str())) {
          pugi::xml_node fa = fattr_doc.child("f");
          std::string t = fa.attribute("t").as_string();
          std::string ref = fa.attribute("ref").as_string();
          if (t == "array" && !ref.empty() && !cc_f[i].empty()) {
            size_t colon = ref.find(':');
            std::string to = colon == std::string::npos ? ref : ref.substr(colon + 1);
            cell.append_attribute("table:number-matrix-columns-spanned") = cell_to_colint(to) - oc.col + 1;
            cell.append_attribute("table:number-matrix-rows-spanned") = cell_to_rowint(to) - oc.row + 1;
          }
          // the cells of a shared formula use the formula of the first cell, moved to their position
          auto anchor = shared.find(fa.attribute("si").as_string());
          if (t == "shared" && cc_f[i].empty() && anchor != shared.end()) {
            formula = ods_formula_to_of(anchor->second.f, oc.row - anchor->second.row, oc.col - anchor->second.col, ctx);
          }
        }
        if (!cc_f[i].empty()) formula = ods_formula_to_of(ods_unescape(cc_f[i]), 0, 0, ctx);
        if (!formula.empty()) cell.append_attribute("table:formula") = formula.c_str();

        bool write_value = formula.empty() || !v.empty();

        if (!write_value) {
        } else if (type == "s" || type == "inlineStr") {
          cell.append_attribute("office:value-type") = "string";
          if (type == "s") {
            size_t si = static_cast<size_t>(std::strtoul(v.c_str(), nullptr, 10));
            if (si < shared_strings.size()) string_text(cell, shared_strings[si]);
          } else {
            string_text(cell, cc_is[i]);
          }
          if (!cell.child("text:p")) cell.append_child("text:p");
        } else if (type == "str" || type == "e") {
          // errors are strings with an empty value and the error as text
          cell.append_attribute("office:value-type") = "string";
          cell.append_attribute("office:string-value") = type == "e" ? "" : v.c_str();
          if (type == "e") cell.append_attribute("calcext:value-type") = "error";
          ods_cell_text(cell, v);
        } else if (type == "b") {
          bool val = v == "1" || v == "true";
          cell.append_attribute("office:value-type") = "boolean";
          cell.append_attribute("office:boolean-value") = val ? "true" : "false";
          ods_cell_text(cell, val ? "TRUE" : "FALSE");
        } else if (type == "d") {
          cell.append_attribute("office:value-type") = "date";
          cell.append_attribute("office:date-value") = v.substr(0, v.find_first_of("Z+")).c_str();
        } else if (!v.empty()) {
          double val = std::strtod(v.c_str(), nullptr);
          ods_kind kind = static_cast<size_t>(oc.style) < n_xfs ? xf_kind[static_cast<size_t>(oc.style)] : ods_kind_number;
          if (kind == ods_kind_date) {
            cell.append_attribute("office:value-type") = "date";
            cell.append_attribute("office:date-value") = ods_date_value(val, date1904).c_str();
          } else if (kind == ods_kind_time && val >= 0.0) {
            cell.append_attribute("office:value-type") = "time";
            cell.append_attribute("office:time-value") = ods_time_value(val).c_str();
          } else {
            cell.append_attribute("office:value-type") = kind == ods_kind_percent ? "percentage" : kind == ods_kind_currency ? "currency" : "float";
            cell.append_attribute("office:value") = v.c_str();
          }
        }

        auto link = links.find(std::make_pair(r, entry.first));
        if (link != links.end()) {
          if (!cell.child("text:p")) ods_cell_text(cell, v);
          std::vector<pugi::xml_node> paragraphs;
          for (pugi::xml_node p : cell.children("text:p")) paragraphs.push_back(p);
          for (pugi::xml_node p : paragraphs) {
            pugi::xml_node a = cell.insert_child_after("text:p", p).append_child("text:a");
            a.append_attribute("xlink:href") = link->second.first.c_str();
            a.append_attribute("xlink:type") = "simple";
            if (!link->second.second.empty()) a.append_attribute("office:title") = link->second.second.c_str();
            for (pugi::xml_node child : p.children()) a.append_copy(child);
            cell.remove_child(p);
          }
        }
      }
      if (!row_cell_style.empty() && col <= MAX_OOXML_COL_INT) {
        pugi::xml_node rest = row.append_child("table:table-cell");
        rest.append_attribute("table:style-name") = row_cell_style.c_str();
        if (col < MAX_OOXML_COL_INT) rest.append_attribute("table:number-columns-repeated") = MAX_OOXML_COL_INT - col + 1;
      }
    }
    // the rows after the last one have the default height of the sheet
    if (next_row <= MAX_OOXML_ROW_INT) empty_rows(next_row, MAX_OOXML_ROW_INT);

    // conditional formats, one for each range, with the rules in the order of their priority
    const std::vector<utf8_string>& cf_sqref = cf_sqref_list;
    const std::vector<utf8_string>& cf_rule = cf_rule_list;
    std::vector<std::string> cf_order;
    std::map<std::string, std::vector<std::pair<int64_t, std::string>>> cf_rules;
    for (size_t k = 0; k < cf_sqref.size() && k < cf_rule.size(); ++k) {
      pugi::xml_document rd;
      if (!rd.load_string(cf_rule[k].c_str())) continue;
      if (cf_rules.find(cf_sqref[k]) == cf_rules.end()) cf_order.push_back(cf_sqref[k]);
      cf_rules[cf_sqref[k]].push_back(std::make_pair(rd.child("cfRule").attribute("priority").as_llong(0), std::string(cf_rule[k])));
    }
    if (!cf_order.empty()) {
      pugi::xml_node formats = table.append_child("calcext:conditional-formats");
      for (const std::string& sqref : cf_order) {
        std::vector<ods_range> ranges = ods_sqref(sqref);
        if (ranges.empty()) continue;
        std::string target;
        for (const ods_range& range : ranges) target += (target.empty() ? "" : " ") + ods_range_address(sheet_quoted, range);
        pugi::xml_node format = formats.append_child("calcext:conditional-format");
        format.append_attribute("calcext:target-range-address") = target.c_str();
        std::vector<std::pair<int64_t, std::string>> sorted = cf_rules[sqref];
        std::stable_sort(sorted.begin(), sorted.end(), [](const std::pair<int64_t, std::string>& a, const std::pair<int64_t, std::string>& b) {
          return a.first < b.first;
        });
        ods_formula_ctx fctx;
        fctx.tables = &tables_info;
        fctx.sheet = sheet_quoted;
        fctx.row = ranges[0].r1;
        fctx.col = ranges[0].c1;
        for (const auto& rule : sorted) {
          pugi::xml_document rd;
          rd.load_string(rule.second.c_str());
          ods_conditional_format(format, rd.child("cfRule"), ranges[0], sheet_quoted, fctx, theme_colors, palette, bar_extensions);
        }
        if (!format.first_child()) formats.remove_child(format);
      }
      if (!formats.first_child()) table.remove_child(formats);
    }

    // sparkline groups of the extension list
    {
      auto local = [](pugi::xml_node node) {
        const char* name = node.name();
        const char* colon = std::strchr(name, ':');
        return std::string(colon ? colon + 1 : name);
      };
      static const char* const flags[][2] = {
        {"dateAxis", "calcext:date-axis"}, {"markers", "calcext:markers"}, {"high", "calcext:high"}, {"low", "calcext:low"},
        {"first", "calcext:first"}, {"last", "calcext:last"}, {"negative", "calcext:negative"},
        {"displayXAxis", "calcext:display-x-axis"}, {"displayHidden", "calcext:display-hidden"}, {"rightToLeft", "calcext:right-to-left"}
      };
      static const char* const colors[][2] = {
        {"colorSeries", "calcext:color-series"}, {"colorNegative", "calcext:color-negative"}, {"colorAxis", "calcext:color-axis"},
        {"colorMarkers", "calcext:color-markers"}, {"colorFirst", "calcext:color-first"}, {"colorLast", "calcext:color-last"},
        {"colorHigh", "calcext:color-high"}, {"colorLow", "calcext:color-low"}
      };
      pugi::xml_document ext_doc;
      pugi::xml_node groups;
      for (pugi::xml_node ext : ods_load_list(ext_doc, Rcpp::as<std::vector<utf8_string>>(sheet["ext"]))) {
        if (std::strcmp(ext.attribute("uri").as_string(), "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}") != 0) continue;
        for (pugi::xml_node group : ext.first_child().children()) {
          if (local(group) != "sparklineGroup") continue;
          pugi::xml_node out = pugi::xml_node();
          for (pugi::xml_node line : group.children()) {
            if (local(line) != "sparklines") continue;
            for (pugi::xml_node sparkline : line.children()) {
              std::string data, cell;
              for (pugi::xml_node part : sparkline.children()) {
                if (local(part) == "f") data = ods_chart_range(part.text().get());
                if (local(part) == "sqref") cell = part.text().get();
              }
              std::vector<ods_range> at = ods_sqref(cell);
              if (data.empty() || at.empty()) continue;
              if (!out) {
                if (!groups) groups = table.append_child("calcext:sparkline-groups");
                out = groups.append_child("calcext:sparkline-group");
                out.append_attribute("calcext:type") = group.attribute("type").as_string("line");
                out.append_attribute("calcext:line-width") = (ods_num(group.attribute("lineWeight").as_double(0.75)) + "pt").c_str();
                for (const auto& flag : flags) {
                  if (group.attribute(flag[0]).as_bool()) out.append_attribute(flag[1]) = "true";
                }
                out.append_attribute("calcext:display-empty-cells-as") = group.attribute("displayEmptyCellsAs").as_string("zero");
                out.append_attribute("calcext:min-axis-type") = group.attribute("minAxisType").as_string("individual");
                out.append_attribute("calcext:max-axis-type") = group.attribute("maxAxisType").as_string("individual");
                if (group.attribute("manualMin")) out.append_attribute("calcext:manual-min") = group.attribute("manualMin").as_string();
                if (group.attribute("manualMax")) out.append_attribute("calcext:manual-max") = group.attribute("manualMax").as_string();
                for (const auto& color : colors) {
                  for (pugi::xml_node child : group.children()) {
                    if (local(child) != color[0]) continue;
                    std::string hex = ods_color(child, theme_colors, palette);
                    if (!hex.empty()) out.append_attribute(color[1]) = hex.c_str();
                  }
                }
                out.append_child("calcext:sparklines");
              }
              pugi::xml_node node = out.child("calcext:sparklines").append_child("calcext:sparkline");
              node.append_attribute("calcext:cell-address") = ods_range_address(sheet_quoted, ods_range{at[0].r1, at[0].c1, at[0].r1, at[0].c1}).c_str();
              node.append_attribute("calcext:data-range") = data.c_str();
            }
          }
        }
      }
    }

    // the autofilter of the sheet as anonymous database range
    pugi::xml_document af_doc;
    af_doc.load_string(Rcpp::as<utf8_string>(sheet["autofilter"]).c_str());
    std::vector<ods_range> af = ods_sqref(af_doc.child("autoFilter").attribute("ref").as_string());
    if (!af.empty()) {
      autofilters.push_back(std::make_pair("__Anonymous_Sheet_DB__" + std::to_string(s), ods_range_address(sheet_quoted, af[0])));
    }

    pugi::xml_document sheet_rels_doc;
    for (pugi::xml_node rel : ods_load_list(sheet_rels_doc, Rcpp::as<std::vector<utf8_string>>(sheet["rels"]))) {
      std::string type = rel.attribute("Type").as_string(), target = rel.attribute("Target").as_string();
      const std::string pivot_type = "/pivotTable";
      if (type.size() < pivot_type.size() || type.compare(type.size() - pivot_type.size(), pivot_type.size(), pivot_type) != 0) continue;
      for (size_t k = 0; k < pivot_nodes.size(); ++k) {
        if (target != "../pivotTables/pivotTable" + std::to_string(k + 1) + ".xml" || k >= pivot_rels.size()) continue;
        std::string cache_target = pivot_rels[k].child("Relationship").attribute("Target").as_string();
        for (size_t m = 0; m < cache_nodes.size(); ++m) {
          if (cache_target != "../pivotCache/pivotCacheDefinition" + std::to_string(m + 1) + ".xml") continue;
          std::vector<utf8_string> first_values;
          if (m < static_cast<size_t>(pivot_first.size())) first_values = Rcpp::as<std::vector<utf8_string>>(pivot_first[static_cast<R_xlen_t>(m)]);
          ods_data_pilot(pilots, pivot_nodes[k], cache_nodes[m], sheet_quoted, tables_info, first_values);
        }
      }
    }
  }

  if (!validations.first_child()) spreadsheet.remove_child(validations);

  // defined names, without the names of the spreadsheet software (_xlnm.)
  pugi::xml_document names_doc;
  std::vector<pugi::xml_node> name_nodes = ods_load_list(names_doc, defined_names);
  pugi::xml_node global_names;
  std::vector<pugi::xml_node> local_names(tables.size());
  for (pugi::xml_node dn : name_nodes) {
    std::string name = dn.attribute("name").as_string();
    if (name.empty() || name.compare(0, 6, "_xlnm.") == 0) continue;
    pugi::xml_node container;
    if (pugi::xml_attribute local = dn.attribute("localSheetId")) {
      auto pos = std::find(sheet_index.begin(), sheet_index.end(), local.as_int());
      if (pos == sheet_index.end()) continue;
      size_t k = static_cast<size_t>(pos - sheet_index.begin());
      if (!local_names[k]) local_names[k] = tables[k].append_child("table:named-expressions");
      container = local_names[k];
    } else {
      if (!global_names) global_names = spreadsheet.append_child("table:named-expressions");
      container = global_names;
    }
    ods_formula_ctx ctx;
    ctx.tables = &tables_info;
    std::string expression = ods_formula_to_of(dn.text().get(), 0, 0, ctx);
    bool range = expression.compare(0, 5, "of:=[") == 0 && expression.back() == ']' &&
                 expression.find('[', 5) == std::string::npos && expression.find("#REF!") == std::string::npos;
    if (range) {
      pugi::xml_node nr = container.append_child("table:named-range");
      nr.append_attribute("table:name") = name.c_str();
      nr.append_attribute("table:cell-range-address") = expression.substr(5, expression.size() - 6).c_str();
    } else {
      pugi::xml_node ne = container.append_child("table:named-expression");
      ne.append_attribute("table:name") = name.c_str();
      ne.append_attribute("table:expression") = expression.c_str();
    }
  }

  // autofilters as anonymous and tables as named database ranges, after the defined names
  if (!autofilters.empty() || !tables_info.empty()) {
    pugi::xml_node database_ranges = spreadsheet.append_child("table:database-ranges");
    for (const auto& af : autofilters) {
      pugi::xml_node db = database_ranges.append_child("table:database-range");
      db.append_attribute("table:name") = af.first.c_str();
      db.append_attribute("table:target-range-address") = af.second.c_str();
      db.append_attribute("table:display-filter-buttons") = "true";
    }
    for (const ods_table& t : tables_info) {
      pugi::xml_node db = database_ranges.append_child("table:database-range");
      db.append_attribute("table:name") = t.name.c_str();
      ods_range range = {t.r1, t.c1, t.r2, t.c2};
      db.append_attribute("table:target-range-address") = ods_range_address(t.sheet, range).c_str();
      if (t.filter) db.append_attribute("table:display-filter-buttons") = "true";
      if (t.headers == 0) db.append_attribute("table:contains-header") = "false";
    }
  }
  if (pilots.first_child()) spreadsheet.append_copy(pilots);
  if (!font_faces.first_child()) root.remove_child(font_faces);
  ods_save(content, path + "content.xml");

  // the default style is xf 0
  pugi::xml_document styles;
  ods_declaration(styles);
  pugi::xml_node sroot = styles.append_child("office:document-styles");
  for (const auto& ns : ods_namespaces) sroot.append_attribute(ns[0]) = ns[1];
  sroot.append_attribute("office:version") = version.c_str();
  pugi::xml_node office_styles = sroot.append_child("office:styles");
  pugi::xml_node def = office_styles.append_child("style:style");
  def.append_attribute("style:name") = "Default";
  def.append_attribute("style:family") = "table-cell";
  if (n_xfs > 0) {
    pugi::xml_node xf = xf_nodes[0];
    size_t font_id = xf.attribute("fontId").as_uint(), fill_id = xf.attribute("fillId").as_uint(), border_id = xf.attribute("borderId").as_uint();
    ods_cell_style(def, xf,
                   font_id < font_nodes.size() ? font_nodes[font_id] : empty,
                   fill_id < fill_nodes.size() ? fill_nodes[fill_id] : empty,
                   border_id < border_nodes.size() ? border_nodes[border_id] : empty,
                   theme_colors, palette);
  }

  for (pugi::xml_node gradient : gradient_doc.children()) office_styles.append_copy(gradient);

  // the differential formats of conditional formats as named styles
  pugi::xml_document dxf_doc;
  std::vector<utf8_string> dxf_list = Rcpp::as<std::vector<utf8_string>>(book["dxfs"]);
  dxf_list.insert(dxf_list.end(), extra_dxfs.begin(), extra_dxfs.end());
  std::vector<pugi::xml_node> dxf_nodes = ods_load_list(dxf_doc, dxf_list);
  for (size_t k = 0; k < dxf_nodes.size(); ++k) {
    pugi::xml_node cs = office_styles.append_child("style:style");
    cs.append_attribute("style:name") = ("ConditionalStyle_" + std::to_string(k + 1)).c_str();
    cs.append_attribute("style:family") = "table-cell";
    cs.append_attribute("style:parent-style-name") = "Default";
    pugi::xml_node dxf = dxf_nodes[k];
    ods_cell_style(cs, dxf, dxf.child("font"), dxf.child("fill"), dxf.child("border"), theme_colors, palette, true);
  }

  // page layouts with headers and footers. xlsx header margins are the
  // distance to the page edge, ODF places the header inside the page margin.
  pugi::xml_node sautos = sroot.append_child("office:automatic-styles");
  // the text styles of the formatted parts of headers and footers
  std::map<std::string, std::string> header_styles;
  pugi::xml_node masters = sroot.append_child("office:master-styles");
  for (size_t k = 0; k < pages.size(); ++k) {
    const ods_page& page = pages[k];
    // a variant (odd, even or first page) has a header or footer if one of its parts is not empty
    auto has_part = [&](size_t variant, size_t hf) {
      return !page.hf[variant * 6 + hf * 3].empty() || !page.hf[variant * 6 + hf * 3 + 1].empty() || !page.hf[variant * 6 + hf * 3 + 2].empty();
    };
    bool has_header = has_part(0, 0) || has_part(1, 0) || has_part(2, 0);
    bool has_footer = has_part(0, 1) || has_part(1, 1) || has_part(2, 1);
    char buf[32];
    std::string pm = "pm" + std::to_string(k + 1);
    pugi::xml_node layout = sautos.append_child("style:page-layout");
    layout.append_attribute("style:name") = pm.c_str();
    pugi::xml_node props = layout.append_child("style:page-layout-properties");
    if (page.paper_w > 0.0) {
      std::snprintf(buf, sizeof(buf), "%.2fmm", page.landscape ? page.paper_h : page.paper_w);
      props.append_attribute("fo:page-width") = buf;
      std::snprintf(buf, sizeof(buf), "%.2fmm", page.landscape ? page.paper_w : page.paper_h);
      props.append_attribute("fo:page-height") = buf;
    }
    props.append_attribute("style:print-orientation") = page.landscape ? "landscape" : "portrait";
    double margins[4] = {has_header ? page.header : page.top, has_footer ? page.footer : page.bottom, page.left, page.right};
    const char* margin_names[4] = {"fo:margin-top", "fo:margin-bottom", "fo:margin-left", "fo:margin-right"};
    for (size_t m = 0; m < 4; ++m) {
      std::snprintf(buf, sizeof(buf), "%.4fin", margins[m]);
      props.append_attribute(margin_names[m]) = buf;
    }
    if (page.fit) {
      if (page.fit_x > 0) props.append_attribute("style:scale-to-X") = page.fit_x;
      if (page.fit_y > 0) props.append_attribute("style:scale-to-Y") = page.fit_y;
    } else if (page.scale != 100) {
      props.append_attribute("style:scale-to") = (std::to_string(page.scale) + "%").c_str();
    }
    std::string print = std::string(page.comments ? "annotations " : "") + "charts drawings " + (page.grid ? "grid " : "") +
                        (page.headings ? "headers " : "") + "objects zero-values";
    props.append_attribute("style:print") = print.c_str();
    if (page.across) props.append_attribute("style:print-page-order") = "ltr";
    if (page.first_page > 0) props.append_attribute("style:first-page-number") = std::to_string(page.first_page).c_str();
    if (page.center_h || page.center_v) {
      props.append_attribute("style:table-centering") = page.center_h && page.center_v ? "both" : page.center_h ? "horizontal" : "vertical";
    }
    for (int32_t hf = 0; hf < 2; ++hf) {
      if (!(hf == 0 ? has_header : has_footer)) continue;
      pugi::xml_node hf_props = layout.append_child(hf == 0 ? "style:header-style" : "style:footer-style").append_child("style:header-footer-properties");
      std::snprintf(buf, sizeof(buf), "%.4fin", std::max(hf == 0 ? page.top - page.header : page.bottom - page.footer, 0.0));
      hf_props.append_attribute("fo:min-height") = buf;
      hf_props.append_attribute(hf == 0 ? "fo:margin-bottom" : "fo:margin-top") = "0in";
    }

    pugi::xml_node master = masters.append_child("style:master-page");
    master.append_attribute("style:name") = ("mp" + std::to_string(k + 1)).c_str();
    master.append_attribute("style:page-layout-name") = pm.c_str();
    static const char* const regions[3] = {"style:region-left", "style:region-center", "style:region-right"};
    // odd pages are right pages, even pages left pages. With different even or first pages both their
    // header and footer are written, without them consumers use those of the odd pages.
    static const char* const variants[3][2] = {
      {"style:header", "style:footer"}, {"style:header-left", "style:footer-left"}, {"style:header-first", "style:footer-first"}
    };
    bool different[3] = {true, has_part(1, 0) || has_part(1, 1), has_part(2, 0) || has_part(2, 1)};
    for (size_t hf = 0; hf < 2; ++hf) {
      if (!(hf == 0 ? has_header : has_footer)) continue;
      for (size_t variant = 0; variant < 3; ++variant) {
        if (!different[variant]) continue;
        pugi::xml_node node = master.append_child(variants[variant][hf]);
        for (size_t region = 0; region < 3; ++region) {
          const std::string& code = page.hf[variant * 6 + hf * 3 + region];
          if (!code.empty()) ods_header_region(node.append_child(regions[region]), code, sautos, header_styles);
        }
      }
    }
  }
  ods_save(styles, path + "styles.xml");

  // document properties
  pugi::xml_document meta, core_doc;
  ods_declaration(meta);
  pugi::xml_node mroot = meta.append_child("office:document-meta");
  mroot.append_attribute("xmlns:office") = ods_namespaces[0][1];
  mroot.append_attribute("xmlns:meta") = ods_namespaces[7][1];
  mroot.append_attribute("xmlns:dc") = "http://purl.org/dc/elements/1.1/";
  mroot.append_attribute("office:version") = version.c_str();
  pugi::xml_node office_meta = mroot.append_child("office:meta");
  office_meta.append_child("meta:generator").text().set("openxlsx2");
  core_doc.load_string(Rcpp::as<utf8_string>(book["core"]).c_str());
  pugi::xml_node core = core_doc.child("cp:coreProperties");
  static const char* const properties[][2] = {
    {"dc:title", "dc:title"}, {"dc:subject", "dc:subject"}, {"dc:description", "dc:description"},
    {"cp:keywords", "meta:keyword"}, {"dc:creator", "meta:initial-creator"}, {"cp:lastModifiedBy", "dc:creator"},
    {"dcterms:created", "meta:creation-date"}, {"dcterms:modified", "dc:date"}
  };
  for (const auto& prop : properties) {
    std::string value = core.child_value(prop[0]);
    if (value.empty()) continue;
    if (std::strncmp(prop[0], "dcterms:", 8) == 0 && value.back() == 'Z') value.pop_back();
    office_meta.append_child(prop[1]).text().set(value.c_str());
  }
  if (*core.child_value("cp:category")) {
    pugi::xml_node category = office_meta.append_child("meta:user-defined");
    category.append_attribute("meta:name") = "Category";
    category.text().set(core.child_value("cp:category"));
  }
  ods_save(meta, path + "meta.xml");

  // view settings: freeze panes, grid lines, zoom and the active sheet
  pugi::xml_document settings, book_views_doc;
  ods_declaration(settings);
  pugi::xml_node set_root = settings.append_child("office:document-settings");
  set_root.append_attribute("xmlns:office") = ods_namespaces[0][1];
  set_root.append_attribute("xmlns:config") = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
  set_root.append_attribute("xmlns:ooo") = "http://openoffice.org/2004/office";
  set_root.append_attribute("office:version") = version.c_str();
  pugi::xml_node view_set = set_root.append_child("office:settings").append_child("config:config-item-set");
  view_set.append_attribute("config:name") = "ooo:view-settings";
  pugi::xml_node view_map = view_set.append_child("config:config-item-map-indexed");
  view_map.append_attribute("config:name") = "Views";
  pugi::xml_node view_entry = view_map.append_child("config:config-item-map-entry");
  ods_config_item(view_entry, "ViewId", "string", "view1");
  pugi::xml_node table_map = view_entry.append_child("config:config-item-map-named");
  table_map.append_attribute("config:name") = "Tables";
  // the left and top positions are the first column and row of the left and top part of a split view, the
  // right and bottom positions those of the right and bottom part
  for (size_t k = 0; k < views.size(); ++k) {
    const ods_view& view = views[k];
    pugi::xml_node te = table_map.append_child("config:config-item-map-entry");
    te.append_attribute("config:name") = sheet_names[k].c_str();
    ods_config_item(te, "CursorPositionX", "int", std::to_string(view.cursor_col));
    ods_config_item(te, "CursorPositionY", "int", std::to_string(view.cursor_row));
    ods_config_item(te, "HorizontalSplitMode", "short", view.cols > 0 ? (view.split ? "1" : "2") : "0");
    ods_config_item(te, "VerticalSplitMode", "short", view.rows > 0 ? (view.split ? "1" : "2") : "0");
    ods_config_item(te, "HorizontalSplitPosition", "int", std::to_string(view.cols));
    ods_config_item(te, "VerticalSplitPosition", "int", std::to_string(view.rows));
    ods_config_item(te, "ActiveSplitRange", "short", view.cols > 0 && view.rows > 0 ? "3" : view.cols > 0 ? "1" : "2");
    ods_config_item(te, "PositionLeft", "int", std::to_string(view.cols > 0 ? 0 : view.left));
    ods_config_item(te, "PositionRight", "int", std::to_string(view.cols > 0 ? view.pane_left : 0));
    ods_config_item(te, "PositionTop", "int", std::to_string(view.rows > 0 ? 0 : view.top));
    ods_config_item(te, "PositionBottom", "int", std::to_string(view.rows > 0 ? view.pane_top : view.top));
    ods_config_item(te, "ZoomType", "short", "0");
    ods_config_item(te, "ZoomValue", "int", std::to_string(view.zoom));
    ods_config_item(te, "PageViewZoomValue", "int", std::to_string(view.page_zoom));
    ods_config_item(te, "ShowGrid", "boolean", view.grid ? "true" : "false");
    ods_config_item(te, "ShowZeroValues", "boolean", view.zeros ? "true" : "false");
    ods_config_item(te, "HasColumnRowHeaders", "boolean", view.headers ? "true" : "false");
  }
  // the settings of the document view are those of the active sheet
  book_views_doc.load_string(Rcpp::as<utf8_string>(book["book_views"]).c_str());
  pugi::xml_node book_view = book_views_doc.child("bookViews").child("workbookView");
  int32_t active = book_view.attribute("activeTab").as_int(0);
  std::vector<int32_t>::const_iterator active_pos = std::find(sheet_index.begin(), sheet_index.end(), active);
  size_t active_sheet = active_pos != sheet_index.end() ? static_cast<size_t>(active_pos - sheet_index.begin()) : 0;
  if (active_sheet < views.size()) {
    const ods_view& view = views[active_sheet];
    ods_config_item(view_entry, "ActiveTable", "string", sheet_names[active_sheet]);
    ods_config_item(view_entry, "ZoomType", "short", "0");
    ods_config_item(view_entry, "ZoomValue", "int", std::to_string(view.zoom));
    ods_config_item(view_entry, "PageViewZoomValue", "int", std::to_string(view.page_zoom));
    ods_config_item(view_entry, "ShowPageBreakPreview", "boolean", view.page_break ? "true" : "false");
    ods_config_item(view_entry, "ShowZeroValues", "boolean", view.zeros ? "true" : "false");
    ods_config_item(view_entry, "ShowGrid", "boolean", view.grid ? "true" : "false");
    ods_config_item(view_entry, "HasColumnRowHeaders", "boolean", view.headers ? "true" : "false");
    ods_config_item(view_entry, "HasSheetTabs", "boolean", book_view.attribute("showSheetTabs").as_bool(true) ? "true" : "false");
    ods_config_item(view_entry, "IsOutlineSymbolsSet", "boolean", view.outline ? "true" : "false");
    if (!view.grid_color.empty()) {
      ods_config_item(view_entry, "GridColor", "int", std::to_string(std::strtol(view.grid_color.c_str() + 1, nullptr, 16)));
    }
  }
  ods_save(settings, path + "settings.xml");

  pugi::xml_document manifest;
  ods_declaration(manifest);
  pugi::xml_node man = manifest.append_child("manifest:manifest");
  man.append_attribute("xmlns:manifest") = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
  man.append_attribute("manifest:version") = version.c_str();
  static const char* const entries[][2] = {
    {"/", "application/vnd.oasis.opendocument.spreadsheet"},
    {"content.xml", "text/xml"},
    {"styles.xml", "text/xml"},
    {"meta.xml", "text/xml"},
    {"settings.xml", "text/xml"}
  };
  for (const auto& entry : entries) {
    pugi::xml_node fe = man.append_child("manifest:file-entry");
    fe.append_attribute("manifest:full-path") = entry[0];
    if (std::strcmp(entry[0], "/") == 0) fe.append_attribute("manifest:version") = version.c_str();
    fe.append_attribute("manifest:media-type") = entry[1];
  }
  for (const std::string& picture : pictures) {
    pugi::xml_node fe = man.append_child("manifest:file-entry");
    fe.append_attribute("manifest:full-path") = picture.c_str();
    fe.append_attribute("manifest:media-type") = ods_media_type(picture).c_str();
  }
  for (const std::string& object : objects) {
    pugi::xml_node fe = man.append_child("manifest:file-entry");
    fe.append_attribute("manifest:full-path") = (object + "/").c_str();
    fe.append_attribute("manifest:media-type") = "application/vnd.oasis.opendocument.chart";
    fe = man.append_child("manifest:file-entry");
    fe.append_attribute("manifest:full-path") = (object + "/content.xml").c_str();
    fe.append_attribute("manifest:media-type") = "text/xml";
    fe = man.append_child("manifest:file-entry");
    fe.append_attribute("manifest:full-path") = (object + "/styles.xml").c_str();
    fe.append_attribute("manifest:media-type") = "text/xml";
  }
  ods_save(manifest, path + "META-INF/manifest.xml");

  std::ofstream mimetype(path + "mimetype", std::ios::out | std::ios::binary);
  mimetype << "application/vnd.oasis.opendocument.spreadsheet";

  // the pictures and chart objects that are used
  std::vector<utf8_string> used(pictures.begin(), pictures.end());
  used.insert(used.end(), objects.begin(), objects.end());
  return used;
}

// reading: ODF -> pseudo xlsx parts

static void ods_text(pugi::xml_node node, std::string& out);

// the text of one node of a paragraph
static void ods_text_node(pugi::xml_node child, std::string& out) {
  if (child.type() == pugi::node_pcdata || child.type() == pugi::node_cdata) {
    bool space = false;
    for (const char* p = child.value(); *p; ++p) {
      if (*p == ' ' || *p == '\n' || *p == '\r' || *p == '\t') {
        if (!space) out += ' ';
        space = true;
      } else {
        out += *p;
        space = false;
      }
    }
  } else if (child.type() == pugi::node_element) {
    const char* name = child.name();
    if (std::strcmp(name, "text:s") == 0) out.append(static_cast<size_t>(child.attribute("text:c").as_uint(1)), ' ');
    else if (std::strcmp(name, "text:tab") == 0) out += '\t';
    else if (std::strcmp(name, "text:line-break") == 0) out += '\n';
    else if (std::strcmp(name, "text:note") == 0 || std::strcmp(name, "office:annotation") == 0 || std::strcmp(name, "text:ruby-text") == 0) return;
    else ods_text(child, out);
  }
}

static void ods_text(pugi::xml_node node, std::string& out) {
  for (pugi::xml_node child : node.children()) ods_text_node(child, out);
}

// the text of a paragraph as runs of the text styles of its spans
static void ods_text_runs(pugi::xml_node node, const std::string& style, std::vector<std::pair<std::string, std::string>>& runs) {
  for (pugi::xml_node child : node.children()) {
    const char* name = child.name();
    if (std::strcmp(name, "text:span") == 0) {
      ods_text_runs(child, child.attribute("text:style-name").as_string(style.c_str()), runs);
      continue;
    }
    if (std::strcmp(name, "text:a") == 0) {
      ods_text_runs(child, style, runs);
      continue;
    }
    std::string piece;
    ods_text_node(child, piece);
    if (piece.empty()) continue;
    if (!runs.empty() && runs.back().first == style) runs.back().second += piece;
    else runs.push_back(std::make_pair(style, piece));
  }
}

static std::string ods_cell_string(pugi::xml_node cell) {
  std::string out;
  bool first = true;
  for (pugi::xml_node p : cell.children("text:p")) {
    if (!first) out += '\n';
    ods_text(p, out);
    first = false;
  }
  return out;
}

// literal text in a format code
static std::string ods_code_literal(const std::string& text) {
  if (text.find_first_not_of(" -/:.,()$+!^&'~{}<>=") == std::string::npos) return text;
  std::string out = "\"";
  for (char c : text) {
    if (c == '"') out += "\\\"";
    else out += c;
  }
  return out + "\"";
}

static std::string ods_section_code(pugi::xml_node style) {
  std::string code;
  std::string name = style.name();
  bool truncate = std::string(style.attribute("number:truncate-on-overflow").as_string("true")) != "false";
  bool elapsed_done = false;

  std::string color = ods_lower(style.child("style:text-properties").attribute("fo:color").as_string());
  if (color == "#ff0000") code += "[Red]";
  else if (color == "#0000ff") code += "[Blue]";
  else if (color == "#00ff00") code += "[Green]";
  else if (color == "#000000") code += "[Black]";
  else if (color == "#ffffff") code += "[White]";
  else if (color == "#ffff00") code += "[Yellow]";
  else if (color == "#ff00ff") code += "[Magenta]";
  else if (color == "#00ffff") code += "[Cyan]";

  for (pugi::xml_node child : style.children()) {
    std::string part = child.name();
    bool long_style = std::string(child.attribute("number:style").as_string()) == "long";
    std::string token;

    if (part == "number:text") {
      std::string text = child.text().get();
      if (name == "number:percentage-style") {
        size_t pct = text.find('%');
        if (pct != std::string::npos) {
          if (pct > 0) code += ods_code_literal(text.substr(0, pct));
          code += "%";
          if (pct + 1 < text.size()) code += ods_code_literal(text.substr(pct + 1));
          continue;
        }
      }
      code += ods_code_literal(text);
    } else if (part == "number:number") {
      if (!child.attribute("number:decimal-places")) {
        code += "General";
        continue;
      }
      uint32_t min_int = child.attribute("number:min-integer-digits").as_uint(0);
      uint32_t decimals = child.attribute("number:decimal-places").as_uint(0);
      uint32_t min_dec = std::min(child.attribute("number:min-decimal-places").as_uint(decimals), decimals);
      std::string digits;
      if (child.attribute("number:grouping").as_bool()) {
        size_t len = std::max<size_t>(4, min_int);
        digits = std::string(len - min_int, '#') + std::string(min_int, '0');
        digits.insert(digits.size() - 3, ",");
      } else {
        digits = min_int == 0 ? "#" : std::string(min_int, '0');
      }
      // the scaling by thousands follows the last digits
      std::string scaling;
      for (double factor = child.attribute("number:display-factor").as_double(1.0); factor >= 1000.0; factor /= 1000.0) scaling += ",";
      if (decimals == 0) digits += scaling;
      // texts embedded in the integer digits at their position, counted in digits left of the decimal separator;
      // texts left of all digits come first
      std::vector<std::pair<int64_t, std::string>> embedded;
      for (pugi::xml_node text : child.children("number:embedded-text")) {
        embedded.push_back(std::make_pair(text.attribute("number:position").as_llong(0), ods_code_literal(text.text().get())));
      }
      std::stable_sort(embedded.begin(), embedded.end(), [](const std::pair<int64_t, std::string>& a, const std::pair<int64_t, std::string>& b) {
        return a.first < b.first;
      });
      std::string leading;
      for (const auto& text : embedded) {
        int64_t seen = 0;
        size_t pos = digits.size();
        while (pos > 0 && seen < text.first) {
          --pos;
          if (digits[pos] == '#' || digits[pos] == '0' || digits[pos] == '?') ++seen;
        }
        if (seen < text.first) leading = text.second + leading;
        else digits.insert(pos, text.second);
      }
      digits = leading + digits;
      if (decimals > 0) digits += "." + std::string(min_dec, '0') + std::string(decimals - min_dec, '#') + scaling;
      code += digits;
    } else if (part == "number:scientific-number") {
      uint32_t min_int = std::max(child.attribute("number:min-integer-digits").as_uint(1), 1u);
      uint32_t decimals = child.attribute("number:decimal-places").as_uint(0);
      code += std::string(min_int, '0');
      if (decimals > 0) code += "." + std::string(decimals, '0');
      code += "E+" + std::string(std::max(child.attribute("number:min-exponent-digits").as_uint(2), 1u), '0');
    } else if (part == "number:fraction") {
      if (child.attribute("number:min-integer-digits")) code += "# ";
      code += std::string(std::max(child.attribute("number:min-numerator-digits").as_uint(1), 1u), '?') + "/";
      if (child.attribute("number:denominator-value")) code += child.attribute("number:denominator-value").as_string();
      else code += std::string(std::max(child.attribute("number:min-denominator-digits").as_uint(1), 1u), '?');
    } else if (part == "number:currency-symbol") {
      code += "[$" + std::string(child.text().get()) + "]";
    } else if (part == "number:text-content") {
      code += "@";
    } else if (part == "number:fill-character") {
      code += "*" + std::string(child.text().get());
    } else if (part == "number:boolean") {
      code += "General";
    } else if (part == "number:day") {
      code += long_style ? "dd" : "d";
    } else if (part == "number:day-of-week") {
      code += long_style ? "dddd" : "ddd";
    } else if (part == "number:month") {
      bool textual = child.attribute("number:textual").as_bool();
      code += textual ? (long_style ? "mmmm" : "mmm") : (long_style ? "mm" : "m");
    } else if (part == "number:year") {
      code += long_style ? "yyyy" : "yy";
    } else if (part == "number:hours" || part == "number:minutes" || part == "number:seconds") {
      token = std::string(long_style ? 2u : 1u, part == "number:hours" ? 'h' : part == "number:minutes" ? 'm' : 's');
      if (!truncate && !elapsed_done) {
        token = "[" + token + "]";
        elapsed_done = true;
      }
      code += token;
      uint32_t decimals = child.attribute("number:decimal-places").as_uint(0);
      if (part == "number:seconds" && decimals > 0) code += "." + std::string(decimals, '0');
    } else if (part == "number:am-pm") {
      code += "AM/PM";
    }
  }
  return code;
}

static std::string ods_format_code(pugi::xml_node style, const std::map<std::string, pugi::xml_node>& data_styles) {
  std::string main = ods_section_code(style);
  std::vector<std::pair<std::string, std::string>> maps;
  for (pugi::xml_node map : style.children("style:map")) {
    std::string cond = map.attribute("style:condition").as_string();
    auto mapped = data_styles.find(map.attribute("style:apply-style-name").as_string());
    if (mapped == data_styles.end()) continue;
    maps.push_back(std::make_pair(cond, ods_section_code(mapped->second)));
  }
  if (maps.empty()) return main;
  if (maps.size() == 1 && maps[0].first == "value()>=0") return maps[0].second + ";" + main;
  if (maps.size() == 2 && maps[0].first == "value()>0" && maps[1].first == "value()<0") {
    return maps[0].second + ";" + maps[1].second + ";" + main;
  }
  // positive, negative and zero numbers and text
  if (maps.size() == 3 && maps[0].first == "value()>0" && maps[1].first == "value()<0" && maps[2].first == "value()=0") {
    return maps[0].second + ";" + maps[1].second + ";" + maps[2].second + ";" + main;
  }
  std::string code;
  for (size_t k = 0; k < maps.size() && k < 2; ++k) {
    std::string cond = maps[k].first;
    if (cond.compare(0, 7, "value()") == 0) cond = cond.substr(7);
    code += "[" + cond + "]" + maps[k].second + ";";
  }
  return code + main;
}

static int32_t ods_index(std::map<std::string, int32_t>& index, std::vector<std::string>& list, const std::string& xml) {
  auto it = index.find(xml);
  if (it != index.end()) return it->second;
  int32_t id = static_cast<int32_t>(list.size());
  index[xml] = id;
  list.push_back(xml);
  return id;
}

static std::string ods_rgb(const char* hex) {
  std::string x = hex;
  if (x.size() != 7 || x[0] != '#') return "";
  std::string out = "FF" + x.substr(1);
  for (char& c : out) c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
  return out;
}

struct ods_xlsx_styles {
  std::map<std::string, pugi::xml_node> cell, data, text;
  std::map<std::string, std::string> font_face;
  pugi::xml_node default_cell;

  std::map<std::string, int32_t> font_index, fill_index, border_index, xf_index, numfmt_index, style_xf;
  std::vector<std::string> font_list, fill_list, border_list, xf_list, numfmt_list;

  ods_xlsx_styles() {
    ods_index(fill_index, fill_list, "<fill><patternFill patternType=\"none\"/></fill>");
    ods_index(fill_index, fill_list, "<fill><patternFill patternType=\"gray125\"/></fill>");
    ods_index(border_index, border_list, "<border><left/><right/><top/><bottom/><diagonal/></border>");
  }

  std::vector<pugi::xml_node> chain(const std::string& name) {
    std::vector<pugi::xml_node> out;
    std::string cur = name;
    for (int32_t depth = 0; depth < 16 && !cur.empty(); ++depth) {
      auto it = cell.find(cur);
      if (it == cell.end()) break;
      out.push_back(it->second);
      cur = it->second.attribute("style:parent-style-name").as_string();
    }
    if (default_cell) out.push_back(default_cell);
    return out;
  }

  static pugi::xml_attribute prop(const std::vector<pugi::xml_node>& styles, const char* props, const char* attr) {
    for (pugi::xml_node st : styles) {
      pugi::xml_attribute a = st.child(props).attribute(attr);
      if (a) return a;
    }
    return pugi::xml_attribute();
  }

  static std::string border_side(const char* value, std::string& color) {
    std::string val = value;
    if (val.empty() || val == "none" || val == "hidden") return "";
    double width = 0.74;
    std::string kind = "solid";
    size_t pos = 0;
    while (pos < val.size()) {
      size_t end = val.find(' ', pos);
      if (end == std::string::npos) end = val.size();
      std::string token = val.substr(pos, end - pos);
      pos = end + 1;
      if (token.empty()) continue;
      if (token[0] == '#') color = ods_rgb(token.c_str());
      else if (std::isdigit(static_cast<unsigned char>(token[0])) || token[0] == '.') width = ods_length_pt(token.c_str());
      else if (token == "thin") width = 0.74;
      else if (token == "medium") width = 1.76;
      else if (token == "thick") width = 2.49;
      else kind = token;
    }
    if (kind == "none" || kind == "hidden") return "";
    if (kind == "double" || kind == "double-thin") return "double";
    if (kind == "dotted") return "dotted";
    if (kind == "dashed" || kind == "fine-dashed") return width >= 1.25 ? "mediumDashed" : "dashed";
    if (kind == "dash-dot") return width >= 1.25 ? "mediumDashDot" : "dashDot";
    if (kind == "dash-dot-dot") return width >= 1.25 ? "mediumDashDotDot" : "dashDotDot";
    if (width <= 0.1) return "hair";
    if (width < 1.25) return "thin";
    if (width < 2.25) return "medium";
    return "thick";
  }

  int32_t numfmt(const std::string& name) {
    auto it = data.find(name);
    if (it == data.end()) return 0;
    std::string code = ods_format_code(it->second, data);
    if (code.empty() || code == "General") return 0;
    for (uint32_t id = 1; id < 50; ++id) {
      const char* builtin = xls_builtin_format(id);
      if (builtin && code == builtin) return static_cast<int32_t>(id);
    }
    auto found = numfmt_index.find(code);
    if (found != numfmt_index.end()) return found->second;
    int32_t id = 164 + static_cast<int32_t>(numfmt_list.size());
    numfmt_index[code] = id;
    numfmt_list.push_back(code);
    return id;
  }

  // the font of the text properties of a style chain, name_tag is name for fonts and rFont for runs
  void font_properties(const std::vector<pugi::xml_node>& styles, pugi::xml_node font, const char* name_tag) {
  std::string weight = prop(styles, "style:text-properties", "fo:font-weight").as_string();
  if (weight == "bold" || std::atoi(weight.c_str()) >= 600) font.append_child("b");
  std::string font_style = prop(styles, "style:text-properties", "fo:font-style").as_string();
  if (font_style == "italic" || font_style == "oblique") font.append_child("i");
  std::string strike = prop(styles, "style:text-properties", "style:text-line-through-style").as_string();
  if (!strike.empty() && strike != "none") font.append_child("strike");
  if (prop(styles, "style:text-properties", "style:text-outline").as_bool()) font.append_child("outline");
  std::string shadow = prop(styles, "style:text-properties", "fo:text-shadow").as_string();
  if (!shadow.empty() && shadow != "none") font.append_child("shadow");
  std::string underline = prop(styles, "style:text-properties", "style:text-underline-style").as_string();
  if (!underline.empty() && underline != "none") {
    pugi::xml_node u = font.append_child("u");
    if (std::string(prop(styles, "style:text-properties", "style:text-underline-type").as_string()) == "double") u.append_attribute("val") = "double";
  }
  std::string position = prop(styles, "style:text-properties", "style:text-position").as_string();
  if (position.compare(0, 5, "super") == 0 || (!position.empty() && std::atof(position.c_str()) > 0.0)) {
    font.append_child("vertAlign").append_attribute("val") = "superscript";
  } else if (position.compare(0, 3, "sub") == 0 || (!position.empty() && std::atof(position.c_str()) < 0.0)) {
    font.append_child("vertAlign").append_attribute("val") = "subscript";
  }
  double size = ods_length_pt(prop(styles, "style:text-properties", "fo:font-size").as_string("10pt"));
  font.append_child("sz").append_attribute("val") = ods_num(size > 0.0 ? size : 10.0).c_str();
  std::string color = ods_rgb(prop(styles, "style:text-properties", "fo:color").as_string());
  if (!color.empty()) font.append_child("color").append_attribute("rgb") = color.c_str();
  std::string family;
  pugi::xml_attribute font_name = prop(styles, "style:text-properties", "style:font-name");
  if (font_name && font_face.count(font_name.as_string())) family = font_face[font_name.as_string()];
  if (family.empty()) family = prop(styles, "style:text-properties", "fo:font-family").as_string();
  if (family.empty() && font_name) family = font_name.as_string();
  family.erase(std::remove(family.begin(), family.end(), '\''), family.end());
  family.erase(std::remove(family.begin(), family.end(), '"'), family.end());
  if (family.empty()) family = "Liberation Sans";
  font.append_child(name_tag).append_attribute("val") = family.c_str();
  }

  // the run properties of a span of rich text in a cell of a cell style
  std::string run_properties(const std::string& span, const std::string& cell_style) {
    // cells without a style have the default cell style
    std::vector<pugi::xml_node> styles = chain(cell_style.empty() && cell.count("Default") ? std::string("Default") : cell_style);
    auto it = text.find(span);
    if (it != text.end()) styles.insert(styles.begin(), it->second);
    pugi::xml_document doc;
    pugi::xml_node rpr = doc.append_child("rPr");
    font_properties(styles, rpr, "rFont");
    return ods_xml_string(rpr);
  }

  int32_t xf(const std::string& name) {
    auto cached = style_xf.find(name);
    if (cached != style_xf.end()) return cached->second;

    std::vector<pugi::xml_node> styles = chain(name);
    pugi::xml_document doc;

    pugi::xml_node font = doc.append_child("font");
    font_properties(styles, font, "name");
    int32_t font_id = ods_index(font_index, font_list, ods_xml_string(font));

    int32_t fill_id = 0;
    std::string background = ods_rgb(prop(styles, "style:table-cell-properties", "fo:background-color").as_string());
    if (!background.empty()) {
      pugi::xml_node fill = doc.append_child("fill");
      pugi::xml_node pattern = fill.append_child("patternFill");
      pattern.append_attribute("patternType") = "solid";
      pattern.append_child("fgColor").append_attribute("rgb") = background.c_str();
      pattern.append_child("bgColor").append_attribute("indexed") = 64;
      fill_id = ods_index(fill_index, fill_list, ods_xml_string(fill));
    }

    static const char* const sides[4][2] = {
      {"left", "fo:border-left"}, {"right", "fo:border-right"}, {"top", "fo:border-top"}, {"bottom", "fo:border-bottom"}
    };
    pugi::xml_node border = doc.append_child("border");
    const char* all = prop(styles, "style:table-cell-properties", "fo:border").as_string();
    for (const auto& side : sides) {
      pugi::xml_attribute specific = prop(styles, "style:table-cell-properties", side[1]);
      std::string side_color = "FF000000";
      std::string kind = border_side(specific ? specific.as_string() : all, side_color);
      pugi::xml_node node = border.append_child(side[0]);
      if (kind.empty()) continue;
      node.append_attribute("style") = kind.c_str();
      node.append_child("color").append_attribute("rgb") = side_color.c_str();
    }
    pugi::xml_node diagonal = border.append_child("diagonal");
    std::string diagonal_color = "FF000000";
    std::string down = border_side(prop(styles, "style:table-cell-properties", "style:diagonal-tl-br").as_string(), diagonal_color);
    std::string up = border_side(prop(styles, "style:table-cell-properties", "style:diagonal-bl-tr").as_string(), diagonal_color);
    if (!down.empty() || !up.empty()) {
      diagonal.append_attribute("style") = (down.empty() ? up : down).c_str();
      diagonal.append_child("color").append_attribute("rgb") = diagonal_color.c_str();
      if (!down.empty()) border.append_attribute("diagonalDown") = 1;
      if (!up.empty()) border.append_attribute("diagonalUp") = 1;
    }
    int32_t border_id = ods_index(border_index, border_list, ods_xml_string(border));

    int32_t numfmt_id = 0;
    for (pugi::xml_node st : styles) {
      if (pugi::xml_attribute ds = st.attribute("style:data-style-name")) {
        numfmt_id = numfmt(ds.as_string());
        break;
      }
    }

    pugi::xml_node alignment = doc.append_child("alignment");
    std::string source = prop(styles, "style:table-cell-properties", "style:text-align-source").as_string();
    std::string text_align = prop(styles, "style:paragraph-properties", "fo:text-align").as_string();
    if (prop(styles, "style:table-cell-properties", "style:repeat-content").as_bool()) {
      alignment.append_attribute("horizontal") = "fill";
    } else if (source != "value-type") {
      if (text_align == "start" || text_align == "left") alignment.append_attribute("horizontal") = "left";
      else if (text_align == "center") alignment.append_attribute("horizontal") = "center";
      else if (text_align == "end" || text_align == "right") alignment.append_attribute("horizontal") = "right";
      else if (text_align == "justify") alignment.append_attribute("horizontal") = "justify";
    }
    std::string vertical = prop(styles, "style:table-cell-properties", "style:vertical-align").as_string();
    if (vertical == "top") alignment.append_attribute("vertical") = "top";
    else if (vertical == "middle") alignment.append_attribute("vertical") = "center";
    double margin = ods_length_pt(prop(styles, "style:paragraph-properties", "fo:margin-left").as_string("0pt"));
    if (margin >= 9.0) alignment.append_attribute("indent") = static_cast<int32_t>(std::lround(margin / 9.0));
    if (std::string(prop(styles, "style:table-cell-properties", "fo:wrap-option").as_string()) == "wrap") alignment.append_attribute("wrapText") = 1;
    if (prop(styles, "style:table-cell-properties", "style:shrink-to-fit").as_bool()) alignment.append_attribute("shrinkToFit") = 1;
    if (std::string(prop(styles, "style:table-cell-properties", "style:direction").as_string()) == "ttb") {
      alignment.append_attribute("textRotation") = 255;
    } else {
      int32_t angle = static_cast<int32_t>(std::lround(std::atof(prop(styles, "style:table-cell-properties", "style:rotation-angle").as_string("0"))));
      angle = ((angle % 360) + 360) % 360;
      if (angle > 0 && angle <= 90) alignment.append_attribute("textRotation") = angle;
      else if (angle >= 270) alignment.append_attribute("textRotation") = 450 - angle;
    }

    pugi::xml_node protection = doc.append_child("protection");
    std::string protect = prop(styles, "style:table-cell-properties", "style:cell-protect").as_string("protected");
    if (protect == "none") protection.append_attribute("locked") = 0;
    if (protect.find("hidden") != std::string::npos) protection.append_attribute("hidden") = 1;

    pugi::xml_node xf_node = doc.append_child("xf");
    xf_node.append_attribute("numFmtId") = numfmt_id;
    xf_node.append_attribute("fontId") = font_id;
    xf_node.append_attribute("fillId") = fill_id;
    xf_node.append_attribute("borderId") = border_id;
    xf_node.append_attribute("xfId") = 0;
    if (numfmt_id) xf_node.append_attribute("applyNumberFormat") = 1;
    if (font_id) xf_node.append_attribute("applyFont") = 1;
    if (fill_id) xf_node.append_attribute("applyFill") = 1;
    if (border_id) xf_node.append_attribute("applyBorder") = 1;
    if (alignment.first_attribute()) {
      xf_node.append_attribute("applyAlignment") = 1;
      xf_node.append_copy(alignment);
    }
    if (protection.first_attribute()) {
      xf_node.append_attribute("applyProtection") = 1;
      xf_node.append_copy(protection);
    }
    int32_t id = ods_index(xf_index, xf_list, ods_xml_string(xf_node));
    style_xf[name] = id;
    return id;
  }

  // a cell format with a built-in number format if it has none, for values of a type that needs one
  int32_t xf_typed(int32_t xf, int32_t builtin) {
    if (xf < 0 || static_cast<size_t>(xf) >= xf_list.size()) return xf;
    pugi::xml_document doc;
    doc.load_string(xf_list[static_cast<size_t>(xf)].c_str());
    pugi::xml_node node = doc.child("xf");
    if (node.attribute("numFmtId").as_int() != 0) return xf;
    node.attribute("numFmtId") = builtin;
    if (!node.attribute("applyNumberFormat")) node.append_attribute("applyNumberFormat") = 1;
    return ods_index(xf_index, xf_list, ods_xml_string(node));
  }
};

// an ODF cell range address (Sheet1.A1:Sheet1.A3) as absolute xlsx reference (Sheet1!$A$1:$A$3)
// an ODF cell range address list as chart range; several ranges are a union in parentheses
static std::string ods_xlsx_chart_range(const std::string& odf, const std::map<std::string, std::string>& renames) {
  std::vector<std::string> parts;
  size_t pos = 0;
  while (pos < odf.size()) {
    size_t end = pos;
    bool quoted = false;
    while (end < odf.size() && (quoted || odf[end] != ' ')) {
      if (odf[end] == '\'') quoted = !quoted;
      ++end;
    }
    std::string address = odf.substr(pos, end - pos);
    pos = end + 1;
    if (address.empty()) continue;
    std::string ref = ods_of_ref(address, renames);
    size_t bang = ref.rfind('!');
    if (bang == std::string::npos || ref == "#REF!") return "";
    std::string out = ref.substr(0, bang + 1);
    ods_ref_part p1, p2;
    size_t k = ods_parse_ref_part(ref, bang + 1, p1);
    if (k == bang + 1) return "";
    p1.col_abs = p1.row_abs = true;
    ods_write_ref_part(p1, 0, 0, out);
    if (k < ref.size() && ref[k] == ':' && ods_parse_ref_part(ref, k + 1, p2) > k + 1) {
      p2.col_abs = p2.row_abs = true;
      out += ":";
      ods_write_ref_part(p2, 0, 0, out);
    }
    parts.push_back(out);
  }
  if (parts.size() == 1) return parts[0];
  std::string out;
  for (const std::string& part : parts) out += (out.empty() ? "(" : ",") + part;
  return out.empty() ? out : out + ")";
}

// The preset dash of a dash: the preset of its name or the preset nearest to its dashes and gaps
// relative to the line width
static const char* ods_dash_preset(pugi::xml_node dash, double line_w) {
  std::string name = dash.attribute("draw:name").as_string();
  for (const ods_dash& d : ods_dashes) {
    if (name == d.name) return d.name;
  }
  auto relative = [&](const char* length) {
    std::string value = length;
    if (!value.empty() && value.back() == '%') return std::strtod(value.c_str(), nullptr);
    return ods_length_pt(length) / std::max(line_w, 0.75) * 100.0;
  };
  int32_t dots2 = dash.attribute("draw:dots2").as_int();
  double length1 = relative(dash.attribute("draw:dots1-length").as_string("100%")), distance = relative(dash.attribute("draw:distance").as_string("100%"));
  const char* preset = nullptr;
  double best = 0.0;
  for (const ods_dash& d : ods_dashes) {
    if ((d.dots2 > 0) != (dots2 > 0) || (dots2 > 0 && d.dots2 != std::min(dots2, 2))) continue;
    double diff = std::fabs(d.length1 - length1) + std::fabs(d.distance - distance);
    if (!preset || diff < best) {
      preset = d.name;
      best = diff;
    }
  }
  return preset;
}

// the size, weight, style, color and font of text properties as run properties
static void ods_xlsx_run_props(pugi::xml_node tp, pugi::xml_node rpr) {
  std::string size = tp.attribute("fo:font-size").as_string();
  if (!size.empty() && size.back() != '%') rpr.append_attribute("sz") = static_cast<int64_t>(std::llround(ods_length_pt(size.c_str()) * 100.0));
  if (tp.attribute("fo:font-weight")) rpr.append_attribute("b") = std::string(tp.attribute("fo:font-weight").as_string()) == "normal" ? 0 : 1;
  if (tp.attribute("fo:font-style")) rpr.append_attribute("i") = std::string(tp.attribute("fo:font-style").as_string()) == "italic" ? 1 : 0;
  std::string color = ods_rgb(tp.attribute("fo:color").as_string());
  if (!color.empty()) rpr.append_child("a:solidFill").append_child("a:srgbClr").append_attribute("val") = color.substr(2).c_str();
  std::string family = tp.attribute("fo:font-family").as_string();
  if (family.size() > 1 && (family[0] == '\'' || family[0] == '"')) family = family.substr(1, family.size() - 2);
  if (!family.empty()) rpr.append_child("a:latin").append_attribute("typeface") = family.c_str();
}

// a title with the text properties of its style
static void ods_xlsx_title(pugi::xml_node parent, const std::string& text, pugi::xml_node tp) {
  pugi::xml_node title = parent.append_child("c:title");
  pugi::xml_node rich = title.append_child("c:tx").append_child("c:rich");
  rich.append_child("a:bodyPr");
  pugi::xml_node run = rich.append_child("a:p").append_child("a:r");
  if (tp) ods_xlsx_run_props(tp, run.append_child("a:rPr"));
  if (run.child("a:rPr") && !run.child("a:rPr").first_attribute() && !run.child("a:rPr").first_child()) run.remove_child("a:rPr");
  run.append_child("a:t").text().set(text.c_str());
  title.append_child("c:overlay").append_attribute("val") = 0;
}

// The DrawingML chart for an ODF chart object, with the plot types, series,
// categories, secondary axes, titles, legend, colors, markers, data labels,
// data point colors, trendlines and axis scaling.
static std::string ods_chart_xml(const std::string& file, const std::map<std::string, std::string>& renames) {
  pugi::xml_document doc;
  if (!doc.load_file(file.c_str(), pugi::parse_default, pugi::encoding_utf8)) return "";
  ods_normalize_document(doc);
  pugi::xml_node root = doc.child("office:document-content");
  pugi::xml_node chart = root.child("office:body").child("office:chart").child("chart:chart");
  if (!chart) return "";

  std::map<std::string, pugi::xml_node> styles, data_styles, dashes;
  for (pugi::xml_node st : root.child("office:automatic-styles").children()) {
    if (std::strcmp(st.name(), "style:style") == 0) styles[st.attribute("style:name").as_string()] = st;
    else if (std::strncmp(st.name(), "number:", 7) == 0) data_styles[st.attribute("style:name").as_string()] = st;
  }
  // the dashes of the styles of the object
  pugi::xml_document styles_doc;
  if (styles_doc.load_file((file.substr(0, file.rfind('/') + 1) + "styles.xml").c_str(), pugi::parse_default, pugi::encoding_utf8)) {
    ods_normalize_document(styles_doc);
    for (pugi::xml_node dash : styles_doc.document_element().child("office:styles").children("draw:stroke-dash")) {
      dashes[dash.attribute("draw:name").as_string()] = dash;
    }
  }
  auto style_prop = [&](pugi::xml_node node, const char* props, const char* attr) {
    auto it = styles.find(node.attribute("chart:style-name").as_string());
    return it == styles.end() ? pugi::xml_attribute() : it->second.child(props).attribute(attr);
  };
  auto style_child = [&](pugi::xml_node node, const char* props) {
    auto it = styles.find(node.attribute("chart:style-name").as_string());
    return it == styles.end() ? pugi::xml_node() : it->second.child(props);
  };
  // the fill and line of an element as shape properties; elements without a stroke have a line if they are lines
  auto shape_props = [&](pugi::xml_node node, pugi::xml_node parent, bool with_fill, bool line_default) {
    pugi::xml_node gp = style_child(node, "style:graphic-properties");
    pugi::xml_node sppr = parent.append_child("c:spPr");
    std::string fill = gp.attribute("draw:fill").as_string();
    std::string fill_color = ods_rgb(gp.attribute("draw:fill-color").as_string());
    if (with_fill && fill == "none") {
      sppr.append_child("a:noFill");
    } else if (with_fill && (fill == "solid" || fill.empty()) && !fill_color.empty()) {
      pugi::xml_node clr = sppr.append_child("a:solidFill").append_child("a:srgbClr");
      clr.append_attribute("val") = fill_color.substr(2).c_str();
      double opacity = std::strtod(gp.attribute("draw:opacity").as_string("100"), nullptr);
      if (opacity < 100.0) clr.append_child("a:alpha").append_attribute("val") = static_cast<int64_t>(std::llround(std::max(opacity, 0.0) * 1000.0));
    }
    std::string stroke = gp.attribute("draw:stroke").as_string(line_default ? "solid" : "none");
    std::string stroke_color = ods_rgb(gp.attribute("svg:stroke-color").as_string());
    double width = ods_length_pt(gp.attribute("svg:stroke-width").as_string("0pt"));
    if (stroke == "none" && gp.attribute("draw:stroke")) {
      sppr.append_child("a:ln").append_child("a:noFill");
    } else if (stroke != "none" && (!stroke_color.empty() || width > 0.0 || stroke == "dash")) {
      pugi::xml_node ln = sppr.append_child("a:ln");
      if (width > 0.0) ln.append_attribute("w") = static_cast<int64_t>(std::llround(width * ods_emu_per_pt));
      if (!stroke_color.empty()) {
        pugi::xml_node clr = ln.append_child("a:solidFill").append_child("a:srgbClr");
        clr.append_attribute("val") = stroke_color.substr(2).c_str();
        double opacity = std::strtod(gp.attribute("svg:stroke-opacity").as_string("100"), nullptr);
        if (opacity < 100.0) clr.append_child("a:alpha").append_attribute("val") = static_cast<int64_t>(std::llround(std::max(opacity, 0.0) * 1000.0));
      }
      auto dash = dashes.find(gp.attribute("draw:stroke-dash").as_string());
      const char* preset = stroke == "dash" && dash != dashes.end() ? ods_dash_preset(dash->second, width) : nullptr;
      if (preset) ln.append_child("a:prstDash").append_attribute("val") = preset;
    }
    if (!sppr.first_child()) parent.remove_child(sppr);
  };
  // the text properties of an element as text properties of its paragraphs
  auto text_props = [&](pugi::xml_node node, pugi::xml_node parent) {
    pugi::xml_node tp = style_child(node, "style:text-properties");
    if (!tp) return;
    pugi::xml_node txpr = parent.append_child("c:txPr");
    txpr.append_child("a:bodyPr");
    txpr.append_child("a:lstStyle");
    pugi::xml_node p = txpr.append_child("a:p");
    ods_xlsx_run_props(tp, p.append_child("a:pPr").append_child("a:defRPr"));
    p.append_child("a:endParaRPr").append_attribute("lang") = "en-US";
  };
  // a number format that is not linked to the source
  auto number_format = [&](pugi::xml_node node, pugi::xml_node parent) {
    auto it = styles.find(node.attribute("chart:style-name").as_string());
    if (it == styles.end() || style_prop(node, "style:chart-properties", "chart:link-data-style-to-source").as_bool(true)) return;
    auto ds = data_styles.find(it->second.attribute("style:data-style-name").as_string());
    if (ds == data_styles.end()) return;
    std::string code = ods_format_code(ds->second, data_styles);
    if (code.empty()) return;
    pugi::xml_node num_fmt = parent.append_child("c:numFmt");
    num_fmt.append_attribute("formatCode") = code.c_str();
    num_fmt.append_attribute("sourceLinked") = 0;
  };
  auto color_of = [&](pugi::xml_node node, bool line) {
    std::string color = style_prop(node, "style:graphic-properties", "draw:fill-color").as_string();
    if (line) color = style_prop(node, "style:graphic-properties", "svg:stroke-color").as_string(color.c_str());
    return color.size() == 7 ? ods_upper(color.substr(1)) : std::string();
  };

  std::string main_class = chart.attribute("chart:class").as_string();
  pugi::xml_node plot = chart.child("chart:plot-area");
  bool vertical = style_prop(plot, "style:chart-properties", "chart:vertical").as_bool();
  bool stacked = style_prop(plot, "style:chart-properties", "chart:stacked").as_bool();
  bool percent = style_prop(plot, "style:chart-properties", "chart:percentage").as_bool();
  bool three_d = style_prop(plot, "style:chart-properties", "chart:three-dimensional").as_bool();
  bool smooth = std::string(style_prop(plot, "style:chart-properties", "chart:interpolation").as_string()) == "cubic-spline";

  pugi::xml_node x_axis, y_axis, y2_axis;
  for (pugi::xml_node ax : plot.children("chart:axis")) {
    std::string dim = ax.attribute("chart:dimension").as_string();
    std::string name = ax.attribute("chart:name").as_string();
    if (dim == "x" && !x_axis && name != "secondary-x") x_axis = ax;
    if (dim == "y" && name == "secondary-y") y2_axis = ax;
    else if (dim == "y" && !y_axis) y_axis = ax;
  }
  // the cells of the local table of the chart: numeric or not and the value; the descriptions of cells are the
  // ranges of the consumer the cells start, such as the ranges of a pivot chart
  std::vector<std::vector<std::pair<bool, std::string>>> local;
  std::map<std::string, std::pair<size_t, size_t>> local_starts;
  std::function<void(pugi::xml_node, std::string&)> local_text = [&](pugi::xml_node node, std::string& out) {
    for (pugi::xml_node child : node.children()) {
      if (child.type() == pugi::node_pcdata) out += child.value();
      else if (std::strcmp(child.name(), "draw:g") != 0) local_text(child, out);
    }
  };
  std::function<void(pugi::xml_node)> local_rows = [&](pugi::xml_node node) {
    for (pugi::xml_node child : node.children()) {
      if (std::strcmp(child.name(), "table:table-row") != 0) {
        local_rows(child);
        continue;
      }
      std::vector<std::pair<bool, std::string>> row;
      for (pugi::xml_node cell : child.children("table:table-cell")) {
        std::string type = cell.attribute("office:value-type").as_string();
        bool numeric = type == "float" || type == "percentage" || type == "currency";
        std::string value = cell.attribute("office:value").as_string();
        if (!numeric) local_text(cell, value);
        std::string desc = cell.child("draw:g").child_value("svg:desc");
        if (!desc.empty()) local_starts.insert(std::make_pair(desc, std::make_pair(local.size() + 1, row.size() + 1)));
        uint32_t repeat = std::min<uint32_t>(std::max<uint32_t>(cell.attribute("table:number-columns-repeated").as_uint(1), 1), 1024);
        for (uint32_t k = 0; k < repeat; ++k) row.push_back(std::make_pair(numeric, value));
      }
      local.push_back(row);
    }
  };
  for (pugi::xml_node table : chart.children("table:table")) {
    if (std::string(table.attribute("table:name").as_string()) == "local-table") local_rows(table);
  }
  // a range of the sheet, or the local cells of a range the consumer described: the cell of a label or the cells
  // down to the last row
  auto chart_range = [&](const char* odf, bool label) {
    std::string range = ods_xlsx_chart_range(odf, renames);
    auto start = local_starts.find(odf);
    if (!range.empty() || start == local_starts.end()) return range;
    std::string col = "$" + int_to_col(static_cast<uint32_t>(start->second.second));
    range = "local-table!" + col + "$" + std::to_string(start->second.first);
    if (!label && start->second.first < local.size()) range += ":" + col + "$" + std::to_string(local.size());
    return range;
  };
  std::string categories = chart_range(x_axis.child("chart:categories").attribute("table:cell-range-address").as_string(), false);

  // a reference to the sheet, or the values of a reference to the local table
  auto chart_ref = [&](pugi::xml_node parent, const char* name, const std::string& range, bool numeric) {
    pugi::xml_node node = parent.append_child(name);
    std::string prefix = range.compare(0, 14, "'local-table'!") == 0 ? "'local-table'!" : "local-table!";
    if (range.compare(0, prefix.size(), prefix) != 0) {
      node.append_child(numeric ? "c:numRef" : "c:strRef").append_child("c:f").text().set(range.c_str());
      return;
    }
    ods_ref_part p1, p2;
    size_t k = ods_parse_ref_part(range, prefix.size(), p1);
    p2 = p1;
    if (k < range.size() && range[k] == ':') ods_parse_ref_part(range, k + 1, p2);
    std::vector<std::string> values;
    bool all_numeric = true;
    for (int64_t r = p1.row; r >= 1 && r <= p2.row && r <= static_cast<int64_t>(local.size()); ++r) {
      const std::vector<std::pair<bool, std::string>>& row = local[static_cast<size_t>(r - 1)];
      for (int64_t c = p1.col; c >= 1 && c <= p2.col; ++c) {
        std::pair<bool, std::string> value = c <= static_cast<int64_t>(row.size()) ? row[static_cast<size_t>(c - 1)] : std::make_pair(true, std::string());
        if (!value.first && !value.second.empty()) all_numeric = false;
        values.push_back(value.second);
      }
    }
    if (std::strcmp(name, "c:tx") == 0) {
      node.append_child("c:v").text().set(values.empty() ? "" : values[0].c_str());
      return;
    }
    pugi::xml_node lit = node.append_child(numeric && all_numeric ? "c:numLit" : "c:strLit");
    lit.append_child("c:ptCount").append_attribute("val") = static_cast<uint32_t>(values.size());
    for (size_t i = 0; i < values.size(); ++i) {
      if (values[i].empty()) continue;
      pugi::xml_node pt = lit.append_child("c:pt");
      pt.append_attribute("idx") = static_cast<uint32_t>(i);
      pt.append_child("c:v").text().set(values[i].c_str());
    }
  };

  struct ods_series {
    std::string cls, values, label, domain, domain2;
    pugi::xml_node ser;
    bool secondary;
  };
  std::vector<ods_series> series;
  // the plot types with the axes they are on
  std::vector<std::pair<std::string, bool>> groups;
  for (pugi::xml_node ser : plot.children("chart:series")) {
    ods_series s;
    s.ser = ser;
    s.cls = ser.attribute("chart:class").as_string(main_class.c_str());
    if (main_class == "chart:stock") s.cls = main_class;
    s.secondary = std::string(ser.attribute("chart:attached-axis").as_string()) == "secondary-y" && main_class != "chart:stock";
    s.values = chart_range(ser.attribute("chart:values-cell-range-address").as_string(), false);
    s.label = chart_range(ser.attribute("chart:label-cell-address").as_string(), true);
    int32_t k = 0;
    for (pugi::xml_node domain : ser.children("chart:domain")) {
      std::string range = chart_range(domain.attribute("table:cell-range-address").as_string(), false);
      if (k++ == 0) s.domain = range;
      else s.domain2 = range;
    }
    if (s.values.empty()) continue;
    std::pair<std::string, bool> group = std::make_pair(s.cls, s.secondary);
    if (std::find(groups.begin(), groups.end(), group) == groups.end()) groups.push_back(group);
    series.push_back(s);
  }
  if (series.empty()) return "";

  // stock series are open, low, high and close in ODF, but open, high, low and close in xlsx
  if (main_class == "chart:stock" && series.size() >= 3) {
    size_t low = series.size() == 4 ? 1 : 0;
    std::swap(series[low], series[low + 1]);
  }

  pugi::xml_document out;
  pugi::xml_node cs = out.append_child("c:chartSpace");
  cs.append_attribute("xmlns:c") = "http://schemas.openxmlformats.org/drawingml/2006/chart";
  cs.append_attribute("xmlns:a") = "http://schemas.openxmlformats.org/drawingml/2006/main";
  cs.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  cs.append_child("c:roundedCorners").append_attribute("val") = 0;
  pugi::xml_node c = cs.append_child("c:chart");
  std::string title = ods_cell_string(chart.child("chart:title"));
  if (!title.empty()) ods_xlsx_title(c, title, style_child(chart.child("chart:title"), "style:text-properties"));
  c.append_child("c:autoTitleDeleted").append_attribute("val") = title.empty() ? 1 : 0;
  pugi::xml_node pa = c.append_child("c:plotArea");
  pa.append_child("c:layout");

  static const char* const trend_types[][2] = {
    {"linear", "linear"}, {"exponential", "exp"}, {"logarithmic", "log"}, {"polynomial", "poly"},
    {"power", "power"}, {"moving-average", "movingAvg"}
  };
  bool has_axes = false, has_secondary = false, only_xy = true;
  uint32_t idx = 0;
  for (const std::pair<std::string, bool>& group : groups) {
    const std::string& cls = group.first;
    const char* element = nullptr;
    if (cls == "chart:bar") element = three_d ? "c:bar3DChart" : "c:barChart";
    else if (cls == "chart:line") element = "c:lineChart";
    else if (cls == "chart:area") element = three_d ? "c:area3DChart" : "c:areaChart";
    else if (cls == "chart:circle") element = three_d ? "c:pie3DChart" : "c:pieChart";
    else if (cls == "chart:ring") element = "c:doughnutChart";
    else if (cls == "chart:scatter") element = "c:scatterChart";
    else if (cls == "chart:bubble") element = "c:bubbleChart";
    else if (cls == "chart:stock") element = "c:stockChart";
    else if (cls == "chart:radar" || cls == "chart:filled-radar") element = "c:radarChart";
    if (!element) continue;
    bool circle = cls == "chart:circle" || cls == "chart:ring";
    bool scatter = cls == "chart:scatter", bubble = cls == "chart:bubble", stock = cls == "chart:stock";
    bool bar = cls == "chart:bar", line = cls == "chart:line", radar = cls == "chart:radar" || cls == "chart:filled-radar";
    if (!scatter && !bubble) only_xy = false;
    const char* grouping = percent ? "percentStacked" : stacked ? "stacked" : "standard";

    pugi::xml_node g = pa.append_child(element);
    if (bar) {
      g.append_child("c:barDir").append_attribute("val") = vertical ? "bar" : "col";
      g.append_child("c:grouping").append_attribute("val") = percent || stacked ? grouping : "clustered";
    } else if (line || cls == "chart:area") {
      g.append_child("c:grouping").append_attribute("val") = grouping;
    } else if (scatter) {
      g.append_child("c:scatterStyle").append_attribute("val") = smooth ? "smoothMarker" : "lineMarker";
    } else if (radar) {
      g.append_child("c:radarStyle").append_attribute("val") = cls == "chart:filled-radar" ? "filled" : "marker";
    }
    if (!stock) g.append_child("c:varyColors").append_attribute("val") = circle ? 1 : 0;

    for (const ods_series& s : series) {
      if (s.cls != cls || s.secondary != group.second) continue;
      pugi::xml_node ser = g.append_child("c:ser");
      ser.append_child("c:idx").append_attribute("val") = idx;
      ser.append_child("c:order").append_attribute("val") = idx;
      ++idx;
      if (!s.label.empty()) chart_ref(ser, "c:tx", s.label, false);

      // lines of line series have the color of the fill without their own color
      bool line_series = line || scatter || cls == "chart:radar" || stock;
      if (stock) {
        ser.append_child("c:spPr").append_child("a:ln").append_child("a:noFill");
      } else if (!circle) {
        shape_props(s.ser, ser, !line_series, line_series);
        std::string color = color_of(s.ser, line_series);
        pugi::xml_node sppr = ser.child("c:spPr");
        if (line_series && !color.empty() && !sppr.child("a:ln").child("a:noFill") && !sppr.child("a:ln").child("a:solidFill")) {
          if (!sppr) sppr = ser.append_child("c:spPr");
          pugi::xml_node ln = sppr.child("a:ln") ? sppr.child("a:ln") : sppr.append_child("a:ln");
          ln.prepend_child("a:solidFill").append_child("a:srgbClr").append_attribute("val") = color.c_str();
        }
      }
      if (bar || bubble) ser.append_child("c:invertIfNegative").append_attribute("val") = 0;
      if (circle) {
        if (pugi::xml_attribute offset = style_prop(s.ser, "style:chart-properties", "chart:pie-offset")) ser.append_child("c:explosion").append_attribute("val") = offset.as_int();
      }
      if (line || scatter || radar || stock) {
        // the named symbols of spreadsheet software and their size
        static const char* const symbols[][2] = {
          {"circle", "circle"}, {"square", "square"}, {"diamond", "diamond"}, {"arrow-up", "triangle"}, {"x", "x"},
          {"asterisk", "star"}, {"horizontal-bar", "dash"}, {"plus", "plus"}
        };
        std::string symbol = style_prop(s.ser, "style:chart-properties", "chart:symbol-type").as_string();
        std::string symbol_name = style_prop(s.ser, "style:chart-properties", "chart:symbol-name").as_string();
        double symbol_size = ods_length_pt(style_prop(s.ser, "style:chart-properties", "chart:symbol-width").as_string("0pt"));
        const char* named = nullptr;
        for (const auto& sym : symbols) {
          if (symbol == "named-symbol" && symbol_name == sym[0]) named = sym[1];
        }
        if (symbol == "none" || (stock && symbol.empty())) {
          ser.append_child("c:marker").append_child("c:symbol").append_attribute("val") = "none";
        } else if (named || symbol_size > 0.0) {
          pugi::xml_node marker = ser.append_child("c:marker");
          if (named) marker.append_child("c:symbol").append_attribute("val") = named;
          if (symbol_size > 0.0) marker.append_child("c:size").append_attribute("val") = std::min<int64_t>(std::max<int64_t>(std::llround(symbol_size), 2), 72);
        }
      }

      // data points with their own color or offset
      int64_t point = 0;
      for (pugi::xml_node dp : s.ser.children("chart:data-point")) {
        int64_t rep = std::max<int64_t>(dp.attribute("chart:repeated").as_llong(1), 1);
        std::string point_color = color_of(dp, false);
        pugi::xml_attribute offset = style_prop(dp, "style:chart-properties", "chart:pie-offset");
        if (!point_color.empty() || (circle && offset)) {
          for (int64_t k = 0; k < rep && k < 1024; ++k) {
            pugi::xml_node pt = ser.append_child("c:dPt");
            pt.append_child("c:idx").append_attribute("val") = point + k;
            if (bar || bubble) pt.append_child("c:invertIfNegative").append_attribute("val") = 0;
            if (bubble) pt.append_child("c:bubble3D").append_attribute("val") = 0;
            if (circle && offset) pt.append_child("c:explosion").append_attribute("val") = offset.as_int();
            if (!point_color.empty()) shape_props(dp, pt, true, false);
          }
        }
        point += rep;
      }

      // data labels
      std::string number = style_prop(s.ser, "style:chart-properties", "chart:data-label-number").as_string();
      bool show_text = style_prop(s.ser, "style:chart-properties", "chart:data-label-text").as_bool();
      bool show_series = style_prop(s.ser, "style:chart-properties", "chart:data-label-series").as_bool();
      bool show_value = number == "value" || number == "value-and-percentage";
      bool show_percent = number == "percentage" || number == "value-and-percentage";
      if (show_value || show_percent || show_text || show_series) {
        pugi::xml_node labels = ser.append_child("c:dLbls");
        if (show_value) number_format(s.ser, labels);
        text_props(s.ser, labels);
        // the positions of labels of the plot type
        static const char* const positions[][2] = {
          {"center", "ctr"}, {"inside", "inEnd"}, {"near-origin", "inBase"}, {"outside", "outEnd"}, {"left", "l"},
          {"right", "r"}, {"top", "t"}, {"bottom", "b"}, {"avoid-overlap", "bestFit"}
        };
        std::string placement = style_prop(s.ser, "style:chart-properties", "chart:label-position").as_string();
        std::string allowed = circle ? " ctr inEnd outEnd bestFit " : bar ? (stacked || percent ? " ctr inEnd inBase " : " ctr inEnd inBase outEnd ") :
                              line || scatter ? " ctr l r t b " : bubble ? " ctr l r t b " : "";
        for (const auto& position : positions) {
          if (placement == position[0] && allowed.find(" " + std::string(position[1]) + " ") != std::string::npos) {
            labels.append_child("c:dLblPos").append_attribute("val") = position[1];
          }
        }
        labels.append_child("c:showLegendKey").append_attribute("val") = 0;
        labels.append_child("c:showVal").append_attribute("val") = show_value ? 1 : 0;
        labels.append_child("c:showCatName").append_attribute("val") = show_text ? 1 : 0;
        labels.append_child("c:showSerName").append_attribute("val") = show_series ? 1 : 0;
        labels.append_child("c:showPercent").append_attribute("val") = show_percent ? 1 : 0;
        labels.append_child("c:showBubbleSize").append_attribute("val") = 0;
      }

      // regression curves as trendlines
      if (!circle && !radar && !stock) {
        for (pugi::xml_node curve : s.ser.children("chart:regression-curve")) {
          std::string type = style_prop(curve, "style:chart-properties", "chart:regression-type").as_string();
          const char* trend = nullptr;
          for (const auto& tt : trend_types) {
            if (type == tt[0]) trend = tt[1];
          }
          if (!trend) continue;
          pugi::xml_node tl = ser.append_child("c:trendline");
          std::string name = style_prop(curve, "style:chart-properties", "chart:regression-name").as_string();
          if (!name.empty()) tl.append_child("c:name").text().set(name.c_str());
          tl.append_child("c:trendlineType").append_attribute("val") = trend;
          if (type == "polynomial") {
            int64_t degree = style_prop(curve, "style:chart-properties", "chart:regression-max-degree").as_llong(2);
            tl.append_child("c:order").append_attribute("val") = std::min<int64_t>(std::max<int64_t>(degree, 2), 6);
          }
          if (type == "moving-average") {
            int64_t period = style_prop(curve, "style:chart-properties", "chart:regression-period").as_llong(2);
            tl.append_child("c:period").append_attribute("val") = std::max<int64_t>(period, 2);
          }
          pugi::xml_node equation = curve.child("chart:equation");
          if (equation.attribute("chart:display-r-square").as_bool()) tl.append_child("c:dispRSqr").append_attribute("val") = 1;
          if (equation.attribute("chart:display-equation").as_bool()) tl.append_child("c:dispEq").append_attribute("val") = 1;
        }
      }

      if (scatter || bubble) {
        // bubble series: the y values are the first and the x values the second domain
        std::string x_values = bubble ? s.domain2 : s.domain, y_values = bubble ? s.domain : s.values;
        if (!x_values.empty()) chart_ref(ser, "c:xVal", x_values, true);
        chart_ref(ser, "c:yVal", y_values, true);
        if (bubble) {
          chart_ref(ser, "c:bubbleSize", s.values, true);
          ser.append_child("c:bubble3D").append_attribute("val") = 0;
        } else {
          ser.append_child("c:smooth").append_attribute("val") = smooth ? 1 : 0;
        }
      } else {
        if (!categories.empty()) chart_ref(ser, "c:cat", categories, false);
        chart_ref(ser, "c:val", s.values, true);
        if (line) ser.append_child("c:smooth").append_attribute("val") = smooth ? 1 : 0;
      }
    }

    if (bar) {
      pugi::xml_attribute gap = style_prop(y_axis, "style:chart-properties", "chart:gap-width");
      pugi::xml_attribute overlap = style_prop(y_axis, "style:chart-properties", "chart:overlap");
      if (gap || stacked || percent) g.append_child("c:gapWidth").append_attribute("val") = gap ? gap.as_int() : 150;
      if (!three_d && (overlap || stacked || percent)) g.append_child("c:overlap").append_attribute("val") = overlap ? overlap.as_int() : 100;
    }
    if (stock) {
      if (plot.child("chart:stock-range-line")) g.append_child("c:hiLowLines");
      if (plot.child("chart:stock-gain-marker") || series.size() == 4) {
        pugi::xml_node bars = g.append_child("c:upDownBars");
        bars.append_child("c:gapWidth").append_attribute("val") = 150;
        bars.append_child("c:upBars");
        bars.append_child("c:downBars");
      }
    }
    if (line) g.append_child("c:marker").append_attribute("val") = 1;
    if (bubble) g.append_child("c:bubbleScale").append_attribute("val") = 100;
    if (circle) {
      // the angle of the first slice clockwise from twelve o'clock
      int32_t offset = style_prop(plot, "style:chart-properties", "chart:angle-offset").as_int(90);
      if (!three_d) g.append_child("c:firstSliceAng").append_attribute("val") = ((90 - offset) % 360 + 360) % 360;
      if (cls == "chart:ring") g.append_child("c:holeSize").append_attribute("val") = 50;
    } else {
      int32_t first = group.second ? 3 : 1;
      g.append_child("c:axId").append_attribute("val") = first;
      g.append_child("c:axId").append_attribute("val") = first + 1;
      if (cls == "chart:area" && three_d) g.append_child("c:axId").append_attribute("val") = 5;
      if (bar && three_d) g.append_child("c:axId").append_attribute("val") = 5;
      has_axes = true;
      if (group.second) has_secondary = true;
    }
  }

  if (!pa.child("c:layout").next_sibling()) return "";

  // the axes with their scaling, grids and titles; the secondary category axis is hidden
  auto add_axis = [&](pugi::xml_node source, int32_t id, int32_t cross, bool category, bool horizontal, bool secondary) {
    pugi::xml_node ax = pa.append_child(category ? "c:catAx" : "c:valAx");
    ax.append_child("c:axId").append_attribute("val") = id;
    pugi::xml_node scaling = ax.append_child("c:scaling");
    if (style_prop(source, "style:chart-properties", "chart:logarithmic").as_bool()) scaling.append_child("c:logBase").append_attribute("val") = 10;
    scaling.append_child("c:orientation").append_attribute("val") = style_prop(source, "style:chart-properties", "chart:reverse-direction").as_bool() ? "maxMin" : "minMax";
    if (pugi::xml_attribute max = style_prop(source, "style:chart-properties", "chart:maximum")) scaling.append_child("c:max").append_attribute("val") = max.as_string();
    if (pugi::xml_attribute min = style_prop(source, "style:chart-properties", "chart:minimum")) scaling.append_child("c:min").append_attribute("val") = min.as_string();
    bool hidden = (secondary && category) || (source && !style_prop(source, "style:chart-properties", "chart:display-label").as_bool(true));
    ax.append_child("c:delete").append_attribute("val") = hidden ? 1 : 0;
    ax.append_child("c:axPos").append_attribute("val") = horizontal ? "b" : (secondary ? "r" : "l");
    for (const char* cls : {"major", "minor"}) {
      for (pugi::xml_node grid : source.children("chart:grid")) {
        if (std::string(grid.attribute("chart:class").as_string("major")) != cls) continue;
        pugi::xml_node lines = ax.append_child(cls[1] == 'a' ? "c:majorGridlines" : "c:minorGridlines");
        shape_props(grid, lines, false, true);
      }
    }
    std::string text = ods_cell_string(source.child("chart:title"));
    if (!text.empty()) ods_xlsx_title(ax, text, style_child(source.child("chart:title"), "style:text-properties"));
    number_format(source, ax);
    // the tick marks inside and outside of the axis, the position of the labels
    for (const char* tick : {"major", "minor"}) {
      pugi::xml_attribute inner = style_prop(source, "style:chart-properties", ("chart:tick-marks-" + std::string(tick) + "-inner").c_str());
      pugi::xml_attribute outer = style_prop(source, "style:chart-properties", ("chart:tick-marks-" + std::string(tick) + "-outer").c_str());
      if (!inner && !outer) continue;
      const char* mark = inner.as_bool() ? (outer.as_bool() ? "cross" : "in") : (outer.as_bool() ? "out" : "none");
      ax.append_child(tick[1] == 'a' ? "c:majorTickMark" : "c:minorTickMark").append_attribute("val") = mark;
    }
    std::string label_position = style_prop(source, "style:chart-properties", "chart:axis-label-position").as_string();
    if (label_position == "outside-start") ax.append_child("c:tickLblPos").append_attribute("val") = "low";
    else if (label_position == "outside-end") ax.append_child("c:tickLblPos").append_attribute("val") = "high";
    if (source) shape_props(source, ax, false, true);
    if (source) text_props(source, ax);
    ax.append_child("c:crossAx").append_attribute("val") = cross;
    // the position of the axis on the crossing axis
    std::string position = style_prop(source, "style:chart-properties", "chart:axis-position").as_string();
    if (secondary && !category) {
      ax.append_child("c:crosses").append_attribute("val") = "max";
    } else if (position == "end" || position == "start") {
      ax.append_child("c:crosses").append_attribute("val") = position == "end" ? "max" : "min";
    } else if (!position.empty() && position.find_first_not_of("0123456789.eE+-") == std::string::npos && std::strtod(position.c_str(), nullptr) != 0.0) {
      ax.append_child("c:crossesAt").append_attribute("val") = position.c_str();
    }
    if (!category) {
      double major = style_prop(source, "style:chart-properties", "chart:interval-major").as_double();
      int64_t divisor = style_prop(source, "style:chart-properties", "chart:interval-minor-divisor").as_llong();
      if (major > 0.0) ax.append_child("c:majorUnit").append_attribute("val") = style_prop(source, "style:chart-properties", "chart:interval-major").as_string();
      if (major > 0.0 && divisor > 1) ax.append_child("c:minorUnit").append_attribute("val") = ods_num(major / static_cast<double>(divisor)).c_str();
    }
  };
  if (has_axes) {
    bool category = !only_xy;
    add_axis(x_axis, 1, 2, category, !vertical, false);
    add_axis(y_axis, 2, 1, false, vertical, false);
    if (has_secondary) {
      add_axis(pugi::xml_node(), 3, 4, category, !vertical, true);
      add_axis(y2_axis, 4, 3, false, vertical, true);
    }
    if (three_d) {
      pugi::xml_node ser_ax = pa.append_child("c:serAx");
      ser_ax.append_child("c:axId").append_attribute("val") = 5;
      ser_ax.append_child("c:scaling").append_child("c:orientation").append_attribute("val") = "minMax";
      ser_ax.append_child("c:delete").append_attribute("val") = 1;
      ser_ax.append_child("c:axPos").append_attribute("val") = "b";
      ser_ax.append_child("c:crossAx").append_attribute("val") = 2;
    }
  }

  if (pugi::xml_node legend = chart.child("chart:legend")) {
    std::string pos = legend.attribute("chart:legend-position").as_string("end");
    const char* legend_pos = "r";
    if (pos == "start") legend_pos = "l";
    else if (pos == "top") legend_pos = "t";
    else if (pos == "bottom") legend_pos = "b";
    else if (pos == "top-end") legend_pos = "tr";
    pugi::xml_node lg = c.append_child("c:legend");
    lg.append_child("c:legendPos").append_attribute("val") = legend_pos;
    lg.append_child("c:overlay").append_attribute("val") = 0;
    shape_props(legend, lg, true, false);
    text_props(legend, lg);
  }
  c.append_child("c:plotVisOnly").append_attribute("val") = 1;
  std::string empty_cells = style_prop(plot, "style:chart-properties", "chart:treat-empty-cells").as_string();
  if (!empty_cells.empty()) c.append_child("c:dispBlanksAs").append_attribute("val") = empty_cells == "use-zero" ? "zero" : empty_cells == "ignore" ? "span" : "gap";
  // the fill and border of the wall are those of the plot area, those of the chart of the chart space
  shape_props(plot.child("chart:wall"), pa, true, false);
  shape_props(chart, cs, true, false);
  return ods_xml_string(cs);
}

// header and footer regions as codes (&P, &N, &D, &T, &A, &F), the text properties of spans as font, size, color
// and text decoration codes where they change
static std::string ods_header_code(pugi::xml_node region, const ods_xlsx_styles& st) {
  std::string code;
  bool first = true;
  std::string family, color, size;
  bool bold = false, italic = false, underline = false, strike = false;
  for (pugi::xml_node p : region.children("text:p")) {
    if (!first) code += "\n";
    first = false;
    for (pugi::xml_node n : ods_descendants(p, "")) {
      std::string name = n.name();
      bool text = n.type() == pugi::node_pcdata;
      if (!text && name != "text:s" && name != "text:page-number" && name != "text:page-count" && name != "text:date" &&
          name != "text:time" && name != "text:sheet-name" && name != "text:file-name" && name != "text:title") continue;
      if (text) {
        std::string parent = n.parent().name();
        if (parent == "text:page-number" || parent == "text:page-count" || parent == "text:date" ||
            parent == "text:time" || parent == "text:sheet-name" || parent == "text:file-name" || parent == "text:title") continue;
      }
      // the text properties of the innermost span
      pugi::xml_node tp;
      for (pugi::xml_node up = text ? n.parent() : n; up && up != p; up = up.parent()) {
        if (std::strcmp(up.name(), "text:span") != 0) continue;
        auto ts = st.text.find(up.attribute("text:style-name").as_string());
        if (ts != st.text.end()) tp = ts->second.child("style:text-properties");
        if (tp) break;
      }
      std::string new_family = tp.attribute("fo:font-family").as_string();
      auto face = st.font_face.find(tp.attribute("style:font-name").as_string());
      if (face != st.font_face.end()) new_family = face->second;
      else if (new_family.empty()) new_family = tp.attribute("style:font-name").as_string();
      new_family.erase(std::remove(new_family.begin(), new_family.end(), '\''), new_family.end());
      new_family.erase(std::remove(new_family.begin(), new_family.end(), '"'), new_family.end());
      std::string weight = tp.attribute("fo:font-weight").as_string(), font_style = tp.attribute("fo:font-style").as_string();
      bool new_bold = weight == "bold" || std::atoi(weight.c_str()) >= 600, new_italic = font_style == "italic" || font_style == "oblique";
      if (new_family != family || new_bold != bold || new_italic != italic) {
        code += "&\"" + (new_family.empty() ? std::string("-") : new_family) + "," +
                (new_bold && new_italic ? "Bold Italic" : new_bold ? "Bold" : new_italic ? "Italic" : "Regular") + "\"";
        family = new_family;
        bold = new_bold;
        italic = new_italic;
      }
      std::string font_size = tp.attribute("fo:font-size").as_string();
      if (!font_size.empty() && font_size.back() != '%' && font_size != size) {
        code += "&" + ods_num(std::round(ods_length_pt(font_size.c_str())));
        size = font_size;
      }
      std::string rgb = ods_rgb(tp.attribute("fo:color").as_string());
      if (!rgb.empty() && rgb != color) {
        code += "&K" + rgb.substr(2);
        color = rgb;
      }
      std::string underline_style = tp.attribute("style:text-underline-style").as_string();
      bool new_underline = !underline_style.empty() && underline_style != "none";
      if (new_underline != underline) code += "&U";
      underline = new_underline;
      std::string strike_style = tp.attribute("style:text-line-through-style").as_string();
      bool new_strike = !strike_style.empty() && strike_style != "none";
      if (new_strike != strike) code += "&S";
      strike = new_strike;

      if (text) {
        for (const char* ch = n.value(); *ch; ++ch) {
          if (*ch == '&') code += "&&";
          else code += *ch;
        }
      } else if (name == "text:s") {
        code.append(static_cast<size_t>(n.attribute("text:c").as_uint(1)), ' ');
      } else if (name == "text:page-number") {
        code += "&P";
      } else if (name == "text:page-count") {
        code += "&N";
      } else if (name == "text:date") {
        code += "&D";
      } else if (name == "text:time") {
        code += "&T";
      } else if (name == "text:sheet-name") {
        code += "&A";
      } else {
        code += std::string(n.attribute("text:display").as_string()) == "path" ? "&Z" : "&F";
      }
    }
  }
  return code;
}

struct ods_in_frame {
  int64_t row = 0, col = 0;
  double x = 0, y = 0, width = 0, height = 0, end_x = 0, end_y = 0, rotation = 0;
  std::string end_cell, image, object, shape, text, fill, line, align, control, geometry;
  bool rich = false;
  // the clipped part of a picture (left, top, right, bottom) in 1/100000 and the gradient of a shape
  double crop[4] = {0, 0, 0, 0};
  std::string gradient_start, gradient_end;
  double gradient_angle = 0;
  // the flips, fill, line, effects and text box of a shape as DrawingML
  std::string properties;
};

static ods_in_frame ods_read_frame(pugi::xml_node frame) {
  ods_in_frame fr;
  fr.x = ods_length_pt(frame.attribute("svg:x").as_string("0pt"));
  fr.y = ods_length_pt(frame.attribute("svg:y").as_string("0pt"));
  fr.width = ods_length_pt(frame.attribute("svg:width").as_string("0pt"));
  fr.height = ods_length_pt(frame.attribute("svg:height").as_string("0pt"));
  std::string end = frame.attribute("table:end-cell-address").as_string();
  size_t dot = end.rfind('.');
  if (dot != std::string::npos) {
    for (char ch : end.substr(dot + 1)) {
      if (ch != '$') fr.end_cell += ch;
    }
  }
  fr.end_x = ods_length_pt(frame.attribute("table:end-x").as_string("0pt"));
  fr.end_y = ods_length_pt(frame.attribute("table:end-y").as_string("0pt"));

  // a rotation around the center, the position is where the unrotated frame starts
  std::string transform = frame.attribute("draw:transform").as_string();
  size_t rotate = transform.find("rotate");
  if (rotate != std::string::npos) {
    size_t open = transform.find('(', rotate);
    double a = open == std::string::npos ? 0.0 : std::strtod(transform.c_str() + open + 1, nullptr);
    double tx = 0.0, ty = 0.0;
    size_t translate = transform.find("translate");
    if (translate != std::string::npos) {
      size_t start = transform.find('(', translate);
      size_t close = transform.find(')', start);
      std::string args = start == std::string::npos || close == std::string::npos ? "" : transform.substr(start + 1, close - start - 1);
      size_t first = args.find_first_not_of(' ');
      size_t space = args.find_first_of(" ,", first);
      if (first != std::string::npos && space != std::string::npos) {
        tx = ods_length_pt(args.substr(first, space - first).c_str());
        ty = ods_length_pt(args.substr(args.find_first_not_of(" ,", space)).c_str());
      }
    }
    fr.rotation = -a * 180.0 / 3.14159265358979323846;
    fr.x = std::max(0.0, tx - (fr.width / 2.0 - (fr.width / 2.0 * std::cos(a) + fr.height / 2.0 * std::sin(a))));
    fr.y = std::max(0.0, ty - (fr.height / 2.0 - (-fr.width / 2.0 * std::sin(a) + fr.height / 2.0 * std::cos(a))));
    fr.end_cell.clear();
  }
  // an object comes with a replacement image, the object is used
  fr.object = frame.child("draw:object").attribute("xlink:href").as_string();
  if (fr.object.compare(0, 2, "./") == 0) fr.object = fr.object.substr(2);
  if (fr.object.empty()) fr.image = frame.child("draw:image").attribute("xlink:href").as_string();
  return fr;
}

// the paragraphs of a cell as xlsx rich text runs, empty if no span has a text style
static std::string ods_cell_rich(pugi::xml_node cell, ods_xlsx_styles& st, const std::string& cell_style) {
  std::vector<std::pair<std::string, std::string>> runs;
  bool first = true;
  for (pugi::xml_node p : cell.children("text:p")) {
    if (!first) {
      if (runs.empty()) runs.push_back(std::make_pair(std::string(), std::string()));
      runs.back().second += '\n';
    }
    ods_text_runs(p, "", runs);
    first = false;
  }
  bool styled = false;
  for (const auto& run : runs) {
    if (st.text.count(run.first)) styled = true;
  }
  if (!styled) return "";
  pugi::xml_document doc;
  for (const auto& run : runs) {
    pugi::xml_node r = doc.append_child("r");
    if (st.text.count(run.first)) {
      pugi::xml_document rpr;
      rpr.load_string(st.run_properties(run.first, cell_style).c_str());
      r.append_copy(rpr.first_child());
    }
    pugi::xml_node t = r.append_child("t");
    t.append_attribute("xml:space") = "preserve";
    t.text().set(run.second.c_str());
  }
  std::string out;
  for (pugi::xml_node r : doc.children()) out += ods_xml_string(r);
  return out;
}

struct ods_in_cell {
  int32_t row;
  int32_t col;
  int32_t xf;
  std::string t, v, f, text, rich;
  int64_t matrix_rows = 0, matrix_cols = 0;
};

struct ods_in_run {
  int64_t row, rows, col, cols;
  int32_t xf;
};

struct ods_in_col {
  int64_t min, max;
  double width;
  bool hidden;
  std::string style;
  int64_t level;
  // the column follows a collapsed group
  bool collapsed = false;
};

struct ods_in_row {
  int64_t row, rows;
  double height;
  bool hidden;
  int32_t xf;
  int64_t level;
  // the height of rows with optimal height, used to place frames
  double optimal = -1.0;
  // the row follows a collapsed group
  bool collapsed = false;
};

// the outline level of a row or column: the number of groups around it
static int64_t ods_outline_level(pugi::xml_node node, const char* group) {
  int64_t level = 0;
  for (pugi::xml_node p = node.parent(); p && std::strcmp(p.name(), "table:table") != 0; p = p.parent()) {
    if (std::strcmp(p.name(), group) == 0) ++level;
  }
  return std::min<int64_t>(level, 7);
}

// a row or column after a collapsed group, or after a group that ends with one
static bool ods_after_collapsed(pugi::xml_node node, const char* group) {
  pugi::xml_node prev = node.previous_sibling();
  while (prev && std::strcmp(prev.name(), group) == 0) {
    if (std::strcmp(prev.attribute("table:display").as_string(), "false") == 0) return true;
    prev = prev.last_child();
  }
  return false;
}

// an ODF cell range address ('S 1'.A1:'S 1'.B2) as sheet name and xlsx range
// a string as R string literal
static std::string ods_r_string(const std::string& x) {
  std::string out = "\"";
  for (char c : x) {
    if (c == '\\' || c == '"') out += '\\';
    out += c;
  }
  return out + "\"";
}

static bool ods_split_range(const std::string& address, std::string& sheet, std::string& ref) {
  std::string xlsx = ods_of_ref(address, std::map<std::string, std::string>());
  size_t bang = xlsx.rfind('!');
  if (bang == std::string::npos || xlsx == "#REF!") return false;
  ref = xlsx.substr(bang + 1);
  std::string quoted = xlsx.substr(0, bang);
  sheet.clear();
  if (!quoted.empty() && quoted[0] == '\'') {
    for (size_t i = 1; i + 1 < quoted.size(); ++i) {
      sheet += quoted[i];
      if (quoted[i] == '\'' && quoted[i + 1] == '\'') ++i;
    }
  } else {
    sheet = quoted;
  }
  return true;
}

// the text argument of a condition such as contains-text("x")
static std::string ods_text_argument(const std::string& arg) {
  std::string out;
  if (arg.size() < 2 || arg.front() != '"') return arg;
  for (size_t i = 1; i + 1 < arg.size(); ++i) {
    out += arg[i];
    if (arg[i] == '"' && arg[i + 1] == '"') ++i;
  }
  return out;
}

// a conditional format rule of the calcext: namespace as xlsx cfRule
static bool ods_cf_rule(pugi::xml_node entry, pugi::xml_node rule, const std::string& top_left, int64_t dxf, const std::map<std::string, std::string>& renames) {
  std::string name = entry.name();
  static const char* const cfvo_types[][2] = {
    {"minimum", "min"}, {"maximum", "max"}, {"auto-minimum", "min"}, {"auto-maximum", "max"},
    {"number", "num"}, {"percent", "percent"}, {"percentile", "percentile"}, {"formula", "formula"}
  };
  auto cfvo = [&](pugi::xml_node parent, pugi::xml_node e) {
    std::string type = e.attribute("calcext:type").as_string("number");
    const char* xtype = "num";
    for (const auto& t : cfvo_types) {
      if (type == t[0]) xtype = t[1];
    }
    pugi::xml_node v = parent.append_child("cfvo");
    v.append_attribute("type") = xtype;
    if (std::strcmp(xtype, "min") != 0 && std::strcmp(xtype, "max") != 0) {
      std::string value = e.attribute("calcext:value").as_string("0");
      v.append_attribute("val") = (std::strcmp(xtype, "formula") == 0 ? ods_formula_to_xlsx(value, renames) : value).c_str();
    }
  };

  if (name == "calcext:color-scale") {
    rule.append_attribute("type") = "colorScale";
    pugi::xml_node scale = rule.append_child("colorScale");
    for (pugi::xml_node e : entry.children("calcext:color-scale-entry")) cfvo(scale, e);
    for (pugi::xml_node e : entry.children("calcext:color-scale-entry")) {
      scale.append_child("color").append_attribute("rgb") = ods_rgb(e.attribute("calcext:color").as_string("#000000")).c_str();
    }
    return true;
  }
  if (name == "calcext:data-bar") {
    rule.append_attribute("type") = "dataBar";
    pugi::xml_node bar = rule.append_child("dataBar");
    for (pugi::xml_node e : entry.children("calcext:formatting-entry")) cfvo(bar, e);
    bar.append_child("color").append_attribute("rgb") = ods_rgb(entry.attribute("calcext:positive-color").as_string("#638ec6")).c_str();
    return true;
  }
  if (name == "calcext:icon-set") {
    rule.append_attribute("type") = "iconSet";
    pugi::xml_node set = rule.append_child("iconSet");
    set.append_attribute("iconSet") = entry.attribute("calcext:icon-set-type").as_string("3TrafficLights1");
    if (!entry.attribute("calcext:show-value").as_bool(true)) set.append_attribute("showValue") = 0;
    for (pugi::xml_node e : entry.children("calcext:formatting-entry")) cfvo(set, e);
    return true;
  }
  if (name == "calcext:date-is") {
    // the periods of spreadsheet software with their formulas, years as expressions
    static const char* const periods[][3] = {
      {"today", "today", "FLOOR(#,1)=TODAY()"}, {"yesterday", "yesterday", "FLOOR(#,1)=TODAY()-1"},
      {"tomorrow", "tomorrow", "FLOOR(#,1)=TODAY()+1"}, {"last-7-days", "last7Days", "AND(TODAY()-FLOOR(#,1)<=6,FLOOR(#,1)<=TODAY())"},
      {"this-week", "thisWeek", "AND(TODAY()-ROUNDDOWN(#,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(#,0)-TODAY()<=7-WEEKDAY(TODAY()))"},
      {"last-week", "lastWeek", "AND(TODAY()-ROUNDDOWN(#,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(#,0)<(WEEKDAY(TODAY())+7))"},
      {"next-week", "nextWeek", "AND(ROUNDDOWN(#,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(#,0)-TODAY()<(15-WEEKDAY(TODAY())))"},
      {"this-month", "thisMonth", "AND(MONTH(#)=MONTH(TODAY()),YEAR(#)=YEAR(TODAY()))"},
      {"last-month", "lastMonth", "AND(MONTH(#)=MONTH(EDATE(TODAY(),0-1)),YEAR(#)=YEAR(EDATE(TODAY(),0-1)))"},
      {"next-month", "nextMonth", "AND(MONTH(#)=MONTH(EDATE(TODAY(),0+1)),YEAR(#)=YEAR(EDATE(TODAY(),0+1)))"},
      {"this-year", "", "YEAR(#)=YEAR(TODAY())"}, {"last-year", "", "YEAR(#)=YEAR(TODAY())-1"}, {"next-year", "", "YEAR(#)=YEAR(TODAY())+1"}
    };
    std::string date = entry.attribute("calcext:date").as_string();
    for (const auto& period : periods) {
      if (date != period[0]) continue;
      std::string formula;
      for (const char* c = period[2]; *c; ++c) {
        if (*c == '#') formula += top_left;
        else formula += *c;
      }
      rule.append_attribute("type") = *period[1] ? "timePeriod" : "expression";
      rule.append_attribute("dxfId") = dxf;
      if (*period[1]) rule.append_attribute("timePeriod") = period[1];
      rule.append_child("formula").text().set(formula.c_str());
      return true;
    }
    return false;
  }
  if (name != "calcext:condition") return false;

  std::string value = entry.attribute("calcext:value").as_string();
  size_t open = value.find('(');
  std::string fn = open == std::string::npos ? value : value.substr(0, open);
  size_t end = 0;
  std::vector<std::string> args;
  if (open != std::string::npos) args = ods_arguments(value, open + 1, end);

  static const char* const ops[][2] = {
    {"!=", "notEqual"}, {">=", "greaterThanOrEqual"}, {"<=", "lessThanOrEqual"}, {"=", "equal"}, {">", "greaterThan"}, {"<", "lessThan"}
  };
  for (const auto& op : ops) {
    size_t len = std::strlen(op[0]);
    if (value.compare(0, len, op[0]) != 0) continue;
    rule.append_attribute("type") = "cellIs";
    rule.append_attribute("dxfId") = dxf;
    rule.append_attribute("operator") = op[1];
    rule.append_child("formula").text().set(ods_formula_to_xlsx(value.substr(len), renames).c_str());
    return true;
  }

  std::string text = args.empty() ? "" : ods_text_argument(args[0]);
  std::string quoted;
  for (char c : text) quoted += c == '"' ? std::string("\"\"") : std::string(1, c);
  quoted = "\"" + quoted + "\"";

  if ((fn == "between" || fn == "not-between") && args.size() == 2) {
    rule.append_attribute("type") = "cellIs";
    rule.append_attribute("dxfId") = dxf;
    rule.append_attribute("operator") = fn == "between" ? "between" : "notBetween";
    rule.append_child("formula").text().set(ods_formula_to_xlsx(args[0], renames).c_str());
    rule.append_child("formula").text().set(ods_formula_to_xlsx(args[1], renames).c_str());
  } else if (fn == "formula-is" && !args.empty()) {
    rule.append_attribute("type") = "expression";
    rule.append_attribute("dxfId") = dxf;
    rule.append_child("formula").text().set(ods_formula_to_xlsx(args[0], renames).c_str());
  } else if (fn == "contains-text" || fn == "not-contains-text" || fn == "begins-with" || fn == "ends-with") {
    const char* type = fn == "contains-text" ? "containsText" : fn == "not-contains-text" ? "notContainsText" : fn == "begins-with" ? "beginsWith" : "endsWith";
    std::string formula;
    if (fn == "contains-text") formula = "NOT(ISERROR(SEARCH(" + quoted + "," + top_left + ")))";
    else if (fn == "not-contains-text") formula = "ISERROR(SEARCH(" + quoted + "," + top_left + "))";
    else if (fn == "begins-with") formula = "LEFT(" + top_left + ",LEN(" + quoted + "))=" + quoted;
    else formula = "RIGHT(" + top_left + ",LEN(" + quoted + "))=" + quoted;
    rule.append_attribute("type") = type;
    rule.append_attribute("dxfId") = dxf;
    rule.append_attribute("operator") = type;
    rule.append_attribute("text") = text.c_str();
    rule.append_child("formula").text().set(formula.c_str());
  } else if (fn == "duplicate" || fn == "unique") {
    rule.append_attribute("type") = fn == "duplicate" ? "duplicateValues" : "uniqueValues";
    rule.append_attribute("dxfId") = dxf;
  } else if (fn == "top-elements" || fn == "bottom-elements" || fn == "top-percent" || fn == "bottom-percent") {
    rule.append_attribute("type") = "top10";
    rule.append_attribute("dxfId") = dxf;
    if (fn.compare(0, 6, "bottom") == 0) rule.append_attribute("bottom") = 1;
    if (fn.find("percent") != std::string::npos) rule.append_attribute("percent") = 1;
    rule.append_attribute("rank") = args.empty() ? "10" : args[0].c_str();
  } else if (fn == "above-average" || fn == "below-average" || fn == "above-equal-average" || fn == "below-equal-average") {
    rule.append_attribute("type") = "aboveAverage";
    rule.append_attribute("dxfId") = dxf;
    if (fn.compare(0, 5, "below") == 0) rule.append_attribute("aboveAverage") = 0;
    if (fn.find("equal") != std::string::npos) rule.append_attribute("equalAverage") = 1;
  } else if (fn == "is-error" || fn == "is-no-error") {
    rule.append_attribute("type") = fn == "is-error" ? "containsErrors" : "notContainsErrors";
    rule.append_attribute("dxfId") = dxf;
    rule.append_child("formula").text().set((fn == "is-error" ? "ISERROR(" + top_left + ")" : "NOT(ISERROR(" + top_left + "))").c_str());
  } else {
    return false;
  }
  return true;
}

// the differential format of a named cell style used by a conditional format
static std::string ods_dxf(pugi::xml_node style, const std::map<std::string, std::string>& font_face) {
  pugi::xml_document doc;
  pugi::xml_node dxf = doc.append_child("dxf");
  pugi::xml_node tp = style.child("style:text-properties");
  pugi::xml_node tcp = style.child("style:table-cell-properties");

  pugi::xml_node font = dxf.append_child("font");
  std::string weight = tp.attribute("fo:font-weight").as_string();
  if (weight == "bold" || std::atoi(weight.c_str()) >= 600) font.append_child("b");
  std::string font_style = tp.attribute("fo:font-style").as_string();
  if (font_style == "italic" || font_style == "oblique") font.append_child("i");
  std::string strike = tp.attribute("style:text-line-through-style").as_string();
  if (!strike.empty() && strike != "none") font.append_child("strike");
  std::string underline = tp.attribute("style:text-underline-style").as_string();
  if (!underline.empty() && underline != "none") font.append_child("u");
  std::string color = ods_rgb(tp.attribute("fo:color").as_string());
  if (!color.empty()) font.append_child("color").append_attribute("rgb") = color.c_str();
  auto face = font_face.find(tp.attribute("style:font-name").as_string());
  std::string family = face != font_face.end() ? face->second : tp.attribute("fo:font-family").as_string();
  family.erase(std::remove(family.begin(), family.end(), '\''), family.end());
  if (!family.empty()) font.append_child("name").append_attribute("val") = family.c_str();
  if (!font.first_child()) dxf.remove_child(font);

  std::string background = ods_rgb(tcp.attribute("fo:background-color").as_string());
  if (!background.empty()) dxf.append_child("fill").append_child("patternFill").append_child("bgColor").append_attribute("rgb") = background.c_str();

  pugi::xml_node border = dxf.append_child("border");
  static const char* const sides[4][2] = {
    {"left", "fo:border-left"}, {"right", "fo:border-right"}, {"top", "fo:border-top"}, {"bottom", "fo:border-bottom"}
  };
  for (const auto& side : sides) {
    pugi::xml_attribute specific = tcp.attribute(side[1]);
    std::string side_color = "FF000000";
    std::string kind = ods_xlsx_styles::border_side(specific ? specific.as_string() : tcp.attribute("fo:border").as_string(), side_color);
    if (kind.empty()) continue;
    pugi::xml_node node = border.append_child(side[0]);
    node.append_attribute("style") = kind.c_str();
    node.append_child("color").append_attribute("rgb") = side_color.c_str();
  }
  if (!border.first_child()) dxf.remove_child(border);
  return ods_xml_string(dxf);
}

// a shape as preset geometry of DrawingML; other shapes are rectangles
// The outline of an enhanced geometry as custom geometry of DrawingML. The equations are evaluated for
// the view box and the size of the shape (in points), arcs and quadrants become cubic Bezier curves.
static std::string ods_enhanced_custgeom(pugi::xml_node geometry, double width_pt, double height_pt) {
  const double pi = 3.14159265358979323846;
  std::vector<double> box;
  std::string view = geometry.attribute("svg:viewBox").as_string("0 0 21600 21600");
  for (size_t pos = 0; pos < view.size();) {
    char* end = nullptr;
    double v = std::strtod(view.c_str() + pos, &end);
    if (end == view.c_str() + pos) break;
    box.push_back(v);
    pos = static_cast<size_t>(end - view.c_str());
  }
  if (box.size() != 4 || !(box[2] > 0.0) || !(box[3] > 0.0)) return "";
  std::vector<double> modifiers;
  std::string mods = geometry.attribute("draw:modifiers").as_string();
  for (size_t pos = 0; pos < mods.size();) {
    char* end = nullptr;
    double v = std::strtod(mods.c_str() + pos, &end);
    if (end == mods.c_str() + pos) break;
    modifiers.push_back(v);
    pos = static_cast<size_t>(end - mods.c_str());
  }

  // the equations reference each other by name; they are evaluated on demand
  std::map<std::string, std::string> formulas;
  std::map<std::string, double> values;
  for (pugi::xml_node eq : geometry.children("draw:equation")) formulas[eq.attribute("draw:name").as_string()] = eq.attribute("draw:formula").as_string();
  std::set<std::string> evaluating;
  std::function<double(const std::string&, size_t&)> expression;
  std::function<double(const std::string&)> equation = [&](const std::string& name) {
    std::map<std::string, double>::const_iterator known = values.find(name);
    if (known != values.end()) return known->second;
    std::map<std::string, std::string>::const_iterator f = formulas.find(name);
    if (f == formulas.end() || evaluating.count(name)) return 0.0;
    evaluating.insert(name);
    size_t pos = 0;
    double v = expression(f->second, pos);
    evaluating.erase(name);
    values[name] = v;
    return v;
  };
  auto skip = [](const std::string& x, size_t& pos) {
    while (pos < x.size() && (x[pos] == ' ' || x[pos] == '\t')) ++pos;
  };
  std::function<double(const std::string&, size_t&)> primary = [&](const std::string& x, size_t& pos) -> double {
    skip(x, pos);
    if (pos >= x.size()) return 0.0;
    char c = x[pos];
    if (c == '(') {
      ++pos;
      double v = expression(x, pos);
      skip(x, pos);
      if (pos < x.size() && x[pos] == ')') ++pos;
      return v;
    }
    if (c == '-') {
      ++pos;
      return -primary(x, pos);
    }
    if (c == '+') {
      ++pos;
      return primary(x, pos);
    }
    if (c == '?') {
      size_t start = ++pos;
      while (pos < x.size() && (std::isalnum(static_cast<unsigned char>(x[pos])) || x[pos] == '_')) ++pos;
      return equation(x.substr(start, pos - start));
    }
    if (c == '$') {
      size_t start = ++pos;
      while (pos < x.size() && std::isdigit(static_cast<unsigned char>(x[pos]))) ++pos;
      size_t k = static_cast<size_t>(std::strtoul(x.substr(start, pos - start).c_str(), nullptr, 10));
      return k < modifiers.size() ? modifiers[k] : 0.0;
    }
    if (std::isdigit(static_cast<unsigned char>(c)) || c == '.') {
      char* end = nullptr;
      double v = std::strtod(x.c_str() + pos, &end);
      pos = static_cast<size_t>(end - x.c_str());
      return v;
    }
    size_t start = pos;
    while (pos < x.size() && std::isalnum(static_cast<unsigned char>(x[pos]))) ++pos;
    std::string name = x.substr(start, pos - start);
    skip(x, pos);
    std::vector<double> args;
    if (pos < x.size() && x[pos] == '(') {
      ++pos;
      while (pos < x.size()) {
        args.push_back(expression(x, pos));
        skip(x, pos);
        if (pos < x.size() && x[pos] == ',') {
          ++pos;
          continue;
        }
        if (pos < x.size() && x[pos] == ')') ++pos;
        break;
      }
    }
    args.resize(std::max<size_t>(args.size(), 3), 0.0);
    if (name == "pi") return pi;
    if (name == "left") return box[0];
    if (name == "top") return box[1];
    if (name == "right") return box[0] + box[2];
    if (name == "bottom") return box[1] + box[3];
    if (name == "width") return box[2];
    if (name == "height") return box[3];
    if (name == "logwidth") return width_pt / 72.0 * 2540.0;
    if (name == "logheight") return height_pt / 72.0 * 2540.0;
    if (name == "xstretch" || name == "ystretch" || name == "hasstroke" || name == "hasfill") return 1.0;
    if (name == "abs") return std::fabs(args[0]);
    if (name == "sqrt") return std::sqrt(std::max(args[0], 0.0));
    if (name == "sin") return std::sin(args[0]);
    if (name == "cos") return std::cos(args[0]);
    if (name == "tan") return std::tan(args[0]);
    if (name == "atan") return std::atan(args[0]);
    if (name == "atan2") return std::atan2(args[0], args[1]);
    if (name == "min") return std::min(args[0], args[1]);
    if (name == "max") return std::max(args[0], args[1]);
    if (name == "if") return args[0] > 0.0 ? args[1] : args[2];
    return 0.0;
  };
  std::function<double(const std::string&, size_t&)> term = [&](const std::string& x, size_t& pos) {
    double v = primary(x, pos);
    for (;;) {
      skip(x, pos);
      if (pos >= x.size() || (x[pos] != '*' && x[pos] != '/')) return v;
      char op = x[pos++];
      double r = primary(x, pos);
      v = op == '*' ? v * r : (r != 0.0 ? v / r : 0.0);
    }
  };
  expression = [&](const std::string& x, size_t& pos) {
    double v = term(x, pos);
    for (;;) {
      skip(x, pos);
      if (pos >= x.size() || (x[pos] != '+' && x[pos] != '-')) return v;
      char op = x[pos++];
      double r = term(x, pos);
      v = op == '+' ? v + r : v - r;
    }
  };

  // the path: commands and their parameters; a parameter is a number, an equation or a modifier
  std::string path = geometry.attribute("draw:enhanced-path").as_string();
  if (path.empty()) return "";
  std::vector<std::pair<char, std::vector<double>>> commands;
  for (size_t pos = 0; pos < path.size();) {
    char c = path[pos];
    if (c == ' ' || c == ',') {
      ++pos;
    } else if (std::isalpha(static_cast<unsigned char>(c))) {
      commands.push_back(std::make_pair(c, std::vector<double>()));
      ++pos;
    } else {
      // a reference to an equation or modifier, or a number
      size_t start = pos;
      if (c == '?' || c == '$') {
        ++pos;
        while (pos < path.size() && std::isalnum(static_cast<unsigned char>(path[pos]))) ++pos;
      } else {
        while (pos < path.size() && path[pos] != ' ' && path[pos] != ',' && !(std::isalpha(static_cast<unsigned char>(path[pos])) && path[pos] != 'e')) ++pos;
      }
      if (pos == start) {
        ++pos;
        continue;
      }
      std::string token = path.substr(start, pos - start);
      size_t at = 0;
      if (!commands.empty()) commands.back().second.push_back(token[0] == '?' || token[0] == '$' ? primary(token, at) : std::strtod(token.c_str(), nullptr));
    }
  }

  // DrawingML coordinates are integers: the view box is scaled to at least 21600 units
  double scale = std::max(1.0, 21600.0 / std::max(box[2], box[3]));
  auto coord = [&](double v, size_t axis) { return static_cast<int64_t>(std::llround((v - box[axis]) * scale)); };
  pugi::xml_document doc;
  pugi::xml_node geom = doc.append_child("a:custGeom");
  geom.append_child("a:avLst");
  geom.append_child("a:gdLst");
  geom.append_child("a:ahLst");
  geom.append_child("a:cxnLst");
  pugi::xml_node rect = geom.append_child("a:rect");
  rect.append_attribute("l") = "l";
  rect.append_attribute("t") = "t";
  rect.append_attribute("r") = "r";
  rect.append_attribute("b") = "b";
  pugi::xml_node list = geom.append_child("a:pathLst");
  pugi::xml_node out;
  double px = 0.0, py = 0.0;
  auto new_path = [&]() {
    out = list.append_child("a:path");
    out.append_attribute("w") = static_cast<int64_t>(std::llround(box[2] * scale));
    out.append_attribute("h") = static_cast<int64_t>(std::llround(box[3] * scale));
  };
  auto pt = [&](pugi::xml_node parent, double x, double y) {
    pugi::xml_node p = parent.append_child("a:pt");
    p.append_attribute("x") = coord(x, 0);
    p.append_attribute("y") = coord(y, 1);
  };
  auto move = [&](double x, double y) {
    if (!out) new_path();
    pt(out.append_child("a:moveTo"), x, y);
    px = x;
    py = y;
  };
  auto line = [&](double x, double y) {
    if (!out) new_path();
    pt(out.append_child("a:lnTo"), x, y);
    px = x;
    py = y;
  };
  auto curve = [&](double x1, double y1, double x2, double y2, double x3, double y3) {
    if (!out) new_path();
    pugi::xml_node c = out.append_child("a:cubicBezTo");
    pt(c, x1, y1);
    pt(c, x2, y2);
    pt(c, x3, y3);
    px = x3;
    py = y3;
  };
  // an arc of an ellipse from angle t0 to t1 (radians, in the direction of y downwards) as curves
  auto arc = [&](double cx, double cy, double rx, double ry, double t0, double t1, bool connect, bool moving) {
    double sx = cx + rx * std::cos(t0), sy = cy + ry * std::sin(t0);
    if (moving) move(sx, sy);
    else if (connect) line(sx, sy);
    int32_t parts = std::max(1, static_cast<int32_t>(std::ceil(std::fabs(t1 - t0) / (pi / 2.0) - 1e-9)));
    for (int32_t k = 0; k < parts; ++k) {
      double a0 = t0 + (t1 - t0) * k / parts, a1 = t0 + (t1 - t0) * (k + 1) / parts;
      double kk = 4.0 / 3.0 * std::tan((a1 - a0) / 4.0);
      double x0 = cx + rx * std::cos(a0), y0 = cy + ry * std::sin(a0), x3 = cx + rx * std::cos(a1), y3 = cy + ry * std::sin(a1);
      curve(x0 - kk * rx * std::sin(a0), y0 + kk * ry * std::cos(a0), x3 + kk * rx * std::sin(a1), y3 - kk * ry * std::cos(a1), x3, y3);
    }
  };
  for (const auto& command : commands) {
    const std::vector<double>& a = command.second;
    char c = command.first;
    if (c == 'M') {
      for (size_t k = 0; k + 1 < a.size(); k += 2) {
        if (k == 0) move(a[k], a[k + 1]);
        else line(a[k], a[k + 1]);
      }
    } else if (c == 'L') {
      for (size_t k = 0; k + 1 < a.size(); k += 2) line(a[k], a[k + 1]);
    } else if (c == 'C') {
      for (size_t k = 0; k + 5 < a.size(); k += 6) curve(a[k], a[k + 1], a[k + 2], a[k + 3], a[k + 4], a[k + 5]);
    } else if (c == 'Q') {
      for (size_t k = 0; k + 3 < a.size(); k += 4) {
        curve(px + 2.0 / 3.0 * (a[k] - px), py + 2.0 / 3.0 * (a[k + 1] - py), a[k + 2] + 2.0 / 3.0 * (a[k] - a[k + 2]),
              a[k + 3] + 2.0 / 3.0 * (a[k + 1] - a[k + 3]), a[k + 2], a[k + 3]);
      }
    } else if (c == 'Z') {
      if (out) out.append_child("a:close");
    } else if (c == 'N') {
      out = pugi::xml_node();
    } else if (c == 'F' || c == 'S') {
      if (!out) new_path();
      out.append_attribute(c == 'F' ? "fill" : "stroke") = c == 'F' ? "none" : "0";
    } else if (c == 'T' || c == 'U') {
      // center, radii and angles in degrees in the direction of y downwards; some producers write angles in 1/65536 degree
      for (size_t k = 0; k + 5 < a.size(); k += 6) {
        double t0 = a[k + 4], t1 = a[k + 5];
        if (std::fabs(t0) >= 3600.0 || std::fabs(t1) >= 3600.0) {
          t0 /= 65536.0;
          t1 /= 65536.0;
        }
        while (t1 <= t0) t1 += 360.0;
        arc(a[k], a[k + 1], a[k + 2], a[k + 3], t0 / 180.0 * pi, t1 / 180.0 * pi, c == 'T' && k == 0 && out, c == 'U' && k == 0);
      }
    } else if (c == 'A' || c == 'B' || c == 'W' || c == 'V') {
      // the ellipse in a rectangle from the direction of the start point to the direction of the end point
      for (size_t k = 0; k + 7 < a.size(); k += 8) {
        double cx = (a[k] + a[k + 2]) / 2.0, cy = (a[k + 1] + a[k + 3]) / 2.0;
        double rx = std::fabs(a[k + 2] - a[k]) / 2.0, ry = std::fabs(a[k + 3] - a[k + 1]) / 2.0;
        if (!(rx > 0.0) || !(ry > 0.0)) continue;
        double t0 = std::atan2((a[k + 5] - cy) / ry, (a[k + 4] - cx) / rx), t1 = std::atan2((a[k + 7] - cy) / ry, (a[k + 6] - cx) / rx);
        bool clockwise = c == 'W' || c == 'V';
        if (clockwise) {
          while (t1 <= t0) t1 += 2.0 * pi;
        } else {
          while (t1 >= t0) t1 -= 2.0 * pi;
        }
        arc(cx, cy, rx, ry, t0, t1, (c == 'A' || c == 'W') && out, c == 'B' || c == 'V');
      }
    } else if (c == 'X' || c == 'Y') {
      // quarter ellipses that start horizontally or vertically, alternating
      bool horizontal = c == 'X';
      for (size_t k = 0; k + 1 < a.size(); k += 2) {
        const double kk = 0.5522847498;
        double x = a[k], y = a[k + 1];
        if (horizontal) curve(px + kk * (x - px), py, x, y + kk * (py - y), x, y);
        else curve(px, py + kk * (y - py), x + kk * (px - x), y, x, y);
        horizontal = !horizontal;
      }
    }
  }
  if (!list.first_child()) return "";
  return ods_xml_string(geom);
}

static std::string ods_shape_type(pugi::xml_node shape) {
  std::string name = shape.name();
  if (name == "draw:ellipse" || name == "draw:circle") return "ellipse";
  if (name == "draw:rect") return shape.attribute("draw:corner-radius") ? "roundRect" : "rect";
  std::string type = shape.child("draw:enhanced-geometry").attribute("draw:type").as_string();
  for (const auto& s : ods_shapes) {
    if (type == s[1]) return s[0];
  }
  if (type.compare(0, 6, "ooxml-") == 0) return type.substr(6);
  return "rect";
}

// the errors shown as text of cells, also with German names, and their xlsx errors
static const char* const ods_errors[][2] = {
  {"#NULL!", "#NULL!"}, {"#DIV/0!", "#DIV/0!"}, {"#VALUE!", "#VALUE!"}, {"#REF!", "#REF!"}, {"#NAME?", "#NAME?"},
  {"#NUM!", "#NUM!"}, {"#N/A", "#N/A"}, {"#WERT!", "#VALUE!"}, {"#BEZUG!", "#REF!"}, {"#ZAHL!", "#NUM!"}, {"#NV", "#N/A"}
};

// a formula cell of a producer without the calc extension with an error as its text and a value of 0
static bool ods_formula_error(pugi::xml_node cell, pugi::xml_attribute formula) {
  if (!formula || cell.attribute("calcext:value-type")) return false;
  std::string text = ods_cell_string(cell);
  for (const auto& error : ods_errors) {
    if (text == error[0]) return true;
  }
  return false;
}

// all_sheets loads all tables, otherwise only the one selected by its name or
// position in sheet_selected; data_only leaves out frames, shapes, comments and pivot tables.
// [[Rcpp::export]]
Rcpp::List read_ods_parts(native_path content_path, native_path styles_path, native_path settings_path, native_path meta_path, native_path dir,
                          bool all_sheets, Rcpp::RObject sheet_selected, bool data_only) {
  uint32_t flags = pugi::parse_default | pugi::parse_ws_pcdata;
  pugi::xml_document content, styles_doc;
  pugi::xml_parse_result res = content.load_file(content_path.c_str(), flags, pugi::encoding_utf8);
  if (!res) Rcpp::stop("could not read content.xml: %s", res.description());
  ods_normalize_document(content);
  pugi::xml_node croot = content.document_element();
  pugi::xml_node sroot = croot;
  if (!styles_path.empty()) {
    res = styles_doc.load_file(styles_path.c_str(), flags, pugi::encoding_utf8);
    if (!res) Rcpp::stop("could not read styles.xml: %s", res.description());
    ods_normalize_document(styles_doc);
    sroot = styles_doc.document_element();
  }

  ods_xlsx_styles st;
  pugi::xml_node containers[] = {
    croot.child("office:automatic-styles"), croot.child("office:styles"),
    sroot.child("office:styles"), sroot.child("office:automatic-styles")
  };
  std::map<std::string, pugi::xml_node> col_styles, row_styles, table_styles, graphic_styles, paragraph_styles, gradient_styles, dash_styles, marker_styles;
  for (pugi::xml_node container : containers) {
    for (pugi::xml_node gradient : container.children("draw:gradient")) {
      gradient_styles.insert(std::make_pair(std::string(gradient.attribute("draw:name").as_string()), gradient));
    }
    for (pugi::xml_node dash : container.children("draw:stroke-dash")) {
      dash_styles.insert(std::make_pair(std::string(dash.attribute("draw:name").as_string()), dash));
    }
    for (pugi::xml_node marker : container.children("draw:marker")) {
      marker_styles.insert(std::make_pair(std::string(marker.attribute("draw:name").as_string()), marker));
    }
    for (pugi::xml_node node : container.children()) {
      std::string name = node.name();
      std::string style_name = node.attribute("style:name").as_string();
      std::string family = node.attribute("style:family").as_string();
      if (name == "style:style") {
        if (family == "table-cell") st.cell.insert(std::make_pair(style_name, node));
        else if (family == "text") st.text.insert(std::make_pair(style_name, node));
        else if (family == "table-column") col_styles.insert(std::make_pair(style_name, node));
        else if (family == "table-row") row_styles.insert(std::make_pair(style_name, node));
        else if (family == "table") table_styles.insert(std::make_pair(style_name, node));
        else if (family == "graphic") graphic_styles.insert(std::make_pair(style_name, node));
        else if (family == "paragraph") paragraph_styles.insert(std::make_pair(style_name, node));
      } else if (name == "style:default-style" && family == "table-cell") {
        if (!st.default_cell) st.default_cell = node;
      } else if (name.compare(0, 7, "number:") == 0) {
        st.data.insert(std::make_pair(style_name, node));
      }
    }
  }
  for (pugi::xml_node root : {croot, sroot}) {
    for (pugi::xml_node face : root.child("office:font-face-decls").children("style:font-face")) {
      st.font_face.insert(std::make_pair(std::string(face.attribute("style:name").as_string()), std::string(face.attribute("svg:font-family").as_string())));
    }
  }

  // xf 0 is the default style, its font gives the width of a column unit in points
  st.xf("Default");
  pugi::xml_document default_font_doc;
  default_font_doc.load_string(st.font_list.empty() ? "" : st.font_list[0].c_str());
  pugi::xml_node default_font = default_font_doc.child("font");
  double unit_pt = ods_digit_width(default_font.child("name").attribute("val").as_string("Calibri"), default_font.child("sz").attribute("val").as_double(11.0)) * 0.75;

  pugi::xml_node spreadsheet = croot.child("office:body").child("office:spreadsheet");

  // external references are cached in tables named after the file and the sheet, scenarios are
  // tables with the values of another table; these are no sheets
  std::vector<pugi::xml_node> caches;
  for (pugi::xml_node table : spreadsheet.children("table:table")) {
    if (table.child("table:table-source") && table.attribute("table:name").as_string()[0] == '\'') caches.push_back(table);
    else if (table.child("table:scenario")) caches.push_back(table);
  }
  for (pugi::xml_node table : caches) spreadsheet.remove_child(table);

  // tables without name are named after their position
  {
    std::set<std::string> existing;
    for (pugi::xml_node table : spreadsheet.children("table:table")) existing.insert(table.attribute("table:name").as_string());
    int32_t position = 0;
    for (pugi::xml_node table : spreadsheet.children("table:table")) {
      ++position;
      if (*table.attribute("table:name").as_string()) continue;
      std::string name = "Sheet" + std::to_string(position);
      for (int32_t k = 2; existing.count(name); ++k) name = "Sheet" + std::to_string(position) + "_" + std::to_string(k);
      existing.insert(name);
      if (!table.attribute("table:name")) table.append_attribute("table:name");
      table.attribute("table:name") = name.c_str();
    }
  }

  // xlsx sheet names have at most 31 characters, none of []*?:/\ and are unique regardless of case
  std::map<std::string, std::string> renames;
  {
    auto lower = [](std::string x) {
      for (char& ch : x) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
      return x;
    };
    auto truncate = [](const std::string& x, size_t chars) {
      size_t count = 0, k = 0;
      for (; k < x.size(); ++k) {
        if ((static_cast<unsigned char>(x[k]) & 0xC0) != 0x80 && count++ == chars) break;
      }
      return x.substr(0, k);
    };
    std::set<std::string> taken;
    std::vector<std::string> invalid;
    for (pugi::xml_node table : spreadsheet.children("table:table")) {
      std::string name = table.attribute("table:name").as_string();
      if (truncate(name, 31) == name && name.find_first_of("[]*?:/\\") == std::string::npos) taken.insert(lower(name));
      else invalid.push_back(name);
    }
    for (const std::string& name : invalid) {
      std::string clean = name;
      for (char& ch : clean) {
        if (std::strchr("[]*?:/\\", ch)) ch = '_';
      }
      std::string candidate = truncate(clean, 31);
      for (int32_t k = 2; taken.count(lower(candidate)); ++k) {
        std::string suffix = "_" + std::to_string(k);
        candidate = truncate(clean, 31 - suffix.size()) + suffix;
      }
      taken.insert(lower(candidate));
      renames[name] = candidate;
    }
  }

  std::string null_date = spreadsheet.child("table:calculation-settings").child("table:null-date").attribute("table:date-value").as_string("1899-12-30");
  bool date1904 = null_date.compare(0, 10, "1904-01-01") == 0;

  std::string path = std::string(dir) + "/";
  std::vector<std::string> names;
  std::vector<bool> hidden_sheets;
  int32_t dropped = 0;

  struct ods_name {
    std::string name, local, formula;
  };
  std::vector<ods_name> defined;
  std::vector<std::pair<pugi::xml_node, std::string>> name_containers;
  name_containers.push_back(std::make_pair(spreadsheet.child("table:named-expressions"), std::string()));
  R_xlen_t iter = 0;

  // view settings: freeze panes, grid lines, zoom, the cursor, the top left cell and the active sheet
  struct ods_view {
    int64_t rows = 0, cols = 0, zoom = 0, page_zoom = 0, cursor_col = 0, cursor_row = 0, left = 0, top = 0;
    int32_t grid = -1;
    bool zeros = true, headers = true, split = false;
  };
  std::map<std::string, ods_view> views;
  std::string active_table;
  bool page_break = false, show_zeros = true, show_headers = true, show_tabs = true, show_outline = true;
  int64_t grid_color = -1;
  pugi::xml_document settings_doc;
  if (!settings_path.empty() && settings_doc.load_file(settings_path.c_str(), pugi::parse_default, pugi::encoding_utf8)) {
    ods_normalize_document(settings_doc);
    auto item = [](pugi::xml_node entry, const char* name) {
      for (pugi::xml_node it : entry.children("config:config-item")) {
        if (std::strcmp(it.attribute("config:name").as_string(), name) == 0) return std::string(it.text().get());
      }
      return std::string();
    };
    for (pugi::xml_node set : settings_doc.document_element().child("office:settings").children("config:config-item-set")) {
      if (std::strcmp(set.attribute("config:name").as_string(), "ooo:view-settings") != 0) continue;
      for (pugi::xml_node map : set.children("config:config-item-map-indexed")) {
        if (std::strcmp(map.attribute("config:name").as_string(), "Views") != 0) continue;
        pugi::xml_node entry = map.child("config:config-item-map-entry");
        active_table = item(entry, "ActiveTable");
        page_break = item(entry, "ShowPageBreakPreview") == "true";
        show_zeros = item(entry, "ShowZeroValues") != "false";
        show_headers = item(entry, "HasColumnRowHeaders") != "false";
        show_tabs = item(entry, "HasSheetTabs") != "false";
        show_outline = item(entry, "IsOutlineSymbolsSet") != "false";
        std::string grid_color_item = item(entry, "GridColor");
        if (!grid_color_item.empty()) grid_color = std::atoll(grid_color_item.c_str());
        std::string global_grid = item(entry, "ShowGrid");
        for (pugi::xml_node named : entry.children("config:config-item-map-named")) {
          if (std::strcmp(named.attribute("config:name").as_string(), "Tables") != 0) continue;
          for (pugi::xml_node te : named.children("config:config-item-map-entry")) {
            ods_view view;
            // frozen rows and columns, or a split in pixels
            std::string mode_cols = item(te, "HorizontalSplitMode"), mode_rows = item(te, "VerticalSplitMode");
            bool split_cols = mode_cols == "1" || mode_cols == "2", split_rows = mode_rows == "1" || mode_rows == "2";
            view.split = mode_cols == "1" || mode_rows == "1";
            if (split_cols) view.cols = std::atoll(item(te, "HorizontalSplitPosition").c_str());
            if (split_rows) view.rows = std::atoll(item(te, "VerticalSplitPosition").c_str());
            if (view.split && mode_cols == "2") view.cols = 0;
            if (view.split && mode_rows == "2") view.rows = 0;
            std::string grid = item(te, "ShowGrid");
            if (grid.empty()) grid = global_grid;
            if (!grid.empty()) view.grid = grid == "true" ? 1 : 0;
            // zero values and headers are settings of the document, some consumers have them for each sheet
            std::string zeros = item(te, "ShowZeroValues"), headers = item(te, "HasColumnRowHeaders");
            view.zeros = zeros.empty() ? show_zeros : zeros != "false";
            view.headers = headers.empty() ? show_headers : headers != "false";
            view.zoom = std::atoll(item(te, "ZoomValue").c_str());
            view.page_zoom = std::atoll(item(te, "PageViewZoomValue").c_str());
            view.cursor_col = std::max<int64_t>(std::atoll(item(te, "CursorPositionX").c_str()), 0);
            view.cursor_row = std::max<int64_t>(std::atoll(item(te, "CursorPositionY").c_str()), 0);
            // the first column and row of the view without split, of the right and bottom part with it
            view.left = std::max<int64_t>(std::atoll(item(te, split_cols ? "PositionRight" : "PositionLeft").c_str()), 0);
            view.top = std::max<int64_t>(std::atoll(item(te, "PositionBottom").c_str()), 0);
            views[te.attribute("config:name").as_string()] = view;
          }
        }
      }
    }
  }

  std::string base = std::string(content_path).substr(0, std::string(content_path).rfind('/') + 1);
  std::vector<pugi::xml_node> page_containers = {sroot.child("office:automatic-styles"), croot.child("office:automatic-styles"), sroot.child("office:styles")};
  std::vector<pugi::xml_node> master_containers = {sroot.child("office:master-styles"), croot.child("office:master-styles")};
  Rcpp::List sheet_extras;

  std::map<std::string, pugi::xml_node> validations;
  for (pugi::xml_node cv : spreadsheet.child("table:content-validations").children("table:content-validation")) {
    validations[cv.attribute("table:name").as_string()] = cv;
  }

  // database ranges: anonymous ranges of a sheet with filter buttons are autofilters, named ranges are tables
  struct ods_db_range {
    std::string name, ref;
    bool filter, header;
  };
  std::map<std::string, std::vector<ods_db_range>> db_ranges;
  for (pugi::xml_node db : spreadsheet.child("table:database-ranges").children("table:database-range")) {
    std::string sheet, ref;
    if (!ods_split_range(db.attribute("table:target-range-address").as_string(), sheet, ref)) continue;
    ods_db_range range = {db.attribute("table:name").as_string(), ref,
                          db.attribute("table:display-filter-buttons").as_bool(), db.attribute("table:contains-header").as_bool(true)};
    if (range.ref.find(':') == std::string::npos) range.ref += ":" + range.ref;
    db_ranges[sheet].push_back(range);
  }
  std::vector<std::string> dxf_list;
  std::map<std::string, int64_t> dxf_index;
  size_t table_parts = 0;
  std::set<std::string> media_extensions;
  int32_t background_images = 0;

  // data pilot tables are created as pivot tables of the loaded workbook. The
  // table body starts below the page fields and an empty row.
  static const char* const pilot_functions[][2] = {
    {"sum", "sum"}, {"count", "count"}, {"average", "average"}, {"max", "max"}, {"min", "min"},
    {"product", "product"}, {"countnums", "countNums"}, {"stdev", "stdDev"}, {"stdevp", "stdDevp"},
    {"var", "var"}, {"varp", "varp"}
  };
  std::map<std::string, Rcpp::List> pilots;
  for (pugi::xml_node dp : data_only ? pugi::xml_node() : spreadsheet.child("table:data-pilot-tables")) {
    if (std::strcmp(dp.name(), "table:data-pilot-table") != 0) continue;
    std::string target_sheet, target, source_sheet, source;
    if (!ods_split_range(dp.attribute("table:target-range-address").as_string(), target_sheet, target)) continue;
    if (!ods_split_range(dp.child("table:source-cell-range").attribute("table:cell-range-address").as_string(), source_sheet, source)) continue;
    std::vector<utf8_string> rows, cols, filter, data, fun, choose_names, choose_values, show_data_as;
    int64_t sort_row = 0, sort_col = 0;
    std::string data_layout, layout;
    for (pugi::xml_node field : dp.children("table:data-pilot-field")) {
      std::string orientation = field.attribute("table:orientation").as_string();
      std::string name = field.attribute("table:source-field-name").as_string();
      if (field.attribute("table:is-data-layout-field").as_bool()) {
        data_layout = orientation;
        continue;
      }
      // the sort order of a row or column field, the first sorted field is kept
      pugi::xml_node sort_info = field.child("table:data-pilot-level").child("table:data-pilot-sort-info");
      std::string sort_order = std::string(sort_info.attribute("table:sort-mode").as_string()) == "name" ? sort_info.attribute("table:order").as_string("ascending") : "";
      int64_t sort_sign = sort_order == "descending" ? -1 : 1;
      if (orientation == "row" && layout.empty()) {
        pugi::xml_node layout_info = field.child("table:data-pilot-level").child("table:data-pilot-layout-info");
        std::string mode = layout_info.attribute("table:layout-mode").as_string();
        if (std::string(layout_info.attribute("loext:layout-mode").as_string()) == "compact-layout") layout = "compact";
        else if (mode == "outline-subtotals-top" || mode == "outline-subtotals-bottom") layout = "outline";
        else if (!mode.empty()) layout = "tabular";
      }
      if (orientation == "row") rows.push_back(name);
      else if (orientation == "column") cols.push_back(name);
      else if (orientation == "page") filter.push_back(name);
      if (orientation == "row" && !sort_order.empty() && sort_row == 0) sort_row = sort_sign * static_cast<int64_t>(rows.size());
      if (orientation == "column" && !sort_order.empty() && sort_col == 0) sort_col = sort_sign * static_cast<int64_t>(cols.size());
      if (orientation == "data") {
        static const char* const shown[][2] = {
          {"member-difference", "difference"}, {"member-percentage", "percent"}, {"member-percentage-difference", "percentDiff"},
          {"running-total", "runTotal"}, {"row-percentage", "percentOfRow"}, {"column-percentage", "percentOfCol"},
          {"total-percentage", "percentOfTotal"}, {"index", "index"}
        };
        std::string type = field.child("table:data-pilot-field-reference").attribute("table:type").as_string(), show = "normal";
        for (const auto& sh : shown) {
          if (type == sh[0]) show = sh[1];
        }
        show_data_as.push_back(show);
        data.push_back(name);
        std::string function = field.attribute("table:function").as_string("sum");
        std::string mapped = "sum";
        for (const auto& fn : pilot_functions) {
          if (function == fn[0]) mapped = fn[1];
        }
        fun.push_back(mapped);
      }
      if (orientation != "row" && orientation != "column" && orientation != "page") continue;
      // a selected page or the hidden members as filter of the field
      std::string selected = field.attribute("table:selected-page").as_string(), choose;
      if (field.attribute("loext:ignore-selected-page").as_bool()) selected.clear();
      // the members are compared with %in%, which handles missing values
      std::string members;
      if (!selected.empty()) {
        members = ods_r_string(selected);
      } else {
        for (pugi::xml_node member : field.child("table:data-pilot-level").child("table:data-pilot-members").children("table:data-pilot-member")) {
          if (member.attribute("table:display").as_bool(true)) continue;
          members += (members.empty() ? "" : ", ") + ods_r_string(member.attribute("table:name").as_string());
        }
      }
      if (!selected.empty()) choose = "x %in% c(" + members + ")";
      else if (!members.empty()) choose = "(x %in% c(" + members + ")) == FALSE";
      if (choose.empty()) continue;
      choose_names.push_back(name);
      choose_values.push_back(choose);
    }
    if (data.empty()) continue;
    std::string top_left = target.substr(0, target.find(':'));
    int64_t body_row = cell_to_rowint(top_left);
    if (!filter.empty()) body_row += static_cast<int64_t>(filter.size()) + 1;
    std::string grand = dp.attribute("table:grand-total").as_string("both");
    pilots[target_sheet].push_back(Rcpp::List::create(
      Rcpp::Named("name") = utf8_string(dp.attribute("table:name").as_string()),
      Rcpp::Named("target") = utf8_string(target),
      Rcpp::Named("dims") = utf8_string(int_to_col(cell_to_colint(top_left)) + std::to_string(body_row)),
      Rcpp::Named("source_sheet") = utf8_string(source_sheet), Rcpp::Named("source") = utf8_string(source),
      Rcpp::Named("rows") = rows, Rcpp::Named("cols") = cols, Rcpp::Named("filter") = filter,
      Rcpp::Named("data") = data, Rcpp::Named("fun") = fun,
      Rcpp::Named("choose_names") = choose_names, Rcpp::Named("choose") = choose_values,
      Rcpp::Named("data_on_rows") = data_layout == "row",
      Rcpp::Named("row_grand_totals") = grand == "both" || grand == "row",
      Rcpp::Named("col_grand_totals") = grand == "both" || grand == "column",
      Rcpp::Named("show_data_as") = show_data_as, Rcpp::Named("sort_row") = static_cast<double>(sort_row),
      Rcpp::Named("sort_col") = static_cast<double>(sort_col), Rcpp::Named("layout") = utf8_string(layout)
    ));
  }

  for (pugi::xml_node table : spreadsheet.children("table:table")) {
    names.push_back(table.attribute("table:name").as_string());
    // the cells of tables that are not selected are not loaded
    bool keep = all_sheets;
    if (Rf_isString(sheet_selected) && Rf_xlength(sheet_selected) > 0) keep = names.back() == Rcpp::as<std::vector<utf8_string>>(sheet_selected)[0];
    if (Rf_isNumeric(sheet_selected) && Rf_xlength(sheet_selected) > 0) keep = static_cast<double>(names.size()) == Rcpp::as<std::vector<double>>(sheet_selected)[0];
    name_containers.push_back(std::make_pair(table.child("table:named-expressions"), std::to_string(names.size() - 1)));
    auto ts = table_styles.find(table.attribute("table:style-name").as_string());
    hidden_sheets.push_back(ts != table_styles.end() &&
                            std::string(ts->second.child("style:table-properties").attribute("table:display").as_string()) == "false");
    bool table_protected = table.attribute("table:protected").as_bool();

    std::vector<ods_in_frame> frames;
    struct ods_in_valid {
      std::string name;
      int64_t r1, c1, r2, c2;
    };
    std::vector<ods_in_valid> valid_cells;
    // frames and shapes of the sheet or of a cell; groups are not kept, their elements are
    std::function<void(pugi::xml_node, int64_t, int64_t)> read_frames;
    read_frames = [&](pugi::xml_node parent, int64_t frame_row, int64_t frame_col) {
      for (pugi::xml_node frame : parent.children()) {
        std::string frame_name = frame.name();
        if (frame_name == "draw:g") {
          read_frames(frame, frame_row, frame_col);
          continue;
        }
        bool line = frame_name == "draw:line" || frame_name == "draw:connector";
        bool shape = line || frame_name == "draw:custom-shape" || frame_name == "draw:rect" || frame_name == "draw:ellipse" || frame_name == "draw:circle";
        if (frame_name != "draw:frame" && frame_name != "draw:control" && !shape) continue;
        ods_in_frame fr = ods_read_frame(frame);
        fr.row = frame_row;
        fr.col = frame_col;
        fr.control = frame.attribute("draw:control").as_string();
        if (frame_name == "draw:control" && fr.control.empty()) continue;
        pugi::xml_document props_doc;
        pugi::xml_node props = props_doc.append_child("p");
        if (line) {
          // the box of the end points, a line from the right or bottom is mirrored; the end follows from the size
          double x1 = ods_length_pt(frame.attribute("svg:x1").as_string("0pt")), y1 = ods_length_pt(frame.attribute("svg:y1").as_string("0pt"));
          double x2 = ods_length_pt(frame.attribute("svg:x2").as_string("0pt")), y2 = ods_length_pt(frame.attribute("svg:y2").as_string("0pt"));
          fr.x = std::min(x1, x2);
          fr.y = std::min(y1, y2);
          fr.width = std::fabs(x2 - x1);
          fr.height = std::fabs(y2 - y1);
          fr.end_cell.clear();
          if (x1 > x2) props.append_attribute("flipH") = 1;
          if (y1 > y2) props.append_attribute("flipV") = 1;
          std::string type = frame.attribute("draw:type").as_string("standard");
          fr.shape = frame_name == "draw:line" || type == "line" ? "straightConnector1" : type == "curve" ? "curvedConnector3" : "bentConnector3";
        }
        if (shape && !line) {
          fr.shape = ods_shape_type(frame);
          pugi::xml_node eg = frame.child("draw:enhanced-geometry");
          if (eg.attribute("draw:mirror-horizontal").as_bool()) props.append_attribute("flipH") = 1;
          if (eg.attribute("draw:mirror-vertical").as_bool()) props.append_attribute("flipV") = 1;
          // shapes without a preset keep their outline
          std::string type = frame.child("draw:enhanced-geometry").attribute("draw:type").as_string();
          bool preset = type.compare(0, 6, "ooxml-") == 0;
          for (const auto& known : ods_shapes) {
            if (type == known[1]) preset = true;
          }
          if (frame_name == "draw:custom-shape" && !preset) {
            fr.geometry = ods_enhanced_custgeom(frame.child("draw:enhanced-geometry"), ods_length_pt(frame.attribute("svg:width").as_string("0pt")),
                                                ods_length_pt(frame.attribute("svg:height").as_string("0pt")));
          }
          std::string rich = ods_cell_rich(frame, st, "");
          fr.rich = !rich.empty();
          fr.text = fr.rich ? rich : ods_cell_string(frame);
          auto ps = paragraph_styles.find(frame.child("text:p").attribute("text:style-name").as_string());
          std::string align = ps != paragraph_styles.end() ? ps->second.child("style:paragraph-properties").attribute("fo:text-align").as_string() : "";
          if (align == "center") fr.align = "center";
          else if (align == "end" || align == "right") fr.align = "right";
          else if (align == "justify") fr.align = "justify";
          auto gs = graphic_styles.find(frame.attribute("draw:style-name").as_string());
          pugi::xml_node gp = gs != graphic_styles.end() ? gs->second.child("style:graphic-properties") : pugi::xml_node();
          fr.fill = std::string(gp.attribute("draw:fill").as_string()) == "none" ? "none" : ods_rgb(gp.attribute("draw:fill-color").as_string());
          fr.line = std::string(gp.attribute("draw:stroke").as_string()) == "none" ? "none" : ods_rgb(gp.attribute("svg:stroke-color").as_string());
          // a gradient as linear gradient of its start and end color
          auto gradient = gradient_styles.find(gp.attribute("draw:fill-gradient-name").as_string());
          if (std::string(gp.attribute("draw:fill").as_string()) == "gradient" && gradient != gradient_styles.end()) {
            fr.gradient_start = ods_rgb(gradient->second.attribute("draw:start-color").as_string());
            fr.gradient_end = ods_rgb(gradient->second.attribute("draw:end-color").as_string());
            // the angle is in tenths of a degree without unit
            std::string angle = gradient->second.attribute("draw:angle").as_string("0");
            double degree = std::strtod(angle.c_str(), nullptr);
            if (angle.find_first_not_of("0123456789.-") == std::string::npos) degree /= 10.0;
            else if (angle.find("rad") != std::string::npos) degree *= 180.0 / 3.14159265358979323846;
            else if (angle.find("grad") != std::string::npos) degree *= 0.9;
            fr.gradient_angle = std::fmod(std::fmod(90.0 - degree, 360.0) + 360.0, 360.0);
            if (fr.fill.empty() || fr.fill == "none") fr.fill = fr.gradient_start;
          }
        }
        if (shape) {
          auto gs = graphic_styles.find(frame.attribute("draw:style-name").as_string());
          pugi::xml_node gp = gs != graphic_styles.end() ? gs->second.child("style:graphic-properties") : pugi::xml_node();
          if (line) {
            fr.fill = "none";
            fr.line = std::string(gp.attribute("draw:stroke").as_string()) == "none" ? "none" : ods_rgb(gp.attribute("svg:stroke-color").as_string());
          }
          auto emu = [](double pt) { return static_cast<int64_t>(std::llround(pt * ods_emu_per_pt)); };
          // a color with its opacity in percent
          auto solid_fill = [](pugi::xml_node target, const std::string& rgb, const char* opacity) {
            pugi::xml_node clr = target.append_child("a:srgbClr");
            clr.append_attribute("val") = rgb.substr(2).c_str();
            double alpha = std::strtod(opacity, nullptr);
            if (alpha < 100.0) clr.append_child("a:alpha").append_attribute("val") = static_cast<int64_t>(std::llround(std::max(alpha, 0.0) * 1000.0));
          };
          std::string fill = gp.attribute("draw:fill").as_string();
          std::string fill_color = ods_rgb(gp.attribute("draw:fill-color").as_string());
          if (line || fill == "none") props.append_child("a:noFill");
          else if (fill == "solid" && !fill_color.empty()) solid_fill(props.append_child("a:solidFill"), fill_color, gp.attribute("draw:opacity").as_string("100"));
          // gradients from the border to a center as path gradients from the center to the border
          auto gradient = gradient_styles.find(gp.attribute("draw:fill-gradient-name").as_string());
          std::string gradient_style = gradient != gradient_styles.end() ? gradient->second.attribute("draw:style").as_string() : "";
          if (fill == "gradient" && !line && (gradient_style == "radial" || gradient_style == "ellipsoid" || gradient_style == "square" || gradient_style == "rectangular")) {
            std::string outer = ods_rgb(gradient->second.attribute("draw:start-color").as_string());
            std::string center = ods_rgb(gradient->second.attribute("draw:end-color").as_string());
            if (!outer.empty() && !center.empty()) {
              pugi::xml_node grad = props.append_child("a:gradFill");
              pugi::xml_node list = grad.append_child("a:gsLst");
              const std::string colors[] = {center, outer};
              for (size_t k = 0; k < 2; ++k) {
                pugi::xml_node stop = list.append_child("a:gs");
                stop.append_attribute("pos") = k == 0 ? 0 : 100000;
                stop.append_child("a:srgbClr").append_attribute("val") = colors[k].substr(2).c_str();
              }
              pugi::xml_node path_node = grad.append_child("a:path");
              path_node.append_attribute("path") = gradient_style == "radial" || gradient_style == "ellipsoid" ? "circle" : "rect";
              double cx = std::strtod(gradient->second.attribute("draw:cx").as_string("50"), nullptr);
              double cy = std::strtod(gradient->second.attribute("draw:cy").as_string("50"), nullptr);
              pugi::xml_node rect = path_node.append_child("a:fillToRect");
              rect.append_attribute("l") = static_cast<int64_t>(std::llround(cx * 1000.0));
              rect.append_attribute("t") = static_cast<int64_t>(std::llround(cy * 1000.0));
              rect.append_attribute("r") = static_cast<int64_t>(std::llround((100.0 - cx) * 1000.0));
              rect.append_attribute("b") = static_cast<int64_t>(std::llround((100.0 - cy) * 1000.0));
              fr.gradient_start.clear();
              fr.gradient_end.clear();
            }
          }

          // the line with its width, dash and ends
          std::string stroke = gp.attribute("draw:stroke").as_string();
          double line_w = ods_length_pt(gp.attribute("svg:stroke-width").as_string("0pt"));
          if (stroke == "none") {
            props.append_child("a:ln").append_child("a:noFill");
          } else if (stroke == "solid" || stroke == "dash") {
            pugi::xml_node ln = props.append_child("a:ln");
            if (line_w > 0.0) ln.append_attribute("w") = emu(line_w);
            std::string stroke_color = ods_rgb(gp.attribute("svg:stroke-color").as_string("#000000"));
            if (!stroke_color.empty()) solid_fill(ln.append_child("a:solidFill"), stroke_color, gp.attribute("svg:stroke-opacity").as_string("100"));
            auto dash = dash_styles.find(gp.attribute("draw:stroke-dash").as_string());
            if (stroke == "dash" && dash != dash_styles.end()) {
              const char* preset = ods_dash_preset(dash->second, line_w);
              if (preset) ln.append_child("a:prstDash").append_attribute("val") = preset;
            }
            // the line ends: the markers of spreadsheet software by their name, others by their shape names
            static const char* const ends[2][2] = {{"draw:marker-start", "a:headEnd"}, {"draw:marker-end", "a:tailEnd"}};
            for (const auto& end : ends) {
              auto marker = marker_styles.find(gp.attribute(end[0]).as_string());
              if (marker == marker_styles.end()) continue;
              std::string name = marker->second.attribute("draw:display-name").as_string(marker->first.c_str());
              for (size_t pos = name.find("_20_"); pos != std::string::npos; pos = name.find("_20_")) name.replace(pos, 4, " ");
              const char* type = "triangle";
              int32_t size = 0;
              bool known = false;
              for (const ods_marker& m : ods_markers) {
                size_t n = std::strlen(m.name);
                if (name.compare(0, n, m.name) == 0 && (name.size() == n || name[n] == ' ')) {
                  type = m.type;
                  known = true;
                  size = name.size() > n ? std::atoi(name.c_str() + n + 1) : 0;
                }
              }
              if (!known) {
                std::string lower = ods_lower(name);
                if (lower.find("circle") != std::string::npos || lower.find("oval") != std::string::npos) type = "oval";
                else if (lower.find("square 45") != std::string::npos || lower.find("diamond") != std::string::npos) type = "diamond";
                else if (lower.find("line") != std::string::npos || lower.find("open") != std::string::npos) type = "arrow";
                else if (lower.find("concave") != std::string::npos || lower.find("stealth") != std::string::npos) type = "stealth";
              }
              static const char* const sizes[] = {"sm", "med", "lg"};
              int32_t w = 1, len = 1;
              if (size >= 1 && size <= 9) {
                w = (size - 1) / 3;
                len = (size - 1) % 3;
              } else {
                pugi::xml_attribute width = gp.attribute((std::string(end[0]) + "-width").c_str());
                double ratio = ods_length_pt(width.as_string("0pt")) / std::max(line_w, 2.0);
                if (width) w = ratio <= 2.5 ? 0 : ratio <= 4.0 ? 1 : 2;
              }
              pugi::xml_node node = ln.append_child(end[1]);
              node.append_attribute("type") = type;
              node.append_attribute("w") = sizes[w];
              node.append_attribute("len") = sizes[len];
            }
            if (!ln.child("a:headEnd") && ln.child("a:tailEnd")) ln.insert_child_before("a:headEnd", ln.child("a:tailEnd")).append_attribute("type") = "none";
          }

          // the shadow, glow and soft edges
          pugi::xml_node effects = props.append_child("a:effectLst");
          double glow = ods_length_pt(gp.attribute("loext:glow-radius").as_string("0pt"));
          std::string glow_color = ods_rgb(gp.attribute("loext:glow-color").as_string());
          if (glow > 0.0 && !glow_color.empty()) {
            pugi::xml_node node = effects.append_child("a:glow");
            node.append_attribute("rad") = emu(glow);
            solid_fill(node, glow_color, ods_num(100.0 - std::strtod(gp.attribute("loext:glow-transparency").as_string("0"), nullptr)).c_str());
          }
          if (std::string(gp.attribute("draw:shadow").as_string()) == "visible") {
            double sx = ods_length_pt(gp.attribute("draw:shadow-offset-x").as_string("0pt")), sy = ods_length_pt(gp.attribute("draw:shadow-offset-y").as_string("0pt"));
            double angle = std::atan2(sy, sx) * 180.0 / 3.14159265358979323846;
            if (angle < 0.0) angle += 360.0;
            pugi::xml_node node = effects.append_child("a:outerShdw");
            node.append_attribute("blurRad") = emu(ods_length_pt(gp.attribute("loext:shadow-blur").as_string("0pt")));
            node.append_attribute("dist") = emu(std::sqrt(sx * sx + sy * sy));
            node.append_attribute("dir") = static_cast<int64_t>(std::llround(angle * 60000.0)) % 21600000;
            node.append_attribute("rotWithShape") = 0;
            std::string shadow_color = ods_rgb(gp.attribute("draw:shadow-color").as_string("#808080"));
            solid_fill(node, shadow_color.empty() ? std::string("FF808080") : shadow_color, gp.attribute("draw:shadow-opacity").as_string("100"));
          }
          double soft = ods_length_pt(gp.attribute("loext:softedge-radius").as_string("0pt"));
          if (soft > 0.0) effects.append_child("a:softEdge").append_attribute("rad") = emu(soft);
          if (!effects.first_child()) props.remove_child(effects);

          // the text box: insets, vertical anchor, wrapping and growing with the text
          if (!line) {
            pugi::xml_node body = props.append_child("a:bodyPr");
            static const char* const insets[4][2] = {{"fo:padding-left", "lIns"}, {"fo:padding-top", "tIns"}, {"fo:padding-right", "rIns"}, {"fo:padding-bottom", "bIns"}};
            for (const auto& inset : insets) {
              pugi::xml_attribute padding = gp.attribute(inset[0]);
              if (!padding) padding = gp.attribute("fo:padding");
              if (padding) body.append_attribute(inset[1]) = emu(ods_length_pt(padding.as_string()));
            }
            std::string valign = gp.attribute("draw:textarea-vertical-align").as_string();
            if (valign == "top" || valign == "middle" || valign == "bottom") body.append_attribute("anchor") = valign == "top" ? "t" : valign == "middle" ? "ctr" : "b";
            std::string wrap = gp.attribute("fo:wrap-option").as_string();
            if (!wrap.empty()) body.append_attribute("wrap") = wrap == "no-wrap" ? "none" : "square";
            if (gp.attribute("draw:auto-grow-height").as_bool()) body.append_child("a:spAutoFit");
          }
          fr.properties = ods_xml_string(props);
        } else {
          // the clip of a picture is relative to the size of the image
          auto gs = graphic_styles.find(frame.attribute("draw:style-name").as_string());
          std::string clip = gs != graphic_styles.end() ? gs->second.child("style:graphic-properties").attribute("fo:clip").as_string() : "";
          double image_w = 0.0, image_h = 0.0;
          std::string image = frame.child("draw:image").attribute("xlink:href").as_string();
          if (clip.compare(0, 5, "rect(") == 0 && !image.empty() && ods_image_size(base + image, image_w, image_h)) {
            std::vector<double> sides;
            size_t pos = 5;
            while (pos < clip.size() && sides.size() < 4) {
              size_t end = clip.find_first_of(",)", pos);
              if (end == std::string::npos) break;
              std::string value = clip.substr(pos, end - pos);
              value.erase(0, value.find_first_not_of(' '));
              sides.push_back(value == "auto" ? 0.0 : ods_length_pt(value.c_str()));
              pos = end + 1;
            }
            if (sides.size() == 4) {
              fr.crop[1] = sides[0] / image_h * 100000.0;
              fr.crop[2] = sides[1] / image_w * 100000.0;
              fr.crop[3] = sides[2] / image_h * 100000.0;
              fr.crop[0] = sides[3] / image_w * 100000.0;
            }
          }
        }
        frames.push_back(fr);
      }
    };
    if (keep && !data_only) read_frames(table.child("table:shapes"), 0, 0);
    std::vector<utf8_string> comment_refs, comment_authors, comment_texts;
    std::vector<bool> comment_rich;
    std::vector<std::string> link_refs, link_targets, link_tips;

    std::vector<ods_in_col> cols;
    int64_t c = 1;
    // repeats that run (nearly) to the end of the sheet fill it; consumers had sheets of 1024 columns and 65536 rows
    const int64_t row_filler = 65536 - 1024, col_filler = 1024 - 64;
    // the default sizes of consumers that do not write them
    double default_col_pt = 64.0, default_row_pt = 12.8;
    // the rows and columns with a page break before them
    std::vector<int64_t> row_breaks, col_breaks;
    // header rows and columns are the print titles
    int64_t title_r1 = 0, title_r2 = 0, title_c1 = 0, title_c2 = 0;
    for (pugi::xml_node col : ods_descendants(table, "table:table-column", "table:table-row")) {
      int64_t rep = std::max<int64_t>(col.attribute("table:number-columns-repeated").as_llong(1), 1);
      if (c > MAX_OOXML_COL_INT) break;
      if (std::strcmp(col.parent().name(), "table:table-header-columns") == 0) {
        if (title_c1 == 0) title_c1 = c;
        title_c2 = std::min<int64_t>(c + rep - 1, MAX_OOXML_COL_INT);
      }
      ods_in_col ic = {c, std::min<int64_t>(c + rep - 1, MAX_OOXML_COL_INT), -1.0, false, col.attribute("table:default-cell-style-name").as_string(),
                       ods_outline_level(col, "table:table-column-group")};
      auto cs = col_styles.find(col.attribute("table:style-name").as_string());
      if (cs != col_styles.end()) ic.width = ods_length_pt(cs->second.child("style:table-column-properties").attribute("style:column-width").as_string(""));
      ic.hidden = std::string(col.attribute("table:visibility").as_string()) == "collapse";
      if (cs != col_styles.end() && std::strcmp(cs->second.child("style:table-column-properties").attribute("fo:break-before").as_string(), "page") == 0) {
        col_breaks.push_back(c);
      }
      // the columns that run to the end of the sheet have the default width
      if (ic.max >= 1024 && ic.width > 0.0) default_col_pt = ic.width;
      ic.collapsed = ods_after_collapsed(col, "table:table-column-group");
      if (ic.collapsed && ic.max > ic.min) {
        ods_in_col first = ic;
        first.max = ic.min;
        cols.push_back(first);
        ic.min += 1;
        ic.collapsed = false;
      }
      cols.push_back(ic);
      c += rep;
    }
    auto col_default = [&](int64_t col) {
      for (const ods_in_col& ic : cols) {
        if (col >= ic.min && col <= ic.max) return ic.style;
      }
      return std::string();
    };

    std::vector<ods_in_cell> cells;
    std::vector<ods_in_run> runs;
    std::vector<ods_in_row> row_attrs, optimal_rows;
    // the cells of cell styles with style maps, conditional formats without the calc extension
    bool map_formats = !table.child("calcext:conditional-formats");
    std::map<std::string, std::vector<ods_range>> map_cells;
    std::map<std::string, std::string> map_styles;
    std::map<std::pair<std::string, std::pair<int64_t, int64_t>>, size_t> map_last;
    std::vector<ods_merge> merges;
    int64_t max_row = 0, max_col = 0;

    int64_t r = 1;
    for (pugi::xml_node row : keep ? ods_descendants(table, "table:table-row") : std::vector<pugi::xml_node>()) {
      if (r > MAX_OOXML_ROW_INT) break;
      int64_t rows = std::min<int64_t>(std::max<int64_t>(row.attribute("table:number-rows-repeated").as_llong(1), 1), MAX_OOXML_ROW_INT - r + 1);
      std::string row_default = row.attribute("table:default-cell-style-name").as_string();
      if (std::strcmp(row.parent().name(), "table:table-header-rows") == 0) {
        if (title_r1 == 0) title_r1 = r;
        title_r2 = r + rows - 1;
      }

      ods_in_row ir = {r, rows, -1.0, std::string(row.attribute("table:visibility").as_string()) == "collapse", 0,
                       ods_outline_level(row, "table:table-row-group")};
      auto rs = row_styles.find(row.attribute("table:style-name").as_string());
      if (rs != row_styles.end()) {
        pugi::xml_node props = rs->second.child("style:table-row-properties");
        if (std::strcmp(props.attribute("fo:break-before").as_string(), "page") == 0 && r > 1) row_breaks.push_back(r);
        if (!props.attribute("style:use-optimal-row-height").as_bool(false)) {
          ir.height = ods_length_pt(props.attribute("style:row-height").as_string(""));
        } else {
          ir.optimal = ods_length_pt(props.attribute("style:row-height").as_string(""));
        }
      }
      if (!row_default.empty()) ir.xf = st.xf(row_default);
      bool row_formatted = !row_default.empty();
      ir.collapsed = ods_after_collapsed(row, "table:table-row-group");
      // the rows that run to the end of the sheet have the default height
      if (r + rows - 1 >= 65536 && (ir.height > 0.0 || ir.optimal > 0.0)) default_row_pt = ir.height > 0.0 ? ir.height : ir.optimal;

      c = 1;
      for (pugi::xml_node cell : row.children()) {
        checkInterrupt(iter);
        if (c > MAX_OOXML_COL_INT) break;
        std::string cname = cell.name();
        if (cname != "table:table-cell" && cname != "table:covered-table-cell") continue;
        int64_t rep = std::min<int64_t>(std::max<int64_t>(cell.attribute("table:number-columns-repeated").as_llong(1), 1), MAX_OOXML_COL_INT - c + 1);

        pugi::xml_attribute style_attr = cell.attribute("table:style-name");
        if (map_formats) {
          std::string map_style = style_attr ? std::string(style_attr.as_string()) : col_default(c);
          auto ms = st.cell.find(map_style);
          if (ms != st.cell.end() && ms->second.child("style:map")) {
            // cell styles with the same conditions share the conditional format
            std::string signature;
            for (pugi::xml_node map : ms->second.children("style:map")) signature += ods_xml_string(map);
            map_styles[signature] = map_style;
            std::vector<ods_range>& ranges = map_cells[signature];
            int64_t r2 = std::min<int64_t>(r + rows - 1, MAX_OOXML_ROW_INT), c2 = std::min<int64_t>(c + rep - 1, MAX_OOXML_COL_INT);
            auto last = map_last.find(std::make_pair(signature, std::make_pair(c, c2)));
            if (last != map_last.end() && ranges[last->second].r2 + 1 == r) {
              ranges[last->second].r2 = r2;
            } else {
              map_last[std::make_pair(signature, std::make_pair(c, c2))] = ranges.size();
              ranges.push_back(ods_range{r, c, r2, c2});
            }
          }
        }
        std::string vt = cell.attribute("office:value-type").as_string();
        pugi::xml_attribute formula = cell.attribute("table:formula");
        bool has_value = !vt.empty() || formula || cell.child("text:p");

        int64_t col_span = cell.attribute("table:number-columns-spanned").as_llong(1);
        int64_t row_span = cell.attribute("table:number-rows-spanned").as_llong(1);
        if (col_span > 1 || row_span > 1) {
          ods_merge m = {static_cast<int32_t>(r), static_cast<int32_t>(c),
                         static_cast<int32_t>(std::min<int64_t>(r + row_span - 1, MAX_OOXML_ROW_INT)),
                         static_cast<int32_t>(std::min<int64_t>(c + col_span - 1, MAX_OOXML_COL_INT))};
          merges.push_back(m);
          max_row = std::max<int64_t>(max_row, m.r2);
          max_col = std::max<int64_t>(max_col, m.c2);
        }

        // comments, images, charts and hyperlinks of the cell
        for (pugi::xml_node child : cell.children()) {
          std::string child_name = child.name();
          if (!data_only && child_name == "office:annotation") {
            comment_refs.push_back(int_to_col(c) + std::to_string(r));
            comment_authors.push_back(std::string(child.child_value("dc:creator")));
            std::string rich = ods_cell_rich(child, st, "");
            comment_texts.push_back(rich.empty() ? ods_cell_string(child) : rich);
            comment_rich.push_back(!rich.empty());
          }
        }
        if (!data_only) read_frames(cell, r, c);
        // the cells of a validation, neighbors in a row are joined
        if (pugi::xml_attribute cvn = cell.attribute("table:content-validation-name")) {
          std::string vname = cvn.as_string();
          if (!valid_cells.empty() && valid_cells.back().name == vname && valid_cells.back().r1 == r &&
              valid_cells.back().r2 == r + rows - 1 && valid_cells.back().c2 + 1 == c) {
            valid_cells.back().c2 = c + rep - 1;
          } else {
            ods_in_valid v = {vname, r, c, r + rows - 1, c + rep - 1};
            valid_cells.push_back(v);
          }
        }
        pugi::xml_node link;
        for (pugi::xml_node para : cell.children("text:p")) {
          link = ods_descendant(para, "text:a");
          if (link) break;
        }
        if (link) {
          link_refs.push_back(int_to_col(c) + std::to_string(r));
          link_targets.push_back(link.attribute("xlink:href").as_string());
          link_tips.push_back(link.attribute("office:title").as_string());
        }

        if (!has_value) {
          if (style_attr && std::string(style_attr.as_string()) != col_default(c)) {
            ods_in_run run = {r, rows, c, rep, st.xf(style_attr.as_string())};
            // a formatted row: its formatted cells run to the end of the sheet
            if (c + rep - 1 >= col_filler && (c == 1 || ir.xf == 0 || ir.xf == run.xf)) {
              ir.xf = run.xf;
              row_formatted = true;
            } else {
              runs.push_back(run);
            }
          }
          c += rep;
          continue;
        }

        ods_in_cell ic;
        ic.row = 0;
        ic.col = 0;
        ic.xf = 0;
        if (formula) {
          std::string fs = formula.as_string();
          size_t colon = fs.find(':');
          std::string prefix = colon == std::string::npos || colon > fs.find('=') ? "" : fs.substr(0, colon);
          if (prefix == "msoxl") ic.f = fs.substr(colon + 2);
          else if (prefix == "of" || prefix == "oooc" || prefix.empty()) ic.f = ods_formula_to_xlsx(prefix.empty() ? fs : fs.substr(colon + 1), renames);
          else ++dropped;
          ic.matrix_rows = cell.attribute("table:number-matrix-rows-spanned").as_llong(0);
          ic.matrix_cols = cell.attribute("table:number-matrix-columns-spanned").as_llong(0);
        }
        if (std::strcmp(cell.attribute("calcext:value-type").as_string(), "error") == 0 || vt == "error" || ods_formula_error(cell, formula)) {
          // the error is the text of the cell, other errors are value errors; some producers write errors as value type
          std::string text = *cell.attribute("office:string-value").as_string() ? cell.attribute("office:string-value").as_string() : ods_cell_string(cell);
          ic.t = "e";
          ic.v = "#VALUE!";
          for (const auto& error : ods_errors) {
            if (text == error[0]) ic.v = error[1];
          }
        } else if (vt == "float" || vt == "percentage" || vt == "currency") {
          ic.v = cell.attribute("office:value").as_string();
        } else if (vt == "date") {
          double serial = 0.0;
          if (ods_parse_date(cell.attribute("office:date-value").as_string(), date1904, serial)) ic.v = ods_num(serial);
        } else if (vt == "time") {
          double serial = 0.0;
          if (ods_parse_time(cell.attribute("office:time-value").as_string(), serial)) ic.v = ods_num(serial);
        } else if (vt == "boolean") {
          ic.t = "b";
          ic.v = cell.attribute("office:boolean-value").as_bool() ? "1" : "0";
        } else {
          pugi::xml_attribute sv = cell.attribute("office:string-value");
          ic.text = sv ? std::string(sv.as_string()) : ods_cell_string(cell);
          if (ic.f.empty()) {
            ic.t = "inlineStr";
            if (!sv) ic.rich = ods_cell_rich(cell, st, style_attr ? std::string(style_attr.as_string()) : col_default(c));
          } else {
            ic.t = "str";
            ic.v = ic.text;
          }
        }

        // dates, times and percentages without a number format of their style have a built-in one:
        // dates with or without time, times as durations
        int32_t builtin = 0;
        if (vt == "date") builtin = ic.v.find('.') == std::string::npos ? 14 : 22;
        else if (vt == "time") builtin = 46;
        else if (vt == "percentage") builtin = 10;
        for (int64_t i = 0; i < rows; ++i) {
          for (int64_t k = 0; k < rep; ++k) {
            std::string sname = style_attr ? std::string(style_attr.as_string()) : col_default(c + k);
            if (sname.empty()) sname = row_default;
            ic.row = static_cast<int32_t>(r + i);
            ic.col = static_cast<int32_t>(c + k);
            ic.xf = sname.empty() ? 0 : st.xf(sname);
            if (builtin > 0 && !ic.v.empty()) ic.xf = st.xf_typed(ic.xf, builtin);
            cells.push_back(ic);
          }
        }
        max_row = std::max<int64_t>(max_row, r + rows - 1);
        max_col = std::max<int64_t>(max_col, c + rep - 1);
        c += rep;
      }
      // empty cells with the style of a formatted row are part of the row
      if (row_formatted) {
        runs.erase(std::remove_if(runs.begin(), runs.end(), [&](const ods_in_run& run) { return run.row == r && run.xf == ir.xf; }), runs.end());
      }
      if (ir.height > 0.0 || ir.hidden || ir.xf > 0 || ir.level > 0 || ir.collapsed) row_attrs.push_back(ir);
      else if (ir.optimal > 0.0) optimal_rows.push_back(ir);
      r += rows;
    }

    // Formatted empty cells and rows extend the used range. Repeats that run
    // (nearly) to the end of the sheet or over more than 1024 rows fill the
    // rest of it and are kept inside the used range only.
    for (const ods_in_run& run : runs) {
      if (run.row + run.rows - 1 >= row_filler || run.col + run.cols - 1 >= col_filler || run.rows > 1024) continue;
      max_row = std::max<int64_t>(max_row, run.row + run.rows - 1);
      max_col = std::max<int64_t>(max_col, run.col + run.cols - 1);
    }
    for (const ods_in_row& ir : row_attrs) {
      if (ir.row + ir.rows - 1 < row_filler) max_row = std::max<int64_t>(max_row, ir.row + ir.rows - 1);
    }
    for (const ods_in_run& run : runs) {
      for (int64_t i = run.row; i < run.row + run.rows && i <= max_row; ++i) {
        for (int64_t k = run.col; k < run.col + run.cols && k <= max_col; ++k) {
          ods_in_cell ic;
          ic.row = static_cast<int32_t>(i);
          ic.col = static_cast<int32_t>(k);
          ic.xf = run.xf;
          cells.push_back(ic);
        }
      }
    }
    std::stable_sort(cells.begin(), cells.end(), [](const ods_in_cell& a, const ods_in_cell& b) {
      return a.row != b.row ? a.row < b.row : a.col < b.col;
    });

    std::map<int32_t, const ods_in_row*> row_map;
    for (const ods_in_row& ir : row_attrs) {
      for (int64_t i = ir.row; i < ir.row + ir.rows && i <= max_row; ++i) row_map[static_cast<int32_t>(i)] = &ir;
    }

    // the cells covered by images and charts; frames anchored to the sheet are
    // placed by their position
    auto col_width = [&](int64_t col) {
      for (const ods_in_col& ic : cols) {
        if (col >= ic.min && col <= ic.max) return ic.hidden ? 0.0 : ic.width > 0.0 ? ic.width : default_col_pt;
      }
      return default_col_pt;
    };
    auto row_height = [&](int64_t row) {
      for (const ods_in_row& ir : row_attrs) {
        if (row >= ir.row && row < ir.row + ir.rows) return ir.hidden ? 0.0 : ir.height > 0.0 ? ir.height : ir.optimal > 0.0 ? ir.optimal : default_row_pt;
      }
      for (const ods_in_row& ir : optimal_rows) {
        if (row >= ir.row && row < ir.row + ir.rows) return ir.optimal;
      }
      return default_row_pt;
    };
    std::vector<utf8_string> image_dims, image_files, chart_dims, chart_xml, shape_dims, shape_types, shape_texts, shape_fills, shape_lines, shape_aligns,
        shape_geometries, shape_properties;
    std::vector<bool> shape_rich, control_checked;
    std::vector<double> shape_rotations, image_rotations, gradient_angles, image_crop[4];
    std::vector<double> image_anchors, chart_anchors, shape_anchors;
    std::vector<utf8_string> gradient_starts, gradient_ends;
    std::vector<utf8_string> control_dims, control_types, control_texts, control_links, control_ranges;
    for (const ods_in_frame& fr : frames) {
      int64_t col = fr.col > 0 ? fr.col : 1, row = fr.row > 0 ? fr.row : 1;
      double x = fr.x, y = fr.y;
      if (fr.col == 0) {
        while (col < MAX_OOXML_COL_INT && x + 0.01 >= col_width(col)) x -= col_width(col++);
        while (row < MAX_OOXML_ROW_INT && y + 0.01 >= row_height(row)) y -= row_height(row++);
      }
      // a frame that ends at the edge of the end cell does not cover it
      std::string to = fr.end_cell;
      if (!to.empty() && (fr.end_x < 1.0 || fr.end_y < 1.0)) {
        int64_t ec = std::max<int64_t>(cell_to_colint(to) - (fr.end_x < 1.0 ? 1 : 0), col);
        int64_t er = std::max<int64_t>(cell_to_rowint(to) - (fr.end_y < 1.0 ? 1 : 0), row);
        to = int_to_col(ec) + std::to_string(er);
      }
      if (to.empty()) {
        double ex = x + fr.width, ey = y + fr.height;
        int64_t ec = col, er = row;
        while (ec < MAX_OOXML_COL_INT && ex > col_width(ec) + 0.01) ex -= col_width(ec++);
        while (er < MAX_OOXML_ROW_INT && ey > row_height(er) + 0.01) ey -= row_height(er++);
        to = int_to_col(ec) + std::to_string(er);
      }
      std::string dims = int_to_col(col) + std::to_string(row) + ":" + to;
      // the anchor of the frame: the start and end cell (0-based) and the offsets in them in EMU, rounded
      // to hundredths of a point; the end is the end cell of the frame or follows from its size
      double offsets[8] = {static_cast<double>(col - 1), x, static_cast<double>(row - 1), y, 0.0, 0.0, 0.0, 0.0};
      if (!fr.end_cell.empty()) {
        offsets[4] = static_cast<double>(cell_to_colint(fr.end_cell) - 1);
        offsets[5] = fr.end_x;
        offsets[6] = static_cast<double>(cell_to_rowint(fr.end_cell) - 1);
        offsets[7] = fr.end_y;
      } else {
        double ex = x + fr.width, ey = y + fr.height;
        int64_t ec = col, er = row;
        while (ec < MAX_OOXML_COL_INT && ex > col_width(ec) + 0.01) ex -= col_width(ec++);
        while (er < MAX_OOXML_ROW_INT && ey > row_height(er) + 0.01) ey -= row_height(er++);
        // an end at the edge of a cell is the start of the next one
        if (ec < MAX_OOXML_COL_INT && ex + 0.01 >= col_width(ec)) {
          ex = 0.0;
          ++ec;
        }
        if (er < MAX_OOXML_ROW_INT && ey + 0.01 >= row_height(er)) {
          ey = 0.0;
          ++er;
        }
        offsets[4] = static_cast<double>(ec - 1);
        offsets[5] = ex;
        offsets[6] = static_cast<double>(er - 1);
        offsets[7] = ey;
      }
      for (size_t k = 1; k < 8; k += 2) offsets[k] = std::round(std::max(offsets[k], 0.0) * 100.0) / 100.0 * ods_emu_per_pt;
      if (!fr.control.empty()) {
        // check boxes, option buttons and drop down lists are kept
        pugi::xml_node control;
        for (pugi::xml_node fc : ods_descendants(table.child("office:forms"), "")) {
          if (fc.type() != pugi::node_element) continue;
          if (fr.control == fc.attribute("xml:id").as_string() || fr.control == fc.attribute("form:id").as_string()) control = fc;
        }
        std::string control_name = control.name(), type;
        if (control_name == "form:checkbox") type = "Checkbox";
        else if (control_name == "form:radio") type = "Radio";
        else if (control_name == "form:listbox" || control_name == "form:combobox") type = "Drop";
        if (type.empty()) continue;
        control_dims.push_back(dims);
        control_types.push_back(type);
        control_texts.push_back(std::string(control.attribute("form:label").as_string()));
        control_links.push_back(ods_xlsx_chart_range(control.attribute("form:linked-cell").as_string(), renames));
        control_ranges.push_back(ods_xlsx_chart_range(control.attribute("form:source-cell-range").as_string(), renames));
        control_checked.push_back(std::string(control.attribute("form:current-state").as_string()) == "checked" ||
                                  control.attribute("form:current-selected").as_bool());
      } else if (!fr.shape.empty()) {
        shape_dims.push_back(dims);
        shape_anchors.insert(shape_anchors.end(), offsets, offsets + 8);
        shape_types.push_back(fr.shape);
        shape_geometries.push_back(fr.geometry);
        shape_properties.push_back(fr.properties);
        shape_texts.push_back(fr.text);
        shape_fills.push_back(fr.fill);
        shape_lines.push_back(fr.line);
        shape_rich.push_back(fr.rich);
        shape_aligns.push_back(fr.align.empty() ? std::string("left") : fr.align);
        shape_rotations.push_back(fr.rotation);
        gradient_starts.push_back(fr.gradient_start);
        gradient_ends.push_back(fr.gradient_end);
        gradient_angles.push_back(fr.gradient_angle);
      } else if (!fr.image.empty()) {
        image_dims.push_back(dims);
        image_anchors.insert(image_anchors.end(), offsets, offsets + 8);
        image_files.push_back(fr.image);
        for (size_t k = 0; k < 4; ++k) image_crop[k].push_back(std::round(fr.crop[k]));
        image_rotations.push_back(fr.rotation);
      } else if (!fr.object.empty()) {
        std::string xml = ods_chart_xml(base + fr.object + "/content.xml", renames);
        if (xml.empty()) continue;
        chart_dims.push_back(dims);
        chart_anchors.insert(chart_anchors.end(), offsets, offsets + 8);
        chart_xml.push_back(xml);
      }
    }
    sheet_extras.push_back(Rcpp::List::create(
      Rcpp::Named("comment_ref") = comment_refs, Rcpp::Named("comment_author") = comment_authors,
      Rcpp::Named("comment_text") = comment_texts, Rcpp::Named("comment_rich") = comment_rich,
      Rcpp::Named("image_dims") = image_dims,
      Rcpp::Named("image_file") = image_files, Rcpp::Named("chart_dims") = chart_dims,
      Rcpp::Named("chart_xml") = chart_xml, Rcpp::Named("shape_dims") = shape_dims,
      Rcpp::Named("shape_type") = shape_types, Rcpp::Named("shape_text") = shape_texts,
      Rcpp::Named("shape_fill") = shape_fills, Rcpp::Named("shape_line") = shape_lines,
      Rcpp::Named("shape_rich") = shape_rich, Rcpp::Named("shape_align") = shape_aligns,
      Rcpp::Named("shape_extra") = Rcpp::List::create(
        Rcpp::Named("rotation") = shape_rotations, Rcpp::Named("geometry") = shape_geometries,
        Rcpp::Named("gradient_start") = gradient_starts, Rcpp::Named("gradient_end") = gradient_ends, Rcpp::Named("gradient_angle") = gradient_angles,
        Rcpp::Named("properties") = shape_properties
      ),
      Rcpp::Named("anchors") = Rcpp::List::create(
        Rcpp::Named("image") = image_anchors, Rcpp::Named("chart") = chart_anchors, Rcpp::Named("shape") = shape_anchors
      ),
      Rcpp::Named("image_extra") = Rcpp::List::create(
        Rcpp::Named("l") = image_crop[0], Rcpp::Named("t") = image_crop[1], Rcpp::Named("r") = image_crop[2],
        Rcpp::Named("b") = image_crop[3], Rcpp::Named("rotation") = image_rotations
      ),
      Rcpp::Named("controls") = Rcpp::List::create(
        Rcpp::Named("dims") = control_dims, Rcpp::Named("type") = control_types, Rcpp::Named("text") = control_texts,
        Rcpp::Named("link") = control_links, Rcpp::Named("range") = control_ranges, Rcpp::Named("checked") = control_checked
      ),
      Rcpp::Named("pivots") = pilots[names.back()]
    ));

    // page setup: the master page of the table and its page layout
    std::string master_name = ts != table_styles.end() ? ts->second.attribute("style:master-page-name").as_string("Default") : "Default";
    pugi::xml_node master, layout;
    for (pugi::xml_node container : master_containers) {
      for (pugi::xml_node mp : container.children("style:master-page")) {
        if (!master && master_name == mp.attribute("style:name").as_string()) master = mp;
      }
    }
    for (pugi::xml_node container : page_containers) {
      for (pugi::xml_node pl : container.children("style:page-layout")) {
        if (!layout && std::string(master.attribute("style:page-layout-name").as_string()) == pl.attribute("style:name").as_string()) layout = pl;
      }
    }

    pugi::xml_document ws, ext_doc;
    ods_declaration(ws);
    pugi::xml_node worksheet = ws.append_child("worksheet");
    // the extensions of the worksheet: conditional formats and sparkline groups of later versions
    pugi::xml_node ext_list = ext_doc.append_child("extLst");
    pugi::xml_node ext_formats;
    worksheet.append_attribute("xmlns") = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    worksheet.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    pugi::xml_node page_props = layout.child("style:page-layout-properties");
    bool fit = page_props.attribute("style:scale-to-X") || page_props.attribute("style:scale-to-Y");
    if (fit) worksheet.append_child("sheetPr").append_child("pageSetUpPr").append_attribute("fitToPage") = 1;

    // the right to left direction and the color of the tab
    pugi::xml_node table_props = ts != table_styles.end() ? ts->second.child("style:table-properties") : pugi::xml_node();
    bool right_to_left = std::strcmp(table_props.attribute("style:writing-mode").as_string(), "rl-tb") == 0;
    std::string tab_color = table_props.attribute("table:tab-color").as_string(table_props.attribute("tableooo:tab-color").as_string());
    if (tab_color.size() == 7 && tab_color[0] == '#') {
      if (!worksheet.child("sheetPr")) worksheet.prepend_child("sheetPr");
      worksheet.child("sheetPr").prepend_child("tabColor").append_attribute("rgb") = ("FF" + ods_upper(tab_color.substr(1))).c_str();
    }

    pugi::xml_node sheet_view = worksheet.append_child("sheetViews").append_child("sheetView");
    auto view = views.find(names.back());
    bool view_page_break = page_break && names.back() == active_table;
    if (view != views.end() ? !view->second.zeros : !show_zeros) sheet_view.append_attribute("showZeros") = 0;
    if (view != views.end() ? !view->second.headers : !show_headers) sheet_view.append_attribute("showRowColHeaders") = 0;
    if (!show_outline) sheet_view.append_attribute("showOutlineSymbols") = 0;
    if (right_to_left) sheet_view.append_attribute("rightToLeft") = 1;
    // the grid color of the document is the nearest color of the palette, the default color is light gray
    if (grid_color >= 0 && grid_color != 12632256) {
      int64_t nearest = 0, best = -1;
      for (int64_t k = 0; k < 56; ++k) {
        int64_t dr = static_cast<int64_t>((XlsDefaultPalette[k] >> 16) & 0xFF) - ((grid_color >> 16) & 0xFF);
        int64_t dg = static_cast<int64_t>((XlsDefaultPalette[k] >> 8) & 0xFF) - ((grid_color >> 8) & 0xFF);
        int64_t db = static_cast<int64_t>(XlsDefaultPalette[k] & 0xFF) - (grid_color & 0xFF);
        int64_t distance = dr * dr + dg * dg + db * db;
        if (best < 0 || distance < best) {
          best = distance;
          nearest = k + 8;
        }
      }
      sheet_view.append_attribute("defaultGridColor") = 0;
      sheet_view.append_attribute("colorId") = nearest;
    }
    if (view != views.end()) {
      if (view->second.grid == 0) sheet_view.append_attribute("showGridLines") = 0;
      if (view_page_break) sheet_view.append_attribute("view") = "pageBreakPreview";
      int64_t zoom = view_page_break ? view->second.page_zoom : view->second.zoom;
      if (zoom > 0) sheet_view.append_attribute("zoomScale") = zoom;
      if (view_page_break && view->second.zoom > 0) sheet_view.append_attribute("zoomScaleNormal") = view->second.zoom;
      if (view->second.page_zoom > 0 && view->second.page_zoom != 60) sheet_view.append_attribute("zoomScaleSheetLayoutView") = view->second.page_zoom;
    }
    if (names.back() == active_table) sheet_view.append_attribute("tabSelected") = 1;
    int64_t left = 0, top = 0;
    if (view != views.end()) {
      left = std::min<int64_t>(view->second.left, MAX_OOXML_COL_INT - 1);
      top = std::min<int64_t>(view->second.top, MAX_OOXML_ROW_INT - 1);
    }
    bool frozen = view != views.end() && (view->second.rows > 0 || view->second.cols > 0);
    if (!frozen && (left > 0 || top > 0)) sheet_view.append_attribute("topLeftCell") = (int_to_col(left + 1) + std::to_string(top + 1)).c_str();
    sheet_view.append_attribute("workbookViewId") = 0;
    const char* active_pane = nullptr;
    if (frozen) {
      // a split that is not frozen is in pixels, xlsx has it in twentieths of a point
      bool split = view->second.split;
      int64_t frows = std::min<int64_t>(view->second.rows, MAX_OOXML_ROW_INT - 1), fcols = std::min<int64_t>(view->second.cols, MAX_OOXML_COL_INT - 1);
      pugi::xml_node pane = sheet_view.append_child("pane");
      if (fcols > 0) pane.append_attribute("xSplit") = split ? view->second.cols * 15 : fcols;
      if (frows > 0) pane.append_attribute("ySplit") = split ? view->second.rows * 15 : frows;
      int64_t pane_left = fcols > 0 ? (split ? left : std::max(left, fcols)) : 0, pane_top = frows > 0 ? (split ? top : std::max(top, frows)) : 0;
      pane.append_attribute("topLeftCell") = (int_to_col(pane_left + 1) + std::to_string(pane_top + 1)).c_str();
      active_pane = fcols > 0 && frows > 0 ? "bottomRight" : frows > 0 ? "bottomLeft" : "topRight";
      pane.append_attribute("activePane") = active_pane;
      pane.append_attribute("state") = split ? "split" : "frozen";
    }
    if (view != views.end() && (active_pane || view->second.cursor_col > 0 || view->second.cursor_row > 0)) {
      pugi::xml_node selection = sheet_view.append_child("selection");
      if (active_pane) selection.append_attribute("pane") = active_pane;
      if (view->second.cursor_col > 0 || view->second.cursor_row > 0) {
        std::string cursor = int_to_col(std::min<int64_t>(view->second.cursor_col, MAX_OOXML_COL_INT - 1) + 1) +
                             std::to_string(std::min<int64_t>(view->second.cursor_row, MAX_OOXML_ROW_INT - 1) + 1);
        selection.append_attribute("activeCell") = cursor.c_str();
        selection.append_attribute("sqref") = cursor.c_str();
      }
    }

    int64_t outline_rows = 0, outline_cols = 0;
    for (const ods_in_row& ir : row_attrs) outline_rows = std::max(outline_rows, ir.level);
    for (const ods_in_col& ic : cols) outline_cols = std::max(outline_cols, ic.level);
    pugi::xml_node sheet_format = worksheet.append_child("sheetFormatPr");
    sheet_format.append_attribute("defaultColWidth") = ods_num(std::round(default_col_pt / unit_pt * 1000.0) / 1000.0).c_str();
    sheet_format.append_attribute("defaultRowHeight") = ods_num(std::round(default_row_pt * 100.0) / 100.0).c_str();
    if (outline_rows > 0) sheet_format.append_attribute("outlineLevelRow") = outline_rows;
    if (outline_cols > 0) sheet_format.append_attribute("outlineLevelCol") = outline_cols;

    // columns of the default width are written only with other attributes
    pugi::xml_node xcols = worksheet.append_child("cols");
    for (const ods_in_col& ic : cols) {
      int32_t xf = ic.style.empty() ? 0 : st.xf(ic.style);
      bool custom = ic.width > 0.0 && std::fabs(ic.width - default_col_pt) > 0.25;
      if (!custom && !ic.hidden && xf == 0 && ic.level == 0 && !ic.collapsed) continue;
      pugi::xml_node col = xcols.append_child("col");
      col.append_attribute("min") = ic.min;
      col.append_attribute("max") = ic.max;
      double width = (custom ? ic.width : default_col_pt) / unit_pt;
      col.append_attribute("width") = ods_num(std::round(width * 1000.0) / 1000.0).c_str();
      if (custom) col.append_attribute("customWidth") = 1;
      if (ic.level > 0) col.append_attribute("outlineLevel") = ic.level;
      if (ic.hidden) col.append_attribute("hidden") = 1;
      if (ic.collapsed) col.append_attribute("collapsed") = 1;
      if (xf > 0) col.append_attribute("style") = xf;
    }
    if (!xcols.first_child()) worksheet.remove_child(xcols);

    pugi::xml_node sheet_data = worksheet.append_child("sheetData");
    std::set<int32_t> row_numbers;
    for (const ods_in_cell& ic : cells) row_numbers.insert(ic.row);
    for (const auto& ir : row_map) row_numbers.insert(ir.first);

    size_t ci = 0;
    for (int32_t rn : row_numbers) {
      pugi::xml_node row = sheet_data.append_child("row");
      row.append_attribute("r") = rn;
      auto attr = row_map.find(rn);
      if (attr != row_map.end()) {
        if (attr->second->xf > 0) {
          row.append_attribute("s") = attr->second->xf;
          row.append_attribute("customFormat") = 1;
        }
        if (attr->second->height > 0.0) {
          row.append_attribute("ht") = ods_num(std::round(attr->second->height * 100.0) / 100.0).c_str();
          row.append_attribute("customHeight") = 1;
        }
        if (attr->second->hidden) row.append_attribute("hidden") = 1;
        if (attr->second->level > 0) row.append_attribute("outlineLevel") = attr->second->level;
        if (attr->second->collapsed && rn == attr->second->row) row.append_attribute("collapsed") = 1;
      }
      int32_t last = 0;
      for (; ci < cells.size() && cells[ci].row == rn; ++ci) {
        const ods_in_cell& ic = cells[ci];
        if (ic.col == last) continue;
        last = ic.col;
        pugi::xml_node cell = row.append_child("c");
        cell.append_attribute("r") = (int_to_col(ic.col) + std::to_string(ic.row)).c_str();
        if (ic.xf > 0) cell.append_attribute("s") = ic.xf;
        if (!ic.t.empty()) cell.append_attribute("t") = ic.t.c_str();
        if (ic.t == "inlineStr") {
          pugi::xml_node is = cell.append_child("is");
          if (!ic.rich.empty()) {
            pugi::xml_document rich_doc;
            rich_doc.load_string(("<is>" + ic.rich + "</is>").c_str(), pugi::parse_default | pugi::parse_ws_pcdata);
            for (pugi::xml_node run : rich_doc.child("is").children()) is.append_copy(run);
            continue;
          }
          pugi::xml_node t = is.append_child("t");
          t.append_attribute("xml:space") = "preserve";
          t.text().set(ic.text.c_str());
          continue;
        }
        if (!ic.f.empty()) {
          pugi::xml_node f = cell.append_child("f");
          if (ic.matrix_rows > 0 || ic.matrix_cols > 0) {
            int64_t to_row = std::min<int64_t>(ic.row + std::max<int64_t>(ic.matrix_rows, 1) - 1, MAX_OOXML_ROW_INT);
            int64_t to_col = std::min<int64_t>(ic.col + std::max<int64_t>(ic.matrix_cols, 1) - 1, MAX_OOXML_COL_INT);
            f.append_attribute("t") = "array";
            f.append_attribute("ref") = (int_to_col(ic.col) + std::to_string(ic.row) + ":" + int_to_col(to_col) + std::to_string(to_row)).c_str();
          }
          f.text().set(ic.f.c_str());
        }
        if (!ic.v.empty()) cell.append_child("v").text().set(ic.v.c_str());
      }
    }

    if (table_protected) worksheet.append_child("sheetProtection").append_attribute("sheet") = 1;

    // print ranges and print titles as defined names of the sheet
    std::string xlsx_sheet = "'";
    for (char ch : renames.count(names.back()) ? renames[names.back()] : names.back()) {
      xlsx_sheet += ch;
      if (ch == '\'') xlsx_sheet += ch;
    }
    xlsx_sheet += "'!";
    std::string print_area, print_targets = table.attribute("table:print-ranges").as_string();
    size_t print_pos = 0;
    while (print_pos < print_targets.size()) {
      size_t end = print_pos;
      bool quoted = false;
      while (end < print_targets.size() && (quoted || print_targets[end] != ' ')) {
        if (print_targets[end] == '\'') quoted = !quoted;
        ++end;
      }
      std::string range_sheet, ref;
      if (ods_split_range(print_targets.substr(print_pos, end - print_pos), range_sheet, ref)) {
        std::string from = ref.substr(0, ref.find(':')), to = ref.substr(ref.find(':') + 1);
        std::string address = "$" + int_to_col(cell_to_colint(from)) + "$" + std::to_string(cell_to_rowint(from));
        if (to != from) address += ":$" + int_to_col(cell_to_colint(to)) + "$" + std::to_string(cell_to_rowint(to));
        print_area += (print_area.empty() ? "" : ",") + xlsx_sheet + address;
      }
      print_pos = end + 1;
    }
    std::string local = std::to_string(names.size() - 1);
    if (!print_area.empty()) defined.push_back(ods_name{"_xlnm.Print_Area", local, print_area});
    std::string titles;
    if (title_c1 > 0) titles = xlsx_sheet + "$" + int_to_col(title_c1) + ":$" + int_to_col(title_c2);
    if (title_r1 > 0) titles += (titles.empty() ? "" : ",") + xlsx_sheet + "$" + std::to_string(title_r1) + ":$" + std::to_string(title_r2);
    if (!titles.empty()) defined.push_back(ods_name{"_xlnm.Print_Titles", local, titles});

    std::vector<ods_db_range> sheet_tables;
    for (const ods_db_range& db : db_ranges[names.back()]) {
      if (db.name.compare(0, 22, "__Anonymous_Sheet_DB__") != 0) {
        sheet_tables.push_back(db);
      } else if (db.filter && !worksheet.child("autoFilter")) {
        worksheet.append_child("autoFilter").append_attribute("ref") = db.ref.c_str();
      }
    }

    if (!merges.empty()) {
      pugi::xml_node merge_cells = worksheet.append_child("mergeCells");
      merge_cells.append_attribute("count") = static_cast<uint32_t>(merges.size());
      for (const ods_merge& m : merges) {
        std::string ref = int_to_col(m.c1) + std::to_string(m.r1) + ":" + int_to_col(m.c2) + std::to_string(m.r2);
        merge_cells.append_child("mergeCell").append_attribute("ref") = ref.c_str();
      }
    }
    // conditional formats; the styles of the conditions are differential formats
    int64_t priority = 0;
    auto dxf_of = [&](const std::string& style_name) {
      int64_t dxf = -1;
      auto cs = st.cell.find(style_name);
      if (cs == st.cell.end()) {
        for (cs = st.cell.begin(); cs != st.cell.end(); ++cs) {
          if (style_name == cs->second.attribute("style:display-name").as_string()) break;
        }
      }
      if (cs != st.cell.end()) {
        std::string xml = ods_dxf(cs->second, st.font_face);
        auto known = dxf_index.find(xml);
        if (known != dxf_index.end()) {
          dxf = known->second;
        } else {
          dxf = static_cast<int64_t>(dxf_list.size());
          dxf_index[xml] = dxf;
          dxf_list.push_back(xml);
        }
      }
      return dxf;
    };
    for (pugi::xml_node format : table.child("calcext:conditional-formats").children("calcext:conditional-format")) {
      std::string targets = format.attribute("calcext:target-range-address").as_string(), sqref;
      size_t pos = 0;
      while (pos < targets.size()) {
        size_t end = pos;
        bool quoted = false;
        while (end < targets.size() && (quoted || targets[end] != ' ')) {
          if (targets[end] == '\'') quoted = !quoted;
          ++end;
        }
        std::string sheet, ref;
        if (ods_split_range(targets.substr(pos, end - pos), sheet, ref)) sqref += (sqref.empty() ? "" : " ") + ref;
        pos = end + 1;
      }
      if (sqref.empty()) continue;
      std::string top_left = sqref.substr(0, sqref.find_first_of(": "));
      pugi::xml_node cf = worksheet.append_child("conditionalFormatting");
      cf.append_attribute("sqref") = sqref.c_str();
      for (pugi::xml_node entry : format.children()) {
        int64_t dxf = dxf_of(entry.attribute(entry.attribute("calcext:style") ? "calcext:style" : "calcext:apply-style-name").as_string());
        pugi::xml_node rule = cf.append_child("cfRule");
        if (!ods_cf_rule(entry, rule, top_left, dxf, renames)) {
          cf.remove_child(rule);
          continue;
        }
        if (dxf < 0) rule.remove_attribute("dxfId");
        rule.append_attribute("priority") = ++priority;

        // data bars with negative or axis colors, a solid fill or lengths and icon sets of later versions are extensions
        std::string kind = entry.name(), icons = entry.attribute("calcext:icon-set-type").as_string();
        bool extended_icons = icons == "3Stars" || icons == "3Triangles" || icons == "5Boxes";
        if (kind != "calcext:data-bar" && !extended_icons) continue;
        std::string number = std::to_string(priority);
        std::string guid = "{00000000-0000-0000-0000-" + std::string(number.size() < 12 ? 12 - number.size() : 0, '0') + number + "}";
        if (!ext_formats) {
          pugi::xml_node ext = ext_list.append_child("ext");
          ext.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
          ext.append_attribute("uri") = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
          ext_formats = ext.append_child("x14:conditionalFormattings");
          ext_formats.append_attribute("xmlns:xm") = "http://schemas.microsoft.com/office/excel/2006/main";
        }
        pugi::xml_node x14_cf = ext_formats.append_child("x14:conditionalFormatting");
        pugi::xml_node x14_rule = x14_cf.append_child("x14:cfRule");
        x14_rule.append_attribute("type") = extended_icons ? "iconSet" : "dataBar";
        if (extended_icons) x14_rule.append_attribute("priority") = priority;
        x14_rule.append_attribute("id") = guid.c_str();
        pugi::xml_node inner = x14_rule.append_child(extended_icons ? "x14:iconSet" : "x14:dataBar");
        if (extended_icons) {
          inner.append_attribute("iconSet") = icons.c_str();
          if (!entry.attribute("calcext:show-value").as_bool(true)) inner.append_attribute("showValue") = 0;
          for (pugi::xml_node cfvo : rule.child("iconSet").children("cfvo")) {
            pugi::xml_node out = inner.append_child("x14:cfvo");
            out.append_attribute("type") = cfvo.attribute("type").as_string();
            out.append_child("xm:f").text().set(cfvo.attribute("val").as_string("0"));
          }
          cf.remove_child(rule);
        } else {
          inner.append_attribute("minLength") = entry.attribute("calcext:min-length").as_int(0);
          inner.append_attribute("maxLength") = entry.attribute("calcext:max-length").as_int(100);
          if (!entry.attribute("calcext:gradient").as_bool(true)) inner.append_attribute("gradient") = 0;
          std::string position = entry.attribute("calcext:axis-position").as_string("automatic");
          if (position != "automatic") inner.append_attribute("axisPosition") = position.c_str();
          for (pugi::xml_node cfvo : rule.child("dataBar").children("cfvo")) {
            std::string type = cfvo.attribute("type").as_string();
            pugi::xml_node out = inner.append_child("x14:cfvo");
            out.append_attribute("type") = type == "min" ? "autoMin" : type == "max" ? "autoMax" : type.c_str();
            if (type != "min" && type != "max") out.append_child("xm:f").text().set(cfvo.attribute("val").as_string("0"));
          }
          std::string negative = ods_rgb(entry.attribute("calcext:negative-color").as_string("#ff0000"));
          std::string axis = ods_rgb(entry.attribute("calcext:axis-color").as_string("#000000"));
          if (!negative.empty()) inner.append_child("x14:negativeFillColor").append_attribute("rgb") = negative.c_str();
          if (!axis.empty()) inner.append_child("x14:axisColor").append_attribute("rgb") = axis.c_str();
          pugi::xml_node rule_ext = rule.append_child("extLst").append_child("ext");
          rule_ext.append_attribute("uri") = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";
          rule_ext.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
          rule_ext.append_child("x14:id").text().set(guid.c_str());
        }
        x14_cf.append_child("xm:sqref").text().set(sqref.c_str());
      }
      if (!cf.first_child()) worksheet.remove_child(cf);
    }

    // the style maps of cell styles: comparisons of the cell content and formulas
    static const char* const map_ops[][2] = {
      {"!=", "notEqual"}, {">=", "greaterThanOrEqual"}, {"<=", "lessThanOrEqual"}, {"=", "equal"}, {">", "greaterThan"}, {"<", "lessThan"}
    };
    for (const auto& mc : map_cells) {
      std::string sqref;
      for (const ods_range& range : mc.second) {
        std::string ref = int_to_col(range.c1) + std::to_string(range.r1);
        if (range.r2 != range.r1 || range.c2 != range.c1) ref += ":" + int_to_col(range.c2) + std::to_string(range.r2);
        sqref += (sqref.empty() ? "" : " ") + ref;
      }
      pugi::xml_node cf = worksheet.append_child("conditionalFormatting");
      cf.append_attribute("sqref") = sqref.c_str();
      for (pugi::xml_node map : st.cell[map_styles[mc.first]].children("style:map")) {
        std::string cond = map.attribute("style:condition").as_string();
        if (cond.compare(0, 3, "of:") == 0) cond = cond.substr(3);
        if (cond.compare(0, 5, "oooc:") == 0) cond = cond.substr(5);
        pugi::xml_node rule = cf.append_child("cfRule");
        size_t end = 0;
        const std::string cell_content = "cell-content()", between = "cell-content-is-between(", not_between = "cell-content-is-not-between(";
        const std::string formula = "is-true-formula(";
        if (cond.compare(0, between.size(), between) == 0 || cond.compare(0, not_between.size(), not_between) == 0) {
          bool is_between = cond.compare(0, between.size(), between) == 0;
          std::vector<std::string> args = ods_arguments(cond, is_between ? between.size() : not_between.size(), end);
          if (args.size() != 2) {
            cf.remove_child(rule);
            continue;
          }
          rule.append_attribute("type") = "cellIs";
          rule.append_attribute("operator") = is_between ? "between" : "notBetween";
          rule.append_child("formula").text().set(ods_formula_to_xlsx(args[0], renames).c_str());
          rule.append_child("formula").text().set(ods_formula_to_xlsx(args[1], renames).c_str());
        } else if (cond.compare(0, formula.size(), formula) == 0) {
          std::vector<std::string> args = ods_arguments(cond, formula.size(), end);
          if (args.empty()) {
            cf.remove_child(rule);
            continue;
          }
          rule.append_attribute("type") = "expression";
          rule.append_child("formula").text().set(ods_formula_to_xlsx(args[0], renames).c_str());
        } else if (cond.compare(0, cell_content.size(), cell_content) == 0) {
          std::string rest = cond.substr(cell_content.size());
          const char* op = nullptr;
          size_t len = 0;
          for (const auto& mo : map_ops) {
            if (rest.compare(0, std::strlen(mo[0]), mo[0]) == 0) {
              op = mo[1];
              len = std::strlen(mo[0]);
              break;
            }
          }
          if (!op) {
            cf.remove_child(rule);
            continue;
          }
          rule.append_attribute("type") = "cellIs";
          rule.append_attribute("operator") = op;
          rule.append_child("formula").text().set(ods_formula_to_xlsx(rest.substr(len), renames).c_str());
        } else {
          cf.remove_child(rule);
          continue;
        }
        int64_t dxf = dxf_of(map.attribute("style:apply-style-name").as_string());
        if (dxf >= 0) rule.insert_attribute_after("dxfId", rule.attribute("type")) = dxf;
        rule.append_attribute("priority") = ++priority;
      }
      if (!cf.first_child()) worksheet.remove_child(cf);
    }

    // data validations, cells of the same validation in consecutive rows are joined
    std::stable_sort(valid_cells.begin(), valid_cells.end(), [](const ods_in_valid& a, const ods_in_valid& b) {
      if (a.name != b.name) return a.name < b.name;
      if (a.c1 != b.c1) return a.c1 < b.c1;
      if (a.c2 != b.c2) return a.c2 < b.c2;
      return a.r1 < b.r1;
    });
    std::vector<ods_in_valid> joined;
    for (const ods_in_valid& v : valid_cells) {
      if (!joined.empty() && joined.back().name == v.name && joined.back().c1 == v.c1 && joined.back().c2 == v.c2 && joined.back().r2 + 1 == v.r1) {
        joined.back().r2 = v.r2;
      } else {
        joined.push_back(v);
      }
    }
    pugi::xml_node data_validations;
    for (size_t k = 0; k < joined.size();) {
      std::string vname = joined[k].name, sqref;
      for (; k < joined.size() && joined[k].name == vname; ++k) {
        std::string ref = int_to_col(joined[k].c1) + std::to_string(joined[k].r1);
        if (joined[k].r1 != joined[k].r2 || joined[k].c1 != joined[k].c2) ref += ":" + int_to_col(joined[k].c2) + std::to_string(joined[k].r2);
        sqref += (sqref.empty() ? "" : " ") + ref;
      }
      auto cv = validations.find(vname);
      if (cv == validations.end()) continue;
      pugi::xml_document dv_doc;
      pugi::xml_node dv = dv_doc.append_child("dataValidation");
      if (!ods_validation_xlsx(cv->second.attribute("table:condition").as_string(), dv, renames)) continue;
      dv.append_attribute("allowBlank") = cv->second.attribute("table:allow-empty-cell").as_bool(true) ? 1 : 0;
      if (std::string(cv->second.attribute("table:display-list").as_string()) == "none") dv.append_attribute("showDropDown") = 1;
      pugi::xml_node help = cv->second.child("table:help-message"), error = cv->second.child("table:error-message");
      dv.append_attribute("showInputMessage") = help.attribute("table:display").as_bool() ? 1 : 0;
      dv.append_attribute("showErrorMessage") = error.attribute("table:display").as_bool() ? 1 : 0;
      std::string message_type = error.attribute("table:message-type").as_string("stop");
      if (message_type == "warning" || message_type == "information") dv.append_attribute("errorStyle") = message_type.c_str();
      if (*error.attribute("table:title").as_string()) dv.append_attribute("errorTitle") = error.attribute("table:title").as_string();
      std::string error_text = ods_cell_string(error), help_text = ods_cell_string(help);
      if (!error_text.empty()) dv.append_attribute("error") = error_text.c_str();
      if (*help.attribute("table:title").as_string()) dv.append_attribute("promptTitle") = help.attribute("table:title").as_string();
      if (!help_text.empty()) dv.append_attribute("prompt") = help_text.c_str();
      dv.append_attribute("sqref") = sqref.c_str();
      if (!data_validations) data_validations = worksheet.append_child("dataValidations");
      data_validations.append_copy(dv);
    }
    if (data_validations) data_validations.prepend_attribute("count") = static_cast<uint32_t>(std::distance(data_validations.children().begin(), data_validations.children().end()));

    // hyperlinks: targets starting with # are references inside the workbook
    pugi::xml_document sheet_rels;
    pugi::xml_node srels = sheet_rels.append_child("Relationships");
    srels.append_attribute("xmlns") = "http://schemas.openxmlformats.org/package/2006/relationships";
    if (!link_refs.empty()) {
      pugi::xml_node hyperlinks = worksheet.append_child("hyperlinks");
      for (size_t k = 0; k < link_refs.size(); ++k) {
        pugi::xml_node hl = hyperlinks.append_child("hyperlink");
        hl.append_attribute("ref") = link_refs[k].c_str();
        const std::string& target = link_targets[k];
        if (!target.empty() && target[0] == '#') {
          std::string location = target.substr(1);
          // some producers write the target in the xlsx syntax with "!" after the sheet
          bool quoted = false, xlsx_syntax = false;
          for (char ch : location) {
            if (ch == '\'') quoted = !quoted;
            else if (ch == '!' && !quoted) xlsx_syntax = true;
          }
          if (!xlsx_syntax && location.find('.') != std::string::npos) location = ods_of_ref(location, renames);
          hl.append_attribute("location") = location.c_str();
        } else {
          std::string id = "rId" + std::to_string(k + 1);
          hl.append_attribute("r:id") = id.c_str();
          pugi::xml_node rel = srels.append_child("Relationship");
          rel.append_attribute("Id") = id.c_str();
          rel.append_attribute("Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
          rel.append_attribute("Target") = target.c_str();
          rel.append_attribute("TargetMode") = "External";
        }
        if (!link_tips[k].empty()) hl.append_attribute("tooltip") = link_tips[k].c_str();
      }
    }

    // print options, margins, page setup, header and footer
    if (layout) {
      // the printed parts are a list of words
      std::set<std::string> print;
      std::string words = page_props.attribute("style:print").as_string();
      for (size_t pos = 0; pos < words.size();) {
        size_t end = words.find(' ', pos);
        if (end == std::string::npos) end = words.size();
        if (end > pos) print.insert(words.substr(pos, end - pos));
        pos = end + 1;
      }
      std::string centering = page_props.attribute("style:table-centering").as_string("none");
      pugi::xml_node print_options = worksheet.append_child("printOptions");
      if (centering == "horizontal" || centering == "both") print_options.append_attribute("horizontalCentered") = 1;
      if (centering == "vertical" || centering == "both") print_options.append_attribute("verticalCentered") = 1;
      if (print.count("headers")) print_options.append_attribute("headings") = 1;
      if (print.count("grid")) print_options.append_attribute("gridLines") = 1;
      if (!print_options.first_attribute()) worksheet.remove_child(print_options);

      pugi::xml_node header = master.child("style:header"), footer = master.child("style:footer");
      bool has_header = header && header.attribute("style:display").as_bool(true);
      bool has_footer = footer && footer.attribute("style:display").as_bool(true);
      pugi::xml_node header_props = layout.child("style:header-style").child("style:header-footer-properties");
      pugi::xml_node footer_props = layout.child("style:footer-style").child("style:header-footer-properties");
      double margin_top = ods_length_pt(page_props.attribute("fo:margin-top").as_string("54pt")) / 72.0;
      double margin_bottom = ods_length_pt(page_props.attribute("fo:margin-bottom").as_string("54pt")) / 72.0;
      double header_height = (ods_length_pt(header_props.attribute("fo:min-height").as_string("0pt")) +
                              ods_length_pt(header_props.attribute("fo:margin-bottom").as_string("0pt"))) / 72.0;
      double footer_height = (ods_length_pt(footer_props.attribute("fo:min-height").as_string("0pt")) +
                              ods_length_pt(footer_props.attribute("fo:margin-top").as_string("0pt"))) / 72.0;
      pugi::xml_node margins = worksheet.append_child("pageMargins");
      margins.append_attribute("left") = ods_num(ods_length_pt(page_props.attribute("fo:margin-left").as_string("50.4pt")) / 72.0).c_str();
      margins.append_attribute("right") = ods_num(ods_length_pt(page_props.attribute("fo:margin-right").as_string("50.4pt")) / 72.0).c_str();
      margins.append_attribute("top") = ods_num(has_header ? margin_top + header_height : margin_top).c_str();
      margins.append_attribute("bottom") = ods_num(has_footer ? margin_bottom + footer_height : margin_bottom).c_str();
      margins.append_attribute("header") = ods_num(has_header ? margin_top : 0.3).c_str();
      margins.append_attribute("footer") = ods_num(has_footer ? margin_bottom : 0.3).c_str();

      pugi::xml_node setup = worksheet.append_child("pageSetup");
      double width = ods_length_pt(page_props.attribute("fo:page-width").as_string("0pt")) / 72.0 * 25.4;
      double height = ods_length_pt(page_props.attribute("fo:page-height").as_string("0pt")) / 72.0 * 25.4;
      for (const auto& ps : ods_papers) {
        bool same = std::fabs(ps.width - width) < 1.5 && std::fabs(ps.height - height) < 1.5;
        bool turned = std::fabs(ps.width - height) < 1.5 && std::fabs(ps.height - width) < 1.5;
        if (same || turned) {
          setup.append_attribute("paperSize") = ps.id;
          break;
        }
      }
      std::string orientation = page_props.attribute("style:print-orientation").as_string(width > height ? "landscape" : "portrait");
      setup.append_attribute("orientation") = orientation == "landscape" ? "landscape" : "portrait";
      double scale = std::atof(page_props.attribute("style:scale-to").as_string("100"));
      if (scale > 0.0 && scale != 100.0) setup.append_attribute("scale") = static_cast<int32_t>(std::lround(scale));
      if (fit) {
        setup.append_attribute("fitToWidth") = page_props.attribute("style:scale-to-X").as_int(0);
        setup.append_attribute("fitToHeight") = page_props.attribute("style:scale-to-Y").as_int(0);
      }
      if (std::strcmp(page_props.attribute("style:print-page-order").as_string(), "ltr") == 0) setup.append_attribute("pageOrder") = "overThenDown";
      if (print.count("annotations")) setup.append_attribute("cellComments") = "atEnd";
      int32_t first_page = page_props.attribute("style:first-page-number").as_int(0);
      if (first_page > 0) {
        setup.append_attribute("firstPageNumber") = first_page;
        setup.append_attribute("useFirstPageNumber") = 1;
      }

      static const char* const regions[3] = {"style:region-left", "style:region-center", "style:region-right"};
      static const char* const codes[3] = {"&L", "&C", "&R"};
      // headers and footers of odd (right), even (left) and first pages, of the calc extension before ODF 1.3
      static const char* const variants[3][2] = {
        {"style:header", "style:footer"}, {"style:header-left", "style:footer-left"}, {"style:header-first", "style:footer-first"}
      };
      static const char* const parts[3][2] = {{"oddHeader", "oddFooter"}, {"evenHeader", "evenFooter"}, {"firstHeader", "firstFooter"}};
      std::string codes_of[3][2];
      for (int32_t variant = 0; variant < 3; ++variant) {
        for (int32_t hf = 0; hf < 2; ++hf) {
          if (!(hf == 0 ? has_header : has_footer)) continue;
          pugi::xml_node node = variant == 0 ? (hf == 0 ? header : footer) : master.child(variants[variant][hf]);
          if (!node && variant == 2) node = master.child(hf == 0 ? "loext:header-first" : "loext:footer-first");
          if (!node || !node.attribute("style:display").as_bool(true)) continue;
          std::string code;
          if (!node.child("style:region-left") && !node.child("style:region-center") && !node.child("style:region-right")) {
            std::string text = ods_header_code(node, st);
            if (!text.empty()) code = "&C" + text;
          } else {
            for (int32_t region = 0; region < 3; ++region) {
              std::string text = ods_header_code(node.child(regions[region]), st);
              if (!text.empty()) code += codes[region] + text;
            }
          }
          codes_of[variant][hf] = code;
        }
      }
      bool even = !codes_of[1][0].empty() || !codes_of[1][1].empty(), first = !codes_of[2][0].empty() || !codes_of[2][1].empty();
      pugi::xml_node header_footer;
      for (int32_t variant = 0; variant < 3; ++variant) {
        for (int32_t hf = 0; hf < 2; ++hf) {
          if (codes_of[variant][hf].empty()) continue;
          if (!header_footer) header_footer = worksheet.append_child("headerFooter");
          header_footer.append_child(parts[variant][hf]).text().set(codes_of[variant][hf].c_str());
        }
      }
      if (header_footer && first) header_footer.prepend_attribute("differentFirst") = 1;
      if (header_footer && even) header_footer.prepend_attribute("differentOddEven") = 1;
    }

    // manual page breaks after a row or column
    for (int32_t dim = 0; dim < 2; ++dim) {
      const std::vector<int64_t>& breaks = dim == 0 ? row_breaks : col_breaks;
      if (breaks.empty()) continue;
      pugi::xml_node node = worksheet.append_child(dim == 0 ? "rowBreaks" : "colBreaks");
      node.append_attribute("count") = static_cast<uint32_t>(breaks.size());
      node.append_attribute("manualBreakCount") = static_cast<uint32_t>(breaks.size());
      for (int64_t brk : breaks) {
        pugi::xml_node b = node.append_child("brk");
        b.append_attribute("id") = brk - 1;
        b.append_attribute("max") = dim == 0 ? MAX_OOXML_COL_INT - 1 : MAX_OOXML_ROW_INT - 1;
        b.append_attribute("man") = 1;
      }
    }

    // the relationships after those of the hyperlinks
    size_t rel_id = link_refs.size();

    // the background image of the table is the background picture of the sheet
    if (ts != table_styles.end()) {
      std::string image = ts->second.child("style:table-properties").child("style:background-image").attribute("xlink:href").as_string();
      std::string ext = image.substr(image.rfind('.') == std::string::npos ? image.size() : image.rfind('.') + 1);
      for (char& ch : ext) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
      std::ifstream in((base + image).c_str(), std::ios::binary);
      if (!image.empty() && in && (ext == "png" || ext == "jpeg" || ext == "jpg" || ext == "gif" || ext == "bmp")) {
        std::string target = "image" + std::to_string(++background_images) + "." + ext;
        std::ofstream out((path + "xl/media/" + target).c_str(), std::ios::binary);
        out << in.rdbuf();
        media_extensions.insert(ext);
        std::string picture_id = "rId" + std::to_string(++rel_id);
        worksheet.append_child("picture").append_attribute("r:id") = picture_id.c_str();
        pugi::xml_node rel = srels.append_child("Relationship");
        rel.append_attribute("Id") = picture_id.c_str();
        rel.append_attribute("Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        rel.append_attribute("Target") = ("../media/" + target).c_str();
      }
    }

    // tables, the column names are the cells of the header row
    pugi::xml_node table_parts_node;
    for (const ods_db_range& db : sheet_tables) {
      size_t colon = db.ref.find(':');
      std::string from = db.ref.substr(0, colon), to = db.ref.substr(colon + 1);
      int64_t tr1 = cell_to_rowint(from), tc1 = cell_to_colint(from), tc2 = cell_to_colint(to);
      std::vector<std::string> column_names;
      for (int64_t tc = tc1; tc <= tc2; ++tc) {
        std::string column;
        if (db.header) {
          for (const ods_in_cell& cell : cells) {
            if (cell.row == tr1 && cell.col == tc) column = cell.t == "inlineStr" ? cell.text : cell.v;
          }
        }
        if (column.empty()) column = "Column" + std::to_string(tc - tc1 + 1);
        std::string unique = column;
        for (int32_t n = 2; std::find(column_names.begin(), column_names.end(), unique) != column_names.end(); ++n) unique = column + std::to_string(n);
        column_names.push_back(unique);
      }

      std::string id = std::to_string(++table_parts);
      pugi::xml_document tdoc;
      ods_declaration(tdoc);
      pugi::xml_node tn = tdoc.append_child("table");
      tn.append_attribute("xmlns") = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
      tn.append_attribute("id") = id.c_str();
      tn.append_attribute("name") = db.name.c_str();
      tn.append_attribute("displayName") = db.name.c_str();
      tn.append_attribute("ref") = db.ref.c_str();
      if (!db.header) tn.append_attribute("headerRowCount") = 0;
      if (db.filter) tn.append_child("autoFilter").append_attribute("ref") = db.ref.c_str();
      pugi::xml_node tcs = tn.append_child("tableColumns");
      tcs.append_attribute("count") = static_cast<uint32_t>(column_names.size());
      for (size_t k = 0; k < column_names.size(); ++k) {
        pugi::xml_node tc = tcs.append_child("tableColumn");
        tc.append_attribute("id") = static_cast<uint32_t>(k + 1);
        tc.append_attribute("name") = column_names[k].c_str();
      }
      ods_save(tdoc, path + "xl/tables/table" + id + ".xml");

      if (!table_parts_node) table_parts_node = worksheet.append_child("tableParts");
      std::string table_id = "rId" + std::to_string(++rel_id);
      table_parts_node.append_child("tablePart").append_attribute("r:id") = table_id.c_str();
      pugi::xml_node rel = srels.append_child("Relationship");
      rel.append_attribute("Id") = table_id.c_str();
      rel.append_attribute("Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
      rel.append_attribute("Target") = ("../tables/table" + id + ".xml").c_str();
    }
    if (table_parts_node) table_parts_node.prepend_attribute("count") = static_cast<uint32_t>(sheet_tables.size());

    // sparkline groups of the calc extension as extension of the worksheet
    if (pugi::xml_node groups = table.child("calcext:sparkline-groups")) {
      static const char* const spark_flags[][2] = {
        {"calcext:date-axis", "dateAxis"}, {"calcext:markers", "markers"}, {"calcext:high", "high"}, {"calcext:low", "low"},
        {"calcext:first", "first"}, {"calcext:last", "last"}, {"calcext:negative", "negative"},
        {"calcext:display-x-axis", "displayXAxis"}, {"calcext:display-hidden", "displayHidden"}, {"calcext:right-to-left", "rightToLeft"}
      };
      static const char* const colors[][2] = {
        {"calcext:color-series", "x14:colorSeries"}, {"calcext:color-negative", "x14:colorNegative"}, {"calcext:color-axis", "x14:colorAxis"},
        {"calcext:color-markers", "x14:colorMarkers"}, {"calcext:color-first", "x14:colorFirst"}, {"calcext:color-last", "x14:colorLast"},
        {"calcext:color-high", "x14:colorHigh"}, {"calcext:color-low", "x14:colorLow"}
      };
      pugi::xml_node ext = ext_list.append_child("ext");
      ext.append_attribute("xmlns:x14") = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
      ext.append_attribute("uri") = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
      pugi::xml_node out_groups = ext.append_child("x14:sparklineGroups");
      out_groups.append_attribute("xmlns:xm") = "http://schemas.microsoft.com/office/excel/2006/main";
      for (pugi::xml_node group : groups.children("calcext:sparkline-group")) {
        pugi::xml_node out = out_groups.append_child("x14:sparklineGroup");
        std::string type = group.attribute("calcext:type").as_string("line");
        if (type != "line") out.append_attribute("type") = type.c_str();
        double weight = ods_length_pt(group.attribute("calcext:line-width").as_string("0.75pt"));
        if (std::fabs(weight - 0.75) > 0.001) out.append_attribute("lineWeight") = ods_num(weight).c_str();
        out.append_attribute("displayEmptyCellsAs") = group.attribute("calcext:display-empty-cells-as").as_string("zero");
        for (const auto& flag : spark_flags) {
          if (group.attribute(flag[0]).as_bool()) out.append_attribute(flag[1]) = 1;
        }
        std::string min_type = group.attribute("calcext:min-axis-type").as_string("individual");
        std::string max_type = group.attribute("calcext:max-axis-type").as_string("individual");
        if (min_type != "individual") out.append_attribute("minAxisType") = min_type.c_str();
        if (max_type != "individual") out.append_attribute("maxAxisType") = max_type.c_str();
        if (group.attribute("calcext:manual-min")) out.append_attribute("manualMin") = group.attribute("calcext:manual-min").as_string();
        if (group.attribute("calcext:manual-max")) out.append_attribute("manualMax") = group.attribute("calcext:manual-max").as_string();
        for (const auto& color : colors) {
          std::string rgb = ods_rgb(group.attribute(color[0]).as_string());
          if (!rgb.empty()) out.append_child(color[1]).append_attribute("rgb") = rgb.c_str();
        }
        pugi::xml_node lines = out.append_child("x14:sparklines");
        for (pugi::xml_node sparkline : group.child("calcext:sparklines").children("calcext:sparkline")) {
          std::string data = ods_of_ref(sparkline.attribute("calcext:data-range").as_string(), renames);
          std::string cell = ods_of_ref(sparkline.attribute("calcext:cell-address").as_string(), renames);
          cell = cell.substr(cell.rfind('!') == std::string::npos ? 0 : cell.rfind('!') + 1);
          if (data.empty() || cell.empty() || data == "#REF!") continue;
          pugi::xml_node line = lines.append_child("x14:sparkline");
          line.append_child("xm:f").text().set(data.c_str());
          line.append_child("xm:sqref").text().set(cell.c_str());
        }
        if (!lines.first_child()) out_groups.remove_child(out);
      }
      if (!out_groups.first_child()) ext_list.remove_child(ext);
    }
    if (ext_list.first_child()) worksheet.append_copy(ext_list);

    if (srels.first_child()) ods_save(sheet_rels, path + "xl/worksheets/_rels/sheet" + std::to_string(names.size()) + ".xml.rels");
    ods_save(ws, path + "xl/worksheets/sheet" + std::to_string(names.size()) + ".xml");
  }

  if (names.empty()) Rcpp::stop("the file has no tables");

  std::string buf = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
  if (!st.numfmt_list.empty()) {
    pugi::xml_document nf_doc;
    pugi::xml_node nfs = nf_doc.append_child("numFmts");
    nfs.append_attribute("count") = static_cast<uint32_t>(st.numfmt_list.size());
    for (size_t i = 0; i < st.numfmt_list.size(); ++i) {
      pugi::xml_node nf = nfs.append_child("numFmt");
      nf.append_attribute("numFmtId") = static_cast<uint32_t>(164 + i);
      nf.append_attribute("formatCode") = st.numfmt_list[i].c_str();
    }
    buf += ods_xml_string(nfs);
  }
  buf += "<fonts count=\"" + std::to_string(st.font_list.size()) + "\">";
  for (const std::string& x : st.font_list) buf += x;
  buf += "</fonts><fills count=\"" + std::to_string(st.fill_list.size()) + "\">";
  for (const std::string& x : st.fill_list) buf += x;
  buf += "</fills><borders count=\"" + std::to_string(st.border_list.size()) + "\">";
  for (const std::string& x : st.border_list) buf += x;
  buf += "</borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>";
  buf += "<cellXfs count=\"" + std::to_string(st.xf_list.size()) + "\">";
  for (const std::string& x : st.xf_list) buf += x;
  buf += "</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>";
  if (!dxf_list.empty()) {
    buf += "<dxfs count=\"" + std::to_string(dxf_list.size()) + "\">";
    for (const std::string& x : dxf_list) buf += x;
    buf += "</dxfs>";
  }
  buf += "</styleSheet>";
  pugi::xml_document xstyles;
  if (!xstyles.load_string(buf.c_str())) Rcpp::stop("could not create the styles");
  ods_declaration(xstyles);
  ods_save(xstyles, path + "xl/styles.xml");

  pugi::xml_document wb, wb_rels, rels, types;
  ods_declaration(wb);
  pugi::xml_node workbook = wb.append_child("workbook");
  workbook.append_attribute("xmlns") = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  workbook.append_attribute("xmlns:r") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  if (date1904) workbook.append_child("workbookPr").append_attribute("date1904") = 1;
  if (spreadsheet.attribute("table:structure-protected").as_bool()) {
    workbook.append_child("workbookProtection").append_attribute("lockStructure") = 1;
  }
  auto active = std::find(names.begin(), names.end(), active_table);
  pugi::xml_node book_view = workbook.append_child("bookViews").append_child("workbookView");
  if (!show_tabs) book_view.append_attribute("showSheetTabs") = 0;
  if (active != names.end()) book_view.append_attribute("activeTab") = static_cast<int32_t>(active - names.begin());
  pugi::xml_node xsheets = workbook.append_child("sheets");

  for (const auto& container : name_containers) {
    for (pugi::xml_node ne : container.first.children()) {
      ods_name dn = {ne.attribute("table:name").as_string(), container.second, ""};
      std::string type = ne.name();
      if (type == "table:named-range") {
        dn.formula = ods_of_ref(ne.attribute("table:cell-range-address").as_string(), renames);
      } else if (type == "table:named-expression") {
        std::string expression = ne.attribute("table:expression").as_string();
        if (expression.compare(0, 3, "of:") == 0) dn.formula = ods_formula_to_xlsx(expression.substr(3), renames);
        else if (expression.compare(0, 5, "oooc:") == 0) dn.formula = ods_formula_to_xlsx(expression.substr(5), renames);
        else if (expression.find(':') == std::string::npos || expression.find(':') > expression.find('=')) dn.formula = ods_formula_to_xlsx(expression, renames);
      }
      if (!dn.name.empty() && !dn.formula.empty()) defined.push_back(dn);
    }
  }

  ods_declaration(wb_rels);
  pugi::xml_node wrels = wb_rels.append_child("Relationships");
  wrels.append_attribute("xmlns") = "http://schemas.openxmlformats.org/package/2006/relationships";

  ods_declaration(types);
  pugi::xml_node ctypes = types.append_child("Types");
  ctypes.append_attribute("xmlns") = "http://schemas.openxmlformats.org/package/2006/content-types";
  pugi::xml_node def = ctypes.append_child("Default");
  def.append_attribute("Extension") = "rels";
  def.append_attribute("ContentType") = "application/vnd.openxmlformats-package.relationships+xml";
  def = ctypes.append_child("Default");
  def.append_attribute("Extension") = "xml";
  def.append_attribute("ContentType") = "application/xml";
  for (const std::string& ext : media_extensions) {
    def = ctypes.append_child("Default");
    def.append_attribute("Extension") = ext.c_str();
    def.append_attribute("ContentType") = ("image/" + std::string(ext == "jpg" ? "jpeg" : ext)).c_str();
  }
  pugi::xml_node ovr = ctypes.append_child("Override");
  ovr.append_attribute("PartName") = "/xl/workbook.xml";
  ovr.append_attribute("ContentType") = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
  ovr = ctypes.append_child("Override");
  ovr.append_attribute("PartName") = "/xl/styles.xml";
  ovr.append_attribute("ContentType") = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

  for (size_t i = 0; i < names.size(); ++i) {
    std::string id = std::to_string(i + 1);
    pugi::xml_node sheet = xsheets.append_child("sheet");
    std::map<std::string, std::string>::const_iterator renamed = renames.find(names[i]);
    sheet.append_attribute("name") = (renamed == renames.end() ? names[i] : renamed->second).c_str();
    sheet.append_attribute("sheetId") = id.c_str();
    if (hidden_sheets[i]) sheet.append_attribute("state") = "hidden";
    sheet.append_attribute("r:id") = ("rId" + id).c_str();

    pugi::xml_node rel = wrels.append_child("Relationship");
    rel.append_attribute("Id") = ("rId" + id).c_str();
    rel.append_attribute("Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    rel.append_attribute("Target") = ("worksheets/sheet" + id + ".xml").c_str();

    ovr = ctypes.append_child("Override");
    ovr.append_attribute("PartName") = ("/xl/worksheets/sheet" + id + ".xml").c_str();
    ovr.append_attribute("ContentType") = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
  }
  for (size_t k = 1; k <= table_parts; ++k) {
    ovr = ctypes.append_child("Override");
    ovr.append_attribute("PartName") = ("/xl/tables/table" + std::to_string(k) + ".xml").c_str();
    ovr.append_attribute("ContentType") = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
  }
  if (!defined.empty()) {
    pugi::xml_node xnames = workbook.append_child("definedNames");
    for (const ods_name& dn : defined) {
      pugi::xml_node xn = xnames.append_child("definedName");
      xn.append_attribute("name") = dn.name.c_str();
      if (!dn.local.empty()) xn.append_attribute("localSheetId") = dn.local.c_str();
      xn.text().set(dn.formula.c_str());
    }
  }

  pugi::xml_node rel = wrels.append_child("Relationship");
  rel.append_attribute("Id") = ("rId" + std::to_string(names.size() + 1)).c_str();
  rel.append_attribute("Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
  rel.append_attribute("Target") = "styles.xml";

  ods_declaration(rels);
  pugi::xml_node prels = rels.append_child("Relationships");
  prels.append_attribute("xmlns") = "http://schemas.openxmlformats.org/package/2006/relationships";
  rel = prels.append_child("Relationship");
  rel.append_attribute("Id") = "rId1";
  rel.append_attribute("Type") = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
  rel.append_attribute("Target") = "xl/workbook.xml";

  ods_save(wb, path + "xl/workbook.xml");
  ods_save(wb_rels, path + "xl/_rels/workbook.xml.rels");
  // document properties
  pugi::xml_document meta_doc, core_doc;
  if (!meta_path.empty() && meta_doc.load_file(meta_path.c_str(), pugi::parse_default, pugi::encoding_utf8)) {
    ods_normalize_document(meta_doc);
    pugi::xml_node office_meta = meta_doc.document_element().child("office:meta");
    ods_declaration(core_doc);
    pugi::xml_node core = core_doc.append_child("cp:coreProperties");
    core.append_attribute("xmlns:cp") = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    core.append_attribute("xmlns:dc") = "http://purl.org/dc/elements/1.1/";
    core.append_attribute("xmlns:dcterms") = "http://purl.org/dc/terms/";
    core.append_attribute("xmlns:xsi") = "http://www.w3.org/2001/XMLSchema-instance";
    static const char* const properties[][2] = {
      {"dc:title", "dc:title"}, {"dc:subject", "dc:subject"}, {"meta:initial-creator", "dc:creator"},
      {"dc:description", "dc:description"}, {"dc:creator", "cp:lastModifiedBy"}
    };
    for (const auto& prop : properties) {
      std::string value = office_meta.child_value(prop[0]);
      if (!value.empty()) core.append_child(prop[1]).text().set(value.c_str());
    }
    std::string keywords;
    for (pugi::xml_node kw : office_meta.children("meta:keyword")) keywords += (keywords.empty() ? "" : ", ") + std::string(kw.text().get());
    if (!keywords.empty()) core.append_child("cp:keywords").text().set(keywords.c_str());
    for (int32_t k = 0; k < 2; ++k) {
      std::string value = office_meta.child_value(k == 0 ? "meta:creation-date" : "dc:date");
      if (value.empty()) continue;
      value = value.substr(0, value.find('.'));
      if (value.back() != 'Z') value += "Z";
      pugi::xml_node date = core.append_child(k == 0 ? "dcterms:created" : "dcterms:modified");
      date.append_attribute("xsi:type") = "dcterms:W3CDTF";
      date.text().set(value.c_str());
    }
    for (pugi::xml_node ud : office_meta.children("meta:user-defined")) {
      if (std::string(ud.attribute("meta:name").as_string()) == "Category") core.append_child("cp:category").text().set(ud.text().get());
    }
    ods_save(core_doc, path + "docProps/core.xml");
    rel = prels.append_child("Relationship");
    rel.append_attribute("Id") = "rId2";
    rel.append_attribute("Type") = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    rel.append_attribute("Target") = "docProps/core.xml";
    ovr = ctypes.append_child("Override");
    ovr.append_attribute("PartName") = "/docProps/core.xml";
    ovr.append_attribute("ContentType") = "application/vnd.openxmlformats-package.core-properties+xml";
  }

  ods_save(rels, path + "_rels/.rels");
  ods_save(types, path + "[Content_Types].xml");

  return Rcpp::List::create(Rcpp::Named("dropped") = dropped, Rcpp::Named("sheets") = sheet_extras);
}

// Crops and rotates the picture or sets the gradient of the shape of the last
// anchor of a drawing. Crops are in 1/100000 of the picture, angles in degrees.
// The properties of a shape are its flips, fill, line, effects and text box.
// [[Rcpp::export]]
utf8_string ods_adjust_drawing(utf8_string xml, std::vector<double> crop, double rotation,
                               utf8_string gradient_start, utf8_string gradient_end, double gradient_angle, utf8_string geometry,
                               std::vector<double> offsets, utf8_string properties) {
  pugi::xml_document doc;
  if (!doc.load_string(xml.c_str(), pugi::parse_default | pugi::parse_ws_pcdata)) return xml;
  pugi::xml_node anchor = doc.child("xdr:wsDr").last_child();

  // the start and end cell of the anchor and the offsets in them in EMU
  if (offsets.size() == 8 && std::strcmp(anchor.name(), "xdr:twoCellAnchor") == 0) {
    const char* const parts[8][2] = {
      {"xdr:from", "xdr:col"}, {"xdr:from", "xdr:colOff"}, {"xdr:from", "xdr:row"}, {"xdr:from", "xdr:rowOff"},
      {"xdr:to", "xdr:col"}, {"xdr:to", "xdr:colOff"}, {"xdr:to", "xdr:row"}, {"xdr:to", "xdr:rowOff"}
    };
    for (size_t k = 0; k < 8; ++k) {
      pugi::xml_node node = anchor.child(parts[k][0]).child(parts[k][1]);
      if (node) node.text().set(static_cast<int64_t>(std::llround(std::max(offsets[k], 0.0))));
    }
  }

  pugi::xml_node el = anchor.child("xdr:pic");
  if (!el) el = anchor.child("xdr:sp");
  pugi::xml_node sppr = el.child("xdr:spPr");
  if (!el || !sppr) return ods_xml_string(doc.document_element());

  if (rotation != 0.0) {
    pugi::xml_node xfrm = sppr.child("a:xfrm");
    if (!xfrm) {
      xfrm = sppr.prepend_child("a:xfrm");
      pugi::xml_node off = xfrm.append_child("a:off");
      off.append_attribute("x") = 0;
      off.append_attribute("y") = 0;
      pugi::xml_node ext = xfrm.append_child("a:ext");
      ext.append_attribute("cx") = 0;
      ext.append_attribute("cy") = 0;
    }
    pugi::xml_attribute rot = xfrm.attribute("rot");
    if (!rot) rot = xfrm.prepend_attribute("rot");
    rot = static_cast<int64_t>(std::llround(rotation * 60000.0));
  }

  pugi::xml_node blip = el.child("xdr:blipFill").child("a:blip");
  if (blip && crop.size() == 4 && (crop[0] != 0.0 || crop[1] != 0.0 || crop[2] != 0.0 || crop[3] != 0.0)) {
    pugi::xml_node src = el.child("xdr:blipFill").insert_child_after("a:srcRect", blip);
    static const char* const sides[] = {"l", "t", "r", "b"};
    for (size_t k = 0; k < 4; ++k) {
      if (crop[k] != 0.0) src.append_attribute(sides[k]) = static_cast<int64_t>(std::llround(crop[k]));
    }
  }

  pugi::xml_document props_doc;
  pugi::xml_node props = props_doc.load_string(properties.c_str()) ? props_doc.child("p") : pugi::xml_node();
  if (el.name() == std::string("xdr:sp") && props) {
    pugi::xml_node xfrm = sppr.child("a:xfrm");
    for (const char* flip : {"flipH", "flipV"}) {
      if (!props.attribute(flip) || !xfrm) continue;
      pugi::xml_attribute attr = xfrm.attribute(flip);
      if (!attr) attr = xfrm.append_attribute(flip);
      attr = 1;
    }
    // the fill, line and effects replace those of the shape in the order of the shape properties
    pugi::xml_node geom = sppr.child("a:prstGeom") ? sppr.child("a:prstGeom") : sppr.child("a:custGeom");
    pugi::xml_node last = geom;
    for (const char* name : {"a:noFill", "a:solidFill", "a:gradFill", "a:ln", "a:effectLst"}) {
      pugi::xml_node node = props.child(name);
      bool is_fill = std::strcmp(name, "a:ln") != 0 && std::strcmp(name, "a:effectLst") != 0;
      pugi::xml_node old = sppr.child(name);
      if (is_fill) old = sppr.child("a:noFill") ? sppr.child("a:noFill") : sppr.child("a:solidFill") ? sppr.child("a:solidFill") : sppr.child("a:gradFill");
      if (old) last = old;
      if (!node) continue;
      last = last ? sppr.insert_copy_after(node, last) : sppr.append_copy(node);
      if (old) sppr.remove_child(old);
    }
    pugi::xml_node body = el.child("xdr:txBody").child("a:bodyPr");
    if (pugi::xml_node pbody = props.child("a:bodyPr")) {
      for (pugi::xml_attribute attr : pbody.attributes()) {
        pugi::xml_attribute target = body.attribute(attr.name());
        if (!target) target = body.append_attribute(attr.name());
        target = attr.value();
      }
      for (pugi::xml_node child : pbody.children()) body.append_copy(child);
    }
  }

  // a custom geometry instead of the preset
  pugi::xml_document geometry_doc;
  if (el.name() == std::string("xdr:sp") && !geometry.empty() && geometry_doc.load_string(geometry.c_str()) && sppr.child("a:prstGeom")) {
    sppr.insert_copy_before(geometry_doc.first_child(), sppr.child("a:prstGeom"));
    sppr.remove_child(sppr.child("a:prstGeom"));
  }

  if (el.name() == std::string("xdr:sp") && !gradient_start.empty() && !gradient_end.empty()) {
    pugi::xml_node fill = sppr.child("a:solidFill");
    if (!fill) fill = sppr.child("a:noFill");
    pugi::xml_node grad = fill ? sppr.insert_child_before("a:gradFill", fill) : sppr.insert_child_after("a:gradFill", sppr.child("a:prstGeom") ? sppr.child("a:prstGeom") : sppr.child("a:custGeom"));
    if (fill) sppr.remove_child(fill);
    pugi::xml_node list = grad.append_child("a:gsLst");
    const std::string colors[] = {gradient_start, gradient_end};
    for (size_t k = 0; k < 2; ++k) {
      pugi::xml_node gs = list.append_child("a:gs");
      gs.append_attribute("pos") = k == 0 ? 0 : 100000;
      gs.append_child("a:srgbClr").append_attribute("val") = colors[k].substr(colors[k].size() > 6 ? colors[k].size() - 6 : 0).c_str();
    }
    pugi::xml_node lin = grad.append_child("a:lin");
    lin.append_attribute("ang") = static_cast<int64_t>(std::llround(gradient_angle * 60000.0));
    lin.append_attribute("scaled") = 0;
  }
  return ods_xml_string(doc.document_element());
}

static std::string ods_base64_encode(const std::string& in) {
  static const char* const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  std::string out;
  out.reserve((in.size() + 2) / 3 * 4);
  for (size_t i = 0; i < in.size(); i += 3) {
    uint32_t n = static_cast<uint32_t>(static_cast<unsigned char>(in[i])) << 16;
    if (i + 1 < in.size()) n |= static_cast<uint32_t>(static_cast<unsigned char>(in[i + 1])) << 8;
    if (i + 2 < in.size()) n |= static_cast<uint32_t>(static_cast<unsigned char>(in[i + 2]));
    out += chars[(n >> 18) & 63];
    out += chars[(n >> 12) & 63];
    out += i + 1 < in.size() ? chars[(n >> 6) & 63] : '=';
    out += i + 2 < in.size() ? chars[n & 63] : '=';
  }
  return out;
}

// characters outside of the alphabet, such as line breaks, are skipped
static std::string ods_base64_decode(const char* in) {
  std::string out;
  uint32_t n = 0;
  int32_t bits = 0;
  for (const char* p = in; *p && *p != '='; ++p) {
    char c = *p;
    int32_t v = c >= 'A' && c <= 'Z' ? c - 'A' : c >= 'a' && c <= 'z' ? c - 'a' + 26 : c >= '0' && c <= '9' ? c - '0' + 52 :
                c == '+' ? 62 : c == '/' ? 63 : -1;
    if (v < 0) continue;
    n = (n << 6) | static_cast<uint32_t>(v);
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out += static_cast<char>((n >> bits) & 0xFF);
    }
  }
  return out;
}

static std::string ods_read_file(const std::string& path) {
  std::ifstream in(path.c_str(), std::ios::binary);
  if (!in) Rcpp::stop("could not read %s", path);
  return std::string((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

static void ods_write_file(const std::string& path, const std::string& data) {
  std::ofstream out(path.c_str(), std::ios::binary);
  if (!out) Rcpp::stop("could not write %s", path);
  out.write(data.data(), static_cast<std::streamsize>(data.size()));
}

// the namespace declarations of a root element added to another element
static void ods_copy_namespaces(pugi::xml_node from, pugi::xml_node to) {
  for (pugi::xml_attribute attr : from.attributes()) {
    std::string name = attr.name();
    if ((name == "xmlns" || name.compare(0, 6, "xmlns:") == 0) && !to.attribute(attr.name())) to.append_attribute(attr.name()) = attr.value();
  }
}

// The parts of an ods file in dir as flat ods file: the pictures as binary data and the
// chart objects as documents in their frames.
// [[Rcpp::export]]
void ods_flatten(native_path dir, native_path file) {
  std::string path = std::string(dir) + "/";
  uint32_t flags = pugi::parse_default | pugi::parse_ws_pcdata;
  pugi::xml_document parts[4];
  static const char* const names[4] = {"meta.xml", "settings.xml", "styles.xml", "content.xml"};
  for (size_t k = 0; k < 4; ++k) {
    if (!parts[k].load_file((path + names[k]).c_str(), flags, pugi::encoding_utf8)) Rcpp::stop("could not read %s", names[k]);
  }
  pugi::xml_node meta = parts[0].document_element(), settings = parts[1].document_element();
  pugi::xml_node styles = parts[2].document_element(), content = parts[3].document_element();

  pugi::xml_document doc;
  ods_declaration(doc);
  pugi::xml_node root = doc.append_child("office:document");
  for (pugi::xml_node part : {content, styles, meta, settings}) ods_copy_namespaces(part, root);
  root.append_attribute("office:version") = content.attribute("office:version").as_string("1.4");
  root.append_attribute("office:mimetype") = "application/vnd.oasis.opendocument.spreadsheet";

  root.append_copy(meta.child("office:meta"));
  root.append_copy(settings.child("office:settings"));
  // the font faces of both parts
  pugi::xml_node faces = root.append_child("office:font-face-decls");
  std::set<std::string> face_names;
  for (pugi::xml_node part : {styles, content}) {
    for (pugi::xml_node face : part.child("office:font-face-decls").children()) {
      if (face_names.insert(face.attribute("style:name").as_string()).second) faces.append_copy(face);
    }
  }
  if (!faces.first_child()) root.remove_child(faces);
  root.append_copy(styles.child("office:styles"));
  // the automatic styles of both parts: page layouts of the styles, cell, row, column and other styles of the content
  pugi::xml_node autos = root.append_child("office:automatic-styles");
  for (pugi::xml_node part : {styles, content}) {
    for (pugi::xml_node style : part.child("office:automatic-styles").children()) autos.append_copy(style);
  }
  root.append_copy(styles.child("office:master-styles"));
  pugi::xml_node body = root.append_copy(content.child("office:body"));

  // the pictures as binary data and the chart objects as documents
  std::function<void(pugi::xml_node)> embed = [&](pugi::xml_node node) {
    for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
      if (child.type() != pugi::node_element) continue;
      std::string href = child.attribute("xlink:href").as_string();
      bool picture = href.compare(0, 9, "Pictures/") == 0;
      bool object = std::strcmp(child.name(), "draw:object") == 0 && href.compare(0, 2, "./") == 0;
      if (picture || object) {
        for (const char* attr : {"xlink:href", "xlink:type", "xlink:show", "xlink:actuate"}) child.remove_attribute(attr);
      }
      if (picture) {
        child.prepend_child("office:binary-data").text().set(ods_base64_encode(ods_read_file(path + href)).c_str());
        continue;
      }
      if (object) {
        std::string object_dir = path + href.substr(2) + "/";
        pugi::xml_document object_content, object_styles;
        if (!object_content.load_file((object_dir + "content.xml").c_str(), flags, pugi::encoding_utf8)) Rcpp::stop("could not read %s", href);
        object_styles.load_file((object_dir + "styles.xml").c_str(), flags, pugi::encoding_utf8);
        pugi::xml_node inner = child.append_child("office:document");
        ods_copy_namespaces(object_content.document_element(), inner);
        ods_copy_namespaces(object_styles.document_element(), inner);
        for (pugi::xml_attribute attr : root.attributes()) {
          if (inner.attribute(attr.name()) && std::string(inner.attribute(attr.name()).value()) == attr.value()) inner.remove_attribute(attr.name());
        }
        inner.append_attribute("office:version") = object_content.document_element().attribute("office:version").as_string("1.4");
        inner.append_attribute("office:mimetype") = "application/vnd.oasis.opendocument.chart";
        pugi::xml_node object_style_list = object_styles.document_element().child("office:styles");
        if (object_style_list && object_style_list.first_child()) inner.append_copy(object_style_list);
        inner.append_copy(object_content.document_element().child("office:automatic-styles"));
        inner.append_copy(object_content.document_element().child("office:body"));
        continue;
      }
      embed(child);
    }
  };
  embed(root.child("office:styles"));
  embed(autos);
  embed(body);
  ods_save(doc, file);
}

// The parts of an ods file in dir for a flat ods file: the binary data of pictures as files in
// Pictures and the chart documents of frames as objects.
// [[Rcpp::export]]
void ods_unflatten(native_path file, native_path dir) {
  std::string path = std::string(dir) + "/";
  pugi::xml_document doc;
  pugi::xml_parse_result res = doc.load_file(std::string(file).c_str(), pugi::parse_default | pugi::parse_ws_pcdata);
  if (!res) Rcpp::stop("could not read %s: %s", std::string(file), res.description());
  ods_normalize_document(doc);
  pugi::xml_node root = doc.document_element();

  // the pictures by the type of their data, the objects in the order of the document
  Rcpp::Function dir_create("dir.create");
  int32_t pictures = 0, objects = 0;
  std::function<void(pugi::xml_node)> extract = [&](pugi::xml_node node) {
    for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
      if (child.type() != pugi::node_element) continue;
      pugi::xml_node data = child.child("office:binary-data");
      pugi::xml_node inner = child.child("office:document");
      if (data) {
        std::string bytes = ods_base64_decode(data.text().get());
        const char* ext = ".bin";
        if (bytes.compare(0, 4, "\x89PNG") == 0) ext = ".png";
        else if (bytes.compare(0, 3, "\xFF\xD8\xFF") == 0) ext = ".jpeg";
        else if (bytes.compare(0, 4, "GIF8") == 0) ext = ".gif";
        else if (bytes.compare(0, 2, "BM") == 0) ext = ".bmp";
        else if (bytes.compare(0, 4, "\xD7\xCD\xC6\x9A") == 0) ext = ".wmf";
        else if (bytes.size() > 44 && bytes.compare(40, 4, " EMF") == 0) ext = ".emf";
        else if (bytes.find("<svg") != std::string::npos) ext = ".svg";
        if (pictures == 0) dir_create(path + "Pictures");
        std::string href = "Pictures/image" + std::to_string(++pictures) + ext;
        ods_write_file(path + href, bytes);
        child.remove_child(data);
        child.append_attribute("xlink:href") = href.c_str();
        child.append_attribute("xlink:type") = "simple";
        continue;
      }
      if (std::strcmp(child.name(), "draw:object") == 0 && inner) {
        std::string name = "Object " + std::to_string(++objects);
        dir_create(path + name);
        pugi::xml_document object_content, object_styles;
        ods_declaration(object_content);
        pugi::xml_node oc = object_content.append_child("office:document-content");
        ods_copy_namespaces(root, oc);
        ods_copy_namespaces(inner, oc);
        oc.append_attribute("office:version") = inner.attribute("office:version").as_string("1.4");
        for (const char* part : {"office:font-face-decls", "office:automatic-styles", "office:body"}) {
          if (inner.child(part)) oc.append_copy(inner.child(part));
        }
        ods_save(object_content, path + name + "/content.xml");
        ods_declaration(object_styles);
        pugi::xml_node os = object_styles.append_child("office:document-styles");
        ods_copy_namespaces(root, os);
        ods_copy_namespaces(inner, os);
        os.append_attribute("office:version") = inner.attribute("office:version").as_string("1.4");
        if (inner.child("office:styles")) os.append_copy(inner.child("office:styles"));
        ods_save(object_styles, path + name + "/styles.xml");
        child.remove_child(inner);
        child.append_attribute("xlink:href") = ("./" + name).c_str();
        child.append_attribute("xlink:type") = "simple";
        continue;
      }
      extract(child);
    }
  };
  extract(root);

  // the parts with the namespaces of the document
  static const char* const parts[4][2] = {
    {"content.xml", "office:document-content"}, {"styles.xml", "office:document-styles"},
    {"meta.xml", "office:document-meta"}, {"settings.xml", "office:document-settings"}
  };
  static const char* const children[4][5] = {
    {"office:scripts", "office:font-face-decls", "office:automatic-styles", "office:body", ""},
    {"office:font-face-decls", "office:styles", "office:automatic-styles", "office:master-styles", ""},
    {"office:meta", "", "", "", ""},
    {"office:settings", "", "", "", ""}
  };
  for (size_t k = 0; k < 4; ++k) {
    pugi::xml_document part;
    ods_declaration(part);
    pugi::xml_node out = part.append_child(parts[k][1]);
    ods_copy_namespaces(root, out);
    out.append_attribute("office:version") = root.attribute("office:version").as_string("1.4");
    for (const char* name : children[k]) {
      if (*name && root.child(name)) out.append_copy(root.child(name));
    }
    ods_save(part, path + parts[k][0]);
  }
}
