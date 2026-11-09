test_that("ods files store the mimetype first and uncompressed", {
  fl <- tempfile(fileext = ".ods")
  write_ods(wb_workbook()$add_worksheet()$add_data(x = 1), fl)

  files <- utils::unzip(fl, list = TRUE)
  expect_equal(files$Name[1], "mimetype")
  expect_true(all(c("content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml") %in% files$Name))

  header <- readBin(fl, "raw", 38L)
  expect_equal(header[9:10], as.raw(c(0, 0)))
  expect_equal(rawToChar(header[31:38]), "mimetype")
})

test_that("values survive writing and loading ods files", {
  df <- data.frame(
    num  = c(1.5, -2, 1e10),
    chr  = c("a  b", " lead", "line1\nline2"),
    lgl  = c(TRUE, FALSE, NA),
    date = as.Date(c("2026-11-09", "1999-12-31", "1904-01-01")),
    time = as.POSIXct(c("2026-11-09 12:34:56", "2000-01-01 00:00:00", "2026-01-01 23:59:59"), tz = "UTC"),
    stringsAsFactors = FALSE
  )
  df$chr[2] <- "\u00e4\u00f6\u00fc \u00df"

  wb <- wb_workbook()$add_worksheet("Data")$add_data(x = df)
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)

  wb2 <- load_ods(fl)
  expect_equal(unname(wb2$get_sheet_names()), "Data")

  got <- wb_to_df(wb2)
  expect_equal(got$num, df$num)
  expect_equal(got$chr, df$chr)
  expect_equal(got$lgl, df$lgl)
  expect_equal(got$date, df$date)
  expect_equal(as.numeric(got$time), as.numeric(df$time))
})

test_that("styles survive writing and loading ods files", {
  wb <- wb_workbook()$
    add_worksheet("Styled")$
    add_data(x = data.frame(a = c(0.25, 0.5), b = c("x", "y")))$
    add_fill(dims = "A1:B1", color = wb_color("FFFFFF00"))$
    add_font(dims = "A2", bold = TRUE, color = wb_color("FFFF0000"), size = 14)$
    add_border(dims = "B2")$
    add_numfmt(dims = "A2:A3", numfmt = "0.0%")$
    add_cell_style(dims = "B3", horizontal = "center", wrap_text = TRUE)$
    merge_cells(dims = "A5:B6")$
    set_col_widths(cols = 2, widths = 20)$
    set_row_heights(rows = 3, heights = 30)$
    add_worksheet("Hidden", visible = FALSE)$
    add_data(x = 1)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  wb2 <- load_ods(fl)

  expect_equal(unname(wb2$get_sheet_visibility()), c("visible", "hidden"))
  expect_equal(wb2$worksheets[[1]]$mergeCells, "<mergeCell ref=\"A5:B6\"/>")

  styles <- wb2$styles_mgr$styles
  expect_match(paste(styles$fills, collapse = ""), "FFFFFF00")
  expect_match(paste(styles$fonts, collapse = ""), "<b/>", fixed = TRUE)
  expect_match(paste(styles$fonts, collapse = ""), "FFFF0000")
  expect_match(paste(styles$borders, collapse = ""), "thin")
  expect_match(paste(styles$numFmts, collapse = ""), "0.0%", fixed = TRUE)
  expect_match(paste(styles$cellXfs, collapse = ""), "horizontal=\"center\"", fixed = TRUE)
  expect_match(paste(styles$cellXfs, collapse = ""), "wrapText=\"1\"", fixed = TRUE)

  cols <- wb2$worksheets[[1]]$cols_attr
  width <- as.numeric(xml_attr(cols[grepl("min=\"2\"", cols)], "col")[[1]][["width"]])
  expect_equal(width, 20.711, tolerance = 0.01)

  ra <- wb2$worksheets[[1]]$sheet_data$row_attr
  expect_equal(as.numeric(ra[ra$r == "3", "ht"]), 30, tolerance = 0.01)
})

test_that("formulas are translated to and from OpenFormula", {
  wb <- wb_workbook()$
    add_worksheet("S 1")$
    add_data(x = c(1, 2, 3))$
    add_worksheet("S2")$
    add_data(x = 4)$
    add_formula(dims = "B1", x = "IF('S 1'!A1>0,\"a\"&\"b\",\"c\")")$
    add_formula(dims = "B2", x = "SUM('S 1'!$A$1:A3,4)")$
    add_formula(dims = "B3", x = "_xlfn.CONCAT(A1,TRUE)+SUM((A1,A1))")$
    add_formula(dims = "B4", x = "SUM(A1:A3 A2:A4)+NORMSDIST(1)+_xlfn.IFNA(1,2)")$
    add_formula(dims = "C1:C3", x = "A1*2", shared = TRUE)$
    add_named_region(sheet = "S 1", dims = "A1:A3", name = "vals")$
    add_formula(sheet = "S2", dims = "D1", x = "SUM(vals)")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)

  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "of:=IF([$'S 1'.A1]>0;&quot;a&quot;&amp;&quot;b&quot;;&quot;c&quot;)", fixed = TRUE)
  expect_match(content, "of:=SUM([$'S 1'.$A$1:.A3];4)", fixed = TRUE)
  expect_match(content, "of:=COM.MICROSOFT.CONCAT([.A1];TRUE())+SUM(([.A1]~[.A1]))", fixed = TRUE)
  expect_match(content, "of:=SUM([.A1:.A3]![.A2:.A4])+LEGACY.NORMSDIST(1)+IFNA(1;2)", fixed = TRUE)
  expect_match(content, "of:=[.A3]*2", fixed = TRUE)
  expect_match(content, "table:named-range table:name=\"vals\" table:cell-range-address=\"$'S 1'.$A$1:.$A$3\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  got <- wb_to_df(wb2, sheet = 2, col_names = FALSE, show_formula = TRUE)
  expect_equal(
    as.character(got$B),
    c(
      "IF('S 1'!A1>0,\"a\"&\"b\",\"c\")", "SUM('S 1'!$A$1:A3,4)",
      "_xlfn.CONCAT(A1,TRUE)+SUM((A1,A1))", "SUM(A1:A3 A2:A4)+NORMSDIST(1)+_xlfn.IFNA(1,2)"
    )
  )
  expect_equal(as.character(got$C[1:3]), c("A1*2", "A2*2", "A3*2"))
  expect_equal(as.character(got$D[1]), "SUM(vals)")
  expect_equal(wb2$get_named_regions()$name, "vals")
  expect_equal(wb2$get_named_regions()$coords, "A1:A3")
})

test_that("structured references are written as references to the table cells", {
  wb <- wb_workbook()$
    add_worksheet("data")$
    add_data_table(x = data.frame(a = 1:3, b = 4:6), table_name = "tab")$
    add_formula(dims = "D2", x = "SUM(tab[a])")$
    add_formula(dims = "D3", x = "COUNTA(tab[#All])+SUM(tab[[a]:[b]])")$
    add_formula(dims = "D4", x = "[1]Sheet1!A1")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  wb2 <- load_ods(fl)

  got <- wb_to_df(wb2, col_names = FALSE, show_formula = TRUE, dims = "D2:D4")
  expect_equal(
    as.character(got$D),
    c("SUM(data!$A$2:$A$4)", "COUNTA(data!$A$1:$B$4)+SUM(data!$A$2:$B$4)", "#REF!")
  )
})

test_that("flat ods files are loaded", {
  fods <- '<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:automatic-styles>
  <number:number-style style:name="N1">
    <number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>
  </number:number-style>
  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N1">
    <style:table-cell-properties fo:background-color="#00ff00"/>
  </style:style>
  <style:style style:name="co1" style:family="table-column">
    <style:table-column-properties style:column-width="1in"/>
  </style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
<table:table table:name="T">
  <table:table-column table:style-name="co1" table:number-columns-repeated="2"/>
  <table:table-column table:number-columns-repeated="16382" table:default-cell-style-name="ce1"/>
  <table:table-row>
    <table:table-cell office:value-type="float" office:value="1234.5" table:style-name="ce1"><text:p>1,234.50</text:p></table:table-cell>
    <table:table-cell table:formula="of:=[.A1]*2+SUM({1;2|3;4})" office:value-type="float" office:value="2479"><text:p>2479</text:p></table:table-cell>
    <table:table-cell office:value-type="string"><text:p>a<text:s text:c="2"/>b</text:p><text:p>c</text:p></table:table-cell>
    <table:table-cell table:formula="ooow:=1+1" office:value-type="float" office:value="2"/>
    <table:table-cell table:number-columns-repeated="16380"/>
  </table:table-row>
  <table:table-row table:number-rows-repeated="2">
    <table:table-cell office:value-type="boolean" office:boolean-value="true" table:number-columns-repeated="2"/>
    <table:table-cell table:number-columns-repeated="16382"/>
  </table:table-row>
  <table:table-row table:number-rows-repeated="1048573">
    <table:table-cell table:number-columns-repeated="16384"/>
  </table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document>'
  fl <- tempfile(fileext = ".fods")
  writeLines(fods, fl, useBytes = TRUE)

  expect_warning(wb <- load_ods(fl), "1 formulas")

  got <- wb_to_df(wb, col_names = FALSE)
  expect_equal(dim(got), c(3L, 4L))
  expect_equal(got$A, c("1234.5", "TRUE", "TRUE"))
  expect_equal(got$B, c("2479", "TRUE", "TRUE"))
  expect_equal(got$C[1], "a  b\nc")
  expect_equal(got$D[1], 2)
  expect_equal(as.character(wb_to_df(wb, col_names = FALSE, show_formula = TRUE, dims = "B1")$B), "A1*2+SUM({1,2;3,4})")

  cc <- wb$worksheets[[1]]$sheet_data$cc
  xf <- xml_attr(wb$styles_mgr$styles$cellXfs[as.integer(cc$c_s[cc$r == "A1"]) + 1], "xf")[[1]]
  expect_equal(xf[["numFmtId"]], "4")
  expect_equal(xf[["fillId"]], "2")
  cols <- wb$worksheets[[1]]$cols_attr
  expect_match(cols[1], "max=\"2\"", fixed = TRUE)
  expect_match(cols[2], "style=\"1\"", fixed = TRUE)
})

test_that("number formats are converted in both directions", {
  codes <- c(
    "0.00", "#,##0", "0.0%", "0.00E+00", "# ?/?", "yyyy-mm-dd", "dd.mm.yyyy hh:mm:ss",
    "[h]:mm:ss", "mmm d, yyyy", "#,##0.00;[Red]-#,##0.00", "0.0;[Red]-0.0;\"zero\"",
    "[<100]0.0;[>=100]0;General", "0.0,,\" M\"", "[$\u20ac]#,##0.00", "@"
  )
  wb <- wb_workbook()$add_worksheet()
  for (i in seq_along(codes)) {
    wb$add_data(dims = paste0("A", i), x = 1)
    wb$add_numfmt(dims = paste0("A", i), numfmt = codes[i])
  }
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  wb2 <- load_ods(fl)

  styles <- wb2$styles_mgr$styles
  xf_ids <- vapply(styles$cellXfs, function(xf) xml_attr(xf, "xf")[[1]][["numFmtId"]], "")
  nfs <- xml_attr(styles$numFmts, "numFmt")
  nfs <- list(
    numFmtId   = vapply(nfs, function(nf) nf[["numFmtId"]], ""),
    formatCode = vapply(nfs, function(nf) nf[["formatCode"]], "")
  )
  cc <- wb2$worksheets[[1]]$sheet_data$cc
  ids <- xf_ids[as.integer(cc$c_s) + 1]
  got <- vapply(ids, function(id) {
    if (id %in% nfs$numFmtId) return(nfs$formatCode[nfs$numFmtId == id])
    builtin <- c("1" = "0", "2" = "0.00", "3" = "#,##0", "11" = "0.00E+00", "12" = "# ?/?", "46" = "[h]:mm:ss", "49" = "@")
    unname(builtin[id])
  }, NA_character_, USE.NAMES = FALSE)

  expect_equal(
    got,
    c(
      "0.00", "#,##0", "0.0%", "0.00E+00", "# ?/?", "yyyy-mm-dd", "dd.mm.yyyy hh:mm:ss",
      "[h]:mm:ss", "mmm d, yyyy", "#,##0.00;[Red]-#,##0.00", "0.0;[Red]-0.0;&quot;zero&quot;",
      "[&lt;100]0.0;[>=100]0;General", "0.0,,&quot; M&quot;", "[$\u20ac]#,##0.00", "@"
    )
  )
})

test_that("comments, hyperlinks, views, page setup and protection survive ods files", {
  wb <- wb_workbook(creator = "Jan", title = "ods title", subject = "ods subject", category = "cat")$
    add_worksheet("S 1")$
    add_data(x = head(mtcars[1:3], 3))$
    add_data(dims = "A6", x = "link")$
    add_hyperlink(dims = "A6", target = "https://example.com", tooltip = "tip")$
    add_comment(dims = "B2", comment = wb_comment("first\nsecond", author = "Jan"))$
    freeze_pane(first_active_row = 2, first_active_col = 2)$
    set_sheetview(zoom_scale = 150, show_grid_lines = FALSE)$
    page_setup(orientation = "landscape", paper_size = 9, left = 0.5, right = 0.5, top = 1, bottom = 1)$
    set_header_footer(header = c(NA, "Center &[Tab]", "&[Page] of &[Pages]"), footer = c("foot", NA, NA))$
    protect_worksheet(protect = TRUE)$
    add_worksheet("other")$
    add_data(x = 1)
  wb$protect(protect = TRUE, lock_structure = TRUE)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  wb2 <- load_ods(fl)
  ws <- wb2$worksheets[[1]]

  expect_equal(ws$freezePane, "<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>")
  expect_match(ws$sheetViews, "showGridLines=\"0\" zoomScale=\"150\"", fixed = TRUE)
  expect_match(ws$pageSetup, "paperSize=\"9\" orientation=\"landscape\"", fixed = TRUE)
  expect_equal(ws$pageMargins, "<pageMargins left=\"0.5\" right=\"0.5\" top=\"1\" bottom=\"1\" header=\"0.3\" footer=\"0.3\"/>")
  expect_equal(ws$headerFooter$oddHeader, c("", "Center &amp;A", "&amp;P of &amp;N"))
  expect_equal(ws$headerFooter$oddFooter, c("foot", "", ""))
  expect_equal(ws$sheetProtection, "<sheetProtection sheet=\"1\"/>")
  expect_equal(wb2$workbook$workbookProtection, "<workbookProtection lockStructure=\"1\"/>")

  expect_equal(ws$hyperlinks, "<hyperlink ref=\"A6\" r:id=\"rId1\" tooltip=\"tip\"/>")
  expect_match(paste(wb2$worksheets_rels[[1]], collapse = ""), "https://example.com", fixed = TRUE)

  comment <- wb2$comments[[1]][[1]]
  expect_equal(comment$ref, "B2")
  expect_equal(comment$author, "Jan")
  expect_equal(paste(comment$comment, collapse = ""), "Jan:\nfirst\nsecond")

  expect_match(wb2$core, "<dc:title>ods title</dc:title>", fixed = TRUE)
  expect_match(wb2$core, "<dc:creator>Jan</dc:creator>", fixed = TRUE)
  expect_match(wb2$core, "<cp:category>cat</cp:category>", fixed = TRUE)
})

test_that("images and charts survive ods files", {
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
    "<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Bars</a:t></a:r></a:p></c:rich></c:tx></c:title>",
    "<c:autoTitleDeleted val=\"0\"/><c:plotArea><c:layout/><c:barChart><c:barDir val=\"col\"/>",
    "<c:grouping val=\"stacked\"/><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:tx><c:strRef><c:f>data!$B$1</c:f></c:strRef></c:tx>",
    "<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>",
    "<c:cat><c:strRef><c:f>data!$A$2:$A$4</c:f></c:strRef></c:cat>",
    "<c:val><c:numRef><c:f>data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>",
    "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>",
    "<c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>",
    "<c:axPos val=\"l\"/><c:majorGridlines/><c:crossAx val=\"1\"/></c:valAx></c:plotArea>",
    "<c:legend><c:legendPos val=\"b\"/></c:legend><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
  )
  wb <- wb_workbook()$
    add_worksheet("data")$
    add_data(x = data.frame(k = c("a", "b", "c"), x = c(3, 1, 4)))$
    add_image(dims = "D2:F8", file = img)$
    add_chart_xml(dims = "H2:N15", xml = chart)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  files <- utils::unzip(fl, list = TRUE)$Name
  expect_true("Pictures/image1.jpg" %in% files)
  expect_true("Object 1/content.xml" %in% files)

  wb2 <- load_ods(fl)
  expect_length(wb2$media, 1)
  expect_match(wb2$drawings[[1]], "<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>", fixed = TRUE)
  expect_match(wb2$drawings[[1]], "<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row>", fixed = TRUE)
  expect_equal(nrow(wb2$charts), 1)
  xml <- wb2$charts$chart[1]
  expect_match(xml, "<c:barDir val=\"col\"/><c:grouping val=\"stacked\"/>", fixed = TRUE)
  expect_match(xml, "<c:f>data!$B$2:$B$4</c:f>", fixed = TRUE)
  expect_match(xml, "<c:f>data!$A$2:$A$4</c:f>", fixed = TRUE)
  expect_match(xml, "<a:t>Bars</a:t>", fixed = TRUE)
  expect_match(xml, "<a:srgbClr val=\"FF0000\"/>", fixed = TRUE)
  expect_match(xml, "<c:legendPos val=\"b\"/>", fixed = TRUE)
})

test_that("tables, filters, validations, conditional formats, outlines and shapes roundtrip", {
  wb <- wb_workbook()$
    add_worksheet("S 1")$
    add_data(x = head(mtcars[1:5], 12))$
    add_filter(rows = 1, cols = 1:5)$
    group_rows(rows = 3:6, collapsed = TRUE)$
    group_cols(cols = 3:4)$
    add_data_validation(dims = "G2:G5", type = "whole", operator = "between", value = c(1, 10))$
    add_data_validation(dims = "H2:H5", type = "list", value = '"a,b,c"')$
    add_data_validation(dims = "J2:J5", type = "textLength", operator = "lessThanOrEqual", value = 3)$
    add_data_validation(dims = "K2:K5", type = "custom", value = "A2>3")$
    add_dxfs_style(name = "high", font_color = wb_color("FF0000"), bg_fill = wb_color("FFFF00"))$
    add_conditional_formatting(dims = "A2:A13", rule = ">20", type = "expression", style = "high")$
    add_conditional_formatting(dims = "B2:B13", type = "colorScale", style = c("red", "green"), rule = c(4, 8))$
    add_conditional_formatting(dims = "C2:C13", type = "dataBar", style = "blue")$
    add_conditional_formatting(dims = "E2:E13", type = "topN", params = list(rank = 3))$
    add_worksheet("tab")$
    add_data_table(x = head(iris), table_name = "irisTab")$
    add_drawing(dims = "H2:K8", xml = create_shape(shape = "ellipse", text = "oval", fill_color = wb_color("FF0000")))

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "table:condition=\"of:cell-content-text-length()&lt;=3\"", fixed = TRUE)
  expect_match(content, "<calcext:data-bar", fixed = TRUE)
  expect_match(content, "draw:type=\"ooxml-ellipse\"", fixed = TRUE)
  expect_match(content, "<table:table-row-group", fixed = TRUE)

  wb2 <- load_ods(fl)
  ws <- wb2$worksheets[[1]]
  expect_equal(ws$autoFilter, "<autoFilter ref=\"A1:E1\"/>")
  dv <- ws$dataValidations
  expect_length(dv, 4)
  expect_match(dv[1], "type=\"whole\" operator=\"between\"", fixed = TRUE)
  expect_match(dv[1], "<formula1>1</formula1><formula2>10</formula2>", fixed = TRUE)
  expect_match(dv[2], "<formula1>\"a,b,c\"</formula1>", fixed = TRUE)
  expect_match(dv[3], "type=\"textLength\" operator=\"lessThanOrEqual\"", fixed = TRUE)
  expect_match(dv[4], "<formula1>A2&gt;3</formula1>", fixed = TRUE)

  cf <- ws$conditionalFormatting
  expect_equal(cf$sqref, c("A2:A13", "B2:B13", "C2:C13", "E2:E13"))
  expect_match(cf$cf[1], "<formula>A2&gt;20</formula>", fixed = TRUE)
  expect_match(cf$cf[2], "<cfvo type=\"num\" val=\"4\"/><cfvo type=\"num\" val=\"8\"/>", fixed = TRUE)
  expect_match(cf$cf[3], "<dataBar>", fixed = TRUE)
  expect_match(cf$cf[4], "type=\"top10\"", fixed = TRUE)
  expect_match(cf$cf[4], "rank=\"3\"", fixed = TRUE)
  expect_match(wb2$styles_mgr$styles$dxfs[1], "<bgColor rgb=\"FFFFFF00\"/>", fixed = TRUE)

  ra <- ws$sheet_data$row_attr
  expect_equal(ra$outlineLevel[ra$r %in% 3:6], c("1", "1", "1", ""))
  expect_equal(ra$hidden[ra$r %in% 3:6], c("1", "1", "1", ""))
  expect_equal(grep("outlineLevel=\"1\"", ws$cols_attr), grep("min=\"3\"", ws$cols_attr))

  expect_equal(wb2$tables$tab_ref, "A1:E7")
  expect_equal(wb_to_df(wb2, sheet = "tab")$Species[1], "setosa")
  drawing <- wb2$drawings[[length(wb2$drawings)]]
  expect_match(drawing, "<a:prstGeom prst=\"ellipse\">", fixed = TRUE)
  expect_match(drawing, "<a:t>oval</a:t>", fixed = TRUE)
  expect_match(drawing, "FF0000", fixed = TRUE)
})

test_that("pivot tables roundtrip", {
  df <- mtcars[1:20, c("cyl", "gear", "am", "mpg", "hp")]
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = df)$add_worksheet("pv")
  x <- wb_data(wb, sheet = "data")
  wb$add_pivot_table(
    x, sheet = "pv", dims = "A3", rows = "cyl", cols = "gear", data = c("mpg", "hp"),
    fun = c("max", "stdDev"), filter = "am",
    params = list(name = "pt1", col_grand_totals = FALSE, choose = c(cyl = "x != 4"))
  )
  wb$add_pivot_table(x, sheet = "pv", dims = "J3", rows = c("cyl", "gear"), data = "mpg", fun = "count")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "table:target-range-address=\"'pv'.A1:'pv'.A3\" table:grand-total=\"row\"", fixed = TRUE)
  expect_match(content, "<table:data-pilot-member table:name=\"4\" table:display=\"false\"/>", fixed = TRUE)
  expect_match(content, "table:source-field-name=\"hp\" table:orientation=\"data\" table:function=\"stdev\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  expect_length(wb2$pivotTables, 2)
  for (k in 1:2) {
    a <- wb$pivotTables[k]
    b <- wb2$pivotTables[k]
    for (node in c("location", "rowFields", "colFields", "pageFields", "dataFields")) {
      expect_equal(xml_node(b, "pivotTableDefinition", node), xml_node(a, "pivotTableDefinition", node))
    }
  }
  expect_match(wb2$pivotTables[1], "colGrandTotals=\"0\"", fixed = TRUE)
  expect_match(wb2$pivotTables[1], "<item x=\"1\" h=\"1\"/>", fixed = TRUE)
})

test_that("rich text in cells and comments roundtrips", {
  txt <- fmt_txt("bold ", bold = TRUE) + fmt_txt("red", color = wb_color("FFFF0000")) +
    fmt_txt(" big\nline2", size = 16, font = "Courier New")
  comment <- wb_comment(text = fmt_txt("c bold", bold = TRUE) + fmt_txt(" c it", italic = TRUE), author = "me")
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = txt, dims = "A1")$
    add_data(x = "plain  two", dims = "A2")$
    add_comment(dims = "B2", comment = comment)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "<style:font-face style:name=\"Courier New\" svg:font-family=\"'Courier New'\"/>", fixed = TRUE)
  expect_match(content, "<style:text-properties style:font-name=\"Courier New\" fo:font-size=\"16pt\"/>", fixed = TRUE)
  expect_match(content, "<text:span text:style-name=\"T1\">bold<text:s/></text:span>", fixed = TRUE)

  wb2 <- load_ods(fl)
  expect_equal(wb_to_df(wb2, col_names = FALSE, dims = "A1:A2")$A, c("bold red big\nline2", "plain  two"))
  is <- wb2$worksheets[[1]]$sheet_data$cc$is
  expect_match(is[1], "<r><rPr><b/>", fixed = TRUE)
  expect_match(is[1], "<color rgb=\"FFFF0000\"/>", fixed = TRUE)
  expect_match(is[1], "<sz val=\"16\"/><color rgb=\"FF000000\"/><rFont val=\"Courier New\"/></rPr><t xml:space=\"preserve\"> big\nline2</t>", fixed = TRUE)
  expect_equal(is[2], "<is><t xml:space=\"preserve\">plain  two</t></is>")
  text <- wb2$comments[[1]][[1]]$comment
  expect_match(text[3], "<b/>", fixed = TRUE)
  expect_match(text[3], "<i/>", fixed = TRUE)
  expect_match(text[3], " c it</t>", fixed = TRUE)
})

test_that("print areas and print titles roundtrip", {
  wb <- wb_workbook()$add_worksheet("S 1")$add_data(x = mtcars)$
    page_setup(print_title_rows = 1:2, print_title_cols = 1)$
    add_worksheet("b")$add_data(x = 1:3)
  wb$add_named_region(sheet = 1, dims = "A1:D10", name = "_xlnm.Print_Area", local_sheet = TRUE)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "table:print-ranges=\"'S 1'.A1:'S 1'.D10\"", fixed = TRUE)
  expect_match(content, "<table:table-header-columns><table:table-column table:style-name=\"co1\"/></table:table-header-columns>", fixed = TRUE)
  expect_match(content, "<table:table-header-rows><table:table-row table:style-name=\"ro1\">", fixed = TRUE)

  wb2 <- load_ods(fl)
  names <- wb2$workbook$definedNames
  expect_true(any(grepl(">'S 1'!$A$1:$D$10</definedName>", names, fixed = TRUE)))
  expect_true(any(grepl(">'S 1'!$A:$A,'S 1'!$1:$2</definedName>", names, fixed = TRUE)))
})

test_that("bubble, stock and combo charts with secondary axes and trendlines roundtrip", {
  ns <- paste(
    "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
  )
  axes <- function(id1, id2, cat = TRUE, pos2 = "l", scaling = "", crosses = "autoZero") {
    first <- if (cat) "c:catAx" else "c:valAx"
    paste0(
      "<", first, "><c:axId val=\"", id1, "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>",
      "<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"", id2, "\"/></", first, ">",
      "<c:valAx><c:axId val=\"", id2, "\"/><c:scaling><c:orientation val=\"minMax\"/>", scaling, "</c:scaling>",
      "<c:delete val=\"0\"/><c:axPos val=\"", pos2, "\"/><c:majorGridlines/><c:crossAx val=\"", id1, "\"/>",
      "<c:crosses val=\"", crosses, "\"/></c:valAx>"
    )
  }
  ser <- function(i, col, extra = "", inner = "") {
    paste0(
      "<c:ser><c:idx val=\"", i, "\"/><c:order val=\"", i, "\"/><c:tx><c:strRef><c:f>data!$", col, "$1</c:f></c:strRef></c:tx>",
      extra, "<c:cat><c:numRef><c:f>data!$A$2:$A$6</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>data!$", col, "$2:$", col,
      "$6</c:f></c:numRef></c:val>", inner, "</c:ser>"
    )
  }
  chart <- function(x) {
    paste0(
      "<c:chartSpace ", ns, "><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>", x,
      "</c:plotArea><c:legend><c:legendPos val=\"r\"/></c:legend><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
    )
  }
  bubble <- chart(paste0(
    "<c:bubbleChart><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:xVal><c:numRef><c:f>data!$A$2:$A$6</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>data!$B$2:$B$6</c:f></c:numRef></c:yVal>",
    "<c:bubbleSize><c:numRef><c:f>data!$C$2:$C$6</c:f></c:numRef></c:bubbleSize></c:ser>",
    "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:bubbleChart>", axes(1, 2, cat = FALSE)
  ))
  stock <- chart(paste0(
    "<c:stockChart>", ser(0, "B"), ser(1, "C"), ser(2, "D"), ser(3, "E"),
    "<c:hiLowLines/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:stockChart>", axes(1, 2)
  ))
  combo <- chart(paste0(
    "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>",
    ser(
      0, "B", "<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>",
      "<c:trendline><c:trendlineType val=\"poly\"/><c:order val=\"2\"/><c:dispRSqr val=\"1\"/><c:dispEq val=\"1\"/></c:trendline>"
    ),
    "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>",
    "<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls><c:gapWidth val=\"80\"/>",
    "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>", ser(1, "C", "", "<c:smooth val=\"1\"/>"),
    "<c:marker val=\"1\"/><c:axId val=\"3\"/><c:axId val=\"4\"/></c:lineChart>",
    axes(1, 2, scaling = "<c:max val=\"10\"/><c:min val=\"0\"/>"), axes(3, 4, pos2 = "r", crosses = "max")
  ))
  df <- data.frame(x = 1:5, a = c(3, 5, 2, 6, 4), b = c(2, 4, 3, 5, 1), c = c(1, 2, 1, 2, 1), d = c(3, 4, 2, 5, 3))
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = df)$
    add_chart_xml(dims = "G1:M10", xml = bubble)$
    add_chart_xml(dims = "G12:M22", xml = stock)$
    add_chart_xml(dims = "G24:M34", xml = combo)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, exdir = exdir)
  object <- function(k) paste(readLines(file.path(exdir, paste("Object", k), "content.xml"), warn = FALSE), collapse = "")
  expect_match(object(1), "chart:values-cell-range-address=\"data.$C$2:data.$C$6\" chart:class=\"chart:bubble\"", fixed = TRUE)
  expect_match(object(1), "\"><chart:domain table:cell-range-address=\"data.$B$2:data.$B$6\"/><chart:domain table:cell-range-address=\"data.$A$2:data.$A$6\"/>", fixed = TRUE)
  expect_match(object(2), "chart:japanese-candle-stick=\"true\"", fixed = TRUE)
  expect_match(object(2), "<chart:stock-range-line/>", fixed = TRUE)
  expect_match(object(3), "chart:attached-axis=\"secondary-y\"", fixed = TRUE)
  expect_match(object(3), "chart:regression-type=\"polynomial\" chart:regression-max-degree=\"2\"", fixed = TRUE)
  expect_match(object(3), "<chart:equation chart:display-equation=\"true\" chart:display-r-square=\"true\"/>", fixed = TRUE)
  expect_match(object(3), "chart:data-label-number=\"value\"", fixed = TRUE)
  expect_match(object(3), "chart:minimum=\"0\" chart:maximum=\"10\"", fixed = TRUE)
  expect_match(object(3), "chart:interpolation=\"cubic-spline\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  charts <- wb2$charts$chart
  expect_length(charts, 3)
  expect_match(charts[1], "<c:xVal><c:numRef><c:f>data!$A$2:$A$6</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>data!$B$2:$B$6</c:f></c:numRef></c:yVal><c:bubbleSize><c:numRef><c:f>data!$C$2:$C$6</c:f>", fixed = TRUE)
  values <- regmatches(charts[2], gregexpr("<c:val><c:numRef><c:f>[^<]*", charts[2]))[[1]]
  expect_equal(sub(".*<c:f>", "", values), c("data!$B$2:$B$6", "data!$C$2:$C$6", "data!$D$2:$D$6", "data!$E$2:$E$6"))
  expect_match(charts[2], "<c:hiLowLines/>", fixed = TRUE)
  expect_match(charts[3], "<c:trendline><c:trendlineType val=\"poly\"/><c:order val=\"2\"/><c:dispRSqr val=\"1\"/><c:dispEq val=\"1\"/></c:trendline>", fixed = TRUE)
  expect_match(charts[3], "<c:showVal val=\"1\"/>", fixed = TRUE)
  expect_match(charts[3], "<c:max val=\"10\"/><c:min val=\"0\"/>", fixed = TRUE)
  expect_match(charts[3], "<c:axId val=\"3\"/><c:axId val=\"4\"/></c:lineChart>", fixed = TRUE)
  expect_match(charts[3], "<c:axPos val=\"r\"/><c:crossAx val=\"3\"/><c:crosses val=\"max\"/>", fixed = TRUE)
  expect_match(charts[3], "<c:smooth val=\"1\"/>", fixed = TRUE)
})

test_that("chartsheets are written as sheets with their chart", {
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><c:chart><c:plotArea>",
    "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:val><c:numRef><c:f>data!$A$1:$A$3</c:f></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>",
    "</c:plotArea></c:chart></c:chartSpace>"
  )
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = 1:3, col_names = FALSE)$
    add_chartsheet(sheet = "chart")$add_chart_xml(xml = chart)

  fl <- tempfile(fileext = ".ods")
  expect_silent(write_ods(wb, fl))
  wb2 <- load_ods(fl)
  expect_equal(unname(wb2$get_sheet_names()), c("data", "chart"))
  expect_equal(nrow(wb2$charts), 1)
  expect_match(wb2$charts$chart, "<c:f>data!$A$1:$A$3</c:f>", fixed = TRUE)
})

test_that("shape text, alignment, rotation and groups roundtrip", {
  text <- fmt_txt("bold", bold = TRUE) + fmt_txt(" red", color = wb_color("FF0000"))
  wb <- wb_workbook()$add_worksheet()$
    add_drawing(dims = "B2:E6", xml = create_shape(shape = "rect", text = text, text_align = "center", rotation = 30))

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "draw:transform=\"rotate (-0.5235987756) translate (", fixed = TRUE)
  expect_match(content, "<text:p text:style-name=\"P1\"><text:span text:style-name=\"T1\">bold</text:span>", fixed = TRUE)
  expect_match(content, "<style:paragraph-properties fo:text-align=\"center\"/>", fixed = TRUE)

  wb2 <- load_ods(fl)
  drawing <- wb2$drawings[[1]]
  expect_match(drawing, "<a:xfrm rot=\"1800000\">", fixed = TRUE)
  expect_match(drawing, "<a:pPr algn=\"ctr\"/>", fixed = TRUE)
  expect_match(drawing, "<a:rPr b=\"1\"", fixed = TRUE)
  expect_match(drawing, "<a:srgbClr val=\"FF0000\">", fixed = TRUE)
  expect_match(drawing, "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>", fixed = TRUE)
  expect_match(drawing, "<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row>", fixed = TRUE)

  sp <- function(id, x, prst) {
    paste0(
      "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"", id, "\" name=\"S", id, "\"/><xdr:cNvSpPr/></xdr:nvSpPr>",
      "<xdr:spPr><a:xfrm><a:off x=\"", x, "\" y=\"0\"/><a:ext cx=\"1800000\" cy=\"1000000\"/></a:xfrm>",
      "<a:prstGeom prst=\"", prst, "\"><a:avLst/></a:prstGeom><a:ln w=\"38100\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>"
    )
  }
  group <- paste0(
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><xdr:twoCellAnchor>",
    "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
    "<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
    "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"2\" name=\"Group 1\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm>",
    "<a:off x=\"0\" y=\"0\"/><a:ext cx=\"4000000\" cy=\"2000000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"4000000\" cy=\"2000000\"/>",
    "</a:xfrm></xdr:grpSpPr>", sp(3, 0, "rect"), sp(4, 2200000, "ellipse"),
    "</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"
  )
  wb <- wb_workbook()$add_worksheet()$add_drawing(dims = "B2:F10", xml = group)
  write_ods(wb, fl)
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "<table:shapes><draw:g draw:z-index=\"0\"><draw:custom-shape draw:z-index=\"0\" draw:name=\"S3\"", fixed = TRUE)
  expect_match(content, "svg:stroke-width=\"3pt\"", fixed = TRUE)
  drawing <- load_ods(fl)$drawings[[1]]
  expect_match(drawing, "prst=\"rect\"", fixed = TRUE)
  expect_match(drawing, "prst=\"ellipse\"", fixed = TRUE)
})

test_that("pivot table sorting and shown values roundtrip", {
  df <- mtcars[1:20, c("cyl", "gear", "am", "mpg", "hp")]
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = df)$add_worksheet("pv")
  x <- wb_data(wb, sheet = "data")
  wb$add_pivot_table(
    x, sheet = "pv", dims = "A3", rows = "cyl", cols = "gear", data = c("mpg", "hp"), fun = c("sum", "average"),
    params = list(show_data_as = c("percentOfTotal", "normal"), sort_row = -1)
  )
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "<table:data-pilot-sort-info table:sort-mode=\"name\" table:order=\"descending\"/>", fixed = TRUE)
  expect_match(content, "table:type=\"total-percentage\"/>", fixed = TRUE)

  pt <- load_ods(fl)$pivotTables
  expect_match(pt, "<pivotField axis=\"axisRow\" showAll=\"0\" sortType=\"descending\">", fixed = TRUE)
  expect_match(pt, "fld=\"3\" showDataAs=\"percentOfTotal\"", fixed = TRUE)
})

test_that("form controls roundtrip", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = c("a", "b", "c"), dims = "D1:D3", col_names = FALSE)$
    add_form_control(dims = "B2", type = "Checkbox", text = "check me", link = "A1", checked = TRUE)$
    add_form_control(dims = "B4", type = "Drop", range = "D1:D3", link = "A2")$
    add_form_control(dims = "B6", type = "Radio", text = "r1")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "<form:checkbox form:name=\"Checkbox 1\"", fixed = TRUE)
  expect_match(content, "form:label=\"check me\" form:linked-cell=\"'Sheet 1'.A1\"", fixed = TRUE)
  expect_match(content, "form:source-cell-range=\"'Sheet 1'.D1:'Sheet 1'.D3\"", fixed = TRUE)
  expect_match(content, "<draw:control svg:width=", fixed = TRUE)

  wb2 <- load_ods(fl)
  expect_length(wb2$ctrlProps, 3)
  vml <- wb2$vml[[1]]
  anchors <- regmatches(vml, gregexpr("<x:Anchor>[^<]*</x:Anchor>", vml))[[1]]
  expect_equal(anchors, c(
    "<x:Anchor>1, 0, 1, 0, 2, 0, 2, 0</x:Anchor>", "<x:Anchor>1, 0, 3, 0, 2, 0, 4, 0</x:Anchor>",
    "<x:Anchor>1, 0, 5, 0, 2, 0, 6, 0</x:Anchor>"
  ))
  expect_match(vml, "<x:FmlaRange>'Sheet 1'!$D$1:$D$3</x:FmlaRange>", fixed = TRUE)
  expect_match(vml, "check me", fixed = TRUE)
})

test_that("chart ranges with several areas roundtrip", {
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><c:chart><c:plotArea>",
    "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:val><c:numRef><c:f>(data!$A$1:$A$2,data!$A$4:$A$5)</c:f></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>",
    "</c:plotArea></c:chart></c:chartSpace>"
  )
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = 1:5, col_names = FALSE)$add_chart_xml(dims = "C1:H10", xml = chart)
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "Object 1/content.xml", exdir = exdir)
  object <- paste(readLines(file.path(exdir, "Object 1", "content.xml"), warn = FALSE), collapse = "")
  expect_match(object, "chart:values-cell-range-address=\"data.$A$1:data.$A$2 data.$A$4:data.$A$5\"", fixed = TRUE)
  expect_match(load_ods(fl)$charts$chart, "<c:f>(data!$A$1:$A$2,data!$A$4:$A$5)</c:f>", fixed = TRUE)
})

test_that("cropped and rotated pictures and gradients roundtrip", {
  img <- tempfile(fileext = ".png")
  grDevices::png(img, width = 200, height = 100)
  graphics::par(mar = c(0, 0, 0, 0))
  graphics::plot.new()
  grDevices::dev.off()
  wb <- wb_workbook()$add_worksheet()$add_image(dims = "B2:E8", file = img)
  drawing <- wb$drawings[[1]]
  drawing <- sub("r:embed=\"rId1\"/>", "r:embed=\"rId1\"/><a:srcRect l=\"10000\" t=\"20000\" r=\"30000\" b=\"5000\"/>", drawing, fixed = TRUE)
  drawing <- sub(
    "<xdr:spPr><a:prstGeom", "<xdr:spPr><a:xfrm rot=\"1200000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom",
    drawing, fixed = TRUE
  )
  wb$drawings[[1]] <- drawing
  shape <- create_shape(shape = "rect", fill_color = wb_color("FF0000"))
  shape <- sub(
    "<a:solidFill><a:srgbClr val=\"FF0000\"><a:alpha val=\"99000\"/></a:srgbClr></a:solidFill><a:ln>",
    paste0(
      "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs>",
      "</a:gsLst><a:lin ang=\"0\" scaled=\"0\"/></a:gradFill><a:ln>"
    ),
    shape, fixed = TRUE
  )
  wb$add_drawing(dims = "G2:J8", xml = shape)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "styles.xml"), exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  styles <- paste(readLines(file.path(exdir, "styles.xml"), warn = FALSE), collapse = "")
  expect_match(content, "fo:clip=\"rect(0.529167cm, 1.587500cm, 0.132292cm, 0.529167cm)\"", fixed = TRUE)
  expect_match(content, "draw:transform=\"rotate (-0.3490658504)", fixed = TRUE)
  expect_match(content, "draw:fill=\"gradient\" draw:fill-gradient-name=\"Gradient_1\"", fixed = TRUE)
  expect_match(styles, "<draw:gradient draw:name=\"Gradient_1\" draw:style=\"linear\" draw:start-color=\"#ff0000\" draw:end-color=\"#0000ff\" draw:angle=\"90deg\"", fixed = TRUE)

  drawing <- load_ods(fl)$drawings[[1]]
  expect_match(drawing, "<a:srcRect l=\"10000\" t=\"20000\" r=\"30000\" b=\"5000\"/>", fixed = TRUE)
  expect_match(drawing, "<a:xfrm rot=\"1200000\">", fixed = TRUE)
  expect_match(drawing, "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst><a:lin ang=\"0\" scaled=\"0\"/></a:gradFill>", fixed = TRUE)
})

test_that("conditional formats are written as style maps and loaded from them", {
  wb <- wb_workbook()$add_worksheet("S 1")$add_data(x = head(mtcars[1:4], 6))$
    add_fill(dims = "A2", color = wb_color("FFFFFF00"))$
    add_dxfs_style(name = "high", font_color = wb_color("FF0000"))$
    add_conditional_formatting(dims = "A2:A7", rule = ">20", style = "high")$
    add_conditional_formatting(dims = "B2:B7", type = "between", rule = c(4, 6), style = "high")$
    add_conditional_formatting(dims = "C2:C7", type = "expression", rule = "C2>D2", style = "high")$
    add_worksheet("S 2")$add_data(x = 1:3)$
    add_conditional_formatting(dims = "A2:A4", rule = ">1", style = "high")
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "<style:map style:condition=\"is-true-formula([.A2]>20)\" style:apply-style-name=\"ConditionalStyle_1\" style:base-cell-address=\"'S 1'.A2\"/>", fixed = TRUE)
  expect_match(content, "style:condition=\"cell-content-is-between(4,6)\"", fixed = TRUE)
  names <- regmatches(content, gregexpr("<style:style style:name=\"[^\"]*\"", content))[[1]]
  expect_false(anyDuplicated(names) > 0)
  # a cell with its own style gets a copy of the style with the condition
  expect_match(content, "fo:color=\"#000000\"/><style:map style:condition=\"is-true-formula([.A2]>20)\"", fixed = TRUE)
  expect_match(content, "style:condition=\"is-true-formula([.C2]>[.D2])\"", fixed = TRUE)

  # without the formats of the calc extension the style maps are loaded
  content <- gsub("<calcext:conditional-formats>.*?</calcext:conditional-formats>", "", content, perl = TRUE)
  writeLines(content, file.path(exdir, "content.xml"))
  plain <- tempfile(fileext = ".ods")
  old <- setwd(exdir)
  on.exit(setwd(old), add = TRUE)
  utils::zip(plain, "mimetype", flags = "-0Xq")
  utils::zip(plain, setdiff(list.files(".", recursive = TRUE), "mimetype"), flags = "-r9Xq")
  setwd(old)
  wb2 <- load_ods(plain)
  # only the cells with content carry the conditions
  expect_equal(wb2$worksheets[[2]]$conditionalFormatting$sqref, "A2:A3")
  cf <- wb2$worksheets[[1]]$conditionalFormatting
  expect_setequal(cf$sqref, c("A2:A7", "B2:B7", "C2:C7"))
  expect_true(any(grepl("<cfRule type=\"expression\" dxfId=\"0\" priority=\"[0-9]+\"><formula>A2&gt;20</formula></cfRule>", cf$cf)))
  expect_true(any(grepl("operator=\"between\".*<formula>4</formula><formula>6</formula>", cf$cf)))
  expect_true(any(grepl("<formula>C2&gt;D2</formula>", cf$cf, fixed = TRUE)))
})

test_that("the background picture of a sheet roundtrips", {
  img <- tempfile(fileext = ".png")
  grDevices::png(img, width = 20, height = 10)
  graphics::par(mar = c(0, 0, 0, 0))
  graphics::plot.new()
  grDevices::dev.off()
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1:3)$add_image(dims = "C2:D4", file = img)
  # the image becomes the background picture of the sheet
  wb$worksheets[[1]]$picture <- "<picture r:id=\"rId9\"/>"
  wb$worksheets_rels[[1]] <- c(
    wb$worksheets_rels[[1]],
    "<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/>"
  )
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "<style:table-properties table:display=\"true\"><style:background-image xlink:href=\"Pictures/image1.png\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  expect_equal(wb2$worksheets[[1]]$picture, "<picture r:id=\"rId1\"/>")
  expect_match(wb2$worksheets_rels[[1]][1], "Target=\"../media/image1.png\"", fixed = TRUE)
  expect_length(wb2$media, 2)
})

test_that("sheets are selected and data is loaded without the rest", {
  wb <- wb_workbook()$add_worksheet("a")$add_data(x = head(mtcars, 3))$
    add_comment(dims = "A1", comment = wb_comment(text = "note", author = "me"))$
    add_worksheet("b")$add_data(x = head(iris, 3))
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)

  wb2 <- load_ods(fl, sheet = "b")
  expect_equal(unname(wb2$get_sheet_names()), c("a", "b"))
  expect_null(wb2$worksheets[[1]]$sheet_data$cc)
  expect_equal(wb_to_df(wb2, sheet = "b")$Species, rep("setosa", 3))

  wb3 <- load_ods(fl, sheet = 1, data_only = TRUE)
  expect_length(wb3$comments, 0)
  expect_equal(wb_to_df(wb3, sheet = 1)$mpg, head(mtcars, 3)$mpg)

  expect_equal(read_ods(fl, sheet = "b"), wb_to_df(wb, sheet = "b"), ignore_attr = TRUE)
  expect_equal(read_ods(fl)$cyl, head(mtcars, 3)$cyl)
  expect_warning(ods_open(fl, interactive = FALSE), "will not open file")
  expect_warning(ods_open(wb, interactive = FALSE), "will not open file")
})

test_that("the ODF version is written and formulas of ODF 1.1 are loaded", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1:3)$add_formula(x = "SUM(A1:A3)", dims = "B1")
  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "office:version=\"1.4\"", fixed = TRUE)
  write_ods(wb, fl, version = "1.3")
  utils::unzip(fl, exdir = exdir, overwrite = TRUE)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  manifest <- paste(readLines(file.path(exdir, "META-INF", "manifest.xml"), warn = FALSE), collapse = "")
  expect_match(content, "office:version=\"1.3\"", fixed = TRUE)
  expect_match(manifest, "manifest:version=\"1.3\"", fixed = TRUE)
  expect_error(write_ods(wb, fl, version = "1.2"))

  # ODF 1.1 formulas have the oooc: prefix
  writeLines(gsub("of:=", "oooc:=", content, fixed = TRUE), file.path(exdir, "content.xml"))
  old <- setwd(exdir)
  on.exit(setwd(old), add = TRUE)
  odf11 <- tempfile(fileext = ".ods")
  utils::zip(odf11, "mimetype", flags = "-0Xq")
  utils::zip(odf11, setdiff(list.files(".", recursive = TRUE), "mimetype"), flags = "-r9Xq")
  setwd(old)
  expect_silent(wb2 <- load_ods(odf11))
  expect_equal(wb2$worksheets[[1]]$sheet_data$cc$f[wb2$worksheets[[1]]$sheet_data$cc$r == "B1"], "SUM(A1:A3)")
})

test_that("frames, charts, pivot tables and shapes carry what other consumers need", {
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
    "<c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>",
    "<c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:tx><c:strRef><c:f>data!$B$1</c:f></c:strRef></c:tx>",
    "<c:cat><c:strRef><c:f>data!$A$2:$A$4</c:f></c:strRef></c:cat>",
    "<c:val><c:numRef><c:f>data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>",
    "<c:ser><c:idx val=\"1\"/><c:order val=\"1\"/><c:tx><c:v>fixed</c:v></c:tx>",
    "<c:spPr><a:solidFill><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:solidFill></c:spPr>",
    "<c:val><c:numLit><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>7</c:v></c:pt><c:pt idx=\"2\"><c:v>9</c:v></c:pt>",
    "</c:numLit></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>",
    "<c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>",
    "<c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"
  )
  df <- mtcars[1:20, c("cyl", "gear", "am", "mpg")]
  wb <- wb_workbook()$
    add_worksheet("data")$
    add_data(x = data.frame(k = c("a", "b", "c"), x = c(3, 1, 4)))$
    add_chart_xml(dims = "D2:H10", xml = chart)$
    add_drawing(dims = "J2:L8", xml = create_shape(shape = "smileyFace"))$
    add_worksheet("pivot")$
    add_data(x = df)$
    add_worksheet("pv")
  wb$add_pivot_table(wb_data(wb, sheet = "pivot"), sheet = "pv", dims = "A3", rows = "cyl", cols = "gear", data = "mpg", filter = "am")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  expect_true("Object 1/styles.xml" %in% utils::unzip(fl, list = TRUE)$Name)
  exdir <- tempfile()
  utils::unzip(fl, exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  object <- paste(readLines(file.path(exdir, "Object 1", "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  expect_match(content, "<draw:frame draw:z-index=\"0\" draw:name=\"[^\"]*\"[^>]*draw:style-name=\"gr[0-9]+\"")
  expect_match(content, "draw:notify-on-update-of-ranges=\"data.$A$2:data.$A$4 data.$B$1 data.$B$2:data.$B$4\"", fixed = TRUE)
  expect_match(content, "table:buttons=\"'pv'.A1 'pv'.A4 'pv'.B3\" table:show-filter-button=\"false\"", fixed = TRUE)
  expect_match(content, "draw:type=\"ooxml-smileyFace\" draw:enhanced-path=\"M 10800 0 X 21600 10800", fixed = TRUE)
  expect_match(object, "<table:table table:name=\"local-table\">", fixed = TRUE)
  expect_match(object, "draw:fill-color=\"#0c445e\"", fixed = TRUE)
  expect_match(object, "<table:table-cell office:value-type=\"string\"><text:p>b</text:p></table:table-cell>", fixed = TRUE)
  expect_match(object, "chart:values-cell-range-address=\"local-table.$C$2:local-table.$C$4\" chart:label-cell-address=\"local-table.$C$1\"",
               fixed = TRUE)

  wb2 <- load_ods(fl)
  xml <- wb2$charts$chart[1]
  expect_match(xml, "<c:f>data!$B$2:$B$4</c:f>", fixed = TRUE)
  expect_match(xml, "<c:tx><c:v>fixed</c:v></c:tx>", fixed = TRUE)
  expect_match(xml, "<c:numLit><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>7</c:v></c:pt><c:pt idx=\"2\"><c:v>9</c:v></c:pt></c:numLit>",
               fixed = TRUE)
  expect_match(wb2$drawings[[1]], "prst=\"smileyFace\"", fixed = TRUE)
})

test_that("column widths, default sizes, outline groups and row styles roundtrip", {
  wb <- wb_workbook()$add_worksheet("s")$add_data(x = matrix(1:40, 10, 4), col_names = FALSE)
  wb$add_fill(dims = "A8", color = wb_color("yellow"))
  ra <- wb$worksheets[[1]]$sheet_data$row_attr
  ra$outlineLevel[ra$r %in% c("3", "4", "5")] <- "1"
  ra$hidden[ra$r %in% c("3", "4", "5")] <- "1"
  ra$collapsed[ra$r == "6"] <- "1"
  ra$s[ra$r == "8"] <- wb$worksheets[[1]]$sheet_data$cc$c_s[wb$worksheets[[1]]$sheet_data$cc$r == "A8"]
  ra$customFormat[ra$r == "8"] <- "1"
  wb$worksheets[[1]]$sheet_data$row_attr <- ra
  wb$worksheets[[1]]$cols_attr <- c(
    "<col min=\"2\" max=\"3\" width=\"12\" customWidth=\"1\" outlineLevel=\"1\" hidden=\"1\"/>",
    "<col min=\"4\" max=\"4\" width=\"9.140625\" collapsed=\"1\"/>"
  )
  wb$worksheets[[1]]$sheetFormatPr <- "<sheetFormatPr baseColWidth=\"8\" defaultRowHeight=\"15\" outlineLevelRow=\"1\" outlineLevelCol=\"1\"/>"

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  # a column unit is the digit width of the default font, 7 pixels at 96 dpi
  expect_match(content, "style:column-width=\"0.6667in\"", fixed = TRUE)
  expect_match(content, "style:column-width=\"0.8750in\"", fixed = TRUE)
  expect_match(content, "<table:table-column-group table:display=\"false\">", fixed = TRUE)
  expect_match(content, "<table:table-row-group table:display=\"false\">", fixed = TRUE)
  expect_match(content, "table:number-columns-repeated=\"16380\"/>", fixed = TRUE)
  expect_match(content, "table:number-rows-repeated=\"1048566\">", fixed = TRUE)
  expect_match(content, "<table:table-cell table:style-name=\"ce[0-9]+\" table:number-columns-repeated=\"16380\"/>")

  wb2 <- load_ods(fl)
  ws <- wb2$worksheets[[1]]
  expect_match(ws$sheetFormatPr, "defaultColWidth=\"9.143\" defaultRowHeight=\"15\"", fixed = TRUE)
  expect_equal(ws$cols_attr, c(
    "<col min=\"2\" max=\"3\" width=\"12\" customWidth=\"1\" outlineLevel=\"1\" hidden=\"1\"/>",
    "<col min=\"4\" max=\"4\" width=\"9.143\" collapsed=\"1\"/>"
  ))
  ra2 <- ws$sheet_data$row_attr
  expect_equal(ra2$hidden[ra2$r %in% c("3", "4", "5")], c("1", "1", "1"))
  expect_equal(ra2$collapsed[ra2$r == "6"], "1")
  expect_equal(ra2$customFormat[ra2$r == "8"], "1")
})

test_that("views, tab colors and the direction of sheets roundtrip", {
  wb <- wb_workbook()$
    add_worksheet("a", tab_color = wb_color("FF0000"))$
    add_data(x = 1:3)$
    set_sheetview(zoom_scale = 80, top_left_cell = "C10", show_zeros = FALSE, right_to_left = TRUE)$
    add_worksheet("b")$
    add_data(x = 1:3)$
    set_sheetview(view = "pageBreakPreview", zoom_scale = 70, zoom_scale_normal = 120)
  wb$set_active_sheet(2)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "settings.xml"), exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  settings <- paste(readLines(file.path(exdir, "settings.xml"), warn = FALSE), collapse = "")
  expect_match(content, "style:writing-mode=\"rl-tb\" table:tab-color=\"#ff0000\"", fixed = TRUE)
  expect_match(settings, "<config:config-item config:name=\"PositionLeft\" config:type=\"int\">2</config:config-item>", fixed = TRUE)
  expect_match(settings, "<config:config-item config:name=\"ShowPageBreakPreview\" config:type=\"boolean\">true</config:config-item>", fixed = TRUE)
  expect_match(settings, "xmlns:ooo=\"http://openoffice.org/2004/office\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  expect_equal(wb2$worksheets[[1]]$sheetViews,
               "<sheetViews><sheetView showZeros=\"0\" rightToLeft=\"1\" zoomScale=\"80\" topLeftCell=\"C10\" workbookViewId=\"0\"/></sheetViews>")
  expect_match(wb2$worksheets[[2]]$sheetViews, "view=\"pageBreakPreview\" zoomScale=\"70\" zoomScaleNormal=\"120\"", fixed = TRUE)
  expect_equal(wb2$worksheets[[1]]$sheetPr, "<sheetPr><tabColor rgb=\"FFFF0000\"/></sheetPr>")
})

test_that("files of other producers are loaded", {
  fods <- paste0('<?xml version="1.0" encoding="UTF-8"?>
<document xmlns="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:xl="http://www.w3.org/1999/xlink"
  o:version="1.3" o:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<body><spreadsheet>
<t:table>
  <t:table-row>
    <t:table-cell t:formula="of:=1/0" o:value-type="float" o:value="0"><x:p>#DIV/0!</x:p></t:table-cell>
    <t:table-cell t:formula="of:=[\'A very long name of a sheet with: and more\'.A1]+1" o:value-type="date" o:date-value="2024-01-02"><x:p>2</x:p></t:table-cell>
    <t:table-cell o:value-type="string"><x:p>\u6771<x:ruby><x:ruby-base>\u4eac</x:ruby-base><x:ruby-text>kyo</x:ruby-text></x:ruby></x:p></t:table-cell>
  </t:table-row>
</t:table>
<t:table t:name="A very long name of a sheet with: and more">
  <t:table-row><t:table-cell o:value-type="float" o:value="1"><x:p>1</x:p></t:table-cell></t:table-row>
</t:table>
<t:table t:name="Scenario"><t:scenario t:scenario-ranges="Scenario.A1" t:is-active="false"/>
  <t:table-row><t:table-cell/></t:table-row>
</t:table>
<t:table t:name="\'file:///other.ods\'#Sheet1"><t:table-source xl:href="file:///other.ods" t:table-name="Sheet1" t:mode="copy-results-only"/>
  <t:table-row><t:table-cell/></t:table-row>
</t:table>
</spreadsheet></body></document>')
  fl <- tempfile(fileext = ".fods")
  writeLines(enc2utf8(fods), fl, useBytes = TRUE)

  wb <- load_ods(fl)
  expect_equal(wb$sheet_names, c("Sheet1", "A very long name of a sheet wit"), ignore_attr = TRUE)
  cc <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(cc$c_t[cc$r == "A1"], "e")
  expect_equal(cc$v[cc$r == "A1"], "#DIV/0!")
  expect_equal(cc$f[cc$r == "B1"], "'A very long name of a sheet wit'!A1+1")
  expect_match(wb$styles_mgr$styles$cellXfs[as.integer(cc$c_s[cc$r == "B1"]) + 1], "numFmtId=\"14\"", fixed = TRUE)
  expect_equal(cc$is[cc$r == "C1"], "<is><t xml:space=\"preserve\">\u6771\u4eac</t></is>")

  fl2 <- tempfile(fileext = ".ods")
  write_ods(wb, fl2)
  exdir <- tempfile()
  utils::unzip(fl2, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "office:string-value=\"\" calcext:value-type=\"error\"><text:p>#DIV/0!</text:p>", fixed = TRUE)
})

test_that("page breaks, print options and headers of even and first pages roundtrip", {
  wb <- wb_workbook()$
    add_worksheet("s")$
    add_data(x = matrix(1:200, 40, 5))$
    add_page_break(row = 10)$
    add_page_break(col = 3)$
    set_header_footer(header = c("L", "C", "R"), even_header = c("eL", NA, NA), first_footer = c(NA, NA, "ff"))
  wb$worksheets[[1]]$printOptions <- "<printOptions horizontalCentered=\"1\" headings=\"1\" gridLines=\"1\"/>"
  wb$worksheets[[1]]$pageSetup <- paste0(
    "<pageSetup paperSize=\"9\" orientation=\"landscape\" pageOrder=\"overThenDown\" ",
    "useFirstPageNumber=\"1\" firstPageNumber=\"5\" cellComments=\"atEnd\"/>"
  )

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "styles.xml"), exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  styles <- paste(readLines(file.path(exdir, "styles.xml"), warn = FALSE), collapse = "")
  expect_match(content, "<style:table-row-properties style:row-height=\"[0-9.]+pt\" fo:break-before=\"page\"")
  expect_match(content, "<style:table-column-properties fo:break-before=\"page\"", fixed = TRUE)
  expect_match(styles, "style:print=\"annotations charts drawings grid headers objects zero-values\"", fixed = TRUE)
  expect_match(styles, "style:print-page-order=\"ltr\" style:first-page-number=\"5\" style:table-centering=\"horizontal\"", fixed = TRUE)
  expect_match(styles, "<style:header-left><style:region-left><text:p>eL</text:p></style:region-left></style:header-left>", fixed = TRUE)
  expect_match(styles, "<style:footer-first><style:region-right><text:p>ff</text:p></style:region-right></style:footer-first>", fixed = TRUE)

  wb2 <- load_ods(fl)
  ws <- wb2$worksheets[[1]]
  expect_equal(ws$rowBreaks, "<brk id=\"10\" max=\"16383\" man=\"1\"/>")
  expect_equal(ws$colBreaks, "<brk id=\"3\" max=\"1048575\" man=\"1\"/>")
  expect_equal(ws$printOptions, "<printOptions horizontalCentered=\"1\" headings=\"1\" gridLines=\"1\"/>")
  expect_match(ws$pageSetup, "pageOrder=\"overThenDown\" cellComments=\"atEnd\" firstPageNumber=\"5\" useFirstPageNumber=\"1\"", fixed = TRUE)
  expect_equal(ws$headerFooter$oddHeader, c("L", "C", "R"))
  expect_equal(ws$headerFooter$evenHeader, c("eL", "", ""))
  expect_equal(ws$headerFooter$firstFooter, c("", "", "ff"))
})

test_that("split views, grid colors and sparklines roundtrip", {
  wb <- wb_workbook()$add_worksheet("S 1")$add_data(x = matrix(c(1, -2, 3, 4, 5, 6, 7, 8, -9), 3), col_names = FALSE)
  wb$worksheets[[1]]$sheetViews <- paste0(
    "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\" defaultGridColor=\"0\" colorId=\"10\">",
    "<pane xSplit=\"2400\" ySplit=\"1800\" topLeftCell=\"C8\" activePane=\"bottomRight\"/>",
    "<selection pane=\"bottomRight\" activeCell=\"D10\" sqref=\"D10\"/></sheetView></sheetViews>"
  )
  sl <- create_sparklines("S 1", "A1:C1", "D1", type = "column", markers = "1", color_series = wb_color("FF00FF"))
  wb$add_sparklines(sparklines = sl)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "settings.xml"), exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  settings <- paste(readLines(file.path(exdir, "settings.xml"), warn = FALSE), collapse = "")
  expect_match(settings, "<config:config-item config:name=\"HorizontalSplitMode\" config:type=\"short\">1</config:config-item>", fixed = TRUE)
  expect_match(settings, "<config:config-item config:name=\"HorizontalSplitPosition\" config:type=\"int\">160</config:config-item>", fixed = TRUE)
  expect_match(settings, "<config:config-item config:name=\"GridColor\" config:type=\"int\">16711680</config:config-item>", fixed = TRUE)
  expect_match(content, "<calcext:sparkline calcext:cell-address=\"'S 1'.D1\" calcext:data-range=\"'S 1'.A1:'S 1'.C1\"/>", fixed = TRUE)
  expect_match(content, "calcext:type=\"column\" calcext:line-width=\"0.75pt\" calcext:markers=\"true\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  ws <- wb2$worksheets[[1]]
  expect_match(ws$sheetViews, "defaultGridColor=\"0\" colorId=\"10\"", fixed = TRUE)
  expect_equal(ws$freezePane, paste0(
    "<pane xSplit=\"2400\" ySplit=\"1800\" topLeftCell=\"C8\" activePane=\"bottomRight\" state=\"split\"/>",
    "<selection pane=\"bottomRight\" activeCell=\"D10\" sqref=\"D10\"/>"
  ))
  expect_match(ws$extLst, "<x14:sparklineGroup type=\"column\" displayEmptyCellsAs=\"gap\" markers=\"1\"><x14:colorSeries rgb=\"FFFF00FF\"/>", fixed = TRUE)
  expect_match(ws$extLst, "<xm:f>'S 1'!A1:C1</xm:f><xm:sqref>D1</xm:sqref>", fixed = TRUE)
})

test_that("extensions of data validations and conditional formats roundtrip", {
  wb <- wb_workbook()$add_worksheet("a")$add_data(x = -5:5)$add_worksheet("b")$add_data(x = c("x", "y", "z"))
  wb$add_conditional_formatting(sheet = "a", dims = "A1:A11", type = "dataBar", style = c("FF0000", "00FF00"), params = list(gradient = FALSE))
  ext <- paste0(
    "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\">",
    "<x14:dataValidations xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" count=\"1\">",
    "<x14:dataValidation type=\"list\" allowBlank=\"1\" showErrorMessage=\"1\">",
    "<x14:formula1><xm:f>b!$A$2:$A$4</xm:f></x14:formula1><xm:sqref>B1:B5</xm:sqref></x14:dataValidation></x14:dataValidations></ext>",
    "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\">",
    "<x14:conditionalFormattings xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:conditionalFormatting>",
    "<x14:cfRule type=\"expression\" priority=\"2\" id=\"{00000000-0000-0000-0000-000000000001}\"><xm:f>A1=b!$A$2</xm:f>",
    "<x14:dxf><fill><patternFill><bgColor rgb=\"FFFFFF00\"/></patternFill></fill></x14:dxf></x14:cfRule>",
    "<x14:cfRule type=\"iconSet\" priority=\"3\" id=\"{00000000-0000-0000-0000-000000000002}\"><x14:iconSet iconSet=\"3Stars\">",
    "<x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"percent\"><xm:f>33</xm:f></x14:cfvo>",
    "<x14:cfvo type=\"percent\"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>C1:C11</xm:sqref>",
    "</x14:conditionalFormatting></x14:conditionalFormattings></ext>"
  )
  wb$worksheets[[1]]$extLst <- c(wb$worksheets[[1]]$extLst, ext)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "styles.xml"), exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  styles <- paste(readLines(file.path(exdir, "styles.xml"), warn = FALSE), collapse = "")
  expect_match(content, "table:condition=\"of:cell-content-is-in-list([$'b'.$A$2:.$A$4])\"", fixed = TRUE)
  expect_match(content, "calcext:negative-color=\"#ff0000\" calcext:axis-color=\"#000000\" calcext:gradient=\"false\"", fixed = TRUE)
  expect_match(content, "calcext:value=\"formula-is([.A1]=[$'b'.$A$2])\"", fixed = TRUE)
  expect_match(content, "<calcext:icon-set calcext:icon-set-type=\"3Stars\">", fixed = TRUE)
  expect_match(styles, "<style:table-cell-properties fo:background-color=\"#ffff00\"/>", fixed = TRUE)

  wb2 <- load_ods(fl)
  ws <- wb2$worksheets[[1]]
  expect_match(ws$dataValidations, "<formula1>'b'!$A$2:$A$4</formula1>", fixed = TRUE)
  expect_match(ws$conditionalFormatting$cf[1], "<x14:id>{00000000-0000-0000-0000-000000000001}</x14:id>", fixed = TRUE)
  expect_match(ws$extLst, "<x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"0\">", fixed = TRUE)
  expect_match(ws$extLst, "<x14:iconSet iconSet=\"3Stars\">", fixed = TRUE)
  expect_match(ws$conditionalFormatting$cf[2], "<formula>A1='b'!$A$2</formula>", fixed = TRUE)
})

test_that("custom geometries and the position of frames in their cells roundtrip", {
  geometry <- paste0(
    "<a:custGeom><a:avLst/><a:gdLst><a:gd name=\"half\" fmla=\"*/ h 1 2\"/></a:gdLst><a:pathLst>",
    "<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"50\"/></a:moveTo>",
    "<a:arcTo wR=\"50\" hR=\"50\" stAng=\"10800000\" swAng=\"10800000\"/>",
    "<a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo><a:lnTo><a:pt x=\"0\" y=\"100\"/></a:lnTo><a:close/></a:path>",
    "<a:path fill=\"none\"><a:moveTo><a:pt x=\"0\" y=\"half\"/></a:moveTo><a:lnTo><a:pt x=\"w\" y=\"half\"/></a:lnTo></a:path>",
    "</a:pathLst></a:custGeom>"
  )
  shape <- sub("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>", geometry, create_shape(shape = "rect"), fixed = TRUE)
  wb <- wb_workbook()$add_worksheet()$add_drawing(dims = "B2:E10", xml = shape)
  wb$drawings[[1]] <- sub("<xdr:colOff>0</xdr:colOff>", "<xdr:colOff>127000</xdr:colOff>", wb$drawings[[1]], fixed = TRUE)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "draw:type=\"non-primitive\" draw:enhanced-path=\"M 0 [0-9]+ C ")
  expect_match(content, "Z N M 0 [0-9]+ L [0-9]+ [0-9]+ F N\"")

  wb2 <- load_ods(fl)
  drawing <- wb2$drawings[[1]]
  expect_match(drawing, "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>127000</xdr:colOff>", fixed = TRUE)
  expect_match(drawing, "<a:custGeom>", fixed = TRUE)
  expect_match(drawing, "<a:path w=\"[0-9]+\" h=\"[0-9]+\" fill=\"none\"><a:moveTo>")
  expect_false(grepl("prstGeom", drawing, fixed = TRUE))
})

test_that("lines, connectors, dashes, line ends, effects and text boxes roundtrip", {
  anchor <- function(c0, c1, body) {
    paste0(
      "<xdr:twoCellAnchor><xdr:from><xdr:col>", c0, "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
      "<xdr:to><xdr:col>", c1, "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>", body,
      "<xdr:clientData/></xdr:twoCellAnchor>"
    )
  }
  style <- paste0(
    "<xdr:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>",
    "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></xdr:style>"
  )
  connector <- paste0(
    "<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr><xdr:cNvPr id=\"2\" name=\"C1\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>",
    "<a:xfrm flipH=\"1\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom>",
    "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:prstDash val=\"dash\"/>",
    "<a:headEnd type=\"none\"/><a:tailEnd type=\"triangle\" w=\"lg\" len=\"lg\"/></a:ln></xdr:spPr>", style, "</xdr:cxnSp>"
  )
  bent <- sub("straightConnector1", "bentConnector3", sub("<a:xfrm flipH=\"1\">", "<a:xfrm flipV=\"1\">", connector, fixed = TRUE), fixed = TRUE)
  bent <- sub("name=\"C1\"", "name=\"C2\"", sub("triangle", "oval", bent, fixed = TRUE), fixed = TRUE)
  box <- paste0(
    "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"4\" name=\"S1\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>",
    "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
    "<a:solidFill><a:srgbClr val=\"4472C4\"><a:alpha val=\"50000\"/></a:srgbClr></a:solidFill>",
    "<a:ln w=\"38100\"><a:solidFill><a:srgbClr val=\"000000\"><a:alpha val=\"40000\"/></a:srgbClr></a:solidFill><a:prstDash val=\"lgDashDot\"/></a:ln>",
    "<a:effectLst><a:glow rad=\"139700\"><a:srgbClr val=\"70AD47\"><a:alpha val=\"40000\"/></a:srgbClr></a:glow>",
    "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"2700000\" algn=\"tl\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"40000\"/></a:srgbClr></a:outerShdw>",
    "<a:softEdge rad=\"63500\"/></a:effectLst></xdr:spPr>", style,
    "<xdr:txBody><a:bodyPr lIns=\"180000\" tIns=\"36000\" rIns=\"180000\" bIns=\"36000\" anchor=\"ctr\" wrap=\"none\"><a:spAutoFit/></a:bodyPr>",
    "<a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>box</a:t></a:r></a:p></xdr:txBody></xdr:sp>"
  )
  wb <- wb_workbook()$add_worksheet()$add_drawing(dims = "B2:E7", xml = create_shape())
  wb$drawings[[1]] <- paste0(
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
    anchor(1, 4, connector), anchor(5, 8, bent), anchor(9, 12, box), "</xdr:wsDr>"
  )

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "styles.xml"), exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  styles <- paste(readLines(file.path(exdir, "styles.xml"), warn = FALSE), collapse = "")
  expect_match(content, "<draw:line draw:z-index=\"0\" draw:name=\"C1\" draw:style-name=\"gr[0-9]+\" svg:x1=\"5.0800cm\" svg:y1=\"0.0000cm\" svg:x2=\"0.0000cm\"")
  expect_match(content, "draw:stroke=\"dash\" draw:stroke-dash=\"dash\" svg:stroke-color=\"#ff0000\" svg:stroke-width=\"2pt\"", fixed = TRUE)
  expect_match(content, "draw:marker-end=\"msArrowEnd_20_9\" draw:marker-end-width=\"10pt\"", fixed = TRUE)
  expect_match(content, "draw:mirror-vertical=\"true\" svg:viewBox=\"0 0 21600 21600\" draw:type=\"ooxml-bentConnector3\"", fixed = TRUE)
  expect_match(content, "draw:marker-end=\"msArrowOvalEnd_20_9\" draw:marker-end-width=\"10pt\" draw:marker-end-center=\"true\"", fixed = TRUE)
  expect_match(content, "draw:fill-color=\"#4472c4\" draw:opacity=\"50%\" draw:stroke=\"dash\" draw:stroke-dash=\"lgDashDot\"", fixed = TRUE)
  expect_match(content, "svg:stroke-opacity=\"40%\"", fixed = TRUE)
  expect_match(content, "draw:shadow=\"visible\" draw:shadow-offset-x=\"2.12132", fixed = TRUE)
  expect_match(content, "loext:glow-radius=\"11pt\" loext:glow-color=\"#70ad47\" loext:glow-transparency=\"60%\" loext:softedge-radius=\"5pt\"", fixed = TRUE)
  expect_match(content, "draw:textarea-vertical-align=\"middle\" fo:padding-left=\"14.173228pt\"", fixed = TRUE)
  expect_match(content, "fo:wrap-option=\"no-wrap\" draw:auto-grow-height=\"true\"", fixed = TRUE)
  expect_match(content, "<style:text-properties fo:font-size=\"11pt\" fo:color=\"#ffffff\"/>", fixed = TRUE)
  expect_match(styles, "<draw:stroke-dash draw:name=\"lgDashDot\" draw:style=\"rect\" draw:dots1=\"1\" draw:dots1-length=\"800%\" draw:dots2=\"1\"", fixed = TRUE)
  expect_match(styles, "<draw:marker draw:name=\"msArrowEnd_20_9\" draw:display-name=\"msArrowEnd 9\"", fixed = TRUE)

  drawing <- load_ods(fl)$drawings[[1]]
  expect_match(drawing, "<a:xfrm rot=\"0\" flipH=\"1\">", fixed = TRUE)
  expect_match(drawing, "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:prstDash val=\"dash\"/><a:headEnd type=\"none\"/><a:tailEnd type=\"triangle\" w=\"lg\" len=\"lg\"/></a:ln>", fixed = TRUE)
  expect_match(drawing, "<a:xfrm rot=\"0\" flipV=\"1\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"bentConnector3\">", fixed = TRUE)
  expect_match(drawing, "<a:headEnd type=\"none\"/><a:tailEnd type=\"oval\" w=\"lg\" len=\"lg\"/>", fixed = TRUE)
  expect_match(drawing, "<a:solidFill><a:srgbClr val=\"4472C4\"><a:alpha val=\"50000\"/></a:srgbClr></a:solidFill><a:ln w=\"38100\">", fixed = TRUE)
  expect_match(drawing, "<a:effectLst><a:glow rad=\"139700\"><a:srgbClr val=\"70AD47\"><a:alpha val=\"40000\"/></a:srgbClr></a:glow>", fixed = TRUE)
  expect_match(drawing, "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"2700000\" rotWithShape=\"0\">", fixed = TRUE)
  expect_match(drawing, "<a:softEdge rad=\"63500\"/></a:effectLst>", fixed = TRUE)
  expect_match(drawing, "anchor=\"ctr\" lIns=\"180000\" tIns=\"36000\" rIns=\"180000\" bIns=\"36000\" wrap=\"none\"><a:spAutoFit/></a:bodyPr>", fixed = TRUE)
})

test_that("text, lines, markers, labels, number formats and tick marks of charts roundtrip", {
  ns <- paste(
    "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
  )
  txpr <- "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"900\"><a:solidFill><a:srgbClr val=\"7F7F7F\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>"
  ser <- function(i, col, extra = "", labels = "") {
    paste0(
      "<c:ser><c:idx val=\"", i, "\"/><c:order val=\"", i, "\"/><c:tx><c:strRef><c:f>data!$", col, "$1</c:f></c:strRef></c:tx>", extra, labels,
      "<c:cat><c:numRef><c:f>data!$A$2:$A$4</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>data!$", col, "$2:$", col, "$4</c:f></c:numRef></c:val></c:ser>"
    )
  }
  line <- paste0(
    "<c:chartSpace ", ns, "><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz=\"1600\" b=\"1\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill></a:rPr>",
    "<a:t>Lines</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"0\"/><c:plotArea><c:layout/>",
    "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>",
    ser(0, "B", "<c:spPr><a:ln w=\"38100\"><a:solidFill><a:srgbClr val=\"00B050\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></c:spPr><c:marker><c:symbol val=\"diamond\"/><c:size val=\"9\"/></c:marker>"),
    ser(1, "C", "", paste0(
      "<c:dLbls><c:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/><c:dLblPos val=\"t\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/>",
      "<c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>"
    )),
    "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:majorTickMark val=\"none\"/>",
    "<c:minorTickMark val=\"none\"/><c:tickLblPos val=\"low\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/>",
    "<c:majorGridlines><c:spPr><a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"D9D9D9\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></c:spPr></c:majorGridlines>",
    "<c:numFmt formatCode=\"0.00%\" sourceLinked=\"0\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>", txpr,
    "<c:crossAx val=\"1\"/><c:majorUnit val=\"1\"/><c:minorUnit val=\"0.5\"/></c:valAx><c:spPr><a:solidFill><a:srgbClr val=\"F2F2F2\"/></a:solidFill></c:spPr></c:plotArea>",
    "<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"0\"/></c:legend><c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"gap\"/></c:chart></c:chartSpace>"
  )
  pie <- paste0(
    "<c:chartSpace ", ns, "><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/><c:pieChart><c:varyColors val=\"1\"/>",
    ser(0, "B", "<c:explosion val=\"0\"/><c:dPt><c:idx val=\"1\"/><c:bubble3D val=\"0\"/><c:explosion val=\"20\"/></c:dPt>"),
    "<c:firstSliceAng val=\"90\"/></c:pieChart></c:plotArea><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
  )
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = data.frame(x = 1:3, a = c(3, 5, 2), b = c(2, 4, 3)))
  wb$add_chart_xml(dims = "E1:L15", xml = line)$add_chart_xml(dims = "E17:L30", xml = pie)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("Object 1/content.xml", "Object 1/styles.xml", "Object 2/content.xml"), exdir = exdir)
  object <- function(i, part = "content.xml") paste(readLines(file.path(exdir, paste("Object", i), part), warn = FALSE), collapse = "")
  expect_match(object(1), "<style:text-properties fo:font-size=\"16pt\" fo:font-weight=\"bold\" fo:color=\"#c00000\"/>", fixed = TRUE)
  expect_match(object(1), "chart:symbol-type=\"named-symbol\" chart:symbol-name=\"diamond\" chart:symbol-width=\"9pt\" chart:symbol-height=\"9pt\"", fixed = TRUE)
  expect_match(object(1), "svg:stroke-color=\"#00b050\" draw:stroke=\"dash\" draw:stroke-dash=\"dash\" svg:stroke-width=\"3pt\"", fixed = TRUE)
  expect_match(object(1), "chart:label-position=\"top\"", fixed = TRUE)
  expect_match(object(1), "chart:interval-minor-divisor=\"2\"", fixed = TRUE)
  expect_match(object(1), "chart:tick-marks-major-inner=\"false\" chart:tick-marks-major-outer=\"true\"", fixed = TRUE)
  expect_match(object(1), "chart:axis-label-position=\"outside-start\"", fixed = TRUE)
  expect_match(object(1), "chart:link-data-style-to-source=\"false\"", fixed = TRUE)
  expect_match(object(1), "<number:percentage-style style:name=\"N[0-9]+\">")
  expect_match(object(1), "<chart:wall chart:style-name=\"ch[0-9]+\"/>")
  expect_match(object(1), "chart:treat-empty-cells=\"leave-gap\"", fixed = TRUE)
  expect_match(object(1, "styles.xml"), "<draw:stroke-dash draw:name=\"sysDash\"", fixed = TRUE)
  expect_match(object(2), "chart:angle-offset=\"0\"", fixed = TRUE)
  expect_match(object(2), "chart:reverse-direction=\"true\"", fixed = TRUE)
  expect_match(object(2), "<style:chart-properties chart:pie-offset=\"20\"/>", fixed = TRUE)

  charts <- load_ods(fl)$charts$chart
  expect_match(charts[1], "<a:rPr sz=\"1600\" b=\"1\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill></a:rPr><a:t>Lines</a:t>", fixed = TRUE)
  expect_match(charts[1], "<c:spPr><a:ln w=\"38100\"><a:solidFill><a:srgbClr val=\"00B050\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></c:spPr><c:marker><c:symbol val=\"diamond\"/><c:size val=\"9\"/></c:marker>", fixed = TRUE)
  expect_match(charts[1], "<c:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/><c:dLblPos val=\"t\"/>", fixed = TRUE)
  expect_match(charts[1], "<c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"low\"/>", fixed = TRUE)
  expect_match(charts[1], "<c:numFmt formatCode=\"0.00%\" sourceLinked=\"0\"/><c:majorTickMark val=\"out\"/>", fixed = TRUE)
  expect_match(charts[1], "<a:defRPr sz=\"900\"><a:solidFill><a:srgbClr val=\"7F7F7F\"/></a:solidFill></a:defRPr>", fixed = TRUE)
  expect_match(charts[1], "<c:majorUnit val=\"1\"/><c:minorUnit val=\"0.5\"/>", fixed = TRUE)
  expect_match(charts[1], "<a:prstDash val=\"sysDash\"/></a:ln></c:spPr></c:majorGridlines>", fixed = TRUE)
  expect_match(charts[1], "<c:spPr><a:solidFill><a:srgbClr val=\"F2F2F2\"/></a:solidFill></c:spPr></c:plotArea>", fixed = TRUE)
  expect_match(charts[1], "<c:dispBlanksAs val=\"gap\"/>", fixed = TRUE)
  expect_match(charts[2], "<c:dPt><c:idx val=\"1\"/><c:explosion val=\"20\"/>", fixed = TRUE)
  expect_match(charts[2], "<c:firstSliceAng val=\"90\"/>", fixed = TRUE)
})

test_that("pivot charts show the points of their cache and border width keywords are read", {
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:pivotSource><c:name>[x.xlsx]data!PivotTable1</c:name>",
    "<c:fmtId val=\"0\"/></c:pivotSource><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/><c:barChart><c:barDir val=\"col\"/>",
    "<c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:cat><c:strRef><c:f>data!$A$2:$A$4</c:f><c:strCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt>",
    "<c:pt idx=\"1\"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat>",
    "<c:val><c:numRef><c:f>data!$B$2:$B$4</c:f><c:numCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt>",
    "<c:pt idx=\"1\"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>",
    "</c:plotArea><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
  )
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = data.frame(x = c("a", "b", "Grand Total"), y = c(1, 2, 3)))
  wb$add_chart_xml(dims = "D1:J10", xml = chart)
  wb$add_border(dims = "A6", left_border = "mediumDashed", right_border = NULL, top_border = NULL, bottom_border = NULL)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = c("content.xml", "Object 1/content.xml"), exdir = exdir)
  object <- paste(readLines(file.path(exdir, "Object 1", "content.xml"), warn = FALSE), collapse = "")
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(object, "chart:values-cell-range-address=\"data.$B$2:data.$B$3\"", fixed = TRUE)
  expect_match(object, "table:cell-range-address=\"data.$A$2:data.$A$3\"", fixed = TRUE)
  expect_false(grepl("loext:data-pilot-source", object, fixed = TRUE))
  expect_match(content, "fo:border-left=\"2pt dashed #000000\"", fixed = TRUE)

  fods <- paste0(
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
    "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" ",
    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" ",
    "office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\"><office:automatic-styles>",
    "<style:style style:name=\"ce1\" style:family=\"table-cell\"><style:table-cell-properties fo:border-left=\"thick solid #ff0000\" ",
    "fo:border-top=\"medium dashed #ff0000\"/></style:style></office:automatic-styles><office:body><office:spreadsheet>",
    "<table:table table:name=\"T\"><table:table-row><table:table-cell table:style-name=\"ce1\" office:value-type=\"float\" office:value=\"1\">",
    "<text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>"
  )
  fl <- tempfile(fileext = ".fods")
  writeLines(fods, fl, useBytes = TRUE)
  borders <- paste(load_ods(fl)$styles_mgr$styles$borders, collapse = "")
  expect_match(borders, "<left style=\"thick\">", fixed = TRUE)
  expect_match(borders, "<top style=\"mediumDashed\">", fixed = TRUE)
})

test_that("flat files are written with embedded pictures and charts and loaded again", {
  img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/>",
    "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:spPr><a:ln><a:prstDash val=\"dash\"/></a:ln></c:spPr>",
    "<c:cat><c:strRef><c:f>data!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>",
    "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>",
    "</c:plotArea><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
  )
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = data.frame(x = c("a", "b", "c"), y = 1:3))
  wb$add_image(dims = "D2:F8", file = img)$add_chart_xml(dims = "H2:N12", xml = chart)

  fl <- tempfile(fileext = ".fods")
  write_ods(wb, fl)
  flat <- paste(readLines(fl, warn = FALSE), collapse = "")
  expect_match(flat, "<office:document xmlns:office=", fixed = TRUE)
  expect_match(flat, "office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\"", fixed = TRUE)
  expect_match(flat, "<draw:image><office:binary-data>/9j/", fixed = TRUE)
  expect_match(flat, "office:mimetype=\"application/vnd.oasis.opendocument.chart\"><office:styles><draw:stroke-dash draw:name=\"dash\"", fixed = TRUE)
  expect_match(flat, "<table:calculation-settings table:case-sensitive=\"false\" table:automatic-find-labels=\"false\" table:use-regular-expressions=\"false\" table:use-wildcards=\"true\"/>", fixed = TRUE)
  expect_false(grepl("xlink:href=\"Pictures/", flat, fixed = TRUE))

  wb2 <- load_ods(fl)
  expect_equal(wb_to_df(wb2), wb_to_df(wb))
  expect_length(wb2$media, 1)
  expect_identical(readBin(unlist(wb2$media), "raw", file.size(img)), readBin(img, "raw", file.size(img)))
  expect_match(wb2$charts$chart, "<a:prstDash val=\"dash\"/>", fixed = TRUE)
})

test_that("errors, number styles, text effects and repeated content of other producers are read", {
  fods <- paste0(
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
    "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" ",
    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" ",
    "xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" ",
    "office:version=\"1.2\" office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\"><office:automatic-styles>",
    "<number:number-style style:name=\"N1\"><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" ",
    "number:min-integer-digits=\"1\" number:grouping=\"true\" number:display-factor=\"1000\">",
    "<number:embedded-text number:position=\"0\">k</number:embedded-text></number:number></number:number-style>",
    "<number:text-style style:name=\"N2P0\"/><number:text-style style:name=\"N2P1\"/><number:text-style style:name=\"N2P2\"/>",
    "<number:text-style style:name=\"N2\"><style:map style:condition=\"value()&gt;0\" style:apply-style-name=\"N2P0\"/>",
    "<style:map style:condition=\"value()&lt;0\" style:apply-style-name=\"N2P1\"/><style:map style:condition=\"value()=0\" style:apply-style-name=\"N2P2\"/>",
    "</number:text-style>",
    "<number:number-style style:name=\"N3\"><number:text>-</number:text><number:fill-character> </number:fill-character>",
    "<number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/></number:number-style>",
    "<style:style style:name=\"ce1\" style:family=\"table-cell\" style:data-style-name=\"N1\"/>",
    "<style:style style:name=\"ce2\" style:family=\"table-cell\" style:data-style-name=\"N2\"/>",
    "<style:style style:name=\"ce3\" style:family=\"table-cell\" style:data-style-name=\"N3\"/>",
    "<style:style style:name=\"ce4\" style:family=\"table-cell\"><style:text-properties style:text-outline=\"true\" fo:text-shadow=\"1pt 1pt\"/></style:style>",
    "<style:style style:name=\"ce5\" style:family=\"table-cell\"><style:table-cell-properties style:repeat-content=\"true\"/></style:style>",
    "</office:automatic-styles><office:body><office:spreadsheet><table:table table:name=\"T\"><table:table-row>",
    "<table:table-cell office:value-type=\"error\" office:string-value=\"#N/A\" table:formula=\"of:=NA()\"><text:p>#N/A</text:p></table:table-cell>",
    "<table:table-cell table:style-name=\"ce1\" office:value-type=\"float\" office:value=\"12000\"><text:p>12k</text:p></table:table-cell>",
    "<table:table-cell table:style-name=\"ce2\" office:value-type=\"float\" office:value=\"1\"/>",
    "<table:table-cell table:style-name=\"ce3\" office:value-type=\"float\" office:value=\"1\"><text:p>- 1.00</text:p></table:table-cell>",
    "<table:table-cell table:style-name=\"ce4\" office:value-type=\"string\"><text:p>a</text:p></table:table-cell>",
    "<table:table-cell table:style-name=\"ce5\" office:value-type=\"string\"><text:p>-</text:p></table:table-cell>",
    "</table:table-row></table:table></office:spreadsheet></office:body></office:document>"
  )
  fl <- tempfile(fileext = ".fods")
  writeLines(fods, fl, useBytes = TRUE)
  wb <- load_ods(fl)
  expect_match(paste(wb$worksheets[[1]]$sheet_data$cc$c_t[1], wb$worksheets[[1]]$sheet_data$cc$v[1]), "e #N/A", fixed = TRUE)
  formats <- paste(wb$styles_mgr$styles$numFmts, collapse = "")
  expect_match(formats, "formatCode=\"#,##0,&quot;k&quot;\"", fixed = TRUE)
  expect_match(formats, "formatCode=\";;;\"", fixed = TRUE)
  expect_match(formats, "formatCode=\"-* 0.00\"", fixed = TRUE)
  expect_match(paste(wb$styles_mgr$styles$fonts, collapse = ""), "<outline/><shadow/>", fixed = TRUE)
  expect_match(paste(wb$styles_mgr$styles$cellXfs, collapse = ""), "<alignment horizontal=\"fill\"/>", fixed = TRUE)
})

test_that("formatted headers and radial gradients roundtrip", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1)
  wb$set_header_footer(header = c("&\"Arial,Bold\"left &A", "&14center", NA), footer = c(NA, NA, "&\"Courier New,Italic\"&KFF0000red &U&&"))
  shape <- create_shape(shape = "ellipse")
  shape <- sub("<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>", paste0(
    "<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FFFF00\"/></a:gs>",
    "<a:gs pos=\"100000\"><a:srgbClr val=\"00B050\"/></a:gs></a:gsLst><a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/>",
    "</a:path></a:gradFill>"
  ), shape, fixed = TRUE)
  wb$add_drawing(dims = "B2:D6", xml = shape)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "styles.xml", exdir = exdir)
  styles <- paste(readLines(file.path(exdir, "styles.xml"), warn = FALSE), collapse = "")
  expect_match(styles, "<text:span text:style-name=\"HT1\">left <text:sheet-name/></text:span>", fixed = TRUE)
  expect_match(styles, "<style:text-properties fo:font-family=\"Arial\" fo:font-weight=\"bold\"/>", fixed = TRUE)
  expect_match(styles, "draw:style=\"radial\" draw:start-color=\"#00b050\" draw:end-color=\"#ffff00\" draw:cx=\"50%\" draw:cy=\"50%\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  expect_equal(
    wb2$worksheets[[1]]$headerFooter$oddHeader[1:2],
    c("&amp;\"Arial,Bold\"left &amp;A", "&amp;14center")
  )
  expect_equal(wb2$worksheets[[1]]$headerFooter$oddFooter[3], "&amp;\"Courier New,Italic\"&amp;KFF0000red &amp;U&amp;&amp;")
  expect_match(wb2$drawings[[1]], "<a:gs pos=\"0\"><a:srgbClr val=\"FFFF00\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"00B050\"/></a:gs></a:gsLst><a:path path=\"circle\">", fixed = TRUE)
})

test_that("pivot charts refer to their data pilot table and described ranges are read from the local table", {
  df <- data.frame(group = c("a", "b", "a", "b"), value = 1:4)
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = df)
  wb$add_pivot_table(wb_data(wb), dims = "A1", rows = "group", data = "value", fun = "sum", pivot_table = "pivot")
  chart <- paste0(
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:pivotSource><c:name>[x.xlsx]Sheet 1!pivot</c:name>",
    "<c:fmtId val=\"0\"/></c:pivotSource><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/><c:barChart><c:barDir val=\"col\"/>",
    "<c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>",
    "<c:cat><c:strRef><c:f>'Sheet 1'!$A$2:$A$4</c:f><c:strCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt>",
    "<c:pt idx=\"1\"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat>",
    "<c:val><c:numRef><c:f>'Sheet 1'!$B$2:$B$4</c:f><c:numCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>4</c:v></c:pt>",
    "<c:pt idx=\"1\"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
    "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>",
    "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>",
    "</c:plotArea><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>"
  )
  wb$add_chart_xml(sheet = 2, dims = "D1:J10", xml = chart)

  fl <- tempfile(fileext = ".fods")
  write_ods(wb, fl)
  flat <- paste(readLines(fl, warn = FALSE), collapse = "")
  expect_match(flat, "loext:data-pilot-source=\"pivot\"", fixed = TRUE)
  expect_match(flat, "<table:data-pilot-table table:name=\"pivot\"", fixed = TRUE)
  expect_match(flat, "chart:values-cell-range-address=\"'Sheet 1'.$B$2:'Sheet 1'.$B$3\"", fixed = TRUE)

  # the ranges of the data pilot table as a consumer describes them in the local table
  flat <- sub("chart:values-cell-range-address=\"'Sheet 1'.$B$2:'Sheet 1'.$B$3\"", "chart:values-cell-range-address=\"PT@data 0\"", flat, fixed = TRUE)
  flat <- sub("table:cell-range-address=\"'Sheet 1'.$A$2:'Sheet 1'.$A$3\"", "table:cell-range-address=\"PT@categories\"", flat, fixed = TRUE)
  parts <- strsplit(flat, "<table:table table:name=\"local-table\">", fixed = TRUE)[[1]]
  parts[2] <- sub("<text:p>a</text:p></table:table-cell>", "<text:p>a</text:p><draw:g><svg:desc>PT@categories</svg:desc></draw:g></table:table-cell>", parts[2], fixed = TRUE)
  parts[2] <- sub("<text:p>4</text:p></table:table-cell>", "<text:p>4</text:p><draw:g><svg:desc>PT@data 0</svg:desc></draw:g></table:table-cell>", parts[2], fixed = TRUE)
  flat <- paste(parts, collapse = "<table:table table:name=\"local-table\">")
  expect_match(flat, "\"PT@data 0\"", fixed = TRUE)
  writeLines(flat, fl, useBytes = TRUE)
  ch <- load_ods(fl)$charts$chart
  expect_match(ch, "<c:cat><c:strLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt><c:pt idx=\"1\"><c:v>b</c:v></c:pt></c:strLit></c:cat>", fixed = TRUE)
  expect_match(ch, "<c:val><c:numLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>4</c:v></c:pt><c:pt idx=\"1\"><c:v>6</c:v></c:pt></c:numLit></c:val>", fixed = TRUE)
})

test_that("error and date period conditions are native conditions", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = data.frame(x = c(1, NA), y = Sys.Date() + 0:1))
  wb$add_style(create_dxfs_style(bg_fill = wb_color("FFFFC7CE")), "bad")
  wb$add_style(create_dxfs_style(bg_fill = wb_color("FFC6EFCE")), "good")
  wb$add_style(create_dxfs_style(bg_fill = wb_color("FFFFEB9C")), "neutral")
  wb$add_conditional_formatting(dims = "A2:A3", type = "containsErrors", style = "bad")
  wb$add_conditional_formatting(dims = "A2:A3", type = "notContainsErrors", style = "good")
  wb$add_conditional_formatting(dims = "B2:B3", rule = "B2>0", style = "neutral")
  wb$worksheets[[1]]$conditionalFormatting$cf[3] <- paste0(
    "<cfRule type=\"timePeriod\" dxfId=\"2\" priority=\"3\" timePeriod=\"yesterday\">",
    "<formula>FLOOR(B2,1)=TODAY()-1</formula></cfRule>"
  )

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "calcext:value=\"is-error\"", fixed = TRUE)
  expect_match(content, "calcext:value=\"is-no-error\"", fixed = TRUE)
  expect_match(content, "<calcext:date-is calcext:date=\"yesterday\" calcext:style=\"ConditionalStyle_3\"/>", fixed = TRUE)

  wb2 <- load_ods(fl)
  cf <- wb2$worksheets[[1]]$conditionalFormatting$cf
  expect_match(cf[1], "type=\"containsErrors\"", fixed = TRUE)
  expect_match(cf[1], "<formula>ISERROR(A2)</formula>", fixed = TRUE)
  expect_match(cf[2], "type=\"notContainsErrors\"", fixed = TRUE)
  expect_match(cf[3], "type=\"timePeriod\" dxfId=\"2\" timePeriod=\"yesterday\"", fixed = TRUE)
  expect_match(cf[3], "<formula>FLOOR(B2,1)=TODAY()-1</formula>", fixed = TRUE)
})

test_that("medium and thick borders keep their weight in other consumers", {
  wb <- wb_workbook()$add_worksheet()$add_data(x = 1)
  wb$add_border(dims = "B2", left_border = "medium", right_border = "thick", top_border = "thin", bottom_border = "mediumDashed")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "fo:border-left=\"2pt solid #000000\" fo:border-right=\"thick solid #000000\"", fixed = TRUE)
  expect_match(content, "fo:border-top=\"0.74pt solid #000000\" fo:border-bottom=\"2pt dashed #000000\"", fixed = TRUE)

  wb2 <- load_ods(fl)
  borders <- wb2$styles_mgr$styles$borders
  expect_true(any(grepl("<left style=\"medium\">.*<right style=\"thick\">.*<top style=\"thin\">.*<bottom style=\"mediumDashed\">", borders)))
})

test_that("merged cells carry the borders of their edges", {
  wb <- wb_workbook()$add_worksheet()$add_data(dims = "B2", x = "a")
  wb$add_border(dims = "B2:D4", left_color = wb_color("FF9BBB59"), right_color = wb_color("FF9BBB59"),
                top_color = wb_color("FF9BBB59"), bottom_color = wb_color("FF9BBB59"))
  wb$merge_cells(dims = "B2:D4")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "table:number-rows-spanned=\"3\" table:style-name=\"cm1\"", fixed = TRUE)
  expect_match(content, paste0(
    "<style:style style:name=\"cm1\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:table-cell-properties ",
    "fo:border-left=\"0.74pt solid #9bbb59\" fo:border-top=\"0.74pt solid #9bbb59\" ",
    "fo:border-right=\"0.74pt solid #9bbb59\" fo:border-bottom=\"0.74pt solid #9bbb59\"/>"
  ), fixed = TRUE)
})

test_that("the output of a pivot table lies in its location", {
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = mtcars[, c("cyl", "mpg")])$add_worksheet("pv")
  wb$add_pivot_table(wb_data(wb, sheet = "data"), sheet = "pv", dims = "B2", rows = "cyl", data = "mpg")
  location <- xml_node(wb$pivotTables[1], "pivotTableDefinition", "location")
  wb$pivotTables[1] <- sub(location, xml_attr_mod(location, c(ref = "B2:C6")), wb$pivotTables[1], fixed = TRUE)
  wb$add_data(sheet = "pv", x = data.frame(cyl = c(4, 6, 8, NA), mpg = c(293.3, 138.2, 211.4, 642.9)), dims = "B2")

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  wb2 <- load_ods(fl)
  expect_match(xml_node(wb2$pivotTables[1], "pivotTableDefinition", "location"), "<location ref=\"B2:C6\"", fixed = TRUE)
})

test_that("the layout of pivot tables roundtrips", {
  wb <- wb_workbook()$add_worksheet("data")$add_data(x = mtcars[, c("cyl", "gear", "mpg", "hp")])$add_worksheet("pv")
  x <- wb_data(wb, sheet = "data")
  wb$add_pivot_table(x, sheet = "pv", dims = "A1", rows = "cyl", data = c("mpg", "hp"), params = list(outline = TRUE))
  wb$add_pivot_table(x, sheet = "pv", dims = "E1", rows = "cyl", data = "mpg", params = list(outline = TRUE, compact = FALSE))
  wb$add_pivot_table(x, sheet = "pv", dims = "I1", rows = "gear", data = "mpg")
  location <- xml_node(wb$pivotTables[1], "pivotTableDefinition", "location")
  wb$pivotTables[1] <- sub(location, xml_attr_mod(location, c(ref = "A1:C5")), wb$pivotTables[1], fixed = TRUE)

  fl <- tempfile(fileext = ".ods")
  write_ods(wb, fl)
  exdir <- tempfile()
  utils::unzip(fl, files = "content.xml", exdir = exdir)
  content <- paste(readLines(file.path(exdir, "content.xml"), warn = FALSE), collapse = "")
  expect_match(content, "table:layout-mode=\"outline-subtotals-top\" loext:layout-mode=\"compact-layout\"", fixed = TRUE)
  expect_match(content, "table:layout-mode=\"outline-subtotals-top\"/>", fixed = TRUE)
  expect_match(content, "table:layout-mode=\"tabular-layout\"/>", fixed = TRUE)

  wb2 <- load_ods(fl)
  pt <- wb2$pivotTables
  expect_match(pt[1], "outline=\"1\"", fixed = TRUE)
  expect_false(grepl("compact=\"0\"", pt[1], fixed = TRUE))
  expect_match(xml_node(pt[1], "pivotTableDefinition", "location"), "<location ref=\"A1:C5\" firstHeaderRow=\"0\" firstDataRow=\"1\"", fixed = TRUE)
  expect_match(pt[2], "outline=\"1\"", fixed = TRUE)
  expect_match(pt[2], "compact=\"0\"", fixed = TRUE)
  expect_false(grepl("outline=\"1\"", pt[3], fixed = TRUE))
})
