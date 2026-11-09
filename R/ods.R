#' @importFrom utils download.file
getFile <- function(xlsxFile) {

  ## Is this a file or URL (code taken from read.table())
  on.exit(try(close(fl), silent = TRUE), add = TRUE)
  fl <- file(description = xlsxFile)

  ## If URL download
  if (inherits(fl, "url")) {
    tmpFile <- tempfile(fileext = ".xlsx")
    download.file(url = xlsxFile, destfile = tmpFile, cacheOK = FALSE, mode = "wb", quiet = TRUE)
    xlsxFile <- tmpFile
  }

  xlsxFile
}

#' Write a workbook to an OpenDocument Spreadsheet file
#'
#' Converts a `wbWorkbook` into an OpenDocument Spreadsheet (`ods`) file. The
#' conversion covers cell values, rich text, formulas, defined names, print
#' areas and print titles, cell styles (fonts, fills, borders, alignment,
#' number formats), merged cells, column widths, row heights, row and column
#' outlines, hidden sheets, comments, hyperlinks, images, charts, chartsheets,
#' shapes, form controls, tables, autofilters, pivot tables, conditional
#' formatting, data validation, freeze panes, grid lines, zoom, page setup with
#' headers and footers, sheet and workbook protection and the document
#' properties. Slicers, timelines and threaded comments have no counterpart in
#' OpenDocument and are not written.
#'
#' Shapes, lines and connectors are written with their geometry, flips, text,
#' text formatting, alignment, text box insets and anchor, rotation, fill
#' (colors with transparency or a linear or radial gradient), line (width,
#' dash, line ends and transparency), shadow, glow and soft edges; grouped
#' shapes are written as a group placed on the sheet. Charts are written with
#' their titles, legend, axes, text properties, fills, lines, markers and data
#' labels. Pictures are written with
#' their cropping and rotation, the background picture of a sheet as background
#' image of the table. Conditional formats are written in the calc extension of
#' the office application and, as far as they compare the cell value or
#' evaluate a formula, also as conditions of the cell styles. Check boxes, option buttons, drop down lists, list boxes, buttons and
#' labels are written as controls of a form. Chartsheets are written as sheets
#' without cells that hold their chart. Tables are written as database ranges,
#' without their table style. Pivot tables are written as data pilot tables
#' with their row, column, page and data fields, functions, shown values, sort
#' order, hidden items and grand totals. Pivot tables created with
#' [openxlsx2::wb_add_pivot_table()] have no output yet, the spreadsheet
#' software shows them once they are refreshed. Charts are converted with
#' their plot type (bar, line, area, pie, doughnut, scatter, radar, bubble,
#' stock), series, categories, secondary axes, axis scaling, titles, legend,
#' colors, markers, data labels, data point colors and trendlines. Other chart
#' formatting is not converted. Passwords of protected sheets and workbooks
#' are not written, their hashes cannot be converted.
#'
#' Formulas are translated into the OpenFormula syntax. Shared formulas are
#' written into each of their cells, structured references into references to
#' the cells of the table. References to other files are written as `#REF!`.
#'
#' @param wb A `wbWorkbook`
#' @param file A path to the `ods` file, a file with the extension `fods` is
#'   written as flat `fods` file with the pictures and charts embedded
#' @param overwrite If `FALSE`, an existing file is not overwritten.
#' @param version The OpenDocument version of the file, `"1.4"` or `"1.3"`.
#'   Files of both versions have the same content; office applications that
#'   predate OpenDocument 1.4 expect version `"1.3"`.
#' @return The `wbWorkbook`, invisibly
#' @seealso [load_ods()]
#' @examples
#' wb <- openxlsx2::wb_workbook()$add_worksheet()$add_data(x = mtcars)
#' write_ods(wb, tempfile(fileext = ".ods"))
#' @export
write_ods <- function(wb, file, overwrite = TRUE, version = c("1.4", "1.3")) {
  version <- match.arg(version)
  if (!inherits(wb, "wbWorkbook")) {
    stop("wb must be a wbWorkbook", call. = FALSE)
  }

  if (file.exists(file) && !overwrite) {
    stop("File already exists!", call. = FALSE)
  }

  ods_dir <- tempfile("odsx_")
  on.exit(unlink(ods_dir, recursive = TRUE), add = TRUE)
  dir.create(file.path(ods_dir, "META-INF"), recursive = TRUE)

  # chartsheets are written as sheets without cells that hold their chart
  sel <- seq_along(wb$worksheets)
  empty_sheet <- wb_workbook()$add_worksheet()$worksheets[[1]]$sheet_data

  sheets <- lapply(sel, function(i) {
    ws <- wb$worksheets[[i]]
    sheet_data <- if (wb$is_chartsheet[i]) empty_sheet else ws$sheet_data
    drawing <- ws$relships$drawing
    comments <- ws$relships$comments
    hf <- if (is.list(ws$headerFooter)) ws$headerFooter else list()
    # the left, center and right parts of the odd, even and first page headers and footers
    parts <- c("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")
    header_footer <- vapply(1:18, function(k) {
      part <- hf[[parts[(k - 1) %/% 3 + 1]]]
      pos <- (k - 1) %% 3 + 1
      if (length(part) >= pos && length(part[[pos]]) && !is.na(part[[pos]][1])) paste(part[[pos]], collapse = "") else ""
    }, "")
    cf <- ws$conditionalFormatting
    list(
      cc            = sheet_data$cc,
      row_attr      = sheet_data$row_attr,
      cols          = as.character(ws$cols_attr),
      merges        = as.character(ws$mergeCells),
      format        = paste(ws$sheetFormatPr, collapse = ""),
      protection    = paste(ws$sheetProtection, collapse = ""),
      pane          = as.character(ws$freezePane),
      views         = paste(ws$sheetViews, collapse = ""),
      page_setup    = paste(ws$pageSetup, collapse = ""),
      page_margins  = paste(ws$pageMargins, collapse = ""),
      sheet_pr      = paste(ws$sheetPr, collapse = ""),
      print_options = paste(ws$printOptions, collapse = ""),
      header_footer = header_footer,
      row_breaks    = as.character(ws$rowBreaks),
      col_breaks    = as.character(ws$colBreaks),
      comments      = if (length(comments)) wb$comments[[comments]] else NULL,
      hyperlinks    = as.character(ws$hyperlinks),
      rels          = as.character(wb$worksheets_rels[[i]]),
      drawing       = if (length(drawing)) paste(wb$drawings[[drawing]], collapse = "") else "",
      drawing_rels  = if (length(drawing)) as.character(wb$drawings_rels[[drawing]]) else character(),
      picture       = paste(ws$picture, collapse = ""),
      vml           = if (length(ws$relships$vmlDrawing)) paste(wb$vml[[ws$relships$vmlDrawing]], collapse = "") else "",
      validations   = as.character(ws$dataValidations),
      cf_sqref      = if (length(cf)) as.character(cf$sqref) else character(),
      cf_rule       = if (length(cf)) as.character(cf$cf) else character(),
      autofilter    = paste(ws$autoFilter, collapse = ""),
      ext           = as.character(ws$extLst)
    )
  })

  # all pictures and chart objects are prepared, unused ones are removed
  charts <- as.character(wb$charts$chart)
  if (length(wb$media)) {
    dir.create(file.path(ods_dir, "Pictures"))
    file.copy(unlist(wb$media), file.path(ods_dir, "Pictures", names(wb$media)))
  }
  for (k in seq_along(charts)) dir.create(file.path(ods_dir, paste("Object", k)))

  styles <- wb$styles_mgr$styles
  tabs <- wb$tables
  if (!is.null(tabs)) tabs <- tabs[tabs$tab_act == 1, ]

  # the first data row of the source of each pivot cache, page fields require a selected value
  pivot_first <- lapply(as.character(wb$pivotDefinitions), function(cache) {
    source <- xml_attr(cache, "pivotCacheDefinition", "cacheSource", "worksheetSource")[[1]]
    if (!all(c("sheet", "ref") %in% names(source)) || !source[["sheet"]] %in% wb$get_sheet_names()) {
      return(character())
    }
    first <- wb_to_df(wb, sheet = source[["sheet"]], dims = source[["ref"]], convert = FALSE, skip_empty_rows = FALSE)
    if (!nrow(first)) return(character())
    first <- vapply(first, function(x) as.character(x[1]), "")
    first[is.na(first)] <- ""
    unname(first)
  })

  used <- write_ods_parts(
    dir            = ods_dir,
    sheet_names    = wb$get_sheet_names()[sel],
    hidden         = wb$get_sheet_visibility()[sel] != "visible",
    sheets         = sheets,
    shared_strings = as.character(unlist(wb$sharedStrings)),
    xfs            = as.character(styles$cellXfs),
    fonts          = as.character(styles$fonts),
    fills          = as.character(styles$fills),
    borders        = as.character(styles$borders),
    numfmts        = as.character(styles$numFmts),
    indexed_colors = as.character(styles$indexedColors),
    theme          = paste(wb$theme, collapse = ""),
    workbook_pr    = paste(wb$workbook$workbookPr, collapse = ""),
    defined_names  = as.character(wb$workbook$definedNames),
    sheet_index    = sel - 1L,
    table_xml      = as.character(tabs$tab_xml),
    table_sheets   = as.character(wb$get_sheet_names()[tabs$tab_sheet]),
    book           = list(
      charts              = charts,
      version             = version,
      dxfs                = as.character(styles$dxfs),
      core                = paste(wb$core, collapse = ""),
      book_views          = paste(wb$workbook$bookViews, collapse = ""),
      workbook_protection = paste(wb$workbook$workbookProtection, collapse = ""),
      pivot_tables        = as.character(wb$pivotTables),
      pivot_tables_rels   = as.character(wb$pivotTables.xml.rels),
      pivot_caches        = as.character(wb$pivotDefinitions),
      pivot_first         = pivot_first
    )
  )
  prepared <- c(file.path("Pictures", names(wb$media)), paste("Object", seq_along(charts)))
  unlink(file.path(ods_dir, setdiff(prepared, used)), recursive = TRUE)
  if (!length(list.files(file.path(ods_dir, "Pictures")))) unlink(file.path(ods_dir, "Pictures"), recursive = TRUE)

  # a flat file holds all parts, otherwise the mimetype is the first entry of the archive and is not compressed
  out <- file.path(normalizePath(dirname(file), mustWork = TRUE), basename(file))
  if (file.exists(out)) unlink(out)
  if (endsWith(tolower(file), ".fods")) {
    ods_flatten(ods_dir, out)
  } else {
    zip_dir(out, ods_dir, stored = "mimetype")
  }

  invisible(wb)
}

#' Load an OpenDocument Spreadsheet file
#'
#' Loads an OpenDocument Spreadsheet (`ods` or flat `fods`) file into a
#' `wbWorkbook`. The file is converted into the parts of an xlsx file, which
#' are then loaded with [openxlsx2::wb_load()]. The conversion covers cell
#' values, rich text, formulas, defined names, print ranges and print titles,
#' cell styles (fonts, fills, borders, alignment, number formats), merged
#' cells, column widths, row heights, row and column outlines, hidden sheets,
#' comments, hyperlinks, images, charts, shapes, form controls, database ranges
#' (as tables or autofilters), data pilot tables (as pivot tables), conditional
#' formatting, data validation, freeze panes, grid lines, zoom, page setup with
#' headers and footers, sheet and workbook protection and the document
#' properties.
#'
#' Shapes, lines and connectors are loaded with their geometry, flips, text,
#' text formatting, alignment, text box insets and anchor, rotation, fill
#' (colors with transparency or a gradient from the start to the end color),
#' line (color, width, dash and line ends), shadow, glow and soft edges; the
#' shapes of a group are loaded as single shapes. Pictures are
#' loaded with their cropping and rotation, the background image of a table as
#' background picture of the sheet. Conditions of cell styles are loaded as
#' conditional formats if the file has none of the calc extension. Check
#' boxes, option buttons and drop down lists are loaded, other controls are
#' not. Pivot tables are created with [openxlsx2::wb_add_pivot_table()], their
#' output cells keep the values of the file until the pivot table is refreshed.
#'
#' Formulas in the OpenFormula syntax are translated into the xlsx syntax.
#' Formulas with the `msoxl:` namespace prefix are loaded unchanged. Formulas
#' in other syntaxes are not loaded, only their values, and a warning reports
#' how many formulas were dropped.
#'
#' As with [openxlsx2::wb_load()], the workbook contains all sheets if a
#' `sheet` is selected, but only the cells of the selected sheet are loaded.
#' With `data_only = TRUE` the workbook is reduced to the cells and styles;
#' comments, images, charts, shapes, form controls and pivot tables are not
#' loaded.
#'
#' @param file A path to an existing `ods` or `fods` file
#' @param sheet Optional, the name or position of the sheet to load
#' @param data_only If `TRUE`, only the data is loaded
#' @param ... additional arguments passed to [openxlsx2::wb_load()]
#' @return A `wbWorkbook`
#' @seealso [write_ods()], [read_ods()], [ods_open()]
#' @examples
#' fl <- tempfile(fileext = ".ods")
#' write_ods(openxlsx2::wb_workbook()$add_worksheet()$add_data(x = mtcars), fl)
#' wb <- load_ods(fl)
#' wb <- load_ods(fl, sheet = 1, data_only = TRUE)
#' @export
load_ods <- function(file, sheet, data_only = FALSE, ...) {

  file <- getFile(file)
  if (length(file) != 1 || !file.exists(file)) {
    stop("File does not exist.", call. = FALSE)
  }

  ods_dir <- tempfile("odsx_")
  on.exit(unlink(ods_dir, recursive = TRUE), add = TRUE)
  xlsx_dir <- file.path(ods_dir, "xlsx")
  for (part_dir in c("_rels", "docProps", "xl/_rels", "xl/media", "xl/tables", "xl/worksheets/_rels")) {
    dir.create(file.path(xlsx_dir, part_dir), recursive = TRUE)
  }

  zip_magic <- as.raw(c(0x50, 0x4B, 0x03, 0x04))
  if (identical(readBin(file, "raw", 4L), zip_magic)) {
    unzip_dir <- file.path(ods_dir, "ods")
    withCallingHandlers(
      utils::unzip(file, exdir = unzip_dir),
      warning = function(w) stop("Unable to open and load file: ", file, call. = FALSE)
    )
    mimetype <- file.path(unzip_dir, "mimetype")
    if (file.exists(mimetype)) {
      type <- readChar(mimetype, file.size(mimetype), useBytes = TRUE)
      if (!startsWith(type, "application/vnd.oasis.opendocument.spreadsheet")) {
        stop("File does not appear to be an ods file: ", file, call. = FALSE)
      }
    }
    manifest <- file.path(unzip_dir, "META-INF", "manifest.xml")
    if (file.exists(manifest) && any(grepl("encryption-data", readLines(manifest, warn = FALSE), fixed = TRUE))) {
      stop("Encrypted ods files are not supported: ", file, call. = FALSE)
    }
    content <- file.path(unzip_dir, "content.xml")
    if (!file.exists(content)) {
      stop("File does not appear to be an ods file: ", file, call. = FALSE)
    }
    parts <- file.path(unzip_dir, c("styles.xml", "settings.xml", "meta.xml"))
    parts[!file.exists(parts)] <- ""
  } else {
    # the parts of a flat file with its pictures and chart objects
    unzip_dir <- file.path(ods_dir, "ods")
    dir.create(unzip_dir)
    ods_unflatten(path.expand(file), unzip_dir)
    content <- file.path(unzip_dir, "content.xml")
    parts <- file.path(unzip_dir, c("styles.xml", "settings.xml", "meta.xml"))
  }

  all_sheets <- missing(sheet)
  info <- read_ods_parts(
    content, parts[1], parts[2], parts[3], xlsx_dir,
    all_sheets, if (all_sheets) NULL else sheet, data_only
  )
  if (info$dropped > 0) {
    warning(
      info$dropped, " formulas in an unknown syntax were not loaded, only their values.",
      call. = FALSE
    )
  }

  xlsx <- file.path(ods_dir, "ods.xlsx")
  zip_dir(xlsx, xlsx_dir)

  wb <- if (all_sheets) wb_load(xlsx, data_only = data_only, ...) else wb_load(xlsx, sheet = sheet, data_only = data_only, ...)
  wb$path <- file
  if (data_only) return(wb)

  # comments, images, charts and shapes are added to the loaded workbook
  for (i in seq_along(info$sheets)) {
    extra <- info$sheets[[i]]
    for (k in seq_along(extra$comment_ref)) {
      text <- extra$comment_text[k]
      if (extra$comment_rich[k]) class(text) <- c("character", "fmt_txt")
      comment <- wb_comment(text = text, author = extra$comment_author[k])
      wb$add_comment(sheet = i, dims = extra$comment_ref[k], comment = comment)
    }
    for (k in seq_along(extra$image_dims)) {
      image <- file.path(unzip_dir, extra$image_file[k])
      if (!file.exists(image)) next
      wb$add_image(sheet = i, dims = extra$image_dims[k], file = image)
      crop <- c(extra$image_extra$l[k], extra$image_extra$t[k], extra$image_extra$r[k], extra$image_extra$b[k])
      drawing <- wb$worksheets[[i]]$relships$drawing
      wb$drawings[[drawing]] <- ods_adjust_drawing(
        wb$drawings[[drawing]], crop, extra$image_extra$rotation[k], "", "", 0, "", extra$anchors$image[(8 * k - 7):(8 * k)], ""
      )
    }
    for (k in seq_along(extra$chart_dims)) {
      wb$add_chart_xml(sheet = i, dims = extra$chart_dims[k], xml = extra$chart_xml[k])
      drawing <- wb$worksheets[[i]]$relships$drawing
      wb$drawings[[drawing]] <- ods_adjust_drawing(
        wb$drawings[[drawing]], numeric(), 0, "", "", 0, "", extra$anchors$chart[(8 * k - 7):(8 * k)], ""
      )
    }
    for (k in seq_along(extra$shape_dims)) {
      fill <- extra$shape_fill[k]
      line <- extra$shape_line[k]
      text <- extra$shape_text[k]
      if (extra$shape_rich[k]) class(text) <- c("character", "fmt_txt")
      shape_args <- list(
        shape             = extra$shape_type[k],
        name              = paste("Shape", k),
        text              = text,
        fill_color        = if (nzchar(fill) && fill != "none") wb_color(fill),
        fill_transparency = if (fill == "none") 100 else 0,
        line_color        = if (nzchar(line) && line != "none") wb_color(line),
        line_transparency = if (line == "none") 100 else 0,
        text_align        = extra$shape_align[k],
        rotation          = round(extra$shape_extra$rotation[k], 4),
        id                = k
      )
      # geometries without a preset of spreadsheet software are rectangles
      shape <- tryCatch(do.call(create_shape, shape_args), error = function(e) NULL)
      if (is.null(shape)) {
        shape_args$shape <- "rect"
        shape <- do.call(create_shape, shape_args)
      }
      wb$add_drawing(sheet = i, dims = extra$shape_dims[k], xml = shape)
      shape_extra <- extra$shape_extra
      has_gradient <- nzchar(shape_extra$gradient_start[k]) && nzchar(shape_extra$gradient_end[k])
      drawing <- wb$worksheets[[i]]$relships$drawing
      wb$drawings[[drawing]] <- ods_adjust_drawing(
        wb$drawings[[drawing]], numeric(), 0,
        if (has_gradient) shape_extra$gradient_start[k] else "", if (has_gradient) shape_extra$gradient_end[k] else "",
        shape_extra$gradient_angle[k], shape_extra$geometry[k], extra$anchors$shape[(8 * k - 7):(8 * k)], shape_extra$properties[k]
      )
    }
  }

  # check boxes, option buttons and drop down lists
  for (i in seq_along(info$sheets)) {
    controls <- info$sheets[[i]]$controls
    for (k in seq_along(controls$dims)) {
      wb$add_form_control(
        sheet   = i,
        dims    = controls$dims[k],
        type    = controls$type[k],
        text    = if (nzchar(controls$text[k])) controls$text[k],
        link    = if (nzchar(controls$link[k])) controls$link[k],
        range   = if (nzchar(controls$range[k])) controls$range[k],
        checked = controls$checked[k]
      )
    }
  }

  # pivot tables are created from their definition, the cells of their output keep their values;
  # their source is only complete if all sheets are loaded
  for (i in seq_along(info$sheets)[all_sheets]) {
    for (pivot in info$sheets[[i]]$pivots) {
      x <- wb_data(wb, sheet = pivot$source_sheet, dims = pivot$source)
      known <- pivot$data %in% names(x)
      pivot$data <- pivot$data[known]
      pivot$fun <- pivot$fun[known]
      pivot$show_data_as <- pivot$show_data_as[known]
      pivot$rows <- intersect(pivot$rows, names(x))
      pivot$cols <- intersect(pivot$cols, names(x))
      pivot$filter <- intersect(pivot$filter, names(x))
      if (!length(pivot$data)) next
      params <- list(
        name             = pivot$name,
        data_on_rows     = pivot$data_on_rows,
        row_grand_totals = pivot$row_grand_totals,
        col_grand_totals = pivot$col_grand_totals
      )
      if (length(pivot$choose)) params$choose <- stats::setNames(pivot$choose, pivot$choose_names)
      if (any(pivot$show_data_as != "normal")) params$show_data_as <- pivot$show_data_as
      if (pivot$sort_row != 0) params$sort_row <- pivot$sort_row
      if (pivot$sort_col != 0) params$sort_col <- pivot$sort_col
      if (pivot$layout %in% c("compact", "outline")) params$outline <- TRUE
      if (pivot$layout == "outline") params$compact <- FALSE
      args <- list(x = x, sheet = i, dims = pivot$dims, data = pivot$data, fun = pivot$fun, params = params)
      if (length(pivot$rows)) args$rows <- pivot$rows
      if (length(pivot$cols)) args$cols <- pivot$cols
      if (length(pivot$filter)) args$filter <- pivot$filter
      do.call(wb$add_pivot_table, args)

      # the output cells lie in the location of the pivot table, otherwise a refresh finds them in its way.
      # Without column fields the hidden values row of an outline leaves the header in the first row.
      corners <- strsplit(pivot$target, ":", fixed = TRUE)[[1]]
      if (length(corners) == 2 && corners[2] != pivot$dims) {
        n <- length(wb$pivotTables)
        location <- xml_node(wb$pivotTables[n], "pivotTableDefinition", "location")
        attrs <- c(ref = paste0(pivot$dims, ":", corners[2]))
        if (pivot$layout %in% c("compact", "outline") && !length(pivot$cols) && length(pivot$data) > 1 && !pivot$data_on_rows) {
          attrs <- c(attrs, firstHeaderRow = "0", firstDataRow = "1")
        }
        wb$pivotTables[n] <- sub(location, xml_attr_mod(location, attrs), wb$pivotTables[n], fixed = TRUE)
      }
    }
  }

  wb
}

#' Read a sheet of an OpenDocument Spreadsheet file into a data frame
#'
#' Loads the cells of a sheet of an `ods` or `fods` file with [load_ods()] and
#' returns them as data frame with [openxlsx2::wb_to_df()].
#'
#' @param file A path to an existing `ods` or `fods` file
#' @param sheet The name or position of the sheet
#' @param ... additional arguments passed to [openxlsx2::wb_to_df()]
#' @return A data frame
#' @seealso [load_ods()]
#' @examples
#' fl <- tempfile(fileext = ".ods")
#' write_ods(openxlsx2::wb_workbook()$add_worksheet()$add_data(x = mtcars), fl)
#' df <- read_ods(fl)
#' @export
read_ods <- function(file, sheet = 1, ...) {
  wb <- load_ods(file, sheet = sheet, data_only = TRUE)
  wb_to_df(wb, sheet = sheet, ...)
}

#' Open a workbook or an OpenDocument Spreadsheet file
#'
#' Opens an `ods` file, or a `wbWorkbook` written to a temporary `ods` file,
#' with the spreadsheet software of the system, as [openxlsx2::xl_open()] does
#' for xlsx files.
#'
#' @param x A `wbWorkbook` or a path to a file
#' @param interactive If `FALSE`, the file is not opened. By default it is
#'   opened in interactive sessions.
#' @return Nothing
#' @seealso [write_ods()], [openxlsx2::xl_open()]
#' @examples
#' if (interactive()) {
#'   ods_open(openxlsx2::wb_workbook()$add_worksheet()$add_data(x = mtcars))
#' }
#' @export
ods_open <- function(x, interactive = NA) {
  if (inherits(x, "wbWorkbook")) {
    file <- tempfile(fileext = ".ods")
    write_ods(x, file)
    x <- file
  }
  xl_open(x, interactive = interactive)
}

# zips the files of a directory, the stored file is the first entry and is not compressed
zip_dir <- function(out, dir, stored = character()) {
  files <- setdiff(list.files(dir, all.files = TRUE, no.. = TRUE), stored)
  zip_tool <- if (Sys.getenv("R_ZIPCMD") != "") Sys.getenv("R_ZIPCMD") else Sys.which("zip")
  if (zip_tool != "") {
    old_wd <- setwd(dir)
    on.exit(setwd(old_wd), add = TRUE)
    if (length(stored)) utils::zip(out, stored, flags = "-0Xq", zip = zip_tool)
    utils::zip(out, files, flags = "-r9Xq", zip = zip_tool)
  } else {
    if (!requireNamespace("zip", quietly = TRUE)) {
      stop(
        "Saving the file requires a zip tool. No system zip tool was found ",
        "and the 'zip' package is not installed.",
        call. = FALSE
      )
    }
    if (length(stored)) {
      zip::zip(out, stored, root = dir, compression_level = 0, mode = "cherry-pick")
      zip::zip_append(out, files, root = dir, compression_level = 9, include_directories = FALSE)
    } else {
      zip::zip(out, files, root = dir, compression_level = 9, include_directories = FALSE)
    }
  }
  invisible(out)
}
