package convertator.pptx

import convertator.model.{ConversionConfig, PageContent, PageImage, PageMode, PageTable, TextElement, TextLine}

import org.apache.poi.xslf.usermodel.*
import org.apache.poi.sl.usermodel.{PictureData, TextParagraph as SLTextParagraph}
import org.apache.poi.sl.usermodel.TableCell.BorderEdge

import java.awt.Color
import java.awt.geom.Rectangle2D
import java.io.{FileOutputStream, OutputStream}
import scala.collection.mutable.ListBuffer

object SlideBuilder:

  def convert(pages: Seq[PageContent], config: ConversionConfig, output: OutputStream): Unit =
    val ppt = new XMLSlideShow()
    try
      ppt.setPageSize(java.awt.Dimension(config.slideWidth.toInt, config.slideHeight.toInt))
      setDocumentLanguage(ppt, "uk-UA")
      val overflowTotal = new java.util.concurrent.atomic.AtomicInteger(0)
      for page <- pages do renderPage(ppt, page, config, overflowTotal)
      val total = overflowTotal.get()
      if total > 0 then
        println(s"  ℹ  Content split across $total additional slide(s)")
      ppt.write(output)
    finally ppt.close()

  def convert(pages: Seq[PageContent], config: ConversionConfig, outputPath: String): Unit =
    val fos = new FileOutputStream(outputPath)
    try convert(pages, config, fos)
    finally fos.close()

  private def renderPage(
    ppt: XMLSlideShow, page: PageContent, cfg: ConversionConfig,
    overflowTotal: java.util.concurrent.atomic.AtomicInteger
  ): Unit =
    val availW = cfg.slideWidth - 2 * cfg.marginX
    val availH = cfg.slideHeight - 2 * cfg.marginY
    val (posScale, fontScale) = cfg.pageMode match
      case PageMode.Fit =>
        val sx = (availW / page.pageWidth).toFloat
        val sy = (availH / page.pageHeight).toFloat
        val ps = math.min(sx, sy)
        (ps, ps * cfg.fontSizeScale)
      case PageMode.Flow =>
        val ps = (availW / page.pageWidth).toFloat
        (ps, cfg.fontSizeScale)

    buildSlides(ppt, page.lines, page.images, page.tables, cfg, posScale, fontScale, availW, availH, overflowTotal)

  // ---- content item --------------------------------------------------------

  private sealed trait ContentItem:
    def y: Float
    def imgOrder: Int  // 0=text, 1=image
  private case class T(line: TextLine) extends ContentItem:
    val y = line.y; val imgOrder = 1
  private case class I(img: PageImage) extends ContentItem:
    val y = img.y; val imgOrder = 0
  private case class Tbl(tbl: PageTable) extends ContentItem:
    val y = tbl.y; val imgOrder = 0

  private val MaxImgH = 0.35

  // ---- slide building -----------------------------------------------------

  private def buildSlides(
    ppt: XMLSlideShow, lines: Seq[TextLine], images: Seq[PageImage], tables: Seq[PageTable],
    cfg: ConversionConfig, posScale: Float, fontScale: Float,
    availW: Double, availH: Double,
    overflowTotal: java.util.concurrent.atomic.AtomicInteger
  ): Unit =
    val safetyMargin = 40.0
    val usableH      = availH - safetyMargin
    val maxImgH      = usableH * MaxImgH

    val items = ListBuffer.from(
      (lines.map(T(_)) ++ images.map(I(_)) ++ tables.map(Tbl(_)))
        .sortBy(c => (if c.y < 0f then Float.MaxValue else c.y, c.imgOrder))
    )

    while items.nonEmpty do
      val slide = newSlide(ppt)
      var tb: XSLFTextBox = null  // created lazily when first text item is reached
      var curY  = 0.0
      var idx   = 0
      var overflow = false

      while idx < items.length && !overflow do
        items(idx) match
          case T(line) =>
            val fs        = line.maxFontSize * fontScale
            // Estimate visual lines after word-wrap in PowerPoint.
            // Line spacing = 120% of font -> fs * 1.2 pt per visual line.
            // Use char width factor 0.65 (conservative for Cyrillic).
            val avgCharW  = (fs * 0.65).max(1.0)
            val chPerLine = (availW / avgCharW).toInt.max(1)
            // total rendered chars = sum of text lengths + estimated spaces between words
            val approxSpaces = line.elements.length
            val totalChars   = line.elements.map(_.text.length).sum + approxSpaces
            val visLines     = math.ceil(totalChars.toDouble / chPerLine).toInt.max(1)
            val lineH        = visLines * (fs * 1.2) + cfg.lineSpacing
            val fits         = curY + lineH <= usableH

            if fits then
              // lazy-create text box at current vertical position (after any images)
              if tb == null then
                val gap = if curY > 0 then 14.0 else 0.0
                tb = newTextBox(slide, cfg, availW, availH - curY - gap, curY + gap)
              // whole line fits - add as one paragraph
              addLineAsParagraph(tb, line, cfg, posScale, fontScale)
              curY += lineH; idx += 1
            else if idx == 0 then
              // first item on the slide but too long - split across slides
              if tb == null then
                val gap = if curY > 0 then 14.0 else 0.0
                tb = newTextBox(slide, cfg, availW, availH - curY - gap, curY + gap)
              val spaceLeft   = usableH - curY
              val linesThatFit = (spaceLeft / (fs * 1.2)).toInt.max(1)
              val charsThatFit = linesThatFit * chPerLine
              val (fitElems, remainElems) = splitElements(line.elements, charsThatFit)
              if fitElems.nonEmpty then
                addLineAsParagraph(tb, TextLine(fitElems, line.y, line.maxFontSize),
                  cfg, posScale, fontScale)
              if remainElems.nonEmpty then
                // replace current item with remainder for the next slide
                items.update(idx, T(TextLine(remainElems, line.y, line.maxFontSize)))
              else
                idx += 1
              // Do NOT set overflow - we handled this item. Other items can still
              // be placed on this slide (images, shorter text after the split).
            else
              // doesn't fit and not first item on slide - overflow
              overflow = true; overflowTotal.incrementAndGet()

          case I(img) =>
            var iw = img.width.toDouble * posScale
            var ih = img.height.toDouble * posScale
            if iw > availW then
              val s = availW / iw; iw = availW; ih = ih * s
            if ih > maxImgH then
              val s = maxImgH / ih; ih = maxImgH; iw = iw * s
            val fits = curY + ih <= usableH

            if tb == null then
              // No text yet on this slide - place image normally
              if fits || idx == 0 then
                try
                  val pt = mimeToPictureType(img.contentType)
                  val pd = ppt.addPicture(img.data, pt)
                  slide.createPicture(pd).setAnchor(new Rectangle2D.Double(
                    cfg.marginX + (availW - iw) / 2, cfg.marginY + curY, iw, ih))
                catch case e: Exception => Console.err.println(s"  \u26a0 Image: ${e.getMessage}")
                curY += ih; idx += 1
              if !fits then
                overflow = true; overflowTotal.incrementAndGet()
            else
              // Text already on this slide - overflow image to next slide
              // to prevent overlapping text and images
              overflow = true
              overflowTotal.incrementAndGet()

          case Tbl(tbl) =>
            val tw = tbl.columnWidths.sum.toDouble
            // Scale table if wider than available space
            val scaleW = if tw > availW then availW / tw else 1.0
            val scaledHeights = tbl.rows.map(r => r.height.toDouble * scaleW)
            val th = scaledHeights.sum
            val fits = curY + th <= usableH

            if tb == null then
              if fits then
                // Whole table fits - render all rows
                try
                  drawTable(slide, tbl, cfg, availW, scaleW, curY, 0, tbl.rows.length)
                catch case e: Exception => Console.err.println(s"  ⚠ Table: ${e.getMessage}")
                curY += th; idx += 1
              else if idx == 0 then
                // Table too tall - split rows across slides
                val spaceLeft = usableH - curY
                var accH = 0.0
                var splitIdx = 0
                while splitIdx < scaledHeights.length &&
                      accH + scaledHeights(splitIdx) <= spaceLeft do
                  accH += scaledHeights(splitIdx)
                  splitIdx += 1

                if splitIdx == 0 then
                  // No rows fit in remaining space - overflow whole remainder
                  // to the next slide
                  overflow = true; overflowTotal.incrementAndGet()
                else
                  try
                    drawTable(slide, tbl, cfg, availW, scaleW, curY, 0, splitIdx)
                  catch case e: Exception => Console.err.println(s"  ⚠ Table: ${e.getMessage}")
                  curY += accH

                  if splitIdx < tbl.rows.length then
                    // Keep remaining rows for the next slide
                    val remaining = PageTable(tbl.y, tbl.columnWidths, tbl.rows.drop(splitIdx))
                    items.update(idx, Tbl(remaining))
                  else
                    idx += 1
                  // Do NOT set overflow - we handled this table by splitting it
              else
                // Table doesn't fit and not first item - overflow
                overflow = true; overflowTotal.incrementAndGet()
            else
              // Text already on slide - overflow table to prevent overlap
              overflow = true
              overflowTotal.incrementAndGet()

      items.remove(0, idx)

  // ---- helpers ------------------------------------------------------------

  private def mimeToPictureType(mime: String): PictureData.PictureType =
    mime.toLowerCase match
      case "image/png"    => PictureData.PictureType.PNG
      case "image/jpeg" | "image/jpg" => PictureData.PictureType.JPEG
      case "image/gif"    => PictureData.PictureType.GIF
      case "image/bmp"    => PictureData.PictureType.BMP
      case "image/svg+xml" => PictureData.PictureType.SVG
      case s if s.contains("emf") => PictureData.PictureType.EMF
      case s if s.contains("wmf") => PictureData.PictureType.WMF
      case _              => PictureData.PictureType.PNG

  private def setDocumentLanguage(ppt: XMLSlideShow, lang: String): Unit =
    val pres   = ppt.getCTPresentation
    val dts    = if pres.isSetDefaultTextStyle then pres.getDefaultTextStyle
                 else pres.addNewDefaultTextStyle()
    val defPPr = if dts.isSetDefPPr then dts.getDefPPr else dts.addNewDefPPr()
    val defRPr = if defPPr.isSetDefRPr then defPPr.getDefRPr else defPPr.addNewDefRPr()
    defRPr.setLang(lang); defRPr.setAltLang(lang)

  private def newSlide(ppt: XMLSlideShow): XSLFSlide = ppt.createSlide

  private def newTextBox(slide: XSLFSlide, cfg: ConversionConfig, availW: Double, boxH: Double, marginYOff: Double): XSLFTextBox =
    val tb = slide.createTextBox()
    tb.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT)
    val boxY = cfg.marginY + marginYOff
    val boxW = math.max(1.0, availW)
    val boxHt = math.max(1.0, boxH)
    tb.setAnchor(new Rectangle2D.Double(cfg.marginX, boxY, boxW, boxHt))
    tb.setWordWrap(true)
    val fp = tb.getTextParagraphs
    if !fp.isEmpty then
      val cp = fp.get(0).getXmlObject
      (if cp.isSetEndParaRPr then cp.getEndParaRPr else cp.addNewEndParaRPr()).setLang("uk-UA")
    tb

  /** Render rows [startRow, endRow) of a PageTable as an XSLFTable on the slide. */
  private def drawTable(
    slide: XSLFSlide, tbl: PageTable, cfg: ConversionConfig,
    availW: Double, scaleW: Double, curY: Double,
    startRow: Int, endRow: Int
  ): Unit =
    if tbl.rows.isEmpty || tbl.columnWidths.isEmpty then return
    val slice = tbl.rows.slice(startRow, endRow)
    if slice.isEmpty then return
    val numRows = slice.length
    val numCols = tbl.columnWidths.length
    val scaledWidths = tbl.columnWidths.map(w => (w * scaleW).toFloat)

    val totalW = scaledWidths.sum.toDouble
    val scaledHeights = slice.map(r => r.height.toDouble * scaleW)
    val totalH = scaledHeights.sum
    val xPos = cfg.marginX + (availW - totalW) / 2  // center horizontally

    val xs = slide.createTable(numRows, numCols)
    xs.setAnchor(new Rectangle2D.Double(xPos, cfg.marginY + curY, totalW, totalH))

    // Set column widths
    for (w, col) <- scaledWidths.zipWithIndex do
      xs.setColumnWidth(col, w.toDouble)

    // Set explicit row heights so cell text doesn't overflow
    val xRows = xs.getRows()
    for (ri <- 0 until math.min(slice.length, xRows.size())) do
      xRows.get(ri).setHeight(scaledHeights(ri))

    val borderColor = new Color(180, 180, 180)  // light gray

    // Populate cells
    for (row, ri) <- slice.zipWithIndex do
      for (cell, ci) <- row.cells.zipWithIndex do
        val xslfCell = xs.getCell(ri, ci)
        xslfCell.clearText()

        // Add thin gray borders
        xslfCell.setBorderWidth(BorderEdge.top, 0.5)
        xslfCell.setBorderColor(BorderEdge.top, borderColor)
        xslfCell.setBorderWidth(BorderEdge.bottom, 0.5)
        xslfCell.setBorderColor(BorderEdge.bottom, borderColor)
        xslfCell.setBorderWidth(BorderEdge.left, 0.5)
        xslfCell.setBorderColor(BorderEdge.left, borderColor)
        xslfCell.setBorderWidth(BorderEdge.right, 0.5)
        xslfCell.setBorderColor(BorderEdge.right, borderColor)
        // White background so text behind the table (from overlapping text box)
        // does not show through transparent cells
        xslfCell.setFillColor(Color.WHITE)

        if cell.textElements.nonEmpty then
          val para = xslfCell.addNewTextParagraph()
          para.setTextAlign(SLTextParagraph.TextAlign.LEFT)
          para.setLineSpacing(110.0)
          val (merged, rep) = mergeCellText(cell.textElements)
          val run = para.addNewTextRun()
          run.setText(merged)
          val fs = math.max(6.0, (rep.fontSize * 1.3).toDouble) // slightly larger for tables
          run.setFontSize(fs)
          run.setFontColor(Color.BLACK)
          run.setFontFamily(cleanFontName(rep.fontName))
          if rep.bold then run.setBold(true)
          if rep.italic then run.setItalic(true)
          if cfg.preserveUnderline && rep.underline then run.setUnderlined(true)
          run.getRPr(true).setLang("uk-UA")

  /** Merge cell text elements into a single text string, preserving the dominant
    * formatting representative. */
  private def mergeCellText(elems: List[TextElement]): (String, TextElement) =
    val sb = new StringBuilder
    var rep: TextElement = null
    for e <- elems do
      if sb.nonEmpty && e.x > 0f then sb.append(' ')
      sb.append(e.text)
      if rep == null || e.fontSize > rep.fontSize then rep = e
    val text = sb.toString
    (if text.isEmpty then " " else text,
     if rep == null then TextElement(" ", 0f, 0f, 11f, "Calibri") else rep)

  private def elemRight(e: TextElement): Float =
    e.x + (if e.width > 0f then e.width else e.text.length * e.fontSize * 0.55f)

  private def buildWords(elements: List[TextElement]): List[(String, TextElement)] =
    val sorted = elements.sortBy(_.x)
    val result = List.newBuilder[(String, TextElement)]
    var i = 0
    while i < sorted.length do
      val start = sorted(i)
      val sb = new StringBuilder(start.text); val rep = start; var right = elemRight(start)
      var j = i + 1; var done = false
      while j < sorted.length && !done do
        val n = sorted(j); val gap = n.x - right
        val th = math.max(start.fontSize, n.fontSize) * 0.3f
        if gap <= th then
          if gap > 1.0f then sb.append(' ')
          sb.append(n.text); right = math.max(right, elemRight(n)); j += 1
        else done = true
      result += ((sb.toString(), rep)); i = j
    result.result()

  private def addLineAsParagraph(
    tb: XSLFTextBox, line: TextLine, cfg: ConversionConfig,
    posScale: Float, fontScale: Float
  ): Unit =
    val words = buildWords(line.elements)
    if words.isEmpty then return
    val para = tb.addNewTextParagraph()
    para.setTextAlign(SLTextParagraph.TextAlign.LEFT)
    para.setLineSpacing(120.0); para.setSpaceBefore(0.0); para.setSpaceAfter(cfg.lineSpacing)
    var first = true
    for (wt, rep) <- words do
      if !first then
        val sp = para.addNewTextRun()
        sp.setText(" "); sp.setFontSize(math.max(4.0, (rep.fontSize * fontScale).toDouble))
        sp.setFontColor(Color.BLACK); sp.setFontFamily(cleanFontName(rep.fontName))
        sp.getRPr(true).setLang("uk-UA")
      val run = para.addNewTextRun()
      run.setText(wt)
      val fs = math.max(4.0, (rep.fontSize * fontScale).toDouble)
      run.setFontSize(fs); run.setFontColor(Color.BLACK)
      run.setFontFamily(cleanFontName(rep.fontName))
      if rep.bold then run.setBold(true)
      if rep.italic then run.setItalic(true)
      if cfg.preserveUnderline && rep.underline then run.setUnderlined(true)
      run.getRPr(true).setLang("uk-UA"); first = false
    val cp = para.getXmlObject
    (if cp.isSetEndParaRPr then cp.getEndParaRPr else cp.addNewEndParaRPr()).setLang("uk-UA")

  private def splitElements(elems: List[TextElement], maxChars: Int): (List[TextElement], List[TextElement]) =
    var rem = maxChars
    val before = List.newBuilder[TextElement]
    val after  = List.newBuilder[TextElement]
    for e <- elems do
      if rem >= e.text.length then
        before += e
        rem -= e.text.length
      else if rem > 0 then
        val (head, tail) = e.text.splitAt(rem)
        val avgW = e.fontSize * 0.55f
        before += e.copy(text = head)
        after  += e.copy(text = tail, x = e.x + head.length.toFloat * avgW)
        rem = 0
      else
        after += e
    (before.result(), after.result())

  private def cleanFontName(raw: String): String =
    val c = if raw.contains('+') then raw.substring(raw.indexOf('+') + 1) else raw
    c.split("[-,\\.]").headOption.getOrElse(c).trim

end SlideBuilder
