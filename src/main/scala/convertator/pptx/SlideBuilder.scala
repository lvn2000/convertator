package convertator.pptx

import convertator.model.{ConversionConfig, PageContent, PageImage, PageMode, PageTable, TextElement, TextLine}
import convertator.pptx.SlidePlacers.*
import convertator.pptx.SlidePlacers.given

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
        println(s"  \u2139  Content split across $total additional slide(s)")
      ppt.write(output)
    finally ppt.close()

  def convert(pages: Seq[PageContent], config: ConversionConfig, outputPath: String): Unit =
    val fos = new FileOutputStream(outputPath)
    try convert(pages, config, fos)
    finally fos.close()

  // ---- Slide placers moved to convertator.pptx.SlidePlacers -----------------
  // Given instances for SlidePlacer were extracted to SlidePlacers.scala to
  // keep rendering logic focused. See convertator.pptx.SlidePlacers for the
  // implementations.



  private val MaxImgH = 0.35

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
    val ctx          = PlaceContext(cfg, availW, usableH, fontScale, posScale, maxImgH)

    // Build sorted list of content items using the PageItem ADT and Positioned
    val buf = scala.collection.mutable.ArrayBuffer.empty[convertator.model.PageItem]
    buf.addAll(lines.map(convertator.model.PageItem.TextLineItem))
    buf.addAll(images.map(convertator.model.PageItem.ImageItem))
    buf.addAll(tables.map(convertator.model.PageItem.TableItem))

    def sortKey(pi: convertator.model.PageItem): (Float, Int) =
      val pos = summon[convertator.model.Positioned[convertator.model.PageItem]]
      val y = pos.y(pi)
      val ord = pos.imgOrder(pi)
      (if y < 0f then Float.MaxValue else y, ord)

    // Sort in-place
    buf.sortInPlaceBy(sortKey)

    while buf.nonEmpty do
      val slide = newSlide(ppt)
      var tb: XSLFTextBox = null  // created lazily when first text item is reached
      var curY  = 0.0
      var idx   = 0
      var overflow = false

      while idx < buf.length && !overflow do
        buf(idx) match
          case convertator.model.PageItem.TextLineItem(line) =>
            val placer = summon[SlidePlacer[TextLine]]
            val h = placer.height(line, ctx)
            val fits = curY + h <= ctx.usableH
            if fits then
              if tb == null then
                val gap = if curY > 0 then 14.0 else 0.0
                tb = newTextBox(slide, ctx.cfg, ctx.availW, ctx.usableH - curY - gap, curY + gap)
              placer.place(line, slide, ppt, ctx, curY, tb)
              curY += h; idx += 1
            else if idx == 0 then
              if tb == null then
                val gap = if curY > 0 then 14.0 else 0.0
                tb = newTextBox(slide, ctx.cfg, ctx.availW, ctx.usableH - curY - gap, curY + gap)
              val splitRes: Option[(TextLine, Option[TextLine])] = placer.trySplit(line, ctx, curY)
              splitRes match
                case Some(t) =>
                  val (portion, remainder) = t
                  val ph = summon[SlidePlacer[TextLine]].height(portion, ctx)
                  summon[SlidePlacer[TextLine]].place(portion, slide, ppt, ctx, curY, tb)
                  curY += ph
                  remainder match
                    case Some(rem) => buf.update(idx, convertator.model.PageItem.TextLineItem(rem))
                    case None      => idx += 1
                case None =>
                  overflow = true
                  overflowTotal.incrementAndGet()
            else
              overflow = true
              overflowTotal.incrementAndGet()

          case convertator.model.PageItem.ImageItem(img) =>
            val placer = summon[SlidePlacer[PageImage]]
            val h = placer.height(img, ctx)
            val fits = curY + h <= ctx.usableH
            if tb == null then
              if fits then
                placer.place(img, slide, ppt, ctx, curY, tb)
                curY += h; idx += 1
              else if idx == 0 then
                  val splitImg: Option[(PageImage, Option[PageImage])] = placer.trySplit(img, ctx, curY)
                  splitImg match
                    case Some(t) =>
                      val (portion, remainder) = t
                      val ph = placer.height(portion, ctx)
                      placer.place(portion, slide, ppt, ctx, curY, tb)
                      curY += ph
                      remainder match
                        case Some(rem) => buf.update(idx, convertator.model.PageItem.ImageItem(rem))
                        case None      => idx += 1
                    case None =>
                      overflow = true; overflowTotal.incrementAndGet()
              else
                overflow = true; overflowTotal.incrementAndGet()
            else
              overflow = true; overflowTotal.incrementAndGet()

          case convertator.model.PageItem.TableItem(tbl) =>
            val placer = summon[SlidePlacer[PageTable]]
            val h = placer.height(tbl, ctx)
            val fits = curY + h <= ctx.usableH
            if tb == null then
              if fits then
                placer.place(tbl, slide, ppt, ctx, curY, tb)
                curY += h; idx += 1
              else if idx == 0 then
                  val splitTbl: Option[(PageTable, Option[PageTable])] = placer.trySplit(tbl, ctx, curY)
                  splitTbl match
                    case Some(t) =>
                      val (portion, remainder) = t
                      val ph = placer.height(portion, ctx)
                      placer.place(portion, slide, ppt, ctx, curY, tb)
                      curY += ph
                      remainder match
                        case Some(rem) => buf.update(idx, convertator.model.PageItem.TableItem(rem))
                        case None      => idx += 1
                    case None =>
                      overflow = true; overflowTotal.incrementAndGet()
              else
                overflow = true; overflowTotal.incrementAndGet()
            else
              overflow = true; overflowTotal.incrementAndGet()

      buf.remove(0, idx)

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
  def drawTable(
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

  def addLineAsParagraph(
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
