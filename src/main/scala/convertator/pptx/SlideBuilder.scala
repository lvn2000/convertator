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
        println(s"  \u2139  Content split across $total additional slide(s)")
      ppt.write(output)
    finally ppt.close()

  def convert(pages: Seq[PageContent], config: ConversionConfig, outputPath: String): Unit =
    val fos = new FileOutputStream(outputPath)
    try convert(pages, config, fos)
    finally fos.close()

  // ---- given instances ----------------------------------------------------

  private given SlidePlacer[TextLine] with
    def height(line: TextLine, ctx: PlaceContext): Double =
      val fs = line.maxFontSize * ctx.fontScale
      val avgCharW = (fs * 0.65).max(1.0)
      val chPerLine = (ctx.availW / avgCharW).toInt.max(1)
      val approxSpaces = line.elements.length
      val totalChars = line.elements.map(_.text.length).sum + approxSpaces
      val visLines = math.ceil(totalChars.toDouble / chPerLine).toInt.max(1)
      visLines * (fs * 1.2) + ctx.cfg.lineSpacing

    def place(line: TextLine, slide: XSLFSlide, ppt: XMLSlideShow,
              ctx: PlaceContext, curY: Double, tb: XSLFTextBox): XSLFTextBox =
      addLineAsParagraph(tb, line, ctx.cfg, ctx.posScale, ctx.fontScale)
      tb

    def trySplit(line: TextLine, ctx: PlaceContext, curY: Double): Option[(TextLine, Option[TextLine])] =
      val fs = line.maxFontSize * ctx.fontScale
      val avgCharW = (fs * 0.65).max(1.0)
      val chPerLine = (ctx.availW / avgCharW).toInt.max(1)
      val spaceLeft = ctx.usableH - curY
      val linesThatFit = (spaceLeft / (fs * 1.2)).toInt.max(1)
      val charsThatFit = linesThatFit * chPerLine
      val (fitElems, remainElems) = splitElements(line.elements, charsThatFit)
      if fitElems.isEmpty then None
      else
        val portion = TextLine(fitElems, line.y, line.maxFontSize)
        val remainder = if remainElems.nonEmpty then Some(TextLine(remainElems, line.y, line.maxFontSize)) else None
        Some((portion, remainder))

  private given SlidePlacer[PageImage] with
    def height(img: PageImage, ctx: PlaceContext): Double =
      var iw = img.width.toDouble * ctx.posScale
      var ih = img.height.toDouble * ctx.posScale
      if iw > ctx.availW then
        val s = ctx.availW / iw; ih = ih * s
      if ih > ctx.maxImgH then
        val s = ctx.maxImgH / ih; ih = ctx.maxImgH
      ih

    def place(img: PageImage, slide: XSLFSlide, ppt: XMLSlideShow,
              ctx: PlaceContext, curY: Double, tb: XSLFTextBox): XSLFTextBox =
      var iw = img.width.toDouble * ctx.posScale
      var ih = img.height.toDouble * ctx.posScale
      if iw > ctx.availW then
        val s = ctx.availW / iw; iw = ctx.availW; ih = ih * s
      if ih > ctx.maxImgH then
        val s = ctx.maxImgH / ih; ih = ctx.maxImgH; iw = iw * s
      try
        val pt = mimeToPictureType(img.contentType)
        val pd = ppt.addPicture(img.data, pt)
        slide.createPicture(pd).setAnchor(new Rectangle2D.Double(
          ctx.cfg.marginX + (ctx.availW - iw) / 2, ctx.cfg.marginY + curY, iw, ih))
      catch case e: Exception => Console.err.println(s"  \u26a0 Image: ${e.getMessage}")
      tb

    def trySplit(img: PageImage, ctx: PlaceContext, curY: Double): None.type = None

  private given SlidePlacer[PageTable] with
    def height(tbl: PageTable, ctx: PlaceContext): Double =
      val tw = tbl.columnWidths.sum.toDouble
      val scaleW = if tw > ctx.availW then ctx.availW / tw else 1.0
      tbl.rows.map(r => r.height.toDouble * scaleW).sum

    def place(tbl: PageTable, slide: XSLFSlide, ppt: XMLSlideShow,
              ctx: PlaceContext, curY: Double, tb: XSLFTextBox): XSLFTextBox =
      val tw = tbl.columnWidths.sum.toDouble
      val scaleW = if tw > ctx.availW then ctx.availW / tw else 1.0
      try
        drawTable(slide, tbl, ctx.cfg, ctx.availW, scaleW, curY, 0, tbl.rows.length)
      catch case e: Exception => Console.err.println(s"  \u26a0 Table: ${e.getMessage}")
      tb

    def trySplit(tbl: PageTable, ctx: PlaceContext, curY: Double): Option[(PageTable, Option[PageTable])] =
      val tw = tbl.columnWidths.sum.toDouble
      val scaleW = if tw > ctx.availW then ctx.availW / tw else 1.0
      val scaledHeights = tbl.rows.map(r => r.height.toDouble * scaleW)
      val spaceLeft = ctx.usableH - curY
      var accH = 0.0
      var splitIdx = 0
      while splitIdx < scaledHeights.length && accH + scaledHeights(splitIdx) <= spaceLeft do
        accH += scaledHeights(splitIdx)
        splitIdx += 1
      if splitIdx == 0 then None  // even one row doesn't fit — overflow
      else
        val portion = PageTable(tbl.y, tbl.columnWidths, tbl.rows.take(splitIdx))
        val remainder =
          if splitIdx < tbl.rows.length then
            Some(PageTable(tbl.y, tbl.columnWidths, tbl.rows.drop(splitIdx)))
          else None
        Some((portion, remainder))

  // ---- rendering ---------------------------------------------------------

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

    // Build sorted list of content items using the Positioned type class
    val items = ListBuffer.empty[TextLine | PageImage | PageTable]
    items.addAll(lines)
    items.addAll(images)
    items.addAll(tables)
    // Sort using explicit pattern matching to avoid Ordering ambiguity on union types
    def sortKey(c: TextLine | PageImage | PageTable): (Float, Int) =
      c match
        case l: TextLine  => (if l.y < 0f then Float.MaxValue else l.y, 1)
        case i: PageImage => (if i.y < 0f then Float.MaxValue else i.y, 0)
        case t: PageTable => (if t.y < 0f then Float.MaxValue else t.y, 0)
    items.sortBy(sortKey)

    while items.nonEmpty do
      val slide = newSlide(ppt)
      var tb: XSLFTextBox = null  // created lazily when first text item is reached
      var curY  = 0.0
      var idx   = 0
      var overflow = false

      // Shared logic for images and tables
      def processNonText[A <: (PageImage | PageTable)](a: A, placer: SlidePlacer[A],
        slide: XSLFSlide, ppt: XMLSlideShow, ctx: PlaceContext
      ): Unit =
        val h = placer.height(a, ctx)
        val fits = curY + h <= ctx.usableH
        if tb == null then
          if fits then
            placer.place(a, slide, ppt, ctx, curY, tb)
            curY += h; idx += 1
          else if idx == 0 then
            placer.trySplit(a, ctx, curY) match
              case Some((portion, remainder)) =>
                val ph = placer.height(portion, ctx)
                placer.place(portion, slide, ppt, ctx, curY, tb)
                curY += ph
                remainder match
                  case Some(rem) => items.update(idx, rem)
                  case None      => idx += 1
              case None =>
                overflow = true; overflowTotal.incrementAndGet()
          else
            overflow = true; overflowTotal.incrementAndGet()
        else
          overflow = true
          overflowTotal.incrementAndGet()

      while idx < items.length && !overflow do
        items(idx) match
          case line: TextLine =>
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
              placer.trySplit(line, ctx, curY) match
                case Some((portion, remainder)) =>
                  val ph = summon[SlidePlacer[TextLine]].height(portion, ctx)
                  summon[SlidePlacer[TextLine]].place(portion, slide, ppt, ctx, curY, tb)
                  curY += ph
                  remainder match
                    case Some(rem) => items.update(idx, rem)
                    case None      => idx += 1
                case None =>
                  overflow = true; overflowTotal.incrementAndGet()
            else
              overflow = true; overflowTotal.incrementAndGet()

          case img: PageImage =>
            processNonText(img, summon[SlidePlacer[PageImage]], slide, ppt, ctx)

          case tbl: PageTable =>
            processNonText(tbl, summon[SlidePlacer[PageTable]], slide, ppt, ctx)

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
