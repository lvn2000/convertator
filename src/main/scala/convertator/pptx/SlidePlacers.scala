package convertator.pptx

import convertator.model.{ConversionConfig, PageContent, PageImage, PageMode, PageTable, TextElement, TextLine}

import org.apache.poi.xslf.usermodel.*
import org.apache.poi.sl.usermodel.{PictureData, TextParagraph as SLTextParagraph}
import org.apache.poi.sl.usermodel.TableCell.BorderEdge

import java.awt.Color
import java.awt.geom.Rectangle2D
import scala.collection.mutable.ListBuffer

/** Given instances for SlidePlacer for the supported page content types.
  * These were previously defined inline within SlideBuilder; extracting them
  * keeps SlideBuilder focused on orchestration.
  */
object SlidePlacers:

  private val MaxImgH = 0.35

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

  // ---- SlidePlacer instances -------------------------------------------------

  given SlidePlacer[TextLine] with
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
      SlideBuilder.addLineAsParagraph(tb, line, ctx.cfg, ctx.posScale, ctx.fontScale)
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

  given SlidePlacer[PageImage] with
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
        val pd = org.apache.poi.xslf.usermodel.XMLSlideShow.SYSTEM_PICTURE_FACTORY.addPicture(img.data, pt)
        // Use ppt.addPicture instead of custom factory where possible
        val pd2 = ppt.addPicture(img.data, pt)
        slide.createPicture(pd2).setAnchor(new Rectangle2D.Double(
          ctx.cfg.marginX + (ctx.availW - iw) / 2, ctx.cfg.marginY + curY, iw, ih))
      catch case e: Exception => Console.err.println(s"  \u26a0 Image: ${e.getMessage}")
      tb

    def trySplit(img: PageImage, ctx: PlaceContext, curY: Double): None.type = None

  given SlidePlacer[PageTable] with
    def height(tbl: PageTable, ctx: PlaceContext): Double =
      val tw = tbl.columnWidths.sum.toDouble
      val scaleW = if tw > ctx.availW then ctx.availW / tw else 1.0
      tbl.rows.map(r => r.height.toDouble * scaleW).sum

    def place(tbl: PageTable, slide: XSLFSlide, ppt: XMLSlideShow,
              ctx: PlaceContext, curY: Double, tb: XSLFTextBox): XSLFTextBox =
      val tw = tbl.columnWidths.sum.toDouble
      val scaleW = if tw > ctx.availW then ctx.availW / tw else 1.0
      try
        SlideBuilder.drawTable(slide, tbl, ctx.cfg, ctx.availW, scaleW, curY, 0, tbl.rows.length)
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
      if splitIdx == 0 then None
      else
        val portion = PageTable(tbl.y, tbl.columnWidths, tbl.rows.take(splitIdx))
        val remainder = if splitIdx < tbl.rows.length then Some(PageTable(tbl.y, tbl.columnWidths, tbl.rows.drop(splitIdx))) else None
        Some((portion, remainder))

end SlidePlacers
