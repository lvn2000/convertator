package convertator.pptx

import convertator.model.{ConversionConfig, PageContent, PageMode, TextElement, TextLine}

import org.apache.poi.xslf.usermodel.*
import org.apache.poi.sl.usermodel.TextParagraph as SLTextParagraph

import java.awt.Color
import java.io.{FileOutputStream, OutputStream}
import scala.collection.mutable.ListBuffer

/** Builds a PPTX presentation from a sequence of [[PageContent]] pages. */
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

  // ---- modes ---------------------------------------------------------------

  private def renderPage(
    ppt: XMLSlideShow, page: PageContent, cfg: ConversionConfig,
    overflowTotal: java.util.concurrent.atomic.AtomicInteger
  ): Unit =
    val availW = cfg.slideWidth - 2 * cfg.marginX
    val availH = cfg.slideHeight - 2 * cfg.marginY
    cfg.pageMode match
      case PageMode.Fit  => renderFit(ppt, page, cfg, availW, availH, overflowTotal)
      case PageMode.Flow => renderFlow(ppt, page, cfg, availW, availH, overflowTotal)

  private def renderFit(
    ppt: XMLSlideShow, page: PageContent, cfg: ConversionConfig,
    availW: Double, availH: Double,
    overflowTotal: java.util.concurrent.atomic.AtomicInteger
  ): Unit =
    val scaleX    = (availW / page.pageWidth).toFloat
    val scaleY    = (availH / page.pageHeight).toFloat
    val posScale  = math.min(scaleX, scaleY)
    val fontScale = posScale * cfg.fontSizeScale
    renderLines(ppt, page.lines, cfg, posScale, fontScale, availW, availH, overflowTotal)

  private def renderFlow(
    ppt: XMLSlideShow, page: PageContent, cfg: ConversionConfig,
    availW: Double, availH: Double,
    overflowTotal: java.util.concurrent.atomic.AtomicInteger
  ): Unit =
    val posScale  = (availW / page.pageWidth).toFloat
    val fontScale = cfg.fontSizeScale
    renderLines(ppt, page.lines, cfg, posScale, fontScale, availW, availH, overflowTotal)

  // ---- word grouping -------------------------------------------------------

  private def elemRight(e: TextElement): Float =
    e.x + (if e.width > 0f then e.width else e.text.length * e.fontSize * 0.55f)

  private def buildWords(elements: List[TextElement]): List[(String, TextElement)] =
    val sorted = elements.sortBy(_.x)
    val result = List.newBuilder[(String, TextElement)]
    var i = 0
    while i < sorted.length do
      val start = sorted(i)
      val sb    = new StringBuilder(start.text)
      val rep   = start
      var right = elemRight(start)
      var j     = i + 1
      var done  = false
      while j < sorted.length && !done do
        val next = sorted(j)
        val gap  = next.x - right
        val threshold = math.max(start.fontSize, next.fontSize) * 0.3f
        if gap <= threshold then
          if gap > 1.0f then sb.append(' ')
          sb.append(next.text)
          right = math.max(right, elemRight(next))
          j += 1
        else done = true
      result += ((sb.toString(), rep))
      i = j
    result.result()

  // ---- line rendering ------------------------------------------------------

  /**
   * Render lines using a conservative estimate of line height to determine
   * when content overflows. Each slide gets ONE text box with one paragraph
   * per PDF line. The estimate is intentionally generous so that text never
   * extends beyond the slide.
   *
   * Estimated line height = fontSize × 1.5 + lineSpacing × 2
   */
  private def renderLines(
    ppt: XMLSlideShow, lines: Seq[TextLine], cfg: ConversionConfig,
    posScale: Float, fontScale: Float,
    availW: Double, availH: Double,
    overflowTotal: java.util.concurrent.atomic.AtomicInteger
  ): Unit =
    val safetyMargin = 16.0
    val usableH      = availH - safetyMargin
    val remaining    = ListBuffer.from(lines)

    while remaining.nonEmpty do
      val slide = newSlide(ppt)
      val tb    = newTextBox(slide, cfg, availW, availH)
      var curY  = 0.0
      var idx   = 0
      var overflow = false

      while idx < remaining.length && !overflow do
        val fs    = remaining(idx).maxFontSize * fontScale
        val lineH = fs * 1.5 + cfg.lineSpacing * 2

        if curY + lineH > usableH then
          // Leave this line for the next slide
          overflow = true
          overflowTotal.incrementAndGet()
        else
          addLineAsParagraph(tb, remaining(idx), cfg, posScale, fontScale)
          curY += lineH
          idx += 1

      // Remove only the lines that were actually added (0 to idx-1)
      remaining.remove(0, idx)

  // ---- language -----------------------------------------------------------

  /** Set the default document language in the presentation XML so that
    * PowerPoint uses the correct proofing tools for the text. */
  private def setDocumentLanguage(ppt: XMLSlideShow, lang: String): Unit =
    val pres = ppt.getCTPresentation
    val dts  = if pres.isSetDefaultTextStyle then pres.getDefaultTextStyle
               else pres.addNewDefaultTextStyle()
    val defPPr = if dts.isSetDefPPr then dts.getDefPPr else dts.addNewDefPPr()
    val defRPr = if defPPr.isSetDefRPr then defPPr.getDefRPr else defPPr.addNewDefRPr()
    defRPr.setLang(lang)
    defRPr.setAltLang(lang)

  // -------------------------------------------------------------------------

  private def newSlide(ppt: XMLSlideShow): XSLFSlide =
    ppt.createSlide

  private def newTextBox(slide: XSLFSlide, cfg: ConversionConfig, availW: Double, availH: Double): XSLFTextBox =
    val tb = slide.createTextBox()
    tb.setShapeType(org.apache.poi.sl.usermodel.ShapeType.RECT)
    tb.setAnchor(new java.awt.Rectangle(
      cfg.marginX.toInt, cfg.marginY.toInt,
      math.max(1, availW.toInt),
      math.max(1, availH.toInt)
    ))
    tb.setWordWrap(true)
    // Set Ukrainian on the initial empty paragraph POI creates
    val firstParas = tb.getTextParagraphs
    if !firstParas.isEmpty then
      val ctPara = firstParas.get(0).getXmlObject
      val endRPr = if ctPara.isSetEndParaRPr then ctPara.getEndParaRPr else ctPara.addNewEndParaRPr()
      endRPr.setLang("uk-UA")
    tb

  private def addLineAsParagraph(
    tb: XSLFTextBox, line: TextLine, cfg: ConversionConfig,
    posScale: Float, fontScale: Float
  ): Unit =
    val words = buildWords(line.elements)
    if words.isEmpty then return

    val para = tb.addNewTextParagraph()
    para.setTextAlign(SLTextParagraph.TextAlign.LEFT)
    // Explicit spacing — eliminates PowerPoint's default paragraph gaps
    para.setSpaceBefore(0.0)
    para.setSpaceAfter(cfg.lineSpacing)

    var first = true
    for (wordText, rep) <- words do
      if !first then
        val sp = para.addNewTextRun()
        sp.setText(" ")
        sp.setFontSize(math.max(4.0, (rep.fontSize * fontScale).toDouble))
        sp.setFontColor(Color.BLACK)
        sp.setFontFamily(cleanFontName(rep.fontName))
        sp.getRPr(true).setLang("uk-UA")

      val run = para.addNewTextRun()
      run.setText(wordText)
      val outFs = math.max(4.0, (rep.fontSize * fontScale).toDouble)
      run.setFontSize(outFs)
      run.setFontColor(Color.BLACK)
      run.setFontFamily(cleanFontName(rep.fontName))
      if rep.bold then run.setBold(true)
      if rep.italic then run.setItalic(true)
      if cfg.preserveUnderline && rep.underline then run.setUnderlined(true)
      run.getRPr(true).setLang("uk-UA")
      first = false

    // Also set language on the end-of-paragraph marker (POI defaults to en-US)
    val ctPara = para.getXmlObject
    val endRPr = if ctPara.isSetEndParaRPr then ctPara.getEndParaRPr else ctPara.addNewEndParaRPr()
    endRPr.setLang("uk-UA")

  private def cleanFontName(raw: String): String =
    val cleaned = if raw.contains('+') then raw.substring(raw.indexOf('+') + 1) else raw
    cleaned.split("[-,\\.]").headOption.getOrElse(cleaned).trim

end SlideBuilder
