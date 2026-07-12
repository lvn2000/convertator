package convertator

import convertator.model.{ConversionConfig, PageContent, TextElement}
import convertator.pdf.PdfReader
import convertator.pptx.SlideBuilder

import java.io.File

/** Top-level orchestrator for the PDF → PPTX conversion pipeline. */
object Converter:

  /**
   * Convert a PDF file to a PPTX file.
   *
   * @param pdfPath  path to the input PDF
   * @param pptxPath path where the output PPTX will be written
   * @param config   optional conversion tuning (default values are used if omitted)
   */
  def run(pdfPath: String, pptxPath: String, config: ConversionConfig = ConversionConfig()): Unit =
    val pdfFile = new File(pdfPath)
    if !pdfFile.exists() || !pdfFile.isFile then
      throw new IllegalArgumentException(s"Input PDF does not exist or is not a file: $pdfPath")

    println(s"📖 Reading PDF: $pdfPath")
    val pages = PdfReader.read(pdfFile)
    println(s"   → ${pages.length} page(s) extracted")

    // Compute effective font scale — targetFontSize overrides fontSizeScale
    val effectiveCfg = resolveFontScale(pages, config)

    println(s"💼 Writing PPTX: $pptxPath")
    SlideBuilder.convert(pages, effectiveCfg, pptxPath)
    println(s"✅ Done — output written to $pptxPath")

  /** If [[ConversionConfig.targetFontSize]] is set, calculate the scale factor
    * that makes the most common ("body") font size match the target.
    * Otherwise return the config unchanged.
    */
  private def resolveFontScale(pages: Seq[PageContent], cfg: ConversionConfig): ConversionConfig =
    cfg.targetFontSize match
      case None => cfg
      case Some(targetPt) =>
        val bodySize = findBodyFontSize(pages)
        if bodySize <= 0f then
          Console.err.println(s"⚠  Could not determine body font size, falling back to fontSizeScale=${cfg.fontSizeScale}")
          cfg
        else
          val scale = targetPt / bodySize
          println(s"   → Body font size in PDF: ${"%.1f".format(bodySize)}pt → target ${"%.0f".format(targetPt)}pt (scale=${"%.2f".format(scale)})")
          cfg.copy(fontSizeScale = scale, targetFontSize = None)

  /** Find the most common font size by weighing each element's font size
    * by its character count (the "body text" size).
    */
  private def findBodyFontSize(pages: Seq[PageContent]): Float =
    // Map fontSize → total characters
    val tally = scala.collection.mutable.Map.empty[Float, Int].withDefaultValue(0)
    for
      page <- pages
      line <- page.lines
      elem <- line.elements
    do
      tally(elem.fontSize) += elem.text.length

    if tally.isEmpty then 0f
    else tally.maxBy(_._2)._1

end Converter
