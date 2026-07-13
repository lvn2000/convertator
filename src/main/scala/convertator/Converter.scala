package convertator

import convertator.model.{ConversionConfig, PageContent, TextElement}
import convertator.docx.WordReader
import convertator.pdf.PdfReader
import convertator.pptx.SlideBuilder

import java.io.File

/** Top-level orchestrator for the conversion pipeline. */
object Converter:

  /**
   * Convert a document (PDF or DOCX) to a PPTX file.
   *
   * @param inputPath  path to the input file (.pdf or .docx)
   * @param pptxPath   path where the output PPTX will be written
   * @param config     optional conversion tuning
   */
  def run(inputPath: String, pptxPath: String, config: ConversionConfig = ConversionConfig()): Unit =
    val inputFile = new File(inputPath)
    if !inputFile.exists() || !inputFile.isFile then
      throw new IllegalArgumentException(s"Input file does not exist or is not a file: $inputPath")

    val ext = inputPath.toLowerCase.takeRight(5).dropWhile(_ != '.')
    val (readerLabel, pages) = ext match
      case ".pdf"   =>
        println(s"📖 Reading PDF: $inputPath")
        val p = PdfReader.read(inputFile)
        (s"page(s)", p)
      case ".docx"  =>
        println(s"📖 Reading Word: $inputPath")
        val p = WordReader.read(inputFile)
        (s"paragraph(s)", p)
      case _ =>
        throw new IllegalArgumentException(s"Unsupported format: $ext (use .pdf or .docx)")

    val paraCount = pages.flatMap(_.lines).length
    println(s"   → ${pages.length} section(s), $paraCount $readerLabel extracted")

    // Compute effective font scale — targetFontSize overrides fontSizeScale
    val effectiveCfg = resolveFontScale(pages, config)

    println(s"💼 Writing PPTX: $pptxPath")
    SlideBuilder.convert(pages, effectiveCfg, pptxPath)
    println(s"✅ Done — output written to $pptxPath")

  /** If [[ConversionConfig.targetFontSize]] is set, calculate the scale factor
    * that makes the most common ("body") font size match the target.
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
          println(s"   → Body font size in document: ${"%.1f".format(bodySize)}pt → target ${"%.0f".format(targetPt)}pt (scale=${"%.2f".format(scale)})")
          cfg.copy(fontSizeScale = scale, targetFontSize = None)

  private def findBodyFontSize(pages: Seq[PageContent]): Float =
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
