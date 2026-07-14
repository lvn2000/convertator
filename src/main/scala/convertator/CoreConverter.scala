package convertator

import convertator.model.ConversionConfig
import convertator.pdf.PdfReader
import convertator.docx.WordReader
import convertator.pptx.SlideBuilder
import convertator.model.PageContent

import java.io.File

/** Core conversion logic separated from CLI/IO concerns. Returns Either to
  * make failures explicit and easier to test. */
object CoreConverter:

  def convert(inputPath: String, pptxPath: String, cfg: ConversionConfig): Either[Throwable, Unit] =
    try
      val inputFile = new File(inputPath)
      if !inputFile.exists() || !inputFile.isFile then
        return Left(new IllegalArgumentException(s"Input file does not exist or is not a file: $inputPath"))

      val ext = inputPath.toLowerCase.takeRight(5).dropWhile(_ != '.')
      val pages: Seq[PageContent] = ext match
        case ".pdf"   => PdfReader.read(inputFile)
        case ".docx"  => WordReader.read(inputFile)
        case _         => return Left(new IllegalArgumentException(s"Unsupported format: $ext (use .pdf or .docx)"))

      // Compute effective font scale — targetFontSize overrides fontSizeScale
      val effectiveCfg = resolveFontScale(pages, cfg)

      SlideBuilder.convert(pages, effectiveCfg, pptxPath)
      Right(())
    catch
      case e: Throwable => Left(e)

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

end CoreConverter
