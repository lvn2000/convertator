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
    println(s"💼 Writing PPTX: $pptxPath")
    CoreConverter.convert(inputPath, pptxPath, config) match
      case Right(_) => println(s"✅ Done — output written to $pptxPath")
      case Left(e)  => throw e

end Converter
