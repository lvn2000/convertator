package convertator.docx

import convertator.model.{PageContent, TextElement, TextLine}

import org.apache.poi.xwpf.usermodel.{UnderlinePatterns, XWPFDocument}

import java.io.{File, FileInputStream, InputStream}
import scala.jdk.CollectionConverters.*

/** Reads a Word (.docx) file and extracts text elements with styling per paragraph. */
object WordReader:

  // Standard page dimensions for layout calculations (A4 in points)
  private val PageWidth: Float  = 595.28f
  private val PageHeight: Float = 841.89f

  /** Read content from a .docx file. The whole document is treated as one
    * continuous page — the slide splitter handles overflow to new slides. */
  def read(input: File): Seq[PageContent] =
    val fis = new FileInputStream(input)
    try read(fis)
    finally fis.close()

  def read(input: InputStream): Seq[PageContent] =
    val doc = new XWPFDocument(input)
    try extract(doc)
    finally doc.close()

  private val TwipToPt: Float = 1f / 20f // 1 Twip = 1/20 pt

  private def extract(doc: XWPFDocument): Seq[PageContent] =
    val lines = doc.getParagraphs.asScala.zipWithIndex.flatMap { (para, lineIdx) =>
      val yPos = lineIdx.toFloat * 15f
      val leftIndent =
        Option(para.getIndentationLeft).map(_.toFloat * TwipToPt * 0f).getOrElse(0f)

      val runs = para.getRuns.asScala.toList
      if runs.isEmpty then
        Some(TextLine(Nil, y = yPos, maxFontSize = 11f))
      else
        // Give each run a staggered X so buildWords doesn't merge them
        var curX = leftIndent
        val elements = runs.flatMap { run =>
          val text = run.getText(0)
          if text == null || text.isBlank then None
          else
            val rawFs  = run.getFontSize
            val fs     = if rawFs == -1 then 11f else rawFs.toFloat // already in points
            val fName  = Option(run.getFontName).filter(_.nonEmpty).getOrElse("Calibri")
            val bold   = run.isBold
            val italic = run.isItalic
            val uline  = run.getUnderline != UnderlinePatterns.NONE

            val elem = Some(TextElement(
              text     = text,
              x        = curX,
              y        = yPos,
              fontSize = fs,
              fontName = fName,
              width    = text.length * fs * 0.6f,
              bold     = bold,
              italic   = italic,
              underline = uline
            ))
            curX += text.length * fs * 0.6f + fs * 0.5f // stagger > buildWords threshold
            elem
        }

        if elements.isEmpty then None
        else
          val maxFs = elements.map(_.fontSize).max
          Some(TextLine(elements.toList, y = yPos, maxFontSize = maxFs))
    }

    Seq(PageContent(lines.toList, PageWidth, PageHeight))

end WordReader
