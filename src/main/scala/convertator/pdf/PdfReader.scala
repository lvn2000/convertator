package convertator.pdf

import convertator.model.{PageContent, TextElement, TextLine}

import org.apache.pdfbox.pdmodel.PDDocument
import org.apache.pdfbox.text.{PDFTextStripper, TextPosition}

import java.io.{File, InputStream}
import scala.collection.mutable
import scala.jdk.CollectionConverters.*

/** Reads a PDF and extracts text elements with positioning per page. */
object PdfReader:

  /** Extract content from every page of a PDF file. */
  def read(input: File): Seq[PageContent] =
    val doc = PDDocument.load(input)
    try readDoc(doc)
    finally doc.close()

  /** Extract content from every page of a PDF input stream. */
  def read(input: InputStream): Seq[PageContent] =
    val doc = PDDocument.load(input)
    try readDoc(doc)
    finally doc.close()

  private def readDoc(doc: PDDocument): Seq[PageContent] =
    val pageCount = doc.getNumberOfPages
    (0 until pageCount).map { pageIdx =>
      extractPage(doc, pageIdx)
    }

  /** Extract one page at the given 0-based index. */
  private def extractPage(doc: PDDocument, pageIdx: Int): PageContent =
    val page  = doc.getPage(pageIdx)
    val media = page.getMediaBox
    val pageW = media.getWidth
    val pageH = media.getHeight

    val stripper = new PositionAwareStripper
    stripper.setStartPage(pageIdx + 1)
    stripper.setEndPage(pageIdx + 1)
    stripper.setSortByPosition(true) // left-to-right, top-to-bottom

    // Suppress the extracted text string (we only care about positions)
    val _ = stripper.getText(doc)

    val raw = stripper.elements.toList

    // Group elements into lines by their Y coordinate (allow ~3pt tolerance)
    val grouped: List[List[TextElement]] =
      groupByY(raw, tolerance = 3.0f)

    val lines = grouped.map { elems =>
      val maxFs = elems.map(_.fontSize).max
      // The representative Y is the average (or first element's Y)
      val avgY  = elems.map(_.y).sum / elems.length
      TextLine(elems, avgY, maxFs)
    }

    PageContent(lines, pageW, pageH)

  // ---- helpers -----------------------------------------------------------

  private def groupByY(elems: List[TextElement], tolerance: Float): List[List[TextElement]] =
    if elems.isEmpty then Nil
    else
      val sorted = elems.sortBy(_.y)
      val groups = mutable.ListBuffer.empty[List[TextElement]]
      var current = mutable.ListBuffer(sorted.head)
      for e <- sorted.tail do
        // same line if y difference is within tolerance
        if math.abs(e.y - current.head.y) <= tolerance then
          current += e
        else
          groups += current.sortBy(_.x).toList
          current = mutable.ListBuffer(e)
      groups += current.sortBy(_.x).toList
      groups.toList

  // ---- custom text stripper -----------------------------------------------

  private class PositionAwareStripper extends PDFTextStripper:
    val elements = mutable.ListBuffer.empty[TextElement]

    // Called by PDFBox for each text-position (word / fragment)
    override def writeString(text: String, textPositions: java.util.List[TextPosition]): Unit =
      for tp <- textPositions.asScala do
        val txt = tp.getUnicode
        if txt != null && txt.nonEmpty then
          val fontName = Option(tp.getFont).map(_.getName).getOrElse("Unknown")
          elements += TextElement(
            text     = txt,
            x        = tp.getXDirAdj,
            y        = tp.getYDirAdj,
            fontSize = tp.getFontSizeInPt,
            fontName = fontName,
            width    = tp.getWidth,
            bold     = isBold(fontName),
            italic   = isItalic(fontName),
            underline = false // PDF underline detection is unreliable this way
          )

    private def isBold(name: String): Boolean =
      val upper = name.toUpperCase
      upper.contains("BOLD") || upper.contains("DEMI") || upper.contains("HEAVY")

    private def isItalic(name: String): Boolean =
      val upper = name.toUpperCase
      upper.contains("ITALIC") || upper.contains("OBLIQUE") || upper.contains("KURSIV")
  end PositionAwareStripper

end PdfReader
