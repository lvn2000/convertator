package convertator.pdf

import convertator.model.{PageContent, PageImage, TextElement, TextLine}

import org.apache.pdfbox.contentstream.operator.Operator
import org.apache.pdfbox.cos.COSName
import org.apache.pdfbox.pdmodel.PDDocument
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject
import org.apache.pdfbox.text.{PDFTextStripper, TextPosition}

import java.io.{ByteArrayOutputStream, File, InputStream}
import java.util.Base64
import javax.imageio.ImageIO
import scala.collection.mutable
import scala.jdk.CollectionConverters.*

/** Reads a PDF and extracts text elements with positioning per page. */
import convertator.readers.DocumentReader

object PdfReader extends DocumentReader:

  def read(input: File): Seq[PageContent] =
    val doc = PDDocument.load(input)
    try readDoc(doc)
    finally doc.close()

  def read(input: InputStream): Seq[PageContent] =
    val doc = PDDocument.load(input)
    try readDoc(doc)
    finally doc.close()

  private def readDoc(doc: PDDocument): Seq[PageContent] =
    val pageCount = doc.getNumberOfPages
    (0 until pageCount).map { pageIdx =>
      extractPage(doc, pageIdx)
    }

  private def extractPage(doc: PDDocument, pageIdx: Int): PageContent =
    val page  = doc.getPage(pageIdx)
    val media = page.getMediaBox
    val pageW = media.getWidth
    val pageH = media.getHeight

    // 1. Extract text
    val stripper = new PositionAwareStripper
    stripper.setStartPage(pageIdx + 1)
    stripper.setEndPage(pageIdx + 1)
    stripper.setSortByPosition(true)
    val _ = stripper.getText(doc)

    val raw   = stripper.elements.toList
    val groups = groupByY(raw, tolerance = 3.0f)
    val lines = groups.map { elems =>
      val maxFs = elems.map(_.fontSize).max
      val avgY  = elems.map(_.y).sum / elems.length
      TextLine(elems, avgY, maxFs)
    }

    // 2. Extract images
    val imgExtractor = new ImageExtractor
    imgExtractor.processPage(page)
    val images = imgExtractor.images.toList

    PageContent(lines, pageW, pageH, images)

  // ---- helpers -----------------------------------------------------------

  private def groupByY(elems: List[TextElement], tolerance: Float): List[List[TextElement]] =
    if elems.isEmpty then Nil
    else
      val sorted = elems.sortBy(_.y)
      val groups = mutable.ListBuffer.empty[List[TextElement]]
      var current = mutable.ListBuffer(sorted.head)
      for e <- sorted.tail do
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

    override def writeString(text: String, textPositions: java.util.List[TextPosition]): Unit =
      for tp <- textPositions.asScala do
        val txt = tp.getUnicode
        if txt != null && txt.nonEmpty then
          val fontName = Option(tp.getFont).map(_.getName).getOrElse("Unknown")
          elements += TextElement(
            text      = txt,
            x         = tp.getXDirAdj,
            y         = tp.getYDirAdj,
            fontSize  = tp.getFontSizeInPt,
            fontName  = fontName,
            width     = tp.getWidth,
            bold      = isBold(fontName),
            italic    = isItalic(fontName),
            underline = false
          )

    private def isBold(name: String): Boolean =
      val upper = name.toUpperCase
      upper.contains("BOLD") || upper.contains("DEMI") || upper.contains("HEAVY")
    private def isItalic(name: String): Boolean =
      val upper = name.toUpperCase
      upper.contains("ITALIC") || upper.contains("OBLIQUE") || upper.contains("KURSIV")
  end PositionAwareStripper

  // ---- image extractor ---------------------------------------------------

  /** Extracts images from a PDF page by scanning the page resources. */
  private class ImageExtractor:
    val images = mutable.ListBuffer.empty[PageImage]

    def processPage(page: org.apache.pdfbox.pdmodel.PDPage): Unit =
      try
        val resources = page.getResources
        if resources != null then
          for name <- resources.getXObjectNames.asScala do
            try
              resources.getXObject(name) match
                case img: PDImageXObject =>
                  val baos = new ByteArrayOutputStream()
                  ImageIO.write(img.getImage, "png", baos)
                  images += PageImage(
                    data        = baos.toByteArray,
                    contentType = "image/png",
                    x           = 0f,
                    y           = 0f,
                    width       = img.getWidth.toFloat,
                    height      = img.getHeight.toFloat
                  )
                case _ => ()
            catch case _: Exception => ()
      catch case _: Exception => ()
  end ImageExtractor

end PdfReader
