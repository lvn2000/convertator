package convertator.docx

import convertator.model.{PageContent, PageImage, TextElement, TextLine}

import org.apache.poi.xwpf.usermodel.{UnderlinePatterns, XWPFDocument}

import java.io.{File, FileInputStream, InputStream}
import scala.jdk.CollectionConverters.*
import scala.collection.mutable.ListBuffer

object WordReader:

  private val PageWidth: Float  = 595.28f
  private val PageHeight: Float = 841.89f
  private val EmuToPt: Float    = 1f / 12700f

  def read(input: File): Seq[PageContent] =
    val fis = new FileInputStream(input)
    try read(fis)
    finally fis.close()

  def read(input: InputStream): Seq[PageContent] =
    val doc = new XWPFDocument(input)
    try extract(doc)
    finally doc.close()

  private def runHasPictures(run: org.apache.poi.xwpf.usermodel.XWPFRun): Boolean =
    try
      val ct = run.getCTR
      (ct.getDrawingList != null && !ct.getDrawingList.isEmpty) ||
      (ct.getPictList != null && !ct.getPictList.isEmpty)
    catch case _: Exception => false

  /** Extract images from a run's CT drawing XML via a:blip r:embed. */
  private def extractDrawingImages(
    run: org.apache.poi.xwpf.usermodel.XWPFRun,
    doc: XWPFDocument, yPos: Float,
    imgList: ListBuffer[PageImage]
  ): Unit =
    try
      val ct = run.getCTR
      if ct.getDrawingList == null then return
      val pkg = doc.getPackagePart
      for i <- 0 until ct.getDrawingList.size do
        val xml = ct.getDrawingList.get(i).toString
        for m <- """r:embed="([^"]+)"""".r.findAllMatchIn(xml) do
          try
            val rel = pkg.getRelationship(m.group(1))
            if rel != null then
              val part = pkg.getRelatedPart(rel)
              if part != null then
                val data = part.getInputStream.readAllBytes
                if data != null && data.nonEmpty then
                  val ctType = Option(part.getContentType).getOrElse("image/png")
                  val extRe = """<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*/>""".r
                  val sz = extRe.findFirstMatchIn(xml).orElse(
                    """cx="(\d+)" cy="(\d+)"""".r.findFirstMatchIn(xml))
                  val ew = sz.map(m => m.group(1).toFloat * EmuToPt).getOrElse(100f)
                  val eh = sz.map(m => m.group(2).toFloat * EmuToPt).getOrElse(100f)
                  val offRe = """<a:off\s+x="(\d+)"\s+y="(\d+)"\s*/>""".r
                  val off = offRe.findFirstMatchIn(xml)
                  imgList += PageImage(data, ctType,
                    off.map(m => m.group(1).toFloat * EmuToPt).getOrElse(0f),
                    yPos + off.map(m => m.group(2).toFloat * EmuToPt).getOrElse(0f),
                    ew, eh)
          catch case _: Exception => ()
    catch case _: Exception => ()

  /** Extract images from header/footer parts. */
  private def extractHFImages(
    doc: XWPFDocument, imgList: ListBuffer[PageImage]
  ): Unit =
    for hf <- doc.getHeaderList.asScala ++ doc.getFooterList.asScala do
      try
        val pkg = hf.getPackagePart
        for para <- hf.getParagraphs.asScala do
          for run <- para.getRuns.asScala do
            if runHasPictures(run) then
              for pic <- run.getEmbeddedPictures.asScala do
                try
                  val pd = pic.getPictureData
                  if pd != null && pd.getData != null && pd.getData.nonEmpty then
                    val xf = pic.getCTPicture.getSpPr.getXfrm
                    val ew = if xf != null && xf.getExt != null then xf.getExt.getCx.asInstanceOf[Long].toFloat * EmuToPt else 100f
                    val eh = if xf != null && xf.getExt != null then xf.getExt.getCy.asInstanceOf[Long].toFloat * EmuToPt else 100f
                    imgList += PageImage(pd.getData, pd.getPackagePart.getContentType, 0f, 0f, ew, eh)
                catch case _: Exception => ()
              // Also try low-level
              extractDrawingImages(run, doc, 0f, imgList)
      catch case _: Exception => ()

  private def extract(doc: XWPFDocument): Seq[PageContent] =
    val textLines = List.newBuilder[TextLine]
    val imgList   = ListBuffer.empty[PageImage]

    doc.getParagraphs.asScala.zipWithIndex.foreach { (para, lineIdx) =>
      val yPos = lineIdx.toFloat * 15f
      val runs = para.getRuns.asScala.toList

      var curX = 0f
      val elems = runs.flatMap { run =>
        if runHasPictures(run) then
          for pic <- run.getEmbeddedPictures.asScala do
            try
              val pd = pic.getPictureData
              if pd != null then
                val data = pd.getData
                if data != null && data.nonEmpty then
                  val ct = pd.getPackagePart.getContentType
                  val xf = pic.getCTPicture.getSpPr.getXfrm
                  imgList += PageImage(data, ct,
                    if xf != null && xf.getOff != null then xf.getOff.getX.asInstanceOf[Long].toFloat * EmuToPt else 0f,
                    yPos,
                    if xf != null && xf.getExt != null then xf.getExt.getCx.asInstanceOf[Long].toFloat * EmuToPt else 100f,
                    if xf != null && xf.getExt != null then xf.getExt.getCy.asInstanceOf[Long].toFloat * EmuToPt else 100f)
            catch case _: Exception => ()
          extractDrawingImages(run, doc, yPos, imgList)

        val text = run.getText(0)
        if text == null || text.isBlank then None
        else
          val rawFs = run.getFontSize
          val fs = if rawFs == -1 then 11f else rawFs.toFloat
          val fName = Option(run.getFontName).filter(_.nonEmpty).getOrElse("Calibri")
          val bold = run.isBold; val italic = run.isItalic
          val uline = run.getUnderline != UnderlinePatterns.NONE
          val elem = TextElement(text, curX, yPos, fs, fName,
            text.length * fs * 0.6f, bold, italic, uline)
          curX += text.length * fs * 0.6f + fs * 0.5f
          Some(elem)
      }

      if elems.nonEmpty then
        textLines += TextLine(elems, yPos, elems.map(_.fontSize).max)
      else
        textLines += TextLine(Nil, yPos, 11f)
    }

    // Also extract images from headers and footers
    extractHFImages(doc, imgList)

    Seq(PageContent(textLines.result(), PageWidth, PageHeight, imgList.toList))

end WordReader
