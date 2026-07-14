package convertator.docx

import convertator.model.{PageContent, PageImage, PageTable, TableCell, TableRow, TextElement, TextLine}

import org.apache.poi.xwpf.usermodel.{IBodyElement, UnderlinePatterns, XWPFDocument, XWPFParagraph, XWPFTable, XWPFTableRow}

import java.io.{File, FileInputStream, InputStream}
import scala.jdk.CollectionConverters.*
import scala.collection.mutable.ListBuffer

import convertator.readers.DocumentReader

object WordReader extends DocumentReader:

  private val PageWidth: Float  = 595.28f
  private val PageHeight: Float = 841.89f
  private val EmuToPt: Float    = 1f / 12700f
  private val TwipToPt: Float   = 1f / 20f     // 1 twip = 1/1440 inch = 1/20 pt

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

  // ---- paragraph helpers ---------------------------------------------------

  private def extractRunText(
    run: org.apache.poi.xwpf.usermodel.XWPFRun, curX: Float, yPos: Float
  ): Option[TextElement] =
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
      Some(elem)

  private def extractParagraphImages(
    run: org.apache.poi.xwpf.usermodel.XWPFRun,
    doc: XWPFDocument, yPos: Float,
    imgList: ListBuffer[PageImage]
  ): Unit =
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

  /** Extract text elements + images from a single XWPFParagraph. */
  private def extractParagraph(
    para: XWPFParagraph, doc: XWPFDocument, yPos: Float,
    imgList: ListBuffer[PageImage]
  ): Option[TextLine] =
    val runs = para.getRuns.asScala.toList
    var curX = 0f
    val elems = runs.flatMap { run =>
      extractParagraphImages(run, doc, yPos, imgList)
      val base = extractRunText(run, curX, yPos)
      base.foreach(e => curX += e.text.length * e.fontSize * 0.6f + e.fontSize * 0.5f)
      base
    }
    if elems.nonEmpty then Some(TextLine(elems, yPos, elems.map(_.fontSize).max))
    else Some(TextLine(Nil, yPos, 11f))

  // ---- table helpers -------------------------------------------------------

  /** Estimate row height based on cell content and font sizes. */
  private def estimateRowHeight(row: XWPFTableRow, colWidths: List[Float]): Float =
    var maxH = 0f
    for (cell, ci) <- row.getTableCells.asScala.zipWithIndex do
      val cellW = colWidths.lift(ci).getOrElse(200f)
      for para <- cell.getParagraphs.asScala do
        val runs = para.getRuns.asScala.toList
        val maxFs = runs.flatMap(r => Option(r.getFontSize).filter(_ > 0)).reduceOption(_ max _).getOrElse(11)
        val textLen = runs.flatMap(r => Option(r.getText(0))).map(_.length).sum
        // Estimate visual lines in the cell
        val avgCharW = (maxFs.toFloat * 0.55f).max(1f)
        val chPerLine = (cellW / avgCharW).toInt.max(1)
        val visLines = math.ceil(textLen.toDouble / chPerLine).toInt.max(1)
        val cellH = visLines * (maxFs.toFloat * 1.5f) + 8f  // 8pt padding
        maxH = math.max(maxH, cellH)
    math.max(maxH, 20f)  // at least 20pt

  /** Extract column widths from a DOCX table (in points). */
  private def extractColumnWidths(table: XWPFTable): List[Float] =
    try
      val grid = table.getCTTbl.getTblGrid
      if grid != null && grid.getGridColList != null then
        grid.getGridColList.asScala.map { gc =>
          val w = gc.getW
          if w != null then
            try w.toString.toFloat * TwipToPt
            catch case _: Exception => 100f
          else 100f
        }.toList
      else
        // Fallback: equal widths from cell count in first row
        val firstRow = table.getRows.asScala.headOption
        val n = firstRow.map(_.getTableCells.size).getOrElse(1)
        List.fill(n)(400f / n)
    catch case _: Exception =>
      val n = table.getRows.asScala.headOption.map(_.getTableCells.size).getOrElse(1)
      List.fill(n)(400f / n)

  /** Extract a single table cell's text elements. */
  private def extractCellElements(
    cell: org.apache.poi.xwpf.usermodel.XWPFTableCell, yPos: Float
  ): List[TextElement] =
    val elems = List.newBuilder[TextElement]
    var curX = 0f
    for para <- cell.getParagraphs.asScala do
      for run <- para.getRuns.asScala do
        val text = run.getText(0)
        if text != null && !text.isBlank then
          val rawFs = run.getFontSize
          val fs = if rawFs == -1 then 11f else rawFs.toFloat
          val fName = Option(run.getFontName).filter(_.nonEmpty).getOrElse("Calibri")
          val bold = run.isBold; val italic = run.isItalic
          val uline = run.getUnderline != UnderlinePatterns.NONE
          elems += TextElement(text, curX, yPos, fs, fName,
            text.length * fs * 0.6f, bold, italic, uline)
          curX += text.length * fs * 0.6f + fs * 0.5f
    elems.result()

  /** Extract images from a table cell (via its paragraphs). */
  private def extractCellImages(
    cell: org.apache.poi.xwpf.usermodel.XWPFTableCell,
    doc: XWPFDocument, yPos: Float,
    imgList: ListBuffer[PageImage]
  ): Unit =
    for para <- cell.getParagraphs.asScala do
      for run <- para.getRuns.asScala do
        extractParagraphImages(run, doc, yPos, imgList)

  /** Extract a full table from the DOCX. */
  private def extractTable(
    table: XWPFTable, doc: XWPFDocument, yPos: Float,
    imgList: ListBuffer[PageImage]
  ): PageTable =
    val colWidths = extractColumnWidths(table)
    val rows = List.newBuilder[TableRow]

    for tblRow <- table.getRows.asScala do
      val cells = List.newBuilder[TableCell]
      for (cell, colIdx) <- tblRow.getTableCells.asScala.zipWithIndex do
        val textElems = extractCellElements(cell, yPos)
        extractCellImages(cell, doc, yPos, imgList)
        val colW = colWidths.lift(colIdx).getOrElse(100f)
        cells += TableCell(textElems, colW)
      val rowH = estimateRowHeight(tblRow, colWidths)
      rows += TableRow(cells.result(), rowH)

    val rowData = rows.result()
    val heights = rowData.map(_.height)
    PageTable(yPos, colWidths, rowData)

  // ---- main extraction ----------------------------------------------------

  private def extract(doc: XWPFDocument): Seq[PageContent] =
    val textLines = List.newBuilder[TextLine]
    val imgList   = ListBuffer.empty[PageImage]
    val tableList = ListBuffer.empty[PageTable]
    var lineIdx   = 0

    doc.getBodyElements.asScala.foreach {
      case para: XWPFParagraph =>
        val yPos = lineIdx.toFloat * 15f
        extractParagraph(para, doc, yPos, imgList).foreach { tl =>
          textLines += tl
        }
        lineIdx += 1

      case tbl: XWPFTable =>
        val yPos = lineIdx.toFloat * 15f
        val pt = extractTable(tbl, doc, yPos, imgList)
        tableList += pt
        // Approximate the table's "line count" so following elements are
        // positioned correctly. Use at least 2 lines per row.
        val approxLines = math.max(2, pt.rows.length * 2)
        lineIdx += approxLines

      case _ =>
        lineIdx += 1
    }

    // Also extract images from headers and footers
    extractHFImages(doc, imgList)

    Seq(PageContent(textLines.result(), PageWidth, PageHeight,
      imgList.toList, tableList.result()))

end WordReader
