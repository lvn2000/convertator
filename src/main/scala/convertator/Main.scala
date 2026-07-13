package convertator

import convertator.model.{ConversionConfig, PageMode}

import java.io.File
import scala.util.Using

/** CLI entry point. */
object Main:

  private val SourceDir = "src" + File.separator + "main" + File.separator + "resources" + File.separator + "source"
  private val OutputDir = "src" + File.separator + "main" + File.separator + "resources" + File.separator + "output"

  def main(args: Array[String]): Unit =
    val cfg      = parseConfig(args)
    val modeArgs = cfg._1
    val userCfg  = cfg._2

    modeArgs match
      case "--help" :: _ | Nil =>
        println(
          s"""Convertator — PDF → PPTX Converter
             |
             |Usage:
             |  convertator [options]                Batch: process all PDFs/DOCX from $SourceDir
             |  convertator --batch [options]         Same as above (explicit)
             |  convertator <input.pdf|docx> <output.pptx> [options]
             |
             |Options:
             |  --help               Show this help
             |  --mode         <mode>    Page mode: "flow" (default) or "fit"
             |  --font-size    <pts>    Target body text size in pt (default: 18)
             |  --font-scale   <num>     Font-size scaling ratio     (default: auto from --font-size)
             |  --slide-width  <pts>     Slide width in points  (default: 720  = 10")
             |  --slide-height <pts>     Slide height in points (default: 540  = 7.5")
             |  --margin-x     <pts>     Horizontal margin      (default: 48)
             |  --margin-y     <pts>     Vertical margin        (default: 48)
             |  --line-spacing <pts>     Extra gap between lines (default: 4)
             |
             |Modes:
             |  flow  – Keep natural font sizes, overflow across multiple slides (default)
             |  fit   – Scale each PDF page to fit on one slide (1:1 mapping)
             |
             |Supported formats: .pdf, .docx
             |
             |Examples:
             |  # Batch (default when no args given): place files in $SourceDir, get PPTXs in $OutputDir
             |  convertator
             |  convertator --mode flow
             |
             |  # Single file
             |  convertator mydoc.pdf mydoc.pptx
             |  convertator mydoc.docx mydoc.pptx
             |""".stripMargin)
        if modeArgs == Nil then runBatch(userCfg) else sys.exit(1)

      case "--batch" :: rest =>
        runBatch(userCfg)

      case pdfPath :: pptxPath :: Nil =>
        runSingle(pdfPath, pptxPath, userCfg)

      case _ =>
        Console.err.println("❌ Invalid arguments. Use --help to see usage.")
        sys.exit(1)

  // ---- batch mode ----------------------------------------------------------

  private def runBatch(cfg: ConversionConfig): Unit =
    val srcDir = new File(SourceDir)
    if !srcDir.exists() then
      Console.err.println(s"❌ Source directory does not exist: $SourceDir")
      sys.exit(2)

    val outDir = new File(OutputDir)
    outDir.mkdirs()

    val inputFiles = srcDir.listFiles().filter { f =>
      f.isFile && {
        val name = f.getName.toLowerCase
        name.endsWith(".pdf") || name.endsWith(".docx")
      }
    }.toSeq

    if inputFiles.isEmpty then
      println(s"ℹ  No PDF or DOCX files found in $SourceDir")
      return

    println(s"📂 Found ${inputFiles.length} file(s) in $SourceDir")
    println()

    var success = 0
    var failed  = 0

    for inputFile <- inputFiles.sortBy(_.getName) do
      val pptxName = inputFile.getName.replaceAll("\\.(pdf|docx)$", ".pptx")
      val pptxFile = new File(outDir, pptxName)

      print(s"  🔄 ${inputFile.getName} → $OutputDir/$pptxName ... ")
      try
        Converter.run(inputFile.getAbsolutePath, pptxFile.getAbsolutePath, cfg)
        println("✅")
        success += 1
      catch
        case e: Exception =>
          println(s"❌  ${e.getMessage}")
          failed += 1

    println()
    println(s"✅ Done — $success succeeded, $failed failed")

  // ---- single-file mode ----------------------------------------------------

  private def runSingle(pdfPath: String, pptxPath: String, cfg: ConversionConfig): Unit =
    try
      Converter.run(pdfPath, pptxPath, cfg)
    catch
      case e: Exception =>
        Console.err.println(s"❌ Conversion failed: ${e.getMessage}")
        e.printStackTrace(Console.err)
        sys.exit(2)

  // ---- options parser -------------------------------------------------------

  private def parseConfig(args: Array[String]): (List[String], ConversionConfig) =
    var cfg   = ConversionConfig()
    var ahead = List.empty[String]

    val it = args.iterator
    while it.hasNext do
      it.next() match
        case "--batch"        => ahead = "--batch" :: Nil
        case "--mode"         => cfg = cfg.copy(pageMode     = parseMode(it.next()))
        case "--slide-width"  => cfg = cfg.copy(slideWidth   = it.next().toDouble)
        case "--slide-height" => cfg = cfg.copy(slideHeight  = it.next().toDouble)
        case "--margin-x"     => cfg = cfg.copy(marginX      = it.next().toDouble)
        case "--margin-y"     => cfg = cfg.copy(marginY      = it.next().toDouble)
        case "--line-spacing" => cfg = cfg.copy(lineSpacing  = it.next().toDouble)
        case "--font-scale"   => cfg = cfg.copy(fontSizeScale   = it.next().toFloat)
        case "--font-size"    => cfg = cfg.copy(targetFontSize = Some(it.next().toFloat))
        case unknown if !unknown.startsWith("--") =>
          ahead = ahead :+ unknown
        case unknown =>
          Console.err.println(s"⚠  Ignoring unknown option: $unknown")

    (ahead, cfg)

  private def parseMode(s: String): PageMode =
    s.toLowerCase match
      case "fit"  => PageMode.Fit
      case "flow" => PageMode.Flow
      case other =>
        Console.err.println(s"⚠  Unknown mode '$other', falling back to 'fit'")
        PageMode.Fit

end Main
