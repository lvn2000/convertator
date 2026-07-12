name := "convertator"
version := "0.1.0"
scalaVersion := "3.3.4"

libraryDependencies ++= Seq(
  // Apache PDFBox - read PDF files and extract text with positioning
  "org.apache.pdfbox" % "pdfbox" % "2.0.31",
  // Apache POI - create PowerPoint PPTX files
  "org.apache.poi"  % "poi-ooxml"  % "5.2.5"
)

// Package a fat JAR so the app runs standalone with `java -jar`
assembly / mainClass := Some("convertator.Main")
assembly / assemblyMergeStrategy := {
  case PathList("META-INF", xs @ _*) => MergeStrategy.discard
  case _                             => MergeStrategy.first
}
