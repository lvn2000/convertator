package convertator.readers

import java.io.{File, InputStream}
import convertator.model.PageContent

/** Trait for document readers (PDF/DOCX). Allows isolating IO behind a
  * small interface for easier testing and replacement. */
trait DocumentReader:
  def read(input: File): Seq[PageContent]
  def read(input: InputStream): Seq[PageContent]
