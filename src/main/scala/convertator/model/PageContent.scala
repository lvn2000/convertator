package convertator.model

/** Content extracted from a single PDF page. */
case class PageContent(
  lines: List[TextLine],
  pageWidth: Float,
  pageHeight: Float
)
