package convertator.model

/** Content extracted from a single PDF page or document section. */
case class PageContent(
  lines: List[TextLine],
  pageWidth: Float,
  pageHeight: Float,
  images: List[PageImage] = Nil
)
