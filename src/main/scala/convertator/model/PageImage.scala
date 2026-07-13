package convertator.model

/** An image extracted from a source document, positioned on the page. */
case class PageImage(
  /** Raw image bytes (PNG, JPEG, etc.). */
  data: Array[Byte],
  /** MIME content type, e.g. "image/png", "image/jpeg". */
  contentType: String,
  /** X position on the source page in PDF user-space units (points). */
  x: Float,
  /** Y position on the source page. */
  y: Float,
  /** Display width in points. */
  width: Float,
  /** Display height in points. */
  height: Float
)
