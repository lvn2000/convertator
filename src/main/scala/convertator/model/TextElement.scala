package convertator.model

/** A single positioned text fragment extracted from a PDF. */
case class TextElement(
  text: String,
  x: Float,
  y: Float,
  fontSize: Float,
  fontName: String,
  /** Width of this fragment in PDF user-space units (0 if unknown). */
  width: Float = 0f,
  bold: Boolean = false,
  italic: Boolean = false,
  underline: Boolean = false
)
