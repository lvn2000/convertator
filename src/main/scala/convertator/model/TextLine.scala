package convertator.model

/** One line of text made up of one or more [[TextElement]]s (words / fragments). */
case class TextLine(
  elements: List[TextElement],
  y: Float,
  maxFontSize: Float
)
