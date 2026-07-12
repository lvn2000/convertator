package convertator.model

/** How to map a PDF page's content onto slides. */
enum PageMode:
  /** Keep natural font sizes; overflow to continuation slides (default). */
  case Flow
  /** Scale the whole page to fit on one slide. */
  case Fit
end PageMode

/** Tuning parameters for the PDF → PPTX conversion. */
case class ConversionConfig(
  /** Slide width in points (default 720 = 10″). */
  slideWidth: Double = 720.0,
  /** Slide height in points (default 540 = 7.5″). */
  slideHeight: Double = 540.0,
  /** Left / right margin (points). */
  marginX: Double = 48.0,
  /** Top margin (points). */
  marginY: Double = 48.0,
  /** Vertical gap between lines (points). */
  lineSpacing: Double = 4.0,
  /** Whether to underline text that was underlined in the PDF. */
  preserveUnderline: Boolean = true,
  /**
   * Font-size multiplier applied to every text element.
   * In `Fit`  mode this compounds with the page-to-slide scaling.
   * In `Flow` mode this is applied directly (1.0 = original PDF size).
   *
   * Ignored when [[targetFontSize]] is set.
   */
  fontSizeScale: Float = 1.0f,
  /**
   * Optional target body-text size in points.
   * When set, the app finds the most common font size in the PDF and
   * scales everything proportionally so that body text matches this value.
   * Overrides [[fontSizeScale]].
   */
  targetFontSize: Option[Float] = Some(18f),
  /** How to map PDF pages into slides (default Flow = natural font sizes). */
  pageMode: PageMode = PageMode.Flow
)
