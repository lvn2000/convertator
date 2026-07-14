package convertator.pptx

import convertator.model.*

import org.apache.poi.xslf.usermodel.{XMLSlideShow, XSLFSlide, XSLFTextBox}

/** Context passed to [[SlidePlacer]] methods, bundling commonly needed
  * parameters that are derived from [[ConversionConfig]] and slide geometry. */
case class PlaceContext(
  cfg: ConversionConfig,
  availW: Double,
  usableH: Double,
  fontScale: Float,
  posScale: Float,
  maxImgH: Double
)

/** Type class that encapsulates how a content type is measured, placed on a
  * slide, and optionally split when it overflows. */
trait SlidePlacer[A]:
  /** Estimated vertical space this item occupies (in points). */
  def height(a: A, ctx: PlaceContext): Double

  /** Render this item onto the slide at the given vertical offset.
    * @param tb the current text box (may be null — the placer should return
    *           the (possibly newly created) text box. */
  def place(a: A, slide: XSLFSlide, ppt: XMLSlideShow,
            ctx: PlaceContext, curY: Double, tb: XSLFTextBox): XSLFTextBox

  /** When the item doesn't fit and is the first item on the slide, try to
    * split it. Returns `None` if the item cannot be split (it will overflow
    * to the next slide as a whole). Returns `Some((portion, remainder))`
    * where `remainder` is `None` if the entire item was placed. */
  def trySplit(a: A, ctx: PlaceContext, curY: Double): Option[(A, Option[A])]
