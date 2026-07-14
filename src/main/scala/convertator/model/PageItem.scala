package convertator.model

/** Sealed ADT representing any item that can appear on a page. */
sealed trait PageItem
object PageItem:
  final case class TextLineItem(line: TextLine) extends PageItem
  final case class ImageItem(img: PageImage) extends PageItem
  final case class TableItem(tbl: PageTable) extends PageItem

  // Provide a Positioned instance for PageItem by delegating to the
  // existing Positioned instances for underlying types.
  import scala.language.implicitConversions
  given Positioned[PageItem] with
    def y(pi: PageItem): Float = pi match
      case TextLineItem(l) => summon[Positioned[TextLine]].y(l)
      case ImageItem(i)    => summon[Positioned[PageImage]].y(i)
      case TableItem(t)    => summon[Positioned[PageTable]].y(t)

    def imgOrder(pi: PageItem): Int = pi match
      case TextLineItem(l) => summon[Positioned[TextLine]].imgOrder(l)
      case ImageItem(i)    => summon[Positioned[PageImage]].imgOrder(i)
      case TableItem(t)    => summon[Positioned[PageTable]].imgOrder(t)
