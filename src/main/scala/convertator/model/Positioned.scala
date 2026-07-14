package convertator.model

/** Type class: anything that can be positioned on a slide (has a y-coordinate
  * and a display priority relative to other content at the same y). */
trait Positioned[A]:
  def y(a: A): Float
  /** Lower values sort first (images/tables before text at same y). */
  def imgOrder(a: A): Int

object Positioned:
  given Positioned[TextLine] with
    def y(t: TextLine): Float = t.y
    def imgOrder(t: TextLine): Int = 1

  given Positioned[PageImage] with
    def y(i: PageImage): Float = i.y
    def imgOrder(i: PageImage): Int = 0

  given Positioned[PageTable] with
    def y(t: PageTable): Float = t.y
    def imgOrder(t: PageTable): Int = 0

  extension [A: Positioned](a: A)
    def positionedY: Float = summon[Positioned[A]].y(a)
    def positionedOrder: Int = summon[Positioned[A]].imgOrder(a)
