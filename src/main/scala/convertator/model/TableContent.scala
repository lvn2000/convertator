package convertator.model

/** A single cell in a table row. */
case class TableCell(
  /** Text content of the cell (runs with formatting). */
  textElements: List[TextElement],
  /** Column width in points (0 if unknown). */
  width: Float = 0f
)

/** A single row in a table. */
case class TableRow(
  /** Cells in this row. */
  cells: List[TableCell],
  /** Estimated height of this row in points. */
  height: Float = 0f
)

/** A table extracted from a source document. */
case class PageTable(
  /** Y position on the source page. */
  y: Float,
  /** Column widths in points. */
  columnWidths: List[Float],
  /** Rows of the table. */
  rows: List[TableRow]
)
