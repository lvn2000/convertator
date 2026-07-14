# Convertator — PDF / DOCX → PPTX Converter

A Scala 3 application that converts PDF and DOCX files into Microsoft PowerPoint (PPTX) presentations, preserving text content, font sizes, styling (bold/italic/underline), embedded images, and tables.

Uses Apache PDFBox and Apache POI for reading; Apache POI (XSLF) for PPTX generation.

## How it works

1. **Read** the source file — `DocumentReader.read()` delegates to `PdfReader` (PDFBox) or `WordReader` (POI) to extract text with positioning, images, and tables.
2. **Group & assemble** — adjacent text fragments are merged into words by positional gaps; table cells are extracted with their row/column structure.
3. **Unified model** — all content (text/images/tables) is wrapped in the `PageItem` ADT for uniform handling.
4. **Render** onto PPTX slides — `SlideBuilder` uses `SlidePlacer` type class instances to measure, place, and split content. Text can use a single grouped textbox (default) or per-line textboxes (--per-line-textbox) to prevent overlap with tables/images.
5. **Overflow handling** — content that doesn't fit the current slide naturally continues on the next slide. When a single item is too large even as the first item on a slide, it attempts trySplit; if unsplittable, it overflows and is counted in the final message.
6. **Default language** is set to Ukrainian (`uk-UA`) for correct proofing tools.

## Supported formats

| Input  | Output | Reader            |
|--------|--------|-------------------|
| `.pdf` | `.pptx`| `PdfReader`       |
| `.docx`| `.pptx`| `WordReader`      |

## Build

```bash
sbt assembly
```

Produces a fat JAR at:
```
target/out/jvm/scala-3.3.4/convertator/convertator-assembly-0.1.0.jar
```

## Usage — via sbt (development)

Run directly with `sbt` without building a JAR:

```bash
# Batch mode — processes all files in src/main/resources/source/
sbt run

# Single file
sbt 'run "mydoc.pdf" "mydoc.pptx"'
sbt 'run "mydoc.docx" "mydoc.pptx"'

# With options
sbt 'run "mydoc.pdf" "mydoc.pptx" --font-size 22 --mode fit'
```

> **Note:** The `run` command expects two or more arguments. Use quotes around the entire argument list after `run`.

## Usage — fat JAR (deployment)

### Build

```bash
sbt assembly
```

Produces a fat JAR at:
```
target/out/jvm/scala-3.3.4/convertator/convertator-assembly-0.1.0.jar
```

### Batch mode (default)

Place your PDF or DOCX files into `src/main/resources/source/`, then run:

```bash
java -jar target/out/jvm/scala-3.3.4/convertator/convertator-assembly-0.1.0.jar
```

Converted PPTX files appear in `src/main/resources/output/` with the same filename (but `.pptx` extension).

Options can be passed:
```bash
java -jar <jar> --font-size 14
java -jar <jar> --mode fit
```

### Single file

```bash
java -jar <jar> mydoc.pdf mydoc.pptx [options]
```

### Options

| Option           | Description                                     | Default |
|------------------|-------------------------------------------------|---------|
| `--help`         | Show help                                       | —       |
| `--mode`         | Page mode: `flow` or `fit`                      | `flow`  |
| `--font-size`    | Target body text size in points                 | 18      |
| `--font-scale`   | Font-size scaling ratio (overrides `--font-size`) | auto  |
| `--slide-width`  | Slide width in points                           | 720     |
| `--slide-height` | Slide height in points                          | 540     |
| `--margin-x`     | Horizontal margin (points)                      | 48      |
| `--margin-y`     | Vertical margin (points)                        | 48      |
| `--line-spacing` | Gap between lines (points)                      | 4       |
| `--per-line-textbox` | Place each text line in its own textbox (prevents text/table overlap)  | false |

### Modes

| Mode   | Behaviour                                                    |
|--------|--------------------------------------------------------------|
| `flow` | Keeps natural font sizes, **overflows** to new slides as needed (default)  |
| `fit`  | Scales each PDF page proportionally to fit **one slide**     |

Use `flow` for book-like pagination where dense content spans multiple slides.

### Examples

```bash
# Batch mode (default) — processes all PDFs/DOCX in src/main/resources/source/
java -jar <jar>

# With larger text
java -jar <jar> --font-size 22

# Widescreen 16:9
java -jar <jar> doc.pdf doc.pptx --slide-width 960 --slide-height 540

# Fit mode — one slide per PDF page
java -jar <jar> --mode fit

# DOCX file
java -jar <jar> report.docx report.pptx
```

## Project structure

```
src/main/
├── resources/
│   ├── source/         ← Put your input files here
│   └── output/         ← Converted PPTX files appear here
└── scala/convertator/
    ├── Main.scala                  # CLI entry point
    ├── Converter.scala             # Public API orchestrator (delegates to CoreConverter)
    ├── CoreConverter.scala         # Core conversion logic; returns Either[Throwable, Unit]
    ├── model/
    │   ├── ConversionConfig.scala  # Tuning parameters & PageMode enum
    │   ├── PageContent.scala       # Extracted page/section content
    │   ├── PageItem.scala          # ADT: TextLineItem | ImageItem | TableItem (replaces union types)
    │   ├── TextElement.scala       # Positioned text fragment (font, style, coords)
    │   ├── TextLine.scala          # One line of text elements
    │   ├── PageImage.scala         # Extracted image with position & dimensions
    │   ├── TableContent.scala      # Table / row / cell model
    │   └── Positioned.scala        # `Positioned` type class trait
    ├── readers/
    │   └── DocumentReader.scala    # Trait for document readers (PDF/DOCX abstraction)
    ├── pdf/
    │   └── PdfReader.scala         # PDF extraction via PDFBox (extends DocumentReader)
    ├── docx/
    │   └── WordReader.scala        # DOCX extraction via POI (extends DocumentReader)
    └── pptx/
        ├── SlideBuilder.scala      # PPTX generation via POI (rendering orchestrator)
        ├── SlidePlacer.scala       # `SlidePlacer` type class (moved from SlideBuilder)
        └── SlidePlacers.scala      # Given instances for SlidePlacer[TextLine | PageImage | PageTable]
```

## Architecture — type classes & ADT

### `PageItem` ADT (`model/PageItem.scala`)

Unified algebraic data type:
- `TextLineItem(line: TextLine)`
- `ImageItem(img: PageImage)`
- `TableItem(tbl: PageTable)`

Replaces union types (`TextLine | PageImage | PageTable`), enabling stronger compile-time type safety and easier pattern matching.

### `Positioned[A]` (`model/Positioned.scala`)

Provides sorting position (`y`) and display priority (`imgOrder`) for any content type. Instances exist for `TextLine`, `PageImage`, `PageTable`, and `PageItem` (delegates to underlying type).

### `SlidePlacer[A]` (`pptx/SlidePlacer.scala` + `pptx/SlidePlacers.scala`)

Encapsulates three operations per content type:

| Method       | Purpose                                                   |
|--------------|-----------------------------------------------------------|
| `height`     | Estimated vertical space the item occupies (points)       |
| `place`      | Render the item onto the slide at a given y-offset        |
| `trySplit`   | Split overflowing content; returns `None` if unsplittable |

Given instances for `TextLine`, `PageImage`, and `PageTable` are defined in `SlidePlacers.scala`.

Adding a new content type (e.g. `PageChart`) requires:
1. A model case class
2. A `Positioned` given instance
3. A `SlidePlacer` given instance with the three methods
4. A `PageItem` case (e.g., `ChartItem`)
5. A case branch in `buildSlides` match

### `DocumentReader` (`readers/DocumentReader.scala`)

Abstraction trait for document extraction (PDF/DOCX):
```scala
trait DocumentReader:
  def read(input: File): Seq[PageContent]
  def read(input: InputStream): Seq[PageContent]
```

Implemented by `PdfReader` and `WordReader`, allowing pluggable readers for future formats.

### `CoreConverter` (`CoreConverter.scala`)

Pure function returning `Either[Throwable, Unit]` for core conversion logic. Separates I/O concerns (Converter.run) from pure business logic, enabling unit testing without side effects.

## Workflow

1. Place PDF or DOCX files in `src/main/resources/source/`
2. Run `sbt assembly` (one-time build, produces fat JAR)
3. Run `java -jar <jar>` to convert all files
4. Open the resulting PPTX files from `src/main/resources/output/`

## Limitations

- **PDF tables** are treated as positioned text (no cell/row structure). DOCX tables are fully supported.
- **Font mapping** relies on PDF/DOCX font names — fallback fonts in PowerPoint may differ.
- **Nested tables** inside table cells are not extracted (only top-level DOCX tables via `getBodyElements`).
