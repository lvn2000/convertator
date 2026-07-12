# Convertator — PDF → PPTX Converter

A Scala 3 application that converts PDF files into Microsoft PowerPoint (PPTX) presentations, preserving text content, font sizes, and basic styling (bold/italic). Uses Apache PDFBox for PDF reading and Apache POI for PPTX generation.

## How it works

1. **Read** the PDF using Apache PDFBox — extracts every text fragment with its position, font name, font size, and style.
2. **Group** adjacent characters into words, detecting word boundaries by positional gaps (PDFs often position words without space characters).
3. **Render** onto PPTX slides using Apache POI — one text box per slide, one paragraph per PDF line, with word-level text runs. Text overflow is detected and content automatically flows to the next slide.
4. **Default language** is set to Ukrainian (`uk-UA`) for correct proofing tools.

## Build

```bash
sbt assembly
```

Produces a fat JAR at:
```
target/out/jvm/scala-3.3.4/convertator/convertator-assembly-0.1.0.jar
```

## Usage — batch mode (default)

Place your PDF files into `src/main/resources/source/`, then run:

```bash
java -jar target/out/jvm/scala-3.3.4/convertator/convertator-assembly-0.1.0.jar
```

Converted PPTX files appear in `src/main/resources/output/` with the same filename (but `.pptx` extension).

You can also pass options:
```bash
java -jar <jar> --font-size 14
java -jar <jar> --mode fit
```

## Usage — single file

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

### Modes

| Mode   | Behaviour                                                    |
|--------|--------------------------------------------------------------|
| `flow` | Keeps natural font sizes, **overflows** to new slides as needed (default)  |
| `fit`  | Scales each PDF page proportionally to fit **one slide**     |

Use `flow` for book-like pagination where dense content spans multiple slides.

### Examples

```bash
# Batch mode (default) — processes all PDFs in src/main/resources/source/
java -jar <jar>

# With larger text
java -jar <jar> --font-size 22

# Widescreen 16:9
java -jar <jar> doc.pdf doc.pptx --slide-width 960 --slide-height 540

# Fit mode — one slide per PDF page
java -jar <jar> --mode fit
```

## Project structure

```
src/main/
├── resources/
│   ├── source/         ← Put your PDF files here
│   └── output/         ← Converted PPTX files appear here
└── scala/convertator/
    ├── Main.scala                  # CLI entry point
    ├── Converter.scala             # Orchestrator
    ├── model/
    │   ├── ConversionConfig.scala  # Tuning parameters & PageMode enum
    │   ├── PageContent.scala       # Extracted page content
    │   ├── TextElement.scala       # Positioned text fragment
    │   └── TextLine.scala          # Line of text fragments
    ├── pdf/
    │   └── PdfReader.scala         # PDF extraction via PDFBox
    └── pptx/
        └── SlideBuilder.scala      # PPTX generation via POI
```

## Workflow

1. Place PDF files in `src/main/resources/source/`
2. Run `sbt assembly` (one-time build, produces fat JAR)
3. Run `java -jar <jar>` to convert all PDFs
4. Open the resulting PPTX files from `src/main/resources/output/`

## Limitations

- **Images** are not extracted (text-only conversion).
- **Tables** are treated as positioned text (no cell/row structure).
- **Font mapping** relies on PDF font names — fallback fonts in PowerPoint may differ.
