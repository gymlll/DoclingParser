from pathlib import Path

from docling.document_converter import DocumentConverter


def main() -> None:
    project_root = Path(__file__).resolve().parent
    template_dir = project_root / "template"
    output_dir = project_root / "outputq"
    output_dir.mkdir(parents=True, exist_ok=True)

    xlsx_files = sorted(template_dir.glob("*.xlsx"))
    print(xlsx_files)
    if not xlsx_files:
        raise FileNotFoundError(f"No .xlsx file found in: {template_dir}")

    source = xlsx_files[0]
    converter = DocumentConverter()
    result = converter.convert(str(source))
    markdown_text = result.document.export_to_markdown()

    markdown_path = output_dir / f"{source.stem}.md"
    markdown_path.write_text(markdown_text, encoding="utf-8")

    print(f"Source: {source}")
    print(f"Saved markdown to: {markdown_path}")
    print(f"Markdown chars: {len(markdown_text)}")


if __name__ == "__main__":
    main()