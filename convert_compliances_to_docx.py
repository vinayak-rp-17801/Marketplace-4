from __future__ import annotations

from pathlib import Path

from docx import Document


SOURCE_DIR = Path("compliances")
OUTPUT_DIR = Path("compliances-docx")


def parse_markdown(content: str) -> tuple[str, str, str, list[str]]:
    lines = [line.rstrip() for line in content.splitlines()]

    title = ""
    tagline = ""
    overview_lines: list[str] = []
    key_features: list[str] = []

    section = None
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        if line.startswith("# ") and not title:
            title = line[2:].strip()
            section = None
            continue

        if line.startswith("**Tagline:**"):
            tagline = line.replace("**Tagline:**", "", 1).strip()
            section = None
            continue

        if line == "## Overview":
            section = "overview"
            continue

        if line == "## Key Features":
            section = "key_features"
            continue

        if section == "overview":
            overview_lines.append(line)
        elif section == "key_features" and line.startswith("- "):
            key_features.append(line[2:].strip())

    return title, tagline, "\n".join(overview_lines).strip(), key_features


def create_docx_from_markdown(markdown_path: Path, output_path: Path) -> None:
    title, tagline, overview, key_features = parse_markdown(markdown_path.read_text(encoding="utf-8"))

    document = Document()

    if title:
        document.add_heading(title, level=1)

    if tagline:
        tagline_paragraph = document.add_paragraph()
        label_run = tagline_paragraph.add_run("Tagline:")
        label_run.bold = True
        tagline_paragraph.add_run(f" {tagline}")

    document.add_heading("Overview", level=2)
    if overview:
        for paragraph_text in overview.split("\n"):
            if paragraph_text.strip():
                document.add_paragraph(paragraph_text)

    if key_features:
        document.add_heading("Key Features", level=2)
        for feature in key_features:
            document.add_paragraph(feature, style="List Bullet")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for markdown_file in sorted(SOURCE_DIR.glob("*.md")):
        output_file = OUTPUT_DIR / f"{markdown_file.stem}.docx"
        create_docx_from_markdown(markdown_file, output_file)


if __name__ == "__main__":
    main()
