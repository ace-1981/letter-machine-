"""Entry point for template creation — prefers Word COM builder."""

from pathlib import Path

from scripts.build_template_with_word import create_fidelity_template


def create_template(output_path: Path) -> Path:
    return create_fidelity_template(output_path)


if __name__ == "__main__":
    out = Path(__file__).resolve().parent.parent / "templates" / "תחשיב זכויות אישי.docx"
    print(f"Created: {create_template(out)}")
