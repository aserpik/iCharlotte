from pathlib import Path


def _requirements_lines():
    requirements_path = Path(__file__).resolve().parents[1] / "requirements.txt"
    return [
        line.strip()
        for line in requirements_path.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.strip().startswith("#")
    ]


def _declared_package_names():
    names = set()
    for line in _requirements_lines():
        package_spec = line.split("#", 1)[0].strip()
        name = package_spec.split(";", 1)[0]
        for separator in ("==", ">=", "<=", "~=", "!=", ">", "<"):
            name = name.split(separator, 1)[0]
        names.add(name.strip().lower())
    return names


def test_pdf_ocr_runtime_dependencies_are_declared():
    declared = _declared_package_names()

    assert {
        "pypdf",
        "python-docx",
        "pymupdf",
        "pytesseract",
        "pdf2image",
    }.issubset(declared)
