from pathlib import Path
root = Path(__file__).resolve().parent
for source in (root / "src").rglob("*"):
    if source.suffix not in (".bas", ".cls"):
        continue
    target = root / "import_icin" / source.relative_to(root / "src")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(source.read_text(encoding="utf-8-sig"), encoding="cp1254")
print("VBE icin Windows-1254 dosyalari hazirlandi.")
