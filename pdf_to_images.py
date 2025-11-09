import fitz
import argparse
from pathlib import Path

def pdf_to_pngs(pdf_path, outdir, dpi=300, fmt="png"):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        out = outdir / f"page_{i:03d}.{fmt}"
        pix.save(out.as_posix())
        print("Saved", out)
    print("Done.")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--outdir", default="pages")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--fmt", default="png", choices=["png","jpg","jpeg","tif","tiff","webp"])
    args = ap.parse_args()
    pdf_to_pngs(args.pdf, args.outdir, dpi=args.dpi, fmt=args.fmt)
