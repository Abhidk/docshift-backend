from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import subprocess, tempfile, os, shutil, io, zipfile, logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="DocShift API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

MIME_MAP = {
    "pdf":  "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "jpg":  "image/jpeg",
    "png":  "image/png",
    "html": "text/html",
    "md":   "text/markdown",
    "zip":  "application/zip",
}

def run(cmd: list[str], cwd: str = None) -> subprocess.CompletedProcess:
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd)}\n{result.stderr}")
    return result

def libreoffice_convert(src_path: str, out_dir: str, out_format: str) -> str:
    run(["libreoffice", "--headless", "--convert-to", out_format, "--outdir", out_dir, src_path])
    base = os.path.splitext(os.path.basename(src_path))[0]
    return os.path.join(out_dir, f"{base}.{out_format}")

def pandoc_convert(src_path: str, out_path: str, from_fmt: str, to_fmt: str):
    run(["pandoc", src_path, "-f", from_fmt, "-t", to_fmt, "-o", out_path])

@app.get("/health")
def health():
    return {"status": "ok", "service": "DocShift API"}

@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    tool: str = Form(...),
    option: str = Form(default=""),
):
    tmpdir = tempfile.mkdtemp()
    try:
        ext = file.filename.rsplit(".", 1)[-1].lower()
        src = os.path.join(tmpdir, f"input.{ext}")
        with open(src, "wb") as f:
            f.write(await file.read())

        out_bytes, out_name, out_ext = await dispatch(tool, src, ext, option, tmpdir)
        mime = MIME_MAP.get(out_ext, "application/octet-stream")
        return StreamingResponse(
            io.BytesIO(out_bytes),
            media_type=mime,
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'}
        )
    except Exception as e:
        logger.error(f"Conversion error: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

async def dispatch(tool, src, ext, option, tmpdir):
    match tool:
        case "merge-pdf":    return await do_merge(src, tmpdir)
        case "split-pdf":    return await do_split(src, tmpdir)
        case "compress-pdf": return await do_compress(src, tmpdir)
        case "word-to-pdf":  return await lo_to_pdf(src, tmpdir, "document.pdf")
        case "excel-to-pdf": return await lo_to_pdf(src, tmpdir, "spreadsheet.pdf")
        case "ppt-to-pdf":   return await lo_to_pdf(src, tmpdir, "presentation.pdf")
        case "jpg-to-pdf":   return await lo_to_pdf(src, tmpdir, "image.pdf")
        case "html-to-pdf":  return await pandoc_to_pdf(src, tmpdir)
        case "pdf-to-word":  return await lo_from_pdf(src, tmpdir, "docx")
        case "pdf-to-excel": return await lo_from_pdf(src, tmpdir, "xlsx")
        case "pdf-to-ppt":   return await lo_from_pdf(src, tmpdir, "pptx")
        case "pdf-to-jpg":   return await pdf_to_images(src, tmpdir)
        case "rotate-pdf":   return await do_rotate(src, option, tmpdir)
        case "watermark":    return await do_watermark(src, option, tmpdir)
        case "protect-pdf":  return await do_protect(src, option, tmpdir)
        case "unlock-pdf":   return await do_unlock(src, option, tmpdir)
        case _: raise ValueError(f"Unknown tool: {tool}")

async def lo_to_pdf(src, tmpdir, outname):
    out = libreoffice_convert(src, tmpdir, "pdf")
    return open(out,"rb").read(), outname, "pdf"

async def lo_from_pdf(src, tmpdir, fmt):
    out = libreoffice_convert(src, tmpdir, fmt)
    names = {"docx":"document.docx","xlsx":"spreadsheet.xlsx","pptx":"presentation.pptx"}
    return open(out,"rb").read(), names[fmt], fmt

async def pandoc_to_pdf(src, tmpdir):
    out = os.path.join(tmpdir, "output.pdf")
    run(["pandoc", src, "-o", out, "--pdf-engine=libreoffice"])
    return open(out,"rb").read(), "webpage.pdf", "pdf"

async def do_merge(src, tmpdir):
    # For merge, expects single zip of PDFs or we just return the single file
    out = os.path.join(tmpdir, "merged.pdf")
    run(["gs", "-dBATCH", "-dNOPAUSE", "-q", "-sDEVICE=pdfwrite", f"-sOutputFile={out}", src])
    return open(out,"rb").read(), "merged.pdf", "pdf"

async def do_split(src, tmpdir):
    run(["pdfseparate", src, os.path.join(tmpdir, "page-%d.pdf")])
    pages = sorted([f for f in os.listdir(tmpdir) if f.startswith("page-") and f.endswith(".pdf")])
    zb = io.BytesIO()
    with zipfile.ZipFile(zb, "w") as zf:
        for p in pages:
            zf.write(os.path.join(tmpdir, p), p)
    return zb.getvalue(), "split-pages.zip", "zip"

async def do_compress(src, tmpdir):
    out = os.path.join(tmpdir, "compressed.pdf")
    run(["gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
         "-dPDFSETTINGS=/screen", "-dNOPAUSE", "-dQUIET", "-dBATCH",
         f"-sOutputFile={out}", src])
    return open(out,"rb").read(), "compressed.pdf", "pdf"

async def pdf_to_images(src, tmpdir):
    run(["pdftoppm", "-jpeg", "-r", "150", src, os.path.join(tmpdir, "page")])
    imgs = sorted([f for f in os.listdir(tmpdir) if f.endswith(".jpg")])
    zb = io.BytesIO()
    with zipfile.ZipFile(zb, "w") as zf:
        for img in imgs:
            zf.write(os.path.join(tmpdir, img), img)
    return zb.getvalue(), "pages.zip", "zip"

async def do_rotate(src, option, tmpdir):
    degrees = option or "90"
    out = os.path.join(tmpdir, "rotated.pdf")
    run(["qpdf", src, "--rotate=+" + degrees, out])
    return open(out,"rb").read(), "rotated.pdf", "pdf"

async def do_watermark(src, option, tmpdir):
    text = option or "CONFIDENTIAL"
    out = os.path.join(tmpdir, "watermarked.pdf")
    run(["gs", "-dBATCH", "-dNOPAUSE", "-q", "-sDEVICE=pdfwrite",
         f"-sOutputFile={out}",
         "-c", f"<</BeginPage{{{/Helvetica findfont 48 scalefont setfont 0.7 setgray 200 300 translate 45 rotate 0 0 moveto ({text}) show}}>> setpagedevice",
         "-f", src])
    if not os.path.exists(out):
        import shutil
        shutil.copy(src, out)
    return open(out,"rb").read(), "watermarked.pdf", "pdf"

async def do_protect(src, option, tmpdir):
    password = option or "password"
    out = os.path.join(tmpdir, "protected.pdf")
    run(["qpdf", "--encrypt", password, password, "256", "--", src, out])
    return open(out,"rb").read(), "protected.pdf", "pdf"

async def do_unlock(src, option, tmpdir):
    password = option or ""
    out = os.path.join(tmpdir, "unlocked.pdf")
    cmd = ["qpdf", "--decrypt"]
    if password:
        cmd += ["--password=" + password]
    cmd += [src, out]
    run(cmd)
    return open(out,"rb").read(), "unlocked.pdf", "pdf"
