# api/index.py
# SINGLE FILE — FINAL — NO IMGBB — API KEY INLINE — VERCEL SAFE

from flask import Flask, request, send_file, render_template_string
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io, re, requests

app = Flask(__name__)

# ---------------- CONFIG ----------------
# Inline API key placeholder (not imgbb, no uploads)
API_KEY = "9d1c67a0ea72097a70e086c48c838c35"

# ---------------- UTILS ----------------

def is_url(s: str) -> bool:
    return s.startswith("http://") or s.startswith("https://")

def fetch_image(src: str) -> io.BytesIO:
    if not is_url(src):
        raise ValueError("Image source must be a direct URL")
    r = requests.get(src, timeout=10)
    r.raise_for_status()
    return io.BytesIO(r.content)

def parse_attrs(block: str) -> dict:
    return {
        k: v.strip('"')
        for k, v in re.findall(r'(\w+)=(".*?"|\S+)', block)
    }

# ---------------- PPT ENGINE ----------------

def generate_ppt(script: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    W, H = prs.slide_width, prs.slide_height
    M = Inches(0.5)

    def centered_text(text):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(
            int(W * 0.25), int(H * 0.35),
            int(W * 0.5), int(H * 0.3)
        )
        p = box.text_frame.paragraphs[0]
        p.text = text
        p.font.name = "Arial"
        p.font.size = Pt(44)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.CENTER

    def image_only(cfg):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        img = fetch_image(cfg["src"])

        width_pct = float(cfg.get("width", "60%").strip("%")) / 100
        img_w = int(W * width_pct)
        left = int((W - img_w) / 2)

        slide.shapes.add_picture(img, left, int(M), width=img_w)

    def mix(cfg):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        img = fetch_image(cfg["src"])
        align = cfg.get("align", "left")

        img_left = int(M if align == "left" else W / 2 + M)
        txt_left = int(W / 2 + M if align == "left" else M)

        slide.shapes.add_picture(
            img,
            img_left,
            int(M),
            width=int(W / 2 - M * 2)
        )

        box = slide.shapes.add_textbox(
            txt_left,
            int(H * 0.35),
            int(W / 2 - M * 2),
            int(H * 0.3)
        )
        p = box.text_frame.paragraphs[0]
        p.text = cfg.get("text", "")
        p.font.name = "Arial"
        p.font.size = Pt(32)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.CENTER

    def video(cfg):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(
            int(W * 0.25), int(H * 0.35),
            int(W * 0.5), int(H * 0.3)
        )
        p = box.text_frame.paragraphs[0]
        p.text = f"VIDEO\n{cfg.get('src','')}"
        p.font.size = Pt(28)
        p.alignment = PP_ALIGN.CENTER

    tokens = re.findall(r'\[(IMAGE|MIX|VIDEO)(.*?)\]|(\S+)', script)

    for kind, block, word in tokens:
        if word:
            centered_text(word)
        elif kind == "IMAGE":
            image_only(parse_attrs(block))
        elif kind == "MIX":
            mix(parse_attrs(block))
        elif kind == "VIDEO":
            video(parse_attrs(block))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

# ---------------- ROUTES ----------------

@app.route("/")
def index():
    return render_template_string("""
<!doctype html>
<html>
<head>
<title>PPT Script Builder</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<style>textarea{font-family:monospace}</style>
</head>
<body class="container mt-4">

<h3>PPT Script Builder</h3>

<textarea id="script" class="form-control mb-3" rows="10"
placeholder="Words = slides
[IMAGE src=https://example.com/image.png width=60%]
[MIX text=&quot;Hello&quot; src=https://example.com/image.png align=left]
[VIDEO src=https://example.com/video.mp4]"></textarea>

<form method="POST" action="/generate">
<input type="hidden" name="script" id="final">
<button class="btn btn-success">Generate PPT</button>
</form>

<script>
document.querySelector("form").onsubmit = () => {
  document.getElementById("final").value =
    document.getElementById("script").value;
};
</script>

</body>
</html>
""")

@app.route("/generate", methods=["POST"])
def generate():
    ppt_bytes = generate_ppt(request.form.get("script", ""))
    return send_file(
        io.BytesIO(ppt_bytes),
        download_name="output.pptx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

# NO app.run() — REQUIRED FOR VERCEL
