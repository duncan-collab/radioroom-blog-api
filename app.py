"""
The Radio Room — Blog Post Publisher API
Deploy to Render.com. Accepts a .docx POST, converts to HTML, FTPs to SiteGround.

Environment variables (set in Render dashboard):
  FTP_HOST       — SiteGround FTP hostname, e.g. ftp.theradioroom.ie
  FTP_USER       — FTP username
  FTP_PASS       — FTP password
  FTP_BLOG_PATH  — Server path to blog folder, e.g. /public_html/blog
  API_KEY        — A secret string Make will send to authenticate requests
"""

import os, io, re, json, ftplib, tempfile, datetime
try:
    import paramiko
except ImportError:
    paramiko = None
from flask import Flask, request, jsonify

import mammoth
from docx import Document

app = Flask(__name__)

# ── Config ───────────────────────────────────────────────────────────────────
FTP_HOST      = os.environ.get("FTP_HOST", "")
FTP_USER      = os.environ.get("FTP_USER", "")
FTP_PASS      = os.environ.get("FTP_PASS", "")
FTP_BLOG_PATH = os.environ.get("FTP_BLOG_PATH", "/public_html/blog")
API_KEY       = os.environ.get("API_KEY", "changeme")

# ── Helpers (same logic as docx-to-post.py) ──────────────────────────────────
def slugify(text):
    text = text.lower().strip()
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "-", text)
    return re.sub(r"^-+|-+$", "", text)

def extract_metadata(doc):
    meta = {
        "title": "", "date": datetime.date.today().strftime("%d %B %Y").lstrip("0"),
        "author": "The Radio Room", "category": "", "excerpt": "",
    }
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            label = row.cells[0].text.strip().upper()
            value = row.cells[1].text.strip()
            if value.startswith("[") or not value:
                continue
            if label == "TITLE":      meta["title"]    = value
            elif label == "DATE":     meta["date"]     = value
            elif label == "AUTHOR":   meta["author"]   = value
            elif label == "CATEGORY": meta["category"] = value
            elif label == "EXCERPT":  meta["excerpt"]  = value
    return meta

def convert_body(docx_path):
    style_map = """
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
    """
    with open(docx_path, "rb") as f:
        result = mammoth.convert_to_html(f, style_map=style_map)
    html = result.value
    html = re.sub(r"<table>.*?</table>", "", html, flags=re.DOTALL)
    html = re.sub(r"<p>[^<]*─+[^<]*</p>", "", html)
    html = re.sub(r"<p>[^<]*YOUR POST STARTS BELOW[^<]*</p>", "", html, flags=re.IGNORECASE)
    html = re.sub(r"<p>\s*</p>", "", html)
    return html.strip()

def build_post_html(meta, body_html):
    title    = meta["title"] or "Untitled Post"
    date     = meta["date"]
    author   = meta["author"]
    category = meta["category"]
    excerpt  = meta["excerpt"]
    cat_badge = f'<span class="post-cat">{category}</span>' if category else ""
    year = datetime.date.today().year

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title} — The Radio Room</title>
  <meta name="description" content="{excerpt}">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {{--white:#FFFFFF;--cream:#F5F7FA;--navy:#1D3780;--teal:#29D4E6;--text:#1A1A2E;--muted:#6B7280;--border:#E5E7EB;}}
    *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
    html{{scroll-behavior:smooth;}}
    body{{font-family:'Inter',system-ui,sans-serif;background:var(--white);color:var(--text);line-height:1.6;-webkit-font-smoothing:antialiased;}}
    nav{{position:fixed;top:0;left:0;right:0;height:90px;background:var(--white);display:flex;align-items:center;justify-content:space-between;padding:0 5%;z-index:100;box-shadow:0 1px 0 var(--border);}}
    nav .logo img{{height:78px;width:auto;}}
    nav ul{{list-style:none;display:flex;gap:2.25rem;}}
    nav ul a{{color:var(--text);text-decoration:none;font-size:.9rem;font-weight:500;transition:color .2s;}}
    nav ul a:hover{{color:var(--navy);}}
    .btn-nav{{background:var(--navy);color:#fff;padding:.55rem 1.4rem;border-radius:7px;font-weight:700;font-size:.875rem;text-decoration:none;}}
    .post-header{{background:var(--cream);padding:8rem 5% 4rem;text-align:center;}}
    .post-cat{{display:inline-block;background:var(--teal);color:var(--navy);font-size:.72rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;padding:.3rem .9rem;border-radius:100px;margin-bottom:1.25rem;}}
    .post-header h1{{font-size:clamp(1.8rem,3.5vw,2.8rem);font-weight:900;letter-spacing:-.03em;color:var(--navy);max-width:760px;margin:0 auto 1rem;line-height:1.15;}}
    .post-meta{{font-size:.875rem;color:var(--muted);}}
    .post-meta span+span::before{{content:' · ';}}
    .post-body{{max-width:740px;margin:0 auto;padding:4rem 5%;}}
    .post-body h1{{display:none;}}
    .post-body h2{{font-size:1.5rem;font-weight:800;color:var(--navy);letter-spacing:-.02em;margin:2.5rem 0 .9rem;padding-bottom:.5rem;border-bottom:2px solid var(--teal);}}
    .post-body h3{{font-size:1.15rem;font-weight:700;color:var(--navy);margin:2rem 0 .6rem;}}
    .post-body p{{font-size:1.05rem;color:var(--muted);line-height:1.8;margin-bottom:1.25rem;}}
    .post-body ul,.post-body ol{{margin:1rem 0 1.5rem 1.75rem;}}
    .post-body li{{font-size:1.05rem;color:var(--muted);line-height:1.7;margin-bottom:.4rem;}}
    .post-body strong{{color:var(--text);font-weight:700;}}
    .post-body a{{color:var(--navy);text-decoration:underline;}}
    .back-link{{display:inline-flex;align-items:center;gap:.5rem;color:var(--navy);text-decoration:none;font-weight:600;font-size:.9rem;margin-bottom:2.5rem;}}
    .footer-cta{{background:var(--navy);text-align:center;padding:5rem 5%;}}
    .footer-cta h2{{font-size:clamp(1.5rem,2.5vw,2.2rem);font-weight:800;color:#fff;margin-bottom:.75rem;}}
    .footer-cta p{{color:rgba(255,255,255,.6);font-size:1rem;max-width:420px;margin:0 auto 2rem;}}
    .btn-teal{{background:var(--teal);color:var(--navy);padding:.9rem 2rem;border-radius:8px;font-weight:700;font-size:.95rem;text-decoration:none;display:inline-block;}}
    footer{{background:#0F1F4A;padding:2rem 5%;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:1rem;}}
    footer img{{height:38px;}}
    footer p{{font-size:.8rem;color:rgba(255,255,255,.35);}}
    @media(max-width:768px){{nav ul{{display:none;}}}}
  </style>
</head>
<body>
<nav>
  <a href="../index.html" class="logo"><img src="../logo.png" alt="The Radio Room"></a>
  <ul>
    <li><a href="../index.html">Home</a></li>
    <li><a href="index.html">Blog</a></li>
    <li><a href="../index.html#contact">Contact</a></li>
  </ul>
  <a href="../index.html#contact" class="btn-nav">Book a Call</a>
</nav>
<div class="post-header">
  {cat_badge}
  <h1>{title}</h1>
  <p class="post-meta"><span>{date}</span><span>{author}</span></p>
</div>
<article class="post-body">
  <a href="index.html" class="back-link">← Back to Blog</a>
  {body_html}
</article>
<section class="footer-cta">
  <h2>Your audio campaign starts here.</h2>
  <p>Independent guidance across radio, streaming, and podcasts.</p>
  <a href="../index.html#contact" class="btn-teal">Get In Touch</a>
</section>
<footer>
  <img src="../logo.png" alt="The Radio Room">
  <p>© {year} The Radio Room. All rights reserved.</p>
</footer>
</body>
</html>"""

def build_index_html(posts):
    year = datetime.date.today().year
    cards = ""
    for p in sorted(posts, key=lambda x: x.get("date_iso", ""), reverse=True):
        cat = f'<span class="post-cat">{p["category"]}</span>' if p.get("category") else ""
        cards += f"""
    <a href="{p['slug']}.html" class="card">
      {cat}
      <h2>{p['title']}</h2>
      <p class="excerpt">{p['excerpt']}</p>
      <div class="card-footer">
        <span class="date">{p['date']}</span>
        <span class="read-more">Read post →</span>
      </div>
    </a>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Blog — The Radio Room</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root{{--white:#FFFFFF;--cream:#F5F7FA;--navy:#1D3780;--teal:#29D4E6;--text:#1A1A2E;--muted:#6B7280;--border:#E5E7EB;}}
    *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
    body{{font-family:'Inter',system-ui,sans-serif;background:var(--white);color:var(--text);-webkit-font-smoothing:antialiased;}}
    nav{{position:fixed;top:0;left:0;right:0;height:90px;background:var(--white);display:flex;align-items:center;justify-content:space-between;padding:0 5%;z-index:100;box-shadow:0 1px 0 var(--border);}}
    nav .logo img{{height:78px;}} nav ul{{list-style:none;display:flex;gap:2.25rem;}}
    nav ul a{{color:var(--text);text-decoration:none;font-size:.9rem;font-weight:500;}}
    .btn-nav{{background:var(--navy);color:#fff;padding:.55rem 1.4rem;border-radius:7px;font-weight:700;font-size:.875rem;text-decoration:none;}}
    .page-header{{background:var(--cream);padding:8rem 5% 4rem;text-align:center;}}
    .label{{font-size:.72rem;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--teal);margin-bottom:.9rem;}}
    .page-header h1{{font-size:clamp(2rem,4vw,3rem);font-weight:900;color:var(--navy);letter-spacing:-.03em;margin-bottom:.75rem;}}
    .page-header p{{color:var(--muted);font-size:1rem;max-width:480px;margin:0 auto;}}
    .posts{{padding:4rem 5%;max-width:1100px;margin:0 auto;}}
    .grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:1.75rem;}}
    .card{{background:var(--cream);border:1px solid var(--border);border-radius:18px;padding:2rem;text-decoration:none;display:flex;flex-direction:column;transition:all .3s;}}
    .card:hover{{border-color:var(--teal);transform:translateY(-4px);box-shadow:0 12px 40px rgba(41,212,230,.12);}}
    .post-cat{{display:inline-block;background:var(--teal);color:var(--navy);font-size:.68rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;padding:.25rem .75rem;border-radius:100px;margin-bottom:1rem;}}
    .card h2{{font-size:1.15rem;font-weight:800;color:var(--navy);letter-spacing:-.02em;margin-bottom:.75rem;line-height:1.3;}}
    .excerpt{{font-size:.9rem;color:var(--muted);line-height:1.65;flex:1;margin-bottom:1.25rem;}}
    .card-footer{{display:flex;justify-content:space-between;align-items:center;padding-top:1rem;border-top:1px solid var(--border);}}
    .date{{font-size:.8rem;color:var(--muted);}} .read-more{{font-size:.8rem;font-weight:700;color:var(--navy);}}
    .footer-cta{{background:var(--navy);text-align:center;padding:5rem 5%;}}
    .footer-cta h2{{font-size:clamp(1.5rem,2.5vw,2.2rem);font-weight:800;color:#fff;margin-bottom:.75rem;}}
    .footer-cta p{{color:rgba(255,255,255,.6);font-size:1rem;max-width:420px;margin:0 auto 2rem;}}
    .btn-teal{{background:var(--teal);color:var(--navy);padding:.9rem 2rem;border-radius:8px;font-weight:700;font-size:.95rem;text-decoration:none;display:inline-block;}}
    footer{{background:#0F1F4A;padding:2rem 5%;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:1rem;}}
    footer img{{height:38px;}} footer p{{font-size:.8rem;color:rgba(255,255,255,.35);}}
    @media(max-width:900px){{.grid{{grid-template-columns:repeat(2,1fr);}}}}
    @media(max-width:600px){{nav ul{{display:none;}}.grid{{grid-template-columns:1fr;}}}}
  </style>
</head>
<body>
<nav>
  <a href="../index.html" class="logo"><img src="../logo.png" alt="The Radio Room"></a>
  <ul><li><a href="../index.html">Home</a></li><li><a href="index.html">Blog</a></li><li><a href="../index.html#contact">Contact</a></li></ul>
  <a href="../index.html#contact" class="btn-nav">Book a Call</a>
</nav>
<div class="page-header">
  <div class="label">Insights & Ideas</div>
  <h1>The Radio Room Blog</h1>
  <p>Audio advertising tips, industry news and campaign ideas for Irish businesses.</p>
</div>
<section class="posts">
  <div class="grid">{'<p style="color:#6B7280;grid-column:1/-1;text-align:center;padding:3rem 0">No posts yet.</p>' if not posts else cards}</div>
</section>
<section class="footer-cta">
  <h2>Your audio campaign starts here.</h2>
  <p>Independent guidance across radio, streaming, and podcasts.</p>
  <a href="../index.html#contact" class="btn-teal">Get In Touch</a>
</section>
<footer>
  <img src="../logo.png" alt="The Radio Room">
  <p>© {year} The Radio Room. All rights reserved.</p>
</footer>
</body>
</html>"""

# ── SFTP helpers ──────────────────────────────────────────────────────────────
def sftp_read_json(sftp, path):
    try:
        with sftp.file(path, "r") as f:
            return json.loads(f.read().decode("utf-8"))
    except Exception:
        return []

def sftp_write(sftp, path, content):
    with sftp.file(path, "w") as f:
        f.write(content.encode("utf-8"))

def ensure_sftp_dir(sftp, path):
    try:
        sftp.mkdir(path)
    except Exception:
        pass  # already exists

# ── API routes ────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "The Radio Room Blog Publisher"})

@app.route("/convert", methods=["POST"])
def convert():
    # Auth check
    if request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    # Accept file either as multipart OR as raw binary body
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_path = tmp.name
    tmp.close()  # Close handle immediately so save can write cleanly

    if "file" in request.files:
        request.files["file"].save(tmp_path)
    elif request.data:
        with open(tmp_path, "wb") as f:
            f.write(request.data)
    else:
        os.unlink(tmp_path)
        return jsonify({"error": "No file provided. Send .docx as multipart field 'file' or as raw binary body."}), 400

    file_size = os.path.getsize(tmp_path)
    with open(tmp_path, "rb") as f:
        first_bytes = f.read(4).hex()  # ZIP/DOCX starts with 504b0304

    try:
        # Convert
        doc  = Document(tmp_path)
        meta = extract_metadata(doc)

        if not meta["title"]:
            return jsonify({"error": "No TITLE found in the metadata table. Please fill it in."}), 400

        slug      = slugify(meta["title"])
        body_html = convert_body(tmp_path)
        post_html = build_post_html(meta, body_html)

        # Update posts list — accepts current posts.json from Make via form field
        current_posts_json = request.form.get("posts_json", "[]")
        try:
            posts = json.loads(current_posts_json)
        except Exception:
            posts = []

        posts = [p for p in posts if p.get("slug") != slug]
        posts.append({
            "slug":     slug,
            "title":    meta["title"],
            "date":     meta["date"],
            "date_iso": datetime.date.today().isoformat(),
            "author":   meta["author"],
            "category": meta["category"],
            "excerpt":  meta["excerpt"],
        })

        index_html  = build_index_html(posts)
        posts_json  = json.dumps(posts, indent=2)
        post_filename = f"{slug}.html"

        # Return all three files for Make to upload via FTP
        return jsonify({
            "success":      True,
            "slug":         slug,
            "title":        meta["title"],
            "post_filename": post_filename,
            "post_html":    post_html,
            "index_html":   index_html,
            "posts_json":   posts_json,
            "message":      f"Converted OK — upload {post_filename}, index.html and posts.json"
        })

    except Exception as e:
        return jsonify({"error": str(e), "file_size_bytes": file_size, "first_bytes_hex": first_bytes}), 500

    finally:
        os.unlink(tmp_path)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
