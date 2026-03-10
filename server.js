const express = require("express");
const cors = require("cors");
const https = require("https");
const http = require("http");
const { execFile } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json({ limit: "20mb" }));

// ── Zillow search proxy ───────────────────────────────────────────────────────
app.get("/api/search", (req, res) => {
  const { url, page, apiKey } = req.query;
  if (!url) return res.status(400).json({ error: "Missing url parameter" });
  if (!apiKey) return res.status(400).json({ error: "Missing apiKey parameter" });

  const encodedUrl = encodeURIComponent(url);
  const reqPath = `/api/search/byurl?url=${encodedUrl}&page=${page || 1}`;

  const options = {
    hostname: "real-estate101.p.rapidapi.com",
    path: reqPath,
    method: "GET",
    headers: {
      "x-rapidapi-host": "real-estate101.p.rapidapi.com",
      "x-rapidapi-key": apiKey,
    },
  };

  const proxyReq = https.request(options, (proxyRes) => {
    let data = "";
    proxyRes.on("data", (chunk) => (data += chunk));
    proxyRes.on("end", () => {
      res.status(proxyRes.statusCode).set("Content-Type", "application/json").send(data);
    });
  });

  proxyReq.on("error", (err) => res.status(500).json({ error: err.message }));
  proxyReq.end();
});

// ── Excel export ──────────────────────────────────────────────────────────────
app.post("/api/export/xlsx", async (req, res) => {
  const { listings } = req.body;
  if (!listings || !listings.length) return res.status(400).json({ error: "No listings provided" });

  // Write listings JSON to a temp file
  const tmpDir = os.tmpdir();
  const jsonPath = path.join(tmpDir, `listings_${Date.now()}.json`);
  const xlsxPath = path.join(tmpDir, `zillow_${Date.now()}.xlsx`);

  fs.writeFileSync(jsonPath, JSON.stringify(listings));

  // Python script to build the xlsx
  const pyScript = `
import json, sys, os, urllib.request, tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

json_path = sys.argv[1]
out_path  = sys.argv[2]

with open(json_path) as f:
    listings = json.load(f)

wb = Workbook()
ws = wb.active
ws.title = "Zillow Listings"

# ── column config ──────────────────────────────────────────────────────────
COLS = [
    ("Photo",         22),
    ("Address",       32),
    ("Price",         13),
    ("Beds",           7),
    ("Baths",          7),
    ("Sqft",          12),
    ("Lot Size",      13),
    ("Days Listed",   12),
    ("Offer Review",  16),
    ("HOA",           13),
    ("Sewer/Septic",  16),
    ("Zillow Link",   40),
]

IMG_COL   = 1   # A
LINK_COL  = 12  # L
ROW_H     = 80  # points (~107px)
HEADER_H  = 22

# header style
hdr_fill  = PatternFill("solid", fgColor="1A1A2A")
hdr_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
cell_align = Alignment(vertical="center", wrap_text=True)
thin = Side(style="thin", color="E2DDD6")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for ci, (label, width) in enumerate(COLS, start=1):
    ws.column_dimensions[get_column_letter(ci)].width = width
    c = ws.cell(row=1, column=ci, value=label)
    c.font = hdr_font
    c.fill = hdr_fill
    c.alignment = hdr_align
    c.border = border

ws.row_dimensions[1].height = HEADER_H
ws.freeze_panes = "A2"

# alternating row fills
fill_even = PatternFill("solid", fgColor="F7F5F0")
fill_odd  = PatternFill("solid", fgColor="FFFFFF")
link_font = Font(color="1A6B3C", underline="single", name="Arial", size=9)
data_font = Font(name="Arial", size=9)
price_font = Font(name="Arial", size=9, bold=True, color="134F2D")

img_cache = {}

def fetch_image(url):
    if url in img_cache:
        return img_cache[url]
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=8).read()
        suffix = ".jpg" if "jpg" in url.lower() or "jpeg" in url.lower() else ".png"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(data); tmp.close()
        img_cache[url] = tmp.name
        return tmp.name
    except Exception:
        img_cache[url] = None
        return None

for ri, listing in enumerate(listings, start=2):
    row = ri
    fill = fill_even if ri % 2 == 0 else fill_odd
    ws.row_dimensions[row].height = ROW_H

    # Photo col (A) — leave text empty, image anchored below
    ws.cell(row=row, column=1, value="").fill = fill

    # Address
    c = ws.cell(row=row, column=2, value=listing.get("address",""))
    c.font = data_font; c.alignment = cell_align; c.fill = fill; c.border = border

    # Price
    c = ws.cell(row=row, column=3, value=listing.get("price",""))
    c.font = price_font; c.alignment = Alignment(horizontal="right", vertical="center"); c.fill = fill; c.border = border

    # Beds / Baths
    for ci, key in [(4,"beds"),(5,"baths")]:
        c = ws.cell(row=row, column=ci, value=listing.get(key,""))
        c.font = data_font; c.alignment = Alignment(horizontal="center", vertical="center"); c.fill = fill; c.border = border

    # Sqft / Lot / Days / OfferReview / HOA / Sewer
    for ci, key in [(6,"sqft"),(7,"lot"),(8,"days"),(9,"offerReview"),(10,"hoa"),(11,"sewer")]:
        val = listing.get(key, "")
        if val == "—": val = ""
        c = ws.cell(row=row, column=ci, value=val)
        c.font = data_font; c.alignment = cell_align; c.fill = fill; c.border = border

    # Zillow hyperlink
    link = listing.get("zillowLink","")
    c = ws.cell(row=row, column=LINK_COL, value=listing.get("address","View Listing") if link else "")
    if link:
        c.hyperlink = link
        c.font = link_font
    else:
        c.font = data_font
    c.alignment = cell_align; c.fill = fill; c.border = border

    # Embed photo
    img_url = listing.get("img","")
    if img_url:
        img_path = fetch_image(img_url)
        if img_path:
            try:
                img = XLImage(img_path)
                # fit within cell: col A width ~22 chars = ~160px, row height ~107px
                img.width  = 120
                img.height = 90
                # anchor: col A, current row (0-indexed)
                cell_ref = f"A{row}"
                ws.add_image(img, cell_ref)
            except Exception:
                pass

# clean up temp json
try:
    for p in img_cache.values():
        if p: os.unlink(p)
except: pass

wb.save(out_path)
print("OK")
`;

  const pyPath = path.join(tmpDir, `build_xlsx_${Date.now()}.py`);
  fs.writeFileSync(pyPath, pyScript);

  execFile("python3", [pyPath, jsonPath, xlsxPath], { timeout: 120000 }, (err, stdout, stderr) => {
    // cleanup temp py + json
    try { fs.unlinkSync(pyPath); fs.unlinkSync(jsonPath); } catch (_) {}

    if (err) {
      console.error("Python error:", stderr);
      return res.status(500).json({ error: "Excel generation failed", detail: stderr });
    }

    if (!fs.existsSync(xlsxPath)) {
      return res.status(500).json({ error: "Excel file not created" });
    }

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=zillow_listings.xlsx");

    const stream = fs.createReadStream(xlsxPath);
    stream.pipe(res);
    stream.on("end", () => { try { fs.unlinkSync(xlsxPath); } catch (_) {} });
    stream.on("error", (e) => { res.status(500).json({ error: e.message }); });
  });
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
