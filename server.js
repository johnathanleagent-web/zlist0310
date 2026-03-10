const express = require("express");
const cors = require("cors");
const https = require("https");
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

// ── Check Python availability ─────────────────────────────────────────────────
app.get("/api/health", (req, res) => {
  const { execSync } = require("child_process");
  const results = {};
  ["python3", "python"].forEach(cmd => {
    try {
      const ver = execSync(`${cmd} --version 2>&1`).toString().trim();
      results[cmd] = ver;
    } catch (e) {
      results[cmd] = "not found";
    }
  });
  // check openpyxl
  try {
    const check = execSync(`python3 -c "import openpyxl; print(openpyxl.__version__)" 2>&1`).toString().trim();
    results["openpyxl"] = check;
  } catch (e) {
    results["openpyxl"] = "not installed: " + e.message;
  }
  try {
    const check = execSync(`python3 -c "import PIL; print(PIL.__version__)" 2>&1`).toString().trim();
    results["Pillow"] = check;
  } catch (e) {
    results["Pillow"] = "not installed";
  }
  res.json(results);
});

// ── Excel export ──────────────────────────────────────────────────────────────
app.post("/api/export/xlsx", (req, res) => {
  const { listings } = req.body;
  if (!listings || !listings.length) return res.status(400).json({ error: "No listings provided" });

  const tmpDir = os.tmpdir();
  const ts = Date.now();
  const jsonPath = path.join(tmpDir, `listings_${ts}.json`);
  const xlsxPath = path.join(tmpDir, `zillow_${ts}.xlsx`);
  const pyPath   = path.join(tmpDir, `build_xlsx_${ts}.py`);

  fs.writeFileSync(jsonPath, JSON.stringify(listings));

  const pyScript = `
import json, sys, os, urllib.request, tempfile, traceback

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
except ImportError as e:
    print("IMPORT_ERROR:" + str(e), flush=True)
    sys.exit(1)

json_path = sys.argv[1]
out_path  = sys.argv[2]

try:
    with open(json_path) as f:
        listings = json.load(f)

    wb = Workbook()
    ws = wb.active
    ws.title = "Zillow Listings"

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
        ("Zillow Link",   45),
    ]

    ROW_H    = 75
    HEADER_H = 22

    hdr_fill  = PatternFill("solid", fgColor="1A1A2A")
    hdr_font  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_aln  = Alignment(vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="E2DDD6")
    bdr       = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_even = PatternFill("solid", fgColor="F7F5F0")
    fill_odd  = PatternFill("solid", fgColor="FFFFFF")
    link_font = Font(color="1A6B3C", underline="single", name="Arial", size=9)
    data_font = Font(name="Arial", size=9)
    price_fnt = Font(name="Arial", size=9, bold=True, color="134F2D")

    for ci, (label, width) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = width
        c = ws.cell(row=1, column=ci, value=label)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = hdr_align
        c.border = bdr

    ws.row_dimensions[1].height = HEADER_H
    ws.freeze_panes = "A2"

    img_cache = {}

    def fetch_image(url):
        if not url:
            return None
        if url in img_cache:
            return img_cache[url]
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            data = urllib.request.urlopen(req, timeout=8).read()
            suffix = ".jpg"
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(data)
            tmp.close()
            img_cache[url] = tmp.name
            return tmp.name
        except Exception as e:
            print(f"IMG_SKIP {url[:60]}: {e}", flush=True)
            img_cache[url] = None
            return None

    for ri, listing in enumerate(listings, start=2):
        fill = fill_even if ri % 2 == 0 else fill_odd
        ws.row_dimensions[ri].height = ROW_H

        # Col A: photo placeholder
        ws.cell(row=ri, column=1, value="").fill = fill

        # Col B: address
        c = ws.cell(row=ri, column=2, value=listing.get("address", ""))
        c.font = data_font; c.alignment = cell_aln; c.fill = fill; c.border = bdr

        # Col C: price
        c = ws.cell(row=ri, column=3, value=listing.get("price", ""))
        c.font = price_fnt
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.fill = fill; c.border = bdr

        # Cols D-E: beds/baths
        for ci, key in [(4, "beds"), (5, "baths")]:
            c = ws.cell(row=ri, column=ci, value=str(listing.get(key, "") or ""))
            c.font = data_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = fill; c.border = bdr

        # Cols F-K: sqft, lot, days, offerReview, hoa, sewer
        for ci, key in [(6,"sqft"),(7,"lot"),(8,"days"),(9,"offerReview"),(10,"hoa"),(11,"sewer")]:
            val = listing.get(key, "") or ""
            if val == "\\u2014" or val == "-": val = ""
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = data_font; c.alignment = cell_aln; c.fill = fill; c.border = bdr

        # Col L: hyperlink
        link = listing.get("zillowLink", "") or ""
        addr = listing.get("address", "View on Zillow") or "View on Zillow"
        c = ws.cell(row=ri, column=12, value=addr if link else "")
        if link:
            c.hyperlink = link
            c.font = link_font
        else:
            c.font = data_font
        c.alignment = cell_aln; c.fill = fill; c.border = bdr

        # Embed photo in col A
        img_url = listing.get("img", "") or ""
        if img_url:
            img_path = fetch_image(img_url)
            if img_path:
                try:
                    img_obj = XLImage(img_path)
                    img_obj.width  = 115
                    img_obj.height = 80
                    ws.add_image(img_obj, f"A{ri}")
                except Exception as e:
                    print(f"IMG_EMBED_ERR row {ri}: {e}", flush=True)

    # cleanup temp images
    for p in img_cache.values():
        if p:
            try: os.unlink(p)
            except: pass

    wb.save(out_path)
    print("SUCCESS", flush=True)

except Exception as e:
    print("SCRIPT_ERROR:" + traceback.format_exc(), flush=True)
    sys.exit(1)
`;

  fs.writeFileSync(pyPath, pyScript);

  // Try python3 first, fall back to python
  const tryExport = (cmd) => {
    execFile(cmd, [pyPath, jsonPath, xlsxPath], { timeout: 180000 }, (err, stdout, stderr) => {
      const output = (stdout || "") + (stderr || "");
      console.log(`[xlsx] ${cmd} stdout:`, stdout);
      console.log(`[xlsx] ${cmd} stderr:`, stderr);

      if (err && cmd === "python3") {
        console.log("[xlsx] python3 failed, trying python...");
        return tryExport("python");
      }

      try { fs.unlinkSync(pyPath); } catch (_) {}
      try { fs.unlinkSync(jsonPath); } catch (_) {}

      if (err) {
        return res.status(500).json({ error: "Python not available", detail: output });
      }

      if (output.includes("IMPORT_ERROR")) {
        // openpyxl not installed — try to install it inline
        const { execSync } = require("child_process");
        try {
          console.log("[xlsx] Installing openpyxl...");
          execSync("pip install openpyxl Pillow --quiet 2>&1", { timeout: 60000 });
          // retry
          return tryExport(cmd);
        } catch (installErr) {
          return res.status(500).json({ error: "openpyxl not installed and auto-install failed", detail: installErr.message });
        }
      }

      if (!fs.existsSync(xlsxPath)) {
        return res.status(500).json({ error: "Excel file not created", detail: output });
      }

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=zillow_listings.xlsx");

      const stream = fs.createReadStream(xlsxPath);
      stream.pipe(res);
      stream.on("end", () => { try { fs.unlinkSync(xlsxPath); } catch (_) {} });
      stream.on("error", (e) => res.status(500).json({ error: e.message }));
    });
  };

  tryExport("python3");
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
