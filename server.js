const express = require("express");
const cors = require("cors");
const https = require("https");

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

// Proxy endpoint — receives zillowUrl, page, and apiKey from the frontend
app.get("/api/search", (req, res) => {
  const { url, page, apiKey } = req.query;

  if (!url) return res.status(400).json({ error: "Missing url parameter" });
  if (!apiKey) return res.status(400).json({ error: "Missing apiKey parameter" });

  const encodedUrl = encodeURIComponent(url);
  const path = `/api/search/byurl?url=${encodedUrl}&page=${page || 1}`;

  const options = {
    hostname: "real-estate101.p.rapidapi.com",
    path,
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

  proxyReq.on("error", (err) => {
    res.status(500).json({ error: err.message });
  });

  proxyReq.end();
});

app.listen(PORT, () => console.log(`Proxy server running on port ${PORT}`));
