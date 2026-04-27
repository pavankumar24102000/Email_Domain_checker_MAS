 
const express = require("express");
const multer  = require("multer");
const cors    = require("cors");
const xlsx    = require("xlsx");
const axios   = require("axios");
const cheerio = require("cheerio");
const path    = require("path");
const fs      = require("fs");

const app    = express();
const upload = multer({ dest: "uploads/" });
app.use(cors());

const SKIP_VALUES = ["row labels", "email", "grand total", "status", ""];
async function checkEmail(email) {
  return new Promise(async (resolve) => {
    
    // Force resolve after 6 seconds no matter what
    const timer = setTimeout(() => {
      console.log(`⏱️ Timeout forced for ${email}`);
      resolve({ status: "Timeout", color: "YELLOW" });
    }, 6000);

    try {
      const { data } = await axios.get(`https://verifymail.io/email/${email}`, {
        timeout: 5000,
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
          "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
          "Accept-Language": "en-US,en;q=0.5"
        }
      });

      clearTimeout(timer);
      const $ = cheerio.load(data);
      const h2 = $("h2.display-5.text-center").text().trim().toLowerCase();

      if (h2.includes("safe"))                                    return resolve({ status: "Safe",                 color: "GREEN"  });
      if (h2.includes("temporary") || h2.includes("disposable")) return resolve({ status: "Disposable/Temporary", color: "RED"    });

      const meta = $("meta[name='description']").attr("content") || "";
      if (meta.toLowerCase().includes("safe"))       return resolve({ status: "Safe",                 color: "GREEN"  });
      if (meta.toLowerCase().includes("disposable")) return resolve({ status: "Disposable/Temporary", color: "RED"    });

      resolve({ status: "Unknown", color: "YELLOW" });

    } catch (err) {
      clearTimeout(timer);
      console.log(`⚠️ Skipping ${email}: ${err.message}`);
      resolve({ status: "Failed", color: "YELLOW" });
    }
  });
}

app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheet    = workbook.Sheets[workbook.SheetNames[0]];
    const rows     = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const emails = rows
      .map(row => (row[0] || "").toString().trim())
      .filter(e => !SKIP_VALUES.includes(e.toLowerCase()) && e !== "");

    console.log(`📧 Processing ${emails.length} emails...`);

    const results = [];
    for (const email of emails) {
      console.log(`Checking: ${email}`);
      const result = await checkEmail(email);
      results.push({ email, ...result });
      await new Promise(r => setTimeout(r, 1000)); // 1s delay to avoid blocks
    }

    // Build output Excel
    const wb  = xlsx.utils.book_new();

    const wsData = [
      ["Email", "Status", "Result"],
      ...results.map(r => [r.email, r.status, r.color])
    ];
    xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(wsData), "Results");

    const safe       = results.filter(r => r.color === "GREEN").length;
    const disposable = results.filter(r => r.color === "RED").length;
    const unknown    = results.filter(r => r.color === "YELLOW").length;

    const summaryData = [
      ["Category",               "Count"],
      ["✅ Safe",                 safe],
      ["❌ Disposable/Temporary", disposable],
      ["⚠️ Unknown/Failed",       unknown],
      ["📧 Total",                results.length]
    ];
    xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(summaryData), "Summary");

    const outputPath = path.join("uploads", `output_${Date.now()}.xlsx`);
    xlsx.writeFile(wb, outputPath);

  const absPath = path.resolve(outputPath);
console.log("Sending file:", absPath);
res.download(absPath, "output_scrape.xlsx", (err) => {
      if (err) console.error("Download error:", err);
      try { fs.unlinkSync(req.file.path); } catch(e) {}
      try { fs.unlinkSync(outputPath); } catch(e) {}
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/", (req, res) => res.sendFile(path.resolve("index.html")));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));