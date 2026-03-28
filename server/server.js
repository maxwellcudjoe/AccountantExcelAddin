const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const https = require("https");
const http = require("http");
require("dotenv").config();

const authRoutes = require("./routes/auth");
const subscriptionRoutes = require("./routes/subscriptions");

const app = express();

// ── CORS ─────────────────────────────────────────────────────────────────────
const allowedOrigins = [
  "https://localhost:3000",
  "https://ledgerflow-pro.azurewebsites.net",
];

app.use(
  cors({
    origin: (origin, callback) => {
      if (!origin || allowedOrigins.includes(origin)) {
        callback(null, true);
      } else {
        callback(new Error(`CORS policy: origin ${origin} not allowed`));
      }
    },
    credentials: true,
  })
);

// ── Body parsing ──────────────────────────────────────────────────────────────
app.use(express.json());

// ── Routes ────────────────────────────────────────────────────────────────────
app.use("/api/auth", authRoutes);
app.use("/api/subscriptions", subscriptionRoutes);

// ── Health check ──────────────────────────────────────────────────────────────
app.get("/api/health", (_req, res) => res.json({ status: "ok", app: "LedgerFlow Pro" }));

// ── Static files (production) ─────────────────────────────────────────────────
const distPath = path.join(__dirname, "../dist");
if (fs.existsSync(distPath)) {
  app.use(express.static(distPath));
  app.get("*", (_req, res) => res.sendFile(path.join(distPath, "taskpane.html")));
}

// ── Server startup ────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 7264;

// Azure IISNode passes a named pipe as PORT
if (typeof PORT === "string" && PORT.includes("\\")) {
  http.createServer(app).listen(PORT, () => {
    console.log(`LedgerFlow server listening on named pipe`);
  });
} else {
  // Try HTTPS dev certs locally
  const certDir = path.join(
    process.env.USERPROFILE || process.env.HOME || "",
    ".office-addin-dev-certs"
  );
  const keyPath = path.join(certDir, "localhost.key");
  const certPath = path.join(certDir, "localhost.crt");

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    https
      .createServer({ key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }, app)
      .listen(PORT, () => console.log(`LedgerFlow HTTPS server running on port ${PORT}`));
  } else {
    http
      .createServer(app)
      .listen(PORT, () => console.log(`LedgerFlow HTTP server running on port ${PORT}`));
  }
}
