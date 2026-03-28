const express = require("express");
const { poolPromise, sql } = require("../db");
const { verifyToken } = require("../middleware/auth");

const router = express.Router();

// GET /api/subscriptions/plans
router.get("/plans", async (_req, res) => {
  try {
    const pool = await poolPromise;
    const result = await pool.request().query("SELECT id, name, features, price_monthly FROM lf_plans");
    res.json(result.recordset);
  } catch (err) {
    console.error("Plans error:", err.message);
    res.status(500).json({ error: "Failed to fetch plans" });
  }
});

// GET /api/subscriptions/current
router.get("/current", verifyToken, async (req, res) => {
  try {
    const pool = await poolPromise;

    const result = await pool
      .request()
      .input("tenantId", sql.Int, req.user.tenantId)
      .query(`
        SELECT s.id, s.status, s.expires_at, p.name AS plan, p.features
        FROM lf_subscriptions s
        JOIN lf_plans p ON s.plan_id = p.id
        WHERE s.tenant_id = @tenantId AND s.status = 'active'
        ORDER BY s.created_at DESC
      `);

    if (result.recordset.length === 0) {
      return res.json({ plan: "free", features: "reconcile,closing,vat" });
    }

    const sub = result.recordset[0];
    res.json({
      plan: sub.plan,
      features: sub.features.split(","),
      status: sub.status,
      expiresAt: sub.expires_at,
    });
  } catch (err) {
    console.error("Subscription error:", err.message);
    res.status(500).json({ error: "Failed to fetch subscription" });
  }
});

// POST /api/subscriptions/upgrade
router.post("/upgrade", verifyToken, async (req, res) => {
  const { planName } = req.body;

  if (!planName) {
    return res.status(400).json({ error: "planName is required" });
  }

  try {
    const pool = await poolPromise;

    const planResult = await pool
      .request()
      .input("planName", sql.NVarChar, planName)
      .query("SELECT id FROM lf_plans WHERE name = @planName");

    if (planResult.recordset.length === 0) {
      return res.status(404).json({ error: "Plan not found" });
    }

    const planId = planResult.recordset[0].id;

    // Deactivate existing subscription
    await pool
      .request()
      .input("tenantId", sql.Int, req.user.tenantId)
      .query("UPDATE lf_subscriptions SET status = 'cancelled' WHERE tenant_id = @tenantId AND status = 'active'");

    // Create new subscription
    await pool
      .request()
      .input("tenantId", sql.Int, req.user.tenantId)
      .input("planId", sql.Int, planId)
      .query("INSERT INTO lf_subscriptions (tenant_id, plan_id, status) VALUES (@tenantId, @planId, 'active')");

    res.json({ success: true, plan: planName });
  } catch (err) {
    console.error("Upgrade error:", err.message);
    res.status(500).json({ error: "Upgrade failed" });
  }
});

module.exports = router;
