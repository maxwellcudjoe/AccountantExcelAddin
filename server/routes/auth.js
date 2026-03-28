const express = require("express");
const bcrypt = require("bcryptjs");
const { poolPromise, sql } = require("../db");
const { signToken, verifyToken } = require("../middleware/auth");

const router = express.Router();

// POST /api/auth/register
router.post("/register", async (req, res) => {
  const { email, password, fullName, firmName } = req.body;

  if (!email || !password || !fullName) {
    return res.status(400).json({ error: "email, password and fullName are required" });
  }

  try {
    const pool = await poolPromise;

    // Check existing user
    const existing = await pool
      .request()
      .input("email", sql.NVarChar, email)
      .query("SELECT id FROM lf_users WHERE email = @email");

    if (existing.recordset.length > 0) {
      return res.status(409).json({ error: "Email already registered" });
    }

    // Create tenant
    const tenantResult = await pool
      .request()
      .input("name", sql.NVarChar, firmName || `${fullName}'s Firm`)
      .query("INSERT INTO lf_tenants (name) OUTPUT INSERTED.id VALUES (@name)");

    const tenantId = tenantResult.recordset[0].id;

    // Hash password
    const passwordHash = await bcrypt.hash(password, 12);

    // Create user
    const userResult = await pool
      .request()
      .input("tenantId", sql.Int, tenantId)
      .input("email", sql.NVarChar, email)
      .input("passwordHash", sql.NVarChar, passwordHash)
      .input("fullName", sql.NVarChar, fullName)
      .input("role", sql.NVarChar, "admin")
      .query(`
        INSERT INTO lf_users (tenant_id, email, password_hash, full_name, role)
        OUTPUT INSERTED.id
        VALUES (@tenantId, @email, @passwordHash, @fullName, @role)
      `);

    const userId = userResult.recordset[0].id;

    // Assign free plan
    const freePlan = await pool
      .request()
      .query("SELECT id FROM lf_plans WHERE name = 'free'");

    if (freePlan.recordset.length > 0) {
      await pool
        .request()
        .input("tenantId", sql.Int, tenantId)
        .input("planId", sql.Int, freePlan.recordset[0].id)
        .query("INSERT INTO lf_subscriptions (tenant_id, plan_id) VALUES (@tenantId, @planId)");
    }

    const token = signToken({ sub: userId, email, fullName, role: "admin", tenantId });
    res.status(201).json({ token });
  } catch (err) {
    console.error("Register error:", err.message);
    res.status(500).json({ error: "Registration failed" });
  }
});

// POST /api/auth/login
router.post("/login", async (req, res) => {
  const { email, password } = req.body;

  if (!email || !password) {
    return res.status(400).json({ error: "email and password are required" });
  }

  try {
    const pool = await poolPromise;

    const result = await pool
      .request()
      .input("email", sql.NVarChar, email)
      .query("SELECT id, tenant_id, password_hash, full_name, role FROM lf_users WHERE email = @email");

    if (result.recordset.length === 0) {
      return res.status(401).json({ error: "Invalid email or password" });
    }

    const user = result.recordset[0];
    const valid = await bcrypt.compare(password, user.password_hash);
    if (!valid) {
      return res.status(401).json({ error: "Invalid email or password" });
    }

    const token = signToken({
      sub: user.id,
      email,
      fullName: user.full_name,
      role: user.role,
      tenantId: user.tenant_id,
    });

    res.json({ token });
  } catch (err) {
    console.error("Login error:", err.message);
    res.status(500).json({ error: "Login failed" });
  }
});

// GET /api/auth/me
router.get("/me", verifyToken, async (req, res) => {
  try {
    const pool = await poolPromise;

    const result = await pool
      .request()
      .input("id", sql.Int, req.user.sub)
      .query("SELECT id, email, full_name, role, tenant_id FROM lf_users WHERE id = @id");

    if (result.recordset.length === 0) {
      return res.status(404).json({ error: "User not found" });
    }

    res.json(result.recordset[0]);
  } catch (err) {
    console.error("Me error:", err.message);
    res.status(500).json({ error: "Failed to fetch user" });
  }
});

module.exports = router;
