const sql = require("mssql");
require("dotenv").config();

const config = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_NAME,
  options: {
    encrypt: true,
    trustServerCertificate: false,
  },
  pool: {
    max: 10,
    min: 0,
    idleTimeoutMillis: 30000,
  },
};

const poolPromise = new sql.ConnectionPool(config)
  .connect()
  .then(async (pool) => {
    console.log("SQL connected");
    await createTables(pool);
    return pool;
  })
  .catch((err) => {
    console.error("SQL connection failed:", err.message);
    process.exit(1);
  });

async function createTables(pool) {
  await pool.request().query(`
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='lf_tenants' AND xtype='U')
    CREATE TABLE lf_tenants (
      id INT IDENTITY(1,1) PRIMARY KEY,
      name NVARCHAR(255) NOT NULL,
      created_at DATETIME DEFAULT GETDATE()
    )
  `);

  await pool.request().query(`
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='lf_users' AND xtype='U')
    CREATE TABLE lf_users (
      id INT IDENTITY(1,1) PRIMARY KEY,
      tenant_id INT REFERENCES lf_tenants(id),
      email NVARCHAR(255) UNIQUE NOT NULL,
      password_hash NVARCHAR(255) NOT NULL,
      full_name NVARCHAR(255),
      role NVARCHAR(50) DEFAULT 'member',
      created_at DATETIME DEFAULT GETDATE()
    )
  `);

  await pool.request().query(`
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='lf_plans' AND xtype='U')
    CREATE TABLE lf_plans (
      id INT IDENTITY(1,1) PRIMARY KEY,
      name NVARCHAR(100) NOT NULL,
      features NVARCHAR(500) NOT NULL,
      price_monthly DECIMAL(10,2) DEFAULT 0,
      created_at DATETIME DEFAULT GETDATE()
    )
  `);

  await pool.request().query(`
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='lf_subscriptions' AND xtype='U')
    CREATE TABLE lf_subscriptions (
      id INT IDENTITY(1,1) PRIMARY KEY,
      tenant_id INT REFERENCES lf_tenants(id),
      plan_id INT REFERENCES lf_plans(id),
      status NVARCHAR(50) DEFAULT 'active',
      expires_at DATETIME,
      created_at DATETIME DEFAULT GETDATE()
    )
  `);

  // Seed default plans if empty
  const { recordset } = await pool.request().query("SELECT COUNT(*) AS cnt FROM lf_plans");
  if (recordset[0].cnt === 0) {
    await pool.request().query(`
      INSERT INTO lf_plans (name, features, price_monthly) VALUES
        ('free',  'reconcile,closing,vat', 0),
        ('pro',   'reconcile,closing,vat,payroll,assets,reports,budget', 19.99),
        ('firm',  'reconcile,closing,vat,payroll,assets,reports,budget,multiuser', 49.99)
    `);
    console.log("Default plans seeded");
  }
}

module.exports = { sql, poolPromise };
