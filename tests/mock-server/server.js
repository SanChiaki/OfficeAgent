// Mock SSO Login Server + Business API Server
// Usage: node server.js
//   SSO server  → http://localhost:3100
//   Business API → http://localhost:3200

const express = require("express");
const cookieParser = require("cookie-parser");

// ---------------------------------------------------------------------------
// Shared mock data
// ---------------------------------------------------------------------------

const performances = [
  { name: "张三", department: "销售部", score: 85, period: "2025-Q4" },
  { name: "李四", department: "技术部", score: 92, period: "2025-Q4" },
  { name: "王五", department: "市场部", score: 78, period: "2025-Q4" },
  { name: "赵六", department: "产品部", score: 88, period: "2025-Q4" },
];

const uploadedProjects = {};

// ---------------------------------------------------------------------------
// SSO Login Server :3100
// ---------------------------------------------------------------------------

const ssoApp = express();
ssoApp.use(express.urlencoded({ extended: true }));
ssoApp.use(cookieParser());

ssoApp.get("/login", (_req, res) => {
  res.type("html").send(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8"/>
  <title>SSO 登录</title>
  <style>
    body { font-family: "Segoe UI", sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; background: #f1f5f9; }
    .card { background: #fff; border-radius: 12px; padding: 32px; width: 360px; box-shadow: 0 2px 12px rgba(0,0,0,.1); }
    h1 { font-size: 1.25rem; margin: 0 0 24px; text-align: center; }
    label { display: block; margin-bottom: 4px; font-weight: 600; font-size: .9rem; }
    input { width: 100%; padding: 10px 12px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: .95rem; box-sizing: border-box; margin-bottom: 16px; }
    button { width: 100%; padding: 10px; background: #2563eb; color: #fff; border: none; border-radius: 8px; font-size: 1rem; cursor: pointer; }
    button:hover { background: #1d4ed8; }
    .error { color: #dc2626; font-size: .85rem; text-align: center; margin-bottom: 12px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>内网 SSO 登录</h1>
    ${_req.query.error ? '<p class="error">用户名和密码不能为空</p>' : ""}
    <form method="POST" action="/login">
      <label for="username">用户名</label>
      <input id="username" name="username" placeholder="admin"/>
      <label for="password">密码</label>
      <input id="password" name="password" type="password" placeholder="任意密码"/>
      <button type="submit">登 录</button>
    </form>
  </div>
</body>
</html>`);
});

ssoApp.post("/login", (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) {
    return res.redirect("/login?error=1");
  }

  // Set cookies visible to WebView2 CookieManager
  res.cookie("session_token", `tok_${username}_${Date.now()}`, {
    httpOnly: false,
    maxAge: 86400000,
  });
  res.cookie("user_name", username, {
    httpOnly: false,
    maxAge: 86400000,
  });

  // Redirect to business API (different port ⇒ different Authority ⇒ triggers cookie capture)
  res.redirect(302, "http://localhost:3200/logged-in");
});

ssoApp.listen(3100, () => {
  console.log("[SSO]      http://localhost:3100/login");
});

// ---------------------------------------------------------------------------
// Business API Server :3200
// ---------------------------------------------------------------------------

const apiApp = express();
apiApp.use(express.json());
apiApp.use(cookieParser());

// ---- Auth middleware for /api/* ----

function requireAuth(req, res, next) {
  if (!req.cookies?.session_token) {
    return res.status(401).json({ code: "unauthorized", message: "未登录，请先通过 SSO 登录。" });
  }
  next();
}

// ---- Landing page after SSO redirect ----

apiApp.get("/logged-in", (req, res) => {
  const user = req.cookies?.user_name || "未知用户";
  res.type("html").send(`<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="utf-8"/><title>登录成功</title>
<style>
  body { font-family: "Segoe UI", sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; background: #f0fdf4; }
  .card { background: #fff; border-radius: 12px; padding: 40px; text-align: center; box-shadow: 0 2px 12px rgba(0,0,0,.08); }
  h1 { color: #166534; font-size: 1.5rem; }
  p { color: #475569; }
</style>
</head>
<body><div class="card">
  <h1>✅ 登录成功</h1>
  <p>欢迎，${user}！此窗口将自动关闭。</p>
</div></body>
</html>`);
});

// ---- Performance API ----

apiApp.get("/api/performance", requireAuth, (_req, res) => {
  res.json(performances);
});

apiApp.get("/api/performance/:name", requireAuth, (req, res) => {
  const item = performances.find((p) => p.name === req.params.name);
  if (!item) {
    return res.status(404).json({ code: "not_found", message: `未找到「${req.params.name}」的绩效记录。` });
  }
  res.json(item);
});

apiApp.post("/api/performance", requireAuth, (req, res) => {
  const { name, department, score, period } = req.body || {};
  if (!name) {
    return res.status(400).json({ code: "bad_request", message: "name 字段必填。" });
  }

  const idx = performances.findIndex((p) => p.name === name);
  if (idx >= 0) {
    performances[idx] = { ...performances[idx], ...(department && { department }), ...(score != null && { score: Number(score) }), ...(period && { period }) };
    return res.json({ success: true, message: `已更新「${name}」的绩效。`, data: performances[idx] });
  }

  const entry = { name, department: department || "未知", score: Number(score) || 0, period: period || "2025-Q4" };
  performances.push(entry);
  res.json({ success: true, message: `已创建「${name}」的绩效。`, data: entry });
});

// ---- Upload data (compatible with existing upload_data skill) ----

apiApp.post("/upload_data", requireAuth, (req, res) => {
  const { projectName, records } = req.body || {};
  if (!projectName) {
    return res.status(400).json({ code: "bad_request", message: "projectName 字段必填。" });
  }
  if (!Array.isArray(records) || records.length === 0) {
    return res.status(400).json({ code: "bad_request", message: "records 必须是非空数组。" });
  }

  if (!uploadedProjects[projectName]) {
    uploadedProjects[projectName] = [];
  }
  uploadedProjects[projectName].push(...records);

  res.json({
    savedCount: records.length,
    message: `成功上传 ${records.length} 条记录到「${projectName}」。`,
  });
});

// ---- Download data ----

apiApp.get("/api/download/:projectName", requireAuth, (req, res) => {
  const { projectName } = req.params;

  if (projectName === "performance") {
    return res.json(performances);
  }

  const data = uploadedProjects[projectName];
  if (!data || data.length === 0) {
    return res.status(404).json({ code: "not_found", message: `项目「${projectName}」没有可下载的数据。` });
  }

  res.json(data);
});

// ---- Start ----

apiApp.listen(3200, () => {
  console.log("[Business] http://localhost:3200/api/performance");
  console.log("[Business] http://localhost:3200/api/download/:project");
  console.log("[Business] http://localhost:3200/upload_data");
  console.log("\nReady. Configure the add-in with:");
  console.log("  SSO URL  = http://localhost:3100/login");
  console.log("  Base URL = http://localhost:3200");
  console.log("  API Key  = (leave empty, uses SSO cookies)");
});
