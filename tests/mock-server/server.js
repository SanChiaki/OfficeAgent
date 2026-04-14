// Mock SSO Login Server + Business API Server
// Usage: node server.js
//   SSO server  -> http://localhost:3100
//   Business API -> http://localhost:3200
//
// Login flow:
//   1. GET  /login         -> renders login form page
//   2. On form submit, page fetch(POST) /rest/login
//   3. POST /rest/login    -> returns 200 with Set-Cookie
//      C# SsoLoginPopup uses WebResourceResponseReceived to detect /rest/login returning 200.

const express = require("express");
const cookieParser = require("cookie-parser");
const fs = require("fs");
const path = require("path");

const performances = [
  { name: "张三", department: "销售部", score: 85, period: "2025-Q4" },
  { name: "李四", department: "技术部", score: 92, period: "2025-Q4" },
  { name: "王五", department: "市场部", score: 78, period: "2025-Q4" },
  { name: "赵六", department: "产品部", score: 88, period: "2025-Q4" },
];

const uploadedProjects = {};

const connectorRows = [
  { row_id: "row-1", owner_name: "张三", start_12345678: "2026-01-02", end_12345678: "2026-01-05" },
  { row_id: "row-2", owner_name: "李四", start_12345678: "2026-01-10", end_12345678: "2026-01-15" },
];

const connectorHeadList = [
  { fieldKey: "row_id", headerText: "ID", headType: "single", isId: true },
  { fieldKey: "owner_name", headerText: "负责人", headType: "single" },
  {
    headType: "activity",
    activityId: "12345678",
    activityName: "测试活动111",
  },
];

// ---------------------------------------------------------------------------
// SSO Login Server :3100
// ---------------------------------------------------------------------------

const ssoApp = express();
ssoApp.use(express.urlencoded({ extended: true }));
ssoApp.use(express.json());
ssoApp.use(cookieParser());

var loginPageHtml = fs.readFileSync(path.join(__dirname, "login-page.html"), "utf8");

ssoApp.get("/login", function (_req, res) {
  res.type("html").send(loginPageHtml);
});

// The endpoint C# detects as the login success marker.
// Returns 200 + Set-Cookie
ssoApp.post("/rest/login", function (req, res) {
  const { username, password } = req.body || {};
  if (!username || !password) {
    return res.status(400).json({ error: "用户名和密码不能为空。" });
  }
  res.cookie("session_token", "tok_" + username + "_" + Date.now(), {
    httpOnly: false,
    maxAge: 86400000,
    path: "/",
  });
  res.cookie("user_name", username, {
    httpOnly: false,
    maxAge: 86400000,
    path: "/",
  });
  res.json({ ok: true, user_name: username });
});

ssoApp.listen(3100, function () {
  console.log("[SSO]      http://localhost:3100/login");
});

// ---------------------------------------------------------------------------
// Business API Server :3200
// ---------------------------------------------------------------------------

const apiApp = express();
apiApp.use(express.json());
apiApp.use(cookieParser());

function requireAuth(req, res, next) {
  if (!req.cookies || !req.cookies.session_token) {
    return res.status(401).json({ code: "unauthorized", message: "未登录，请先通过 SSO 登录。" });
  }
  next();
}

apiApp.get("/logged-in", function (req, res) {
  var user = (req.cookies && req.cookies.user_name) || "未知用户";
  res.type("html").send(renderLoggedIn(user));
});

apiApp.get("/api/performance", requireAuth, function (_req, res) {
  res.json(performances);
});

apiApp.get("/api/performance/:name", requireAuth, function (req, res) {
  var item = performances.find(function (p) { return p.name === req.params.name; });
  if (!item) {
    return res.status(404).json({ code: "not_found", message: '未找到\u300c' + req.params.name + '\u300d的绩效记录。' });
  }
  res.json(item);
});

apiApp.post("/api/performance", requireAuth, function (req, res) {
  var name = (req.body || {}).name;
  if (!name) {
    return res.status(400).json({ code: "bad_request", message: "name 字段必填。" });
  }
  var department = req.body.department;
  var score = req.body.score;
  var period = req.body.period;
  var idx = performances.findIndex(function (p) { return p.name === name; });
  if (idx >= 0) {
    var cur = performances[idx];
    if (department) cur.department = department;
    if (score != null) cur.score = Number(score);
    if (period) cur.period = period;
    return res.json({ success: true, message: '已更新\u300c' + name + '\u300d的绩效。', data: cur });
  }
  var entry = {
    name: name,
    department: department || "未知",
    score: Number(score) || 0,
    period: period || "2025-Q4",
  };
  performances.push(entry);
  res.json({ success: true, message: '已创建\u300c' + name + '\u300d的绩效。', data: entry });
});

apiApp.post("/upload_data", requireAuth, function (req, res) {
  var projectName = (req.body || {}).projectName;
  var records = (req.body || {}).records;
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
    message: '成功上传 ' + records.length + ' 条记录到\u300c' + projectName + '\u300d。',
  });
});

apiApp.post("/head", requireAuth, function (_req, res) {
  res.json({ headList: connectorHeadList });
});

apiApp.post("/find", requireAuth, function (req, res) {
  var ids = Array.isArray(req.body?.ids) ? req.body.ids : [];
  var fieldKeys = Array.isArray(req.body?.fieldKeys) ? req.body.fieldKeys : [];
  var result = connectorRows;

  if (ids.length > 0) {
    result = result.filter(function (row) {
      return ids.indexOf(row.row_id) >= 0;
    });
  }

  if (fieldKeys.length > 0) {
    result = result.map(function (row) {
      var filtered = { row_id: row.row_id };
      fieldKeys.forEach(function (key) {
        if (Object.prototype.hasOwnProperty.call(row, key)) {
          filtered[key] = row[key];
        }
      });
      if (!Object.prototype.hasOwnProperty.call(filtered, "row_id")) {
        filtered.row_id = row.row_id;
      }
      return filtered;
    });
  }

  res.json(result);
});

apiApp.post("/batchSave", requireAuth, function (req, res) {
  var items = Array.isArray(req.body) ? req.body : [];
  if (items.length === 0) {
    return res.status(400).json({ code: "bad_request", message: "items 必须为非空数组。" });
  }
  items.forEach(function (item) {
    if (!item) {
      return;
    }

    var rowId = item.id || item.Id;
    var fieldKey = item.fieldKey || item.FieldKey;
    var value = item.value != null ? item.value : item.Value;
    if (!rowId || !fieldKey) {
      return;
    }

    var target = connectorRows.find(function (row) { return row.row_id === rowId; });
    if (target) {
      target[fieldKey] = value != null ? value : "";
    }
  });

  res.json({
    savedCount: items.length,
    message: "Received batchSave with " + items.length + " item(s).",
  });
});

apiApp.get("/api/download/:projectName", requireAuth, function (req, res) {
  var projectName = req.params.projectName;
  if (projectName === "performance") {
    return res.json(performances);
  }
  var data = uploadedProjects[projectName];
  if (!data || data.length === 0) {
    return res.status(404).json({ code: "not_found", message: '项目\u300c' + projectName + '\u300d没有可下载的数据。' });
  }
  res.json(data);
});

apiApp.listen(3200, function () {
  console.log("[Business] http://localhost:3200/api/performance");
  console.log("[Business] http://localhost:3200/api/download/:project");
  console.log("[Business] http://localhost:3200/upload_data");
  console.log("\nReady. Configure the add-in with:");
  console.log("  Base URL              = <LLM service URL>");
  console.log("  Business Base URL     = http://localhost:3200");
  console.log("  SSO URL               = http://localhost:3100/login");
  console.log("  登录成功路径           = /rest/login");
  console.log("  API Key               = (留空，使用 SSO cookies)");
});

function renderLoggedIn(user) {
  return '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"/><title>登录成功</title>'
    + '<style>body{font-family:"Segoe UI",sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:#f0fdf4;}'
    + '.card{background:#fff;border-radius:12px;padding:40px;text-align:center;box-shadow:0 2px 12px rgba(0,0,0,.08);}'
    + 'h1{color:#166534;font-size:1.5rem;}p{color:#475569;}</style></head>'
    + '<body><div class="card"><h1>✅ 登录成功</h1>'
    + '<p>欢迎，' + user + '！此窗口将自动关闭。</p></div></body></html>';
}
