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

const connectorProjectData = {
  performance: createConnectorProject(
    "performance",
    "绩效项目",
    "12345678",
    "测试活动111",
    [
      { rowId: "row-1", ownerName: "张三", startDate: "2026-01-02", endDate: "2026-01-05" },
      { rowId: "row-2", ownerName: "李四", startDate: "2026-01-10", endDate: "2026-01-15" },
    ]),
  "delivery-tracker": createConnectorProject(
    "delivery-tracker",
    "交付跟踪项目",
    "22334455",
    "交付阶段",
    [
      { rowId: "delivery-row-1", ownerName: "交付一组", startDate: "2026-02-01", endDate: "2026-02-03" },
      { rowId: "delivery-row-2", ownerName: "交付二组", startDate: "2026-02-04", endDate: "2026-02-06" },
      { rowId: "delivery-row-3", ownerName: "交付三组", startDate: "2026-02-07", endDate: "2026-02-09" },
      { rowId: "delivery-row-4", ownerName: "交付四组", startDate: "2026-02-10", endDate: "2026-02-12" },
      { rowId: "delivery-row-5", ownerName: "交付五组", startDate: "2026-02-13", endDate: "2026-02-15" },
      { rowId: "delivery-row-6", ownerName: "交付六组", startDate: "2026-02-16", endDate: "2026-02-18" },
      { rowId: "delivery-row-7", ownerName: "交付七组", startDate: "2026-02-19", endDate: "2026-02-21" },
      { rowId: "delivery-row-8", ownerName: "交付八组", startDate: "2026-02-22", endDate: "2026-02-24" },
      { rowId: "delivery-row-9", ownerName: "交付九组", startDate: "2026-02-25", endDate: "2026-02-27" },
      { rowId: "delivery-row-10", ownerName: "交付十组", startDate: "2026-02-28", endDate: "2026-03-02" },
    ]),
  "customer-onboarding": createConnectorProject(
    "customer-onboarding",
    "客户上线项目",
    "99887766",
    "上线流程",
    [
      { rowId: "onboarding-row-1", ownerName: "客户成功一组", startDate: "2026-03-01", endDate: "2026-03-03" },
      { rowId: "onboarding-row-2", ownerName: "客户成功二组", startDate: "2026-03-04", endDate: "2026-03-06" },
      { rowId: "onboarding-row-3", ownerName: "客户成功三组", startDate: "2026-03-07", endDate: "2026-03-09" },
      { rowId: "onboarding-row-4", ownerName: "客户成功四组", startDate: "2026-03-10", endDate: "2026-03-12" },
      { rowId: "onboarding-row-5", ownerName: "客户成功五组", startDate: "2026-03-13", endDate: "2026-03-15" },
      { rowId: "onboarding-row-6", ownerName: "客户成功六组", startDate: "2026-03-16", endDate: "2026-03-18" },
      { rowId: "onboarding-row-7", ownerName: "客户成功七组", startDate: "2026-03-19", endDate: "2026-03-21" },
      { rowId: "onboarding-row-8", ownerName: "客户成功八组", startDate: "2026-03-22", endDate: "2026-03-24" },
      { rowId: "onboarding-row-9", ownerName: "客户成功九组", startDate: "2026-03-25", endDate: "2026-03-27" },
      { rowId: "onboarding-row-10", ownerName: "客户成功十组", startDate: "2026-03-28", endDate: "2026-03-30" },
    ]),
  "large-activity-benchmark": createLargeActivityBenchmarkProject(),
};

const connectorProjects = Object.keys(connectorProjectData).map(function (projectId) {
  var project = connectorProjectData[projectId];
  return {
    projectId: project.projectId,
    displayName: project.displayName,
  };
});

function createConnectorProject(projectId, displayName, activityId, activityName, rows) {
  return {
    projectId: projectId,
    displayName: displayName,
    headList: createConnectorHeadList(activityId, activityName),
    rows: createConnectorRows(activityId, rows),
  };
}

function createConnectorHeadList(activityId, activityName) {
  return [
    { fieldKey: "row_id", headerText: "ID", headType: "single", isId: true },
    { fieldKey: "owner_name", headerText: "负责人", headType: "single" },
    {
      headType: "activity",
      activityId: activityId,
      activityName: activityName,
    },
  ];
}

function createConnectorRows(activityId, rows) {
  var startFieldKey = "start_" + activityId;
  var endFieldKey = "end_" + activityId;

  return rows.map(function (row) {
    return {
      row_id: row.rowId,
      owner_name: row.ownerName,
      [startFieldKey]: row.startDate,
      [endFieldKey]: row.endDate,
    };
  });
}

function createLargeActivityBenchmarkProject() {
  return {
    projectId: "large-activity-benchmark",
    displayName: "大数据活动压测项目",
    headList: createLargeActivityBenchmarkHeadList(),
    rows: createLargeActivityBenchmarkRows(10000),
  };
}

function createLargeActivityBenchmarkHeadList() {
  return [
    { fieldKey: "row_id", headerText: "ID", headType: "single", isId: true },
    { fieldKey: "owner_name", headerText: "负责人", headType: "single" },
    { fieldKey: "region", headerText: "区域", headType: "single" },
    { fieldKey: "priority", headerText: "优先级", headType: "single" },
    { fieldKey: "status", headerText: "状态", headType: "single" },
    { headType: "activity", activityId: "benchmarka", activityName: "基准阶段A" },
    { headType: "activity", activityId: "benchmarkb", activityName: "基准阶段B" },
  ];
}

function createLargeActivityBenchmarkRows(rowCount) {
  var regions = ["华东", "华北", "华南", "西南", "海外"];
  var priorities = ["P0", "P1", "P2", "P3"];
  var statuses = ["未开始", "进行中", "已完成", "已暂停"];
  var rows = [];

  for (var index = 0; index < rowCount; index++) {
    var rowNumber = index + 1;
    var benchmarkRowId = "benchmark-row-" + String(rowNumber).padStart(5, "0");
    var ownerName = "区域负责人" + String((index % 300) + 1).padStart(3, "0");
    var dayOffset = index % 28;

    rows.push({
      row_id: benchmarkRowId,
      owner_name: ownerName,
      region: regions[index % regions.length],
      priority: priorities[index % priorities.length],
      status: statuses[index % statuses.length],
      name_benchmarka: "阶段A-" + String(rowNumber).padStart(5, "0"),
      start_benchmarka: createIsoDate(2026, 1, 1 + dayOffset),
      end_benchmarka: createIsoDate(2026, 1, 3 + dayOffset),
      name_benchmarkb: "阶段B-" + String(rowNumber).padStart(5, "0"),
      start_benchmarkb: createIsoDate(2026, 1, 4 + dayOffset),
      end_benchmarkb: createIsoDate(2026, 1, 6 + dayOffset),
    });
  }

  return rows;
}

function createIsoDate(year, month, day) {
  var date = new Date(Date.UTC(year, month - 1, day));
  return date.toISOString().slice(0, 10);
}

function getConnectorProject(projectId) {
  if (!projectId) {
    return null;
  }

  return connectorProjectData[projectId] || null;
}

function resolveConnectorProject(projectId, res) {
  if (!projectId) {
    res.status(400).json({ code: "bad_request", message: "projectId 字段必填。" });
    return null;
  }

  var project = getConnectorProject(projectId);
  if (!project) {
    res.status(404).json({ code: "not_found", message: '未找到项目\u300c' + projectId + '\u300d。' });
    return null;
  }

  return project;
}

function getBatchSaveItems(body) {
  if (Array.isArray(body)) {
    return body;
  }

  if (body && Array.isArray(body.items)) {
    return body.items;
  }

  return [];
}

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
apiApp.use(express.json({ limit: "20mb" }));
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

apiApp.get("/projects", requireAuth, function (_req, res) {
  res.json(connectorProjects);
});

apiApp.post("/head", requireAuth, function (req, res) {
  var project = resolveConnectorProject((req.body || {}).projectId, res);
  if (!project) {
    return;
  }

  res.json({ headList: project.headList });
});

apiApp.post("/find", requireAuth, function (req, res) {
  var project = resolveConnectorProject((req.body || {}).projectId, res);
  if (!project) {
    return;
  }

  var ids = Array.isArray(req.body?.ids)
    ? req.body.ids
    : Array.isArray(req.body?.rowIds)
      ? req.body.rowIds
      : [];
  var fieldKeys = Array.isArray(req.body?.fieldKeys) ? req.body.fieldKeys : [];
  var result = project.rows;

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
  var items = getBatchSaveItems(req.body);
  if (items.length === 0) {
    return res.status(400).json({ code: "bad_request", message: "items 必须为非空数组。" });
  }

  for (var i = 0; i < items.length; i++) {
    var batchItem = items[i];
    var batchProjectId = batchItem && (batchItem.projectId || batchItem.ProjectId);
    if (!batchProjectId) {
      return res.status(400).json({ code: "bad_request", message: "batchSave item.projectId 字段必填。" });
    }

    if (!getConnectorProject(batchProjectId)) {
      return res.status(404).json({ code: "not_found", message: '未找到项目\u300c' + batchProjectId + '\u300d。' });
    }
  }

  items.forEach(function (item) {
    if (!item) {
      return;
    }

    var projectId = item.projectId || item.ProjectId;
    var rowId = item.id || item.Id;
    var fieldKey = item.fieldKey || item.FieldKey;
    var value = item.value != null ? item.value : item.Value;
    if (!rowId || !fieldKey) {
      return;
    }

    var project = getConnectorProject(projectId);
    var target = project.rows.find(function (row) { return row.row_id === rowId; });
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
  console.log("[Business] http://localhost:3200/projects");
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
