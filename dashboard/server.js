const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const path = require('path');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Agent definitions from CLAUDE.md ──
const AGENTS = {
  'architecture-design': {
    name: '建築設計部',
    nameEn: 'Architecture & Design',
    agents: [
      { id: 'design-lead', name: '設計リード 高橋', role: '設計全体の統括・方針決定' },
      { id: 'drawing-specialist', name: '図面スペシャリスト 佐藤', role: '図面製作・CAD作業' },
      { id: 'grasshopper-modeler', name: 'Grasshopperモデラー 田中', role: 'パラメトリックモデリング・GH定義' },
      { id: 'design-reviewer', name: '設計レビュアー 山本', role: '設計品質チェック・法規確認' },
    ],
  },
  'manufacturing': {
    name: '製造・加工部',
    nameEn: 'Manufacturing',
    agents: [
      { id: 'cnc-programmer', name: 'CNCプログラマー 中村', role: 'CNC加工データ・Gコード作成' },
      { id: 'machining-planner', name: '加工計画担当 木村', role: '加工手順・材料取り計画' },
      { id: 'robot-arm-specialist', name: 'ロボットアーム担当 林', role: 'ロボットアーム制御・ティーチング' },
      { id: 'quality-inspector', name: '品質検査担当 渡辺', role: '加工品質検査・精度管理' },
    ],
  },
  'project-management': {
    name: 'プロジェクト管理部',
    nameEn: 'Project Management',
    agents: [
      { id: 'project-manager', name: 'プロジェクトマネージャー 鈴木', role: 'PJ全体統括・意思決定支援' },
      { id: 'schedule-coordinator', name: 'スケジュール調整担当 伊藤', role: '工程表・マイルストーン管理' },
      { id: 'risk-analyst', name: 'リスク分析担当 小林', role: 'リスク特定・対策立案' },
      { id: 'stakeholder-communicator', name: 'ステークホルダー対応 加藤', role: '顧客・協力会社との連絡' },
    ],
  },
  'business-strategy': {
    name: '事業戦略・営業部',
    nameEn: 'Business Strategy',
    agents: [
      { id: 'sales-strategist', name: '営業戦略担当 吉田', role: '営業戦略・顧客開拓' },
      { id: 'market-analyst', name: '市場分析担当 松本', role: '市場調査・競合分析' },
      { id: 'proposal-writer', name: '提案書作成担当 井上', role: '提案書・見積書作成' },
      { id: 'partnership-developer', name: 'パートナーシップ担当 山田', role: '連携先開拓・関係構築' },
    ],
  },
  'ai-technology': {
    name: 'AI・テクノロジー部',
    nameEn: 'AI & Technology',
    agents: [
      { id: 'ai-product-developer', name: 'AIプロダクト開発者 藤田', role: '社内AIツール開発' },
      { id: 'ai-strategy-planner', name: 'AI戦略プランナー 石井', role: 'AI事業戦略・ロードマップ' },
      { id: 'tech-researcher', name: '技術リサーチャー 前田', role: '最新技術調査・PoC' },
      { id: 'system-architect', name: 'システムアーキテクト 岡田', role: 'システム設計・インフラ' },
    ],
  },
  'communications': {
    name: '広報・コミュニケーション部',
    nameEn: 'Communications',
    agents: [
      { id: 'instagram-manager', name: 'Instagram運用担当 西村', role: 'Instagram戦略・投稿管理' },
      { id: 'content-creator', name: 'コンテンツクリエイター 三浦', role: 'テキスト・ビジュアル制作' },
      { id: 'brand-guardian', name: 'ブランドガーディアン 太田', role: 'ブランド一貫性・品質管理' },
    ],
  },
  'operations': {
    name: '業務改善・オペレーション部',
    nameEn: 'Operations',
    agents: [
      { id: 'workflow-designer', name: 'ワークフロー設計担当 藤原', role: '業務フロー設計・改善' },
      { id: 'process-analyst', name: 'プロセス分析担当 原田', role: '現状分析・ボトルネック特定' },
      { id: 'automation-engineer', name: '自動化エンジニア 村上', role: '業務自動化・ツール連携' },
      { id: 'documentation-manager', name: 'ドキュメント管理担当 清水', role: '社内ドキュメント整備' },
    ],
  },
};

// ── In-memory state ──
let tasks = loadJSON('tasks.json', []);
let logs = loadJSON('logs.json', []);

function loadJSON(filename, fallback) {
  const filepath = path.join(DATA_DIR, filename);
  try {
    if (fs.existsSync(filepath)) {
      return JSON.parse(fs.readFileSync(filepath, 'utf8'));
    }
  } catch { /* ignore */ }
  return fallback;
}

function saveJSON(filename, data) {
  fs.writeFileSync(path.join(DATA_DIR, filename), JSON.stringify(data, null, 2));
}

function broadcast(type, payload) {
  const msg = JSON.stringify({ type, payload });
  wss.clients.forEach((client) => {
    if (client.readyState === 1) client.send(msg);
  });
}

// ── API Routes ──

// Get all departments and agents
app.get('/api/agents', (req, res) => {
  res.json(AGENTS);
});

// Get all tasks
app.get('/api/tasks', (req, res) => {
  res.json(tasks);
});

// Create a new task
app.post('/api/tasks', (req, res) => {
  const { agentId, department, title, description } = req.body;
  if (!agentId || !title) {
    return res.status(400).json({ error: 'agentId and title are required' });
  }
  const task = {
    id: `task-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    agentId,
    department: department || '',
    title,
    description: description || '',
    status: 'pending', // pending | running | completed | failed
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  tasks.push(task);
  saveJSON('tasks.json', tasks);
  broadcast('task:created', task);
  res.status(201).json(task);
});

// Update task status
app.patch('/api/tasks/:id', (req, res) => {
  const task = tasks.find((t) => t.id === req.params.id);
  if (!task) return res.status(404).json({ error: 'Task not found' });
  const { status, result } = req.body;
  if (status) task.status = status;
  if (result !== undefined) task.result = result;
  task.updatedAt = new Date().toISOString();
  saveJSON('tasks.json', tasks);
  broadcast('task:updated', task);
  res.json(task);
});

// Delete a task
app.delete('/api/tasks/:id', (req, res) => {
  const idx = tasks.findIndex((t) => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Task not found' });
  const [removed] = tasks.splice(idx, 1);
  saveJSON('tasks.json', tasks);
  broadcast('task:deleted', removed);
  res.json(removed);
});

// Get logs
app.get('/api/logs', (req, res) => {
  const { agentId, limit } = req.query;
  let filtered = logs;
  if (agentId) filtered = filtered.filter((l) => l.agentId === agentId);
  const n = parseInt(limit) || 100;
  res.json(filtered.slice(-n));
});

// Post a log entry (agents call this to report)
app.post('/api/logs', (req, res) => {
  const { agentId, department, level, message, taskId } = req.body;
  if (!agentId || !message) {
    return res.status(400).json({ error: 'agentId and message are required' });
  }
  const entry = {
    id: `log-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    agentId,
    department: department || '',
    taskId: taskId || null,
    level: level || 'info', // info | warn | error | success
    message,
    timestamp: new Date().toISOString(),
  };
  logs.push(entry);
  // Keep last 1000 logs in memory
  if (logs.length > 1000) logs = logs.slice(-1000);
  saveJSON('logs.json', logs);
  broadcast('log:new', entry);
  res.status(201).json(entry);
});

// Summary stats
app.get('/api/stats', (req, res) => {
  const totalAgents = Object.values(AGENTS).reduce((sum, d) => sum + d.agents.length, 0);
  const tasksByStatus = { pending: 0, running: 0, completed: 0, failed: 0 };
  tasks.forEach((t) => { tasksByStatus[t.status] = (tasksByStatus[t.status] || 0) + 1; });
  const recentLogs = logs.slice(-50);
  res.json({
    departments: Object.keys(AGENTS).length,
    totalAgents,
    tasks: tasksByStatus,
    totalTasks: tasks.length,
    recentLogCount: recentLogs.length,
  });
});

// ── WebSocket ──
wss.on('connection', (ws) => {
  ws.send(JSON.stringify({ type: 'connected', payload: { message: 'VUILD Agent Dashboard connected' } }));
});

// ── Start ──
server.listen(PORT, () => {
  console.log(`VUILD Agent Dashboard running at http://localhost:${PORT}`);
});
