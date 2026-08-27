const express = require('express');
const Database = require('better-sqlite3');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
const PORT = 8989;

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET,POST,PUT,PATCH,DELETE,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

app.use('/api', (req, res, next) => {
  res.header('Cache-Control', 'no-store');
  next();
});

app.use(express.json());
app.use(express.static('public'));

app.get('/records.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'records.html'));
});

app.get('/lookup', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'lookup.html'));
});

app.get('/print', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'print.html'));
});

app.get('/config', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'config.html'));
});

// 人脸库API
app.get('/api/faces', (req, res) => {
  const visitors = db.prepare('SELECT id, name, phone, photo_path FROM visitors WHERE photo_path IS NOT NULL ORDER BY id DESC').all();
  res.json(visitors);
});

const dbDir = path.join(__dirname, 'database');
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir, { recursive: true });
}
const db = new Database(path.join(dbDir, 'visitors.db'));

// 启用 WAL 模式，支持读写并发
db.pragma('journal_mode = WAL');
db.pragma('busy_timeout = 5000');

db.exec(`
  CREATE TABLE IF NOT EXISTS visitors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT,
    name TEXT NOT NULL,
    phone TEXT NOT NULL,
    identity TEXT NOT NULL,
    photo_path TEXT,
    purpose TEXT,
    purpose_detail TEXT,
    class_info TEXT,
    student_info TEXT,
    item_photo_path TEXT,
    item_description TEXT,
    company TEXT,
    from_school TEXT,
    host TEXT,
    reason TEXT,
    car_info TEXT,
    checkin_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    checkout_at DATETIME
  )
`);

db.exec(`
  CREATE TABLE IF NOT EXISTS interview_stats (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    person_name TEXT NOT NULL,
    count INTEGER DEFAULT 0
  )
`);

// 如果code列不存在，添加它
try {
  db.exec("ALTER TABLE visitors ADD COLUMN code TEXT");
} catch(e) {
  // 列已存在
}

// 数据库字段迁移
const tableInfo = db.prepare("PRAGMA table_info(visitors)").all();
const columns = tableInfo.map(c => c.name);

const newColumns = [
  { name: 'identity', sql: 'ALTER TABLE visitors ADD COLUMN identity TEXT' },
  { name: 'photo_path', sql: 'ALTER TABLE visitors ADD COLUMN photo_path TEXT' },
  { name: 'purpose', sql: 'ALTER TABLE visitors ADD COLUMN purpose TEXT' },
  { name: 'purpose_detail', sql: 'ALTER TABLE visitors ADD COLUMN purpose_detail TEXT' },
  { name: 'class_info', sql: 'ALTER TABLE visitors ADD COLUMN class_info TEXT' },
  { name: 'student_info', sql: 'ALTER TABLE visitors ADD COLUMN student_info TEXT' },
  { name: 'item_photo_path', sql: 'ALTER TABLE visitors ADD COLUMN item_photo_path TEXT' },
  { name: 'item_description', sql: 'ALTER TABLE visitors ADD COLUMN item_description TEXT' },
  { name: 'company', sql: 'ALTER TABLE visitors ADD COLUMN company TEXT' },
  { name: 'from_school', sql: 'ALTER TABLE visitors ADD COLUMN from_school TEXT' },
  { name: 'host', sql: 'ALTER TABLE visitors ADD COLUMN host TEXT' },
  { name: 'reason', sql: 'ALTER TABLE visitors ADD COLUMN reason TEXT' }
];

newColumns.forEach(col => {
  if (!columns.includes(col.name)) {
    try { db.exec(col.sql); } catch(e) {}
  }
});

// 自定义字段存储
try {
  if (!columns.includes('custom_fields')) {
    db.exec("ALTER TABLE visitors ADD COLUMN custom_fields TEXT");
  }
} catch(e) {}

// ========== 系统配置 / 账号 ==========
db.exec(`
  CREATE TABLE IF NOT EXISTS settings (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    data TEXT NOT NULL,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'parent',
    name TEXT,
    phone TEXT,
    photo_path TEXT,
    grade TEXT,
    class_name TEXT
  )
`);

// 用户表字段迁移（兼容旧库）
try {
  const uCols = db.prepare("PRAGMA table_info(users)").all().map(c => c.name);
  const uNew = [
    { name: 'phone', sql: "ALTER TABLE users ADD COLUMN phone TEXT" },
    { name: 'photo_path', sql: "ALTER TABLE users ADD COLUMN photo_path TEXT" },
    { name: 'grade', sql: "ALTER TABLE users ADD COLUMN grade TEXT" },
    { name: 'class_name', sql: "ALTER TABLE users ADD COLUMN class_name TEXT" }
  ];
  uNew.forEach(c => { if (!uCols.includes(c.name)) { try { db.exec(c.sql); } catch(e) {} } });
} catch(e) {}
db.exec(`
  CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    expires_at DATETIME
  )
`);

// ========== 索引 ==========
db.exec(`
  CREATE INDEX IF NOT EXISTS idx_visitors_code ON visitors(code);
  CREATE INDEX IF NOT EXISTS idx_visitors_name ON visitors(name);
  CREATE INDEX IF NOT EXISTS idx_visitors_phone ON visitors(phone);
  CREATE INDEX IF NOT EXISTS idx_visitors_host ON visitors(host);
  CREATE INDEX IF NOT EXISTS idx_visitors_identity ON visitors(identity);
  CREATE INDEX IF NOT EXISTS idx_visitors_checkin ON visitors(checkin_at);
  CREATE INDEX IF NOT EXISTS idx_sessions_user_id ON sessions(user_id);
`);

const DEFAULT_CONFIG = {
  schoolName: '北辰幼儿园',
  welcomeText: '访客登记',
  contactPhone: '',
  logoPath: '',
  theme: { primary: '#3b82f6', secondary: '#1d4ed8' },
  identityTypes: [
    { value: '家长', label: '学生家长', enabled: true },
    { value: '维修工', label: '维修工人', enabled: true },
    { value: '外来老师', label: '来访人', enabled: true },
    { value: '其他', label: '其他人员', enabled: true }
  ],
  fields: [
    { key: 'name', label: '姓名', type: 'text', enabled: true, required: true },
    { key: 'phone', label: '手机号', type: 'tel', enabled: true, required: true },
    { key: 'photo', label: '人脸图像', type: 'photo', enabled: true, required: true }
  ],
  purposes: {
    '家长': ['接送', '沟通', '家长会', '陪餐', '其他'],
    '维修工': [],
    '外来老师': ['督导检查', '观摩交流', '其他'],
    '其他': ['其他']
  },
  school: {
    type: '小学',
    classesPerGrade: 4,
    extraClasses: []
  },
  customFields: []
};

const CONFIG_FILE = path.join(__dirname, 'config.json');

function loadConfigFromFile() {
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
    }
  } catch(e) {}
  return null;
}

function saveConfigToFile(data) {
  try {
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(data, null, 2), 'utf8');
  } catch(e) {
    console.error('保存配置文件失败:', e.message);
  }
}

if (!db.prepare('SELECT id FROM settings WHERE id = 1').get()) {
  const fileConfig = loadConfigFromFile();
  const configToSave = fileConfig || DEFAULT_CONFIG;
  db.prepare('INSERT INTO settings (id, data) VALUES (1, ?)').run(JSON.stringify(configToSave));
  saveConfigToFile(configToSave);
}

function hashPassword(pw) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(pw, salt, 64).toString('hex');
  return salt + ':' + hash;
}
function verifyPassword(pw, stored) {
  try {
    const parts = stored.split(':');
    if (parts.length !== 2) return false;
    const hash = crypto.scryptSync(pw, parts[0], 64).toString('hex');
    return crypto.timingSafeEqual(Buffer.from(parts[1], 'hex'), Buffer.from(hash, 'hex'));
  } catch(e) { return false; }
}

if (!db.prepare('SELECT id FROM users WHERE username = ?').get('admin')) {
  db.prepare('INSERT INTO users (username, password, role, name) VALUES (?, ?, ?, ?)')
    .run('admin', hashPassword('1234'), 'principal', '校长');
}

function parseCookies(req) {
  const header = req.headers.cookie;
  const out = {};
  if (!header) return out;
  header.split(';').forEach(p => {
    const i = p.indexOf('=');
    if (i > -1) out[p.slice(0, i).trim()] = decodeURIComponent(p.slice(i + 1).trim());
  });
  return out;
}

function getSession(req) {
  const token = parseCookies(req).sid;
  if (!token) return null;
  const row = db.prepare(`
    SELECT u.id, u.username, u.role, u.name, u.phone, u.photo_path, u.grade, u.class_name, s.expires_at
    FROM sessions s JOIN users u ON s.user_id = u.id
    WHERE s.token = ?
  `).get(token);
  if (!row) return null;
  if (row.expires_at && new Date(row.expires_at) < new Date()) {
    db.prepare('DELETE FROM sessions WHERE token = ?').run(token);
    return null;
  }
  delete row.expires_at;
  return row;
}

function requireAuth(role) {
  return (req, res, next) => {
    const user = getSession(req);
    if (!user) return res.status(401).json({ error: '请先登录' });
    if (role && user.role !== role) return res.status(403).json({ error: '没有权限' });
    req.user = user;
    next();
  };
}

function getConfig() {
  const row = db.prepare('SELECT data FROM settings WHERE id = 1').get();
  if (!row) return DEFAULT_CONFIG;
  try { return JSON.parse(row.data); } catch(e) { return DEFAULT_CONFIG; }
}

const uploadDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: uploadDir,
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  }
});
const upload = multer({ storage });

app.post('/upload', upload.single('photo'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: '请选择图片文件' });
  }
  res.json({ photoPath: '/uploads/' + req.file.filename });
});

app.get('/api/interview-stats', (req, res) => {
  const stats = db.prepare('SELECT person_name, count FROM interview_stats ORDER BY count DESC').all();
  res.json(stats);
});

// ========== 认证 ==========
// 实时检测用户名是否已存在
app.get('/api/auth/check-username', (req, res) => {
  const { username } = req.query;
  if (!username || !username.trim()) return res.json({ exists: false });
  const exists = !!db.prepare('SELECT id FROM users WHERE username = ?').get(username.trim());
  res.json({ exists });
});

app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: '请输入账号和密码' });
  const user = db.prepare('SELECT * FROM users WHERE username = ?').get(username);
  if (!user || !verifyPassword(password, user.password)) {
    return res.status(401).json({ error: '账号或密码错误' });
  }
  const token = crypto.randomBytes(32).toString('hex');
  const expires = new Date(Date.now() + 7 * 24 * 3600 * 1000).toISOString();
  db.prepare('INSERT INTO sessions (token, user_id, expires_at) VALUES (?, ?, ?)').run(token, user.id, expires);
  res.setHeader('Set-Cookie', 'sid=' + token + '; HttpOnly; Path=/; Max-Age=604800; SameSite=Lax');
  res.json({ id: user.id, username: user.username, role: user.role, name: user.name, phone: user.phone, photo_path: user.photo_path, grade: user.grade, class_name: user.class_name });
});

// 注册（访客 / 老师 / 校长）
app.post('/api/auth/register', (req, res) => {
  const { username, password, role, name, phone, photoPath, grade, class_name } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: '请填写账号和密码' });
  if (!['parent', 'teacher', 'principal'].includes(role)) return res.status(400).json({ error: '角色无效' });
  if (db.prepare('SELECT id FROM users WHERE username = ?').get(username)) {
    return res.status(400).json({ error: '账号已存在' });
  }
  if (role === 'parent' && (!name || !phone)) {
    return res.status(400).json({ error: '访客需填写姓名和手机号' });
  }
  if (role === 'teacher' && !name) {
    return res.status(400).json({ error: '老师需填写姓名' });
  }
  const result = db.prepare('INSERT INTO users (username, password, role, name, phone, photo_path, grade, class_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?)')
    .run(username, hashPassword(password), role, name || null, phone || null, photoPath || null, grade || null, class_name || null);
  const user = db.prepare('SELECT id, username, role, name, phone, photo_path, grade, class_name FROM users WHERE id = ?').get(result.lastInsertRowid);
  const token = crypto.randomBytes(32).toString('hex');
  const expires = new Date(Date.now() + 7 * 24 * 3600 * 1000).toISOString();
  db.prepare('INSERT INTO sessions (token, user_id, expires_at) VALUES (?, ?, ?)').run(token, user.id, expires);
  res.setHeader('Set-Cookie', 'sid=' + token + '; HttpOnly; Path=/; Max-Age=604800; SameSite=Lax');
  res.json(user);
});

// 教师列表（被访人）
app.get('/api/teachers', (req, res) => {
  const teachers = db.prepare("SELECT id, name, grade, class_name FROM users WHERE role = 'teacher' ORDER BY id").all();
  res.json(teachers);
});

// 班级结构（根据学校类型动态生成年级）
const GRADE_CN = { 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九' };
app.get('/api/school-classes', (req, res) => {
  const cfg = getConfig();
  const school = cfg.school || { type: '小学', classesPerGrade: 4, extraClasses: [] };
  let grades = [];
  if (school.type === '幼儿园') {
    grades = ['小班', '中班', '大班'].concat(school.extraClasses || []);
  } else {
    let start = 1, end = 6;
    if (school.type === '初中+小学') end = 9;
    else if (school.type === '初中') { start = 7; end = 9; }
    for (let g = start; g <= end; g++) grades.push(GRADE_CN[g] + '年级');
  }
  res.json({ type: school.type, classesPerGrade: school.classesPerGrade || 4, grades });
});

// 更新个人资料
app.put('/api/auth/profile', requireAuth(), (req, res) => {
  const { name, phone, photoPath, grade, class_name } = req.body || {};
  db.prepare('UPDATE users SET name = ?, phone = ?, photo_path = ?, grade = ?, class_name = ? WHERE id = ?')
    .run(
      name !== undefined ? name : req.user.name,
      phone !== undefined ? phone : req.user.phone,
      photoPath !== undefined ? photoPath : req.user.photo_path,
      grade !== undefined ? grade : req.user.grade,
      class_name !== undefined ? class_name : req.user.class_name,
      req.user.id
    );
  const user = db.prepare('SELECT id, username, role, name, phone, photo_path, grade, class_name FROM users WHERE id = ?').get(req.user.id);
  res.json(user);
});

app.post('/api/auth/logout', (req, res) => {
  const token = parseCookies(req).sid;
  if (token) db.prepare('DELETE FROM sessions WHERE token = ?').run(token);
  res.setHeader('Set-Cookie', 'sid=; HttpOnly; Path=/; Max-Age=0');
  res.json({ message: '已退出' });
});

app.get('/api/auth/me', (req, res) => {
  const user = getSession(req);
  if (!user) return res.status(401).json({ error: '未登录' });
  res.json(user);
});

// ========== 配置 ==========
app.get('/api/config', (req, res) => {
  res.json(getConfig());
});

app.put('/api/config', requireAuth('principal'), (req, res) => {
  const data = req.body;
  if (!data || typeof data !== 'object') return res.status(400).json({ error: '配置无效' });
  db.prepare('UPDATE settings SET data = ?, updated_at = CURRENT_TIMESTAMP WHERE id = 1')
    .run(JSON.stringify(data));
  saveConfigToFile(data);
  res.json({ message: '配置已保存', config: getConfig() });
});

app.post('/api/config/logo', requireAuth('principal'), upload.single('logo'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择图片文件' });
  res.json({ logoPath: '/uploads/' + req.file.filename });
});

// ========== 账号管理（仅校长） ==========
app.get('/api/users', requireAuth('principal'), (req, res) => {
  const users = db.prepare('SELECT id, username, role, name FROM users ORDER BY id').all();
  res.json(users);
});

app.post('/api/users', requireAuth('principal'), (req, res) => {
  const { username, password, role, name } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: '请填写账号和密码' });
  if (db.prepare('SELECT id FROM users WHERE username = ?').get(username)) {
    return res.status(400).json({ error: '账号已存在' });
  }
  const r = ['parent', 'teacher', 'principal'].includes(role) ? role : 'parent';
  const result = db.prepare('INSERT INTO users (username, password, role, name) VALUES (?, ?, ?, ?)')
    .run(username, hashPassword(password), r, name || null);
  res.json({ id: result.lastInsertRowid });
});

app.delete('/api/users/:id', requireAuth('principal'), (req, res) => {
  const id = parseInt(req.params.id);
  const user = db.prepare('SELECT * FROM users WHERE id = ?').get(id);
  if (!user) return res.status(400).json({ error: '账号不存在' });
  if (user.username === 'admin') return res.status(400).json({ error: '不能删除默认管理员账号' });
  if (req.user.id === id) return res.status(400).json({ error: '不能删除当前登录的账号' });
  db.prepare('DELETE FROM users WHERE id = ?').run(id);
  db.prepare('DELETE FROM sessions WHERE user_id = ?').run(id);
  res.json({ message: '已删除' });
});

app.put('/api/users/password', requireAuth('principal'), (req, res) => {
  const { username, newPassword } = req.body || {};
  if (!username || !newPassword) return res.status(400).json({ error: '请填写账号和新密码' });
  const user = db.prepare('SELECT * FROM users WHERE username = ?').get(username);
  if (!user) return res.status(400).json({ error: '账号不存在' });
  db.prepare('UPDATE users SET password = ? WHERE id = ?').run(hashPassword(newPassword), user.id);
  res.json({ message: '密码已修改' });
});

const VISITOR_FIELDS = 'id, code, name, phone, identity, photo_path, purpose, purpose_detail, class_info, student_info, item_photo_path, item_description, company, from_school, host, reason, car_info, custom_fields, checkin_at, checkout_at';

app.get('/api/visitors', (req, res) => {
  const { name, phone, identity, host, code, startDate, endDate, page = '1', pageSize = '20' } = req.query;
  
  let sql = `SELECT ${VISITOR_FIELDS} FROM visitors WHERE 1=1`;
  const params = [];
  
  if (code) {
    sql += ' AND code = ?';
    params.push(code);
  }
  if (name) {
    sql += ' AND name LIKE ?';
    params.push(`%${name}%`);
  }
  if (phone) {
    sql += ' AND phone LIKE ?';
    params.push(`%${phone}%`);
  }
  if (identity) {
    sql += ' AND identity = ?';
    params.push(identity);
  }
  if (host) {
    sql += ' AND host LIKE ?';
    params.push(`%${host}%`);
  }
  if (startDate) {
    sql += ' AND checkin_at >= ?';
    params.push(startDate);
  }
  if (endDate) {
    sql += ' AND checkin_at <= ?';
    params.push(endDate + ' 23:59:59');
  }
  
  const countSql = sql.replace(VISITOR_FIELDS, 'COUNT(*) as total');
  const countResult = db.prepare(countSql).get(...params);
  const total = countResult.total;
  
  sql += ' ORDER BY checkin_at DESC LIMIT ? OFFSET ?';
  params.push(parseInt(pageSize), (parseInt(page) - 1) * parseInt(pageSize));
  
  const visitors = db.prepare(sql).all(...params);
  res.json({ data: visitors, total });
});

// 按姓名搜索最近一条记录（用于自动填充）
app.get('/api/visitors/search-by-name', (req, res) => {
  const { name } = req.query;
  if (!name || !name.trim()) return res.json(null);
  const visitor = db.prepare('SELECT name, phone, photo_path, identity FROM visitors WHERE name = ? ORDER BY checkin_at DESC LIMIT 1').get(name.trim());
  res.json(visitor || null);
});

function generateCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code = '';
  for (var i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

app.post('/api/visitors', (req, res) => {
  const { name, phone, identity, photoPath, purpose, purposeDetail, classInfo, studentInfo, itemPhotoPath, itemDescription, company, fromSchool, host, reason, carInfo, customFields } = req.body;
  
  if (!name || !phone || !identity) {
    return res.status(400).json({ error: '请填写所有必填字段' });
  }
  
  const phoneRegex = /^1[3-9]\d{9}$/;
  if (!phoneRegex.test(phone)) {
    return res.status(400).json({ error: '手机号格式不正确' });
  }
  
  const code = generateCode();
  
  const stmt = db.prepare(`
    INSERT INTO visitors (code, name, phone, identity, photo_path, purpose, purpose_detail, class_info, student_info, item_photo_path, item_description, company, from_school, host, reason, car_info, custom_fields)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  
  const result = stmt.run(code, name, phone, identity, photoPath || null, purpose || null, purposeDetail || null, classInfo || null, studentInfo || null, itemPhotoPath || null, itemDescription || null, company || null, fromSchool || null, host || null, reason || null, carInfo || null, customFields ? JSON.stringify(customFields) : null);
  
  if (purpose === '采访' && purposeDetail) {
    const existing = db.prepare('SELECT * FROM interview_stats WHERE person_name = ?').get(purposeDetail);
    if (existing) {
      db.prepare('UPDATE interview_stats SET count = count + 1 WHERE person_name = ?').run(purposeDetail);
    } else {
      db.prepare('INSERT INTO interview_stats (person_name, count) VALUES (?, 1)').run(purposeDetail);
    }
  }
  
  res.json({ id: result.lastInsertRowid, code: code, message: '登记成功' });
});

app.patch('/api/visitors/:id', (req, res) => {
  const { id } = req.params;
  
  const stmt = db.prepare(`
    UPDATE visitors SET checkout_at = CURRENT_TIMESTAMP WHERE id = ? AND checkout_at IS NULL
  `);
  
  const result = stmt.run(id);
  if (result.changes === 0) {
    return res.status(400).json({ error: '签离失败或已签离' });
  }
  res.json({ message: '签离成功' });
});

app.delete('/api/visitors/:id', (req, res) => {
  const { id } = req.params;
  
  const stmt = db.prepare('DELETE FROM visitors WHERE id = ?');
  const result = stmt.run(id);
  
  if (result.changes === 0) {
    return res.status(400).json({ error: '删除失败' });
  }
  res.json({ message: '删除成功' });
});

app.get('/api/export', (req, res) => {
  const { name, host, startDate, endDate } = req.query;
  
  let sql = `SELECT ${VISITOR_FIELDS} FROM visitors WHERE 1=1`;
  const params = [];
  
  if (name) {
    sql += ' AND name LIKE ?';
    params.push(`%${name}%`);
  }
  if (host) {
    sql += ' AND host LIKE ?';
    params.push(`%${host}%`);
  }
  if (startDate) {
    sql += ' AND checkin_at >= ?';
    params.push(startDate);
  }
  if (endDate) {
    sql += ' AND checkin_at <= ?';
    params.push(endDate + ' 23:59:59');
  }
  
  sql += ' ORDER BY checkin_at DESC';
  const visitors = db.prepare(sql).all(...params);
  
  const data = visitors.map(v => {
    var classInfo = v.class_info || '';
    var studentInfo = v.student_info || '';
    var targetInfo = '';
    if (v.identity === '家长') {
      targetInfo = classInfo + (studentInfo ? ' ' + studentInfo : '');
    } else if (v.identity === '维修工') {
      targetInfo = v.company || '';
    } else if (v.identity === '外来老师') {
      targetInfo = v.from_school || '';
    } else {
      targetInfo = v.host || '';
    }
    
    var duration = '-';
    if (v.checkout_at) {
      var seconds = (new Date(v.checkout_at) - new Date(v.checkin_at)) / 1000;
      if (seconds < 60) duration = Math.round(seconds) + '秒';
      else if (seconds < 3600) duration = Math.floor(seconds / 60) + '分钟';
      else duration = Math.floor(seconds / 3600) + '小时' + Math.floor((seconds % 3600) / 60) + '分钟';
    }
    
    var purpose = v.purpose || '';
    if (v.purpose_detail) purpose += '：' + v.purpose_detail;
    
    return {
      访客码: v.code || '',
      姓名: v.name,
      手机号: v.phone,
      身份: v.identity || '访客',
      访问对象: targetInfo,
      来校事由: purpose,
      物品描述: v.item_description || '',
      签到时间: v.checkin_at,
      签离时间: v.checkout_at || '',
      停留时长: duration,
      状态: v.checkout_at ? '已签离' : '在访中'
    };
  });
  
  const headers = Object.keys(data[0] || {});
  const csvRows = [headers.join(',')];
  data.forEach(row => {
    csvRows.push(headers.map(h => '"' + String(row[h] || '').replace(/"/g, '""') + '"').join(','));
  });
  const csv = csvRows.join('\n');
  const bom = '\uFEFF';
  res.setHeader('Content-Disposition', 'attachment; filename=visitors.csv');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.send(bom + csv);
});

process.on('uncaughtException', (err) => {
  console.error('未捕获异常:', err);
});
process.on('unhandledRejection', (err) => {
  console.error('未处理Promise:', err);
});

function gracefulShutdown() {
  console.log('正在关闭服务...');
  try {
    db.pragma('wal_checkpoint(TRUNCATE)');
    db.close();
    console.log('数据库已安全关闭');
  } catch(e) {
    console.error('关闭数据库失败:', e.message);
  }
  process.exit(0);
}

process.on('SIGTERM', gracefulShutdown);
process.on('SIGINT', gracefulShutdown);

app.listen(PORT, '0.0.0.0', () => {
  console.log(`服务运行在 http://0.0.0.0:${PORT}`);
});