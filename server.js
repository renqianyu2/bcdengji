const express = require('express');
const Database = require('better-sqlite3');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 9090;

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET,POST,PATCH,DELETE,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
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

// 人脸库API
app.get('/api/faces', (req, res) => {
  const visitors = db.prepare('SELECT id, name, phone, photo_path FROM visitors WHERE photo_path IS NOT NULL ORDER BY id DESC').all();
  res.json(visitors);
});

const db = new Database('database/visitors.db');

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

const storage = multer.diskStorage({
  destination: 'public/uploads/',
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

app.get('/api/visitors', (req, res) => {
  const { name, phone, identity, host, code, startDate, endDate, page = '1', pageSize = '20' } = req.query;
  
  let sql = 'SELECT * FROM visitors WHERE 1=1';
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
    sql += ' AND date(checkin_at) >= ?';
    params.push(startDate);
  }
  if (endDate) {
    sql += ' AND date(checkin_at) <= ?';
    params.push(endDate);
  }
  
  const countSql = sql.replace('SELECT *', 'SELECT COUNT(*) as total');
  const countResult = db.prepare(countSql).get(...params);
  const total = countResult.total;
  
  sql += ' ORDER BY checkin_at DESC LIMIT ? OFFSET ?';
  params.push(parseInt(pageSize), (parseInt(page) - 1) * parseInt(pageSize));
  
  const visitors = db.prepare(sql).all(...params);
  res.json({ data: visitors, total });
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
  const { name, phone, identity, photoPath, purpose, purposeDetail, classInfo, studentInfo, itemPhotoPath, itemDescription, company, fromSchool, host, reason, carInfo } = req.body;
  
  if (!name || !phone || !identity) {
    return res.status(400).json({ error: '请填写所有必填字段' });
  }
  
  const phoneRegex = /^1[3-9]\d{9}$/;
  if (!phoneRegex.test(phone)) {
    return res.status(400).json({ error: '手机号格式不正确' });
  }
  
  const code = generateCode();
  
  const stmt = db.prepare(`
    INSERT INTO visitors (code, name, phone, identity, photo_path, purpose, purpose_detail, class_info, student_info, item_photo_path, item_description, company, from_school, host, reason, car_info)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  
  const result = stmt.run(code, name, phone, identity, photoPath || null, purpose || null, purposeDetail || null, classInfo || null, studentInfo || null, itemPhotoPath || null, itemDescription || null, company || null, fromSchool || null, host || null, reason || null, carInfo || null);
  
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
  
  let sql = 'SELECT * FROM visitors WHERE 1=1';
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
    sql += ' AND date(checkin_at) >= ?';
    params.push(startDate);
  }
  if (endDate) {
    sql += ' AND date(checkin_at) <= ?';
    params.push(endDate);
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
  
  const ws = xlsx.utils.json_to_sheet(data);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, '访客记录');
  
  const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename=visitors.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`服务运行在 http://0.0.0.0:${PORT}`);
});