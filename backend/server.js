const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 3001;

// Data directory: use DATA_DIR env var (set this to a Railway volume mount path)
const dataDir = process.env.DATA_DIR || __dirname;
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const db = new Database(path.join(dataDir, 'artworks.db'));
db.exec(`
  CREATE TABLE IF NOT EXISTS artworks (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    author TEXT NOT NULL,
    image_data TEXT NOT NULL,
    mime_type TEXT NOT NULL DEFAULT 'image/jpeg',
    likes INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE TABLE IF NOT EXISTS likes (
    artwork_id TEXT NOT NULL,
    liker_name TEXT NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (artwork_id, liker_name)
  );
`);

app.use(express.json());

// Serve built frontend
const publicDir = path.join(__dirname, 'public');
if (fs.existsSync(publicDir)) {
  app.use(express.static(publicDir));
}

// Multer: memory storage (images stored as base64 in SQLite)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    cb(null, allowed.includes(file.mimetype));
  }
});

// GET all artworks (metadata only, no image data)
app.get('/api/artworks', (req, res) => {
  const artworks = db.prepare(`
    SELECT id, title, author, mime_type, likes, created_at
    FROM artworks ORDER BY likes DESC, created_at DESC
  `).all();
  res.json(artworks);
});

// GET artwork image
app.get('/api/artworks/:id/image', (req, res) => {
  const row = db.prepare('SELECT image_data, mime_type FROM artworks WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).end();
  const buf = Buffer.from(row.image_data, 'base64');
  res.setHeader('Content-Type', row.mime_type);
  res.setHeader('Cache-Control', 'public, max-age=86400');
  res.send(buf);
});

// POST upload artwork
app.post('/api/artworks', upload.single('image'), (req, res) => {
  const { title, author } = req.body;
  if (!req.file || !title || !author) {
    return res.status(400).json({ error: '必要な情報が足りません' });
  }
  const id = uuidv4();
  const imageData = req.file.buffer.toString('base64');
  db.prepare(`
    INSERT INTO artworks (id, title, author, image_data, mime_type) VALUES (?, ?, ?, ?, ?)
  `).run(id, title.trim(), author.trim(), imageData, req.file.mimetype);

  const artwork = db.prepare(`
    SELECT id, title, author, mime_type, likes, created_at FROM artworks WHERE id = ?
  `).get(id);
  res.status(201).json(artwork);
});

// POST like an artwork
app.post('/api/artworks/:id/like', (req, res) => {
  const { id } = req.params;
  const { liker_name } = req.body;
  if (!liker_name) return res.status(400).json({ error: '名前を教えてください' });

  const artwork = db.prepare('SELECT id FROM artworks WHERE id = ?').get(id);
  if (!artwork) return res.status(404).json({ error: '作品が見つかりません' });

  try {
    db.prepare('INSERT INTO likes (artwork_id, liker_name) VALUES (?, ?)').run(id, liker_name.trim());
    db.prepare('UPDATE artworks SET likes = likes + 1 WHERE id = ?').run(id);
  } catch {
    return res.status(409).json({ error: 'すでにいいね！しています' });
  }

  const updated = db.prepare(`
    SELECT id, title, author, mime_type, likes, created_at FROM artworks WHERE id = ?
  `).get(id);
  res.json(updated);
});

// DELETE artwork
app.delete('/api/artworks/:id', (req, res) => {
  const artwork = db.prepare('SELECT id FROM artworks WHERE id = ?').get(req.params.id);
  if (!artwork) return res.status(404).json({ error: '作品が見つかりません' });

  db.prepare('DELETE FROM likes WHERE artwork_id = ?').run(req.params.id);
  db.prepare('DELETE FROM artworks WHERE id = ?').run(req.params.id);
  res.json({ message: '削除しました' });
});

// SPA fallback
app.get('*', (req, res) => {
  const index = path.join(publicDir, 'index.html');
  if (fs.existsSync(index)) {
    res.sendFile(index);
  } else {
    res.status(404).json({ error: 'Not found' });
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
