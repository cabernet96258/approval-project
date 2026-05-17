const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = 3001;

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

// Initialize SQLite database
const db = new Database(path.join(__dirname, 'artworks.db'));
db.exec(`
  CREATE TABLE IF NOT EXISTS artworks (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    author TEXT NOT NULL,
    filename TEXT NOT NULL,
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

app.use(cors());
app.use(express.json());
app.use('/uploads', express.static(uploadsDir));

// Multer setup
const storage = multer.diskStorage({
  destination: uploadsDir,
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, uuidv4() + ext);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
  fileFilter: (req, file, cb) => {
    const allowed = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    cb(null, allowed.includes(file.mimetype));
  }
});

// GET all artworks
app.get('/api/artworks', (req, res) => {
  const artworks = db.prepare(`
    SELECT * FROM artworks ORDER BY likes DESC, created_at DESC
  `).all();
  res.json(artworks);
});

// POST upload artwork
app.post('/api/artworks', upload.single('image'), (req, res) => {
  const { title, author } = req.body;
  if (!req.file || !title || !author) {
    return res.status(400).json({ error: '必要な情報が足りません' });
  }
  const id = uuidv4();
  db.prepare(`
    INSERT INTO artworks (id, title, author, filename) VALUES (?, ?, ?, ?)
  `).run(id, title.trim(), author.trim(), req.file.filename);

  const artwork = db.prepare('SELECT * FROM artworks WHERE id = ?').get(id);
  res.status(201).json(artwork);
});

// POST like an artwork
app.post('/api/artworks/:id/like', (req, res) => {
  const { id } = req.params;
  const { liker_name } = req.body;
  if (!liker_name) return res.status(400).json({ error: '名前を教えてください' });

  const artwork = db.prepare('SELECT * FROM artworks WHERE id = ?').get(id);
  if (!artwork) return res.status(404).json({ error: '作品が見つかりません' });

  try {
    db.prepare('INSERT INTO likes (artwork_id, liker_name) VALUES (?, ?)').run(id, liker_name.trim());
    db.prepare('UPDATE artworks SET likes = likes + 1 WHERE id = ?').run(id);
  } catch (e) {
    return res.status(409).json({ error: 'すでにいいね！しています' });
  }

  const updated = db.prepare('SELECT * FROM artworks WHERE id = ?').get(id);
  res.json(updated);
});

// DELETE artwork
app.delete('/api/artworks/:id', (req, res) => {
  const { id } = req.params;
  const artwork = db.prepare('SELECT * FROM artworks WHERE id = ?').get(id);
  if (!artwork) return res.status(404).json({ error: '作品が見つかりません' });

  const filePath = path.join(uploadsDir, artwork.filename);
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);

  db.prepare('DELETE FROM likes WHERE artwork_id = ?').run(id);
  db.prepare('DELETE FROM artworks WHERE id = ?').run(id);
  res.json({ message: '削除しました' });
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
