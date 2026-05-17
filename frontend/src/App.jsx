import { useState, useEffect } from 'react';
import Gallery from './components/Gallery';
import UploadForm from './components/UploadForm';

export default function App() {
  const [artworks, setArtworks] = useState([]);
  const [userName, setUserName] = useState(() => localStorage.getItem('userName') || '');
  const [nameInput, setNameInput] = useState('');
  const [view, setView] = useState('gallery'); // 'gallery' | 'upload'
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchArtworks();
  }, []);

  const fetchArtworks = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/artworks');
      const data = await res.json();
      setArtworks(data);
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  const handleSetName = (e) => {
    e.preventDefault();
    if (nameInput.trim()) {
      const name = nameInput.trim();
      localStorage.setItem('userName', name);
      setUserName(name);
    }
  };

  const handleLike = async (artworkId) => {
    if (!userName) return;
    try {
      const res = await fetch(`/api/artworks/${artworkId}/like`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ liker_name: userName })
      });
      if (res.ok) {
        const updated = await res.json();
        setArtworks(prev => prev.map(a => a.id === artworkId ? updated : a)
          .sort((a, b) => b.likes - a.likes || new Date(b.created_at) - new Date(a.created_at)));
      } else {
        const err = await res.json();
        if (err.error === 'すでにいいね！しています') {
          alert('もう「いいね！」しているよ！');
        }
      }
    } catch (e) {
      console.error(e);
    }
  };

  const handleUpload = async (formData) => {
    const res = await fetch('/api/artworks', { method: 'POST', body: formData });
    if (!res.ok) throw new Error('アップロードに失敗しました');
    const newArtwork = await res.json();
    setArtworks(prev => [newArtwork, ...prev]);
    setView('gallery');
  };

  const handleDelete = async (artworkId) => {
    if (!confirm('本当に削除しますか？')) return;
    const res = await fetch(`/api/artworks/${artworkId}`, { method: 'DELETE' });
    if (res.ok) {
      setArtworks(prev => prev.filter(a => a.id !== artworkId));
    }
  };

  if (!userName) {
    return (
      <div className="name-screen">
        <div className="name-card">
          <div className="name-emoji">🎨</div>
          <h1>みんなのさくひんギャラリー</h1>
          <p>まず、あなたのなまえをおしえてね！</p>
          <form onSubmit={handleSetName}>
            <input
              type="text"
              value={nameInput}
              onChange={e => setNameInput(e.target.value)}
              placeholder="なまえをいれてね"
              maxLength={20}
              autoFocus
            />
            <button type="submit" disabled={!nameInput.trim()}>
              はじめる！
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-left">
          <h1>🎨 みんなのさくひんギャラリー</h1>
        </div>
        <div className="header-right">
          <span className="user-badge">👤 {userName}さん</span>
          <button
            className="change-name-btn"
            onClick={() => { localStorage.removeItem('userName'); setUserName(''); }}
          >
            なまえをかえる
          </button>
        </div>
      </header>

      <nav className="tab-nav">
        <button
          className={view === 'gallery' ? 'tab active' : 'tab'}
          onClick={() => setView('gallery')}
        >
          🖼️ ギャラリー ({artworks.length})
        </button>
        <button
          className={view === 'upload' ? 'tab active' : 'tab'}
          onClick={() => setView('upload')}
        >
          ✨ さくひんをのせる
        </button>
      </nav>

      <main className="main-content">
        {view === 'gallery' ? (
          <Gallery
            artworks={artworks}
            loading={loading}
            userName={userName}
            onLike={handleLike}
            onDelete={handleDelete}
          />
        ) : (
          <UploadForm userName={userName} onUpload={handleUpload} onCancel={() => setView('gallery')} />
        )}
      </main>
    </div>
  );
}
