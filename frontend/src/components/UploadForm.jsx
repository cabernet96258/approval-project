import { useState, useRef } from 'react';

export default function UploadForm({ userName, onUpload, onCancel }) {
  const [title, setTitle] = useState('');
  const [file, setFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState('');
  const fileInputRef = useRef();

  const handleFileChange = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    if (!f.type.startsWith('image/')) {
      setError('がぞうファイルをえらんでね');
      return;
    }
    setFile(f);
    setError('');
    const reader = new FileReader();
    reader.onload = (ev) => setPreview(ev.target.result);
    reader.readAsDataURL(f);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file || !title.trim()) return;
    setUploading(true);
    setError('');
    try {
      const formData = new FormData();
      formData.append('image', file);
      formData.append('title', title.trim());
      formData.append('author', userName);
      await onUpload(formData);
    } catch (err) {
      setError('アップロードにしっぱいしました。もういちどためしてね。');
      setUploading(false);
    }
  };

  return (
    <div className="upload-form-container">
      <div className="upload-card">
        <h2>✨ さくひんをのせよう！</h2>
        <form onSubmit={handleSubmit}>
          <div className="form-group">
            <label>さくひんのなまえ</label>
            <input
              type="text"
              value={title}
              onChange={e => setTitle(e.target.value)}
              placeholder="れい：はるのはな"
              maxLength={50}
              required
            />
          </div>

          <div className="form-group">
            <label>がぞうをえらぶ</label>
            <div
              className={`drop-zone ${preview ? 'has-preview' : ''}`}
              onClick={() => fileInputRef.current.click()}
            >
              {preview ? (
                <img src={preview} alt="プレビュー" className="preview-image" />
              ) : (
                <div className="drop-zone-placeholder">
                  <span className="drop-icon">📷</span>
                  <span>ここをおして<br/>がぞうをえらぼう</span>
                </div>
              )}
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              onChange={handleFileChange}
              style={{ display: 'none' }}
            />
          </div>

          {error && <p className="error-msg">{error}</p>}

          <div className="form-actions">
            <button type="button" className="cancel-btn" onClick={onCancel}>
              もどる
            </button>
            <button
              type="submit"
              className="submit-btn"
              disabled={!file || !title.trim() || uploading}
            >
              {uploading ? 'アップロードちゅう...' : '🚀 のせる！'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
