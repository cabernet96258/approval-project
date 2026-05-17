import ArtworkCard from './ArtworkCard';

export default function Gallery({ artworks, loading, userName, onLike, onDelete }) {
  if (loading) {
    return (
      <div className="loading">
        <div className="spinner">🌟</div>
        <p>よみこんでいます...</p>
      </div>
    );
  }

  if (artworks.length === 0) {
    return (
      <div className="empty-state">
        <div className="empty-emoji">🎨</div>
        <h2>まださくひんがないよ</h2>
        <p>さいしょのさくひんをのせてみよう！</p>
      </div>
    );
  }

  const topArtwork = artworks[0];

  return (
    <div>
      {artworks.length > 0 && topArtwork.likes > 0 && (
        <div className="ranking-banner">
          🏆 いちばん「いいね！」が多いのは <strong>{topArtwork.author}さん</strong> の「{topArtwork.title}」！ ({topArtwork.likes}いいね！)
        </div>
      )}
      <div className="gallery-grid">
        {artworks.map((artwork, index) => (
          <ArtworkCard
            key={artwork.id}
            artwork={artwork}
            rank={index + 1}
            userName={userName}
            onLike={onLike}
            onDelete={onDelete}
          />
        ))}
      </div>
    </div>
  );
}
