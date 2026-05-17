export default function ArtworkCard({ artwork, rank, userName, onLike, onDelete }) {
  const isOwner = artwork.author === userName;
  const rankEmoji = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : '';

  return (
    <div className={`artwork-card ${rank <= 3 && artwork.likes > 0 ? 'top-card' : ''}`}>
      {rankEmoji && artwork.likes > 0 && (
        <div className="rank-badge">{rankEmoji}</div>
      )}
      <div className="artwork-image-wrapper">
        <img
          src={`/api/artworks/${artwork.id}/image`}
          alt={artwork.title}
          className="artwork-image"
          loading="lazy"
        />
      </div>
      <div className="artwork-info">
        <h3 className="artwork-title">{artwork.title}</h3>
        <p className="artwork-author">✏️ {artwork.author}さん</p>
        <div className="artwork-actions">
          <button
            className="like-btn"
            onClick={() => onLike(artwork.id)}
            disabled={isOwner}
            title={isOwner ? '自分のさくひんにはいいね！できません' : `いいね！する`}
          >
            ❤️ いいね！{artwork.likes > 0 && <span className="like-count">{artwork.likes}</span>}
          </button>
          {isOwner && (
            <button className="delete-btn" onClick={() => onDelete(artwork.id)}>
              🗑️ けす
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
