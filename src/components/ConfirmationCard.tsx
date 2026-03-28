interface ConfirmationCardProps {
  title: string;
  detail: string;
  onConfirm(): void;
  onCancel(): void;
}

export function ConfirmationCard({ title, detail, onConfirm, onCancel }: ConfirmationCardProps) {
  return (
    <article className="confirmation-card">
      <h2>{title}</h2>
      <p>{detail}</p>
      <div className="confirmation-actions">
        <button type="button" onClick={onCancel}>
          取消
        </button>
        <button type="button" onClick={onConfirm}>
          确认
        </button>
      </div>
    </article>
  );
}
