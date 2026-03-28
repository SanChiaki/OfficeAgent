interface ConfirmationCardProps {
  summary: string;
  error?: string | null;
  isExecuting?: boolean;
  onConfirm(): void;
  onCancel(): void;
}

export function ConfirmationCard({
  summary,
  error,
  isExecuting = false,
  onConfirm,
  onCancel,
}: ConfirmationCardProps) {
  return (
    <article className="confirmation-card">
      <p>{summary}</p>
      {error ? (
        <p role="alert" className="confirmation-error">
          {error}
        </p>
      ) : null}
      <div className="confirmation-actions">
        <button type="button" onClick={onConfirm} disabled={isExecuting}>
          {"\u786e\u8ba4"}
        </button>
        <button type="button" onClick={onCancel} disabled={isExecuting}>
          {"\u53d6\u6d88"}
        </button>
      </div>
    </article>
  );
}
