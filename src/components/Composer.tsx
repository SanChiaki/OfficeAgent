interface ComposerProps {
  value: string;
  onChange(value: string): void;
  onSubmit(): void;
}

export function Composer({ value, onChange, onSubmit }: ComposerProps) {
  return (
    <div className="composer">
      <textarea
        aria-label="消息输入框"
        placeholder="输入你的问题或命令"
        value={value}
        onChange={(event) => onChange(event.target.value)}
      />
      <button type="button" className="composer-send" onClick={onSubmit}>
        发送
      </button>
    </div>
  );
}
