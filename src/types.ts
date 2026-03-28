export type ChatRole = "user" | "assistant" | "system";

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
}

export interface SelectionContext {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}
