export type ChatRole = "user" | "assistant" | "system";

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
}

export interface ChatSession {
  id: string;
  title: string;
  messages: ChatMessage[];
}

export interface SelectionContext {
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}
