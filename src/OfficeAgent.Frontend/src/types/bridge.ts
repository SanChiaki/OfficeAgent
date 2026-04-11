export interface BridgeRequestEnvelope<TPayload = unknown> {
  type: string;
  requestId: string;
  payload?: TPayload;
}

export interface BridgeErrorPayload {
  code: string;
  message: string;
}

export interface AppSettings {
  apiKey: string;
  baseUrl: string;
  businessBaseUrl: string;
  model: string;
  ssoUrl: string;
  ssoLoginSuccessPath: string;
}

export interface LoginResult {
  success: boolean;
  error?: string;
}

export interface LoginStatus {
  isLoggedIn: boolean;
  ssoUrl: string;
}

export interface ChatMessage {
  id: string;
  role: string;
  content: string;
  createdAtUtc: string;
}

export interface ChatSession {
  id: string;
  title: string;
  createdAtUtc: string;
  updatedAtUtc: string;
  messages: ChatMessage[];
}

export interface SessionState {
  activeSessionId: string;
  sessions: ChatSession[];
}

export interface SelectionContext {
  hasSelection: boolean;
  workbookName: string;
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  isContiguous: boolean;
  headerPreview: string[];
  sampleRows: string[][];
  warningMessage?: string | null;
}

export interface ExcelCommand {
  commandType: string;
  sheetName?: string;
  targetAddress?: string;
  newSheetName?: string;
  values?: string[][];
  confirmed: boolean;
}

export interface ExcelCommandPreview {
  title: string;
  summary: string;
  details: string[];
}

export interface ExcelTableData {
  sheetName: string;
  address: string;
  headers: string[];
  rows: string[][];
}

export interface ExcelCommandResult {
  commandType: string;
  requiresConfirmation: boolean;
  status: string;
  message: string;
  preview?: ExcelCommandPreview;
  table?: ExcelTableData;
  selectionContext?: SelectionContext;
}

export interface UploadPreview {
  projectName: string;
  sheetName: string;
  address: string;
  headers: string[];
  rows: string[][];
  records: Array<Record<string, string>>;
}

export interface SkillRequestEnvelope {
  userInput: string;
  skillName?: string;
  confirmed: boolean;
  uploadPreview?: UploadPreview;
}

export interface SkillResult {
  route: string;
  skillName?: string;
  requiresConfirmation: boolean;
  status: string;
  message: string;
  preview?: ExcelCommandPreview;
  uploadPreview?: UploadPreview;
}

export interface AgentPlanStep {
  type: string;
  args?: Record<string, unknown>;
}

export interface AgentPlan {
  summary: string;
  steps: AgentPlanStep[];
}

export interface PlannerResponse {
  mode: string;
  assistantMessage: string;
  step?: AgentPlanStep;
  plan?: AgentPlan;
}

export interface PlanExecutionJournalStep {
  type: string;
  title: string;
  status: string;
  message?: string;
  errorMessage?: string;
}

export interface PlanExecutionJournal {
  hasFailures: boolean;
  errorMessage: string;
  steps: PlanExecutionJournalStep[];
}

export interface ConversationTurn {
  role: string;
  content: string;
}

export interface AgentRequestEnvelope {
  userInput: string;
  confirmed: boolean;
  sessionId?: string;
  plan?: AgentPlan;
  conversationHistory?: ConversationTurn[];
}

export interface AgentResult {
  route: string;
  requiresConfirmation: boolean;
  status: string;
  message: string;
  planner?: PlannerResponse;
  journal?: PlanExecutionJournal;
}

export interface BridgeResponseEnvelope<TPayload = unknown> {
  type: string;
  requestId: string;
  ok: boolean;
  payload?: TPayload;
  error?: BridgeErrorPayload;
}

export interface PingPayload {
  host: string;
  version: string;
}

export interface BridgeEventEnvelope<TPayload = unknown> {
  type: string;
  payload?: TPayload;
}

export interface WebViewMessageEventLike {
  data: unknown;
}

export interface WebViewHostLike {
  addEventListener: (
    type: 'message',
    listener: (event: WebViewMessageEventLike) => void,
  ) => void;
  removeEventListener: (
    type: 'message',
    listener: (event: WebViewMessageEventLike) => void,
  ) => void;
  postMessage: (message: unknown) => void;
}

declare global {
  interface Window {
    chrome?: {
      webview?: WebViewHostLike;
    };
  }
}
