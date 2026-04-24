import type { UiLocale } from '../types/bridge';

export type UiStrings = {
  chatHeaderLabel: string;
  bridgeConnecting: string;
  bridgeConnected: (host: string, version: string) => string;
  bridgeUnavailable: (message: string) => string;
  settingsLoadFailed: string;
  settingsSaveFailed: string;
  loginFailed: string;
  appHeadingFallback: string;
  untitledSessionTitle: string;
  welcomeMessage: string;
  messageThreadLabel: string;
  loadingThinking: string;
  messageComposerLabel: string;
  messagePlaceholder: string;
  selectionCapsuleLabel: string;
  noSelection: string;
  send: string;
  openSettings: string;
  close: string;
  openSessionsDrawer: string;
  closeSessionsDrawer: string;
  sessionsDrawerLabel: string;
  sessionsTitle: string;
  newSession: string;
  noSessions: string;
  renameSession: string;
  confirmRename: string;
  cancelRename: string;
  deleteSession: string;
  settingsDialogLabel: string;
  settingsEyebrow: string;
  settingsTitle: string;
  showApiKey: string;
  hideApiKey: string;
  loginSuccessPath: string;
  loggedIn: string;
  loggedOut: string;
  loginInProgress: string;
  login: string;
  logout: string;
  cancel: string;
  save: string;
  deleteSessionDialogTitle: string;
  deleteSessionPrompt: (title: string) => string;
  confirmCardLabel: string;
  confirmCardEyebrow: string;
  confirmCardTitle: string;
  confirm: string;
  cancellationExcel: string;
  cancellationSkill: string;
  cancellationPlan: string;
  requestFailed: (message: string) => string;
  excelRequestFallback: string;
  skillRequestFallback: string;
  agentRequestFallback: string;
  planPreviewTitle: string;
  formatPlanStepAddWorksheet: (sheetName: string) => string;
  formatPlanStepWriteRange: (targetAddress: string) => string;
  formatPlanStepRenameWorksheet: (sheetName: string, newSheetName: string) => string;
  formatPlanStepDeleteWorksheet: (sheetName: string) => string;
  formatPlanStepUploadData: string;
};

export const UNTITLED_SESSION_STORAGE_TITLE = 'New chat';
export const LEGACY_SYSTEM_UNTITLED_SESSION_TITLES = ['New chat', 'Untitled'] as const;

export const uiStrings: Record<UiLocale, UiStrings> = {
  zh: {
    chatHeaderLabel: '聊天页眉',
    bridgeConnecting: '正在连接宿主...',
    bridgeConnected: (host, version) => `已连接 ${host} (${version})`,
    bridgeUnavailable: (message) => `宿主不可用: ${message}`,
    settingsLoadFailed: '无法从宿主加载设置。',
    settingsSaveFailed: '保存设置失败。',
    loginFailed: '登录失败。',
    appHeadingFallback: 'ISDP AI',
    untitledSessionTitle: '未命名会话',
    welcomeMessage: '欢迎使用ISDP，我是能和Excel交互的Agent。你选中的单元格会被我优先识别，尽情尝试吧~',
    messageThreadLabel: '消息线程',
    loadingThinking: '正在思考…',
    messageComposerLabel: '消息输入框',
    messagePlaceholder: '输入消息...',
    selectionCapsuleLabel: '选区胶囊',
    noSelection: '未选中',
    send: '发送',
    openSettings: '打开设置',
    close: '关闭',
    openSessionsDrawer: '打开会话列表',
    closeSessionsDrawer: '关闭会话列表',
    sessionsDrawerLabel: '会话抽屉',
    sessionsTitle: '会话',
    newSession: '新建会话',
    noSessions: '暂无会话',
    renameSession: '重命名会话',
    confirmRename: '确认重命名',
    cancelRename: '取消重命名',
    deleteSession: '删除会话',
    settingsDialogLabel: '设置对话框',
    settingsEyebrow: '配置',
    settingsTitle: '设置',
    showApiKey: '显示 API Key',
    hideApiKey: '隐藏 API Key',
    loginSuccessPath: '登录成功路径',
    loggedIn: '已登录',
    loggedOut: '未登录',
    loginInProgress: '登录中...',
    login: '登录',
    logout: '登出',
    cancel: '取消',
    save: '保存',
    deleteSessionDialogTitle: '删除会话',
    deleteSessionPrompt: (title) => `确定要删除「${title}」吗？此操作不可撤销。`,
    confirmCardLabel: '确认 Excel 操作',
    confirmCardEyebrow: '待确认的写入操作',
    confirmCardTitle: '确认 Excel 操作',
    confirm: '确认',
    cancellationExcel: '已取消待处理的 Excel 操作。',
    cancellationSkill: '已取消待处理的上传操作。',
    cancellationPlan: '已取消待执行的计划。',
    requestFailed: (message) => `请求失败：${message}`,
    excelRequestFallback: 'Excel 命令执行失败。',
    skillRequestFallback: 'Skill 执行失败。',
    agentRequestFallback: 'Agent 执行失败。',
    planPreviewTitle: '执行计划',
    formatPlanStepAddWorksheet: (sheetName) => `新增工作表 ${sheetName}`.trim(),
    formatPlanStepWriteRange: (targetAddress) => `写入范围 ${targetAddress}`.trim(),
    formatPlanStepRenameWorksheet: (sheetName, newSheetName) => `重命名工作表 ${sheetName} 为 ${newSheetName}`.trim(),
    formatPlanStepDeleteWorksheet: (sheetName) => `删除工作表 ${sheetName}`.trim(),
    formatPlanStepUploadData: '上传所选数据',
  },
  en: {
    chatHeaderLabel: 'Chat header',
    bridgeConnecting: 'Connecting to host...',
    bridgeConnected: (host, version) => `Connected to ${host} (${version})`,
    bridgeUnavailable: (message) => `Host unavailable: ${message}`,
    settingsLoadFailed: 'Unable to load settings from the host.',
    settingsSaveFailed: 'Failed to save settings.',
    loginFailed: 'Login failed.',
    appHeadingFallback: 'ISDP AI',
    untitledSessionTitle: 'Untitled',
    welcomeMessage: 'Welcome to ISDP. I am an agent that can work with Excel. I will prioritize your current selection, so try anything you need.',
    messageThreadLabel: 'Message thread',
    loadingThinking: 'Thinking…',
    messageComposerLabel: 'Message composer',
    messagePlaceholder: 'Type a message...',
    selectionCapsuleLabel: 'Selection capsule',
    noSelection: 'No selection',
    send: 'Send',
    openSettings: 'Open settings',
    close: 'Close',
    openSessionsDrawer: 'Open sessions drawer',
    closeSessionsDrawer: 'Close sessions drawer',
    sessionsDrawerLabel: 'Sessions drawer',
    sessionsTitle: 'Sessions',
    newSession: 'New session',
    noSessions: 'No sessions yet',
    renameSession: 'Rename session',
    confirmRename: 'Confirm rename',
    cancelRename: 'Cancel rename',
    deleteSession: 'Delete session',
    settingsDialogLabel: 'Settings dialog',
    settingsEyebrow: 'Configuration',
    settingsTitle: 'Settings',
    showApiKey: 'Show API Key',
    hideApiKey: 'Hide API Key',
    loginSuccessPath: 'Login success path',
    loggedIn: 'Logged in',
    loggedOut: 'Logged out',
    loginInProgress: 'Logging in...',
    login: 'Log in',
    logout: 'Log out',
    cancel: 'Cancel',
    save: 'Save',
    deleteSessionDialogTitle: 'Delete session',
    deleteSessionPrompt: (title) => `Delete "${title}"? This action cannot be undone.`,
    confirmCardLabel: 'Confirm Excel action',
    confirmCardEyebrow: 'Pending workbook change',
    confirmCardTitle: 'Confirm Excel action',
    confirm: 'Confirm',
    cancellationExcel: 'Cancelled the pending Excel action.',
    cancellationSkill: 'Cancelled the pending upload.',
    cancellationPlan: 'Cancelled the pending plan.',
    requestFailed: (message) => `Request failed: ${message}`,
    excelRequestFallback: 'Excel command execution failed.',
    skillRequestFallback: 'Skill execution failed.',
    agentRequestFallback: 'Agent execution failed.',
    planPreviewTitle: 'Execution plan',
    formatPlanStepAddWorksheet: (sheetName) => `Add worksheet ${sheetName}`.trim(),
    formatPlanStepWriteRange: (targetAddress) => `Write range ${targetAddress}`.trim(),
    formatPlanStepRenameWorksheet: (sheetName, newSheetName) => `Rename worksheet ${sheetName} to ${newSheetName}`.trim(),
    formatPlanStepDeleteWorksheet: (sheetName) => `Delete worksheet ${sheetName}`.trim(),
    formatPlanStepUploadData: 'Upload selected data',
  },
};

export function getUiStrings(locale: UiLocale): UiStrings {
  return uiStrings[locale];
}

export function isLegacySystemUntitledSessionTitle(title: string) {
  return LEGACY_SYSTEM_UNTITLED_SESSION_TITLES.includes(title as (typeof LEGACY_SYSTEM_UNTITLED_SESSION_TITLES)[number]);
}
