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
  apiKeyFieldLabel: string;
  baseUrlFieldLabel: string;
  businessBaseUrlFieldLabel: string;
  apiFormatFieldLabel: string;
  apiFormatOpenAiCompatible: string;
  apiFormatAnthropicMessages: string;
  modelFieldLabel: string;
  ssoUrlFieldLabel: string;
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
  excelAddWorksheetPreviewTitle: string;
  excelRenameWorksheetPreviewTitle: string;
  excelDeleteWorksheetPreviewTitle: string;
  excelWriteRangePreviewTitle: string;
  formatExcelAddWorksheetPreviewSummary: (sheetName: string) => string;
  formatExcelRenameWorksheetPreviewSummary: (sheetName: string, newSheetName: string) => string;
  formatExcelDeleteWorksheetPreviewSummary: (sheetName: string) => string;
  formatExcelWriteRangePreviewSummary: (rowCount: number, columnCount: number, targetAddress: string) => string;
  formatWorkbookDetail: (workbookName: string) => string;
  uploadPreviewTitle: string;
  formatUploadPreviewSummary: (rowCount: number, projectName: string) => string;
  formatUploadPreviewSourceDetail: (sheetName: string, address: string) => string;
  formatUploadPreviewFieldsDetail: (headers: string[]) => string;
  formatJournalStatus: (status: string) => string;
  cancellationExcel: string;
  cancellationSkill: string;
  cancellationPlan: string;
  requestFailed: (message: string) => string;
  excelRequestFallback: string;
  skillRequestFallback: string;
  agentRequestFallback: string;
  browserPreviewExcelConfirmMessage: string;
  browserPreviewChatFallback: string;
  browserPreviewLoginUnavailable: string;
  browserPreviewPlanPreparedMessage: string;
  browserPreviewPlanExecutedMessage: string;
  browserPreviewPlanSummary: string;
  browserPreviewReadSelectionMessage: (sheetName: string, address: string) => string;
  browserPreviewWorksheetCreatedMessage: (sheetName: string) => string;
  browserPreviewWorksheetRenamedMessage: (sheetName: string, newSheetName: string) => string;
  browserPreviewWorksheetDeletedMessage: (sheetName: string) => string;
  browserPreviewWriteRangeCompletedMessage: (rowCount: number, targetAddress: string) => string;
  browserPreviewUploadReviewMessage: (projectName: string) => string;
  browserPreviewUploadCompletedMessage: (projectName: string, rowCount: number) => string;
  browserPreviewUnsupportedCommandMessage: (commandType: string) => string;
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
    appHeadingFallback: 'xISDP AI',
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
    apiKeyFieldLabel: 'API 密钥',
    baseUrlFieldLabel: '基础 URL',
    businessBaseUrlFieldLabel: '业务基础 URL',
    apiFormatFieldLabel: 'API 格式',
    apiFormatOpenAiCompatible: 'OpenAI 兼容',
    apiFormatAnthropicMessages: 'Anthropic Messages',
    modelFieldLabel: '模型',
    ssoUrlFieldLabel: 'SSO 地址',
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
    excelAddWorksheetPreviewTitle: '新增工作表',
    excelRenameWorksheetPreviewTitle: '重命名工作表',
    excelDeleteWorksheetPreviewTitle: '删除工作表',
    excelWriteRangePreviewTitle: '写入范围',
    formatExcelAddWorksheetPreviewSummary: (sheetName) => `新增工作表“${sheetName}”`,
    formatExcelRenameWorksheetPreviewSummary: (sheetName, newSheetName) => `将工作表“${sheetName}”重命名为“${newSheetName}”`,
    formatExcelDeleteWorksheetPreviewSummary: (sheetName) => `删除工作表“${sheetName}”`,
    formatExcelWriteRangePreviewSummary: (rowCount, columnCount, targetAddress) => `向 ${targetAddress} 写入 ${rowCount} 行 ${columnCount} 列数据`,
    formatWorkbookDetail: (workbookName) => `工作簿：${workbookName}`,
    uploadPreviewTitle: '上传所选数据',
    formatUploadPreviewSummary: (rowCount, projectName) => `上传 ${rowCount} 行数据到 ${projectName}`,
    formatUploadPreviewSourceDetail: (sheetName, address) => `来源：${sheetName}!${address}`,
    formatUploadPreviewFieldsDetail: (headers) => `字段：${headers.join(', ')}`,
    formatJournalStatus: (status) => {
      switch (status.trim().toLowerCase()) {
        case 'completed':
          return '已完成';
        case 'failed':
          return '已失败';
        case 'preview':
          return '待确认';
        case 'running':
          return '进行中';
        case 'pending':
          return '待处理';
        default:
          return status;
      }
    },
    cancellationExcel: '已取消待处理的 Excel 操作。',
    cancellationSkill: '已取消待处理的上传操作。',
    cancellationPlan: '已取消待执行的计划。',
    requestFailed: (message) => `请求失败：${message}`,
    excelRequestFallback: 'Excel 命令执行失败。',
    skillRequestFallback: 'Skill 执行失败。',
    agentRequestFallback: 'Agent 执行失败。',
    browserPreviewExcelConfirmMessage: '确认此 Excel 操作后再修改工作簿。',
    browserPreviewChatFallback: '暂未实现通用对话路由，请使用 /upload_data ... 或直接的 Excel 命令。',
    browserPreviewLoginUnavailable: 'SSO 登录仅在 Excel 任务窗格内可用。',
    browserPreviewPlanPreparedMessage: '我已经准备好执行计划，请确认后再修改 Excel。',
    browserPreviewPlanExecutedMessage: '计划执行成功。',
    browserPreviewPlanSummary: '创建 Summary 工作表并写入当前选中数据。',
    browserPreviewReadSelectionMessage: (sheetName, address) => `已读取 ${sheetName} ${address} 的选区。`,
    browserPreviewWorksheetCreatedMessage: (sheetName) => `已创建工作表“${sheetName}”。`,
    browserPreviewWorksheetRenamedMessage: (sheetName, newSheetName) => `已将工作表“${sheetName}”重命名为“${newSheetName}”。`,
    browserPreviewWorksheetDeletedMessage: (sheetName) => `已删除工作表“${sheetName}”。`,
    browserPreviewWriteRangeCompletedMessage: (rowCount, targetAddress) => `已向 ${targetAddress} 写入 ${rowCount} 行数据。`,
    browserPreviewUploadReviewMessage: (projectName) => `请先确认发往${projectName}的上传内容。`,
    browserPreviewUploadCompletedMessage: (projectName, rowCount) => `${projectName} 的预览上传已完成（${rowCount} 行）。`,
    browserPreviewUnsupportedCommandMessage: (commandType) => `浏览器预览暂不支持 ${commandType}。`,
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
    appHeadingFallback: 'xISDP AI',
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
    apiKeyFieldLabel: 'API Key',
    baseUrlFieldLabel: 'Base URL',
    businessBaseUrlFieldLabel: 'Business Base URL',
    apiFormatFieldLabel: 'API Format',
    apiFormatOpenAiCompatible: 'OpenAI Compatible',
    apiFormatAnthropicMessages: 'Anthropic Messages',
    modelFieldLabel: 'Model',
    ssoUrlFieldLabel: 'SSO URL',
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
    excelAddWorksheetPreviewTitle: 'Add worksheet',
    excelRenameWorksheetPreviewTitle: 'Rename worksheet',
    excelDeleteWorksheetPreviewTitle: 'Delete worksheet',
    excelWriteRangePreviewTitle: 'Write range',
    formatExcelAddWorksheetPreviewSummary: (sheetName) => `Add worksheet "${sheetName}"`,
    formatExcelRenameWorksheetPreviewSummary: (sheetName, newSheetName) => `Rename worksheet "${sheetName}" to "${newSheetName}"`,
    formatExcelDeleteWorksheetPreviewSummary: (sheetName) => `Delete worksheet "${sheetName}"`,
    formatExcelWriteRangePreviewSummary: (rowCount, columnCount, targetAddress) => `Write ${rowCount} row(s) x ${columnCount} column(s) to ${targetAddress}`,
    formatWorkbookDetail: (workbookName) => `Workbook: ${workbookName}`,
    uploadPreviewTitle: 'Upload selected data',
    formatUploadPreviewSummary: (rowCount, projectName) => `Upload ${rowCount} row(s) to ${projectName}`,
    formatUploadPreviewSourceDetail: (sheetName, address) => `Source: ${sheetName}!${address}`,
    formatUploadPreviewFieldsDetail: (headers) => `Fields: ${headers.join(', ')}`,
    formatJournalStatus: (status) => status,
    cancellationExcel: 'Cancelled the pending Excel action.',
    cancellationSkill: 'Cancelled the pending upload.',
    cancellationPlan: 'Cancelled the pending plan.',
    requestFailed: (message) => `Request failed: ${message}`,
    excelRequestFallback: 'Excel command execution failed.',
    skillRequestFallback: 'Skill execution failed.',
    agentRequestFallback: 'Agent execution failed.',
    browserPreviewExcelConfirmMessage: 'Confirm this Excel action before the workbook is modified.',
    browserPreviewChatFallback: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
    browserPreviewLoginUnavailable: 'SSO login is only available inside the Excel task pane.',
    browserPreviewPlanPreparedMessage: 'I prepared a plan. Review it before Excel is changed.',
    browserPreviewPlanExecutedMessage: 'Plan executed successfully.',
    browserPreviewPlanSummary: 'Create a Summary sheet and write the selected rows.',
    browserPreviewReadSelectionMessage: (sheetName, address) => `Read selection from ${sheetName} ${address}.`,
    browserPreviewWorksheetCreatedMessage: (sheetName) => `Worksheet "${sheetName}" created.`,
    browserPreviewWorksheetRenamedMessage: (sheetName, newSheetName) => `Worksheet "${sheetName}" renamed to "${newSheetName}".`,
    browserPreviewWorksheetDeletedMessage: (sheetName) => `Worksheet "${sheetName}" deleted.`,
    browserPreviewWriteRangeCompletedMessage: (rowCount, targetAddress) => `Wrote ${rowCount} row(s) to ${targetAddress}.`,
    browserPreviewUploadReviewMessage: (projectName) => `Review the upload payload before sending it to ${projectName}.`,
    browserPreviewUploadCompletedMessage: (projectName, rowCount) => `Preview-only upload completed for ${projectName} (${rowCount} row(s)).`,
    browserPreviewUnsupportedCommandMessage: (commandType) => `Browser preview does not support ${commandType}.`,
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
