import { nativeBridge } from '../bridge/nativeBridge';
import type { AnalyticsPayload, SelectionContext, UiLocale } from '../types/bridge';

type CommonPanelAnalyticsContext = {
  sessionId?: string;
  uiLocale: UiLocale;
  selectionContext?: SelectionContext | null;
};

export function trackPanelEvent(
  eventName: string,
  common: CommonPanelAnalyticsContext,
  properties: Record<string, unknown> = {},
  businessContext: Record<string, unknown> = {},
) {
  const payload: AnalyticsPayload = {
    eventName,
    source: 'panel',
    properties: {
      sessionId: common.sessionId ?? '',
      uiLocale: common.uiLocale,
      hasSelection: Boolean(common.selectionContext?.hasSelection),
      sheetName: common.selectionContext?.sheetName ?? '',
      workbookName: common.selectionContext?.workbookName ?? '',
      ...properties,
    },
    businessContext,
  };

  void nativeBridge.trackAnalytics(payload).catch(() => {
    // Analytics must not interrupt panel interactions.
  });
}
