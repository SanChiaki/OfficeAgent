# Module Index

本文件是功能模块文档的统一入口。

如果你在新会话里继续迭代某个模块，建议先读这里，再跳到对应模块的“当前功能快照”文档。

## 使用方式

1. 先找到目标模块
2. 先读该模块的当前功能快照
3. 再按需要继续读设计说明、实施计划、测试清单、接入指南

## 模块索引

| 模块名 | 当前功能快照 | 相关设计 / 计划 | 相关测试 / 接入文档 |
| --- | --- | --- | --- |
| Ribbon Sync | [docs/modules/ribbon-sync-current-behavior.md](./modules/ribbon-sync-current-behavior.md) | [docs/superpowers/specs/2026-05-08-ai-column-mapping-design.md](./superpowers/specs/2026-05-08-ai-column-mapping-design.md)<br>[docs/superpowers/plans/2026-05-08-ai-column-mapping-implementation-plan.md](./superpowers/plans/2026-05-08-ai-column-mapping-implementation-plan.md)<br>[docs/superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md](./superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md)<br>[docs/superpowers/plans/2026-04-14-office-agent-metadata-layout-implementation-plan.md](./superpowers/plans/2026-04-14-office-agent-metadata-layout-implementation-plan.md)<br>[docs/superpowers/plans/2026-04-14-ribbon-sync-multi-system-project-loading.md](./superpowers/plans/2026-04-14-ribbon-sync-multi-system-project-loading.md) | [docs/vsto-manual-test-checklist.md](./vsto-manual-test-checklist.md)<br>[docs/ribbon-sync-real-system-integration-guide.md](./ribbon-sync-real-system-integration-guide.md)<br>[docs/ribbon-button-custom-icons-guide.md](./ribbon-button-custom-icons-guide.md)<br>[tests/mock-server/README.md](../tests/mock-server/README.md) |
| Task Pane | [docs/modules/task-pane-current-behavior.md](./modules/task-pane-current-behavior.md) | [docs/superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md](./superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md) | [docs/vsto-manual-test-checklist.md](./vsto-manual-test-checklist.md)<br>[src/OfficeAgent.Frontend/src/App.test.tsx](../src/OfficeAgent.Frontend/src/App.test.tsx)<br>[src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts](../src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts) |
| Analytics | [docs/modules/task-pane-current-behavior.md](./modules/task-pane-current-behavior.md)<br>[docs/modules/ribbon-sync-current-behavior.md](./modules/ribbon-sync-current-behavior.md) | [docs/superpowers/specs/2026-05-12-office-agent-analytics-instrumentation-design.md](./superpowers/specs/2026-05-12-office-agent-analytics-instrumentation-design.md)<br>[docs/superpowers/plans/2026-05-12-office-agent-analytics-instrumentation-implementation-plan.md](./superpowers/plans/2026-05-12-office-agent-analytics-instrumentation-implementation-plan.md) | [docs/vsto-manual-test-checklist.md](./vsto-manual-test-checklist.md)<br>[tests/mock-server/README.md](../tests/mock-server/README.md) |

## 维护约定

- 每个功能模块都应在 `docs/modules/` 下维护一份 `*-current-behavior.md`
- 如果模块行为发生变化，应同步更新对应快照文档
- 如果新增模块，应同时更新本索引
- 如果某模块只有设计文档，没有当前功能快照，则不适合作为后续迭代入口
