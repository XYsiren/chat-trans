# 公式评测与指标采集说明

## 公式回归测试

评测集位于 `tests/formula-regression.test.js`，当前覆盖 12 类常见输入：块/行内公式、`\[...\]`、独占 `$...$`、金额干扰项、`math` 围栏、普通代码围栏、未闭合公式、旧式 LaTeX 命令、集合/量词等。

执行：

```bash
node tests/formula-regression.test.js
node tests/metrics.test.js
```

新增解析规则时，必须先补充一条“此前会失败”的样例，再修改解析器；测试断言包括块类型、公式 LaTeX 原文和闭合状态。

建议扩展到 100–200 条人工标注样本，并为每个样本补充 `expectedFormulaCount`、`expectedBlockTypes`、`expectedLatex`。公式块识别准确率可定义为：`类型与边界均正确的样本数 / 全部样本数`；若将每个公式块单独标注，则可同时计算 Precision、Recall 与 F1。

## 事件与隐私

事件由 `miniprogram/services/metrics-service.js` 记录，云函数为 `metrics`。记录的仅是草稿 ID、匿名用户 ID（仅云端）、尝试 ID、建议 ID、块数和公式数；**不会上传正文、公式原文、文件内容或用户个人信息**。

首次启用前：

1. 在云开发控制台创建集合 `metricsEvents`；
2. 对 `cloudfunctions/metrics` 执行“上传并部署：云端安装依赖”；
3. 真机完成一次“识别 → 修复 → 导出 → 打开文件”流程，确认集合中出现事件。

核心事件：

| 事件 | 触发时机 |
|---|---|
| `document_processing_started` | 用户成功完成一次结构识别 |
| `repair_suggestion_shown` | 一条可一键修复建议首次展示 |
| `repair_suggestion_accepted` | 用户采用修复 |
| `repair_suggestion_dismissed` | 用户选择暂不提示 |
| `repair_undone` | 用户撤销刚才的修复 |
| `word_export_succeeded` / `word_export_failed` | Word 导出结果 |
| `document_delivery_completed` | Word 成功打开 |

## 三项指标口径

指标计算函数位于 `miniprogram/utils/metrics.js`。

- 修复建议采纳率 = 去重后的“至少被采用一次的建议数” / “首次展示的建议数”。
- 误修复率 = 去重后的“被撤销或被标记有问题的修复数” / “已采用的修复数”。当前产品已通过“撤销上次修复”采集第一类误修复信号；后续可补充“修复结果不正确”反馈以增强覆盖率。
- 端到端文档交付完成率 = 去重后的“成功打开 Word 的处理尝试数” / “成功开始处理的尝试数”。这比单纯导出成功率更接近真正交付。

`tests/metrics.test.js` 包含去重和公式正确性断言。分析时以 `attemptId` 或 `suggestionId`/`repairId` 去重；不得直接按原始事件行数相除。

## 简历写法

完成至少一轮真实用户测试后，再填写具体数字。例如：

> 构建 120 条 LaTeX 公式回归集与事件埋点体系，跟踪公式解析准确率、修复建议采纳率、误修复率及端到端 Word 交付完成率；基于错例迭代公式分隔符与兼容命令解析规则。

请使用实际测试样本量和实际结果，避免写入未经验证的百分比。
