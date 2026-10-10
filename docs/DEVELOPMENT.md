# AI 学术内容结构化与 Word 交付助手：迭代开发设计文档

## 1. 文档目标

本文档定义**下一阶段需要新增或改造**的功能，不重复说明现有代码。目标是将当前“文本解析并导出 Word”的原型升级为具备 AI 校验、可控修改、模板化交付与质量检查能力的产品。

产品原则：AI 负责发现问题和提出建议，用户始终拥有最终决定权；所有导出内容都可追溯到用户确认后的版本。

## 2. 迭代总览

```text
输入 AI 内容
  → 结构化解析（扩展）
  → 公式校验 / AI 修复建议
  → 文档健康度检查
  → 用户逐项确认修改
  → 选择导出模板
  → 高保真 Word 生成与下载
```

| 版本 | 开发模块 | 用户价值 | 优先级 |
|---|---|---|---|
| v0.2 | 草稿、完整预览、模板选择 | 完成可反复使用的基础体验 | P0 |
| v0.3 | 公式规范化与公式问题清单 | 减少最核心的公式混乱 | P0 |
| v0.4 | AI 修改建议与差异确认 | 在不失控的前提下提升内容质量 | P0 |
| v0.5 | 文档健康度检查 | 导出前发现遗漏与结构错误 | P1 |
| v1.0 | 公式高保真/可编辑导出、数据监控 | 建立产品差异化并支撑正式使用 | P1 |

## 3. 公共技术改造

### 3.1 模块拆分

将当前单页的解析、预览和导出逻辑拆分，避免后续功能都堆积在 `pages/ai-word/index.js`：

```text
miniprogram/
├─ pages/editor/                 # 内容编辑、结构预览、问题处理
├─ pages/template-select/        # 模板选择与参数设置
├─ pages/export-result/          # 导出状态和下载
├─ services/
│  ├─ document.js                # 云函数调用封装
│  ├─ draft.js                   # 本地/云端草稿读写
│  └─ ai-review.js               # AI 检查接口调用
├─ utils/
│  ├─ parser.js                  # Markdown/公式边界解析
│  ├─ sanitizer.js               # 内容清洗和长度限制
│  └─ diff.js                    # 建议前后差异计算
└─ components/
   ├─ document-block/            # 各类块的预览组件
   ├─ issue-card/                # 问题与修改建议卡片
   └─ template-card/             # 模板卡片
```

云函数应拆分为 `documentParse`、`aiReview`、`generateWord`，而不是继续在 `quickstartFunctions` 中追加业务分支。这样能单独设置超时、权限、日志和限流。

### 3.2 统一文档模型

后续所有功能围绕 `DocumentDraft` 工作，原始输入与已确认内容必须分开保存。

```ts
type DocumentDraft = {
  id: string;
  title: string;
  sourceText: string;            // 用户原始输入，不被 AI 覆盖
  blocks: DocumentBlock[];       // 当前已确认结构
  issues: DocumentIssue[];       // 待处理/已处理问题
  templateId: 'homework' | 'lab-report' | 'study-note';
  status: 'editing' | 'reviewing' | 'ready' | 'exporting' | 'exported';
  revision: number;
  createdAt: number;
  updatedAt: number;
};

type DocumentBlock = {
  id: string;
  type: 'heading' | 'paragraph' | 'quote' | 'code' | 'list' | 'table' | 'formula';
  content: string | string[][];
  level?: number;
  formulaFormat?: 'latex' | 'plain';
  sourceRange?: { start: number; end: number };
};

type DocumentIssue = {
  id: string;
  category: 'formula' | 'structure' | 'placeholder' | 'reference';
  severity: 'error' | 'warning' | 'info';
  blockId?: string;
  message: string;
  suggestion?: { before: string; after: string; reason: string; confidence: number };
  status: 'pending' | 'accepted' | 'rejected' | 'fixed';
};
```

接口传输时仅使用 JSON；`rendered`、富文本节点等视图字段不得写入数据库或发送给导出服务。

## 4. v0.2：编辑、预览、草稿与模板

### 4.1 完整结构预览

**需求**：预览页需要分别展示标题、列表、代码、公式、表格和引用，点击任意区块可回到编辑状态。

**实现**：新增 `document-block` 组件，按 `block.type` 选择渲染器。公式块在本版本先以“公式标签 + 原始 LaTeX”展示；无法渲染的内容不能隐藏。

**验收**：混合包含六种块类型的输入，预览顺序、文本内容和换行与解析结果一致；任一块都可定位编辑。

### 4.2 草稿持久化

**需求**：用户退出小程序后可恢复未导出的文档，默认保存最近编辑版本。

**实现路径**：

1. 每次停止输入 800ms 后写入 `wx.setStorage`；
2. 用户主动点击“保存草稿”时写入云数据库 `documentDrafts`；
3. 云端草稿以用户 openid 为分区，查询和更新都必须携带所有者条件；
4. 冲突策略采用“最新保存版本”，并保留 `revision` 供后续扩展。

**验收**：杀死小程序再进入，能恢复最近内容；不同用户不能访问彼此草稿；删除草稿有二次确认。

### 4.3 模板系统

**首批模板**：课程作业、实验报告、学术笔记。

模板是导出配置而非对正文的二次改写：

```js
{
  id: 'lab-report',
  name: '实验报告',
  cover: true,
  styles: {
    font: '宋体',
    bodySize: 12,
    lineSpacing: 1.5,
    headingNumbering: true,
    formulaAlignment: 'center'
  },
  requiredSections: ['实验目的', '实验原理', '实验步骤', '实验结果', '结论']
}
```

**验收**：切换模板不改变用户正文；生成的 Word 在不同模板下有可见、稳定的排版差异。

## 5. v0.3：公式规范化与校验

### 5.1 支持范围

需要新增识别以下公式边界：块级 `$$...$$`、`\\[...\\]`，行内 `$...$`、`\\(...\\)`。其中美元符号出现在金额文本时不视为公式，例如“价格为 $10”。

公式解析结果应作为独立 `formula` 块或段落内的 `inlineFormula` 片段保存，不能继续混在普通字符串中。

### 5.2 非 AI 规则校验

先实现确定性检查，避免所有问题都交由大模型：

- `()[]{}` 是否成对；
- `\\frac`、`\\sqrt`、`\\text` 等命令的大括号参数是否完整；
- `^` 和 `_` 后是否缺少内容；
- 块级公式开始/结束符是否匹配；
- 全角运算符、中文标点是否需要规范化；
- 公式中是否存在明显非法控制字符。

规则校验只生成问题和修复建议；当存在歧义时，`confidence` 必须低于 0.8，默认不可自动应用。

### 5.3 公式问题交互

问题卡片应包含原始公式、建议公式、错误原因、置信度、接受/忽略按钮。接受后只替换对应 `blockId` 的对应内容，记录一次版本变更；忽略后不再对同一版本重复提示。

**验收**：

- 输入未闭合括号或未闭合公式边界，能定位到对应公式；
- 接受建议后预览和导出内容同步更新；
- 忽略建议后不改动原文；
- 规则无法判断时不会伪造“已修复”。

## 6. v0.4：AI 文档审校与可控改写

### 6.1 功能范围

AI 仅处理以下任务：公式修复建议、标题层级建议、文本正式化、表格转结构化数据、引用/占位内容提醒。首版不提供开放式聊天入口，以避免范围失控。

### 6.2 后端调用链

```text
小程序选择“AI 检查”
  → aiReview 云函数：身份校验、限流、脱敏、截断
  → 大模型 API
  → 严格校验 JSON Schema
  → 返回 issue 列表
  → 前端逐项展示并等待用户确认
```

API Key 必须仅保存于云函数环境变量。请求应限制单次字符数、最大块数量和超时；对超长文档进行分块审校，再合并结果。

### 6.3 AI 输出协议

模型必须返回 JSON，禁止直接返回可执行代码或 HTML：

```json
{
  "version": "1.0",
  "issues": [
    {
      "blockId": "formula-3",
      "category": "formula",
      "severity": "warning",
      "message": "分式右花括号缺失",
      "suggestion": {
        "before": "\\frac{a}{b",
        "after": "\\frac{a}{b}",
        "reason": "补全缺失的右花括号",
        "confidence": 0.91
      }
    }
  ]
}
```

服务端校验 `blockId` 必须属于当前文档，`before` 必须匹配该块的当前内容；任一条件不满足时丢弃建议，防止模型错误改写无关文本。

### 6.4 版本与撤销

每一次“接受 AI 建议”生成一条变更记录：`draftId`、`blockId`、修改前后文本、建议 ID、时间。用户可撤销最近一次或从“修改记录”恢复任意版本。

**验收**：AI 不能直接覆盖原文；所有接受的改动可撤销；AI 服务异常时仍可使用基础导出功能。

## 7. v0.5：文档健康度检查

### 7.1 检查项和规则

| 分类 | 规则 | 严重级别 |
|---|---|---|
| 公式 | 无法解析、边界不闭合、疑似非法命令 | 错误/警告 |
| 结构 | 标题级别跳跃、文档无一级标题 | 警告 |
| 编号 | 图/表/公式编号重复或不连续 | 警告 |
| 占位 | `XXX`、`待补充`、`TODO` | 警告 |
| 引用 | 出现“如下图/如下表”但对应对象不存在 | 警告 |
| 导出 | 模板必填章节缺失 | 错误 |

### 7.2 健康度计算

初始分为 100。错误每项扣 15 分，警告每项扣 5 分，最低为 0。评分仅作为辅助，页面必须同时展示具体问题，不能只给分数。

导出规则：存在错误时提示“建议先修复”，用户仍可选择“继续导出”；存在警告时允许直接导出。

**验收**：任意检查项可跳转到对应块；修复问题后自动重新计算；用户可清楚了解导出风险。

## 8. v1.0：高保真 Word 导出

### 8.1 导出能力路线

| 阶段 | 技术方案 | 适用范围 | 限制 |
|---|---|---|---|
| 1 | 样式化文本 + 公式原文 | 最快上线 | 公式不可编辑 |
| 2 | LaTeX 渲染为 PNG/SVG 并插入 | 视觉正确性优先 | 公式为图片 |
| 3 | LaTeX/MathML 转 OMML | Word 原生公式 | 开发与兼容成本高 |

v1.0 以阶段 2 为最低交付标准，并将阶段 3 作为后续增强。产品界面必须明确展示“图片公式”或“可编辑公式”的实际状态。

### 8.2 导出服务重构

新建 `generateWord` 云函数，输入为已确认的 `DocumentDraft` 快照和模板 ID；后端重新校验块类型与容量，不信任客户端传入样式。输出文件路径使用 `userId/draftId/uuid.docx`，设置 7～30 天自动清理策略。

Word 映射要求：标题映射样式级别；有序/无序列表分开处理；代码块使用等宽字体和背景色；表格使用标准表头样式；公式按选择的导出策略插入并支持居中、编号。

**验收**：在 Microsoft Word 与 WPS 均可打开；三种模板的层级、列表、代码、表格、公式都可读；导出失败可重试且不会生成重复的无效记录。

## 9. 数据、安全与监控

### 9.1 数据集合

- `documentDrafts`：草稿正文与当前版本；
- `documentRevisions`：用户或 AI 变更记录；
- `exportJobs`：导出任务状态、文件 ID、过期时间；
- `usageMetrics`：不含原文的操作事件与性能数据。

每条业务数据必须保存 `ownerOpenId`；客户端查询不能仅按文档 ID 读取。用户原始内容默认不用于模型训练，只有在用户主动提交 AI 审校时才发送必要文本给模型服务。

### 9.2 关键埋点

`draft_created`、`parse_completed`、`issue_found`、`suggestion_accepted`、`suggestion_rejected`、`template_selected`、`export_started`、`export_succeeded`、`export_failed`。

核心看板：导出成功率、平均导出耗时、公式问题率、AI 建议采纳率、健康检查修复率、7 日复用率。

## 10. 开发顺序与验收门禁

1. 先完成文档数据模型和模块拆分；
2. 实现预览、草稿和模板，确保基础链路稳定；
3. 上线规则型公式检查，再接入 AI 审校；
4. 增加健康检查和导出前确认；
5. 最后投入公式渲染/原生公式导出。

每项功能合并前必须通过：单元测试（规则和数据转换）、云函数接口测试、真机流程测试、Word/WPS 文件验证。AI 功能额外要求 JSON Schema 校验、超时降级与“全部建议可撤销”。

