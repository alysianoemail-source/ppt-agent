# PPT Agent 部署交接说明

这是一个五步工作流的PPT制作agent，在Claude.ai artifact里已验证可用。目标：部署成Vercel上的独立网页，任何人（无需Claude账号）可用，文本生成走DeepSeek。

## 工作流（业务逻辑，勿改）

1. **信息收集**：目标/受众/总体观点/板块/配色（人填表单）
2. **大纲骨架**：AI生成每页核心论点，人确认逻辑（不通过可重生成）
3. **每页排布**：AI生成每页观点式标题+2-3条核心观点+内容说明，人review（prompt里有"三条观点测试+反例+防编造占位"的质量规则，是调教过的，勿删）
4. **逐页细化**：AI只推荐版式（不改已审定文字），支持手绘草图上传识别、修改意见最小化修改、人工直接改文字、每页最多8版的版本管理可随时切回
5. **Agent评审**：独立评审角色审全稿，意见可一键采纳生成新版本
- 第4/5步均有导出：真实.pptx（pptxgenjs前端生成）+ HTML

## 文件结构

```
src/PPTAgent.jsx   — 完整前端（单文件React组件，Tailwind样式）
api/chat.js        — Vercel Serverless统一LLM代理（已写好，直接用）
.env.example       — 环境变量模板
package.json       — 依赖清单
```

## 必改的两处适配（artifact专属API，外部环境没有）

### 1. AI调用：callClaude → 走 /api/chat

`src/PPTAgent.jsx` 里的 `callClaude(prompt, image)` 目前直连 `https://api.anthropic.com/v1/messages`（artifact平台代理，无需key）。部署版改为调自己的后端：

```js
async function callClaude(prompt, image) {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, image: image || undefined }),
  });
  if (!response.ok) throw new Error(`API请求失败 (${response.status})`);
  const { text, error } = await response.json();
  if (error) throw new Error(error);
  // ↓ 以下JSON提取逻辑保持原样（从原函数复制）
  const cleaned = text.trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```\s*$/, "");
  const start = Math.min(...["[", "{"].map((c) => (cleaned.indexOf(c) === -1 ? Infinity : cleaned.indexOf(c))));
  const end = Math.max(cleaned.lastIndexOf("]"), cleaned.lastIndexOf("}"));
  if (start === Infinity || end === -1 || end <= start) throw new Error("AI返回的内容不是有效JSON，请重试");
  try {
    return JSON.parse(cleaned.slice(start, end + 1));
  } catch (e) {
    throw new Error("JSON解析失败，请重试");
  }
}
```

函数签名和返回值不变，5个工具函数（toolGenerateOutline等）一行都不用动。

### 2. 持久化：window.storage → localStorage

`window.storage` 是Claude artifact专属API，浏览器里不存在。全局替换（共两处调用，在根组件的两个useEffect里）：

```js
// window.storage.get(STORAGE_KEY, false) 替换为：
const raw = localStorage.getItem(STORAGE_KEY);
if (raw) setState(JSON.parse(raw));

// window.storage.set(STORAGE_KEY, JSON.stringify(state), false) 替换为：
localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
```

## 建议改（非必须）

- **pptxgenjs改为npm依赖**：现在是运行时从cdnjs动态注入script（`loadPptxLib`函数）。部署版 `npm i pptxgenjs` 后 `import PptxGenJS from "pptxgenjs"`，删掉loadPptxLib，exportPptx里直接用。
- **Tailwind**：组件用了Tailwind工具类，脚手架需配好Tailwind（Vite + tailwindcss标准配置即可），无自定义主题依赖。

## 部署步骤

1. Vite + React脚手架，装依赖，接入本组件为主页面
2. 完成上述两处适配
3. 本地 `vercel dev` 验证5步全流程 + 草图上传 + pptx导出
4. Vercel控制台配环境变量（见.env.example），部署

## 已知注意事项

- **手绘草图功能依赖视觉模型**：DeepSeek是纯文本模型看不了图。api/chat.js已做路由：带图请求自动走Anthropic。若只配DEEPSEEK_API_KEY，草图功能会报错提示（其余功能正常），属预期行为。
- **跨模型JSON稳定性**：prompt都要求"只返回紧凑JSON"，Claude下已稳定；换DeepSeek后若出现解析失败率升高，优先在api/chat.js给DeepSeek加 `response_format: { type: "json_object" }` 参数。
- **max_tokens**：api/chat.js里Anthropic是1500。第3步"每页排布"输出最大，若页数超过7-8页出现截断报错，适当调大。
- **草图不持久化**：草图只存内存（避免base64撑爆localStorage），刷新需重传，界面已有提示，是刻意设计。

