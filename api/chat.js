// api/chat.js —— Vercel Serverless Function
// 统一LLM代理：前端所有AI调用都打到这里，key只存在服务端环境变量，不暴露给浏览器
//
// 路由规则：
// - 纯文本请求 → DeepSeek（便宜，中文强）
// - 带图片请求（手绘草图功能）→ Anthropic（DeepSeek是纯文本模型，看不了图）
// - 对应key未配置时自动降级到另一家

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });
  const { prompt, image } = req.body || {};
  if (!prompt) return res.status(400).json({ error: "缺少prompt" });

  // 路线A：DeepSeek（纯文本）
  if (!image && process.env.DEEPSEEK_API_KEY) {
    try {
      const r = await fetch("https://api.deepseek.com/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${process.env.DEEPSEEK_API_KEY}`,
        },
        body: JSON.stringify({
          // 注意：deepseek-chat 模型名在 2026/07/24 后停用，届时查文档换新名，或用环境变量覆盖
          model: process.env.DEEPSEEK_MODEL || "deepseek-chat",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.4,
        }),
      });
      if (!r.ok) return res.status(502).json({ error: "DeepSeek请求失败", detail: await r.text() });
      const d = await r.json();
      return res.status(200).json({ text: d.choices?.[0]?.message?.content ?? "" });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  // 路线B：Anthropic（支持图片；也是DeepSeek未配置时的文本兜底）
  if (process.env.ANTHROPIC_API_KEY) {
    try {
      const content = image
        ? [
            { type: "image", source: { type: "base64", media_type: image.media_type, data: image.data } },
            { type: "text", text: prompt },
          ]
        : prompt;
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": process.env.ANTHROPIC_API_KEY,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: process.env.ANTHROPIC_MODEL || "claude-sonnet-4-6",
          max_tokens: 1500,
          messages: [{ role: "user", content }],
        }),
      });
      if (!r.ok) return res.status(502).json({ error: "Anthropic请求失败", detail: await r.text() });
      const d = await r.json();
      const text = (d.content || [])
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("\n");
      return res.status(200).json({ text });
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(500).json({
    error: image
      ? "带图片的请求需要配置 ANTHROPIC_API_KEY（DeepSeek不支持图片输入）"
      : "未配置任何模型API key（DEEPSEEK_API_KEY 或 ANTHROPIC_API_KEY）",
  });
}
