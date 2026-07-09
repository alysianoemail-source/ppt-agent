import React, { useState, useEffect, useRef } from "react";
import {
  Check,
  ChevronRight,
  ChevronLeft,
  Plus,
  X,
  RotateCcw,
  AlertTriangle,
  Pencil,
  Sparkles,
  Loader2,
} from "lucide-react";
import PptxGenJS from "pptxgenjs";

// ============================================================
// Agent 三层结构：
// 1. TOOLS —— 两个真实AI调用（generateOutline / refinePage）
// 2. ORCHESTRATION —— 四步状态机，每步有准入门槛
// 3. MEMORY —— window.storage 持久化，刷新不丢
// ============================================================

// ---- design tokens (Apple Settings style) ---------------------------
const bg = "#F5F5F7";
const cardBg = "#FFFFFF";
const divider = "#E5E5EA";
const ink = "#1D1D1F";
const inkSecondary = "#6E6E73";
const inkTertiary = "#AEAEB2";
const blue = "#0071E3";
const green = "#34C759";
const greenDeep = "#1F7A3D";
const orange = "#FF9500";
const orangeDeep = "#B86B00";
const grey = "#AEAEB2";
const FONT = '-apple-system, BlinkMacSystemFont, "SF Pro Text", system-ui, sans-serif';

const STEP_META = [
  { id: 1, label: "信息收集" },
  { id: 2, label: "大纲骨架" },
  { id: 3, label: "每页排布" },
  { id: 4, label: "逐页细化" },
  { id: 5, label: "Agent评审" },
];

const STORAGE_KEY = "ppt-agent-state-v2"; // 结构变了，换key避免旧存档冲突

const defaultState = {
  step: 1,
  form: {
    goal: "",
    audience: "",
    mainPoint: "",
    sections: [""],
    visualPriority: null,
    color: "",
    colorRole: null,
  },
  outline: null, // 真实AI生成后填入 [{section, point}]
  outlineStatus: null,
  pagePlan: null, // [{section, title, points[], detail}] —— 每页排布，供人review
  pagePlanStatus: null,
  pages: [],
  review: null, // 评审agent的结果 {overall:[], pages:[{idx, issues:[]}]}
};

// ============================================================
// TOOLS 层：真实AI调用
// ============================================================

async function callClaude(prompt, image) {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, image: image || undefined }),
  });
  if (!response.ok) throw new Error(`API请求失败 (${response.status})`);
  const { text, error } = await response.json();
  if (error) throw new Error(error);
  // 健壮提取：截取第一个 [ 或 { 到最后一个 ] 或 } 之间的内容再解析
  const cleaned = text.trim().replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```\s*$/, "");
  const start = Math.min(...["[", "{"].map((c) => (cleaned.indexOf(c) === -1 ? Infinity : cleaned.indexOf(c))));
  const end = Math.max(cleaned.lastIndexOf("]"), cleaned.lastIndexOf("}"));
  if (start === Infinity || end === -1 || end <= start) {
    throw new Error("AI返回的内容不是有效JSON（可能内容被截断），请重试");
  }
  try {
    return JSON.parse(cleaned.slice(start, end + 1));
  } catch (e) {
    throw new Error("JSON解析失败（内容可能被截断），请点重试");
  }
}

// 工具①：生成大纲（第2步）
async function toolGenerateOutline(form) {
  const sections = form.sections.filter((s) => s.trim());
  const prompt = `你是PPT内容策划助手。根据以下信息，为每个板块生成一句话核心论点。
只返回JSON数组，不要任何多余文字、不要markdown代码块。每项包含：
- section：板块名称（原样使用输入的板块名）
- point：这一页的核心论点，一句话，必须能支撑总体观点

目标：${form.goal}
受众：${form.audience}
总体观点：${form.mainPoint}
板块：${sections.join("、")}`;
  return callClaude(prompt);
}

// 工具②：每页排布（第3步）——逐页生成，每次只出一页，避免多页JSON超长被截断
async function toolGeneratePagePlanForPage(outlineItem, outline, form, pageIndex) {
  const prompt = `你是PPT内容策划助手。基于已确认的大纲，为其中第${pageIndex + 1}页生成排布，供人审核。

标题和核心观点必须是真正的"观点"，用以下三条测试逐条自检，不通过就重写：
1. 可反驳测试：一个持怀疑态度的听众有可能不同意这句话吗？描述已发生的事实（"我们半年前做了X"）没人能反对，不算观点，必须改写成可以被质疑的断言。
2. 新信息测试：句中每个分句必须有独立的信息增量。同义反复、语气强调（如"不是估算是实测"只是在强调前半句）直接删掉或换成真实信息（如数据来源、口径）。
3. 具体锚点测试：观点必须包含至少一个可检验的锚点（数字、比较对象、时间范围）。"质量没有下降"是空断言，必须写成"合格率98%对97%持平"这种。

反例参照（真实审核意见）：
- 差："30%提效目标已兑现，不是估算是实测"（后半句零信息）→ 好："30%提效已兑现，来自两团队3个月工时实测"
- 差："半年前我们用一个假设押注AI提效"（叙述不是观点）→ 好："半年前的押注已验证：AI提效在我们的场景成立"
- 差："省时没有牺牲质量"（无锚点空断言）→ 好："效率翻四倍且质量持平：抽检合格率98%对97%"

重要：锚点里的数字只能来自用户提供的信息，没有依据的数字一律写成【需补数据：XX】占位，禁止编造。

只返回一个紧凑JSON对象（不是数组），不要任何多余文字、不要markdown代码块。包含字段：
- section：板块名称（原样保留：${outlineItem.section}）
- title：这一页的标题，通过上述三条测试的观点句，20字以内
- points：2到3条核心观点，字符串数组，每条25字以内，逐条通过三条测试
- detail：主要内容说明，1-2句话（60字以内）——写明这页用什么数据/例子/论证思路

整体背景——目标：${form.goal}；受众：${form.audience}；总体观点：${form.mainPoint}
全稿大纲（供把握上下文，只生成第${pageIndex + 1}页）：
${outline.map((p, i) => `${i + 1}. ${p.section}：${p.point}`).join("\n")}

本页（第${pageIndex + 1}页）：${outlineItem.section}：${outlineItem.point}`;
  return callClaude(prompt);
}

// 工具③：版式推荐（第4步）——只定"怎么摆"，不碰已审定的内容
async function toolLayoutPage(page, form, sketch) {
  const prompt = `你是PPT版式助手。这一页的文字内容已经人工审定，你的任务只有一件：为它选版式。禁止改写、新增、删减任何已审定文字。
${sketch ? "用户上传了手绘草图（见图片），优先按草图里的版式意图（分栏、图文位置、流程方向等）选择最接近的版式类型。\n" : ""}只返回JSON对象，不要任何多余文字、不要markdown代码块。包含字段：
- layout：必须是这四个值之一："左图右文" / "三栏对比" / "时间轴" / "纯文字要点"
- reason：一句话说明为什么这个版式适合这页内容${sketch ? "，并指出草图哪个特征对应这个选择" : ""}

本页已审定标题：${page.title || page.point}
本页已审定核心观点：${(page.points || []).join("；")}
本页内容说明：${page.detail || "无"}
用户版式直觉补充：${page.content || "无"}
视觉复杂度偏好：${form.visualPriority === "design" ? "兼顾设计感" : "内容优先，版式从简"}`;
  return callClaude(prompt, sketch);
}

// 工具④：按修改意见修改单页（第4步）——最小修改原则，人主导
async function toolEditPage(current, instruction, form) {
  const prompt = `你是PPT单页修改助手。基于当前版本和用户的修改意见做最小修改：只改意见明确涉及的部分，其余文字逐字保留，禁止顺手润色。
没有依据的数字一律写成【需补数据：XX】占位，禁止编造。
只返回JSON对象，不要任何多余文字、不要markdown代码块。包含字段：
- layout：必须是这四个值之一："左图右文" / "三栏对比" / "时间轴" / "纯文字要点"（意见没提版式就保持原值）
- title：标题
- bullets：2到5条要点，字符串数组

当前版本：
版式：${current.layout}
标题：${current.title}
要点：${(current.bullets || []).join("；")}

用户修改意见：${instruction}

整体背景——目标：${form.goal}；受众：${form.audience}；总体观点：${form.mainPoint}`;
  return callClaude(prompt);
}

// ============================================================
// 版式渲染器：把 layout/title/bullets 显示成"像PPT页"的样子
// ============================================================

function PageRenderer({ data, colorHint, large }) {
  if (!data?.layout) return null;
  const accent = colorHint || blue;
  // large：放大查看模式，字号/间距按比例放大
  const s = large
    ? { title: 26, body: 16, pad: 36, gap: 20, dot: 14 }
    : { title: 15, body: 12, pad: 16, gap: 12, dot: 10 };

  const frame = {
    background: "#FFFFFF",
    border: `1px solid ${divider}`,
    borderRadius: large ? 14 : 10,
    padding: s.pad,
    aspectRatio: "16/9",
    overflow: "hidden",
    width: "100%",
  };

  const titleStyle = { color: ink, borderLeft: `3px solid ${accent}`, paddingLeft: s.pad / 3, fontSize: s.title, fontWeight: 600, marginBottom: s.gap };
  const bodyStyle = { color: inkSecondary, fontSize: s.body };

  if (data.layout === "左图右文") {
    return (
      <div style={{ ...frame, display: "flex", gap: s.gap }}>
        <div className="w-1/3 rounded-md flex items-center justify-center" style={{ background: bg, color: inkTertiary, fontSize: s.body - 1 }}>
          图片占位
        </div>
        <div className="flex-1 min-w-0">
          <h3 style={titleStyle} className="truncate">{data.title}</h3>
          <ul style={{ ...bodyStyle, display: "flex", flexDirection: "column", gap: s.gap / 2 }}>
            {data.bullets.map((b, i) => (
              <li key={i} className="flex gap-1.5">
                <span style={{ color: accent }}>•</span>
                <span className={large ? "" : "line-clamp-1"}>{b}</span>
              </li>
            ))}
          </ul>
        </div>
      </div>
    );
  }

  if (data.layout === "三栏对比") {
    return (
      <div style={frame}>
        <h3 style={titleStyle}>{data.title}</h3>
        <div className="grid grid-cols-3" style={{ gap: s.gap / 1.5 }}>
          {data.bullets.slice(0, 3).map((b, i) => (
            <div key={i} className={`rounded-md leading-snug ${large ? "" : "line-clamp-3"}`} style={{ background: bg, color: inkSecondary, fontSize: s.body - 1, padding: s.pad / 2 }}>
              {b}
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (data.layout === "时间轴") {
    return (
      <div style={frame}>
        <h3 style={titleStyle}>{data.title}</h3>
        <div className="flex items-start gap-1" style={{ marginTop: s.gap }}>
          {data.bullets.map((b, i) => (
            <React.Fragment key={i}>
              <div className="flex-1 min-w-0">
                <div className="rounded-full" style={{ width: s.dot, height: s.dot, background: accent, marginBottom: s.gap / 2 }} />
                <p className={`leading-snug ${large ? "" : "line-clamp-2"}`} style={{ color: inkSecondary, fontSize: s.body - 1 }}>
                  {b}
                </p>
              </div>
              {i < data.bullets.length - 1 && <div className="flex-shrink-0" style={{ height: 2, width: s.gap, background: divider, marginTop: s.dot / 2 }} />}
            </React.Fragment>
          ))}
        </div>
      </div>
    );
  }

  // 纯文字要点（默认）
  return (
    <div style={frame}>
      <h3 style={titleStyle}>{data.title}</h3>
      <ul style={{ ...bodyStyle, display: "flex", flexDirection: "column", gap: s.gap / 2 }}>
        {data.bullets.map((b, i) => (
          <li key={i} className="flex gap-1.5">
            <span style={{ color: accent }}>•</span>
            <span className={large ? "" : "line-clamp-1"}>{b}</span>
          </li>
        ))}
      </ul>
    </div>
  );
}

// ---- UI primitives -------------------------------------------------

function SegmentedControl({ options, value, onChange }) {
  return (
    <div className="flex p-[2px] rounded-[10px] gap-[2px]" style={{ background: "#E8E8ED" }}>
      {options.map((opt) => {
        const active = value === opt.value;
        return (
          <button
            key={opt.value}
            onClick={() => onChange(opt.value)}
            className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded-[8px] text-[14px] font-medium transition-all"
            style={{
              background: active ? "#FFFFFF" : "transparent",
              color: active ? ink : inkSecondary,
              boxShadow: active ? "0 1px 2px rgba(0,0,0,0.08)" : "none",
            }}
          >
            {opt.icon && <opt.icon size={14} style={{ color: active ? opt.color || blue : inkTertiary }} />}
            {opt.label}
          </button>
        );
      })}
    </div>
  );
}

function GroupCard({ children }) {
  return (
    <div className="rounded-[14px] overflow-hidden" style={{ background: cardBg, boxShadow: "0 1px 2px rgba(0,0,0,0.04)" }}>
      {children}
    </div>
  );
}

function GroupRow({ children, last }) {
  return (
    <div className="px-4 py-3.5" style={{ borderBottom: last ? "none" : `0.5px solid ${divider}` }}>
      {children}
    </div>
  );
}

function FieldLabel({ children, optional }) {
  return (
    <label className="block text-[13px] font-medium mb-1.5" style={{ color: inkSecondary }}>
      {children}
      {optional && <span className="ml-1.5 font-normal" style={{ color: inkTertiary }}>（可选）</span>}
    </label>
  );
}

const inputStyle = {
  background: "transparent",
  border: "none",
  color: ink,
  fontSize: 15,
  fontFamily: FONT,
  padding: 0,
  width: "100%",
  outline: "none",
};

function PrimaryButton({ disabled, onClick, children, loading }) {
  return (
    <button
      disabled={disabled || loading}
      onClick={onClick}
      className="flex-1 flex items-center justify-center gap-2 py-3.5 rounded-full font-semibold text-[15px] transition-opacity"
      style={{ background: disabled ? "#D2D2D7" : blue, color: "#FFFFFF", cursor: disabled || loading ? "not-allowed" : "pointer" }}
    >
      {loading ? <Loader2 size={16} className="animate-spin" /> : null}
      {children}
    </button>
  );
}

function BackButton({ onClick }) {
  return (
    <button
      onClick={onClick}
      className="flex items-center justify-center gap-1.5 px-5 py-3.5 rounded-full text-[15px] font-medium"
      style={{ background: "#FFFFFF", color: ink, boxShadow: "0 1px 2px rgba(0,0,0,0.04)" }}
    >
      <ChevronLeft size={16} /> 上一步
    </button>
  );
}

// ---- stepper ---------------------------------------------------------

function Stepper({ step }) {
  return (
    <div className="flex items-center w-full mb-8 select-none">
      {STEP_META.map((s, i) => {
        const done = s.id < step;
        const current = s.id === step;
        return (
          <React.Fragment key={s.id}>
            <div className="flex flex-col items-center gap-2" style={{ minWidth: 60 }}>
              <div
                className="w-8 h-8 rounded-full flex items-center justify-center font-semibold text-[13px] transition-colors"
                style={{
                  background: done || current ? blue : "#FFFFFF",
                  border: done || current ? "none" : `1.5px solid ${divider}`,
                  color: done || current ? "#FFFFFF" : inkTertiary,
                }}
              >
                {done ? <Check size={15} /> : s.id}
              </div>
              <span className="text-[12px] font-medium text-center leading-tight" style={{ color: done || current ? ink : inkTertiary }}>
                {s.label}
              </span>
            </div>
            {i < STEP_META.length - 1 && (
              <div className="flex-1 h-[2px] mx-1 rounded-full" style={{ background: s.id < step ? blue : divider, marginBottom: 22 }} />
            )}
          </React.Fragment>
        );
      })}
    </div>
  );
}

// ---- Step 1 ------------------------------------------------------------

function StepIntake({ form, setForm, onGenerate, generating }) {
  const updateSection = (i, val) => {
    const next = [...form.sections];
    next[i] = val;
    setForm({ ...form, sections: next });
  };
  const addSection = () => setForm({ ...form, sections: [...form.sections, ""] });
  const removeSection = (i) => {
    if (form.sections.length === 1) return;
    setForm({ ...form, sections: form.sections.filter((_, idx) => idx !== i) });
  };
  const canProceed = form.goal.trim() && form.audience.trim() && form.mainPoint.trim() && form.sections.some((s) => s.trim());

  return (
    <div className="space-y-6">
      <GroupCard>
        <GroupRow>
          <FieldLabel>PPT目标 — 这次要达成什么</FieldLabel>
          <textarea value={form.goal} onChange={(e) => setForm({ ...form, goal: e.target.value })} placeholder="例：说服管理层批准Q3新渠道预算" rows={2} className="resize-none" style={inputStyle} />
        </GroupRow>
        <GroupRow>
          <FieldLabel>给谁看</FieldLabel>
          <input value={form.audience} onChange={(e) => setForm({ ...form, audience: e.target.value })} placeholder="例：管理层，已经知道渠道现状，但对新渠道ROI存疑" style={inputStyle} />
        </GroupRow>
        <GroupRow>
          <FieldLabel>总体观点 — 要让对方记住的一句话</FieldLabel>
          <input value={form.mainPoint} onChange={(e) => setForm({ ...form, mainPoint: e.target.value })} placeholder="例：新渠道三个月内能把获客成本降三成" style={inputStyle} />
        </GroupRow>
        <GroupRow last>
          <FieldLabel optional>配色偏好</FieldLabel>
          <input value={form.color} onChange={(e) => setForm({ ...form, color: e.target.value, colorRole: null })} placeholder="例：橙白配色（留空则由AI建议安全方案）" style={inputStyle} />
          {form.color.trim() && (
            <div className="mt-3">
              <SegmentedControl
                options={[
                  { value: "accent", label: "白底+点缀" },
                  { value: "primary", label: "该色为主" },
                  { value: "ai", label: "AI决定" },
                ]}
                value={form.colorRole}
                onChange={(v) => setForm({ ...form, colorRole: v })}
              />
            </div>
          )}
        </GroupRow>
      </GroupCard>

      <GroupCard>
        <GroupRow last>
          <FieldLabel>视觉复杂度</FieldLabel>
          <SegmentedControl
            options={[
              { value: "content", label: "内容优先，省心" },
              { value: "design", label: "兼顾设计感" },
            ]}
            value={form.visualPriority}
            onChange={(v) => setForm({ ...form, visualPriority: v })}
          />
          <p className="text-[12px] mt-2.5 leading-relaxed" style={{ color: inkTertiary }}>
            选哪个都不放松文字溢出 / 对齐 / 留白这几条底线
          </p>
        </GroupRow>
      </GroupCard>

      <GroupCard>
        <div className="px-4 pt-3.5 pb-1">
          <FieldLabel>内容板块</FieldLabel>
        </div>
        {form.sections.map((sec, i) => (
          <GroupRow key={i}>
            <div className="flex items-center gap-3">
              <span className="text-[13px] font-medium w-5 text-center" style={{ color: inkTertiary }}>
                {i + 1}
              </span>
              <input value={sec} onChange={(e) => updateSection(i, e.target.value)} placeholder="例：现状与问题" style={inputStyle} />
              <button onClick={() => removeSection(i)} className="p-1 rounded-full hover:opacity-60" style={{ color: inkTertiary }} aria-label="删除板块">
                <X size={15} />
              </button>
            </div>
          </GroupRow>
        ))}
        <div className="px-4 py-3">
          <button onClick={addSection} className="flex items-center gap-1.5 text-[14px] font-medium" style={{ color: blue }}>
            <Plus size={15} /> 添加板块
          </button>
        </div>
      </GroupCard>

      <div className="flex">
        <PrimaryButton disabled={!canProceed} loading={generating} onClick={onGenerate}>
          {generating ? "AI正在生成大纲…" : "生成内容大纲"} {!generating && <ChevronRight size={16} />}
        </PrimaryButton>
      </div>
    </div>
  );
}

// ---- Step 2 -------------------------------------------------------------

function StepOutline({ outline, error, outlineStatus, setOutlineStatus, onBack, onRegenerate, onNext, regenerating, nextLoading }) {
  return (
    <div className="space-y-6">
      {error && (
        <div className="rounded-[14px] px-4 py-3 flex items-center gap-2.5 text-[13px]" style={{ background: `${orange}12`, color: orangeDeep }}>
          <AlertTriangle size={14} style={{ flexShrink: 0 }} />
          生成失败：{error}。可以点"重新生成"再试。
        </div>
      )}

      {outline && (
        <GroupCard>
          {outline.map((p, i) => (
            <GroupRow key={i} last={i === outline.length - 1}>
              <p className="text-[13px] font-medium mb-1" style={{ color: inkSecondary }}>
                {p.section}
              </p>
              <p className="text-[15px] leading-relaxed" style={{ color: ink }}>
                {p.point}
              </p>
            </GroupRow>
          ))}
        </GroupCard>
      )}

      <GroupCard>
        <GroupRow last>
          <FieldLabel>逻辑顺序OK吗？</FieldLabel>
          <SegmentedControl
            options={[
              { value: "approved", label: "没问题，继续", color: green },
              { value: "fix", label: "有遗漏/顺序不对", color: orange, icon: AlertTriangle },
            ]}
            value={outlineStatus}
            onChange={setOutlineStatus}
          />
          {outlineStatus === "fix" && (
            <div className="mt-3">
              <button onClick={onRegenerate} disabled={regenerating} className="flex items-center gap-1.5 text-[14px] font-medium" style={{ color: blue }}>
                {regenerating ? <Loader2 size={14} className="animate-spin" /> : <RotateCcw size={14} />}
                {regenerating ? "重新生成中…" : "回第1步改信息后重新生成，或直接点此重试"}
              </button>
            </div>
          )}
        </GroupRow>
      </GroupCard>

      <div className="flex gap-2.5">
        <BackButton onClick={onBack} />
        <PrimaryButton disabled={outlineStatus !== "approved"} loading={nextLoading} onClick={onNext}>
          {nextLoading ? "AI正在生成每页排布…" : "生成每页排布"} {!nextLoading && <ChevronRight size={16} />}
        </PrimaryButton>
      </div>
    </div>
  );
}

// ---- Step 3：每页排布 review ------------------------------------------------

function StepPagePlan({ pagePlan, error, pagePlanStatus, setPagePlanStatus, onBack, onRegenerate, onRetryPage, onNext, regenerating, progress }) {
  const allReady = pagePlan && pagePlan.length > 0 && pagePlan.every((p) => !p.pending && !p.failed);
  return (
    <div className="space-y-6">
      {regenerating && progress?.total > 0 && (
        <div className="rounded-[14px] px-4 py-3 flex items-center gap-2.5 text-[13px]" style={{ background: `${blue}10`, color: blue }}>
          <Loader2 size={14} className="animate-spin" style={{ flexShrink: 0 }} />
          逐页生成中：{progress.done} / {progress.total} 页（生成一页显示一页，不用等全部）
        </div>
      )}

      {error && !regenerating && (
        <div className="rounded-[14px] px-4 py-3 flex items-center gap-2.5 text-[13px]" style={{ background: `${orange}12`, color: orangeDeep }}>
          <AlertTriangle size={14} style={{ flexShrink: 0 }} />
          {error}
        </div>
      )}

      {!pagePlan && (
        <GroupCard>
          <GroupRow last>
            <div className="py-4 text-center">
              <p className="text-[14px] mb-3" style={{ color: inkSecondary }}>
                还没有生成每页排布
              </p>
              <button
                onClick={onRegenerate}
                disabled={regenerating}
                className="inline-flex items-center gap-1.5 text-[14px] font-medium px-4 py-2 rounded-full"
                style={{ background: `${blue}10`, color: blue }}
              >
                {regenerating ? <Loader2 size={14} className="animate-spin" /> : <RotateCcw size={14} />}
                {regenerating ? "生成中…" : "生成每页排布"}
              </button>
            </div>
          </GroupRow>
        </GroupCard>
      )}

      {pagePlan &&
        pagePlan.map((p, i) => {
          if (p.pending) {
            return (
              <GroupCard key={i}>
                <GroupRow last>
                  <p className="text-[12px] font-medium mb-1" style={{ color: inkTertiary }}>
                    第{i + 1}页 · {p.section}
                  </p>
                  {regenerating ? (
                    <div className="flex items-center gap-2 py-2 text-[14px]" style={{ color: inkSecondary }}>
                      <Loader2 size={14} className="animate-spin" /> 生成中…
                    </div>
                  ) : (
                    <button
                      onClick={() => onRetryPage(i)}
                      className="inline-flex items-center gap-1.5 text-[13px] font-medium px-3 py-1.5 rounded-full"
                      style={{ background: `${blue}10`, color: blue }}
                    >
                      <RotateCcw size={13} /> 生成这一页（上次可能被中断）
                    </button>
                  )}
                </GroupRow>
              </GroupCard>
            );
          }
          if (p.failed) {
            return (
              <GroupCard key={i}>
                <GroupRow last>
                  <p className="text-[12px] font-medium mb-1" style={{ color: inkTertiary }}>
                    第{i + 1}页 · {p.section}
                  </p>
                  <p className="text-[13px] mb-2" style={{ color: orangeDeep }}>
                    这一页生成失败{p.error ? `：${p.error}` : ""}
                  </p>
                  <button
                    onClick={() => onRetryPage(i)}
                    className="inline-flex items-center gap-1.5 text-[13px] font-medium px-3 py-1.5 rounded-full"
                    style={{ background: `${blue}10`, color: blue }}
                  >
                    <RotateCcw size={13} /> 重试这一页
                  </button>
                </GroupRow>
              </GroupCard>
            );
          }
          return (
            <GroupCard key={i}>
              <GroupRow>
                <p className="text-[12px] font-medium mb-1" style={{ color: inkTertiary }}>
                  第{i + 1}页 · {p.section}
                </p>
                <p className="text-[16px] font-semibold leading-snug" style={{ color: ink }}>
                  {p.title}
                </p>
              </GroupRow>
              <GroupRow>
                <FieldLabel>核心观点</FieldLabel>
                <ul className="space-y-1.5">
                  {(p.points || []).map((pt, j) => (
                    <li key={j} className="flex gap-2 text-[14px] leading-relaxed" style={{ color: ink }}>
                      <span style={{ color: blue, flexShrink: 0 }}>{j + 1}.</span>
                      <span>{pt}</span>
                    </li>
                  ))}
                </ul>
              </GroupRow>
              <GroupRow last>
                <FieldLabel>主要内容</FieldLabel>
                <p className="text-[14px] leading-relaxed" style={{ color: inkSecondary }}>
                  {p.detail}
                </p>
              </GroupRow>
            </GroupCard>
          );
        })}

      <GroupCard>
        <GroupRow last>
          <FieldLabel>每页的标题、观点、内容安排OK吗？</FieldLabel>
          <SegmentedControl
            options={[
              { value: "approved", label: "没问题，继续", color: green },
              { value: "fix", label: "有页要调整", color: orange, icon: AlertTriangle },
            ]}
            value={pagePlanStatus}
            onChange={setPagePlanStatus}
          />
          {pagePlanStatus === "fix" && (
            <div className="mt-3">
              <button onClick={onRegenerate} disabled={regenerating} className="flex items-center gap-1.5 text-[14px] font-medium" style={{ color: blue }}>
                {regenerating ? <Loader2 size={14} className="animate-spin" /> : <RotateCcw size={14} />}
                {regenerating ? "重新生成中…" : "回上一步改大纲后重新生成，或直接点此重试"}
              </button>
            </div>
          )}
        </GroupRow>
      </GroupCard>

      <div className="flex gap-2.5">
        <BackButton onClick={onBack} />
        <PrimaryButton disabled={pagePlanStatus !== "approved" || !allReady} onClick={onNext}>
          进入逐页细化 <ChevronRight size={16} />
        </PrimaryButton>
      </div>
    </div>
  );
}

// ---- Step 4 --------------------------------------------------------------

const PAGE_STATUS = {
  todo: { label: "待细化", color: grey },
  draft: { label: "草稿中", color: blue },
  fix: { label: "需修改", color: orange },
  done: { label: "已确认", color: green },
};

// ============================================================
// 导出层：HTML真导出 + pptx定稿复制（回对话生成真实文件）
// ============================================================

function escHTML(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function getCurVersion(p) {
  return p.versions?.length ? p.versions[p.activeIdx ?? p.versions.length - 1] : null;
}

function buildDeckHTML(pages, form) {
  const accent = (form.color || "").includes("橙") ? "#C2691F" : "#0071E3";
  const slides = pages
    .map((p, i) => {
      const cur = getCurVersion(p);
      const title = escHTML(cur?.title || p.title || p.point);
      const bullets = (cur?.bullets || p.points || []).map(escHTML);
      const layout = cur?.layout || "纯文字要点";
      let body = "";
      if (layout === "左图右文") {
        body = `<div class="row"><div class="img">图片占位</div><ul>${bullets.map((b) => `<li>${b}</li>`).join("")}</ul></div>`;
      } else if (layout === "三栏对比") {
        body = `<div class="cols">${bullets.slice(0, 3).map((b) => `<div class="col">${b}</div>`).join("")}</div>`;
      } else if (layout === "时间轴") {
        body = `<div class="tl">${bullets.map((b) => `<div class="tli"><span class="dot"></span><p>${b}</p></div>`).join("")}</div>`;
      } else {
        body = `<ul>${bullets.map((b) => `<li>${b}</li>`).join("")}</ul>`;
      }
      return `<section class="slide"><div class="tag">${i + 1} / ${pages.length} · ${escHTML(p.section)}</div><h2>${title}</h2>${body}</section>`;
    })
    .join("\n");
  return `<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escHTML(form.mainPoint || "PPT草稿")}</title><style>
:root{--accent:${accent}}
body{margin:0;background:#ECECEE;font-family:-apple-system,BlinkMacSystemFont,"PingFang SC",system-ui,sans-serif}
.slide{aspect-ratio:16/9;max-width:1080px;margin:28px auto;background:#fff;padding:52px 60px;box-sizing:border-box;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.08);overflow:hidden}
.tag{font-size:13px;color:#9a9a9f;margin-bottom:14px}
h2{font-size:30px;margin:0 0 26px;color:#1d1d1f;border-left:5px solid var(--accent);padding-left:16px;line-height:1.3}
ul{margin:0;padding-left:24px;font-size:19px;line-height:1.9;color:#3c3c40}
.row{display:flex;gap:36px;align-items:flex-start}
.img{flex:0 0 32%;aspect-ratio:1;background:#F5F5F7;border-radius:10px;display:flex;align-items:center;justify-content:center;color:#aeaeb2;font-size:14px}
.cols{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px}
.col{background:#F5F5F7;border-radius:10px;padding:22px;font-size:16px;line-height:1.7;color:#3c3c40}
.tl{display:flex;gap:22px;margin-top:12px}
.tli{flex:1}.dot{display:block;width:14px;height:14px;border-radius:50%;background:var(--accent);margin-bottom:12px}
.tli p{margin:0;font-size:15px;line-height:1.7;color:#3c3c40}
@media print{body{background:#fff}.slide{box-shadow:none;page-break-after:always;margin:0 auto}}
</style></head><body>${slides}</body></html>`;
}

function buildDeckText(pages, form) {
  const lines = [
    `请根据以下定稿生成一份.pptx文件（16:9）：`,
    ``,
    `全局：目标=${form.goal}；受众=${form.audience}；总体观点=${form.mainPoint}；配色=${form.color || "由你建议安全方案"}${form.colorRole === "accent" ? "（白底+点缀）" : form.colorRole === "primary" ? "（该色为主）" : ""}；视觉复杂度=${form.visualPriority === "design" ? "兼顾设计感" : "内容优先从简"}`,
    ``,
  ];
  pages.forEach((p, i) => {
    const cur = getCurVersion(p);
    lines.push(`第${i + 1}页（${p.section}）｜版式：${cur?.layout || "由你决定"}`);
    lines.push(`标题：${cur?.title || p.title || p.point}`);
    (cur?.bullets || p.points || []).forEach((b) => lines.push(`- ${b}`));
    lines.push(``);
  });
  lines.push(`注意：文字内容已人工审定，请勿改写；【需补数据】占位保留原样。`);
  return lines.join("\n");
}

function downloadFile(name, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function exportPptx(pages, form) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "W16x9", width: 13.33, height: 7.5 });
  pptx.layout = "W16x9";
  const accent = (form.color || "").includes("橙") ? "C2691F" : "0071E3";

  pages.forEach((p, i) => {
    const cur = getCurVersion(p);
    const title = cur?.title || p.title || p.point || "";
    const bullets = cur?.bullets || p.points || [];
    const layout = cur?.layout || "纯文字要点";
    const slide = pptx.addSlide();

    slide.addText(`${i + 1} / ${pages.length} · ${p.section || ""}`, { x: 0.6, y: 0.35, w: 12, h: 0.4, fontSize: 11, color: "9A9A9F" });
    slide.addShape("rect", { x: 0.6, y: 0.85, w: 0.09, h: 0.75, fill: { color: accent } });
    slide.addText(title, { x: 0.85, y: 0.8, w: 11.8, h: 0.9, fontSize: 26, bold: true, color: "1D1D1F" });

    if (layout === "左图右文") {
      slide.addShape("roundRect", { x: 0.6, y: 2.0, w: 4.2, h: 4.6, fill: { color: "F5F5F7" }, rectRadius: 0.08 });
      slide.addText("图片占位", { x: 0.6, y: 2.0, w: 4.2, h: 4.6, align: "center", fontSize: 12, color: "AEAEB2" });
      slide.addText(
        bullets.map((b) => ({ text: b, options: { bullet: { code: "2022" }, breakLine: true } })),
        { x: 5.2, y: 2.0, w: 7.4, h: 4.6, fontSize: 16, color: "3C3C40", lineSpacingMultiple: 1.4, valign: "top" }
      );
    } else if (layout === "三栏对比") {
      bullets.slice(0, 3).forEach((b, j) => {
        const x = 0.6 + j * 4.25;
        slide.addShape("roundRect", { x, y: 2.0, w: 3.95, h: 4.4, fill: { color: "F5F5F7" }, rectRadius: 0.08 });
        slide.addText(b, { x: x + 0.25, y: 2.25, w: 3.45, h: 3.9, fontSize: 14, color: "3C3C40", valign: "top", lineSpacingMultiple: 1.3 });
      });
    } else if (layout === "时间轴") {
      const n = Math.max(bullets.length, 1);
      const w = 12.1 / n;
      bullets.forEach((b, j) => {
        const x = 0.6 + j * w;
        slide.addShape("ellipse", { x, y: 2.4, w: 0.22, h: 0.22, fill: { color: accent } });
        if (j < n - 1) slide.addShape("line", { x: x + 0.3, y: 2.51, w: w - 0.4, h: 0, line: { color: "E5E5EA", width: 2 } });
        slide.addText(b, { x, y: 2.8, w: w - 0.3, h: 3.4, fontSize: 13, color: "3C3C40", valign: "top", lineSpacingMultiple: 1.3 });
      });
    } else {
      slide.addText(
        bullets.map((b) => ({ text: b, options: { bullet: { code: "2022" }, breakLine: true } })),
        { x: 0.85, y: 2.0, w: 11.8, h: 4.8, fontSize: 17, color: "3C3C40", lineSpacingMultiple: 1.5, valign: "top" }
      );
    }
  });

  await pptx.writeFile({ fileName: "ppt-draft.pptx" });
}

function ExportBar({ pages, form }) {
  const [pptxLoading, setPptxLoading] = useState(false);
  const [pptxError, setPptxError] = useState(null);
  const ready = pages.some((p) => getCurVersion(p));

  const handleHTML = () => downloadFile("ppt-draft.html", buildDeckHTML(pages, form), "text/html;charset=utf-8");
  const handlePptx = async () => {
    setPptxLoading(true);
    setPptxError(null);
    try {
      await exportPptx(pages, form);
    } catch (err) {
      setPptxError(err.message); // 降级：显示定稿文本供手动复制回对话生成
    } finally {
      setPptxLoading(false);
    }
  };

  return (
    <GroupCard>
      <GroupRow last>
        <FieldLabel>导出</FieldLabel>
        <div className="flex gap-2.5 flex-wrap">
          <button
            onClick={handlePptx}
            disabled={!ready || pptxLoading}
            className="flex items-center gap-1.5 text-[14px] font-medium px-4 py-2 rounded-full"
            style={{ background: ready ? blue : `${grey}20`, color: ready ? "#FFF" : inkTertiary }}
          >
            {pptxLoading ? <Loader2 size={14} className="animate-spin" /> : null}
            {pptxLoading ? "生成中…" : "导出 .pptx"}
          </button>
          <button
            onClick={handleHTML}
            disabled={!ready}
            className="flex items-center gap-1.5 text-[14px] font-medium px-4 py-2 rounded-full"
            style={{ background: ready ? `${blue}10` : `${grey}15`, color: ready ? blue : inkTertiary }}
          >
            导出 HTML（可打开/打印）
          </button>
        </div>
        {!ready && (
          <p className="text-[12px] mt-2" style={{ color: inkTertiary }}>
            至少一页生成版式后才能导出
          </p>
        )}
        {pptxError && (
          <div className="mt-2">
            <p className="text-[12px] mb-1.5" style={{ color: orangeDeep }}>
              pptx生成失败：{pptxError}。降级方案：复制下面的定稿粘贴给Claude对话生成真实.pptx
            </p>
            <textarea readOnly rows={8} className="w-full rounded-[10px] p-3 text-[12px]" style={{ background: bg, border: `1px solid ${divider}`, color: ink }} value={buildDeckText(pages, form)} onFocus={(e) => e.target.select()} />
          </div>
        )}
      </GroupRow>
    </GroupCard>
  );
}

function StepPages({ pages, setPages, form, onBack, onNext }) {
  const [openId, setOpenId] = useState(null);
  const [loadingIds, setLoadingIds] = useState({});
  const [sketches, setSketches] = useState({}); // 草图只存内存，刷新后需重传（避免撑爆存储）
  const [instructions, setInstructions] = useState({}); // 每页的修改意见输入
  const [editingId, setEditingId] = useState(null); // 正在手动编辑的页
  const [editDraft, setEditDraft] = useState({ title: "", bulletsText: "" });
  const [zoomId, setZoomId] = useState(null); // 放大查看的页
  const allDone = pages.length > 0 && pages.every((p) => p.status === "done");
  const updatePage = (id, patch) => setPages(pages.map((p) => (p.id === id ? { ...p, ...patch } : p)));

  // 版本管理：每页一个版本栈，activeIdx指向当前展示的版本
  const getCur = (p) => (p.versions?.length ? p.versions[p.activeIdx ?? p.versions.length - 1] : null);
  const pushVersion = (page, data, label) => {
    const versions = [...(page.versions || []), { ...data, label, ts: Date.now() }].slice(-8); // 最多留8版
    updatePage(page.id, { versions, activeIdx: versions.length - 1, status: "draft", aiError: null });
  };

  const handleSketchUpload = (pageId, file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const [meta, data] = String(reader.result).split(",");
      const match = meta.match(/data:(.*);base64/);
      if (match) setSketches((s) => ({ ...s, [pageId]: { media_type: match[1], data } }));
    };
    reader.readAsDataURL(file);
  };

  // AI推荐版式：文字沿用审定内容，AI只选layout
  const handleAIAssist = async (page) => {
    setLoadingIds((s) => ({ ...s, [page.id]: true }));
    try {
      const sketch = sketches[page.id] || null;
      const result = await toolLayoutPage(page, form, sketch);
      pushVersion(page, { layout: result.layout, title: page.title || page.point, bullets: page.points || [], reason: result.reason }, sketch ? "AI版式(草图)" : "AI版式");
    } catch (err) {
      updatePage(page.id, { aiError: err.message });
    } finally {
      setLoadingIds((s) => ({ ...s, [page.id]: false }));
    }
  };

  // AI按修改意见改：最小修改，产生新版本
  const handleInstructionEdit = async (page) => {
    const instruction = (instructions[page.id] || "").trim();
    const cur = getCur(page);
    if (!instruction || !cur) return;
    setLoadingIds((s) => ({ ...s, [page.id]: true }));
    try {
      const result = await toolEditPage(cur, instruction, form);
      pushVersion(page, { layout: result.layout, title: result.title, bullets: result.bullets, reason: `按意见修改：${instruction.slice(0, 30)}` }, "AI修改");
      setInstructions((s) => ({ ...s, [page.id]: "" }));
    } catch (err) {
      updatePage(page.id, { aiError: err.message });
    } finally {
      setLoadingIds((s) => ({ ...s, [page.id]: false }));
    }
  };

  // 手动编辑：进入/保存
  const startManualEdit = (page) => {
    const cur = getCur(page);
    if (!cur) return;
    setEditingId(page.id);
    setEditDraft({ title: cur.title, bulletsText: (cur.bullets || []).join("\n") });
  };
  const saveManualEdit = (page) => {
    const cur = getCur(page);
    const bullets = editDraft.bulletsText.split("\n").map((s) => s.trim()).filter(Boolean);
    pushVersion(page, { layout: cur.layout, title: editDraft.title.trim(), bullets, reason: "人工直接修改" }, "手动");
    setEditingId(null);
  };

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-3">
        {pages.map((p) => {
          const meta = PAGE_STATUS[p.status];
          const isOpen = openId === p.id;
          const isLoading = loadingIds[p.id];
          const cur = getCur(p);
          const isEditing = editingId === p.id;
          return (
            <div
              key={p.id}
              className="rounded-[14px] p-4 cursor-pointer transition-all"
              style={{ background: cardBg, boxShadow: isOpen ? `0 0 0 2px ${meta.color}` : "0 1px 2px rgba(0,0,0,0.04)", gridColumn: isOpen ? "1 / -1" : "auto" }}
              onClick={() => setOpenId(isOpen ? null : p.id)}
            >
              <div className="flex items-center justify-between mb-2.5">
                <span className="text-[13px] font-medium" style={{ color: inkSecondary }}>
                  {p.section}
                </span>
                <span className="flex items-center gap-1.5 text-[12px] font-medium" style={{ color: inkSecondary }}>
                  <span className="w-[7px] h-[7px] rounded-full" style={{ background: meta.color }} />
                  {meta.label}
                </span>
              </div>
              <p className="text-[15px] leading-snug line-clamp-2" style={{ color: ink }}>
                {cur?.title || p.title || p.point}
              </p>

              {isOpen && (
                <div className="mt-4 space-y-3.5" onClick={(e) => e.stopPropagation()}>
                  {cur && !isEditing && (
                    <div>
                      <PageRenderer data={cur} />
                      <div className="flex items-center justify-between mt-2">
                        <div className="flex items-center gap-1.5 flex-wrap">
                          {(p.versions || []).map((v, i) => {
                            const active = i === (p.activeIdx ?? p.versions.length - 1);
                            return (
                              <button
                                key={v.ts}
                                onClick={() => updatePage(p.id, { activeIdx: i })}
                                className="text-[11px] font-medium px-2 py-0.5 rounded-full"
                                style={{ background: active ? blue : `${grey}20`, color: active ? "#FFF" : inkSecondary }}
                                title={v.label}
                              >
                                v{i + 1}
                              </button>
                            );
                          })}
                          {p.versions?.length > 1 && (
                            <span className="text-[11px]" style={{ color: inkTertiary }}>
                              {cur.label}
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-3">
                          <button onClick={() => setZoomId(p.id)} className="text-[12px] font-medium" style={{ color: blue }}>
                            放大查看
                          </button>
                          <button onClick={() => startManualEdit(p)} className="text-[12px] font-medium" style={{ color: blue }}>
                            直接改文字
                          </button>
                        </div>
                      </div>
                      {cur.reason && (
                        <p className="text-[12px] leading-relaxed mt-1.5" style={{ color: inkTertiary }}>
                          {cur.reason}
                        </p>
                      )}
                    </div>
                  )}

                  {isEditing && (
                    <div className="rounded-[10px] px-3 py-2.5 space-y-2.5" style={{ background: bg }}>
                      <div>
                        <FieldLabel>标题</FieldLabel>
                        <input value={editDraft.title} onChange={(e) => setEditDraft({ ...editDraft, title: e.target.value })} style={inputStyle} />
                      </div>
                      <div>
                        <FieldLabel>要点（一行一条）</FieldLabel>
                        <textarea
                          value={editDraft.bulletsText}
                          onChange={(e) => setEditDraft({ ...editDraft, bulletsText: e.target.value })}
                          rows={4}
                          className="resize-none"
                          style={inputStyle}
                        />
                      </div>
                      <div className="flex gap-2">
                        <button onClick={() => saveManualEdit(p)} className="text-[13px] font-medium px-3.5 py-1.5 rounded-full" style={{ background: green, color: "#FFF" }}>
                          保存为新版本
                        </button>
                        <button onClick={() => setEditingId(null)} className="text-[13px] font-medium px-3.5 py-1.5 rounded-full" style={{ background: `${grey}20`, color: inkSecondary }}>
                          取消
                        </button>
                      </div>
                    </div>
                  )}

                  {cur && !isEditing && (
                    <div className="rounded-[10px] px-3 py-2.5" style={{ background: bg }}>
                      <FieldLabel>修改意见（AI按最小修改原则执行，只改你提到的部分）</FieldLabel>
                      <textarea
                        value={instructions[p.id] || ""}
                        onChange={(e) => setInstructions((s) => ({ ...s, [p.id]: e.target.value }))}
                        rows={2}
                        placeholder="例：第二条要点太空，补上试点团队的名字；标题里的'验证'换成'兑现'"
                        className="resize-none"
                        style={inputStyle}
                      />
                      <button
                        onClick={() => handleInstructionEdit(p)}
                        disabled={isLoading || !(instructions[p.id] || "").trim()}
                        className="mt-2 flex items-center gap-1.5 text-[13px] font-medium px-3.5 py-1.5 rounded-full"
                        style={{ background: (instructions[p.id] || "").trim() ? blue : `${grey}20`, color: (instructions[p.id] || "").trim() ? "#FFF" : inkTertiary }}
                      >
                        {isLoading ? <Loader2 size={13} className="animate-spin" /> : <Pencil size={13} />}
                        让AI按意见修改（生成新版本）
                      </button>
                    </div>
                  )}

                  {!cur && (
                    <>
                      <div className="rounded-[10px] px-3 py-2.5" style={{ background: bg }}>
                        <FieldLabel>版式直觉补充（可选，文字描述你想要的摆法）</FieldLabel>
                        <textarea
                          value={p.content}
                          onChange={(e) => updatePage(p.id, { content: e.target.value })}
                          rows={2}
                          placeholder="例：左边放对比图，右边两条结论"
                          className="resize-none"
                          style={inputStyle}
                        />
                      </div>

                      <div className="rounded-[10px] px-3 py-2.5 flex items-center justify-between gap-3" style={{ background: bg }}>
                        <div className="min-w-0">
                          <FieldLabel optional>手绘草图</FieldLabel>
                          <p className="text-[12px]" style={{ color: sketches[p.id] ? greenDeep : inkTertiary }}>
                            {sketches[p.id] ? "已上传 ✓（只在本次会话有效，刷新后需重传）" : "拍照上传你的鬼图，AI会按草图意图选版式"}
                          </p>
                        </div>
                        <label className="shrink-0 text-[13px] font-medium px-3 py-1.5 rounded-full cursor-pointer" style={{ background: `${blue}10`, color: blue }}>
                          {sketches[p.id] ? "重新上传" : "选择图片"}
                          <input type="file" accept="image/*" className="hidden" onChange={(e) => handleSketchUpload(p.id, e.target.files?.[0])} />
                        </label>
                      </div>
                    </>
                  )}

                  <button
                    onClick={() => handleAIAssist(p)}
                    disabled={isLoading}
                    className="w-full flex items-center justify-center gap-2 py-2.5 rounded-[10px] text-[14px] font-medium"
                    style={{ background: `${blue}10`, color: blue }}
                  >
                    {isLoading ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                    {isLoading ? "AI分析中…" : cur ? "换个版式再推荐一次（生成新版本）" : sketches[p.id] ? "让AI按草图推荐版式" : "让AI推荐版式"}
                  </button>

                  {p.aiError && (
                    <p className="text-[12px]" style={{ color: orangeDeep }}>
                      生成失败：{p.aiError}
                    </p>
                  )}

                  <SegmentedControl
                    options={[
                      { value: "done", label: "满意，下一页", color: green },
                      { value: "fix", label: "内容/版式要改", color: orange, icon: Pencil },
                    ]}
                    value={p.status === "done" || p.status === "fix" ? p.status : null}
                    onChange={(v) => updatePage(p.id, { status: v })}
                  />
                </div>
              )}
            </div>
          );
        })}
      </div>

      {zoomId && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-6" style={{ background: "rgba(0,0,0,0.55)" }} onClick={() => setZoomId(null)}>
          <div className="w-full" style={{ maxWidth: 960 }} onClick={(e) => e.stopPropagation()}>
            <PageRenderer data={getCur(pages.find((p) => p.id === zoomId))} large />
            <p className="text-center text-[13px] mt-3" style={{ color: "#FFFFFFAA" }}>
              点击空白处关闭
            </p>
          </div>
        </div>
      )}

      <ExportBar pages={pages} form={form} />

      <div className="flex gap-2.5">
        <BackButton onClick={onBack} />
        <PrimaryButton disabled={!allDone} onClick={onNext}>
          全部确认，进入Agent评审 <ChevronRight size={16} />
        </PrimaryButton>
      </div>
    </div>
  );
}

// ---- Step 4 ----------------------------------------------------------------

// 工具⑤：评审agent（第5步）——新调用+苛刻评审角色，冷眼看全稿
async function toolReviewDeck(pages, form) {
  const deckText = pages
    .map((p, i) => {
      const cur = p.versions?.length ? p.versions[p.activeIdx ?? p.versions.length - 1] : null;
      return `第${i + 1}页(${p.section}) 版式:${cur?.layout || "未定"} 标题:${cur?.title || p.title} 要点:${(cur?.bullets || p.points || []).join("；")}`;
    })
    .join("\n");
  const prompt = `你是一位苛刻的PPT评审，独立审阅这份已完成的稿子。你没有参与制作，不需要给制作者留面子，只对最终效果负责。
从这几个角度挑问题：观点是否空洞无锚点、跨页术语是否一致、逻辑是否有断层、对该受众的说服力缺口、【需补数据】占位是否遗留。
只返回紧凑JSON对象，不要任何多余文字、不要markdown代码块：
- overall：字符串数组，最多3条全局问题（每条30字内）
- pages：数组，只列出有问题的页，每项 {idx:页码数字, issues:[{problem:问题(25字内), fix:具体修改建议(30字内)}]}，每页最多2条
没有问题的维度不要硬凑。

受众：${form.audience}
总体观点：${form.mainPoint}
全稿：
${deckText}`;
  return callClaude(prompt);
}

function StepReview({ pages, setPages, form, review, setReview, onBack }) {
  const [reviewing, setReviewing] = useState(false);
  const [error, setError] = useState(null);
  const [fixingKey, setFixingKey] = useState(null); // "pageIdx-issueIdx"

  const runReview = async () => {
    setReviewing(true);
    setError(null);
    try {
      const result = await toolReviewDeck(pages, form);
      setReview(result);
    } catch (err) {
      setError(err.message);
    } finally {
      setReviewing(false);
    }
  };

  // 采纳意见：把fix作为修改指令发给工具④，在对应页生成新版本
  const adoptFix = async (pageIdx, issueIdx, issue) => {
    const page = pages[pageIdx - 1];
    if (!page) return;
    const cur = page.versions?.length ? page.versions[page.activeIdx ?? page.versions.length - 1] : null;
    if (!cur) return;
    const key = `${pageIdx}-${issueIdx}`;
    setFixingKey(key);
    try {
      const result = await toolEditPage(cur, `${issue.problem}。修改要求：${issue.fix}`, form);
      const versions = [...(page.versions || []), { layout: result.layout, title: result.title, bullets: result.bullets, reason: `评审采纳：${issue.fix}`, label: "评审修改", ts: Date.now() }].slice(-8);
      setPages(pages.map((p) => (p.id === page.id ? { ...p, versions, activeIdx: versions.length - 1 } : p)));
    } catch (err) {
      setError(err.message);
    } finally {
      setFixingKey(null);
    }
  };

  return (
    <div className="space-y-6">
      <div className="rounded-[14px] px-4 py-3 flex items-center gap-2.5 text-[13px]" style={{ background: `${blue}10`, color: blue }}>
        <Sparkles size={14} style={{ flexShrink: 0 }} />
        第4步的稿子已是"出版稿"。这里由评审agent（独立调用，苛刻评审角色）冷眼审全稿，采纳的意见会直接在对应页生成新版本。
      </div>

      {error && (
        <div className="rounded-[14px] px-4 py-3 text-[13px]" style={{ background: `${orange}12`, color: orangeDeep }}>
          {error}
        </div>
      )}

      {!review && (
        <GroupCard>
          <GroupRow last>
            <div className="py-4 text-center">
              <button
                onClick={runReview}
                disabled={reviewing}
                className="inline-flex items-center gap-2 text-[15px] font-semibold px-6 py-3 rounded-full"
                style={{ background: blue, color: "#FFF" }}
              >
                {reviewing ? <Loader2 size={15} className="animate-spin" /> : <Sparkles size={15} />}
                {reviewing ? "评审agent审阅中…" : "开始评审"}
              </button>
            </div>
          </GroupRow>
        </GroupCard>
      )}

      {review && (
        <>
          {(review.overall || []).length > 0 && (
            <GroupCard>
              <div className="px-4 pt-3.5 pb-1">
                <FieldLabel>全局问题</FieldLabel>
              </div>
              {review.overall.map((o, i) => (
                <GroupRow key={i} last={i === review.overall.length - 1}>
                  <p className="text-[14px] leading-relaxed" style={{ color: ink }}>
                    {o}
                  </p>
                </GroupRow>
              ))}
            </GroupCard>
          )}

          {(review.pages || []).map((rp) => (
            <GroupCard key={rp.idx}>
              <div className="px-4 pt-3.5 pb-1">
                <FieldLabel>
                  第{rp.idx}页 · {pages[rp.idx - 1]?.section || ""}
                </FieldLabel>
              </div>
              {(rp.issues || []).map((issue, j) => {
                const key = `${rp.idx}-${j}`;
                return (
                  <GroupRow key={j} last={j === rp.issues.length - 1}>
                    <p className="text-[14px] mb-1" style={{ color: ink }}>
                      {issue.problem}
                    </p>
                    <p className="text-[13px] mb-2" style={{ color: inkSecondary }}>
                      建议：{issue.fix}
                    </p>
                    <button
                      onClick={() => adoptFix(rp.idx, j, issue)}
                      disabled={fixingKey === key}
                      className="inline-flex items-center gap-1.5 text-[13px] font-medium px-3 py-1.5 rounded-full"
                      style={{ background: `${green}14`, color: greenDeep }}
                    >
                      {fixingKey === key ? <Loader2 size={13} className="animate-spin" /> : <Check size={13} />}
                      采纳并让AI修改（新版本）
                    </button>
                  </GroupRow>
                );
              })}
            </GroupCard>
          ))}

          {(review.pages || []).length === 0 && (review.overall || []).length === 0 && (
            <GroupCard>
              <GroupRow last>
                <div className="flex items-center gap-2.5 font-medium text-[15px]" style={{ color: greenDeep }}>
                  <span className="w-[22px] h-[22px] rounded-full flex items-center justify-center shrink-0" style={{ background: green }}>
                    <Check size={13} style={{ color: "#FFFFFF" }} />
                  </span>
                  评审agent没有发现问题——可以把定稿带回对话生成.pptx了
                </div>
              </GroupRow>
            </GroupCard>
          )}

          <button onClick={runReview} disabled={reviewing} className="flex items-center gap-1.5 text-[14px] font-medium" style={{ color: blue }}>
            {reviewing ? <Loader2 size={14} className="animate-spin" /> : <RotateCcw size={14} />}
            {reviewing ? "重新评审中…" : "修改后重新评审一轮"}
          </button>
        </>
      )}

      <ExportBar pages={pages} form={form} />

      <div className="flex gap-2.5">
        <BackButton onClick={onBack} />
      </div>
    </div>
  );
}

// ============================================================
// ORCHESTRATION 层：状态机 + MEMORY 层：storage持久化
// ============================================================

export default function PPTAgent() {
  const [state, setState] = useState(defaultState);
  const [loaded, setLoaded] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [outlineError, setOutlineError] = useState(null);
  const [planGenerating, setPlanGenerating] = useState(false);
  const [planError, setPlanError] = useState(null);
  const [planProgress, setPlanProgress] = useState({ done: 0, total: 0 });
  const saveTimer = useRef(null);

  // MEMORY：启动时从 localStorage 恢复上次进度
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setState(JSON.parse(raw));
    } catch (e) {
      // 第一次使用，没有存档或数据损坏，正常
    } finally {
      setLoaded(true);
    }
  }, []);

  // MEMORY：任何状态变化250ms后自动存档到 localStorage
  useEffect(() => {
    if (!loaded) return;
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      } catch (e) {
        console.error("存档失败", e);
      }
    }, 250);
  }, [state, loaded]);

  const setForm = (form) => setState({ ...state, form });
  const setPages = (pages) => setState({ ...state, pages });
  const setReview = (review) => setState({ ...state, review });
  const goTo = (step) => setState({ ...state, step });

  // ORCHESTRATION：第1步→第2步，触发真实AI调用
  const handleGenerateOutline = async () => {
    setGenerating(true);
    setOutlineError(null);
    try {
      const outline = await toolGenerateOutline(state.form);
      setState({ ...state, step: 2, outline, outlineStatus: null });
    } catch (err) {
      setOutlineError(err.message);
      setState({ ...state, step: 2, outline: null, outlineStatus: null });
    } finally {
      setGenerating(false);
    }
  };

  // ORCHESTRATION：第2步确认→第3步，逐页生成排布（每页独立调用，避免整批截断）
  const handleGeneratePagePlan = async () => {
    const outline = state.outline || [];
    setPlanGenerating(true);
    setPlanError(null);
    setPlanProgress({ done: 0, total: outline.length });
    // 先进入第3步并放好占位骨架，逐页填充，生成一页立刻能看一页
    setState((s) => ({ ...s, step: 3, pagePlan: outline.map((o) => ({ section: o.section, pending: true })), pagePlanStatus: null }));
    let failCount = 0;
    for (let i = 0; i < outline.length; i++) {
      try {
        const page = await toolGeneratePagePlanForPage(outline[i], outline, state.form, i);
        setState((s) => ({ ...s, pagePlan: s.pagePlan.map((p, j) => (j === i ? { ...page, pending: false } : p)) }));
      } catch (err) {
        failCount++;
        setState((s) => ({ ...s, pagePlan: s.pagePlan.map((p, j) => (j === i ? { section: outline[i].section, pending: false, failed: true, error: err.message } : p)) }));
      }
      setPlanProgress({ done: i + 1, total: outline.length });
    }
    if (failCount > 0) setPlanError(`${failCount}页生成失败，点对应页的"重试这一页"即可，无需整批重来`);
    setPlanGenerating(false);
  };

  // 单页重试：只重新生成失败的那一页
  const handleRetryPlanPage = async (i) => {
    const outline = state.outline || [];
    if (!outline[i]) return;
    setState((s) => ({ ...s, pagePlan: s.pagePlan.map((p, j) => (j === i ? { section: outline[i].section, pending: true } : p)) }));
    try {
      const page = await toolGeneratePagePlanForPage(outline[i], outline, state.form, i);
      setState((s) => ({ ...s, pagePlan: s.pagePlan.map((p, j) => (j === i ? { ...page, pending: false } : p)) }));
    } catch (err) {
      setState((s) => ({ ...s, pagePlan: s.pagePlan.map((p, j) => (j === i ? { section: outline[i].section, pending: false, failed: true, error: err.message } : p)) }));
    }
  };

  // ORCHESTRATION：第3步review通过→第4步，把已审定排布转成页面卡片
  const handleEnterPages = () => {
    const pages = (state.pagePlan || []).map((p, i) => ({
      id: `p${i + 1}`,
      section: p.section,
      title: p.title,
      points: p.points,
      detail: p.detail,
      point: p.title, // 兼容字段
      content: "",
      rendered: null,
      status: "todo",
    }));
    setState({ ...state, step: 4, pages });
  };

  if (!loaded) {
    return (
      <div className="flex items-center justify-center h-64" style={{ background: bg, color: inkSecondary, fontFamily: FONT }}>
        <span className="text-[15px]">加载中…</span>
      </div>
    );
  }

  return (
    <div className="w-full min-h-[600px] p-6 sm:p-10" style={{ background: bg, fontFamily: FONT }}>
      <div className="max-w-2xl mx-auto">
        <h1 className="text-[28px] font-semibold mb-8" style={{ color: ink, letterSpacing: "-0.015em" }}>
          PPT 工作流
        </h1>

        <Stepper step={state.step} />

        {state.step === 1 && <StepIntake form={state.form} setForm={setForm} onGenerate={handleGenerateOutline} generating={generating} />}

        {state.step === 2 && (
          <StepOutline
            outline={state.outline}
            error={outlineError}
            outlineStatus={state.outlineStatus}
            setOutlineStatus={(outlineStatus) => setState({ ...state, outlineStatus })}
            onBack={() => goTo(1)}
            onRegenerate={handleGenerateOutline}
            regenerating={generating}
            onNext={handleGeneratePagePlan}
            nextLoading={planGenerating}
          />
        )}

        {state.step === 3 && (
          <StepPagePlan
            pagePlan={state.pagePlan}
            error={planError}
            pagePlanStatus={state.pagePlanStatus}
            setPagePlanStatus={(pagePlanStatus) => setState({ ...state, pagePlanStatus })}
            onBack={() => goTo(2)}
            onRegenerate={handleGeneratePagePlan}
            onRetryPage={handleRetryPlanPage}
            regenerating={planGenerating}
            progress={planProgress}
            onNext={handleEnterPages}
          />
        )}

        {state.step === 4 && <StepPages pages={state.pages} setPages={setPages} form={state.form} onBack={() => goTo(3)} onNext={() => goTo(5)} />}

        {state.step === 5 && (
          <StepReview
            pages={state.pages}
            setPages={setPages}
            form={state.form}
            review={state.review}
            setReview={setReview}
            onBack={() => goTo(4)}
          />
        )}
      </div>
    </div>
  );
}
