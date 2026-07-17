// Sequentially tests every candidate Cloudflare Workers AI model against the same
// natural-language shift-notes prompt used by parseNaturalLanguageEntries() in ai.js,
// so you can compare latency / JSON validity / reasoning behavior side by side.
//
// Usage:
//   1. Fill in CLOUDFLARE_ACCOUNT_ID and CLOUDFLARE_API_TOKEN in .env
//   2. node test-cloudflare-models.mjs

import { readFileSync } from 'node:fs';

function loadEnv(path) {
  const env = {};
  for (const line of readFileSync(path, 'utf8').split('\n')) {
    const m = line.match(/^\s*([\w.-]+)\s*=\s*(.*?)\s*$/);
    if (m) env[m[1]] = m[2];
  }
  return env;
}

const env = loadEnv(new URL('.env', import.meta.url));
const ACCOUNT_ID = env.CLOUDFLARE_ACCOUNT_ID;
const API_TOKEN = env.CLOUDFLARE_API_TOKEN;

if (!ACCOUNT_ID || !API_TOKEN) {
  console.error('Missing CLOUDFLARE_ACCOUNT_ID or CLOUDFLARE_API_TOKEN in .env');
  process.exit(1);
}

const URL_ = `https://api.cloudflare.com/client/v4/accounts/${ACCOUNT_ID}/ai/v1/chat/completions`;

// Same catalog as the comment block above CLOUDFLARE_MODEL in ai.js.
const MODELS = [
  '@cf/meta/llama-3.1-8b-instruct',
  '@cf/meta/llama-3.1-8b-instruct-fp8',
  '@cf/meta/llama-3.2-3b-instruct',
  '@cf/meta/llama-3.3-70b-instruct-fp8-fast',
  '@cf/mistralai/mistral-small-3.1-24b-instruct',
  '@cf/qwen/qwen2.5-coder-32b-instruct',
  '@cf/meta/llama-4-scout-17b-16e-instruct',
  '@cf/google/gemma-4-26b-a4b-it',
  '@cf/openai/gpt-oss-20b',
  '@cf/openai/gpt-oss-120b',
  '@cf/qwen/qwen3-30b-a3b-fp8',
  '@cf/deepseek-ai/deepseek-r1-distill-qwen-32b',
  '@cf/qwen/qwq-32b'
];

const INPUT_TEXT = `13.7.26
1.Rasheedh 8am-4pm
2.Ravi 7am-3pm
3.senthil 7am- 3pm
4.suganathan 7am-3pm
5.velan 7am-3pm
6.sundar 7am-3pm
7.Nilabdeen 7am-6pm`;

const EMPLOYEES = ['Rasheedh', 'Ravi', 'Senthil', 'Suganathan', 'Velan', 'Sundar', 'Nilabdeen'];

// Mirrors the prompt built in parseNaturalLanguageEntries() (ai.js).
const SYSTEM_MSG =
  'You are a precise JSON-only utility assistant. You must output raw, valid JSON only. ' +
  'Do not wrap your response in markdown code blocks (fences) and do not include conversational text or explanations.';

const USER_MSG =
  'You convert informal timesheet notes into structured JSON for a construction/industrial contractor in Qatar.\n\n' +
  'CONTEXT:\n' +
  '- Today is 2026-07-13 (Monday). Resolve relative dates like "today", "yesterday", "last Monday" against this.\n' +
  '- The target sheet "JULY 2026" covers July 2026. When the text gives only a day number, resolve it into this month and year.\n' +
  '- Known employees on this sheet (use these EXACT names when you are confident of a match):\n  ' + JSON.stringify(EMPLOYEES) + '\n' +
  '- Known job orders (prefer these exact values when the text clearly refers to one):\n  []\n\n' +
  '- Expect date formats like "DD.MM.YY", "DD.MM.YYYY", "DD/MM/YYYY". A date line at the top (e.g. "14.7.26") sets the date for ALL following entries unless another date appears.\n' +
  'INPUT NOTES (may be WhatsApp-style shorthand, mixed English/Tamil, abbreviations, tables, or messy lists):\n' +
  '"""\n' + INPUT_TEXT + '\n"""\n\n' +
  'RULES:\n' +
  '1. Output times in strict 24-hour "HH:mm". Interpret common shorthand: "8 to 5" = 08:00 to 17:00, "7-7" = 07:00 to 19:00, "8pm to 4am" = 20:00 to 04:00, "8am-4pm" = 08:00 to 16:00, "7am-3pm" = 07:00 to 15:00, "7am-5pm" = 07:00 to 17:00.\n' +
  '2. Dates in strict "YYYY-MM-DD".\n' +
  '3. Group people who share the SAME date, start, end and job order into ONE entry with multiple names in "employees".\n' +
  '4. For employee names: if a name in the notes clearly matches a known employee (ignoring case, spacing, partial spelling, or minor typos), return the EXACT known name. If you cannot confidently match, return the name exactly as written and set confidence to "low".\n' +
  '5. If the job order is not stated and cannot be inferred, use an empty string "".\n' +
  '6. Never invent employees, dates, or times that are not in the notes. If something essential is missing or ambiguous, skip that entry and add an explanation to "warnings".\n' +
  '7. When you guessed anything (AM vs PM, which month, which person), say so in that entry\'s "note" and lower its confidence.\n' +
  '8. Numbered list format like "1.Rasheedh 8am-4pm" means each line is one person with their shift. The number is just a list marker, not part of the name.\n\n' +
  'Return ONLY a valid JSON object matching this structure:\n' +
  '{\n' +
  '  "entries": [\n' +
  '    { "date": "YYYY-MM-DD", "employees": ["NAME"], "startTime": "HH:mm", "endTime": "HH:mm", "jobOrder": "", "confidence": "high", "note": "" }\n' +
  '  ],\n' +
  '  "warnings": []\n' +
  '}';

function extractMessageText(content) {
  if (!content) return '';
  if (typeof content === 'string') return content.trim();
  if (Array.isArray(content)) {
    return content.map((p) => (p && typeof p.text === 'string') ? p.text : '').join('').trim();
  }
  return '';
}

async function testModel(model) {
  const payload = {
    model,
    messages: [
      { role: 'system', content: SYSTEM_MSG },
      { role: 'user', content: USER_MSG }
    ],
    temperature: 0.2,
    max_tokens: 6000,
    reasoning_effort: 'low',
    chat_template_kwargs: { thinking: false },
    response_format: { type: 'json_object' }
  };

  const started = Date.now();
  let res, text;
  try {
    res = await fetch(URL_, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${API_TOKEN}`,
        'cf-aig-gateway-id': 'default'
      },
      body: JSON.stringify(payload)
    });
    text = await res.text();
  } catch (networkErr) {
    return { model, ok: false, ms: Date.now() - started, error: 'Network error: ' + networkErr.message };
  }
  const ms = Date.now() - started;

  if (!res.ok) {
    let detail = text;
    try {
      const parsed = JSON.parse(text);
      detail = parsed.errors?.[0]?.message || parsed.error?.message || text;
    } catch { /* not JSON */ }
    return { model, ok: false, ms, error: `HTTP ${res.status}: ${detail.slice(0, 300)}` };
  }

  let resJson;
  try {
    resJson = JSON.parse(text);
  } catch (e) {
    return { model, ok: false, ms, error: 'Unparsable response body: ' + e.message };
  }

  const choice = resJson.choices?.[0];
  const rawText = extractMessageText(choice?.message?.content);
  const usage = resJson.usage;
  const finishReason = choice?.finish_reason;

  if (!rawText) {
    return { model, ok: false, ms, error: `Empty content (finish_reason: ${finishReason})`, usage, finishReason };
  }

  const cleaned = rawText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/, '');
  let data;
  try {
    data = JSON.parse(cleaned);
  } catch (e) {
    return { model, ok: false, ms, error: 'JSON.parse failed: ' + e.message, rawText: cleaned.slice(0, 500), usage, finishReason };
  }

  const entryCount = Array.isArray(data.entries) ? data.entries.length : 0;
  return { model, ok: true, ms, entryCount, usage, finishReason, data };
}

console.log(`Testing ${MODELS.length} models sequentially against ${INPUT_TEXT.split('\n').length}-line shift note...\n`);

const results = [];
for (const model of MODELS) {
  process.stdout.write(`→ ${model} ... `);
  const r = await testModel(model);
  results.push(r);
  if (r.ok) {
    console.log(`OK  ${r.ms}ms  entries=${r.entryCount}  tokens=${r.usage?.completion_tokens ?? '?'}  finish=${r.finishReason}`);
  } else {
    console.log(`FAIL ${r.ms}ms  ${r.error}`);
  }
}

console.log('\n=== SUMMARY ===');
console.table(results.map((r) => ({
  model: r.model,
  ok: r.ok,
  ms: r.ms,
  entries: r.entryCount ?? '-',
  completion_tokens: r.usage?.completion_tokens ?? '-',
  finish_reason: r.finishReason ?? '-',
  error: r.error ?? ''
})));

console.log('\n=== FULL ENTRIES (successful models) ===');
for (const r of results) {
  if (r.ok) {
    console.log(`\n--- ${r.model} ---`);
    console.log(JSON.stringify(r.data, null, 2));
  }
}
