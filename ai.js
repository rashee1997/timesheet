/**
 * CLOUDFLARE.GS — all AI plumbing via Cloudflare AI Gateway
 * -------------------------------------------------------------------------
 * This file owns:
 *   1. Credentials storage (setupCloudflareCredentials, askAndSetCloudflareCredentials)
 *      and the shared low-level caller (callCloudflareAiJson) with retries,
 *      backoff, fence-stripping, and JSON validation.
 *   2. parseNaturalLanguageEntries()  — Converts pasted free-text shift notes
 *      into structured entries utilizing the Cloudflare edge model.
 *
 * Shares CONFIG, parseSheetNameForMonthYear(), getEmployeesForSheet(),
 * getJobOrderSuggestions() etc. from Code.gs (same project, same global scope).
 */

// ---------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------

/**
 * Highly capable, multi-lingual text generation model on Cloudflare Workers AI.
 * Free tier: 10,000 Neurons/day total across all models (~15-25 calls/day for an 8B model).
 * Swap the string below to try alternatives — all run on the free tier.
 *
 * NON-REASONING (no hidden "thinking" pass — fast, won't hit the empty-content/finish_reason:"length" bug):
 *   '@cf/meta/llama-3.1-8b-instruct'                 8B,  fast, solid JSON-following, function calling
 *   '@cf/meta/llama-3.1-8b-instruct-fp8'             8B,  same but quantized (cheaper/faster)
 *   '@cf/meta/llama-3.2-3b-instruct'                 3B,  smaller/faster, less accurate on messy input
 *   '@cf/meta/llama-3.3-70b-instruct-fp8-fast'        70B, much stronger, still non-reasoning, function calling
 *   '@cf/mistralai/mistral-small-3.1-24b-instruct'   24B, good instruction following, function calling
 *   '@cf/qwen/qwen2.5-coder-32b-instruct'            32B, no function calling, strict about structure
 *   '@cf/meta/llama-4-scout-17b-16e-instruct'        17B MoE, vision-capable (used below as CLOUDFLARE_VISION_MODEL)
 *
 * REASONING / "thinking" models (higher quality but can burn max_tokens on hidden reasoning
 * before answering — see the finish_reason:"length" handling in callCloudflareAiJson):
 *   '@cf/google/gemma-4-26b-a4b-it'                  26B MoE (4B active), current default, vision-capable
 *   '@cf/openai/gpt-oss-20b'                         20B, function calling
 *   '@cf/openai/gpt-oss-120b'                        120B, function calling
 *   '@cf/qwen/qwen3-30b-a3b-fp8'                      30B MoE, function calling
 *   '@cf/deepseek-ai/deepseek-r1-distill-qwen-32b'   32B, reasoning-focused, no function calling
 *   '@cf/qwen/qwq-32b'                                32B, reasoning-focused, no function calling
 *
 * Exact IDs (esp. vendor prefix) can drift — confirm against
 * https://developers.cloudflare.com/workers-ai/models/ before relying on one; an unknown/renamed
 * id just fails fast with a 400 from callCloudflareAiJson's non-2xx error path.
 */
const CLOUDFLARE_MODEL = '@cf/openai/gpt-oss-120b';
const CLOUDFLARE_ACCOUNT_ID_PROPERTY = 'CLOUDFLARE_ACCOUNT_ID';
const CLOUDFLARE_API_TOKEN_PROPERTY = 'CLOUDFLARE_API_TOKEN';

const CLOUDFLARE_MAX_ATTEMPTS = 3;
const CLOUDFLARE_TIMEOUT_BACKOFF_MS = [500, 1500, 3500]; // wait before attempt 2, 3...
const CLOUDFLARE_FETCH_TIMEOUT_MS = 60000;               // per-attempt timeout (60s for multimodal)
const AI_PARSE_MAX_ENTRIES = 80;                         // sanity cap on parsed entries
const AI_PARSE_MAX_INPUT_CHARS = 6000;                  // sanity cap on pasted text


// ---------------------------------------------------------------------
// CREDENTIAL STORAGE FUNCTIONS
// ---------------------------------------------------------------------

/**
 * OPTION A: Run this function directly from the Apps Script editor dropdown.
 * Replace the placeholder text with your actual credentials before running.
 */
function setupCloudflareCredentials() {
  const accountId = 'YOUR_CLOUDFLARE_ACCOUNT_ID_HERE'; 
  const apiToken = 'YOUR_CLOUDFLARE_API_TOKEN_HERE';
  
  if (accountId === 'YOUR_CLOUDFLARE_ACCOUNT_ID_HERE' || apiToken === 'YOUR_CLOUDFLARE_API_TOKEN_HERE') {
    throw new Error('Please replace the placeholder values with your actual Cloudflare credentials before running this function.');
  }
  
  setCloudflareCredentials(accountId, apiToken);
}

/**
 * OPTION B: Run this function to trigger secure pop-up dialog boxes in the
 * active Google Sheet to enter your credentials without editing code.
 */
function askAndSetCloudflareCredentials() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const accountResponse = ui.prompt('Cloudflare Configuration', 'Please enter your Cloudflare Account ID:', ui.ButtonSet.OK_CANCEL);
    if (accountResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Configuration cancelled.');
      return;
    }
    const accountId = accountResponse.getResponseText().trim();
    if (!accountId) {
      ui.alert('Account ID cannot be empty.');
      return;
    }

    const tokenResponse = ui.prompt('Cloudflare Configuration', 'Please enter your Cloudflare API Token:', ui.ButtonSet.OK_CANCEL);
    if (tokenResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Configuration cancelled.');
      return;
    }
    const apiToken = tokenResponse.getResponseText().trim();
    if (!apiToken) {
      ui.alert('API Token cannot be empty.');
      return;
    }

    setCloudflareCredentials(accountId, apiToken);
    ui.alert('Cloudflare credentials saved successfully!');
  } catch (e) {
    Logger.log('Could not open spreadsheet UI prompts. If running from the Apps Script editor, use setupCloudflareCredentials() instead.');
    throw e;
  }
}

/**
 * Programmatic helper to write values to Script Properties.
 */
function setCloudflareCredentials(accountId, apiToken) {
  if (!accountId || !apiToken) {
    throw new Error('Both Account ID and API Token must be provided.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CLOUDFLARE_ACCOUNT_ID_PROPERTY, accountId.trim());
  props.setProperty(CLOUDFLARE_API_TOKEN_PROPERTY, apiToken.trim());
  Logger.log('Cloudflare credentials successfully saved to Script Properties.');
  return 'Cloudflare credentials successfully saved.';
}


// ---------------------------------------------------------------------
// SHARED LOW-LEVEL CALLER
// ---------------------------------------------------------------------

/**
 * Calls Cloudflare Workers AI Gateway expecting a JSON response.
 * Returns { success: true, data: <parsed object> } or { success: false, error: <human readable> }.
 */
function callCloudflareAiJson(prompt, temperature, messages) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accountId = scriptProperties.getProperty(CLOUDFLARE_ACCOUNT_ID_PROPERTY);
  const apiToken = scriptProperties.getProperty(CLOUDFLARE_API_TOKEN_PROPERTY);

  if (!accountId || !apiToken) {
    return {
      success: false,
      error: 'No Cloudflare credentials configured. Run setupCloudflareCredentials() or askAndSetCloudflareCredentials() to set them up.'
    };
  }

  const url = 'https://api.cloudflare.com/client/v4/accounts/' + accountId + '/ai/v1/chat/completions';

  const payload = {
    model: CLOUDFLARE_MODEL,
    messages: messages || [
      {
        role: 'system',
        content: 'You are a precise JSON-only utility assistant. You must output raw, valid JSON only. Do not wrap your response in markdown code blocks (fences) and do not include conversational text or explanations.'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: (typeof temperature === 'number') ? temperature : 0.4,
    max_tokens: 6000,
    reasoning_effort: 'low',
    chat_template_kwargs: { thinking: false },
    response_format: { type: 'json_object' }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiToken,
      'cf-aig-gateway-id': 'default'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let lastError = 'Unknown error.';

  for (let attempt = 0; attempt < CLOUDFLARE_MAX_ATTEMPTS; attempt++) {
    if (attempt > 0) {
      Utilities.sleep(CLOUDFLARE_TIMEOUT_BACKOFF_MS[attempt - 1] || 3000);
    }

    try {
      const response = UrlFetchApp.fetch(url, Object.assign({}, options, { muteHttpExceptions: true, timeout: CLOUDFLARE_FETCH_TIMEOUT_MS }));
      const code = response.getResponseCode();
      const resText = response.getContentText();

      if (code === 429 || code >= 500) {
        lastError = 'Cloudflare API busy (status ' + code + '). Retrying...';
        Logger.log(lastError);
        continue;
      }

      if (code !== 200) {
        let errDetails = resText;
        try {
          const parsed = JSON.parse(resText);
          if (parsed.errors && parsed.errors[0]) {
            errDetails = parsed.errors[0].message;
          } else if (parsed.error && parsed.error.message) {
            errDetails = parsed.error.message;
          }
        } catch (ignored) { /* resText was not structured JSON */ }
        return { success: false, error: 'Cloudflare API Error (Status ' + code + '): ' + errDetails };
      }

      let resJson;
      try {
        resJson = JSON.parse(resText);
      } catch (parseErr) {
        return { success: false, error: 'Cloudflare returned an unparsable response format: ' + parseErr };
      }

      let rawText = extractMessageText(resJson.choices && resJson.choices[0] && resJson.choices[0].message && resJson.choices[0].message.content);
      if (!rawText && resJson.result && resJson.result.response) {
        rawText = String(resJson.result.response).trim();
      }
      if (!rawText) {
        const finishReason = resJson.choices && resJson.choices[0] && resJson.choices[0].finish_reason;
        lastError = (finishReason === 'length')
          ? 'Model spent its whole token budget on internal reasoning and never produced an answer. Retrying...'
          : 'Cloudflare response payload contained no text content. Retrying...';
        Logger.log(lastError);
        continue;
      }

      if (rawText.startsWith('```')) {
        rawText = rawText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/, '');
      }

      let data;
      try {
        data = JSON.parse(rawText);
      } catch (jsonErr) {
        return { success: false, error: 'Failed to parse JSON: ' + jsonErr + '. Raw text received: ' + rawText.slice(0, 300) };
      }

      return { success: true, data: data };
    } catch (e) {
      lastError = 'Execution/Network Error: ' + e;
      Logger.log('callCloudflareAiJson attempt ' + (attempt + 1) + ' failed: ' + lastError);
    }
  }

  return { success: false, error: lastError + ' (gave up after ' + CLOUDFLARE_MAX_ATTEMPTS + ' attempts)' };
}


// ---------------------------------------------------------------------
// NATURAL-LANGUAGE ENTRY PARSER
// ---------------------------------------------------------------------

/**
 * Converts pasted free-text shift notes into structured entries.
 */
function parseNaturalLanguageEntries(text, sheetName) {
  try {
    if (!text || String(text).trim() === '') {
      return { success: false, error: 'Paste some shift notes first.' };
    }
    let input = String(text).trim();
    if (input.length > AI_PARSE_MAX_INPUT_CHARS) {
      input = input.slice(0, AI_PARSE_MAX_INPUT_CHARS);
    }

    const empResult = getEmployeesForSheet(sheetName);
    if (empResult.error) return { success: false, error: empResult.error };
    const employeeNames = empResult.employees.map(function (e) { return e.name; });

    let jobOrders = [];
    try {
      jobOrders = getJobOrderSuggestions().slice(0, 40);
    } catch (e) {
      Logger.log('parseNaturalLanguageEntries: job order lookup failed: ' + e);
    }

    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const weekday = Utilities.formatDate(now, tz, 'EEEE');

    const detected = parseSheetNameForMonthYear(sheetName);
    let sheetPeriodLine;
    if (detected) {
      const periodDate = new Date(detected.year, detected.monthIndex, 1);
      sheetPeriodLine = 'The target sheet "' + sheetName + '" covers ' +
        Utilities.formatDate(periodDate, tz, 'MMMM yyyy') +
        '. When the text gives only a day number (e.g. "12th", "on 5"), resolve it into this month and year.';
    } else {
      sheetPeriodLine = 'The target sheet is "' + sheetName + '" (month could not be detected from its name). ' +
        'When the text gives only a day number, resolve it into the current month and year.';
    }

    const prompt =
      'You convert informal timesheet notes into structured JSON for a construction/industrial contractor in Qatar.\n\n' +
      'CONTEXT:\n' +
      '- Today is ' + todayStr + ' (' + weekday + '). Resolve relative dates like "today", "yesterday", "last Monday" against this.\n' +
      '- ' + sheetPeriodLine + '\n' +
      '- Known employees on this sheet (use these EXACT names when you are confident of a match):\n  ' + JSON.stringify(employeeNames) + '\n' +
      '- Known job orders (prefer these exact values when the text clearly refers to one):\n  ' + JSON.stringify(jobOrders) + '\n\n' +
      '- Expect date formats like "DD.MM.YY", "DD.MM.YYYY", "DD/MM/YYYY". A date line at the top (e.g. "14.7.26") sets the date for ALL following entries unless another date appears.\n' +
      'INPUT NOTES (may be WhatsApp-style shorthand, mixed English/Tamil, abbreviations, tables, or messy lists):\n' +
      '"""\n' + input + '\n"""\n\n' +
      'RULES:\n' +
      '1. Output times in strict 24-hour "HH:mm". Interpret common shorthand: "8 to 5" = 08:00 to 17:00, "7-7" = 07:00 to 19:00, "8pm to 4am" = 20:00 to 04:00, "8am-4pm" = 08:00 to 16:00, "7am-3pm" = 07:00 to 15:00, "7am-5pm" = 07:00 to 17:00.\n' +
      '2. Dates in strict "YYYY-MM-DD".\n' +
      '3. Group people who share the SAME date, start, end and job order into ONE entry with multiple names in "employees".\n' +
      '3a. Start and end times are PER-PERSON, never a header. Read each person’s own line for their own time. If a line gives no time at all for that person, and only then, reuse the most recent time that appeared above it. But if a person’s line states its OWN time (e.g. "8 to 4pm"), you MUST use that person’s stated time even if it differs from the person listed just before or after them - never overwrite it with a neighboring line’s time. Unlike job orders, times do NOT apply retroactively or as a running default once someone states their own.\n' +
      '4. For employee names: if a name in the notes clearly matches a known employee (ignoring case, spacing, partial spelling, or minor typos like "suganathan" vs "SUGUNADAN"/"SUGANANTHAN"), return the EXACT known name. If you cannot confidently match, return the name exactly as written in the notes and set confidence to "low" with a note explaining.\n' +
      '5. If the job order is not stated and cannot be inferred, use an empty string "".\n' +
      '6. Never invent employees, dates, or times that are not in the notes. If something essential is missing or ambiguous, skip that entry and add an explanation to "warnings".\n' +
      '7. When you guessed anything (AM vs PM, which month, which person), say so in that entry\u2019s "note" and lower its confidence.\n' +
      '8. Numbered list format like "1.Rasheedh 8am-4pm" or "2.Ravi 7am-3pm" means each line is one person with their shift. The number is just a list marker, not part of the name.\n' +
      '9. Job orders can be written in ANY of these forms - handle all of them: (a) inline on a person\'s own line, e.g. "1.Rasheedh 8am-4pm GW"; (b) on its own line by itself (just a site/client/project name, with or without a label like "Job:", "Site:", "For:", "@"), acting as a HEADER that applies to every entry from that point until a new job-order line, a new date line, or the text ends - exactly like how a date line at the top governs all following entries; (c) a header line placed AFTER a whole numbered list, applying retroactively to that entire list when no per-line job orders were given. A job order stated directly on a person\'s own line always overrides a header for that person. Never leave a job order blank just because it appeared on a separate line from the shift details - trace back (or forward) to find the nearest applicable header.\n' +
      '10. The pasted text may open or close with an instruction line meant for YOU rather than being part of the timesheet data, e.g. "only Sundar", "extract only Rasheedh\'s time", "just show Ravi and Ganesh". If such an instruction names one or more employees, treat it as a FILTER: parse the entire pasted timesheet as usual (across all dates and people) but only include entries for the named employee(s) in the output, using the same fuzzy name matching as rule 4. Do not treat the instruction line itself as a date, job order, or employee entry. If no such instruction is present, output entries for everyone as normal.\n\n' +
      'Return ONLY a valid JSON object matching this structure:\n' +
      '{\n' +
      '  "entries": [\n' +
      '    { "date": "YYYY-MM-DD", "employees": ["NAME"], "startTime": "HH:mm", "endTime": "HH:mm", "jobOrder": "", "confidence": "high", "note": "" }\n' +
      '  ],\n' +
      '  "warnings": []\n' +
      '}';

    const result = callCloudflareAiJson(prompt, 0.2);
    if (!result.success) return { success: false, error: result.error };

    const data = result.data;
    if (!data || !Array.isArray(data.entries)) {
      return { success: false, error: 'QuickFill response did not contain an "entries" list. Try parsing again.' };
    }

    // ---- Server-side validation of the AI output ----
    const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
    const TIME_RE = /^([01]?\d|2[0-3]):[0-5]\d$/;

    const warnings = Array.isArray(data.warnings)
      ? data.warnings.filter(function (w) { return typeof w === 'string' && w.trim() !== ''; })
      : [];
    const validEntries = [];

    data.entries.slice(0, AI_PARSE_MAX_ENTRIES).forEach(function (e, i) {
      const label = 'QuickFill entry #' + (i + 1);
      if (!e || typeof e !== 'object') { warnings.push(label + ' was malformed and was dropped.'); return; }
      if (typeof e.date !== 'string' || !DATE_RE.test(e.date)) {
        warnings.push(label + ' had an invalid date ("' + e.date + '") and was dropped.');
        return;
      }
      if (typeof e.startTime !== 'string' || !TIME_RE.test(e.startTime) ||
          typeof e.endTime !== 'string' || !TIME_RE.test(e.endTime)) {
        warnings.push(label + ' had invalid times and was dropped.');
        return;
      }
      const emps = Array.isArray(e.employees)
        ? e.employees.map(function (n) { return String(n || '').trim(); }).filter(function (n) { return n !== ''; })
        : [];
      if (emps.length === 0) {
        warnings.push(label + ' had no employees and was dropped.');
        return;
      }

      validEntries.push({
        date: e.date,
        employees: emps,
        startTime: padTime(e.startTime),
        endTime: padTime(e.endTime),
        jobOrder: (typeof e.jobOrder === 'string') ? e.jobOrder.trim() : '',
        confidence: (e.confidence === 'high' || e.confidence === 'medium' || e.confidence === 'low') ? e.confidence : 'medium',
        note: (typeof e.note === 'string') ? e.note.trim() : ''
      });
    });

    if (data.entries.length > AI_PARSE_MAX_ENTRIES) {
      warnings.push('Only the first ' + AI_PARSE_MAX_ENTRIES + ' entries were kept — split very large batches into smaller pastes.');
    }
    if (validEntries.length === 0) {
      return {
        success: false,
        error: 'QuickFill could not extract any complete entries from that text.' +
          (warnings.length ? '\n' + warnings.join('\n') : '')
      };
    }

    return { success: true, entries: validEntries, warnings: warnings };
  } catch (err) {
    Logger.log('parseNaturalLanguageEntries error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

/**
 * Normalizes an OpenAI-compat message.content field to a plain trimmed string.
 * Some models (e.g. llama-4-scout) return content as an array of parts
 * (e.g. [{ type: "text", text: "..." }]) instead of a plain string.
 */
function extractMessageText(content) {
  if (!content) return '';
  if (typeof content === 'string') return content.trim();
  if (Array.isArray(content)) {
    return content
      .map(function (part) { return (part && typeof part.text === 'string') ? part.text : ''; })
      .join('')
      .trim();
  }
  return '';
}

/** Pads short times (e.g. "8:30" -> "08:30") for safe browser interface binding. */
function padTime(t) {
  const parts = String(t).split(':');
  return ('0' + parts[0]).slice(-2) + ':' + parts[1];
}


// ---------------------------------------------------------------------
// VISION (IMAGE) API — Multimodal Cloudflare model for screenshot parsing
// ---------------------------------------------------------------------

// Routed through the AI Gateway (not the raw Workers AI endpoint) because these are
// BYOK Google AI Studio models — the Gemini key is stored as a Gateway secret, so no
// provider Authorization header is sent here; Cloudflare injects it.
const CLOUDFLARE_GATEWAY_ID = 'default';
const GEMINI_VISION_MODEL_PRIMARY = 'google-ai-studio/gemini-3.5-flash';
const GEMINI_VISION_MODEL_FALLBACK = 'google-ai-studio/gemini-3.1-flash-lite';
const AI_PARSE_MAX_IMAGE_SIZE_BYTES = 4 * 1024 * 1024; // 4 MB after client-side resize

/**
 * Calls one Gateway-routed model, retrying on transient (429/5xx/empty-content) failures.
 * Returns { success: true, data } or { success: false, error, retryable: <bool for next model> }.
 */
function callGatewayChatCompletion(url, apiToken, model, messages, temperature) {
  const payload = {
    model: model,
    messages: messages,
    temperature: (typeof temperature === 'number') ? temperature : 0.2,
    max_tokens: 6000,
    response_format: { type: 'json_object' }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 60000
  };

  let lastError = 'Unknown error.';

  for (let attempt = 0; attempt < CLOUDFLARE_MAX_ATTEMPTS; attempt++) {
    if (attempt > 0) {
      Utilities.sleep(CLOUDFLARE_TIMEOUT_BACKOFF_MS[attempt - 1] || 3000);
    }

    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const resText = response.getContentText();

      if (code === 429 || code >= 500) {
        lastError = model + ' busy (status ' + code + '). Retrying...';
        Logger.log(lastError);
        continue;
      }

      if (code !== 200) {
        let errDetails = resText;
        try {
          const parsed = JSON.parse(resText);
          if (parsed.errors && parsed.errors[0]) {
            errDetails = parsed.errors[0].message;
          } else if (parsed.error && parsed.error.message) {
            errDetails = parsed.error.message;
          }
        } catch (ignored) {}
        return { success: false, error: model + ' Error (Status ' + code + '): ' + errDetails };
      }

      let resJson;
      try {
        resJson = JSON.parse(resText);
      } catch (parseErr) {
        return { success: false, error: model + ' returned an unparsable response: ' + parseErr };
      }

      let rawText = extractMessageText(resJson.choices && resJson.choices[0] && resJson.choices[0].message && resJson.choices[0].message.content);
      if (!rawText) {
        const finishReason = resJson.choices && resJson.choices[0] && resJson.choices[0].finish_reason;
        lastError = model + ' returned no text content (finish_reason: ' + finishReason + '). Retrying...';
        Logger.log(lastError);
        continue;
      }

      if (rawText.startsWith('```')) {
        rawText = rawText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/, '');
      }

      let data;
      try {
        data = JSON.parse(rawText);
      } catch (jsonErr) {
        return { success: false, error: model + ': failed to parse JSON: ' + jsonErr + '. Raw: ' + rawText.slice(0, 300) };
      }

      return { success: true, data: data };
    } catch (e) {
      lastError = model + ' Execution/Network Error: ' + e;
      Logger.log('callGatewayChatCompletion attempt ' + (attempt + 1) + ' failed: ' + lastError);
    }
  }

  return { success: false, error: lastError + ' (gave up after ' + CLOUDFLARE_MAX_ATTEMPTS + ' attempts)' };
}

/**
 * Calls the primary Gemini vision model, falling back to the lite model on failure.
 * Returns { success: true, data: <parsed object> } or { success: false, error: <human readable> }.
 */
function callCloudflareAiVision(messages, temperature) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accountId = scriptProperties.getProperty(CLOUDFLARE_ACCOUNT_ID_PROPERTY);
  const apiToken = scriptProperties.getProperty(CLOUDFLARE_API_TOKEN_PROPERTY);

  if (!accountId || !apiToken) {
    return {
      success: false,
      error: 'No Cloudflare credentials configured. Run setupCloudflareCredentials() or askAndSetCloudflareCredentials() to set them up.'
    };
  }

  const url = 'https://gateway.ai.cloudflare.com/v1/' + accountId + '/' + CLOUDFLARE_GATEWAY_ID + '/compat/chat/completions';

  const primary = callGatewayChatCompletion(url, apiToken, GEMINI_VISION_MODEL_PRIMARY, messages, temperature);
  if (primary.success) return primary;

  Logger.log(GEMINI_VISION_MODEL_PRIMARY + ' failed, falling back to ' + GEMINI_VISION_MODEL_FALLBACK + ': ' + primary.error);
  const fallback = callGatewayChatCompletion(url, apiToken, GEMINI_VISION_MODEL_FALLBACK, messages, temperature);
  if (fallback.success) return fallback;

  return { success: false, error: 'Both Gemini models failed. ' + GEMINI_VISION_MODEL_PRIMARY + ': ' + primary.error + ' | ' + GEMINI_VISION_MODEL_FALLBACK + ': ' + fallback.error };
}


/**
 * Converts pasted text + screenshot into structured entries using a vision model.
 */
function parseNaturalLanguageEntriesWithImage(text, imageBase64, mimeType, sheetName) {
  try {
    if (!imageBase64) {
      return { success: false, error: 'No image data received.' };
    }
    if (imageBase64.length > AI_PARSE_MAX_IMAGE_SIZE_BYTES) {
      return { success: false, error: 'Image too large after compression (' + Math.round(imageBase64.length / 1024) + ' KB). Try a smaller screenshot.' };
    }

    var userText = String(text || '').trim();
    if (userText.length > AI_PARSE_MAX_INPUT_CHARS) {
      userText = userText.slice(0, AI_PARSE_MAX_INPUT_CHARS);
    }

    const empResult = getEmployeesForSheet(sheetName);
    if (empResult.error) return { success: false, error: empResult.error };
    const employeeNames = empResult.employees.map(function (e) { return e.name; });

    let jobOrders = [];
    try {
      jobOrders = getJobOrderSuggestions().slice(0, 40);
    } catch (e) {
      Logger.log('parseNaturalLanguageEntriesWithImage: job order lookup failed: ' + e);
    }

    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const weekday = Utilities.formatDate(now, tz, 'EEEE');

    const detected = parseSheetNameForMonthYear(sheetName);
    let sheetPeriodLine;
    if (detected) {
      const periodDate = new Date(detected.year, detected.monthIndex, 1);
      sheetPeriodLine = 'The target sheet "' + sheetName + '" covers ' +
        Utilities.formatDate(periodDate, tz, 'MMMM yyyy') +
        '. When the text gives only a day number (e.g. "12th", "on 5"), resolve it into this month and year.';
    } else {
      sheetPeriodLine = 'The target sheet is "' + sheetName + '" (month could not be detected from its name). ' +
        'When the text gives only a day number, resolve it into the current month and year.';
    }

    const systemMsg =
      'You are a precise JSON-only utility assistant. You must output raw, valid JSON only. Do not wrap your response in markdown code blocks (fences) and do not include conversational text or explanations.';

    const userMsg =
      'You convert informal timesheet notes (possibly from a screenshot of WhatsApp chat) into structured JSON for a construction/industrial contractor in Qatar.\n\n' +
      'CONTEXT:\n' +
      '- Today is ' + todayStr + ' (' + weekday + '). Resolve relative dates like "today", "yesterday", "last Monday" against this.\n' +
      '- ' + sheetPeriodLine + '\n' +
      '- Known employees on this sheet (use these EXACT names when you are confident of a match):\n  ' + JSON.stringify(employeeNames) + '\n' +
      '- Known job orders (prefer these exact values when the text clearly refers to one):\n  ' + JSON.stringify(jobOrders) + '\n\n' +
      '- Expect date formats like "DD.MM.YY", "DD.MM.YYYY", "DD/MM/YYYY". A date line at the top (e.g. "14.7.26") sets the date for ALL following entries unless another date appears.\n' +
      'ADDITIONAL TEXT FROM USER (may be empty or corrections):\n"""\n' + userText + '\n"""\n\n' +
      'RULES:\n' +
      '1. Output times in strict 24-hour "HH:mm". Interpret common shorthand: "8 to 5" = 08:00 to 17:00, "7-7" = 07:00 to 19:00, "8pm to 4am" = 20:00 to 04:00, "8am-4pm" = 08:00 to 16:00, "7am-3pm" = 07:00 to 15:00, "7am-5pm" = 07:00 to 17:00.\n' +
      '2. Dates in strict "YYYY-MM-DD".\n' +
      '3. Group people who share the SAME date, start, end and job order into ONE entry with multiple names in "employees".\n' +
      '3a. Start and end times are PER-PERSON, never a header. Read each person’s own line for their own time. If a line gives no time at all for that person, and only then, reuse the most recent time that appeared above it. But if a person’s line states its OWN time (e.g. "8 to 4pm"), you MUST use that person’s stated time even if it differs from the person listed just before or after them - never overwrite it with a neighboring line’s time. Unlike job orders, times do NOT apply retroactively or as a running default once someone states their own.\n' +
      '4. For employee names: if a name in the image clearly matches a known employee (ignoring case, spacing, partial spelling, or minor typos like "suganathan" vs "SUGUNADAN"/"SUGANANTHAN"), return the EXACT known name. If you cannot confidently match, return the name as written and set confidence to "low" with a note.\n' +
      '5. If the job order is not stated and cannot be inferred, use an empty string "".\n' +
      '6. Never invent employees, dates, or times that are not in the image/text. If something is missing or ambiguous, skip that entry and add a "warnings" entry.\n' +
      '7. When you guessed anything (AM vs PM, which month, which person), say so in that entry\'s "note" and lower its confidence.\n' +
      '8. Numbered list format like "1.Rasheedh 8am-4pm" or "2.Ravi 7am-3pm" means each line is one person with their shift. The number is just a list marker, not part of the name.\n' +
      '9. Job orders can be written in ANY of these forms - handle all of them: (a) inline on a person\'s own line, e.g. "1.Rasheedh 8am-4pm GW"; (b) on its own line by itself (just a site/client/project name, with or without a label like "Job:", "Site:", "For:", "@"), acting as a HEADER that applies to every entry from that point until a new job-order line, a new date line, or the text ends - exactly like how a date line at the top governs all following entries; (c) a header line placed AFTER a whole numbered list, applying retroactively to that entire list when no per-line job orders were given. A job order stated directly on a person\'s own line always overrides a header for that person. Never leave a job order blank just because it appeared on a separate line from the shift details - trace back (or forward) to find the nearest applicable header.\n' +
      '10. The "ADDITIONAL TEXT FROM USER" above may contain an instruction meant for YOU rather than shift data, e.g. "only Sundar", "extract only Rasheedh\'s time", "just show Ravi and Ganesh". If such an instruction names one or more employees, treat it as a FILTER: parse the entire image/timesheet as usual (across all dates and people) but only include entries for the named employee(s) in the output, using the same fuzzy name matching as rule 4. If no such instruction is present, output entries for everyone as normal.\n\n' +
      'Return ONLY a valid JSON object matching this structure:\n' +
      '{\n' +
      '  "entries": [\n' +
      '    { "date": "YYYY-MM-DD", "employees": ["NAME"], "startTime": "HH:mm", "endTime": "HH:mm", "jobOrder": "", "confidence": "high", "note": "" }\n' +
      '  ],\n' +
      '  "warnings": []\n' +
      '}';

    var imageDataUrl = 'data:' + mimeType + ';base64,' + imageBase64;

    const messages = [
      { role: 'system', content: systemMsg },
      {
        role: 'user',
        content: [
          { type: 'text', text: userMsg },
          { type: 'image_url', image_url: { url: imageDataUrl } }
        ]
      }
    ];

    const result = callCloudflareAiJson(null, 0.2, messages);
    if (!result.success) return { success: false, error: result.error };

    const data = result.data;
    if (!data || !Array.isArray(data.entries)) {
      return { success: false, error: 'QuickFill response did not contain an "entries" list. Try again.' };
    }

    // ---- Server-side validation (same as text-only parser) ----
    const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
    const TIME_RE = /^([01]?\d|2[0-3]):[0-5]\d$/;

    const warnings = Array.isArray(data.warnings)
      ? data.warnings.filter(function (w) { return typeof w === 'string' && w.trim() !== ''; })
      : [];

    if (userText) {
      warnings.push('QuickFill analyzed the screenshot' + (userText ? ' combined with your text notes' : '') + '. Please verify all entries.');
    }

    const validEntries = [];

    data.entries.slice(0, AI_PARSE_MAX_ENTRIES).forEach(function (e, i) {
      const label = 'QuickFill entry #' + (i + 1);
      if (!e || typeof e !== 'object') { warnings.push(label + ' was malformed and was dropped.'); return; }
      if (typeof e.date !== 'string' || !DATE_RE.test(e.date)) {
        warnings.push(label + ' had an invalid date ("' + e.date + '") and was dropped.');
        return;
      }
      if (typeof e.startTime !== 'string' || !TIME_RE.test(e.startTime) ||
          typeof e.endTime !== 'string' || !TIME_RE.test(e.endTime)) {
        warnings.push(label + ' had invalid times and was dropped.');
        return;
      }
      const emps = Array.isArray(e.employees)
        ? e.employees.map(function (n) { return String(n || '').trim(); }).filter(function (n) { return n !== ''; })
        : [];
      if (emps.length === 0) {
        warnings.push(label + ' had no employees and was dropped.');
        return;
      }

      validEntries.push({
        date: e.date,
        employees: emps,
        startTime: padTime(e.startTime),
        endTime: padTime(e.endTime),
        jobOrder: (typeof e.jobOrder === 'string') ? e.jobOrder.trim() : '',
        confidence: (e.confidence === 'high' || e.confidence === 'medium' || e.confidence === 'low') ? e.confidence : 'medium',
        note: (typeof e.note === 'string') ? e.note.trim() : ''
      });
    });

    if (data.entries.length > AI_PARSE_MAX_ENTRIES) {
      warnings.push('Only the first ' + AI_PARSE_MAX_ENTRIES + ' entries were kept.');
    }
    if (validEntries.length === 0) {
      return {
        success: false,
        error: 'QuickFill could not extract any complete entries from that image.' +
          (warnings.length ? '\n' + warnings.join('\n') : '')
      };
    }

    return { success: true, entries: validEntries, warnings: warnings };
  } catch (err) {
    Logger.log('parseNaturalLanguageEntriesWithImage error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}