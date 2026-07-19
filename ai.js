/**
 * CLOUDFLARE.GS — all AI plumbing via Cloudflare AI Gateway
 * -------------------------------------------------------------------------
 * This file owns:
 *   1. Credentials storage (setupCloudflareCredentials, askAndSetCloudflareCredentials)
 *      and the shared low-level caller with retries, backoff, and JSON validation.
 *   2. parseNaturalLanguageEntries() & parseNaturalLanguageEntriesWithImage() —
 *      converts pasted shift notes/images into structured entries.
 */

// ---------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------
const CLOUDFLARE_MODEL = '@cf/openai/gpt-oss-120b';
const CLOUDFLARE_ACCOUNT_ID_PROPERTY = 'CLOUDFLARE_ACCOUNT_ID';
const CLOUDFLARE_API_TOKEN_PROPERTY = 'CLOUDFLARE_API_TOKEN';

const CLOUDFLARE_MAX_ATTEMPTS = 3;
const CLOUDFLARE_TIMEOUT_BACKOFF_MS = [500, 1500, 3500];
const CLOUDFLARE_FETCH_TIMEOUT_MS = 60000;
const AI_PARSE_MAX_ENTRIES = 80;
const AI_PARSE_MAX_INPUT_CHARS = 6000;

const CLOUDFLARE_GATEWAY_ID = 'default';
const GEMINI_VISION_MODEL_PRIMARY = 'google-ai-studio/gemini-3.5-flash';
const GEMINI_VISION_MODEL_FALLBACK = 'google-ai-studio/gemini-3.1-flash-lite';
const AI_PARSE_MAX_IMAGE_SIZE_BYTES = 4 * 1024 * 1024;


// ---------------------------------------------------------------------
// CREDENTIAL STORAGE FUNCTIONS
// ---------------------------------------------------------------------

function setupCloudflareCredentials() {
  const accountId = 'YOUR_CLOUDFLARE_ACCOUNT_ID_HERE';
  const apiToken = 'YOUR_CLOUDFLARE_API_TOKEN_HERE';

  if (accountId === 'YOUR_CLOUDFLARE_ACCOUNT_ID_HERE' || apiToken === 'YOUR_CLOUDFLARE_API_TOKEN_HERE') {
    throw new Error('Please replace the placeholder values with actual Cloudflare credentials.');
  }
  setCloudflareCredentials(accountId, apiToken);
}

function askAndSetCloudflareCredentials() {
  try {
    const ui = SpreadsheetApp.getUi();

    const accountResponse = ui.prompt('Cloudflare Configuration', 'Please enter your Cloudflare Account ID:', ui.ButtonSet.OK_CANCEL);
    if (accountResponse.getSelectedButton() !== ui.Button.OK) return;
    const accountId = accountResponse.getResponseText().trim();

    const tokenResponse = ui.prompt('Cloudflare Configuration', 'Please enter your Cloudflare API Token:', ui.ButtonSet.OK_CANCEL);
    if (tokenResponse.getSelectedButton() !== ui.Button.OK) return;
    const apiToken = tokenResponse.getResponseText().trim();

    if (!accountId || !apiToken) {
      ui.alert('Account ID and API Token cannot be empty.');
      return;
    }

    setCloudflareCredentials(accountId, apiToken);
    ui.alert('Cloudflare credentials saved successfully!');
  } catch (e) {
    Logger.log('Could not open spreadsheet UI prompts. Run setupCloudflareCredentials() instead.');
    throw e;
  }
}

function setCloudflareCredentials(accountId, apiToken) {
  if (!accountId || !apiToken) {
    throw new Error('Both Account ID and API Token must be provided.');
  }
  PropertiesService.getScriptProperties().setProperties({
    [CLOUDFLARE_ACCOUNT_ID_PROPERTY]: accountId.trim(),
    [CLOUDFLARE_API_TOKEN_PROPERTY]: apiToken.trim()
  });
  Logger.log('Cloudflare credentials successfully saved to Script Properties.');
  return 'Cloudflare credentials successfully saved.';
}


// ---------------------------------------------------------------------
// SHARED LOW-LEVEL CALLER
// ---------------------------------------------------------------------

/**
 * Handles HTTP requests with retries, status code evaluation, backoff, and JSON extraction.
 */
function executeFetchWithRetry(url, headers, payload, modelName) {
  let lastError = 'Unknown error.';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: CLOUDFLARE_FETCH_TIMEOUT_MS
  };

  for (let attempt = 0; attempt < CLOUDFLARE_MAX_ATTEMPTS; attempt++) {
    if (attempt > 0) {
      Utilities.sleep(CLOUDFLARE_TIMEOUT_BACKOFF_MS[attempt - 1] || 3000);
    }

    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const resText = response.getContentText();

      if (code === 429 || code >= 500) {
        lastError = `${modelName} busy (status ${code}). Retrying...`;
        Logger.log(lastError);
        continue;
      }

      if (code !== 200) {
        let errDetails = resText;
        try {
          const parsed = JSON.parse(resText);
          errDetails = parsed.errors?.[0]?.message || parsed.error?.message || resText;
        } catch (ignored) {}
        return { success: false, error: `${modelName} Error (Status ${code}): ${errDetails}` };
      }

      const resJson = JSON.parse(resText);
      let rawText = extractMessageText(resJson.choices?.[0]?.message?.content);
      if (!rawText && resJson.result?.response) {
        rawText = String(resJson.result.response).trim();
      }

      if (!rawText) {
        const finishReason = resJson.choices?.[0]?.finish_reason;
        lastError = (finishReason === 'length')
          ? 'Model spent token budget on reasoning. Retrying...'
          : 'Empty response payload. Retrying...';
        Logger.log(lastError);
        continue;
      }

      if (rawText.startsWith('```')) {
        rawText = rawText.replace(/^```(?:json)?\n?/i, '').replace(/\n?```$/, '');
      }

      return { success: true, data: JSON.parse(rawText) };
    } catch (e) {
      lastError = `Network/Parsing Error: ${e}`;
      Logger.log(`Fetch attempt ${attempt + 1} failed: ${lastError}`);
    }
  }
  return { success: false, error: `${lastError} (gave up after ${CLOUDFLARE_MAX_ATTEMPTS} attempts)` };
}

function callCloudflareAiJson(prompt, temperature, messages) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accountId = scriptProperties.getProperty(CLOUDFLARE_ACCOUNT_ID_PROPERTY);
  const apiToken = scriptProperties.getProperty(CLOUDFLARE_API_TOKEN_PROPERTY);

  if (!accountId || !apiToken) {
    return { success: false, error: 'No Cloudflare credentials configured.' };
  }

  const url = 'https://api.cloudflare.com/client/v4/accounts/' + accountId + '/ai/v1/chat/completions';
  const payload = {
    model: CLOUDFLARE_MODEL,
    messages: messages || [
      { role: 'system', content: 'You are a precise JSON-only utility assistant. Output raw, valid JSON only. Do not use code fences.' },
      { role: 'user', content: prompt }
    ],
    temperature: (typeof temperature === 'number') ? temperature : 0.4,
    max_tokens: 6000,
    reasoning_effort: 'low',
    chat_template_kwargs: { thinking: false },
    response_format: { type: 'json_object' }
  };

  const headers = {
    'Authorization': 'Bearer ' + apiToken,
    'cf-aig-gateway-id': 'default'
  };

  return executeFetchWithRetry(url, headers, payload, CLOUDFLARE_MODEL);
}


// ---------------------------------------------------------------------
// CONTEXT & VALIDATION HELPERS
// ---------------------------------------------------------------------

/**
 * Builds the comprehensive prompt rules, merging dynamic sheet context, employees, and job orders.
 */
function getContextAndRules(sheetName, userText) {
  const empResult = getEmployeesForSheet(sheetName);
  const employeeNames = empResult.error ? [] : empResult.employees.map(e => e.name);

  let jobOrders = [];
  try {
    jobOrders = getJobOrderSuggestions().slice(0, 40);
  } catch (e) {
    Logger.log('Job order lookup bypassed: ' + e);
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const weekday = Utilities.formatDate(now, tz, 'EEEE');

  const detected = parseSheetNameForMonthYear(sheetName);
  const sheetPeriodLine = detected
    ? `The target sheet "${sheetName}" covers ${Utilities.formatDate(new Date(detected.year, detected.monthIndex, 1), tz, 'MMMM yyyy')}. Resolve day numbers (e.g. "12th") into this month and year.`
    : `The target sheet is "${sheetName}". Resolve day numbers into current month and year.`;

  return `You convert timesheet notes into structured JSON for a contractor in Qatar.\n\n` +
    `CONTEXT:\n` +
    `- Today is ${todayStr} (${weekday}). Resolve relative dates like "today" or "yesterday" against this.\n` +
    `- ${sheetPeriodLine}\n` +
    `- Known employees: ${JSON.stringify(employeeNames)}\n` +
    `- Known job orders: ${JSON.stringify(jobOrders)}\n\n` +
    `RULES:\n` +
    `1. Output times in strict 24-hour "HH:mm". Interpret shorthands (e.g., "8 to 5" = 08:00 to 17:00, "7-7" = 07:00 to 19:00).\n` +
    `2. Dates in strict "YYYY-MM-DD".\n` +
    `3. Group people sharing the SAME date, start, end, and job order into ONE entry under "employees".\n` +
    `3a. Start and end times are PER-PERSON. If a line gives no time, reuse the preceding time. If a line states its own time, you must use it.\n` +
    `4. Fuzzy match employee names to Known employees. If matching is low confidence, use the written name, set confidence to "low", and note why.\n` +
    `5. If a job order is missing and cannot be inferred, use an empty string "".\n` +
    `6. Do not invent details. Skip incomplete entries and add an explanation to "warnings".\n` +
    `7. Lower confidence and add a "note" if you make guesses.\n` +
    `8. Numbered lists imply separate shifts for each person.\n` +
    `9. Job orders can be inline, standalone headers (applies forward), or trailing headers (applies retroactively). Inline overrides headers.\n` +
    `10. Treat instructions like "only Sundar" as a filter: parse everything but only output entries matching those employees.\n\n` +
    `Return ONLY a valid JSON object matching this structure:\n` +
    `{\n` +
    `  "entries": [\n    { "date": "YYYY-MM-DD", "employees": ["NAME"], "startTime": "HH:mm", "endTime": "HH:mm", "jobOrder": "", "confidence": "high", "note": "" }\n  ],\n` +
    `  "warnings": []\n` +
    `}`;
}

/**
 * Validates parsed entries, sanitizes data types, and normalizes times.
 */
function validateAndCleanEntries(data, initialWarnings) {
  const warnings = Array.isArray(initialWarnings) ? initialWarnings : [];
  if (!data || !Array.isArray(data.entries)) return null;

  const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
  const TIME_RE = /^([01]?\d|2[0-3]):[0-5]\d$/;
  const validEntries = [];

  data.entries.slice(0, AI_PARSE_MAX_ENTRIES).forEach((e, i) => {
    const label = `QuickFill entry #${i + 1}`;
    if (!e || typeof e !== 'object') { warnings.push(`${label} was malformed and dropped.`); return; }
    if (typeof e.date !== 'string' || !DATE_RE.test(e.date)) { warnings.push(`${label} had an invalid date and was dropped.`); return; }
    if (typeof e.startTime !== 'string' || !TIME_RE.test(e.startTime) || typeof e.endTime !== 'string' || !TIME_RE.test(e.endTime)) {
      warnings.push(`${label} had invalid times and was dropped.`);
      return;
    }

    const emps = Array.isArray(e.employees) ? e.employees.map(n => String(n || '').trim()).filter(Boolean) : [];
    if (!emps.length) { warnings.push(`${label} had no employees and was dropped.`); return; }

    validEntries.push({
      date: e.date,
      employees: emps,
      startTime: padTime(e.startTime),
      endTime: padTime(e.endTime),
      jobOrder: (typeof e.jobOrder === 'string') ? e.jobOrder.trim() : '',
      confidence: ['high', 'medium', 'low'].includes(e.confidence) ? e.confidence : 'medium',
      note: (typeof e.note === 'string') ? e.note.trim() : ''
    });
  });

  if (data.entries.length > AI_PARSE_MAX_ENTRIES) {
    warnings.push(`Only the first ${AI_PARSE_MAX_ENTRIES} entries were kept.`);
  }

  return { validEntries, warnings };
}


// ---------------------------------------------------------------------
// NATURAL-LANGUAGE ENTRY PARSERS (TEXT & IMAGE)
// ---------------------------------------------------------------------

function parseNaturalLanguageEntries(text, sheetName) {
  try {
    if (!text || String(text).trim() === '') {
      return { success: false, error: 'Paste some shift notes first.' };
    }
    const input = String(text).trim().slice(0, AI_PARSE_MAX_INPUT_CHARS);
    const contextAndRules = getContextAndRules(sheetName);
    const prompt = `${contextAndRules}\n\nINPUT NOTES:\n"""\n${input}\n"""`;

    const result = callCloudflareAiJson(prompt, 0.2);
    if (!result.success) return { success: false, error: result.error };

    const initialWarnings = Array.isArray(result.data.warnings)
      ? result.data.warnings.filter(w => typeof w === 'string' && w.trim() !== '')
      : [];

    const validation = validateAndCleanEntries(result.data, initialWarnings);
    if (!validation) {
      return { success: false, error: 'QuickFill response did not contain an "entries" list. Try parsing again.' };
    }

    if (validation.validEntries.length === 0) {
      return {
        success: false,
        error: 'QuickFill could not extract any complete entries.\n' + validation.warnings.join('\n')
      };
    }

    return { success: true, entries: validation.validEntries, warnings: validation.warnings };
  } catch (err) {
    Logger.log('parseNaturalLanguageEntries error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}


// ---------------------------------------------------------------------
// VISION (IMAGE) API — Multimodal Cloudflare model for screenshot parsing
// ---------------------------------------------------------------------

function callGatewayChatCompletion(url, apiToken, model, messages, temperature) {
  const payload = {
    model: model,
    messages: messages,
    temperature: (typeof temperature === 'number') ? temperature : 0.2,
    max_tokens: 6000,
    response_format: { type: 'json_object' }
  };
  const headers = { 'Authorization': 'Bearer ' + apiToken };
  return executeFetchWithRetry(url, headers, payload, model);
}

function callCloudflareAiVision(messages, temperature) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accountId = scriptProperties.getProperty(CLOUDFLARE_ACCOUNT_ID_PROPERTY);
  const apiToken = scriptProperties.getProperty(CLOUDFLARE_API_TOKEN_PROPERTY);

  if (!accountId || !apiToken) {
    return { success: false, error: 'No Cloudflare credentials configured.' };
  }

  const url = 'https://gateway.ai.cloudflare.com/v1/' + accountId + '/' + CLOUDFLARE_GATEWAY_ID + '/compat/chat/completions';

  const primary = callGatewayChatCompletion(url, apiToken, GEMINI_VISION_MODEL_PRIMARY, messages, temperature);
  if (primary.success) return primary;

  Logger.log(GEMINI_VISION_MODEL_PRIMARY + ' failed, falling back to ' + GEMINI_VISION_MODEL_FALLBACK + ': ' + primary.error);
  return callGatewayChatCompletion(url, apiToken, GEMINI_VISION_MODEL_FALLBACK, messages, temperature);
}

function parseNaturalLanguageEntriesWithImage(text, imageBase64, mimeType, sheetName) {
  try {
    if (!imageBase64) return { success: false, error: 'No image data received.' };
    if (imageBase64.length > AI_PARSE_MAX_IMAGE_SIZE_BYTES) {
      return { success: false, error: 'Image too large. Try a smaller screenshot.' };
    }

    const userText = String(text || '').trim().slice(0, AI_PARSE_MAX_INPUT_CHARS);
    const contextAndRules = getContextAndRules(sheetName);

    const systemMsg = 'You are a precise JSON-only utility assistant. You must output raw, valid JSON only. Do not use markdown code blocks (fences).';
    const userMsg = contextAndRules + (userText ? '\n\nADDITIONAL TEXT FROM USER:\n"""\n' + userText + '\n"""' : '');
    const imageDataUrl = 'data:' + mimeType + ';base64,' + imageBase64;

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

    // Uses the gateway-routed fallback model architecture
    const result = callCloudflareAiVision(messages, 0.2);
    if (!result.success) return { success: false, error: result.error };

    const initialWarnings = Array.isArray(result.data.warnings)
      ? result.data.warnings.filter(w => typeof w === 'string' && w.trim() !== '')
      : [];
    if (userText) {
      initialWarnings.push('QuickFill analyzed the screenshot combined with your text notes. Please verify all entries.');
    }

    const validation = validateAndCleanEntries(result.data, initialWarnings);
    if (!validation) {
      return { success: false, error: 'QuickFill response did not contain an "entries" list. Try again.' };
    }

    if (validation.validEntries.length === 0) {
      return {
        success: false,
        error: 'QuickFill could not extract any complete entries from that image.\n' + validation.warnings.join('\n')
      };
    }

    return { success: true, entries: validation.validEntries, warnings: validation.warnings };
  } catch (err) {
    Logger.log('parseNaturalLanguageEntriesWithImage error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}


// ---------------------------------------------------------------------
// UTILITY FUNCTIONS
// ---------------------------------------------------------------------

function extractMessageText(content) {
  if (!content) return '';
  if (typeof content === 'string') return content.trim();
  if (Array.isArray(content)) {
    return content
      .map(part => (part && typeof part.text === 'string') ? part.text : '')
      .join('')
      .trim();
  }
  return '';
}

function padTime(t) {
  const parts = String(t).split(':');
  return ('0' + parts[0]).slice(-2) + ':' + parts[1];
}
