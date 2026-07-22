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
// QUICKFILL PROMPT TEMPLATE
// ---------------------------------------------------------------------

/**
 * Template for the QuickFill prompt. Placeholders are replaced with dynamic context at runtime.
 */
const QUICKFILL_PROMPT_TEMPLATE = `You are an expert timesheet parser for a contractor in Qatar. Your ONLY job is to convert unstructured shift notes into strict JSON.
STRICT REQUIREMENTS:
- Output ONLY valid JSON. No explanations, apologies, or markdown.
- If you cannot parse an entry, SKIP IT and add a warning explaining why.

CONTEXT:
- Today is {today} ({weekday}).
- Yesterday was {yesterday}. Tomorrow is {tomorrow}.
- {sheetPeriodLine}
- Known employees: {employeeNames}.
  Fuzzy match names to this list (case-insensitive). If no close match, use the original name and set confidence="low".
- Known job orders: {jobOrders}.
  Job orders may contain spaces, hyphens, or slashes (e.g., "JO-123", "Project Alpha - Phase 1").

--- PARSING RULES ---
### DATES:
- Default year: {defaultYear}.
- Default month: {defaultMonth}.
- Accept these date formats (convert all to YYYY-MM-DD):
  - DD/MM/YYYY or DD-MM-YYYY (e.g., "13/07/2026" -> "2026-07-13")
  - MM/DD/YYYY (e.g., "07/13/2026" -> "2026-07-13")
  - D.M.YYYY or D-M-YYYY (e.g., "13.7.2026" -> "2026-07-13")
  - Day names: "Monday", "Mon", "Today" -> {today}, "Yesterday" -> {yesterday}, "Tomorrow" -> {tomorrow}
  - Day numbers: "13th", "13", "thirteenth" -> Resolve to the current month/year (or sheet month if specified).
  - Month names: "July 13" -> "{defaultYear}-07-13" (use sheet year if available).
  - Relative: "next Monday", "last Friday" -> Calculate from {today}.
- If date cannot be resolved, SKIP the entry and warn.

### TIMES:
- Output ALWAYS in 24-hour "HH:mm" format (e.g., "08:00", "17:30").
- Accept these input formats:
  - "8am-4pm" -> "08:00"-"16:00"
  - "8:00 AM to 5:00 PM" -> "08:00"-"17:00"
  - "8-5" -> "08:00"-"17:00" (assume AM-PM if no suffix)
  - "0800-1600" -> "08:00"-"16:00"
  - "8:00 to 17:00" -> "08:00"-"17:00"
  - "8am" -> "08:00"
  - "1700" -> "17:00"
- If time is invalid (e.g., "25:00"), SKIP the entry and warn.

### EMPLOYEES:
- Match names to {employeeNames} CASE-INSENSITIVELY.
  - If no match within 2 edits (Levenshtein), use the original name and set confidence="low".
- Handle lists:
  - "Name1, Name2" -> ["Name1", "Name2"]
  - "Name1 & Name2" -> ["Name1", "Name2"]
  - "Name1 + Name2" -> ["Name1", "Name2"]
- If no employees found, SKIP the entry and warn.

### JOB ORDERS:
- Extract from:
  - Inline: "Name1 8am-4pm JO-123" -> jobOrder: "JO-123"
  - Header line: "Job: JO-123" -> Apply "JO-123" to all subsequent entries until another header.
  - Trailing: "JO-123" on a line after employees -> Apply retroactively to previous entries.
- Multi-word job orders: Preserve full text (e.g., "Project Alpha - Phase 1").
- If no job order, use empty string "".

### GROUPING:
- MANDATORY: if two or more people share the SAME date, startTime, endTime, AND jobOrder (including "no job order" for all of them), they MUST appear together in ONE entry's "employees" array — this rule applies REGARDLESS of whether the input used a numbered list, a bullet list, or separate lines. List formatting only signals "these are distinct people/shifts to parse," not "these must stay in separate output entries."
- Numbered lists (e.g., "1. Name1 8am-4pm") and bullet lists (e.g., "- Name1 8am-4pm") each describe one person's shift to extract — but if two such lines end up with identical date/startTime/endTime/jobOrder, merge them into one entry with both names in "employees". Never emit two entries for an identical date/start/end/job combination.

### CONFIDENCE & NOTES:
- confidence:
  - "high": Exact match to known employees/job orders, unambiguous times/dates.
  - "medium": Minor guesses (e.g., time format conversion, fuzzy employee match).
  - "low": Major assumptions (e.g., resolved ambiguous date, no employee match).
- note: Explain any guesses.

### EDGE CASES:
- Lines starting with "#" or "//" are COMMENTS -> Ignore.
- Lines with only a date (e.g., "13.7.26") -> Apply to subsequent entries until another date.
- Lines with only a job order (e.g., "JO-123") -> Apply to subsequent entries until another job order.
- "All" or "Everyone" -> Use ALL known employees for that entry.
- "OT" or "Overtime" in job order -> Preserve as-is.
- Shifts crossing midnight (e.g., "10pm-6am") -> Keep as-is; do NOT adjust times.

### OUTPUT STRUCTURE:
Return ONLY this JSON format:
{
  "entries": [
    {
      "date": "YYYY-MM-DD",
      "employees": ["NAME1", "NAME2"],
      "startTime": "HH:mm",
      "endTime": "HH:mm",
      "jobOrder": "JO-123",
      "confidence": "high" | "medium" | "low",
      "note": "Explanation if confidence is not high"
    }
  ],
  "warnings": []
}`;

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
 * Calculates Levenshtein distance between two strings (for fuzzy matching).
 */
function levenshteinDistance(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  const matrix = [];

  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

/**
 * Returns dynamic context for the sheet (employees, job orders, dates).
 * This data is injected into the LLM prompt at runtime.
 */
function getSheetContext(sheetName) {
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
  const yesterdayStr = Utilities.formatDate(new Date(now.getTime() - 24 * 60 * 60 * 1000), tz, 'yyyy-MM-dd');
  const tomorrowStr = Utilities.formatDate(new Date(now.getTime() + 24 * 60 * 60 * 1000), tz, 'yyyy-MM-dd');
  const weekday = Utilities.formatDate(now, tz, 'EEEE');

  const detected = parseSheetNameForMonthYear(sheetName);
  const sheetPeriodLine = detected
    ? `The target sheet "${sheetName}" covers ${Utilities.formatDate(new Date(detected.year, detected.monthIndex, 1), tz, 'MMMM yyyy')}.`
    : `The target sheet is "${sheetName}".`;

  return {
    todayStr,
    yesterdayStr,
    tomorrowStr,
    weekday,
    sheetPeriodLine,
    employeeNames,
    jobOrders,
    detected,
    defaultYear: detected ? detected.year : now.getFullYear(),
    defaultMonth: detected ? MONTHS[detected.monthIndex].full : Utilities.formatDate(now, tz, 'MMMM')
  };
}

// Context cache for performance (5-minute TTL)
const CONTEXT_CACHE_TTL_MS = 5 * 60 * 1000;
const contextCache = {};

/**
 * Gets sheet context with caching to avoid repeated fetches.
 */
function getSheetContextCached(sheetName) {
  const cached = contextCache[sheetName];
  if (cached && (Date.now() - cached.timestamp) < CONTEXT_CACHE_TTL_MS) {
    return cached.data;
  }

  const context = getSheetContext(sheetName);
  contextCache[sheetName] = { data: context, timestamp: Date.now() };
  return context;
}

/**
 * Validates parsed entries, sanitizes data types, and normalizes times.
 * Also validates employees against the current sheet and performs fuzzy matching.
 */
function validateAndCleanEntries(data, initialWarnings, sheetName) {
  const warnings = Array.isArray(initialWarnings) ? initialWarnings : [];
  if (!data || !Array.isArray(data.entries)) return null;

  const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
  const TIME_RE = /^([01]?\d|2[0-3]):[0-5]\d$/;
  const validEntries = [];

  // Fetch current employees for the sheet to validate against
  let validEmployeeNames = [];
  let validEmployeeNamesLower = [];
  if (sheetName) {
    const empResult = getEmployeesForSheet(sheetName);
    if (!empResult.error) {
      validEmployeeNames = empResult.employees.map(e => e.name);
      validEmployeeNamesLower = validEmployeeNames.map(n => n.toLowerCase());
    }
  }

  data.entries.slice(0, AI_PARSE_MAX_ENTRIES).forEach((e, i) => {
    const label = `QuickFill entry #${i + 1}`;
    if (!e || typeof e !== 'object') {
      warnings.push(`${label} was malformed and dropped.`);
      return;
    }

    // --- Date validation ---
    if (typeof e.date !== 'string' || !DATE_RE.test(e.date)) {
      warnings.push(`${label} had an invalid date "${e.date}" and was dropped.`);
      return;
    }

    // --- Time validation ---
    if (typeof e.startTime !== 'string' || !TIME_RE.test(e.startTime)) {
      warnings.push(`${label} had an invalid start time "${e.startTime}" and was dropped.`);
      return;
    }
    if (typeof e.endTime !== 'string' || !TIME_RE.test(e.endTime)) {
      warnings.push(`${label} had an invalid end time "${e.endTime}" and was dropped.`);
      return;
    }

    // --- Employee validation ---
    const emps = Array.isArray(e.employees) ? e.employees.map(n => String(n || '').trim()).filter(Boolean) : [];
    if (!emps.length) {
      warnings.push(`${label} had no employees and was dropped.`);
      return;
    }

    // Validate each employee against the sheet's current list and fuzzy match
    const validatedEmps = [];
    for (const emp of emps) {
      const lowerEmp = emp.toLowerCase();

      // Check for exact match (case-insensitive)
      const exactMatchIndex = validEmployeeNamesLower.indexOf(lowerEmp);
      if (exactMatchIndex !== -1) {
        validatedEmps.push(validEmployeeNames[exactMatchIndex]);
        continue;
      }

      // Check for "All" or "Everyone"
      if (lowerEmp === 'all' || lowerEmp === 'everyone') {
        validatedEmps.push(...validEmployeeNames);
        continue;
      }

      // Fuzzy match: find the closest match using Levenshtein distance
      if (validEmployeeNames.length > 0) {
        let bestMatch = null;
        let bestDistance = Infinity;
        for (let j = 0; j < validEmployeeNames.length; j++) {
          const distance = levenshteinDistance(lowerEmp, validEmployeeNamesLower[j]);
          if (distance < bestDistance) {
            bestDistance = distance;
            bestMatch = validEmployeeNames[j];
          }
        }
        if (bestDistance <= 2 && bestMatch) {
          warnings.push(`Fuzzy matched '${emp}' to '${bestMatch}' (confidence: medium).`);
          validatedEmps.push(bestMatch);
        } else {
          warnings.push(`Employee '${emp}' not found in sheet '${sheetName}' and was dropped.`);
        }
      } else {
        // No valid employees in sheet, use original name
        validatedEmps.push(emp);
      }
    }

    if (validatedEmps.length === 0) {
      warnings.push(`${label} had no valid employees and was dropped.`);
      return;
    }

    // --- Job order validation ---
    let jobOrder = '';
    if (typeof e.jobOrder === 'string') {
      jobOrder = e.jobOrder.trim();
      if (jobOrder.length > 200) {
        warnings.push(`${label} had a job order longer than 200 characters (truncated).`);
        jobOrder = jobOrder.substring(0, 200);
      }
    }

    // --- Confidence validation ---
    const confidence = ['high', 'medium', 'low'].includes(e.confidence)
      ? e.confidence
      : 'medium';

    // --- Note validation ---
    const note = (typeof e.note === 'string') ? e.note.trim().substring(0, 500) : '';

    validEntries.push({
      date: e.date,
      employees: validatedEmps,
      startTime: padTime(e.startTime),
      endTime: padTime(e.endTime),
      jobOrder: jobOrder,
      confidence: confidence,
      note: note
    });
  });

  if (data.entries.length > AI_PARSE_MAX_ENTRIES) {
    warnings.push(`Only the first ${AI_PARSE_MAX_ENTRIES} entries were kept.`);
  }

  return { validEntries: mergeEntriesSharingShift_(validEntries), warnings };
}

/**
 * Deterministic safety net for grouping — does NOT rely on the model
 * obeying the "group same date/time/job into one entry" prompt rule.
 * The prompt asks for grouping already (see the GROUPING section of
 * QUICKFILL_PROMPT_TEMPLATE), but prompt compliance from an LLM is never
 * guaranteed — this is especially true when the input is a numbered or
 * bulleted list, since the model can end up treating "one line per person"
 * as "one output entry per person" even when their date/time/job all
 * match. This runs AFTER employee validation/fuzzy-matching, so it merges
 * on the final, already-resolved employee names — entries sharing the same
 * date + startTime + endTime + jobOrder (job order compared
 * case-insensitively; "" and "no job" bucket together) are combined into
 * one entry with a deduplicated "employees" array regardless of how the
 * model split them. Confidence is downgraded to the lowest of the merged
 * set, and per-entry notes are combined so nothing silently disappears.
 */
function mergeEntriesSharingShift_(entries) {
  const CONFIDENCE_RANK = { high: 2, medium: 1, low: 0 };
  const byKey = {};
  const order = [];

  entries.forEach(e => {
    const jobKey = (e.jobOrder || '').trim().toLowerCase();
    const key = e.date + '|' + e.startTime + '|' + e.endTime + '|' + jobKey;

    if (!byKey[key]) {
      byKey[key] = {
        date: e.date,
        employees: [],
        startTime: e.startTime,
        endTime: e.endTime,
        jobOrder: e.jobOrder,
        confidence: e.confidence,
        note: e.note
      };
      order.push(key);
    }

    const bucket = byKey[key];
    e.employees.forEach(name => {
      if (!bucket.employees.includes(name)) bucket.employees.push(name);
    });
    if (CONFIDENCE_RANK[e.confidence] < CONFIDENCE_RANK[bucket.confidence]) {
      bucket.confidence = e.confidence;
    }
    if (e.note && e.note !== bucket.note) {
      bucket.note = bucket.note ? bucket.note + ' | ' + e.note : e.note;
    }
  });

  return order.map(key => byKey[key]);
}


// ---------------------------------------------------------------------
// NATURAL-LANGUAGE ENTRY PARSERS (TEXT & IMAGE)
// ---------------------------------------------------------------------

function parseNaturalLanguageEntries(text, sheetName) {
  try {
    if (!text || String(text).trim() === '') {
      return { success: false, error: 'Paste some shift notes first.' };
    }

    // Get dynamic context for the sheet
    const context = getSheetContextCached(sheetName);
    const input = String(text).trim().slice(0, AI_PARSE_MAX_INPUT_CHARS);

    // Replace placeholders in the template with dynamic context
    const prompt = QUICKFILL_PROMPT_TEMPLATE
      .replace(/{today}/g, context.todayStr)
      .replace(/{yesterday}/g, context.yesterdayStr)
      .replace(/{tomorrow}/g, context.tomorrowStr)
      .replace(/{weekday}/g, context.weekday)
      .replace(/{sheetPeriodLine}/g, context.sheetPeriodLine)
      .replace(/{employeeNames}/g, JSON.stringify(context.employeeNames))
      .replace(/{jobOrders}/g, JSON.stringify(context.jobOrders))
      .replace(/{defaultYear}/g, context.defaultYear)
      .replace(/{defaultMonth}/g, context.defaultMonth)
      + `\n\nINPUT NOTES:\n"""\n${input}\n"""`;

    const result = callCloudflareAiJson(prompt, 0.2);
    if (!result.success) return { success: false, error: result.error };

    const initialWarnings = Array.isArray(result.data.warnings)
      ? result.data.warnings.filter(w => typeof w === 'string' && w.trim() !== '')
      : [];

    const validation = validateAndCleanEntries(result.data, initialWarnings, sheetName);
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

    // Get dynamic context for the sheet
    const context = getSheetContextCached(sheetName);
    const userText = String(text || '').trim().slice(0, AI_PARSE_MAX_INPUT_CHARS);

    // Replace placeholders in the template with dynamic context
    const userMsg = QUICKFILL_PROMPT_TEMPLATE
      .replace(/{today}/g, context.todayStr)
      .replace(/{yesterday}/g, context.yesterdayStr)
      .replace(/{tomorrow}/g, context.tomorrowStr)
      .replace(/{weekday}/g, context.weekday)
      .replace(/{sheetPeriodLine}/g, context.sheetPeriodLine)
      .replace(/{employeeNames}/g, JSON.stringify(context.employeeNames))
      .replace(/{jobOrders}/g, JSON.stringify(context.jobOrders))
      .replace(/{defaultYear}/g, context.defaultYear)
      .replace(/{defaultMonth}/g, context.defaultMonth)
      + (userText ? '\n\nADDITIONAL TEXT FROM USER:\n"""\n' + userText + '\n"""' : '');

    const imageDataUrl = 'data:' + mimeType + ';base64,' + imageBase64;

    const messages = [
      { role: 'system', content: 'You are a precise JSON-only utility assistant. You must output raw, valid JSON only. Do not use markdown code blocks (fences).' },
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

    const validation = validateAndCleanEntries(result.data, initialWarnings, sheetName);
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