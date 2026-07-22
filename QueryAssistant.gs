/**
 * QueryAssistant.gs — Multi-turn natural-language timesheet queries
 * ------------------------------------------------------------------
 * Uses OpenAI multi-turn format. Frontend sends the full messages[]
 * array; backend:
 *   1. Extracts structured intent (action, filters, dates) via AI
 *   2. Queries sheet data using existing functions
 *   3. Generates a natural-language answer via AI
 *
 * Reuses: callCloudflareAiJson (ai.gs), generateOtReport (ReportCode.gs),
 * listTimesheetEntries (EntriesList.gs).
 */

/**
 * @param {Object} payload  { messages: [{ role: "user"|"assistant", content }], context?: { currentDateRange, currentFilters } }
 * @returns {Object} { success, role: "assistant", content, entries?, report? }
 */
function askTimesheetQuery(payload) {
  try {
    if (!checkAiCredentials()) {
      return { success: false, error: 'AI not configured.' };
    }

    var messages = payload.messages;
    if (!messages || !messages.length) {
      return { success: false, error: 'No messages provided.' };
    }

    // Step 1 — extract structured intent from the conversation, with optional context
    var intent = extractQueryIntent(messages, payload.context);
    if (intent.error) return { success: false, error: intent.error };

    // Step 2 — fetch data from sheets
    var queryResult = executeDataQuery(intent);

    // Step 3 — generate natural-language answer
    var answer = generateQueryAnswer(messages, intent, queryResult);

    return {
      success: true,
      role: 'assistant',
      content: answer.content,
      entries: answer.entries || null,
      report: answer.report || null
    };
  } catch (err) {
    Logger.log('askTimesheetQuery error: ' + err);
    return { success: false, error: err.message || String(err) };
  }
}

// ---------------------------------------------------------------------------
// Step 1 — Intent extraction
// ---------------------------------------------------------------------------

function extractQueryIntent(messages, context) {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var weekday = Utilities.formatDate(new Date(), tz, 'EEEE');

  // Build context information for the prompt
  var contextInfo = '';
  if (context && context.currentDateRange) {
    contextInfo += 'Current view context: Date range is ' + context.currentDateRange + '. ';
  }
  if (context && context.currentFilters) {
    var filters = context.currentFilters;
    var filterParts = [];
    if (filters.employees && filters.employees.length > 0) {
      filterParts.push('Employees: ' + filters.employees.join(', '));
    }
    if (filters.jobOrder) {
      filterParts.push('Job Order: ' + filters.jobOrder);
    }
    if (filters.otOnly) {
      filterParts.push('OT Only: true');
    }
    if (filters.minHours) {
      filterParts.push('Min Hours: ' + filters.minHours);
    }
    if (filters.maxHours) {
      filterParts.push('Max Hours: ' + filters.maxHours);
    }
    if (filters.search) {
      filterParts.push('Search: ' + filters.search);
    }
    if (filterParts.length > 0) {
      contextInfo += 'Current filters: ' + filterParts.join('; ') + '. ';
    }
  }

  var systemPrompt = [
    'You extract structured query intent from timesheet questions. Return ONLY valid JSON.',
    '',
    'Today is ' + today + ' (' + weekday + ').',
    contextInfo ? contextInfo : '',
    'If the user query is vague (e.g., "show me entries", "what about today"), use the current view context as the default filters.',
    'Available actions: list_entries | get_report',
    'Date format: yyyy-MM-dd. Interpret relative dates ("last week", "this month") relative to today.',
    '',
    'Return: {',
    '  "action": "list_entries|get_report",',
    '  "startDate": "yyyy-MM-dd",',
    '  "endDate": "yyyy-MM-dd",',
    '  "employees": ["name"] or [],',
    '  "jobOrder": "string or empty",',
    '  "minHours": number or null,',
    '  "maxHours": number or null,',
    '  "search": "string or empty",',
    '  "question": "rephrase the user intent as a clear question"',
    '}',
    '',
    'Examples:',
    '- "Show me Rasheedh entries from last week" -> { action:"list_entries", startDate:"2026-07-13", endDate:"2026-07-19", employees:["Rasheedh"] }',
    '- "Who worked over 10 hours yesterday?" -> { action:"list_entries", minHours:10 }',
    '- "How much OT this month?" -> { action:"get_report" } with month range',
    '- "What about Ravi?" -> inherit context from prior assistant messages, keep same dates, change employee to ["Ravi"]',
    '- "Show me entries" (with current context) -> use current date range and filters'
  ].join('\n');

  var extractMessages = [{ role: 'system', content: systemPrompt }].concat(
    messages.map(function (m) { return { role: m.role, content: m.content }; })
  );

  var result = callCloudflareAiJson(null, 0.1, extractMessages);
  if (!result.success) return { error: 'Could not understand the question: ' + result.error };

  var intent = result.data;
  if (!intent.action) return { error: 'Could not determine what to look up from that question.' };

  // Default date range (current month) if not specified
  if (!intent.startDate) {
    var now = new Date();
    intent.startDate = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), tz, 'yyyy-MM-dd');
  }
  if (!intent.endDate) {
    intent.endDate = today;
  }

  // If context has date range and intent doesn't specify, use context
  if (context && context.currentDateRange && !intent.startDate && !intent.endDate) {
    var dateParts = context.currentDateRange.split(' to ');
    if (dateParts.length === 2) {
      intent.startDate = dateParts[0];
      intent.endDate = dateParts[1];
    }
  }

  // If context has filters and intent doesn't override them, use context
  if (context && context.currentFilters) {
    var ctxFilters = context.currentFilters;
    if (!intent.employees && ctxFilters.employees) {
      intent.employees = ctxFilters.employees;
    }
    if (!intent.jobOrder && ctxFilters.jobOrder) {
      intent.jobOrder = ctxFilters.jobOrder;
    }
    if (intent.minHours === undefined && ctxFilters.minHours !== undefined) {
      intent.minHours = ctxFilters.minHours;
    }
    if (intent.maxHours === undefined && ctxFilters.maxHours !== undefined) {
      intent.maxHours = ctxFilters.maxHours;
    }
    if (!intent.search && ctxFilters.search) {
      intent.search = ctxFilters.search;
    }
    if (intent.otOnly === undefined && ctxFilters.otOnly !== undefined) {
      intent.otOnly = ctxFilters.otOnly;
    }
  }

  return intent;
}

// ---------------------------------------------------------------------------
// Step 2 — Data query execution
// ---------------------------------------------------------------------------

function executeDataQuery(intent) {
  if (intent.action === 'get_report') {
    try {
      var report = generateOtReport(intent.startDate, intent.endDate, intent.employees || []);
      return { type: 'report', data: report };
    } catch (e) {
      return { type: 'error', message: e.message };
    }
  }

  // Default: list_entries
  try {
    var result = listTimesheetEntries({
      startDate: intent.startDate,
      endDate: intent.endDate,
      employees: intent.employees && intent.employees.length > 0 ? intent.employees : undefined,
      jobOrder: intent.jobOrder || undefined,
      minHours: intent.minHours || undefined,
      maxHours: intent.maxHours || undefined,
      search: intent.search || undefined,
      page: 1,
      pageSize: 1000
    });
    return { type: 'entries', data: result };
  } catch (e) {
    return { type: 'error', message: e.message };
  }
}

// ---------------------------------------------------------------------------
// Step 3 — Answer generation
// ---------------------------------------------------------------------------

// Deterministic per-employee rollup, so the model narrates facts instead of counting raw rows itself.
function summarizeEntriesByEmployee(entries) {
  var byEmployee = {};
  entries.forEach(function (e) {
    var key = e.employeeName;
    if (!byEmployee[key]) {
      byEmployee[key] = { employeeName: key, daysWorked: 0, totalHours: 0, otHours: 0, dates: {} };
    }
    var bucket = byEmployee[key];
    if (!bucket.dates[e.date]) {
      bucket.dates[e.date] = true;
      bucket.daysWorked++;
    }
    bucket.totalHours += e.totalHours || 0;
    bucket.otHours += e.otHours || 0;
  });
  return Object.keys(byEmployee).map(function (key) {
    var b = byEmployee[key];
    return {
      employeeName: b.employeeName,
      daysWorked: b.daysWorked,
      totalHours: Math.round(b.totalHours * 100) / 100,
      otHours: Math.round(b.otHours * 100) / 100
    };
  }).sort(function (a, b) { return b.daysWorked - a.daysWorked; });
}

function generateQueryAnswer(conversation, intent, queryResult) {
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var dataBlock;
  if (queryResult.type === 'entries') {
    var entries = queryResult.data.entries || [];
    var summary = summarizeEntriesByEmployee(entries);
    dataBlock = 'Retrieved ' + entries.length + ' entries from ' + intent.startDate + ' to ' + intent.endDate + '.\n\n' +
      'EMPLOYEE SUMMARY (pre-computed, authoritative — use these numbers for any counting/comparison question):\n' +
      JSON.stringify(summary, null, 2) +
      '\n\nSample raw entries (first 30 of ' + entries.length + '):\n' +
      JSON.stringify(entries.slice(0, 30), null, 2);
  } else if (queryResult.type === 'report') {
    dataBlock = 'Report ' + intent.startDate + ' to ' + intent.endDate +
      ':\n\n' + JSON.stringify(queryResult.data, null, 2);
  } else {
    dataBlock = 'Error fetching data: ' + queryResult.message;
  }

  var systemPrompt = [
    'You are a timesheet data assistant. Answer concisely using the data provided.',
    'Today is ' + todayStr + '.',
    '',
    'DATA:',
    dataBlock,
    '',
    'Return ONLY valid JSON: { "answer": "your natural language response here" }',
    'Rules: Use numbers. Summarize clearly. If no data matches, say so. Do not output JSON in the answer field.',
    'Format the answer field as Markdown. When listing multiple entries or multiple employees, use a Markdown table with a header row (e.g. | Employee | Date | Hours |). Use **bold** for key totals.'
  ].join('\n');

  var answerMessages = [{ role: 'system', content: systemPrompt }].concat(
    conversation.map(function (m) { return { role: m.role, content: m.content }; })
  );

  var result = callCloudflareAiJson(null, 0.3, answerMessages);
  var content = result.success
    ? (result.data.answer || String(result.data))
    : 'Sorry, I could not generate an answer right now.';

  return {
    content: content,
    entries: queryResult.type === 'entries' ? (queryResult.data.entries || []) : null,
    report: queryResult.type === 'report' ? queryResult.data : null
  };
}
