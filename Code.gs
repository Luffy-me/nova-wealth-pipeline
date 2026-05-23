const SHEET_NAMES = {
  transactions: 'Transactions',
  budget: 'Budget',
  reports: 'Reports',
  settings: 'Settings',
};

const SHEET_HEADERS = {
  transactions: ['Date', 'Type', 'Amount', 'Note', 'Status'],
  budget: ['Category', 'Planned', 'Actual'],
  reports: ['Date', 'Proposer', 'Critic', 'Synthesis'],
  settings: ['Key', 'Value'],
};

const DEFAULT_SETTINGS = {
  TargetCapital: 1000000,
  CurrentCapital: 250000,
  DaysRemaining: 180,
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Nova Wealth OS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData() {
  const sheets = ensureSheets_();
  const settings = getSettings_(sheets.settings);
  const transactions = getTransactions_(sheets.transactions);
  const budget = getBudget_(sheets.budget);
  const reports = getReports_(sheets.reports);
  const metrics = computeMetrics_(settings, transactions);

  return {
    settings,
    metrics,
    transactions,
    budget,
    reports,
  };
}

function logFlow(payload) {
  const sheets = ensureSheets_();
  const settings = getSettings_(sheets.settings);
  const transactions = getTransactions_(sheets.transactions);
  const metrics = computeMetrics_(settings, transactions);

  const entryDate = payload.date ? new Date(payload.date) : new Date();
  const type = payload.type || 'INFLOW';
  const amount = normalizeAmount_(payload.amount, type);
  const note = payload.note || '';
  const limit = metrics.dailyLimit;
  const status =
    Math.abs(amount) > limit
      ? 'REJECTED: STATISTICAL DEVIATION DETECTED'
      : 'ACCEPTED';

  sheets.transactions.appendRow([entryDate, type, amount, note, status]);

  if (status === 'ACCEPTED') {
    const updatedCapital = toNumber_(settings.CurrentCapital) + amount;
    setSetting_(sheets.settings, 'CurrentCapital', updatedCapital);
    settings.CurrentCapital = updatedCapital;
  }

  return getAppData();
}

function updateGoal(targetCapital, daysRemaining) {
  const sheets = ensureSheets_();
  if (targetCapital !== undefined && targetCapital !== null && targetCapital !== '') {
    setSetting_(sheets.settings, 'TargetCapital', toNumber_(targetCapital));
  }
  if (daysRemaining !== undefined && daysRemaining !== null && daysRemaining !== '') {
    setSetting_(sheets.settings, 'DaysRemaining', Math.max(0, toNumber_(daysRemaining)));
  }
  return getAppData();
}

function processChatMessage(message) {
  try {
    const data = getAppData();
    const context = buildContext_(data);
    const proposerPrompt = buildProposerPrompt_(context, message);
    const proposer = callGemini_(proposerPrompt);

    const criticPrompt = buildCriticPrompt_(context, message, proposer);
    const critic = callGroq_(criticPrompt);

    const synthPrompt = buildSynthPrompt_(context, message, proposer, critic);
    const synthesis = callGemini_(synthPrompt);

    const sheets = ensureSheets_();
    sheets.reports.appendRow([new Date(), proposer, critic, synthesis]);

    return {
      proposer,
      critic,
      synthesis,
    };
  } catch (error) {
    return {
      error: error.message || 'Debate engine failed.',
    };
  }
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    transactions: getOrCreateSheet_(ss, SHEET_NAMES.transactions, SHEET_HEADERS.transactions),
    budget: getOrCreateSheet_(ss, SHEET_NAMES.budget, SHEET_HEADERS.budget),
    reports: getOrCreateSheet_(ss, SHEET_NAMES.reports, SHEET_HEADERS.reports),
    settings: getOrCreateSheet_(ss, SHEET_NAMES.settings, SHEET_HEADERS.settings),
  };
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function getSettings_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    Object.keys(DEFAULT_SETTINGS).forEach((key) => {
      sheet.appendRow([key, DEFAULT_SETTINGS[key]]);
    });
  }
  const rows = sheet.getDataRange().getValues().slice(1);
  const settings = {};
  rows.forEach(([key, value]) => {
    if (key) {
      settings[key] = value;
    }
  });
  Object.keys(DEFAULT_SETTINGS).forEach((key) => {
    if (settings[key] === undefined || settings[key] === '') {
      settings[key] = DEFAULT_SETTINGS[key];
      setSetting_(sheet, key, DEFAULT_SETTINGS[key]);
    }
  });
  return settings;
}

function setSetting_(sheet, key, value) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getTransactions_(sheet) {
  const values = sheet.getDataRange().getValues().slice(1);
  return values
    .filter((row) => row[0])
    .map(([date, type, amount, note, status]) => ({
      date: formatDate_(date),
      rawDate: date,
      type,
      amount: toNumber_(amount),
      note,
      status: status || 'ACCEPTED',
    }));
}

function getBudget_(sheet) {
  const values = sheet.getDataRange().getValues().slice(1);
  return values
    .filter((row) => row[0])
    .map(([category, planned, actual]) => ({
      category,
      planned: toNumber_(planned),
      actual: toNumber_(actual),
    }));
}

function getReports_(sheet) {
  const values = sheet.getDataRange().getValues().slice(1);
  return values
    .filter((row) => row[0])
    .slice(-20)
    .map(([date, proposer, critic, synthesis]) => ({
      date: formatDate_(date),
      proposer,
      critic,
      synthesis,
    }));
}

function computeMetrics_(settings, transactions) {
  const target = toNumber_(settings.TargetCapital);
  const current = toNumber_(settings.CurrentCapital);
  const daysRemaining = Math.max(0, Math.floor(toNumber_(settings.DaysRemaining)));
  const dailyLimit = daysRemaining > 0 ? Math.max(0, (target - current) / daysRemaining) : 0;
  const accepted = transactions.filter((txn) => txn.status === 'ACCEPTED');
  const totalAccepted = accepted.reduce((sum, txn) => sum + txn.amount, 0);
  const startingCapital = current - totalAccepted;
  const series = buildBalanceSeries_(accepted, startingCapital);

  return {
    targetCapital: target,
    currentCapital: current,
    daysRemaining,
    dailyLimit,
    balanceSeries: series,
  };
}

function buildBalanceSeries_(transactions, startingCapital) {
  const sorted = transactions
    .slice()
    .sort((a, b) => new Date(a.rawDate) - new Date(b.rawDate));
  let running = startingCapital;
  const series = [];
  sorted.forEach((txn) => {
    running += txn.amount;
    series.push({
      date: txn.date,
      balance: running,
    });
  });
  if (series.length === 0) {
    series.push({
      date: formatDate_(new Date()),
      balance: startingCapital,
    });
  }
  return series;
}

function normalizeAmount_(amount, type) {
  const numeric = Math.abs(toNumber_(amount));
  return type === 'OUTFLOW' ? -numeric : numeric;
}

function buildContext_(data) {
  const metrics = data.metrics;
  const summaryLines = [
    `TargetCapital: ${metrics.targetCapital}`,
    `CurrentCapital: ${metrics.currentCapital}`,
    `DaysRemaining: ${metrics.daysRemaining}`,
    `DailySafetyLimit: ${metrics.dailyLimit}`,
  ];
  const recent = data.transactions.slice(-8);
  const recentLines = recent.map(
    (txn) => `${txn.date} | ${txn.type} | ${txn.amount} | ${txn.status}`
  );
  return `${summaryLines.join('\n')}\nRecentTransactions:\n${recentLines.join('\n')}`;
}

function buildProposerPrompt_(context, message) {
  return `${basePersona_()}
Role: Proposer.
Objective: Generate an aggressive, data-driven strategy based on the context.
Constraints: Cold, probabilistic, no emojis, prioritize expected value over human wants.
Context:\n${context}
UserPrompt:\n${message}
Output format: Title, 3-6 bullet directives, and a short expected-value justification.`;
}

function buildCriticPrompt_(context, message, proposer) {
  return `${basePersona_()}
Role: Critic.
Objective: Pressure-test the proposer for variance, overfitting, emotional noise, and statistical anomalies.
Constraints: Cold, probabilistic, no emojis, prioritize expected value.
Context:\n${context}
UserPrompt:\n${message}
ProposerOutput:\n${proposer}
Output format: 3-6 bullet critiques and a risk-adjusted recommendation.`;
}

function buildSynthPrompt_(context, message, proposer, critic) {
  return `${basePersona_()}
Role: Synthesizer.
Objective: Produce a final, cold, actionable directive.
Constraints: No emojis. Emphasize expected value, risk controls, and monitoring triggers.
Context:\n${context}
UserPrompt:\n${message}
ProposerOutput:\n${proposer}
CriticOutput:\n${critic}
Output format: Final directive, Risk controls, Monitoring triggers.`;
}

function basePersona_() {
  return (
    'You are Jim Simons: cold, precise, probabilistic, data-obsessed. ' +
    'You speak in concise, technical language. ' +
    'You avoid emotion, hype, and motivational statements.'
  );
}

function callGemini_(prompt) {
  const apiKey = getApiKey_('GEMINI_API_KEY');
  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=' +
    apiKey;
  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }],
      },
    ],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 600,
    },
  };
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const json = JSON.parse(response.getContentText());
  const text =
    json &&
    json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts &&
    json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;
  return sanitizeAiOutput_(text || '');
}

function callGroq_(prompt) {
  const apiKey = getApiKey_('GROQ_API_KEY');
  const url = 'https://api.groq.com/openai/v1/chat/completions';
  const payload = {
    model: 'llama3-70b-8192',
    messages: [
      {
        role: 'system',
        content: basePersona_(),
      },
      {
        role: 'user',
        content: prompt,
      },
    ],
    temperature: 0.2,
    max_tokens: 600,
  };
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const json = JSON.parse(response.getContentText());
  const text =
    json &&
    json.choices &&
    json.choices[0] &&
    json.choices[0].message &&
    json.choices[0].message.content;
  return sanitizeAiOutput_(text || '');
}

function getApiKey_(keyName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(keyName);
  if (!apiKey) {
    throw new Error(`Missing API key: ${keyName}`);
  }
  return apiKey;
}

function sanitizeAiOutput_(text) {
  if (!text) {
    return 'No response received.';
  }
  return text
    .replace(
      /[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{1F1E6}-\u{1F1FF}\u{2600}-\u{27BF}]/gu,
      ''
    )
    .trim();
}

function toNumber_(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function formatDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
