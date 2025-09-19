const SPREADSHEET_ID = 'PUT_YOUR_SHEET_ID_HERE'; // required

const SHEET_DEFINITIONS = {
  Users: {
    headers: [
      'id',
      'firstName',
      'lastName',
      'email',
      'passwordHash',
      'role',
      'parentUserId',
      'notificationEmail',
      'tutorialSeen',
      'isActive',
      'createdAt',
      'updatedAt'
    ]
  },
  Tasks: {
    headers: [
      'id',
      'title',
      'description',
      'assignedTo',
      'priority',
      'status',
      'dueDate',
      'estimatedHours',
      'tags',
      'links',
      'createdBy',
      'createdAt',
      'updatedAt',
      'totalSeconds'
    ]
  },
  TimeLog: {
    headers: [
      'id',
      'taskId',
      'userEmail',
      'startedAt',
      'stoppedAt',
      'durationSeconds'
    ]
  },
  Tools: {
    headers: [
      'id',
      'iconUrl',
      'name',
      'toolUrl',
      'isActive',
      'createdAt',
      'updatedAt',
      'sortOrder'
    ]
  }
};

const PRIORITIES = ['Low', 'Normal', 'High'];
const STATUSES = ['Todo', 'Doing', 'Done', 'Blocked'];
const ROLE_ORDER = ['MASTER_ADMIN', 'ADMIN', 'MANAGER', 'USER'];
const BRAND_LOGO_URL = 'https://images.emojiterra.com/google/noto-emoji/unicode-16.0/color/svg/1f300.svg';

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'exportTasks') {
    const params = {
      view: e.parameter.view,
      q: e.parameter.q,
      includeDone: e.parameter.includeDone === 'true',
      statusFilters: e.parameter.status ? e.parameter.status.split(',') : [],
      overdueOnly: e.parameter.overdue === 'true'
    };
    return exportTasksCSV(params);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AuraFlow');
}

function getEnv_() {
  const env = ensureSpreadsheet_();
  if (!env.success) {
    return env;
  }
  const data = {
    spreadsheetId: SPREADSHEET_ID,
    users: { sheet: 'Users', headers: SHEET_DEFINITIONS.Users.headers },
    tasks: { sheet: 'Tasks', headers: SHEET_DEFINITIONS.Tasks.headers },
    timeLog: { sheet: 'TimeLog', headers: SHEET_DEFINITIONS.TimeLog.headers },
    tools: { sheet: 'Tools', headers: SHEET_DEFINITIONS.Tools.headers },
    appUrl: getAppUrl_()
  };
  return { success: true, data: data };
}

function getEnv() {
  return getEnv_();
}
function ensureSpreadsheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'PUT_YOUR_SHEET_ID_HERE') {
    return { success: false, message: 'INVALID_SPREADSHEET_ID' };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = {};
    Object.keys(SHEET_DEFINITIONS).forEach(function (name) {
      const headers = SHEET_DEFINITIONS[name].headers;
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
      }
      ensureHeaders_(sheet, headers);
      sheets[name] = { sheet: sheet, headers: headers };
    });
    return { success: true, data: { ss: ss, sheets: sheets } };
  } catch (err) {
    return { success: false, message: 'INVALID_SPREADSHEET_ID' };
  }
}

function ensureHeaders_(sheet, headers) {
  const width = headers.length;
  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, width).setValues([headers]);
  sheet.setFrozenRows(1);
}

function uuid_() {
  return Utilities.getUuid();
}

function todayISO_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function nowISODateTime_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function hash_(s) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s || '');
  return Utilities.base64Encode(digest);
}

function getSessionContext_(options) {
  options = options || {};
  const env = ensureSpreadsheet_();
  if (!env.success) {
    return env;
  }
  const data = env.data;
  const email = getCurrentEmail_();
  if (!email) {
    return { success: false, message: 'GOOGLE_SIGN_IN_REQUIRED' };
  }
  const usersSheet = data.sheets.Users.sheet;
  const headers = SHEET_DEFINITIONS.Users.headers;
  const table = loadTable_(usersSheet, headers);
  const entry = table.find(function (row) {
    return (row.record.email || '').toLowerCase() === email.toLowerCase();
  });
  if (!entry) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  if (!toBool_(entry.record.isActive)) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const user = sanitizeUser_(entry.record);
  if (options.includePrefs) {
    user.preferences = getNotificationPrefs_(user.id);
  }
  const context = {
    env: data,
    email: email,
    user: user,
    userEntry: entry,
    usersTable: table
  };
  if (options.includeUsers) {
    context.users = table.map(function (row) {
      return sanitizeUser_(row.record);
    });
  }
  return { success: true, data: context };
}

function getCurrentEmail_() {
  return (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').trim();
}

function isRoleAtLeast_(role, minimum) {
  role = role || 'USER';
  minimum = minimum || 'USER';
  return ROLE_ORDER.indexOf(role) <= ROLE_ORDER.indexOf(minimum);
}

function sanitizeUser_(record) {
  return {
    id: record.id || '',
    firstName: record.firstName || '',
    lastName: record.lastName || '',
    email: record.email || '',
    role: record.role || 'USER',
    parentUserId: record.parentUserId || '',
    notificationEmail: record.notificationEmail || record.email || '',
    tutorialSeen: toBool_(record.tutorialSeen),
    isActive: toBool_(record.isActive),
    createdAt: record.createdAt || '',
    updatedAt: record.updatedAt || ''
  };
}

function currentSession() {
  const ctx = getSessionContext_({ includePrefs: true });
  if (!ctx.success) {
    return ctx;
  }
  return {
    success: true,
    data: {
      email: ctx.data.email,
      user: ctx.data.user
    }
  };
}

function adminQuickAddSelf(payload) {
  payload = payload || {};
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const me = ctx.data.user;
  if (!isRoleAtLeast_(me.role, 'ADMIN')) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const email = ctx.data.email.toLowerCase();
  const exists = ctx.data.usersTable.find(function (entry) {
    return (entry.record.email || '').toLowerCase() === email;
  });
  if (exists) {
    return { success: false, message: 'ALREADY_EXISTS' };
  }
  const now = todayISO_();
  const record = {
    id: uuid_(),
    firstName: payload.firstName || '',
    lastName: payload.lastName || '',
    email: ctx.data.email,
    passwordHash: '',
    role: ROLE_ORDER.indexOf(payload.role) >= 0 ? payload.role : 'ADMIN',
    parentUserId: me.id,
    notificationEmail: payload.notificationEmail || ctx.data.email,
    tutorialSeen: 'FALSE',
    isActive: 'TRUE',
    createdAt: now,
    updatedAt: now
  };
  appendRecord_(ctx.data.env.sheets.Users.sheet, SHEET_DEFINITIONS.Users.headers, record);
  return { success: true, data: sanitizeUser_(record) };
}

function createUserFromUI(user) {
  user = user || {};
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const creator = ctx.data.user;
  if (!isRoleAtLeast_(creator.role, 'MANAGER')) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const required = ['firstName', 'email', 'role', 'notificationEmail'];
  const missing = required.filter(function (field) {
    return !user[field];
  });
  if (missing.length) {
    return { success: false, message: 'MISSING_' + missing.join('_') };
  }
  const email = (user.email || '').toLowerCase();
  const exists = ctx.data.usersTable.find(function (entry) {
    return (entry.record.email || '').toLowerCase() === email;
  });
  if (exists) {
    return { success: false, message: 'EMAIL_IN_USE' };
  }
  const now = todayISO_();
  const record = {
    id: uuid_(),
    firstName: user.firstName || '',
    lastName: user.lastName || '',
    email: user.email,
    passwordHash: user.passwordHash || '',
    role: ROLE_ORDER.indexOf(user.role) >= 0 ? user.role : 'USER',
    parentUserId: creator.id,
    notificationEmail: user.notificationEmail || user.email,
    tutorialSeen: user.tutorialSeen ? 'TRUE' : 'FALSE',
    isActive: 'TRUE',
    createdAt: now,
    updatedAt: now
  };
  appendRecord_(ctx.data.env.sheets.Users.sheet, SHEET_DEFINITIONS.Users.headers, record);
  return { success: true, data: sanitizeUser_(record) };
}

function listUsers(params) {
  params = params || {};
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  let users = ctx.data.users || [];
  const viewer = ctx.data.user;
  if (!isRoleAtLeast_(viewer.role, 'ADMIN')) {
    if (viewer.role === 'MANAGER') {
      users = users.filter(function (u) {
        return u.id === viewer.id || u.parentUserId === viewer.id;
      });
    } else {
      users = users.filter(function (u) {
        return u.id === viewer.id;
      });
    }
  }
  return { success: true, data: users };
}

function saveEmailSettings(payload) {
  payload = payload || {};
  const ctx = getSessionContext_({ includePrefs: true });
  if (!ctx.success) {
    return ctx;
  }
  const record = ctx.data.userEntry.record;
  if (payload.notificationEmail) {
    record.notificationEmail = payload.notificationEmail;
  }
  if (payload.tutorialSeen === true) {
    record.tutorialSeen = 'TRUE';
  }
  record.updatedAt = todayISO_();
  writeRecord_(ctx.data.env.sheets.Users.sheet, SHEET_DEFINITIONS.Users.headers, ctx.data.userEntry.rowNumber, record);
  saveNotificationPrefs_(ctx.data.user.id, {
    instant: payload.instant !== false,
    daily: payload.daily !== false,
    weekly: payload.weekly !== false
  });
  const updated = sanitizeUser_(record);
  updated.preferences = getNotificationPrefs_(ctx.data.user.id);
  return { success: true, data: updated };
}

function getTasks(params) {
  params = params || {};
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const requester = ctx.data.user;
  const env = ctx.data.env;
  const tasksSheet = env.sheets.Tasks.sheet;
  const headers = SHEET_DEFINITIONS.Tasks.headers;
  const table = loadTable_(tasksSheet, headers);
  const users = ctx.data.users || ctx.data.usersTable.map(function (entry) {
    return sanitizeUser_(entry.record);
  });
  const visibility = buildVisibilityEmails_(requester, users);
  const view = (params.view || 'my').toLowerCase();
  let allowed;
  if (view === 'all') {
    allowed = visibility.all;
  } else if (view === 'team') {
    allowed = visibility.team.length ? visibility.team : visibility.all;
  } else {
    allowed = visibility.my.length ? visibility.my : visibility.all;
  }
  const includeDone = params.includeDone === true;
  const statusFilters = (params.statusFilters || []).filter(function (s) {
    return STATUSES.indexOf(s) !== -1;
  });
  const overdueOnly = params.overdueOnly === true;
  const search = (params.q || '').toLowerCase();
  const sortBy = ['priority', 'status'].indexOf(params.sortBy) !== -1 ? params.sortBy : 'dueDate';
  const sortDir = (params.sortDir || 'asc').toLowerCase() === 'desc' ? 'desc' : 'asc';
  const pageSize = Math.min(200, Math.max(1, params.pageSize || 50));
  const startIndex = Math.max(0, parseInt(params.pageToken, 10) || 0);
  const today = todayISO_();

  let rows = table.map(function (entry) {
    return sanitizeTask_(entry.record);
  });

  rows = rows.filter(function (task) {
    if (!task.assignedTo) {
      return false;
    }
    if (allowed.length && allowed.indexOf(task.assignedTo.toLowerCase()) === -1) {
      return false;
    }
    if (!includeDone && task.status === 'Done') {
      return false;
    }
    if (statusFilters.length && statusFilters.indexOf(task.status) === -1) {
      return false;
    }
    if (overdueOnly) {
      if (!task.dueDate || task.status === 'Done' || task.dueDate >= today) {
        return false;
      }
    }
    if (search) {
      const haystack = [
        task.title,
        task.description,
        task.assignedTo,
        task.priority,
        task.status,
        task.tags,
        task.links
      ]
        .join(' ')
        .toLowerCase();
      if (haystack.indexOf(search) === -1) {
        return false;
      }
    }
    return true;
  });

  rows.sort(function (a, b) {
    const dir = sortDir === 'desc' ? -1 : 1;
    if (sortBy === 'priority') {
      return (PRIORITIES.indexOf(a.priority) - PRIORITIES.indexOf(b.priority)) * dir;
    }
    if (sortBy === 'status') {
      return (STATUSES.indexOf(a.status) - STATUSES.indexOf(b.status)) * dir;
    }
    const av = a.dueDate || '';
    const bv = b.dueDate || '';
    if (!av && !bv) {
      return a.title.localeCompare(b.title) * dir;
    }
    if (!av) {
      return 1 * dir;
    }
    if (!bv) {
      return -1 * dir;
    }
    if (av === bv) {
      return a.title.localeCompare(b.title) * dir;
    }
    return (av > bv ? 1 : -1) * dir;
  });

  const totals = computeTaskTotals_(rows, requester.email);
  const page = rows.slice(startIndex, startIndex + pageSize);
  const nextPageToken = startIndex + pageSize < rows.length ? String(startIndex + pageSize) : '';

  return {
    success: true,
    data: {
      rows: page,
      nextPageToken: nextPageToken,
      totals: totals
    }
  };
}

function createTask(task) {
  task = task || {};
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const requester = ctx.data.user;
  const users = ctx.data.users || ctx.data.usersTable.map(function (entry) {
    return sanitizeUser_(entry.record);
  });
  const assignedEmail = (task.assignedTo || '').toLowerCase();
  const assignee = users.find(function (u) {
    return u.email.toLowerCase() === assignedEmail && u.isActive;
  });
  if (!assignee) {
    return { success: false, message: 'INVALID_ASSIGNEE' };
  }
  const record = {
    id: uuid_(),
    title: task.title || 'Untitled Task',
    description: task.description || '',
    assignedTo: assignee.email,
    priority: PRIORITIES.indexOf(task.priority) >= 0 ? task.priority : 'Normal',
    status: STATUSES.indexOf(task.status) >= 0 ? task.status : 'Todo',
    dueDate: normalizeDate_(task.dueDate),
    estimatedHours: Number(task.estimatedHours) || 0,
    tags: task.tags || '',
    links: task.links || '',
    createdBy: requester.email,
    createdAt: todayISO_(),
    updatedAt: todayISO_(),
    totalSeconds: 0
  };
  appendRecord_(ctx.data.env.sheets.Tasks.sheet, SHEET_DEFINITIONS.Tasks.headers, record);
  const cc = [];
  if (requester.notificationEmail) {
    cc.push(requester.notificationEmail);
  }
  if (requester.email && cc.indexOf(requester.email) === -1) {
    cc.push(requester.email);
  }
  sendOnAssignment_(record, assignee.email, cc);
  return { success: true, data: sanitizeTask_(record) };
}

function updateTask(task) {
  task = task || {};
  if (!task.id) {
    return { success: false, message: 'TASK_ID_REQUIRED' };
  }
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const env = ctx.data.env;
  const sheet = env.sheets.Tasks.sheet;
  const headers = SHEET_DEFINITIONS.Tasks.headers;
  const table = loadTable_(sheet, headers);
  const entry = table.find(function (row) {
    return row.record.id === task.id;
  });
  if (!entry) {
    return { success: false, message: 'NOT_FOUND' };
  }
  const requester = ctx.data.user;
  if (!canInteractWithTask_(requester, entry.record, ctx.data.usersTable)) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const users = ctx.data.users || ctx.data.usersTable.map(function (u) {
    return sanitizeUser_(u.record);
  });
  const previousAssignee = entry.record.assignedTo;
  const nextAssigneeEmail = (task.assignedTo || previousAssignee || '').toLowerCase();
  const nextAssignee = users.find(function (u) {
    return u.email.toLowerCase() === nextAssigneeEmail && u.isActive;
  });
  if (!nextAssignee) {
    return { success: false, message: 'INVALID_ASSIGNEE' };
  }
  entry.record.title = task.title !== undefined ? task.title : entry.record.title;
  entry.record.description = task.description !== undefined ? task.description : entry.record.description;
  entry.record.assignedTo = nextAssignee.email;
  if (STATUSES.indexOf(task.status) !== -1) {
    entry.record.status = task.status;
  }
  if (PRIORITIES.indexOf(task.priority) !== -1) {
    entry.record.priority = task.priority;
  }
  const normalizedDue = normalizeDate_(task.dueDate);
  if (normalizedDue || task.dueDate === '') {
    entry.record.dueDate = normalizedDue;
  }
  if (task.estimatedHours !== undefined) {
    entry.record.estimatedHours = Number(task.estimatedHours) || 0;
  }
  if (task.tags !== undefined) {
    entry.record.tags = task.tags;
  }
  if (task.links !== undefined) {
    entry.record.links = task.links;
  }
  entry.record.updatedAt = todayISO_();
  writeRecord_(sheet, headers, entry.rowNumber, entry.record);
  if ((previousAssignee || '').toLowerCase() !== entry.record.assignedTo.toLowerCase()) {
    const cc = [];
    if (requester.notificationEmail) {
      cc.push(requester.notificationEmail);
    }
    if (requester.email && cc.indexOf(requester.email) === -1) {
      cc.push(requester.email);
    }
    sendOnAssignment_(entry.record, entry.record.assignedTo, cc);
  }
  return { success: true, data: sanitizeTask_(entry.record) };
}

function markDone(payload) {
  payload = payload || {};
  if (!payload.taskId) {
    return { success: false, message: 'TASK_ID_REQUIRED' };
  }
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const env = ctx.data.env;
  const sheet = env.sheets.Tasks.sheet;
  const headers = SHEET_DEFINITIONS.Tasks.headers;
  const table = loadTable_(sheet, headers);
  const entry = table.find(function (row) {
    return row.record.id === payload.taskId;
  });
  if (!entry) {
    return { success: false, message: 'NOT_FOUND' };
  }
  if (!canInteractWithTask_(ctx.data.user, entry.record, ctx.data.usersTable)) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  entry.record.status = 'Done';
  entry.record.updatedAt = todayISO_();
  writeRecord_(sheet, headers, entry.rowNumber, entry.record);
  return { success: true, data: sanitizeTask_(entry.record) };
}

function startTimer(taskId) {
  if (!taskId) {
    return { success: false, message: 'TASK_ID_REQUIRED' };
  }
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const env = ctx.data.env;
  const tasksSheet = env.sheets.Tasks.sheet;
  const headers = SHEET_DEFINITIONS.Tasks.headers;
  const table = loadTable_(tasksSheet, headers);
  const entry = table.find(function (row) {
    return row.record.id === taskId;
  });
  if (!entry) {
    return { success: false, message: 'NOT_FOUND' };
  }
  if (!canInteractWithTask_(ctx.data.user, entry.record, ctx.data.usersTable)) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const record = {
    id: uuid_(),
    taskId: taskId,
    userEmail: ctx.data.user.email,
    startedAt: nowISODateTime_(),
    stoppedAt: '',
    durationSeconds: 0
  };
  appendRecord_(env.sheets.TimeLog.sheet, SHEET_DEFINITIONS.TimeLog.headers, record);
  return { success: true, data: record };
}

function stopTimer(taskId) {
  if (!taskId) {
    return { success: false, message: 'TASK_ID_REQUIRED' };
  }
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const env = ctx.data.env;
  const timeSheet = env.sheets.TimeLog.sheet;
  const headers = SHEET_DEFINITIONS.TimeLog.headers;
  const table = loadTable_(timeSheet, headers);
  const entry = table.reverse().find(function (row) {
    return row.record.taskId === taskId && (row.record.userEmail || '').toLowerCase() === ctx.data.user.email.toLowerCase() && !row.record.stoppedAt;
  });
  if (!entry) {
    return { success: false, message: 'TIMER_NOT_FOUND' };
  }
  const stopTime = nowISODateTime_();
  const duration = Math.max(1, Math.floor((new Date(stopTime) - new Date(entry.record.startedAt)) / 1000));
  entry.record.stoppedAt = stopTime;
  entry.record.durationSeconds = duration;
  writeRecord_(timeSheet, headers, entry.rowNumber, entry.record);
  accumulateTaskTime_(env, taskId);
  return { success: true, data: entry.record };
}

function bulkImportTasks(rows) {
  rows = rows || [];
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  if (!isRoleAtLeast_(ctx.data.user.role, 'MANAGER')) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const users = ctx.data.users || ctx.data.usersTable.map(function (entry) {
    return sanitizeUser_(entry.record);
  });
  const sheet = ctx.data.env.sheets.Tasks.sheet;
  const headers = SHEET_DEFINITIONS.Tasks.headers;
  const errors = [];
  let count = 0;
  rows.forEach(function (row, index) {
    const assignedTo = (row.assignedTo || '').toLowerCase();
    const assignee = users.find(function (u) {
      return u.email.toLowerCase() === assignedTo && u.isActive;
    });
    if (!assignee) {
      errors.push({ index: index, message: 'INVALID_ASSIGNEE' });
      return;
    }
    const record = {
      id: uuid_(),
      title: row.title || 'Untitled Task',
      description: row.description || '',
      assignedTo: assignee.email,
      priority: PRIORITIES.indexOf(row.priority) >= 0 ? row.priority : 'Normal',
      status: STATUSES.indexOf(row.status) >= 0 ? row.status : 'Todo',
      dueDate: normalizeDate_(row.dueDate),
      estimatedHours: Number(row.estimatedHours) || 0,
      tags: row.tags || '',
      links: row.links || '',
      createdBy: ctx.data.user.email,
      createdAt: todayISO_(),
      updatedAt: todayISO_(),
      totalSeconds: 0
    };
    appendRecord_(sheet, headers, record);
    count++;
  });
  return { success: true, count: count, errors: errors };
}

function exportTasksCSV(params) {
  params = params || {};
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ContentService.createTextOutput(JSON.stringify(ctx));
  }
  const response = getTasks(params);
  if (!response.success) {
    return ContentService.createTextOutput(JSON.stringify(response));
  }
  const headers = ['title', 'description', 'assignedTo', 'priority', 'status', 'dueDate', 'estimatedHours', 'tags', 'links'];
  const lines = [headers.join(',')];
  response.data.rows.forEach(function (task) {
    const row = headers.map(function (key) {
      const value = task[key] !== undefined && task[key] !== null ? String(task[key]) : '';
      if (value.indexOf(',') > -1 || value.indexOf('"') > -1 || value.indexOf('\n') > -1) {
        return '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    });
    lines.push(row.join(','));
  });
  return ContentService.createTextOutput(lines.join('\n')).setMimeType(ContentService.MimeType.CSV);
}

function getDashboardMetrics() {
  const ctx = getSessionContext_({ includeUsers: true });
  if (!ctx.success) {
    return ctx;
  }
  const env = ctx.data.env;
  const tasksSheet = env.sheets.Tasks.sheet;
  const headers = SHEET_DEFINITIONS.Tasks.headers;
  const table = loadTable_(tasksSheet, headers);
  const users = ctx.data.users || ctx.data.usersTable.map(function (entry) {
    return sanitizeUser_(entry.record);
  });
  const visibility = buildVisibilityEmails_(ctx.data.user, users);
  const allowed = visibility.all;
  const tasks = table
    .map(function (row) {
      return sanitizeTask_(row.record);
    })
    .filter(function (task) {
      return allowed.indexOf(task.assignedTo.toLowerCase()) !== -1;
    });
  const totals = computeTaskTotals_(tasks, ctx.data.user.email);
  const today = todayISO_();
  const dueSoonLimit = addDays_(today, 3);
  const overdue = tasks.filter(function (task) {
    return task.dueDate && task.dueDate < today && task.status !== 'Done';
  }).length;
  const dueSoon = tasks.filter(function (task) {
    return task.dueDate && task.dueDate >= today && task.dueDate <= dueSoonLimit && task.status !== 'Done';
  }).length;
  const myTasks = tasks.filter(function (task) {
    return task.assignedTo.toLowerCase() === ctx.data.user.email.toLowerCase();
  });
  return {
    success: true,
    data: {
      total: tasks.length,
      byStatus: totals.byStatus,
      overdue: overdue,
      dueSoon: dueSoon,
      myTotals: computeTaskTotals_(myTasks, ctx.data.user.email),
      teamTotals: totals
    }
  };
}

function listTools(params) {
  params = params || {};
  const ctx = getSessionContext_({ includeUsers: false });
  if (!ctx.success) {
    return ctx;
  }
  const includeInactive = params.includeInactive && isRoleAtLeast_(ctx.data.user.role, 'ADMIN');
  const sheet = ctx.data.env.sheets.Tools.sheet;
  const headers = SHEET_DEFINITIONS.Tools.headers;
  const table = loadTable_(sheet, headers);
  const tools = table
    .map(function (row) {
      return sanitizeTool_(row.record);
    })
    .filter(function (tool) {
      return includeInactive || tool.isActive;
    })
    .sort(function (a, b) {
      return a.sortOrder - b.sortOrder;
    });
  return { success: true, data: tools };
}

function createTool(tool) {
  tool = tool || {};
  const ctx = getSessionContext_({ includeUsers: false });
  if (!ctx.success) {
    return ctx;
  }
  if (!isRoleAtLeast_(ctx.data.user.role, 'ADMIN')) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const record = {
    id: uuid_(),
    iconUrl: tool.iconUrl || '',
    name: tool.name || '',
    toolUrl: tool.toolUrl || '',
    isActive: tool.isActive === false ? 'FALSE' : 'TRUE',
    createdAt: todayISO_(),
    updatedAt: todayISO_(),
    sortOrder: Number(tool.sortOrder) || 0
  };
  appendRecord_(ctx.data.env.sheets.Tools.sheet, SHEET_DEFINITIONS.Tools.headers, record);
  return { success: true, data: sanitizeTool_(record) };
}

function updateTool(tool) {
  tool = tool || {};
  if (!tool.id) {
    return { success: false, message: 'TOOL_ID_REQUIRED' };
  }
  const ctx = getSessionContext_({ includeUsers: false });
  if (!ctx.success) {
    return ctx;
  }
  if (!isRoleAtLeast_(ctx.data.user.role, 'ADMIN')) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const sheet = ctx.data.env.sheets.Tools.sheet;
  const headers = SHEET_DEFINITIONS.Tools.headers;
  const table = loadTable_(sheet, headers);
  const entry = table.find(function (row) {
    return row.record.id === tool.id;
  });
  if (!entry) {
    return { success: false, message: 'NOT_FOUND' };
  }
  if (tool.iconUrl !== undefined) {
    entry.record.iconUrl = tool.iconUrl;
  }
  if (tool.name !== undefined) {
    entry.record.name = tool.name;
  }
  if (tool.toolUrl !== undefined) {
    entry.record.toolUrl = tool.toolUrl;
  }
  if (tool.isActive !== undefined) {
    entry.record.isActive = tool.isActive ? 'TRUE' : 'FALSE';
  }
  if (tool.sortOrder !== undefined) {
    entry.record.sortOrder = Number(tool.sortOrder) || 0;
  }
  entry.record.updatedAt = todayISO_();
  writeRecord_(sheet, headers, entry.rowNumber, entry.record);
  return { success: true, data: sanitizeTool_(entry.record) };
}

function deleteTool(payload) {
  const id = typeof payload === 'string' ? payload : (payload && payload.id);
  if (!id) {
    return { success: false, message: 'TOOL_ID_REQUIRED' };
  }
  const ctx = getSessionContext_({ includeUsers: false });
  if (!ctx.success) {
    return ctx;
  }
  if (!isRoleAtLeast_(ctx.data.user.role, 'ADMIN')) {
    return { success: false, message: 'ACCESS_DENIED' };
  }
  const sheet = ctx.data.env.sheets.Tools.sheet;
  const headers = SHEET_DEFINITIONS.Tools.headers;
  const table = loadTable_(sheet, headers);
  const entry = table.find(function (row) {
    return row.record.id === id;
  });
  if (!entry) {
    return { success: false, message: 'NOT_FOUND' };
  }
  entry.record.isActive = 'FALSE';
  entry.record.updatedAt = todayISO_();
  writeRecord_(sheet, headers, entry.rowNumber, entry.record);
  return { success: true, data: sanitizeTool_(entry.record) };
}

function sendOnAssignment_(task, assigneeEmail, ccEmails) {
  if (!assigneeEmail) {
    return { success: false, message: 'NO_ASSIGNEE' };
  }
  const url = getAppUrl_();
  const cc = dedupeArray_((ccEmails || []).filter(Boolean));
  const subject = '[AuraFlow] Task Assigned: ' + (task.title || 'Task');
  const due = task.dueDate ? '<p><strong>Due Date:</strong> ' + task.dueDate + '</p>' : '';
  const estimated = Number(task.estimatedHours) ? '<p><strong>Estimated Hours:</strong> ' + Number(task.estimatedHours) + '</p>' : '';
  const html = [
    '<div style="font-family:Roboto,Arial,sans-serif;color:#202124;">',
    '<div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;">',
    '<img src="' + BRAND_LOGO_URL + '" alt="AuraFlow" style="height:28px;width:28px;">',
    '<span style="font-size:20px;font-weight:600;">AuraFlow</span>',
    '</div>',
    '<div style="background:#ffffff;border:1px solid #e0e0e0;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,0.08);">',
    '<h2 style="margin:0 0 8px;font-size:18px;">New Task Assigned</h2>',
    '<p style="margin:0 0 12px;color:#5f6368;">' + (task.createdBy ? 'Assigned by ' + task.createdBy : 'New assignment') + '</p>',
    '<p><strong>Title:</strong> ' + (task.title || 'Task') + '</p>',
    (task.description ? '<p><strong>Description:</strong><br>' + sanitizeHtml_(task.description) + '</p>' : ''),
    due,
    estimated,
    '<p><strong>Status:</strong> ' + (task.status || 'Todo') + '</p>',
    '<a href="' + url + '" style="display:inline-block;margin-top:12px;padding:10px 18px;background:#1a73e8;color:#fff;border-radius:6px;text-decoration:none;font-weight:500;">View Task</a>',
    '</div>',
    brandFooterHtml_(),
    '</div>'
  ].join('');
  try {
    MailApp.sendEmail({
      to: assigneeEmail,
      subject: subject,
      htmlBody: html,
      body: 'A task has been assigned to you in AuraFlow. View it here: ' + url,
      cc: cc.join(','),
      name: 'AuraFlow'
    });
  } catch (err) {
    Logger.log('Assignment email failed: ' + err);
  }
  return { success: true };
}

function dailyDigestRunner() {
  const env = ensureSpreadsheet_();
  if (!env.success) {
    return env;
  }
  sendDigestEmails_(env.data, 'daily');
  return { success: true };
}

function weeklyDigestRunner() {
  const env = ensureSpreadsheet_();
  if (!env.success) {
    return env;
  }
  sendDigestEmails_(env.data, 'weekly');
  return { success: true };
}

function initializeApp(options) {
  options = options || {};
  const env = ensureSpreadsheet_();
  if (!env.success) {
    return env;
  }
  let triggerResult;
  try {
    triggerResult = setupTriggers();
  } catch (err) {
    triggerResult = { success: false, message: err.message };
  }
  let demoResult = null;
  if (options.createDemoData || options.seedDemoData) {
    demoResult = createDemoData(options);
  }
  const response = {
    success: true,
    message: 'APP_INITIALIZED',
    data: {
      sheetsEnsured: Object.keys(env.data.sheets),
      triggers: triggerResult,
      demo: demoResult
    }
  };
  if (triggerResult && triggerResult.success === false) {
    response.success = false;
    response.message = triggerResult.message || 'TRIGGER_SETUP_FAILED';
  }
  if (demoResult && demoResult.success === false) {
    response.success = false;
    response.message = demoResult.message || 'DEMO_DATA_FAILED';
  }
  return response;
}

function createDemoData(options) {
  options = options || {};
  const env = ensureSpreadsheet_();
  if (!env.success) {
    return env;
  }
  try {
    const data = env.data;
    const today = todayISO_();
    const timezone = Session.getScriptTimeZone();
    const summary = {
      usersCreated: 0,
      tasksCreated: 0,
      timeLogsCreated: 0,
      toolsCreated: 0
    };

    const usersSheet = data.sheets.Users.sheet;
    const userHeaders = SHEET_DEFINITIONS.Users.headers;
    const userRows = loadTable_(usersSheet, userHeaders);
    const userByEmail = {};
    userRows.forEach(function (entry) {
      const email = (entry.record.email || '').toLowerCase();
      if (!email) {
        return;
      }
      if (!entry.record.id) {
        entry.record.id = uuid_();
        if (!entry.record.createdAt) {
          entry.record.createdAt = today;
        }
        entry.record.updatedAt = today;
        writeRecord_(usersSheet, userHeaders, entry.rowNumber, entry.record);
      }
      userByEmail[email] = entry;
    });

    const sampleUsers = [
      {
        email: 'mananvermabusiness@gmail.com',
        firstName: 'Manan',
        lastName: 'Verma',
        role: 'MASTER_ADMIN',
        notificationEmail: 'mananvermabusiness@gmail.com',
        tutorialSeen: true
      },
      {
        email: 'operations@auraflow.app',
        firstName: 'Operations',
        lastName: 'Lead',
        role: 'ADMIN',
        notificationEmail: 'operations@auraflow.app',
        parentEmail: 'mananvermabusiness@gmail.com'
      },
      {
        email: 'manager@auraflow.app',
        firstName: 'Taylor',
        lastName: 'Lee',
        role: 'MANAGER',
        notificationEmail: 'manager@auraflow.app',
        parentEmail: 'operations@auraflow.app'
      },
      {
        email: 'teammate@auraflow.app',
        firstName: 'Jordan',
        lastName: 'Chen',
        role: 'USER',
        notificationEmail: 'teammate@auraflow.app',
        parentEmail: 'manager@auraflow.app'
      }
    ];

    function getParentId(parentEmail) {
      if (!parentEmail) {
        return '';
      }
      const parentEntry = userByEmail[parentEmail.toLowerCase()];
      return parentEntry && parentEntry.record.id ? parentEntry.record.id : '';
    }

    sampleUsers.forEach(function (def) {
      const key = def.email.toLowerCase();
      let entry = userByEmail[key];
      if (entry) {
        let mutated = false;
        if (!entry.record.id) {
          entry.record.id = uuid_();
          mutated = true;
        }
        if (!entry.record.notificationEmail && def.notificationEmail) {
          entry.record.notificationEmail = def.notificationEmail;
          mutated = true;
        }
        if (!entry.record.parentUserId && def.parentEmail) {
          const parentId = getParentId(def.parentEmail);
          if (parentId) {
            entry.record.parentUserId = parentId;
            mutated = true;
          }
        }
        if (!entry.record.isActive) {
          entry.record.isActive = 'TRUE';
          mutated = true;
        }
        if (mutated) {
          entry.record.updatedAt = today;
          writeRecord_(usersSheet, userHeaders, entry.rowNumber, entry.record);
        }
        return;
      }
      const record = {
        id: uuid_(),
        firstName: def.firstName || '',
        lastName: def.lastName || '',
        email: def.email,
        passwordHash: '',
        role: def.role || 'USER',
        parentUserId: getParentId(def.parentEmail),
        notificationEmail: def.notificationEmail || def.email,
        tutorialSeen: def.tutorialSeen ? 'TRUE' : 'FALSE',
        isActive: 'TRUE',
        createdAt: today,
        updatedAt: today
      };
      const rowNumber = appendRecord_(usersSheet, userHeaders, record);
      entry = { rowNumber: rowNumber, record: record };
      userByEmail[key] = entry;
      summary.usersCreated += 1;
    });

    function getUserRecord(email) {
      if (!email) {
        return null;
      }
      const entry = userByEmail[email.toLowerCase()];
      return entry ? entry.record : null;
    }

    const tasksSheet = data.sheets.Tasks.sheet;
    const taskHeaders = SHEET_DEFINITIONS.Tasks.headers;
    const taskRows = loadTable_(tasksSheet, taskHeaders);
    const tasksByTitle = {};
    taskRows.forEach(function (entry) {
      const titleKey = (entry.record.title || '').toLowerCase();
      if (!titleKey) {
        return;
      }
      if (!entry.record.id) {
        entry.record.id = uuid_();
        if (!entry.record.createdAt) {
          entry.record.createdAt = today;
        }
        entry.record.updatedAt = today;
        writeRecord_(tasksSheet, taskHeaders, entry.rowNumber, entry.record);
      }
      tasksByTitle[titleKey] = entry;
    });

    const sampleTasks = [
      {
        title: 'Finalize AuraFlow onboarding checklist',
        description: 'Document the onboarding flow and share it with new collaborators.',
        assignedTo: 'teammate@auraflow.app',
        priority: 'High',
        status: 'Doing',
        dueOffset: 2,
        estimatedHours: 6,
        tags: 'onboarding,process',
        links: 'https://auraflow.one/wiki/onboarding',
        createdBy: 'operations@auraflow.app',
        logs: [
          { userEmail: 'teammate@auraflow.app', startTime: '09:00', durationMinutes: 90, dateOffset: -1 },
          { userEmail: 'teammate@auraflow.app', startTime: '11:00', durationMinutes: 60, dateOffset: -1 }
        ]
      },
      {
        title: 'Prepare weekly executive summary',
        description: 'Compile highlights, blockers, and wins for leadership review.',
        assignedTo: 'manager@auraflow.app',
        priority: 'Normal',
        status: 'Todo',
        dueOffset: 3,
        estimatedHours: 4,
        tags: 'reporting,leadership',
        links: 'https://auraflow.one/docs/executive-summary',
        createdBy: 'mananvermabusiness@gmail.com',
        logs: [
          { userEmail: 'manager@auraflow.app', startTime: '14:00', durationMinutes: 45, dateOffset: 0 }
        ]
      },
      {
        title: 'Backlog triage and grooming',
        description: 'Review incoming requests, prioritize, and clarify requirements.',
        assignedTo: 'operations@auraflow.app',
        priority: 'Low',
        status: 'Blocked',
        dueOffset: -1,
        estimatedHours: 2,
        tags: 'planning,triage',
        links: 'https://auraflow.one/boards/backlog',
        createdBy: 'mananvermabusiness@gmail.com',
        logs: []
      }
    ];

    const seededTasks = [];

    sampleTasks.forEach(function (def) {
      const assigned = getUserRecord(def.assignedTo);
      if (!assigned) {
        return;
      }
      const titleKey = (def.title || '').toLowerCase();
      if (!titleKey) {
        return;
      }
      if (tasksByTitle[titleKey]) {
        seededTasks.push({ def: def, record: tasksByTitle[titleKey].record, created: false });
        return;
      }
      const dueDate = typeof def.dueOffset === 'number' ? addDays_(today, def.dueOffset) : (def.dueDate || '');
      const logs = def.logs || [];
      const totalSeconds = logs.reduce(function (sum, log) {
        return sum + (Number(log.durationMinutes) || 0) * 60;
      }, 0);
      const record = {
        id: uuid_(),
        title: def.title,
        description: def.description || '',
        assignedTo: def.assignedTo,
        priority: PRIORITIES.indexOf(def.priority) >= 0 ? def.priority : 'Normal',
        status: STATUSES.indexOf(def.status) >= 0 ? def.status : 'Todo',
        dueDate: dueDate || '',
        estimatedHours: Number(def.estimatedHours) || 0,
        tags: def.tags || '',
        links: def.links || '',
        createdBy: def.createdBy || def.assignedTo,
        createdAt: today,
        updatedAt: today,
        totalSeconds: totalSeconds
      };
      const rowNumber = appendRecord_(tasksSheet, taskHeaders, record);
      const entry = { rowNumber: rowNumber, record: record };
      tasksByTitle[titleKey] = entry;
      seededTasks.push({ def: def, record: record, created: true });
      summary.tasksCreated += 1;
    });

    const timeLogSheet = data.sheets.TimeLog.sheet;
    const timeLogHeaders = SHEET_DEFINITIONS.TimeLog.headers;
    const existingLogs = loadTable_(timeLogSheet, timeLogHeaders);
    const logKeys = {};
    existingLogs.forEach(function (entry) {
      const key = (entry.record.taskId || '') + '|' + (entry.record.startedAt || '');
      logKeys[key] = true;
    });

    function buildDateFromParts(isoDate, timeString) {
      if (!isoDate) {
        return new Date();
      }
      const dateParts = isoDate.split('-');
      const timeParts = (timeString || '09:00').split(':');
      const year = Number(dateParts[0]);
      const month = Number(dateParts[1]) - 1;
      const day = Number(dateParts[2]);
      const hour = Number(timeParts[0]) || 0;
      const minute = Number(timeParts[1]) || 0;
      return new Date(year, month, day, hour, minute, 0, 0);
    }

    seededTasks.forEach(function (taskEntry) {
      if (!taskEntry.created || !taskEntry.def.logs || !taskEntry.def.logs.length) {
        return;
      }
      taskEntry.def.logs.forEach(function (logDef) {
        const logDate = addDays_(today, typeof logDef.dateOffset === 'number' ? logDef.dateOffset : 0);
        if (!logDate) {
          return;
        }
        const startDate = buildDateFromParts(logDate, logDef.startTime || '09:00');
        const durationSeconds = (Number(logDef.durationMinutes) || 0) * 60;
        const stopDate = new Date(startDate.getTime() + durationSeconds * 1000);
        const startedAt = Utilities.formatDate(startDate, timezone, "yyyy-MM-dd'T'HH:mm:ssXXX");
        const stoppedAt = Utilities.formatDate(stopDate, timezone, "yyyy-MM-dd'T'HH:mm:ssXXX");
        const key = taskEntry.record.id + '|' + startedAt;
        if (logKeys[key]) {
          return;
        }
        logKeys[key] = true;
        const record = {
          id: uuid_(),
          taskId: taskEntry.record.id,
          userEmail: logDef.userEmail || taskEntry.record.assignedTo,
          startedAt: startedAt,
          stoppedAt: stoppedAt,
          durationSeconds: durationSeconds
        };
        appendRecord_(timeLogSheet, timeLogHeaders, record);
        summary.timeLogsCreated += 1;
      });
    });

    const toolsSheet = data.sheets.Tools.sheet;
    const toolHeaders = SHEET_DEFINITIONS.Tools.headers;
    const toolRows = loadTable_(toolsSheet, toolHeaders);
    const toolsByName = {};
    toolRows.forEach(function (entry) {
      const key = (entry.record.name || '').toLowerCase();
      if (key) {
        toolsByName[key] = entry;
      }
    });

    const sampleTools = [
      {
        name: 'AuraFlow Help Center',
        iconUrl: BRAND_LOGO_URL,
        toolUrl: 'https://auraflow.one/help',
        isActive: true,
        sortOrder: 1
      },
      {
        name: 'Team Stand-up Room',
        iconUrl: 'https://www.gstatic.com/images/branding/product/1x/meet_2020q4_48dp.png',
        toolUrl: 'https://meet.google.com/',
        isActive: true,
        sortOrder: 2
      },
      {
        name: 'Product Feedback Board',
        iconUrl: 'https://www.gstatic.com/images/icons/material/system_gm/1x/assignment_turned_in_gm_blue_48dp.png',
        toolUrl: 'https://workspace.google.com/marketplace/category/works-with-drive',
        isActive: true,
        sortOrder: 3
      }
    ];

    sampleTools.forEach(function (def) {
      const key = (def.name || '').toLowerCase();
      if (!key || toolsByName[key]) {
        return;
      }
      const record = {
        id: uuid_(),
        iconUrl: def.iconUrl || BRAND_LOGO_URL,
        name: def.name,
        toolUrl: def.toolUrl || '',
        isActive: def.isActive === false ? 'FALSE' : 'TRUE',
        createdAt: today,
        updatedAt: today,
        sortOrder: def.sortOrder !== undefined ? Number(def.sortOrder) : 0
      };
      appendRecord_(toolsSheet, toolHeaders, record);
      summary.toolsCreated += 1;
    });

    const totalCreated = summary.usersCreated + summary.tasksCreated + summary.timeLogsCreated + summary.toolsCreated;
    const message = totalCreated ? 'DEMO_DATA_CREATED' : 'DEMO_DATA_ALREADY_PRESENT';
    return { success: true, message: message, data: summary };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const managed = ['dailyDigestRunner', 'weeklyDigestRunner'];
  triggers.forEach(function (trigger) {
    if (managed.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('dailyDigestRunner')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .inTimezone('Asia/Kolkata')
    .create();
  ScriptApp.newTrigger('weeklyDigestRunner')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(10)
    .inTimezone('Asia/Kolkata')
    .create();
  return { success: true, message: 'TRIGGERS_CONFIGURED' };
}

function sendTestEmail() {
  const ctx = getSessionContext_({ includePrefs: true });
  if (!ctx.success) {
    return ctx;
  }
  const url = getAppUrl_();
  const html = [
    '<div style="font-family:Roboto,Arial,sans-serif;color:#202124;">',
    '<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">',
    '<img src="' + BRAND_LOGO_URL + '" alt="AuraFlow" style="height:28px;width:28px;">',
    '<span style="font-size:20px;font-weight:600;">AuraFlow</span>',
    '</div>',
    '<p>This is a sample notification from AuraFlow. Your notification email is set to <strong>' + (ctx.data.user.notificationEmail || ctx.data.user.email) + '</strong>.</p>',
    '<p><a href="' + url + '" style="color:#1a73e8;">Open AuraFlow</a></p>',
    brandFooterHtml_(),
    '</div>'
  ].join('');
  try {
    MailApp.sendEmail({
      to: ctx.data.user.notificationEmail || ctx.data.user.email,
      subject: '[AuraFlow] Notification Test',
      htmlBody: html,
      body: 'This is a test notification from AuraFlow. Open the app here: ' + url,
      name: 'AuraFlow'
    });
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function sendDigestEmails_(env, mode) {
  const users = loadTable_(env.sheets.Users.sheet, SHEET_DEFINITIONS.Users.headers)
    .map(function (row) { return sanitizeUser_(row.record); })
    .filter(function (user) { return user.email && user.isActive; });
  if (!users.length) {
    return;
  }
  const tasks = loadTable_(env.sheets.Tasks.sheet, SHEET_DEFINITIONS.Tasks.headers)
    .map(function (row) { return sanitizeTask_(row.record); });
  const today = todayISO_();
  const windowDays = mode === 'weekly' ? 7 : 1;
  const windowEnd = addDays_(today, windowDays);
  users.forEach(function (user) {
    const visibility = buildVisibilityEmails_(user, users);
    const allowed = visibility.all;
    const visibleTasks = tasks.filter(function (task) {
      return allowed.indexOf(task.assignedTo.toLowerCase()) !== -1;
    });
    const overdue = visibleTasks.filter(function (task) {
      return task.dueDate && task.dueDate < today && task.status !== 'Done';
    });
    const upcoming = visibleTasks.filter(function (task) {
      return task.dueDate && task.dueDate >= today && task.dueDate <= windowEnd && task.status !== 'Done';
    });
    if (!overdue.length && !upcoming.length) {
      return;
    }
    const htmlSections = [];
    if (overdue.length) {
      htmlSections.push('<h3 style="margin:16px 0 8px;">Overdue Tasks</h3>');
      htmlSections.push(renderTaskDigestList_(overdue));
    }
    if (upcoming.length) {
      htmlSections.push('<h3 style="margin:16px 0 8px;">Upcoming Tasks</h3>');
      htmlSections.push(renderTaskDigestList_(upcoming));
    }
    const html = [
      '<div style="font-family:Roboto,Arial,sans-serif;color:#202124;">',
      '<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">',
      '<img src="' + BRAND_LOGO_URL + '" alt="AuraFlow" style="height:28px;width:28px;">',
      '<span style="font-size:20px;font-weight:600;">AuraFlow</span>',
      '</div>',
      '<p style="margin:0 0 12px;">Here is your ' + (mode === 'weekly' ? 'weekly' : 'daily') + ' task digest.</p>',
      htmlSections.join(''),
      '<p style="margin:16px 0 0;">View all tasks in AuraFlow to update progress.</p>',
      '<p><a href="' + getAppUrl_() + '" style="color:#1a73e8;">Open AuraFlow</a></p>',
      brandFooterHtml_(),
      '</div>'
    ].join('');
    try {
      MailApp.sendEmail({
        to: user.notificationEmail || user.email,
        subject: '[AuraFlow] ' + (mode === 'weekly' ? 'Weekly' : 'Daily') + ' Digest',
        htmlBody: html,
        body: 'Open AuraFlow to review your tasks: ' + getAppUrl_(),
        name: 'AuraFlow'
      });
    } catch (err) {
      Logger.log('Digest email failed for ' + user.email + ': ' + err);
    }
  });
}

function renderTaskDigestList_(tasks) {
  if (!tasks.length) {
    return '<p>No tasks.</p>';
  }
  const items = tasks.slice(0, 20).map(function (task) {
    return '<li style="margin-bottom:6px;">' +
      '<strong>' + sanitizeHtml_(task.title) + '</strong>' +
      (task.dueDate ? ' &mdash; due ' + task.dueDate : '') +
      ' <span style="color:#5f6368;">(' + task.status + ')</span>' +
      '</li>';
  }).join('');
  return '<ul style="padding-left:20px;margin:0;">' + items + '</ul>';
}

function appendRecord_(sheet, headers, record) {
  const row = headers.map(function (key) {
    return record[key] !== undefined && record[key] !== null ? record[key] : '';
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function writeRecord_(sheet, headers, rowNumber, record) {
  const row = headers.map(function (key) {
    return record[key] !== undefined && record[key] !== null ? record[key] : '';
  });
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
}

function loadTable_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function (row, index) {
    const record = {};
    headers.forEach(function (header, col) {
      record[header] = row[col];
    });
    return { rowNumber: index + 2, record: record };
  });
}

function sanitizeTask_(record) {
  return {
    id: record.id || '',
    title: record.title || '',
    description: record.description || '',
    assignedTo: (record.assignedTo || '').trim(),
    priority: PRIORITIES.indexOf(record.priority) >= 0 ? record.priority : 'Normal',
    status: STATUSES.indexOf(record.status) >= 0 ? record.status : 'Todo',
    dueDate: record.dueDate || '',
    estimatedHours: Number(record.estimatedHours) || 0,
    tags: record.tags || '',
    links: record.links || '',
    createdBy: record.createdBy || '',
    createdAt: record.createdAt || '',
    updatedAt: record.updatedAt || '',
    totalSeconds: Number(record.totalSeconds) || 0
  };
}

function sanitizeTool_(record) {
  return {
    id: record.id || '',
    iconUrl: record.iconUrl || BRAND_LOGO_URL,
    name: record.name || '',
    toolUrl: record.toolUrl || '',
    isActive: toBool_(record.isActive),
    createdAt: record.createdAt || '',
    updatedAt: record.updatedAt || '',
    sortOrder: Number(record.sortOrder) || 0
  };
}

function computeTaskTotals_(tasks, email) {
  const totals = {
    count: tasks.length,
    byStatus: {
      Todo: 0,
      Doing: 0,
      Done: 0,
      Blocked: 0
    },
    totalEstimatedHours: 0,
    totalSeconds: 0,
    mySeconds: 0
  };
  const lowerEmail = (email || '').toLowerCase();
  tasks.forEach(function (task) {
    if (totals.byStatus[task.status] !== undefined) {
      totals.byStatus[task.status] += 1;
    }
    totals.totalEstimatedHours += Number(task.estimatedHours) || 0;
    totals.totalSeconds += Number(task.totalSeconds) || 0;
    if (task.assignedTo && task.assignedTo.toLowerCase() === lowerEmail) {
      totals.mySeconds += Number(task.totalSeconds) || 0;
    }
  });
  return totals;
}

function buildVisibilityEmails_(viewer, users) {
  const my = viewer.email ? [viewer.email.toLowerCase()] : [];
  let team = [];
  let all = [];
  if (isRoleAtLeast_(viewer.role, 'ADMIN')) {
    all = users.filter(function (u) { return u.isActive; }).map(function (u) { return u.email.toLowerCase(); });
    team = all.filter(function (email) { return my.indexOf(email) === -1; });
  } else if (viewer.role === 'MANAGER') {
    team = users.filter(function (u) { return u.isActive && u.parentUserId === viewer.id; }).map(function (u) { return u.email.toLowerCase(); });
    all = dedupeArray_(my.concat(team));
  } else {
    all = my.slice();
  }
  return {
    my: dedupeArray_(my),
    team: dedupeArray_(team),
    all: dedupeArray_(all.length ? all : my)
  };
}

function addDays_(isoDate, days) {
  if (!isoDate) {
    return '';
  }
  const date = new Date(isoDate + 'T00:00:00');
  date.setDate(date.getDate() + days);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeDate_(value) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }
  return '';
}

function canInteractWithTask_(viewer, taskRecord, usersTable) {
  if (isRoleAtLeast_(viewer.role, 'ADMIN')) {
    return true;
  }
  const users = usersTable.map(function (entry) {
    return sanitizeUser_(entry.record);
  });
  const visibility = buildVisibilityEmails_(viewer, users);
  const assigned = (taskRecord.assignedTo || '').toLowerCase();
  return assigned && visibility.all.indexOf(assigned) !== -1;
}

function accumulateTaskTime_(env, taskId) {
  const timeSheet = env.sheets.TimeLog.sheet;
  const timeHeaders = SHEET_DEFINITIONS.TimeLog.headers;
  const logs = loadTable_(timeSheet, timeHeaders).filter(function (row) {
    return row.record.taskId === taskId;
  });
  const total = logs.reduce(function (sum, row) {
    return sum + (Number(row.record.durationSeconds) || 0);
  }, 0);
  const taskSheet = env.sheets.Tasks.sheet;
  const taskHeaders = SHEET_DEFINITIONS.Tasks.headers;
  const tasks = loadTable_(taskSheet, taskHeaders);
  const entry = tasks.find(function (row) {
    return row.record.id === taskId;
  });
  if (entry) {
    entry.record.totalSeconds = total;
    entry.record.updatedAt = todayISO_();
    writeRecord_(taskSheet, taskHeaders, entry.rowNumber, entry.record);
  }
}

function toBool_(value) {
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'string') {
    return value.toUpperCase() === 'TRUE';
  }
  return !!value;
}

function getNotificationPrefs_(userId) {
  if (!userId) {
    return { instant: true, daily: true, weekly: true };
  }
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('notify_' + userId);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      return {
        instant: parsed.instant !== false,
        daily: parsed.daily !== false,
        weekly: parsed.weekly !== false
      };
    } catch (err) {
      return { instant: true, daily: true, weekly: true };
    }
  }
  return { instant: true, daily: true, weekly: true };
}

function saveNotificationPrefs_(userId, prefs) {
  if (!userId) {
    return;
  }
  const merged = {
    instant: prefs.instant !== false,
    daily: prefs.daily !== false,
    weekly: prefs.weekly !== false
  };
  PropertiesService.getScriptProperties().setProperty('notify_' + userId, JSON.stringify(merged));
}

function dedupeArray_(list) {
  const seen = {};
  return (list || []).filter(function (item) {
    const key = String(item).toLowerCase();
    if (seen[key]) {
      return false;
    }
    seen[key] = true;
    return true;
  });
}

function sanitizeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function brandFooterHtml_() {
  return '<div style="margin-top:24px;font-size:12px;color:#5f6368;line-height:1.4;">' +
    'AuraFlow by Aura One Labs<br>' +
    '<span style="display:inline-flex;align-items:center;gap:4px;">' +
    '<img src="' + BRAND_LOGO_URL + '" alt="AuraFlow" style="height:14px;width:14px;">' +
    'Rare Aura Media Group</span>' +
    '</div>';
}

function getAppUrl_() {
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) {
      return url;
    }
  } catch (err) {
    Logger.log('App URL unavailable: ' + err);
  }
  return 'https://script.google.com/macros/s/' + ScriptApp.getScriptId() + '/exec';
}

