/**
 * Booking with approval
 * ---------------------
 * Guest picks a free slot -> a [PENDING] hold is placed in your Google Calendar
 * -> you get an email with Approve (Meet / Teams / Zoom) and Reject
 * -> on approval the guest gets a calendar invite + confirmation email with the meeting link.
 *
 * Availability = free in Google Calendar AND (if connected) in Outlook / Microsoft 365.
 * Outlook / Teams is optional. Without the MS_* script properties it runs Google-only.
 * Zoom is optional. Without the ZOOM_* script properties the Zoom button is hidden.
 */

const DEFAULTS = {
  OWNER_NAME: 'Isaac Touka Tchoumkeu',
  OWNER_EMAIL: '',                 // empty = the Google account that deploys the script
  CALENDAR_ID: 'primary',
  TIMEZONE: 'Europe/Berlin',
  SLOT_MINUTES: 30,
  BUFFER_MINUTES: 10,              // free time required before and after each meeting
  MIN_NOTICE_HOURS: 12,            // earliest bookable slot from now
  DAYS_AHEAD: 21,                  // how far ahead guests can book (max 60 with Outlook)
  // Working hours per weekday, 1 = Monday ... 7 = Sunday, in TIMEZONE
  WORKING_HOURS: {
    1: [['09:00', '12:00'], ['13:00', '17:00']],
    2: [['09:00', '12:00'], ['13:00', '17:00']],
    3: [['09:00', '12:00'], ['13:00', '17:00']],
    4: [['09:00', '12:00'], ['13:00', '17:00']],
    5: [['09:00', '12:00'], ['13:00', '16:00']],
  },
  PENDING_EXPIRY_HOURS: 48,        // unanswered requests are released after this
  MAX_PENDING_PER_EMAIL: 2,        // simple spam protection
  MIRROR_TO_OUTLOOK: true,         // when approving with Meet or Zoom, also block the time in Outlook
  SEND_CONFIRMATION_EMAIL: true,   // confirmation email on top of the calendar invite
  MEETING_TITLE: 'Meeting with {name}',
  BOOKING_ENABLED: true,           // false = booking page shows PAUSED_MESSAGE
  PAUSED_MESSAGE: 'Bookings are paused at the moment. Please check back later.',
  CHECK_GOOGLE_BUSY: true,         // count busy times of the main Google calendar
  EXTRA_BUSY_CALENDARS: [],        // more Google calendar IDs whose busy times count
  CHECK_OUTLOOK_BUSY: true,        // count Outlook busy times (when Outlook is connected)
  PLATFORM_MEET: true,             // which meeting platforms you can approve with
  PLATFORM_TEAMS: true,
  PLATFORM_ZOOM: true,
};

/**
 * Effective settings = DEFAULTS above, overridden by what you save on the admin page.
 * The values above are only the starting point once the admin page is used.
 */
let SETTINGS_CACHE_ = null;
function C() {
  if (!SETTINGS_CACHE_) {
    let saved = {};
    try { saved = JSON.parse(props_().getProperty('SETTINGS') || '{}'); } catch (err) { saved = {}; }
    SETTINGS_CACHE_ = Object.assign({}, DEFAULTS, saved);
  }
  return SETTINGS_CACHE_;
}

function meetOn_() { return !!C().PLATFORM_MEET; }
function teamsOn_() { return msEnabled_() && !!C().PLATFORM_TEAMS; }
function zoomOn_() { return zoomEnabled_() && !!C().PLATFORM_ZOOM; }
function calId_(req) { return (req && req.calendarId) || C().CALENDAR_ID; }

/* ============================ Web entry ============================ */

function doGet(e) {
  const p = (e && e.parameter) || {};

  // Hidden admin page: only opens with the right key, otherwise the normal booking page is shown
  if (p.action === 'admin' && adminKeyOk_(p.key)) {
    const t = HtmlService.createTemplateFromFile('Admin');
    t.data = safeJson_({ key: p.key });
    return t.evaluate()
      .setTitle('Booking admin')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (p.action === 'review') {
    const req = getRequest_(p.id);
    if (!req || req.token !== p.token) {
      return simplePage_('Link not valid', 'This booking request was not found or the link is incomplete.');
    }
    const t = HtmlService.createTemplateFromFile('Approve');
    t.data = safeJson_({
      id: req.id,
      token: req.token,
      name: req.name,
      email: req.email,
      topic: req.topic,
      when: formatRange_(req.start, req.end, C().TIMEZONE),
      status: req.status,
      platform: req.platform || '',
      link: req.meetingLink || '',
      preselect: ['meet', 'teams', 'zoom', 'reject'].indexOf(p.do) >= 0 ? p.do : '',
      meetEnabled: meetOn_(),
      outlookEnabled: teamsOn_(),
      zoomEnabled: zoomOn_(),
    });
    return t.evaluate()
      .setTitle('Review booking request')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const t = HtmlService.createTemplateFromFile('Booking');
  t.data = safeJson_({
    ownerName: C().OWNER_NAME, slotMinutes: C().SLOT_MINUTES,
    paused: !C().BOOKING_ENABLED, pausedMessage: C().PAUSED_MESSAGE,
  });
  return t.evaluate()
    .setTitle('Book a meeting with ' + C().OWNER_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ============================ Called from the booking page ============================ */

function getSlots() {
  if (!C().BOOKING_ENABLED) return [];
  const now = new Date();
  const earliest = now.getTime() + C().MIN_NOTICE_HOURS * 3600e3;
  const today = Utilities.formatDate(now, C().TIMEZONE, 'yyyy-MM-dd').split('-').map(Number);
  const rangeStart = zoned_(today[0], today[1], today[2], 0, 0);
  const rangeEnd = zoned_(today[0], today[1], today[2] + C().DAYS_AHEAD + 1, 0, 0);
  const busy = getBusy_(rangeStart, rangeEnd);
  const slotMs = C().SLOT_MINUTES * 60e3;
  const slots = [];

  for (let i = 0; i <= C().DAYS_AHEAD; i++) {
    const noon = zoned_(today[0], today[1], today[2] + i, 12, 0);
    const dow = Number(Utilities.formatDate(noon, C().TIMEZONE, 'u')); // 1 = Mon ... 7 = Sun
    const parts = Utilities.formatDate(noon, C().TIMEZONE, 'yyyy-MM-dd').split('-').map(Number);
    (C().WORKING_HOURS[dow] || []).forEach(function (win) {
      const from = win[0].split(':').map(Number);
      const to = win[1].split(':').map(Number);
      let s = zoned_(parts[0], parts[1], parts[2], from[0], from[1]).getTime();
      const winEnd = zoned_(parts[0], parts[1], parts[2], to[0], to[1]).getTime();
      while (s + slotMs <= winEnd) {
        if (s >= earliest && isFree_(s, s + slotMs, busy)) {
          slots.push({ start: new Date(s).toISOString(), end: new Date(s + slotMs).toISOString() });
        }
        s += slotMs;
      }
    });
  }
  return slots;
}

function submitBooking(form) {
  form = form || {};
  if (form.website) return { ok: true }; // honeypot: bots fill hidden fields
  if (!C().BOOKING_ENABLED) throw new Error(C().PAUSED_MESSAGE);

  const name = clean_(form.name, 100);
  const email = clean_(form.email, 200).toLowerCase();
  const topic = clean_(form.topic, 1000);
  const start = new Date(form.start);
  let guestTz = clean_(form.timezone, 60);

  if (!name) throw new Error('Please enter your name.');
  if (!/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(email)) throw new Error('Please enter a valid email address.');
  if (isNaN(start.getTime())) throw new Error('Please pick a time slot.');
  try { Utilities.formatDate(start, guestTz || C().TIMEZONE, 'HH:mm'); } catch (err) { guestTz = ''; }

  const end = new Date(start.getTime() + C().SLOT_MINUTES * 60e3);
  let req;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const pendingForEmail = allRequests_().filter(function (r) {
      return r.status === 'pending' && r.email === email;
    }).length;
    if (pendingForEmail >= C().MAX_PENDING_PER_EMAIL) {
      throw new Error('You already have open requests. Please wait for a reply before booking again.');
    }
    const stillFree = getSlots().some(function (s) { return s.start === start.toISOString(); });
    if (!stillFree) throw new Error('Sorry, this time was just taken. Please pick another slot.');

    const id = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    const token = Utilities.getUuid();
    const hold = Calendar.Events.insert({
      summary: '[PENDING] ' + name,
      description: 'Booking request, waiting for your approval.\n\nName: ' + name + '\nEmail: ' + email +
        '\nTopic: ' + (topic || '-') + '\n\nReview: ' + reviewUrl_(id, token, ''),
      start: { dateTime: start.toISOString() },
      end: { dateTime: end.toISOString() },
      transparency: 'opaque',
      colorId: '5', // yellow
      reminders: { useDefault: false, overrides: [] },
    }, C().CALENDAR_ID);

    req = {
      id: id, token: token, name: name, email: email, topic: topic, guestTz: guestTz,
      start: start.toISOString(), end: end.toISOString(),
      status: 'pending', googleEventId: hold.id, calendarId: C().CALENDAR_ID, created: new Date().toISOString(),
    };
    saveRequest_(req);
  } finally {
    lock.releaseLock();
  }

  notifyOwner_(req);
  emailGuest_(req, 'Request received: ' + shortWhen_(req),
    'Thanks ' + esc_(req.name) + ', your request for <b>' + esc_(guestWhen_(req)) + '</b> has been received.' +
    '<br><br>You will get a confirmation with the meeting link as soon as it is approved.');
  return { ok: true, when: guestWhen_(req) };
}

/* ============================ Called from the review page ============================ */

function approveRequest(id, token, platform) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let req;
  try {
    req = checkRequest_(id, token);
    const title = C().MEETING_TITLE.replace('{name}', req.name);
    const desc = 'Topic: ' + (req.topic || '-') + '\nGuest: ' + req.name + ' <' + req.email + '>';
    let link = '';

    if (platform === 'teams') {
      if (!teamsOn_()) throw new Error('Teams is not connected or is switched off on the admin page.');
      const ev = msCreateEvent_({
        subject: title, body: desc, start: req.start, end: req.end,
        attendees: [{ email: req.email, name: req.name }], teams: true,
      });
      link = (ev.onlineMeeting && ev.onlineMeeting.joinUrl) || msWaitJoinUrl_(ev.id);
      req.outlookEventId = ev.id;
      // keep the Google hold as a confirmed block with the Teams link
      Calendar.Events.patch({
        summary: title + ' (Teams)', location: link, colorId: '9',
        description: desc + '\n\nTeams: ' + link,
      }, calId_(req), req.googleEventId);
    } else if (platform === 'zoom') {
      if (!zoomOn_()) throw new Error('Zoom is not connected or is switched off on the admin page.');
      const zm = zoomCreateMeeting_(title, req);
      link = zm.join_url;
      req.zoomMeetingId = zm.id;
      // the Google hold becomes the real event with the Zoom link, Google sends the invite
      Calendar.Events.patch({
        summary: title + ' (Zoom)', location: link, colorId: '7',
        description: desc + '\n\nZoom: ' + link +
          (zm.id ? '\nMeeting ID: ' + zm.id : '') + (zm.password ? '\nPasscode: ' + zm.password : ''),
        attendees: [{ email: req.email, displayName: req.name }],
      }, calId_(req), req.googleEventId, { sendUpdates: 'all' });
      mirrorToOutlook_(req, title + ' (Zoom)', desc + '\n\nZoom: ' + link);
    } else {
      if (!meetOn_()) throw new Error('Google Meet is switched off on the admin page.');
      platform = 'meet';
      let ev = Calendar.Events.patch({
        summary: title, description: desc, colorId: '10',
        attendees: [{ email: req.email, displayName: req.name }],
        conferenceData: { createRequest: { requestId: req.id, conferenceSolutionKey: { type: 'hangoutsMeet' } } },
      }, calId_(req), req.googleEventId, { conferenceDataVersion: 1, sendUpdates: 'all' });
      link = meetLink_(ev);
      for (let i = 0; !link && i < 5; i++) {
        Utilities.sleep(1500);
        ev = Calendar.Events.get(calId_(req), req.googleEventId);
        link = meetLink_(ev);
      }
      mirrorToOutlook_(req, title + ' (Google Meet)', desc + '\n\nMeet: ' + link);
    }

    req.status = 'approved';
    req.platform = platform;
    req.meetingLink = link;
    req.decided = new Date().toISOString();
    saveRequest_(req);
  } finally {
    lock.releaseLock();
  }

  if (C().SEND_CONFIRMATION_EMAIL) {
    const tool = { teams: 'Microsoft Teams', zoom: 'Zoom', meet: 'Google Meet' }[req.platform];
    emailGuest_(req, 'Confirmed: ' + shortWhen_(req),
      'Hi ' + esc_(req.name) + ', your meeting with ' + esc_(C().OWNER_NAME) + ' is confirmed.' +
      '<br><br><b>When:</b> ' + esc_(guestWhen_(req)) +
      '<br><b>Where:</b> ' + tool +
      (req.topic ? '<br><b>Topic:</b> ' + esc_(req.topic) : '') +
      '<br><br>A calendar invite is on its way as well.',
      req.meetingLink ? [{ label: 'Join with ' + tool, url: req.meetingLink }] : []);
  }
  return { ok: true, link: req.meetingLink, platform: req.platform };
}

function rejectRequest(id, token, message) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let req;
  try {
    req = checkRequest_(id, token);
    removeGoogleHold_(req);
    req.status = 'rejected';
    req.decided = new Date().toISOString();
    saveRequest_(req);
  } finally {
    lock.releaseLock();
  }
  const note = clean_(message, 1000);
  emailGuest_(req, 'About your request for ' + shortWhen_(req),
    'Hi ' + esc_(req.name) + ', unfortunately the requested time (' + esc_(guestWhen_(req)) + ') does not work.' +
    (note ? '<br><br><i>' + esc_(note).replace(/\n/g, '<br>') + '</i>' : '') +
    '<br><br>Feel free to pick another time.',
    [{ label: 'Pick another time', url: webAppUrl_() }]);
  return { ok: true };
}

/* ============================ Maintenance ============================ */

/** Runs hourly (see installTriggers). Releases unanswered holds and clears old records. */
function cleanupExpired() {
  const now = Date.now();
  allRequests_().forEach(function (r) {
    if (r.status === 'pending') {
      const tooOld = now - new Date(r.created).getTime() > C().PENDING_EXPIRY_HOURS * 3600e3;
      const tooLate = new Date(r.start).getTime() < now + 3600e3;
      if (tooOld || tooLate) {
        removeGoogleHold_(r);
        r.status = 'expired';
        saveRequest_(r);
        emailGuest_(r, 'About your request for ' + shortWhen_(r),
          'Hi ' + esc_(r.name) + ', your request for ' + esc_(guestWhen_(r)) +
          ' could not be confirmed in time, so the slot was released. Sorry about that.',
          [{ label: 'Pick another time', url: webAppUrl_() }]);
      }
    } else if (new Date(r.end).getTime() < now - 30 * 86400e3) {
      props_().deleteProperty('req_' + r.id);
    }
  });
}

/** Run once after deploying. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupExpired') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cleanupExpired').timeBased().everyHours(1).create();
  Logger.log('Hourly cleanup trigger installed.');
}

/** Run from the editor to authorize and check everything. */
function testSetup() {
  Logger.log('Owner email: ' + ownerEmail_());
  Logger.log('Web app URL: ' + (webAppUrl_() || '(deploy first)'));
  Logger.log('Outlook connected: ' + msEnabled_());
  if (msEnabled_()) {
    msToken_();
    Logger.log('Microsoft login OK. Outlook busy blocks next 7 days: ' +
      msBusy_(new Date(), new Date(Date.now() + 7 * 86400e3)).length);
  }
  Logger.log('Zoom connected: ' + zoomEnabled_());
  if (zoomEnabled_()) { zoomToken_(); Logger.log('Zoom login OK.'); }
  Logger.log('Free slots found: ' + getSlots().length);
}

/** Lists all requests in the log. */
function listRequests() {
  allRequests_().sort(function (a, b) { return a.start < b.start ? -1 : 1; }).forEach(function (r) {
    Logger.log(r.status.toUpperCase() + '  ' + formatRange_(r.start, r.end, C().TIMEZONE) + '  ' + r.name + ' <' + r.email + '>');
  });
}

/* ============================ Admin page ============================ */

/** Run once from the editor: creates the secret admin key and logs your private admin URL. */
function setupAdmin() {
  let key = props_().getProperty('ADMIN_KEY');
  if (!key) {
    key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props_().setProperty('ADMIN_KEY', key);
  }
  Logger.log('Your private admin page (bookmark it, do not share it):\n' + webAppUrl_() + '?action=admin&key=' + key);
}

/** Run from the editor if the admin URL ever leaks. The old link stops working immediately. */
function rotateAdminKey() {
  props_().deleteProperty('ADMIN_KEY');
  setupAdmin();
}

function adminKeyOk_(key) {
  const k = props_().getProperty('ADMIN_KEY');
  return !!k && typeof key === 'string' && key.length >= 32 && key === k;
}

function requireAdmin_(key) {
  if (!adminKeyOk_(key)) {
    Utilities.sleep(1500); // slows down guessing
    throw new Error('Not allowed.');
  }
}

const SECRET_KEYS_ = ['MS_CLIENT_SECRET', 'ZOOM_CLIENT_SECRET'];
const CONNECTION_KEYS_ = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MS_USER',
  'ZOOM_ACCOUNT_ID', 'ZOOM_CLIENT_ID', 'ZOOM_CLIENT_SECRET', 'ZOOM_USER', 'WEB_APP_URL'];

function adminGetState(key) {
  requireAdmin_(key);
  const p = props_();

  let calendars = [];
  let calendarError = '';
  try {
    let pageToken;
    do {
      const r = Calendar.CalendarList.list({ maxResults: 250, pageToken: pageToken });
      (r.items || []).forEach(function (c) {
        calendars.push({ id: c.id, name: c.summaryOverride || c.summary, primary: !!c.primary, access: c.accessRole });
      });
      pageToken = r.nextPageToken;
    } while (pageToken);
  } catch (err) { calendarError = String(err.message || err); }

  const now = Date.now();
  const requests = allRequests_()
    .filter(function (r) { return r.status === 'pending' || (r.status === 'approved' && new Date(r.end).getTime() > now); })
    .sort(function (a, b) { return a.start < b.start ? -1 : 1; })
    .map(function (r) {
      return {
        name: r.name, email: r.email, topic: r.topic, status: r.status, platform: r.platform || '',
        link: r.meetingLink || '', when: formatRange_(r.start, r.end, C().TIMEZONE),
        reviewUrl: reviewUrl_(r.id, r.token, ''),
      };
    });

  const connections = {};
  CONNECTION_KEYS_.forEach(function (k) {
    const v = p.getProperty(k) || '';
    connections[k] = SECRET_KEYS_.indexOf(k) >= 0 ? (v ? 'set' : '') : v; // secrets are never sent to the page
  });

  return {
    settings: C(),
    calendars: calendars,
    calendarError: calendarError,
    outlookConnected: msEnabled_(),
    zoomConnected: zoomEnabled_(),
    connections: connections,
    requests: requests,
    bookingUrl: webAppUrl_(),
    ownerEmail: ownerEmail_(),
  };
}

function adminSave(key, s) {
  requireAdmin_(key);
  s = s || {};
  const out = {};
  function int(v, min, max, label) {
    const n = Math.round(Number(v));
    if (!isFinite(n) || n < min || n > max) throw new Error(label + ' must be between ' + min + ' and ' + max + '.');
    return n;
  }

  out.OWNER_NAME = clean_(s.OWNER_NAME, 100) || DEFAULTS.OWNER_NAME;
  out.OWNER_EMAIL = clean_(s.OWNER_EMAIL, 200);
  if (out.OWNER_EMAIL && !/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(out.OWNER_EMAIL)) throw new Error('Notification email is not valid.');
  out.MEETING_TITLE = clean_(s.MEETING_TITLE, 150) || DEFAULTS.MEETING_TITLE;
  out.PAUSED_MESSAGE = clean_(s.PAUSED_MESSAGE, 500) || DEFAULTS.PAUSED_MESSAGE;

  out.TIMEZONE = clean_(s.TIMEZONE, 60) || DEFAULTS.TIMEZONE;
  try { Utilities.formatDate(new Date(), out.TIMEZONE, 'Z'); } catch (err) { throw new Error('Unknown time zone: ' + out.TIMEZONE); }

  out.CALENDAR_ID = clean_(s.CALENDAR_ID, 300) || 'primary';
  try {
    const cal = Calendar.CalendarList.get(out.CALENDAR_ID);
    if (cal.accessRole !== 'owner' && cal.accessRole !== 'writer') throw new Error('no write access');
  } catch (err) {
    throw new Error('The main calendar must be one you can edit (' + out.CALENDAR_ID + ').');
  }
  out.EXTRA_BUSY_CALENDARS = (Array.isArray(s.EXTRA_BUSY_CALENDARS) ? s.EXTRA_BUSY_CALENDARS : [])
    .map(function (x) { return clean_(x, 300); }).filter(Boolean).slice(0, 20);

  out.SLOT_MINUTES = int(s.SLOT_MINUTES, 5, 480, 'Meeting length');
  out.BUFFER_MINUTES = int(s.BUFFER_MINUTES, 0, 240, 'Buffer');
  out.MIN_NOTICE_HOURS = int(s.MIN_NOTICE_HOURS, 0, 720, 'Minimum notice');
  out.DAYS_AHEAD = int(s.DAYS_AHEAD, 1, 60, 'Days ahead');
  out.PENDING_EXPIRY_HOURS = int(s.PENDING_EXPIRY_HOURS, 1, 720, 'Request expiry');
  out.MAX_PENDING_PER_EMAIL = int(s.MAX_PENDING_PER_EMAIL, 1, 20, 'Open requests per email');
  out.WORKING_HOURS = parseHours_(s.WORKING_HOURS);

  ['BOOKING_ENABLED', 'CHECK_GOOGLE_BUSY', 'CHECK_OUTLOOK_BUSY', 'MIRROR_TO_OUTLOOK', 'SEND_CONFIRMATION_EMAIL',
    'PLATFORM_MEET', 'PLATFORM_TEAMS', 'PLATFORM_ZOOM'].forEach(function (k) { out[k] = s[k] === true; });

  if (!out.PLATFORM_MEET && !(out.PLATFORM_TEAMS && msEnabled_()) && !(out.PLATFORM_ZOOM && zoomEnabled_())) {
    throw new Error('Keep at least one connected meeting platform switched on.');
  }

  props_().setProperty('SETTINGS', JSON.stringify(out));
  SETTINGS_CACHE_ = null;
  return adminGetState(key);
}

/** {1: "09:00-12:00, 13:00-17:00", 6: ""} -> {1: [["09:00","12:00"],["13:00","17:00"]]} */
function parseHours_(wh) {
  const names = ['', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const out = {};
  for (let d = 1; d <= 7; d++) {
    const text = clean_(wh && wh[d], 200);
    if (!text) continue;
    const wins = text.split(',').map(function (part) {
      const m = part.trim().match(/^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/);
      if (!m) throw new Error(names[d] + ': use the format 09:00-12:00, 13:00-17:00');
      const from = Number(m[1]) * 60 + Number(m[2]);
      const to = Number(m[3]) * 60 + Number(m[4]);
      if (Number(m[2]) > 59 || Number(m[4]) > 59 || to > 1440 || from >= to) throw new Error(names[d] + ': ' + part.trim() + ' is not a valid time range.');
      return [from, to];
    }).sort(function (a, b) { return a[0] - b[0]; });
    for (let i = 1; i < wins.length; i++) {
      if (wins[i][0] < wins[i - 1][1]) throw new Error(names[d] + ': time ranges overlap.');
    }
    const hhmm = function (n) { return ('0' + Math.floor(n / 60)).slice(-2) + ':' + ('0' + (n % 60)).slice(-2); };
    out[d] = wins.map(function (w) { return [hhmm(w[0]), hhmm(w[1])]; });
  }
  return out;
}

/** Saves connection values. Empty fields keep the stored value, null removes it. */
function adminSaveConnections(key, values) {
  requireAdmin_(key);
  const p = props_();
  CONNECTION_KEYS_.forEach(function (k) {
    if (!values || !(k in values)) return;
    const v = values[k];
    if (v === null) p.deleteProperty(k);
    else if (typeof v === 'string' && v.trim()) p.setProperty(k, v.trim().slice(0, 500));
  });
  const cache = CacheService.getScriptCache();
  cache.remove('ms_token');
  cache.remove('zoom_token');
  return adminGetState(key);
}

function adminTest(key) {
  requireAdmin_(key);
  const results = [];
  function run(name, fn) {
    try { results.push({ name: name, ok: true, message: fn() }); }
    catch (err) { results.push({ name: name, ok: false, message: String(err.message || err) }); }
  }
  const start = new Date(), end = new Date(Date.now() + 7 * 86400e3);
  run('Google Calendar', function () {
    const fb = Calendar.Freebusy.query({ timeMin: start.toISOString(), timeMax: end.toISOString(), items: [{ id: C().CALENDAR_ID }] });
    const cal = fb.calendars[Object.keys(fb.calendars)[0]] || {};
    if (cal.errors && cal.errors.length) throw new Error(JSON.stringify(cal.errors));
    return 'OK, ' + (cal.busy || []).length + ' busy blocks in the next 7 days';
  });
  if (msEnabled_()) run('Outlook', function () { return 'OK, ' + msBusy_(start, end).length + ' busy blocks in the next 7 days'; });
  else results.push({ name: 'Outlook', ok: false, message: 'Not connected' });
  if (zoomEnabled_()) run('Zoom', function () { zoomToken_(); return 'OK, login works'; });
  else results.push({ name: 'Zoom', ok: false, message: 'Not connected' });
  run('Booking page', function () {
    const n = getSlots().length;
    return C().BOOKING_ENABLED ? n + ' free slots shown to guests' : 'Paused, guests see the pause message';
  });
  return results;
}

/* ============================ Availability ============================ */

function getBusy_(start, end) {
  const busy = [];
  // Pending and approved requests always block their slot, whatever the calendar switches say
  allRequests_().forEach(function (r) {
    if (r.status === 'pending' || r.status === 'approved') {
      busy.push({ start: new Date(r.start).getTime(), end: new Date(r.end).getTime() });
    }
  });

  const ids = (C().CHECK_GOOGLE_BUSY ? [C().CALENDAR_ID] : []).concat(C().EXTRA_BUSY_CALENDARS || [])
    .filter(function (id, i, a) { return id && a.indexOf(id) === i; });
  if (ids.length) {
    const fb = Calendar.Freebusy.query({
      timeMin: start.toISOString(), timeMax: end.toISOString(),
      items: ids.map(function (id) { return { id: id }; }),
    });
    Object.keys(fb.calendars || {}).forEach(function (k) {
      const cal = fb.calendars[k];
      if (cal.errors && cal.errors.length) throw new Error('Cannot read calendar ' + k + ': ' + JSON.stringify(cal.errors));
      (cal.busy || []).forEach(function (b) {
        busy.push({ start: new Date(b.start).getTime(), end: new Date(b.end).getTime() });
      });
    });
  }
  if (msEnabled_() && C().CHECK_OUTLOOK_BUSY) msBusy_(start, end).forEach(function (b) { busy.push(b); });
  return busy;
}

function isFree_(s, e, busy) {
  const buf = C().BUFFER_MINUTES * 60e3;
  return !busy.some(function (b) { return b.start < e + buf && b.end > s - buf; });
}

/* ============================ Microsoft 365 (Outlook + Teams) ============================ */

function msEnabled_() {
  const p = props_();
  return !!(p.getProperty('MS_TENANT_ID') && p.getProperty('MS_CLIENT_ID') &&
            p.getProperty('MS_CLIENT_SECRET') && p.getProperty('MS_USER'));
}

function msToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('ms_token');
  if (cached) return cached;
  const p = props_();
  const res = UrlFetchApp.fetch('https://login.microsoftonline.com/' + p.getProperty('MS_TENANT_ID') + '/oauth2/v2.0/token', {
    method: 'post',
    payload: {
      client_id: p.getProperty('MS_CLIENT_ID'),
      client_secret: p.getProperty('MS_CLIENT_SECRET'),
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    },
    muteHttpExceptions: true,
  });
  const j = JSON.parse(res.getContentText());
  if (!j.access_token) throw new Error('Microsoft login failed: ' + (j.error_description || res.getContentText()));
  cache.put('ms_token', j.access_token, Math.min(21600, Math.max(60, (j.expires_in || 3600) - 300)));
  return j.access_token;
}

function msFetch_(method, path, body) {
  const opts = {
    method: method,
    headers: { Authorization: 'Bearer ' + msToken_(), Prefer: 'outlook.timezone="UTC"' },
    muteHttpExceptions: true,
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  const res = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0' + path, opts);
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code >= 300) throw new Error('Microsoft Graph error ' + code + ': ' + text);
  return text ? JSON.parse(text) : {};
}

function msUserPath_() {
  return '/users/' + encodeURIComponent(props_().getProperty('MS_USER'));
}

function msBusy_(start, end) {
  const user = props_().getProperty('MS_USER');
  const j = msFetch_('post', msUserPath_() + '/calendar/getSchedule', {
    schedules: [user],
    startTime: { dateTime: toGraph_(start), timeZone: 'UTC' },
    endTime: { dateTime: toGraph_(end), timeZone: 'UTC' },
    availabilityViewInterval: 15,
  });
  const sched = (j.value || [])[0] || {};
  if (sched.error) throw new Error('Outlook availability error: ' + JSON.stringify(sched.error));
  return (sched.scheduleItems || [])
    .filter(function (i) { return i.status !== 'free' && i.status !== 'workingElsewhere'; })
    .map(function (i) { return { start: fromGraph_(i.start.dateTime), end: fromGraph_(i.end.dateTime) }; });
}

function msCreateEvent_(o) {
  const body = {
    subject: o.subject,
    body: { contentType: 'text', content: o.body || '' },
    start: { dateTime: toGraph_(new Date(o.start)), timeZone: 'UTC' },
    end: { dateTime: toGraph_(new Date(o.end)), timeZone: 'UTC' },
    showAs: 'busy',
    attendees: (o.attendees || []).map(function (a) {
      return { emailAddress: { address: a.email, name: a.name }, type: 'required' };
    }),
  };
  if (o.teams) { body.isOnlineMeeting = true; body.onlineMeetingProvider = 'teamsForBusiness'; }
  return msFetch_('post', msUserPath_() + '/events', body);
}

function msWaitJoinUrl_(eventId) {
  for (let i = 0; i < 5; i++) {
    Utilities.sleep(1500);
    const ev = msFetch_('get', msUserPath_() + '/events/' + encodeURIComponent(eventId) + '?$select=onlineMeeting');
    if (ev.onlineMeeting && ev.onlineMeeting.joinUrl) return ev.onlineMeeting.joinUrl;
  }
  return '';
}

/** When approving with Meet or Zoom, also block the time in Outlook (no attendees). */
function mirrorToOutlook_(req, subject, body) {
  if (!C().MIRROR_TO_OUTLOOK || !msEnabled_()) return;
  const block = msCreateEvent_({ subject: subject, body: body, start: req.start, end: req.end, attendees: [], teams: false });
  req.outlookEventId = block.id;
}

/* ============================ Zoom ============================ */

function zoomEnabled_() {
  const p = props_();
  return !!(p.getProperty('ZOOM_ACCOUNT_ID') && p.getProperty('ZOOM_CLIENT_ID') && p.getProperty('ZOOM_CLIENT_SECRET'));
}

function zoomToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('zoom_token');
  if (cached) return cached;
  const p = props_();
  const res = UrlFetchApp.fetch('https://zoom.us/oauth/token?grant_type=account_credentials&account_id=' +
    encodeURIComponent(p.getProperty('ZOOM_ACCOUNT_ID')), {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(p.getProperty('ZOOM_CLIENT_ID') + ':' + p.getProperty('ZOOM_CLIENT_SECRET')) },
    muteHttpExceptions: true,
  });
  const j = JSON.parse(res.getContentText());
  if (!j.access_token) throw new Error('Zoom login failed: ' + (j.reason || j.error || res.getContentText()));
  cache.put('zoom_token', j.access_token, Math.min(21600, Math.max(60, (j.expires_in || 3600) - 300)));
  return j.access_token;
}

function zoomCreateMeeting_(title, req) {
  const user = props_().getProperty('ZOOM_USER') || 'me';
  const res = UrlFetchApp.fetch('https://api.zoom.us/v2/users/' + encodeURIComponent(user) + '/meetings', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + zoomToken_() },
    payload: JSON.stringify({
      topic: title,
      type: 2, // scheduled meeting
      start_time: new Date(req.start).toISOString().replace(/\.\d{3}Z$/, 'Z'),
      duration: Math.round((new Date(req.end) - new Date(req.start)) / 60000),
      timezone: C().TIMEZONE,
      agenda: (req.topic || '').slice(0, 2000),
      settings: { join_before_host: false, waiting_room: true, host_video: true, participant_video: true },
    }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('Zoom error ' + code + ': ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

function toGraph_(d) { return d.toISOString().replace('Z', ''); }
function fromGraph_(s) { return new Date(String(s).replace(/(\.\d{3})\d*$/, '$1') + 'Z').getTime(); }

/* ============================ Emails ============================ */

function notifyOwner_(req) {
  const buttons = [];
  if (meetOn_()) buttons.push({ label: 'Approve with Google Meet', url: reviewUrl_(req.id, req.token, 'meet') });
  if (teamsOn_()) buttons.push({ label: 'Approve with Teams', url: reviewUrl_(req.id, req.token, 'teams') });
  if (zoomOn_()) buttons.push({ label: 'Approve with Zoom', url: reviewUrl_(req.id, req.token, 'zoom') });
  buttons.push({ label: 'Reject', url: reviewUrl_(req.id, req.token, 'reject'), muted: true });
  MailApp.sendEmail({
    to: ownerEmail_(),
    subject: 'Booking request: ' + req.name + ', ' + shortWhen_(req),
    htmlBody: emailHtml_('New booking request',
      '<b>' + esc_(req.name) + '</b> (' + esc_(req.email) + ') would like to meet.' +
      '<br><br><b>When:</b> ' + esc_(formatRange_(req.start, req.end, C().TIMEZONE)) +
      '<br><b>Topic:</b> ' + (req.topic ? esc_(req.topic).replace(/\n/g, '<br>') : '-') +
      '<br><br>The slot is held in your calendar until you decide. Each button opens a page where you confirm.',
      buttons),
    name: 'Booking',
  });
}

function emailGuest_(req, subject, html, buttons) {
  MailApp.sendEmail({
    to: req.email,
    subject: subject,
    htmlBody: emailHtml_(subject, html, buttons || []),
    name: C().OWNER_NAME,
    replyTo: ownerEmail_(),
  });
}

function emailHtml_(title, body, buttons) {
  const btns = buttons.map(function (b) {
    const bg = b.muted ? '#f1f1f4' : '#1f5eff';
    const fg = b.muted ? '#333' : '#fff';
    return '<a href="' + esc_(b.url) + '" style="display:inline-block;margin:6px 8px 0 0;padding:11px 18px;border-radius:8px;' +
      'background:' + bg + ';color:' + fg + ';text-decoration:none;font-weight:600">' + esc_(b.label) + '</a>';
  }).join('');
  return '<div style="font-family:Segoe UI,Helvetica,Arial,sans-serif;font-size:15px;line-height:1.5;color:#1d1d24;max-width:560px">' +
    '<h2 style="font-size:19px;margin:0 0 14px">' + esc_(title) + '</h2>' + body +
    (btns ? '<div style="margin-top:20px">' + btns + '</div>' : '') + '</div>';
}

/* ============================ Helpers ============================ */

function checkRequest_(id, token) {
  const req = getRequest_(id);
  if (!req || req.token !== token) throw new Error('Request not found.');
  if (req.status !== 'pending') throw new Error('This request is already ' + req.status + '.');
  return req;
}

function removeGoogleHold_(req) {
  try { Calendar.Events.remove(calId_(req), req.googleEventId); } catch (err) { /* already gone */ }
}

function meetLink_(ev) {
  if (ev.hangoutLink) return ev.hangoutLink;
  const eps = (ev.conferenceData && ev.conferenceData.entryPoints) || [];
  const v = eps.filter(function (x) { return x.entryPointType === 'video'; })[0];
  return v ? v.uri : '';
}

function props_() { return PropertiesService.getScriptProperties(); }
function saveRequest_(r) { props_().setProperty('req_' + r.id, JSON.stringify(r)); }
function getRequest_(id) {
  if (!id || !/^[a-f0-9]{12}$/.test(id)) return null;
  const s = props_().getProperty('req_' + id);
  return s ? JSON.parse(s) : null;
}
function allRequests_() {
  const all = props_().getProperties();
  return Object.keys(all).filter(function (k) { return k.indexOf('req_') === 0; })
    .map(function (k) { return JSON.parse(all[k]); });
}

function ownerEmail_() { return C().OWNER_EMAIL || Session.getEffectiveUser().getEmail(); }
function webAppUrl_() { return props_().getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl(); }
function reviewUrl_(id, token, action) {
  return webAppUrl_() + '?action=review&id=' + id + '&token=' + encodeURIComponent(token) + (action ? '&do=' + action : '');
}

/** Date for a wall-clock time in C().TIMEZONE (month 1-12, day may overflow). */
function zoned_(y, mo, d, h, mi) {
  const guess = Date.UTC(y, mo - 1, d, h, mi);
  let t = guess - tzOffsetMs_(new Date(guess));
  t = guess - tzOffsetMs_(new Date(t));
  return new Date(t);
}
function tzOffsetMs_(date) {
  const z = Utilities.formatDate(date, C().TIMEZONE, 'Z'); // e.g. +0200
  const sign = z.charAt(0) === '-' ? -1 : 1;
  return sign * (parseInt(z.substr(1, 2), 10) * 60 + parseInt(z.substr(3, 2), 10)) * 60000;
}

function formatRange_(startIso, endIso, tz) {
  const s = new Date(startIso), e = new Date(endIso);
  return Utilities.formatDate(s, tz, 'EEEE, d MMMM yyyy, HH:mm') + ' - ' + Utilities.formatDate(e, tz, 'HH:mm') + ' (' + tz + ')';
}
function guestWhen_(req) { return formatRange_(req.start, req.end, req.guestTz || C().TIMEZONE); }
function shortWhen_(req) {
  return Utilities.formatDate(new Date(req.start), req.guestTz || C().TIMEZONE, 'EEE d MMM, HH:mm');
}

function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function safeJson_(o) { return JSON.stringify(o).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026'); }

function simplePage_(title, msg) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;max-width:520px;margin:60px auto;padding:0 16px">' +
    '<h2>' + esc_(title) + '</h2><p>' + esc_(msg) + '</p></div>')
    .setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
