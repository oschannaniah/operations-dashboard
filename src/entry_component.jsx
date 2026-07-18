import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  LayoutGrid, ListChecks, CalendarDays, Users, DollarSign,
  ChevronLeft, ChevronRight, Plus, Circle, CheckCircle2, AlertTriangle,
  Link2, Building2, ShieldCheck, X, StickyNote, Activity as ActivityIcon,
  MessageSquare, Mic, Image as ImageIcon, Trash2, Clock, Bell, Camera, FileText
} from "lucide-react";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, HeadingLevel, TextRun, WidthType } from "docx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { createClient } from "@supabase/supabase-js";

const CENTRAL_TEAM = ["Hannaniah Owens", "Kassi Bourgeois", "Natalie Benton", "Lauren Bostic"];

const TIERS = [
  { tier: 1, role: "Central Operations Director", scope: "All campuses, full edit" },
  { tier: 2, role: "Campus Operations Director", scope: "Full edit, own campus Ops branch" },
  { tier: 3, role: "Campus Support Dir / Campus Admin Dir", scope: "Own sub-branch only (Phase 3-4 campuses)" },
  { tier: 4, role: "Coordinator (Facilities, Safety, Events, Admin)", scope: "Assigned items only" },
  { tier: 5, role: "Support role (Reception, Info Center, Hospitality, Cafe)", scope: "Task checkbox only" },
];

const STAGES = ["Pending", "Started", "In Progress", "Stalled", "Completed"];

const DEFAULT_ROLE_OPTIONS = [
  "Campus Operations Director", "Campus Support Director", "Campus Admin Director",
  "Facilities Coordinator", "Safety Coordinator", "Events Coordinator", "Admin Coordinator",
  "Production Coordinator", "Worship Director", "Reception", "Info Center", "Hospitality", "Cafe",
];

// The four team lanes a phase-4 (fully staffed) campus org chart is organized into, each with
// its own role/title list — sourced directly from the reference org chart. "Add Team Role"
// builds a person straight into one of these, as a second, more structured way to build the
// roster alongside the simpler free-text "Add Team Member" flow.
const TEAM_LANES = ["Operations", "Ministry", "Next Gen", "Programming"];

const LANE_ROLE_OPTIONS = {
  "Operations": [
    "Operations Director", "Campus Support Director", "Facilities Coordinator", "Safety Coordinator",
    "Events Coordinator", "Campus Admin Director", "Admin Coordinator", "Reception", "Info Center",
    "Hospitality", "Cafe",
  ],
  "Ministry": [
    "Ministry Director", "Connections Director", "Outreach Coordinator", "Groups Coordinator",
    "Next Steps Coordinator", "Life Care Director", "Prayer Coordinator", "Inner Healing Coordinator",
    "Pastoral Care Coordinator", "Guest Services Director", "Greeters Coordinator", "Parking Coordinator",
    "Usher Coordinator",
  ],
  "Next Gen": [
    "Next Gen Director", "Kids Director", "Elementary Coordinator", "Nursery Coordinator",
    "MDO Coordinator", "Young Adults Director", "Student's Director", "HS Coordinator", "MS Coordinator",
  ],
  "Programming": [
    "Programming Director", "Worship Director", "Production Director", "Production Coordinator",
    "Media Coordinator",
  ],
};

// Computed once at page load (module scope, not a hook) — accurate for the life of the page,
// and correct again on every fresh load/refresh. Was previously hardcoded to a fixed past
// date, which quietly threw off overdue-project detection, recurring task/project
// regeneration, and every activity/notification timestamp along with the Calendar tab.
const _now = new Date();
const TODAY_STR = `${_now.getFullYear()}-${String(_now.getMonth() + 1).padStart(2, "0")}-${String(_now.getDate()).padStart(2, "0")}`;

// ---------- Backend: Supabase (Postgres + Auth + Row Level Security + Edge Functions) ----------
// Replaces the old Google Sheets/Apps Script backend entirely. The URL and anon key are both
// meant to be public (same trust level as the Google Client ID below) — every actual security
// boundary is enforced server-side by Postgres Row Level Security policies, not by keeping
// these secret. See supabase/migrations/ in this repo for the schema/policies, and
// supabase/functions/ for the three Edge Functions this app calls.
const SUPABASE_URL = "https://pcuadpgamkoaytksbkcl.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBjdWFkcGdhbWtvYXl0a3Nia2NsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQzMTc1NzcsImV4cCI6MjA5OTg5MzU3N30.V5lMQXyX2N8Cl66KayZZOSSatn0-VR_BgjRs7kHtE4A";
// OSC's row in the "organizations" table — seeded once in migrations/0003_seed.sql. Hardcoded
// here only because this deployment is (for now) OSC-only; a real multi-tenant signup flow
// would resolve this from an invite/org code instead of a constant.
const OSC_ORG_ID = "00000000-0000-0000-0000-000000000001";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Set once auth state is wired up in OpsDashboard — calling this signs the user all the way
// out (e.g. if a request ever comes back with an expired/invalid session).
let handleUnauthorized = () => {};

// The rest of this app (every apiGet/apiCreate/apiUpdate/apiDelete call site, built up over
// the Sheets-backed version of this app) speaks the old "sheet name" + camelCase-field
// vocabulary. Rather than rewrite every one of those call sites, these two maps translate at
// the boundary: sheet name -> real Postgres table name, and camelCase <-> the table's actual
// snake_case columns — so every existing call site keeps working unchanged.
const TABLE_MAP = { Projects: "projects", Subtasks: "subtasks", Staff: "staff", Users: "profiles", CampusConfig: "campus_config", Notifications: "notifications", Campuses: "campuses", CentralThreads: "central_threads" };

// Every table's primary key is "id" except campus_config, whose natural key (campus_id) is
// also its actual Postgres primary key column — there's no separate surrogate "id" there.
const PK_MAP = { campus_config: "campus_id" };

const FIELD_MAPS = {
  projects: { organizationId: "organization_id", createdBy: "created_by", completedOn: "completed_on", sharedWith: "shared_with", dueTime: "due_time" },
  subtasks: { projectId: "project_id", createdBy: "created_by", dueTime: "due_time" },
  staff: { organizationId: "organization_id", campusId: "campus_id", reportsTo: "reports_to", nextMeeting: "next_meeting", lastContact: "last_contact", calendarSynced: "calendar_synced", userId: "user_id" },
  profiles: { organizationId: "organization_id", firstName: "first_name", lastName: "last_name", campusId: "campus_id", googleCalendarIds: "google_calendar_ids", googleCalendarNames: "google_calendar_names" },
  campus_config: { campusId: "campus_id", slidesLink: "slides_link" },
  notifications: { organizationId: "organization_id", forUser: "for_user_name", projectId: "project_id" },
  campuses: { organizationId: "organization_id", lead: "lead_name", leadProfileId: "lead_profile_id" },
  central_threads: { organizationId: "organization_id", threadKey: "thread_key" },
};

function toSnakeRow(table, obj) {
  const map = FIELD_MAPS[table] || {};
  const out = {};
  for (const [k, v] of Object.entries(obj)) out[map[k] || k] = v;
  return out;
}

function toCamelRow(table, row) {
  const map = FIELD_MAPS[table] || {};
  const reverse = Object.fromEntries(Object.entries(map).map(([k, v]) => [v, k]));
  const out = {};
  for (const [k, v] of Object.entries(row)) out[reverse[k] || k] = v;
  return out;
}

function throwOrHandle(error) {
  if (!error) return;
  if (error.code === "PGRST301" || /JWT/i.test(error.message || "")) handleUnauthorized();
  throw new Error(error.message || String(error));
}

async function apiGet(sheet, params = {}) {
  const table = TABLE_MAP[sheet] || sheet;
  let query = supabase.from(table).select("*");
  const map = FIELD_MAPS[table] || {};
  Object.entries(params).forEach(([key, value]) => { query = query.eq(map[key] || key, value); });
  const { data, error } = await query;
  throwOrHandle(error);
  return (data || []).map((row) => toCamelRow(table, row));
}

async function apiCreate(sheet, data) {
  const table = TABLE_MAP[sheet] || sheet;
  const { data: inserted, error } = await supabase.from(table).insert(toSnakeRow(table, data)).select().single();
  throwOrHandle(error);
  return toCamelRow(table, inserted);
}

async function apiUpdate(sheet, id, data) {
  const table = TABLE_MAP[sheet] || sheet;
  const { error } = await supabase.from(table).update(toSnakeRow(table, data)).eq(PK_MAP[table] || "id", id);
  throwOrHandle(error);
  return { ok: true };
}

async function apiDelete(sheet, id) {
  const table = TABLE_MAP[sheet] || sheet;
  const { error } = await supabase.from(table).delete().eq(PK_MAP[table] || "id", id);
  throwOrHandle(error);
  return { ok: true };
}

// The three Edge Functions (supabase/functions/) — anything requiring the service role key
// (admin account creation, the permanent Google connection, reading Slides with it) runs
// server-side here instead of directly against a table.
async function invokeFunction(name, action, data) {
  const { data: result, error } = await supabase.functions.invoke(name, { body: { action, data } });
  if (error) throw new Error(error.message || String(error));
  if (result && result.error) throw new Error(result.error);
  return result;
}
const apiAdminAccounts = (action, data) => invokeFunction("admin-accounts", action, data);
const apiConnectGoogleCalendar = (code) => invokeFunction("google-calendar", "connect", { code });
const apiListGoogleCalendars = () => invokeFunction("google-calendar", "listCalendars", {});
const apiListGoogleEvents = (calendarIds) => invokeFunction("google-calendar", "listEvents", { calendarIds });
const apiDisconnectGoogleCalendar = () => invokeFunction("google-calendar", "disconnect", {});
// Reads the campus's current-phase slide out of its Google Slides deck and returns
// { people: [{ name, role, reportsTo }] } — see importOrgChartFromSlides_ in Code.gs.
const apiImportOrgChartFromSlides = (campusId, url, phase) => invokeFunction("import-org-chart", null, { campusId, url, phase });


// Prevents the page behind a modal from scrolling. Kept deliberately simple — just
// locking overflow on html/body — because heavier techniques (like pinning body with
// position:fixed) can themselves create a new containing block for descendant fixed
// elements in some WebKit versions, which risks breaking the modal's own positioning.
// Paired with the outer-scrolls-not-inner modal pattern below, this is what actually
// needs to hold for the fix to be reliable on real iOS Safari/Chrome.
function useLockBodyScroll() {
  useEffect(() => {
    const prevBody = document.body.style.overflow;
    const prevHtml = document.documentElement.style.overflow;
    document.body.style.overflow = "hidden";
    document.documentElement.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = prevBody;
      document.documentElement.style.overflow = prevHtml;
    };
  }, []);
}

// Encodes the app's current "location" — which role is viewing, which tab, which campus,
// and any open modal ("window": a project, the new-project form, or a central callout detail)
// — into the URL's query string, and reads it back out. This is what makes the browser's
// Back/Forward buttons recall every page and every modal instead of leaving the app or doing
// nothing. Implemented with plain query params rather than a router library so the exact same
// file works both as a Claude artifact preview and as the deployed static site.
function readNavFromUrl() {
  if (typeof window === "undefined") return { role: "central", tab: "overview", campus: null, project: null, isNew: false, detail: null };
  const params = new URLSearchParams(window.location.search);
  return {
    role: params.get("role") || "central",
    tab: params.get("tab") || "overview",
    campus: params.get("campus") || null,
    project: params.get("project") ? Number(params.get("project")) : null,
    isNew: params.get("new") === "1",
    detail: params.get("detail") || null,
  };
}

function buildUrlFromNav(nav) {
  const params = new URLSearchParams();
  if (nav.role && nav.role !== "central") params.set("role", nav.role);
  if (nav.tab && nav.tab !== "overview") params.set("tab", nav.tab);
  if (nav.campus) params.set("campus", nav.campus);
  if (nav.project) params.set("project", String(nav.project));
  if (nav.isNew) params.set("new", "1");
  if (nav.detail) params.set("detail", nav.detail);
  const qs = params.toString();
  const path = window.location.pathname;
  return qs ? `${path}?${qs}` : path;
}

const RECUR_FREQ = [
  { id: "none", label: "Does not repeat" },
  { id: "daily", label: "Daily" },
  { id: "weekly", label: "Weekly" },
  { id: "biweekly", label: "Bi-weekly" },
  { id: "monthly", label: "Monthly" },
  { id: "custom", label: "Custom" },
];

function describeRecurrence(r) {
  if (!r || r.freq === "none") return null;
  const freqLabel = r.freq === "custom" ? `every ${r.customEvery || 1} ${r.customUnit || "weeks"}` : RECUR_FREQ.find((f) => f.id === r.freq)?.label.toLowerCase();
  const end = r.endType === "count" ? ` · ${r.count || 1}x left` : r.endType === "date" ? ` · until ${r.untilDate}` : "";
  return `${freqLabel}${end}`;
}

function addInterval(dateStr, r) {
  const base = dateStr || TODAY_STR;
  const d = new Date(base + "T00:00:00");
  if (r.freq === "daily") d.setDate(d.getDate() + 1);
  else if (r.freq === "weekly") d.setDate(d.getDate() + 7);
  else if (r.freq === "biweekly") d.setDate(d.getDate() + 14);
  else if (r.freq === "monthly") d.setMonth(d.getMonth() + 1);
  else if (r.freq === "custom") {
    const n = Number(r.customEvery) || 1;
    if (r.customUnit === "days") d.setDate(d.getDate() + n);
    else if (r.customUnit === "weeks") d.setDate(d.getDate() + 7 * n);
    else if (r.customUnit === "months") d.setMonth(d.getMonth() + n);
  }
  return d.toISOString().slice(0, 10);
}

// Recurrence is schedule-anchored, not completion-anchored: a task set to recur every
// Monday only regenerates once the calendar actually reaches the next Monday — completing
// it early (e.g. the Wednesday before) does not spawn or reset anything ahead of schedule.
// This is a pure read-time resolver: it never mutates stored state on its own.
function resolveSubtask(s) {
  if (!s.recurrence || s.recurrence.freq === "none" || !s.done) return s;
  const r = s.recurrence;
  const nextDue = addInterval(s.due, r);
  if (TODAY_STR < nextDue) return s; // scheduled date hasn't arrived yet — stays completed as-is
  let carryCount = r.count;
  if (r.endType === "count") {
    carryCount = (Number(r.count) || 1) - 1;
    if (carryCount <= 0) return s; // recurrence has run its course
  }
  if (r.endType === "date" && r.untilDate && nextDue > r.untilDate) return s;
  return { ...s, due: nextDue, done: false, recurrence: { ...r, count: carryCount } };
}

// Same schedule-anchored regeneration, but for a whole recurring project: only once the
// next scheduled date arrives does a Done project reopen for its next cycle, with every
// sub-task reset for the new round. Sub-task-level recurrence is resolved independently
// on top, so a project and its individual tasks can each run their own schedule.
function resolveProject(p) {
  const subtasks = (p.subtasks || []).map(resolveSubtask);
  let result = { ...p, subtasks };
  if (p.recurrence && p.recurrence.freq !== "none" && p.stage === "Completed") {
    const r = p.recurrence;
    const nextDue = addInterval(p.due, r);
    if (TODAY_STR >= nextDue) {
      let carryCount = r.count;
      let exhausted = false;
      if (r.endType === "count") {
        carryCount = (Number(r.count) || 1) - 1;
        if (carryCount <= 0) exhausted = true;
      }
      if (r.endType === "date" && r.untilDate && nextDue > r.untilDate) exhausted = true;
      if (!exhausted) {
        result = { ...result, due: nextDue, stage: "Pending", recurrence: { ...r, count: carryCount }, subtasks: subtasks.map((s) => ({ ...s, done: false })) };
      }
    }
  }
  return result;
}

const seedProjects = [
  {
    id: 1, campus: "abv", title: "Develop Jack Thomas — Worship Director track", stage: "In Progress", createdAt: "2026-06-10", completedOn: null,
    owner: "Angel Lormand", createdBy: "Angel Lormand", team: ["Jack Thomas"], collaborators: ["Worship Team"], due: "2026-08-15",
    cost: 1200, spent: 450, shared: false,
    subtasks: [
      { id: 101, t: "Shadow Life.Church production model", done: true, cost: 0, spent: 0, createdBy: "Angel Lormand", photos: [], notes: [] },
      { id: 102, t: "First led rehearsal", done: true, cost: 0, spent: 0, createdBy: "Angel Lormand", photos: [], notes: [] },
      { id: 103, t: "First led Sunday", done: false, cost: 0, spent: 0, createdBy: "Angel Lormand", photos: [], notes: [] },
    ],
    photos: [],
    notes: [{ ts: "2026-07-08 09:12", author: "Katie", text: "Jack shadowed the 9am service — strong instinct for transitions." }],
  },
  {
    id: 2, campus: "abv", title: "Worship Night production plan", stage: "In Progress", createdAt: "2026-06-20", completedOn: null,
    owner: "Angel Lormand", createdBy: "Angel Lormand", team: ["Production Coor"], collaborators: ["Worship Team"], due: "2026-07-22",
    cost: 800, spent: 300, shared: false,
    subtasks: [
      { id: 201, t: "Run-of-show drafted", done: true, cost: 0, spent: 0, createdBy: "Angel Lormand", photos: [], notes: [] },
      { id: 202, t: "Volunteer staffing confirmed", done: false, cost: 0, spent: 0, createdBy: "Angel Lormand", photos: [], notes: [] },
      { id: 203, t: "Sound check", done: false, cost: 0, spent: 0, createdBy: "Angel Lormand", photos: [], notes: [] },
      { id: 204, t: "Production team huddle", done: false, cost: 0, spent: 0, due: "2026-07-16", createdBy: "Angel Lormand", photos: [], notes: [], recurrence: { freq: "weekly", endType: "date", untilDate: "2026-08-30" } },
    ],
    photos: [],
    notes: [{ ts: "2026-07-09 14:40", author: "Katie", text: "Confirmed 6 volunteers, still need 2 for parking." }],
  },
  {
    id: 3, campus: "laf", title: "Guest connect flow — NFC to CCB", stage: "Stalled", createdAt: "2026-06-01", completedOn: null,
    owner: "Katie Wilbanks", createdBy: "Katie Wilbanks", team: ["Admin Coor"], collaborators: ["Connections Team"], due: "2026-07-30",
    cost: 2400, spent: 1900, shared: true, sharedWith: ["opl", "mid"],
    subtasks: [
      { id: 301, t: "NFC cards ordered", done: true, cost: 0, spent: 0, createdBy: "Katie Wilbanks", photos: [], notes: [] },
      { id: 302, t: "Cognito Forms field mapping", done: false, cost: 0, spent: 0, createdBy: "Katie Wilbanks", photos: [], notes: [] },
      { id: 303, t: "Pushpay API write-access confirmed", done: false, cost: 0, spent: 0, createdBy: "Katie Wilbanks", photos: [], notes: [] },
    ],
    photos: [],
    notes: [{ ts: "2026-07-07 11:02", author: "Shaina", text: "Blocked on Pushpay confirming API tier access." }],
  },
  {
    id: 4, campus: "opl", title: "OD transition planning", stage: "Pending", owner: "Shaina Broussard", createdBy: "Shaina Broussard", createdAt: "2026-07-01", completedOn: null, team: [], collaborators: [],
    due: "2026-08-01", cost: 0, spent: 0, shared: false,
    subtasks: [{ id: 401, t: "Draft comms", done: false, cost: 0, spent: 0, createdBy: "Shaina Broussard", photos: [], notes: [] }],
    photos: [], notes: [],
  },
  {
    id: 5, campus: "ynv", title: "Safety Team SOP drills — Q3", stage: "Pending", owner: "Pastor Josh Mesa", createdBy: "Pastor Josh Mesa", createdAt: "2026-07-05", completedOn: null, team: ["Safety Coor"],
    collaborators: [], due: "2026-08-10", cost: 300, spent: 0, shared: true, sharedWith: ["nib"],
    subtasks: [{ id: 501, t: "Schedule drill weekend", done: false, cost: 0, spent: 0, createdBy: "Pastor Josh Mesa", photos: [], notes: [] }],
    photos: [], notes: [],
  },
  {
    id: 6, campus: "central", title: "Church Health Scorecard — Phase 2 rollout", stage: "In Progress", section: "Central Initiatives", createdAt: "2026-05-15", completedOn: null,
    owner: "Lauren / Hannaniah", createdBy: "Hannaniah Owens", team: [], collaborators: ["Directional Team"], due: "2026-09-01",
    cost: 5000, spent: 2100, shared: true, sharedWith: ["all"],
    subtasks: [
      { id: 601, t: "Weekly dashboard live", done: true, cost: 0, spent: 0, createdBy: "Hannaniah Owens", photos: [], notes: [] },
      { id: 602, t: "Quarterly report template", done: false, cost: 0, spent: 0, createdBy: "Hannaniah Owens", photos: [], notes: [] },
      { id: 603, t: "Finance AI agent prototype", done: false, cost: 0, spent: 0, createdBy: "Hannaniah Owens", photos: [], notes: [] },
    ],
    photos: [],
    notes: [{ ts: "2026-07-05 08:30", author: "Hannaniah", text: "Strong reception at Directional — moving to Phase 2 build." }],
  },
  {
    id: 7, campus: "central", title: "Bollard & wayfinding signage — OpC", stage: "In Progress", owner: "Shaina Broussard", createdBy: "Shaina Broussard", section: "Central Initiatives", createdAt: "2026-06-25", completedOn: null,
    team: ["Facilities Coor"], collaborators: [], due: "2026-08-05", cost: 3200, spent: 3200, shared: false,
    subtasks: [
      { id: 701, t: "Marker schedule M1-M7 finalized", done: true, cost: 0, spent: 0, createdBy: "Shaina Broussard", photos: [], notes: [] },
      { id: 702, t: "Install", done: false, cost: 0, spent: 0, createdBy: "Shaina Broussard", photos: [], notes: [] },
    ],
    photos: [], notes: [],
  },
];

const MOCK_GOOGLE_CALENDARS = [
  "Primary — work email",
  "Campus Events",
  "Production & Rehearsal Schedule",
  "Safety Team Drills",
  "Facilities Maintenance",
  "Personal",
];

// ---------- Real per-user Google Calendar connection (read-only, permanent) ----------
// Each person connects their OWN Google Workspace calendars to their OWN Calendar tab. This is
// a genuine, permanent connection, not a per-session one: we use Google Identity Services' OAuth
// CODE flow (not the simpler token flow) to get a one-time authorization code, which the backend
// (Code.gs) exchanges for an access + refresh token and stores the refresh token server-side —
// never in the browser. From then on, the browser just asks the backend for calendars/events;
// Google is never called directly from here again, and no re-consent is ever needed unless the
// person revokes access themselves at myaccount.google.com/permissions.
//
// Requires a real OAuth Client ID from Google Cloud Console (APIs & Services > Credentials >
// Create OAuth client ID > Web application), with the Google Calendar API enabled and this
// site's URL(s) added under "Authorized JavaScript origins". Paste it in below — the matching
// Client Secret goes only in Code.gs's Script Properties, never here.
const GOOGLE_CALENDAR_CLIENT_ID = "554435571384-jj0t9lfk8gh3kenu327phiam0c6d39b1.apps.googleusercontent.com";
const GOOGLE_CALENDAR_SCOPE = "https://www.googleapis.com/auth/calendar.readonly";

let googleCodeClient = null;
let googleCodeCallback = () => {};

function requestGoogleAuthCode(callback) {
  if (typeof window === "undefined" || !window.google?.accounts?.oauth2) {
    callback({ error: "Google sign-in isn't available right now — the page may still be loading, try again in a moment." });
    return;
  }
  if (!googleCodeClient) {
    googleCodeClient = window.google.accounts.oauth2.initCodeClient({
      client_id: GOOGLE_CALENDAR_CLIENT_ID,
      scope: GOOGLE_CALENDAR_SCOPE,
      ux_mode: "popup",
      access_type: "offline", // ask for a refresh token, not just a short-lived access token
      callback: (resp) => googleCodeCallback(resp),
    });
  }
  googleCodeCallback = callback;
  googleCodeClient.requestCode();
}

// Merges raw Google event objects (from possibly several calendars) into the same
// { "YYYY-MM-DD": [{ time, title, attendees }] } shape the rest of CalendarPanel already reads.
function groupGoogleEventsByDate(events) {
  const byDate = {};
  events.forEach((ev) => {
    const startRaw = ev.start?.dateTime || ev.start?.date;
    if (!startRaw) return;
    const dateStr = startRaw.slice(0, 10);
    const timeLabel = ev.start?.dateTime
      ? new Date(ev.start.dateTime).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
      : "All day";
    if (!byDate[dateStr]) byDate[dateStr] = [];
    byDate[dateStr].push({ time: timeLabel, title: ev.summary || "(no title)", attendees: (ev.attendees || []).map((a) => a.displayName || a.email).filter(Boolean) });
  });
  return byDate;
}

const seedStaff = {
  abv: [
    { id: 1, name: "Jack Thomas", roles: ["Worship Director"], nextMeeting: "2026-07-15 · 1:1 w/ Angel Lormand", lastContact: "Yesterday", flag: null, calendarSynced: true, calendars: ["Primary — work email", "Production & Rehearsal Schedule"] },
    { id: 2, name: "Production Coor", roles: ["Production Coordinator"], nextMeeting: "2026-07-22 · Worship Night", lastContact: "3 days ago", flag: "Missed check-in due 7/8", calendarSynced: false, calendars: [] },
    { id: 3, name: "Safety Coor", roles: ["Safety Coordinator"], nextMeeting: "No meeting scheduled", lastContact: "2 weeks ago", flag: "Overdue contact", calendarSynced: false, calendars: [] },
  ],
  opl: [{ id: 4, name: "Admin Coor", roles: ["Admin Coordinator"], nextMeeting: "2026-07-16 · Weekly sync", lastContact: "Today", flag: null, calendarSynced: true, calendars: ["Primary — work email"] }],
  laf: [
    { id: 5, name: "Admin Coor", roles: ["Admin Coordinator", "Reception"], nextMeeting: "2026-07-14 · NFC rollout", lastContact: "Today", flag: null, calendarSynced: true, calendars: ["Primary — work email", "Campus Events"] },
    { id: 6, name: "Facilities Coor", roles: ["Facilities Coordinator"], nextMeeting: "2026-07-18 · Walkthrough", lastContact: "5 days ago", flag: null, calendarSynced: false, calendars: [] },
  ],
};

const seedNotesInit = {
  abv: [
    { id: 1, text: "Ask Pastor Don about Worship Night parking overflow plan.", ts: "2026-07-09 08:14" },
    { id: 2, text: "Follow up: background check renewal for 2 volunteers", ts: "2026-07-08 16:02" },
  ],
  laf: [{ id: 3, text: "Cognito Forms pricing tier — confirm before Friday", ts: "2026-07-09 10:20" }],
};

const seedActivity = [
  { id: 1, ts: "2 min ago", actor: "Angel Lormand", action: "completed subtask", target: "\"First led rehearsal\" — Jack Thomas development" },
  { id: 2, ts: "18 min ago", actor: "Katie Wilbanks", action: "added a note to", target: "Guest connect flow — NFC to CCB" },
  { id: 3, ts: "1 hr ago", actor: "Shaina Broussard", action: "uploaded a photo to", target: "Bollard & wayfinding signage — OpC" },
  { id: 4, ts: "3 hrs ago", actor: "Lauren Smith", action: "created task", target: "Schedule Q3 safety drill" },
  { id: 5, ts: "Yesterday", actor: "AI Meeting Import", action: "synced transcript into", target: "Church Health Scorecard — Phase 2 rollout" },
  { id: 6, ts: "Yesterday", actor: "Jared Robicheaux", action: "marked overdue", target: "Submit budget variance note — New Iberia" },
];

const seedCalendar = {
  "2026-07-10": [
    { time: "9:00 AM", title: "Staff Huddle", attendees: ["Angel Lormand", "Production Coor", "Safety Coor"] },
    { time: "11:00 AM", title: "1:1 — Jack Thomas", attendees: ["Angel Lormand", "Jack Thomas"] },
    { time: "2:00 PM", title: "Vendor call: Stage lighting", attendees: ["Angel Lormand"] },
  ],
  "2026-07-11": [{ time: "10:00 AM", title: "Central Ops / OD Touchpoint", attendees: ["Hannaniah Owens", "Angel Lormand", "Shaina Broussard", "Katie Wilbanks"] }],
  "2026-07-13": [{ time: "9:00 AM", title: "Sunday Services", attendees: ["Angel Lormand", "Jack Thomas", "Production Coor"] }],
  "2026-07-16": [{ time: "9:00 AM", title: "Production team huddle", attendees: ["Angel Lormand", "Production Coor"] }],
  "2026-07-17": [{ time: "10:00 AM", title: "Central Ops / OD Touchpoint", attendees: ["Hannaniah Owens", "Angel Lormand", "Shaina Broussard", "Katie Wilbanks"] }],
  "2026-07-22": [{ time: "6:00 PM", title: "Worship Night", attendees: ["Angel Lormand", "Jack Thomas", "Production Coor", "Worship Team"] }],
  "2026-08-01": [{ time: "9:00 AM", title: "OD transition planning", attendees: ["Shaina Broussard"] }],
  "2026-08-05": [{ time: "1:00 PM", title: "Bollard signage install", attendees: ["Shaina Broussard", "Facilities Coor"] }],
  "2026-08-09": [{ time: "10:00 AM", title: "Quarterly Men's Community Gathering", attendees: ["Katie Wilbanks", "Pastor Josh Mesa", "Lauren Smith"] }],
  "2026-08-10": [{ time: "8:00 AM", title: "Safety Drill Weekend", attendees: ["Pastor Josh Mesa", "Jared Robicheaux", "Safety Coor"] }],
};

const fmtMoney = (n) => `$${n.toLocaleString()}`;

// Rolls up the true project budget: the project's own estimate/spend plus every
// sub-task's estimate/spend underneath it. This is what all budget widgets should
// read from now, rather than a static per-campus number.
function projectBudget(p) {
  const subtaskCost = (p.subtasks || []).reduce((sum, s) => sum + (Number(s.cost) || 0), 0);
  const subtaskSpent = (p.subtasks || []).reduce((sum, s) => sum + (Number(s.spent) || 0), 0);
  return { total: (Number(p.cost) || 0) + subtaskCost, spent: (Number(p.spent) || 0) + subtaskSpent };
}

// Completed projects are excluded from every budget outlook — their spending is finalized and
// closed out, so they'd only clutter forward-looking "how are we tracking" totals. Their own
// final cost/spent/outcome is still fully preserved on the project record itself and shown
// wherever that specific project is viewed (see budgetOutcomeLabel) — this only affects sums.
function rollupBudget(projectsList) {
  return projectsList.filter((p) => p.stage !== "Completed").reduce(
    (acc, p) => {
      const b = projectBudget(p);
      acc.total += b.total;
      acc.spent += b.spent;
      return acc;
    },
    { total: 0, spent: 0 }
  );
}

// "under" / "at" / "over" budget, based on the full rollup (project + every sub-task).
function budgetStatus(project) {
  const b = projectBudget(project);
  if (b.total === 0) return "at";
  if (b.spent < b.total) return "under";
  if (b.spent === b.total) return "at";
  return "over";
}
const BUDGET_STATUS_LABEL = { under: "Under budget", at: "On budget", over: "Over budget" };
const BUDGET_STATUS_COLOR = { under: "#5E9E8A", at: "#B8862F", over: "#C15B5B" };

function tierForRole(roleId) { return roleId === "central" ? TIERS[0] : TIERS[1]; }

// Shown whenever there's no valid auth in state — either a returning user logging back in,
// or (only reachable while the Users sheet is completely empty — see handleBootstrap_ in
// Code.gs) the very first Central account being created for a fresh deploy.
function LoginScreen() {
  const [mode, setMode] = useState("login"); // "login" | "register"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [phone, setPhone] = useState("");
  const [error, setError] = useState("");
  const [checkEmail, setCheckEmail] = useState(false);
  const [loading, setLoading] = useState(false);

  const inputClass = "w-full bg-[#F7F6FB] border border-[#D8D5EC] rounded-md px-3 py-2 text-[13.5px] outline-none focus:border-[#2B4C7E]";
  const labelClass = "text-[11.5px] font-medium text-[#6B6980] mb-1 block";

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setCheckEmail(false); setLoading(true);
    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        // onAuthStateChange (wired up in OpsDashboard) picks up the new session from here.
      } else {
        if (password.length < 8) throw new Error("Password must be at least 8 characters.");
        const { data, error } = await supabase.auth.signUp({
          email,
          password,
          options: { data: { organization_id: OSC_ORG_ID, first_name: firstName, last_name: lastName, phone } },
        });
        if (error) throw error;
        // Email confirmation is required on this project — signUp() succeeds but returns no
        // session until the link in that email is clicked, so there's nothing for
        // onAuthStateChange to pick up yet. Without this, the form just looked like it did
        // nothing.
        if (!data.session) setCheckEmail(true);
      }
    } catch (err) {
      setError(err?.message || String(err));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen w-full flex items-center justify-center px-4" style={{ fontFamily: "'Inter', sans-serif", background: "#F7F6FB", color: "#2A2A3A" }}>
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fraunces:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" />
      <div className="w-full max-w-[380px]">
        <div className="flex items-center gap-3 mb-6 justify-center">
          <div className="w-9 h-9 rounded-md flex items-center justify-center shrink-0" style={{ background: "#2B4C7E" }}><Building2 size={18} color="#F7F6FB" /></div>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif" }} className="text-[19px] font-semibold leading-none tracking-tight">OpsCore</div>
            <div className="text-[11px] mt-0.5 text-[#6B6980]">Reaching People · Building Lives</div>
          </div>
        </div>

        <div className="bg-[#FFFFFF] rounded-xl p-6" style={{ border: "1px solid #E3E1F0" }}>
          <h1 className="text-[16px] font-semibold mb-1">{mode === "login" ? "Sign in" : "Create your account"}</h1>
          <p className="text-[12px] text-[#6B6980] mb-4">
            {mode === "login" ? "Use the email and password you signed up with." : "Anyone can create a login — your Central Operations Director will assign your campus and role afterward."}
          </p>

          <form onSubmit={submit} className="space-y-3">
            {mode === "register" && (
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className={labelClass}>First name</label>
                  <input required value={firstName} onChange={(e) => setFirstName(e.target.value)} className={inputClass} />
                </div>
                <div>
                  <label className={labelClass}>Last name</label>
                  <input required value={lastName} onChange={(e) => setLastName(e.target.value)} className={inputClass} />
                </div>
              </div>
            )}
            <div>
              <label className={labelClass}>Email</label>
              <input required type="email" value={email} onChange={(e) => setEmail(e.target.value)} className={inputClass} />
            </div>
            {mode === "register" && (
              <div>
                <label className={labelClass}>Phone (optional)</label>
                <input value={phone} onChange={(e) => setPhone(e.target.value)} className={inputClass} />
              </div>
            )}
            <div>
              <label className={labelClass}>Password</label>
              <input required type="password" value={password} onChange={(e) => setPassword(e.target.value)} className={inputClass} minLength={mode === "register" ? 8 : undefined} />
            </div>

            {error && <div className="text-[12px] rounded-md px-3 py-2" style={{ background: "#C15B5B1A", color: "#C15B5B" }}>{error}</div>}
            {checkEmail && <div className="text-[12px] rounded-md px-3 py-2" style={{ background: "#5E9E8A1A", color: "#5E9E8A" }}>Account created — check your email for a confirmation link, then come back and sign in.</div>}

            <button type="submit" disabled={loading}
              className="w-full text-[13.5px] font-medium rounded-md px-3 py-2.5 mt-1"
              style={{ background: "#2B4C7E", color: "#F7F6FB", opacity: loading ? 0.7 : 1 }}>
              {loading ? "Please wait…" : mode === "login" ? "Sign in" : "Create account & sign in"}
            </button>
          </form>

          <button onClick={() => { setMode(mode === "login" ? "register" : "login"); setError(""); }}
            className="w-full text-center text-[11.5px] mt-4" style={{ color: "#B8862F" }}>
            {mode === "login" ? "New here? Create an account" : "Already have an account? Sign in"}
          </button>
        </div>
      </div>
    </div>
  );
}

// Shown after a successful login/registration whose account has tier "unassigned" — a real
// login that simply hasn't been given a campus + role yet. Nothing to do here but wait; once
// Central assigns them (Team Accounts panel), the change takes effect on next sign-in, since
// tier/campusId are baked into the token issued at login time, not re-checked live.
function PendingAssignmentScreen({ user, onSignOut }) {
  return (
    <div className="min-h-screen w-full flex items-center justify-center px-4" style={{ fontFamily: "'Inter', sans-serif", background: "#F7F6FB", color: "#2A2A3A" }}>
      <div className="w-full max-w-[380px] text-center">
        <div className="flex items-center gap-3 mb-6 justify-center">
          <div className="w-9 h-9 rounded-md flex items-center justify-center shrink-0" style={{ background: "#2B4C7E" }}><Building2 size={18} color="#F7F6FB" /></div>
          <div className="text-left">
            <div style={{ fontFamily: "'Fraunces', serif" }} className="text-[19px] font-semibold leading-none tracking-tight">OpsCore</div>
            <div className="text-[11px] mt-0.5 text-[#6B6980]">Reaching People · Building Lives</div>
          </div>
        </div>
        <div className="bg-[#FFFFFF] rounded-xl p-6" style={{ border: "1px solid #E3E1F0" }}>
          <div className="w-10 h-10 rounded-full flex items-center justify-center mx-auto mb-3" style={{ background: "#B8862F22" }}>
            <Clock size={18} color="#B8862F" />
          </div>
          <h1 className="text-[15px] font-semibold mb-2">Almost there, {user.firstName}</h1>
          <p className="text-[12.5px] text-[#6B6980] leading-relaxed mb-4">
            Your account is created, but no one's assigned your campus and role yet. Once your Central Operations Director does that in Team Accounts, sign out and back in to pick up your access.
          </p>
          <button onClick={onSignOut} className="text-[12.5px] font-medium rounded-md px-4 py-2" style={{ background: "#E3E1F0", color: "#2A2A3A" }}>Sign out</button>
        </div>
      </div>
    </div>
  );
}

export default function OpsDashboard() {
  const [role, setRole] = useState(() => readNavFromUrl().role);
  const [tab, setTab] = useState(() => readNavFromUrl().tab);
  const [projects, setProjects] = useState(seedProjects);
  const [notesByCampus, setNotesByCampus] = useState(seedNotesInit);
  const [newNote, setNewNote] = useState("");
  const [selectedCampus, setSelectedCampus] = useState(() => readNavFromUrl().campus);
  const [openProject, setOpenProject] = useState(() => readNavFromUrl().project);
  const [showNewProject, setShowNewProject] = useState(() => readNavFromUrl().isNew);
  const [detail, setDetail] = useState(() => readNavFromUrl().detail); // central callout detail modal
  const [calDate, setCalDate] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  });
  const [calView, setCalView] = useState("day");
  const [staffByCampus, setStaffByCampus] = useState(seedStaff);
  const [campuses, setCampuses] = useState([]);
  const [users, setUsers] = useState([]);
  const [campusSlidesLinks, setCampusSlidesLinks] = useState({});
  const [orgChartByCampus, setOrgChartByCampus] = useState({});
  const [pendingReassignments, setPendingReassignments] = useState([]);
  const [roleOptions, setRoleOptions] = useState(DEFAULT_ROLE_OPTIONS);
  const [myCalendarSynced, setMyCalendarSynced] = useState(true);
  const [sections, setSections] = useState(["General", "Central Initiatives"]);
  const [centralThreads, setCentralThreads] = useState({});
  const [notifications, setNotifications] = useState([]);
  const [activityLog, setActivityLog] = useState([]);
  const [backendStatus, setBackendStatus] = useState("loading"); // "loading" | "connected" | "offline"
  const [backendError, setBackendError] = useState("");
  const [syncing, setSyncing] = useState(false);
  const [bootstrapProgress, setBootstrapProgress] = useState(null); // { done, total } while seeding a fresh Sheet
  const [showNotifications, setShowNotifications] = useState(false);
  const [openStaffProfile, setOpenStaffProfile] = useState(null);
  // Auth now rides on Supabase's own session (it persists/refreshes itself — no manual
  // localStorage handling needed), with the campus/role/tier fields the rest of this app
  // expects on auth.user pulled from that person's own "profiles" row.
  const [auth, setAuthState] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const suppressPush = useRef(false);

  const setAuth = (next) => setAuthState(next); // null = signed out
  handleUnauthorized = () => { supabase.auth.signOut(); setAuthState(null); };

  useEffect(() => {
    let cancelled = false;
    const loadForSession = async (session) => {
      if (!session) { if (!cancelled) { setAuthState(null); setAuthLoading(false); } return; }
      const { data: profileRow } = await supabase.from("profiles").select("*").eq("id", session.user.id).single();
      if (cancelled) return;
      setAuthState(profileRow ? { token: session.access_token, user: toCamelRow("profiles", profileRow) } : null);
      setAuthLoading(false);
    };
    supabase.auth.getSession().then(({ data }) => loadForSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => loadForSession(session));
    return () => { cancelled = true; sub.subscription.unsubscribe(); };
  }, []);

  // Restore state when the browser's Back/Forward buttons are used.
  useEffect(() => {
    const onPopState = () => {
      const nav = readNavFromUrl();
      suppressPush.current = true;
      setRole(nav.role);
      setTab(nav.tab);
      setSelectedCampus(nav.campus);
      setOpenProject(nav.project);
      setShowNewProject(nav.isNew);
      setDetail(nav.detail);
    };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  // Push a new history entry whenever the visible page or open modal changes, so Back/Forward
  // can step through every page and window that's been visited — unless this change came FROM
  // a Back/Forward press itself, in which case the URL already matches and we skip re-pushing.
  useEffect(() => {
    if (suppressPush.current) { suppressPush.current = false; return; }
    const nav = { role, tab, campus: selectedCampus, project: openProject, isNew: showNewProject, detail };
    const url = buildUrlFromNav(nav);
    if (url !== window.location.pathname + window.location.search) {
      window.history.pushState(nav, "", url);
    }
  }, [role, tab, selectedCampus, openProject, showNewProject, detail]);

  const activeCampusId = role === "central" ? selectedCampus : role;
  const activeCampus = campuses.find((c) => c.id === activeCampusId);
  const tier = tierForRole(role);
  const currentViewerName = auth?.user
    ? (`${auth.user.firstName || ""} ${auth.user.lastName || ""}`.trim() || auth.user.email)
    : "—";

  // A logged-in campus OD/staff account is permanently scoped to its own campus — the role
  // switcher in the header is Central-only. This just keeps local `role` state in sync with
  // that identity; the backend enforces the actual scoping independently either way.
  useEffect(() => {
    if (auth?.user && auth.user.tier !== "central" && auth.user.tier !== "unassigned" && role !== auth.user.campusId) {
      setRole(auth.user.campusId);
      setSelectedCampus(null);
    }
  }, [auth]);

  // Who's available to assign a project/task to depends on scope, not on who's currently
  // logged in: a project scoped to one campus can only pull from that campus's own roster,
  // even if Central is the one creating it. Only a project explicitly marked cross-campus
  // can reach into other campuses or the Central team.
  const campusRoster = useMemo(() => {
    if (!activeCampusId) {
      return CENTRAL_TEAM.map((name) => ({ name, roles: ["Central Team"] }));
    }
    const campus = campuses.find((c) => c.id === activeCampusId);
    const lead = campus ? [{ name: campus.lead, roles: ["Campus Operations Director"] }] : [];
    const campusStaff = (staffByCampus[activeCampusId] || []).map((s) => ({ name: s.name, roles: s.roles || [] }));
    return [...lead, ...campusStaff];
  }, [activeCampusId, staffByCampus]);

  const fullRoster = useMemo(() => {
    const central = CENTRAL_TEAM.map((name) => ({ name, roles: ["Central Team"] }));
    const leads = campuses.map((c) => ({ name: c.lead, roles: ["Campus Operations Director"] }));
    const allStaff = Object.values(staffByCampus).flat().map((s) => ({ name: s.name, roles: s.roles || [] }));
    const combined = [...central, ...leads, ...allStaff];
    const seen = new Set();
    return combined.filter((p) => (seen.has(p.name) ? false : (seen.add(p.name), true)));
  }, [staffByCampus]);

  const roster = fullRoster; // used where scope doesn't matter (e.g. Reports)
  const myNotifications = notifications.filter((n) => n.forUser === currentViewerName);

  // Read-time schedule resolution: reflects any recurring project/sub-task that has reached
  // its next scheduled date, without ever mutating the underlying stored state on its own.
  // Fire-and-forget backend sync: local state updates instantly (optimistic), the Sheet
  // catches up in the background. If it fails, the banner shows it — nothing is lost locally,
  // but it won't be saved until the connection is healthy again.
  const syncToBackend = (promise) => {
    setSyncing(true);
    promise
      .then(() => { setBackendStatus("connected"); setBackendError(""); })
      .catch((err) => { setBackendStatus("offline"); setBackendError(err?.message || String(err)); })
      .finally(() => setSyncing(false));
  };

  const bootstrapBackend = async (onProgress) => {
    let done = 0;
    const failures = [];
    const total = seedProjects.length + seedProjects.reduce((n, p) => n + p.subtasks.length, 0) + Object.values(seedStaff).reduce((n, arr) => n + arr.length, 0);
    const tryCreate = async (sheet, data, label) => {
      try {
        await apiCreate(sheet, data);
      } catch (err) {
        failures.push(`${label}: ${err?.message || err}`);
      }
      done++; onProgress && onProgress(done, total);
    };
    for (const p of seedProjects) {
      const { subtasks, ...projectRow } = p;
      await tryCreate("Projects", projectRow, p.title);
      for (const s of subtasks) await tryCreate("Subtasks", { ...s, projectId: p.id }, s.t);
    }
    for (const [campusId, people] of Object.entries(seedStaff)) {
      for (const person of people) await tryCreate("Staff", { ...person, campusId }, person.name);
    }
    if (failures.length > 0) throw new Error(`${failures.length} row(s) failed to create. First: ${failures[0]}`);
  };

  const clearAllProjects = async () => {
    setSyncing(true);
    try {
      const [projectRows, subtaskRows] = await Promise.all([apiGet("Projects"), apiGet("Subtasks")]);
      for (const p of projectRows) await apiDelete("Projects", p.id);
      for (const s of subtaskRows) await apiDelete("Subtasks", s.id);
      setProjects([]);
      setBackendStatus("connected"); setBackendError("");
    } catch (err) {
      setBackendStatus("offline"); setBackendError(err?.message || String(err));
    } finally {
      setSyncing(false);
    }
  };

  useEffect(() => {
    if (!auth) return;
    let cancelled = false;
    (async () => {
      // Each sheet is fetched independently (allSettled, not all) — one sheet erroring (a
      // missing tab, a transient timeout) must not wipe out the others. Previously a single
      // failed fetch in a Promise.all rejected the whole batch and silently reset everything
      // (staff, projects, users) back to nothing, which is exactly what happened when the
      // CampusConfig tab didn't exist yet.
      const [projectsResult, subtasksResult, staffResult, usersResult, campusConfigResult, campusesResult, notificationsResult, centralThreadsResult] = await Promise.allSettled([
        apiGet("Projects"), apiGet("Subtasks"), apiGet("Staff"), apiGet("Users"), apiGet("CampusConfig"), apiGet("Campuses"), apiGet("Notifications"), apiGet("CentralThreads"),
      ]);
      if (cancelled) return;

      const failures = [];
      const valueOr = (result, label, fallback) => {
        if (result.status === "fulfilled") return result.value;
        failures.push(`${label}: ${result.reason?.message || result.reason}`);
        return fallback;
      };

      const subtaskRows = valueOr(subtasksResult, "Subtasks", []);
      const userRows = valueOr(usersResult, "Users", []);
      const campusConfigRows = valueOr(campusConfigResult, "CampusConfig", []);
      const campusRows = valueOr(campusesResult, "Campuses", []);

      if (usersResult.status === "fulfilled") setUsers(userRows.map((u) => ({ ...u, id: Number(u.id) })));

      // RLS scopes notifications to the signed-in user automatically (for_user_name =
      // my_name()) and central_threads to central tier only — no client-side filtering needed.
      if (notificationsResult.status === "fulfilled") setNotifications(notificationsResult.value);

      if (centralThreadsResult.status === "fulfilled") {
        const byKey = {};
        centralThreadsResult.value.forEach((t) => { byKey[t.threadKey] = { tags: t.tags || [], messages: t.messages || [] }; });
        setCentralThreads(byKey);
      }

      // Replaces the old hardcoded campuses const — Central can now add/rename campuses
      // through the database instead of a code change. The "central" pseudo-campus row
      // (seeded to satisfy the campus_id FK elsewhere) isn't a real campus, so it's excluded
      // from what the UI shows as a campus.
      if (campusesResult.status === "fulfilled") setCampuses(campusRows.filter((c) => c.id !== "central"));

      if (campusConfigResult.status === "fulfilled") {
        const slidesLinks = {};
        campusConfigRows.forEach((r) => { if (r.slidesLink) slidesLinks[r.campusId] = r.slidesLink; });
        setCampusSlidesLinks(slidesLinks);
      }

      if (projectsResult.status === "fulfilled") {
        const projectRows = projectsResult.value;
        if (projectRows.length === 0) {
          // Empty Sheet means empty dashboard — no auto-reseeding. Demo data is only ever
          // loaded on purpose now (see "Load Demo Data" in Reports), never silently.
          setProjects([]);
        } else {
          const subtasksByProject = {};
          subtaskRows.forEach((s) => {
            const pid = Number(s.projectId);
            if (!subtasksByProject[pid]) subtasksByProject[pid] = [];
            subtasksByProject[pid].push({
              ...s, id: Number(s.id),
              done: s.done === true || s.done === "true",
              cost: Number(s.cost) || 0, spent: Number(s.spent) || 0,
              photos: Array.isArray(s.photos) ? s.photos : [],
              notes: Array.isArray(s.notes) ? s.notes : [],
              recurrence: s.recurrence && typeof s.recurrence === "object" ? s.recurrence : null,
            });
          });
          const loadedProjects = projectRows.map((p) => ({
            ...p,
            id: Number(p.id),
            cost: Number(p.cost) || 0,
            spent: Number(p.spent) || 0,
            shared: p.shared === true || p.shared === "true",
            subtasks: subtasksByProject[Number(p.id)] || [],
            team: Array.isArray(p.team) ? p.team : [],
            collaborators: Array.isArray(p.collaborators) ? p.collaborators : [],
            sharedWith: Array.isArray(p.sharedWith) ? p.sharedWith : [],
            photos: Array.isArray(p.photos) ? p.photos : [],
            notes: Array.isArray(p.notes) ? p.notes : [],
            recurrence: p.recurrence && typeof p.recurrence === "object" ? p.recurrence : null,
          }));
          setProjects(loadedProjects);
        }
      }

      // Staff: always reflects exactly what the Sheet has, including "nothing" — previously
      // an empty Staff sheet was silently ignored and the old seed roster stayed forever.
      if (staffResult.status === "fulfilled") {
        const staffRows = staffResult.value;
        if (staffRows.length === 0) {
          setStaffByCampus({});
        } else {
          const byCampus = {};
          staffRows.forEach((s) => {
            if (!byCampus[s.campusId]) byCampus[s.campusId] = [];
            byCampus[s.campusId].push({
              ...s, id: Number(s.id),
              roles: Array.isArray(s.roles) ? s.roles : [],
              calendars: Array.isArray(s.calendars) ? s.calendars : [],
              calendarSynced: s.calendarSynced === true || s.calendarSynced === "true",
            });
          });
          setStaffByCampus(byCampus);
        }
      }

      if (failures.length > 0) {
        setBackendStatus("offline");
        setBackendError(`${failures.length} sheet(s) failed to load. First: ${failures[0]}`);
      } else {
        setBackendStatus("connected");
      }
    })();
    return () => { cancelled = true; };
  }, [auth]);

  const loadDemoData = async () => {
    setBackendStatus("bootstrapping");
    try {
      await bootstrapBackend((done, total) => setBootstrapProgress({ done, total }));
      setBootstrapProgress(null);
      // reload from the Sheet now that it has data, rather than re-deriving locally
      const [projectRows, subtaskRows] = await Promise.all([apiGet("Projects"), apiGet("Subtasks")]);
      const subtasksByProject = {};
      subtaskRows.forEach((s) => {
        const pid = Number(s.projectId);
        if (!subtasksByProject[pid]) subtasksByProject[pid] = [];
        subtasksByProject[pid].push({ ...s, id: Number(s.id), done: s.done === true || s.done === "true", cost: Number(s.cost) || 0, spent: Number(s.spent) || 0 });
      });
      setProjects(projectRows.map((p) => ({ ...p, id: Number(p.id), cost: Number(p.cost) || 0, spent: Number(p.spent) || 0, subtasks: subtasksByProject[Number(p.id)] || [] })));
      setStaffByCampus(seedStaff);
      setBackendStatus("connected"); setBackendError("");
    } catch (err) {
      setBackendStatus("offline"); setBackendError(err?.message || String(err));
    }
  };

  const displayProjects = useMemo(() => projects.map(resolveProject), [projects]);

  const scopedProjects = useMemo(() => {
    if (!activeCampusId) return displayProjects;
    return displayProjects.filter(
      (p) => p.campus === activeCampusId || p.campus === "central" ||
        (p.shared && (p.sharedWith?.includes(activeCampusId) || p.sharedWith?.includes("all")))
    );
  }, [displayProjects, activeCampusId]);

  const staff = staffByCampus[activeCampusId] || [];
  // Central, viewing the org-wide overview (no campus drilled into), still gets a real staff
  // roster of its own — the Central Operations Team — via the same Staff & Team machinery
  // every campus uses, just scoped to the "central" pseudo-campus id instead of a real one.
  const staffCampusId = activeCampusId || (role === "central" ? "central" : null);
  const staffForPanel = staffByCampus[staffCampusId] || [];
  const notes = notesByCampus[activeCampusId] || [];
  const dayEvents = seedCalendar[calDate] || [];

  // Real activity log — every meaningful edit, add, or completion appends here, and the
  // Activity tab reads directly from this (not a static mock list).
  const logActivity = (action, target, campus, meta) => {
    setActivityLog((prev) => [
      { id: Date.now() + Math.random(), ts: TODAY_STR + " " + new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }), date: TODAY_STR, actor: currentViewerName, action, target, campus: campus || activeCampusId || "central", ...meta },
      ...prev,
    ]);
  };

  // Notifies everyone meaningfully connected to a project — owner, creator, team, that
  // campus's own Operations Director, plus anyone explicitly named (e.g. a subtask's
  // assignees) — never the person who just took the action. `summary` is the ready-to-read
  // description of what happened; recipients see "<actor> <summary>" in their bell.
  const notifyInvolved = (project, summary, extraRecipients) => {
    if (!project) return;
    const campus = campuses.find((c) => c.id === project.campus);
    const names = new Set([
      project.owner, project.createdBy, ...(project.team || []), ...(extraRecipients || []),
      campus?.lead,
    ].filter(Boolean));
    names.delete(currentViewerName);
    if (names.size === 0) return;
    const rows = Array.from(names).map((forUser) => ({
      forUser, actor: currentViewerName, summary, projectId: project.id, ts: TODAY_STR, read: false,
    }));
    setNotifications((prev) => [...rows.map((r) => ({ ...r, id: `${Date.now()}-${Math.random()}` })), ...prev]);
    syncToBackend(Promise.all(rows.map((r) => apiCreate("Notifications", { ...r, organizationId: OSC_ORG_ID }))));
  };

  // Bulk-marks every one of the viewer's own notifications read — a single UPDATE rather than
  // one apiUpdate per row, since RLS already scopes this to the signed-in user's own rows.
  const markAllNotificationsRead = () => {
    setNotifications((prev) => prev.map((n) => n.forUser === currentViewerName ? { ...n, read: true } : n));
    syncToBackend(
      supabase.from("notifications").update({ read: true }).eq("for_user_name", currentViewerName).eq("read", false)
        .then(({ error }) => { if (error) throw new Error(error.message); })
    );
  };

  const addProjectNote = (projId, text) => {
    if (!text.trim()) return;
    setProjects((ps) => ps.map((p) => p.id === projId
      ? { ...p, notes: [{ ts: TODAY_STR + " " + new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }), author: currentViewerName, text: text.trim() }, ...(p.notes || [])] }
      : p));
  };

  const addSubtaskNote = (projId, idx, text) => {
    if (!text.trim()) return;
    setProjects((ps) => ps.map((p) => {
      if (p.id !== projId) return p;
      const subtasks = [...p.subtasks];
      const s = subtasks[idx];
      subtasks[idx] = { ...s, notes: [{ ts: TODAY_STR, author: currentViewerName, text: text.trim() }, ...(s.notes || [])] };
      return { ...p, subtasks };
    }));
  };

  const addProjectPhoto = (projId, dataUrl, caption) =>
    setProjects((ps) => ps.map((p) => p.id === projId
      ? { ...p, photos: [{ id: Date.now(), dataUrl, caption, uploadedBy: currentViewerName, ts: TODAY_STR }, ...(p.photos || [])] }
      : p));

  const addSubtaskPhoto = (projId, idx, dataUrl, caption) =>
    setProjects((ps) => ps.map((p) => {
      if (p.id !== projId) return p;
      const subtasks = [...p.subtasks];
      const s = subtasks[idx];
      subtasks[idx] = { ...s, photos: [{ id: Date.now(), dataUrl, caption, uploadedBy: currentViewerName, ts: TODAY_STR }, ...(s.photos || [])] };
      return { ...p, subtasks };
    }));


  const cycleStage = (id) =>
    setProjects((ps) => ps.map((p) => p.id === id ? { ...p, stage: STAGES[(STAGES.indexOf(p.stage) + 1) % STAGES.length] } : p));

  const setProjectStage = (id, stage) => {
    let newCompletedOn = null;
    setProjects((ps) => ps.map((p) => {
      if (p.id !== id) return p;
      const resolved = resolveProject(p); // catch the project up to its current scheduled cycle first
      const wasDone = resolved.stage === "Completed";
      if (stage === "Completed" && !wasDone) {
        notifyInvolved(resolved, `marked "${resolved.title}" complete (ended ${BUDGET_STATUS_LABEL[budgetStatus(resolved)].toLowerCase()})`);
        logActivity("completed project", resolved.title, resolved.campus, { itemType: "project", createdBy: resolved.createdBy, budgetStatusVal: budgetStatus(resolved), projectId: resolved.id });
      } else if (stage !== resolved.stage) {
        notifyInvolved(resolved, `moved "${resolved.title}" to ${stage}`);
        logActivity(`moved "${resolved.title}" to ${stage}`, "", resolved.campus);
      }
      newCompletedOn = stage === "Completed" ? TODAY_STR : (stage !== "Completed" ? null : resolved.completedOn);
      return { ...resolved, stage, completedOn: newCompletedOn };
    }));
    syncToBackend(apiUpdate("Projects", id, { stage, completedOn: newCompletedOn }));
  };

  const setProjectDue = (id, due, dueTime) => {
    setProjects((ps) => ps.map((p) => {
      if (p.id !== id) return p;
      notifyInvolved(p, `changed the deadline on "${p.title}" to ${due}${dueTime ? ` at ${dueTime}` : ""}`);
      logActivity(`changed the deadline on "${p.title}"`, `now due ${due}`, p.campus);
      return { ...p, due, dueTime: dueTime || null };
    }));
    syncToBackend(apiUpdate("Projects", id, { due, dueTime: dueTime || null }));
  };

  const addProject = (data) => {
    logActivity("created project", data.title, data.campus || activeCampusId || "central");
    const newProject = {
      id: Date.now(),
      campus: data.campus || activeCampusId || "central",
      title: data.title,
      stage: data.stage || "Pending",
      owner: data.owner || "Unassigned",
      createdBy: currentViewerName,
      createdAt: TODAY_STR,
      completedOn: null,
      team: data.team ? data.team.split(",").map((s) => s.trim()).filter(Boolean) : [],
      collaborators: data.collaborators ? data.collaborators.split(",").map((s) => s.trim()).filter(Boolean) : [],
      due: data.due,
      dueTime: data.dueTime || null,
      cost: Number(data.cost) || 0,
      spent: 0,
      shared: !!data.shared,
      section: data.section || "General",
      recurrence: data.recurrence?.freq === "none" ? null : data.recurrence,
      subtasks: [],
      photos: [],
      notes: [],
    };
    setProjects((ps) => [...ps, newProject]);
    notifyInvolved(newProject, `created the project "${newProject.title}"`);
    const { subtasks, ...projectRow } = newProject;
    syncToBackend(apiCreate("Projects", projectRow));
  };

  // Marking a recurring task done never spawns or resets anything immediately — it just
  // records completion for the current scheduled cycle. Regeneration only happens once
  // TODAY_STR actually reaches the next scheduled date (see resolveSubtask/resolveProject).
  const toggleSubtask = (projId, idx) => {
    let updatedSubtask = null;
    setProjects((ps) => ps.map((p) => {
      if (p.id !== projId) return p;
      const resolved = resolveProject(p); // self-heal to the current cycle before toggling
      const subtasks = [...resolved.subtasks];
      const target = subtasks[idx];
      const newDone = !target.done;
      subtasks[idx] = { ...target, done: newDone };
      updatedSubtask = subtasks[idx];
      if (newDone) {
        notifyInvolved(resolved, `completed "${target.t}" on "${resolved.title}"`, target.assignees);
        logActivity("completed task", `${target.t} (${resolved.title})`, resolved.campus, { itemType: "task", createdBy: target.createdBy, projectId: resolved.id });
      } else {
        notifyInvolved(resolved, `reopened "${target.t}" on "${resolved.title}"`, target.assignees);
        logActivity("reopened task", target.t, resolved.campus);
      }
      return { ...resolved, subtasks };
    }));
    if (updatedSubtask) syncToBackend(apiUpdate("Subtasks", updatedSubtask.id, { done: updatedSubtask.done, due: updatedSubtask.due, recurrence: updatedSubtask.recurrence }));
  };

  const addSubtask = (projId, text, recurrence, due, cost, dueTime, assignees) => {
    const proj = projects.find((p) => p.id === projId);
    logActivity("added task", `"${text}" to ${proj?.title || "a project"}`, proj?.campus);
    const newSubtask = {
      id: Date.now(), t: text, done: false, due: due || null, dueTime: dueTime || null, assignees: assignees || [],
      createdBy: currentViewerName, photos: [], notes: [], cost: Number(cost) || 0, spent: 0, recurrence: recurrence?.freq === "none" ? null : recurrence,
    };
    setProjects((ps) => ps.map((p) => p.id === projId ? { ...p, subtasks: [...p.subtasks, newSubtask] } : p));
    if (proj) {
      const summary = (assignees || []).length > 0
        ? `assigned you "${text}" on "${proj.title}"${due ? `, due ${due}${dueTime ? ` at ${dueTime}` : ""}` : ""}`
        : `added the task "${text}" to "${proj.title}"`;
      notifyInvolved(proj, summary, assignees);
    }
    syncToBackend(apiCreate("Subtasks", { ...newSubtask, projectId: projId }));
  };

  const assignSubtask = (projId, idx, assignees) => {
    let subtaskId = null; let subtaskTitle = "";
    const proj = projects.find((p) => p.id === projId);
    setProjects((ps) => ps.map((p) => {
      if (p.id !== projId) return p;
      const subtasks = [...p.subtasks];
      subtaskTitle = subtasks[idx].t;
      subtasks[idx] = { ...subtasks[idx], assignees };
      subtaskId = subtasks[idx].id;
      return { ...p, subtasks };
    }));
    if (proj) notifyInvolved(proj, `assigned you to "${subtaskTitle}" on "${proj.title}"`, assignees);
    logActivity(`assigned "${subtaskTitle}"`, assignees.join(", "), proj?.campus);
    if (subtaskId) syncToBackend(apiUpdate("Subtasks", subtaskId, { assignees }));
  };

  const setSubtaskDue = (projId, idx, due, dueTime) => {
    let subtaskId = null;
    const proj = projects.find((p) => p.id === projId);
    setProjects((ps) => ps.map((p) => {
      if (p.id !== projId) return p;
      const subtasks = [...p.subtasks];
      logActivity(`changed the deadline on "${subtasks[idx].t}"`, `now due ${due}`, p.campus);
      if (proj) notifyInvolved(proj, `changed the deadline on "${subtasks[idx].t}" to ${due}${dueTime ? ` at ${dueTime}` : ""}`, subtasks[idx].assignees);
      subtasks[idx] = { ...subtasks[idx], due, dueTime: dueTime || null };
      subtaskId = subtasks[idx].id;
      return { ...p, subtasks };
    }));
    if (subtaskId) syncToBackend(apiUpdate("Subtasks", subtaskId, { due, dueTime: dueTime || null }));
  };

  const updateSubtaskBudget = (projId, idx, cost, spent) => {
    let subtaskId = null;
    const proj = projects.find((p) => p.id === projId);
    setProjects((ps) => ps.map((p) => {
      if (p.id !== projId) return p;
      const subtasks = [...p.subtasks];
      logActivity("updated task budget", subtasks[idx].t, p.campus);
      if (proj) notifyInvolved(proj, `updated the budget on "${subtasks[idx].t}"`, subtasks[idx].assignees);
      subtasks[idx] = { ...subtasks[idx], cost: Number(cost) || 0, spent: Number(spent) || 0 };
      subtaskId = subtasks[idx].id;
      return { ...p, subtasks };
    }));
    if (subtaskId) syncToBackend(apiUpdate("Subtasks", subtaskId, { cost: Number(cost) || 0, spent: Number(spent) || 0 }));
  };

  const updateProjectBudget = (projId, cost, spent) => {
    const proj = projects.find((p) => p.id === projId);
    logActivity("updated project budget", proj?.title, proj?.campus);
    if (proj) notifyInvolved(proj, `updated the budget on "${proj.title}"`);
    setProjects((ps) => ps.map((p) => p.id === projId ? { ...p, cost: Number(cost) || 0, spent: Number(spent) || 0 } : p));
    syncToBackend(apiUpdate("Projects", projId, { cost: Number(cost) || 0, spent: Number(spent) || 0 }));
  };

  const deleteProject = (projId) => {
    const proj = projects.find((p) => p.id === projId);
    if (proj) {
      notifyInvolved(proj, `deleted the project "${proj.title}"`);
      logActivity(`deleted "${proj.title}"`, "", proj.campus);
    }
    setProjects((ps) => ps.filter((p) => p.id !== projId));
    syncToBackend(apiDelete("Projects", projId));
  };

  const deleteSubtask = (projId, idx) => {
    const proj = projects.find((p) => p.id === projId);
    const subtask = proj?.subtasks?.[idx];
    if (proj && subtask) {
      notifyInvolved(proj, `deleted the task "${subtask.t}" from "${proj.title}"`, subtask.assignees);
      logActivity(`deleted task "${subtask.t}"`, proj.title, proj.campus);
    }
    setProjects((ps) => ps.map((p) => p.id === projId ? { ...p, subtasks: p.subtasks.filter((_, i) => i !== idx) } : p));
    if (subtask) syncToBackend(apiDelete("Subtasks", subtask.id));
  };

  const addSection = (name) => setSections((prev) => prev.includes(name) ? prev : [...prev, name]);
  const setProjectSection = (projId, section) => {
    setProjects((ps) => ps.map((p) => p.id === projId ? { ...p, section } : p));
    syncToBackend(apiUpdate("Projects", projId, { section }));
  };

  // One row per thread (organization_id, thread_key unique) — every mutation upserts the
  // thread's whole tags/messages state, same "small object, whole-row upsert" shape as
  // setCampusSlidesLink above.
  const persistCentralThread = (threadKey, thread) => {
    syncToBackend(
      supabase.from("central_threads")
        .upsert({ organization_id: OSC_ORG_ID, thread_key: String(threadKey), tags: thread.tags, messages: thread.messages }, { onConflict: "organization_id,thread_key" })
        .then(({ error }) => { if (error) throw new Error(error.message); })
    );
  };

  // Central-only tagging & discussion — never exposed to campus-scoped views.
  const addCentralTag = (projId, tag) => {
    const clean = tag.trim().replace(/^#/, "");
    if (!clean) return;
    setCentralThreads((prev) => {
      const t = prev[projId] || { tags: [], messages: [] };
      if (t.tags.includes(clean)) return prev;
      const updated = { ...t, tags: [...t.tags, clean] };
      persistCentralThread(projId, updated);
      return { ...prev, [projId]: updated };
    });
  };
  const removeCentralTag = (projId, tag) =>
    setCentralThreads((prev) => {
      const t = prev[projId] || { tags: [], messages: [] };
      const updated = { ...t, tags: t.tags.filter((x) => x !== tag) };
      persistCentralThread(projId, updated);
      return { ...prev, [projId]: updated };
    });
  const addCentralMessage = (threadKey, text, assignedTo) => {
    if (!text.trim()) return;
    setCentralThreads((prev) => {
      const t = prev[threadKey] || { tags: [], messages: [] };
      const msg = { author: currentViewerName, text: text.trim(), ts: TODAY_STR + " " + new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }), assignedTo: assignedTo || null };
      const updated = { ...t, messages: [...t.messages, msg] };
      persistCentralThread(threadKey, updated);
      return { ...prev, [threadKey]: updated };
    });
    if (assignedTo && assignedTo !== currentViewerName) {
      const note = { forUser: assignedTo, actor: currentViewerName, summary: `assigned you a central note: "${text.trim().slice(0, 60)}${text.trim().length > 60 ? "…" : ""}"`, projectId: null, ts: TODAY_STR, read: false };
      setNotifications((prev) => [{ ...note, id: `${Date.now()}-${Math.random()}` }, ...prev]);
      syncToBackend(apiCreate("Notifications", { ...note, organizationId: OSC_ORG_ID }));
    }
  };

  const addStaff = (campusId, name, roles) => {
    const newPerson = { id: Date.now(), name, roles: roles.slice(0, 2), nextMeeting: "Not yet scheduled", lastContact: "Just added", flag: null, calendarSynced: false, calendars: [] };
    setStaffByCampus((prev) => ({ ...prev, [campusId]: [...(prev[campusId] || []), newPerson] }));
    syncToBackend(apiCreate("Staff", { ...newPerson, campusId }));
  };

  // Second, more structured way to build the roster: a person added straight into one of the
  // 4 team lanes, with contact info and an optional reports-to captured up front — versus the
  // simpler free-text "Add Team Member" flow above, which stays as-is alongside this.
  const addTeamRole = (campusId, lane, data) => {
    const newPerson = {
      id: Date.now(),
      name: `${data.firstName} ${data.lastName}`.trim(),
      roles: [data.role],
      lane,
      email: data.email || "",
      phone: data.phone || "",
      reportsTo: data.reportsTo || null,
      nextMeeting: "Not yet scheduled", lastContact: "Just added", flag: null, calendarSynced: false, calendars: [],
    };
    setStaffByCampus((prev) => ({ ...prev, [campusId]: [...(prev[campusId] || []), newPerson] }));
    logActivity("added team role", `${newPerson.name} — ${data.role}`, campusId);
    syncToBackend(apiCreate("Staff", { ...newPerson, campusId }));
  };
  const removeStaff = (campusId, id) => {
    setStaffByCampus((prev) => ({ ...prev, [campusId]: (prev[campusId] || []).filter((s) => s.id !== id) }));
    syncToBackend(apiDelete("Staff", id));
  };
  const updateStaffRoles = (campusId, id, roles) => {
    setStaffByCampus((prev) => ({ ...prev, [campusId]: (prev[campusId] || []).map((s) => s.id === id ? { ...s, roles: roles.slice(0, 2) } : s) }));
    syncToBackend(apiUpdate("Staff", id, { roles: roles.slice(0, 2) }));
  };
  const setStaffCalendars = (campusId, id, calendars) => {
    setStaffByCampus((prev) => ({ ...prev, [campusId]: (prev[campusId] || []).map((s) => s.id === id ? { ...s, calendars, calendarSynced: calendars.length > 0 } : s) }));
    syncToBackend(apiUpdate("Staff", id, { calendars, calendarSynced: calendars.length > 0 }));
  };
  const addRoleOption = (roleName) => {
    const trimmed = roleName.trim();
    if (trimmed && !roleOptions.includes(trimmed)) setRoleOptions((prev) => [...prev, trimmed]);
  };

  // Persists which of their own real Google Calendars a user has chosen to show — not the
  // OAuth connection itself (that's session-only, see requestGoogleCalendarToken). Anyone can
  // update their own Users row (Code.gs already allows this: auth.userId === id), so this
  // needs no backend change.
  const saveMyGoogleCalendars = (calendarIds, calendarNames) => {
    setUsers((prev) => prev.map((u) => u.id === auth.user.id ? { ...u, googleCalendarIds: calendarIds, googleCalendarNames: calendarNames } : u));
    setAuth({ ...auth, user: { ...auth.user, googleCalendarIds: calendarIds, googleCalendarNames: calendarNames } });
    syncToBackend(apiUpdate("Users", auth.user.id, { googleCalendarIds: calendarIds, googleCalendarNames: calendarNames }));
  };

  // Reflects the permanent connect/disconnect Code.gs already performed server-side (the
  // actual connect/disconnect API calls happen in CalendarPanel) into local state, so the rest
  // of the app immediately sees authUser.googleCalendarConnected change without a reload.
  const setMyGoogleConnected = (connected) => {
    setUsers((prev) => prev.map((u) => u.id === auth.user.id ? { ...u, googleCalendarConnected: connected } : u));
    setAuth({ ...auth, user: { ...auth.user, googleCalendarConnected: connected } });
  };

  // Persists a campus's Slides org-chart link to the CampusConfig sheet (id = campusId, so
  // this is a plain generic-CRUD upsert: create the row the first time, update it after that).
  const setCampusSlidesLink = (campusId, url) => {
    const hadRow = campusSlidesLinks[campusId] !== undefined;
    setCampusSlidesLinks((prev) => ({ ...prev, [campusId]: url }));
    syncToBackend(hadRow
      ? apiUpdate("CampusConfig", campusId, { slidesLink: url })
      : apiCreate("CampusConfig", { campusId, slidesLink: url }));
  };

  // Mirrors tierForRole_ in Code.gs — kept in sync locally purely so the Team Accounts panel
  // can show the resulting tier immediately, without waiting on a round trip. The backend
  // always recomputes this itself from campusId/role on write and is the actual source of truth.
  const tierFromAccess = (campusId, role) => campusId === "central" ? "central" : (role === "Campus Operations Director" ? "od" : "staff");

  // Creates a real login — a normal, ordinary account, exactly like every account after the
  // very first. Only Central can do this (enforced server-side too, via the admin-accounts
  // Edge Function — creating an account for someone else with a chosen password needs the
  // service role key, which never runs in the browser). The new profile's id is assigned by
  // Supabase Auth, so unlike other creates here this can't be applied optimistically — local
  // state waits for the Edge Function to return the real id.
  const addUserAccount = (data) => {
    const tier = tierFromAccess(data.campusId, data.role);
    setSyncing(true);
    apiAdminAccounts("create", { firstName: data.firstName, lastName: data.lastName, email: data.email, phone: data.phone || "", campusId: data.campusId, role: data.role, password: data.password })
      .then((result) => {
        const newUser = { id: result.id, firstName: data.firstName, lastName: data.lastName, email: data.email, phone: data.phone || "", campusId: data.campusId, role: data.role, tier, createdAt: TODAY_STR };
        setUsers((prev) => [...prev, newUser]);
        logActivity("created a login for", `${data.firstName} ${data.lastName}`.trim(), data.campusId);
        setBackendStatus("connected"); setBackendError("");
      })
      .catch((err) => { setBackendStatus("offline"); setBackendError(err?.message || String(err)); })
      .finally(() => setSyncing(false));
  };

  // Reassigns which campus (or Central) and role a login has access to. Routed through
  // admin-accounts rather than a direct profiles update — the column-grant policy that lets
  // Central update others' rows would, if used to also let it change tier/campus/role, equally
  // let anyone self-elevate, so that recompute happens server-side under the service role.
  const updateUserAccess = (id, campusId, role) => {
    const tier = tierFromAccess(campusId, role);
    setSyncing(true);
    apiAdminAccounts("updateAccess", { userId: id, campusId, role })
      .then(() => {
        setUsers((prev) => prev.map((u) => u.id === id ? { ...u, campusId, role, tier } : u));
        logActivity("updated account access for", users.find((u) => u.id === id)?.email || String(id), campusId);
        setBackendStatus("connected"); setBackendError("");
      })
      .catch((err) => { setBackendStatus("offline"); setBackendError(err?.message || String(err)); })
      .finally(() => setSyncing(false));
  };

  const removeUserAccount = (id) => {
    setSyncing(true);
    apiAdminAccounts("delete", { userId: id })
      .then(() => { setUsers((prev) => prev.filter((u) => u.id !== id)); setBackendStatus("connected"); setBackendError(""); })
      .catch((err) => { setBackendStatus("offline"); setBackendError(err?.message || String(err)); })
      .finally(() => setSyncing(false));
  };

  // Associates a Staff & Team roster entry with a real login (Users row) — the two stay
  // separate systems (Staff is directory/assignment info, Users is auth), this just cross-
  // references them so a roster entry can show and manage the account tied to that person.
  const linkStaffUser = (campusId, staffId, userId) => {
    setStaffByCampus((prev) => ({ ...prev, [campusId]: (prev[campusId] || []).map((s) => s.id === staffId ? { ...s, userId } : s) }));
    syncToBackend(apiUpdate("Staff", staffId, { userId }));
  };

  // Creates a brand-new login for someone already on the roster and links it in the same
  // step — the common case (this person doesn't have an account yet) in one action instead of
  // creating it in Team Accounts and then separately linking it here. Same server-assigned-id
  // constraint as addUserAccount, so the link only happens once the real id comes back.
  const createAndLinkUser = (campusId, staffId, data) => {
    const staffPerson = (staffByCampus[campusId] || []).find((s) => s.id === staffId);
    const role = staffPerson?.roles?.[0] || "";
    const tier = tierFromAccess(campusId, role);
    setSyncing(true);
    apiAdminAccounts("create", { firstName: data.firstName, lastName: data.lastName, email: data.email, phone: data.phone || "", campusId, role, password: data.password })
      .then((result) => {
        const newUser = { id: result.id, firstName: data.firstName, lastName: data.lastName, email: data.email, phone: data.phone || "", campusId, role, tier, createdAt: TODAY_STR };
        setUsers((prev) => [...prev, newUser]);
        linkStaffUser(campusId, staffId, newUser.id);
        setBackendStatus("connected"); setBackendError("");
      })
      .catch((err) => { setBackendStatus("offline"); setBackendError(err?.message || String(err)); })
      .finally(() => setSyncing(false));
  };

  // Commits a reviewed org-chart roster for a campus. Any role that's unchanged or brand new
  // gets applied straight to the staff roster. Any role that's held by a DIFFERENT person than
  // last time gets queued as a reassignment prompt instead of silently overwriting — the OD
  // has to confirm before we touch existing task ownership.
  const commitOrgChart = (campusId, people) => {
    const prevChart = orgChartByCampus[campusId] || [];
    const prevByRole = {};
    prevChart.forEach((p) => { prevByRole[p.role] = p.name; });

    const toApplyNow = [];
    const toQueue = [];
    people.forEach((p) => {
      const prevName = prevByRole[p.role];
      if (prevName && prevName !== p.name) {
        toQueue.push({ id: `${campusId}-${p.role}-${Date.now()}-${Math.random()}`, campusId, role: p.role, reportsTo: p.reportsTo, oldName: prevName, newName: p.name });
      } else {
        toApplyNow.push(p);
      }
    });

    setOrgChartByCampus((prev) => ({ ...prev, [campusId]: people }));

    setStaffByCampus((prev) => {
      const existing = prev[campusId] || [];
      let updated = [...existing];
      toApplyNow.forEach((p) => {
        const match = updated.find((s) => s.name === p.name);
        if (match) {
          const roles = match.roles?.includes(p.role) ? match.roles : [...(match.roles || []), p.role].slice(0, 2);
          updated = updated.map((s) => s.id === match.id ? { ...s, roles, reportsTo: p.reportsTo } : s);
        } else {
          updated.push({ id: Date.now() + Math.random(), name: p.name, roles: [p.role], reportsTo: p.reportsTo, nextMeeting: "Not yet scheduled", lastContact: "Imported from org chart", flag: null, calendarSynced: false, calendars: [] });
        }
      });
      return { ...prev, [campusId]: updated };
    });

    if (toQueue.length > 0) setPendingReassignments((prev) => [...prev, ...toQueue]);
  };

  // OD confirms a role reassignment: updates the staff roster (swap the role from the outgoing
  // person to the incoming one) and reassigns any open projects/sub-tasks in that campus still
  // owned by the outgoing name over to the incoming name.
  const resolveReassignment = (id, confirm) => {
    const item = pendingReassignments.find((r) => r.id === id);
    if (!item) return;
    setPendingReassignments((prev) => prev.filter((r) => r.id !== id));
    if (!confirm) return;

    setStaffByCampus((prev) => {
      const existing = prev[item.campusId] || [];
      let updated = existing.map((s) =>
        s.name === item.oldName ? { ...s, roles: (s.roles || []).filter((r) => r !== item.role) } : s
      );
      const incoming = updated.find((s) => s.name === item.newName);
      if (incoming) {
        const roles = incoming.roles?.includes(item.role) ? incoming.roles : [...(incoming.roles || []), item.role].slice(0, 2);
        updated = updated.map((s) => s.id === incoming.id ? { ...s, roles, reportsTo: item.reportsTo } : s);
      } else {
        updated.push({ id: Date.now() + Math.random(), name: item.newName, roles: [item.role], reportsTo: item.reportsTo, nextMeeting: "Not yet scheduled", lastContact: "Imported from org chart", flag: null, calendarSynced: false, calendars: [] });
      }
      return { ...prev, [item.campusId]: updated };
    });

    setProjects((ps) => ps.map((p) => {
      if (p.campus !== item.campusId) return p;
      const owner = p.owner === item.oldName ? item.newName : p.owner;
      const team = (p.team || []).map((t) => t === item.oldName ? item.newName : t);
      return { ...p, owner, team };
    }));
  };

  const addNote = () => {
    if (!newNote.trim() || !activeCampusId) return;
    const n = { id: Date.now(), text: newNote.trim(), ts: "2026-07-10 " + new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) };
    setNotesByCampus((prev) => ({ ...prev, [activeCampusId]: [n, ...(prev[activeCampusId] || [])] }));
    setNewNote("");
  };
  const removeNote = (id) =>
    setNotesByCampus((prev) => ({ ...prev, [activeCampusId]: (prev[activeCampusId] || []).filter((n) => n.id !== id) }));

  const orgBudget = rollupBudget(displayProjects);
  const orgBudgetUsed = orgBudget.spent;
  const orgBudgetTotal = orgBudget.total;

  const navItems = [
    { id: "overview", label: role === "central" && !selectedCampus ? "All Campuses" : "Dashboard", icon: LayoutGrid },
    { id: "projects", label: "Projects & Tasks", icon: ListChecks },
    { id: "budget", label: "Budget", icon: DollarSign },
    { id: "staff", label: "Staff & Team", icon: Users },
    { id: "calendar", label: "Calendar", icon: CalendarDays },
    { id: "notes", label: "Notes", icon: StickyNote },
    { id: "activity", label: "Activity", icon: ActivityIcon },
    { id: "events", label: "Cross-Campus Events", icon: Link2 },
    ...(auth?.user?.tier === "central" ? [{ id: "accounts", label: "User Management", icon: ShieldCheck }] : []),
    { id: "reports", label: "Reports", icon: FileText },
  ];

  if (authLoading) return null; // Supabase checking for an existing session — near-instant, avoids a login-screen flash
  if (!auth) return <LoginScreen />;
  if (auth.user.tier === "unassigned") return <PendingAssignmentScreen user={auth.user} onSignOut={() => supabase.auth.signOut()} />;

  return (
    <div style={{ fontFamily: "'Inter', sans-serif", background: "#F7F6FB", color: "#2A2A3A" }} className="min-h-screen w-full">
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fraunces:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap" />
      <style dangerouslySetInnerHTML={{ __html: `
      /* Explicit fallback rules using attribute selectors — no escaping needed,
         guarantees theme colors render regardless of Tailwind JIT behavior. */
      [class*="bg-[#2B4C7E]"] { background-color: #2B4C7E !important; }
      [class*="bg-[#B8862F]"] { background-color: #B8862F !important; }
      [class*="bg-[#E3E1F0]"] { background-color: #E3E1F0 !important; }
      [class*="bg-[#EFEAFB]"] { background-color: #EFEAFB !important; }
      [class*="bg-[#EFEEFA]"] { background-color: #EFEEFA !important; }
      [class*="bg-[#F3EFFC]"] { background-color: #F3EFFC !important; }
      [class*="bg-[#FFFFFF]"] { background-color: #FFFFFF !important; }
      [class*="border-[#5E9E8A]"] { border-color: #5E9E8A !important; }
      [class*="border-[#B8862F]"] { border-color: #B8862F !important; }
      [class*="border-[#D8D5EC]"] { border-color: #D8D5EC !important; }
      [class*="border-[#D9CFF0]"] { border-color: #D9CFF0 !important; }
      [class*="border-[#E3E1F0]"] { border-color: #E3E1F0 !important; }
      [class*="text-[#2A2A3A]"] { color: #2A2A3A !important; }
      [class*="text-[#2B4C7E]"] { color: #2B4C7E !important; }
      [class*="text-[#5E9E8A]"] { color: #5E9E8A !important; }
      [class*="text-[#6B6980]"] { color: #6B6980 !important; }
      [class*="text-[#8B889C]"] { color: #8B889C !important; }
      [class*="text-[#A8A5BE]"] { color: #A8A5BE !important; }
      [class*="text-[#B8862F]"] { color: #B8862F !important; }
      [class*="text-[#C15B5B]"] { color: #C15B5B !important; }
      [class*="text-[#F7F6FB]"] { color: #F7F6FB !important; }
      [class*="hover:bg-[#EFEEFA]"]:hover { background-color: #EFEEFA !important; }
      [class*="hover:border-[#C7C3E0]"]:hover { border-color: #C7C3E0 !important; }
      [class*="hover:text-[#2A2A3A]"]:hover { color: #2A2A3A !important; }
      [class*="hover:text-[#B8862F]"]:hover { color: #B8862F !important; }
      [class*="hover:text-[#C15B5B]"]:hover { color: #C15B5B !important; }
      ` }} />

      <header className="px-6 py-4 flex items-center justify-between sticky top-0 backdrop-blur z-20 gap-3 flex-wrap" style={{ background: "rgba(43,76,126,0.97)" }}>
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-md flex items-center justify-center shrink-0" style={{ background: "#F7F6FB" }}><Building2 size={17} color="#2B4C7E" /></div>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", color: "#F7F6FB" }} className="text-[17px] font-semibold leading-none tracking-tight">OpsCore</div>
            <div className="text-[11px] mt-0.5 tracking-wide" style={{ color: "#C9D6E8" }}>Reaching People · Building Lives</div>
          </div>
          {role === "central" && selectedCampus && (
            <button onClick={() => { setSelectedCampus(null); setTab("overview"); }}
              className="flex items-center gap-1.5 text-[12px] font-medium rounded-md px-3 py-1.5 ml-1"
              style={{ background: "#F7F6FB", color: "#2B4C7E" }}>
              <LayoutGrid size={13} /> Return to Dashboard
            </button>
          )}
        </div>
        <div className="flex items-center gap-2">
          <div className="relative">
            <button onClick={() => setShowNotifications((v) => !v)} className="relative w-8 h-8 rounded-md flex items-center justify-center" style={{ background: "rgba(247,246,251,0.14)", border: "1px solid rgba(247,246,251,0.35)" }}>
              <Bell size={15} color="#F7F6FB" />
              {myNotifications.some((n) => !n.read) && (
                <span className="absolute -top-1 -right-1 w-4 h-4 rounded-full flex items-center justify-center text-[9px] font-bold" style={{ background: "#C15B5B", color: "#FFFFFF" }}>
                  {myNotifications.filter((n) => !n.read).length}
                </span>
              )}
            </button>
            {showNotifications && (
              <div className="absolute right-0 top-10 w-[300px] max-h-[380px] overflow-y-auto rounded-lg shadow-xl z-30" style={{ background: "#FFFFFF", border: "1px solid #D8D5EC" }}>
                <div className="px-3 py-2.5 border-b border-[#E3E1F0] flex items-center justify-between">
                  <span className="text-[12.5px] font-bold text-[#2A2A3A]">Notifications</span>
                  {myNotifications.length > 0 && (
                    <button onClick={markAllNotificationsRead} className="text-[10.5px] text-[#B8862F]">Mark all read</button>
                  )}
                </div>
                {myNotifications.length === 0 && <div className="text-[11.5px] text-[#6B6980] px-3 py-6 text-center">Nothing yet, {currentViewerName}.</div>}
                {myNotifications.map((n) => (
                  <div key={n.id} className="px-3 py-2.5 border-b border-[#E3E1F0] last:border-0" style={{ background: n.read ? "transparent" : "#5E9E8A0F" }}>
                    <div className="text-[12px] text-[#2A2A3A]"><span className="font-medium">{n.actor}</span> {n.summary}</div>
                    <div className="text-[10px] text-[#8B889C] mt-0.5">{n.ts}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
          {auth.user.tier === "central" ? (
            <select value={role} onChange={(e) => { setRole(e.target.value); setSelectedCampus(null); setTab("overview"); }}
              className="text-[13px] rounded-md px-2.5 py-1.5 outline-none"
              style={{ background: "rgba(247,246,251,0.14)", border: "1px solid rgba(247,246,251,0.35)", color: "#F7F6FB" }}>
              <option value="central" style={{ color: "#2A2A3A" }}>Central Operations Director</option>
              {campuses.map((c) => <option key={c.id} value={c.id} style={{ color: "#2A2A3A" }}>{c.name} ({c.abbr}) — Campus OD</option>)}
            </select>
          ) : (
            <div className="text-[12.5px] px-2.5 py-1.5" style={{ color: "#C9D6E8" }}>{activeCampus?.name} ({activeCampus?.abbr})</div>
          )}
          <div className="flex items-center gap-2 pl-2 ml-1" style={{ borderLeft: "1px solid rgba(247,246,251,0.25)" }}>
            <span className="text-[12px]" style={{ color: "#C9D6E8" }}>{currentViewerName}</span>
            <button onClick={() => setAuth(null)} className="text-[11.5px] rounded-md px-2.5 py-1.5"
              style={{ background: "rgba(247,246,251,0.14)", border: "1px solid rgba(247,246,251,0.35)", color: "#F7F6FB" }}>
              Sign out
            </button>
          </div>
        </div>
      </header>

      <div className="flex">
        <nav className="w-[188px] shrink-0 border-r border-[#E3E1F0] px-3 py-5 hidden md:flex flex-col gap-1">
          {navItems.map((item) => (
            <button key={item.id} onClick={() => setTab(item.id)}
              className={`flex items-center gap-2.5 px-3 py-2 rounded-md text-[13px] text-left transition ${tab === item.id ? "bg-[#E3E1F0] text-[#2A2A3A]" : "text-[#6B6980] hover:text-[#2A2A3A] hover:bg-[#EFEEFA]"}`}>
              <item.icon size={15} strokeWidth={2} />{item.label}
            </button>
          ))}
          {role === "central" && selectedCampus && (
            <button onClick={() => setSelectedCampus(null)} className="mt-4 flex items-center gap-2 px-3 py-2 rounded-md text-[12px] text-[#B8862F] hover:bg-[#EFEEFA]">
              <ChevronLeft size={14} /> All campuses
            </button>
          )}
          <div className="mt-auto pt-4 border-t border-[#E3E1F0]">
            <div className="flex items-center gap-1.5 px-1 mb-1">
              <span className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: backendStatus === "connected" ? "#5E9E8A" : (backendStatus === "loading" || backendStatus === "bootstrapping") ? "#B8862F" : "#C15B5B" }} />
              <span className="text-[10px] font-medium" style={{ color: backendStatus === "connected" ? "#5E9E8A" : (backendStatus === "loading" || backendStatus === "bootstrapping") ? "#B8862F" : "#C15B5B" }}>
                {backendStatus === "connected" ? (syncing ? "Saving…" : "Synced to Sheet")
                  : backendStatus === "loading" ? "Connecting…"
                  : backendStatus === "bootstrapping" ? `Setting up your Sheet… ${bootstrapProgress ? `${bootstrapProgress.done}/${bootstrapProgress.total}` : ""}`
                  : "Offline — not saving"}
              </span>
            </div>
            {backendStatus === "offline" && <div className="text-[9.5px] text-[#C15B5B] leading-snug px-1">{backendError || "Couldn't reach the backend."}</div>}
          </div>
        </nav>

        <main className="flex-1 min-w-0 px-6 py-6 max-w-[1180px]">
          <div className="flex md:hidden gap-1 mb-5 overflow-x-auto">
            {navItems.map((item) => (
              <button key={item.id} onClick={() => setTab(item.id)}
                className={`px-3 py-1.5 rounded-full text-[12px] whitespace-nowrap border ${tab === item.id ? "bg-[#B8862F] text-[#F7F6FB] border-[#B8862F]" : "border-[#D8D5EC] text-[#6B6980]"}`}>
                {item.label}
              </button>
            ))}
          </div>

          {tab === "overview" && role === "central" && !selectedCampus && (
            <CentralOverview campuses={campuses} orgBudgetUsed={orgBudgetUsed} orgBudgetTotal={orgBudgetTotal}
              projects={displayProjects} staffByCampus={staffByCampus} onSelectCampus={(id) => { setSelectedCampus(id); setTab("overview"); }}
              centralThreads={centralThreads} onAddTag={addCentralTag} onRemoveTag={removeCentralTag} onAddMessage={addCentralMessage}
              currentViewerName={currentViewerName} onOpenProject={setOpenProject}
              detail={detail} setDetail={setDetail} />
          )}

          {tab === "overview" && (role !== "central" || selectedCampus) && activeCampus && (
            <CampusDashboard
              campus={activeCampus} projects={scopedProjects} staff={staff} notes={notes} activity={activityLog}
              dayEvents={dayEvents} calDate={calDate}
              onOpenProject={setOpenProject} onGoTab={setTab} onOpenProfile={setOpenStaffProfile}
            />
          )}

          {tab === "projects" && (
            <ProjectsBoard projects={scopedProjects} campusLabel={activeCampus ? activeCampus.name : "All Campuses"}
              onCycle={cycleStage} onOpen={setOpenProject} sections={sections} onAddSection={addSection} onSetSection={setProjectSection}
              onNewProject={() => setShowNewProject(true)} />
          )}

          {tab === "budget" && (
            <BudgetPanel
              projects={activeCampusId ? displayProjects.filter((p) => p.campus === activeCampusId) : displayProjects}
              campuses={campuses} campusLabel={activeCampus ? activeCampus.name : "All Campuses"}
              onOpenProject={setOpenProject} onSelectCampus={activeCampusId ? null : (id) => { setSelectedCampus(id); setTab("budget"); }}
              accentColor={activeCampus?.color}
            />
          )}

          {tab === "staff" && (
            <StaffPanel
              staff={staffForPanel} campusLabel={activeCampus ? activeCampus.name : (staffCampusId === "central" ? "Central Operations Team" : "Select a campus")} full
              campusId={staffCampusId} roleOptions={roleOptions}
              onAdd={addStaff} onAddTeamRole={addTeamRole} onRemove={removeStaff} onUpdateRoles={updateStaffRoles}
              onSetCalendars={setStaffCalendars} onAddRoleOption={addRoleOption}
              slidesLink={campusSlidesLinks[staffCampusId] || ""}
              onSetSlidesLink={(url) => setCampusSlidesLink(staffCampusId, url)}
              campusPhase={activeCampus?.phase}
              onCommitOrgChart={commitOrgChart}
              pendingReassignments={pendingReassignments.filter((r) => r.campusId === staffCampusId)}
              onResolveReassignment={resolveReassignment}
              onOpenProfile={setOpenStaffProfile}
              users={users} onLinkUser={linkStaffUser} onCreateAndLinkUser={createAndLinkUser}
            />
          )}

          {tab === "calendar" && (
            <CalendarPanel calDate={calDate} setCalDate={setCalDate} calView={calView} setCalView={setCalView} dayEvents={dayEvents} full
              authUser={auth.user} onSaveGoogleCalendars={saveMyGoogleCalendars} onSetGoogleConnected={setMyGoogleConnected} />
          )}

          {tab === "notes" && (
            <NotesPanel notes={notes} newNote={newNote} setNewNote={setNewNote} addNote={addNote} removeNote={removeNote}
              campusLabel={activeCampus ? activeCampus.name : "Select a campus"} full />
          )}

          {tab === "activity" && (
            <ActivityPanel
              activity={activeCampusId ? activityLog.filter((a) => a.campus === activeCampusId) : activityLog}
              full projects={displayProjects} currentViewerName={currentViewerName}
            />
          )}

          {tab === "events" && <EventsList campusId={activeCampusId} campuses={campuses} />}

          {tab === "accounts" && auth.user.tier === "central" && (
            <AccountsPanel users={users} campuses={campuses} roleOptions={roleOptions}
              onCreate={addUserAccount} onUpdateAccess={updateUserAccess} onRemove={removeUserAccount}
              currentUserId={auth.user.id} />
          )}

          {tab === "reports" && (
            <ReportsPanel projects={displayProjects} campuses={campuses} staffByCampus={staffByCampus} roster={roster} isCentral={role === "central" && !selectedCampus} onClearAllProjects={clearAllProjects} onLoadDemoData={loadDemoData} />
          )}
        </main>
      </div>

      {openProject && (
        <ProjectModal
          project={displayProjects.find((p) => p.id === openProject)} onClose={() => setOpenProject(null)}
          onToggleSubtask={toggleSubtask} onAddSubtask={addSubtask} onSetStage={setProjectStage}
          onAddProjectNote={addProjectNote} onAddSubtaskNote={addSubtaskNote}
          onAddProjectPhoto={addProjectPhoto} onAddSubtaskPhoto={addSubtaskPhoto}
          onUpdateProjectBudget={updateProjectBudget} onUpdateSubtaskBudget={updateSubtaskBudget}
          onSetDue={setProjectDue} onSetSubtaskDue={setSubtaskDue}
          onAssignSubtask={assignSubtask} onDeleteSubtask={deleteSubtask} onDeleteProject={deleteProject}
          roster={fullRoster}
        />
      )}
      {showNewProject && <NewProjectModal onClose={() => setShowNewProject(false)} onCreate={(data) => { addProject(data); setShowNewProject(false); }} sections={sections} campusRoster={campusRoster} fullRoster={fullRoster} campusLabel={activeCampus ? `${activeCampus.name} (${activeCampus.abbr})` : "Central"} />}
      {openStaffProfile && <StaffProfileModal name={openStaffProfile} projects={displayProjects} onClose={() => setOpenStaffProfile(null)} onOpenProject={(id) => { setOpenStaffProfile(null); setOpenProject(id); }} />}
    </div>
  );
}

// Central-only collaboration space. This entire component only ever renders inside CentralOverview,
// which itself only renders when role === "central" — campus-scoped views never mount this,
// so tagged discussions are structurally invisible to any campus OD.
function CentralTeamWindow({ projects, campuses, centralThreads, onAddTag, onRemoveTag, onAddMessage, currentViewerName }) {
  const [pickerOpen, setPickerOpen] = useState(false);
  const [browseCampus, setBrowseCampus] = useState("");
  const [browseProjectId, setBrowseProjectId] = useState("");
  const [browseSubtaskId, setBrowseSubtaskId] = useState("");
  const [noteText, setNoteText] = useState("");
  const [assignTo, setAssignTo] = useState("");
  const [tagInput, setTagInput] = useState("");
  const [messageDrafts, setMessageDrafts] = useState({});
  const [assignDrafts, setAssignDrafts] = useState({});

  const sortedCampuses = [...campuses].sort((a, b) => a.name.localeCompare(b.name));
  const projectsInBrowseCampus = projects.filter((p) => p.campus === browseCampus).sort((a, b) => a.title.localeCompare(b.title));
  const browseProject = projects.find((p) => p.id === Number(browseProjectId));

  const threadKeyFor = (projId, subtaskId) => subtaskId ? `${projId}:${subtaskId}` : `${projId}`;

  const threadLabel = (key) => {
    const [projId, subId] = String(key).split(":");
    const p = projects.find((pr) => pr.id === Number(projId));
    if (!p) return "Unknown item";
    if (!subId) return p.title;
    const s = (p.subtasks || []).find((st) => String(st.id) === subId);
    return `${p.title} → ${s ? s.t : "task"}`;
  };

  const activeThreadIds = Object.keys(centralThreads).filter((k) => (centralThreads[k]?.tags?.length || 0) > 0 || (centralThreads[k]?.messages?.length || 0) > 0);

  const submitNote = () => {
    if (!browseProjectId || !noteText.trim()) return;
    const key = threadKeyFor(browseProjectId, browseSubtaskId);
    onAddMessage(key, noteText, assignTo);
    setNoteText(""); setAssignTo(""); setPickerOpen(false);
    setBrowseCampus(""); setBrowseProjectId(""); setBrowseSubtaskId("");
  };

  const sendMessage = (key) => {
    const text = messageDrafts[key];
    if (!text?.trim()) return;
    onAddMessage(key, text, assignDrafts[key] || "");
    setMessageDrafts((prev) => ({ ...prev, [key]: "" }));
    setAssignDrafts((prev) => ({ ...prev, [key]: "" }));
  };

  return (
    <div className="rounded-lg p-4 mb-2" style={{ background: "#EFEAFB", border: "1px solid #D9CFF0" }}>
      <div className="flex items-center justify-between mb-1">
        <div className="flex items-center gap-2 text-[13px] font-medium"><ShieldCheck size={14} style={{ color: "#6B4FA0" }} /> Central Team</div>
        <div className="flex items-center gap-2">
          <span className="text-[10px] text-[#6B6980]">Posting as</span>
          <span className="text-[11px] font-medium px-2 py-1 rounded-md" style={{ background: "#F3EFFC", border: "1px solid #D9CFF0" }}>{currentViewerName}</span>
        </div>
      </div>
      <p className="text-[11px] text-[#6B6980] mb-3">Private to Central Ops — browse any project or task by campus, leave a note, and optionally assign it to a teammate. Campus teams never see this.</p>

      {!pickerOpen ? (
        <button onClick={() => setPickerOpen(true)} className="flex items-center gap-1.5 text-[11.5px] rounded-md px-3 py-1.5 font-medium mb-3" style={{ background: "#6B4FA0", color: "#F3EFFC" }}>
          <Plus size={13} /> Note a project or task
        </button>
      ) : (
        <div className="rounded-md p-3 mb-3 space-y-2" style={{ background: "#F3EFFC", border: "1px solid #D9CFF0" }}>
          <select value={browseCampus} onChange={(e) => { setBrowseCampus(e.target.value); setBrowseProjectId(""); setBrowseSubtaskId(""); }}
            className="w-full bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2 py-1.5 text-[11.5px] outline-none">
            <option value="">Select campus…</option>
            <option value="central">Central — Org-Wide</option>
            {sortedCampuses.map((c) => <option key={c.id} value={c.id}>{c.name} ({c.abbr})</option>)}
          </select>

          {browseCampus && (
            <select value={browseProjectId} onChange={(e) => { setBrowseProjectId(e.target.value); setBrowseSubtaskId(""); }}
              className="w-full bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2 py-1.5 text-[11.5px] outline-none">
              <option value="">Select project…</option>
              {projectsInBrowseCampus.map((p) => <option key={p.id} value={p.id}>{p.title}</option>)}
            </select>
          )}

          {browseProject && (
            <select value={browseSubtaskId} onChange={(e) => setBrowseSubtaskId(e.target.value)}
              className="w-full bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2 py-1.5 text-[11.5px] outline-none">
              <option value="">Whole project (no specific task)</option>
              {(browseProject.subtasks || []).map((s) => <option key={s.id} value={s.id}>Task: {s.t}</option>)}
            </select>
          )}

          {browseProjectId && (
            <>
              <textarea value={noteText} onChange={(e) => setNoteText(e.target.value)} placeholder="Your note…" rows={2}
                className="w-full bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2.5 py-1.5 text-[11.5px] outline-none resize-none" />
              <select value={assignTo} onChange={(e) => setAssignTo(e.target.value)} className="w-full bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2 py-1.5 text-[11.5px] outline-none">
                <option value="">Don't assign — just a note</option>
                {CENTRAL_TEAM.filter((n) => n !== currentViewerName).map((n) => <option key={n} value={n}>Assign to {n}</option>)}
              </select>
            </>
          )}

          <div className="flex gap-2">
            <button onClick={submitNote} className="text-[11.5px] rounded-md px-3 py-1.5 font-medium" style={{ background: "#6B4FA0", color: "#F3EFFC" }}>Save Note</button>
            <button onClick={() => { setPickerOpen(false); setBrowseCampus(""); setBrowseProjectId(""); setBrowseSubtaskId(""); setNoteText(""); setAssignTo(""); }} className="text-[11.5px] text-[#6B6980] px-2">Cancel</button>
          </div>
        </div>
      )}

      {activeThreadIds.length === 0 && <div className="text-[11px] text-[#8B889C] py-3 text-center">No central notes yet.</div>}

      <div className="space-y-3">
        {activeThreadIds.map((key) => {
          const thread = centralThreads[key] || { tags: [], messages: [] };
          return (
            <div key={key} className="rounded-md p-3" style={{ background: "#F3EFFC", border: "1px solid #D9CFF0" }}>
              <div className="flex items-center justify-between mb-1.5 gap-2">
                <span className="text-[12.5px] font-medium truncate">{threadLabel(key)}</span>
              </div>
              {thread.tags.length > 0 && (
                <div className="flex flex-wrap gap-1.5 mb-2">
                  {thread.tags.map((t) => (
                    <span key={t} className="text-[10px] px-2 py-0.5 rounded-full flex items-center gap-1" style={{ background: "#6B4FA033", color: "#6B4FA0" }}>
                      #{t}
                      <button onClick={() => onRemoveTag(key, t)} className="hover:text-[#2A2A3A]"><X size={9} /></button>
                    </span>
                  ))}
                </div>
              )}
              <div className="space-y-1.5 mb-2 max-h-[140px] overflow-y-auto">
                {thread.messages.map((m, i) => (
                  <div key={i} className="text-[11.5px]">
                    <span className="font-medium">{m.author}</span> <span className="text-[#6B6980] text-[10px]">{m.ts}</span>
                    {m.assignedTo && <span className="text-[10px] px-1.5 py-0.5 rounded-full ml-1" style={{ background: "#B8862F22", color: "#B8862F" }}>→ {m.assignedTo}</span>}
                    <div>{m.text}</div>
                  </div>
                ))}
                {thread.messages.length === 0 && <div className="text-[10.5px] text-[#8B889C]">No discussion yet.</div>}
              </div>
              <div className="flex gap-2 flex-wrap">
                <input value={messageDrafts[key] || ""} onChange={(e) => setMessageDrafts((prev) => ({ ...prev, [key]: e.target.value }))}
                  onKeyDown={(e) => e.key === "Enter" && sendMessage(key)} placeholder="Add to the central discussion…"
                  className="flex-1 min-w-[140px] bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2.5 py-1.5 text-[11.5px] outline-none" />
                <select value={assignDrafts[key] || ""} onChange={(e) => setAssignDrafts((prev) => ({ ...prev, [key]: e.target.value }))}
                  className="bg-[#EFEAFB] border border-[#D9CFF0] rounded-md px-2 py-1.5 text-[11px] outline-none">
                  <option value="">No assignment</option>
                  {CENTRAL_TEAM.filter((n) => n !== currentViewerName).map((n) => <option key={n} value={n}>→ {n}</option>)}
                </select>
                <button onClick={() => sendMessage(key)} className="text-[11px] px-3 rounded-md font-medium" style={{ background: "#6B4FA0", color: "#F3EFFC" }}>Send</button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function CentralOverview({ campuses, orgBudgetUsed, orgBudgetTotal, projects, staffByCampus, onSelectCampus, centralThreads, onAddTag, onRemoveTag, onAddMessage, currentViewerName, onOpenProject, detail, setDetail }) {
  const sharedProjects = projects.filter((p) => p.shared);
  return (
    <div>
      <div className="mb-6">
        <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(20px,4vw,26px)] font-semibold tracking-tight">All Campuses</h1>
        <p className="text-[13px] text-[#6B6980] mt-1">Organization-wide standing, at a glance.</p>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-7">
        <SummaryCard onClick={() => setDetail("projects")} icon={ListChecks} label="Open Projects" value={projects.filter((p) => p.stage !== "Completed").length} sub="org-wide" color="#2B4C7E" />
        <SummaryCard onClick={() => setDetail("budget")} icon={DollarSign} label="Budget Used" value={`${Math.round((orgBudgetUsed / orgBudgetTotal) * 100)}%`} sub={`${fmtMoney(orgBudgetUsed)} of ${fmtMoney(orgBudgetTotal)}`} color="#B8862F" />
        <SummaryCard onClick={() => setDetail("campuses")} icon={Building2} label="Campuses" value={campuses.length} sub="active sites" color="#5E9E8A" />
        <SummaryCard onClick={() => setDetail("shared")} icon={Link2} label="Shared Projects" value={sharedProjects.length} sub="cross-campus" color="#6B4FA0" />
      </div>

      {detail && (
        <SummaryDetailModal
          type={detail} onClose={() => setDetail(null)}
          campuses={campuses} projects={projects} sharedProjects={sharedProjects} staffByCampus={staffByCampus}
          orgBudgetUsed={orgBudgetUsed} orgBudgetTotal={orgBudgetTotal}
          onOpenProject={(id) => { setDetail(null); onOpenProject(id); }}
          onSelectCampus={(id) => { setDetail(null); onSelectCampus(id); }}
        />
      )}

      <CentralTeamWindow projects={projects} campuses={campuses} centralThreads={centralThreads} onAddTag={onAddTag} onRemoveTag={onRemoveTag}
        onAddMessage={onAddMessage} currentViewerName={currentViewerName} />

      <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3 mb-8 mt-7">
        {campuses.map((c) => {
          const campusProjects = projects.filter((p) => p.stage !== "Completed" && (p.campus === c.id || (p.shared && (p.sharedWith?.includes(c.id) || p.sharedWith?.includes("all")))));
          const overdue = campusProjects.filter((p) => p.due < TODAY_STR).length;
          const cb = rollupBudget(projects.filter((p) => p.campus === c.id));
          const cbPct = cb.total ? Math.round((cb.spent / cb.total) * 100) : 0;
          const staffCount = (staffByCampus[c.id] || []).length;
          return (
            <button key={c.id} onClick={() => onSelectCampus(c.id)} className="text-left bg-[#FFFFFF] rounded-lg p-4 transition group"
              style={{ border: `1.5px solid ${c.color}55` }}
              onMouseEnter={(e) => e.currentTarget.style.borderColor = c.color}
              onMouseLeave={(e) => e.currentTarget.style.borderColor = `${c.color}55`}>
              <div className="flex items-center justify-between mb-3">
                <div className="min-w-0">
                  <div className="text-[15px] font-medium truncate flex items-center gap-1.5" style={{ color: c.color }}>{c.name} <span className="text-[10px] font-normal px-1.5 py-0.5 rounded" style={{ background: `${c.color}22` }}>{c.abbr}</span></div>
                  <div className="text-[11.5px] text-[#6B6980] truncate">OD: {c.lead} · Phase {c.phase}</div>
                </div>
                <div className="text-right shrink-0 pl-2">
                  <div className="text-[16px] font-medium" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{campusProjects.length}</div>
                  <div className="text-[9.5px] text-[#6B6980]">open{overdue > 0 ? ` · ${overdue} overdue` : ""}</div>
                </div>
              </div>
              <div className="mb-1.5">
                <div className="flex justify-between text-[11px] text-[#6B6980] mb-1"><span>Budget</span><span>{cbPct}%</span></div>
                <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden"><div className="h-full rounded-full" style={{ width: `${cbPct}%`, background: c.color }} /></div>
              </div>
              <div className="flex gap-3 text-[11px] text-[#6B6980] mt-2.5"><span>{staffCount} staff</span></div>
            </button>
          );
        })}
      </div>
      <h2 className="text-[14px] font-medium mb-3 flex items-center gap-1.5"><Link2 size={14} /> Cross-Campus & Central Projects</h2>
      <div className="space-y-2">
        {sharedProjects.map((p) => (
          <div key={p.id} className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg px-4 py-3 flex items-center justify-between">
            <div>
              <div className="text-[13.5px]">{p.title}</div>
              <div className="text-[11px] text-[#6B6980] mt-0.5">Owner: {p.owner} · Due {p.due} · Involves {p.sharedWith?.includes("all") ? "all campuses" : p.sharedWith?.map((id) => campuses.find((c) => c.id === id)?.name).join(", ")}</div>
            </div>
            <StageBadge stage={p.stage} />
          </div>
        ))}
      </div>
    </div>
  );
}

function SummaryCard({ icon: Icon, label, value, sub, color, onClick }) {
  const c = color || "#2A2A3A";
  const Tag = onClick ? "button" : "div";
  return (
    <Tag onClick={onClick} className="rounded-lg p-3.5 text-left w-full transition hover:brightness-95 active:scale-[0.98]" style={{ background: `${c}18`, border: `1px solid ${c}55` }}>
      <div className="flex items-center gap-1.5 text-[clamp(10px,2vw,11px)] mb-2" style={{ color: c }}><Icon size={13} className="shrink-0" /> <span className="truncate">{label}</span></div>
      <div style={{ color: c, fontFamily: "'JetBrains Mono', monospace" }} className="text-[clamp(17px,4.2vw,22px)] font-semibold leading-none truncate">{value}</div>
      <div className="text-[clamp(9.5px,1.8vw,10.5px)] mt-1.5 truncate" style={{ color: `${c}CC` }}>{sub}</div>
    </Tag>
  );
}

function CampusDashboard({ campus, projects, staff, notes, activity, dayEvents, calDate, onOpenProject, onGoTab, onOpenProfile }) {
  return (
    <div>
      <div className="mb-5 flex items-start justify-between">
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif", color: campus.color }} className="text-[clamp(19px,4vw,24px)] font-semibold tracking-tight flex items-center gap-2">{campus.name} <span className="text-[13px] font-normal px-2 py-0.5 rounded-full" style={{ background: `${campus.color}22` }}>{campus.abbr}</span></h1>
          <p className="text-[12.5px] text-[#6B6980] mt-1">Campus Operations Director: {campus.lead} · Phase {campus.phase}</p>
        </div>
        <div className="text-right">
          <div className="text-[20px] font-medium" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{projects.filter((p) => p.stage !== "Completed").length}</div>
          <div className="text-[10px] text-[#6B6980]">open projects</div>
        </div>
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <Window title="Staff & Team" icon={Users} onExpand={() => onGoTab("staff")} accentColor={campus.color}>
          <StaffPanel staff={staff} compact onOpenProfile={onOpenProfile} />
        </Window>

        <Window title="Projects / Tasks" icon={ListChecks} onExpand={() => onGoTab("projects")} accentColor={campus.color}>
          <div className="space-y-2">
            {projects.slice(0, 4).map((p) => (
              <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2.5 hover:border-[#C7C3E0]">
                <div className="flex items-center justify-between mb-1.5 gap-2">
                  <span className="text-[12.5px] flex items-center gap-1.5 min-w-0 truncate">{p.shared && <Link2 size={11} className="text-[#B8862F] shrink-0" />}<span className="truncate">{p.title}</span></span>
                  <span className="shrink-0"><StageBadge stage={p.stage} /></span>
                </div>
                <ProgressBar subtasks={p.subtasks} />
                <div className="text-[10px] text-[#6B6980] mt-1.5">{fmtMoney(projectBudget(p).spent)} of {fmtMoney(projectBudget(p).total)} spent · Due {p.due}</div>
              </button>
            ))}
          </div>
        </Window>

        <Window title="Budget" icon={DollarSign} onExpand={() => onGoTab("budget")} accentColor={campus.color}>
          <CampusBudgetSummary projects={projects.filter((p) => p.campus === campus.id)} accentColor={campus.color} />
        </Window>

        <Window title="Calendar" icon={CalendarDays} onExpand={() => onGoTab("calendar")} accentColor={campus.color}>
          <div className="text-[11px] text-[#6B6980] mb-2">{calDate} · Today</div>
          <div className="space-y-1.5">
            {dayEvents.length === 0 && <div className="text-[11.5px] text-[#8B889C]">Nothing scheduled.</div>}
            {dayEvents.map((e, i) => (
              <div key={i} className="flex items-center gap-2 text-[12px] bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2">
                <Clock size={12} className="text-[#6B6980]" /><span className="text-[#6B6980] w-[74px] shrink-0">{e.time}</span><span>{e.title}</span>
              </div>
            ))}
          </div>
        </Window>

        <Window title="Notes" icon={StickyNote} onExpand={() => onGoTab("notes")} accentColor={campus.color}>
          <div className="space-y-1.5">
            {notes.slice(0, 4).map((n) => (
              <div key={n.id} className="bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2">
                <div className="text-[12px]">{n.text}</div>
                <div className="text-[10px] text-[#8B889C] mt-1">{n.ts}</div>
              </div>
            ))}
            {notes.length === 0 && <div className="text-[11.5px] text-[#8B889C]">No notes yet.</div>}
          </div>
        </Window>
      </div>

      <div className="mt-4">
        <Window title="Activity" icon={ActivityIcon} onExpand={() => onGoTab("activity")} accentColor={campus.color}>
          <ActivityFeed activity={activity.slice(0, 5)} />
        </Window>
      </div>
    </div>
  );
}

function Window({ title, icon: Icon, onExpand, children, accentColor }) {
  return (
    <div className="bg-[#FFFFFF] rounded-lg p-4" style={{ border: `1px solid ${accentColor ? accentColor + "66" : "#E3E1F0"}` }}>
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2 text-[13px] font-bold"><Icon size={14} style={{ color: accentColor || "#B8862F" }} />{title}</div>
        <button onClick={onExpand} className="text-[10.5px] text-[#6B6980] hover:text-[#B8862F] flex items-center gap-0.5">Expand <ChevronRight size={12} /></button>
      </div>
      {children}
    </div>
  );
}

function ProgressBar({ subtasks }) {
  if (!subtasks?.length) return null;
  const pct = Math.round((subtasks.filter((s) => s.done).length / subtasks.length) * 100);
  return (
    <div>
      <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden"><div className="h-full rounded-full bg-[#2B4C7E]" style={{ width: `${pct}%` }} /></div>
      <div className="text-[10px] text-[#6B6980] mt-1">{subtasks.filter((s) => s.done).length}/{subtasks.length} subtasks · {pct}%</div>
    </div>
  );
}

function StaffPanel({ staff, campusLabel, compact, full, campusId, roleOptions, onAdd, onAddTeamRole, onRemove, onUpdateRoles, onSetCalendars, onAddRoleOption, slidesLink, onSetSlidesLink, campusPhase, onCommitOrgChart, pendingReassignments, onResolveReassignment, onOpenProfile, users, onLinkUser, onCreateAndLinkUser }) {
  const [addingLane, setAddingLane] = useState(null); // which lane's "Add Team Role" modal is open
  const [linkingId, setLinkingId] = useState(null); // which staff row's login-linker is open
  const [editingId, setEditingId] = useState(null);
  const [calendarPickerId, setCalendarPickerId] = useState(null);
  const [adding, setAdding] = useState(false);
  const [newName, setNewName] = useState("");
  const [newRoles, setNewRoles] = useState([]);
  const [editingSlidesLink, setEditingSlidesLink] = useState(false);
  const [slidesDraft, setSlidesDraft] = useState(slidesLink || "");
  const [importStatus, setImportStatus] = useState("idle"); // idle | loading | review | error
  const [importError, setImportError] = useState("");
  const [importRows, setImportRows] = useState([]); // { name, role, reportsTo }

  const submitAdd = () => {
    if (!newName.trim() || !campusId) return;
    onAdd(campusId, newName.trim(), newRoles);
    setNewName(""); setNewRoles([]); setAdding(false);
  };

  const saveSlidesLink = () => {
    onSetSlidesLink(slidesDraft.trim());
    setEditingSlidesLink(false);
  };

  const runImport = async () => {
    setImportStatus("loading");
    setImportError("");
    try {
      const result = await apiImportOrgChartFromSlides(campusId, slidesLink, campusPhase);
      setImportRows(result.people.map((p) => ({ name: p.name || "", role: p.role || "", reportsTo: p.reportsTo || "" })));
      setImportStatus("review");
    } catch (e) {
      setImportError(e?.message || String(e));
      setImportStatus("error");
    }
  };

  const updateImportRow = (i, field, value) => setImportRows((prev) => prev.map((r, idx) => idx === i ? { ...r, [field]: value } : r));
  const removeImportRow = (i) => setImportRows((prev) => prev.filter((_, idx) => idx !== i));

  const commitImport = () => {
    const clean = importRows.filter((r) => r.name.trim() && r.role.trim());
    if (clean.length === 0) return;
    onCommitOrgChart(campusId, clean);
    setImportStatus("idle");
    setImportRows([]);
  };

  return (
    <div>
      {full && (
        <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
          <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight">Staff & Team — {campusLabel}</h1>
          {campusId && (
            <button onClick={() => setAdding((v) => !v)} className="flex items-center gap-1.5 text-[12px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium">
              <Plus size={14} /> Add Team Member
            </button>
          )}
        </div>
      )}

      {full && campusId && (
        <div className="rounded-lg p-3 mb-4" style={{ background: "#EFEEFA", border: "1px solid #D8D5EC" }}>
          <div className="flex items-center justify-between mb-1.5 flex-wrap gap-1.5">
            <div className="text-[11.5px] font-medium text-[#2A2A3A]">Campus Org Chart {campusPhase ? `— Phase ${campusPhase}` : ""} (Google Slides)</div>
            {!editingSlidesLink && (
              <button onClick={() => { setSlidesDraft(slidesLink || ""); setEditingSlidesLink(true); }} className="text-[10.5px] text-[#B8862F]">
                {slidesLink ? "Change link" : "Add link"}
              </button>
            )}
          </div>
          {editingSlidesLink ? (
            <div className="flex gap-2 flex-wrap">
              <input value={slidesDraft} onChange={(e) => setSlidesDraft(e.target.value)} placeholder="Paste the Google Slides share link…"
                className="flex-1 min-w-[200px] bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[11.5px] outline-none focus:border-[#B8862F]" />
              <button onClick={saveSlidesLink} className="text-[11px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-2.5 py-1.5 font-medium">Save</button>
              <button onClick={() => setEditingSlidesLink(false)} className="text-[11px] text-[#6B6980] px-2">Cancel</button>
            </div>
          ) : slidesLink ? (
            <div className="flex items-center gap-3 flex-wrap">
              <a href={slidesLink} target="_blank" rel="noreferrer" className="text-[11.5px] flex items-center gap-1" style={{ color: "#2B4C7E" }}>
                <Link2 size={12} /> View current phase, roles & titles deck
              </a>
              {importStatus !== "review" && (
                <button onClick={runImport} disabled={importStatus === "loading"} className="text-[11px] rounded-md px-2.5 py-1 font-medium" style={{ background: "#B8862F", color: "#F7F6FB", opacity: importStatus === "loading" ? 0.7 : 1 }}>
                  {importStatus === "loading" ? "Reading slide…" : "Import from this deck"}
                </button>
              )}
            </div>
          ) : (
            <div className="text-[11px] text-[#8B889C]">No deck linked yet — add the Slides link showing this campus's current phase, staff roles, and titles.</div>
          )}

          {importStatus === "error" && (
            <div className="text-[10.5px] mt-2" style={{ color: "#B5462E" }}>{importError}</div>
          )}

          {importStatus === "review" && (
            <div className="mt-3">
              <div className="text-[10.5px] text-[#6B6980] mb-2">Read from Phase {campusPhase} of the deck — review before committing to the roster.</div>
              <div className="space-y-1.5 mb-2">
                {importRows.map((r, i) => (
                  <div key={i} className="grid grid-cols-[1fr_1fr_1fr_auto] gap-1.5">
                    <input value={r.name} onChange={(e) => updateImportRow(i, "name", e.target.value)} placeholder="Name"
                      className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
                    <input value={r.role} onChange={(e) => updateImportRow(i, "role", e.target.value)} placeholder="Role"
                      className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
                    <input value={r.reportsTo} onChange={(e) => updateImportRow(i, "reportsTo", e.target.value)} placeholder="Reports to"
                      className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
                    <button onClick={() => removeImportRow(i)} className="text-[#8B889C] hover:text-[#C15B5B] px-1"><Trash2 size={13} /></button>
                  </div>
                ))}
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={commitImport} className="text-[11px] rounded-md px-3 py-1.5 font-medium" style={{ background: "#B8862F", color: "#F7F6FB" }}>Commit to Roster</button>
                <button onClick={() => { setImportStatus("idle"); setImportRows([]); }} className="text-[11px] text-[#6B6980] px-2 py-1.5">Cancel</button>
              </div>
            </div>
          )}
        </div>
      )}

      {full && campusId && pendingReassignments?.length > 0 && (
        <div className="rounded-lg p-3 mb-4" style={{ background: "#B8862F18", border: "1px solid #B8862F55" }}>
          <div className="text-[11.5px] font-medium mb-2" style={{ color: "#8A6420" }}>Org chart update — confirm role changes</div>
          <div className="space-y-2">
            {pendingReassignments.map((r) => (
              <div key={r.id} className="bg-[#FFFFFF] rounded-md p-2.5 flex items-center justify-between flex-wrap gap-2" style={{ border: "1px solid #E3E1F0" }}>
                <div className="text-[12px]">
                  <span className="font-medium">{r.role}</span> — was <span className="line-through text-[#8B889C]">{r.oldName}</span>, now <span className="font-medium">{r.newName}</span>
                </div>
                <div className="flex gap-2 shrink-0">
                  <button onClick={() => onResolveReassignment(r.id, true)} className="text-[11px] rounded-md px-2.5 py-1 font-medium" style={{ background: "#B8862F", color: "#F7F6FB" }}>Reassign</button>
                  <button onClick={() => onResolveReassignment(r.id, false)} className="text-[11px] text-[#6B6980] px-2 py-1">Not now</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {full && campusId && (
        <div className="mb-5">
          <div className="text-[11.5px] font-medium text-[#2A2A3A] mb-2">Team Lanes</div>
          <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {TEAM_LANES.map((lane) => {
              const laneStaff = staff.filter((s) => s.lane === lane);
              return (
                <div key={lane} className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg p-3">
                  <div className="flex items-center justify-between mb-2">
                    <div className="text-[12px] font-medium">{lane}</div>
                    <button onClick={() => setAddingLane(lane)} className="text-[#B8862F] hover:text-[#8A6420]"><Plus size={14} /></button>
                  </div>
                  {laneStaff.length === 0 && <div className="text-[10.5px] text-[#8B889C]">No one yet.</div>}
                  <div className="space-y-1.5">
                    {laneStaff.map((s) => (
                      <div key={s.id} className="text-[11.5px] rounded-md px-2 py-1.5" style={{ background: "#F7F6FB" }}>
                        <div className="font-medium">{s.name}</div>
                        <div className="text-[10.5px] text-[#6B6980]">{(s.roles || []).join(", ")}</div>
                        {s.reportsTo && <div className="text-[10px] text-[#8B889C]">→ reports to {s.reportsTo}</div>}
                      </div>
                    ))}
                  </div>
                  <button onClick={() => setAddingLane(lane)} className="flex items-center gap-1 text-[10.5px] mt-2" style={{ color: "#B8862F" }}>
                    <Plus size={11} /> Add Team Role
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {addingLane && (
        <AddTeamRoleModal
          lane={addingLane}
          roleOptions={LANE_ROLE_OPTIONS[addingLane] || []}
          reportsToOptions={staff}
          onClose={() => setAddingLane(null)}
          onCreate={(data) => { onAddTeamRole(campusId, addingLane, data); setAddingLane(null); }}
        />
      )}

      {full && adding && (
        <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg p-3 mb-4 max-w-[420px]">
          <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Full name"
            className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-3 py-2 text-[12.5px] outline-none focus:border-[#B8862F] mb-2" />
          <div className="text-[10.5px] text-[#6B6980] mb-1.5">Role(s) — up to 2</div>
          <RolePicker selected={newRoles} onChange={setNewRoles} roleOptions={roleOptions} onAddRoleOption={onAddRoleOption} />
          <div className="flex gap-2 mt-3">
            <button onClick={submitAdd} className="text-[12px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium">Save</button>
            <button onClick={() => { setAdding(false); setNewName(""); setNewRoles([]); }} className="text-[12px] text-[#6B6980] px-3 py-1.5">Cancel</button>
          </div>
        </div>
      )}

      <div className="space-y-2">
        {staff.map((s) => (
          <div key={s.id} className={`bg-[#EFEEFA] border border-[#E3E1F0] rounded-md ${compact ? "px-3 py-2" : "px-4 py-3"}`}>
            <div className="flex items-center justify-between flex-wrap gap-y-1.5">
              <div className="flex-1 min-w-0 pr-2">
                <button onClick={() => onOpenProfile && onOpenProfile(s.name)} className="text-[13px] truncate text-left hover:underline hover:text-[#B8862F] block">{s.name}</button>
                <div className="text-[11px] text-[#6B6980] truncate">{(s.roles || []).join(" · ") || "No role set"}</div>
              </div>
              <div className="flex items-center gap-1.5 flex-wrap justify-end">
                {s.flag && <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#C15B5B22] text-[#C15B5B] flex items-center gap-1 whitespace-nowrap"><AlertTriangle size={10} />{s.flag}</span>}
                {full && campusId && (
                  <>
                    {(() => {
                      const linkedUser = s.userId ? users?.find((u) => String(u.id) === String(s.userId)) : null;
                      return linkedUser ? (
                        <span className="text-[9.5px] px-2 py-0.5 rounded-full border border-[#5E9E8A] text-[#5E9E8A] whitespace-nowrap flex items-center gap-1">
                          <ShieldCheck size={10} /> {linkedUser.email}
                        </span>
                      ) : (
                        <button onClick={() => setLinkingId(linkingId === s.id ? null : s.id)} className="text-[9.5px] px-2 py-0.5 rounded-full border border-[#D8D5EC] text-[#8B889C] hover:border-[#B8862F] hover:text-[#B8862F] whitespace-nowrap">
                          Link Login
                        </button>
                      );
                    })()}
                    <button onClick={() => setCalendarPickerId(calendarPickerId === s.id ? null : s.id)}
                      className={`text-[9.5px] px-2 py-0.5 rounded-full border whitespace-nowrap ${s.calendarSynced ? "border-[#5E9E8A] text-[#5E9E8A]" : "border-[#D8D5EC] text-[#8B889C]"}`}>
                      {s.calendarSynced ? `${s.calendars?.length || 0} cal${s.calendars?.length === 1 ? "" : "s"} synced` : "Sync Calendar"}
                    </button>
                    <button onClick={() => setEditingId(editingId === s.id ? null : s.id)} className="text-[10.5px] text-[#6B6980] hover:text-[#B8862F] whitespace-nowrap">Edit</button>
                    <button onClick={() => onRemove(campusId, s.id)} className="text-[#8B889C] hover:text-[#C15B5B] shrink-0"><Trash2 size={13} /></button>
                  </>
                )}
              </div>
            </div>
            <div className="flex gap-4 text-[10.5px] text-[#6B6980] mt-2 flex-wrap">
              <span>Next: {s.nextMeeting}</span>
              <span>Last contact: {s.lastContact}</span>
              {s.reportsTo && <span>Reports to: {s.reportsTo}</span>}
            </div>
            {full && editingId === s.id && (
              <div className="mt-3 pt-3 border-t border-[#E3E1F0]">
                <div className="text-[10.5px] text-[#6B6980] mb-1.5">Role(s) — up to 2</div>
                <RolePicker selected={s.roles || []} onChange={(roles) => onUpdateRoles(campusId, s.id, roles)} roleOptions={roleOptions} onAddRoleOption={onAddRoleOption} />
              </div>
            )}
            {full && calendarPickerId === s.id && (
              <div className="mt-3 pt-3 border-t border-[#E3E1F0]">
                <CalendarSyncPicker staffer={s} onSave={(cals) => { onSetCalendars(campusId, s.id, cals); setCalendarPickerId(null); }} />
              </div>
            )}
            {full && linkingId === s.id && (
              <div className="mt-3 pt-3 border-t border-[#E3E1F0]">
                <StaffLoginLinker staffPerson={s} campusId={campusId} users={users}
                  onLink={(userId) => { onLinkUser(campusId, s.id, userId); setLinkingId(null); }}
                  onCreateAndLink={(data) => { onCreateAndLinkUser(campusId, s.id, data); setLinkingId(null); }}
                  onCancel={() => setLinkingId(null)} />
              </div>
            )}
          </div>
        ))}
        {staff.length === 0 && <div className="text-[11.5px] text-[#8B889C] py-6 text-center">No team members in view.</div>}
      </div>
    </div>
  );
}

function AddTeamRoleModal({ lane, roleOptions, reportsToOptions, onClose, onCreate }) {
  useLockBodyScroll();
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [email, setEmail] = useState("");
  const [phone, setPhone] = useState("");
  const [role, setRole] = useState(roleOptions[0] || "");
  const [reportsTo, setReportsTo] = useState("");

  const inputClass = "w-full bg-[#F7F6FB] border border-[#D8D5EC] rounded-md px-3 py-2 text-[13px] outline-none focus:border-[#2B4C7E]";
  const labelClass = "text-[11px] font-medium text-[#6B6980] mb-1 block";

  const submit = (e) => {
    e.preventDefault();
    if (!firstName.trim() || !lastName.trim()) return;
    onCreate({ firstName: firstName.trim(), lastName: lastName.trim(), email: email.trim(), phone: phone.trim(), role, reportsTo: reportsTo || null });
  };

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center p-4" style={{ background: "rgba(42,42,58,0.45)" }}>
      <div className="bg-[#FFFFFF] rounded-xl p-6 w-full max-w-[420px] max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-[15px] font-semibold">Add Team Role — {lane}</h2>
          <button onClick={onClose} className="text-[#8B889C] hover:text-[#2A2A3A]"><X size={16} /></button>
        </div>
        <form onSubmit={submit} className="space-y-3">
          <div className="grid grid-cols-2 gap-2">
            <div><label className={labelClass}>First name</label><input required value={firstName} onChange={(e) => setFirstName(e.target.value)} className={inputClass} /></div>
            <div><label className={labelClass}>Last name</label><input required value={lastName} onChange={(e) => setLastName(e.target.value)} className={inputClass} /></div>
          </div>
          <div><label className={labelClass}>Email address</label><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} className={inputClass} /></div>
          <div><label className={labelClass}>Phone number</label><input type="tel" value={phone} onChange={(e) => setPhone(e.target.value)} className={inputClass} /></div>
          <div>
            <label className={labelClass}>Role / Title</label>
            <select value={role} onChange={(e) => setRole(e.target.value)} className={inputClass}>
              {roleOptions.map((r) => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div>
            <label className={labelClass}>Reports To (optional)</label>
            <select value={reportsTo} onChange={(e) => setReportsTo(e.target.value)} className={inputClass}>
              <option value="">— None —</option>
              {reportsToOptions.map((s) => <option key={s.id} value={s.name}>{s.name} — {(s.roles || []).join(", ") || "no role set"}</option>)}
            </select>
          </div>
          <div className="flex gap-2 pt-1">
            <button type="submit" className="text-[13px] font-medium rounded-md px-4 py-2" style={{ background: "#2B4C7E", color: "#F7F6FB" }}>Add to {lane}</button>
            <button type="button" onClick={onClose} className="text-[13px] text-[#6B6980] px-2">Cancel</button>
          </div>
        </form>
      </div>
    </div>
  );
}

function RolePicker({ selected, onChange, roleOptions, onAddRoleOption }) {
  const [customRole, setCustomRole] = useState("");
  const toggle = (r) => {
    if (selected.includes(r)) onChange(selected.filter((x) => x !== r));
    else if (selected.length < 2) onChange([...selected, r]);
  };
  const addCustom = () => {
    if (!customRole.trim()) return;
    onAddRoleOption(customRole.trim());
    if (selected.length < 2) onChange([...selected, customRole.trim()]);
    setCustomRole("");
  };
  return (
    <div>
      <div className="flex flex-wrap gap-1.5 mb-2">
        {roleOptions.map((r) => (
          <button key={r} onClick={() => toggle(r)}
            disabled={!selected.includes(r) && selected.length >= 2}
            className={`text-[10.5px] px-2 py-1 rounded-full border ${selected.includes(r) ? "bg-[#B8862F] text-[#F7F6FB] border-[#B8862F]" : "border-[#D8D5EC] text-[#6B6980]"} ${!selected.includes(r) && selected.length >= 2 ? "opacity-40" : ""}`}>
            {r}
          </button>
        ))}
      </div>
      <div className="flex gap-1.5">
        <input value={customRole} onChange={(e) => setCustomRole(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCustom()}
          placeholder="Add a new role…" className="flex-1 bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[11px] outline-none focus:border-[#B8862F]" />
        <button onClick={addCustom} className="text-[10.5px] text-[#B8862F] px-2">Add</button>
      </div>
      {selected.length >= 2 && <div className="text-[10px] text-[#8B889C] mt-1.5">Max 2 roles selected.</div>}
    </div>
  );
}

function CalendarSyncPicker({ staffer, onSave }) {
  const [connected, setConnected] = useState(!!staffer.calendarSynced);
  const [selected, setSelected] = useState(staffer.calendars || []);

  const toggleCal = (cal) =>
    setSelected((prev) => prev.includes(cal) ? prev.filter((c) => c !== cal) : [...prev, cal]);

  if (!connected) {
    return (
      <div>
        <div className="text-[10.5px] text-[#6B6980] mb-2">Connect this person's Google Workspace account to select which calendars sync into the dashboard.</div>
        <button onClick={() => setConnected(true)} className="text-[11.5px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium">
          Connect Google Workspace
        </button>
      </div>
    );
  }
  return (
    <div>
      <div className="text-[10.5px] text-[#6B6980] mb-2">Calendars available through this Google Workspace login — select any number:</div>
      <div className="space-y-1.5 mb-3">
        {MOCK_GOOGLE_CALENDARS.map((cal) => (
          <label key={cal} className="flex items-center gap-2 text-[12px] cursor-pointer">
            <input type="checkbox" checked={selected.includes(cal)} onChange={() => toggleCal(cal)} className="accent-[#B8862F]" />
            {cal}
          </label>
        ))}
      </div>
      <button onClick={() => onSave(selected)} className="text-[11.5px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium">Save Selection</button>
    </div>
  );
}

// Lets a Staff & Team roster entry (directory info, no login) get connected to a real Users
// login — either an existing account (an unassigned one waiting to be claimed, or one already
// scoped to this campus) or a brand-new one created and linked in the same step.
function StaffLoginLinker({ staffPerson, campusId, users, onLink, onCreateAndLink, onCancel }) {
  const [mode, setMode] = useState("pick"); // "pick" | "create"
  const [selectedUserId, setSelectedUserId] = useState("");
  const nameParts = (staffPerson.name || "").split(" ");
  const [firstName, setFirstName] = useState(nameParts[0] || "");
  const [lastName, setLastName] = useState(nameParts.slice(1).join(" "));
  const [email, setEmail] = useState(staffPerson.email || "");
  const [phone, setPhone] = useState(staffPerson.phone || "");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const eligibleUsers = (users || []).filter((u) => u.campusId === campusId || u.tier === "unassigned");

  const submitCreate = (e) => {
    e.preventDefault();
    if (!firstName.trim() || !lastName.trim() || !email.trim()) { setError("First name, last name, and email are required."); return; }
    if (password.length < 8) { setError("Password must be at least 8 characters."); return; }
    onCreateAndLink({ firstName: firstName.trim(), lastName: lastName.trim(), email: email.trim(), phone: phone.trim(), password });
  };

  return (
    <div className="rounded-md p-3" style={{ background: "#F7F6FB", border: "1px solid #D8D5EC" }}>
      <div className="flex items-center gap-2 mb-2">
        <button onClick={() => setMode("pick")} className={`text-[10.5px] px-2 py-1 rounded-md ${mode === "pick" ? "bg-[#2B4C7E] text-[#F7F6FB]" : "text-[#6B6980]"}`}>Link existing account</button>
        <button onClick={() => setMode("create")} className={`text-[10.5px] px-2 py-1 rounded-md ${mode === "create" ? "bg-[#2B4C7E] text-[#F7F6FB]" : "text-[#6B6980]"}`}>Create new login</button>
      </div>

      {mode === "pick" ? (
        <div className="flex gap-2 flex-wrap">
          <select value={selectedUserId} onChange={(e) => setSelectedUserId(e.target.value)}
            className="flex-1 min-w-[180px] bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none">
            <option value="">Select an account…</option>
            {eligibleUsers.map((u) => (
              <option key={u.id} value={u.id}>{u.firstName} {u.lastName} — {u.email}{u.tier === "unassigned" ? " (unassigned)" : ""}</option>
            ))}
          </select>
          <button disabled={!selectedUserId} onClick={() => onLink(selectedUserId)} className="text-[11px] rounded-md px-3 py-1.5 font-medium disabled:opacity-50" style={{ background: "#2B4C7E", color: "#F7F6FB" }}>Link</button>
          <button onClick={onCancel} className="text-[11px] text-[#6B6980] px-2">Cancel</button>
          {eligibleUsers.length === 0 && <div className="text-[10.5px] text-[#8B889C] w-full">No unlinked accounts found for this campus — try "Create new login" instead.</div>}
        </div>
      ) : (
        <form onSubmit={submitCreate} className="space-y-2">
          <div className="grid grid-cols-2 gap-2">
            <input required value={firstName} onChange={(e) => setFirstName(e.target.value)} placeholder="First name" className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none" />
            <input required value={lastName} onChange={(e) => setLastName(e.target.value)} placeholder="Last name" className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none" />
          </div>
          <input required type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="Email address" className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none" />
          <input type="tel" value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Phone (optional)" className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none" />
          <input required type="password" minLength={8} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Temporary password" className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none" />
          {error && <div className="text-[10.5px]" style={{ color: "#C15B5B" }}>{error}</div>}
          <div className="flex gap-2">
            <button type="submit" className="text-[11px] rounded-md px-3 py-1.5 font-medium" style={{ background: "#2B4C7E", color: "#F7F6FB" }}>Create & Link</button>
            <button type="button" onClick={onCancel} className="text-[11px] text-[#6B6980] px-2">Cancel</button>
          </div>
        </form>
      )}
    </div>
  );
}

function calToDate(s) { return new Date(s + "T00:00:00"); }
function calFmtISO(d) { return d.toISOString().slice(0, 10); }
function calStartOfWeek(dateStr) {
  const d = calToDate(dateStr);
  d.setDate(d.getDate() - d.getDay());
  return d;
}
function calWeekDates(dateStr) {
  const start = calStartOfWeek(dateStr);
  return Array.from({ length: 7 }, (_, i) => { const d = new Date(start); d.setDate(start.getDate() + i); return calFmtISO(d); });
}
function calMonthGrid(dateStr) {
  const d = calToDate(dateStr);
  const year = d.getFullYear(), month = d.getMonth();
  const firstOfMonth = new Date(year, month, 1);
  const startGrid = new Date(firstOfMonth);
  startGrid.setDate(startGrid.getDate() - firstOfMonth.getDay());
  const days = [];
  for (let i = 0; i < 42; i++) { const dd = new Date(startGrid); dd.setDate(startGrid.getDate() + i); days.push(calFmtISO(dd)); }
  return { days, month, year, monthLabel: firstOfMonth.toLocaleDateString(undefined, { month: "long", year: "numeric" }) };
}
function calDatesInRange(startStr, endStr) {
  if (!startStr || !endStr) return [];
  const start = calToDate(startStr), end = calToDate(endStr);
  const days = [];
  let cur = new Date(start);
  let guard = 0;
  while (cur <= end && guard < 366) { days.push(calFmtISO(cur)); cur.setDate(cur.getDate() + 1); guard++; }
  return days;
}

function CalendarPanel({ calDate, setCalDate, calView, setCalView, dayEvents, full, authUser, onSaveGoogleCalendars, onSetGoogleConnected }) {
  const [searchQuery, setSearchQuery] = useState("");
  const [customStart, setCustomStart] = useState(calDate);
  const [customEnd, setCustomEnd] = useState(calDate);

  // Real, permanent Google Calendar connection — see requestGoogleAuthCode and the
  // apiConnectGoogleCalendar/apiListGoogleCalendars/apiListGoogleEvents helpers above.
  // googleEventsByDate is null until events have actually been fetched; until then every view
  // below falls back to the seedCalendar demo data, same as before. Once authUser.
  // googleCalendarConnected is true, this loads automatically on mount — no popup, no click.
  const [connecting, setConnecting] = useState(false);
  const [calError, setCalError] = useState("");
  const [availableCalendars, setAvailableCalendars] = useState(null); // fetched list, once connected
  const [selectedIds, setSelectedIds] = useState(() => authUser?.googleCalendarIds || []);
  const [showPicker, setShowPicker] = useState(false);
  const [googleEventsByDate, setGoogleEventsByDate] = useState(null);
  const [loadingEvents, setLoadingEvents] = useState(false);

  useEffect(() => {
    if (authUser?.googleCalendarConnected && (authUser?.googleCalendarIds || []).length > 0) {
      setLoadingEvents(true);
      apiListGoogleEvents(authUser.googleCalendarIds)
        .then((result) => setGoogleEventsByDate(groupGoogleEventsByDate(result.events || [])))
        .catch((e) => setCalError(e?.message || String(e)))
        .finally(() => setLoadingEvents(false));
    }
    // Only ever needs to run once per mount — reconnecting/reselecting is handled by their
    // own explicit button clicks below, not by this effect re-firing.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const connectGoogleCalendar = () => {
    setConnecting(true);
    setCalError("");
    requestGoogleAuthCode(async (resp) => {
      if (resp.error) { setConnecting(false); setCalError(typeof resp.error === "string" ? resp.error : "Couldn't connect to Google Calendar."); return; }
      try {
        await apiConnectGoogleCalendar(resp.code); // backend exchanges + stores the refresh token permanently
        onSetGoogleConnected(true);
        const { calendars } = await apiListGoogleCalendars();
        setAvailableCalendars(calendars);
        setShowPicker(true);
        if ((authUser?.googleCalendarIds || []).length > 0) {
          setLoadingEvents(true);
          const result = await apiListGoogleEvents(authUser.googleCalendarIds);
          setGoogleEventsByDate(groupGoogleEventsByDate(result.events || []));
          setLoadingEvents(false);
        }
      } catch (e) {
        setCalError(e?.message || String(e));
      } finally {
        setConnecting(false);
      }
    });
  };

  const disconnectGoogleCalendar = async () => {
    setCalError("");
    try {
      await apiDisconnectGoogleCalendar();
      onSetGoogleConnected(false);
      setGoogleEventsByDate(null);
      setAvailableCalendars(null);
      setShowPicker(false);
    } catch (e) {
      setCalError(e?.message || String(e));
    }
  };

  const toggleCalendarId = (id) => setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);

  const saveCalendarSelection = async () => {
    const names = selectedIds.map((id) => availableCalendars?.find((c) => c.id === id)?.name || id);
    onSaveGoogleCalendars(selectedIds, names);
    setShowPicker(false);
    if (selectedIds.length > 0) {
      setLoadingEvents(true);
      try {
        const result = await apiListGoogleEvents(selectedIds);
        setGoogleEventsByDate(groupGoogleEventsByDate(result.events || []));
      } catch (e) {
        setCalError(e?.message || String(e));
      }
      setLoadingEvents(false);
    } else {
      setGoogleEventsByDate({});
    }
  };

  const matchesSearch = (e) => {
    if (!searchQuery.trim()) return true;
    const q = searchQuery.toLowerCase();
    return e.title.toLowerCase().includes(q) || (e.attendees || []).some((a) => a.toLowerCase().includes(q));
  };
  const eventsForDate = (d) => (googleEventsByDate ? (googleEventsByDate[d] || []) : (seedCalendar[d] || [])).filter(matchesSearch);
  const filteredDayEvents = (googleEventsByDate ? (googleEventsByDate[calDate] || []) : dayEvents).filter(matchesSearch);

  const weekDates = calWeekDates(calDate);
  const monthGrid = calMonthGrid(calDate);
  const customDates = calDatesInRange(customStart, customEnd);

  return (
    <div>
      {full && (
        <>
          <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
            <div>
              <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight">Calendar</h1>
              {authUser?.googleCalendarConnected ? (
                <p className="text-[11.5px] mt-1 flex items-center gap-1" style={{ color: "#5E9E8A" }}>
                  <Link2 size={11} /> Connected — showing {selectedIds.length} of your Google calendar{selectedIds.length === 1 ? "" : "s"}
                  <button onClick={async () => { const { calendars } = await apiListGoogleCalendars(); setAvailableCalendars(calendars); setShowPicker((v) => !v); }} className="underline ml-1" style={{ color: "#5E9E8A" }}>change</button>
                </p>
              ) : (
                <p className="text-[11.5px] text-[#6B6980] mt-1 flex items-center gap-1">
                  <Link2 size={11} /> Showing sample data — connect your Google Calendar
                </p>
              )}
            </div>
            <div className="flex items-center gap-2">
              {authUser?.googleCalendarConnected ? (
                <button onClick={disconnectGoogleCalendar} className="text-[11.5px] rounded-md px-3 py-1.5 font-medium whitespace-nowrap" style={{ background: "#E3E1F0", color: "#2A2A3A" }}>
                  Disconnect
                </button>
              ) : (
                <button onClick={connectGoogleCalendar} disabled={connecting}
                  className="text-[11.5px] rounded-md px-3 py-1.5 font-medium whitespace-nowrap" style={{ background: "#2B4C7E", color: "#F7F6FB", opacity: connecting ? 0.7 : 1 }}>
                  {connecting ? "Connecting…" : "Connect Google Calendar"}
                </button>
              )}
              <div className="flex gap-1 bg-[#EFEEFA] border border-[#D8D5EC] rounded-md p-0.5">
                {["day", "week", "month", "custom"].map((v) => (
                  <button key={v} onClick={() => setCalView(v)} className={`text-[11px] px-2.5 py-1 rounded ${calView === v ? "bg-[#B8862F] text-[#F7F6FB]" : "text-[#6B6980]"}`}>{v[0].toUpperCase() + v.slice(1)}</button>
                ))}
              </div>
            </div>
          </div>

          {calError && <div className="text-[11px] mb-3 rounded-md px-3 py-2" style={{ background: "#C15B5B1A", color: "#C15B5B" }}>{calError}</div>}
          {loadingEvents && <div className="text-[11px] text-[#6B6980] mb-3">Loading your calendar events…</div>}

          {showPicker && availableCalendars && (
            <div className="rounded-lg p-3 mb-4" style={{ background: "#EFEEFA", border: "1px solid #D8D5EC" }}>
              <div className="text-[11.5px] font-medium mb-2">Choose which of your calendars to show — 1, some, or all</div>
              <div className="space-y-1.5 mb-3">
                {availableCalendars.map((c) => (
                  <label key={c.id} className="flex items-center gap-2 text-[12px] cursor-pointer">
                    <input type="checkbox" checked={selectedIds.includes(c.id)} onChange={() => toggleCalendarId(c.id)} className="accent-[#B8862F]" />
                    {c.name}{c.primary ? " (primary)" : ""}
                  </label>
                ))}
              </div>
              <div className="flex gap-2">
                <button onClick={saveCalendarSelection} className="text-[11.5px] rounded-md px-3 py-1.5 font-medium" style={{ background: "#B8862F", color: "#F7F6FB" }}>Save Selection</button>
                <button onClick={() => setShowPicker(false)} className="text-[11.5px] text-[#6B6980] px-2">Cancel</button>
              </div>
            </div>
          )}

          <div className="mb-4 flex items-center gap-2 flex-wrap">
            <input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="Search by title or person invited…"
              className="flex-1 min-w-[200px] max-w-[360px] bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[11.5px] outline-none focus:border-[#B8862F]" />
            {searchQuery && <button onClick={() => setSearchQuery("")} className="text-[10.5px] text-[#8B889C]">Clear</button>}
            {calView === "custom" && (
              <div className="flex items-center gap-1.5 text-[11px] text-[#6B6980]">
                <input type="date" value={customStart} onChange={(e) => setCustomStart(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />
                <span>to</span>
                <input type="date" value={customEnd} onChange={(e) => setCustomEnd(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />
              </div>
            )}
          </div>
        </>
      )}

      {(!full || calView === "day") && (
        <div className="space-y-1.5 max-w-[480px]">
          {filteredDayEvents.length === 0 && <div className="text-[11.5px] text-[#8B889C]">Nothing scheduled{searchQuery ? " matching your search" : ""}.</div>}
          {filteredDayEvents.map((e, i) => (
            <div key={i} className="text-[12.5px] bg-[#FFFFFF] border border-[#E3E1F0] rounded-md px-3 py-2.5">
              <div className="flex items-center gap-2">
                <Clock size={13} className="text-[#6B6980]" /><span className="text-[#6B6980] w-[74px] shrink-0">{e.time}</span><span>{e.title}</span>
              </div>
              {e.attendees?.length > 0 && <div className="text-[10px] text-[#8B889C] mt-1 pl-[21px]">{e.attendees.join(", ")}</div>}
            </div>
          ))}
        </div>
      )}

      {full && calView === "week" && (
        <div className="grid grid-cols-7 gap-2 items-start">
          {weekDates.map((d) => {
            const evs = eventsForDate(d);
            return (
              <button key={d} onClick={() => setCalDate(d)} className={`text-left border rounded-md p-2 h-[280px] flex flex-col ${d === calDate ? "border-[#B8862F] bg-[#EFEEFA]" : "border-[#E3E1F0] bg-[#FFFFFF]"}`}>
                <div className="text-[10.5px] font-medium text-[#6B6980] mb-1.5 shrink-0">{d.slice(5)}</div>
                <div className="space-y-1 overflow-y-auto min-h-0">
                  {evs.map((e, i) => (
                    <div key={i} title={`${e.time} — ${e.title}`} className="text-[9.5px] bg-[#E3E1F0] rounded px-1.5 py-1">
                      <div className="text-[8.5px] text-[#6B6980] font-medium leading-tight">{e.time}</div>
                      <div className="truncate leading-tight">{e.title}</div>
                    </div>
                  ))}
                  {evs.length === 0 && <div className="text-[9px] text-[#B8B5C9]">Nothing scheduled</div>}
                </div>
              </button>
            );
          })}
        </div>
      )}

      {full && calView === "month" && (
        <div>
          <div className="text-[12.5px] font-medium mb-2">{monthGrid.monthLabel}</div>
          <div className="grid grid-cols-7 gap-1.5 text-center text-[9.5px] text-[#8B889C] mb-1">
            {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d) => <div key={d}>{d}</div>)}
          </div>
          <div className="grid grid-cols-7 gap-1.5 items-start">
            {monthGrid.days.map((d) => {
              const inMonth = calToDate(d).getMonth() === monthGrid.month;
              const evs = eventsForDate(d);
              return (
                <button key={d} onClick={() => setCalDate(d)}
                  className={`text-left border rounded-md p-1.5 h-[92px] flex flex-col ${d === calDate ? "border-[#B8862F] bg-[#EFEEFA]" : "border-[#E3E1F0] bg-[#FFFFFF]"} ${inMonth ? "" : "opacity-40"}`}>
                  <div className="text-[9.5px] text-[#8B889C] shrink-0">{Number(d.slice(8, 10))}</div>
                  <div className="overflow-y-auto min-h-0">
                    {evs.slice(0, 3).map((e, i) => (
                      <div key={i} title={`${e.time} — ${e.title}`} className="text-[8.5px] bg-[#E3E1F0] rounded px-1 py-0.5 mt-0.5 truncate">{e.title}</div>
                    ))}
                    {evs.length > 3 && <div className="text-[8px] text-[#8B889C] mt-0.5">+{evs.length - 3} more</div>}
                  </div>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {full && calView === "custom" && (
        <div className="space-y-1.5">
          {customDates.length === 0 && <div className="text-[11.5px] text-[#8B889C]">Pick a start and end date above.</div>}
          {customDates.map((d) => {
            const evs = eventsForDate(d);
            if (evs.length === 0) return null;
            return (
              <div key={d}>
                <div className="text-[10.5px] text-[#8B889C] mb-1">{calToDate(d).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" })}</div>
                <div className="space-y-1 mb-2">
                  {evs.map((e, i) => (
                    <div key={i} className="text-[12px] bg-[#FFFFFF] border border-[#E3E1F0] rounded-md px-3 py-2 flex items-center gap-2">
                      <Clock size={12} className="text-[#6B6980] shrink-0" /><span className="text-[#6B6980] w-[64px] shrink-0">{e.time}</span><span className="truncate">{e.title}</span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
          {customDates.length > 0 && customDates.every((d) => eventsForDate(d).length === 0) && (
            <div className="text-[11.5px] text-[#8B889C]">Nothing scheduled in this range{searchQuery ? " matching your search" : ""}.</div>
          )}
        </div>
      )}
    </div>
  );
}

function NotesPanel({ notes, newNote, setNewNote, addNote, removeNote, campusLabel, full }) {
  return (
    <div>
      {full && <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight mb-5">Notes — {campusLabel}</h1>}
      <div className="flex gap-2 mb-4 max-w-[520px]">
        <input value={newNote} onChange={(e) => setNewNote(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addNote()}
          placeholder="Quick note — press Enter to save" className="flex-1 bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-3 py-2 text-[12.5px] outline-none focus:border-[#B8862F]" />
        <button onClick={addNote} className="bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 text-[12px] font-medium flex items-center gap-1"><Plus size={13} /> Add</button>
      </div>
      <div className="space-y-1.5 max-w-[520px]">
        {notes.map((n) => (
          <div key={n.id} className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-md px-3 py-2.5 flex items-start justify-between gap-2">
            <div>
              <div className="text-[12.5px]">{n.text}</div>
              <div className="text-[10px] text-[#8B889C] mt-1">{n.ts}</div>
            </div>
            <button onClick={() => removeNote(n.id)} className="text-[#8B889C] hover:text-[#C15B5B] shrink-0"><Trash2 size={13} /></button>
          </div>
        ))}
        {notes.length === 0 && <div className="text-[11.5px] text-[#8B889C] py-6 text-center">No notes yet.</div>}
      </div>
    </div>
  );
}

function ActivityFeed({ activity }) {
  const iconFor = (a) => a.includes("photo") ? ImageIcon : a.includes("transcript") ? Mic : a.includes("note") ? MessageSquare : a.includes("deadline") ? Clock : a.includes("completed") ? CheckCircle2 : ActivityIcon;
  return (
    <div className="space-y-2">
      {activity.map((a, idx) => {
        const Icon = iconFor(a.action);
        return (
          <div key={a.id} className="flex items-start gap-2.5 text-[12px] pl-2 relative">
            {idx === 0 && <span className="absolute left-0 top-1.5 w-1.5 h-1.5 rounded-full animate-pulse" style={{ background: "#5E9E8A" }} />}
            <Icon size={13} className="text-[#B8862F] mt-0.5 shrink-0" />
            <div><span className="font-medium">{a.actor}</span> <span className="text-[#6B6980]">{a.action}</span> <span>{a.target}</span>
              <div className="text-[10px] text-[#8B889C]">{a.ts}</div>
            </div>
          </div>
        );
      })}
      {activity.length === 0 && <div className="text-[11.5px] text-[#8B889C]">No activity yet — this fills in live as people work.</div>}
    </div>
  );
}

function ActivityPanel({ activity, full, projects, currentViewerName }) {
  const todayCompletions = (activity || []).filter((a) => a.date === TODAY_STR && a.createdBy === currentViewerName && (a.itemType === "project" || a.itemType === "task"));
  const weekDates = calWeekDates(TODAY_STR);
  const weekCompletions = (activity || []).filter((a) => weekDates.includes(a.date) && a.createdBy === currentViewerName && (a.itemType === "project" || a.itemType === "task"));

  return (
    <div>
      {full && (
        <div className="mb-4 flex items-center gap-2">
          <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight">Activity</h1>
          <span className="text-[9.5px] px-2 py-0.5 rounded-full font-medium flex items-center gap-1" style={{ background: "#5E9E8A22", color: "#5E9E8A" }}>
            <span className="w-1.5 h-1.5 rounded-full animate-pulse" style={{ background: "#5E9E8A" }} /> LIVE
          </span>
        </div>
      )}
      {full && <p className="text-[11.5px] text-[#6B6980] mb-4">Real-time thread across every campus — every edit, add, completion, and deadline change.</p>}

      {full && (todayCompletions.length > 0 || weekCompletions.length > 0) && (
        <div className="rounded-lg p-4 mb-5" style={{ background: "#5E9E8A18", border: "1px solid #5E9E8A55" }}>
          <div className="text-[12px] font-bold mb-2" style={{ color: "#3D6B5C" }}>Your completions — things you created that got finished</div>
          <div className="grid sm:grid-cols-2 gap-3">
            <div>
              <div className="text-[10.5px] text-[#6B6980] mb-1.5">Today ({todayCompletions.length})</div>
              <div className="space-y-1.5">
                {todayCompletions.map((a) => (
                  <div key={a.id} className="text-[11.5px] bg-[#FFFFFF] rounded-md px-2.5 py-1.5" style={{ border: "1px solid #E3E1F0" }}>
                    <div>{a.target}</div>
                    <div className="text-[10px] text-[#8B889C] flex items-center gap-1.5 mt-0.5">
                      by {a.actor} · {a.ts.split(" ").slice(1).join(" ")}
                      {a.itemType === "project" && a.budgetStatusVal && (
                        <span className="px-1.5 py-0.5 rounded-full" style={{ background: `${BUDGET_STATUS_COLOR[a.budgetStatusVal]}22`, color: BUDGET_STATUS_COLOR[a.budgetStatusVal] }}>{BUDGET_STATUS_LABEL[a.budgetStatusVal]}</span>
                      )}
                    </div>
                  </div>
                ))}
                {todayCompletions.length === 0 && <div className="text-[10.5px] text-[#8B889C]">Nothing yet today.</div>}
              </div>
            </div>
            <div>
              <div className="text-[10.5px] text-[#6B6980] mb-1.5">This week ({weekCompletions.length})</div>
              <div className="space-y-1.5">
                {weekCompletions.map((a) => (
                  <div key={a.id} className="text-[11.5px] bg-[#FFFFFF] rounded-md px-2.5 py-1.5" style={{ border: "1px solid #E3E1F0" }}>
                    <div>{a.target}</div>
                    <div className="text-[10px] text-[#8B889C] flex items-center gap-1.5 mt-0.5">
                      by {a.actor} · {a.date}
                      {a.itemType === "project" && a.budgetStatusVal && (
                        <span className="px-1.5 py-0.5 rounded-full" style={{ background: `${BUDGET_STATUS_COLOR[a.budgetStatusVal]}22`, color: BUDGET_STATUS_COLOR[a.budgetStatusVal] }}>{BUDGET_STATUS_LABEL[a.budgetStatusVal]}</span>
                      )}
                    </div>
                  </div>
                ))}
                {weekCompletions.length === 0 && <div className="text-[10.5px] text-[#8B889C]">Nothing yet this week.</div>}
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[560px]"><ActivityFeed activity={activity} /></div>
    </div>
  );
}

function progressPct(p) {
  if (!p.subtasks?.length) return 0;
  return Math.round((p.subtasks.filter((s) => s.done).length / p.subtasks.length) * 100);
}

const SORTERS = {
  due: (a, b) => new Date(a.due) - new Date(b.due),
  title: (a, b) => a.title.localeCompare(b.title),
  cost: (a, b) => (b.cost || 0) - (a.cost || 0),
  progress: (a, b) => progressPct(b) - progressPct(a),
  team: (a, b) => (a.owner || "").localeCompare(b.owner || ""),
};

function CampusBudgetSummary({ projects, accentColor }) {
  const b = rollupBudget(projects);
  const pct = b.total ? Math.min(100, Math.round((b.spent / b.total) * 100)) : 0;
  const topSpenders = [...projects].sort((a, b2) => projectBudget(b2).spent - projectBudget(a).spent).slice(0, 3);
  return (
    <div>
      <div className="flex items-baseline justify-between mb-1">
        <span className="text-[16px] font-semibold" style={{ fontFamily: "'JetBrains Mono', monospace", color: accentColor || "#B8862F" }}>{fmtMoney(b.spent)}</span>
        <span className="text-[11px] text-[#6B6980]">of {fmtMoney(b.total)}</span>
      </div>
      <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden mb-3"><div className="h-full rounded-full" style={{ width: `${pct}%`, background: accentColor || "#B8862F" }} /></div>
      <div className="space-y-1">
        {topSpenders.map((p) => {
          const pb = projectBudget(p);
          return (
            <div key={p.id} className="flex items-center justify-between text-[10.5px] text-[#6B6980]">
              <span className="truncate pr-2">{p.title}</span>
              <span className="shrink-0">{fmtMoney(pb.spent)} / {fmtMoney(pb.total)}</span>
            </div>
          );
        })}
        {projects.length === 0 && <div className="text-[11px] text-[#8B889C]">No projects with budget data yet.</div>}
      </div>
    </div>
  );
}

function BudgetPanel({ projects, campuses, campusLabel, onOpenProject, onSelectCampus, accentColor }) {
  const b = rollupBudget(projects);
  const pct = b.total ? Math.min(100, Math.round((b.spent / b.total) * 100)) : 0;
  const color = accentColor || "#B8862F";
  return (
    <div>
      <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight mb-1">Budget — {campusLabel}</h1>
      <p className="text-[12px] text-[#6B6980] mb-5">Tallies every project's estimate and spend, plus every task and sub-task budget underneath it.</p>

      <div className="rounded-lg p-4 mb-5" style={{ background: `${color}18`, border: `1px solid ${color}55` }}>
        <div className="text-[11px] mb-1" style={{ color }}>Total spent</div>
        <div className="text-[26px] font-semibold" style={{ fontFamily: "'JetBrains Mono', monospace", color }}>{fmtMoney(b.spent)} <span className="text-[15px] font-normal">of {fmtMoney(b.total)}</span></div>
        <div className="h-2 rounded-full bg-[#E3E1F0] overflow-hidden mt-2"><div className="h-full rounded-full" style={{ width: `${pct}%`, background: color }} /></div>
      </div>

      {onSelectCampus && (
        <>
          <h2 className="text-[13px] font-bold mb-2">By Campus</h2>
          <div className="space-y-1.5 mb-6">
            {campuses.map((c) => {
              const cb = rollupBudget(projects.filter((p) => p.campus === c.id));
              const cpct = cb.total ? Math.min(100, Math.round((cb.spent / cb.total) * 100)) : 0;
              return (
                <button key={c.id} onClick={() => onSelectCampus(c.id)} className="w-full text-left bg-[#FFFFFF] rounded-lg px-3 py-2.5" style={{ border: `1px solid ${c.color}55` }}>
                  <div className="flex items-center justify-between mb-1.5">
                    <span className="text-[12.5px] font-medium" style={{ color: c.color }}>{c.name} <span className="text-[10px] opacity-70">({c.abbr})</span></span>
                    <span className="text-[11px] text-[#6B6980]">{fmtMoney(cb.spent)} / {fmtMoney(cb.total)}</span>
                  </div>
                  <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden"><div className="h-full rounded-full" style={{ width: `${cpct}%`, background: c.color }} /></div>
                </button>
              );
            })}
          </div>
        </>
      )}

      <h2 className="text-[13px] font-bold mb-2">By Project</h2>
      <div className="space-y-1.5">
        {projects.map((p) => {
          const pb = projectBudget(p);
          const ppct = pb.total ? Math.min(100, Math.round((pb.spent / pb.total) * 100)) : 0;
          return (
            <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg px-3 py-2.5 hover:border-[#C7C3E0]">
              <div className="flex items-center justify-between mb-1 gap-2">
                <span className="text-[12.5px] truncate">{p.title}</span>
                <span className="text-[11px] text-[#6B6980] shrink-0">{fmtMoney(pb.spent)} / {fmtMoney(pb.total)}</span>
              </div>
              <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden mb-1"><div className="h-full rounded-full" style={{ width: `${ppct}%`, background: color }} /></div>
              <div className="text-[10px] text-[#8B889C]">{p.subtasks?.length || 0} task{p.subtasks?.length === 1 ? "" : "s"} rolled in</div>
            </button>
          );
        })}
        {projects.length === 0 && <div className="text-[11.5px] text-[#8B889C] py-6 text-center">No projects in view.</div>}
      </div>
    </div>
  );
}

function ProjectsBoard({ projects, campusLabel, onCycle, onOpen, sections, onAddSection, onSetSection, onNewProject }) {
  const [sortBy, setSortBy] = useState("due");
  const [filterStage, setFilterStage] = useState("All");
  const [filterSection, setFilterSection] = useState("All");
  const [addingSection, setAddingSection] = useState(false);
  const [newSectionName, setNewSectionName] = useState("");
  const [showCompleted, setShowCompleted] = useState(false);

  const submitSection = () => {
    if (!newSectionName.trim()) return;
    onAddSection(newSectionName.trim());
    setNewSectionName(""); setAddingSection(false);
  };

  const completedCount = projects.filter((p) => p.stage === "Completed").length;

  let list = projects.filter(
    (p) => (filterStage === "All" || p.stage === filterStage) && (filterSection === "All" || (p.section || "General") === filterSection) && (showCompleted || p.stage !== "Completed")
  );
  list = [...list].sort(SORTERS[sortBy]);

  const grouped = {};
  sections.forEach((s) => { grouped[s] = []; });
  list.forEach((p) => {
    const s = p.section || "General";
    if (!grouped[s]) grouped[s] = [];
    grouped[s].push(p);
  });

  return (
    <div>
      <div className="mb-4 flex items-center justify-between">
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight">Projects & Tasks — {campusLabel}</h1>
          <p className="text-[12px] text-[#6B6980] mt-1">Click a card to open cost, sub-tasks, and notes. Click the stage pill to advance it.</p>
        </div>
        <button onClick={onNewProject} className="flex items-center gap-1.5 text-[12px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium"><Plus size={14} /> New Project</button>
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-5 bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg px-3 py-2.5">
        <span className="text-[10.5px] text-[#6B6980]">Sort by</span>
        <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11.5px] outline-none">
          <option value="due">Due Date</option>
          <option value="title">Title A–Z</option>
          <option value="cost">Cost (high→low)</option>
          <option value="progress">Progress</option>
          <option value="team">Team Member A–Z</option>
        </select>
        <span className="text-[10.5px] text-[#6B6980] ml-2">Filter by stage</span>
        <select value={filterStage} onChange={(e) => setFilterStage(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11.5px] outline-none">
          <option value="All">All Stages</option>
          {STAGES.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <span className="text-[10.5px] text-[#6B6980] ml-2">Section</span>
        <select value={filterSection} onChange={(e) => setFilterSection(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11.5px] outline-none">
          <option value="All">All Sections</option>
          {sections.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <label className="flex items-center gap-1.5 text-[11px] text-[#6B6980] ml-2 cursor-pointer">
          <input type="checkbox" checked={showCompleted} onChange={(e) => setShowCompleted(e.target.checked)} className="accent-[#B8862F]" />
          Show completed ({completedCount})
        </label>
        <div className="ml-auto flex items-center gap-2">
          {addingSection ? (
            <>
              <input value={newSectionName} onChange={(e) => setNewSectionName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submitSection()}
                placeholder="Section name" autoFocus className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11.5px] outline-none focus:border-[#B8862F]" />
              <button onClick={submitSection} className="text-[11px] text-[#B8862F]">Add</button>
              <button onClick={() => setAddingSection(false)} className="text-[11px] text-[#6B6980]">Cancel</button>
            </>
          ) : (
            <button onClick={() => setAddingSection(true)} className="flex items-center gap-1 text-[11.5px] text-[#B8862F]"><Plus size={12} /> Add Section</button>
          )}
        </div>
      </div>

      <div className="space-y-6">
        {sections.map((section) => (
          <div key={section}>
            <div className="text-[12px] font-medium text-[#2A2A3A] mb-2 flex items-center gap-2">
              {section} <span className="text-[10.5px] text-[#6B6980] font-normal">· {grouped[section]?.length || 0}</span>
            </div>
            <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-2">
              {(grouped[section] || []).map((p) => (
                <div key={p.id} className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg p-3">
                  <button onClick={() => onOpen(p.id)} className="text-left w-full">
                    <div className="flex items-start gap-1.5 mb-2">
                      {p.shared && <Link2 size={12} className="text-[#B8862F] mt-0.5 shrink-0" />}
                      <div className="text-[12.5px] leading-snug">{p.title}</div>
                    </div>
                    <ProgressBar subtasks={p.subtasks} />
                    <div className="text-[10.5px] text-[#6B6980] my-2">{p.owner} · Due {p.due}</div>
                    {p.stage === "Completed" && (
                      <span className="text-[9.5px] px-1.5 py-0.5 rounded-full inline-block mb-1" style={{ background: `${BUDGET_STATUS_COLOR[budgetStatus(p)]}22`, color: BUDGET_STATUS_COLOR[budgetStatus(p)] }}>
                        Ended {BUDGET_STATUS_LABEL[budgetStatus(p)].toLowerCase()}
                      </span>
                    )}
                  </button>
                  <div className="flex items-center justify-between">
                    <button onClick={() => onCycle(p.id)}><StageBadge stage={p.stage} /></button>
                    <select value={p.section || "General"} onChange={(e) => onSetSection(p.id, e.target.value)}
                      className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-1.5 py-0.5 text-[9.5px] text-[#6B6980] outline-none">
                      {sections.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                </div>
              ))}
              {(!grouped[section] || grouped[section].length === 0) && (
                <div className="text-[11px] text-[#A8A5BE] px-1 py-4 text-center border border-dashed border-[#E3E1F0] rounded-lg sm:col-span-2 lg:col-span-3">Nothing here</div>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function ProjectModal({ project, onClose, onToggleSubtask, onAddSubtask, onSetStage, onAddProjectNote, onAddSubtaskNote, onAddProjectPhoto, onAddSubtaskPhoto, onUpdateProjectBudget, onUpdateSubtaskBudget, onSetDue, onSetSubtaskDue, onAssignSubtask, onDeleteSubtask, onDeleteProject, roster }) {
  const [showAddSubtask, setShowAddSubtask] = useState(false);
  const [subtaskText, setSubtaskText] = useState("");
  const [subtaskCost, setSubtaskCost] = useState("");
  const [recurFreq, setRecurFreq] = useState("none");
  const [startDate, setStartDate] = useState(TODAY_STR);
  const [customEvery, setCustomEvery] = useState(1);
  const [customUnit, setCustomUnit] = useState("weeks");
  const [endType, setEndType] = useState("never");
  const [count, setCount] = useState(4);
  const [untilDate, setUntilDate] = useState("2026-08-15");
  const [newNoteText, setNewNoteText] = useState("");
  const [expandedSubtask, setExpandedSubtask] = useState(null);
  const [pendingPhoto, setPendingPhoto] = useState(null); // { scope: "project" | idx, dataUrl }
  const [photoCaption, setPhotoCaption] = useState("");
  const [subtaskNoteDrafts, setSubtaskNoteDrafts] = useState({});
  const [editingBudget, setEditingBudget] = useState(false);
  const [budgetDraft, setBudgetDraft] = useState({ cost: "", spent: "" });
  const [editingSubtaskBudget, setEditingSubtaskBudget] = useState(null);
  const [subtaskBudgetDraft, setSubtaskBudgetDraft] = useState({ cost: "", spent: "" });
  const [subtaskSort, setSubtaskSort] = useState("default");
  const [editingDue, setEditingDue] = useState(false);
  const [dueDraft, setDueDraft] = useState("");
  const [dueTimeDraft, setDueTimeDraft] = useState("");
  const [editingSubtaskDue, setEditingSubtaskDue] = useState(null);
  const [subtaskDueDraft, setSubtaskDueDraft] = useState("");
  const [subtaskDueTimeDraft, setSubtaskDueTimeDraft] = useState("");
  const [editingAssignees, setEditingAssignees] = useState(null);
  const [subtaskAssigneesDraft, setSubtaskAssigneesDraft] = useState([]);
  const [newSubtaskTime, setNewSubtaskTime] = useState("");
  const [newSubtaskAssignees, setNewSubtaskAssignees] = useState([]);
  const [confirmDeleteProject, setConfirmDeleteProject] = useState(false);
  const [confirmDeleteSubtask, setConfirmDeleteSubtask] = useState(null);
  useLockBodyScroll();

  if (!project) return null;
  const pct = project.subtasks?.length ? Math.round((project.subtasks.filter((s) => s.done).length / project.subtasks.length) * 100) : 0;
  const allSubtasksDone = project.subtasks?.length > 0 && project.subtasks.every((s) => s.done);
  const readyToComplete = allSubtasksDone && project.stage !== "Completed";
  const budgetRoll = projectBudget(project);
  const subtaskEntries = (project.subtasks || []).map((s, i) => ({ s, i }));
  if (subtaskSort === "team") subtaskEntries.sort((a, b) => (a.s.createdBy || "").localeCompare(b.s.createdBy || ""));

  const submitSubtask = () => {
    if (!subtaskText.trim()) return;
    const recurrence = recurFreq === "none" ? { freq: "none" } : { freq: recurFreq, customEvery, customUnit, endType, count, untilDate };
    const due = recurFreq !== "none" ? startDate : (startDate || null);
    onAddSubtask(project.id, subtaskText.trim(), recurrence, due, subtaskCost, newSubtaskTime, newSubtaskAssignees);
    setSubtaskText(""); setSubtaskCost(""); setRecurFreq("none"); setEndType("never"); setStartDate(TODAY_STR); setNewSubtaskTime(""); setNewSubtaskAssignees([]); setShowAddSubtask(false);
  };

  const toggleNewSubtaskAssignee = (name) =>
    setNewSubtaskAssignees((prev) => prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name]);

  const startEditBudget = () => { setBudgetDraft({ cost: project.cost || 0, spent: project.spent || 0 }); setEditingBudget(true); };
  const saveBudget = () => { onUpdateProjectBudget(project.id, budgetDraft.cost, budgetDraft.spent); setEditingBudget(false); };

  const startEditSubtaskBudget = (i, s) => { setSubtaskBudgetDraft({ cost: s.cost || 0, spent: s.spent || 0 }); setEditingSubtaskBudget(i); };
  const saveSubtaskBudget = (i) => { onUpdateSubtaskBudget(project.id, i, subtaskBudgetDraft.cost, subtaskBudgetDraft.spent); setEditingSubtaskBudget(null); };

  const submitProjectNote = () => {
    if (!newNoteText.trim()) return;
    onAddProjectNote(project.id, newNoteText);
    setNewNoteText("");
  };

  const handleFileSelect = (scope, file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => { setPendingPhoto({ scope, dataUrl: reader.result }); setPhotoCaption(""); };
    reader.readAsDataURL(file);
  };

  const savePendingPhoto = () => {
    if (!pendingPhoto) return;
    if (pendingPhoto.scope === "project") onAddProjectPhoto(project.id, pendingPhoto.dataUrl, photoCaption);
    else onAddSubtaskPhoto(project.id, pendingPhoto.scope, pendingPhoto.dataUrl, photoCaption);
    setPendingPhoto(null); setPhotoCaption("");
  };

  return (
    <div className="fixed inset-0 z-30 overflow-y-auto p-4" style={{ background: "rgba(0,0,0,0.65)", WebkitOverflowScrolling: "touch" }} onClick={onClose}>
      <div className="border rounded-xl max-w-[560px] w-full my-8 mx-auto p-5" style={{ background: "#FFFFFF", borderColor: "#D8D5EC", color: "#2A2A3A", fontFamily: "'Inter', sans-serif" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between mb-3 gap-2">
          <div className="flex items-center gap-1.5 min-w-0">{project.shared && <Link2 size={13} className="text-[#B8862F] shrink-0" />}<h2 className="text-[clamp(14px,3vw,16px)] font-medium truncate">{project.title}</h2></div>
          <div className="flex items-center gap-2 shrink-0">
            {confirmDeleteProject ? (
              <span className="flex items-center gap-1">
                <button onClick={() => { onDeleteProject(project.id); onClose(); }} className="text-[10.5px] font-medium px-2 py-1 rounded" style={{ background: "#C15B5B", color: "#FFFFFF" }}>Confirm delete</button>
                <button onClick={() => setConfirmDeleteProject(false)} className="text-[10.5px] text-[#8B889C]">Cancel</button>
              </span>
            ) : (
              <button onClick={() => setConfirmDeleteProject(true)} className="text-[#8B889C] hover:text-[#C15B5B]"><Trash2 size={15} /></button>
            )}
            <button onClick={onClose} className="text-[#6B6980] hover:text-[#2A2A3A]"><X size={18} /></button>
          </div>
        </div>
        <div className="flex items-center gap-3 text-[11px] text-[#6B6980] mb-4 flex-wrap">
          <span>Owner: {project.owner}</span><span>·</span>
          {editingDue ? (
            <span className="flex items-center gap-1">
              <input type="date" value={dueDraft} onChange={(e) => setDueDraft(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded px-1.5 py-0.5 text-[10.5px] outline-none" />
              <input type="time" value={dueTimeDraft} onChange={(e) => setDueTimeDraft(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded px-1.5 py-0.5 text-[10.5px] outline-none" />
              <button onClick={() => { onSetDue(project.id, dueDraft, dueTimeDraft); setEditingDue(false); }} className="text-[#B8862F] font-medium">Save</button>
              <button onClick={() => setEditingDue(false)} className="text-[#8B889C]">Cancel</button>
            </span>
          ) : (
            <span className="flex items-center gap-1">
              Due {project.due}{project.dueTime ? ` ${project.dueTime}` : ""}
              <button onClick={() => { setDueDraft(project.due); setDueTimeDraft(project.dueTime || ""); setEditingDue(true); }} className="text-[#B8862F] underline decoration-dotted">edit</button>
            </span>
          )}
          <StageBadge stage={project.stage} />
          {project.stage === "Completed" && (
            <span className="text-[10.5px] px-2 py-0.5 rounded-full" style={{ background: `${BUDGET_STATUS_COLOR[budgetStatus(project)]}22`, color: BUDGET_STATUS_COLOR[budgetStatus(project)] }}>
              Ended {BUDGET_STATUS_LABEL[budgetStatus(project)].toLowerCase()}
            </span>
          )}
        </div>

        {readyToComplete && (
          <div className="rounded-md p-3 mb-4 flex items-center justify-between gap-3 flex-wrap" style={{ background: "#5E9E8A18", border: "1px solid #5E9E8A55" }}>
            <div className="flex items-center gap-2 text-[12.5px]" style={{ color: "#3D6B5C" }}>
              <CheckCircle2 size={15} /> All sub-tasks are complete. Mark this project Completed?
            </div>
            <button onClick={() => onSetStage(project.id, "Completed")} className="text-[11.5px] rounded-md px-3 py-1.5 font-medium shrink-0" style={{ background: "#5E9E8A", color: "#FFFFFF" }}>
              Mark Complete
            </button>
          </div>
        )}

        <div className="grid grid-cols-2 gap-3 mb-4">
          <div className="bg-[#EFEEFA] border border-[#E3E1F0] rounded-md p-3">
            <div className="flex items-center justify-between mb-1">
              <span className="text-[10.5px] text-[#6B6980]">Budget (incl. tasks)</span>
              {!editingBudget && <button onClick={startEditBudget} className="text-[10px] text-[#B8862F]">Edit</button>}
            </div>
            {editingBudget ? (
              <div className="space-y-1.5">
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] text-[#8B889C] w-[38px] shrink-0">Est.</span>
                  <input type="number" value={budgetDraft.cost} onChange={(e) => setBudgetDraft((d) => ({ ...d, cost: e.target.value }))} className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded px-1.5 py-1 text-[11px] outline-none" />
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="text-[10px] text-[#8B889C] w-[38px] shrink-0">Spent</span>
                  <input type="number" value={budgetDraft.spent} onChange={(e) => setBudgetDraft((d) => ({ ...d, spent: e.target.value }))} className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded px-1.5 py-1 text-[11px] outline-none" />
                </div>
                <div className="flex gap-1.5">
                  <button onClick={saveBudget} className="text-[10.5px] bg-[#B8862F] text-[#F7F6FB] rounded px-2 py-1 font-medium">Save</button>
                  <button onClick={() => setEditingBudget(false)} className="text-[10.5px] text-[#6B6980] px-1">Cancel</button>
                </div>
              </div>
            ) : (
              <>
                <div className="text-[clamp(13px,2.6vw,15px)] font-medium truncate" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{fmtMoney(budgetRoll.spent)} / {fmtMoney(budgetRoll.total)}</div>
                <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden mt-2"><div className="h-full rounded-full bg-[#B8862F]" style={{ width: `${budgetRoll.total ? Math.min(100, (budgetRoll.spent / budgetRoll.total) * 100) : 0}%` }} /></div>
                {project.subtasks?.length > 0 && <div className="text-[9px] text-[#8B889C] mt-1">incl. {fmtMoney(budgetRoll.total - (project.cost || 0))} across {project.subtasks.length} task{project.subtasks.length === 1 ? "" : "s"}</div>}
              </>
            )}
          </div>
          <div className="bg-[#EFEEFA] border border-[#E3E1F0] rounded-md p-3">
            <div className="text-[10.5px] text-[#6B6980] mb-1">Progress</div>
            <div className="text-[clamp(13px,2.6vw,15px)] font-medium" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{pct}%</div>
            <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden mt-2"><div className="h-full rounded-full bg-[#2B4C7E]" style={{ width: `${pct}%` }} /></div>
          </div>
        </div>

        {(project.team?.length || project.collaborators?.length) ? (
          <div className="mb-4 text-[11.5px]">
            {project.team?.length > 0 && <div className="mb-1"><span className="text-[#6B6980]">Team: </span>{project.team.join(", ")}</div>}
            {project.collaborators?.length > 0 && <div><span className="text-[#6B6980]">Collaborating with: </span>{project.collaborators.join(", ")}</div>}
          </div>
        ) : null}

        {/* Project-level photos */}
        <div className="mb-4">
          <div className="flex items-center justify-between mb-2">
            <div className="text-[11px] text-[#6B6980]">Photos</div>
            <label className="text-[10.5px] text-[#B8862F] flex items-center gap-1 cursor-pointer">
              <Camera size={12} /> Add photo
              <input type="file" accept="image/*" className="hidden" onChange={(e) => handleFileSelect("project", e.target.files[0])} />
            </label>
          </div>
          {pendingPhoto?.scope === "project" && (
            <div className="bg-[#EFEEFA] border border-[#E3E1F0] rounded-md p-3 mb-2">
              <img src={pendingPhoto.dataUrl} alt="preview" className="max-h-[140px] rounded-md mb-2" />
              <input value={photoCaption} onChange={(e) => setPhotoCaption(e.target.value)} placeholder="Add a memo for this photo…"
                className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[12px] outline-none mb-2" />
              <div className="flex gap-2">
                <button onClick={savePendingPhoto} className="text-[11.5px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium">Save Photo</button>
                <button onClick={() => setPendingPhoto(null)} className="text-[11.5px] text-[#6B6980] px-3 py-1.5">Cancel</button>
              </div>
            </div>
          )}
          {project.photos?.length > 0 && (
            <div className="grid grid-cols-3 gap-2">
              {project.photos.map((ph) => (
                <div key={ph.id} className="rounded-md overflow-hidden border border-[#E3E1F0]">
                  <img src={ph.dataUrl} alt={ph.caption || "task photo"} className="w-full h-[70px] object-cover" />
                  <div className="px-1.5 py-1 bg-[#EFEEFA]">
                    {ph.caption && <div className="text-[9.5px] truncate">{ph.caption}</div>}
                    <div className="text-[8.5px] text-[#8B889C] truncate">{ph.uploadedBy} · {ph.ts}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="mb-4">
          <div className="flex items-center justify-between mb-2 flex-wrap gap-1.5">
            <div className="text-[11px] text-[#6B6980]">Sub-tasks</div>
            <div className="flex items-center gap-2">
              <select value={subtaskSort} onChange={(e) => setSubtaskSort(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-1.5 py-1 text-[10px] outline-none">
                <option value="default">Sort: Default</option>
                <option value="team">Sort: Team Member</option>
              </select>
              <button onClick={() => setShowAddSubtask((v) => !v)} className="text-[10.5px] text-[#B8862F] flex items-center gap-1"><Plus size={12} /> Add sub-task</button>
            </div>
          </div>
          <div className="space-y-1.5 mb-2">
            {subtaskEntries.map(({ s, i }) => (
              <div key={s.id ?? i} className="rounded-md" style={{ background: expandedSubtask === i ? "#EFEEFA" : "transparent" }}>
                <div className="flex items-center gap-2 px-1 py-1 flex-wrap">
                  <button onClick={() => onToggleSubtask(project.id, i)} className="flex items-center gap-2 text-left flex-1 min-w-0 basis-full sm:basis-0">
                    {s.done ? <CheckCircle2 size={14} className="text-[#5E9E8A] shrink-0" /> : <Circle size={14} className="text-[#A8A5BE] shrink-0" />}
                    <span className={`text-[12.5px] truncate min-w-0 flex-1 ${s.done ? "line-through text-[#8B889C]" : ""}`}>{s.t}</span>
                  </button>
                  <div className="flex items-center gap-1.5 shrink-0 flex-wrap justify-end pl-[22px] sm:pl-0">
                    {editingSubtaskDue === i ? (
                      <span className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                        <input type="date" value={subtaskDueDraft} onChange={(e) => setSubtaskDueDraft(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded px-1 py-0.5 text-[9.5px] outline-none" />
                        <input type="time" value={subtaskDueTimeDraft} onChange={(e) => setSubtaskDueTimeDraft(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded px-1 py-0.5 text-[9.5px] outline-none" />
                        <button onClick={() => { onSetSubtaskDue(project.id, i, subtaskDueDraft, subtaskDueTimeDraft); setEditingSubtaskDue(null); }} className="text-[9.5px] text-[#B8862F] font-medium">Save</button>
                      </span>
                    ) : (
                      <span className="flex items-center gap-0.5 whitespace-nowrap">
                        {s.due && <span className="text-[9.5px] text-[#8B889C]">· {s.due}{s.dueTime ? ` ${s.dueTime}` : ""}</span>}
                        <button onClick={(e) => { e.stopPropagation(); setSubtaskDueDraft(s.due || TODAY_STR); setSubtaskDueTimeDraft(s.dueTime || ""); setEditingSubtaskDue(i); }} className="text-[9px] text-[#B8862F] underline decoration-dotted">edit</button>
                      </span>
                    )}
                    {s.assignees?.length > 0 && (
                      <span className="text-[9.5px] px-1.5 py-0.5 rounded-full bg-[#B8862F22] text-[#B8862F] whitespace-nowrap truncate max-w-[120px]">{s.assignees.join(", ")}</span>
                    )}
                    {s.recurrence && describeRecurrence(s.recurrence) && (
                      <span className="text-[9.5px] px-1.5 py-0.5 rounded-full bg-[#2B4C7E22] text-[#2B4C7E] whitespace-nowrap">🔁 {describeRecurrence(s.recurrence)}</span>
                    )}
                    {(s.photos?.length > 0) && <span className="text-[9.5px] text-[#8B889C] flex items-center gap-0.5 whitespace-nowrap"><ImageIcon size={11} />{s.photos.length}</span>}
                    <button onClick={() => setExpandedSubtask(expandedSubtask === i ? null : i)} className="text-[#8B889C] hover:text-[#B8862F] text-[10px] px-1 whitespace-nowrap">
                      {expandedSubtask === i ? "Close" : "Details"}
                    </button>
                    {confirmDeleteSubtask === i ? (
                      <span className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                        <button onClick={() => { onDeleteSubtask(project.id, i); setConfirmDeleteSubtask(null); }} className="text-[9.5px] font-medium px-1.5 py-0.5 rounded" style={{ background: "#C15B5B", color: "#FFFFFF" }}>Confirm</button>
                        <button onClick={() => setConfirmDeleteSubtask(null)} className="text-[9.5px] text-[#8B889C]">Cancel</button>
                      </span>
                    ) : (
                      <button onClick={(e) => { e.stopPropagation(); setConfirmDeleteSubtask(i); }} className="text-[#8B889C] hover:text-[#C15B5B] shrink-0"><Trash2 size={12} /></button>
                    )}
                  </div>
                </div>

                {expandedSubtask === i && (
                  <div className="px-3 pb-3 pt-1 space-y-2.5">
                    {/* sub-task budget */}
                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <span className="text-[10px] text-[#6B6980]">Budget</span>
                        {editingSubtaskBudget !== i && (
                          <button onClick={() => startEditSubtaskBudget(i, s)} className="text-[10px] text-[#B8862F]">Edit</button>
                        )}
                      </div>
                      {editingSubtaskBudget === i ? (
                        <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-md p-2 space-y-1.5">
                          <div className="flex items-center gap-1.5">
                            <span className="text-[10px] text-[#8B889C] w-[38px] shrink-0">Est.</span>
                            <input type="number" value={subtaskBudgetDraft.cost} onChange={(e) => setSubtaskBudgetDraft((d) => ({ ...d, cost: e.target.value }))} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded px-1.5 py-1 text-[11px] outline-none" />
                          </div>
                          <div className="flex items-center gap-1.5">
                            <span className="text-[10px] text-[#8B889C] w-[38px] shrink-0">Spent</span>
                            <input type="number" value={subtaskBudgetDraft.spent} onChange={(e) => setSubtaskBudgetDraft((d) => ({ ...d, spent: e.target.value }))} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded px-1.5 py-1 text-[11px] outline-none" />
                          </div>
                          <div className="flex gap-1.5">
                            <button onClick={() => saveSubtaskBudget(i)} className="text-[10.5px] bg-[#B8862F] text-[#F7F6FB] rounded px-2 py-1 font-medium">Save</button>
                            <button onClick={() => setEditingSubtaskBudget(null)} className="text-[10.5px] text-[#6B6980] px-1">Cancel</button>
                          </div>
                        </div>
                      ) : (
                        <div className="text-[11px] text-[#6B6980]">{fmtMoney(s.spent || 0)} / {fmtMoney(s.cost || 0)} spent</div>
                      )}
                    </div>
                    {/* sub-task assignees */}
                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <span className="text-[10px] text-[#6B6980]">Assigned to</span>
                        {editingAssignees !== i && (
                          <button onClick={() => { setSubtaskAssigneesDraft(s.assignees || []); setEditingAssignees(i); }} className="text-[10px] text-[#B8862F]">Edit</button>
                        )}
                      </div>
                      {editingAssignees === i ? (
                        <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-md p-2">
                          <div className="flex flex-wrap gap-x-3 gap-y-1 mb-2">
                            {(roster || []).map((r) => (
                              <label key={r.name} className="flex items-center gap-1 text-[10.5px] cursor-pointer">
                                <input type="checkbox" checked={subtaskAssigneesDraft.includes(r.name)}
                                  onChange={() => setSubtaskAssigneesDraft((prev) => prev.includes(r.name) ? prev.filter((n) => n !== r.name) : [...prev, r.name])}
                                  className="accent-[#B8862F]" />
                                {r.name}
                              </label>
                            ))}
                          </div>
                          <div className="flex gap-1.5">
                            <button onClick={() => { onAssignSubtask(project.id, i, subtaskAssigneesDraft); setEditingAssignees(null); }} className="text-[10.5px] bg-[#B8862F] text-[#F7F6FB] rounded px-2 py-1 font-medium">Save</button>
                            <button onClick={() => setEditingAssignees(null)} className="text-[10.5px] text-[#6B6980] px-1">Cancel</button>
                          </div>
                        </div>
                      ) : (
                        <div className="text-[11px] text-[#6B6980]">{s.assignees?.length > 0 ? s.assignees.join(", ") : "No one assigned"}</div>
                      )}
                    </div>
                    {/* sub-task photos */}
                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <span className="text-[10px] text-[#6B6980]">Photos</span>
                        <label className="text-[10px] text-[#B8862F] flex items-center gap-1 cursor-pointer">
                          <Camera size={11} /> Add photo
                          <input type="file" accept="image/*" className="hidden" onChange={(e) => handleFileSelect(i, e.target.files[0])} />
                        </label>
                      </div>
                      {pendingPhoto?.scope === i && (
                        <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-md p-2 mb-2">
                          <img src={pendingPhoto.dataUrl} alt="preview" className="max-h-[110px] rounded-md mb-1.5" />
                          <input value={photoCaption} onChange={(e) => setPhotoCaption(e.target.value)} placeholder="Add a memo for this photo…"
                            className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none mb-1.5" />
                          <div className="flex gap-2">
                            <button onClick={savePendingPhoto} className="text-[11px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-2.5 py-1 font-medium">Save</button>
                            <button onClick={() => setPendingPhoto(null)} className="text-[11px] text-[#6B6980] px-2">Cancel</button>
                          </div>
                        </div>
                      )}
                      {s.photos?.length > 0 && (
                        <div className="grid grid-cols-3 gap-1.5">
                          {s.photos.map((ph) => (
                            <div key={ph.id} className="rounded-md overflow-hidden border border-[#E3E1F0]">
                              <img src={ph.dataUrl} alt={ph.caption || "sub-task photo"} className="w-full h-[56px] object-cover" />
                              {ph.caption && <div className="text-[8.5px] px-1 py-0.5 bg-[#FFFFFF] truncate">{ph.caption}</div>}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    {/* sub-task notes */}
                    <div>
                      <div className="text-[10px] text-[#6B6980] mb-1.5">Notes</div>
                      <div className="flex gap-1.5 mb-1.5">
                        <input value={subtaskNoteDrafts[i] || ""} onChange={(e) => setSubtaskNoteDrafts((prev) => ({ ...prev, [i]: e.target.value }))}
                          onKeyDown={(e) => { if (e.key === "Enter") { onAddSubtaskNote(project.id, i, subtaskNoteDrafts[i] || ""); setSubtaskNoteDrafts((prev) => ({ ...prev, [i]: "" })); } }}
                          placeholder="Add a note to this sub-task…" className="flex-1 bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none" />
                        <button onClick={() => { onAddSubtaskNote(project.id, i, subtaskNoteDrafts[i] || ""); setSubtaskNoteDrafts((prev) => ({ ...prev, [i]: "" })); }} className="text-[10.5px] text-[#B8862F] px-2">Add</button>
                      </div>
                      <div className="space-y-1">
                        {(s.notes || []).map((n, ni) => (
                          <div key={ni} className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-md px-2 py-1.5">
                            <div className="text-[11px]">{n.text}</div>
                            <div className="text-[9.5px] text-[#8B889C] mt-0.5">{n.author} · {n.ts}</div>
                          </div>
                        ))}
                        {(!s.notes || s.notes.length === 0) && <div className="text-[10.5px] text-[#8B889C]">No notes yet.</div>}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>

          {showAddSubtask && (
            <div className="bg-[#EFEEFA] border border-[#E3E1F0] rounded-md p-3 space-y-2">
              <input value={subtaskText} onChange={(e) => setSubtaskText(e.target.value)} placeholder="Sub-task title"
                className="w-full bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[12px] outline-none focus:border-[#B8862F]" />
              <div className="flex items-center gap-2">
                <span className="text-[10.5px] text-[#6B6980] w-[60px] shrink-0">Due</span>
                <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />
                <input type="time" value={newSubtaskTime} onChange={(e) => setNewSubtaskTime(e.target.value)} className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />
              </div>
              <div className="flex items-center gap-2">
                <span className="text-[10.5px] text-[#6B6980] w-[60px] shrink-0">Est. cost</span>
                <input type="number" value={subtaskCost} onChange={(e) => setSubtaskCost(e.target.value)} placeholder="0"
                  className="flex-1 bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />
              </div>
              <div className="flex items-start gap-2">
                <span className="text-[10.5px] text-[#6B6980] w-[60px] shrink-0 pt-1">Assign to</span>
                <div className="flex-1 flex flex-wrap gap-x-3 gap-y-1 bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1.5">
                  {(roster || []).map((r) => (
                    <label key={r.name} className="flex items-center gap-1 text-[10.5px] cursor-pointer">
                      <input type="checkbox" checked={newSubtaskAssignees.includes(r.name)} onChange={() => toggleNewSubtaskAssignee(r.name)} className="accent-[#B8862F]" />
                      {r.name}
                    </label>
                  ))}
                  {(!roster || roster.length === 0) && <span className="text-[10px] text-[#8B889C]">No one available to assign yet.</span>}
                </div>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-[10.5px] text-[#6B6980] w-[60px] shrink-0">Repeats</span>
                <select value={recurFreq} onChange={(e) => setRecurFreq(e.target.value)} className="flex-1 bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none">
                  {RECUR_FREQ.map((f) => <option key={f.id} value={f.id}>{f.label}</option>)}
                </select>
              </div>
              {recurFreq === "custom" && (
                <div className="flex items-center gap-2">
                  <span className="text-[10.5px] text-[#6B6980] w-[60px] shrink-0">Every</span>
                  <input type="number" min="1" value={customEvery} onChange={(e) => setCustomEvery(e.target.value)} className="w-[56px] bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />
                  <select value={customUnit} onChange={(e) => setCustomUnit(e.target.value)} className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none">
                    <option value="days">Days</option><option value="weeks">Weeks</option><option value="months">Months</option>
                  </select>
                </div>
              )}
              {recurFreq !== "none" && (
                <div className="flex items-center gap-2">
                  <span className="text-[10.5px] text-[#6B6980] w-[60px] shrink-0">Ends</span>
                  <select value={endType} onChange={(e) => setEndType(e.target.value)} className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none">
                    <option value="never">Never</option><option value="count">After # times</option><option value="date">On date</option>
                  </select>
                  {endType === "count" && <input type="number" min="1" value={count} onChange={(e) => setCount(e.target.value)} className="w-[56px] bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />}
                  {endType === "date" && <input type="date" value={untilDate} onChange={(e) => setUntilDate(e.target.value)} className="bg-[#FFFFFF] border border-[#D8D5EC] rounded-md px-2 py-1 text-[11px] outline-none" />}
                </div>
              )}
              <div className="flex gap-2 pt-1">
                <button onClick={submitSubtask} className="text-[11.5px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 py-1.5 font-medium">Add</button>
                <button onClick={() => setShowAddSubtask(false)} className="text-[11.5px] text-[#6B6980] px-3 py-1.5">Cancel</button>
              </div>
            </div>
          )}
        </div>

        <div>
          <div className="text-[11px] text-[#6B6980] mb-2">Notes (time-stamped, visible to team)</div>
          <div className="flex gap-2 mb-2">
            <input value={newNoteText} onChange={(e) => setNewNoteText(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submitProjectNote()}
              placeholder="Add a note for the team…" className="flex-1 bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[12px] outline-none focus:border-[#B8862F]" />
            <button onClick={submitProjectNote} className="text-[11.5px] bg-[#B8862F] text-[#F7F6FB] rounded-md px-3 font-medium">Add</button>
          </div>
          <div className="space-y-1.5">
            {project.notes?.map((n, i) => (
              <div key={i} className="bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2">
                <div className="text-[12px]">{n.text}</div>
                <div className="text-[10px] text-[#8B889C] mt-1">{n.author} · {n.ts}</div>
              </div>
            ))}
            {(!project.notes || project.notes.length === 0) && <div className="text-[11px] text-[#8B889C]">No notes yet.</div>}
          </div>
        </div>
      </div>
    </div>
  );
}

function StageBadge({ stage }) {
  const map = { "Pending": "#8B889C", "Started": "#B8862F", "In Progress": "#2B4C7E", "Stalled": "#C15B5B", "Completed": "#5E9E8A" };
  return <span className="inline-block text-[10.5px] px-2 py-0.5 rounded-full" style={{ background: `${map[stage]}22`, color: map[stage] }}>{stage}</span>;
}

const seedEvents = [
  { id: 1, title: "Worship Night", date: "2026-07-22", campuses: ["abv"], type: "Campus" },
  { id: 2, title: "Quarterly Men's Community Gathering", date: "2026-08-09", campuses: ["laf", "ynv", "brx"], type: "Cross-Campus" },
  { id: 3, title: "Central Ops / OD Touchpoint", date: "2026-07-17", campuses: ["all"], type: "Central" },
  { id: 4, title: "Safety Drill Weekend", date: "2026-08-10", campuses: ["ynv", "nib"], type: "Cross-Campus" },
];
function EventsList({ campusId, campuses }) {
  const events = seedEvents.filter((e) => !campusId || e.campuses.includes(campusId) || e.campuses.includes("all"));
  const typeColor = { "Campus": "#2B4C7E", "Cross-Campus": "#B8862F", "Central": "#6B4FA0" };
  return (
    <div>
      <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight mb-5">Cross-Campus Events</h1>
      <div className="space-y-2 max-w-[640px]">
        {events.map((e) => (
          <div key={e.id} className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg px-4 py-3 flex items-center justify-between">
            <div>
              <div className="text-[13.5px]">{e.title}</div>
              <div className="text-[11px] text-[#6B6980] mt-0.5">{e.date} · {e.campuses.includes("all") ? "All campuses" : e.campuses.map((id) => campuses.find((c) => c.id === id)?.name).join(", ")}</div>
            </div>
            <span className="text-[10.5px] px-2 py-0.5 rounded-full" style={{ background: `${typeColor[e.type]}22`, color: typeColor[e.type] }}>{e.type}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function NewProjectModal({ onClose, onCreate, sections, campusRoster, fullRoster, campusLabel }) {
  const [title, setTitle] = useState("");
  const [section, setSection] = useState(sections[0] || "General");
  const [stage, setStage] = useState("Pending");
  const [due, setDue] = useState(TODAY_STR);
  const [dueTime, setDueTime] = useState("");
  const [cost, setCost] = useState("");
  const [owner, setOwner] = useState("");
  const [team, setTeam] = useState("");
  const [collaborators, setCollaborators] = useState("");
  const [isCrossCampus, setIsCrossCampus] = useState(false);
  const [recurFreq, setRecurFreq] = useState("none");
  const [customEvery, setCustomEvery] = useState(1);
  const [customUnit, setCustomUnit] = useState("weeks");
  const [endType, setEndType] = useState("never");
  const [count, setCount] = useState(4);
  const [untilDate, setUntilDate] = useState("2026-08-15");
  useLockBodyScroll();

  const roster = isCrossCampus ? fullRoster : campusRoster;

  const submit = () => {
    if (!title.trim()) return;
    const recurrence = recurFreq === "none" ? { freq: "none" } : { freq: recurFreq, customEvery, customUnit, endType, count, untilDate };
    onCreate({ title: title.trim(), section, stage, due, dueTime, cost, owner: owner.trim(), team, collaborators, recurrence, shared: isCrossCampus });
  };

  return (
    <div className="fixed inset-0 z-30 overflow-y-auto p-4" style={{ background: "rgba(0,0,0,0.65)", WebkitOverflowScrolling: "touch" }} onClick={onClose}>
      <div className="border rounded-xl max-w-[520px] w-full my-8 mx-auto p-5" style={{ background: "#FFFFFF", borderColor: "#D8D5EC", color: "#2A2A3A", fontFamily: "'Inter', sans-serif" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between mb-4">
          <h2 className="text-[16px] font-bold">New Project / Task</h2>
          <button onClick={onClose} className="text-[#6B6980] hover:text-[#2A2A3A]"><X size={18} /></button>
        </div>

        <div className="space-y-3">
          <div>
            <label className="text-[10.5px] text-[#6B6980] block mb-1">Title</label>
            <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="What needs to happen?"
              className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-3 py-2 text-[13px] outline-none focus:border-[#B8862F]" />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-[10.5px] text-[#6B6980] block mb-1">Section</label>
              <select value={section} onChange={(e) => setSection(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12.5px] outline-none">
                {sections.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <label className="text-[10.5px] text-[#6B6980] block mb-1">Stage</label>
              <select value={stage} onChange={(e) => setStage(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12.5px] outline-none">
                {STAGES.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>

          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="text-[10.5px] text-[#6B6980] block mb-1">Due date</label>
              <input type="date" value={due} onChange={(e) => setDue(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12.5px] outline-none" />
            </div>
            <div>
              <label className="text-[10.5px] text-[#6B6980] block mb-1">Due time (optional)</label>
              <input type="time" value={dueTime} onChange={(e) => setDueTime(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12.5px] outline-none" />
            </div>
            <div>
              <label className="text-[10.5px] text-[#6B6980] block mb-1">Estimated cost ($)</label>
              <input type="number" value={cost} onChange={(e) => setCost(e.target.value)} placeholder="0" className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12.5px] outline-none" />
            </div>
          </div>

          <label className="flex items-center gap-2 text-[12px] text-[#2A2A3A] cursor-pointer bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-3 py-2">
            <input type="checkbox" checked={isCrossCampus} onChange={(e) => { setIsCrossCampus(e.target.checked); setOwner(""); setTeam(""); }} className="accent-[#B8862F]" />
            This is a cross-campus project — pull people from other campuses or Central
          </label>

          <div>
            <label className="text-[10.5px] text-[#6B6980] block mb-1">Owner</label>
            <select value={owner} onChange={(e) => setOwner(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-3 py-2 text-[12.5px] outline-none">
              <option value="">Select who owns this…</option>
              {roster.map((p) => <option key={p.name} value={p.name}>{p.name}{p.roles?.length ? ` — ${p.roles[0]}` : ""}</option>)}
            </select>
          </div>
          <div>
            <label className="text-[10.5px] text-[#6B6980] block mb-1">Team — available people {isCrossCampus ? "org-wide" : `for ${campusLabel}`}</label>
            <div className="flex flex-wrap gap-1.5 bg-[#EFEEFA] border border-[#D8D5EC] rounded-md p-2">
              {roster.length === 0 && <span className="text-[11px] text-[#8B889C]">No roster available — add staff first.</span>}
              {roster.map((p) => {
                const selected = team.split(",").map((s) => s.trim()).includes(p.name);
                return (
                  <button key={p.name} type="button"
                    onClick={() => {
                      const list = team.split(",").map((s) => s.trim()).filter(Boolean);
                      const next = selected ? list.filter((n) => n !== p.name) : [...list, p.name];
                      setTeam(next.join(", "));
                    }}
                    className={`text-[10.5px] px-2 py-1 rounded-full border ${selected ? "bg-[#B8862F] text-[#F7F6FB] border-[#B8862F]" : "border-[#D8D5EC] text-[#6B6980]"}`}>
                    {p.name}
                  </button>
                );
              })}
            </div>
          </div>
          <div className="grid grid-cols-1 gap-3">
            <div>
              <label className="text-[10.5px] text-[#6B6980] block mb-1">Collaborating teams</label>
              <input value={collaborators} onChange={(e) => setCollaborators(e.target.value)} placeholder="e.g. Worship Team" className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12px] outline-none" />
            </div>
          </div>

          <div className="border-t border-[#E3E1F0] pt-3">
            <label className="text-[10.5px] text-[#6B6980] block mb-1.5">Repeats</label>
            <select value={recurFreq} onChange={(e) => setRecurFreq(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-2 text-[12.5px] outline-none mb-2">
              {RECUR_FREQ.map((f) => <option key={f.id} value={f.id}>{f.label}</option>)}
            </select>
            {recurFreq === "custom" && (
              <div className="flex items-center gap-2 mb-2">
                <span className="text-[10.5px] text-[#6B6980]">Every</span>
                <input type="number" min="1" value={customEvery} onChange={(e) => setCustomEvery(e.target.value)} className="w-[56px] bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none" />
                <select value={customUnit} onChange={(e) => setCustomUnit(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none">
                  <option value="days">Days</option><option value="weeks">Weeks</option><option value="months">Months</option>
                </select>
              </div>
            )}
            {recurFreq !== "none" && (
              <div className="flex items-center gap-2">
                <span className="text-[10.5px] text-[#6B6980]">Ends</span>
                <select value={endType} onChange={(e) => setEndType(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none">
                  <option value="never">Never</option><option value="count">After # times</option><option value="date">On date</option>
                </select>
                {endType === "count" && <input type="number" min="1" value={count} onChange={(e) => setCount(e.target.value)} className="w-[56px] bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none" />}
                {endType === "date" && <input type="date" value={untilDate} onChange={(e) => setUntilDate(e.target.value)} className="bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none" />}
              </div>
            )}
          </div>
        </div>

        <div className="flex gap-2 mt-5">
          <button onClick={submit} className="text-[12.5px] rounded-md px-4 py-2 font-medium" style={{ background: "#B8862F", color: "#FFFFFF" }}>Create Project</button>
          <button onClick={onClose} className="text-[12.5px] text-[#6B6980] px-3 py-2">Cancel</button>
        </div>
      </div>
    </div>
  );
}

function SummaryDetailModal({ type, onClose, campuses, projects, sharedProjects, staffByCampus, orgBudgetUsed, orgBudgetTotal, onOpenProject, onSelectCampus }) {
  useLockBodyScroll();
  const titles = { projects: "Open Projects — Org-Wide", budget: "Budget Used — Org-Wide", campuses: "Campuses", shared: "Shared Projects — Cross-Campus" };
  const campusName = (id) => campuses.find((c) => c.id === id)?.name || id;

  return (
    <div className="fixed inset-0 z-30 overflow-y-auto p-4" style={{ background: "rgba(0,0,0,0.65)", WebkitOverflowScrolling: "touch" }} onClick={onClose}>
      <div className="border rounded-xl max-w-[560px] w-full my-8 mx-auto p-5" style={{ background: "#FFFFFF", borderColor: "#D8D5EC", color: "#2A2A3A", fontFamily: "'Inter', sans-serif" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between mb-4">
          <h2 className="text-[clamp(14px,3vw,16px)] font-bold">{titles[type]}</h2>
          <button onClick={onClose} className="text-[#6B6980] hover:text-[#2A2A3A]"><X size={18} /></button>
        </div>

        {type === "projects" && (
          <div className="space-y-4">
            {(() => {
              const open = projects.filter((p) => p.stage !== "Completed");
              const central = open.filter((p) => p.campus === "central");
              const sortedCampuses = [...campuses].sort((a, b) => a.name.localeCompare(b.name));
              return (
                <>
                  {central.length > 0 && (
                    <div>
                      <div className="text-[11px] font-bold text-[#6B4FA0] mb-1.5">Central — Org-Wide</div>
                      <div className="space-y-1.5">
                        {central.map((p) => (
                          <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2.5 hover:border-[#B8862F]">
                            <div className="flex items-center justify-between gap-2">
                              <span className="text-[12.5px] truncate">{p.title}</span>
                              <StageBadge stage={p.stage} />
                            </div>
                            <div className="text-[10.5px] text-[#6B6980] mt-1">{p.owner} · Due {p.due}</div>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                  {sortedCampuses.map((c) => {
                    const campusProjects = open.filter((p) => p.campus === c.id);
                    if (campusProjects.length === 0) return null;
                    return (
                      <div key={c.id}>
                        <div className="text-[11px] font-bold mb-1.5" style={{ color: c.color }}>{c.name} <span className="font-normal opacity-70">({c.abbr})</span></div>
                        <div className="space-y-1.5">
                          {campusProjects.map((p) => (
                            <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2.5 hover:border-[#B8862F]">
                              <div className="flex items-center justify-between gap-2">
                                <span className="text-[12.5px] truncate">{p.title}</span>
                                <StageBadge stage={p.stage} />
                              </div>
                              <div className="text-[10.5px] text-[#6B6980] mt-1">{p.owner} · Due {p.due}</div>
                            </button>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </>
              );
            })()}
          </div>
        )}

        {type === "budget" && (
          <div>
            <div className="rounded-md p-3 mb-3" style={{ background: "#B8862F18", border: "1px solid #B8862F55" }}>
              <div className="text-[11px]" style={{ color: "#B8862F" }}>Org total — every project, task & sub-task budget</div>
              <div className="text-[20px] font-semibold" style={{ fontFamily: "'JetBrains Mono', monospace", color: "#B8862F" }}>{fmtMoney(orgBudgetUsed)} <span className="text-[13px] font-normal">of {fmtMoney(orgBudgetTotal)}</span></div>
            </div>
            <div className="space-y-1.5">
              {campuses.map((c) => {
                const b = rollupBudget(projects.filter((p) => p.campus === c.id));
                return (
                  <button key={c.id} onClick={() => onSelectCampus(c.id)} className="w-full text-left bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2.5 hover:border-[#B8862F]">
                    <div className="flex items-center justify-between mb-1.5">
                      <span className="text-[12.5px] font-medium" style={{ color: c.color }}>{c.name} <span className="text-[10px] opacity-70">({c.abbr})</span></span>
                      <span className="text-[11px] text-[#6B6980]">{fmtMoney(b.spent)} / {fmtMoney(b.total)}</span>
                    </div>
                    <div className="h-1.5 rounded-full bg-[#E3E1F0] overflow-hidden"><div className="h-full rounded-full" style={{ width: `${b.total ? Math.min(100, (b.spent / b.total) * 100) : 0}%`, background: c.color }} /></div>
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {type === "campuses" && (
          <div className="space-y-1.5">
            {campuses.map((c) => (
              <button key={c.id} onClick={() => onSelectCampus(c.id)} className="w-full text-left rounded-md px-3 py-2.5" style={{ background: `${c.color}12`, border: `1px solid ${c.color}55` }}>
                <div className="flex items-center justify-between">
                  <span className="text-[12.5px] font-medium" style={{ color: c.color }}>{c.name} <span className="text-[10px] opacity-70">({c.abbr})</span></span>
                  <span className="text-[10.5px] text-[#6B6980]">Phase {c.phase}</span>
                </div>
                <div className="text-[10.5px] text-[#6B6980] mt-1">OD: {c.lead} · {c.abbr} · {(staffByCampus[c.id] || []).length} staff</div>
              </button>
            ))}
          </div>
        )}

        {type === "shared" && (
          <div className="space-y-1.5">
            {sharedProjects.length === 0 && <div className="text-[12px] text-[#6B6980] py-4 text-center">No shared projects right now.</div>}
            {sharedProjects.map((p) => (
              <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left bg-[#EFEEFA] border border-[#E3E1F0] rounded-md px-3 py-2.5 hover:border-[#6B4FA0]">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[12.5px] truncate flex items-center gap-1"><Link2 size={11} style={{ color: "#6B4FA0" }} />{p.title}</span>
                  <StageBadge stage={p.stage} />
                </div>
                <div className="text-[10.5px] text-[#6B6980] mt-1">
                  Owner: {p.owner} · Due {p.due} · Involves {p.sharedWith?.includes("all") ? "all campuses" : (p.sharedWith || []).map(campusName).join(", ")}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Best-available "latest activity" signal for a project: the most recent note on the
// project or any of its sub-tasks, falling back to when it was created.
function latestActivity(p) {
  const dates = [
    ...(p.notes || []).map((n) => n.ts?.slice(0, 10)),
    ...(p.subtasks || []).flatMap((s) => (s.notes || []).map((n) => n.ts?.slice(0, 10))),
    p.createdAt,
  ].filter(Boolean);
  return dates.length ? dates.sort().slice(-1)[0] : p.createdAt || null;
}

function ReportsPanel({ projects, campuses, staffByCampus, roster, isCentral, onClearAllProjects, onLoadDemoData }) {
  const [confirmingClear, setConfirmingClear] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [selectedCampuses, setSelectedCampuses] = useState([]);
  const [staffFilter, setStaffFilter] = useState("");
  const [titleFilter, setTitleFilter] = useState("");
  const [costMin, setCostMin] = useState("");
  const [costMax, setCostMax] = useState("");
  const [addedFrom, setAddedFrom] = useState("");
  const [addedTo, setAddedTo] = useState("");
  const [completedFrom, setCompletedFrom] = useState("");
  const [completedTo, setCompletedTo] = useState("");
  const [activityFrom, setActivityFrom] = useState("");
  const [activityTo, setActivityTo] = useState("");
  const [crossCampusOnly, setCrossCampusOnly] = useState(false);

  const campusName = (id) => id === "central" ? "Central" : (campuses.find((c) => c.id === id)?.name || id);

  const toggleCampus = (id) => setSelectedCampuses((prev) => prev.includes(id) ? prev.filter((c) => c !== id) : [...prev, id]);

  const results = useMemo(() => {
    return projects.filter((p) => {
      if (selectedCampuses.length && !selectedCampuses.includes(p.campus)) return false;
      if (staffFilter.trim()) {
        const q = staffFilter.toLowerCase();
        const inTeam = (p.team || []).some((t) => t.toLowerCase().includes(q));
        if (!p.owner?.toLowerCase().includes(q) && !inTeam) return false;
      }
      if (titleFilter.trim() && !p.title.toLowerCase().includes(titleFilter.toLowerCase())) return false;
      const budget = projectBudget(p);
      if (costMin !== "" && budget.total < Number(costMin)) return false;
      if (costMax !== "" && budget.total > Number(costMax)) return false;
      if (addedFrom && (!p.createdAt || p.createdAt < addedFrom)) return false;
      if (addedTo && (!p.createdAt || p.createdAt > addedTo)) return false;
      if (completedFrom && (!p.completedOn || p.completedOn < completedFrom)) return false;
      if (completedTo && (!p.completedOn || p.completedOn > completedTo)) return false;
      const la = latestActivity(p);
      if (activityFrom && (!la || la < activityFrom)) return false;
      if (activityTo && (!la || la > activityTo)) return false;
      if (crossCampusOnly && !p.shared) return false;
      return true;
    }).sort((a, b) => a.title.localeCompare(b.title));
  }, [projects, selectedCampuses, staffFilter, titleFilter, costMin, costMax, addedFrom, addedTo, completedFrom, completedTo, activityFrom, activityTo, crossCampusOnly]);

  const clearFilters = () => {
    setSelectedCampuses([]); setStaffFilter(""); setTitleFilter(""); setCostMin(""); setCostMax("");
    setAddedFrom(""); setAddedTo(""); setCompletedFrom(""); setCompletedTo(""); setActivityFrom(""); setActivityTo(""); setCrossCampusOnly(false);
  };

  const exportDocx = async () => {
    const headerRow = new TableRow({
      children: ["Title", "Campus", "Stage", "Owner", "Budget", "Spent", "Added", "Completed", "Latest Activity"].map((h) =>
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, bold: true })] })] })
      ),
    });
    const rows = results.map((p) => {
      const b = projectBudget(p);
      const cells = [p.title, campusName(p.campus), p.stage, p.owner, fmtMoney(b.total), fmtMoney(b.spent), p.createdAt || "—", p.completedOn || "—", latestActivity(p) || "—"];
      return new TableRow({ children: cells.map((text) => new TableCell({ children: [new Paragraph(String(text))] })) });
    });
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({ text: "OpsCore Custom Report", heading: HeadingLevel.HEADING1 }),
          new Paragraph({ text: `Generated ${TODAY_STR} · ${results.length} result${results.length === 1 ? "" : "s"}` }),
          new Paragraph({ text: "" }),
          new Table({ rows: [headerRow, ...rows], width: { size: 100, type: WidthType.PERCENTAGE } }),
        ],
      }],
    });
    const blob = await Packer.toBlob(doc);
    downloadBlob(blob, `opscore-report-${TODAY_STR}.docx`);
  };

  const exportPdf = () => {
    const docPdf = new jsPDF();
    docPdf.setFontSize(16);
    docPdf.text("OpsCore Custom Report", 14, 16);
    docPdf.setFontSize(10);
    docPdf.text(`Generated ${TODAY_STR} · ${results.length} result${results.length === 1 ? "" : "s"}`, 14, 22);
    autoTable(docPdf, {
      startY: 28,
      head: [["Title", "Campus", "Stage", "Owner", "Budget", "Spent", "Added", "Completed", "Latest Activity"]],
      body: results.map((p) => {
        const b = projectBudget(p);
        return [p.title, campusName(p.campus), p.stage, p.owner, fmtMoney(b.total), fmtMoney(b.spent), p.createdAt || "—", p.completedOn || "—", latestActivity(p) || "—"];
      }),
      styles: { fontSize: 8 },
      headStyles: { fillColor: [43, 76, 126] },
    });
    docPdf.save(`opscore-report-${TODAY_STR}.pdf`);
  };

  return (
    <div>
      <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(18px,3.6vw,22px)] font-semibold tracking-tight mb-1">Reports</h1>
      <p className="text-[12px] text-[#6B6980] mb-4">Filter across every campus, project, and task, then export the results.</p>

      <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg p-4 mb-5">
        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Campus</div>
            <div className="flex flex-wrap gap-1.5">
              {[...campuses].sort((a, b) => a.name.localeCompare(b.name)).map((c) => (
                <button key={c.id} onClick={() => toggleCampus(c.id)}
                  className={`text-[10.5px] px-2 py-1 rounded-full border ${selectedCampuses.includes(c.id) ? "text-white" : "text-[#6B6980]"}`}
                  style={selectedCampuses.includes(c.id) ? { background: c.color, borderColor: c.color } : { borderColor: "#D8D5EC" }}>
                  {c.abbr}
                </button>
              ))}
              <button onClick={() => toggleCampus("central")}
                className={`text-[10.5px] px-2 py-1 rounded-full border ${selectedCampuses.includes("central") ? "text-white" : "text-[#6B6980]"}`}
                style={selectedCampuses.includes("central") ? { background: "#6B4FA0", borderColor: "#6B4FA0" } : { borderColor: "#D8D5EC" }}>
                Central
              </button>
            </div>
          </div>

          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Staff / team member</div>
            <input value={staffFilter} onChange={(e) => setStaffFilter(e.target.value)} list="report-roster" placeholder="Name…"
              className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[12px] outline-none" />
            <datalist id="report-roster">
              {roster.map((p) => <option key={p.name} value={p.name} />)}
            </datalist>
          </div>

          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Project title contains</div>
            <input value={titleFilter} onChange={(e) => setTitleFilter(e.target.value)} placeholder="Search title…"
              className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2.5 py-1.5 text-[12px] outline-none" />
          </div>

          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Budget range ($)</div>
            <div className="flex items-center gap-1.5">
              <input type="number" value={costMin} onChange={(e) => setCostMin(e.target.value)} placeholder="Min" className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none" />
              <span className="text-[#8B889C]">–</span>
              <input type="number" value={costMax} onChange={(e) => setCostMax(e.target.value)} placeholder="Max" className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[12px] outline-none" />
            </div>
          </div>

          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Date added</div>
            <div className="flex items-center gap-1.5">
              <input type="date" value={addedFrom} onChange={(e) => setAddedFrom(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
              <span className="text-[#8B889C]">–</span>
              <input type="date" value={addedTo} onChange={(e) => setAddedTo(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
            </div>
          </div>

          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Date completed</div>
            <div className="flex items-center gap-1.5">
              <input type="date" value={completedFrom} onChange={(e) => setCompletedFrom(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
              <span className="text-[#8B889C]">–</span>
              <input type="date" value={completedTo} onChange={(e) => setCompletedTo(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
            </div>
          </div>

          <div>
            <div className="text-[10.5px] text-[#6B6980] mb-1.5">Latest activity</div>
            <div className="flex items-center gap-1.5">
              <input type="date" value={activityFrom} onChange={(e) => setActivityFrom(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
              <span className="text-[#8B889C]">–</span>
              <input type="date" value={activityTo} onChange={(e) => setActivityTo(e.target.value)} className="w-full bg-[#EFEEFA] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11px] outline-none" />
            </div>
          </div>

          <div className="flex items-end">
            <label className="flex items-center gap-2 text-[12px] text-[#2A2A3A] cursor-pointer">
              <input type="checkbox" checked={crossCampusOnly} onChange={(e) => setCrossCampusOnly(e.target.checked)} className="accent-[#B8862F]" />
              Cross-campus items only
            </label>
          </div>
        </div>

        <div className="flex items-center gap-2 mt-4 flex-wrap">
          <button onClick={clearFilters} className="text-[11.5px] text-[#6B6980] px-3 py-1.5">Clear filters</button>
          <div className="ml-auto flex gap-2">
            <button onClick={exportPdf} disabled={results.length === 0} className="flex items-center gap-1.5 text-[12px] rounded-md px-3 py-1.5 font-medium disabled:opacity-40" style={{ background: "#2B4C7E", color: "#F7F6FB" }}>
              <FileText size={13} /> Export PDF
            </button>
            <button onClick={exportDocx} disabled={results.length === 0} className="flex items-center gap-1.5 text-[12px] rounded-md px-3 py-1.5 font-medium disabled:opacity-40" style={{ background: "#B8862F", color: "#F7F6FB" }}>
              <FileText size={13} /> Export DOCX
            </button>
          </div>
        </div>
      </div>

      <div className="text-[11.5px] text-[#6B6980] mb-2">{results.length} result{results.length === 1 ? "" : "s"}</div>
      <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg overflow-x-auto">
        <table className="w-full text-[11.5px]">
          <thead>
            <tr className="border-b border-[#E3E1F0] text-left text-[#6B6980]">
              <th className="px-3 py-2 font-medium">Title</th>
              <th className="px-3 py-2 font-medium">Campus</th>
              <th className="px-3 py-2 font-medium">Stage</th>
              <th className="px-3 py-2 font-medium">Owner</th>
              <th className="px-3 py-2 font-medium">Budget</th>
              <th className="px-3 py-2 font-medium">Added</th>
              <th className="px-3 py-2 font-medium">Completed</th>
              <th className="px-3 py-2 font-medium">Latest Activity</th>
            </tr>
          </thead>
          <tbody>
            {results.map((p) => {
              const b = projectBudget(p);
              return (
                <tr key={p.id} className="border-b border-[#E3E1F0] last:border-0">
                  <td className="px-3 py-2 max-w-[200px] truncate">{p.title}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{campusName(p.campus)}</td>
                  <td className="px-3 py-2"><StageBadge stage={p.stage} /></td>
                  <td className="px-3 py-2 whitespace-nowrap">{p.owner}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{fmtMoney(b.spent)} / {fmtMoney(b.total)}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{p.createdAt || "—"}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{p.completedOn || "—"}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{latestActivity(p) || "—"}</td>
                </tr>
              );
            })}
            {results.length === 0 && (
              <tr><td colSpan={8} className="px-3 py-8 text-center text-[#8B889C]">No projects match these filters.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {isCentral && (
        <div className="mt-8 rounded-lg p-4" style={{ background: "#B8862F0F", border: "1px solid #B8862F55" }}>
          <div className="text-[12px] font-bold mb-1" style={{ color: "#8A6420" }}>Demo data</div>
          <p className="text-[11px] text-[#6B6980] mb-3">Loads sample projects, tasks, and staff into your Sheet — useful for testing. Only runs when you click this; the dashboard never seeds data automatically.</p>
          <button onClick={onLoadDemoData} className="text-[12px] rounded-md px-3 py-1.5 font-medium" style={{ background: "#B8862F", color: "#FFFFFF" }}>
            Load Demo Data
          </button>
        </div>
      )}

      {isCentral && (
        <div className="mt-4 rounded-lg p-4" style={{ background: "#C15B5B0F", border: "1px solid #C15B5B55" }}>
          <div className="text-[12px] font-bold mb-1" style={{ color: "#8A3A2E" }}>Danger zone</div>
          <p className="text-[11px] text-[#6B6980] mb-3">Permanently deletes every project and task from the Sheet and resets the dashboard to empty — use this once, when you're ready to stop using demo data and start tracking real work.</p>
          {!confirmingClear ? (
            <button onClick={() => setConfirmingClear(true)} className="text-[12px] rounded-md px-3 py-1.5 font-medium" style={{ background: "#C15B5B", color: "#FFFFFF" }}>
              Clear All Projects
            </button>
          ) : (
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-[12px] font-medium" style={{ color: "#8A3A2E" }}>Are you sure? This can't be undone.</span>
              <button
                disabled={clearing}
                onClick={async () => { setClearing(true); await onClearAllProjects(); setClearing(false); setConfirmingClear(false); }}
                className="text-[12px] rounded-md px-3 py-1.5 font-medium disabled:opacity-50" style={{ background: "#C15B5B", color: "#FFFFFF" }}>
                {clearing ? "Clearing…" : "Yes, delete everything"}
              </button>
              <button onClick={() => setConfirmingClear(false)} className="text-[12px] text-[#6B6980] px-3 py-1.5">Cancel</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// Central-only. Login accounts (Users sheet, real auth) are deliberately kept separate from
// the Staff directory (contact/assignment info, no login) — this panel is specifically about
// who can sign in and what campus/role their token carries, not who's on a campus team.
function AccountsPanel({ users, campuses, roleOptions, onCreate, onUpdateAccess, onRemove, currentUserId }) {
  const [showCreate, setShowCreate] = useState(false);
  const sorted = [...users].sort((a, b) => `${a.firstName} ${a.lastName}`.localeCompare(`${b.firstName} ${b.lastName}`));

  return (
    <div>
      <div className="flex items-start justify-between mb-6 gap-3 flex-wrap">
        <div>
          <h1 style={{ fontFamily: "'Fraunces', serif" }} className="text-[clamp(20px,4vw,26px)] font-semibold tracking-tight">User Management</h1>
          <p className="text-[13px] text-[#6B6980] mt-1">Create logins and control which campus and role each person's account has access to.</p>
        </div>
        <button onClick={() => setShowCreate(true)} className="flex items-center gap-1.5 text-[12.5px] font-medium rounded-md px-3 py-2" style={{ background: "#2B4C7E", color: "#F7F6FB" }}>
          <Plus size={14} /> New Account
        </button>
      </div>

      {sorted.length === 0 && <div className="text-[12.5px] text-[#8B889C] py-8 text-center">No accounts yet besides yours.</div>}

      <div className="space-y-2">
        {sorted.map((u) => (
          <AccountRow key={u.id} user={u} campuses={campuses} roleOptions={roleOptions}
            isSelf={u.id === currentUserId}
            onUpdateAccess={(campusId, role) => onUpdateAccess(u.id, campusId, role)}
            onRemove={() => onRemove(u.id)} />
        ))}
      </div>

      {showCreate && (
        <CreateAccountModal campuses={campuses} roleOptions={roleOptions}
          onClose={() => setShowCreate(false)}
          onCreate={(data) => { onCreate(data); setShowCreate(false); }} />
      )}
    </div>
  );
}

function AccountRow({ user, campuses, roleOptions, isSelf, onUpdateAccess, onRemove }) {
  const [confirmDelete, setConfirmDelete] = useState(false);
  const campusId = user.campusId || ""; // "" = not yet assigned
  const role = user.role || "";
  const tierLabel = campusId === "central" ? "central" : !campusId ? "unassigned" : (role === "Campus Operations Director" ? "od" : "staff");

  return (
    <div className="bg-[#FFFFFF] border border-[#E3E1F0] rounded-lg px-4 py-3 flex items-center justify-between gap-3 flex-wrap">
      <div className="min-w-0">
        <div className="text-[13.5px] font-medium truncate">{user.firstName} {user.lastName} {isSelf && <span className="text-[10.5px] text-[#8B889C] font-normal">(you)</span>}</div>
        <div className="text-[11.5px] text-[#6B6980] truncate">{user.email}</div>
      </div>
      <div className="flex items-center gap-2 flex-wrap">
        <select value={campusId} disabled={isSelf}
          onChange={(e) => onUpdateAccess(e.target.value, (!e.target.value || e.target.value === "central") ? "" : (role || roleOptions[0]))}
          className="bg-[#F7F6FB] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none disabled:opacity-60">
          <option value="">Unassigned</option>
          <option value="central">Central</option>
          {campuses.map((c) => <option key={c.id} value={c.id}>{c.name} ({c.abbr})</option>)}
        </select>
        {campusId && campusId !== "central" && (
          <select value={role} disabled={isSelf} onChange={(e) => onUpdateAccess(campusId, e.target.value)}
            className="bg-[#F7F6FB] border border-[#D8D5EC] rounded-md px-2 py-1.5 text-[11.5px] outline-none disabled:opacity-60">
            {roleOptions.map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        )}
        <span className="text-[10px] px-2 py-1 rounded-full" style={{ background: tierLabel === "unassigned" ? "#B8862F22" : "#E3E1F0", color: tierLabel === "unassigned" ? "#B8862F" : "#6B6980" }}>
          {tierLabel} tier
        </span>
        {!isSelf && (
          confirmDelete ? (
            <span className="flex items-center gap-1">
              <button onClick={onRemove} className="text-[11px] px-2 py-1 rounded-md font-medium" style={{ background: "#C15B5B", color: "#FFFFFF" }}>Confirm</button>
              <button onClick={() => setConfirmDelete(false)} className="text-[11px] text-[#6B6980] px-1">Cancel</button>
            </span>
          ) : (
            <button onClick={() => setConfirmDelete(true)} className="text-[#8B889C] hover:text-[#C15B5B] p-1"><Trash2 size={14} /></button>
          )
        )}
      </div>
    </div>
  );
}

function CreateAccountModal({ campuses, roleOptions, onClose, onCreate }) {
  useLockBodyScroll();
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [email, setEmail] = useState("");
  const [phone, setPhone] = useState("");
  const [campusId, setCampusId] = useState("central");
  const [role, setRole] = useState(roleOptions[0] || "");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const inputClass = "w-full bg-[#F7F6FB] border border-[#D8D5EC] rounded-md px-3 py-2 text-[13px] outline-none focus:border-[#2B4C7E]";
  const labelClass = "text-[11px] font-medium text-[#6B6980] mb-1 block";

  const submit = (e) => {
    e.preventDefault();
    if (password.length < 8) { setError("Password must be at least 8 characters."); return; }
    onCreate({ firstName, lastName, email, phone, campusId, role: campusId === "central" ? "" : role, password });
  };

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center p-4" style={{ background: "rgba(42,42,58,0.45)" }}>
      <div className="bg-[#FFFFFF] rounded-xl p-6 w-full max-w-[420px] max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-[15px] font-semibold">New Account</h2>
          <button onClick={onClose} className="text-[#8B889C] hover:text-[#2A2A3A]"><X size={16} /></button>
        </div>
        <form onSubmit={submit} className="space-y-3">
          <div className="grid grid-cols-2 gap-2">
            <div><label className={labelClass}>First name</label><input required value={firstName} onChange={(e) => setFirstName(e.target.value)} className={inputClass} /></div>
            <div><label className={labelClass}>Last name</label><input required value={lastName} onChange={(e) => setLastName(e.target.value)} className={inputClass} /></div>
          </div>
          <div><label className={labelClass}>Email</label><input required type="email" value={email} onChange={(e) => setEmail(e.target.value)} className={inputClass} /></div>
          <div><label className={labelClass}>Phone (optional)</label><input value={phone} onChange={(e) => setPhone(e.target.value)} className={inputClass} /></div>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className={labelClass}>Campus access</label>
              <select value={campusId} onChange={(e) => setCampusId(e.target.value)} className={inputClass}>
                <option value="central">Central</option>
                {campuses.map((c) => <option key={c.id} value={c.id}>{c.name} ({c.abbr})</option>)}
              </select>
            </div>
            {campusId !== "central" && (
              <div>
                <label className={labelClass}>Role</label>
                <select value={role} onChange={(e) => setRole(e.target.value)} className={inputClass}>
                  {roleOptions.map((r) => <option key={r} value={r}>{r}</option>)}
                </select>
              </div>
            )}
          </div>
          <div>
            <label className={labelClass}>Temporary password</label>
            <input required type="password" minLength={8} value={password} onChange={(e) => setPassword(e.target.value)} className={inputClass} />
          </div>
          {error && <div className="text-[12px] rounded-md px-3 py-2" style={{ background: "#C15B5B1A", color: "#C15B5B" }}>{error}</div>}
          <div className="flex gap-2 pt-1">
            <button type="submit" className="text-[13px] font-medium rounded-md px-4 py-2" style={{ background: "#2B4C7E", color: "#F7F6FB" }}>Create Account</button>
            <button type="button" onClick={onClose} className="text-[13px] text-[#6B6980] px-2">Cancel</button>
          </div>
        </form>
      </div>
    </div>
  );
}

function StaffProfileModal({ name, projects, onClose, onOpenProject }) {
  useLockBodyScroll();
  if (!name) return null;

  const isInvolved = (p) => p.owner === name || (p.team || []).includes(name) || (p.subtasks || []).some((s) => s.createdBy === name);

  const involved = projects.filter(isInvolved);
  const completedProjects = involved.filter((p) => p.stage === "Completed").sort((a, b) => (b.completedOn || "").localeCompare(a.completedOn || "")).slice(0, 8);
  const completedTasks = [];
  involved.forEach((p) => (p.subtasks || []).forEach((s) => { if (s.done && s.createdBy === name) completedTasks.push({ ...s, projectTitle: p.title, projectId: p.id }); }));
  completedTasks.sort((a, b) => (b.due || "").localeCompare(a.due || ""));

  const current = involved.filter((p) => p.stage !== "Completed");
  const needsAttention = current.filter((p) => p.stage === "Stalled" || (p.due && p.due < TODAY_STR) || (p.subtasks || []).some((s) => !s.done && s.due && s.due < TODAY_STR));

  return (
    <div className="fixed inset-0 z-30 overflow-y-auto p-4" style={{ background: "rgba(0,0,0,0.65)", WebkitOverflowScrolling: "touch" }} onClick={onClose}>
      <div className="border rounded-xl max-w-[560px] w-full my-8 mx-auto p-5" style={{ background: "#FFFFFF", borderColor: "#D8D5EC", color: "#2A2A3A", fontFamily: "'Inter', sans-serif" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between mb-4">
          <h2 className="text-[17px] font-bold">{name}</h2>
          <button onClick={onClose} className="text-[#6B6980] hover:text-[#2A2A3A]"><X size={18} /></button>
        </div>

        {needsAttention.length > 0 && (
          <div className="rounded-md p-3 mb-4" style={{ background: "#C15B5B18", border: "1px solid #C15B5B55" }}>
            <div className="text-[11.5px] font-medium mb-2" style={{ color: "#8A3A2E" }}>Needs attention</div>
            <div className="space-y-1.5">
              {needsAttention.map((p) => (
                <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left text-[12px] bg-[#FFFFFF] rounded-md px-2.5 py-1.5 flex items-center justify-between" style={{ border: "1px solid #E3E1F0" }}>
                  <span className="truncate">{p.title}</span>
                  <StageBadge stage={p.stage} />
                </button>
              ))}
            </div>
          </div>
        )}

        <div className="mb-4">
          <div className="text-[11px] text-[#6B6980] mb-2 font-medium">Current projects ({current.length})</div>
          <div className="space-y-1.5">
            {current.map((p) => (
              <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left text-[12px] bg-[#EFEEFA] rounded-md px-2.5 py-2 flex items-center justify-between gap-2 hover:border-[#B8862F]" style={{ border: "1px solid #E3E1F0" }}>
                <span className="truncate">{p.title}</span>
                <StageBadge stage={p.stage} />
              </button>
            ))}
            {current.length === 0 && <div className="text-[11.5px] text-[#8B889C]">Nothing currently in progress.</div>}
          </div>
        </div>

        <div className="mb-2">
          <div className="text-[11px] text-[#6B6980] mb-2 font-medium">Recently completed</div>
          <div className="space-y-1.5">
            {completedProjects.map((p) => {
              const status = budgetStatus(p);
              return (
                <button key={p.id} onClick={() => onOpenProject(p.id)} className="w-full text-left text-[12px] bg-[#FFFFFF] rounded-md px-2.5 py-2 flex items-center justify-between gap-2" style={{ border: "1px solid #E3E1F0" }}>
                  <span className="truncate">{p.title} <span className="text-[10px] text-[#8B889C]">· project · {p.completedOn}</span></span>
                  <span className="text-[9.5px] px-1.5 py-0.5 rounded-full shrink-0" style={{ background: `${BUDGET_STATUS_COLOR[status]}22`, color: BUDGET_STATUS_COLOR[status] }}>{BUDGET_STATUS_LABEL[status]}</span>
                </button>
              );
            })}
            {completedTasks.slice(0, 8).map((s, i) => (
              <button key={i} onClick={() => onOpenProject(s.projectId)} className="w-full text-left text-[12px] bg-[#FFFFFF] rounded-md px-2.5 py-2" style={{ border: "1px solid #E3E1F0" }}>
                {s.t} <span className="text-[10px] text-[#8B889C]">· task on {s.projectTitle}</span>
              </button>
            ))}
            {completedProjects.length === 0 && completedTasks.length === 0 && <div className="text-[11.5px] text-[#8B889C]">Nothing completed yet.</div>}
          </div>
        </div>
      </div>
    </div>
  );
}
