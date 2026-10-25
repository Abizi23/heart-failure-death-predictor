
/* GANTT AI ASSISTANT - Outlook task pane */

const $ = (id) => document.getElementById(id);
const API_BASE = "https://localhost:8000";

let currentEmail = {
  subject: "",
  sender: "",
  body: ""
};

let events = [];
let nextEventId = 1;

const DASHBOARD_URL = "http://localhost:8501";

// Open the Streamlit dashboard — tries every available method
function openDashboard() {
  // Method 1: Official Office.js API (works in modern Outlook)
  if (
    Office.context &&
    Office.context.ui &&
    typeof Office.context.ui.openBrowserWindow === "function"
  ) {
    Office.context.ui.openBrowserWindow(DASHBOARD_URL);
    return;
  }
  // Method 2: Plain window.open fallback
  window.open(DASHBOARD_URL, "_blank");
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    $("status").textContent = "Open this add-in inside Outlook.";
    return;
  }

  $("analyze-btn").addEventListener("click", analyzeEmail);
  $("dashboard-btn").addEventListener("click", openDashboard);

  loadSelectedEmail();
  await loadSavedEvents();
});

// --------------------------------------------------
// Load the currently selected Outlook email
// --------------------------------------------------

function loadSelectedEmail() {
  const item = Office.context.mailbox?.item;

  if (!item) {
    currentEmail = { subject: "", sender: "", body: "" };
    $("email-subject").textContent = "No email selected";
    $("email-sender").textContent = "";
    $("email-summary").textContent = "Open an email to begin.";
    return;
  }

  currentEmail.subject = item.subject || "(No subject)";

  currentEmail.sender = item.from
    ? `${item.from.displayName || ""} ${item.from.emailAddress
      ? `<${item.from.emailAddress}>`
      : ""
      }`.trim()
    : "Sender unavailable";

  $("email-subject").textContent = currentEmail.subject;
  $("email-sender").textContent = `From: ${currentEmail.sender}`;
  $("email-summary").textContent = "Reading email...";

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      $("status").textContent = "Could not read this email's body.";
      $("email-summary").textContent = "Email body unavailable.";
      return;
    }

    currentEmail.body = normalizeText(result.value || "");

    $("email-summary").textContent =
      "Email loaded. Click Analyze this email to request AI analysis.";

    $("status").textContent = "Email loaded. Ready for AI analysis.";
  });
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// --------------------------------------------------
// Backend request helper
// --------------------------------------------------

async function apiRequest(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {})
    }
  });

  const text = await response.text();
  let data = {};

  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error(`Unexpected server response (${response.status}).`);
    }
  }

  if (!response.ok) {
    throw new Error(
      typeof data.detail === "string"
        ? data.detail
        : `Request failed (${response.status})`
    );
  }

  return data;
}

// --------------------------------------------------
// Convert a database record into a table event
// --------------------------------------------------

function mapDatabaseEvent(record) {
  return {
    id: nextEventId++,
    databaseId: record.id,
    eventType: record.event_type || "task",
    task: record.task || "",
    startDate: record.start_date || "",
    startTime: record.start_time || "",
    endDate: record.end_date || "",
    endTime: record.end_time || "",
    hours: record.hours == null ? "" : String(record.hours),
    project: record.project || "",
    resources: record.resources || "",
    location: record.location || "",
    aiWarnings: Array.isArray(record.warnings)
      ? [...record.warnings]
      : [],
    status: record.status || "Proposed",
    emailSubject: record.email_subject || "",
    saving: false
  };
}

// --------------------------------------------------
// Load saved events when Outlook opens
// --------------------------------------------------

async function loadSavedEvents() {
  $("status").textContent = "Loading saved events...";

  try {
    const records = await apiRequest("/events");

    if (!Array.isArray(records)) {
      throw new Error("The server returned an invalid events list.");
    }

    events = records.map(mapDatabaseEvent);
    renderEvents();

    $("status").textContent = events.length
      ? `Loaded ${events.length} saved event(s) from the database.`
      : "No saved events yet. Analyze an email to begin.";

  } catch (error) {
    console.error("Could not load saved events:", error);
    $("status").textContent =
      `Could not load saved events: ${error.message}`;
  }
}

// --------------------------------------------------
// Format an event for saving
// --------------------------------------------------

function eventPayload(event, status = event.status) {
  return {
    event_type: event.eventType,
    task: event.task,
    start_date: event.startDate,
    start_time: event.startTime,
    end_date: event.endDate,
    end_time: event.endTime,
    hours: event.hours === "" ? null : Number(event.hours),
    project: event.project,
    resources: event.resources,
    location: event.location,
    warnings: event.aiWarnings,
    status,
    email_subject: event.emailSubject || currentEmail.subject
  };
}

// --------------------------------------------------
// Analyze an email and save proposed events
// --------------------------------------------------

async function analyzeEmail() {
  if (!currentEmail.subject && !currentEmail.body) {
    $("status").textContent =
      "No email content is available. Open an email first.";
    return;
  }

  const button = $("analyze-btn");
  button.disabled = true;
  button.textContent = "Analyzing email...";

  $("status").textContent = "AI is analyzing the email...";
  $("email-summary").textContent = "Please wait...";

  try {
    const result = await apiRequest("/analyze-email", {
      method: "POST",
      body: JSON.stringify({
        subject: currentEmail.subject,
        body: currentEmail.body
      })
    });

    if (
      typeof result.summary !== "string" ||
      !Array.isArray(result.events)
    ) {
      throw new Error("Unexpected response format from the AI.");
    }

    $("email-summary").textContent = result.summary;

    let savedCount = 0;

    for (const record of result.events) {
      const event = {
        id: nextEventId++,
        databaseId: null,
        eventType: record.event_type || "task",
        task: record.task || "",
        startDate: record.start_date || "",
        startTime: record.start_time || "",
        endDate: record.end_date || "",
        endTime: record.end_time || "",
        hours: record.hours == null ? "" : String(record.hours),
        project: record.project || "",
        resources: record.resources || "",
        location: record.location || "",
        aiWarnings: Array.isArray(record.warnings)
          ? [...record.warnings]
          : [],
        status: "Proposed",
        emailSubject: currentEmail.subject,
        saving: false
      };

      events.unshift(event);
      renderEvents();

      try {
        const saved = await apiRequest("/events", {
          method: "POST",
          body: JSON.stringify(eventPayload(event, "Proposed"))
        });

        event.databaseId = saved.id;
        event.status = saved.status || "Proposed";
        event.aiWarnings = Array.isArray(saved.warnings)
          ? [...saved.warnings]
          : event.aiWarnings;

        savedCount++;
      } catch (error) {
        console.error("Could not save event:", error);
        event.aiWarnings.push("Database save failed");
      }

      renderEvents();
    }

    if (result.events.length === 0) {
      $("status").textContent =
        "No clear project events were identified.";
    } else {
      $("status").textContent =
        `AI identified ${result.events.length} event(s); ` +
        `${savedCount} saved to the database. Opening dashboard...`;

      // Automatically open the dashboard so the user sees the new events
      openDashboard();
    }

  } catch (error) {
    console.error("Email analysis failed:", error);
    $("email-summary").textContent = "The email could not be analyzed.";
    $("status").textContent = `Analysis failed: ${error.message}`;

  } finally {
    button.disabled = false;
    button.textContent = "Analyze this email";
  }
}

// --------------------------------------------------
// Warnings
// --------------------------------------------------

function getWarnings(event) {
  const missing = [];

  if (!event.eventType.trim()) missing.push("event type");
  if (!event.task.trim()) missing.push("event name");
  if (!event.startDate) missing.push("start date");

  if (
    event.startDate &&
    event.endDate &&
    event.endDate < event.startDate
  ) {
    missing.push("end date is earlier than start date");
  }

  return missing;
}

function updateWarningCell(row, event) {
  const warningCell = row.cells[0];
  const missing = getWarnings(event);

  const allWarnings = [
    ...new Set([...missing, ...event.aiWarnings])
  ];

  warningCell.className =
    `warning ${allWarnings.length ? "" : "clear"}`;

  warningCell.textContent = allWarnings.length ? "⚠" : "✓";
  warningCell.title = allWarnings.length
    ? `Review: ${allWarnings.join(", ")}`
    : "No warnings detected";
}

// --------------------------------------------------
// Render the event table
// --------------------------------------------------

function renderEvents() {
  const tbody = $("events-body");
  tbody.replaceChildren();

  $("event-count").textContent =
    `${events.length} event${events.length === 1 ? "" : "s"} found`;

  if (events.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");

    cell.colSpan = 12;
    cell.className = "empty";
    cell.textContent = "No saved events. Analyze an email to begin.";

    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  for (const event of events) {
    const row = document.createElement("tr");

    const warningCell = document.createElement("td");
    warningCell.className = "warning";
    warningCell.textContent = "⚠";
    row.appendChild(warningCell);

    addSelectCell(row, event, "eventType", "Event type", [
      { value: "meeting", label: "Meeting" },
      { value: "task", label: "Task" },
      { value: "milestone", label: "Milestone" },
      { value: "schedule_change", label: "Schedule change" }
    ]);

    addInputCell(row, event, "task", "Event / task", true);
    addInputCell(row, event, "startDate", "Start date", false, "date");
    addInputCell(row, event, "startTime", "Start time", false, "time");
    addInputCell(row, event, "endDate", "End date", false, "date");
    addInputCell(row, event, "endTime", "End time", false, "time");
    addInputCell(row, event, "hours", "Hours", false, "number");
    addInputCell(row, event, "project", "Project name", true);
    addInputCell(row, event, "resources", "Resources", true);
    addInputCell(row, event, "location", "Location", true);

    updateWarningCell(row, event);

    const decisionCell = document.createElement("td");
    const buttons = document.createElement("div");
    buttons.className = "decision-buttons";

    const accept = document.createElement("button");
    accept.className = "accept";
    accept.textContent = event.saving ? "Saving..." : "Accept";
    accept.disabled = event.saving;
    accept.addEventListener("click", () => {
      decideEvent(event.id, "Accepted");
    });

    const reject = document.createElement("button");
    reject.className = "reject";
    reject.textContent = event.saving ? "Saving..." : "Reject";
    reject.disabled = event.saving;
    reject.addEventListener("click", () => {
      decideEvent(event.id, "Rejected");
    });

    const rowStatus = document.createElement("div");
    rowStatus.className = "row-status";
    rowStatus.textContent = event.status;

    buttons.append(accept, reject, rowStatus);
    decisionCell.appendChild(buttons);
    row.appendChild(decisionCell);

    tbody.appendChild(row);
  }
}

function addInputCell(
  row,
  event,
  field,
  label,
  wide = false,
  type = "text"
) {
  const cell = document.createElement("td");
  const input = document.createElement("input");

  input.type = type;
  input.value = event[field] || "";
  input.setAttribute("aria-label", label);
  input.title = label;

  if (wide) input.classList.add("wide");

  if (type === "number") {
    input.min = "0";
    input.step = "0.5";
  }

  input.addEventListener("change", () => {
    event[field] = input.value.trim();
    updateWarningCell(row, event);
  });

  cell.appendChild(input);
  row.appendChild(cell);
}

function addSelectCell(row, event, field, label, options) {
  const cell = document.createElement("td");
  const select = document.createElement("select");

  select.setAttribute("aria-label", label);
  select.title = label;

  for (const item of options) {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    select.appendChild(option);
  }

  select.value = options.some(
    (item) => item.value === event[field]
  ) ? event[field] : "task";

  select.addEventListener("change", () => {
    event[field] = select.value;
    updateWarningCell(row, event);
  });

  cell.appendChild(select);
  row.appendChild(cell);
}

// --------------------------------------------------
// Save Accept / Reject decisions to SQLite
// --------------------------------------------------

async function decideEvent(id, decision) {
  const event = events.find((item) => item.id === id);
  if (!event || event.saving) return;

  if (decision === "Accepted") {
    const missing = getWarnings(event);

    if (missing.length) {
      $("status").textContent =
        `Please fix these fields before accepting: ${missing.join(", ")}.`;
      return;
    }
  }

  if (event.databaseId == null) {
    $("status").textContent =
      "This event has not been saved. Check the backend and try again.";
    return;
  }

  event.saving = true;
  renderEvents();

  $("status").textContent =
    `Saving event as ${decision.toLowerCase()}...`;

  try {
    const saved = await apiRequest(
      `/events/${event.databaseId}`,
      {
        method: "PATCH",
        body: JSON.stringify(eventPayload(event, decision))
      }
    );

    event.status = saved.status;
    event.aiWarnings = Array.isArray(saved.warnings)
      ? [...saved.warnings]
      : event.aiWarnings;

    event.saving = false;
    renderEvents();

    $("status").textContent =
      `Event "${event.task}" saved as ${decision}. Opening dashboard...`;

    // Open dashboard so the user sees the updated status immediately
    openDashboard();

  } catch (error) {
    event.saving = false;
    renderEvents();

    $("status").textContent =
      `Could not save decision: ${error.message}`;
  }
}
