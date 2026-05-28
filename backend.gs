// ── CONFIG ──────────────────────────────────────────────────
const SHEET_NAME   = "Goals";
const GROQ_API_KEY = "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";     
const GROQ_MODEL   = "llama-3.3-70b-versatile";
const GROQ_URL     = "https://api.groq.com/openai/v1/chat/completions";

// Column index constants (1-based)
const COL = {
  ID:            1,
  NAME:          2,
  TARGET:        3,
  SAVED:         4,
  DEADLINE:      5,
  PURPOSE:       6,
  CREATED_AT:    7,
  EMAIL:         8,
  ALERT_SENT:    9
};

// ── ENTRY POINT ─────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("💰 AI Savings Goal Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── SHEET HELPERS ────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "ID", "Goal Name", "Target Amount", "Saved Amount",
      "Deadline", "Purpose", "Created At", "Email", "Alert Sent"
    ]);
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#1a1a2e");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateId() {
  return "GOAL_" + new Date().getTime();
}

// ── CRUD OPERATIONS ──────────────────────────────────────────

/**
 * Add a new savings goal to the sheet.
 * Called from the frontend via google.script.run
 */
function addGoal(goalData) {
  try {
    const sheet = getSheet();
    const id = generateId();
    const now = new Date().toISOString();
    sheet.appendRow([
      id,
      goalData.name,
      parseFloat(goalData.target),
      parseFloat(goalData.initialSaved) || 0,
      goalData.deadline,
      goalData.purpose || "",
      now,
      goalData.email || "",
      false
    ]);
    return { success: true, id: id, message: "Goal created successfully!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Fetch all goals from the sheet and return as JSON array.
 */
function getAllGoals() {
  try {
    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, goals: [] };

    const goals = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[COL.ID - 1]) continue;           // skip empty rows

      const target   = parseFloat(row[COL.TARGET - 1])  || 0;
      const saved    = parseFloat(row[COL.SAVED - 1])   || 0;
      const progress = target > 0 ? Math.min((saved / target) * 100, 100) : 0;
      const deadline = row[COL.DEADLINE - 1];
      const daysLeft = deadline
        ? Math.ceil((new Date(deadline) - new Date()) / (1000 * 60 * 60 * 24))
        : null;

      goals.push({
        rowIndex: i + 1,          // 1-based sheet row
        id:       row[COL.ID - 1],
        name:     row[COL.NAME - 1],
        target:   target,
        saved:    saved,
        deadline: deadline ? new Date(deadline).toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" }) : "",
        deadlineRaw: deadline ? new Date(deadline).toISOString().split("T")[0] : "",
        purpose:  row[COL.PURPOSE - 1],
        createdAt:row[COL.CREATED_AT - 1],
        email:    row[COL.EMAIL - 1],
        progress: Math.round(progress),
        daysLeft: daysLeft,
        status:   progress >= 100 ? "completed" : (daysLeft !== null && daysLeft < 0 ? "overdue" : "active")
      });
    }
    return { success: true, goals: goals };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Update (add to) the saved amount for a specific goal row.
 */
function updateSavings(rowIndex, additionalAmount) {
  try {
    const sheet   = getSheet();
    const cell    = sheet.getRange(rowIndex, COL.SAVED);
    const current = parseFloat(cell.getValue()) || 0;
    const newVal  = current + parseFloat(additionalAmount);
    cell.setValue(newVal);
    return { success: true, newSaved: newVal, message: "Savings updated!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Delete a goal row by sheet row index.
 */
function deleteGoal(rowIndex) {
  try {
    getSheet().deleteRow(rowIndex);
    return { success: true, message: "Goal deleted." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ── AI INSIGHTS (GROQ) ───────────────────────────────────────

/**
 * Call Groq AI to generate personalised savings advice.
 * @param {object} goalInfo  { name, target, saved, deadline, purpose, daysLeft }
 */
function getAIInsights(goalInfo) {
  try {
    const progress = goalInfo.target > 0
      ? ((goalInfo.saved / goalInfo.target) * 100).toFixed(1)
      : 0;
    const remaining = goalInfo.target - goalInfo.saved;

    const prompt = `You are a professional financial advisor specializing in personal savings goals.
A user has the following savings goal:
- Goal Name: ${goalInfo.name}
- Purpose: ${goalInfo.purpose || "Not specified"}
- Target Amount: ₹${goalInfo.target.toLocaleString("en-IN")}
- Amount Saved: ₹${goalInfo.saved.toLocaleString("en-IN")} (${progress}% complete)
- Remaining: ₹${remaining.toLocaleString("en-IN")}
- Deadline: ${goalInfo.deadline || "No deadline"}
- Days Left: ${goalInfo.daysLeft !== null ? goalInfo.daysLeft + " days" : "No deadline"}

Provide:
1. A brief assessment of their current progress (1-2 sentences)
2. Three specific, actionable savings tips tailored to this goal
3. A motivational closing statement

Keep the response concise (under 200 words), practical, and encouraging.
Format with clear numbered points.`;

    const payload = {
      model: GROQ_MODEL,
      messages: [
        { role: "system", content: "You are a concise, encouraging financial advisor. Provide practical, personalised savings advice." },
        { role: "user",   content: prompt }
      ],
      max_tokens: 400,
      temperature: 0.7
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + GROQ_API_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(GROQ_URL, options);
    const json     = JSON.parse(response.getContentText());

    if (json.choices && json.choices[0]) {
      return {
        success: true,
        insight: json.choices[0].message.content.trim()
      };
    } else {
      return { success: false, message: "Groq returned no choices. Check API key & model." };
    }
  } catch (e) {
    return { success: false, message: "AI Error: " + e.toString() };
  }
}

// ── AUTOMATED DAILY AUDIT TRIGGER ────────────────────────────
// testing
/**
 * dailyAuditTrigger — Set this as a Time-Driven trigger (daily).
 * Scans all goals; sends email alert if:
 *   - Deadline is within 7 days AND
 *   - Goal is not yet met (saved < target) AND
 *   - Alert hasn't been sent already
 */
function dailyAuditTrigger() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const id        = row[COL.ID - 1];
    if (!id) continue;

    const name      = row[COL.NAME - 1];
    const target    = parseFloat(row[COL.TARGET - 1]) || 0;
    const saved     = parseFloat(row[COL.SAVED - 1])  || 0;
    const deadline  = row[COL.DEADLINE - 1] ? new Date(row[COL.DEADLINE - 1]) : null;
    const email     = row[COL.EMAIL - 1];
    const alertSent = row[COL.ALERT_SENT - 1];

    if (!deadline || !email || alertSent === true || saved >= target) continue;

    deadline.setHours(0, 0, 0, 0);
    const daysLeft = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));

    if (daysLeft >= 0 && daysLeft <= 7) {
      const progress   = target > 0 ? ((saved / target) * 100).toFixed(1) : 0;
      const remaining  = target - saved;
      const dailyNeeded = daysLeft > 0 ? (remaining / daysLeft).toFixed(0) : remaining;

      const subject = `⚠️ Savings Alert: "${name}" deadline in ${daysLeft} day(s)!`;
      const body = `
Hi there,

This is an automated alert from your AI Savings Goal Tracker.

  Goal: ${name}
  Target: ₹${target.toLocaleString("en-IN")}
  Saved: ₹${saved.toLocaleString("en-IN")} (${progress}%)
  Remaining: ₹${remaining.toLocaleString("en-IN")}
  Deadline: ${deadline.toLocaleDateString("en-IN", { day:"2-digit", month:"long", year:"numeric" })}
  Days Left: ${daysLeft}

${daysLeft > 0
  ? ` To meet your goal, you need to save approximately ₹${dailyNeeded}/day for the next ${daysLeft} day(s).`
  : ` Your deadline is TODAY! Make a final contribution to reach your goal.`
}

Stay disciplined and keep going — every rupee counts!

Best regards,
AI Savings Goal Tracker 
      `.trim();

      try {
        GmailApp.sendEmail(email, subject, body);
        sheet.getRange(i + 1, COL.ALERT_SENT).setValue(true);
        Logger.log(`Alert sent to ${email} for goal: ${name}`);
      } catch (mailErr) {
        Logger.log(`Failed to send email for ${name}: ${mailErr}`);
      }
    }
  }
  Logger.log("dailyAuditTrigger completed: " + new Date().toISOString());
}

/**
 * Helper: Manually install the daily trigger (run once from GAS editor).
 */
function installDailyTrigger() {
  // Remove any existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "dailyAuditTrigger") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("dailyAuditTrigger")
    .timeBased()
    .everyDays(1)
    .atHour(8)        // 8 AM every day
    .create();
  Logger.log("Daily audit trigger installed successfully.");
}

// ── DASHBOARD SUMMARY ─────────────────────────────────────────

/**
 * Returns aggregate statistics for the dashboard header.
 */
function getDashboardStats() {
  try {
    const result = getAllGoals();
    if (!result.success) return result;

    const goals = result.goals;
    let totalTarget = 0, totalSaved = 0, completed = 0, active = 0, overdue = 0;

    goals.forEach(g => {
      totalTarget += g.target;
      totalSaved  += g.saved;
      if (g.status === "completed") completed++;
      else if (g.status === "overdue") overdue++;
      else active++;
    });

    return {
      success:      true,
      totalGoals:   goals.length,
      totalTarget:  totalTarget,
      totalSaved:   totalSaved,
      completedGoals: completed,
      activeGoals:  active,
      overdueGoals: overdue,
      overallProgress: totalTarget > 0 ? Math.round((totalSaved / totalTarget) * 100) : 0
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
