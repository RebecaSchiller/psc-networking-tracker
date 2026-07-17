function doGet(e) { return handleRequest(e); }

function doPost(e) {
  var p = JSON.parse(e.postData.contents);
  return handleAction(p);
}

function handleRequest(e) {
  var p = e.parameter;
  return handleAction(p);
}

function authorizeCalendar() {
  CalendarApp.getDefaultCalendar();
}

function installStaleFollowUpTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkStaleFollowUps") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("checkStaleFollowUps")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}

function checkStaleFollowUps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Contacts");
  var rows = sheet.getDataRange().getValues();
  var staleDays = 7;
  var now = new Date();
  var stale = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var name = row[1];
    var company = row[2];
    var status = row[4];
    var addedAt = row[7];
    var tags = row[9] || "";

    if (status !== "Follow up Email") continue;

    var changedDate = addedAt;
    var match = String(tags).match(/statuschanged:([0-9]{4}-[0-9]{2}-[0-9]{2})/);
    if (match) changedDate = match[1];
    if (!changedDate) continue;

    var days = Math.floor((now - new Date(changedDate)) / 86400000);
    if (days >= staleDays) {
      stale.push({ name: name, company: company, days: days });
    }
  }

  if (stale.length === 0) return;

  var lines = stale.map(function(c) {
    return "- " + c.name + (c.company ? " (" + c.company + ")" : "") +
      " \u2014 " + c.days + " days in Follow up Email";
  });

  var body = "These contacts have been sitting in \"Follow up Email\" status for 7+ days:\n\n" +
    lines.join("\n") +
    "\n\nOpen Network Momentum: https://networkmomentum.app";

  MailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    "Network Momentum: " + stale.length + " stale follow-up" + (stale.length > 1 ? "s" : ""),
    body
  );
}

function installWeeklyDigestTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendWeeklyDigest") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("sendWeeklyDigest")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
}

function sendWeeklyDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Contacts");
  var rows = sheet.getDataRange().getValues();
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 86400000);

  var newThisWeek = [];
  var overdue = [];
  var stuckFollowUp = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var name = row[1];
    var company = row[2];
    var status = row[4];
    var lastContact = row[5];
    var addedAt = row[7];
    var tags = row[9] || "";

    if (addedAt) {
      var addedDate = new Date(addedAt);
      if (addedDate >= weekAgo) {
        newThisWeek.push({ name: name, company: company });
      }
    }

    if (lastContact && status !== "Closed" && status !== "Follow-up needed" && status !== "Not contacted") {
      var daysSinceContact = Math.floor((now - new Date(lastContact)) / 86400000);
      if (daysSinceContact >= 7) {
        overdue.push({ name: name, company: company, days: daysSinceContact });
      }
    }

    if (status === "Follow up Email") {
      var changedDate = addedAt;
      var match = String(tags).match(/statuschanged:([0-9]{4}-[0-9]{2}-[0-9]{2})/);
      if (match) changedDate = match[1];
      if (changedDate) {
        var daysInStatus = Math.floor((now - new Date(changedDate)) / 86400000);
        if (daysInStatus >= 5) {
          stuckFollowUp.push({ name: name, company: company, days: daysInStatus });
        }
      }
    }
  }

  if (newThisWeek.length === 0 && overdue.length === 0 && stuckFollowUp.length === 0) return;

  var body = "Your Network Momentum weekly digest:\n";

  body += "\nNew contacts this week (" + newThisWeek.length + "):\n";
  if (newThisWeek.length === 0) {
    body += "- None\n";
  } else {
    newThisWeek.forEach(function(c) {
      body += "- " + c.name + (c.company ? " (" + c.company + ")" : "") + "\n";
    });
  }

  body += "\nOverdue, 7+ days since last contact (" + overdue.length + "):\n";
  if (overdue.length === 0) {
    body += "- None\n";
  } else {
    overdue.forEach(function(c) {
      body += "- " + c.name + (c.company ? " (" + c.company + ")" : "") + " \u2014 " + c.days + "d\n";
    });
  }

  body += "\nStuck in Follow up Email, 5+ days (" + stuckFollowUp.length + "):\n";
  if (stuckFollowUp.length === 0) {
    body += "- None\n";
  } else {
    stuckFollowUp.forEach(function(c) {
      body += "- " + c.name + (c.company ? " (" + c.company + ")" : "") + " \u2014 " + c.days + "d\n";
    });
  }

  body += "\nOpen Network Momentum: https://networkmomentum.app";

  MailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    "Network Momentum weekly digest",
    body
  );
}

function handleAction(p) {
  var out = {};

  if (p.action === "headers") {
    out = { headers: ["ID","Name","Company","Role","Status",
      "LastContact","Notes","AddedAt","Priority","Tags",
      "LinkedIn","NotesHistory","Email"] };

  } else if (p.action === "read") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Contacts");
    var rows = sheet.getDataRange().getValues();
    out = { rows: rows };

  } else if (p.action === "append") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Contacts");
    var row = JSON.parse(p.row);
    sheet.appendRow(row);
    out = { success: true };

  } else if (p.action === "append_batch") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Contacts");
    var newRows = JSON.parse(p.rows);
    if (!newRows || !newRows.length) {
      out = { error: "No rows provided" };
    } else {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      out = { success: true, count: newRows.length };
    }

  } else if (p.action === "update") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Contacts");
    sheet.getRange(parseInt(p.row), parseInt(p.col)).setValue(p.value);
    out = { success: true };

  } else if (p.action === "clear") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Contacts");
    sheet.getRange(parseInt(p.row), 1, 1, 13).clearContent();
    out = { success: true };

  } else if (p.action === "create_event") {
    var title = p.title || "Follow up";
    var description = p.description || "";
    var start = new Date(p.start);
    var minutes = parseInt(p.duration) || 30;
    var end = new Date(start.getTime() + minutes * 60000);
    try {
      var event = CalendarApp.getDefaultCalendar()
        .createEvent(title, start, end, { description: description });
      out = { success: true, eventId: event.getId() };
    } catch (err) {
      out = { error: err.message };
    }

  } else if (p.action === "ai_suggest") {
    var apiKey = p.apiKey;
    var prompt = p.prompt;
    if (!apiKey || !prompt) {
      out = { error: "Missing apiKey or prompt" };
    } else {
      // Provider is detected from the key itself: Anthropic keys start
      // with "sk-ant", anything else is treated as an OpenAI key. No
      // separate provider field is sent by the client.
      var isAnthropic = apiKey.indexOf("sk-ant") === 0;
      var url = isAnthropic
        ? "https://api.anthropic.com/v1/messages"
        : "https://api.openai.com/v1/chat/completions";
      var payload = isAnthropic
        ? JSON.stringify({
            model: "claude-sonnet-4-6",
            max_tokens: 1000,
            messages: [{ role: "user", content: prompt }]
          })
        : JSON.stringify({
            model: "gpt-4o-mini",
            messages: [{ role: "user", content: prompt }]
          });
      var options = {
        method: "post",
        contentType: "application/json",
        headers: isAnthropic
          ? { "x-api-key": apiKey, "anthropic-version": "2023-06-01" }
          : { "Authorization": "Bearer " + apiKey },
        payload: payload,
        muteHttpExceptions: true
      };
      try {
        var resp = UrlFetchApp.fetch(url, options);
        var result = JSON.parse(resp.getContentText());
        if (isAnthropic && result.content && result.content[0]) {
          out = { suggestion: result.content[0].text };
        } else if (!isAnthropic && result.choices && result.choices[0]) {
          out = { suggestion: result.choices[0].message.content };
        } else if (result.error) {
          out = { error: result.error.message || "AI provider returned an error" };
        } else {
          out = { error: "No response from AI" };
        }
      } catch(err) {
        out = { error: err.message };
      }
    }
  } else {
    out = { error: "unknown action" };
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
