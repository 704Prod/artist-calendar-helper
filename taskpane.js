// taskpane.js
// Artist Calendar Helper - v1.0.3

// Artist roster with email addresses
const ARTIST_ROSTER = [
  { name: "704 Prod", email: "prod@thelegacycrew.com" },
  { name: "704 Lenz", email: "lenz@thelegacycrew.com" },
  { name: "Kyla Harmony", email: "kyla@thelegacycrew.com" },
  { name: "Medeuca", email: "medeuca@thelegacycrew.com" },
  { name: "Crosshairs", email: "crosshairs@thelegacycrew.com" },
  { name: "K4", email: "k4@thelegacycrew.com" }
];

// Category definitions and colors (Office.MailboxEnums.CategoryColor)
const CATEGORY_DEFINITIONS = [
  { name: "PERFORMANCE", color: "Preset0" },   // Red
  { name: "STUDIO",       color: "Preset7" },  // Blue
  { name: "COLLAB",       color: "Preset8" },  // Purple
  { name: "VIDEO SHOOT",  color: "Preset5" },  // Teal
  { name: "PHOTOSHOOT",   color: "Preset9" },  // Cranberry
  { name: "PRESS",        color: "Preset3" },  // Yellow
  { name: "MARKETING",    color: "Preset4" },  // Green
  { name: "ADMIN",        color: "Preset12" }, // Gray
  { name: "REHEARSAL",    color: "Preset1" },  // Orange
  { name: "TRAVEL",       color: "Preset2" },  // Brown
  { name: "RELEASE DAY",  color: "Preset10" }, // Steel
  { name: "REST",         color: "Preset14" }  // Black
];

// Default durations in minutes
const CATEGORY_DURATIONS = {
  "PERFORMANCE": 3 * 60,
  "STUDIO": 4 * 60,
  "COLLAB": 3 * 60,
  "VIDEO SHOOT": 6 * 60,
  "PHOTOSHOOT": 2 * 60,
  "PRESS": 2 * 60,
  "MARKETING": 2 * 60,
  "ADMIN": 6 * 60,
  "REHEARSAL": 2 * 60,
  "TRAVEL": 8 * 60,
  "RELEASE DAY": null, // All-day override
  "REST": 8 * 60
};

// State tracking for duration override logic
let lastCategoryApplied = null;
let lastStartTicks = null;
let lastDurationMinutes = null;

// Initialize Office JS
Office.onReady(function () {
  if (Office.context.host === Office.HostType.Outlook) {
    initializeUI();
    ensureMasterCategories();
  }
});

function initializeUI() {
  const step1Next = document.getElementById("step1Next");
  const step2Back = document.getElementById("step2Back");
  const applyButton = document.getElementById("applyButton");

  step1Next.addEventListener("click", onStep1Next);
  step2Back.addEventListener("click", onStep2Back);
  applyButton.addEventListener("click", onApplyClicked);
}

function showStep(stepNumber) {
  const step1 = document.getElementById("step1");
  const step2 = document.getElementById("step2");
  const step1Label = document.getElementById("step1-label");
  const step2Label = document.getElementById("step2-label");

  if (stepNumber === 1) {
    step1.classList.remove("hidden");
    step2.classList.add("hidden");
    step1Label.classList.add("active");
    step2Label.classList.remove("active");
  } else {
    step1.classList.add("hidden");
    step2.classList.remove("hidden");
    step1Label.classList.remove("active");
    step2Label.classList.add("active");
  }

  clearMessages();
}

function onStep1Next() {
  const category = getSelectedCategory();
  if (!category) {
    showError("Category is required.");
    return;
  }
  showStep(2);
}

function onStep2Back() {
  showStep(1);
}

function getSelectedCategory() {
  const select = document.getElementById("categorySelect");
  const value = (select.value || "").trim();
  if (!value) return "";
  return value.toUpperCase();
}

function onApplyClicked() {
  clearMessages();

  const category = getSelectedCategory();
  const artistRaw = document.getElementById("artistInput").value.trim();
  const shortDescription = document.getElementById("shortDescription").value.trim();
  const location = document.getElementById("locationInput").value.trim();
  const manualAttendeesRaw = document.getElementById("manualAttendees").value.trim();

  // Required fields enforcement
  const missing = [];
  if (!category) missing.push("Category");
  if (!artistRaw) missing.push("Artist / Participants");
  if (!shortDescription) missing.push("Short Description");
  if (!location) missing.push("Location");

  if (missing.length > 0) {
    showError("Missing required fields: " + missing.join(", "));
    return;
  }

  const normalizedCategory = normalizeCategory(category);
  const normalizedArtists = normalizeArtists(artistRaw);
  const warnings = evaluateWarnings(shortDescription, location);

  if (warnings.length > 0) {
    showWarnings(warnings);
  }

  // Parse attendees - check if artists in artistRaw match roster for auto-invite
  const attendeeEmails = parseAttendees(artistRaw, manualAttendeesRaw);

  const finalSubject =
    normalizedCategory +
    " — " +
    normalizedArtists +
    " — " +
    shortDescription +
    " — " +
    location;

  const item = Office.context.mailbox.item;

  // Set subject
  item.subject.setAsync(finalSubject, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      showError("Failed to set subject: " + result.error.message);
      return;
    }

    // Set location (Outlook field)
    item.location.setAsync(location, function (locResult) {
      if (locResult.status === Office.AsyncResultStatus.Failed) {
        showError("Failed to set location: " + locResult.error.message);
        return;
      }

      // Add attendees if any
      if (attendeeEmails.length > 0) {
        addAttendees(item, attendeeEmails, function() {
          // Apply categories and duration after attendees
          applyCategoryAndDuration(item, normalizedCategory);
        });
      } else {
        // No attendees, proceed to categories
        applyCategoryAndDuration(item, normalizedCategory);
      }
    });
  });
}

function normalizeArtists(artistRaw) {
  const parts = artistRaw
    .split(",")
    .map(p => p.trim())
    .filter(p => p.length > 0);

  if (parts.length === 0) return artistRaw;
  return parts.join(" x ");
}

/**
 * Parse attendees from artist names and manual input
 * Artists in the roster automatically get invited
 * Returns array of email addresses
 */
function parseAttendees(artistRaw, manualAttendeesRaw) {
  const emails = [];

  // Parse artist names and check if they're in roster
  if (artistRaw) {
    const names = artistRaw
      .split(",")
      .map(n => n.trim())
      .filter(n => n.length > 0);

    names.forEach(name => {
      const artist = ARTIST_ROSTER.find(a => a.name.toLowerCase() === name.toLowerCase());
      if (artist && artist.email) {
        emails.push(artist.email);
        console.log("Auto-inviting roster artist:", name, "->", artist.email);
      }
    });
  }

  // Parse manual attendees (comma or semicolon separated emails)
  if (manualAttendeesRaw) {
    const manualEmails = manualAttendeesRaw
      .split(/[,;]/)
      .map(e => e.trim())
      .filter(e => e.length > 0 && isValidEmail(e));

    emails.push(...manualEmails);
    console.log("Added manual attendees:", manualEmails);
  }

  // Remove duplicates
  return [...new Set(emails)];
}

/**
 * Basic email validation
 */
function isValidEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

/**
 * Add attendees to the appointment as required attendees
 */
function addAttendees(item, emailAddresses, callback) {
  if (!item.requiredAttendees || !item.requiredAttendees.addAsync) {
    console.log("Required attendees API not available");
    if (callback) callback();
    return;
  }

  const attendees = emailAddresses.map(email => ({
    emailAddress: email,
    displayName: email.split("@")[0] // Use email prefix as display name
  }));

  item.requiredAttendees.addAsync(attendees, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Added required attendees:", emailAddresses);
    } else {
      console.log("Failed to add attendees:", result.error);
    }
    if (callback) callback();
  });
}

function evaluateWarnings(shortDescription, location) {
  const warnings = [];

  const emojiRegex = /[\u{1F300}-\u{1FAFF}]/u;
  if (emojiRegex.test(shortDescription)) {
    warnings.push("Short Description contains emojis. Standard advises against emojis.");
  }

  if (shortDescription.indexOf("#") !== -1) {
    warnings.push("Short Description contains hashtags. Standard advises against hashtags.");
  }

  if (
    location &&
    shortDescription.toLowerCase().startsWith(location.toLowerCase())
  ) {
    warnings.push("Short Description appears to start with the location. Standard advises not to put location first.");
  }

  return warnings;
}

function normalizeCategory(categoryInput) {
  const upper = (categoryInput || "").toUpperCase().trim();
  const knownNames = CATEGORY_DEFINITIONS.map(c => c.name);

  if (knownNames.includes(upper)) {
    return upper;
  }

  const lower = upper.toLowerCase();
  const found = knownNames.find(name => name.toLowerCase() === lower);
  if (found) return found;

  showWarnings(["Unrecognized category '" + categoryInput + "'. Auto-corrected to PERFORMANCE."]);
  return "PERFORMANCE";
}

function ensureMasterCategories() {
  const mailbox = Office.context.mailbox;
  if (!mailbox || !mailbox.masterCategories) {
    console.log("Master categories API not available on this client");
    return;
  }

  mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log("Failed to get master categories:", asyncResult.error);
      return;
    }

    const existing = asyncResult.value || [];
    const existingNames = existing.map(c => c.displayName.toUpperCase());
    const toAdd = CATEGORY_DEFINITIONS.filter(c => !existingNames.includes(c.name));

    if (toAdd.length === 0) {
      console.log("All categories already exist");
      return;
    }

    console.log("Adding categories:", toAdd.map(c => c.name));

    // Try multiple color format approaches for compatibility
    const masterCategoriesToAdd = toAdd.map(c => ({
      displayName: c.name,
      color: Office.MailboxEnums.CategoryColor[c.color]
    }));

    mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (addResult) {
      if (addResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
      } else {
        console.log("Failed to add categories:", addResult.error);
        // Fallback: try with string color values
        const fallbackCategories = toAdd.map(c => ({
          displayName: c.name,
          color: c.color
        }));
        mailbox.masterCategories.addAsync(fallbackCategories, function (fallbackResult) {
          if (fallbackResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Successfully added categories with fallback method");
          } else {
            console.log("Fallback also failed:", fallbackResult.error);
          }
        });
      }
    });
  });
}

function applyCategoryAndDuration(item, categoryName) {
  // Apply category label to item
  if (item.categories && item.categories.addAsync) {
    item.categories.addAsync([categoryName], function (catResult) {
      if (catResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Category applied:", categoryName);
      } else {
        console.log("Failed to apply category:", catResult.error);
      }
    });
  } else {
    console.log("Categories API not available on this client");
  }

  // RELEASE DAY: force all-day, no prompt
  if (categoryName === "RELEASE DAY") {
    if (item.isAllDayEvent && item.isAllDayEvent.setAsync) {
      item.isAllDayEvent.setAsync(true, function () {
        showStatus("Subject, location, category, and all-day RELEASE DAY applied.");
      });
    } else {
      showStatus("Subject, location, and category applied. (All-day flag not supported on this client.)");
    }
    lastCategoryApplied = categoryName;
    lastStartTicks = null;
    lastDurationMinutes = null;
    return;
  }

  // All other categories: duration logic
  const durationMinutes = CATEGORY_DURATIONS[categoryName];
  if (!durationMinutes) {
    showStatus("Subject, location, and category applied. (No default duration configured for " + categoryName + ".)");
    return;
  }

  // Get current start/end, decide whether to override
  item.start.getAsync(function (startResult) {
    if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
      showStatus("Subject, location, and category applied. (Failed to get start time.)");
      return;
    }

    const start = startResult.value;
    item.end.getAsync(function (endResult) {
      if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
        showStatus("Subject, location, and category applied. (Failed to get end time.)");
        return;
      }

      const end = endResult.value;
      const currentMinutes = Math.round((end.getTime() - start.getTime()) / 60000);

      let shouldApplyDefault = false;

      if (!lastCategoryApplied) {
        shouldApplyDefault = true;
      } else if (lastCategoryApplied === categoryName &&
        lastStartTicks === start.getTime() &&
        lastDurationMinutes === currentMinutes) {
        shouldApplyDefault = true;
      } else {
        const msg =
          "Detected custom or previous duration.\n" +
          "Apply default " + categoryName + " duration of " + durationMinutes + " minutes?";
        shouldApplyDefault = window.confirm(msg);
      }

      if (!shouldApplyDefault) {
        showStatus("Subject, location, and category applied. Existing time preserved.");
        lastCategoryApplied = categoryName;
        lastStartTicks = start.getTime();
        lastDurationMinutes = currentMinutes;
        return;
      }

      const newEnd = new Date(start.getTime() + durationMinutes * 60000);
      item.end.setAsync(newEnd, function (setResult) {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          showStatus("Subject, location, and category applied. Failed to set default duration.");
          return;
        }

        showStatus(
          "Subject, location, category, and default " +
          categoryName +
          " duration applied (" + durationMinutes + " minutes)."
        );

        lastCategoryApplied = categoryName;
        lastStartTicks = start.getTime();
        lastDurationMinutes = durationMinutes;
      });
    });
  });
}

// Messaging helpers
function clearMessages() {
  const warnings = document.getElementById("warnings");
  const errors = document.getElementById("errors");
  const status = document.getElementById("status");

  [warnings, errors, status].forEach(el => {
    el.classList.add("hidden");
    el.textContent = "";
  });
}

function showWarnings(list) {
  if (!list || !list.length) return;
  const el = document.getElementById("warnings");
  el.innerHTML = list.map(w => "- " + w).join("<br/>");
  el.classList.remove("hidden");
}

function showError(msg) {
  const el = document.getElementById("errors");
  el.textContent = msg;
  el.classList.remove("hidden");
}

function showStatus(msg) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.classList.remove("hidden");
}

