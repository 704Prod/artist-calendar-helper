// taskpane.js

// Category definitions and colors (Office.MailboxEnums.CategoryColor)
// Mapping: based on your PDF + closest preset colors.
const CATEGORY_DEFINITIONS = [
  { name: "PERFORMANCE", color: "Preset0" },   // Red
  { name: "STUDIO",       color: "Preset7" },  // Blue
  { name: "COLLAB",       color: "Preset8" },  // Purple
  { name: "VIDEO SHOOT",  color: "Preset5" },  // Teal
  { name: "PHOTOSHOOT",   color: "Preset9" },  // Cranberry (closest to pink)
  { name: "PRESS",        color: "Preset3" },  // Yellow
  { name: "MARKETING",    color: "Preset4" },  // Green
  { name: "ADMIN",        color: "Preset12" }, // Gray
  { name: "REHEARSAL",    color: "Preset1" },  // Orange
  { name: "TRAVEL",       color: "Preset2" },  // Brown
  { name: "RELEASE DAY",  color: "Preset10" }, // Steel / light blue-ish
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
    document.addEventListener("DOMContentLoaded", initializeUI);
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

      // Apply categories and duration
      applyCategoryAndDuration(item, normalizedCategory);
    });
  });
}

/**
 * Normalize artist input:
 * - Split by comma
 * - Trim
 * - Join with " x "
 */
function normalizeArtists(artistRaw) {
  const parts = artistRaw
    .split(",")
    .map(p => p.trim())
    .filter(p => p.length > 0);

  if (parts.length === 0) return artistRaw; // fallback

  return parts.join(" x ");
}

/**
 * Evaluate style warnings:
 * - Emojis
 * - Hashtags
 * - Location-first in description (matches exact location string at start)
 */
function evaluateWarnings(shortDescription, location) {
  const warnings = [];

  // Basic emoji detection (typical emoji ranges)
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

/**
 * Normalize category:
 * - Uppercase
 * - Auto-correct to closest known category if needed.
 * Currently: exact match only; can be expanded to fuzzy match later.
 */
function normalizeCategory(categoryInput) {
  const upper = (categoryInput || "").toUpperCase().trim();
  const knownNames = CATEGORY_DEFINITIONS.map(c => c.name);

  if (knownNames.includes(upper)) {
    return upper;
  }

  // Simple auto-correct: case-insensitive, trim, then fallback to PERFORMANCE if no match.
  const lower = upper.toLowerCase();
  const found = knownNames.find(name => name.toLowerCase() === lower);
  if (found) return found;

  // Fallback: PERFORMANCE as default if truly unknown.
  showWarnings(["Unrecognized category '" + categoryInput + "'. Auto-corrected to PERFORMANCE."]);
  return "PERFORMANCE";
}

/**
 * Ensure master categories exist with correct colors.
 * Uses mailbox.masterCategories (Mailbox 1.8+).
 */
function ensureMasterCategories() {
  const mailbox = Office.context.mailbox;
  if (!mailbox || !mailbox.masterCategories) return;

  mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      // Not fatal for naming logic; skip.
      return;
    }

    const existing = asyncResult.value || [];
    const existingNames = existing.map(c => c.displayName.toUpperCase());

    const toAdd = CATEGORY_DEFINITIONS.filter(c => !existingNames.includes(c.name));

    if (toAdd.length === 0) return;

    const masterCategoriesToAdd = toAdd.map(c => ({
      displayName: c.name,
      color: "Office.MailboxEnums.CategoryColor." + c.color
    }));

    // Some clients accept string preset names directly; to be safe, simplify:
    const masterCategoriesToAddFixed = toAdd.map(c => ({
      displayName: c.name,
      color: c.color // e.g., "Preset0"
    }));

    mailbox.masterCategories.addAsync(masterCategoriesToAddFixed, function () {
      // Even if this fails, subject/loc/duration still work.
    });
  });
}

/**
 * Apply category and duration rules to the current appointment.
 */
function applyCategoryAndDuration(item, categoryName) {
  // First: apply category label to item
  if (item.categories && item.categories.addAsync) {
    item.categories.addAsync([categoryName], function () {
      // Ignore errors here; naming still applied.
    });
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
        // First time using the helper on this item: apply default directly
        shouldApplyDefault = true;
      } else if (lastCategoryApplied === categoryName &&
        lastStartTicks === start.getTime() &&
        lastDurationMinutes === currentMinutes) {
        // Still at last default -> safe to reapply
        shouldApplyDefault = true;
      } else {
        // Category changed or user modified time -> ask
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

      // Apply default duration
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
