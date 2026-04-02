/**
 * ============================================================
 * ACCESSIBILITY TOOLKIT FOR GOOGLE SLIDES
 * ============================================================
 *
 * Three tools in one script:
 *   1. TITLE GENERATOR  -- fills in missing slide titles
 *   2. READING ORDER    -- fixes the order screen readers
 *                          use to read each slide
 *   3. ALT TEXT         -- writes image descriptions so
 *                          screen readers can describe images
 *
 * RECOMMENDED WORKFLOW:
 *   Run all three in order (the menu has a one-click option).
 *   Titles first -> reading order second -> alt text last.
 *
 * SETUP (one-time):
 *   1. Open your Google Slides presentation
 *   2. Go to Extensions > Apps Script
 *   3. Delete any starter text in the editor
 *   4. Paste this entire script
 *   5. Enable the Advanced Slides Service:
 *      - In the left sidebar, click "Services" (the + icon)
 *      - Scroll down and select "Google Slides API"
 *      - Click "Add"
 *   6. Click the floppy-disk icon to save
 *   7. Close the Apps Script tab and RELOAD your presentation
 *   8. You'll see a new "Accessibility" menu in the toolbar
 *   9. Click "Accessibility > Set API key" and paste your
 *      Anthropic API key (starts with "sk-ant-")
 *
 * After setup, just use the menu -- no need to open the
 * script editor again.
 *
 * ============================================================
 */

// ===============================================================
//  CONFIGURATION
// ===============================================================

const CONFIG = {
  CLAUDE_MODEL: "claude-sonnet-4-20250514",
  CLAUDE_API_URL: "https://api.anthropic.com/v1/messages",

  // The prompt sent to Claude when generating image alt text.
  ALT_TEXT_PROMPT:
    "Write concise, descriptive alt text for this image.\n" +
    "The alt text should:\n" +
    "- Be 1-2 sentences maximum\n" +
    "- Describe the meaningful content and function of the image\n" +
    '- Not start with "Image of" or "Picture of"\n' +
    "- Be useful for someone using a screen reader\n" +
    "- If the image contains text, include the key text content\n" +
    "Respond with ONLY the alt text, no extra commentary.",

  // When true, images that already have alt text are left alone.
  SKIP_EXISTING_ALT_TEXT: true,

  // Show detailed progress in the Apps Script log
  // (View > Executions in the Apps Script editor).
  VERBOSE_LOGGING: true,
};

// ===============================================================
//  MENU & API KEY SETUP
// ===============================================================

/**
 * Builds the custom menu in the Google Slides toolbar.
 * This runs automatically every time you open the presentation.
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu("\u267F Accessibility")
    .addItem("\u25B6 Run all fixes (recommended)", "runAllFixes")
    .addSeparator()
    .addItem("1 \u00B7 Generate missing titles", "runTitleGenerator")
    .addItem("1 \u00B7 Preview missing titles (no changes)", "previewMissingTitles")
    .addSeparator()
    .addItem("2 \u00B7 Fix reading order (all slides)", "fixReadingOrderAllSlides")
    .addItem("2 \u00B7 Fix reading order (this slide)", "fixReadingOrderCurrentSlide")
    .addItem("2 \u00B7 Preview reading order (this slide)", "previewReadingOrderCurrentSlide")
    .addSeparator()
    .addItem("3 \u00B7 Generate alt text for images", "runAltTextGenerator")
    .addItem("3 \u00B7 Audit images missing alt text", "auditMissingAltText")
    .addSeparator()
    .addItem("\u2699 Set API key", "setApiKey")
    .addToUi();
}

/**
 * Asks the user for their Anthropic API key and stores it
 * securely in the script's properties (not visible in the code).
 * Only needs to be done once.
 */
function setApiKey() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt(
    "Set Anthropic API Key",
    'Enter your Anthropic API key (starts with "sk-ant-").\n\n' +
      "This is stored securely in your script properties and is never shared.",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var key = response.getResponseText().trim();
    if (key.length > 0) {
      PropertiesService.getScriptProperties().setProperty("ANTHROPIC_API_KEY", key);
      ui.alert("\u2705 API key saved successfully!");
    }
  }
}

/** Retrieves the stored API key. */
function getApiKey_() {
  var key = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!key) {
    throw new Error(
      'No API key found. Please click "\u267F Accessibility > Set API key" first.'
    );
  }
  return key;
}

// ===============================================================
//  "RUN ALL" -- ONE-CLICK WORKFLOW
// ===============================================================

/**
 * Runs all three tools in the recommended order:
 *   1. Generate missing titles
 *   2. Fix reading order on every slide
 *   3. Generate alt text for all images
 *
 * Shows a combined summary at the end.
 */
function runAllFixes() {
  var ui = SlidesApp.getUi();

  // -- Pre-flight: make sure we have an API key --
  try {
    getApiKey_();
  } catch (e) {
    ui.alert(e.message);
    return;
  }

  // -- Confirmation --
  var confirm = ui.alert(
    "\u267F Run All Accessibility Fixes",
    "This will:\n\n" +
      "  1. Generate missing slide titles\n" +
      "  2. Remove empty body placeholders\n" +
      "  3. Fix the reading order on every slide\n" +
      "  4. Generate alt text for all images\n\n" +
      "This may take a few minutes for large decks.\n\n" +
      "Continue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // -- Step 1: Titles --
  // Ask the user to confirm they've set layouts
  if (!confirmLayoutsReady_()) return;

  log("\u2550\u2550\u2550 STEP 1 OF 4: GENERATING MISSING TITLES \u2550\u2550\u2550");
  var titleStats = runTitleGeneratorInternal_();

  // -- Step 1.5: Remove empty body placeholders --
  log("\n\u2550\u2550\u2550 STEP 2 OF 4: REMOVING EMPTY BODY PLACEHOLDERS \u2550\u2550\u2550");
  var removedBodies = removeEmptyBodyPlaceholders_();

  // -- Step 2: Reading order --
  log("\n\u2550\u2550\u2550 STEP 3 OF 4: FIXING READING ORDER \u2550\u2550\u2550");
  var orderStats = fixReadingOrderAllSlidesInternal_();

  // -- Step 3: Alt text --
  log("\n\u2550\u2550\u2550 STEP 4 OF 4: GENERATING ALT TEXT \u2550\u2550\u2550");
  var altStats = runAltTextGeneratorInternal_();

  // -- Combined summary --
  var summary =
    "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\n" +
    "TITLES\n" +
    "  Generated: " + titleStats.generated + "\n" +
    "  Promoted:  " + titleStats.promoted + "\n" +
    "  Skipped:   " + titleStats.skipped + "\n" +
    "  Errors:    " + titleStats.errors + "\n\n" +
    "CLEANUP\n" +
    "  Empty body placeholders removed: " + removedBodies + "\n\n" +
    "READING ORDER\n" +
    "  Fixed:   " + orderStats.fixed + "\n" +
    "  Skipped: " + orderStats.skipped + "\n" +
    "  Errors:  " + orderStats.errors + "\n\n" +
    "ALT TEXT\n" +
    "  Written: " + altStats.processed + "\n" +
    "  Skipped: " + altStats.skipped + "\n" +
    "  Errors:  " + altStats.errors +
    "\n\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501";

  ui.alert("\u267F All Done!", summary, ui.ButtonSet.OK);
}

// ===============================================================
//  TOOL 1: SLIDE TITLE GENERATOR
// ===============================================================
//
//  Scans the deck for slides with empty title placeholders.
//  If a text box looks like a title (bold, large, near the top),
//  it gets moved into the placeholder. Otherwise Claude
//  generates a title from the slide's content and screenshot.
//
//  PREREQUISITE: Every slide must use a layout that has a title
//  placeholder (e.g. "Title and Body").
// ---------------------------------------------------------------

/** Menu entry point -- runs the title generator with confirmation. */
function runTitleGenerator() {
  var ui = SlidesApp.getUi();
  try { getApiKey_(); } catch (e) { ui.alert(e.message); return; }
  if (!confirmLayoutsReady_()) return;

  var stats = runTitleGeneratorInternal_();

  ui.alert(
    "\u2705 Title Generator \u2014 Done!",
    "Generated:  " + stats.generated + "\n" +
    "Promoted:   " + stats.promoted + "\n" +
    "Already had: " + stats.skipped + "\n" +
    "Errors:     " + stats.errors + "\n" +
    "Total slides: " + stats.total,
    ui.ButtonSet.OK
  );
}

/**
 * Pre-flight check before running fixes.
 * 1. Scans for oversized images (>5 MB) that will fail the
 *    Claude API, and reports them with resize guidance.
 * 2. Asks the user to confirm slide layouts are set.
 * Returns true if the user clicks YES, false otherwise.
 */
function confirmLayoutsReady_() {
  var ui = SlidesApp.getUi();
  var MAX_BYTES = 5 * 1024 * 1024; // 5 MB

  // -- Scan for oversized images --
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var oversized = [];

  for (var i = 0; i < slides.length; i++) {
    var images = slides[i].getImages();
    for (var j = 0; j < images.length; j++) {
      var blob = images[j].getBlob();
      var sizeBytes = blob.getBytes().length;
      if (sizeBytes > MAX_BYTES) {
        var sizeMB = (sizeBytes / (1024 * 1024)).toFixed(1);
        var reductionPct = Math.ceil((1 - (MAX_BYTES / sizeBytes)) * 100);
        oversized.push(
          "Slide " + (i + 1) + ", Image " + (j + 1) +
          " \u2014 " + sizeMB + " MB (reduce by at least " + reductionPct + "%)"
        );
      }
    }
  }

  // -- Build the prompt message --
  var message = "";

  if (oversized.length > 0) {
    message +=
      "\u26A0\uFE0F OVERSIZED IMAGES\n" +
      "The following images exceed the 5 MB limit for alt text\n" +
      "generation. Please resize them before running, or they\n" +
      "will be skipped:\n\n" +
      oversized.join("\n") + "\n\n" +
      "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\n\n";
  }

  message +=
    "SLIDE LAYOUTS\n" +
    "The title generator needs every slide to use a layout\n" +
    "with a title placeholder (e.g. \"Title and Body\").\n\n" +
    "To set this:\n" +
    "  1. Select all slides in the filmstrip (Ctrl+A / Cmd+A)\n" +
    "  2. Right-click \u2192 Apply layout \u2192 choose one with a title\n\n" +
    "Have you completed both of the above?";

  var answer = ui.alert(
    "Before we begin\u2026",
    message,
    ui.ButtonSet.YES_NO
  );
  return answer === ui.Button.YES;
}

/**
 * Core title generation logic. Returns stats object.
 * Separated so it can be called from "Run All" without
 * showing its own popup.
 */
function runTitleGeneratorInternal_() {
  var apiKey          = getApiKey_();
  var presentation    = SlidesApp.getActivePresentation();
  var presentationId  = presentation.getId();
  var slides          = presentation.getSlides();

  var stats = { generated: 0, promoted: 0, skipped: 0, errors: 0, total: slides.length };

  for (var i = 0; i < slides.length; i++) {
    var slide       = slides[i];
    var slideNumber = i + 1;

    // Does this slide have a title placeholder?
    var titlePh = findTitlePlaceholder_(slide);
    if (!titlePh) {
      log("Slide " + slideNumber + ": no title placeholder \u2014 set layout first.");
      stats.errors++;
      continue;
    }

    // Already has a title?
    var existingTitle = titlePh.getText().asString().trim();
    if (existingTitle.length > 0) {
      log("Slide " + slideNumber + ': already has title "' + existingTitle + '" \u2014 skipping.');
      stats.skipped++;
      continue;
    }

    // Try to promote a title-like text box into the placeholder
    var titleLikeShape = findTitleLikeShape_(slide, titlePh);
    if (titleLikeShape) {
      copyTextWithFormatting_(titleLikeShape, titlePh);
      titleLikeShape.remove();
      var promoted = titlePh.getText().asString().trim();
      log("Slide " + slideNumber + ': promoted "' + promoted + '" into title placeholder.');
      stats.promoted++;
      continue;
    }

    // No title text found -- ask Claude to generate one
    log("Slide " + slideNumber + ": generating title with Claude\u2026");

    var textContent     = getSlideTextContent_(slide, titlePh);
    var speakerNotes    = getSpeakerNotes_(slide);
    var pageObjectId    = slide.getObjectId();
    var thumbnailBase64 = getSlideThumbnailBase64_(presentationId, pageObjectId);

    var suggestedTitle = callClaudeForTitle_(
      apiKey, textContent, speakerNotes, thumbnailBase64, slideNumber
    );

    if (suggestedTitle) {
      titlePh.getText().setText(suggestedTitle);
      log("Slide " + slideNumber + ': title set to "' + suggestedTitle + '"');
      stats.generated++;
    } else {
      log("Slide " + slideNumber + ": could not generate a title.");
      stats.errors++;
    }

    Utilities.sleep(1000);
  }

  return stats;
}

/** Dry-run preview -- shows what the title generator would do. */
function previewMissingTitles() {
  var slides          = SlidesApp.getActivePresentation().getSlides();
  var noPlaceholder   = [];
  var promotable      = [];
  var needsGeneration = [];

  for (var i = 0; i < slides.length; i++) {
    var slide   = slides[i];
    var titlePh = findTitlePlaceholder_(slide);

    if (!titlePh) {
      noPlaceholder.push(i + 1);
      continue;
    }

    var existing = titlePh.getText().asString().trim();
    if (existing.length > 0) continue;

    var titleLike = findTitleLikeShape_(slide, titlePh);
    if (titleLike) {
      promotable.push(i + 1);
    } else {
      needsGeneration.push(i + 1);
    }
  }

  if (noPlaceholder.length === 0 && promotable.length === 0 &&
      needsGeneration.length === 0) {
    SlidesApp.getUi().alert("\uD83C\uDF89 All slides already have titles!");
    return;
  }

  var msg = "";
  if (noPlaceholder.length > 0) {
    msg += "\u26A0\uFE0F " + noPlaceholder.length + " slide(s) have NO title placeholder.\n" +
           "Set their layout to \"Title and Body\" first.\n" +
           "Slides: " + noPlaceholder.join(", ") + "\n\n";
  }
  if (promotable.length > 0) {
    msg += "\uD83D\uDD04 " + promotable.length + " slide(s) have title-like text that " +
           "will be moved into the title placeholder:\n" +
           "Slides: " + promotable.join(", ") + "\n\n";
  }
  if (needsGeneration.length > 0) {
    msg += "\uD83E\uDD16 " + needsGeneration.length + " slide(s) need Claude to generate a title:\n" +
           "Slides: " + needsGeneration.join(", ") + "\n\n";
  }
  msg += 'Run "Generate missing titles" to fix them.';

  SlidesApp.getUi().alert(msg);
}

// -- Title Generator: helpers -----------------------------------

function findTitlePlaceholder_(slide) {
  var shapes = slide.getShapes();
  for (var j = 0; j < shapes.length; j++) {
    var type = shapes[j].getPlaceholderType();
    if (type === SlidesApp.PlaceholderType.TITLE ||
        type === SlidesApp.PlaceholderType.CENTERED_TITLE) {
      return shapes[j];
    }
  }
  return null;
}

function findTitleLikeShape_(slide, titlePlaceholder) {
  var shapes     = slide.getShapes();
  var pageHeight = SlidesApp.getActivePresentation().getPageHeight();
  var topThird   = pageHeight / 3;

  // First pass: find the typical body font size
  var bodySizes = [];
  for (var j = 0; j < shapes.length; j++) {
    var shape = shapes[j];
    if (shape.getObjectId() === titlePlaceholder.getObjectId()) continue;
    var type = shape.getPlaceholderType();
    if (type === SlidesApp.PlaceholderType.TITLE ||
        type === SlidesApp.PlaceholderType.CENTERED_TITLE ||
        type === SlidesApp.PlaceholderType.SUBTITLE) continue;
    var text = shape.getText().asString().trim();
    if (!text || text.length < 10) continue;
    var fontSize = getFirstRunFontSize_(shape);
    if (fontSize) bodySizes.push(fontSize);
  }
  var typicalBodySize = bodySizes.length > 0 ? mode_(bodySizes) : 0;

  // Second pass: find title-like text box
  var bestCandidate = null;
  var bestTop       = Infinity;

  for (var j = 0; j < shapes.length; j++) {
    var shape = shapes[j];
    if (shape.getObjectId() === titlePlaceholder.getObjectId()) continue;
    var type = shape.getPlaceholderType();
    if (type === SlidesApp.PlaceholderType.TITLE ||
        type === SlidesApp.PlaceholderType.CENTERED_TITLE ||
        type === SlidesApp.PlaceholderType.SUBTITLE) continue;
    var text = shape.getText().asString().trim();
    if (!text || text.length > 120) continue;
    var topPos   = shape.getTop();
    var isBold   = getFirstRunBold_(shape);
    var fontSize = getFirstRunFontSize_(shape);
    if (topPos > topThird) continue;
    var isBigger = (typicalBodySize > 0 && fontSize && fontSize > typicalBodySize);
    if (!isBold && !isBigger) continue;
    if (topPos < bestTop) {
      bestCandidate = shape;
      bestTop       = topPos;
    }
  }
  return bestCandidate;
}

function copyTextWithFormatting_(sourceShape, destShape) {
  var sourceText = sourceShape.getText();
  var destText   = destShape.getText();

  var paragraphs = sourceText.getParagraphs();
  var runData    = [];
  var fullText   = "";

  for (var p = 0; p < paragraphs.length; p++) {
    if (p > 0) fullText += "\n";
    var runs = paragraphs[p].getRange().getRuns();
    for (var r = 0; r < runs.length; r++) {
      var run     = runs[r];
      var runText = run.asString();
      if (r === runs.length - 1) runText = runText.replace(/\n$/, "");
      if (runText.length === 0) continue;
      var srcStyle = run.getTextStyle();
      runData.push({
        start:      fullText.length,
        end:        fullText.length + runText.length,
        bold:       srcStyle.isBold(),
        italic:     srcStyle.isItalic(),
        underline:  srcStyle.isUnderline(),
        fontSize:   srcStyle.getFontSize(),
        fontFamily: srcStyle.getFontFamily(),
        foreColor:  srcStyle.getForegroundColor()
      });
      fullText += runText;
    }
  }

  destText.setText(fullText);

  for (var i = 0; i < runData.length; i++) {
    var rd        = runData[i];
    var destRange = destText.getRange(rd.start, rd.end);
    var destStyle = destRange.getTextStyle();
    if (rd.bold !== null)       destStyle.setBold(rd.bold);
    if (rd.italic !== null)     destStyle.setItalic(rd.italic);
    if (rd.underline !== null)  destStyle.setUnderline(rd.underline);
    if (rd.fontSize !== null)   destStyle.setFontSize(rd.fontSize);
    if (rd.fontFamily !== null) destStyle.setFontFamily(rd.fontFamily);
    if (rd.foreColor !== null) {
      try { destStyle.setForegroundColor(rd.foreColor); } catch(e) {}
    }
  }
}

function getFirstRunFontSize_(shape) {
  try {
    var paragraphs = shape.getText().getParagraphs();
    if (paragraphs.length === 0) return null;
    var runs = paragraphs[0].getRange().getRuns();
    if (runs.length > 0) {
      var size = runs[0].getTextStyle().getFontSize();
      if (size) return size;
    }
    var fullText = shape.getText().asString();
    if (fullText.length > 0) {
      return shape.getText().getRange(0, 1).getTextStyle().getFontSize();
    }
  } catch (e) {}
  return null;
}

function getFirstRunBold_(shape) {
  try {
    var paragraphs = shape.getText().getParagraphs();
    if (paragraphs.length === 0) return false;
    var runs = paragraphs[0].getRange().getRuns();
    if (runs.length > 0) {
      var bold = runs[0].getTextStyle().isBold();
      if (bold === true) return true;
      if (bold === false) return false;
    }
    var fullText = shape.getText().asString();
    if (fullText.length > 0) {
      var bold = shape.getText().getRange(0, 1).getTextStyle().isBold();
      return bold === true;
    }
  } catch (e) {}
  return false;
}

function mode_(arr) {
  var counts   = {};
  var maxCount = 0;
  var modeVal  = arr[0];
  for (var i = 0; i < arr.length; i++) {
    var v = arr[i];
    counts[v] = (counts[v] || 0) + 1;
    if (counts[v] > maxCount) {
      maxCount = counts[v];
      modeVal  = v;
    }
  }
  return modeVal;
}

function getSlideTextContent_(slide, titlePlaceholder) {
  var parts  = [];
  var shapes = slide.getShapes();
  for (var j = 0; j < shapes.length; j++) {
    var shape = shapes[j];
    if (shape.getObjectId() === titlePlaceholder.getObjectId()) continue;
    var type = shape.getPlaceholderType();
    if (type === SlidesApp.PlaceholderType.TITLE ||
        type === SlidesApp.PlaceholderType.CENTERED_TITLE) continue;
    var text = shape.getText().asString().trim();
    if (text) parts.push(text);
  }
  return parts.join("\n");
}

function getSpeakerNotes_(slide) {
  var notesPage = slide.getNotesPage();
  if (!notesPage) return "";
  var shapes = notesPage.getShapes();
  for (var j = 0; j < shapes.length; j++) {
    if (shapes[j].getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
      return shapes[j].getText().asString().trim();
    }
  }
  return "";
}

function getSlideThumbnailBase64_(presentationId, pageObjectId) {
  try {
    var thumbnail = Slides.Presentations.Pages.getThumbnail(
      presentationId, pageObjectId,
      { "thumbnailProperties.thumbnailSize": "MEDIUM" }
    );
    var imageBlob = UrlFetchApp.fetch(thumbnail.contentUrl).getBlob();
    return Utilities.base64Encode(imageBlob.getBytes());
  } catch (e) {
    log("\u26A0\uFE0F Could not fetch thumbnail: " + e.message);
    return null;
  }
}

function callClaudeForTitle_(apiKey, textContent, speakerNotes, thumbnailBase64, slideNumber) {
  var contentParts = [];

  if (thumbnailBase64) {
    contentParts.push({
      type: "image",
      source: { type: "base64", media_type: "image/png", data: thumbnailBase64 }
    });
  }

  var prompt =
    "You are helping generate a concise, descriptive title for " +
    "slide " + slideNumber + " of a presentation. The slide is " +
    "currently missing its title.\n\n";

  if (textContent) {
    prompt += "TEXT CONTENT ON THE SLIDE:\n" + textContent + "\n\n";
  } else {
    prompt += "(The slide has no body text.)\n\n";
  }

  if (speakerNotes) {
    prompt += "SPEAKER NOTES:\n" + speakerNotes + "\n\n";
  }

  if (thumbnailBase64) {
    prompt += "A screenshot of the slide is attached above.\n\n";
  }

  prompt +=
    "Based on ALL of the above, suggest a single concise title " +
    "for this slide. The title should:\n" +
    "- Be descriptive and meaningful (avoid generic titles)\n" +
    "- Be concise \u2014 ideally under 8 words\n" +
    "- Accurately capture the slide's main point\n" +
    "- Work well as an accessibility label for screen readers\n\n" +
    "Respond with ONLY the title text. No quotes, no explanation, " +
    "no trailing punctuation unless the title is a question.";

  contentParts.push({ type: "text", text: prompt });

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify({
      model: CONFIG.CLAUDE_MODEL,
      max_tokens: 100,
      messages: [{ role: "user", content: contentParts }]
    }),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(CONFIG.CLAUDE_API_URL, options);
    if (response.getResponseCode() !== 200) {
      log("API error: " + response.getContentText());
      return null;
    }
    var data = JSON.parse(response.getContentText());
    return data.content[0].text.trim();
  } catch (e) {
    log("Error calling Claude: " + e.message);
    return null;
  }
}

// ===============================================================
//  CLEANUP: REMOVE EMPTY BODY PLACEHOLDERS
// ===============================================================
//
//  When using "Title and Body" layouts, slides that don't have
//  body text end up with an empty body placeholder. Screen
//  readers will announce these as blank text boxes, which is
//  confusing. This step removes them.
// ---------------------------------------------------------------

/**
 * Scans every slide for BODY placeholders that are empty
 * and removes them. Returns the number removed.
 */
function removeEmptyBodyPlaceholders_() {
  var presentation = SlidesApp.getActivePresentation();
  var slides       = presentation.getSlides();
  var removed      = 0;

  for (var i = 0; i < slides.length; i++) {
    var shapes = slides[i].getShapes();

    // Walk backwards so removing a shape doesn't shift indices
    for (var j = shapes.length - 1; j >= 0; j--) {
      var shape           = shapes[j];
      var placeholderType = shape.getPlaceholderType();

      if (placeholderType === SlidesApp.PlaceholderType.BODY ||
          placeholderType === SlidesApp.PlaceholderType.SUBTITLE) {
        var text = shape.getText().asString().trim();
        if (text.length === 0) {
          log("Slide " + (i + 1) + ": removed empty body placeholder");
          shape.remove();
          removed++;
        }
      }
    }
  }

  log("Removed " + removed + " empty body placeholder(s) total.");
  return removed;
}

// ===============================================================
//  TOOL 2: READING ORDER FIXER
// ===============================================================
//
//  Takes a screenshot of each slide, sends it to Claude along
//  with element metadata, and Claude determines the logical
//  reading order. The script then rearranges the z-order so
//  screen readers read elements in the correct sequence.
// ---------------------------------------------------------------

/** Menu entry: fix reading order on the currently selected slide. */
function fixReadingOrderCurrentSlide() {
  var ui = SlidesApp.getUi();
  try { getApiKey_(); } catch (e) { ui.alert(e.message); return; }

  var presentation = SlidesApp.getActivePresentation();
  var slide        = presentation.getSelection().getCurrentPage().asSlide();
  var slideIndex   = getSlideIndex_(presentation, slide);

  var result = analyzeAndFixSlide_(presentation, slide, slideIndex);
  ui.alert("Done", "Slide " + (slideIndex + 1) + ": " + result, ui.ButtonSet.OK);
}

/** Menu entry: fix reading order on every slide. */
function fixReadingOrderAllSlides() {
  var ui = SlidesApp.getUi();
  try { getApiKey_(); } catch (e) { ui.alert(e.message); return; }

  var confirm = ui.alert(
    "Fix reading order on all slides?",
    "This will analyze and reorder elements on every slide. Continue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var stats = fixReadingOrderAllSlidesInternal_();

  ui.alert(
    "\u2705 Reading Order \u2014 Done!",
    "Fixed:   " + stats.fixed + "\n" +
    "Skipped: " + stats.skipped + "\n" +
    "Errors:  " + stats.errors,
    ui.ButtonSet.OK
  );
}

/** Internal version for "Run All". Returns stats. */
function fixReadingOrderAllSlidesInternal_() {
  var presentation = SlidesApp.getActivePresentation();
  var slides       = presentation.getSlides();

  var stats = { fixed: 0, skipped: 0, errors: 0 };

  for (var i = 0; i < slides.length; i++) {
    try {
      var result = analyzeAndFixSlide_(presentation, slides[i], i);
      if (result.indexOf("Skipped") === 0) {
        stats.skipped++;
      } else {
        stats.fixed++;
      }
      log("Slide " + (i + 1) + ": " + result);
    } catch (e) {
      log("Slide " + (i + 1) + ": ERROR \u2014 " + e.message);
      stats.errors++;
    }
  }

  return stats;
}

/** Menu entry: preview reading order for the current slide. */
function previewReadingOrderCurrentSlide() {
  var ui = SlidesApp.getUi();
  try { getApiKey_(); } catch (e) { ui.alert(e.message); return; }

  var presentation = SlidesApp.getActivePresentation();
  var slide        = presentation.getSelection().getCurrentPage().asSlide();
  var slideIndex   = getSlideIndex_(presentation, slide);

  var elements = getElementMetadata_(slide);
  if (elements.length < 2) {
    ui.alert("This slide has fewer than 2 elements \u2014 nothing to reorder.");
    return;
  }

  var thumbnailUrl = getSlideThumbUrl_(presentation, slide, slideIndex);
  var imageBase64  = fetchImageAsBase64_(thumbnailUrl);
  var orderedIds   = getReadingOrderFromClaude_(imageBase64, elements);

  var preview = "Claude recommends this reading order:\n\n";
  for (var i = 0; i < orderedIds.length; i++) {
    var el = findElementById_(elements, orderedIds[i]);
    var label = el
      ? (el.type + ': "' + truncate_(el.content, 50) + '"')
      : orderedIds[i];
    preview += (i + 1) + ". " + label + "\n";
  }
  preview += '\nUse "Fix reading order" to apply this order.';
  ui.alert("Proposed Reading Order \u2014 Slide " + (slideIndex + 1), preview, ui.ButtonSet.OK);
}

// -- Reading Order: core pipeline --------------------------------

function analyzeAndFixSlide_(presentation, slide, slideIndex) {
  var elements = getElementMetadata_(slide);
  if (elements.length < 2) {
    return "Skipped \u2014 fewer than 2 elements on this slide.";
  }

  var thumbnailUrl = getSlideThumbUrl_(presentation, slide, slideIndex);
  var imageBase64  = fetchImageAsBase64_(thumbnailUrl);
  var orderedIds   = getReadingOrderFromClaude_(imageBase64, elements);

  applyZOrder_(presentation, slide, slideIndex, orderedIds);
  return "Reordered " + orderedIds.length + " elements.";
}

function getElementMetadata_(slide) {
  var pageElements = slide.getPageElements();
  var elements     = [];

  for (var i = 0; i < pageElements.length; i++) {
    var el = pageElements[i];
    elements.push({
      id:      el.getObjectId(),
      type:    getElementType_(el),
      content: getElementContent_(el),
      left:    el.getLeft(),
      top:     el.getTop(),
      width:   el.getWidth(),
      height:  el.getHeight()
    });
  }
  return elements;
}

function getElementType_(element) {
  var type = element.getPageElementType();
  switch (type) {
    case SlidesApp.PageElementType.SHAPE:        return "Shape/TextBox";
    case SlidesApp.PageElementType.IMAGE:        return "Image";
    case SlidesApp.PageElementType.TABLE:        return "Table";
    case SlidesApp.PageElementType.VIDEO:        return "Video";
    case SlidesApp.PageElementType.GROUP:        return "Group";
    case SlidesApp.PageElementType.LINE:         return "Line";
    case SlidesApp.PageElementType.SHEETS_CHART: return "Chart";
    case SlidesApp.PageElementType.WORD_ART:     return "WordArt";
    default:                                     return "Other";
  }
}

function getElementContent_(element) {
  try {
    var type = element.getPageElementType();

    if (type === SlidesApp.PageElementType.SHAPE) {
      var text = element.asShape().getText().asString().trim();
      return text || "[empty shape]";
    }
    if (type === SlidesApp.PageElementType.IMAGE) {
      var desc  = element.asImage().getDescription();
      var title = element.asImage().getTitle();
      return "Image: " + (title || desc || "[no alt text]");
    }
    if (type === SlidesApp.PageElementType.TABLE) {
      var table = element.asTable();
      return "Table (" + table.getNumRows() + " rows x " + table.getNumColumns() + " cols)";
    }
    if (type === SlidesApp.PageElementType.GROUP) {
      return "Group (" + element.asGroup().getChildren().length + " items)";
    }
    return "[" + getElementType_(element) + "]";
  } catch (e) {
    return "[unreadable]";
  }
}

// -- Reading Order: screenshot helpers ---------------------------

function getSlideThumbUrl_(presentation, slide, slideIndex) {
  var presentationId = presentation.getId();
  var pageId         = slide.getObjectId();

  var thumbnail = Slides.Presentations.Pages.getThumbnail(
    presentationId, pageId,
    {
      "thumbnailProperties.thumbnailSize": "LARGE",
      "thumbnailProperties.mimeType": "PNG"
    }
  );
  return thumbnail.contentUrl;
}

function fetchImageAsBase64_(url) {
  var response = UrlFetchApp.fetch(url);
  var blob     = response.getBlob();
  return Utilities.base64Encode(blob.getBytes());
}

// -- Reading Order: Claude call ----------------------------------

function getReadingOrderFromClaude_(imageBase64, elements) {
  var apiKey = getApiKey_();

  var elementList = elements.map(function(el) {
    return "- ID: " + el.id +
           " | Type: " + el.type +
           ' | Content: "' + truncate_(el.content, 80) + '"' +
           " | Position: left=" + Math.round(el.left) + "pt, top=" + Math.round(el.top) + "pt" +
           " | Size: " + Math.round(el.width) + "x" + Math.round(el.height) + "pt";
  }).join("\n");

  var prompt =
    "You are an accessibility expert. I am showing you a screenshot of a Google Slides presentation slide.\n\n" +
    "Here are the elements on this slide:\n\n" +
    elementList + "\n\n" +
    "TASK: Look at the screenshot and determine the correct reading order for a screen reader. " +
    "The reading order should follow the logical visual flow that a sighted person would use \u2014 " +
    "typically: slide title first, then main content top-to-bottom and left-to-right, " +
    "with related elements grouped together (e.g., a caption should come right after its image). " +
    "Decorative elements like background shapes or divider lines should come LAST.\n\n" +
    "RESPOND WITH ONLY a JSON array of element IDs in the correct reading order, from first to last. " +
    'Example: ["id1", "id2", "id3"]\n' +
    "Do not include any other text, explanation, or markdown formatting \u2014 just the JSON array.";

  var payload = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: 1024,
    messages: [{
      role: "user",
      content: [
        {
          type: "image",
          source: { type: "base64", media_type: "image/png", data: imageBase64 }
        },
        { type: "text", text: prompt }
      ]
    }]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CONFIG.CLAUDE_API_URL, options);

  if (response.getResponseCode() !== 200) {
    throw new Error("Claude API returned status " + response.getResponseCode() +
      ": " + response.getContentText());
  }

  var data = JSON.parse(response.getContentText());
  var text = data.content[0].text.trim();

  // Clean up any accidental markdown fencing
  text = text.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();

  var orderedIds = JSON.parse(text);

  // Validate: every returned ID must exist on the slide
  var knownIds = elements.map(function(el) { return el.id; });
  var validatedIds = orderedIds.filter(function(id) {
    return knownIds.indexOf(id) !== -1;
  });

  // Safety net: add any missing IDs at the end
  knownIds.forEach(function(id) {
    if (validatedIds.indexOf(id) === -1) {
      validatedIds.push(id);
    }
  });

  return validatedIds;
}

// -- Reading Order: apply z-order --------------------------------

/**
 * Rearranges elements on the slide so the reading order
 * matches the provided array of IDs.
 *
 * Screen readers read from back (bottom of stack) to front
 * (top of stack). We bring each element to the front in
 * reading-order sequence, so the first-to-read ends up at the
 * back and the last-to-read ends up on top.
 */
function applyZOrder_(presentation, slide, slideIndex, orderedIds) {
  var presentationId = presentation.getId();
  var requests       = [];

  for (var i = 0; i < orderedIds.length; i++) {
    requests.push({
      updatePageElementsZOrder: {
        pageElementObjectIds: [orderedIds[i]],
        operation: "BRING_TO_FRONT"
      }
    });
  }

  Slides.Presentations.batchUpdate(
    { requests: requests },
    presentationId
  );
}

// ===============================================================
//  TOOL 3: IMAGE ALT TEXT GENERATOR
// ===============================================================
//
//  Finds every image in the presentation, sends each one to
//  Claude for a description, and writes the description back
//  as the image's alt text.
// ---------------------------------------------------------------

/** Menu entry point for alt text generation. */
function runAltTextGenerator() {
  var ui = SlidesApp.getUi();
  try { getApiKey_(); } catch (e) { ui.alert(e.message); return; }

  var stats = runAltTextGeneratorInternal_();

  ui.alert(
    "\u2705 Alt Text Generator \u2014 Done!",
    "Written: " + stats.processed + "\n" +
    "Skipped: " + stats.skipped + "\n" +
    "Errors:  " + stats.errors,
    ui.ButtonSet.OK
  );
}

/** Internal version for "Run All". Returns stats. */
function runAltTextGeneratorInternal_() {
  var presentation = SlidesApp.getActivePresentation();
  var slides       = presentation.getSlides();

  log('Presentation: "' + presentation.getName() + '"');
  log("Total slides: " + slides.length);
  log("Skip images with existing alt text: " + CONFIG.SKIP_EXISTING_ALT_TEXT);
  log("\u2500".repeat(50));

  var stats = { processed: 0, skipped: 0, errors: 0 };

  for (var i = 0; i < slides.length; i++) {
    var slide  = slides[i];
    var images = slide.getImages();

    if (images.length === 0) continue;

    log("\nSlide " + (i + 1) + ": Found " + images.length + " image(s)");

    for (var j = 0; j < images.length; j++) {
      var image       = images[j];
      var existingAlt = image.getDescription();

      if (CONFIG.SKIP_EXISTING_ALT_TEXT && existingAlt && existingAlt.trim() !== "") {
        log("  Image " + (j + 1) + ": Already has alt text, skipping");
        stats.skipped++;
        continue;
      }

      try {
        log("  Image " + (j + 1) + ": Sending to Claude...");
        var altText = getAltTextFromClaude_(image);

        image.setDescription(altText);

        var shortTitle = altText.length > 100
          ? altText.substring(0, 97) + "..."
          : altText;
        image.setTitle(shortTitle);

        log('  Image ' + (j + 1) + ': "' + altText + '"');
        stats.processed++;
      } catch (error) {
        log("  Image " + (j + 1) + ": Error \u2014 " + error.message);
        stats.errors++;
      }

      Utilities.sleep(1500);
    }
  }

  return stats;
}

/** Audit: lists images missing alt text without making changes. */
function auditMissingAltText() {
  var presentation = SlidesApp.getActivePresentation();
  var slides       = presentation.getSlides();
  var missing      = [];

  for (var i = 0; i < slides.length; i++) {
    var images = slides[i].getImages();
    for (var j = 0; j < images.length; j++) {
      var desc = images[j].getDescription();
      if (!desc || desc.trim() === "") {
        missing.push("Slide " + (i + 1) + ", Image " + (j + 1));
      }
    }
  }

  var message = missing.length === 0
    ? "\uD83C\uDF89 All images have alt text!"
    : "\u26A0\uFE0F " + missing.length + " image(s) missing alt text:\n\n" +
      missing.join("\n");

  SlidesApp.getUi().alert("Alt Text Audit", message, SlidesApp.getUi().ButtonSet.OK);
}

// -- Alt Text: Claude call ---------------------------------------

function getAltTextFromClaude_(image) {
  var apiKey     = getApiKey_();
  var blob       = image.getBlob();
  var base64Data = Utilities.base64Encode(blob.getBytes());
  var mimeType   = blob.getContentType();

  var supportedTypes = ["image/png", "image/jpeg", "image/gif", "image/webp"];
  if (supportedTypes.indexOf(mimeType) === -1) {
    throw new Error("Unsupported image type: " + mimeType + ". Claude accepts PNG, JPEG, GIF, and WebP.");
  }

  var requestBody = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: 300,
    messages: [{
      role: "user",
      content: [
        {
          type: "image",
          source: { type: "base64", media_type: mimeType, data: base64Data }
        },
        { type: "text", text: CONFIG.ALT_TEXT_PROMPT }
      ]
    }]
  };

  var options = {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  var response     = UrlFetchApp.fetch(CONFIG.CLAUDE_API_URL, options);
  var responseCode = response.getResponseCode();
  var responseBody = JSON.parse(response.getContentText());

  if (responseCode !== 200) {
    var errorMsg = responseBody?.error?.message || ("HTTP " + responseCode);
    throw new Error("Claude API error: " + errorMsg);
  }

  var textBlock = responseBody?.content?.find(function(block) {
    return block.type === "text";
  });
  var altText = textBlock?.text;

  if (!altText) {
    throw new Error("Claude returned an empty response");
  }

  return altText.trim();
}

// ===============================================================
//  SHARED UTILITIES
// ===============================================================

function getSlideIndex_(presentation, slide) {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; i++) {
    if (slides[i].getObjectId() === slide.getObjectId()) return i;
  }
  return 0;
}

function findElementById_(elements, id) {
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].id === id) return elements[i];
  }
  return null;
}

function truncate_(str, maxLen) {
  if (!str) return "";
  if (str.length <= maxLen) return str;
  return str.substring(0, maxLen) + "\u2026";
}

function log(message) {
  if (CONFIG.VERBOSE_LOGGING) {
    Logger.log(message);
  }
}
