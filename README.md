# Accessibility Toolkit for Google Slides

A Google Apps Script that automates three key accessibility fixes for Google Slides presentations, powered by Claude AI.

## Features

1. **Title Generator** -- Fills in missing slide titles. Promotes existing title-like text boxes into title placeholders, or uses Claude to generate descriptive titles from slide content and screenshots.

2. **Reading Order Fixer** -- Fixes the order that screen readers use to read each slide. Sends a screenshot and element metadata to Claude, which determines the logical reading order, then rearranges the z-order accordingly.

3. **Alt Text Generator** -- Writes image descriptions so screen readers can describe images. Sends each image to Claude for a concise, meaningful alt text description.

## Prerequisites

- A Google Slides presentation
- An [Anthropic API key](https://console.anthropic.com/) (starts with `sk-ant-`)

## Setup (one-time)

1. Open your Google Slides presentation
2. In the top menu bar, click **Extensions > Apps Script**
3. A new tab opens with a code editor -- select all the starter text and **delete it**
4. Open `Code.gs` from this repo, **copy everything**, and paste it into the editor
5. In the left sidebar of the Apps Script editor, click the **"+" next to "Services"**
6. Scroll down the list, find **"Google Slides API"**, click it, then click **Add**
7. Click the **floppy disk icon** (or Ctrl+S / Cmd+S) to save
8. **Close the Apps Script tab** and **reload your presentation** (refresh the browser tab)

After the page reloads, you'll see a new **"Accessibility"** menu in your toolbar.

### Setting your API key (also one-time)

9. Click **Accessibility > Set API key**
10. Paste in your Anthropic API key and click OK
11. The first time you run anything, Google will ask you to grant permissions -- click through and allow it

## Usage

### Recommended: Run All Fixes

1. Before running, **select all slides** in the filmstrip on the left (Ctrl+A or Cmd+A), then **right-click > Apply layout** and choose **"Title Only"** or **"Title and Body"** (any layout with a title area)
2. Click **Accessibility > Run all fixes (recommended)**
3. Confirm the layout step, then the script works through all three tools automatically

For large decks it may take a few minutes since it processes each slide and image sequentially.

### Individual Tools

You can also run each tool separately from the menu:

| Menu Item | What it does |
|-----------|-------------|
| Generate missing titles | Fills in titles for slides that don't have them |
| Preview missing titles | Dry-run showing what would change (no edits) |
| Fix reading order (all slides) | Reorders elements on every slide |
| Fix reading order (this slide) | Reorders elements on the current slide only |
| Preview reading order (this slide) | Shows proposed order without applying |
| Generate alt text for images | Writes alt text for all images missing it |
| Audit images missing alt text | Lists which images need alt text (no edits) |

## Notes

- The script uses Claude Sonnet for AI-powered features (title generation, reading order analysis, alt text). You can change the model in the `CONFIG` object at the top of the script.
- API keys are stored securely in Google Apps Script's `PropertiesService` -- they are never embedded in the code.
- Images larger than 5 MB will be flagged during the pre-flight check and skipped during alt text generation.
- Edit your slide theme to reposition title/body placeholders to match your formatting style.

## License

MIT
