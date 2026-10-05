# Sequence template — usage notes

## Starting a deck

Double-click `Sequence-template.potx`. PowerPoint opens a new untitled
presentation built on the template; the .potx itself stays unchanged. To make
the template available permanently, put it in
`Documents\Custom Office Templates` and it appears under File → New → Personal.

## Adding a slide

Home → New Slide, then pick a layout by name. Ten are available:

| Layout | Use it for |
|---|---|
| 1 Cover | The opening slide |
| 2 Section | Resetting attention between topics |
| 3 Agenda | Previewing the sections |
| 4 Content | Any slide composed of text, cards or shapes |
| 5 Content two column | Text beside a list, image or quote |
| 6 Product screenshot | A screenshot with commentary beside it |
| 7 Workflow | Sequential steps |
| 8 Comparison | Before and after, or option A against option B |
| 9 Data | A chart or table that needs interpreting |
| 10 Closing | The decision and next steps |

Click into the dotted regions and type. Position, size, weight and colour are
already set by the layout, so a new slide matches the deck without any
formatting work.

## Colour and type

Every colour in the deck comes from one theme. Design → Variants → Colours
changes all of them at once. Never pick a colour from the "Standard colours"
row at the bottom of the picker — those are fixed values that will not follow
the theme later. Use the top row, which is theme colours.

Headings are Cambria, body text is Calibri. Both ship with Office, so the deck
renders identically on any machine without font files travelling with it.

## Charts and tables

The chart on the Data slide is a native PowerPoint chart. Right-click → Edit
Data opens the figures in a spreadsheet; the chart redraws in theme colours.
Tables are real tables — add rows with Tab from the last cell.

Keep axis and data labels looking the same across every chart in a deck. If
you add a second chart, copy the first one and change its data rather than
inserting a fresh chart, which arrives with PowerPoint's default styling.

## Screenshots

On layout 6, click the picture icon in the placeholder and choose your image.
Crop the screenshot to its content before inserting; whitespace inside the
image reads as a mistake on the slide. Keep screenshots at the same corner
radius and shadow as each other across a deck.

## Editing the footer or slide numbers

The footer text and slide number live on the layouts, not on individual
slides. View → Slide Master, edit them once, close master view, and every
slide updates.

## What to avoid

- Don't drag placeholders to new positions on individual slides; that breaks
  the alignment between slides of the same type.
- Don't paste text with source formatting. Paste as plain text (Ctrl+Shift+V)
  so it inherits the layout's styling.
- Don't shrink body text below 14pt. If text doesn't fit, cut words or split
  the slide.

## Rebuilding the template

`build.js` generates the deck with pptxgenjs and writes the theme afterwards.
`npm install pptxgenjs`, then `node build.js`. Changing the `THEME` object at
the top re-colours the whole system.
