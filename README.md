# Sequence — PowerPoint template system

A reusable branded deck template: eleven layouts on slide masters, a single
theme driving every colour, and named placeholders so new slides inherit
position, weight and colour automatically.

## Download

| File | What it is |
|---|---|
| `Sequence-template.potx` | The template. Open it in PowerPoint to start a new deck. |
| `Sequence-template.pptx` | Ten example slides showing every layout in use. |
| `Sequence-template.pdf` | Preview — look at this first. |

## Layouts

Cover · Section divider · Agenda · Title only · Two column · Stat ·
Chart + commentary · Workflow · Comparison · Table · Closing

## What stays editable

Text, shapes, tables and charts are native PowerPoint objects. The chart on
the chart slide has live data behind it, not a picture. The table is a real
table. Change the theme and the whole deck follows.

## Theme

Headings Cambria, body Calibri — both ship with Office, so the deck renders
the same on any machine. Ink and deep teal with an amber accent.

## Build

`build.js` generates the deck with pptxgenjs and writes the theme in
afterwards. `npm install pptxgenjs`, then `node build.js`.

Designed and built by Adarsh Subham.
