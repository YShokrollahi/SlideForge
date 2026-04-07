# Slide HTML Set

This repository contains a set of clean, structured HTML slides designed to be converted into editable PowerPoint presentations using the SlideForge workflow.

---

## 🎯 Purpose

The goal of this repo is to provide a simple way to:

- Design slides using HTML (fast, flexible, version-controlled)
- Convert them into fully editable PowerPoint slides
- Avoid manual slide design in PowerPoint
- Keep presentations reproducible and easy to update

This is especially useful for:
- research presentations
- lab meetings
- technical demos
- iterative storytelling (edit → regenerate PPT)

---

## ⚙️ How it works

Each slide is written as a standalone HTML file.

SlideForge parses these files and converts:
- text → editable text boxes
- shapes → PowerPoint shapes
- layout → positioned elements

---

## 📁 Structure

```
shared.css        # common styling across slides
slides/
  01_*.html       # one file = one slide
  02_*.html
  ...
```

---

## 🚀 How to use

### 1. Create / edit slides

Edit any file inside:

```
slides/
```

Each slide:
- fixed size: `1280 x 720`
- uses absolute positioning
- uses `data-object="true"` for editable elements

---

### 2. Build PowerPoint

Run:

```bash
node build.js
```

This will:
- read all HTML slides
- convert them into PowerPoint slides
- export a `.pptx` file in the `output/` folder

---

### 3. Open in PowerPoint

The generated slides are:
- fully editable
- text can be changed
- shapes can be resized
- layout remains intact

---

## 🧠 Design Guidelines

To keep slides editable and clean:

Use:
- `data-object-type="textbox"` for text
- `data-object-type="shape"` for boxes/backgrounds

Avoid:
- complex CSS layouts (flex/grid)
- nested structures
- unsupported styling

Keep layout:
- flat
- absolute-positioned

---

## 🖼 Suggested images to add later

- System architecture diagram
- Main UI screenshot
- Example output / export visual
- Case-study workflow visual

---

## ✨ Notes

- Canvas is fixed at `1280 x 720`
- Each HTML file represents one slide
- Styles can be shared using `shared.css`
- Keep designs simple for best PowerPoint compatibility
