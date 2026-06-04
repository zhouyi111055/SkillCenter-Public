# SkillCenter-Public

## Overview

SkillCenter-Public is a centralized repository for reusable public skills, templates, prompts, scripts, and automation assets.

The purpose of this repository is to:

* Collect and manage reusable AI skills.
* Provide standardized skill structures.
* Share public skills across multiple projects.
* Maintain version history and documentation.
* Support Codex and other AI-assisted development workflows.

---

## Repository Structure

```text
SkillCenter-Public
│
├─ ppt-master
├─ excel-master
├─ html-master
├─ svg-master
├─ docx-master
├─ pdf-master
│
└─ README.md
```

Each skill should be maintained as an independent directory.

---

## Skill Directory Standard

Each skill should follow the structure below:

```text
skill-name
│
├─ SKILL.md
├─ README.md
├─ examples
├─ templates
├─ scripts
└─ output-samples
```

### SKILL.md

Defines:

* Purpose
* Inputs
* Outputs
* Usage instructions
* Limitations
* Codex guidance

### README.md

Human-readable documentation.

### examples

Sample input files.

### templates

Reusable templates.

### scripts

Automation scripts and utilities.

### output-samples

Expected output examples.

---

## Skill Naming Convention

Recommended naming style:

```text
ppt-master
excel-master
html-master
svg-master
docx-master
pdf-master
```

Avoid:

```text
MySkill
NewSkill
Skill001
TestSkill
```

Names should clearly describe the capability.

---

## Usage with Codex

Example:

```text
Use skill:

/ppt-master

Input:

/Projects/PF_Project/Input

Output:

/Projects/PF_Project/Output

Please read:

/ppt-master/SKILL.md

before starting.
```

---

## Version Management

Recommended:

```text
main
feature/*
```

Examples:

```text
feature/improve-layout
feature/add-template-support
feature/ppt-export-v2
```

Merge into main after validation.

---

## Contribution Rules

Before adding a skill:

1. Create a dedicated folder.
2. Add SKILL.md.
3. Add at least one example.
4. Add usage documentation.
5. Verify compatibility with Codex.

---

## Public Repository Policy

This repository contains only:

* Public assets
* Open-source skills
* Generic templates
* Reusable prompts

Do NOT upload:

* Customer information
* Internal documents
* Credentials
* Confidential templates
* Proprietary source code

Such assets must be stored in SkillCenter-Private.

---

## Future Expansion

Planned categories:

```text
Presentation Skills
Document Skills
Excel Skills
Data Analysis Skills
Development Skills
Testing Skills
Migration Skills
AI Agent Skills
```

The repository should serve as the central public skill library for all future projects.

## gpt-image2-ppt-skills

Path:

`/gpt-image2-ppt-skills/`

Usage for Web Codex:

When creating editable PowerPoint from an image, first read this folder and refer to its prompt, layout rules, template clone logic, and PPT generation examples.

Do not output full-slide screenshots.
Use editable PowerPoint elements such as text boxes, shapes, tables, lines, arrows, and SVG icons.

# Image to Editable PPT for Web Codex

Before generating PPTX, read:

`/gpt-image2-ppt-skills/`

Use it as reference for:
- image layout analysis
- template clone logic
- editable PPT generation
- style restoration
- shape reconstruction

Requirements:
1. Generate editable PPTX.
2. Do not use one full-page image as slide background.
3. Recreate text as text boxes.
4. Recreate rectangles, lines, arrows, cards, tables, and icons as editable objects.
5. Keep original layout, spacing, color, and hierarchy as close as possible.
6. Prefer PptxGenJS or python-pptx.
7. Output final `.pptx` file.
