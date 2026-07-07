# PRD — Aplikasi Laporan Akuisisi (Digital Forensic)

## Problem Statement
Redesign tampilan aplikasi Flask "Laporan Akuisisi" agar lebih modern, keren, simpel, dan profesional — tanpa mengubah fungsi apapun.

## Stack
- Flask 3.x + Jinja2 templates
- SQLite (akuisisi.db)
- python-docx / docxtpl untuk generate laporan Word
- Frontend: Bootstrap 5, Tom Select, Chart.js, Bootstrap Icons

## Redesign Delivered (Jan 2026)
- **Theme**: Sophisticated dark ("Forensic Dark") — deep midnight navy #0a0e17 + emerald #4ade80 + cyan #22d3ee accents.
- **Typography**: Bricolage Grotesque (display) + Instrument Sans (body) + JetBrains Mono (numeric) — no generic Inter/Roboto.
- **Layout**: Sticky left sidebar (260px) + top breadcrumb bar + content area with soft radial glow + subtle grid overlay backdrop.
- **Templates**:
  - `templates/base.html` — full rewrite (sidebar shell, dark-theme design tokens, Bootstrap overrides for buttons/cards/forms/tables/badges/alerts/Tom Select).
  - `templates/dashboard.html` — modern hero, glowing stat cards, dark-themed Chart.js (doughnut + line + bar), quick-action cards, recent-activity list.
  - `templates/list.html` — modern table w/ inline search + status filter + row count.
  - `templates/form.html` — new hero header, sticky bottom submit bar, refined Bootstrap Table inside cards; ALL field `name`s / `id`s / JS logic preserved (no functional change).

## Preserved
- All Flask routes and Python logic untouched (`app.py` unchanged)
- Every form field name, id, and JS function preserved for backend compatibility (OCR/hash/dasar dinamis/imei/foto handling — all intact)
- API endpoints untouched

## Known Environment Note
- Only in this hosted preview: `/api/*` requests go to a different port via Kubernetes ingress, so Chart.js data fetch may return 502 in preview. On local run (`python app.py`), all endpoints share one port and charts render correctly.

## Backlog / Future
- P1: light theme toggle (icon in topbar)
- P2: keyboard shortcut palette (⌘K) to jump between laporan
- P2: skeleton loaders for chart cards while data fetches
