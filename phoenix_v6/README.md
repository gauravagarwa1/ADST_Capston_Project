# Project Phoenix V9 — Purple Data & Learning Workspace

A local Flask + SQLite prototype for Project Phoenix: student project management, teacher-led evaluation, career guidance, gamification, notifications and a multi-dataset analytics center.

## V9 highlights
- Professional purple/lavender canvas with white surfaces and accessible dark text.
- Teacher workflow: create assignment → student receives it → student submits → teacher evaluates → publish final grade → student receives notification and sees the final result.
- Student workflow: assigned projects, draft/submit, files, AI preliminary feedback, teacher final grade, points and career guidance.
- Analytics: live SQLite charts plus separate uploaded dataset dashboards.
- Data Center: upload multiple CSV/XLS/XLSX/PDF/DOCX/PPTX/images/ZIP files without replacing earlier uploads.
- Dataset dashboards: generated charts, schema, data quality, outlier estimate, missing values, duplicate rows, searchable data explorer and CSV export.
- Notifications: assignment created, submission received and final grade published events.
- Settings and Help pages are functional.
- Responsive UI, hover states, transitions and reduced-motion friendly progressive reveal.

## Run
```bash
pip install -r requirements.txt
python app.py
```
Open `http://127.0.0.1:5000`.

Demo accounts:
- Student: `student@phoenix.local` / `student123`
- Teacher: `teacher@phoenix.local` / `teacher123`

## Important product rule
AI scores are preliminary. The teacher publishes the final grade. Uploaded datasets are separate analytical inputs and never overwrite the live Phoenix database.
