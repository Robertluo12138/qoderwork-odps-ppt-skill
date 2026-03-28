Use the `odps-ppt-report-builder` skill in zero-setup mode.

Rules:
- Assume I am using Qoder Work, not Qoder IDE.
- First decide whether this task can be completed with zero setup or whether the shared Python runtime is needed.
- Do not ask me to install Python, git, Node.js, or any package manager unless absolutely necessary.
- Never install dependencies into the Qoder Work application folder.
- If Python is needed, install packages into the shared runtime under `~/.qoderwork-runtime/` on macOS or `%USERPROFILE%\.qoderwork-runtime\` on Windows.
- Read `assets/company_private_brief.md` first. If it does not exist, ask me to create it from `assets/company_private_brief.template.md`.
- Use only the exact ODPS commands or SQL defined in the private brief.
- Be extremely explicit for Windows users. Give one command at a time and avoid hidden assumptions.
- If a command differs between macOS and Windows, show both and clearly label them.
- Save all intermediate files into one task folder under the current working directory.
- Create at least these files:
  - `raw_schema.txt`
  - `report_context.md`
  - `slide_plan.generated.json`
- Follow the PPT outline in the private brief exactly.
- Do not expose sensitive fields or hidden dimensions in the final report.
- If native `.pptx` rendering is blocked on this machine, produce a fully structured slide manuscript instead of stopping.
- If Python is available, prefer using the bundled scripts to reduce reasoning burden and keep the output stable.
- If I ask to preview the final style quickly, generate an HTML preview or PNG preview contact sheet in addition to the PPT.

Expected output:
- final report file path
- key findings summary
- any blocked steps or missing data
