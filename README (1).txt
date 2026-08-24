NovaCore V7 — Driver Charts + Wide Responsive Workspace

Replace only:
- app.py
- copilot_agent.py

Changes:
- Driver evidence is passed from the agent to the UI.
- WHY/root-cause questions render contribution bar charts for Region/Product/etc.
- The misleading regular line chart is hidden when verified driver charts exist.
- Main workspace expands on desktop and when the Streamlit sidebar is collapsed.
- Tablet/mobile width remains responsive.
- Existing semantic metric resolution and executive insight logic are preserved.
