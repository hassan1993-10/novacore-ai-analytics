NovaCore V2 - Responsive + Performance

Replace only:
- app.py
- copilot_agent.py

Do not change excel_mcp.py, assets, data, requirements, or Streamlit Secrets.

Performance:
- Excel schema cached per app process.
- Data Overview cached for 30 minutes.
- Analytics now uses 1 Copilot call instead of 2.
- Refresh Data clears caches.

Responsive:
- prevents horizontal page zoom/overflow
- mobile input is 16px to prevent iPhone auto-zoom
- mobile columns stack vertically
- larger/darker Arabic typography
- charts use responsive Plotly mode
