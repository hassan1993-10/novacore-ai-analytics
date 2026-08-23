NovaCore V6 — Semantic Analytics + Brand Responsive UI

Replace only:
- app.py
- copilot_agent.py

What changed:
1. Semantic guard prevents unrelated dimensions from leaking into a result.
2. "By year" now aggregates at year grain unless another dimension is explicitly requested.
3. Deep questions (why / causes / analyze / recommendations) get an additional verified Copilot insight pass.
4. Standard questions keep the fast local insight path.
5. Smart chart renderer supports clean trend, ranking, and category comparison visuals.
6. Result table is now a white brand-consistent Supporting Data expander.
7. Black/unreadable buttons under suggested questions and Quick Tools are fixed.
8. Streamlit top strip is minimized.
9. Stronger NovaCore color/font system.
10. Responsive rules target laptop, tablet, and mobile; mobile input uses 16px to prevent auto zoom.

Do not change:
- excel_mcp.py
- requirements.txt
- Streamlit Secrets
- assets/
- data/
