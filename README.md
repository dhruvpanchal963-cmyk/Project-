FITX UX Analytics Dashboard
A professional Streamlit UX/UI analytics dashboard for the FITX project. The app reads the supplied Excel workbook and turns the analysis sheets into an interactive product-analytics experience with KPIs, filters, conversion funnels, a navigation Sankey diagram, form/error analysis, device insights, exit/scroll analysis, correlation heatmaps, and prioritized UX recommendations.
Key insights visible in the workbook
913 unique users, 2,027 sessions, and 53,523 events.
Average activity is about 2.22 sessions per user and 26.41 events per session.
Mobile has the highest displayed repeat-user rate (27.59%), ahead of desktop (25.28%) and tablet (23.38%).
The Classes funnel shows a major conversion problem: 677 Homepage → 107 Classes → 11 Confirmed Booking.
`Join Now` is the strongest displayed CTA by reach/click volume, with 290 unique users and 373 clicks.
Form abandonment and validation errors are important friction areas; required fields and invalid input/email/phone patterns deserve UX attention.
The homepage (`index.html`) is the strongest displayed page by reach, so it should remain the main gateway into high-value journeys.
Dashboard sections
Executive Overview — headline KPIs, executive signals, CTA performance, device retention, funnels and form abandonment.
Audience & Devices — engagement segments and device-level behavior.
Navigation — top page-to-page transitions and an interactive Sankey flow.
Conversion — selectable conversion funnels, CTA reach vs clicks, and button interactions.
Forms & Errors — completion/abandonment rates and validation-error analysis.
Exit & Scroll — exit hotspots and scroll-depth behavior.
Correlation — correlation heatmap and strongest behavioral relationships.
UX Recommendations — filterable high/medium/low-priority insights from the workbook.
Raw Data — inspect any workbook sheet and download it as CSV.
Project structure
```text
FITX-dashboard/
├── app.py
├── requirements.txt
├── README.md
└── FITX_UX_UI_Analysis_Dashboard_Aligned.xlsx
```
> Keep the Excel file in the same folder as `app.py`. The app also supports manual Excel upload from the sidebar.
Run locally
Create a virtual environment (recommended), install the packages, then run Streamlit:
```bash
pip install -r requirements.txt
streamlit run app.py
```
Streamlit will print a local URL in the terminal. Open it in your browser.
Deploy with GitHub + Streamlit Community Cloud
Create a new GitHub repository, for example `fitx-ux-analytics-dashboard`.
Upload these files to the repository:
`app.py`
`requirements.txt`
`README.md`
`FITX_UX_UI_Analysis_Dashboard_Aligned.xlsx`
Commit and push the files to the `main` branch.
Sign in to Streamlit Community Cloud with GitHub.
Create a new app and select your repository.
Set the main file path to `app.py`.
Deploy.
No API key or secret is required because the dashboard reads the Excel workbook directly from the repository.
Recommended GitHub repository description
> Interactive Streamlit dashboard for FITX UX/UI analytics with behavioral KPIs, navigation Sankey flows, conversion funnels, device analysis, form friction, validation errors, correlation analysis and prioritized UX recommendations.
Suggested project title for college/project presentation
FITX UX Intelligence: Interactive Behavioral Analytics and Conversion Optimization Dashboard
Suggested objectives
Analyze user engagement, retention and device behavior across the FITX website.
Visualize navigation pathways and identify important user journeys.
Measure conversion drop-offs in class-booking and membership funnels.
Identify UX friction through form abandonment, validation errors and exit behavior.
Convert behavioral data into prioritized, evidence-based UX recommendations.
Presentation conclusion
The FITX UX Analytics Dashboard converts a multi-sheet analytics workbook into an interactive decision-support system. The analysis shows strong overall activity but also exposes important conversion and form-friction problems. Mobile demonstrates the strongest repeat-user rate, while the Classes and Membership funnels lose a large share of users before completion. By combining navigation flows, CTA performance, device behavior, forms, validation errors, exits, scroll behavior and correlations in one Streamlit interface, the project provides a practical foundation for improving user experience and increasing conversion.
Notes
Funnel and correlation values are read from the workbook as supplied. Where the workbook itself includes scope notes or retained analysis, the dashboard preserves those values rather than inventing new user-level observations.
Correlation indicates association, not causation.
For a public GitHub repository, confirm that the Excel workbook contains no confidential or personally identifiable information before committing it.
