# Smart Salon CRM ✂️

A minimal, offline-first Customer Relationship Management (CRM) dashboard designed and prototyped for a local salon business.

## 🎯 The Problem
The salon was manually tracking customer visits and service due dates using traditional methods. This led to missed follow-ups, lost client retention, and difficulty in identifying overdue customers. They needed a lightweight, internet-independent solution that doesn't require a steep learning curve.

## 💡 The Solution
I designed and prototyped a local Python-based CRM using **Streamlit** and **Pandas**. Instead of a complex cloud database, it uses a simple Excel sheet as the backend, allowing the salon staff to easily edit or backup data without any technical knowledge.

### ✨ Key UX & UI Highlights
*   **Ink-Friendly & Minimal Design:** Dark mode layout (`#0F172A` background) with a subtle grid structure, designed to reduce eye strain for receptionists looking at the screen all day.
*   **Zero-Friction Search:** An instantly accessible search bar to find clients rapidly by name or number without refreshing the page.
*   **Automated Due-Date Logic:** The system automatically calculates the next due date based on gender and specific service type (e.g., Haircut, Facial) and moves clients to 'Due Today' or 'Overdue' dynamically.
*   **Visual Hierarchy:** Distinct visual indicators (red dots for overdue, green badges for new clients) to grab user attention immediately.

## 🛠️ Tech Stack Used (For Prototyping)
*   **Frontend & Logic:** Python (Streamlit)
*   **Database Management:** Pandas (Excel integration)
*   **UI/Styling:** Custom CSS injected via Markdown

## 🚀 How to Run Locally
1. Clone this repository.
2. Ensure you have Python installed.
3. Install the required libraries: `pip install streamlit pandas openpyxl streamlit-option-menu`
4. Run the app: `streamlit run app.py`

*(Note: The uploaded Excel file contains dummy data for privacy and portfolio demonstration purposes. The original business data is kept strictly confidential.)*
