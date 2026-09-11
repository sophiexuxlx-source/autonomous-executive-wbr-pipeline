"""
Autonomous Executive Weekly Business Review (WBR) & AI Intelligence Pipeline
=============================================================================
An end-to-end automated business intelligence system:
1. Queries operational metrics from SQLite database.
2. Synthesizes an executive-grade narrative via Google Gemini API.
3. Updates Microsoft Excel records, refreshes Pivot Tables, and exports native charts via COM automation.
4. Generates an executive dashboard (dispatched to Outlook or exported to standalone HTML).
"""

import os
import sys
import sqlite3
import datetime
import base64
import argparse
from dotenv import load_dotenv

# Ensure UTF-8 output on Windows terminal
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Directory of this script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Load environment variables from .env located in the script directory
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("AI_BUILDER_API_KEY")
AI_MODEL = os.getenv("DEFAULT_AI_MODEL", "gemini-3.6-flash")
RECIPIENT_EMAIL = os.getenv("EXECUTIVE_EMAIL_RECIPIENT", "leadership-team@example.com")

SHIFT_YEAR_TO_2024 = True  # Shift database dates from 2025 to 2024 to align with WBR chart timeline
TEMP_IMAGE_NAME = os.path.join(SCRIPT_DIR, 'temp_chart.png')
HTML_EXPORT_PATH = os.path.join(SCRIPT_DIR, 'wbr_executive_dashboard.html')


def parse_date(date_str, shift_to_2024=True):
    """
    Safely parse ISO date strings (YYYY-MM-DD) into datetime.datetime objects.
    If shift_to_2024 is True and the year is 2025, shifts it to 2024.
    """
    if not date_str:
        return None
    try:
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        if shift_to_2024 and dt.year == 2025:
            dt = dt.replace(year=2024)
        return dt
    except ValueError:
        return date_str


def get_gemini_client():
    """Initializes and returns the Google GenAI client if configured."""
    if not GEMINI_API_KEY or GEMINI_API_KEY.startswith("your_"):
        return None
    try:
        from google import genai
        return genai.Client(api_key=GEMINI_API_KEY)
    except Exception as e:
        print(f"[-] Warning initializing Google GenAI client: {e}")
        return None


def generate_ai_executive_summary(rows, headers):
    """
    Analyzes aggregated WBR metrics and invokes Google Gemini to
    synthesize an executive-level narrative for senior leadership.
    """
    print("\n--- [Stage 2] Generating AI Executive Synthesis ---")
    
    # Calculate key statistical aggregates from query results
    total_events = sum(row[2] for row in rows) if rows else 0
    category_counts = {}
    date_counts = {}
    
    for cat, dt, count in rows:
        category_counts[cat] = category_counts.get(cat, 0) + count
        date_counts[dt] = date_counts.get(dt, 0) + count
        
    dates_sorted = sorted(date_counts.keys())
    start_date = dates_sorted[0] if dates_sorted else "N/A"
    end_date = dates_sorted[-1] if dates_sorted else "N/A"
    peak_date = max(date_counts, key=date_counts.get) if date_counts else "N/A"
    peak_count = date_counts.get(peak_date, 0)
    
    # Format metrics text for the LLM
    category_breakdown_str = "\n".join(
        f"  - {cat}: {count:,} events ({count/total_events*100:.1f}%)" 
        for cat, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
    )
    
    metrics_summary = f"""Reporting Period: {start_date} to {end_date}
Total Operational Incidents: {total_events:,}
Peak Event Date: {peak_date} ({peak_count:,} events)
Category Breakdown:
{category_breakdown_str}"""

    print(f"[+] Aggregated Operational Metrics:\n{metrics_summary}\n")
    
    client = get_gemini_client()
    if not client:
        print("[!] No valid GEMINI_API_KEY found. Using rule-based structured executive template.")
        return f"""
        <p><strong>Executive Overview:</strong> A total of <strong>{total_events:,}</strong> operational incidents were logged across the reporting period ({start_date} to {end_date}).</p>
        <ul>
            <li><strong>Volume Peak:</strong> Peak incident volume occurred on <b>{peak_date}</b> with {peak_count:,} events.</li>
            <li><strong>Leading Category:</strong> Primary drivers concentrated in <b>{list(category_counts.keys())[0] if category_counts else 'N/A'}</b>.</li>
            <li><strong>Recommendation:</strong> Recommend conducting a root-cause anomaly review on peak dates and optimizing resource allocation.</li>
        </ul>
        """

    prompt = f"""You are an Executive Operations & Business Intelligence Architect presenting to C-level executives.
Analyze the following Weekly Business Review (WBR) metrics and produce a concise, high-impact executive summary.

Requirements:
1. Output ONLY clean HTML snippet (using <p>, <ul>, <li>, <strong> tags). Do NOT wrap in ```html code fences.
2. Structure into three sections:
   - <strong>Top-Line Performance Summary</strong>: High-level volume trends across the reporting period.
   - <strong>Critical Anomalies & Category Distribution</strong>: Key volume concentration and peak date spikes.
   - <strong>Strategic Leadership Recommendations</strong>: 2 concrete, prioritized action items for mitigation.
3. Keep the tone professional, authoritative, and data-backed.

Data:
{metrics_summary}"""

    models_to_try = [AI_MODEL, "gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite"]
    models_to_try = list(dict.fromkeys(models_to_try))

    for model in models_to_try:
        try:
            print(f"[+] Invoking Google Gemini API ({model})...")
            response = client.models.generate_content(
                model=model,
                contents=prompt
            )
            ai_content = response.text.strip()
            # Strip any markdown code fences if returned
            if ai_content.startswith("```"):
                lines = ai_content.splitlines()
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].startswith("```"):
                    lines = lines[:-1]
                ai_content = "\n".join(lines).strip()
            print("[✓] AI Executive Narrative generated successfully!")
            return ai_content
        except Exception as e:
            print(f"[-] Model {model} attempt notice: {e}")

    print("[!] Fallback to structured rule-based summary.")
    return f"""
    <p><strong>Executive Overview:</strong> Total of <strong>{total_events:,}</strong> incidents recorded from {start_date} to {end_date}.</p>
    <ul>
        <li><strong>Peak Volume:</strong> {peak_date} ({peak_count:,} events).</li>
        <li><strong>Key Incident Driver:</strong> {list(category_counts.keys())[0] if category_counts else 'Operational'}.</li>
    </ul>
    """


def update_excel_com(abs_excel_path, temp_image_path, headers, rows):
    """
    Executes Windows Excel COM automation to update Data sheet,
    refresh Pivot Table, and export the chart.
    """
    excel = None
    wb = None
    try:
        import win32com.client as win32
    except ImportError:
        print("[-] pywin32 not available on this platform. Skipping Excel COM automation.")
        return False

    try:
        print("\n--- [Stage 3] Windows Excel COM Automation ---")
        print("[+] Connecting to Excel.Application...")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        print(f"[+] Opening workbook: '{os.path.basename(abs_excel_path)}'...")
        wb = excel.Workbooks.Open(abs_excel_path)
        
        # Step A: Overwrite 'Data' sheet
        print("[+] Refreshing 'Data' worksheet with query rows...")
        sheet = wb.Worksheets(1)
        sheet.Cells.ClearContents()
        
        for col_idx, header in enumerate(headers, start=1):
            sheet.Cells(1, col_idx).Value = header
            
        for row_idx, row in enumerate(rows, start=2):
            for col_idx, val in enumerate(row, start=1):
                if isinstance(val, str) and len(val) == 10 and val.count('-') == 2:
                    parsed_val = parse_date(val, shift_to_2024=SHIFT_YEAR_TO_2024)
                    sheet.Cells(row_idx, col_idx).Value = parsed_val
                else:
                    sheet.Cells(row_idx, col_idx).Value = val

        # Step B: Refresh Pivot Table
        print("[+] Refreshing Pivot Table in 'Analysis' sheet...")
        ws_analysis = wb.Sheets('Analysis')
        pivot_table = ws_analysis.PivotTables(1)
        pivot_table.RefreshTable()

        # Step C: Export Chart
        print("[+] Locating chart in 'Visualization' tab...")
        ws_visual = wb.Sheets('Visualization')
        if ws_visual.ChartObjects().Count > 0:
            chart_obj = ws_visual.ChartObjects(1)
            chart = chart_obj.Chart
            print(f"[+] Exporting native chart to: '{temp_image_path}'...")
            chart.Export(Filename=temp_image_path)
        else:
            print("[-] Warning: No chart objects found in 'Visualization' tab.")

        wb.Close(SaveChanges=True)
        print("[✓] Excel workbook refreshed, updated, and saved successfully.")
        return True
    except Exception as e:
        print(f"[-] Excel COM Automation Notice: {e}")
        return False
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass


def build_dashboard_html(ai_summary_html, chart_cid_or_base64, is_inline_cid=True):
    """Compiles the executive-grade HTML dashboard styling."""
    img_tag = (
        f'<img src="cid:{chart_cid_or_base64}" alt="WBR Trend Visualization" style="max-width: 100%; border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">'
        if is_inline_cid
        else f'<img src="{chart_cid_or_base64}" alt="WBR Trend Visualization" style="max-width: 100%; border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">'
    )
    
    return f"""<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8">
        <title>Weekly Business Review (WBR) Executive Dashboard</title>
    </head>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.6; max-width: 800px; margin: 20px auto; background-color: #f8fafc; padding: 20px;">
        <div style="background: linear-gradient(135deg, #0284c7 0%, #1e40af 100%); color: white; padding: 24px; border-radius: 12px 12px 0 0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
            <h2 style="margin: 0; font-size: 22px; font-weight: 700;">Weekly Business Review (WBR) Executive Dashboard</h2>
            <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.9;">Autonomous BI Data Pipeline & AI Strategic Intelligence Synthesis</p>
        </div>
        
        <div style="border: 1px solid #e2e8f0; border-top: none; padding: 28px; border-radius: 0 0 12px 12px; background: #ffffff; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05);">
            <h3 style="color: #0f172a; margin-top: 0; border-bottom: 2px solid #f1f5f9; padding-bottom: 8px; font-size: 17px;">
                🤖 AI Executive Analysis & Strategic Highlights
            </h3>
            <div style="background: #f8fafc; border-left: 4px solid #0284c7; padding: 16px 20px; margin-bottom: 28px; border-radius: 0 8px 8px 0; font-size: 14px;">
                {ai_summary_html}
            </div>

            <h3 style="color: #0f172a; border-bottom: 2px solid #f1f5f9; padding-bottom: 8px; font-size: 17px;">
                📊 Performance Visualization (Excel Native)
            </h3>
            <div style="text-align: center; margin: 20px 0;">
                {img_tag}
            </div>

            <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 28px 0;">
            <p style="font-size: 12px; color: #64748b; margin-bottom: 0;">
                <i>Generated autonomously via WBR Pipeline (SQLite → Google Gemini API → Excel COM Automation → Executive Dashboard).</i>
            </p>
        </div>
    </body>
</html>"""


def dispatch_or_export(ai_summary_html, temp_image_path, recipient=RECIPIENT_EMAIL, force_html_export=False):
    """
    Sends email draft via Microsoft Outlook if available, or exports
    the standalone HTML dashboard directly to disk.
    """
    print("\n--- [Stage 4] Executive Dashboard Delivery ---")
    outlook_success = False

    if not force_html_export and sys.platform == "win32":
        try:
            import win32com.client as win32
            print("[+] Attempting Outlook COM connection...")
            try:
                outlook = win32.GetActiveObject('Outlook.Application')
            except Exception:
                outlook = win32.Dispatch('Outlook.Application')

            mail = outlook.CreateItem(0)
            mail.Subject = 'Executive Weekly Business Review (WBR) - AI Synthesis & Performance Report'
            mail.To = recipient

            if os.path.exists(temp_image_path):
                attachment = mail.Attachments.Add(Source=temp_image_path)
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "WBRChart")
                mail.HTMLBody = build_dashboard_html(ai_summary_html, "WBRChart", is_inline_cid=True)
            else:
                mail.HTMLBody = build_dashboard_html(ai_summary_html, "", is_inline_cid=False)

            mail.Display(False)
            print("[✓] Successfully initialized Outlook email draft ready for executive review!")
            outlook_success = True
        except Exception as e:
            print(f"[-] Outlook COM not available: {e}")

    # Fallback / Universal Export: Save standalone HTML report to disk
    chart_src = ""
    if os.path.exists(temp_image_path):
        with open(temp_image_path, "rb") as img_f:
            b64_data = base64.b64encode(img_f.read()).decode("utf-8")
            chart_src = f"data:image/png;base64,{b64_data}"
    
    html_content = build_dashboard_html(ai_summary_html, chart_src, is_inline_cid=False)
    with open(HTML_EXPORT_PATH, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    print(f"[✓] Standalone Executive Dashboard exported to: '{HTML_EXPORT_PATH}'")
    return True


def execute_pipeline(db_path, sql_path, excel_path, force_html=False):
    if not os.path.exists(db_path):
        print(f"Error: Database file '{db_path}' not found.")
        return False
    if not os.path.exists(sql_path):
        print(f"Error: SQL file '{sql_path}' not found.")
        return False

    temp_image_path = os.path.abspath(TEMP_IMAGE_NAME)
    abs_excel_path = os.path.abspath(excel_path)

    # 1. Read and Execute SQL
    print("--- [Stage 1] SQL Operational Ingestion ---")
    print(f"[+] Loading SQL query from '{os.path.basename(sql_path)}'...")
    with open(sql_path, 'r', encoding='utf-8') as f:
        query = f.read()

    print(f"[+] Querying operational database: '{os.path.basename(db_path)}'...")
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute(query)
        headers = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        print(f"[✓] Retrieved {len(rows)} operational incident records.")
    except sqlite3.Error as e:
        print(f"[-] SQLite Error: {e}")
        return False
    finally:
        if 'conn' in locals():
            conn.close()

    # 2. Generate AI Narrative
    ai_summary_html = generate_ai_executive_summary(rows, headers)

    # 3. Excel COM Automation
    if os.path.exists(abs_excel_path):
        update_excel_com(abs_excel_path, temp_image_path, headers, rows)
    else:
        print(f"[-] Excel template '{abs_excel_path}' not found. Continuing with report generation.")

    # 4. Dispatch Dashboard
    dispatch_or_export(ai_summary_html, temp_image_path, force_html_export=force_html)

    # 5. Clean up temporary chart
    if os.path.exists(temp_image_path):
        try:
            os.remove(temp_image_path)
        except Exception:
            pass

    return True


def main():
    parser = argparse.ArgumentParser(description="Autonomous Executive WBR Pipeline")
    parser.add_argument("--export-html", action="store_true", help="Force standalone HTML dashboard export")
    args = parser.parse_args()

    db = os.path.join(SCRIPT_DIR, 'events.db')
    sql = os.path.join(SCRIPT_DIR, 'data.sql')
    excel = os.path.join(SCRIPT_DIR, 'WBR.xlsx')

    print("=" * 70)
    print("  Autonomous Executive WBR Pipeline (Google Gemini API & Analytics)")
    print("=" * 70)
    print(f"AI Model:     {AI_MODEL}")
    print(f"Database:     {os.path.basename(db)}")
    print(f"Excel Model:  {os.path.basename(excel)}")

    success = execute_pipeline(db, sql, excel, force_html=args.export_html)
    if success:
        print("\n✅ Pipeline execution completed successfully!")
    else:
        print("\n❌ Pipeline encountered execution errors.")


if __name__ == '__main__':
    main()
