# app.py
import streamlit as st
from datetime import date
from docx import Document
import os
from math import log, sqrt, exp, erf
import pandas as pd

# ---------------------------
# Streamlit page setup
# ---------------------------
st.set_page_config(page_title="Term Sheet with Scenario Table", layout="wide")
st.title("Term Sheet Generator — Insert Scenario Table into Word")

# ---------------------------
# Sidebar inputs
# ---------------------------
st.sidebar.header("Template")

# Locate template file next to this script (include .docx by default)
script_dir = os.path.dirname(os.path.abspath(__file__))
default_template = os.path.join(script_dir, "TermSheetTemplate.docx")
template_path = st.sidebar.text_input("Template path", value=default_template)

# Ensure .docx extension is present
if template_path and not template_path.lower().endswith(".docx"):
    template_path = template_path + ".docx"

st.sidebar.header("Trade inputs")
client_name = st.sidebar.text_input("Client name", value="Client A")
valuation_date = st.sidebar.date_input("Valuation date", value=date.today())
maturity_date = st.sidebar.date_input("Maturity date", value=date.today())
spot_now = st.sidebar.number_input("Spot (current)", value=100.0, format="%.4f")
strike = st.sidebar.number_input("Strike", value=80.0, format="%.4f")
implied_vol = st.sidebar.number_input("Implied vol (%)", value=20.0)
rate = st.sidebar.number_input("Risk-free rate (%)", value=5.0)
notional = st.sidebar.number_input("Notional (units)", value=1.0)

st.sidebar.header("Scenario (choose one input method)")
mode = st.sidebar.radio("Scenario input mode", ["Explicit list", "Start/Stop/Step"], index=1)

if mode == "Explicit list":
    spot_list_txt = st.sidebar.text_input(
        "Spot at maturity values (comma separated)",
        value="78,80,82,84,86,88,90,92,94,96",
        help="Enter absolute future spot levels, e.g. 78,80,82,..."
    )
    def parse_list(txt):
        vals = []
        for part in txt.split(","):
            p = part.strip()
            if p == "":
                continue
            try:
                vals.append(float(p))
            except:
                pass
        return vals
    spot_levels = parse_list(spot_list_txt)
else:
    start = st.sidebar.number_input("Start (Spot at maturity)", value=78.0, format="%.4f")
    stop = st.sidebar.number_input("Stop (Spot at maturity)", value=96.0, format="%.4f")
    step = st.sidebar.number_input("Step", value=2.0, format="%.4f")
    spot_levels = []
    v = start
    if step == 0:
        step = 1.0
    # handle positive or negative step safely
    if step > 0:
        while v <= stop + 1e-9:
            spot_levels.append(round(v, 8))
            v = v + step
    else:
        while v >= stop - 1e-9:
            spot_levels.append(round(v, 8))
            v = v + step

# ---------------------------
# Pricing helpers (no SciPy)
# ---------------------------
def norm_cdf(x: float) -> float:
    return 0.5 * (1.0 + erf(x / sqrt(2.0)))

def bs_call_price(S, K, T, r, sigma):
    if T <= 0:
        return max(S - K, 0.0)
    if sigma <= 0:
        return max(S - K * exp(-r * T), 0.0)
    d1 = (log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * sqrt(T))
    d2 = d1 - sigma * sqrt(T)
    return S * norm_cdf(d1) - K * exp(-r * T) * norm_cdf(d2)

# ---------------------------
# Compute premium (for summary)
# ---------------------------
T_days = max((maturity_date - valuation_date).days, 0)
T = T_days / 365.0
sigma = implied_vol / 100.0
r = rate / 100.0
premium = bs_call_price(spot_now, strike, T, r, sigma)

scenario_rows = []
for ST in spot_levels:
    payoff = max(ST - strike, 0.0)
    payoff_notional = payoff * notional
    scenario_rows.append((ST, payoff, payoff_notional))

# Show scenario preview in UI
st.subheader("Scenario table preview")
df = pd.DataFrame(scenario_rows, columns=["Spot at maturity", "Option payoff", "Payoff (× Notional)"])
st.dataframe(df)

# ---------------------------
# Word: insert table at placeholder
# ---------------------------
def insert_table_at_placeholder(doc: Document, placeholder: str, headers: list, rows: list):
    # find paragraph that contains the placeholder
    target_p = None
    for p in doc.paragraphs:
        if placeholder in p.text:
            target_p = p
            break

    # create a new table (this appends at the end; we'll move it)
    tbl = doc.add_table(rows=1, cols=len(headers))
    try:
        tbl.style = 'Table Grid'
    except Exception:
        pass
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = str(h)

    for r in rows:
        row_cells = tbl.add_row().cells
        for i, val in enumerate(r):
            if isinstance(val, float):
                # pretty formatting
                if abs(val - int(val)) < 1e-9:
                    row_cells[i].text = f"{int(val)}"
                else:
                    row_cells[i].text = f"{val:.4f}"
            else:
                row_cells[i].text = str(val)

    # move the created table to the position after the placeholder paragraph (if found)
    tbl_element = tbl._tbl
    body = doc._body._element

    # remove the table element from the end (where add_table placed it)
    try:
        body.remove(tbl_element)
    except Exception:
        # if removal fails, continue (we will still try to insert)
        pass

    if target_p is not None:
        p_element = target_p._p
        # find paragraph index inside body
        idx = list(body).index(p_element)
        body.insert(idx + 1, tbl_element)
        # remove the placeholder paragraph (or clear its text)
        try:
            body.remove(p_element)
        except Exception:
            # as fallback, clear the paragraph runs
            for run in target_p.runs:
                run.text = ""
    else:
        # append at the end
        body.append(tbl_element)

# ---------------------------
# Replace placeholders
# ---------------------------
def replace_simple_placeholders(doc: Document, replacements: dict):
    # paragraphs
    for p in doc.paragraphs:
        text = p.text
        for k, v in replacements.items():
            if k in text:
                # clear existing runs safely
                for run in p.runs:
                    run.text = ""
                p.add_run(str(v))
                # update text var to avoid multiple replacements
                text = p.text

    # tables (cells)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text
                    for k, v in replacements.items():
                        if k in text:
                            for run in p.runs:
                                run.text = ""
                            p.add_run(str(v))
                            text = p.text

# ---------------------------
# Generate Word button
# ---------------------------
if st.button("Generate Term Sheet (Word with scenario table)"):
    # helpful debugging: list files in script dir when template not found
    if not os.path.exists(template_path):
        st.error(f"Template not found at: {template_path}")
        st.info("Files in script directory:")
        try:
            files = os.listdir(script_dir)
            st.write(files)
        except Exception as e:
            st.write("Failed to list directory:", e)
    else:
        try:
            doc = Document(template_path)
        except Exception as e:
            st.error(f"Failed to open template: {e}")
        else:
            headers = ["Spot at maturity", "Option payoff", "Payoff (× Notional)"]
            insert_table_at_placeholder(doc, "{{ScenarioTable}}", headers, scenario_rows)

            replacements = {
                "{{ClientName}}": client_name,
                "{{ValuationDate}}": valuation_date.isoformat(),
                "{{MaturityDate}}": maturity_date.isoformat(),
                "{{Spot}}": f"{spot_now:,.4f}",
                "{{Strike}}": f"{strike:,.4f}",
                "{{ImpliedVol}}": f"{implied_vol:.2f}%",
                "{{RiskFreeRate}}": f"{rate:.2f}%",
                "{{Premium}}": f"{premium:,.4f}",
                "{{Notional}}": f"{notional}"
            }
            replace_simple_placeholders(doc, replacements)

            out_fname = f"TermSheet_{client_name.replace(' ', '_')}_{valuation_date.isoformat()}.docx"
            out_path = os.path.join(script_dir, out_fname)
            doc.save(out_path)

            st.success(f"Term sheet generated: {out_path}")
            with open(out_path, "rb") as f:
                st.download_button(
                    "Download Term Sheet",
                    data=f.read(),
                    file_name=out_fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

st.markdown("---")
st.info("Ensure your template file is named **TermSheetTemplate.docx** (or update the Template path) and contains the placeholder `{{ScenarioTable}}` where you want the scenario table inserted.")
