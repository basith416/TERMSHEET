# app.py
import streamlit as st
from docx import Document
import io
import datetime
import pkgutil

st.set_page_config(page_title="Term Sheet Generator", layout="wide")

st.title("Term Sheet Generator — Word template + Scenario Table")

# --- Helper functions -----------------------------------------------------

def replace_text_in_paragraphs(doc, placeholder, replacement):
    """Replace placeholder text in all paragraphs and runs."""
    for para in doc.paragraphs:
        if placeholder in para.text:
            # replace in runs to preserve styling where possible
            for run in para.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(replacement))

def replace_text_in_tables(doc, placeholder, replacement):
    """Replace placeholder text inside table cells."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    # simple approach: replace full cell.text preserving minimal formatting loss
                    cell_text = cell.text.replace(placeholder, str(replacement))
                    # clear cell and add a single paragraph with replaced text
                    cell.clear()
                    cell_para = cell.paragraphs[0]
                    cell_para.add_run(cell_text)

def find_paragraph_with_placeholder(doc, placeholder):
    """Return the paragraph object that contains the placeholder or None."""
    for para in doc.paragraphs:
        if placeholder in para.text:
            return para
    # also search in tables (cells) and return the first cell paragraph if found
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if placeholder in para.text:
                        return para
    return None

# Note: python-docx doesn't have a public API to "insert table after paragraph" directly.
# We use the low-level _p and table._tbl elements to move the created table next to the placeholder paragraph.
def insert_table_after_paragraph(doc, paragraph, data, col_names=None):
    """
    Insert a table after the given paragraph.
    data: list of rows (each row is a list of cell values)
    col_names: optional list of header names
    """
    # create table at the end of document (we will move it)
    nrows = len(data) + (1 if col_names else 0)
    ncols = len(data[0]) if data else (len(col_names) if col_names else 1)
    table = doc.add_table(rows=nrows, cols=ncols)
    table.style = 'Table Grid'

    row_idx = 0
    if col_names:
        hdr_cells = table.rows[row_idx].cells
        for c, name in enumerate(col_names):
            hdr_cells[c].text = str(name)
        row_idx += 1

    for r, row in enumerate(data):
        cells = table.rows[row_idx + r].cells
        for c, val in enumerate(row):
            cells[c].text = str(val)

    # Move the table element to just after the paragraph
    try:
        paragraph._p.addnext(table._tbl)
    except Exception as e:
        # If moving fails for any reason, leave the table at the end and log
        st.warning(f"Could not insert table at placeholder location; table appended at the end. ({e})")

# small helper to clear cell content (python-docx lacks direct clear so we remove paragraphs)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
def clear_cell(cell):
    cell._tc.clear_content()

# -------------------------------------------------------------------------

st.markdown("**Instructions:** Upload your Word template (docx) containing placeholders: "
            "`{{ClientName}}`, `{{Strike}}`, `{{Spot}}`, and `{{ScenarioTable}}` (for the table).")

# Sidebar inputs
st.sidebar.header("Inputs")
client_name = st.sidebar.text_input("Client name", value="ACME Ltd.")
strike = st.sidebar.number_input("Strike", value=80.0, step=0.5, format="%.2f")
spot = st.sidebar.number_input("Spot (current)", value=78.0, step=0.5, format="%.2f")
min_spot = st.sidebar.number_input("Scenario: min spot", value=max(0.0, strike - 40.0), step=0.5, format="%.2f")
max_spot = st.sidebar.number_input("Scenario: max spot", value=strike + 40.0, step=0.5, format="%.2f")
step_spot = st.sidebar.number_input("Scenario step", value=2.0, step=0.1, format="%.2f")
template_file = st.file_uploader("Upload Word template (.docx)", type=["docx"])

st.write("### Preview inputs")
st.write(f"**Client**: {client_name}  •  **Strike**: {strike}  •  **Spot**: {spot}")
st.write(f"Scenario from {min_spot} to {max_spot} step {step_spot}")

if not template_file:
    st.warning("Upload a Word template file (.docx) with placeholders to enable document generation.")
    st.stop()

# Load the template into python-docx Document
try:
    template_doc = Document(template_file)
except Exception as e:
    st.error(f"Failed to read the uploaded docx template: {e}")
    st.stop()

# Build scenario data (Spot at maturity and payoff)
spots = []
val = float(min_spot)
max_sp = float(max_spot)
s_step = float(step_spot) if step_spot > 0 else 1.0
while val <= max_sp + 1e-9:
    spots.append(round(val, 2))
    val += s_step

# compute payoff for a call
def call_payoff(spot_val, strike_val):
    return round(max(0.0, spot_val - strike_val), 2)

scenario_rows = []
for s in spots:
    payoff = call_payoff(s, strike)
    scenario_rows.append([s, payoff])

# Replace simple placeholders
replace_text_in_paragraphs(template_doc, "{{ClientName}}", client_name)
replace_text_in_paragraphs(template_doc, "{{Strike}}", "{:.2f}".format(strike))
replace_text_in_paragraphs(template_doc, "{{Spot}}", "{:.2f}".format(spot))
replace_text_in_tables(template_doc, "{{ClientName}}", client_name)
replace_text_in_tables(template_doc, "{{Strike}}", "{:.2f}".format(strike))
replace_text_in_tables(template_doc, "{{Spot}}", "{:.2f}".format(spot))

# Insert scenario table at placeholder location, or append if not found
placeholder = "{{ScenarioTable}}"
para = find_paragraph_with_placeholder(template_doc, placeholder)
if para:
    # remove the placeholder text from the paragraph first
    for run in para.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, "")
    insert_table_after_paragraph(template_doc, para, scenario_rows, col_names=["Spot at Maturity", "Payoff"])
else:
    # append at end
    insert_table_after_paragraph(template_doc, template_doc.paragraphs[-1], scenario_rows, col_names=["Spot at Maturity", "Payoff"])
    st.info("Placeholder {{ScenarioTable}} not found — scenario table appended at document end.")

# Save to bytes buffer
output = io.BytesIO()
template_doc.save(output)
output.seek(0)

# Create a filename
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"TermSheet_{client_name.replace(' ', '_')}_{timestamp}.docx"

st.success("Document ready — click the button below to download.")
st.download_button(label="Download Term Sheet (.docx)", data=output.getvalue(),
                   file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
