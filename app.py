# app.py (updated: extra placeholders + "Generate Term Sheet" button)
import streamlit as st
from docx import Document
import io
import datetime

st.set_page_config(page_title="Term Sheet Generator", layout="wide")
st.title("Term Sheet Generator — Word template + Scenario Table")

# -------------------- helper functions --------------------

def replace_text_in_paragraphs_full(doc, placeholder, replacement):
    """
    Replace placeholder text in all paragraphs by operating on the full paragraph.text.
    Note: this replaces the paragraph runs with a single run (styling in that paragraph may be lost).
    """
    for para in doc.paragraphs:
        if placeholder in para.text:
            new_text = para.text.replace(placeholder, str(replacement))
            para.clear()
            para.add_run(new_text)

def replace_text_in_cell_paragraphs_full(cell, placeholder, replacement):
    """Replace placeholder inside each paragraph in a single table cell."""
    for para in cell.paragraphs:
        if placeholder in para.text:
            new_text = para.text.replace(placeholder, str(replacement))
            para.clear()
            para.add_run(new_text)

def replace_text_in_tables_full(doc, placeholder, replacement):
    """Replace placeholder text inside every table cell (full-text replacement)."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_in_cell_paragraphs_full(cell, placeholder, replacement)

def replace_in_headers_and_footers(doc, placeholder, replacement):
    """Replace placeholders inside headers and footers as well (common place for templates)."""
    for section in doc.sections:
        header = section.header
        footer = section.footer
        # header paragraphs
        for para in header.paragraphs:
            if placeholder in para.text:
                new_text = para.text.replace(placeholder, str(replacement))
                para.clear()
                para.add_run(new_text)
        # header tables
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_text_in_cell_paragraphs_full(cell, placeholder, replacement)
        # footer paragraphs
        for para in footer.paragraphs:
            if placeholder in para.text:
                new_text = para.text.replace(placeholder, str(replacement))
                para.clear()
                para.add_run(new_text)
        # footer tables
        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_text_in_cell_paragraphs_full(cell, placeholder, replacement)

def find_paragraph_with_placeholder(doc, placeholder):
    """Return the paragraph object that contains the placeholder or None."""
    for para in doc.paragraphs:
        if placeholder in para.text:
            return para
    # search inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if placeholder in para.text:
                        return para
    # search headers/footers
    for section in doc.sections:
        for para in section.header.paragraphs:
            if placeholder in para.text:
                return para
        for para in section.footer.paragraphs:
            if placeholder in para.text:
                return para
    return None

def insert_table_after_paragraph(doc, paragraph, data, col_names=None, preferred_style_name="Table Grid"):
    """
    Insert a table after the given paragraph.
    data: list of rows (each row is a list of cell values)
    col_names: optional list of header names
    preferred_style_name: style to try to apply; if not present we skip style
    """
    # create table at end of document
    nrows = len(data) + (1 if col_names else 0)
    ncols = len(data[0]) if data else (len(col_names) if col_names else 1)
    table = doc.add_table(rows=nrows, cols=ncols)
    # try to set style safely (some templates may not have named styles)
    try:
        table.style = preferred_style_name
    except Exception:
        pass

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

    # attempt to move the table right after the paragraph
    try:
        paragraph._p.addnext(table._tbl)
    except Exception:
        st.warning("Could not insert table at the placeholder location; table appended at the end.")

# -------------------- UI & inputs --------------------

st.markdown("**Instructions:** Upload your Word template (.docx) containing placeholders: "
            "`{{ClientName}}`, `{{ValuationDate}}`, `{{MaturityDate}}`, `{{Spot}}`, `{{Strike}}`, `{{Premium}}`, "
            "`{{Notional}}`, `{{ImpliedVol}}`, and `{{ScenarioTable}}`.")

st.sidebar.header("Inputs")
client_name = st.sidebar.text_input("Client name", value="ACME Ltd.")
valuation_date = st.sidebar.date_input("Valuation date", value=datetime.date.today())
maturity_date = st.sidebar.date_input("Maturity date", value=(datetime.date.today() + datetime.timedelta(days=30)))
strike = st.sidebar.number_input("Strike", value=80.0, step=0.5, format="%.2f")
spot = st.sidebar.number_input("Spot (current)", value=78.0, step=0.5, format="%.2f")
premium = st.sidebar.number_input("Premium", value=0.0, step=0.01, format="%.4f")
notional = st.sidebar.number_input("Notional", value=1000000.0, step=1000.0, format="%.2f")
implied_vol = st.sidebar.number_input("Implied vol (in %)", value=12.0, step=0.1, format="%.2f")

min_spot = st.sidebar.number_input("Scenario: min spot", value=max(0.0, strike - 40.0), step=0.5, format="%.2f")
max_spot = st.sidebar.number_input("Scenario: max spot", value=strike + 40.0, step=0.5, format="%.2f")
step_spot = st.sidebar.number_input("Scenario step", value=2.0, step=0.1, format="%.2f")

template_file = st.file_uploader("Upload Word template (.docx)", type=["docx"])

st.write("### Preview inputs")
st.write(f"**Client**: {client_name}  •  **Strike**: {strike}  •  **Spot**: {spot}")
st.write(f"Valuation date: {valuation_date}  •  Maturity date: {maturity_date}")
st.write(f"Premium: {premium}  •  Notional: {notional}  •  Implied vol: {implied_vol}%")
st.write(f"Scenario from {min_spot} to {max_spot} step {step_spot}")

if not template_file:
    st.warning("Upload a Word template file (.docx) with placeholders to enable document generation.")
    st.stop()

# Only generate when user clicks button
if st.button("Generate Term Sheet"):

    # Load the template into python-docx Document
    try:
        template_doc = Document(template_file)
    except Exception as e:
        st.error(f"Failed to read the uploaded docx template: {e}")
        st.stop()

    # Build scenario data (Spot at maturity and payoff for a call)
    spots = []
    val = float(min_spot)
    max_sp = float(max_spot)
    s_step = float(step_spot) if step_spot > 0 else 1.0
    # generate discrete spots
    while val <= max_sp + 1e-9:
        spots.append(round(val, 2))
        val += s_step

    def call_payoff(spot_val, strike_val):
        return round(max(0.0, spot_val - strike_val), 2)

    scenario_rows = []
    for s in spots:
        payoff = call_payoff(s, strike)
        scenario_rows.append([s, payoff])

    # Prepare placeholder values (format dates and numeric display nicely)
    placeholders = {
        "{{ClientName}}": client_name,
        "{{ValuationDate}}": valuation_date.strftime("%Y-%m-%d"),
        "{{MaturityDate}}": maturity_date.strftime("%Y-%m-%d"),
        "{{Strike}}": "{:.2f}".format(strike),
        "{{Spot}}": "{:.2f}".format(spot),
        "{{Premium}}": "{:.4f}".format(premium),
        "{{Notional}}": "{:,.2f}".format(notional),
        "{{ImpliedVol}}": "{:.2f}%".format(implied_vol),
    }

    # Replace placeholders in paragraphs, tables, headers/footers
    for ph, val in placeholders.items():
        replace_text_in_paragraphs_full(template_doc, ph, val)
        replace_text_in_tables_full(template_doc, ph, val)
        replace_in_headers_and_footers(template_doc, ph, val)

    # Insert scenario table at placeholder or append
    placeholder = "{{ScenarioTable}}"
    para = find_paragraph_with_placeholder(template_doc, placeholder)
    if para:
        # remove placeholder text in that paragraph (operate on full text)
        if placeholder in para.text:
            new_text = para.text.replace(placeholder, "")
            para.clear()
            para.add_run(new_text)
        insert_table_after_paragraph(template_doc, para, scenario_rows, col_names=["Spot at Maturity", "Payoff"])
    else:
        insert_table_after_paragraph(template_doc, template_doc.paragraphs[-1], scenario_rows, col_names=["Spot at Maturity", "Payoff"])
        st.info("Placeholder {{ScenarioTable}} not found — scenario table appended at document end.")

    # Save to bytes buffer and provide download
    output = io.BytesIO()
    template_doc.save(output)
    output.seek(0)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_client = client_name.replace(" ", "_")
    filename = f"TermSheet_{safe_client}_{timestamp}.docx"

    st.success("Document ready — click the button below to download.")
    st.download_button(
        label="Download Term Sheet (.docx)",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.info("Click **Generate Term Sheet** to create and download the document.")
