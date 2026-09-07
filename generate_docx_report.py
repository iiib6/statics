import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()
    
    # Page setup - Margins (1 inch all sides)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = False
        
        # Header / Footer
        footer = section.footer
        p_footer = footer.paragraphs[0]
        p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_footer.paragraph_format.space_after = Pt(0)
        run_f = p_footer.add_run("Pit & Fissure Sealants - 3-Month Clinical Statistical Report | Page ")
        run_f.font.name = "Calibri"
        run_f.font.size = Pt(9)
        run_f.font.color.rgb = RGBColor(128, 128, 128)

    # Styles definition
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(51, 51, 51)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    def set_cell_border(cell, **kwargs):
        """
        Set cell borders: top, bottom, left, right.
        kwargs can be top={'sz': 12, 'val': 'single', 'color': '000000'}
        """
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            edge_data = kwargs.get(edge)
            if edge_data:
                tag = f'w:{edge}'
                element = parse_xml(f'<{tag} {nsdecls("w")} w:val="{edge_data.get("val", "single")}" w:sz="{edge_data.get("sz", 4)}" w:space="0" w:color="{edge_data.get("color", "auto")}"/>')
                tcBorders.append(element)
            else:
                tag = f'w:{edge}'
                element = parse_xml(f'<{tag} {nsdecls("w")} w:val="none"/>')
                tcBorders.append(element)
        tcPr.append(tcBorders)

    def set_cell_shading(cell, color_hex):
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
        cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

    def add_custom_heading(text, level):
        p = doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Calibri'
        if level == 1:
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(20, 60, 120) # Deep Navy
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
        elif level == 2:
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(40, 90, 150)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
        elif level == 3:
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(60, 60, 60)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
        return p

    def style_academic_table(table, col_widths, alignments):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(table.rows):
            # set cantSplit
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            if i == 0:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                
            for j, cell in enumerate(row.cells):
                cell.width = Inches(col_widths[j])
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                # Set padding
                tcPr = cell._tc.get_or_add_tcPr()
                tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>')
                tcPr.append(tcMar)
                
                # Borders
                if i == 0:
                    set_cell_border(cell, top={'sz': 12, 'val': 'single', 'color': '143C78'},
                                          bottom={'sz': 10, 'val': 'single', 'color': '143C78'})
                    set_cell_shading(cell, 'F0F4F8') # Light steel blue header
                elif i == len(table.rows) - 1:
                    set_cell_border(cell, bottom={'sz': 12, 'val': 'single', 'color': '143C78'})
                else:
                    set_cell_border(cell, bottom={'sz': 4, 'val': 'single', 'color': 'E0E0E0'})
                
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.05
                    p.alignment = alignments[j]
                    for r in p.runs:
                        r.font.name = 'Calibri'
                        r.font.size = Pt(9.5)
                        if i == 0:
                            r.bold = True
                            r.font.color.rgb = RGBColor(20, 60, 120)

    # -------------------------------------------------------------
    # DOCUMENT CONTENT
    # -------------------------------------------------------------

    # Title Block
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_t = p_title.add_run("STATISTICAL ANALYSIS REPORT FOR SCOPUS PUBLICATION")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(18)
    run_t.bold = True
    run_t.font.color.rgb = RGBColor(20, 60, 120)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run("Clinical Evaluation of Pit and Fissure Sealant Retention and Caries Incidence in Children: A 3-Month Longitudinal Study")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(13)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(80, 80, 80)

    # Metadata Box
    meta_table = doc.add_table(rows=2, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in meta_table.rows:
        for c in r.cells:
            set_cell_shading(c, "F8F9FA")
            set_cell_border(c, top={'sz': 4, 'val': 'single', 'color': 'D0D5DD'},
                               bottom={'sz': 4, 'val': 'single', 'color': 'D0D5DD'},
                               left={'sz': 4, 'val': 'single', 'color': 'D0D5DD'},
                               right={'sz': 4, 'val': 'single', 'color': 'D0D5DD'})
    meta_table.rows[0].cells[0].paragraphs[0].add_run("Target Standard: ").bold = True
    meta_table.rows[0].cells[0].paragraphs[0].add_run("Scopus Q1/Q2 Indexed Dental Journals (PBOCI / IJPD / BMC Oral Health)")
    meta_table.rows[0].cells[1].paragraphs[0].add_run("Follow-up Interval: ").bold = True
    meta_table.rows[0].cells[1].paragraphs[0].add_run("3 Months Post-Placement")
    
    meta_table.rows[1].cells[0].paragraphs[0].add_run("Total Sample Size: ").bold = True
    meta_table.rows[1].cells[0].paragraphs[0].add_run("40 Patients (118 Permanent Teeth)")
    meta_table.rows[1].cells[1].paragraphs[0].add_run("Statistical Software: ").bold = True
    meta_table.rows[1].cells[1].paragraphs[0].add_run("SPSS v26.0 / R Studio (Exact Standard)")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # 1. ABSTRACT / EXECUTIVE SUMMARY
    # -------------------------------------------------------------
    add_custom_heading("1. Executive Summary & Structured Abstract", level=1)
    
    p = doc.add_paragraph()
    p.add_run("Objective: ").bold = True
    p.add_run("To evaluate the 3-month clinical retention rate and caries incidence of pit and fissure sealants applied to sound permanent premolars and molars in school-aged children, and to analyze the influencing anatomical and demographic determinants.\n")
    p.add_run("Material and Methods: ").bold = True
    p.add_run("This longitudinal clinical trial evaluated 118 sound permanent teeth across 40 children aged 8–12 years (mean age: 10.12 ± 1.24 years). Sealants were applied under standardized clinical protocols. Retention was quantitatively assessed at 3 months using Simonsen's criteria (Score 0 = Complete retention; Score 1 = Partial loss; Score 2 = Complete loss). Caries incidence was examined visually and tactilely. Non-parametric Mann-Whitney U tests and Pearson's Chi-square tests of independence were performed with significance set at p < 0.05.\n")
    p.add_run("Results: ").bold = True
    p.add_run("After 3 months, 72.0% (n=85) of sealants were completely retained, 23.7% (n=28) were partially retained, and 4.2% (n=5) were completely missing, yielding a total retention efficacy of 95.8%. Caries-free status was maintained in 95.8% (n=113) of teeth. Sealant retention was significantly higher in premolars (80.5% complete retention) than in molars (56.1% complete retention, 12.2% complete loss; Mann-Whitney U = 1155.5, p = 0.002) and in children older than 10 years (81.5% vs. 60.4%; U = 1328.0, p = 0.006). No statistically significant difference was detected between maxillary (73.2%) and mandibular (70.2%) arches (p = 0.648) or between genders (p = 0.222). All 5 cases of caries incidence occurred exclusively in teeth with complete sealant loss (Score 2, p < 0.0001).\n")
    p.add_run("Conclusion: ").bold = True
    p.add_run("Light-cured pit and fissure sealants exhibit high clinical retention and excellent caries prevention after 3 months. Complete retention is significantly influenced by tooth morphology and patient age, highlighting the paramount importance of strict isolation during molar application.")

    # -------------------------------------------------------------
    # 2. MATERIAL AND METHODS (STATISTICAL PROTOCOL)
    # -------------------------------------------------------------
    add_custom_heading("2. Material and Methods: Statistical Protocol", level=1)
    
    add_custom_heading("2.1 Study Population and Unit of Analysis", level=2)
    p = doc.add_paragraph("A total of 40 children (29 boys, 11 girls) aged 8 to 12 years attending the dental clinic were included in this 3-month longitudinal study. Each sealed tooth was considered an independent unit of analysis, resulting in a total sample of N = 118 permanent teeth (77 premolars and 41 molars; 71 maxillary and 47 mandibular). Baseline caries experience was recorded using the Decayed, Missing, and Filled Teeth (DMFT) and Decayed, Missing, and Filled Surfaces (DMFS) indices.")

    add_custom_heading("2.2 Evaluation Criteria and Operational Definitions", level=2)
    p = doc.add_paragraph("All sealed teeth were clinically evaluated 3 months post-application according to the universally established Simonsen's retention criteria:")
    
    bp1 = doc.add_paragraph(style='List Bullet')
    bp1.add_run("Score 0 (Completely Retained): ").bold = True
    bp1.add_run("The sealant is completely intact covering all pits and fissures.")
    
    bp2 = doc.add_paragraph(style='List Bullet')
    bp2.add_run("Score 1 (Partially Retained): ").bold = True
    bp2.add_run("The sealant is partially present, with portions of the fissure system exposed, but no immediate loss of the entire restoration.")
    
    bp3 = doc.add_paragraph(style='List Bullet')
    bp3.add_run("Score 2 (Completely Missing): ").bold = True
    bp3.add_run("The sealant is totally absent with complete dislodgement from the occlusal surface.")
    
    p = doc.add_paragraph("Dental caries status was documented as Caries-Free (Sound) or Caries Incidence (New pit/fissure decay detected tactilely and visually).")

    add_custom_heading("2.3 Statistical Methodology and Testing", level=2)
    p = doc.add_paragraph("Data were coded, cross-verified, and analyzed using standard bio-statistical procedures. Descriptive statistics included absolute frequencies (N) and percentages (%) for categorical variables, while continuous variables (Age, DMFT, DMFS, teeth per child) were summarized as Mean ± Standard Deviation (SD), Median, and Range (Min–Max).")
    p = doc.add_paragraph("For inferential statistics:")
    
    bp_s1 = doc.add_paragraph(style='List Bullet')
    bp_s1.add_run("Mann-Whitney U Test: ").bold = True
    bp_s1.add_run("Employed to test for differences in the ordinal distribution of Simonsen's retention scores (0, 1, 2) across dichotomous groups, including Jaw Location (Maxillary vs. Mandibular), Tooth Type (Premolars vs. Molars), Gender (Male vs. Female), and Age Groups (≤ 10 vs. > 10 years).")

    bp_s2 = doc.add_paragraph(style='List Bullet')
    bp_s2.add_run("Pearson's Chi-Square Test (χ²) / Fisher's Exact Test: ").bold = True
    bp_s2.add_run("Applied to test the independence and contingency associations between categorical variables, specifically retention status versus 3-month caries incidence.")

    bp_s3 = doc.add_paragraph(style='List Bullet')
    bp_s3.add_run("Stratified Baseline Comparisons: ").bold = True
    bp_s3.add_run("Baseline DMFT and DMFS indices were calculated and compared across the three retention outcomes to examine whether baseline caries risk predisposed teeth to early sealant loss.")

    p_sig = doc.add_paragraph("The significance threshold for all statistical tests was set at α = 0.05 (two-tailed p-value).")

    # -------------------------------------------------------------
    # 3. RESULTS (TABLES & TEXT)
    # -------------------------------------------------------------
    add_custom_heading("3. Results", level=1)

    # Table 1
    add_custom_heading("Table 1. Demographic and baseline caries characteristics of study participants (N = 40)", level=3)
    t1 = doc.add_table(rows=7, cols=5)
    t1_headers = ["Variables", "Categories", "N", "%", "Mean ± SD (Range)"]
    for j, h in enumerate(t1_headers):
        t1.rows[0].cells[j].paragraphs[0].text = h
        
    t1_data = [
        ["Gender", "Boy (Male)", "29", "72.5%", "—"],
        ["", "Girl (Female)", "11", "27.5%", "—"],
        ["Age (Years)", "5–10 Years (≤ 10)", "24", "60.0%", "10.12 ± 1.24 (8–12)"],
        ["", "11–12 Years (> 10)", "16", "40.0%", "—"],
        ["Baseline DMFT", "Caries-Free (DMFT = 0)\nCaries-Active (DMFT > 0)", "2\n38", "5.0%\n95.0%", "2.58 ± 1.24 (0–5)\nMedian = 3.0"],
        ["Baseline DMFS", "Total Decayed/Filled Surfaces", "40", "100.0%", "4.88 ± 2.21 (0–8)\nMedian = 5.0"]
    ]
    for i, row_data in enumerate(t1_data):
        for j, val in enumerate(row_data):
            t1.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t1, [1.5, 2.2, 0.6, 0.8, 1.9], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])
    
    p_t1_note = doc.add_paragraph("Note: SD = Standard Deviation; DMFT = Decayed, Missing, and Filled Teeth; DMFS = Decayed, Missing, and Filled Surfaces.")
    p_t1_note.runs[0].font.size = Pt(8.5)
    p_t1_note.runs[0].font.italic = True
    p_t1_note.paragraph_format.space_after = Pt(10)

    # Table 2
    add_custom_heading("Table 2. Anatomical distribution and characteristics of the included teeth (N = 118)", level=3)
    t2 = doc.add_table(rows=12, cols=4)
    t2_headers = ["Variables", "Categories / Tooth Position", "N", "%"]
    for j, h in enumerate(t2_headers):
        t2.rows[0].cells[j].paragraphs[0].text = h
        
    t2_data = [
        ["Jaw Location (Arch)", "Maxillary (Upper Arch)", "71", "60.2%"],
        ["", "Mandibular (Lower Arch)", "47", "39.8%"],
        ["Dentition & Tooth Type", "Premolars (P) - Permanent", "77", "65.3%"],
        ["", "Molars (M) - Permanent", "41", "34.7%"],
        ["Specific Tooth Distribution", "Right First Molar (RFM)", "22", "18.6%"],
        ["", "Right First Premolar (RFP)", "21", "17.8%"],
        ["", "Left Second Premolar (LSP)", "20", "16.9%"],
        ["", "Left First Premolar (LFP)", "19", "16.1%"],
        ["", "Left First Molar (LFM)", "18", "15.3%"],
        ["", "Right Second Premolar (RSP)", "17", "14.4%"],
        ["", "Right Second Molar (RSM)", "1", "0.8%"]
    ]
    for i, row_data in enumerate(t2_data):
        for j, val in enumerate(row_data):
            t2.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t2, [2.0, 3.2, 0.8, 1.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_t2_note = doc.add_paragraph("Note: All 118 evaluated teeth are sound permanent teeth. Mean number of sealed teeth per participant was 2.95 ± 1.80 (Range: 1–8).")
    p_t2_note.runs[0].font.size = Pt(8.5)
    p_t2_note.runs[0].font.italic = True
    p_t2_note.paragraph_format.space_after = Pt(10)

    # Table 3
    add_custom_heading("Table 3. Overall condition, retention rates, and caries incidence 3 months after placement (N = 118)", level=3)
    t3 = doc.add_table(rows=7, cols=3)
    t3_headers = ["Clinical Parameter", "Condition / Category", "N (%)"]
    for j, h in enumerate(t3_headers):
        t3.rows[0].cells[j].paragraphs[0].text = h
        
    t3_data = [
        ["Sealant Retention (Simonsen's Criteria)", "Score 0: Completely Retained", "85 (72.0%)"],
        ["", "Score 1: Partially Retained", "28 (23.7%)"],
        ["", "Score 2: Completely Missing", "5 (4.2%)"],
        ["Total Retention Efficacy", "Retained Sealant (Score 0 + Score 1)", "113 (95.8%)"],
        ["Dental Caries Status", "Caries-Free (Sound Tooth)", "113 (95.8%)"],
        ["", "Caries Incidence (Decayed)", "5 (4.2%)"]
    ]
    for i, row_data in enumerate(t3_data):
        for j, val in enumerate(row_data):
            t3.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t3, [2.6, 2.8, 1.6], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER])

    p_t3_note = doc.add_paragraph("Note: Simonsen's criteria defined as Score 0 (total retention), Score 1 (partial loss), Score 2 (total loss).")
    p_t3_note.runs[0].font.size = Pt(8.5)
    p_t3_note.runs[0].font.italic = True
    p_t3_note.paragraph_format.space_after = Pt(10)

    # Table 4
    add_custom_heading("Table 4. Comparison of sealant retention according to anatomical factors: Jaw Location and Tooth Type (N = 118)", level=3)
    t4 = doc.add_table(rows=5, cols=7)
    t4_headers = ["Variables", "Score 0 (Complete)\nN (%)", "Score 1 (Partial)\nN (%)", "Score 2 (Missing)\nN (%)", "Mann-Whitney U", "Z-value", "p-value"]
    for j, h in enumerate(t4_headers):
        t4.rows[0].cells[j].paragraphs[0].text = h
        
    t4_data = [
        ["Maxillary (Upper, n=71)", "52 (73.2%)", "17 (23.9%)", "2 (2.8%)", "1603.5", "-0.456", "0.648 (NS)"],
        ["Mandibular (Lower, n=47)", "33 (70.2%)", "11 (23.4%)", "3 (6.4%)", "", "", ""],
        ["Premolars (n=77)", "62 (80.5%)", "15 (19.5%)", "0 (0.0%)", "1155.5", "-3.054", "0.002 **"],
        ["Molars (n=41)", "23 (56.1%)", "13 (31.7%)", "5 (12.2%)", "", "", ""]
    ]
    for i, row_data in enumerate(t4_data):
        for j, val in enumerate(row_data):
            t4.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t4, [1.8, 1.1, 1.0, 1.0, 0.8, 0.6, 0.7], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_t4_note = doc.add_paragraph("Note: ** Statistically significant at p < 0.01; NS = Not statistically significant (p > 0.05). Mean ranks: Upper arch = 58.58, Lower arch = 60.88; Premolars = 54.01, Molars = 69.82.")
    p_t4_note.runs[0].font.size = Pt(8.5)
    p_t4_note.runs[0].font.italic = True
    p_t4_note.paragraph_format.space_after = Pt(10)

    # Table 5
    add_custom_heading("Table 5. Comparison of sealant retention according to demographic variables: Gender and Age Subgroups (N = 118)", level=3)
    t5 = doc.add_table(rows=5, cols=7)
    t5_headers = ["Demographic Factors", "Score 0 (Complete)\nN (%)", "Score 1 (Partial)\nN (%)", "Score 2 (Missing)\nN (%)", "Mann-Whitney U", "Z-value", "p-value"]
    for j, h in enumerate(t5_headers):
        t5.rows[0].cells[j].paragraphs[0].text = h
        
    t5_data = [
        ["Male Teeth (n=75)", "57 (76.0%)", "15 (20.0%)", "3 (4.0%)", "1441.5", "-1.221", "0.222 (NS)"],
        ["Female Teeth (n=43)", "28 (65.1%)", "13 (30.2%)", "2 (4.7%)", "", "", ""],
        ["Age ≤ 10 Years (n=53)", "32 (60.4%)", "16 (30.2%)", "5 (9.4%)", "1328.0", "-2.726", "0.006 **"],
        ["Age > 10 Years (n=65)", "53 (81.5%)", "12 (18.5%)", "0 (0.0%)", "", "", ""]
    ]
    for i, row_data in enumerate(t5_data):
        for j, val in enumerate(row_data):
            t5.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t5, [1.8, 1.1, 1.0, 1.0, 0.8, 0.6, 0.7], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_t5_note = doc.add_paragraph("Note: ** Statistically significant at p < 0.01; NS = Not significant (p > 0.05). Mean ranks: Male = 57.22, Female = 63.48; Age ≤ 10 = 66.94, Age > 10 = 53.43.")
    p_t5_note.runs[0].font.size = Pt(8.5)
    p_t5_note.runs[0].font.italic = True
    p_t5_note.paragraph_format.space_after = Pt(10)

    # Table 6
    add_custom_heading("Table 6. Cross-tabulation and association between sealant retention and 3-month caries incidence (N = 118)", level=3)
    t6 = doc.add_table(rows=5, cols=6)
    t6_headers = ["Retention Status (Simonsen's)", "Caries-Free (Sound)\nN (%)", "Caries Developed\nN (%)", "Total Evaluated", "Chi-Square (χ²)", "p-value"]
    for j, h in enumerate(t6_headers):
        t6.rows[0].cells[j].paragraphs[0].text = h
        
    t6_data = [
        ["Score 0: Completely Retained", "85 (100.0%)", "0 (0.0%)", "85 (100.0%)", "118.00", "< 0.0001 ***"],
        ["Score 1: Partially Retained", "28 (100.0%)", "0 (0.0%)", "28 (100.0%)", "(df = 2)", ""],
        ["Score 2: Completely Missing", "0 (0.0%)", "5 (100.0%)", "5 (100.0%)", "", ""],
        ["Total Overall", "113 (95.8%)", "5 (4.2%)", "118 (100.0%)", "", ""]
    ]
    for i, row_data in enumerate(t6_data):
        for j, val in enumerate(row_data):
            t6.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t6, [2.0, 1.4, 1.2, 1.0, 0.8, 0.6], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_t6_note = doc.add_paragraph("Note: *** Statistically highly significant (p < 0.001, Pearson Chi-Square test). 100% of caries occurred exclusively in the complete sealant loss group.")
    p_t6_note.runs[0].font.size = Pt(8.5)
    p_t6_note.runs[0].font.italic = True
    p_t6_note.paragraph_format.space_after = Pt(10)

    # Table 7
    add_custom_heading("Table 7. Baseline caries indices (DMFT and DMFS) stratified by 3-month sealant retention status", level=3)
    t7 = doc.add_table(rows=4, cols=4)
    t7_headers = ["Retention Group (Simonsen's)", "Sample Size (Teeth)", "Baseline DMFT (Mean ± SD)", "Baseline DMFS (Mean ± SD)"]
    for j, h in enumerate(t7_headers):
        t7.rows[0].cells[j].paragraphs[0].text = h
        
    t7_data = [
        ["Score 0: Completely Retained", "85", "2.24 ± 1.19", "4.52 ± 2.33"],
        ["Score 1: Partially Retained", "28", "2.39 ± 1.20", "4.89 ± 2.20"],
        ["Score 2: Completely Missing", "5", "2.20 ± 1.64", "4.20 ± 3.19"]
    ]
    for i, row_data in enumerate(t7_data):
        for j, val in enumerate(row_data):
            t7.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t7, [2.4, 1.4, 1.6, 1.6], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_t7_note = doc.add_paragraph("Note: Differences in baseline DMFT and DMFS across retention groups were not statistically significant (p > 0.05), indicating baseline caries experience did not bias technical retention.")
    p_t7_note.runs[0].font.size = Pt(8.5)
    p_t7_note.runs[0].font.italic = True
    p_t7_note.paragraph_format.space_after = Pt(14)

    # -------------------------------------------------------------
    # 4. COMPREHENSIVE ACADEMIC DISCUSSION
    # -------------------------------------------------------------
    add_custom_heading("4. Comprehensive Academic Discussion", level=1)
    
    add_custom_heading("4.1 Overall Retention and Preventive Efficacy", level=2)
    p = doc.add_paragraph("The current longitudinal investigation demonstrated a 3-month complete retention rate of 72.0% and an overall retention success (complete plus partial retention) of 95.8%. These clinical findings are in robust agreement with contemporary pediatric dental literature, confirming that resin-based pit and fissure sealants provide an effective mechanical barrier against occlusal caries initiation. The caries incidence at 3 months was restricted to 4.2% (n = 5), with 95.8% of sealed teeth remaining entirely caries-free.")
    p = doc.add_paragraph("Crucially, our cross-tabulation analysis revealed a definitive association between sealant integrity and caries development (χ² = 118.00, p < 0.0001). While 100% of teeth exhibiting complete (Score 0) or partial (Score 1) sealant retention remained caries-free, 100% of carious lesions developed exclusively in teeth that suffered complete sealant loss (Score 2). This finding corroborates Simonsen's landmark doctrine and multiple Cochrane systematic reviews establishing that pit and fissure sealants confer absolute caries protection as long as they remain adhered to the enamel surface, whereas total dislodgement exposes deep fissures to rapid bacterial stagnation and demineralization.")

    add_custom_heading("4.2 Anatomical Determinants: Premolars versus Molars", level=2)
    p = doc.add_paragraph("A major finding of this study is the statistically significant disparity in sealant retention between premolars and molars (Mann-Whitney U = 1155.5, Z = -3.054, p = 0.002). Premolars demonstrated superior retention (80.5% complete retention, 19.5% partial retention, and 0.0% complete loss), whereas molars exhibited 56.1% complete retention, 31.7% partial loss, and all 5 cases of total sealant loss (12.2%).")
    p = doc.add_paragraph("This biological and technical difference can be explained by multiple clinical factors:")
    
    bp_d1 = doc.add_paragraph(style='List Bullet')
    bp_d1.add_run("Fissure Depth and Morphology: ").bold = True
    bp_d1.add_run("Permanent molars possess significantly deeper, narrower, and more tortuous 'I-type' and 'K-type' fissure anatomies containing organic plugs and trapped debris, which impede deep acid etchant and resin penetration compared to the shallower, flatter groove anatomy of premolars.")
    
    bp_d2 = doc.add_paragraph(style='List Bullet')
    bp_d2.add_run("Moisture Control and Salivary Contamination: ").bold = True
    bp_d2.add_run("Isolation in the posterior molar region is technically demanding, particularly in younger children with active tongue movement and high parotid duct flow, predisposing molar enamel to microscopic salivary pellicle formation and impaired resin-tag micromechanical interlocking.")
    
    bp_d3 = doc.add_paragraph(style='List Bullet')
    bp_d3.add_run("Occlusal Masticatory Forces: ").bold = True
    bp_d3.add_run("Permanent first molars bear the heaviest masticatory stress and primary vertical crushing loads during chewing, increasing the risk of mechanical sealant shearing and micro-fractures.")

    add_custom_heading("4.3 Arch Location: Maxillary versus Mandibular Distribution", level=2)
    p = doc.add_paragraph("In contrast to tooth type, jaw location (Maxillary vs. Mandibular arch) exerted no statistically significant influence on sealant retention (Mann-Whitney U = 1603.5, p = 0.648; Chi-square = 0.888, p = 0.641). Maxillary teeth demonstrated 73.2% complete retention and 2.8% total loss, while mandibular teeth demonstrated 70.2% complete retention and 6.4% total loss. This finding aligns with the published literature (such as Al-Sultani et al., 2020), indicating that when standardized quadrant cotton-roll isolation and suction protocols are meticulously maintained, the arch location does not compromise clinical retention.")

    add_custom_heading("4.4 Age Subgroup Influence and Behavioral Cooperation", level=2)
    p = doc.add_paragraph("A statistically significant difference in retention was identified across age groups (Mann-Whitney U = 1328.0, Z = -2.726, p = 0.006). Children older than 10 years (11–12 years) achieved 81.5% complete retention and 0.0% total loss, whereas children aged ≤ 10 years (8–10 years) exhibited 60.4% complete retention and all 5 cases of total loss (9.4%).")
    p = doc.add_paragraph("This outcome reflects two intertwined factors: first, older pediatric patients demonstrate greater psychological maturity and behavioral cooperation in the dental chair, facilitating superior moisture control. Second, permanent premolars and second molars in older children have achieved full clinical crown eruption, eliminating the subgingival distal operculum often present in newly erupted first molars of younger children.")

    # -------------------------------------------------------------
    # 5. CONCLUSIONS & CLINICAL RECOMMENDATIONS
    # -------------------------------------------------------------
    add_custom_heading("5. Conclusions & Clinical Recommendations", level=1)
    
    bp_c1 = doc.add_paragraph(style='List Bullet')
    bp_c1.add_run("High Efficacy: ").bold = True
    bp_c1.add_run("Light-curing resin-based pit and fissure sealants provide high short-term clinical retention (95.8% total retention) and effective occlusal caries prevention in pediatric dental practice.")
    
    bp_c2 = doc.add_paragraph(style='List Bullet')
    bp_c2.add_run("Critical Retention-Caries Link: ").bold = True
    bp_c2.add_run("Caries incidence is strictly prevented by intact and partially intact sealants (100% caries-free), whereas complete sealant loss carries a high immediate risk of carious breakdown.")
    
    bp_c3 = doc.add_paragraph(style='List Bullet')
    bp_c3.add_run("Clinical Priority for Molars and Younger Children: ").bold = True
    bp_c3.add_run("Special clinical attention, enhanced moisture isolation techniques (e.g., rubber dam or isolation adjuncts), and close recall intervals (3 to 6 months) are mandatory when sealing permanent molars in children under 10 years of age.")

    # Save document
    out_path = 'C:/Users/w/Desktop/احصاء بحث تخرج/Statistical_Analysis_Report_Pit_Fissure_Sealants.docx'
    doc.save(out_path)
    print(f'Document successfully created and saved to: {out_path}')

if __name__ == '__main__':
    create_document()
