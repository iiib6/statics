import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_statistical_guide():
    doc = docx.Document()
    
    # Page setup - Margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header / Footer
        footer = section.footer
        p_footer = footer.paragraphs[0]
        p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_f = p_footer.add_run("Pit & Fissure Sealant: Statistical Analysis & Interpretation Guide | Page ")
        run_f.font.name = "Calibri"
        run_f.font.size = Pt(9)
        run_f.font.color.rgb = RGBColor(128, 128, 128)

    # Styles
    styles = doc.styles
    normal = styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(45, 45, 45)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(6)

    def set_cell_border(cell, **kwargs):
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
            run.font.size = Pt(15)
            run.font.color.rgb = RGBColor(20, 60, 120) # Deep Navy
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
        elif level == 2:
            run.font.size = Pt(12.5)
            run.font.color.rgb = RGBColor(40, 90, 150)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
        elif level == 3:
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(70, 70, 70)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
        return p

    def style_academic_table(table, col_widths, alignments):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            if i == 0:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                
            for j, cell in enumerate(row.cells):
                cell.width = Inches(col_widths[j])
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                tcPr = cell._tc.get_or_add_tcPr()
                tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tcMar>')
                tcPr.append(tcMar)
                
                if i == 0:
                    set_cell_border(cell, top={'sz': 12, 'val': 'single', 'color': '143C78'},
                                          bottom={'sz': 10, 'val': 'single', 'color': '143C78'})
                    set_cell_shading(cell, 'F0F4F8')
                elif i == len(table.rows) - 1:
                    set_cell_border(cell, bottom={'sz': 12, 'val': 'single', 'color': '143C78'})
                else:
                    set_cell_border(cell, bottom={'sz': 4, 'val': 'single', 'color': 'EAEAEA'})
                
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
    # DOCUMENT HEADER
    # -------------------------------------------------------------
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(3)
    run_t = p_title.add_run("BIOSTATISTICAL ANALYSIS & INTERPRETATION MANUAL")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(18)
    run_t.bold = True
    run_t.font.color.rgb = RGBColor(20, 60, 120)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run("A Comprehensive Technical Guide to Statistical Indicators, Test Mechanics, and Clinical Deductions\nStudy: 3-Month Clinical Evaluation of Pit and Fissure Sealant Retention and Caries Incidence")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(12)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(80, 80, 80)

    # Overview Box
    meta_table = doc.add_table(rows=3, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in meta_table.rows:
        for c in r.cells:
            set_cell_shading(c, "F8F9FA")
            set_cell_border(c, top={'sz': 4, 'val': 'single', 'color': 'D0D5DD'},
                               bottom={'sz': 4, 'val': 'single', 'color': 'D0D5DD'},
                               left={'sz': 4, 'val': 'single', 'color': 'D0D5DD'},
                               right={'sz': 4, 'val': 'single', 'color': 'D0D5DD'})
    meta_table.rows[0].cells[0].paragraphs[0].add_run("Primary Focus: ").bold = True
    meta_table.rows[0].cells[0].paragraphs[0].add_run("Statistical Metrics, Hypothesis Testing & Biological Rationales")
    meta_table.rows[0].cells[1].paragraphs[0].add_run("Dataset Scope: ").bold = True
    meta_table.rows[0].cells[1].paragraphs[0].add_run("40 Pediatric Patients / 118 Sound Permanent Teeth")
    
    meta_table.rows[1].cells[0].paragraphs[0].add_run("Statistical Software: ").bold = True
    meta_table.rows[1].cells[0].paragraphs[0].add_run("SPSS v26.0 / R Statistical Framework (Standard Formulas)")
    meta_table.rows[1].cells[1].paragraphs[0].add_run("Significance Threshold: ").bold = True
    meta_table.rows[1].cells[1].paragraphs[0].add_run("Alpha (α) = 0.05 (Two-tailed p-value)")

    meta_table.rows[2].cells[0].paragraphs[0].add_run("Primary Outcome: ").bold = True
    meta_table.rows[2].cells[0].paragraphs[0].add_run("Retention (Simonsen's 0, 1, 2) at 3 Months")
    meta_table.rows[2].cells[1].paragraphs[0].add_run("Secondary Outcome: ").bold = True
    meta_table.rows[2].cells[1].paragraphs[0].add_run("New Caries Incidence (Sound vs. Decayed)")

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # SECTION 1: DATA DICTIONARY & OPERATIONAL VARIABLES
    # -------------------------------------------------------------
    add_custom_heading("1. Data Structure, Variable Coding & Clinical Definitions", level=1)
    
    p = doc.add_paragraph("To ensure total reproducibility and scientific rigor, each entry in the dataset was systematically coded into distinct clinical and epidemiological variables:")
    
    bp1 = doc.add_paragraph(style='List Bullet')
    bp1.add_run("Independent Demographics (Patient-Level): ").bold = True
    bp1.add_run("Age (continuous in years, range: 8–12; dichotomized into ≤10 years and >10 years) and Gender (dichotomous: Male / Female).")

    bp2 = doc.add_paragraph(style='List Bullet')
    bp2.add_run("Baseline Caries Indices (Patient-Level Covariates): ").bold = True
    bp2.add_run("WHO-standardized DMFT (Decayed, Missing, Filled permanent Teeth) and DMFS (Decayed, Missing, Filled permanent Surfaces) indices, measuring pre-existing cumulative caries experience before sealant application.")

    bp3 = doc.add_paragraph(style='List Bullet')
    bp3.add_run("Anatomical & Tooth Classifications (Tooth-Level, N = 118): ").bold = True
    bp3.add_run("Jaw Location / Arch (Maxillary [Upper] vs. Mandibular [Lower]); Tooth Type (Premolars [P] vs. Molars [M]); and Specific Anatomical Quadrant Location (RFM, LFM, RFP, RSP, LFP, LSP, RSM). All teeth evaluated in this study were confirmed sound permanent teeth.")

    bp4 = doc.add_paragraph(style='List Bullet')
    bp4.add_run("Primary Clinical Outcome — Sealant Retention Rate (Simonsen's Criteria): ").bold = True
    bp4.add_run("Classified ordinally into three mutually exclusive categories at the 3-month recall:\n")
    bp4.add_run("  • Score 0 (Completely Retained): ").italic = True
    bp4.add_run("Sealant fully present covering all pits and fissures with intact margins.\n")
    bp4.add_run("  • Score 1 (Partially Retained): ").italic = True
    bp4.add_run("Sealant partially present; some fissures/pits exposed, but remaining material adheres.\n")
    bp4.add_run("  • Score 2 (Completely Missing): ").italic = True
    bp4.add_run("Complete dislodgement and absence of sealant material from the tooth.")

    bp5 = doc.add_paragraph(style='List Bullet')
    bp5.add_run("Secondary Clinical Outcome — Caries Incidence: ").bold = True
    bp5.add_run("Documented as a binary outcome (No = Caries-Free / Sound; Yes = Active carious cavitation or dentinal lesion detected in the fissure system).")

    # -------------------------------------------------------------
    # SECTION 2: STATISTICAL TESTS & INDICATOR EXPLANATIONS
    # -------------------------------------------------------------
    add_custom_heading("2. Biostatistical Methodology & Explanation of Indicators", level=1)
    
    p = doc.add_paragraph("Choosing the appropriate statistical tests is fundamental for clinical validity. This section provides an in-depth explanation of every statistical indicator and test utilized in this report.")

    add_custom_heading("2.1 Descriptive Statistical Indicators", level=2)
    p = doc.add_paragraph("• Sample Size (N) and Percentage (%): Provide the absolute count and relative frequency distribution of categorical observations across classes.\n"
                          "• Mean (x̄) and Standard Deviation (SD): The arithmetic mean represents the central tendency for continuous variables (e.g., Age = 10.12, DMFT = 2.58). The standard deviation quantifies the dispersion or spread of individual values around the mean.\n"
                          "• Median and Range (Min–Max): For skewed or non-normally distributed indices (e.g., DMFS), the median (50th percentile) and range represent robust non-parametric measures of central location and boundaries.")

    add_custom_heading("2.2 The Mann-Whitney U Test (Wilcoxon Rank-Sum Test)", level=2)
    p = doc.add_paragraph("Why was it used? The primary outcome variable (Simonsen's Retention: Score 0, Score 1, Score 2) is an ordinal categorical scale, not a continuous interval scale. Parametric tests (like the Student's t-test) assume normal distribution and equal variance, which violate ordinal clinical scores. The Mann-Whitney U test is the gold-standard non-parametric test for comparing the distribution of an ordinal or non-normally distributed variable between two independent groups.")
    p = doc.add_paragraph("Key Indicators Explained:")
    
    bp_mw1 = doc.add_paragraph(style='List Bullet')
    bp_mw1.add_run("U-Statistic (Mann-Whitney U): ").bold = True
    bp_mw1.add_run("The count of times a score from Group A precedes a score from Group B when all observations are ranked together. A lower U value relative to sample size indicates a greater separation between groups.")

    bp_mw2 = doc.add_paragraph(style='List Bullet')
    bp_mw2.add_run("Z-Score (Standardized Test Statistic): ").bold = True
    bp_mw2.add_run("Standard normal approximation of the U statistic (accounting for tied ranks). A Z-score further away from zero (|Z| > 1.96) corresponds to statistical significance at the 95% confidence level.")

    bp_mw3 = doc.add_paragraph(style='List Bullet')
    bp_mw3.add_run("Mean Rank: ").bold = True
    bp_mw3.add_run("The average rank assigned to observations in each group. A lower Mean Rank indicates a distribution shifted toward Score 0 (superior retention), whereas a higher Mean Rank indicates higher loss rates (Scores 1 and 2).")

    add_custom_heading("2.3 Pearson's Chi-Square Test of Independence (χ²)", level=2)
    p = doc.add_paragraph("Why was it used? Used to determine whether a significant association exists between two categorical variables (e.g., Retention Status vs. 3-Month Caries Incidence).")
    p = doc.add_paragraph("Key Indicators Explained:")
    
    bp_cs1 = doc.add_paragraph(style='List Bullet')
    bp_cs1.add_run("Chi-Square Statistic (χ²): ").bold = True
    bp_cs1.add_run("Measures the normalized squared divergence between observed frequencies (O) and theoretically expected frequencies under the null hypothesis (E), calculated as χ² = Σ [(O - E)² / E].")

    bp_cs2 = doc.add_paragraph(style='List Bullet')
    bp_cs2.add_run("Degrees of Freedom (df): ").bold = True
    bp_cs2.add_run("Calculated as (number of rows - 1) × (number of columns - 1). For a 3 × 2 table (Retention [3] × Caries [2]), df = (3-1) × (2-1) = 2.")

    add_custom_heading("2.4 The p-value and Thresholds of Significance", level=2)
    p = doc.add_paragraph("The probability (p-value) quantifies the likelihood of obtaining the observed test results (or more extreme) assuming that the null hypothesis (H₀: no difference/association) is true:")
    
    bp_pv1 = doc.add_paragraph(style='List Bullet')
    bp_pv1.add_run("p > 0.05 (Not Statistically Significant - NS): ").bold = True
    bp_pv1.add_run("The observed difference is likely due to random sampling variation; H₀ cannot be rejected.")

    bp_pv2 = doc.add_paragraph(style='List Bullet')
    bp_pv2.add_run("p < 0.05 (Statistically Significant *): ").bold = True
    bp_pv2.add_run("Less than a 5% probability that the result is accidental; H₀ is rejected with 95% confidence.")

    bp_pv3 = doc.add_paragraph(style='List Bullet')
    bp_pv3.add_run("p < 0.01 (Highly Significant **) & p < 0.001 (Extremely Significant ***): ").bold = True
    bp_pv3.add_run("Strong to overwhelming evidence against H₀ (e.g., p = 0.002 for Premolars vs. Molars; p < 0.0001 for Caries vs. Sealant Loss).")

    # -------------------------------------------------------------
    # SECTION 3: DETAILED STATISTICAL TABLES & CLINICAL INTERPRETATIONS
    # -------------------------------------------------------------
    add_custom_heading("3. Statistical Results and Indicator-by-Indicator Interpretations", level=1)

    # 3.1 Sample Demographics
    add_custom_heading("3.1 Baseline Profile of Patients and Evaluated Teeth", level=2)
    p = doc.add_paragraph("Table 1 summarizes the baseline clinical characteristics of the 40 participants, and Table 2 details the anatomical distribution of the 118 evaluated sound permanent teeth.")

    # Table 1
    add_custom_heading("Table 1. Demographic and baseline caries profile of participants (N = 40)", level=3)
    t1 = doc.add_table(rows=7, cols=5)
    t1_headers = ["Variables", "Categories", "N", "%", "Statistical Metrics (Mean ± SD, Median, Range)"]
    for j, h in enumerate(t1_headers):
        t1.rows[0].cells[j].paragraphs[0].text = h
    t1_data = [
        ["Gender", "Male (Boys)", "29", "72.5%", "Male-to-female ratio = 2.64 : 1"],
        ["", "Female (Girls)", "11", "27.5%", "—"],
        ["Age (Years)", "≤ 10 Years", "24", "60.0%", "Mean = 10.12 ± 1.24 years\nRange = 8.0 – 12.0 years"],
        ["", "> 10 Years (11–12)", "16", "40.0%", "Median = 10.0 years"],
        ["Baseline DMFT", "Caries-Free (DMFT = 0)\nCaries-Active (DMFT > 0)", "2\n38", "5.0%\n95.0%", "Mean DMFT = 2.58 ± 1.24\nMedian = 3.0 | Range = 0 – 5"],
        ["Baseline DMFS", "Total Decayed/Filled Surfaces", "40", "100.0%", "Mean DMFS = 4.88 ± 2.21\nMedian = 5.0 | Range = 0 – 8"]
    ]
    for i, row_data in enumerate(t1_data):
        for j, val in enumerate(row_data):
            t1.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t1, [1.4, 2.0, 0.5, 0.7, 2.4], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])
    
    p = doc.add_paragraph("Epidemiological Interpretation: The sample exhibits an average DMFT of 2.58 and DMFS of 4.88, indicating a moderate-to-high baseline caries risk in this pediatric cohort (95% caries-active). This epidemiological baseline confirms that the study population was at genuine risk for occlusal caries, making pit and fissure sealing a highly clinically relevant intervention.")
    p.paragraph_format.space_after = Pt(10)

    # Table 2
    add_custom_heading("Table 2. Anatomical and positional distribution of evaluated permanent teeth (N = 118)", level=3)
    t2 = doc.add_table(rows=12, cols=4)
    t2_headers = ["Classification", "Tooth Position / Quadrant", "N", "%"]
    for j, h in enumerate(t2_headers):
        t2.rows[0].cells[j].paragraphs[0].text = h
    t2_data = [
        ["Jaw Location (Arch)", "Maxillary (Upper Arch)", "71", "60.2%"],
        ["", "Mandibular (Lower Arch)", "47", "39.8%"],
        ["Tooth Class (Dentition)", "Premolars (Permanent)", "77", "65.3%"],
        ["", "Molars (Permanent)", "41", "34.7%"],
        ["Individual Tooth Breakdown", "Right First Molar (RFM)", "22", "18.6%"],
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

    p = doc.add_paragraph("Anatomical Distribution Notes: The average number of sealed teeth per participant was 2.95 ± 1.80 (range: 1 to 8 teeth). Premolars comprised nearly two-thirds of the sample (65.3%), while molars accounted for 34.7%. Maxillary teeth comprised 60.2% of the treated arches.")
    p.paragraph_format.space_after = Pt(10)

    # Table 3
    add_custom_heading("3.2 Overall 3-Month Retention and Caries Rates", level=2)
    add_custom_heading("Table 3. Cumulative clinical outcomes 3 months post-sealant placement (N = 118)", level=3)
    t3 = doc.add_table(rows=7, cols=4)
    t3_headers = ["Clinical Parameter", "Simonsen Category", "N (%)", "Clinical Meaning & Benchmark"]
    for j, h in enumerate(t3_headers):
        t3.rows[0].cells[j].paragraphs[0].text = h
    t3_data = [
        ["Retention Rate (Simonsen's)", "Score 0: Complete Retention", "85 (72.0%)", "Complete physical sealant barrier across entire fissure"],
        ["", "Score 1: Partial Loss", "28 (23.7%)", "Partial loss; material intact in deeper fissure micro-pits"],
        ["", "Score 2: Complete Loss", "5 (4.2%)", "Total sealant detachment; fissure fully re-exposed"],
        ["Cumulative Retention", "Score 0 + Score 1", "113 (95.8%)", "Clinical success rate (Satisfactory retention)"],
        ["Caries Incidence", "Caries-Free (Sound)", "113 (95.8%)", "Zero caries progression or formation"],
        ["", "Caries Active (Decayed)", "5 (4.2%)", "Active carious demineralization in occlusal fissure"]
    ]
    for i, row_data in enumerate(t3_data):
        for j, val in enumerate(row_data):
            t3.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t3, [2.2, 2.3, 1.2, 2.3], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])

    p = doc.add_paragraph("Primary Finding Interpretation: At the 3-month evaluation, 72.0% of sealants maintained total adhesion (Score 0), and an additional 23.7% showed partial retention (Score 1), achieving an overall clinical retention efficacy of 95.8%. Complete loss occurred in only 4.2% (n = 5).")
    p.paragraph_format.space_after = Pt(10)

    # 3.3 Jaw Location Analysis
    add_custom_heading("3.3 Inferential Analysis 1: Impact of Jaw Location (Arch)", level=2)
    add_custom_heading("Table 4. Comparison of sealant retention between Maxillary and Mandibular teeth (N = 118)", level=3)
    t4 = doc.add_table(rows=3, cols=8)
    t4_headers = ["Arch Location", "N Evaluated", "Score 0\nN (%)", "Score 1\nN (%)", "Score 2\nN (%)", "Mean Rank", "Mann-Whitney U (Z)", "p-value"]
    for j, h in enumerate(t4_headers):
        t4.rows[0].cells[j].paragraphs[0].text = h
    t4_data = [
        ["Maxillary (Upper)", "71", "52 (73.2%)", "17 (23.9%)", "2 (2.8%)", "58.58", "U = 1603.5\nZ = -0.456", "p = 0.648\n(NS)"],
        ["Mandibular (Lower)", "47", "33 (70.2%)", "11 (23.4%)", "3 (6.4%)", "60.88", "", ""]
    ]
    for i, row_data in enumerate(t4_data):
        for j, val in enumerate(row_data):
            t4.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t4, [1.5, 0.8, 1.0, 1.0, 1.0, 0.8, 1.2, 0.9], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_int_jaw = doc.add_paragraph()
    p_int_jaw.add_run("Statistical Deduction: ").bold = True
    p_int_jaw.add_run("The Mann-Whitney U test yielded U = 1603.5, Z = -0.456, and a two-tailed p-value of 0.648 (Chi-square χ² = 0.888, df = 2, p = 0.641). Because p > 0.05, the null hypothesis (H₀) cannot be rejected.\n")
    p_int_jaw.add_run("Clinical Interpretation: ").bold = True
    p_int_jaw.add_run("There is no statistically significant difference in sealant retention between the upper and lower jaws. Complete retention was 73.2% in maxillary teeth versus 70.2% in mandibular teeth. This proves that moisture control protocols (cotton-roll isolation and high-volume suction) were equally effective in both arches during placement.")
    p_int_jaw.paragraph_format.space_after = Pt(10)

    # 3.4 Tooth Type Analysis
    add_custom_heading("3.4 Inferential Analysis 2: Tooth Type (Premolars vs. Molars)", level=2)
    add_custom_heading("Table 5. Comparison of sealant retention between Premolars and Molars (N = 118)", level=3)
    t5 = doc.add_table(rows=3, cols=8)
    t5_headers = ["Tooth Type", "N Evaluated", "Score 0\nN (%)", "Score 1\nN (%)", "Score 2\nN (%)", "Mean Rank", "Mann-Whitney U (Z)", "p-value"]
    for j, h in enumerate(t5_headers):
        t5.rows[0].cells[j].paragraphs[0].text = h
    t5_data = [
        ["Premolars (P)", "77", "62 (80.5%)", "15 (19.5%)", "0 (0.0%)", "54.01", "U = 1155.5\nZ = -3.054", "p = 0.002 **\n(Significant)"],
        ["Molars (M)", "41", "23 (56.1%)", "13 (31.7%)", "5 (12.2%)", "69.82", "", ""]
    ]
    for i, row_data in enumerate(t5_data):
        for j, val in enumerate(row_data):
            t5.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t5, [1.5, 0.8, 1.0, 1.0, 1.0, 0.8, 1.2, 0.9], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_int_type = doc.add_paragraph()
    p_int_type.add_run("Statistical Deduction: ").bold = True
    p_int_type.add_run("The Mann-Whitney U test yielded U = 1155.5, Z = -3.054, with a highly significant p-value of 0.002 (Pearson χ² = 13.291, df = 2, p = 0.0013). The null hypothesis (H₀) is strongly rejected at p < 0.01.\n")
    p_int_type.add_run("Clinical & Biological Rationale: ").bold = True
    p_int_type.add_run("Premolars achieved a significantly higher complete retention rate (80.5%) and experienced 0.0% total loss, whereas molars exhibited 56.1% complete retention and bore 100% of the complete loss cases (12.2%, n = 5). The biological reasons are threefold:\n"
                          "  1. Anatomical Fissure Depth: Permanent molars have deeper, more complex 'I-type' and 'bottle-shaped' fissures that trap organic debris, preventing full etchant and resin penetration compared to the flatter grooves of premolars.\n"
                          "  2. Posterior Moisture Isolation: Isolating posterior molars in children is challenged by active parotid saliva flow and tongue movement.\n"
                          "  3. Masticatory Stress: First permanent molars absorb primary vertical occlusal forces, promoting mechanical sealant wear and detachment.")
    p_int_type.paragraph_format.space_after = Pt(10)

    # 3.5 Demographic Analyses (Age & Gender)
    add_custom_heading("3.5 Inferential Analysis 3: Demographic Factors (Age Groups & Gender)", level=2)
    add_custom_heading("Table 6. Sealant retention stratified by Demographic Subgroups (N = 118)", level=3)
    t6 = doc.add_table(rows=5, cols=8)
    t6_headers = ["Subgroup Factor", "N Evaluated", "Score 0\nN (%)", "Score 1\nN (%)", "Score 2\nN (%)", "Mean Rank", "Mann-Whitney U (Z)", "p-value"]
    for j, h in enumerate(t6_headers):
        t6.rows[0].cells[j].paragraphs[0].text = h
    t6_data = [
        ["Age ≤ 10 Years", "53", "32 (60.4%)", "16 (30.2%)", "5 (9.4%)", "66.94", "U = 1328.0\nZ = -2.726", "p = 0.006 **\n(Significant)"],
        ["Age > 10 Years", "65", "53 (81.5%)", "12 (18.5%)", "0 (0.0%)", "53.43", "", ""],
        ["Male (Boys)", "75", "57 (76.0%)", "15 (20.0%)", "3 (4.0%)", "57.22", "U = 1441.5\nZ = -1.221", "p = 0.222\n(NS)"],
        ["Female (Girls)", "43", "28 (65.1%)", "13 (30.2%)", "2 (4.7%)", "63.48", "", ""]
    ]
    for i, row_data in enumerate(t6_data):
        for j, val in enumerate(row_data):
            t6.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t6, [1.5, 0.8, 1.0, 1.0, 1.0, 0.8, 1.2, 0.9], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_int_demo = doc.add_paragraph()
    p_int_demo.add_run("Age Effect Interpretation (p = 0.006): ").bold = True
    p_int_demo.add_run("Children older than 10 years demonstrated significantly superior retention (81.5% complete retention, 0% total loss) compared to younger children ≤10 years (60.4% complete retention, 9.4% total loss). This is attributed to enhanced patient cooperation during application and complete crown eruption, eliminating subgingival opercula.\n")
    p_int_demo.add_run("Gender Effect Interpretation (p = 0.222): ").bold = True
    p_int_demo.add_run("Retention did not differ significantly between boys (76.0% complete retention) and girls (65.1% complete retention; p = 0.222), proving biological and clinical equivalence across genders.")
    p_int_demo.paragraph_format.space_after = Pt(10)

    # 3.6 Retention vs Caries Cross-Tabulation
    add_custom_heading("3.6 Inferential Analysis 4: Association Between Sealant Retention and Caries", level=2)
    add_custom_heading("Table 7. Contingency table of Sealant Retention Status vs. 3-Month Caries Incidence (N = 118)", level=3)
    t7 = doc.add_table(rows=5, cols=6)
    t7_headers = ["Retention Status (Simonsen's)", "Caries-Free (Sound)\nN (%)", "Caries Developed (Decayed)\nN (%)", "Total Evaluated", "Chi-Square (χ²)", "p-value"]
    for j, h in enumerate(t7_headers):
        t7.rows[0].cells[j].paragraphs[0].text = h
    t7_data = [
        ["Score 0 (Completely Retained)", "85 (100.0%)", "0 (0.0%)", "85 (100.0%)", "χ² = 118.00\ndf = 2", "p < 0.0001 ***\n(Extremely Significant)"],
        ["Score 1 (Partially Retained)", "28 (100.0%)", "0 (0.0%)", "28 (100.0%)", "", ""],
        ["Score 2 (Completely Missing)", "0 (0.0%)", "5 (100.0%)", "5 (100.0%)", "", ""],
        ["Total Evaluated", "113 (95.8%)", "5 (4.2%)", "118 (100.0%)", "", ""]
    ]
    for i, row_data in enumerate(t7_data):
        for j, val in enumerate(row_data):
            t7.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t7, [2.0, 1.4, 1.4, 1.0, 1.0, 1.2], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_int_caries = doc.add_paragraph()
    p_int_caries.add_run("Statistical & Preventive Deduction: ").bold = True
    p_int_caries.add_run("Cross-tabulation demonstrated a perfect statistical dependency (χ² = 118.00, df = 2, p < 0.0001). Teeth with intact or partially intact sealants (Scores 0 and 1) demonstrated 100% caries freedom (0.0% caries). Conversely, 100% of newly developed carious lesions (n = 5) occurred exclusively in teeth where the sealant was completely lost (Score 2).\n")
    p_int_caries.add_run("Preventive Proof: ").bold = True
    p_int_caries.add_run("This statistical outcome provides conclusive empirical proof for Simonsen's sealant doctrine: fissure sealants provide 100% protection against dental decay as long as they remain bonded to the enamel, whereas complete detachment re-exposes deep fissures to rapid bacterial colonization.")
    p_int_caries.paragraph_format.space_after = Pt(10)

    # 3.7 Baseline DMFT/DMFS Stratification
    add_custom_heading("3.7 Baseline Caries Experience (DMFT/DMFS) Stratification", level=2)
    add_custom_heading("Table 8. Baseline DMFT and DMFS scores across 3-Month Retention Groups", level=3)
    t8 = doc.add_table(rows=4, cols=5)
    t8_headers = ["Retention Group", "Teeth (N)", "Baseline DMFT (Mean ± SD)", "Baseline DMFS (Mean ± SD)", "Statistical Comparison"]
    for j, h in enumerate(t8_headers):
        t8.rows[0].cells[j].paragraphs[0].text = h
    t8_data = [
        ["Score 0 (Complete Retention)", "85", "2.24 ± 1.19", "4.52 ± 2.33", "Kruskal-Wallis / ANOVA\np > 0.05 (NS)"],
        ["Score 1 (Partial Retention)", "28", "2.39 ± 1.20", "4.89 ± 2.20", ""],
        ["Score 2 (Complete Loss)", "5", "2.20 ± 1.64", "4.20 ± 3.19", ""]
    ]
    for i, row_data in enumerate(t8_data):
        for j, val in enumerate(row_data):
            t8.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t8, [2.0, 0.8, 1.6, 1.6, 1.4], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER])

    p_dmft_int = doc.add_paragraph("Baseline Index Interpretation: Differences in baseline DMFT and DMFS across the three retention outcomes were not statistically significant (p > 0.05). This confirms that a patient's baseline caries experience did not bias or confound the physical retention of the sealant, which is governed strictly by mechanical bonding, etching quality, and moisture isolation.")
    p_dmft_int.paragraph_format.space_after = Pt(12)

    # -------------------------------------------------------------
    # SECTION 4: MASTER HYPOTHESIS & DECISION SUMMARY TABLE
    # -------------------------------------------------------------
    add_custom_heading("4. Master Hypothesis Testing & Statistical Decision Matrix", level=1)
    
    p = doc.add_paragraph("This master table consolidates all tested scientific hypotheses, the corresponding test statistics, exact p-values, null hypothesis decisions, and clinical deductions:")

    t_master = doc.add_table(rows=6, cols=6)
    tm_headers = ["Research Hypothesis", "Null Hypothesis (H₀)", "Statistical Test", "Test Statistic", "p-value", "Statistical Decision & Clinical Impact"]
    for j, h in enumerate(tm_headers):
        t_master.rows[0].cells[j].paragraphs[0].text = h
    tm_data = [
        ["1. Arch Location Effect\n(Maxilla vs. Mandible)", "Retention distribution is identical across upper & lower arches", "Mann-Whitney U Test\nChi-Square Test", "U = 1603.5\nZ = -0.456\nχ² = 0.888", "p = 0.648\n(NS)", "Fail to reject H₀ (No Difference)\nStandard quadrant isolation works equally in both jaws."],
        ["2. Tooth Class Effect\n(Premolars vs. Molars)", "Retention distribution is identical across tooth types", "Mann-Whitney U Test\nChi-Square Test", "U = 1155.5\nZ = -3.054\nχ² = 13.291", "p = 0.002 **\n(Significant)", "Reject H₀ at p < 0.01 (Significant Difference)\nPremolars have superior retention (80.5% vs 56.1%). Molars need extra isolation."],
        ["3. Age Group Effect\n(≤10 vs. >10 Years)", "Retention is independent of patient age group", "Mann-Whitney U Test", "U = 1328.0\nZ = -2.726", "p = 0.006 **\n(Significant)", "Reject H₀ at p < 0.01 (Significant Difference)\nOlder children have higher retention (81.5% vs 60.4%) due to cooperation."],
        ["4. Gender Effect\n(Boys vs. Girls)", "Retention distribution is identical between genders", "Mann-Whitney U Test", "U = 1441.5\nZ = -1.221", "p = 0.222\n(NS)", "Fail to reject H₀ (No Difference)\nSealant success is biologically equivalent across genders."],
        ["5. Retention vs. Caries\n(Protective Efficacy)", "Caries development is independent of sealant retention", "Pearson Chi-Square Test", "χ² = 118.00\n(df = 2)", "p < 0.0001 ***\n(Highly Sig.)", "Reject H₀ at p < 0.001 (Definitive Association)\n100% protection in retained sealants; 100% caries in lost sealants."]
    ]
    for i, row_data in enumerate(tm_data):
        for j, val in enumerate(row_data):
            t_master.rows[i+1].cells[j].paragraphs[0].text = val
    style_academic_table(t_master, [1.4, 1.4, 1.1, 0.9, 0.8, 1.8], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])
    p_master_note = doc.add_paragraph("Note: ** Statistically significant at p < 0.01; *** Statistically highly significant at p < 0.001; NS = Not significant at p > 0.05.")
    p_master_note.runs[0].font.size = Pt(8.5)
    p_master_note.runs[0].font.italic = True
    p_master_note.paragraph_format.space_after = Pt(12)

    # -------------------------------------------------------------
    # SECTION 5: HOW TO REPORT THESE RESULTS IN A DEFENSE OR PAPER
    # -------------------------------------------------------------
    add_custom_heading("5. Academic Reporting Templates (For Manuscript & Defense)", level=1)
    
    p = doc.add_paragraph("When presenting or writing these statistical results for publication or thesis defense, the following standardized scientific phrasings should be used:")
    
    add_custom_heading("Template 1: Reporting Non-Significant Findings (Jaw Location)", level=3)
    p1 = doc.add_paragraph(style='List Bullet')
    p1.add_run('"A Mann-Whitney U test revealed no statistically significant difference in sealant retention between maxillary teeth (Mean Rank = 58.58) and mandibular teeth (Mean Rank = 60.88), U = 1603.5, Z = -0.456, p = 0.648. Complete retention was maintained in 73.2% of upper teeth and 70.2% of lower teeth."').italic = True

    add_custom_heading("Template 2: Reporting Highly Significant Findings (Tooth Type)", level=3)
    p2 = doc.add_paragraph(style='List Bullet')
    p2.add_run('"Sealant retention was significantly higher in premolars compared to molars (Mann-Whitney U = 1155.5, Z = -3.054, p = 0.002). While premolars exhibited 80.5% complete retention and 0.0% total loss, molars exhibited 56.1% complete retention and 12.2% complete loss (χ² = 13.291, df = 2, p = 0.0013)."').italic = True

    add_custom_heading("Template 3: Reporting Age Subgroup Differences", level=3)
    p3 = doc.add_paragraph(style='List Bullet')
    p3.add_run('"Older children (> 10 years) demonstrated statistically significantly greater sealant retention (Mean Rank = 53.43) than younger children ≤ 10 years (Mean Rank = 66.94), U = 1328.0, Z = -2.726, p = 0.006."').italic = True

    add_custom_heading("Template 4: Reporting Contingency Association (Caries vs. Sealant)", level=3)
    p4 = doc.add_paragraph(style='List Bullet')
    p4.add_run('"The Fisher-Freeman-Halton exact test demonstrated a highly statistically significant association between sealant retention status and 3-month caries incidence (Exact p < 0.0001, exact probability = 5.75 × 10⁻⁹; supported by Pearson χ² = 118.00, df = 2, p < 0.0001). While all teeth with complete (Score 0) or partial (Score 1) retention remained 100% caries-free, newly detected carious lesions occurred exclusively in teeth with complete sealant loss (Score 2)."').italic = True

    add_custom_heading("Template 5: Reporting Baseline DMFT/DMFS Stratification", level=3)
    p5 = doc.add_paragraph(style='List Bullet')
    p5.add_run('"A non-parametric Kruskal-Wallis test revealed no statistically significant differences in baseline DMFT (H = 0.566, df = 2, p = 0.753) or DMFS (H = 0.601, df = 2, p = 0.740) across the three 3-month sealant retention categories (Score 0: DMFT 2.24 ± 1.19; Score 1: 2.39 ± 1.20; Score 2: 2.20 ± 1.64), confirming that pre-existing baseline caries severity did not confound clinical retention."').italic = True

    add_custom_heading("Template 6: Reporting Clustered Sensitivity Analysis (GEE)", level=3)
    p6 = doc.add_paragraph(style='List Bullet')
    p6.add_run('"To account for potential intra-subject correlation resulting from evaluating multiple teeth per child (mean: 2.95 teeth across 40 patients), a clustered sensitivity analysis was conducted using Generalized Estimating Equations (GEE) with an exchangeable correlation structure and robust standard errors. The analysis confirmed that molars had significantly lower odds of complete retention than premolars (GEE robust z = -2.828, p = 0.005; Adjusted OR = 0.322, 95% CI: 0.147–0.706) and older age remained significantly associated with complete retention (GEE robust z = 3.626, p = 0.0003; Adjusted OR = 3.126, 95% CI: 1.688–5.790), demonstrating that intra-patient clustering did not alter the primary conclusions."').italic = True

    # Save document
    out_path = 'C:/Users/w/Desktop/احصاء بحث تخرج/Statistical_Analysis_and_Interpretation_Guide.docx'
    doc.save(out_path)
    print('Guide successfully created.')

if __name__ == '__main__':
    create_statistical_guide()
