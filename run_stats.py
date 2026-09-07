import zipfile
import xml.etree.ElementTree as ET
import math
from collections import Counter, defaultdict

# Load data from Excel
with zipfile.ZipFile('C:/Users/w/Desktop/احصاء بحث تخرج/NOOR_120.xlsx', 'r') as z:
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for elem in tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            t = elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            if t is not None and t.text is not None:
                shared_strings.append(t.text)
            else:
                txt = ''.join([t_sub.text for t_sub in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t_sub.text])
                shared_strings.append(txt)

    sheet_tree = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    rows_data = []
    for row in sheet_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
        row_num = int(row.attrib.get('r', 0))
        row_dict = {}
        for c in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            r_coord = c.attrib.get('r', '')
            col_letter = ''.join([ch for ch in r_coord if ch.isalpha()])
            t_type = c.attrib.get('t', '')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else None
            if val is not None and t_type == 's':
                val = shared_strings[int(val)]
            row_dict[col_letter] = val
        rows_data.append((row_num, row_dict))

patients = []
teeth = []
current_p = None

for r_num, r_d in rows_data[1:]:
    age_raw = r_d.get('A')
    gender_raw = r_d.get('B')
    dmfs_raw = r_d.get('E')
    dmft_raw = r_d.get('F')
    
    if age_raw is not None or gender_raw is not None:
        current_p = {
            'id': len(patients) + 1,
            'age': int(float(age_raw)),
            'gender': gender_raw.strip().capitalize(),
            'dmfs': int(float(dmfs_raw)),
            'dmft': int(float(dmft_raw)),
            'teeth': []
        }
        patients.append(current_p)
    
    tooth_raw = r_d.get('C')
    arch_raw = r_d.get('D')
    ret_raw = r_d.get('G')
    caries_raw = r_d.get('H')
    
    t_str = tooth_raw.strip().upper() if tooth_raw else ''
    is_molar = 'M' in t_str
    is_premolar = 'P' in t_str
    tooth_type = 'Molar' if is_molar else ('Premolar' if is_premolar else 'Other')
    
    t_entry = {
        'row': r_num,
        'patient_id': current_p['id'],
        'age': current_p['age'],
        'gender': current_p['gender'],
        'dmfs': current_p['dmfs'],
        'dmft': current_p['dmft'],
        'tooth': t_str,
        'tooth_type': tooth_type,
        'arch': arch_raw.strip().capitalize() if arch_raw else '',
        'retention': int(float(ret_raw)),
        'caries': caries_raw.strip().upper() if caries_raw else ''
    }
    current_p['teeth'].append(t_entry)
    teeth.append(t_entry)

def mean_sd(arr):
    n = len(arr)
    if n == 0: return 0, 0
    m = sum(arr) / n
    var = sum((x - m)**2 for x in arr) / (n - 1) if n > 1 else 0
    return m, math.sqrt(var)

def mann_whitney_u(group1, group2):
    n1 = len(group1)
    n2 = len(group2)
    if n1 == 0 or n2 == 0: return None
    
    combined = [(x, 1) for x in group1] + [(x, 2) for x in group2]
    combined.sort(key=lambda item: item[0])
    
    ranks = [0] * len(combined)
    i = 0
    while i < len(combined):
        j = i
        while j < len(combined) and combined[j][0] == combined[i][0]:
            j += 1
        avg_rank = (i + 1 + j) / 2.0
        for k in range(i, j):
            ranks[k] = avg_rank
        i = j
        
    r1 = sum(ranks[idx] for idx, item in enumerate(combined) if item[1] == 1)
    r2 = sum(ranks[idx] for idx, item in enumerate(combined) if item[1] == 2)
    
    u1 = r1 - (n1 * (n1 + 1)) / 2.0
    u2 = r2 - (n2 * (n2 + 1)) / 2.0
    u = min(u1, u2)
    
    mean_u = (n1 * n2) / 2.0
    tie_counts = Counter([x for x, _ in combined])
    t_sum = sum(cnt**3 - cnt for cnt in tie_counts.values())
    n = n1 + n2
    sigma_u = math.sqrt((n1 * n2 / 12.0) * ((n + 1) - t_sum / (n * (n - 1))))
    
    z = (u - mean_u) / sigma_u
    def norm_cdf(z_val):
        return 0.5 * (1.0 + math.erf(z_val / math.sqrt(2.0)))
    p_val = 2.0 * norm_cdf(-abs(z))
    return {'U': u, 'W': r1, 'Z': z, 'p_value': p_val, 'n1': n1, 'n2': n2, 'mean_rank1': r1/n1, 'mean_rank2': r2/n2}

def chi_square_test(table):
    nrows = len(table)
    ncols = len(table[0])
    row_sums = [sum(row) for row in table]
    col_sums = [sum(table[r][c] for r in range(nrows)) for c in range(ncols)]
    total = sum(row_sums)
    
    chi2 = 0.0
    expected = []
    for r in range(nrows):
        exp_row = []
        for c in range(ncols):
            exp_val = (row_sums[r] * col_sums[c]) / total
            exp_row.append(exp_val)
            if exp_val > 0:
                chi2 += ((table[r][c] - exp_val)**2) / exp_val
        expected.append(exp_row)
        
    df = (nrows - 1) * (ncols - 1)
    if df == 1:
        p_val = 2.0 * (1.0 - 0.5 * (1.0 + math.erf(math.sqrt(chi2) / math.sqrt(2.0))))
    elif df == 2:
        p_val = math.exp(-chi2 / 2.0)
    else:
        z_approx = ((chi2 / df)**(1/3) - (1 - 2/(9*df))) / math.sqrt(2/(9*df))
        p_val = 1.0 - 0.5 * (1.0 + math.erf(z_approx / math.sqrt(2.0)))
        p_val = max(0.0, min(1.0, p_val))
        
    return {'chi2': chi2, 'df': df, 'p_value': p_val, 'expected': expected}

print('=== 1. PATIENT DEMOGRAPHICS (N=40) ===')
p_ages = [p['age'] for p in patients]
m_age, sd_age = mean_sd(p_ages)
print(f'Age: Mean = {m_age:.2f} ± {sd_age:.2f}, Range = {min(p_ages)}-{max(p_ages)}')
age_g1 = len([p for p in patients if p['age'] <= 10])
age_g2 = len([p for p in patients if p['age'] > 10])
print(f'Age <=10: {age_g1} ({age_g1/40*100:.1f}%) | Age >10: {age_g2} ({age_g2/40*100:.1f}%)')
p_boys = len([p for p in patients if p['gender'] == 'Male'])
p_girls = len([p for p in patients if p['gender'] == 'Female'])
print(f'Boys: {p_boys} ({p_boys/40*100:.1f}%) | Girls: {p_girls} ({p_girls/40*100:.1f}%)')

p_dmft = [p['dmft'] for p in patients]
m_dmft, sd_dmft = mean_sd(p_dmft)
print(f'DMFT: Mean = {m_dmft:.2f} ± {sd_dmft:.2f}, Median = {sorted(p_dmft)[20]}, Range = {min(p_dmft)}-{max(p_dmft)}')

p_dmfs = [p['dmfs'] for p in patients]
m_dmfs, sd_dmfs = mean_sd(p_dmfs)
print(f'DMFS: Mean = {m_dmfs:.2f} ± {sd_dmfs:.2f}, Median = {sorted(p_dmfs)[20]}, Range = {min(p_dmfs)}-{max(p_dmfs)}')

print('\n=== 2. TEETH DEMOGRAPHICS & TOOTH DISTRIBUTION (N=118) ===')
t_boys = len([t for t in teeth if t['gender'] == 'Male'])
t_girls = len([t for t in teeth if t['gender'] == 'Female'])
print(f'Gender: Boys = {t_boys} ({t_boys/118*100:.1f}%), Girls = {t_girls} ({t_girls/118*100:.1f}%)')
t_age1 = len([t for t in teeth if t['age'] <= 10])
t_age2 = len([t for t in teeth if t['age'] > 10])
print(f'Age <=10: {t_age1} ({t_age1/118*100:.1f}%), Age >10: {t_age2} ({t_age2/118*100:.1f}%)')
t_up = len([t for t in teeth if t['arch'] == 'Upper'])
t_low = len([t for t in teeth if t['arch'] == 'Lower'])
print(f'Arch: Upper = {t_up} ({t_up/118*100:.1f}%), Lower = {t_low} ({t_low/118*100:.1f}%)')
t_mol = len([t for t in teeth if t['tooth_type'] == 'Molar'])
t_pre = len([t for t in teeth if t['tooth_type'] == 'Premolar'])
print(f'Tooth Type: Molars = {t_mol} ({t_mol/118*100:.1f}%), Premolars = {t_pre} ({t_pre/118*100:.1f}%)')

print('Detailed tooth breakdown:')
t_counts = Counter([t['tooth'] for t in teeth])
for k, v in sorted(t_counts.items(), key=lambda x: -x[1]):
    print(f'  {k}: {v} ({v/118*100:.1f}%)')

print('\n=== 3. OVERALL SEALANT RETENTION & CARIES (N=118) ===')
r0 = len([t for t in teeth if t['retention'] == 0])
r1 = len([t for t in teeth if t['retention'] == 1])
r2 = len([t for t in teeth if t['retention'] == 2])
print(f'Completely Retained (Score 0): {r0} ({r0/118*100:.1f}%)')
print(f'Partially Retained (Score 1): {r1} ({r1/118*100:.1f}%)')
print(f'Completely Missing (Score 2): {r2} ({r2/118*100:.1f}%)')
c_no = len([t for t in teeth if t['caries'] == 'NO'])
c_yes = len([t for t in teeth if t['caries'] == 'YES'])
print(f'Caries-Free (Sound): {c_no} ({c_no/118*100:.1f}%) | Caries (Decayed): {c_yes} ({c_yes/118*100:.1f}%)')

print('\n=== 4. RETENTION BY JAW LOCATION (Upper vs Lower) ===')
u_ret = [t['retention'] for t in teeth if t['arch'] == 'Upper']
l_ret = [t['retention'] for t in teeth if t['arch'] == 'Lower']
u_c = Counter(u_ret)
l_c = Counter(l_ret)
print(f'Upper (N={len(u_ret)}): Complete={u_c[0]} ({u_c[0]/len(u_ret)*100:.1f}%), Partial={u_c[1]} ({u_c[1]/len(u_ret)*100:.1f}%), Missing={u_c[2]} ({u_c[2]/len(u_ret)*100:.1f}%)')
print(f'Lower (N={len(l_ret)}): Complete={l_c[0]} ({l_c[0]/len(l_ret)*100:.1f}%), Partial={l_c[1]} ({l_c[1]/len(l_ret)*100:.1f}%), Missing={l_c[2]} ({l_c[2]/len(l_ret)*100:.1f}%)')
mw_jaw = mann_whitney_u(u_ret, l_ret)
print(f'Mann-Whitney U: U={mw_jaw["U"]:.1f}, Z={mw_jaw["Z"]:.3f}, p={mw_jaw["p_value"]:.4f}, MeanRank_Upper={mw_jaw["mean_rank1"]:.2f}, MeanRank_Lower={mw_jaw["mean_rank2"]:.2f}')
chi_jaw = chi_square_test([[u_c[0], u_c[1], u_c[2]], [l_c[0], l_c[1], l_c[2]]])
print(f'Chi-Square: Chi2={chi_jaw["chi2"]:.3f}, df={chi_jaw["df"]}, p={chi_jaw["p_value"]:.4f}')

print('\n=== 5. RETENTION BY TOOTH TYPE (Premolars vs Molars) ===')
p_ret = [t['retention'] for t in teeth if t['tooth_type'] == 'Premolar']
m_ret = [t['retention'] for t in teeth if t['tooth_type'] == 'Molar']
p_c = Counter(p_ret)
m_c = Counter(m_ret)
print(f'Premolars (N={len(p_ret)}): Complete={p_c[0]} ({p_c[0]/len(p_ret)*100:.1f}%), Partial={p_c[1]} ({p_c[1]/len(p_ret)*100:.1f}%), Missing={p_c[2]} ({p_c[2]/len(p_ret)*100:.1f}%)')
print(f'Molars (N={len(m_ret)}): Complete={m_c[0]} ({m_c[0]/len(m_ret)*100:.1f}%), Partial={m_c[1]} ({m_c[1]/len(m_ret)*100:.1f}%), Missing={m_c[2]} ({m_c[2]/len(m_ret)*100:.1f}%)')
mw_type = mann_whitney_u(p_ret, m_ret)
print(f'Mann-Whitney U: U={mw_type["U"]:.1f}, Z={mw_type["Z"]:.3f}, p={mw_type["p_value"]:.4f}, MeanRank_Premolar={mw_type["mean_rank1"]:.2f}, MeanRank_Molar={mw_type["mean_rank2"]:.2f}')
chi_type = chi_square_test([[p_c[0], p_c[1], p_c[2]], [m_c[0], m_c[1], m_c[2]]])
print(f'Chi-Square: Chi2={chi_type["chi2"]:.3f}, df={chi_type["df"]}, p={chi_type["p_value"]:.4f}')

print('\n=== 6. RETENTION BY GENDER (Male vs Female) ===')
b_ret = [t['retention'] for t in teeth if t['gender'] == 'Male']
g_ret = [t['retention'] for t in teeth if t['gender'] == 'Female']
b_c = Counter(b_ret)
g_c = Counter(g_ret)
print(f'Boys (N={len(b_ret)}): Complete={b_c[0]} ({b_c[0]/len(b_ret)*100:.1f}%), Partial={b_c[1]} ({b_c[1]/len(b_ret)*100:.1f}%), Missing={b_c[2]} ({b_c[2]/len(b_ret)*100:.1f}%)')
print(f'Girls (N={len(g_ret)}): Complete={g_c[0]} ({g_c[0]/len(g_ret)*100:.1f}%), Partial={g_c[1]} ({g_c[1]/len(g_ret)*100:.1f}%), Missing={g_c[2]} ({g_c[2]/len(g_ret)*100:.1f}%)')
mw_gender = mann_whitney_u(b_ret, g_ret)
print(f'Mann-Whitney U: U={mw_gender["U"]:.1f}, Z={mw_gender["Z"]:.3f}, p={mw_gender["p_value"]:.4f}, MeanRank_Boys={mw_gender["mean_rank1"]:.2f}, MeanRank_Girls={mw_gender["mean_rank2"]:.2f}')

print('\n=== 7. RETENTION BY AGE GROUP (<=10 vs >10) ===')
a1_ret = [t['retention'] for t in teeth if t['age'] <= 10]
a2_ret = [t['retention'] for t in teeth if t['age'] > 10]
a1_c = Counter(a1_ret)
a2_c = Counter(a2_ret)
print(f'Age <=10 (N={len(a1_ret)}): Complete={a1_c[0]} ({a1_c[0]/len(a1_ret)*100:.1f}%), Partial={a1_c[1]} ({a1_c[1]/len(a1_ret)*100:.1f}%), Missing={a1_c[2]} ({a1_c[2]/len(a1_ret)*100:.1f}%)')
print(f'Age >10 (N={len(a2_ret)}): Complete={a2_c[0]} ({a2_c[0]/len(a2_ret)*100:.1f}%), Partial={a2_c[1]} ({a2_c[1]/len(a2_ret)*100:.1f}%), Missing={a2_c[2]} ({a2_c[2]/len(a2_ret)*100:.1f}%)')
mw_age = mann_whitney_u(a1_ret, a2_ret)
print(f'Mann-Whitney U: U={mw_age["U"]:.1f}, Z={mw_age["Z"]:.3f}, p={mw_age["p_value"]:.4f}, MeanRank_Age1={mw_age["mean_rank1"]:.2f}, MeanRank_Age2={mw_age["mean_rank2"]:.2f}')

print('\n=== 8. CARIES INCIDENCE vs RETENTION STATUS ===')
ret_caries = defaultdict(lambda: Counter())
for t in teeth:
    ret_caries[t['retention']][t['caries']] += 1

for r in [0, 1, 2]:
    print(f'Retention Score {r}: Caries NO = {ret_caries[r]["NO"]} ({ret_caries[r]["NO"]/(ret_caries[r]["NO"]+ret_caries[r]["YES"])*100:.1f}%), Caries YES = {ret_caries[r]["YES"]} ({ret_caries[r]["YES"]/(ret_caries[r]["NO"]+ret_caries[r]["YES"])*100:.1f}%)')

chi_rc = chi_square_test([
    [ret_caries[0]['NO'], ret_caries[0]['YES']],
    [ret_caries[1]['NO'], ret_caries[1]['YES']],
    [ret_caries[2]['NO'], ret_caries[2]['YES']]
])
print(f'Chi-Square (Caries vs Retention): Chi2={chi_rc["chi2"]:.3f}, df={chi_rc["df"]}, p={chi_rc["p_value"]:.6f}')

print('\n=== 9. DMFT & DMFS vs RETENTION STATUS ===')
for r in [0, 1, 2]:
    r_dmft = [t['dmft'] for t in teeth if t['retention'] == r]
    r_dmfs = [t['dmfs'] for t in teeth if t['retention'] == r]
    m_f, sd_f = mean_sd(r_dmft)
    m_s, sd_s = mean_sd(r_dmfs)
    print(f'Retention Score {r} (N={len(r_dmft)}): DMFT = {m_f:.2f} ± {sd_f:.2f}, DMFS = {m_s:.2f} ± {sd_s:.2f}')
