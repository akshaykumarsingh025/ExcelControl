import re

html = """<table class="table table-bordered"><thead><tr><th></th><th>Name</th><th>Father
name</th><th>Designation</th><th>Date</th><th>PAGE NO.</th><th>TERM</th><th>PAGE
NO.</th></tr></thead><tbody><tr><td>1</td><td>Fathesih</td><td>Banathore</td><td>G.N.L.</td><td>Fathesih</td><td>50NO.</th></tr></thead><tbody><tr><t>1</td><td>Fathesih</td><td>Banathore</td><td>G.N.L.</td><td>Fathesih</td><td>50053581</td><td>908</td><td>IDIB0000501</td><td>INDIAN BANK</td></tr><tr><td>2</td><td>Ram
briksh</td><td>Banathore</td><td>G.N.L.</td><td>Fathesih</td><td>50369536</td><td>880</td><td>IDIB00004641</td><td>INDIAN BANK
BANK</td></tr><tr><td>3</td><td>Umeek</td><td>Fathesih</td><td>G.N.L.</td><td>Fathesih</td><td>5961745</td><td>884<BANK</td></tr><tr><td>3</td><td>Umeek</td><td>Fathesih</td><td>G.N.L.</td><td>Fathesih</td><td>5961745</td><t>884</td><td>IDIB0000501</td><td>INDIAN BANK</td></tr><tr><td>4</td><td>Jay Rizad</td><td>Ram
swayad</td><td>G.N.L.</td><td>Jay Rizad</td><td>05961007</td><td>116</td><td>IDIB00004629</td><td>INDIAN
BANK</td></tr><tr><td>5</td><td>Fooluon</td><td>Hubbal</td><td>G.N.L.</td><td>Ram
Vichar</td><td>69910500</td><td>00472</td><td>IDIB00004630</td><td>INDIAN
BANK</td></tr><tr><td>6</td><td>Ramchazan</td><td>Shidhonath</td><td>G.N.L.</td><td>Sambishi</td><td>71963487</td><BANK</td></tr><tr><td>6</td><td>Ramchazan/td><td>Shidhonath</td><td>G.N.L.</td><td>Sambishi</td><td>71963487</td><td>004277</td><td>IDIB00004501</td><td>INDIAN
BANK</td></tr><tr><td>7</td><td>Rachhe</td><td>Mahadev</td><td>G.N.L.</td><td>Rachhonam</td><td>65300941</td><td>00BANK</td></tr><tr><td>7</td><td>Rachhe</td><td>Mahadev</td><td>G.N.L.<td><td>Rachhonam</td><td>65300941</td><td>004277</td><td>IDIB00004501</td><td>INDIAN BANK</td></tr><tr><td>8</td><td>Roy Jaram</td><td>Ram
Pyare</td><td>G.N.L.</td><td>Rachhonam</td><td>1008200100</td><td>004277</td><td>IDIB00004501</td><td>INDIAN
BANK</td></tr><tr><td>9</td><td>Jaydish</td><td>Ram
Pyare</td><td>G.N.L.</td><td></td><td></td><td>004277</td><td>IDIB00004501</td><td>INDIAN
BANK</td></tr><tr><td>10</td><td>Ajay Kuman</td><td>Lal singh</td><td>G.N.L.</td><td>AJAY
Kuman</td><td>28480100</td><td>009610</td><td>BARBO @ BRABS</td><td>Bank of
Barrada</td></tr><tr><td>11</td><td>chet
Singh</td><td>Motilal</td><td>G.N.L.</td><td>Kisanatya</td><td>5003961</td><td>7063</td><td>IDIB0000501</td><td>INDIAN BANK</td></tr><tr><td>12</td><td>Ram Vilash</td><td>Dhan Singh</td><td>G.N.L.</td><td>Ram
biksh</td><td>6306982</td><td>7063</td><td>IDIB0000501</td><td>INDIAN BANK</td></tr><tr><td>13</td><td>Ram
Jigunh</td><td>Dev
Shah</td><td>G.N.L.</td><td>Kushishya</td><td>6991821</td><td>000918</td><td>BKID0000629</td><td>INDIAN
BANK</td></tr><tr><td>14</td><td>Ramreshh</td><td>Fathe Singh</td><td>G.N.L.</td><td>Sushil
Dhawji</td><td>5074591</td><td>649</td><td>IDIB0000629</td><td>INDIAN
BANK</td></tr><tr><td>15</td><td>Ramesthan</td><td>Singh
Lal</td><td>G.N.L.</td><td>Athawariya</td><td>.50273041</td><td>656</td><td>IDIB0000629</td><td>INDIAN
BANK</td></tr><tr><td>16</td><td>Mustad Singh</td><td>Ram Prasad</td><td>G.N.L.</td><td>Ram
Prasad</td><td>.50352361</td><td>699</td><td>IDIB0000629</td><td>INDIAN BANK</td></tr><tr><td>17</td><td>Jagt
Singh</td><td>Ratanon</td><td>G.N.L.</td><td>Lilahati</td><td>8991010</td><td>000924</td><td>BKID0006929</td><td>Bank of INDIA</td></tr><tr><td>18</td><td>Shu Kuman</td><td>Ram
Sawande</td><td>G.N.L.</td><td>Lilahati</td><td>69290100</td><td>0022</td><td>BKID0006929</td><td>INDIAN
BANK</td></tr><tr><td>19</td><td>Kubler</td><td>Virisah</td><td>G.N.L.</td><td>Athawariya</td><td>5900774</td><td>8BANK</td></tr><tr><td>19</td><td>Kubler</td><td>Virisah</td><td>G.N.L.</td><td>Athawariya</td><td>5900774</td><td>847</td><td>IDIB0000629</td><td>INDIAN BANK</td></tr><tr><td>20</td><td>@m
Prakash</td><td>Lakchman</td><td>G.N.L.</td><td>Sambali</td><td>.59108116</td><td>841</td><td>IDIB0000629</td><td>INDIAN
BANK</td></tr><tr><td>21</td><td>Lal atharjee</td><td>Ram
Singh</td><td>G.N.L.</td><td>Ramraj</td><td>16082010</td><td>000939</td><td>UBIN0014985</td><td>UNION
BANK</td></tr><tr><td>22</td><td>Ram Parasad</td><td>Syambal</td><td>G.N.L.</td><td>Ful
Thariga</td><td>5030410</td><td>716</td><td>IDIB0000501</td><td>INDIAN
BANK</td></tr><tr><td>23</td><td>Satynazoyan</td><td>Jay
Pal</td><td>G.N.L.</td><td>Satyhroyan</td><td>.43650210</td><td>006402</td><td>UBIN00543659</td><td>UNION
BANK</td></tr><tr><td>24</td><td>Phool Singh</td><td>Havivankh</td><td>G.N.L.</td><td>Phool
Singh</td><td>4817001700</td><td>006402</td><td>PUMBO431700</td><td>Panational
bank</td></tr><tr><td>25</td><td>Ram Kuman</td><td>Havivankh</td><td>G.N.L.</td><td>Ram
Kuman</td><td>204102785</td><td>25</td><td>Fino 0001157</td><td>Finno Bynanbank</td></tr><tr><td>26</td><td>Satyaj
Kuman</td><td>Ram Pyare</td><td>G.N.L.</td><td>Sathyj Kuman</td><td>4365021000</td><td>05335</td><td>VBIN
0543659</td><td>UNION BANK</td></tr><tr><td>27</td><td>Shi Groomd</td><td>Lal man</td><td>G.N.L.</td><td>Sony
Kumanzi</td><td>.5910473</td><td>1131</td><td>IDIB00006441</td><td>INDIAN
BANK</td></tr><tr><td>28</td><td>Sunil</td><td>Sony Kumanzi</td><td>G.N.L.</td><td>Sony
Kumanzi</td><td>.7379710960</td><td>80</td><td>IDIB0000642</td><td>INDIAN BANK</td></tr></tbody></table>"""

# Extract all cell content (both td and th)
raw_rows = re.findall(r'<tr>(.*?)</tr>', html, re.DOTALL)

rows_by_sno = {}
for row_html in raw_rows:
    cells = re.findall(r'<t[dh][^>]*>(.*?)(?:</t[dh]>|$)', row_html, re.DOTALL)
    cells = [re.sub(r'\s+', ' ', c.strip()) for c in cells]
    cells = [c for c in cells if c]
    if not cells:
        continue
    if cells[0].isdigit():
        sno = int(cells[0])
        prev = rows_by_sno.get(sno, [cells[0]])
        # merge: take longer of the two
        if len(cells) > len(prev):
            rows_by_sno[sno] = cells
    elif cells[0] in ('Name', 'Father name', 'Designation', 'Date', 'PAGE NO.', 'TERM'):
        continue  # skip header
    else:
        # maybe continuation without serial
        pass

# print merged
print("=== Merged Rows ===")
cols = ['S.No', 'Name', 'Father Name', 'Designation', 'Account Holder', 'Account No', 'Code', 'IFSC', 'Bank Name']
print(' | '.join(cols))
print('-' * 120)
for sno in sorted(rows_by_sno):
    r = rows_by_sno[sno]
    # pad to 9 cols
    r = r + [''] * (9 - len(r))
    print(f" | ".join(f"{c:<20}" if i > 0 else f"{c:<4}" for i, c in enumerate(r[:9])))

print(f"\n=== Summary ===")
print(f"Total rows: {len(rows_by_sno)}")
