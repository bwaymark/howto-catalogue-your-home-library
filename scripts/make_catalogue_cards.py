"""Turn a LibraryThing Excel export into a mail-merge file for catalogue cards.

Usage: python make_catalogue_cards.py LibraryThing_export.xlsx Catalogue_cards_merge.xlsx

Produces one row per card (author, added author, title, Dewey subject), sorted in
filing order, ready for Brother P-touch Editor's database merge or a Word mail merge.
"""
import re, sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SRC = sys.argv[1]
OUT = sys.argv[2]

MAIN = {'0':'Computer science, information & general works','1':'Philosophy & psychology','2':'Religion',
'3':'Social sciences','4':'Language','5':'Science','6':'Technology','7':'Arts & recreation',
'8':'Literature','9':'History & geography'}

SECT = {
'001':'Knowledge','003':'Systems','004':'Computer science','005':'Computer programming, programs & data',
'006':'Special computer methods','028':'Reading & use of other information media','032':'General encyclopaedic works in English',
'070':'News media, journalism & publishing','128':'Humankind','133':'Parapsychology & occultism',
'149':'Other philosophical systems','153':'Conscious mental processes & intelligence',
'155':'Differential & developmental psychology','158':'Applied psychology','174':'Occupational ethics',
'177':'Ethics of social relations','201':'Religious mythology & social theology','202':'Doctrines',
'214':'Theodicy','220':'Bible','222':'Historical books of Old Testament','230':'Christianity & Christian theology',
'231':'God','232':'Jesus Christ & his family','241':'Christian ethics','242':'Devotional literature',
'248':'Christian experience, practice & life','253':'Pastoral office & work','261':'Social theology',
'262':'Ecclesiology','270':'History of Christianity','277':'Christianity in North America',
'284':'Protestant denominations of Continental origin','289':'Other denominations & sects (incl. Quakers)',
'291':'Comparative religion','293':'Germanic religion','294':'Religions of Indic origin',
'297':'Islam, Bábism & Bahá’í Faith','299':'Other religions','302':'Social interaction',
'303':'Social processes','304':'Factors affecting social behaviour','305':'Groups of people',
'306':'Culture & institutions','320':'Political science','321':'Systems of governments & states',
'324':'The political process','330':'Economics','331':'Labour economics','332':'Financial economics',
'333':'Economics of land & energy','335':'Socialism & related systems','338':'Production',
'340':'Law','355':'Military science','363':'Other social problems & services','364':'Criminology',
'370':'Education','381':'Commerce','384':'Communications','385':'Railway transport',
'387':'Water, air & space transport','394':'General customs','398':'Folklore','400':'Language',
'410':'Linguistics','420':'English & Old English languages','422':'English etymology',
'423':'English dictionaries','428':'Standard English usage','468':'Standard Spanish usage',
'495':'Languages of East & Southeast Asia','500':'Science','508':'Natural history','529':'Chronology',
'530':'Physics','539':'Modern physics','546':'Inorganic chemistry','573':'Physiological systems of animals',
'576':'Genetics & evolution','579':'Micro-organisms, fungi & algae','581':'Natural history of plants',
'590':'Animals','591':'Natural history of animals','610':'Medicine & health','612':'Human physiology',
'613':'Personal health & safety','615':'Pharmacology & therapeutics','616':'Diseases','630':'Agriculture',
'631':'Agricultural techniques, equipment & materials','635':'Garden crops (horticulture)',
'636':'Animal husbandry & pets','640':'Home & family management','641':'Food & drink',
'643':'Housing & household equipment','646':'Sewing, clothing & personal living',
'647':'Management of public households (pubs, hotels)','648':'Housekeeping',
'649':'Child rearing & home care of people','650':'Management & auxiliary services','658':'General management',
'659':'Advertising & public relations','668':'Technology of other organic products',
'701':'Philosophy & theory of fine arts','709':'History of the arts','711':'Area planning',
'720':'Architecture','736':'Carving & carvings','738':'Ceramic arts','741':'Drawing & drawings',
'745':'Decorative arts','746':'Textile arts','782':'Vocal music','790':'Recreational & performing arts',
'791':'Public performances','793':'Indoor games & amusements','796':'Athletic & outdoor sports & games',
'808':'Rhetoric & collections of literature','812':'American drama','813':'American fiction',
'818':'American miscellaneous writings','822':'English drama','823':'English fiction','824':'English essays',
'828':'English miscellaneous writings','829':'Old English (Anglo-Saxon) literature','833':'German fiction',
'839':'Other Germanic literatures','843':'French fiction','869':'Portuguese literature',
'895':'Literatures of East & Southeast Asia','900':'History & geography','901':'Philosophy & theory of history',
'909':'World history','910':'Geography & travel','912':'Atlases & maps','914':'Geography of & travel in Europe',
'915':'Geography of & travel in Asia','920':'Biography & genealogy','938':'Ancient Greece',
'940':'History of Europe','941':'History of the British Isles','942':'History of England & Wales',
'943':'History of Germany & central Europe','951':'History of China','952':'History of Japan',
'954':'History of India & South Asia','960':'History of Africa','966':'History of West Africa',
'967':'History of Central Africa','971':'History of Canada','972':'History of Mexico & Central America',
'973':'History of the United States'}

# Fix individual author names that LibraryThing has mangled.
# Keys are the name exactly as exported; a value of None means "no author" (file under title).
# Example: {'strauss-neil': 'Strauss, Neil', 'Unknown': None}
NAMEFIX = {'Unknown': None, 'Collectif': None}
HONORIFICS = re.compile(r'\b(Professor|Prof\.?|Dr\.?|Mr\.?|Mrs\.?|Ms\.?|PhD|Sir)\s*', re.I)

AUTHOR_ROLES = {'author','autor','auteur','autore','editor','compiler','primary contributor',''}

def clean_name(n):
    if pd.isna(n): return None
    n = n.strip()
    if n in NAMEFIX: return NAMEFIX[n]
    n = HONORIFICS.sub('', n).strip(' ,')
    if n.count(',') == 1 and not n.split(',')[1].strip():   # "PhD, Jane Smith" -> "Jane Smith"
        n = n.split(',')[0].strip()
    return n or None

def strip_article(t):
    return re.sub(r'^(the|a|an|le|la|les|l\'|der|die|das|el)\s+', '', t.strip(), flags=re.I)

def code_from(text):
    letters = re.sub(r'[^A-Za-zÀ-ÿ]', '', text)
    return letters[:3].upper()

def author_code(author, title):
    if author:
        return code_from(strip_article(author.split(',')[0]))
    return code_from(strip_article(title))

def call_no(d):
    if pd.isna(d): return ''
    d = d.strip()
    whole, _, frac = d.partition('.')
    whole = whole.zfill(3)
    frac = frac[:3].rstrip('0')
    return f'{whole}.{frac}' if frac else whole

def file_key(s):
    s = strip_article(s or '').lower()
    s = re.sub(r'[^\w\s]', '', s)
    return re.sub(r'\s+', ' ', s).strip()

def imprint(p):
    if pd.isna(p): return ''
    m = re.match(r'^(.*?\(\d{4}\))', p)
    return m.group(1) if m else p.split(',')[0]

df = pd.read_excel(SRC, dtype=str)
rows, missing = [], []
for _, r in df.iterrows():
    title = (r['Title'] or '').strip()
    sc = int(r['Sort Character']) if pd.notna(r['Sort Character']) else 1
    title_file = title[sc-1:] if sc > 1 else title
    author = clean_name(r['Primary Author'])
    cn = call_no(r['Dewey Decimal'])
    code = author_code(author, title_file)
    sect = cn[:3] if cn else ''
    subj = f"{sect} {SECT.get(sect, MAIN.get(sect[:1], ''))}" if sect else ''
    pages = r['Page Count'] if pd.notna(r['Page Count']) else ''
    base = dict(call=cn, code=code, author=author or '', title=title, imprint=imprint(r['Publication']),
                pages=pages, subject=subj, main=MAIN.get(sect[:1], '') if sect else '', bookid=r['Book ID'])
    if not cn:
        missing.append(dict(bookid=r['Book ID'], title=title, author=author or ''))
    # main entry: author if known, otherwise title
    if author:
        rows.append({**base, 'seq':'1 Author', 'heading':author, 'key':file_key(author)+' | '+file_key(title_file)})
    # added authors (co-authors, editors) only
    if pd.notna(r['Secondary Author']):
        names = r['Secondary Author'].split('|')
        roles = (r['Secondary Author Role'] or '').split('|') if pd.notna(r['Secondary Author Role']) else ['']*len(names)
        for i, n in enumerate(names):
            role = roles[i].strip().lower() if i < len(roles) else ''
            n = clean_name(n)
            if n and role in AUTHOR_ROLES and n != author:
                rows.append({**base, 'seq':'2 Added author', 'heading':n, 'key':file_key(n)+' | '+file_key(title_file)})
    # title card
    rows.append({**base, 'seq':'3 Title', 'heading':title_file, 'key':file_key(title_file)})
    # subject card, skipping fiction (8x3) where shelf order already does the job
    is_fiction = sect[:1] == '8' and sect[2:3] == '3' and sect[1:2] != '0'
    if sect and not is_fiction:
        rows.append({**base, 'seq':'4 Subject', 'heading':subj, 'key':cn+' | '+file_key(author or title_file)})

cards = pd.DataFrame(rows).sort_values(['seq','key']).reset_index(drop=True)
cards['seq'] = cards['seq'].str[2:]

COLS = [('Catalogue','seq',14),('Heading','heading',40),('Call number','call',12),('Author code','code',11),
        ('Author','author',28),('Title','title',55),('Imprint','imprint',36),('Pages','pages',7),
        ('Dewey subject','subject',42),('Dewey main class','main',30),('LibraryThing ID','bookid',14)]

wb = Workbook()
F = Font(name='Arial', size=10); FB = Font(name='Arial', size=10, bold=True, color='FFFFFF')
HF = PatternFill('solid', fgColor='3A4A5C')

ws = wb.active; ws.title = 'Cards'
for c,(h,_,w) in enumerate(COLS,1):
    cell = ws.cell(1,c,h); cell.font = FB; cell.fill = HF
    ws.column_dimensions[get_column_letter(c)].width = w
for i, rec in enumerate(cards.to_dict('records'), 2):
    for c,(_,k,_) in enumerate(COLS,1):
        v = rec[k]
        if k == 'pages' and v:
            nums = re.findall(r'\d+', v); v = int(nums[-1]) if nums else ''
        cell = ws.cell(i,c,v); cell.font = F
        if k in ('call','bookid'): cell.number_format = '@'
ws.freeze_panes = 'A2'; ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{len(cards)+1}"

# Summary with live counts
s = wb.create_sheet('Summary')
s.column_dimensions['A'].width = 34; s.column_dimensions['B'].width = 12; s.column_dimensions['C'].width = 70
s['A1'] = 'Cards by catalogue'; s['A1'].font = Font(name='Arial', size=12, bold=True)
n = len(cards)+1
for j, cat in enumerate(['Author','Added author','Title','Subject'], 3):
    s.cell(j,1,cat).font = F
    s.cell(j,2,f'=COUNTIF(Cards!$A$2:$A${n},A{j})').font = F
s['A7'] = 'Total cards'; s['A7'].font = Font(name='Arial', size=10, bold=True)
s['B7'] = '=SUM(B3:B6)'; s['B7'].font = Font(name='Arial', size=10, bold=True)
s['A8'] = 'DK-22205 rolls at 115 mm per label'; s['A8'].font = F
s['B8'] = '=ROUNDUP(B7/ROUNDDOWN(30480/115,0),0)'; s['B8'].font = F
s['C8'] = 'Assumes 30.48 m per roll and 115 mm labels; change 115 if you set a different length.'; s['C8'].font = F
notes = [
 'Notes',
 'Call numbers are LibraryThing’s Dewey numbers cut to three decimal places, so they fit a spine label.',
 'Author code is the first three letters of the author’s surname, or of the title where there is no author.',
 'Added-author cards cover co-authors and editors only, not illustrators, translators or foreword writers.',
 'Subject cards use the Dewey section (three-digit) caption and file in number order, like a classified catalogue.',
 'Fiction (813, 823, 833, 843) has no subject cards; filter Catalogue = Subject on the Cards sheet to check.',
 'Titles file without a leading The/A/An, using LibraryThing’s sort character.',
 'A few author names were tidied (titles such as Dr/Professor removed, names inverted).',
]
for j, t in enumerate(notes, 10):
    s.cell(j,1,t).font = Font(name='Arial', size=10, bold=(j==10))

m = wb.create_sheet('No Dewey')
for c,h in enumerate(['LibraryThing ID','Title','Author'],1):
    cell = m.cell(1,c,h); cell.font = FB; cell.fill = HF
for i, rec in enumerate(missing, 2):
    m.cell(i,1,rec['bookid']).font = F; m.cell(i,2,rec['title']).font = F; m.cell(i,3,rec['author']).font = F
m.column_dimensions['A'].width = 14; m.column_dimensions['B'].width = 55; m.column_dimensions['C'].width = 30
m.cell(len(missing)+3,1,'These have no Dewey number in LibraryThing, so their cards have no call number and no subject card. Add a number in LibraryThing and re-export, or type it into the Cards sheet.').font = F

wb.save(OUT)
print(cards['seq'].value_counts().to_dict(), len(cards), len(missing))
print(sorted({x for x in cards['subject'] if x and x[4:]==''}))
