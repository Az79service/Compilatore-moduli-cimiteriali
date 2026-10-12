import streamlit as st
from PIL import Image, ImageOps, ImageEnhance, UnidentifiedImageError
import io, pytesseract, re, os, subprocess, shutil, tempfile, hmac
from docx import Document
from datetime import date, datetime
from pathlib import Path

st.set_page_config(page_title='Compilatore Moduli Cimiteriali Palermo', page_icon='⚰️', layout='wide')

# Optional shared-password gate for hosted deployments. Keep the password in the
# hosting provider's environment settings; never commit it to the repository.
_ACCESS_PASSWORD = os.environ.get("APP_ACCESS_PASSWORD", "")
# Streamlit Community Cloud exposes secrets through st.secrets, not as
# environment variables. Support both mechanisms without storing credentials
# in source control.
if not _ACCESS_PASSWORD:
    try:
        _ACCESS_PASSWORD = str(st.secrets.get("APP_ACCESS_PASSWORD", "") or "")
    except Exception:
        _ACCESS_PASSWORD = ""
_APP_ENV = os.environ.get("APP_ENV", "production").strip().lower()
# On Community Cloud, the presence of the configured secret also activates
# the password gate even though custom environment variables are unavailable.
if _ACCESS_PASSWORD and _APP_ENV == "development":
    _APP_ENV = "production"
if not _ACCESS_PASSWORD:
    if _APP_ENV in {"production", "prod"}:
        st.error("Configurazione di sicurezza incompleta: l'amministratore deve impostare APP_ACCESS_PASSWORD nell'hosting.")
        st.stop()
    st.warning("Accesso non protetto da password: per un uso locale configura APP_ACCESS_PASSWORD; non pubblicare così l'app su Internet.")
else:
    _entered_password = st.text_input("Password di accesso", type="password", key="app_access_password")
    if not hmac.compare_digest(str(_entered_password or ""), _ACCESS_PASSWORD):
        st.warning("Inserisci la password autorizzata per accedere al compilatore.")
        st.stop()
BASE=Path(__file__).resolve().parent; T=BASE/'templates'; OUT=BASE/'output'  # legacy fallback only; normal downloads use a temporary folder
MODULES={
 'Richiesta operazioni cimiteriali – I-01':'richiesta_operazioni_cimiteriali_i-01.docx',
 'Richiesta cremazione – I-04':'richiesta_cremazione_i-04.docx',
 'Richiesta autorizzazione trasporti':'modello_trasporto.docx',
 'Dichiarazione prossimi congiunti':'dichiarazione_prossimi_congiunti.docx',
 'Dichiarazione parenti per cremazione – I-05-1':'dichiarazione_parenti_cremazione.docx',
 'Richiesta concessione loculo/nicchia':'richiesta_loculo_nicchia.docx',
 'Consenso cremazione – I-05':'consenso_cremazione.docx',
 'Autorizzazione tumulazione salma/resti – I-15':'autorizzazione_tumulazione_I15.docx',
}

def safe_open_uploaded_image(uploaded, max_bytes=20 * 1024 * 1024, max_pixels=40_000_000):
    """Validate an uploaded image before decoding it for OCR.

    Accepts Streamlit UploadedFile-like objects exposing getvalue(). Rejects
    oversized, malformed, unsupported, or excessively large images cleanly.
    """
    raw = uploaded.getvalue()
    if not raw:
        raise ValueError('Il file caricato è vuoto.')
    if len(raw) > max_bytes:
        raise ValueError('Il file supera il limite di 20 MB.')
    try:
        with Image.open(io.BytesIO(raw)) as probe:
            if (probe.format or '').upper() not in {'JPEG', 'PNG', 'WEBP'}:
                raise ValueError('Formato immagine non supportato: usa JPG, PNG o WEBP.')
            width, height = probe.size
            if width < 1 or height < 1 or width * height > max_pixels:
                raise ValueError('Immagine troppo grande: riduci le dimensioni e riprova.')
            probe.verify()
        with Image.open(io.BytesIO(raw)) as image_file:
            image_file.load()
            return image_file.copy()
    except ValueError:
        raise
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError) as exc:
        raise ValueError('Il file non è un’immagine valida o è danneggiato.') from exc


def ocr(im):
    im=ImageOps.exif_transpose(im).convert('L'); im=ImageOps.autocontrast(im); im=ImageEnhance.Contrast(im).enhance(1.5)
    try: return pytesseract.image_to_string(im,lang='ita+eng',config='--psm 6')
    except Exception: return ''

def norm(s): return re.sub(r'\s+',' ',s or '').strip()
def first(p,t):
    m=re.search(p,t,re.I|re.S); return norm(m.group(1)) if m else ''
def parse_date(s):
    s=(s or '').replace('.','/').replace('-','/');
    m=re.search(r'(\d{1,2})/(\d{1,2})/(\d{2,4})',s)
    return f'{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}' if m else ''

def extract_id(text):
    t=norm(text).upper()
    cf=first(r'(?:CODICE FISCALE|C\.F\.)\s*[:\-]?\s*([A-Z0-9]{16})',t)
    dn=first(r'(?:DATA DI NASCITA|NATO/A IL|NASCITA)\s*[:\-]?\s*(\d{1,2}[ /.-]\d{1,2}[ /.-]\d{2,4})',t)
    # Common Italian ID OCR patterns; values remain editable.
    cogn=first(r'COGNOME\s*[:\-]?\s*([A-ZÀ-ÖØ-Ý\' ]{2,})\s+(?:NOME|NOME)',t)
    nome=first(r'NOME\s*[:\-]?\s*([A-ZÀ-ÖØ-Ý\' ]{2,}?)(?=\s+(?:DATA DI NASCITA|CODICE FISCALE|C\.F\.|LUOGO DI NASCITA|NATO/A A|NATO A|INDIRIZZO|RESIDENZA|COGNOME)\b|$)',t)
    luogo=first(r'(?:LUOGO DI NASCITA|NATO/A A|NATO A)\s*[:\-]?\s*([A-ZÀ-ÖØ-Ý\' .()-]{2,}?)(?=\s+(?:DATA DI NASCITA|CODICE FISCALE|C\.F\.|INDIRIZZO|RESIDENZA|COGNOME|NOME)\b|$)',t)
    indir=first(r'(?:VIA|VIA/PIAZZA|INDIRIZZO|RESIDENZA)\s*[:\-]?\s*([A-Z0-9À-ÖØ-Ý\' .,/\-]{3,})',t)
    return {'cognome':cogn,'nome':nome,'codice_fiscale':cf,'data_nascita':parse_date(dn),'luogo_nascita':luogo,'stato_nascita':'ITALIA','cittadinanza':'ITALIANA','indirizzo':indir,'civico':'','cap':'','comune_residenza':'','provincia_residenza':'','sesso':'','email':'','telefono':''}

def cell(t,r,c):
    try: return t.cell(r,c)
    except Exception: return None

def cell_text(c):
    return re.sub(r"\s+", " ", c.text or "").strip()

def put(t,r,c,v):
    """Write only into an empty cell: never erase an official form label."""
    x=cell(t,r,c)
    if x is not None and not cell_text(x) and str(v or '').strip():
        x.text=str(v).strip()

def chars(t,r,cols,v):
    s=re.sub(r'\s+','',str(v or '')).upper()
    seen=set(); k=0
    for c in cols:
        if k>=len(s): break
        x=cell(t,r,c)
        if x is None or x._tc in seen: continue
        seen.add(x._tc)
        if cell_text(x): continue
        put(t,r,c,s[k]); k+=1

def chars_after_label(t,r,label,v):
    """Fill one character per blank box immediately after a printed field label."""
    value=re.sub(r'\s+','',str(v or '')).upper()
    if not value or r>=len(t.rows): return
    row=t.rows[r].cells; started=False; seen=set(); k=0
    for c,x in enumerate(row):
        txt=cell_text(x); ident=x._tc
        if label.lower() in txt.lower():
            started=True; seen.add(ident); continue
        if not started: continue
        if ident in seen: continue
        seen.add(ident)
        if txt:
            break
        put(t,r,c,value[k]); k+=1
        if k>=len(value): break

def put_unique_values(t,r,cols,values):
    seen=set(); k=0
    for c in cols:
        if k>=len(values): break
        x=cell(t,r,c)
        if x is None or x._tc in seen: continue
        seen.add(x._tc)
        if cell_text(x): continue
        put(t,r,c,values[k]); k+=1

def find_label(t, label, rows=None):
    wanted=norm(label).lower()
    indexes=range(len(t.rows)) if rows is None else rows
    for ri in indexes:
        if ri < 0 or ri >= len(t.rows): continue
        for ci, c in enumerate(t.rows[ri].cells):
            txt=cell_text(c).lower()
            if txt and wanted in txt:
                return ri,ci
    return None

def put_below_label(t, label, value, rows=None, occurrence='first'):
    """Place a value in the row below its label, aligned to the label's first cell."""
    if not str(value or '').strip(): return
    hits=[]
    indexes=range(len(t.rows)) if rows is None else rows
    wanted=norm(label).lower()
    for ri in indexes:
        if ri < 0 or ri >= len(t.rows): continue
        for ci,c in enumerate(t.rows[ri].cells):
            if wanted in cell_text(c).lower(): hits.append((ri,ci))
    if not hits: return
    ri,ci=hits[-1] if occurrence=='last' else hits[0]
    if ri+1 < len(t.rows): put(t,ri+1,ci,value)

def put_after_label(t, label, value, rows=None, count=1):
    """Place a value in the blank field following a printed label, using merged-cell width."""
    if not str(value or '').strip(): return
    wanted=norm(label).lower(); indexes=range(len(t.rows)) if rows is None else rows
    for ri in indexes:
        if ri < 0 or ri >= len(t.rows): continue
        row=t.rows[ri].cells
        label_cols=[ci for ci,c in enumerate(row) if wanted in cell_text(c).lower()]
        if not label_cols: continue
        ci=max(label_cols); seen=set(); candidates=[]
        for j in range(ci+1,len(row)):
            x=row[j]; ident=x._tc
            if ident in seen: continue
            seen.add(ident); existing=cell_text(x)
            if existing:
                if wanted in existing.lower(): continue
                break
            span=x._tc.tcPr.gridSpan.val if x._tc.tcPr.gridSpan is not None else 1
            candidates.append((span,j))
        if not candidates: return
        # For text fields (civil number, email, telephone), use the widest blank merged cell.
        _, target=max(candidates)
        if count==1: put(t,ri,target,value)
        else: chars_after_label(t,ri,label,value)
        return

def date_parts(s):
    p=parse_date(s).split('/'); return p if len(p)==3 else ['','','']

def fill_identity_table(t,p,kind='applicant'):
    """Fill only identity-grid layouts checked against the supplied Word templates."""
    cols=len(t.columns); rows=len(t.rows)
    if cols==42 and rows>=13 and kind=='deceased':
        # Dedicated deceased block in “Dichiarazione prossimi congiunti”.
        # Its 42-column geometry differs from the loculo/nicchia form; map only
        # fields aligned with the printed labels and leave CF/sex manual.
        put(t,3,1,p.get('cognome')); put(t,3,26,p.get('nome'))
        put_unique_values(t,5,[28,29,31],date_parts(p.get('data_nascita')))
        put_below_label(t,'Stato',p.get('stato_nascita','ITALIA'),rows=[6])
        put_below_label(t,'Comune e provincia',f"{p.get('comune_residenza','')} ({p.get('provincia_residenza','')})".strip(' ()'),rows=[6])
        put_below_label(t,'Cittadinanza',p.get('cittadinanza','ITALIANA'),rows=[6])
        put_below_label(t,'Luogo di nascita',p.get('luogo_nascita'),rows=[7])
        return
    if cols==47 and rows>=11 and kind=='consent_applicant':
        # The consent form's applicant block shares the 47-column count with I-15,
        # but its merged-cell layout differs. Only name/date are visually confirmed.
        put(t,1,1,p.get('cognome')); put(t,1,28,p.get('nome'))
        put_unique_row_values(t,3,[30,31,33],date_parts(p.get('data_nascita')))
        return
    if cols==47 and rows>=11 and kind=='i15_applicant':
        # Applicant block in form I-15. Populate fields only in blank cells
        # aligned with labels in the supplied DOCX template.
        put(t,1,1,p.get('cognome')); put(t,1,28,p.get('nome'))
        # Keep C.F. manual: the merged cells do not map one-to-one to 16 boxes.
        put_unique_row_values(t,3,[29,30,31],date_parts(p.get('data_nascita')))
        put(t,5,8,p.get('stato_nascita','ITALIA'))
        put(t,5,24,p.get('luogo_nascita'))
        put(t,5,34,p.get('cittadinanza','ITALIANA'))
        put_row_cell(t,7,5,p.get('indirizzo'))
        put_row_cell(t,7,33,p.get('civico'))
        put_row_chars(t,7,[39,40,41,43,44],p.get('cap'))
        # Row 8 contains the printed 'Comune e provincia' label; its value box is row 9.
        put_row_cell(t,9,1,f"{p.get('comune_residenza','')} ({p.get('provincia_residenza','')})".strip())
        put_row_cell(t,9,20,p.get('email'))
        put_row_cell(t,9,38,p.get('telefono'))
        return
    if cols==43 and rows>=9 and kind=='deceased':
        # I-15 secondary identity grid; only the clearly aligned identity
        # fields are automatic until the whole block is visually validated.
        put(t,1,1,p.get('cognome')); put(t,1,28,p.get('nome'))
        # Keep C.F. manual until the individual character boxes are verified.
        put_unique_row_values(t,3,[29,30,32],date_parts(p.get('data_nascita')))
        # The secondary I-15 grid has a different visual row order; keep birth details manual.
        return
    if cols==37 and rows>=8:
        # Compact relative-person block in “Dichiarazione prossimi congiunti”.
        # Name/date/birth fields align with the template labels; CF and sex remain manual.
        put(t,1,1,p.get('cognome')); put(t,1,23,p.get('nome'))
        put_unique_values(t,3,[25,26,27],date_parts(p.get('data_nascita')))
        put_below_label(t,'Stato',p.get('stato_nascita','ITALIA'),rows=[4])
        put_below_label(t,'Comune e provincia',f"{p.get('comune_residenza','')} ({p.get('provincia_residenza','')})".strip(' ()'),rows=[4])
        put_below_label(t,'Cittadinanza',p.get('cittadinanza','ITALIANA'),rows=[4])
        put_below_label(t,'Luogo di nascita',p.get('luogo_nascita'),rows=[5])
        return
    if cols not in (40,44,48):
        # Do not guess coordinates in other grid layouts; preserving the blank template is safer.
        return
    # Name field begins at a different grid column in the 48-column forms.
    put(t,1,1,p.get('cognome')); put(t,1,28 if cols==48 else 26,p.get('nome'))
    # The 48-column forms have a different merged-cell geometry; leave their CF boxes for manual verification.
    if cols != 48:
        chars_after_label(t,3,'C.F.',p.get('codice_fiscale'))
    date_cols={40:range(27,35),44:range(27,35),48:range(26,34)}[cols]
    put_unique_values(t,3,date_cols,date_parts(p.get('data_nascita')))
    # Birth fields use the empty cells immediately below their printed labels.
    if cols==48:
        state_col,place_col,cit_col=9,25,34
    else:
        state_col,place_col,cit_col=8,23,32
    put(t,5,state_col,p.get('stato_nascita','ITALIA'))
    put(t,5,place_col,p.get('luogo_nascita'))
    put(t,5,cit_col,p.get('cittadinanza','ITALIANA'))
    # The deceased-only 9-row grids do not contain address/residence fields.
    if rows>=11:
        put(t,7,5,p.get('indirizzo'))
        if cols==44:
            put_after_label(t,'N. Civ',p.get('civico'),rows=[7])
            chars_after_label(t,7,'CAP',p.get('cap'))
        # In 48-column forms these three fields need a visual cell-by-cell check before automation.
        put(t,9,1,f"{p.get('comune_residenza','')} ({p.get('provincia_residenza','')})".strip())
        if cols==48:
            put(t,9,21,p.get('email'))
        else:
            put_after_label(t,'E-m@il',p.get('email'),rows=[9])
            put_after_label(t,'Tel.',p.get('telefono'),rows=[9])

def fill_transport(doc,app,dec,extra):
    t=doc.tables; fill_identity_table(t[0],app); put(t[1],0,10,extra.get('qualita')); fill_deceased_transport(t[2],dec,extra)
    put(t[3],0,2,extra.get('indirizzo_morte')); put(t[3],1,2,dec.get('indirizzo')); put(t[3],1,7,dec.get('civico')); chars(t[3],1,[12,13,14,15,16],dec.get('cap')); put(t[3],3,1,f"{dec.get('comune_residenza','')} ({dec.get('provincia_residenza','')})".strip(' ()'))
    put(t[6],1,2,extra.get('ora')); put(t[6],1,5,extra.get('data_servizio')); put(t[6],5,2,extra.get('destinazione'))  # vehicle plate boxes are left for manual entry

def fill_deceased_transport(t,p,e):
    fill_identity_table(t,p,kind='deceased')
    put_after_label(t,'Luogo di morte',e.get('luogo_morte'),rows=range(len(t.rows)))
    parts=date_parts(e.get('data_morte'))
    if all(parts):
        hit=find_label(t,'Data di morte',rows=range(len(t.rows)))
        if hit:
            r,c=hit; put_unique_values(t,r,range(c+1,min(len(t.columns),c+10)),parts)

def fill_next_of_kin_deceased(t, p, f):
    """Use the template's printed labels to fill the deceased identity block safely."""
    fill_identity_table(t, p, kind='deceased')

def fill_deceased_grid(t, p, module, f):
    """Fill only deceased grids whose coordinates were confirmed in the supplied templates."""
    if module == 'Dichiarazione prossimi congiunti':
        fill_next_of_kin_deceased(t, p, f)
        return
    fill_identity_table(t, p, kind='deceased')
    if module == 'Dichiarazione parenti per cremazione – I-05-1':
        put_after_label(t,'Luogo di morte',f.get('luogo_morte'),rows=range(len(t.rows)))
        parts=date_parts(f.get('data_morte'))
        if all(parts):
            hit=find_label(t,'Data di morte',rows=range(len(t.rows)))
            if hit:
                r,c=hit; j=c+1; k=0
                while j<len(t.columns) and k<3:
                    if cell_text(t.cell(r,j)):
                        j+=1; continue
                    put(t,r,j,parts[k]); j+=1; k+=1
        return
    if module == 'Richiesta operazioni cimiteriali – I-01':
        put(t, 7, 8, f.get('luogo_morte'))
        dm = date_parts(f.get('data_morte')); put(t, 7, 34, dm[0] if dm[0] else ''); put(t, 7, 35, dm[1] if dm[1] else ''); put(t, 7, 37, dm[2] if dm[2] else '')
    elif module == 'Richiesta cremazione – I-04':
        put(t, 7, 8, f.get('luogo_morte'))
        dm = date_parts(f.get('data_morte')); put(t, 7, 34, dm[0] if dm[0] else ''); put(t, 7, 35, dm[1] if dm[1] else ''); put(t, 7, 37, dm[2] if dm[2] else '')
    elif module == 'Consenso cremazione – I-05':
        put(t, 7, 8, f.get('luogo_morte'))
        dm = date_parts(f.get('data_morte')); put(t, 7, 34, dm[0] if dm[0] else ''); put(t, 7, 35, dm[1] if dm[1] else ''); put(t, 7, 37, dm[2] if dm[2] else '')
    elif module == 'Autorizzazione tumulazione salma/resti – I-15':
        put(t, 7, 8, f.get('luogo_morte'))
        dm = date_parts(f.get('data_morte')); put(t, 7, 32, dm[0] if dm[0] else ''); put(t, 7, 33, dm[1] if dm[1] else ''); put(t, 7, 35, dm[2] if dm[2] else '')

def put_row_cell(t, r, c, value):
    """Write to a row's actual cell list; some official tables have uneven row widths."""
    if r < 0 or r >= len(t.rows) or c < 0 or c >= len(t.rows[r].cells): return
    x = t.rows[r].cells[c]
    if not cell_text(x) and str(value or '').strip(): x.text = str(value).strip()


def put_unique_row_values(t, r, cols, values):
    """Fill distinct blank cells using row coordinates, preserving merged-cell geometry."""
    if r < 0 or r >= len(t.rows): return
    seen = set(); k = 0; row = t.rows[r].cells
    for c in cols:
        if k >= len(values) or c < 0 or c >= len(row): break
        x = row[c]
        if x._tc in seen: continue
        seen.add(x._tc)
        if cell_text(x): continue
        if str(values[k] or '').strip(): x.text = str(values[k]).strip()
        k += 1


def put_row_chars(t, r, cols, value):
    """Fill a sequence of character boxes using actual row cells, without touching labels."""
    value = re.sub(r'\s+', '', str(value or '')).upper()
    if not value or r < 0 or r >= len(t.rows): return
    seen = set(); k = 0; row = t.rows[r].cells
    for c in cols:
        if k >= len(value) or c < 0 or c >= len(row): break
        x = row[c]
        if x._tc in seen: continue
        seen.add(x._tc)
        if cell_text(x): continue
        x.text = value[k]; k += 1


def fill_loculo_nicchia(doc, dec, f):
    """Fill the deceased block and cemetery priority in the loculo/nicchia form."""
    if len(doc.tables) > 1:
        t = doc.tables[1]  # The template places this grid immediately after 'per la salma di:'
        put(t, 1, 1, dec.get('cognome'))
        put(t, 1, 22, dec.get('nome'))
        put_unique_values(t, 3, [2, 5, 7], date_parts(dec.get('data_nascita')))
        # The template's merged cells make birthplace/residence placement unsafe; leave these manual.
        put(t, 11, 3, f.get('luogo_morte', ''))
        death_parts = date_parts(f.get('data_morte'))
        if all(death_parts):
            put_unique_values(t, 11, [36, 37, 38], death_parts)
    # The printed form explicitly requests a rank in the blank cell at column 3.
    # The UI asks for the first preference only, so write only rank 1 and leave the
    # remaining two rows untouched rather than inventing preferences.
    preference = str(f.get('preferenza_cimitero', ''))
    match = re.match(r'\s*([123])', preference)
    if doc.tables and match:
        selected_row = 3 + int(match.group(1))
        if selected_row < len(doc.tables[0].rows):
            put_row_cell(doc.tables[0], selected_row, 3, '1')


def fill_generic_identity(doc, app, dec=None, choice='', f=None, relatives=None):
    relatives=relatives or []
    f=f or {}
    if choice == 'Richiesta concessione loculo/nicchia':
        # Never substitute applicant data for deceased data: keep the deceased block blank if absent.
        fill_loculo_nicchia(doc, dec or {}, f)
        return
    grids = {
        'Richiesta operazioni cimiteriali – I-01': (0, 5),
        'Richiesta cremazione – I-04': (0, 1),
        'Dichiarazione prossimi congiunti': (0, 1),
        'Richiesta concessione loculo/nicchia': (None, None),
        'Consenso cremazione – I-05': (0, 1),
        'Autorizzazione tumulazione salma/resti – I-15': (0, 4),
        # In this declaration the first grid describes the deceased, not the applicant.
        'Dichiarazione parenti per cremazione – I-05-1': (None, 0),
    }
    ai, di = grids.get(choice, (None, None))
    if ai is not None and ai < len(doc.tables):
        t=doc.tables[ai]
        if len(t.rows)>=8 and len(t.columns)>=35:
            applicant_kind = 'i15_applicant' if choice == 'Autorizzazione tumulazione salma/resti – I-15' else ('consent_applicant' if choice == 'Consenso cremazione – I-05' else 'applicant')
            fill_identity_table(t,app,kind=applicant_kind)
    if di is not None and dec and di < len(doc.tables):
        t=doc.tables[di]
        if len(t.rows)>=8 and len(t.columns)>=35: fill_deceased_grid(t,dec,choice,f or {})
    # Fill only identity grids explicitly reserved for relatives in the two declarations.
    if choice == 'Dichiarazione prossimi congiunti':
        # The second table contains the deceased's identity, relationship, and burial details.
        if len(doc.tables)>1:
            deceased_table=doc.tables[1]
            put_after_label(deceased_table,'nella qualità di prossimo congiunto',f.get('qualita',''),rows=[0])
            put_after_label(deceased_table,'in quanto',f.get('grado_parentela',''),rows=[0])
            for label,key in [('Sezione','sezione'),('Lotto','lotto'),('Cimitero','cimitero'),('Fascicolo n°','fascicolo')]:
                put_below_label(deceased_table,label,f.get(key,''),rows=[11])
        slots=list(range(2,10))
        for ti,person in zip(slots,relatives):
            if ti < len(doc.tables):
                fill_identity_table(doc.tables[ti],person)
                put_after_label(doc.tables[ti],'grado di parentela con il defunto',f.get('grado_parentela',''),rows=range(len(doc.tables[ti].rows)))
    elif choice == 'Dichiarazione parenti per cremazione – I-05-1':
        if len(doc.tables)>1 and len(doc.tables[1].rows)>0 and len(doc.tables[1].columns)>1:
            put(doc.tables[1],0,1,(f or {}).get('numero_parenti',''))
        slots=[2,4,6,8,10,11]
        relation_tables=[3,5,7,9,12]
        for pos,(ti,person) in enumerate(zip(slots,relatives)):
            if ti < len(doc.tables): fill_identity_table(doc.tables[ti],person)
            if pos < len(relation_tables) and relation_tables[pos] < len(doc.tables):
                rt=doc.tables[relation_tables[pos]]
                if len(rt.rows)>1 and len(rt.columns)>1 and not cell_text(rt.cell(1,1)):
                    put(rt,1,1,f.get('grado_parentela',''))
    elif choice == 'Autorizzazione tumulazione salma/resti – I-15':
        # The concessionaire and burial-reference fields remain manual. Visual
        # checks showed that merged cells in this official template can make
        # automated placement spill into labels or adjacent boxes. Never infer
        # the concessionaire from the applicant; the operator completes these
        # fields directly in Word after reviewing the source documents.
        pass

def module_fields(name):
    base={
      'qualita':'','data_morte':'','luogo_morte':'','data_servizio':date.today().strftime('%d/%m/%Y'),'ora':'08:00','destinazione':'','targa':'',
      'grado_parentela':'','tipo_operazione':'','tipo_materiale':'','sezione':'','lotto':'','cimitero':'','fascicolo':'','preferenza_cimitero':'1','tipo_sepoltura':'loculo ipogeo in sepoltura',
      'concessionario':'','grado_concessionario':'','numero_parenti':'','dispersione':'','affidatario_ceneri':'','scelta_ceneri':'','impianto_cremazione':'Cimitero di S.M. dei Rotoli','agenzia':'',
      'luogo_data':date.today().strftime('%d/%m/%Y')}
    return base

def save_doc(doc,name,out_dir=None):
    """Save a generated document; production flow passes a per-request temporary directory."""
    target_dir=Path(out_dir) if out_dir else OUT
    target_dir.mkdir(parents=True,exist_ok=True)
    out=target_dir/(re.sub(r'[^A-Za-z0-9_.-]+','_',name)+'.docx')
    doc.save(out)
    return out

def pdf_from_docx(docx,out_dir=None):
    lo=shutil.which('libreoffice') or shutil.which('soffice')
    if not lo: return None
    target_dir=Path(out_dir) if out_dir else Path(docx).parent
    target_dir.mkdir(parents=True,exist_ok=True)
    pdf=target_dir/(Path(docx).stem+'.pdf')
    # Evita di restituire un PDF preesistente o incompleto dopo una conversione fallita.
    try:
        result=subprocess.run(
            [lo,'--headless','--convert-to','pdf','--outdir',str(target_dir),str(docx)],
            capture_output=True,text=True,timeout=60,check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if result.returncode != 0 or not pdf.is_file():
        return None
    try:
        if pdf.stat().st_size < 1000 or pdf.read_bytes()[:5] != b'%PDF-':
            return None
    except OSError:
        return None
    return pdf

st.title('⚰️ Compilatore modulistica cimiteriale – Palermo')
st.caption('Compilatore assistito: controlla i dati riconosciuti e verifica attentamente ogni campo del documento generato.')
with st.expander('Moduli disponibili',expanded=False):
    for x in MODULES: st.write('•',x)

uploads=st.file_uploader('1. Carica una o più carte d’identità',type=['jpg','jpeg','png','webp'],accept_multiple_files=True)
people=[]
if uploads:
    st.subheader('Dati riconosciuti (sempre modificabili)')
    for i,u in enumerate(uploads):
        try:
            uploaded_image = safe_open_uploaded_image(u)
        except ValueError as exc:
            st.error(f'File non elaborato ({u.name}): {exc}')
            continue
        recognized_text = ocr(uploaded_image)
        d = extract_id(recognized_text)
        d['_file'] = u.name
        if not recognized_text.strip() or not any(d.get(k) for k in ('cognome', 'nome', 'codice_fiscale', 'data_nascita', 'luogo_nascita', 'indirizzo')):
            st.warning(f'OCR senza risultati utili per {u.name}. Verifica che Tesseract e il pacchetto lingua italiana siano installati e funzionanti; puoi comunque inserire o correggere manualmente i dati qui sotto.')
        people.append(d)
        with st.expander(f'Persona {i+1} – {u.name}',expanded=True):
            c1,c2,c3=st.columns(3); d['cognome']=c1.text_input('Cognome',d['cognome'],key=f'cog{i}'); d['nome']=c2.text_input('Nome',d['nome'],key=f'nom{i}'); d['codice_fiscale']=c3.text_input('Codice fiscale',d['codice_fiscale'],key=f'cf{i}')
            c1,c2,c3=st.columns(3); d['data_nascita']=c1.text_input('Data di nascita',d['data_nascita'],key=f'dn{i}'); d['luogo_nascita']=c2.text_input('Luogo di nascita',d['luogo_nascita'],key=f'ln{i}'); d['sesso']=c3.selectbox('Sesso',['','M','F'],index=['','M','F'].index(d.get('sesso','')) if d.get('sesso','') in ['','M','F'] else 0,key=f'sx{i}')
            c1,c2,c3=st.columns(3); d['indirizzo']=c1.text_input('Via/Piazza',d['indirizzo'],key=f'ind{i}'); d['civico']=c2.text_input('N. civico',d['civico'],key=f'civ{i}'); d['cap']=c3.text_input('CAP',d['cap'],key=f'cap{i}')
            c1,c2=st.columns(2); d['comune_residenza']=c1.text_input('Comune di residenza',d['comune_residenza'],key=f'com{i}'); d['provincia_residenza']=c2.text_input('Provincia',d['provincia_residenza'],key=f'prov{i}')
            c1,c2,c3=st.columns(3); d['stato_nascita']=c1.text_input('Stato di nascita',d.get('stato_nascita','ITALIA'),key=f'stn{i}'); d['cittadinanza']=c2.text_input('Cittadinanza',d.get('cittadinanza','ITALIANA'),key=f'cit{i}'); d['email']=c3.text_input('E-mail',d.get('email',''),key=f'email{i}')
            d['telefono']=st.text_input('Telefono',d.get('telefono',''),key=f'tel{i}')

choice=st.selectbox('2. Scegli il modulo',list(MODULES))
f=module_fields(choice)
names=[f"{i+1} – {p.get('cognome','')} {p.get('nome','')}" for i,p in enumerate(people)]
app=people[0] if people else {'cognome':'','nome':'','codice_fiscale':'','data_nascita':'','luogo_nascita':'','stato_nascita':'ITALIA','cittadinanza':'ITALIANA','indirizzo':'','civico':'','cap':'','comune_residenza':'','provincia_residenza':'','sesso':'','email':'','telefono':''}
dec=people[1] if len(people)>1 else app.copy()
ai=0; di=min(1,len(people)-1) if people else 0
if names:
    c1,c2=st.columns(2)
    ai=c1.selectbox('Richiedente/dichiarante',range(len(names)),format_func=lambda i:names[i])
    # Do not silently treat the sole uploaded person as both applicant and deceased.
    deceased_options=[-1]+list(range(len(names)))
    deceased_index=2 if len(names)>1 else 0  # second uploaded person by default only when available
    di=c2.selectbox(
        'Defunto/a (se previsto)', deceased_options, index=deceased_index,
        format_func=lambda i: '— Non selezionato —' if i == -1 else names[i]
    )
    app=people[ai]
    dec=people[di] if di >= 0 else {
        'cognome':'','nome':'','codice_fiscale':'','data_nascita':'','luogo_nascita':'',
        'stato_nascita':'','cittadinanza':'','indirizzo':'','civico':'','cap':'',
        'comune_residenza':'','provincia_residenza':'','sesso':'','email':'','telefono':''
    }
relative_indices=[]
if choice in ['Dichiarazione prossimi congiunti','Dichiarazione parenti per cremazione – I-05-1'] and len(people)>2:
    candidates=[i for i in range(len(people)) if i not in {ai,di}]
    relative_indices=st.multiselect('Parenti da inserire nei riquadri successivi (facoltativo)',candidates,format_func=lambda i:names[i],key='relative_indices')

st.subheader('3. Dati specifici della pratica')
c1,c2,c3=st.columns(3)
f['qualita']=c1.text_input('Qualità del richiedente (es. FIGLIA)',f['qualita']); f['grado_parentela']=c2.text_input('Grado di parentela',f['grado_parentela']); f['data_morte']=c3.text_input('Data di morte',f['data_morte'])
c1,c2,c3=st.columns(3); f['luogo_morte']=c1.text_input('Luogo di morte',f['luogo_morte']); f['destinazione']=c2.text_input('Destinazione / Cimitero',f['destinazione']); f['targa']=c3.text_input('Targa',f['targa'])
c1,c2,c3=st.columns(3); f['data_servizio']=c1.text_input('Data pratica/trasporto',f['data_servizio']); f['ora']=c2.text_input('Ora',f['ora']); f['tipo_materiale']=c3.selectbox('Oggetto', ['salma','resti','ceneri'], index=0)
c1,c2,c3=st.columns(3); f['cimitero']=c1.text_input('Cimitero',f['cimitero']); f['sezione']=c2.text_input('Sezione',f['sezione']); f['lotto']=c3.text_input('Lotto',f['lotto'])
c1,c2=st.columns(2); f['fascicolo']=c1.text_input('Fascicolo n°',f['fascicolo']); f['tipo_operazione']=c2.text_input('Operazione cimiteriale',f['tipo_operazione'])

if choice=='Richiesta autorizzazione trasporti':
    st.info('Il modello di trasporto dispone di una mappatura dedicata. La targa e le caselle di scelta del veicolo vanno completate o verificate manualmente.')
elif choice=='Richiesta concessione loculo/nicchia':
    c1,c2=st.columns(2); f['tipo_sepoltura']=c1.selectbox('Preferenza', ['loculo ipogeo in sepoltura','nicchia murale fuori terra']); f['preferenza_cimitero']=c2.selectbox('Preferenza cimitero',['1 – Santa Maria dei Rotoli','2 – Cappuccini','3 – Santa Maria di Gesù'])
elif choice == 'Dichiarazione parenti per cremazione – I-05-1':
    f['numero_parenti']=st.text_input('Numero di parenti dello stesso grado',f['numero_parenti'])
elif choice in ['Richiesta cremazione – I-04','Consenso cremazione – I-05']:
    c1,c2=st.columns(2); f['scelta_ceneri']=c1.text_input('Destinazione/collocazione ceneri',f['scelta_ceneri']); f['dispersione']=c2.text_input('Eventuale dispersione',f['dispersione'])
elif 'tumulazione' in choice.lower():
    c1,c2=st.columns(2); f['concessionario']=c1.text_input('Concessionario / intestatario sepoltura',f['concessionario']); f['grado_concessionario']=c2.text_input('Grado di parentela con concessionario',f['grado_concessionario'])

if choice == 'Richiesta concessione loculo/nicchia':
    st.warning('Nel modulo loculo/nicchia il riquadro anagrafico del richiedente e la scelta loculo/nicchia restano manuali. Nel riquadro del defunto vengono inseriti cognome, nome, data di nascita, luogo e data di morte; viene indicata la prima preferenza del cimitero. Completa e verifica manualmente codice fiscale, sesso, nascita/residenza e tutti gli altri campi non compilati.')
elif choice == 'Autorizzazione tumulazione salma/resti – I-15':
    st.warning('Nel modulo I-15 il richiedente viene compilato parzialmente. Restano manuali il codice fiscale, il sesso, le date, concessionario/beneficiario e tutti i dati della sepoltura (sezione, lotto, cimitero e fascicolo), perché le celle unite possono spostare il testo. Verifica ogni campo prima dell’uso.')
elif choice in ['Richiesta operazioni cimiteriali – I-01','Richiesta cremazione – I-04']:
    st.warning('In questo modello alcuni campi del richiedente restano manuali: codice fiscale, sesso, numero civico, CAP e telefono. Controlla anche i dati del defunto, le date e ogni casella prima di utilizzare il documento.')
elif choice == 'Consenso cremazione – I-05':
    st.warning('Nel consenso cremazione il riquadro del richiedente resta da compilare manualmente; controlla anche i dati del defunto.')
elif choice == 'Dichiarazione prossimi congiunti':
    st.warning('Nel riquadro del defunto vengono compilati alcuni dati anagrafici e di sepoltura. Completa/verifica manualmente codice fiscale, sesso, qualità e grado di parentela e controlla ogni campo prima dell’uso.')
st.warning('Il riconoscimento OCR può sbagliare. Controlla ogni pagina, le caselle e tutti i campi specifici prima di firmare o presentare la pratica. Il sistema non firma né invia automaticamente la pratica.')
with st.expander('Controlli prima della generazione', expanded=True):
    st.write('Verifica i dati trascritti dalle carte d’identità, la scelta del richiedente e del defunto, le date e le opzioni specifiche della pratica.')
    conferma_dati = st.checkbox('Confermo di aver controllato i dati inseriti e di verificare il documento generato prima di utilizzarlo.', key='conferma_dati')

def data_valida(value):
    value=(value or '').strip()
    if not value: return True
    try: datetime.strptime(value, '%d/%m/%Y'); return True
    except ValueError: return False

if st.button('📝 GENERA DOCUMENTO DOCX + PDF',type='primary'):
    if not conferma_dati:
        st.error('Prima di generare il documento, conferma di aver controllato i dati.')
        st.stop()
    date_err=[]
    for label,value in [('Data di morte',f.get('data_morte','')),('Data pratica/trasporto',f.get('data_servizio',''))]:
        if not data_valida(value): date_err.append(label)
    for i,p in enumerate(people):
        if not data_valida(p.get('data_nascita','')): date_err.append(f'Data di nascita della persona {i+1}')
    if date_err:
        st.error('Formato data non valido (usa GG/MM/AAAA): '+', '.join(date_err))
        st.stop()
    src=T/MODULES[choice]
    if not src.exists(): st.error('Template ufficiale non trovato nel pacchetto.'); st.stop()
    doc=Document(src)
    try:
        if choice=='Richiesta autorizzazione trasporti': fill_transport(doc,app,dec,f)
        else: fill_generic_identity(doc,app,dec,choice,f,[people[i] for i in relative_indices])
        # Conservative additional fields by searching visible labels and writing only into clearly blank adjacent cells.
        # This avoids destroying the official wording/layout.
        # Keep identity documents and generated forms out of a persistent app/output folder.
        # A unique temporary directory is created for this generation and removed afterwards.
        with tempfile.TemporaryDirectory(prefix='moduli_cimiteriali_') as temp_dir:
            out=save_doc(doc,f"{choice}_{dec.get('cognome','')}_{date.today().strftime('%Y%m%d')}",out_dir=temp_dir)
            pdf=pdf_from_docx(out,out_dir=temp_dir)
            docx_bytes=out.read_bytes()
            pdf_bytes=pdf.read_bytes() if pdf else None
            docx_name=out.name
            pdf_name=pdf.name if pdf else ''
        st.success('Documento generato. Controllalo prima della firma e dell’invio.')
        st.download_button('⬇️ Scarica DOCX',docx_bytes,file_name=docx_name,mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        if pdf_bytes: st.download_button('⬇️ Scarica PDF',pdf_bytes,file_name=pdf_name,mime='application/pdf')
        else: st.info('PDF non generato: installa LibreOffice per abilitarne la conversione automatica.')
    except Exception as e:
        st.exception(e)

st.divider(); st.caption('Versione di lavoro v54 — test automatici dei modelli superati; collaudo end-to-end Streamlit ancora da completare. Interfaccia non ancora collaudata end-to-end e non pronta per uso amministrativo reale. I modelli originali sono conservati come template. I documenti generati vengono gestiti in una cartella temporanea per la richiesta; controllare integralmente il documento prima dell’uso.')
