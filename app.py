from pathlib import Path
from io import BytesIO
from calendar import month_name
import re, math, shutil, subprocess, tempfile
from decimal import Decimal, ROUND_HALF_UP
import openpyxl as x
import xlrd

MONTHS=list(month_name)[4:]+list(month_name)[1:4]
TAX=['IGST','CGST','SGST','CESS']
def source_cells(data, name):
    if name.lower().endswith('.xls'):
        sh=xlrd.open_workbook(file_contents=data).sheet_by_name('Offset Summary')
        get=lambda r,c: sh.cell_value(r-1,c-1) if r<=sh.nrows and c<=sh.ncols else None
    else:
        sh=x.load_workbook(BytesIO(data),data_only=True)['Offset Summary']
        formulas=x.load_workbook(BytesIO(data),data_only=False)['Offset Summary']
        def get(r,c):
            cell = sh.cell(r,c)
            if cell.value is None and formulas.cell(r,c).data_type == 'f':
                raise ValueError(f'Source formula {cell.coordinate} has no saved value. Recalculate and save the source before uploading.')
            return cell.value
    return get

def fill(data, name):
    get=source_cells(data, name)
    if str(get(1,2)).strip().lower()!='consolidated 3b offset summary':
        raise ValueError('Expected a Consolidated 3B Offset Summary export.')
    def check(r,c,labels):
        if [str(get(r,c+i)).strip().lower() for i in range(len(labels))]!=[v.lower() for v in labels]:
            raise ValueError(f'Unexpected source layout at row {r}, column {c}.')
    check(6,3,['Outward (NRC)','Outward (RC)','Non-Taxable','Total']+TAX)
    check(5,7,['Liability on Outward']); check(22,7,['Liability on Reverse Charge']); check(23,7,TAX)
    for c,label in zip([4,8,12,16,20],['ITC Eligible (NRC)','ITC Eligible (RC)','ITC Reversed','ITC Ineligible','ITC Reclaimed']):
        check(39,c,[label]); check(40,c,TAX)
    def months(start):
        d={str(get(r,2)).strip().lower():r for r in range(start,start+12)}
        if set(d)!={m.lower() for m in MONTHS}: raise ValueError(f'Expected 12 unique month rows starting at row {start}.')
        return d
    out,rc,itc=map(months,[7,24,41])
    def number(r,c):
        v=get(r,c)
        if v is None or v=='': return None
        if type(v) in (int,float) and math.isfinite(v): return v
        raise ValueError(f'Expected a numeric amount or blank at source row {r}, column {c}.')
    wb=x.load_workbook(Path(__file__).with_name('template.xlsx')); s=wb['Offset Summary']
    for r in (2,3):
        v=get(r,2)
        if not isinstance(v,str) or not v.strip() or v.lstrip().startswith('='): raise ValueError('Missing or invalid company / financial year header.')
        s.cell(r,2).value=v.strip()
    def copy(row,col,sr,sc,n):
        for j in range(n): s.cell(row,col+j).value=number(sr,sc+j)
    for i,m in enumerate(MONTHS):
        k=m.lower(); r=7+i
        copy(r,3,out[k],3,1); copy(r,4,out[k],5,1); copy(r,6,out[k],7,4)
        s[f'E{r}']=f'=IF(COUNT(C{r}:D{r})=0,"",SUM(C{r}:D{r}))'
        s[f'S{r}']=f'=SUM(E{r})-SUM(L{r})'
        copy(25+i,4,itc[k],4,4); copy(42+i,10,rc[k],7,4)
        s.cell(42+i,9).value=m
        copy(60+i,10,itc[k],8,4); copy(60+i,4,itc[k],12,4)
        copy(78+i,7,itc[k],16,4); copy(78+i,2,itc[k],20,4)
        s.cell(78+i,1).value=m
        for d in range(4):
            a,b,z=[x.utils.get_column_letter(c+d) for c in [10,14,18]]
            s[f'{z}{60+i}']=f'={a}{60+i}-{b}{60+i}'
    s['B54']='Grand Total'; s['A90']='Grand Total'
    for c in range(10,22):
        l=x.utils.get_column_letter(c); s.cell(72,c).value=f'=SUM({l}60:{l}71)'
    for c in range(18,22): s.cell(59,c).value=None
    s['E19']='=SUM(E7:E18)'; s['P19']='=SUM(P7:P18)'; s['H37']='=SUM(H25:H36)'
    notes=wb.create_sheet('Import Notes'); notes.column_dimensions['A'].width=115
    for line in ['GSTR-3B Offset Summary import', 'Source: '+name, 'Matched months by name (April–March). Source blanks remain blank in imported input cells.', 'Outward total = Outward (NRC) + Non-Taxable; Outward (RC) is excluded as indicated by the REECO heading.', 'As per Books is not imported. Enter books figures before interpreting differences.', 'Totals and differences may show zero or treat blanks as zero; this is not proof that the underlying amounts were reported.', 'Imported: outward values and liability, reverse charge liability, ITC eligible (NRC/RC), reversed, reclaimed and ineligible.']:
        notes.append([line]); notes.cell(notes.max_row,1).data_type='s'
    buf=BytesIO(); wb.save(buf); return buf.getvalue()

def recalculate(data):
    executable = shutil.which('libreoffice') or shutil.which('soffice')
    if not executable:
        raise ValueError('Verification unavailable: install LibreOffice on the app server.')
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        source = root / 'reeco.xlsx'
        source.write_bytes(data)
        output = root / 'output'
        output.mkdir()
        completed = subprocess.run([
            executable, '-env:UserInstallation=' + (root / 'profile').as_uri(),
            '--headless', '--convert-to', 'xlsx', '--outdir', str(output), str(source)
        ], capture_output=True, timeout=90)
        result = output / source.name
        if completed.returncode or not result.is_file():
            raise ValueError('Workbook recalculation failed; download blocked.')
        return result.read_bytes()

def verify(source, name, result):
    get = source_cells(source, name)
    book = x.load_workbook(BytesIO(result), data_only=True)
    try:
        s = book['Offset Summary']
        client, year = (s.cell(r, 2).value for r in (2, 3))
        if (client, year) != tuple(get(r, 2).strip() for r in (2, 3)):
            raise ValueError('Client or financial-year header differs from source.')
        gstin = re.search(r'\b\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]\b', client.upper())
        gstin = gstin.group() if gstin else None
        rows = []
        def amount(v):
            if type(v) not in (int, float) or not math.isfinite(v):
                raise ValueError('Expected a calculated numeric amount; check formulas and source data.')
            return Decimal(str(v))
        def display_amount(v):
            return f'{v.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP):,.2f}'
        def equal(a,b):
            return abs(a-b) <= Decimal('0.0000001')
        def metric(label, start, cols, dest, total):
            lookup = {get(r, 2).strip().casefold(): r for r in range(start, start+12)}
            expected = Decimal(0)
            reported = 0
            for i, month in enumerate(MONTHS):
                month_col = {19:'B', 37:'B', 54:'I', 72:'B', 90:'A'}[total]
                if s[f'{month_col}{total-12+i}'].value != month:
                    raise ValueError(f'{label}: unexpected REECO month label for {month}.')
                values = [get(lookup[month.casefold()], c) for c in cols]
                present = [amount(v) for v in values if v not in (None, '')]
                actual = s[f'{dest}{total-12+i}'].value
                if present:
                    reported += 1
                    want = sum(present)
                    if not equal(amount(actual), want):
                        raise ValueError(f'{label}: mismatch for {month}.')
                    expected += want
                elif actual not in (None, ''):
                    raise ValueError(f'{label}: expected an unreported amount for {month}.')
            actual = amount(s[f'{dest}{total}'].value)
            if not equal(actual, expected):
                raise ValueError(f'{label}: annual total mismatch.')
            rows.append({'Metric': label, 'Source (sum of months)': display_amount(expected) if reported else 'Not reported',
                         'REECO total': display_amount(actual), 'Difference': display_amount(actual-expected) if reported else '—',
                         'Months reported': reported, 'REECO cell': f'{dest}{total}',
                         'Status': 'Match' if reported else 'No source amounts'})
        for label, cols, dest in [('Outward (NRC)', [3], 'C'), ('Non-Taxable', [5], 'D'), ('Outward (NRC + Non-Taxable)', [3,5], 'E')]:
            metric(label, 7, cols, dest, 19)
        for label, start, col, dests, total in [
            ('Outward liability',7,7,'FGHI',19), ('Reverse-charge liability',24,7,'JKLM',54),
            ('ITC Eligible (NRC)',41,4,'DEFG',37), ('ITC Eligible (RC)',41,8,'JKLM',72),
            ('ITC Reversed',41,12,'DEFG',72), ('ITC Ineligible',41,16,'GHIJ',90),
            ('ITC Reclaimed',41,20,'BCDE',90)]:
            for i, tax in enumerate(TAX):
                metric(f'{label} — {tax}', start, [col+i], dests[i], total)
        return {'client': client, 'year': year, 'gstin': gstin, 'source': name, 'rows': rows}
    finally:
        book.close()

def fill_batch(inputs):
    files, errors, used = [], [], set()
    for source, data in inputs:
        try:
            result = recalculate(fill(data, source))
            preview = verify(data, source, result)
            client = preview['client']
            client = re.sub(r"\s*\([0-9A-Z]{15}\)\s*$", "", client, flags=re.I)
            client = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', client)
            base = 'REECO - ' + (client[:120].strip(' .') or 'Client')
            name, n = base + '.xlsx', 2
            while name.casefold() in used:
                name = f'{base} ({n}).xlsx'
                n += 1
            used.add(name.casefold())
            files.append((name, result, preview))
        except Exception as exc:
            errors.append((source, str(exc)))
    return files, errors

if __name__=='__main__':
    import hashlib
    import hmac
    import streamlit as st
    st.title('GST REECO Sheet Automation')
    try:
        protection = st.secrets.get('PASSWORD_PROTECTION', True)
    except FileNotFoundError:
        protection = True
    if type(protection) is not bool:
        st.error('Set PASSWORD_PROTECTION to true or false in Streamlit Secrets, without quotes.')
        st.stop()
    if not protection:
        st.session_state.pop('login_token', None)
    if protection:
        try:
            password = st.secrets['APP_PASSWORD']
        except (FileNotFoundError, KeyError):
            password = None
        if not isinstance(password, str) or len(password.strip()) < 16 or password == 'REPLACE_WITH_YOUR_PASSWORD':
            st.error('The app is locked. Set APP_PASSWORD in Streamlit Secrets to a unique password of at least 16 characters.')
            st.stop()
        token = hashlib.sha256(password.encode()).digest()
        def sign_in():
            supplied = st.session_state.pop('login_password', '')
            if hmac.compare_digest(supplied.encode(), password.encode()):
                st.session_state['login_token'] = token
                st.session_state.pop('login_failed', None)
            else:
                st.session_state['login_failed'] = True
        if st.session_state.get('login_token') != token:
            with st.form('sign_in'):
                st.text_input('App password', type='password', key='login_password')
                st.form_submit_button('Sign in', on_click=sign_in)
            if st.session_state.get('login_failed'):
                st.error('Incorrect password.')
            st.stop()
        if st.sidebar.button('Sign out'):
            st.session_state.clear()
            st.rerun()
    from zipfile import ZipFile, ZIP_DEFLATED
    st.write('Upload one or more Consolidated 3B Offset Summaries. Each input produces its own REECO workbook.')
    uploaded = st.file_uploader('3B Offset Summaries', type=['xls', 'xlsx'], accept_multiple_files=True)
    if uploaded:
        files, errors = fill_batch([(f.name, f.getvalue()) for f in uploaded])
        for name, error in errors:
            st.error(f'{name}: {error}')
        if files:
            st.success(f'{len(files)} of {len(uploaded)} workbooks passed source-to-REECO import checks. Books figures are not verified.')
            identities = [(p['gstin'], p['year'].casefold()) for _, _, p in files if p['gstin']]
            if len(set(identities)) != len(identities):
                st.warning('Duplicate GSTIN and financial-year combination in this batch. Review the source filenames before using either workbook.')
            if len(uploaded) > 1:
                archive = BytesIO()
                with ZipFile(archive, 'w', ZIP_DEFLATED) as bundle:
                    for name, result, _ in files:
                        bundle.writestr(name, result)
                st.download_button('Download all successful workbooks (ZIP)', archive.getvalue(), 'REECO Batch.zip', 'application/zip', on_click='ignore')
            for i, (name, result, preview) in enumerate(files):
                st.download_button(f'Download {name}', result, name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key=f'download_{i}', on_click='ignore')
                st.subheader(preview['client'])
                st.text(f"GSTIN: {preview['gstin'] or 'Not detected'} | Financial year: {preview['year']}\nSource: {preview['source']}")
                if not preview['gstin']:
                    st.warning('GSTIN could not be identified in the source header. Confirm the client manually.')
                st.caption('REECO totals come from the exact workbook above. Source totals are calculated from comparable monthly values, not copied from the export’s printed grand-total row. Amounts are displayed to two decimals.')
                st.dataframe(preview['rows'], hide_index=True)
