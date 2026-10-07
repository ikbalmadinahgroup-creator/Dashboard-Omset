"""Pastikan gaya (cellXfs) yang dipakai script rekap ada di workbook.
Excel me-renumber style saat user menyimpan ulang file -> id gaya hardcode di script
(mis. 245, 250) bisa melebihi jumlah cellXfs dan Excel akan 'repair' file.
ensure_styles() mencari gaya yang identik di workbook, kalau tidak ada ditambahkan,
lalu mengembalikan peta id lama -> id baru."""
import re, json, os
REF=os.path.join(os.path.dirname(__file__),'style_ref.json')

def _parts(st):
    def blk(tag):
        m=re.search(rf'<{tag}[^>]*?(?:/>|>(.*?)</{tag}>)',st,re.S)
        return m
    nf={m.group(1):m.group(2) for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\s*/>',st)}
    def items(tag,item):
        m=re.search(rf'<{tag}\b[^>]*>(.*?)</{tag}>',st,re.S)
        if not m: return []
        return re.findall(rf'<{item}\b[^>]*?/>|<{item}\b[^>]*>.*?</{item}>',m.group(1),re.S)
    return nf,items('fonts','font'),items('fills','fill'),items('borders','border'),items('cellXfs','xf')

def _attr(x,a,d='0'):
    m=re.search(rf'\b{a}="([^"]*)"',x); return m.group(1) if m else d

def resolve(st):
    nf,fonts,fills,borders,xfs=_parts(st)
    out=[]
    for x in xfs:
        n=_attr(x,'numFmtId'); inner=re.sub(r'^<xf\b[^>]*?/?>','',x).replace('</xf>','')
        out.append({'numFmtId':n,'numFmt':nf.get(n),'font':fonts[int(_attr(x,'fontId'))],
                    'fill':fills[int(_attr(x,'fillId'))],'border':borders[int(_attr(x,'borderId'))],'inner':inner})
    return out

def _key(d): return json.dumps([d['numFmt'] or d['numFmtId'],d['font'],d['fill'],d['border'],d['inner']])

def ensure_styles(st, ids):
    """st = isi styles.xml workbook tujuan. Return (st_baru, {id_lama: id_baru})."""
    ref=json.load(open(REF))
    cur=resolve(st); idx={}
    for i,d in enumerate(cur): idx.setdefault(_key(d),i)
    mp={}
    for i in ids:
        d=ref.get(str(i))
        if d is None: continue
        k=_key(d)
        if k in idx: mp[i]=idx[k]; continue
        # tambah komponen
        def add(tag,item,xml):
            nonlocal st
            m=re.search(rf'<{tag} count="(\d+)"[^>]*>',st); n=int(m.group(1))
            st=st.replace(m.group(0),m.group(0).replace(f'count="{n}"',f'count="{n+1}"'),1)
            j=st.index(f'</{tag}>'); st=st[:j]+xml+st[j:]; return n
        fid=add('fonts','font',d['font']); flid=add('fills','fill',d['fill']); bid=add('borders','border',d['border'])
        num=d['numFmtId']
        if d['numFmt'] is not None:
            nf,_,_,_,_=_parts(st)
            hit=[k2 for k2,v in nf.items() if v==d['numFmt']]
            if hit: num=hit[0]
            else:
                ids_=[int(k2) for k2 in nf]; num=str(max(ids_+[199])+1)
                code=d['numFmt']
                if '<numFmts' in st:
                    m=re.search(r'<numFmts count="(\d+)">',st); n=int(m.group(1))
                    st=st.replace(m.group(0),f'<numFmts count="{n+1}">',1)
                    st=st.replace('</numFmts>',f'<numFmt numFmtId="{num}" formatCode="{code}"/></numFmts>',1)
                else:
                    st=st.replace('<fonts',f'<numFmts count="1"><numFmt numFmtId="{num}" formatCode="{code}"/></numFmts><fonts',1)
        xf=(f'<xf numFmtId="{num}" fontId="{fid}" fillId="{flid}" borderId="{bid}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"'
            +(f' applyAlignment="1">{d["inner"]}</xf>' if d['inner'] else '/>'))
        newid=add('cellXfs','xf',xf)
        cur.append(d); idx[k]=newid; mp[i]=newid
    return st, mp

def remap_sheet(x, mp, cells_pred=None):
    """Ganti s="lama" -> s="baru" pada <c> (semua sel di sheet hasil generate)."""
    def rep(m):
        s=int(m.group(2)); return f'{m.group(1)}s="{mp.get(s,s)}"'
    return re.sub(r'(<c r="[A-Z]+\d+"[^>]*?\s)s="(\d+)"',rep,x)
