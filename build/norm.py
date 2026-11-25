import re
SUBS=[(r'\(E\)|\bEAST\b|-E\b',' EAST '),(r'\(W\)|\bWEST\b|-W\b',' WEST '),(r'\(N\)|\bNORTH\b',' NORTH '),(r'\(S\)|\bSOUTH\b',' SOUTH '),
(r'\(U\)|\bURBAN\b|-U\b|\bU\b',' URBAN '),(r'\(R\)|\bRURAL\b|-R\b|\bR\b',' RURAL '),(r'\bNO\.\s*|\bNO\s+(?=\d)',' '),
(r'\b(III|3)\b',' 3 '),(r'\b(II|2)\b',' 2 '),(r'\b(I|1)\b',' 1 '),(r'\bJI\b',' '),(r'\bESD\b|\bESS\b|\bESC\b|\bDIVISION\b|\bCIRCLE\b|\bSECTION\b|\bSUB ?DIVISION\b',' ')]
PH=[('SHARIFF','SARIF'),('SHARIF','SARIF'),('ARRAH','ARA'),('AARA','ARA'),('PURNIA','PURNEA'),('CHHAPRA','CHAPRA'),('BHABHUA','BHABUA'),
('KOSHI','KOSI'),('BAKHTIYARPUR','BAKHTIARPUR'),('DEHRI','DEHARI'),('JAGDEESHPUR','JAGDISHPUR'),('PIRPAINTI','PIRPAINTY')]
PREFIX=r'^(ESSD|ESS|ESD|ESC|SEC|SECTION|SUBDIVISION|SUB DIVISION|DIVISION|CIRCLE|PSS|33/11 ?KV|33/11KV|33 ?KV)[\s\-:.]+'
def key(s):
    s=str(s).upper().strip().replace('_',' ')
    for _ in range(2): s=re.sub(PREFIX,'',s).strip()
    s=re.sub(r'([\s-]NEW|\sN)$','',s).strip()
    s=s.replace('&',' AND ')
    for a,b in SUBS: s=re.sub(a,b,s)
    s=re.sub(r'[^A-Z0-9 ]',' ',s)
    toks=s.split()
    s=' '.join(toks)
    for a,b in PH: s=s.replace(a,b)
    # phonetic squash: remove H after consonant, collapse doubles, EE->I, OO->U, Y->I
    s=s.replace('EE','I').replace('OO','U').replace('Y','I')
    s=re.sub(r'([BCDGJKPST])H',r'\1',s)
    s=re.sub(r'(.)\1+',r'\1',s)
    s=s.replace('AA','A')
    return s.replace(' ','')
