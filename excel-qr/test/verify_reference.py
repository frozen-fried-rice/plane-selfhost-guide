# Python mirror of the VBA QR algorithm (Nayuki-style, byte mode).
# Used only to cross-check correctness against segno. The VBA code mirrors this exactly.

ECCB = {
0:[-1,7,10,15,20,26,18,20,24,30,18,20,24,26,30,22,24,28,30,28,28,28,28,30,30,26,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30],
1:[-1,10,16,26,18,24,16,18,22,22,26,30,22,22,24,24,28,28,26,26,26,26,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28,28],
2:[-1,13,22,18,26,18,24,18,22,20,24,28,26,24,20,30,24,28,28,26,30,28,30,30,30,30,28,30,30,30,30,30,30,30,30,30,30,30,30,30,30],
3:[-1,17,28,22,16,22,28,26,26,24,28,24,28,22,24,24,30,28,28,26,28,30,24,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30,30],
}
NEB = {
0:[-1,1,1,1,1,1,2,2,2,2,4,4,4,4,4,6,6,6,6,7,8,8,9,9,10,12,12,12,13,14,15,16,17,18,19,19,20,21,22,24,25],
1:[-1,1,1,1,2,2,4,4,4,5,5,5,8,9,9,10,10,11,13,14,16,17,17,18,20,21,23,25,26,28,29,31,33,35,37,38,40,43,45,47,49],
2:[-1,1,1,2,2,4,4,6,6,8,8,8,10,12,16,12,17,16,18,21,20,23,23,25,27,29,34,34,35,38,40,43,45,48,51,53,56,59,62,65,68],
3:[-1,1,1,2,4,4,4,5,6,8,8,11,11,16,16,18,16,19,21,25,25,25,34,30,32,35,37,40,42,45,48,51,54,57,60,63,66,70,74,77,81],
}

def gfmul(x,y):
    z=0
    for i in range(7,-1,-1):
        z=z*2
        if z & 0x100: z^=0x11D
        if (y>>i)&1: z^=x
    return z & 0xFF

def rs_divisor(deg):
    result=[0]*deg; result[deg-1]=1; root=1
    for i in range(deg):
        for j in range(deg):
            result[j]=gfmul(result[j],root)
            if j+1<deg: result[j]^=result[j+1]
        root=gfmul(root,2)
    return result

def rs_rem(data,divisor,deg):
    result=[0]*deg
    for b in data:
        factor=(b^result[0])&0xFF
        result=result[1:]+[0]
        for i in range(deg):
            result[i]^=gfmul(divisor[i],factor)
    return result

def raw_modules(ver):
    size=ver*4+17
    r=size*size-8*8*3-(15*2+1)-(size-16)*2
    if ver>=2:
        na=ver//7+2
        r-=(na-1)*(na-1)*25
        r-=(na-2)*2*20
        if ver>=7: r-=36
    return r

def num_data_cw(ver,ecl):
    return raw_modules(ver)//8 - NEB[ecl][ver]*ECCB[ecl][ver]

def cc_bits(ver): return 8 if ver<=9 else 16

def utf8(s):
    return list(s.encode('utf-8'))

def build_data_cw(byts,ver,ecl):
    ndc=num_data_cw(ver,ecl); cap=ndc*8
    bits=[]
    def app(val,nb):
        for i in range(nb-1,-1,-1): bits.append((val>>i)&1)
    app(4,4)
    app(len(byts),cc_bits(ver))
    for b in byts: app(b,8)
    term=min(4,cap-len(bits)); app(0,term)
    app(0,(8-len(bits)%8)%8)
    pad=0xEC
    while len(bits)<cap:
        app(pad,8); pad=0x11 if pad==0xEC else 0xEC
    cw=[]
    for i in range(ndc):
        v=0
        for b in range(8): v=v*2+bits[i*8+b]
        cw.append(v)
    return cw

def add_ecc_interleave(data_cw,ver,ecl):
    nb=NEB[ecl][ver]; be=ECCB[ecl][ver]
    raw=raw_modules(ver)//8
    nshort=nb-(raw%nb); slen=raw//nb
    div=rs_divisor(be)
    blocks=[[0]*(slen+1) for _ in range(nb)]
    k=0
    for j in range(nb):
        dl=slen-be+(0 if j<nshort else 1)
        dat=data_cw[k:k+dl]; k+=dl
        for i in range(dl): blocks[j][i]=dat[i]
        ecc=rs_rem(dat,div,be)
        es=(slen+1)-be
        for i in range(be): blocks[j][es+i]=ecc[i]
    res=[]
    for i in range(slen+1):
        for j in range(nb):
            if not (i==(slen-be) and j<nshort):
                res.append(blocks[j][i])
    return res

def align_positions(ver):
    if ver==1: return []
    na=ver//7+2
    step=26 if ver==32 else ((ver*4+na*2+1)//(na*2-2))*2
    pos=[0]*na; pos[0]=6; p=ver*4+10
    for i in range(na-1,0,-1):
        pos[i]=p; p-=step
    return pos

class M:
    def __init__(s,ver):
        s.size=ver*4+17; s.ver=ver
        s.m=[[False]*s.size for _ in range(s.size)]
        s.f=[[False]*s.size for _ in range(s.size)]
    def setf(s,x,y,d):
        s.m[y][x]=d; s.f[y][x]=True

def draw_finder(mm,cx,cy):
    for dy in range(-4,5):
        for dx in range(-4,5):
            dist=max(abs(dx),abs(dy)); xx=cx+dx; yy=cy+dy
            if 0<=xx<mm.size and 0<=yy<mm.size:
                mm.setf(xx,yy,dist!=2 and dist!=4)

def draw_align(mm,cx,cy):
    for dy in range(-2,3):
        for dx in range(-2,3):
            mm.setf(cx+dx,cy+dy,max(abs(dx),abs(dy))!=1)

def get_bit(x,i): return (x>>i)&1!=0

def draw_format(mm,ecl,msk):
    fb={0:1,1:0,2:3,3:2}[ecl]
    data=fb*8+msk; rem=data
    for _ in range(10):
        hi=(rem>>9)&1; rem=(rem*2)^(hi*0x537)
    bits=((data<<10)|rem)^0x5412
    for i in range(6): mm.setf(8,i,get_bit(bits,i))
    mm.setf(8,7,get_bit(bits,6)); mm.setf(8,8,get_bit(bits,7)); mm.setf(7,8,get_bit(bits,8))
    for i in range(9,15): mm.setf(14-i,8,get_bit(bits,i))
    for i in range(8): mm.setf(mm.size-1-i,8,get_bit(bits,i))
    for i in range(8,15): mm.setf(8,mm.size-15+i,get_bit(bits,i))
    mm.setf(8,mm.size-8,True)

def draw_version(mm):
    if mm.ver<7: return
    rem=mm.ver
    for _ in range(12):
        hi=(rem>>11)&1; rem=(rem*2)^(hi*0x1F25)
    bits=(mm.ver<<12)|rem
    for i in range(18):
        bit=get_bit(bits,i); a=mm.size-11+i%3; b=i//3
        mm.setf(a,b,bit); mm.setf(b,a,bit)

def draw_funcs(mm,ecl):
    for i in range(mm.size):
        mm.setf(6,i,i%2==0); mm.setf(i,6,i%2==0)
    draw_finder(mm,3,3); draw_finder(mm,mm.size-4,3); draw_finder(mm,3,mm.size-4)
    pos=align_positions(mm.ver); n=len(pos)
    for a in range(n):
        for b in range(n):
            if not ((a==0 and b==0) or (a==0 and b==n-1) or (a==n-1 and b==0)):
                draw_align(mm,pos[a],pos[b])
    draw_format(mm,ecl,0); draw_version(mm)

def draw_codewords(mm,data):
    i=0; total=len(data)*8; right=mm.size-1
    while right>=1:
        if right==6: right=5
        for vert in range(mm.size):
            for j in range(2):
                x=right-j; upward=((right+1)&2)==0
                y=(mm.size-1-vert) if upward else vert
                if (not mm.f[y][x]) and i<total:
                    mm.m[y][x]=get_bit(data[i>>3],7-(i&7)); i+=1
        right-=2

def apply_mask(mm,msk):
    for y in range(mm.size):
        for x in range(mm.size):
            if mm.f[y][x]: continue
            if msk==0: inv=(x+y)%2==0
            elif msk==1: inv=y%2==0
            elif msk==2: inv=x%3==0
            elif msk==3: inv=(x+y)%3==0
            elif msk==4: inv=((x//3)+(y//2))%2==0
            elif msk==5: inv=(x*y%2)+(x*y%3)==0
            elif msk==6: inv=((x*y%2)+(x*y%3))%2==0
            else: inv=(((x+y)%2)+(x*y%3))%2==0
            if inv: mm.m[y][x]=not mm.m[y][x]

def finder_add(rh,cur,size):
    if rh[0]==0: cur+=size
    rh[:] = [cur]+rh[:-1]
def finder_count(rh):
    n=rh[1]; core=n>0 and rh[2]==n and rh[3]==n*3 and rh[4]==n and rh[5]==n
    c=0
    if core and rh[0]>=n*4 and rh[6]>=n: c+=1
    if core and rh[6]>=n*4 and rh[0]>=n: c+=1
    return c
def finder_term(color,length,rh,size):
    if color: finder_add(rh,length,size); length=0
    length+=size; finder_add(rh,length,size); return finder_count(rh)

def penalty(mm):
    N1,N2,N3,N4=3,3,40,10; size=mm.size; res=0
    for y in range(size):
        rc=False; rx=0; rh=[0]*7
        for x in range(size):
            if mm.m[y][x]==rc:
                rx+=1
                if rx==5: res+=N1
                elif rx>5: res+=1
            else:
                finder_add(rh,rx,size)
                if not rc: res+=finder_count(rh)*N3
                rc=mm.m[y][x]; rx=1
        res+=finder_term(rc,rx,rh,size)*N3
    for x in range(size):
        rc=False; rx=0; rh=[0]*7
        for y in range(size):
            if mm.m[y][x]==rc:
                rx+=1
                if rx==5: res+=N1
                elif rx>5: res+=1
            else:
                finder_add(rh,rx,size)
                if not rc: res+=finder_count(rh)*N3
                rc=mm.m[y][x]; rx=1
        res+=finder_term(rc,rx,rh,size)*N3
    for y in range(size-1):
        for x in range(size-1):
            c=mm.m[y][x]
            if c==mm.m[y][x+1] and c==mm.m[y+1][x] and c==mm.m[y+1][x+1]: res+=N2
    dark=sum(1 for y in range(size) for x in range(size) if mm.m[y][x])
    total=size*size
    k=(abs(dark*20-total*10)+total-1)//total-1
    res+=k*N4
    return res

def generate(text,ecl,force_mask=None):
    byts=utf8(text); ver=0
    for v in range(1,41):
        if 4+cc_bits(v)+8*len(byts)<=num_data_cw(v,ecl)*8: ver=v; break
    assert ver>0
    dcw=build_data_cw(byts,ver,ecl)
    allcw=add_ecc_interleave(dcw,ver,ecl)
    mm=M(ver); draw_funcs(mm,ecl); draw_codewords(mm,allcw)
    if force_mask is not None:
        best=force_mask
    else:
        best=0; minp=1<<30
        for msk in range(8):
            draw_format(mm,ecl,msk); apply_mask(mm,msk)
            p=penalty(mm)
            if p<minp: minp=p; best=msk
            apply_mask(mm,msk)
    draw_format(mm,ecl,best); apply_mask(mm,best)
    return mm,ver,best

if __name__=='__main__':
    import segno
    ECLNAME={0:'l',1:'m',2:'q',3:'h'}
    tests=["HELLO WORLD","https://example.com/path?id=12345","こんにちは、世界！QR",
           "ABC-1234567890-xyz","日本語テスト用の長めの文字列 1234567890 ABCDEFG"]
    allok=True
    for text in tests:
        for ecl in range(4):
            mm,ver,msk=generate(text,ecl)  # our best-mask choice
            # force segno to same version/ecl/mask, byte mode
            try:
                qr=segno.make(text.encode('utf-8'),error=ECLNAME[ecl],mask=msk,version=ver,mode='byte',boost_error=False)
            except Exception as e:
                print("segno err",e); allok=False; continue
            sm=[[bool(qr.matrix[y][x]) for x in range(mm.size)] for y in range(mm.size)]
            ok = all(mm.m[y][x]==sm[y][x] for y in range(mm.size) for x in range(mm.size))
            # also verify segno picks same optimal mask independently
            qr2=segno.make(text.encode('utf-8'),error=ECLNAME[ecl],version=ver,mode='byte',boost_error=False)
            mask_match = (qr2.mask==msk)
            print(f"ver{ver:2d} ecl{ecl} mask{msk} match={ok} segno_optmask={qr2.mask}({'same' if mask_match else 'DIFF'})  '{text[:20]}'")
            if not ok: allok=False
    print("ALL MATCH" if allok else "MISMATCH FOUND")
