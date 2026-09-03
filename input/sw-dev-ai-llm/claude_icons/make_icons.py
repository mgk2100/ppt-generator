"""Claude 삽화 세트 — 플랫 벡터 SVG 15종 (Physical AI 슬라이드 톤)"""
import os, math
OUT = os.path.join(os.path.dirname(__file__), "svg")
NAVY="#203768"; BLUE="#0E5DB7"; LIGHT="#A6CAEC"; PALE="#F3F6FC"; INK="#1F3A63"
PURPLE="#8E44AD"; PURPLE_L="#D7BDE2"; GREEN="#27AE60"; GREEN_L="#BFE8CF"; BLUE_L="#BBD6F2"
W=200
def svg(body, vb=200):
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {vb} {vb}" width="{vb}" height="{vb}" fill="none" stroke-linecap="round" stroke-linejoin="round">{body}</svg>'
def rr(x,y,w,h,r=8,fill=PALE,stroke=NAVY,sw=4,extra=""):
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" {extra}/>'
def line(x1,y1,x2,y2,stroke=NAVY,sw=4,extra=""):
    return f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{stroke}" stroke-width="{sw}" {extra}/>'
def circ(cx,cy,r,fill=BLUE,stroke="none",sw=0):
    return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'
def path(d,fill="none",stroke=NAVY,sw=4,extra=""):
    return f'<path d="{d}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" {extra}/>'
def text_lines(x,y,w,n,gap=12,color=LIGHT,sw=5,short_last=True):
    s=""
    for i in range(n):
        ww = w*0.6 if (short_last and i==n-1) else w
        s+=line(x,y+i*gap,x+ww,y+i*gap,stroke=color,sw=sw)
    return s
def arrow_right(x,y,l=40,color=BLUE,sw=5):
    return line(x,y,x+l,y,stroke=color,sw=sw)+path(f"M{x+l-12} {y-10} L{x+l} {y} L{x+l-12} {y+10}",stroke=color,sw=sw)
def arrow_up(x,y,l=40,color=BLUE,sw=5):
    return line(x,y,x,y-l,stroke=color,sw=sw)+path(f"M{x-10} {y-l+12} L{x} {y-l} L{x+10} {y-l+12}",stroke=color,sw=sw)
def sparkle(cx,cy,r=10,fill="#FFFFFF"):
    return path(f"M{cx} {cy-r} Q{cx} {cy} {cx+r} {cy} Q{cx} {cy} {cx} {cy+r} Q{cx} {cy} {cx-r} {cy} Q{cx} {cy} {cx} {cy-r} Z",fill=fill,stroke="none")
def doc(x,y,w=70,h=88,fill="#FFFFFF",lines=4,fold=16,code=False):
    d=f"M{x} {y} H{x+w-fold} L{x+w} {y+fold} V{y+h} H{x} Z"
    s=path(d,fill=fill,stroke=NAVY,sw=4)
    s+=path(f"M{x+w-fold} {y} V{y+fold} H{x+w}",stroke=NAVY,sw=3)
    if code:
        cx=x+w/2; cy=y+h/2+6
        s+=path(f"M{cx-14} {cy-14} L{cx-26} {cy} L{cx-14} {cy+14}",stroke=BLUE,sw=5)
        s+=path(f"M{cx+14} {cy-14} L{cx+26} {cy} L{cx+14} {cy+14}",stroke=BLUE,sw=5)
        s+=line(cx+5,cy-18,cx-5,cy+18,stroke=LIGHT,sw=5)
    else:
        s+=text_lines(x+12,y+30,w-24,lines,gap=13)
    return s
def cylinder(x,y,w=90,h=100,fill="#FFFFFF",stroke=NAVY):
    ry=w*0.18
    s=f'<path d="M{x} {y+ry} V{y+h-ry} A{w/2} {ry} 0 0 0 {x+w} {y+h-ry} V{y+ry}" fill="{fill}" stroke="{stroke}" stroke-width="4"/>'
    s+=f'<ellipse cx="{x+w/2}" cy="{y+ry}" rx="{w/2}" ry="{ry}" fill="{LIGHT}" stroke="{stroke}" stroke-width="4"/>'
    for k in (1,2):
        yy=y+ry+ (h-2*ry)*k/3
        s+=f'<path d="M{x} {yy} A{w/2} {ry} 0 0 0 {x+w} {yy}" stroke="{stroke}" stroke-width="3"/>'
    return s
def neural(cx,cy,layers,color,color_l,scale=1.0,node_r=9):
    """layers: list of node counts; centered at cx,cy"""
    s=""; pts=[]
    n=len(layers); dx=52*scale
    for li,cnt in enumerate(layers):
        x=cx+(li-(n-1)/2)*dx
        dy=34*scale
        col=[(x, cy+(j-(cnt-1)/2)*dy) for j in range(cnt)]
        pts.append(col)
    for li in range(n-1):
        for a in pts[li]:
            for b in pts[li+1]:
                s+=line(a[0],a[1],b[0],b[1],stroke=color_l,sw=2.2*scale)
    for col in pts:
        for (x,y) in col:
            s+=circ(x,y,node_r*scale,fill=color,stroke="#FFFFFF",sw=2.5*scale)
    return s
def person(cx,cy,scale=1.0,fill=NAVY):
    r=16*scale
    return circ(cx,cy-22*scale,r,fill=fill)+path(f"M{cx-34*scale} {cy+34*scale} Q{cx-34*scale} {cy-2*scale} {cx} {cy-2*scale} Q{cx+34*scale} {cy-2*scale} {cx+34*scale} {cy+34*scale} Z",fill=fill,stroke="none")
def bubble(x,y,w,h,fill,tail="l",stroke=NAVY):
    r=14
    d=f"M{x+r} {y} H{x+w-r} Q{x+w} {y} {x+w} {y+r} V{y+h-r} Q{x+w} {y+h} {x+w-r} {y+h} "
    if tail=="l":
        d+=f"H{x+34} L{x+18} {y+h+16} L{x+22} {y+h} H{x+r} Q{x} {y+h} {x} {y+h-r} V{y+r} Q{x} {y} {x+r} {y} Z"
    else:
        d+=f"H{x+w-22} L{x+w-18} {y+h+16} L{x+w-34} {y+h} H{x+r} Q{x} {y+h} {x} {y+h-r} V{y+r} Q{x} {y} {x+r} {y} Z"
    return path(d,fill=fill,stroke=stroke,sw=4)
def gear(cx,cy,r=22,teeth=8,fill=BLUE):
    d=""
    for i in range(teeth*2):
        ang=math.pi*2*i/(teeth*2)
        rr_=r if i%2==0 else r*0.72
        x=cx+rr_*math.cos(ang); y=cy+rr_*math.sin(ang)
        d+=("M" if i==0 else "L")+f"{x:.1f} {y:.1f} "
    d+="Z"
    return path(d,fill=fill,stroke="none")+circ(cx,cy,r*0.32,fill="#FFFFFF")
def window(x,y,w,h,dark=False,code=True):
    bg = NAVY if dark else "#FFFFFF"
    s=rr(x,y,w,h,r=8,fill=bg,stroke=NAVY,sw=4)
    s+=line(x,y+18,x+w,y+18,stroke=NAVY if not dark else "#FFFFFF",sw=3)
    for i,c in enumerate(["#E74C3C","#F1C40F","#2ECC71"]):
        s+=circ(x+12+i*11,y+9,3.5,fill=c)
    if code:
        cols=[LIGHT,BLUE,"#FFFFFF"] if dark else [LIGHT,BLUE,NAVY]
        ys=y+34
        for i,(ind,ln,c) in enumerate([(8,0.55,0),(20,0.4,1),(20,0.3,2),(8,0.5,0),(20,0.35,1)]):
            if ys+i*12>y+h-10: break
            s+=line(x+ind,ys+i*12,x+ind+(w-ind-10)*ln,ys+i*12,stroke=cols[c],sw=5)
    return s
def shield(cx,cy,w=110,h=125,fill=BLUE):
    x=cx-w/2; y=cy-h/2
    d=f"M{cx} {y} L{x+w} {y+h*0.22} V{y+h*0.55} Q{x+w} {y+h*0.85} {cx} {y+h} Q{x} {y+h*0.85} {x} {y+h*0.55} V{y+h*0.22} Z"
    return path(d,fill=fill,stroke=NAVY,sw=4)
def thumb(cx,cy,up=True,fill=BLUE):
    # 단순화된 엄지: 손바닥(둥근 사각) + 엄지(둥근 막대)
    s=rr(cx-4,cy-8,34,40,r=8,fill=fill,stroke=NAVY,sw=3.5)
    s+=path(f"M{cx-4} {cy+2} L{cx-4} {cy-20} Q{cx-2} {cy-36} {cx+10} {cy-32} L{cx+14} {cy-8}",fill=fill,stroke=NAVY,sw=3.5)
    s+=rr(cx-24,cy-2,18,34,r=6,fill="#FFFFFF",stroke=NAVY,sw=3.5)
    g=f'<g transform="rotate(180 {cx+4} {cy+8})">' if not up else "<g>"
    return g+s+"</g>"

icons={}
# 01 설계자 ↔ AI 대화
icons["01_designer_chat"]=svg(
    person(46,130,1.0)+
    bubble(74,24,110,54,PALE,tail="l")+text_lines(90,42,78,2,gap=14,color=LIGHT,sw=6)+
    bubble(60,100,120,54,BLUE,tail="r",stroke=BLUE)+
    circ(92,127,6,fill="#FFFFFF")+circ(118,127,6,fill="#FFFFFF")+circ(144,127,6,fill="#FFFFFF")+
    sparkle(172,96,12,fill=BLUE)
)
# 02 문서·코드 업로드
icons["02_upload_docs"]=svg(
    doc(28,70,70,92,lines=4)+doc(102,70,70,92,code=True)+
    rr(72,10,56,44,r=10,fill=BLUE,stroke=BLUE)+arrow_up(100,46,26,color="#FFFFFF",sw=5)+
    line(20,180,180,180,stroke=LIGHT,sw=6)
)
# 03 지식 유닛 DB
icons["03_knowledge_db"]=svg(
    cylinder(20,70,100,110)+
    rr(112,26,68,36,r=6,fill="#FFFFFF")+text_lines(122,38,46,2,gap=12,color=BLUE,sw=5)+
    rr(124,68,68,36,r=6,fill="#FFFFFF")+text_lines(134,80,46,2,gap=12,color=BLUE,sw=5)+
    rr(136,110,58,36,r=6,fill="#FFFFFF")+text_lines(146,122,38,2,gap=12,color=BLUE,sw=5)
)
# 04 정제 깔때기
dots="".join(circ(x,y,r,fill=c) for (x,y,r,c) in [(30,22,7,LIGHT),(62,14,5,NAVY),(90,30,8,LIGHT),(126,16,6,BLUE),(158,28,7,LIGHT),(48,44,5,BLUE),(110,48,6,NAVY),(140,50,5,LIGHT),(76,38,4,NAVY)])
icons["04_refine_funnel"]=svg(
    dots+
    path("M28 64 H172 L118 118 V152 L82 168 V118 Z",fill=PALE,stroke=NAVY,sw=4)+
    line(46,84,154,84,stroke=LIGHT,sw=4)+
    rr(126,132,62,50,r=8,fill="#FFFFFF")+text_lines(136,146,42,3,gap=12,color=BLUE,sw=5)
)
# 05 토큰화
tok=""
cols=[NAVY,BLUE,LIGHT,BLUE,NAVY,LIGHT]
for r_ in range(3):
    for c_ in range(4):
        tok+=rr(116+c_*20,60+r_*26,16,18,r=3,fill=cols[(r_*4+c_)%6],stroke="none",sw=0)
icons["05_tokens"]=svg(
    rr(12,44,86,112,r=8,fill="#FFFFFF")+text_lines(24,64,62,6,gap=14,color=LIGHT,sw=5)+
    arrow_right(100,100,16,color=BLUE,sw=5)+tok
)
# 06/07/08 신경망
icons["06_neural_generic"]=svg(neural(100,100,[3,4,3],PURPLE,PURPLE_L,scale=1.25,node_r=9))
icons["07_neural_domain"]=svg(neural(100,100,[3,4,3],GREEN,GREEN_L,scale=1.25,node_r=9))
icons["08_neural_advanced"]=svg(neural(100,100,[4,5,4],BLUE,BLUE_L,scale=1.15,node_r=8))
# 09 피드백
stars="".join(path(f"M{cx} {y-11} L{cx+3.4} {y-3.5} L{cx+11} {y-3.5} L{cx+5} {y+1.5} L{cx+7} {y+9.5} L{cx} {y+5} L{cx-7} {y+9.5} L{cx-5} {y+1.5} L{cx-11} {y-3.5} L{cx-3.4} {y-3.5} Z",fill=(BLUE if i<4 else LIGHT),stroke="none") for i,cx in enumerate([48,74,100,126,152]) for y in [160])
icons["09_feedback"]=svg(
    rr(14,30,80,92,r=12,fill=PALE)+thumb(44,74,up=True,fill=BLUE)+
    rr(106,30,80,92,r=12,fill="#FFFFFF")+thumb(138,74,up=False,fill=LIGHT)+stars
)
# 10 지식증류·양자화
icons["10_distill_quantize"]=svg(
    neural(56,92,[3,4,3],BLUE,BLUE_L,scale=0.95,node_r=7)+
    arrow_right(104,92,26,color=NAVY,sw=5)+
    neural(160,92,[2,3,2],BLUE,BLUE_L,scale=0.55,node_r=5)+
    rr(126,140,68,40,r=6,fill=PALE)+"".join(rr(132+i*12,148,8,24,r=2,fill=(BLUE if i%2==0 else LIGHT),stroke="none",sw=0) for i in range(5))+
    path("M60 152 L84 152 M76 144 L84 152 L76 160",stroke=LIGHT,sw=4)+path("M20 152 L44 152 M28 144 L20 152 L28 160",stroke=LIGHT,sw=4)
)
# 11 GPU 서버
srv=rr(36,16,128,168,r=10,fill=PALE)
for i in range(4):
    y=28+i*40
    srv+=rr(46,y,108,30,r=5,fill="#FFFFFF")
    srv+=line(56,y+15,110,y+15,stroke=LIGHT,sw=5)
    srv+=circ(128,y+15,4,fill=(GREEN if i!=2 else BLUE))+circ(142,y+15,4,fill=BLUE)
icons["11_gpu_server"]=svg(srv)
# 12 웹·IDE 화면
icons["12_web_ide"]=svg(
    window(60,24,128,112,dark=True,code=True)+
    window(12,64,124,112,dark=False,code=False)+
    rr(24,96,100,14,r=4,fill=PALE,stroke=LIGHT,sw=3)+text_lines(24,124,100,3,gap=13,color=LIGHT,sw=5)
)
# 13 문서 생성
icons["13_doc_gen"]=svg(
    doc(40,22,104,132,lines=5,fold=22)+gear(146,140,26,teeth=8,fill=BLUE)+sparkle(34,40,11,fill=BLUE)
)
# 14 코드 생성
icons["14_code_gen"]=svg(
    window(16,30,168,140,dark=True,code=True)+
    rr(112,112,60,44,r=8,fill=BLUE,stroke=BLUE)+
    path("M130 124 L122 134 L130 144",stroke="#FFFFFF",sw=5)+path("M154 124 L162 134 L154 144",stroke="#FFFFFF",sw=5)+line(146,120,138,148,stroke="#FFFFFF",sw=4)
)
# 15 정적 검증
icons["15_static_check"]=svg(
    shield(100,100,116,136,fill=BLUE)+path("M70 104 L92 126 L134 78",stroke="#FFFFFF",sw=10)
)

# 16 학습 데이터셋 (카드 묶음 + 체크 배지)
icons["16_dataset"]=svg(
    rr(52,26,112,80,r=8,fill="#FFFFFF",stroke=LIGHT)+
    rr(40,48,112,80,r=8,fill="#FFFFFF",stroke=LIGHT)+
    rr(28,70,112,80,r=8,fill="#FFFFFF")+text_lines(44,92,80,3,gap=15,color=BLUE,sw=6)+
    circ(152,140,26,fill=BLUE)+path("M138 140 L148 150 L166 130",stroke="#FFFFFF",sw=7)
)
# 17 사내 LLM 답변 (파란 말풍선 + 반짝 + 코드/문서 결과)
icons["17_llm_answer"]=svg(
    bubble(18,26,132,74,BLUE,tail="l",stroke=BLUE)+sparkle(40,50,12,fill="#FFFFFF")+
    line(62,46,130,46,stroke="#FFFFFF",sw=6)+line(62,64,118,64,stroke="#FFFFFF",sw=6)+line(62,82,104,82,stroke=LIGHT,sw=6)+
    rr(96,110,92,74,r=8,fill="#FFFFFF")+line(96,128,188,128,stroke=NAVY,sw=3)+
    line(108,144,150,144,stroke=LIGHT,sw=5)+line(118,158,170,158,stroke=BLUE,sw=5)+line(118,171,158,171,stroke=NAVY,sw=5)
)
os.makedirs(OUT,exist_ok=True)
for k,v in icons.items():
    with open(os.path.join(OUT,k+".svg"),"w") as f: f.write(v)
print("written", len(icons))
