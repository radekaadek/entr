import json, math
from PIL import Image, ImageDraw, ImageFont
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
d=json.load(open("data.json"))

NAVY="#0B3D5C"; AMB="#F5A524"; GRN="#2BB673"; RED="#D9534F"; GREY="#8A99A6"
FB="/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"; FR="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
def f(sz,b=True): return ImageFont.truetype(FB if b else FR,sz)

# ---- logo
def logo_mark(sz):
    im=Image.new("RGBA",(sz,sz),(0,0,0,0)); g=ImageDraw.Draw(im)
    g.ellipse([sz*.04,sz*.04,sz*.96,sz*.96],fill=NAVY)
    g.arc([sz*.14,sz*.14,sz*.86,sz*.86],150,390,fill=AMB,width=int(sz*.07))
    bolt=[(.54,.2),(.32,.56),(.47,.56),(.42,.82),(.68,.43),(.52,.43),(.6,.2)]
    g.polygon([(x*sz,y*sz) for x,y in bolt],fill="white")
    return im
def logo_full(path,dark=False):
    W,H=1400,420; im=Image.new("RGBA",(W,H),(0,0,0,0)); m=logo_mark(380); im.paste(m,(20,20),m)
    g=ImageDraw.Draw(im); col="white" if dark else NAVY
    g.text((440,110),"Watoskop",font=f(170),fill=col)
    g.text((446,335),"Zobacz, za co płacisz za prąd.",font=f(46,False),fill=AMB if dark else GREY)
    im.save(path)
logo_mark(512).save("assets/logo_mark.png"); logo_full("assets/logo.png"); logo_full("assets/logo_dark.png",True)

# ---- charts
plt.rcParams.update({"font.family":"DejaVu Sans","axes.spines.top":False,"axes.spines.right":False,"font.size":11})
r=d["real"]; mo=list(range(1,25))
cost=[a+b+c for a,b,c in zip(r["var"],r["fix"],r["mkt"])]
fig,ax=plt.subplots(figsize=(8,4.2))
ax.bar(mo[:12],r["rev"][:12],color=AMB,label="Przychody",width=.7); ax.bar(mo[12:],r["rev"][12:],color=AMB,alpha=.45,width=.7)
ax.plot(mo,cost,color=NAVY,lw=2.5,label="Koszty (zmienne+stałe+marketing)")
ax.axvline(12.5,color=GREY,ls="--",lw=1); ax.text(12.3,max(r["rev"])*1.03,"← prognoza 12 mies.",color=GREY,fontsize=9,ha="right")
be=r["kpi"]["Miesiąc progu rentowności (EBIT ≥ 0)"]; ax.axvline(be,color=GRN,lw=1.5); ax.text(be+.3,max(r["rev"])*1.03,f"próg rentowności: M{be}",color=GRN,fontsize=10,ha="left")
ax.set_xlabel("miesiąc"); ax.set_ylabel("PLN / mies."); ax.legend(frameon=False,loc="upper left"); ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v,p:f"{int(v):,}".replace(","," ")))
fig.tight_layout(); fig.savefig("assets/c_revcost.png",dpi=200); plt.close()

fig,ax=plt.subplots(figsize=(8,4.2))
for k,c,l in (("opt",GRN,"Optymistyczny"),("real",NAVY,"Realistyczny"),("pes",RED,"Pesymistyczny")):
    ax.plot(mo[:12],d[k]["cash"][:12],color=c,lw=2.6,label=l)

ax.axhline(0,color=GREY,lw=1); ax.set_xlabel("miesiąc"); ax.set_ylabel("saldo gotówki, PLN"); ax.legend(frameon=False)
ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v,p:f"{int(v):,}".replace(","," ")))
fig.tight_layout(); fig.savefig("assets/c_cash.png",dpi=200); plt.close()

fig,ax=plt.subplots(figsize=(8,4.2))
for k,c,l in (("opt",GRN,"Optymistyczny"),("real",NAVY,"Realistyczny"),("pes",RED,"Pesymistyczny")):
    ax.plot(mo[:12],d[k]["cust"][:12],color=c,lw=2.6,label=l)
ax.set_xlabel("miesiąc"); ax.set_ylabel("klienci płatni"); ax.legend(frameon=False)
fig.tight_layout(); fig.savefig("assets/c_cust.png",dpi=200); plt.close()

# startup cost donut
labs=["Rozwój MVP (dev, UX, API)","Sprzęt i oprogramowanie","Prawo, IP, rejestracja","Marka i reklama","Zapas bramek"]
vals=[24000+6000+3000,12000+3000,900+4000+600,2500+3500,2150]
fig,ax=plt.subplots(figsize=(8,4.2))
w,_=ax.pie(vals,colors=[NAVY,AMB,GRN,"#5B8DB8",GREY],startangle=90,wedgeprops=dict(width=.42,edgecolor="white"))
ax.legend(w,[f"{l}: {v:,}".replace(","," ")+" PLN" for l,v in zip(labs,vals)],loc="center left",bbox_to_anchor=(.95,.5),frameon=False,fontsize=12)
ax.text(0,0,"61 650\nPLN",ha="center",va="center",fontsize=15,fontweight="bold",color=NAVY)
fig.tight_layout(); fig.savefig("assets/c_startup.png",dpi=200,bbox_inches="tight"); plt.close()

# ---- video frames 1280x720 @ 20fps
W,H=1280,720; FPS=20; fr=0
def ease(t): t=max(0,min(1,t)); return t*t*(3-2*t)
def save(im):
    global fr; im.convert("RGB").save(f"frames/f{fr:04d}.png"); fr+=1
def base(bg):
    im=Image.new("RGBA",(W,H),bg); return im,ImageDraw.Draw(im)
def ctext(g,y,t,font,fill): w=g.textlength(t,font=font); g.text(((W-w)/2,y),t,font=font,fill=fill)
mark=logo_mark(120)
# scene 1: problem (0-4s)
for i in range(4*FPS):
    t=i/(4*FPS); im,g=base("#14202B")
    ctext(g,120,"Rachunek za prąd w Twojej firmie:",f(40,False),"#C9D3DB")
    v=int(9480*ease(t*1.6)); s=f"{v:,}".replace(","," ")+" zł"
    ctext(g,230,s,f(150),RED if t>.2 else "white")
    if t>.55:
        a=ease((t-.55)/.25); ctext(g,470,"…ale za co dokładnie płacisz?",f(50),(255,int(255-60*a),int(255-90*a)))
    save(im)
# scene 2: scan (4s)
for i in range(4*FPS):
    t=i/(4*FPS); im,g=base("#F4F7FA")
    g.rounded_rectangle([140,90,520,640],14,fill="white",outline="#B8C4CE",width=3)
    for k in range(11):
        g.rectangle([170,130+k*44,170+(300 if k%3 else 200),142+k*44],fill="#DDE4EA")
    g.text((170,100),"FAKTURA ZA ENERGIĘ",font=f(18),fill=NAVY)
    y=int(110+500*ease(t*1.5)); g.rectangle([140,y,520,y+6],fill=AMB)
    g.text((600,130),"AI czyta rachunek…" if t<.5 else "Watoskop znalazł:",font=f(40),fill=NAVY)
    items=[("Moc umowna za wysoka","−1 240 zł / mies."),("Nieoptymalna taryfa","−620 zł / mies."),("Szczyt poboru w piątki 16:00","−410 zł / mies.")]
    for k,(a1,b1) in enumerate(items):
        p=ease((t-.45-k*.15)/.2)
        if p>0:
            x=600+int((1-p)*80)
            g.rounded_rectangle([x,230+k*110,1220,320+k*110],12,fill=NAVY if p>.9 else "#4A6F89")
            g.text((x+20,240+k*110),a1,font=f(26),fill="white"); g.text((x+20,277+k*110),b1,font=f(28),fill=AMB)
    save(im)
# scene 3: dashboard (4s)
import random; random.seed(3)
bars=[random.randint(40,100) for _ in range(14)]
for i in range(4*FPS):
    t=i/(4*FPS); im,g=base("#0B3D5C")
    g.text((70,50),"Twoje zużycie – na żywo",font=f(40),fill="white")
    g.text((70,110),"z falownika PV i bramki pomiarowej",font=f(24,False),fill="#9FB8CB")
    for k,b in enumerate(bars):
        p=ease(t*2-k*.05); hh=int(b*3.4*p); x=90+k*62
        g.rectangle([x,640-hh,x+44,640],fill=GRN if k%5 else RED)
    g.line([(70,640),(980,640)],fill="#9FB8CB",width=2)
    g.rounded_rectangle([1010,170,1230,330],14,fill="#12527B"); g.text((1030,185),"Oszczędność",font=f(22,False),fill="#9FB8CB")
    sv=int(2270*ease(t*1.4)); g.text((1030,225),f"{sv:,}".replace(","," ")+" zł",font=f(36),fill=AMB); g.text((1030,280),"miesięcznie",font=f(20,False),fill="white")
    g.rounded_rectangle([1010,360,1230,520],14,fill="#12527B"); g.text((1030,375),"Produkcja PV",font=f(22,False),fill="#9FB8CB"); g.text((1030,415),f"{int(31*ease(t*1.4))} kWh",font=f(36),fill=GRN); g.text((1030,470),"dziś",font=f(20,False),fill="white")
    save(im)
# scene 4: brand (4s)
for i in range(4*FPS):
    t=i/(4*FPS); im,g=base("#FFFFFF")
    p=ease(t*3); m=logo_mark(int(260*p)+1); im.paste(m,(int(W/2-m.width/2),int(70+(1-p)*40)),m)
    ctext(g,350,"Watoskop",f(110),NAVY)
    ctext(g,490,"Zobacz, za co płacisz za prąd.",f(44,False),AMB if False else "#C77D00")
    if t>.4:
        a=ease((t-.4)/.3); g.rounded_rectangle([300,580,980,665],42,fill=NAVY)
        ctext(g,596,"Pierwsza analiza rachunku za 0 zł  •  watoskop.pl",f(28),"white")
    save(im)
print("frames",fr)
