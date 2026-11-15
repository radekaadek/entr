import json
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
d=json.load(open("data.json")); real,pes,opt=d["real"],d["pes"],d["opt"]
K=real["kpi"]; k=lambda s,key:s["kpi"][key]
n=lambda v:f"{round(v):,}".replace(","," ")
NAVY=RGBColor(0x0B,0x3D,0x5C); AMB=RGBColor(0xF5,0xA5,0x24); GRN=RGBColor(0x2B,0xB6,0x73); RED=RGBColor(0xD9,0x53,0x4F)
LIGHT=RGBColor(0xF4,0xF7,0xFA); DARK=RGBColor(0x1E,0x2A,0x35); GREY=RGBColor(0x6B,0x7B,0x88); WHITE=RGBColor(255,255,255)
import openpyxl
_f=openpyxl.load_workbook("../Watoskop_biznesplan.xlsx",data_only=True)["Finansowanie"]
PES_MIN=_f["B8"].value; PES_NEED=_f["B9"].value
prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]; A="assets/"
def tb(s,x,y,w,h,text,size=16,bold=False,color=DARK,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=t.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    tf.margin_left=tf.margin_right=Inches(.05)
    lines=text if isinstance(text,list) else [text]
    for i,l in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); r=p.add_run(); r.text=l; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name="Calibri"; p.alignment=align; p.space_after=Pt(4)
    return t
def box(s,x,y,w,h,fill=LIGHT,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    b=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h)); b.fill.solid(); b.fill.fore_color.rgb=fill
    if line: b.line.color.rgb=line
    else: b.line.fill.background()
    if shape==MSO_SHAPE.ROUNDED_RECTANGLE: b.adjustments[0]=.08
    b.shadow.inherit=False; return b
def card(s,x,y,w,h,title,body,fill=LIGHT,tcol=NAVY,size=14,bcol=DARK):
    if isinstance(fill,int): size=fill; fill=LIGHT
    box(s,x,y,w,h,fill); tb(s,x+.15,y+.1,w-.3,.5,title,size+2,True,tcol); tb(s,x+.15,y+.6,w-.3,h-.7,body,size,False,bcol)
def stat(s,x,y,w,val,lab,col=NAVY):
    box(s,x,y,w,1.35,LIGHT); tb(s,x,y+.1,w,.7,val,32 if len(val)<9 else 24,True,col,PP_ALIGN.CENTER); tb(s,x+.1,y+.82,w-.2,.5,lab,12,False,GREY,PP_ALIGN.CENTER)
def slide(title,stage=None):
    s=prs.slides.add_slide(blank); box(s,0,0,13.333,.12,NAVY,shape=MSO_SHAPE.RECTANGLE)
    tb(s,.6,.35,9.7,1.0,title,28,True,NAVY)
    if stage: tb(s,10.5,.2,2.3,.6,stage,11,True,AMB,PP_ALIGN.RIGHT)
    s.shapes.add_picture(A+"logo.png",Inches(11.5),Inches(.7),height=Inches(.38)) if False else None
    tb(s,.6,7.05,6,.3,"Watoskop • Przedsiębiorczość technologiczna",10,False,GREY)
    tb(s,12.0,7.05,.8,.3,str(len(prs.slides)),10,False,GREY,PP_ALIGN.RIGHT); return s
def pic(s,f,x,y,w=None,h=None):
    return s.shapes.add_picture(A+f,Inches(x),Inches(y),width=Inches(w) if w else None,height=Inches(h) if h else None)
def table(s,x,y,w,rows,colw,size=12,hh=.4,rh=.38,hfill=NAVY):
    t=s.shapes.add_table(len(rows),len(rows[0]),Inches(x),Inches(y),Inches(w),Inches(hh+rh*(len(rows)-1))).table
    for j,cw in enumerate(colw): t.columns[j].width=Inches(cw)
    for i,r in enumerate(rows):
        for j,v in enumerate(r):
            c=t.cell(i,j); c.text=str(v); p=c.text_frame.paragraphs[0]; p.font.size=Pt(size); p.font.name="Calibri"
            c.margin_left=c.margin_right=Inches(.06); c.margin_top=c.margin_bottom=Inches(.03)
            c.fill.solid()
            if i==0: c.fill.fore_color.rgb=hfill; p.font.color.rgb=WHITE; p.font.bold=True
            else: c.fill.fore_color.rgb=LIGHT if i%2 else WHITE; p.font.color.rgb=DARK
            if j>0 and i>0 and isinstance(v,str) and len(v)<18 and any(ch.isdigit() for ch in v): p.alignment=PP_ALIGN.RIGHT
    return t
def arrow(s,x,y,w=.5): 
    a=s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,Inches(x),Inches(y),Inches(w),Inches(.4)); a.fill.solid(); a.fill.fore_color.rgb=AMB; a.line.fill.background()

# 1 title
s=prs.slides.add_slide(blank); bg=box(s,0,0,13.333,7.5,NAVY,shape=MSO_SHAPE.RECTANGLE)
pic(s,"logo_dark.png",.8,1.2,w=7.2)
tb(s,.8,3.6,11,1,"System AI analizujący rachunki i zużycie energii w małych firmach",28,True,WHITE)
tb(s,.8,4.6,11,.6,"Od pomysłu do pierwszych pieniędzy na koncie startupu",20,False,AMB)
tb(s,.8,6.2,11,.8,["Projekt z przedmiotu „Przedsiębiorczość technologiczna” • temat 1: Energia pod kontrolą","Autorzy: Radosław Dąbkowski, Paweł Kapela, Jakub Dobrosielski"],14,False,RGBColor(0xC9,0xD3,0xDB))

# 2 problem
s=slide("Małe firmy przepłacają za prąd, bo nie rozumieją rachunków","ETAP 1 • POTRZEBA")
stat(s,.6,1.5,2.9,"2,36 mln","mikro i małych firm w Polsce (PARP 2026)")
stat(s,3.7,1.5,2.9,"≈ 20%","z nich ma rachunek > 1,5 tys. zł/mies. (założenie)",AMB)
stat(s,6.8,1.5,2.9,"8–15%","możliwa oszczędność (deklaracje dostawców)",GRN)
stat(s,9.9,1.5,2.85,"0 etatów","firma 5–30 osób nie ma energetyka",RED)
card(s,.6,3.2,4,3.4,"Problem","• rachunek 2–6 stron, kilkanaście składników (dystrybucja, moc umowna, opłata mocowa…)\n• taryfa i moc umowna dobrane raz, „na oko”\n• szczyty poboru podbijają koszty\n• dane z falownika (urządzenia przy panelach) i instalacji PV (fotowoltaiki) leżą w osobnej aplikacji".replace("\n","\n"),size=14)
card(s,4.8,3.2,4,3.4,"Co dziś robią klienci","• księgowa „przepisuje” fakturę\n• sprzedawca energii proponuje ofertę, nie audyt\n• audyt energetyczny: tysiące zł, jednorazowo\n• brak ciągłej kontroli",size=14)
card(s,9,3.2,3.75,3.4,"Nasza hipoteza","Jeśli pokażemy właścicielowi w 60 sekund, za co płaci i ile może zaoszczędzić, zapłaci 149–299 zł miesięcznie – bo zwrot jest wielokrotny.",NAVY,WHITE,14,WHITE)

# 3 solution
s=slide("Watoskop: wrzucasz rachunek, dostajesz plan oszczędności","ETAP 1 • ROZWIĄZANIE")
steps=[("1. Wgraj rachunek","PDF/zdjęcie lub e-faktura; opcjonalnie dane z falownika/bramki"),("2. AI czyta i liczy","AI (sztuczna inteligencja): OCR (odczyt tekstu ze skanu) + LLM (duży model językowy) wyciąga składniki; silnik reguł sprawdza taryfę, moc umowną, opłaty"),("3. Raport i alerty","Lista działań z kwotą oszczędności; alerty o anomaliach zużycia")]
for i,(t,b) in enumerate(steps):
    x=.6+i*4.2; card(s,x,1.5,3.6,2.2,t,b,LIGHT,NAVY,13)
    if i<2: arrow(s,x+3.65,2.45,.5)
tb(s,.6,3.95,12,.4,"Na czym polega innowacyjność",20,True,NAVY)
card(s,.6,4.45,3.9,2.3,"Procesowa","Audyt energii z dnia na dzień zamiast tygodni; zastępuje ręczną analizę",size=14)
card(s,4.7,4.45,3.9,2.3,"Produktowa","Łączy rachunek (koszt) z danymi z falownika PV (zużycie i produkcja) w jednej analizie",size=14)
card(s,8.8,4.45,3.95,2.3,"Biznesowa","Abonament niski dla mikrofirmy + model „płacisz mniej niż 3% oszczędności”",size=14)

# 4 inverters
s=slide("Skąd bierzemy dane: falowników nie kupujemy, łączymy się z nimi","ETAP 1/5 • DANE I SPRZĘT")
cols=[("Klient ma PV (fotowoltaikę)","Ok. 15% rynku dostępnego (SAM). Klient loguje się kontem producenta (OAuth – bezpieczne logowanie bez podawania hasła nam). Pobieramy produkcję i pobór przez oficjalne API (interfejs do pobierania danych) chmury:\nHuawei FusionSolar • Growatt • SolaX • Fronius Solar.web • SMA Sunny Portal • Solis\nBez zakupu sprzętu, start w 5 minut.",GRN),
      ("Klient bez PV","Bramka pomiarowa z cęgami CT (przekładnikami prądowymi, typ Shelly Pro 3EM): kupujemy od dystrybutora (≈ 190 zł), sprzedajemy 349 zł, konfigurujemy zdalnie. Marża ≈ 134 zł/szt.",AMB),
      ("Bez sprzętu","Import godzinowych danych licznika z portalu OSD (operatora sieci dystrybucyjnej) w pliku CSV (tabelarycznym) oraz wgrane rachunki. Plan Start działa wyłącznie na rachunkach.",NAVY)]
for i,(t,b,c) in enumerate(cols):
    x=.6+i*4.2; box(s,x,1.5,3.95,.12,c,shape=MSO_SHAPE.RECTANGLE); card(s,x,1.62,3.95,3.1,t,b,LIGHT,NAVY,13)
box(s,.6,5.0,12.15,1.8,NAVY)
tb(s,.8,5.1,11.8,.4,"Co jest ryzykiem i jak je ograniczamy",16,True,AMB)
tb(s,.8,5.5,11.8,1.3,["• API producentów mogą się zmieniać lub być płatne → adapter na każdego producenta, start od dwóch (Huawei, Growatt) + zawsze import CSV.","• Zgoda klienta na dostęp do danych (RODO – ochrona danych osobowych, umowa powierzenia); token odwoływalny w każdej chwili.","• Zależność od dostawcy bramek → dwa źródła (Shelly lub odpowiednik), bramka to tylko ≈ 10% przychodu."],12,False,WHITE)

# 5 app architecture
s=slide("Jak tworzymy aplikację: MVP (wersja minimalna) w 4 miesiące, 3 fazy","ETAP 5 • PLAN TECHNICZNY")
layers=[("Klient","Web (React/Next.js, PWA – strona działająca jak aplikacja mobilna)"),("API","Python / FastAPI, uwierzytelnianie, rozliczenia"),("AI","OCR + LLM → JSON rachunku, walidacja, silnik reguł, wykrywanie anomalii"),("Dane","PostgreSQL + TimescaleDB, kolejka zadań"),("Konektory","Huawei • Growatt • CSV OSD • bramka (protokół MQTT dla urządzeń)")]
for i,(a,b) in enumerate(layers):
    y=1.5+i*.95; box(s,.6,y,1.9,.8,NAVY); tb(s,.6,y,1.9,.8,a,16,True,WHITE,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE); box(s,2.55,y,4.9,.8,LIGHT); tb(s,2.65,y,4.7,.8,b,12,False,DARK,anchor=MSO_ANCHOR.MIDDLE)
tb(s,.6,6.3,6.9,.6,"Chmura w UE (RODO), szyfrowanie, kopie zapasowe. M1 = pierwszy miesiąc działalności.",12,False,GREY)
ph=[("Faza 1 • M1–M2","Upload rachunku → raport oszczędności. Wartość od pierwszego dnia."),("Faza 2 • M3","Konektory Huawei i Growatt + import CSV."),("Faza 3 • M4","Bramka pomiarowa, alerty, rozliczanie abonamentów.")]
for i,(a,b) in enumerate(ph): card(s,7.8,1.5+i*1.55,4.95,1.4,a,b,LIGHT,NAVY,13)
tb(s,7.8,6.2,4.95,.7,"Koszt MVP: 33 tys. zł (dev zewn., UX) + 3 tys. zł API; zespół 2 założycieli.",12,True,NAVY)

# 6 market
s=slide("Rynek: ponad 470 tys. firm z istotnym rachunkiem za prąd","ETAP 2 • RYNEK")
tam=2356267*.2; sam=tam*.15; som=sam*.02; arpu=12*(.6*149+.4*299)
for i,(lab,val,vv,w,col) in enumerate([("TAM (cały rynek)",f"{n(tam)} firm",f"≈ {n(tam*arpu/1e6)} mln zł/rok",6.2,NAVY),("SAM (rynek dostępny)",f"{n(sam)} firm",f"≈ {n(sam*arpu/1e6)} mln zł/rok",5.4,RGBColor(0x2E,0x6F,0x9E)),("SOM (cel na 3 lata)",f"{n(som)} firm",f"≈ {n(som*arpu/1e6)} mln zł/rok",4.6,AMB)]):
    y=1.6+i*1.45; x=.6+(6.2-w)/2; box(s,x,y,w,1.25,col); tb(s,x,y+.05,w,.5,f"{lab}: {val}",17,True,WHITE,PP_ALIGN.CENTER); tb(s,x,y+.6,w,.5,vv,15,False,WHITE,PP_ALIGN.CENTER)
card(s,7.2,1.6,5.55,4.35,"Założenia i klienci docelowi","• TAM = cały rynek; SAM = część, którą możemy obsłużyć; SOM = udział, który zdobędziemy\n• Baza: 2,31 mln mikro + 48 tys. małych firm (PARP 2026)\n• 20% ma rachunek > 1,5 tys. zł/mies. (założenie)\n• SAM: 15% to firmy z PV lub 3-fazowym licznikiem gotowe na monitoring\n• SOM: 2% SAM po 3 latach\n• Średni przychód rocznie na klienta: "+n(arpu)+" zł (60% Start / 40% Pro)\n• Segmenty startowe: piekarnie, warsztaty, sklepy, biura, małe produkcje, hotele pensjonatowe",13)
tb(s,.6,6.1,12,.8,"Źródło: Raport PARP o stanie sektora MŚP 2026; kolejne współczynniki to założenia własne – arkusz „Rynek”.",11,False,GREY)

# 7 competition
s=slide("Konkurencja: nikt nie łączy rachunku i danych z falownika dla mikrofirm","ETAP 2 • KONKURENCI")
rows=[["Konkurent","Co oferuje","Czego brakuje","Cena"],
["ScanEnergy","Analiza rachunku ze zdjęcia, porównanie ofert sprzedawców","Brak monitoringu zużycia i PV; cel: zmiana sprzedawcy","brak cennika (prowizje od zmiany sprzedawcy)"],
["Plus Omnitorus „Prąd na żywo”","Monitoring licznikowy, wykrywanie strat w węzłach","Wdrożenie projektowe, nie dla mikrofirmy; bez analizy rachunku","wycena indywidualna"],
["Portale producentów falowników (FusionSolar, Growatt…)","Produkcja i zużycie PV, bezpłatnie","Zero analizy rachunku i taryf; tylko jeden producent","0 zł"],
["Atos Smart Facility","Monitoring i AI dla dużych obiektów","Segment enterprise","projekt, wiele tys. zł"],
["Watoskop","Rachunek + dane PV/bramki + rekomendacje","Nowy gracz, bez marki","149 / 299 zł mies."]]
table(s,.6,1.5,12.15,rows,[2.6,3.6,3.6,2.35],12,.4,.7)
tb(s,.6,6.0,12.1,.9,"Przewaga: jedna aplikacja dla mikrofirmy – cena mniejsza niż 3% rachunku, start w minutę, kilka marek falowników. Informacje o konkurentach z publicznych artykułów; ceny do weryfikacji ofertą.",13,False,GREY)

# 8 pricing + validation
s=slide("Cena 149/299 zł i plan sprawdzenia popytu","ETAP 2 • CENA I POPYT")
rows=[["Plan","Cena / mies.","Zawartość"],["Start","149 zł","analiza rachunków, raport miesięczny, rekomendacje"],["Pro","299 zł","+ dane z falownika/bramki, alerty, benchmark"],["Bramka","349 zł jednorazowo","pomiar dla firm bez PV (marża ≈ 134 zł)"]]
table(s,.6,1.5,7,rows,[1.3,2,3.7],13,.42,.5)
card(s,.6,3.8,7,2.9,"Dlaczego klient zapłaci","Rachunek 8 000 zł × 10% oszczędności = 800 zł/mies.\nCena Pro 299 zł → zwrot 2,7× (Start: 5,4×).\nLTV (wartość klienta w całym okresie) ≈ "+n(K["LTV – wartość klienta w całym okresie (marża brutto abonamentu / churn) [PLN]"])+" zł wobec CAC (kosztu pozyskania klienta) 550 zł (LTV/CAC ≈ "+f"{K['LTV / CAC (wartość klienta ÷ koszt jego pozyskania)']:.0f}"+").",LIGHT,NAVY,14)
card(s,7.9,1.5,4.85,5.2,"Walidacja popytu (do wykonania)","Wyniki ankiety nie są dołączone – zostaną uzupełnione po badaniu, nie szacujemy ich z góry.\n\n• 20 rozmów z właścicielami firm (piekarnie, warsztaty, sklepy)\n• ankieta online, n ≥ 50 (liczba respondentów)\n• kryterium sukcesu: ≥ 30% deklaruje gotowość zapłaty ≥ 149 zł\n• 10 bezpłatnych pilotaży w M4; konwersja 30–80% (3/6/8 klientów w scenariuszach)",LIGHT,NAVY,13)

# 9 brand
s=slide("Marka: Watoskop – „zobacz, za co płacisz za prąd”","ETAP 3 • MARKA")
box(s,.6,1.4,6.6,2.6,WHITE,RGBColor(0xDD,0xE4,0xEA)); pic(s,"logo.png",.9,1.65,w=6.0)
box(s,.6,4.2,3.2,1.8,NAVY); pic(s,"logo_dark.png",.75,4.45,w=2.9); box(s,4.0,4.2,3.2,1.8,AMB); pic(s,"logo_mark.png",4.9,4.35,h=1.5)
card(s,7.5,1.4,5.25,2.6,"Nazwa i hasło","Watoskop = watt + „-skop” (przyrząd do oglądania). Krótka, łatwa do wymówienia, sugeruje przejrzystość.\nHasło: „Zobacz, za co płacisz za prąd.”",LIGHT,NAVY,13)
card(s,7.5,4.2,5.25,2.6,"Kolory i weryfikacja","Granat (zaufanie), bursztyn (energia), zieleń (oszczędność).\nOdrzucono „WattPilot” (znana ładowarka Fronius).\nZnak: zgłoszenie w UPRP (Urząd Patentowy RP) i EUIPO (urząd UE ds. własności intelektualnej), klasy 9, 35, 42; domena watoskop.pl.",LIGHT,NAVY,12)

# 10 legal
s=slide("Założenie firmy: sp. z o.o. przez S24 (rejestracja online), koszt ≈ 0,9 tys. zł","ETAP 4 • FORMALNOŚCI")
rows=[["Krok","Czas","Koszt"],["Umowa spółki w S24 (e-podpis/Profil Zaufany)","1 dzień","w opłatach sądowych"],["Wpis do KRS (Krajowy Rejestr Sądowy), NIP, REGON, zgłoszenie VAT (podatek od towarów i usług)","ok. 1–2 tyg.","ok. 500–600 zł + ogłoszenie w MSiG (Monitor Sądowy i Gospodarczy)"],["Rachunek firmowy, wpłata kapitału 5 000 zł","2–3 dni","0–30 zł/mies."],["Zgłoszenia do ZUS (składki ubezpieczeniowe) – zatrudnienie założycieli","7 dni","składki w kosztach"],["Zgłoszenie znaku w UPRP, domena, RODO, regulamin","2–4 tyg.","ok. 4,6 tys. zł (prawnik+znak)"]]
table(s,.6,1.5,7.4,rows,[3.8,1.4,2.2],12,.4,.62)
card(s,8.2,1.5,4.55,2.45,"Dlaczego sp. z o.o., nie JDG","JDG = jednoosobowa działalność gospodarcza. Dwóch wspólników, ograniczona odpowiedzialność, wiarygodność dla inwestora/pożyczki, CIT 9% (podatek dochodowy od firm, stawka dla małych podatników), udziały dla inwestora.",LIGHT,NAVY,12)
card(s,8.2,4.1,4.55,2.7,"PKD 2025, podatki, wymagania","PKD 2025 (klasyfikacja działalności): 62.10.Z, 62.20.Z, 63.10.Z\nVAT czynny 23%; CIT 9%; księgowość pełna (700 zł/mies.)\nBrak koncesji. RODO: umowa powierzenia; dane z falowników tylko za zgodą.",LIGHT,NAVY,11)

# 11 org
s=slide("Organizacja: dwóch założycieli, dwóch freelancerów, 4 miesiące do MVP","ETAP 5 • ORGANIZACJA")
card(s,.6,1.5,3.9,2.55,"Zespół","• Założyciel 1 (CEO – prezes): sprzedaż, partnerzy, finanse\n• Założyciel 2 (CTO – dyrektor techniczny): produkt, AI, konektory\n• Freelancer UX (projektant interfejsu, 6 tys. zł) i programista backend (24 tys. zł)\n• Od M13: dodatkowy etat obsługi",size=12)
card(s,4.7,1.5,3.9,2.55,"Zasoby","• 2 laptopy (12 tys. zł), narzędzia dev/design (3 tys. zł)\n• chmura UE, API LLM/OCR (3 tys. zł testy)\n• bramki: zakup u dystrybutora na zamówienie",size=12)
card(s,8.8,1.5,3.95,2.55,"Koszt i czas uruchomienia","Nakłady startowe: "+n(K["Nakłady startowe razem [PLN]"])+" zł\nOferta gotowa do sprzedaży: początek M5.",LIGHT,NAVY,13)
pic(s,"c_startup.png",.6,4.2,h=2.7)
tb(s,7.2,4.4,5.5,2.4,["Harmonogram:","M1–M2 prototyp raportu • M3 konektory • M4 bramka + 10 pilotaży • M5 start sprzedaży • M13 pierwszy etat"],14,False,DARK)

# 12 campaign
s=slide("Kampania: wchodzimy tam, gdzie właściciel i tak szuka oszczędności","ETAP 6 • KAMPANIA")
card(s,.6,1.5,3.9,2.4,"Grupa docelowa","Właściciele i managerowie firm 5–30 osób, rachunek > 5 tys. zł; sekundarnie: biura rachunkowe i instalatorzy PV (polecający).",size=13)
card(s,4.7,1.5,3.9,2.4,"Przekaz","„Zobacz, za co płacisz za prąd.”\nPierwsza analiza za 0 zł – bez karty i bez zobowiązań.",NAVY,WHITE,13,WHITE)
card(s,8.8,1.5,3.95,2.4,"Kanały","Google/LinkedIn Ads • Meta remarketing (reklama do osób, które były na stronie) • SEO (pozycjonowanie w Google) + webinary • partnerzy z prowizją • lokalne spotkania firm",size=13)
rows=[["Okres","Działanie"],["M1–M3","strona WWW + lista oczekujących, content, materiał wideo"],["M4","10 pilotaży, referencje, case study"],["M5–M8","reklama płatna + webinary, start partnerów"],["M9–M12","skalowanie kanałów z najniższym CAC, polecenia"]]
table(s,.6,4.1,7.2,rows,[1.2,6],12,.38,.5)
card(s,8.0,4.1,4.75,2.7,"Efekt (realistyczny)","Budżet 12 mies.: "+n(K["Wydatki marketingowe M1–M12 [PLN]"])+" zł\nCAC ≈ 550 zł (mix płatny + partnerzy)\nKonwersja lead→klient ≈ 5%, koszt leada (potencjalnego klienta) ≈ 28 zł\nKlienci płatni po 12 mies.: "+n(K["Klienci płatni M12"]),LIGHT,NAVY,13)

# 13 ad
s=slide("Materiał reklamowy: 16-sekundowy spot „Zobacz, za co płacisz”","ETAP 6 • REKLAMA")
for i in range(4):
    x=.6+i*3.08; pic(s,f"ad{i+1}.png",x,1.5,w=2.95)
caps=["Problem: wysoki rachunek, nie wiadomo za co","AI czyta fakturę i znajduje oszczędności","Dashboard: dane z PV i bramki na żywo","Marka, hasło i oferta: pierwsza analiza za 0 zł"]
for i,c in enumerate(caps): tb(s,.6+i*3.08,3.1,2.95,.8,c,12,False,DARK,PP_ALIGN.CENTER)
card(s,.6,4.1,6,2.7,"Użycie","Plik: Watoskop_reklama.mp4 (1280×720, 16 s). Format na LinkedIn, YouTube, landing page; wersja pionowa 9:16 do reklam w mediach społecznościowych.\nCTA: „Pierwsza analiza rachunku za 0 zł”.",size=13)
rows=[["Pozycja budżetu","PLN / 12 mies."],["Reklama płatna (Google, LinkedIn)","18 700"],["Meta Ads","7 200"],["Content, SEO, webinary","6 800"],["Partnerzy (prowizje)","9 900"],["Wydarzenia i spotkania","4 200"],["Narzędzia (strona WWW, CRM – baza klientów)","8 700"]]
t=sum([0]) 
table(s,6.9,4.1,5.85,rows,[3.9,1.95],11,.34,.34)

# 14 costs
s=slide("Biznesplan: ile kosztuje start i prowadzenie firmy","ETAP 7 • NAKŁADY I KOSZTY")
pic(s,"c_startup.png",.5,1.4,h=3.2)
rows=[["Koszty miesięczne (realistycznie)","PLN"],["Wynagrodzenia założycieli (od M5, 2 os.)","7 000"],["Księgowość","700"],["Hosting i narzędzia","600"],["OC (ubezpieczenie odpowiedzialności) i prawnik","300"],["Pozostałe","200"],["Koszty stałe razem","8 800"],["Zmienne: chmura/AI 14 zł, obsługa 8 zł na klienta, płatności 1,5%","≈ 22 zł/klient"]]
table(s,6.4,1.5,6.35,rows,[4.55,1.8],12,.4,.4)
stat(s,.6,5.0,2.9,n(K["Nakłady startowe razem [PLN]"])+" zł","nakłady startowe",NAVY)
stat(s,3.7,5.0,2.9,n(K["Wydatki marketingowe M1–M12 [PLN]"])+" zł","marketing w 12 mies.",AMB)
stat(s,6.8,5.0,2.9,"CIT 9%","podatek dochodowy (mały podatnik)")
stat(s,9.9,5.0,2.85,"8%","oprocentowanie pożyczki 120 tys. zł")

# 15 sales
s=slide("Po 12 mies.: "+n(K["Klienci płatni M12"])+" klientów i "+n(K["MRR – miesięczny przychód z abonamentów w M12 [PLN]"])+" zł MRR (przychód z abonamentów)","ETAP 7 • PRZYCHODY")
pic(s,"c_revcost.png",.5,1.4,w=7.2)
rows=[["M","Klienci","Przychód zł"]]+[[f"M{m+1}",f"{real['cust'][m]:.0f}",n(real['rev'][m])] for m in range(4,12)]
table(s,8.0,1.5,4.75,rows,[1,1.6,2.15],12,.36,.36)
tb(s,.6,5.2,7.2,1.7,["Założenia: sprzedaż od M5, nowi klienci = budżet marketingowy poprzedniego miesiąca ÷ CAC 550 zł, 6 klientów z pilotaży, churn (odpływ klientów) 3% mies., 40% klientów w planie Pro, 60% z nich kupuje bramkę.","Przychody 12 mies.: "+n(K["Przychody M1–M12 [PLN]"])+" zł."],12,False,GREY)
tb(s,8.0,5.0,4.75,1.8,"Cena Start 149 zł, Pro 299 zł; ARPU (średni przychód na klienta) ≈ "+n(K["ARPU – średni przychód abonamentowy na klienta [PLN/mies., M12]"])+" zł.\nPełny rachunek M1–M24 w arkuszu Watoskop_biznesplan.xlsx.",12,False,DARK)

# 16 cash
s=slide("Płynność: próg rentowności w M"+str(K["Miesiąc progu rentowności (EBIT ≥ 0)"])+", finansowanie 180 tys. zł wystarcza","ETAP 7 • WYNIK I FINANSOWANIE")
pic(s,"c_cash.png",.5,1.4,w=7.0)
stat(s,7.9,1.5,2.4,"M"+str(K["Miesiąc progu rentowności (EBIT ≥ 0)"]),"próg rentowności (wynik operacyjny)",GRN)
stat(s,10.4,1.5,2.35,n(K["Minimalne saldo gotówki w M1–M12 [PLN]"])+" zł","min. saldo gotówki w M1–M12",NAVY)
card(s,7.9,3.1,4.85,3.7,"Źródła finansowania","1. Środki własne założycieli: 60 000 zł\n2. Pożyczka startowa: 120 000 zł (8%, karencja 12 mies.)\nRazem 180 000 zł.\nRealistycznie: najniższe saldo w 24 mies. = "+n(K["Minimalne saldo gotówki w 24 mies. [PLN]"])+" zł, dodatni zapas.\nScenariusz pesymistyczny NIE SPINA SIĘ: przy 180 tys. zł gotówka kończy się po ok. 9 mies., brakuje "+n(max(0,-PES_MIN))+" zł (łączna potrzeba ≈ "+n(PES_NEED)+" zł). Dodatkowe źródła (dotacje, inwestor) nie są założone.",LIGHT,NAVY,11)
tb(s,.6,6.0,7,.8,"Wynik netto 12 mies.: "+n(K["Wynik netto M1–M12 [PLN]"])+" zł (okres budowy produktu i rynku).",13,True,RED)

# 17 scenarios
s=slide("Ryzyko: w scenariuszu pesymistycznym budżet się nie spina","ETAP 8 • RYZYKO")
def row(l,key,fmt=n): return [l]+[fmt(k(x,key)) for x in (pes,real,opt)]
be=lambda x:str(k(x,"Miesiąc progu rentowności (EBIT ≥ 0)"))
rows=[["Wskaźnik","Pesymistyczny","Realistyczny","Optymistyczny"],
 ["Klienci M12"]+[n(k(x,"Klienci płatni M12")) for x in (pes,real,opt)],
 ["Przychody 12 mies. [zł]"]+[n(k(x,"Przychody M1–M12 [PLN]")) for x in (pes,real,opt)],
 ["Nakłady startowe [zł]"]+[n(k(x,"Nakłady startowe razem [PLN]")) for x in (pes,real,opt)],
 ["Próg rentowności"]+[("> M24" if be(x)=="> 24" else "M"+be(x)) for x in (pes,real,opt)],
 ["Min. saldo 24 mies. [zł]"]+[n(k(x,"Minimalne saldo gotówki w 24 mies. [PLN]")) for x in (pes,real,opt)],
 ["Brakujące finansowanie [zł]",n(max(0,-PES_MIN)),"0","0"]]
table(s,.6,1.5,6.7,rows,[2.2,1.5,1.5,1.5],11,.4,.42)
tb(s,.6,5.15,6.6,.4,"Pesymistyczny: CAC 800 zł, churn 5%, koszty +15%, MVP +30%, 30% Pro.",11,False,GREY)
pic(s,"c_cust.png",7.4,1.4,w=5.4)
card(s,.6,5.55,12.15,1.4,"Plan działań, gdy sprzedaż jest niższa","Ograniczyć koszty (pensje założycieli do 0 do M12 oszczędzają ≈ 7 tys. zł/mies.), wstrzymać etat, przesunąć budżet na kanał partnerski z najniższym CAC; scenariusz pesymistyczny się nie spina: przy finansowaniu 180 tys. zł gotówka kończy się po ok. 9 mies., a brakuje "+n(max(0,-PES_MIN))+" zł. Uzupełnienie (dotacja, inwestor) nie jest założone w biznesplanie. Przy braku postępu po M9: pivot (np. wyłącznie usługa audytu rachunków) albo zamknięcie spółki.",LIGHT,NAVY,11)

# 18 summary
s=slide("Czy uruchomiłbym ten startup? Tak, etapowo","ETAP 9 • DECYZJA")
card(s,.6,1.5,6,2.5,"Szanse i przewaga","Realistyczne szanse są dobre: niski koszt MVP, LTV/CAC ≈ 11, rynek >470 tys. firm. Przewaga: jedna aplikacja łączy rachunek i dane z falownika dla mikrofirm.",LIGHT,NAVY,13)
card(s,6.8,1.5,5.95,2.5,"Największe zagrożenie","Skuteczność akwizycji (CAC) i zależność od API producentów falowników. Przy CAC 800 zł i churn 5% firma nie wychodzi na zero w 24 mies.",LIGHT,RED,13)
stat(s,.6,4.2,2.9,"180 tys. zł","potrzebne finansowanie (realistyczne)",NAVY)
stat(s,3.7,4.2,2.9,"M"+str(K["Miesiąc progu rentowności (EBIT ≥ 0)"]),"dodatni wynik operacyjny",GRN)
card(s,6.8,4.2,5.95,2.6,"Decyzja","Zainwestowałbym etapami: najpierw 20 wywiadów i 10 pilotaży (≈ 92 tys. zł wydane do M4), a pełne 180 tys. zł po potwierdzeniu konwersji i CAC ≤ 600 zł.",NAVY,WHITE,13,WHITE)
tb(s,.6,5.9,6,.9,"Pytania, na które odpowiada projekt: pomysł, rynek, firma, MVP, kampania, finanse, ryzyko.",12,False,GREY)

prs.save("../Watoskop_prezentacja.pptx"); print(len(prs.slides))
