from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import LineChart, BarChart, Reference

NAME = "Watoskop"
wb = Workbook()
HDR = PatternFill("solid", fgColor="0B3D5C"); INP = PatternFill("solid", fgColor="FFF4CC")
SUB = PatternFill("solid", fgColor="DCE9F2"); TOT = PatternFill("solid", fgColor="E8F5E9")
WH = Font(color="FFFFFF", bold=True); B = Font(bold=True); BLUE = Font(color="0000FF")
NUM = '#,##0;[Red]-#,##0'; PCT = '0.0%'
def hdr(ws, r, c1, c2):
    for c in range(c1, c2+1):
        ws.cell(r, c).fill = HDR; ws.cell(r, c).font = WH
        ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ---------------- Założenia ----------------
A = wb.active; A.title = "Założenia"
A["A1"] = f"{NAME} – założenia biznesplanu (PLN netto; żółte pola = dane wejściowe)"; A["A1"].font = Font(bold=True, size=14)
for i, h in enumerate(["Parametr", "Pesymistyczny", "Realistyczny", "Optymistyczny", "Uzasadnienie / źródło"], 1):
    A.cell(3, i, h)
hdr(A, 3, 1, 5)
rows = [
 ("price_start", "Cena abonamentu Start [PLN/mies.]", 149, 149, 149, "ok. 2% typowego rachunku 5–10 tys. PLN; oszczędność 10% = 500–1000 PLN/mies. (założenie)"),
 ("price_pro", "Cena abonamentu Pro [PLN/mies.]", 299, 299, 299, "Start + dane z falownika/bramki, alerty, raport miesięczny"),
 ("price_gw", "Cena bramki pomiarowej [PLN, jednorazowo]", 349, 349, 349, "Detal bramek 3-fazowych ok. 400 PLN – podstawowy zakup w modelu pass-through"),
 ("gw_cost", "Koszt zakupu bramki od dystrybutora [PLN]", 215, 190, 180, "Hurtowa cena bramki 3-fazowej z CT"),
 ("gw_ship", "Wysyłka + konfiguracja bramki [PLN]", 25, 25, 25, "kurier + etykieta"),
 ("pro_share", "Udział klientów w planie Pro", 0.30, 0.40, 0.50, "Klienci z PV/falownikiem lub z bramką"),
 ("gw_attach", "Udział nowych klientów Pro kupujących bramkę", 0.50, 0.60, 0.70, "Pozostali Pro mają falownik z API (bez zakupu sprzętu)"),
 ("churn", "Miesięczny churn (odpływ) klientów", 0.05, 0.03, 0.02, "Typowo 2–5% w SaaS dla małych firm"),
 ("cac", "Średni koszt pozyskania płacącego klienta – CAC [PLN]", 800, 550, 400, "Mix: reklama + partnerzy (biura rachunkowe, instalatorzy PV) – patrz arkusz Marketing"),
 ("pilots", "Klienci pilotażowi konwertujący na płatnych w M5 [os.]", 3, 6, 8, "10 bezpłatnych pilotaży w M4, konwersja 30–80%"),
 ("cloud_pc", "Chmura + OCR/LLM na klienta [PLN/mies.]", 16, 14, 12, "Koszt przetwarzania ~3–5 rachunków/mies. + dane z falownika"),
 ("support_pc", "Obsługa klienta na klienta [PLN/mies.]", 10, 8, 7, "Onboarding + support (czas założycieli wyceniony)"),
 ("pay_fee", "Prowizja płatności (% przychodu z abonamentów)", 0.015, 0.015, 0.015, "Bramka płatności / faktury cykliczne"),
 ("dev_pay", "Wynagrodzenie założyciela w fazie MVP M1–M4 [PLN/mies. / os.]", 2000, 2000, 2000, "Koszt pracodawcy; 2 założycieli (umowa o pracę/powołanie)"),
 ("run_pay", "Wynagrodzenie założyciela od M5 [PLN/mies. / os.]", 3500, 3500, 3500, "Poniżej rynku; założyciele inwestują czas"),
 ("accounting", "Księgowość [PLN/mies.]", 700, 700, 700, "Biuro rachunkowe dla sp. z o.o. (pełna księgowość)"),
 ("tools", "Hosting bazowy, narzędzia SaaS [PLN/mies.]", 600, 600, 600, "Chmura UE, monitoring, CRM, e-mail"),
 ("ins", "OC + prawnik na zlecenie [PLN/mies.]", 300, 300, 300, "OC działalności IT + doradztwo"),
 ("other", "Pozostałe koszty stałe [PLN/mies.]", 200, 200, 200, "telefon, drobne"),
 ("hire", "Dodatkowy etat obsługi/sprzedaży od M13 [PLN/mies.]", 3500, 3500, 3500, "Część etatu / praktykant (poza okresem 12 mies.)"),
 ("fixed_mult", "Mnożnik kosztów stałych", 1.15, 1.00, 1.00, "Pesymistycznie: wzrost kosztów o 15%"),
 ("start_mult", "Mnożnik nakładów startowych", 1.30, 1.00, 0.95, "Pesymistycznie: przekroczenie kosztów MVP o 30%"),
 ("cit", "CIT – podatek dochodowy od firm (mały podatnik)", 0.09, 0.09, 0.09, "CIT 9% dla małych podatników"),
 ("equity", "Wkład własny założycieli [PLN]", 60000, 60000, 60000, "Kapitał zakładowy + dopłaty (M1)"),
 ("loan", "Pożyczka startowa [PLN]", 120000, 120000, 120000, "np. pożyczka dla startupów/mikropożyczka; wypłata M1"),
 ("rate", "Oprocentowanie pożyczki (rocznie)", 0.08, 0.08, 0.08, "uproszczenie rynkowe"),
 ("grace", "Karencja spłaty kapitału [mies.]", 12, 12, 12, "Spłata od M13"),
 ("term", "Okres spłaty kapitału po karencji [mies.]", 36, 36, 36, "raty równe"),
 ("dep_base", "Środki trwałe do amortyzacji [PLN]", 12000, 12000, 12000, "2 laptopy"),
 ("dep_m", "Okres amortyzacji [mies.]", 36, 36, 36, "liniowo"),
]
REF = {}
GROUPS = [
 ("PRZYCHODY – ceny i wolumen sprzedaży", "C8E6C9", "E8F5E9", ["price_start","price_pro","price_gw","pro_share","gw_attach","pilots","churn","cac"]),
 ("KOSZTY ZMIENNE – rosną z liczbą klientów", "F8BBB6", "FDECEA", ["gw_cost","gw_ship","cloud_pc","support_pc","pay_fee"]),
 ("KOSZTY STAŁE – miesięczne", "FFCC80", "FFF3E0", ["dev_pay","run_pay","accounting","tools","ins","other","hire","fixed_mult"]),
 ("NAKŁADY STARTOWE I AMORTYZACJA", "B3D4F0", "E3F0FB", ["start_mult","dep_base","dep_m"]),
 ("FINANSOWANIE – wpływy i spłaty", "D1C4E9", "EFE8F7", ["equity","loan","rate","grace","term"]),
 ("PODATKI", "CFD8DC", "ECEFF1", ["cit"]),
]
byk = {r[0]: r for r in rows}
rr = 4
for gname, hcol, fcol, keys in GROUPS:
    for c in range(1, 6):
        cell = A.cell(rr, c); cell.fill = PatternFill("solid", fgColor=hcol); cell.font = B
    A.cell(rr, 1, gname); rr += 1
    for k in keys:
        _, lab, p, r_, o, note = byk[k]; REF[k] = rr
        A.cell(rr, 1, lab); A.cell(rr, 5, note)
        A.row_dimensions[rr].outlineLevel = 1
        for c in (1, 5): A.cell(rr, c).fill = PatternFill("solid", fgColor=fcol)
        for j, v in enumerate((p, r_, o)):
            c = A.cell(rr, 2 + j, v); c.fill = PatternFill("solid", fgColor=fcol); c.font = BLUE
            c.number_format = PCT if isinstance(v, float) and v < 1.5 and k not in ("fixed_mult", "start_mult") else '#,##0.00' if isinstance(v, float) else '#,##0'
        rr += 1
A.cell(rr+1, 1, "Legenda: kolor wiersza = rodzaj pozycji (zielony przychód, czerwony koszt zmienny, pomarańczowy koszt stały, niebieski nakłady, fioletowy finansowanie); niebieska czcionka = dane wejściowe do zmiany.").font = Font(italic=True)
A.sheet_properties.outlinePr.summaryBelow = False
A.column_dimensions["A"].width = 62; A.column_dimensions["E"].width = 80
for c in "BCD": A.column_dimensions[c].width = 16
A.freeze_panes = "B4"

# ---------------- Marketing ----------------
M = wb.create_sheet("Marketing")
M["A1"] = "Budżet marketingowy – kanały [PLN netto/mies.] (M13–M24 = poziom M12)"; M["A1"].font = Font(bold=True, size=14)
M.cell(3, 1, "Kanał")
for m in range(1, 25): M.cell(3, 1 + m, f"M{m}")
hdr(M, 3, 1, 25)
chan = [
 ("Reklama płatna (Google Ads, LinkedIn)", [0,0,500,1000,2000,2000,2200,2200,2200,2200,2200,2200]),
 ("Meta Ads (Facebook/Instagram) – remarketing", [0,0,300,500,800,800,800,800,800,800,800,800]),
 ("Content, SEO, webinary „Jak czytać rachunek za prąd”", [500,500,500,500,600,600,600,600,600,600,600,600]),
 ("Partnerzy: prowizje biura rachunkowe/instalatorzy PV", [0,0,0,300,900,1000,1200,1300,1300,1300,1300,1300]),
 ("Wydarzenia branżowe i lokalne spotkania (stoiska, targi)", [0,0,0,200,500,500,500,500,500,500,500,500]),
 ("Narzędzia: strona WWW, CRM, e-mailing, A/B", [1000,1000,700,500,500,500,500,800,800,800,800,800]),
]
for i, (n, v) in enumerate(chan):
    r = 4 + i; M.cell(r, 1, n)
    for m in range(1, 25):
        c = M.cell(r, 1 + m, v[m-1] if m <= 12 else f"={L(13)}{r}")  # col M = month 12
        if m <= 12: c.fill = INP; c.font = BLUE
        c.number_format = NUM
r = 4 + len(chan); M.cell(r, 1, "RAZEM budżet marketingowy").font = B
for m in range(1, 25):
    c = M.cell(r, 1 + m, f"=SUM({L(1+m)}4:{L(1+m)}{r-1})"); c.font = B; c.fill = TOT; c.number_format = NUM
MKT_ROW = r
M.cell(r+1, 1, "Suma 12 mies."); M.cell(r+1, 2, f"=SUM(B{r}:M{r})").number_format = NUM
M.column_dimensions["A"].width = 58
for m in range(1, 25): M.column_dimensions[L(1+m)].width = 9

# ---------------- Nakłady startowe ----------------
S = wb.create_sheet("Nakłady startowe")
S["A1"] = "Nakłady na rozpoczęcie działalności [PLN netto] i harmonogram"; S["A1"].font = Font(bold=True, size=14)
for i, h in enumerate(["Pozycja", "Kwota [PLN]", "Grupa", "M1 %", "M2 %", "M3 %", "M4 %", "Komentarz"], 1): S.cell(3, i, h)
hdr(S, 3, 1, 8)
items = [
 ("Rejestracja sp. z o.o. (S24, opłata KRS, ogłoszenie)", 900, "Rejestracja", [1,0,0,0], "S24: koszt sądowy ok. 500–600 PLN + ogłoszenie MSiG"),
 ("Laptopy 2 szt.", 12000, "Wyposażenie", [1,0,0,0], "amortyzacja 36 mies."),
 ("Licencje i narzędzia deweloperskie/projektowe", 3000, "Oprogramowanie", [1,0,0,0], "IDE, Figma, repozytorium, monitoring"),
 ("MVP – programista zewnętrzny (backend/konektory)", 24000, "Opracowanie produktu", [.25,.25,.25,.25], "~4 mies. × 6 000 PLN"),
 ("MVP – UX/UI freelancer", 6000, "Opracowanie produktu", [.5,.5,0,0], "makiety, design system"),
 ("Kredyty API (LLM/OCR) i testy na rachunkach", 3000, "Opracowanie produktu", [0,.34,.33,.33], "walidacja ekstrakcji z rachunków"),
 ("Zapas bramek pilotażowych (10 szt.)", 2150, "Opracowanie produktu", [0,0,0,1], "10 × (bramka + wysyłka)"),
 ("Prawnik: regulamin, umowa powierzenia (RODO), umowy B2B", 4000, "Prawo i IP", [.5,.5,0,0], ""),
 ("Znak towarowy (UPRP) + domena", 600, "Prawo i IP", [1,0,0,0], "zgłoszenie znaku ok. 450 PLN"),
 ("Identyfikacja wizualna (logo, księga znaku)", 2500, "Marka", [1,0,0,0], ""),
 ("Materiał reklamowy (film, animacja, landing page)", 3500, "Marketing", [0,0,1,0], "reklama w załączniku"),
]
for i, (n, a, g, sh, cmt) in enumerate(items):
    r = 4 + i
    S.cell(r, 1, n); c = S.cell(r, 2, a); c.fill = INP; c.font = BLUE; c.number_format = NUM
    S.cell(r, 3, g)
    for j in range(4):
        c = S.cell(r, 4 + j, sh[j]); c.number_format = '0%'; c.fill = INP
    S.cell(r, 8, cmt)
last = 3 + len(items)
S.cell(last+1, 1, "RAZEM nakłady startowe (scenariusz realistyczny)").font = B
c = S.cell(last+1, 2, f"=SUM(B4:B{last})"); c.font = B; c.number_format = NUM; c.fill = TOT
S.cell(last+3, 1, "Harmonogram nakładów [PLN, przed mnożnikiem scenariusza]").font = B
S.cell(last+4, 1, "Miesiąc")
for m in range(1, 25): S.cell(last+4, 1+m, f"M{m}")
hdr(S, last+4, 1, 25)
S.cell(last+5, 1, "Nakłady w miesiącu")
for m in range(1, 25):
    if m <= 4:
        col = L(3 + m)
        f = f"=SUMPRODUCT($B$4:$B${last},{col}$4:{col}${last})"
    else: f = 0
    S.cell(last+5, 1+m, f).number_format = NUM
OUT_ROW = last + 5
S.column_dimensions["A"].width = 58; S.column_dimensions["C"].width = 22; S.column_dimensions["H"].width = 50
for m in range(2, 26): S.column_dimensions[L(m)].width = max(S.column_dimensions[L(m)].width or 0, 10)

# ---------------- Model sheets ----------------
def model(title, scol, nice):
    ws = wb.create_sheet(title)
    a = lambda k: f"Założenia!${scol}${REF[k]}"
    ws["A1"] = f"Model finansowy 24 mies. – scenariusz {nice} (M1–M12 = wymagana prognoza; M13–M24 = rozszerzenie do progu rentowności)"
    ws["A1"].font = Font(bold=True, size=13)
    ws.cell(2, 1, "Pozycja [PLN netto]")
    for m in range(1, 25): ws.cell(2, 1+m, f"M{m}")
    ws.cell(2, 26, "Suma M1–M12"); hdr(ws, 2, 1, 26)
    R = {}
    labels = [
     ("mkt","Budżet marketingowy"),("new","Nowi klienci (płatna akwizycja, opóźnienie 1 mies.)"),("pil","Konwersja pilotaży"),
     ("churned","Klienci utraceni (churn)"),("cust","Klienci płatni – koniec miesiąca"),("pro","  w tym Pro"),("start","  w tym Start"),("gw","Sprzedane bramki [szt.]"),
     None,("rev_sub","Przychód z abonamentów"),("rev_gw","Przychód ze sprzedaży bramek"),("rev","PRZYCHODY RAZEM"),None,
     ("v_cloud","Chmura + OCR/LLM"),("v_sup","Obsługa klienta"),("v_pay","Prowizje płatności"),("v_gw","Koszt bramek (zakup + wysyłka)"),("var","KOSZTY ZMIENNE RAZEM"),
     ("gm","Marża brutto"),("gmp","Marża brutto %"),None,
     ("f_pay","Wynagrodzenia założycieli"),("f_acc","Księgowość"),("f_tools","Hosting i narzędzia"),("f_ins","OC i prawnik"),("f_oth","Pozostałe"),("f_hire","Dodatkowy etat (od M13)"),("fix","KOSZTY STAŁE RAZEM"),
     ("mkt2","Marketing"),("ebitda","EBITDA (wynik operacyjny przed amortyzacją)"),("dep","Amortyzacja"),("ebit","EBIT (wynik operacyjny)"),("int","Odsetki od pożyczki"),("ebt","Wynik brutto"),("cum","Wynik narastająco"),("tax","CIT 9%"),("net","WYNIK NETTO"),None,
     ("cfo","Przepływ operacyjny (EBITDA − CIT − odsetki)"),("capex","Nakłady startowe (MVP, sprzęt, rejestracja)"),("eq","Wkład własny"),("loanin","Wypłata pożyczki"),("loanout","Spłata kapitału pożyczki"),("ncf","PRZEPŁYW NETTO"),("cash","SALDO GOTÓWKI"),("loanbal","Saldo pożyczki"),None,
     ("be","Pomocnik: miesiąc EBIT ≥ 0 (po M5)"),
    ]
    r = 3
    for it in labels:
        if it is None: r += 1; continue
        R[it[0]] = r; ws.cell(r, 1, it[1]); r += 1
    for m in range(1, 25):
        c = L(1+m); p = L(m) if m > 1 else None
        f = {}
        f["mkt"] = f"=Marketing!{c}{MKT_ROW}"
        f["new"] = "=0" if m < 5 else f"={p}{R['mkt']}/{a('cac')}"
        f["pil"] = f"={a('pilots')}" if m == 5 else "=0"
        f["churned"] = "=0" if m < 6 else f"={p}{R['cust']}*{a('churn')}"
        f["cust"] = (f"={c}{R['new']}+{c}{R['pil']}" if m == 1 else f"={p}{R['cust']}+{c}{R['new']}+{c}{R['pil']}-{c}{R['churned']}")
        f["pro"] = f"={c}{R['cust']}*{a('pro_share')}"
        f["start"] = f"={c}{R['cust']}-{c}{R['pro']}"
        f["gw"] = f"=({c}{R['new']}+{c}{R['pil']})*{a('pro_share')}*{a('gw_attach')}"
        f["rev_sub"] = f"={c}{R['start']}*{a('price_start')}+{c}{R['pro']}*{a('price_pro')}"
        f["rev_gw"] = f"={c}{R['gw']}*{a('price_gw')}"
        f["rev"] = f"={c}{R['rev_sub']}+{c}{R['rev_gw']}"
        f["v_cloud"] = f"={c}{R['cust']}*{a('cloud_pc')}"
        f["v_sup"] = f"={c}{R['cust']}*{a('support_pc')}"
        f["v_pay"] = f"={c}{R['rev_sub']}*{a('pay_fee')}"
        f["v_gw"] = f"={c}{R['gw']}*({a('gw_cost')}+{a('gw_ship')})"
        f["var"] = f"=SUM({c}{R['v_cloud']}:{c}{R['v_gw']})"
        f["gm"] = f"={c}{R['rev']}-{c}{R['var']}"
        f["gmp"] = f"=IF({c}{R['rev']}=0,0,{c}{R['gm']}/{c}{R['rev']})"
        f["f_pay"] = f"=2*{a('dev_pay') if m<=4 else a('run_pay')}*{a('fixed_mult')}"
        f["f_acc"] = f"={a('accounting')}*{a('fixed_mult')}"
        f["f_tools"] = f"={a('tools')}*{a('fixed_mult')}"
        f["f_ins"] = f"={a('ins')}*{a('fixed_mult')}"
        f["f_oth"] = f"={a('other')}*{a('fixed_mult')}"
        f["f_hire"] = f"={a('hire')}*{a('fixed_mult')}" if m >= 13 else "=0"
        f["fix"] = f"=SUM({c}{R['f_pay']}:{c}{R['f_hire']})"
        f["mkt2"] = f"={c}{R['mkt']}"
        f["ebitda"] = f"={c}{R['gm']}-{c}{R['fix']}-{c}{R['mkt2']}"
        f["dep"] = f"=IF({m}<={a('dep_m')},{a('dep_base')}/{a('dep_m')},0)"
        f["ebit"] = f"={c}{R['ebitda']}-{c}{R['dep']}"
        f["int"] = "=0" if m == 1 else f"={p}{R['loanbal']}*{a('rate')}/12"
        f["ebt"] = f"={c}{R['ebit']}-{c}{R['int']}"
        f["cum"] = f"={c}{R['ebt']}" if m == 1 else f"={p}{R['cum']}+{c}{R['ebt']}"
        f["tax"] = (f"=MAX(0,{a('cit')}*MAX(0,{c}{R['cum']}))" if m == 1 else
                    f"=MAX(0,{a('cit')}*MAX(0,{c}{R['cum']})-SUM($B{R['tax']}:{p}{R['tax']}))")
        f["net"] = f"={c}{R['ebt']}-{c}{R['tax']}"
        f["cfo"] = f"={c}{R['ebitda']}-{c}{R['tax']}-{c}{R['int']}"
        f["capex"] = f"='Nakłady startowe'!{c}{OUT_ROW}*{a('start_mult')}"
        f["eq"] = f"={a('equity')}" if m == 1 else "=0"
        f["loanin"] = f"={a('loan')}" if m == 1 else "=0"
        f["loanout"] = f"=IF({m}>{a('grace')},{a('loan')}/{a('term')},0)"
        f["ncf"] = f"={c}{R['cfo']}-{c}{R['capex']}+{c}{R['eq']}+{c}{R['loanin']}-{c}{R['loanout']}"
        f["cash"] = f"={c}{R['ncf']}" if m == 1 else f"={p}{R['cash']}+{c}{R['ncf']}"
        f["loanbal"] = (f"={c}{R['loanin']}-{c}{R['loanout']}" if m == 1 else f"={p}{R['loanbal']}+{c}{R['loanin']}-{c}{R['loanout']}")
        f["be"] = f"=IF(AND({m}>=5,{c}{R['ebit']}>=0),{m},99)"
        for k, v in f.items():
            cell = ws.cell(R[k], 1+m, v)
            cell.number_format = PCT if k == "gmp" else ('0.0' if k in ("new","churned","cust","pro","start","gw","pil") else NUM)
    for k in ("rev","var","fix","ebitda","ebit","net","ncf","cash","cust","gm"):
        for cc in range(1, 27): ws.cell(R[k], cc).font = B; ws.cell(R[k], cc).fill = TOT
    for k in ("mkt","new","pil","churned","rev_sub","rev_gw","rev","v_cloud","v_sup","v_pay","v_gw","var","gm","f_pay","f_acc","f_tools","f_ins","f_oth","f_hire","fix","mkt2","ebitda","dep","ebit","int","ebt","tax","net","cfo","capex","eq","loanin","loanout","ncf","gw"):
        c = ws.cell(R[k], 26, f"=SUM(B{R[k]}:M{R[k]})"); c.number_format = NUM
    ws.cell(R["cust"], 26, f"=M{R['cust']}").number_format = '0.0'
    ws.cell(R["cash"], 26, f"=M{R['cash']}").number_format = NUM
    # KPI block
    k0 = R["be"] + 2
    ws.cell(k0, 1, "KPI scenariusza").font = B
    kp = [
     ("Miesiąc progu rentowności (EBIT ≥ 0)", f"=IF(MIN(B{R['be']}:Y{R['be']})=99,\"> 24\",MIN(B{R['be']}:Y{R['be']}))"),
     ("Minimalne saldo gotówki w 24 mies. [PLN]", f"=MIN(B{R['cash']}:Y{R['cash']})"),
     ("Minimalne saldo gotówki w M1–M12 [PLN]", f"=MIN(B{R['cash']}:M{R['cash']})"),
     ("Niedobór finansowania (jeśli saldo < 0) [PLN]", f"=MAX(0,-MIN(B{R['cash']}:Y{R['cash']}))"),
     ("Przychody M1–M12 [PLN]", f"=Z{R['rev']}"),
     ("Wynik netto M1–M12 [PLN]", f"=Z{R['net']}"),
     ("Klienci płatni M12", f"=M{R['cust']}"),
     ("MRR – miesięczny przychód z abonamentów w M12 [PLN]", f"=M{R['rev_sub']}"),
     ("Przychód M12 [PLN]", f"=M{R['rev']}"),
     ("Nakłady startowe razem [PLN]", f"=Z{R['capex']}"),
     ("Wydatki marketingowe M1–M12 [PLN]", f"=Z{R['mkt']}"),
     ("Wynik netto M24 (miesięczny) [PLN]", f"=Y{R['net']}"),
     ("Liczba klientów w M24", f"=Y{R['cust']}"),
     ("ARPU – średni przychód abonamentowy na klienta [PLN/mies., M12]", f"=IF(M{R['cust']}=0,0,M{R['rev_sub']}/M{R['cust']})"),
     ("LTV – wartość klienta w całym okresie (marża brutto abonamentu / churn) [PLN]", f"=IF({a('churn')}=0,0,(M{R['rev_sub']}/M{R['cust']}-{a('cloud_pc')}-{a('support_pc')}-M{R['rev_sub']}/M{R['cust']}*{a('pay_fee')})/{a('churn')})"),
     ("LTV / CAC (wartość klienta ÷ koszt jego pozyskania)", f"=B{k0+15}/{a('cac')}"),
    ]
    for i, (lab, fm) in enumerate(kp):
        ws.cell(k0+1+i, 1, lab); ws.cell(k0+1+i, 2, fm).number_format = '#,##0.0' if "LTV / CAC" in lab else NUM
    ws.column_dimensions["A"].width = 52
    for m in range(1, 27): ws.column_dimensions[L(1+m)].width = 10
    ws.column_dimensions["Z"].width = 13
    ws.freeze_panes = "B3"
    return ws, R, k0
sheets = {}
for t, col, nice in (("Model_Realistyczny", "C", "REALISTYCZNY"), ("Model_Pesymistyczny", "B", "PESYMISTYCZNY"), ("Model_Optymistyczny", "D", "OPTYMISTYCZNY")):
    sheets[t] = model(t, col, nice)
ws, R, k0 = sheets["Model_Realistyczny"]

# charts on realistic sheet
def line(ws, title, rows, anchor, cats_row=2, ncol=13):
    ch = LineChart(); ch.title = title; ch.height = 8; ch.width = 18
    for rr in rows:
        ch.add_data(Reference(ws, min_col=1, max_col=ncol, min_row=rr), titles_from_data=True, from_rows=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=ncol, min_row=cats_row)); ws.add_chart(ch, anchor)
base = k0 + 20
line(ws, "Przychody vs koszty (M1–M12)", [R['rev'], R['var']], f"A{base}")
line(ws, "Saldo gotówki (M1–M24)", [R['cash']], f"H{base}", ncol=25)
line(ws, "Klienci płatni (M1–M24)", [R['cust']], f"A{base+18}", ncol=25)

# ---------------- Scenariusze ----------------
SC = wb.create_sheet("Scenariusze")
SC["A1"] = "Porównanie scenariuszy"; SC["A1"].font = Font(bold=True, size=14)
for i, h in enumerate(["Wskaźnik", "Pesymistyczny", "Realistyczny", "Optymistyczny"], 1): SC.cell(3, i, h)
hdr(SC, 3, 1, 4)
names = [("Model_Pesymistyczny"), ("Model_Realistyczny"), ("Model_Optymistyczny")]
for i in range(16):
    SC.cell(4+i, 1, f"='Model_Realistyczny'!A{k0+1+i}")
    for j, n in enumerate(names):
        c = SC.cell(4+i, 2+j, f"='{n}'!B{k0+1+i}"); c.number_format = '#,##0.0' if i == 15 else NUM
SC.column_dimensions["A"].width = 52
for c in "BCD": SC.column_dimensions[c].width = 18
# monthly comparison blocks
SC.cell(22, 1, "Saldo gotówki [PLN]").font = B; SC.cell(22, 1)
SC.cell(23, 1, "Miesiąc")
for m in range(1, 13): SC.cell(23, 1+m, f"M{m}")
hdr(SC, 23, 1, 13)
for j, n in enumerate(names):
    SC.cell(24+j, 1, ["Pesymistyczny", "Realistyczny", "Optymistyczny"][j])
    SC.cell(28+j, 1, ["Pesymistyczny", "Realistyczny", "Optymistyczny"][j])
    for m in range(1, 13):
        SC.cell(24+j, 1+m, f"='{n}'!{L(1+m)}{R['cash']}").number_format = NUM
        SC.cell(28+j, 1+m, f"='{n}'!{L(1+m)}{R['rev']}").number_format = NUM
SC.cell(27, 1, "Przychody miesięczne [PLN]").font = B
SC.cell(26, 1)
ch = LineChart(); ch.title = "Saldo gotówki w scenariuszach (M1–M12)"; ch.height = 8; ch.width = 18
ch.add_data(Reference(SC, min_col=1, max_col=13, min_row=24, max_row=26), titles_from_data=True, from_rows=True)
ch.set_categories(Reference(SC, min_col=2, max_col=13, min_row=23)); SC.add_chart(ch, "F3")
ch2 = BarChart(); ch2.title = "Przychody miesięczne w scenariuszach"; ch2.height = 8; ch2.width = 18
ch2.add_data(Reference(SC, min_col=1, max_col=13, min_row=28, max_row=30), titles_from_data=True, from_rows=True)
ch2.set_categories(Reference(SC, min_col=2, max_col=13, min_row=23)); SC.add_chart(ch2, "F20")
for m in range(2, 14): SC.column_dimensions[L(m)].width = max(SC.column_dimensions[L(m)].width or 0, 11)

# ---------------- Finansowanie ----------------
F = wb.create_sheet("Finansowanie")
F["A1"] = "Źródła finansowania i zapotrzebowanie"; F["A1"].font = Font(bold=True, size=14)
for i, h in enumerate(["Źródło", "Kwota [PLN]", "Uwagi"], 1): F.cell(3, i, h)
hdr(F, 3, 1, 3)
F.append(["Środki własne założycieli (2 osoby)", f"=Założenia!C{REF['equity']}", "kapitał zakładowy + dopłaty; wybrane źródło nr 1"])
F.append(["Pożyczka startowa", f"=Założenia!C{REF['loan']}", "8% rocznie, karencja 12 mies., spłata 36 rat (wybrane źródło nr 2)"])
F.append(["RAZEM zaplanowane finansowanie", "=B4+B5", ""])
F.append(["Minimalne saldo gotówki – realistyczny", f"='Model_Realistyczny'!B{k0+2}", "najniższy punkt salda w 24 mies."])
F.append(["Minimalne saldo gotówki – pesymistyczny", f"='Model_Pesymistyczny'!B{k0+2}", "jeśli < 0 – niedobór"])
F.append(["Minimalne zapotrzebowanie na finansowanie (pesymistyczny)", f"=MAX(0,B4+B5-B8)", "finansowanie potrzebne, by saldo nie spadło poniżej 0"])
F.append(["Rekomendowane finansowanie z buforem 20% (pesymistyczny)", "=ROUNDUP(B9*1.2,-3)", ""])
F.append([]); 
F.append(["Alternatywy (opis): inwestor aniołowy (np. 15–20% udziałów), programy wsparcia startupów (PARP/NCBiR, fundusze regionalne), dotacje/pożyczki rozwojowe.", "", ""])
for r in range(4, 11): F.cell(r, 2).number_format = NUM
F.column_dimensions["A"].width = 62; F.column_dimensions["B"].width = 18; F.column_dimensions["C"].width = 70

# ---------------- Rynek ----------------
K = wb.create_sheet("Rynek")
K["A1"] = "Wielkość rynku (TAM / SAM / SOM)"; K["A1"].font = Font(bold=True, size=14)
for i, h in enumerate(["Pozycja", "Wartość", "Źródło / założenie"], 1): K.cell(3, i, h)
hdr(K, 3, 1, 3)
mk = [
 ("Mikro i małe firmy w Polsce (PARP/GUS, ok. 2,36 mln)", 2356267, "Raport PARP o stanie MŚP 2026: 2 307 838 mikro + 48 429 małych"),
 ("% firm z istotnym rachunkiem za prąd (> 1,5 tys. PLN/mies.)", 0.20, "założenie własne"),
 ("TAM (cały rynek) – liczba firm (potencjalni klienci)", "=B4*B5", ""),
 ("Średni przychód roczny na klienta [PLN] (mix 60% Start / 40% Pro)", "=12*(0.6*149+0.4*299)", "cennik"),
 ("TAM wartość [PLN/rok]", "=B6*B7", ""),
 ("SAM (rynek dostępny dla nas) – firmy z PV (fotowoltaiką) lub gotowe na cyfrowy monitoring (15% TAM)", "=B6*0.15", "założenie: wstępnie klienci z PV lub 3-fazowym liczniekiem"),
 ("SAM wartość [PLN/rok]", "=B9*B7", ""),
 ("SOM (rynek, który realnie zdobędziemy) – cel po 3 latach: 2% SAM (liczba firm)", "=B9*0.02", "realistyczne przejęcie rynku po 3 latach"),
 ("SOM wartość [PLN/rok]", "=B11*B7", ""),
]
for i, (n, v, s) in enumerate(mk):
    K.cell(4+i, 1, n); c = K.cell(4+i, 2, v); K.cell(4+i, 3, s)
    c.number_format = PCT if isinstance(v, float) and v < 1 else NUM
    if not str(v).startswith("="): c.fill = INP
K.column_dimensions["A"].width = 66; K.column_dimensions["B"].width = 18; K.column_dimensions["C"].width = 70


# ---------------- Budżet obok parametrów ----------------
from openpyxl.formatting.rule import CellIsRule, FormulaRule
mods = ["Model_Pesymistyczny", "Model_Realistyczny", "Model_Optymistyczny"]
for i, h in enumerate(["BUDŻET 12 MIES. [PLN]", "Pesymistyczny", "Realistyczny", "Optymistyczny"]): A.cell(3, 7+i, h)
hdr(A, 3, 7, 10)
def Z(m, key): return f"'{m}'!Z{R[key]}"
brows = [
 ("PRZYCHODY", None, "C8E6C9", True),
 ("Abonamenty", lambda m: "="+Z(m,"rev_sub"), "E8F5E9", False),
 ("Sprzedaż bramek", lambda m: "="+Z(m,"rev_gw"), "E8F5E9", False),
 ("Przychody razem", lambda m: "="+Z(m,"rev"), "C8E6C9", True),
 ("KOSZTY", None, "F8BBB6", True),
 ("Koszty zmienne", lambda m: "="+Z(m,"var"), "FDECEA", False),
 ("Koszty stałe", lambda m: "="+Z(m,"fix"), "FFF3E0", False),
 ("Marketing", lambda m: "="+Z(m,"mkt"), "FFF3E0", False),
 ("Nakłady startowe (MVP, sprzęt, rejestracja)", lambda m: "="+Z(m,"capex"), "E3F0FB", False),
 ("Odsetki od pożyczki", lambda m: "="+Z(m,"int"), "ECEFF1", False),
 ("CIT", lambda m: "="+Z(m,"tax"), "ECEFF1", False),
 ("Koszty razem", "SUMCOST", "F8BBB6", True),
 ("WYNIK GOTÓWKOWY (przychody − koszty)", "RESULT", "FFE082", True),
 ("FINANSOWANIE", None, "D1C4E9", True),
 ("Wkład własny + pożyczka", lambda m: f"={Z(m,'eq')}+{Z(m,'loanin')}", "EFE8F7", False),
 ("Spłata kapitału pożyczki w M1–M12", lambda m: "="+Z(m,"loanout"), "EFE8F7", False),
 ("SALDO KOŃCOWE M12 (wynik + finansowanie − spłata)", "END", "FFE082", True),
 ("Saldo z modelu miesięcznego (kontrola)", lambda m: f"='{m}'!M{R['cash']}", "FFFFFF", False),
 ("Różnica (powinno być 0)", "DIFF", "FFFFFF", False),
 ("Najniższe saldo w M1–M12", lambda m: f"=MIN('{m}'!B{R['cash']}:M{R['cash']})", "FFFFFF", False),
 ("CZY BUDŻET SIĘ SPINA?", "STATUS", "FFE082", True),
]
r0 = 4; pos = {}
for i, (lab, fn, col, bold) in enumerate(brows):
    r = r0 + i; pos[lab] = r
    A.cell(r, 7, lab).fill = PatternFill("solid", fgColor=col)
    if bold: A.cell(r, 7).font = B
    for j, m in enumerate(mods):
        c = A.cell(r, 8+j); c.fill = PatternFill("solid", fgColor=col); c.number_format = NUM
        L_ = L(8+j)
        if fn is None: continue
        if fn == "SUMCOST": f_ = f"=SUM({L_}{pos['Koszty zmienne']}:{L_}{pos['CIT']})"
        elif fn == "RESULT": f_ = f"={L_}{pos['Przychody razem']}-{L_}{pos['Koszty razem']}"
        elif fn == "END": f_ = f"={L_}{pos['WYNIK GOTÓWKOWY (przychody − koszty)']}+{L_}{pos['Wkład własny + pożyczka']}-{L_}{pos['Spłata kapitału pożyczki w M1–M12']}"
        elif fn == "DIFF": f_ = f"=ROUND({L_}{pos['SALDO KOŃCOWE M12 (wynik + finansowanie − spłata)']}-{L_}{pos['Saldo z modelu miesięcznego (kontrola)']},2)"
        elif fn == "STATUS": f_ = f'=IF({L_}{pos["Najniższe saldo w M1–M12"]}>=0,"SPINA SIĘ","NIE SPINA SIĘ: brakuje "&TEXT(-{L_}{pos["Najniższe saldo w M1–M12"]},"#,##0")&" zł")'
        else: f_ = fn(m)
        c.value = f_
        if bold: c.font = B
st = pos["CZY BUDŻET SIĘ SPINA?"]
A.conditional_formatting.add(f"H{st}:J{st}", FormulaRule(formula=[f'LEFT(H{st},5)="SPINA"'], fill=PatternFill("solid", bgColor="A5D6A7", fgColor="A5D6A7"), font=Font(bold=True, color="1B5E20")))
A.conditional_formatting.add(f"H{st}:J{st}", FormulaRule(formula=[f'LEFT(H{st},3)="NIE"'], fill=PatternFill("solid", bgColor="EF9A9A", fgColor="EF9A9A"), font=Font(bold=True, color="B71C1C")))
A.cell(st+2, 7, "Budżet = suma M1–M12 z arkuszy modeli. Spina się, gdy saldo gotówki nie spada poniżej 0 w żadnym miesiącu.").font = Font(italic=True)
A.column_dimensions["F"].width = 3; A.column_dimensions["G"].width = 52
for c in "HIJ": A.column_dimensions[c].width = 22

order = ["Założenia", "Rynek", "Nakłady startowe", "Marketing", "Model_Realistyczny", "Model_Pesymistyczny", "Model_Optymistyczny", "Scenariusze", "Finansowanie"]
wb._sheets = [wb[n] for n in order]
wb.save("/home/ard/prg/entr/build/biznesplan_raw.xlsx")
import json; json.dump({"R": R, "k0": k0, "REF": REF, "MKT_ROW": MKT_ROW}, open("/home/ard/prg/entr/build/refs.json", "w"))
print("ok", k0)
