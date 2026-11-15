import openpyxl, json
refs=json.load(open("refs.json")); R=refs["R"]; k0=refs["k0"]
wb=openpyxl.load_workbook("../Watoskop_biznesplan.xlsx",data_only=True)
d={}
for key,sh in (("pes","Model_Pesymistyczny"),("real","Model_Realistyczny"),("opt","Model_Optymistyczny")):
    ws=wb[sh]; o={}
    for k in ("rev","var","fix","mkt","ebitda","ebit","net","cash","cust","new","rev_sub","rev_gw","gm","capex"):
        o[k]=[ws.cell(R[k],c).value for c in range(2,26)]
    o["kpi"]={ws.cell(k0+1+i,1).value:ws.cell(k0+1+i,2).value for i in range(16)}
    d[key]=o
json.dump(d,open("data.json","w"),ensure_ascii=False)
r=d["real"]; print({k:round(v) if isinstance(v,(int,float)) else v for k,v in r["kpi"].items()})
print([round(x) for x in r["ebit"]]); print(sum(r["capex"]))
