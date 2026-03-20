# =========================================================
# 1) IMPORT & SETUP
# =========================================================
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import requests
import geopandas as gpd
import os
from io import BytesIO
from matplotlib.font_manager import FontProperties

sns.set_theme(style="darkgrid")
thai_font = FontProperties(family="Sukhumvit Set")

# =========================================================
# 2) DATA PREPARATION
# =========================================================
data = {
    "Country": [
        "Brunei","Cambodia","Indonesia","Lao PDR","Malaysia",
        "Myanmar","Philippines","Singapore","Thailand","Vietnam"],
    "Per_Capita_MSW_kg_per_day": [1.4,0.55,0.70,0.69,1.17,0.53,0.69,3.763,1.05,0.84],
    "Annual_MSW_ton": [210480,1089429,64000000,77380,12840000,841508,14660000,7514500,26770000,22020000]}
df_asean = pd.DataFrame(data)
df_asean["Annual_MSW_million_ton"] = df_asean["Annual_MSW_ton"] / 1_000_000
url = "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/%E0%B8%82%E0%B8%A2%E0%B8%B0%E0%B8%A1%E0%B8%B9%E0%B8%A5%E0%B8%9D%E0%B8%AD%E0%B8%A2%E0%B8%A3%E0%B8%B2%E0%B8%A2%E0%B8%88%E0%B8%B1%E0%B8%87%E0%B8%AB%E0%B8%A7%E0%B8%B1%E0%B8%94.xlsx"
pop_url = "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/%E0%B8%88%E0%B8%B3%E0%B8%99%E0%B8%A7%E0%B8%99%E0%B8%9B%E0%B8%A3%E0%B8%B0%E0%B8%8A%E0%B8%B2%E0%B8%81%E0%B8%A3%E0%B8%A3%E0%B8%B2%E0%B8%A2%E0%B8%88%E0%B8%B1%E0%B8%87%E0%B8%AB%E0%B8%A7%E0%B8%B1%E0%B8%94.xlsx"
url_disposal = "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/%E0%B8%AA%E0%B8%96%E0%B8%B2%E0%B8%99%E0%B8%97%E0%B8%B5%E0%B9%88%E0%B8%81%E0%B8%B3%E0%B8%88%E0%B8%B1%E0%B8%94%E0%B8%82%E0%B8%A2%E0%B8%B0%E0%B8%A1%E0%B8%B9%E0%B8%A5%E0%B8%9D%E0%B8%AD%E0%B8%A2.xlsx"
df_disposal = pd.read_excel(url_disposal)
url_disposal_cap= "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/%E0%B8%AA%E0%B8%96%E0%B8%B2%E0%B8%99%E0%B8%97%E0%B8%B5%E0%B9%88%E0%B8%81%E0%B8%B3%E0%B8%88%E0%B8%B1%E0%B8%94%E0%B8%82%E0%B8%A2%E0%B8%B0%E0%B8%A1%E0%B8%B9%E0%B8%A5%E0%B8%9D%E0%B8%AD%E0%B8%A2.xlsx"
geo_url = "https://raw.githubusercontent.com/apisit/thailand.json/master/thailand.json"
gdf = gpd.read_file(geo_url)
url_residual = "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/%E0%B8%82%E0%B8%A2%E0%B8%B0%E0%B8%95%E0%B8%81%E0%B8%84%E0%B9%89%E0%B8%B2%E0%B8%87.xlsx"
xls_residual = pd.ExcelFile(url_residual)
url_2566 = "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/factor2566.xlsx"
xls_2566 = pd.ExcelFile(url_2566)
base_url = "https://raw.githubusercontent.com/poonyanood124/Municipal-solid-waste/main/"
url_legal = base_url + "waste_disposal_legal.xlsx"
url_illegal = base_url + "waste_disposal_illegal.xlsx"
url_recovery = base_url + "province_waste_recovery.xlsx"

region_map = {
    "เชียงใหม่":"North","เชียงราย":"North","ลำปาง":"North",
    "ลำพูน":"North","แม่ฮ่องสอน":"North","น่าน":"North",
    "พะเยา":"North","แพร่":"North","อุตรดิตถ์":"North",
    "กำแพงเพชร":"North","ตาก":"North","พิษณุโลก":"North",
    "พิจิตร":"North","เพชรบูรณ์":"North","สุโขทัย":"North",
    "อุทัยธานี":"North","นครสวรรค์":"North",
    "กรุงเทพมหานคร":"Central","นนทบุรี":"Central","ปทุมธานี":"Central",
    "สมุทรปราการ":"Central","สมุทรสาคร":"Central","สมุทรสงคราม":"Central",
    "พระนครศรีอยุธยา":"Central","อ่างทอง":"Central","ลพบุรี":"Central",
    "สระบุรี":"Central","ชัยนาท":"Central","นครปฐม":"Central",
    "สุพรรณบุรี":"Central","สิงห์บุรี":"Central","นครนายก":"Central",
    "ชลบุรี":"East","ระยอง":"East","จันทบุรี":"East","ตราด":"East",
    "ฉะเชิงเทรา":"East","ปราจีนบุรี":"East","สระแก้ว":"East",
    "กาญจนบุรี":"West","ราชบุรี":"West",
    "เพชรบุรี":"West","ประจวบคีรีขันธ์":"West",
    "นครราชสีมา":"Northeast","บุรีรัมย์":"Northeast","สุรินทร์":"Northeast",
    "ศรีสะเกษ":"Northeast","อุบลราชธานี":"Northeast","ยโสธร":"Northeast",
    "ชัยภูมิ":"Northeast","อำนาจเจริญ":"Northeast","หนองบัวลำภู":"Northeast",
    "ขอนแก่น":"Northeast","อุดรธานี":"Northeast","เลย":"Northeast",
    "หนองคาย":"Northeast","มหาสารคาม":"Northeast","ร้อยเอ็ด":"Northeast",
    "กาฬสินธุ์":"Northeast","สกลนคร":"Northeast","นครพนม":"Northeast",
    "มุกดาหาร":"Northeast","บึงกาฬ":"Northeast",
    "ชุมพร":"South","ระนอง":"South","สุราษฎร์ธานี":"South",
    "นครศรีธรรมราช":"South","พังงา":"South","ภูเก็ต":"South",
    "กระบี่":"South","ตรัง":"South","พัทลุง":"South",
    "สงขลา":"South","สตูล":"South","ปัตตานี":"South",
    "ยะลา":"South","นราธิวาส":"South"
}

thai_to_eng = {

    # North
    "เชียงใหม่": "Chiang Mai",
    "เชียงราย": "Chiang Rai",
    "ลำปาง": "Lampang",
    "ลำพูน": "Lamphun",
    "แม่ฮ่องสอน": "Mae Hong Son",
    "น่าน": "Nan",
    "พะเยา": "Phayao",
    "แพร่": "Phrae",
    "อุตรดิตถ์": "Uttaradit",
    "กำแพงเพชร": "Kamphaeng Phet",
    "ตาก": "Tak",
    "พิษณุโลก": "Phitsanulok",
    "พิจิตร": "Phichit",
    "เพชรบูรณ์": "Phetchabun",
    "สุโขทัย": "Sukhothai",
    "อุทัยธานี": "Uthai Thani",
    "นครสวรรค์": "Nakhon Sawan",

    # Central
    "กรุงเทพมหานคร": "Bangkok Metropolis",
    "นนทบุรี": "Nonthaburi",
    "ปทุมธานี": "Pathum Thani",
    "สมุทรปราการ": "Samut Prakan",
    "สมุทรสาคร": "Samut Sakhon",
    "สมุทรสงคราม": "Samut Songkhram",
    "พระนครศรีอยุธยา": "Phra Nakhon Si Ayutthaya",
    "อ่างทอง": "Ang Thong",
    "ลพบุรี": "Lop Buri",
    "สระบุรี": "Saraburi",
    "ชัยนาท": "Chai Nat",
    "นครปฐม": "Nakhon Pathom",
    "สุพรรณบุรี": "Suphan Buri",
    "สิงห์บุรี": "Sing Buri",
    "นครนายก": "Nakhon Nayok",

    # East
    "ชลบุรี": "Chon Buri",
    "ระยอง": "Rayong",
    "จันทบุรี": "Chanthaburi",
    "ตราด": "Trat",
    "ฉะเชิงเทรา": "Chachoengsao",
    "ปราจีนบุรี": "Prachin Buri",
    "สระแก้ว": "Sa Kaeo",

    # West
    "กาญจนบุรี": "Kanchanaburi",
    "ราชบุรี": "Ratchaburi",
    "เพชรบุรี": "Phetchaburi",
    "ประจวบคีรีขันธ์": "Prachuap Khiri Khan",

    # Northeast
    "นครราชสีมา": "Nakhon Ratchasima",
    "บุรีรัมย์": "Buri Ram",
    "สุรินทร์": "Surin",
    "ศรีสะเกษ": "Si Sa Ket",
    "อุบลราชธานี": "Ubon Ratchathani",
    "ยโสธร": "Yasothon",
    "ชัยภูมิ": "Chaiyaphum",
    "อำนาจเจริญ": "Amnat Charoen",
    "หนองบัวลำภู": "Nong Bua Lam Phu",
    "ขอนแก่น": "Khon Kaen",
    "อุดรธานี": "Udon Thani",
    "เลย": "Loei",
    "หนองคาย": "Nong Khai",
    "มหาสารคาม": "Maha Sarakham",
    "ร้อยเอ็ด": "Roi Et",
    "กาฬสินธุ์": "Kalasin",
    "สกลนคร": "Sakon Nakhon",
    "นครพนม": "Nakhon Phanom",
    "มุกดาหาร": "Mukdahan",
    "บึงกาฬ": "Bueng Kan",

    # South
    "ชุมพร": "Chumphon",
    "ระนอง": "Ranong",
    "สุราษฎร์ธานี": "Surat Thani",
    "นครศรีธรรมราช": "Nakhon Si Thammarat",
    "พังงา": "Phangnga",
    "ภูเก็ต": "Phuket",
    "กระบี่": "Krabi",
    "ตรัง": "Trang",
    "พัทลุง": "Phatthalung",
    "สงขลา": "Songkhla",
    "สตูล": "Satun",
    "ปัตตานี": "Pattani",
    "ยะลา": "Yala",
    "นราธิวาส": "Narathiwat"
}

years = ["2563","2564","2565","2566","2567"]
response1 = requests.get(url)
response2 = requests.get(pop_url)
excel_file = BytesIO(response1.content)
pop_excel = BytesIO(response2.content)
region_area = {
    "Northeast": 168854,
    "North": 93691,
    "Central": 91795,
    "South": 70715,
    "West": 53679,
    "East": 34381
}

def load_waste(year):
    sheet = f"ขยะมูลฝอยรายจังหวัด{year}"
    df = pd.read_excel(excel_file, sheet_name=sheet)
    df["Province"] = df["Province"].astype(str).str.strip()
    df["Waste_generated"] = (
        df["Waste_generated"]
        .astype(str)
        .str.replace(",", "")
        .astype(float))
    df["Region"] = df["Province"].map(region_map)
    return df

def load_population(year):
    pop_sheet = f"จำนวนประชากร{year}"
    df_pop = pd.read_excel(pop_excel, sheet_name=pop_sheet)
    df_pop.columns = df_pop.columns.str.strip()
    df_pop["ชื่อจังหวัด"] = (
        df_pop["ชื่อจังหวัด"]
        .astype(str)
        .str.replace("จังหวัด", "")
        .str.strip())
    df_pop = df_pop.rename(columns={
        "ชื่อจังหวัด": "Province",
        "จำนวนประชากรทั้งหมด": "Population"})
    return df_pop[["Province", "Population"]]

def clean_value(x):
    if pd.isna(x) or str(x).strip() == "-":
        return 0
    return float(str(x).replace(",", "").strip())

def load_and_sum(xls, sheet_prefix, real_col_name, new_col_name):
    data = []
    for year in range(2563, 2568):
        sheet_name = f"{sheet_prefix}{year}"
        df = pd.read_excel(xls, sheet_name=sheet_name)
        print(f"\n===== {sheet_name} =====")
        print(df.head())  # 👈 print ดูก่อน
        # clean column name กันพัง
        df.columns = df.columns.str.strip()
        # clean ค่า
        df[real_col_name] = df[real_col_name].apply(clean_value)
        total = df[real_col_name].sum()
        data.append({
            "Year": year,
            new_col_name: total
        })
    return pd.DataFrame(data)

# =========================================================
# 3. CALCULATIONS
# =========================================================
xls_legal = pd.ExcelFile(url_legal)
xls_illegal = pd.ExcelFile(url_illegal)
xls_recovery = pd.ExcelFile(url_recovery)
df_legal = load_and_sum(xls_legal,"waste_disposal_legal","Waste_disposal_legal","Legal")
df_illegal = load_and_sum(xls_illegal,"waste_disposal_illegal","Waste_disposal_illegal","Illegal")
df_recovery = load_and_sum(xls_recovery,"province_waste_recovered_","Waste_recovered","Recovered")

all_region = []
all_area = []
all_capita = []
totals = []
for y in years:
    df = load_waste(y)
    # ---------- waste by region ----------
    region_sum = df.groupby("Region")["Waste_generated"].sum().reset_index()
    region_sum["Waste"] = region_sum["Waste_generated"] * 365 / 1_000_000
    region_sum["Waste_year"] = region_sum["Waste_generated"] * 365
    region_sum["Area_km2"] = region_sum["Region"].map(region_area)
    region_sum["Waste_per_area"] = (region_sum["Waste_year"] / region_sum["Area_km2"]).round(2)
    for _, r in region_sum.iterrows():
        all_region.append({
            "Year": y,
            "Region": r["Region"],
            "Waste": r["Waste"]
        })
        all_area.append({
            "Year": y,
            "Region": r["Region"],
            "Waste_per_area": r["Waste_per_area"]
        })
    # ---------- national trend ----------
    total_day = df["Waste_generated"].sum()
    total_year = total_day * 365 / 1_000_000
    totals.append(total_year)

    # ---------- waste per capita ----------
    df_pop = load_population(y)
    df_merge = pd.merge(df, df_pop, on="Province", how="inner")
    region_capita = df_merge.groupby("Region").agg({
        "Waste_generated": "sum",
        "Population": "sum"}).reset_index()

    region_capita["Waste_per_capita"] = (
        region_capita["Waste_generated"] * 1000
        / region_capita["Population"]).round(2)

    for _, r in region_capita.iterrows():
        all_capita.append({
            "Year": y,
            "Region": r["Region"],
            "Waste_per_capita": r["Waste_per_capita"]})

df_region = pd.DataFrame(all_region)
df_area = pd.DataFrame(all_area)
df_capita = pd.DataFrame(all_capita)
trend_df = pd.DataFrame({"Year": years,"MSW": totals})

xls = pd.ExcelFile(url_disposal_cap)
sheets = {
    "2563": "สถานที่กำจัดขยะ2563",
    "2567": "สถานที่กำจัดขยะ2567"}
map_dfs = {}
for year, sheet_name in sheets.items():
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()
    df["Province"] = (df["Province"].astype(str).str.replace("จังหวัด", "").str.replace("จ.", "").str.strip())
    df["Waste Quatities (ton/day)"] = (df["Waste Quatities (ton/day)"].astype(str).str.replace(",", "").astype(float))
    df_province = (df.groupby("Province")["Waste Quatities (ton/day)"].sum().reset_index())
    df_province = df_province.rename(columns={"Waste Quatities (ton/day)": "Waste_ton_per_day"})
    df_province["Province_EN"] = df_province["Province"].map(thai_to_eng)
    df_province["Waste_ton_per_year"] = df_province["Waste_ton_per_day"] * 365
    map_dfs[year] = df_province
gdf["Province_EN"] = gdf["name"]
vmin = 0
vmax = max(
    map_dfs["2563"]["Waste_ton_per_year"].max(),
    map_dfs["2567"]["Waste_ton_per_year"].max())

residual_sheets = {"2563": "ขยะตกค้าง2563","2567": "ขยะตกค้าง2567"}
map_residual = {}
for year, sheet_name in residual_sheets.items():
    df_res = pd.read_excel(xls_residual, sheet_name=sheet_name)
    df_res.columns = df_res.columns.str.strip()
    df_res["Province"] = (df_res["Province"].astype(str).str.replace("จังหวัด", "").str.replace("จ.", "").str.strip())
    df_res["Residual_waste"] = (df_res["Residual_waste"].astype(str).str.replace(",", "").str.replace("-", "0").str.strip())
    df_res["Residual_waste"] = pd.to_numeric(df_res["Residual_waste"], errors="coerce").fillna(0)
    df_res["Province_EN"] = df_res["Province"].map(thai_to_eng)
    df_res_province = (df_res.groupby("Province_EN")["Residual_waste"].sum().reset_index())
    map_residual[year] = df_res_province

vmin_res = 0
vmax_res = max(map_residual["2563"]["Residual_waste"].max(),map_residual["2567"]["Residual_waste"].max())

years_trend = ["2563","2564","2565","2566","2567"]
map_residual_trend = {}
for yr in years_trend:
    df_tr = pd.read_excel(xls_residual, sheet_name=f"ขยะตกค้าง{yr}")
    df_tr.columns = df_tr.columns.str.strip()
    df_tr["Province"] = (df_tr["Province"].astype(str).str.replace("จังหวัด", "").str.replace("จ.", "").str.strip())
    df_tr["Residual_waste"] = (df_tr["Residual_waste"].astype(str).str.replace(",", "").str.replace("-", "0").str.strip())
    df_tr["Residual_waste"] = pd.to_numeric(df_tr["Residual_waste"], errors="coerce").fillna(0)
    df_tr_province = ( df_tr.groupby("Province")["Residual_waste"].sum().reset_index())
    map_residual_trend[yr] = df_tr_province
list_trend_all = []
for yr, df_tr in map_residual_trend.items():
    temp_df = df_tr.copy()
    temp_df["Year"] = int(yr)
    list_trend_all.append(temp_df)
df_residual_all_years = pd.concat(list_trend_all, ignore_index=True)
df_avg_residual = (df_residual_all_years.groupby("Province")["Residual_waste"].mean().reset_index())
top5_provinces_th = (df_avg_residual.sort_values(by="Residual_waste", ascending=False).head(5)["Province"].tolist())
# print("Top 5 Provinces (TH):", top5_provinces_th)
df_top5_trend = df_residual_all_years[df_residual_all_years["Province"].isin(top5_provinces_th)]
bottom5_provinces_th = (df_avg_residual.sort_values(by="Residual_waste", ascending=True)  .head(5)["Province"].tolist())
# print("Bottom 5 Provinces (TH):", bottom5_provinces_th)
df_bottom5_trend = df_residual_all_years[df_residual_all_years["Province"].isin(bottom5_provinces_th)]

years = ["2563", "2564", "2565", "2566", "2567"]
sheet_names = {y: f"สถานที่กำจัดขยะ{y}" for y in years}
all_data = []
for y in years:
    sheet = sheet_names[y]
    if sheet not in xls.sheet_names:
        continue
    df = pd.read_excel(xls, sheet_name=sheet)
    if "Waste Quatities (ton/day)" not in df.columns:
        continue
    df["Waste Quatities (ton/day)"] = pd.to_numeric(df["Waste Quatities (ton/day)"], errors="coerce")
    df_group = df.groupby("Province", as_index=False)["Waste Quatities (ton/day)"].sum()
    df_group["Year"] = int(y)
    all_data.append(df_group)
df_all = pd.concat(all_data, ignore_index=True)
df_pivot = df_all.pivot_table(index="Province",columns="Year",values="Waste Quatities (ton/day)")
df_pivot["avg"] = df_pivot.mean(axis=1, skipna=True)
top5_provinces_wq = df_pivot["avg"].nlargest(5).index.tolist()
# print("Top 5 Provinces:")
df_top5_trend_wq = df_all[df_all["Province"].isin(top5_provinces_wq)]
bottom5_provinces_wq = df_pivot["avg"].nsmallest(5).index.tolist()
df_bottom5_trend_wq = df_all[df_all["Province"].isin(bottom5_provinces_wq)]

all_data_site = []
for y in years:
    sheet = sheet_names[y]
    if sheet not in xls.sheet_names:
        continue
    df = pd.read_excel(xls, sheet_name=sheet)
    df["Province"] = ( df["Province"].astype(str).str.replace("จังหวัด", "").str.replace("จ.", "").str.strip())
    df_group = df.groupby("Province").size().reset_index(name="Disposal_sites")
    df_group["Year"] = int(y)
    all_data_site.append(df_group)
df_sites_all = pd.concat(all_data_site, ignore_index=True)
df_residual_year = (df_residual_all_years.groupby("Year")["Residual_waste"].sum().reset_index())
df_sites_year = ( df_sites_all.groupby("Year")["Disposal_sites"].sum().reset_index())
df_compare = pd.merge(df_residual_year,df_sites_year,on="Year",how="outer").sort_values("Year")


df_factory2566 = pd.read_excel(xls_2566, sheet_name="จำนวนโรงงาน2566")
df_gpp2566 = pd.read_excel(xls_2566, sheet_name="GPP2566")
df_tourist2566 = pd.read_excel(xls_2566, sheet_name="จำนวนนักท่องเที่ยว2566")
df_population2566 = pd.read_excel(xls_2566, sheet_name="จำนวนประชากร2566")
df_waste2566 = pd.read_excel(xls_2566, sheet_name="ขยะมูลฝอยรายจังหวัด2566")
for df in [df_factory2566, df_gpp2566, df_tourist2566]:
    df.rename(columns={"ภูมิภาค": "Region2566"}, inplace=True)
for df in [df_factory2566, df_gpp2566, df_tourist2566]:
    df["Region2566"] = (df["Region2566"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True))
region_th_to_en_2566 = {"ภาคเหนือ": "North","ภาคกลาง": "Central","ภาคตะวันออกเฉียงเหนือ": "Northeast","ภาคตะวันออก": "East","ภาคใต้": "South"}
for df in [df_factory2566, df_gpp2566, df_tourist2566]:
    df["Region2566"] = df["Region2566"].map(region_th_to_en_2566)
def clean_province(x):
    if pd.isna(x):
        return x
    return str(x).replace("จังหวัด", "").strip()
df_population2566["Province_clean2566"] = df_population2566["ชื่อจังหวัด"].apply(clean_province)
df_waste2566["Province_clean2566"] = df_waste2566["Province"].apply(clean_province)
df_population2566["Region2566"] = df_population2566["Province_clean2566"].map(region_map)
df_waste2566["Region2566"] = df_waste2566["Province_clean2566"].map(region_map)
df_waste2566["Waste_generated"] = (df_waste2566["Waste_generated"].astype(str).str.replace(",", ""))
df_waste2566["Waste_generated"] = pd.to_numeric(df_waste2566["Waste_generated"].astype(str).str.replace(",", ""),errors="coerce")
df_population_region2566 = df_population2566.groupby("Region2566", as_index=False)["จำนวนประชากรทั้งหมด"].sum()
df_waste_region2566 = df_waste2566.groupby("Region2566", as_index=False)["Waste_generated"].sum()
df_plot2566 = df_factory2566.merge(df_waste_region2566, on="Region2566")
df_plot2566 = df_plot2566.merge(df_gpp2566, on="Region2566")
df_plot2566 = df_plot2566.merge(df_tourist2566, on="Region2566")
df_plot2566 = df_plot2566.merge(df_population_region2566, on="Region2566")
df_plot2566.rename(columns={"จำนวนโรงงาน": "Factory2566","GPP": "GPP2566","Foreigners": "Tourist2566","จำนวนประชากรทั้งหมด": "Population2566","Waste_generated": "Waste2566"}, inplace=True)
cols = ["Factory2566","GPP2566","Tourist2566","Population2566","Waste2566"]
for c in cols:
    df_plot2566[c] = (df_plot2566[c].astype(str).str.replace(",", "").str.strip())
    df_plot2566[c] = pd.to_numeric(df_plot2566[c], errors='coerce')
order = ["North","Central","Northeast","East","South"]
df_plot2566["Region2566"] = pd.Categorical(df_plot2566["Region2566"], categories=order, ordered=True)
df_plot2566 = df_plot2566.sort_values("Region2566")

years_scatter = [2563, 2564, 2565, 2566, 2567]
def load_residual_scatter(year):
    sheet = f"ขยะตกค้าง{year}"
    df_residual_scatter = pd.read_excel(xls_residual, sheet_name=sheet)
    df_residual_scatter["Province"] = df_residual_scatter["Province"].astype(str).str.strip()
    df_residual_scatter["Residual_waste"] = pd.to_numeric(df_residual_scatter["Residual_waste"].astype(str).str.replace(",", "").str.strip(),errors="coerce")
    df_residual_scatter["Region"] = df_residual_scatter["Province"].map(region_map)
    return df_residual_scatter
all_data_scatter = []
for y in years_scatter:
    df_waste_scatter = load_waste(y)
    df_residual_scatter = load_residual_scatter(y)
    regions_scatter = df_waste_scatter["Region"].dropna().unique()
    for r in regions_scatter:
        df_w_n_scatter = df_waste_scatter[df_waste_scatter["Region"] == r]
        df_r_n_scatter = df_residual_scatter[df_residual_scatter["Region"] == r]
        total_waste_scatter = df_w_n_scatter["Waste_generated"].sum()
        total_residual_scatter = df_r_n_scatter["Residual_waste"].sum()
        all_data_scatter.append({"Year": y,"Region": r,"Residual": total_residual_scatter,"Generated": total_waste_scatter})

# =========================================================
# 4 PLOTTING
# =========================================================
# ---------------------------------------------------------
# GRAPH 1 : ASEAN Municipal Solid Waste per Capita
# ---------------------------------------------------------
df_plot = df_asean.sort_values("Per_Capita_MSW_kg_per_day", ascending=False)
colors = [ "#2a6fdb" if c == "Thailand" else "#cfcfcf" for c in df_plot["Country"]]
plt.figure(figsize=(12,7))
ax = sns.barplot(data=df_plot, x="Country",y="Per_Capita_MSW_kg_per_day",palette=colors)
plt.title("ASEAN Municipal Solid Waste per Capita", fontsize=18, weight="bold")
plt.ylabel("kg per person per day")
plt.xticks(rotation=40)
for p in ax.patches:
    h = p.get_height()
    ax.text(p.get_x() + p.get_width()/2,h + 0.03,f"{h:.2f}",ha="center",fontsize=11)
plt.figtext(0.99,0.01,"Source: United Nations Environment Programme (2017)",ha="right",fontsize=10,style="italic")
plt.tight_layout()
# path1 = os.path.expanduser("~/Desktop/asean_msw_per_capita.png")
# plt.savefig(path1, dpi=1800)
plt.close()

# ---------------------------------------------------------
# GRAPH 2 : ASEAN Annual Municipal Solid Waste
# ---------------------------------------------------------
df_plot2 = df_asean.sort_values("Annual_MSW_million_ton", ascending=False)
colors2 = [ "#2a6fdb" if c == "Thailand" else "#cfcfcf"for c in df_plot2["Country"]]
plt.figure(figsize=(12,7))
ax = sns.barplot(data=df_plot2,x="Country",y="Annual_MSW_million_ton",palette=colors2)
plt.title("ASEAN Annual Municipal Solid Waste", fontsize=18, weight="bold")
plt.ylabel("Million tons per year")
plt.xticks(rotation=40)
for p in ax.patches:
    h = p.get_height()
    ax.text(p.get_x() + p.get_width()/2,h + 0.2,f"{h:.2f}",ha="center",fontsize=11)
plt.figtext(0.99,0.01,"Source: United Nations Environment Programme (2017)",ha="right",fontsize=10,style="italic")
plt.tight_layout()
# path2 = os.path.expanduser("~/Desktop/asean_annual_msw.png")
# plt.savefig(path2, dpi=1800)
plt.close()

# ---------------------------------------------------------
# GRAPH 3 : Thailand Municipal Solid Waste Trend
# ---------------------------------------------------------
plt.figure(figsize=(10,6))
sns.lineplot(data=trend_df,x="Year",y="MSW",marker="o",linewidth=3,color="#2a6fdb")
plt.title("Trend of Municipal Solid Waste in Thailand (2563–2567)",fontsize=16,weight="bold")
plt.ylabel("Million Tons per Year")
plt.xlabel("Year")
for x,y in zip(trend_df["Year"], trend_df["MSW"]):
    plt.text(x, y+0.2, f"{y:.2f}", ha="center")
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout()
# path3 = os.path.expanduser("~/Desktop/thailand_msw_trend.png")
# plt.savefig(path3, dpi=1800)
plt.close()

# ---------------------------------------------------------
# GRAPH 4 : Municipal Solid Waste by Region (2563-2567)
# ---------------------------------------------------------
plt.figure(figsize=(14,7))
ax = sns.barplot(data=df_region,x="Region",y="Waste",hue="Year",palette="viridis")
plt.title("Municipal Solid Waste by Region in Thailand (2563–2567)",fontsize=18,weight="bold")
plt.ylabel("Million tons per year")
plt.xlabel("Region")
for container in ax.containers:
    ax.bar_label(container, fmt="%.2f", fontsize=9)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout()
# path4 = os.path.expanduser("~/Desktop/thailand_msw_region.png")
# plt.savefig(path4, dpi=300,bbox_inches="tight",facecolor="white")
plt.close()

# ---------------------------------------------------------
# GRAPH 5 : Waste per Capita by Region (2563-2567)
# ---------------------------------------------------------
plt.figure(figsize=(14,7))
ax = sns.barplot(data=df_capita,x="Region",y="Waste_per_capita",hue="Year",palette="viridis")
plt.title("Municipal Solid Waste per Capita by Region in Thailand (2563–2567)",fontsize=18,weight="bold")
plt.ylabel("kg per person per day")
plt.xlabel("Region")
for container in ax.containers:
    ax.bar_label(container, fmt="%.2f", fontsize=9)
plt.figtext(0.99,0.01, "ที่มา: กรมควบคุมมลพิษ และสำนักบริหารการทะเบียน",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout()
# path5 = os.path.expanduser("~/Desktop/thailand_waste_per_capita_region.png")
# plt.savefig(path5,dpi=300,bbox_inches="tight",facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 6 : Municipal Solid Waste per Area by Region in Thailand (2563-2567)
# ---------------------------------------------------------
plt.figure(figsize=(14,7))
ax = sns.barplot(data=df_area,x="Region",y="Waste_per_area",hue="Year",palette="viridis")
plt.title("Municipal Solid Waste per Area by Region in Thailand (2563–2567)",fontsize=18,weight="bold")
plt.ylabel("ton per km² per year")
plt.xlabel("Region")
for container in ax.containers:
    ax.bar_label(container, fmt="%.1f", fontsize=9)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ และกรมอุตอนิยมวิทยา",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout()
# path6 = os.path.expanduser("~/Desktop/thailand_waste_per_area.png")
# plt.savefig(path6,dpi=300,bbox_inches="tight",facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 7 : Top 5 Provinces Waste Trend
# ---------------------------------------------------------
all_province = []
for y in years:
    df = load_waste(y)
    df["Waste_year"] = df["Waste_generated"] * 365
    for _, row in df.iterrows():
        all_province.append({"Year": y,"Province": row["Province"],"Waste": row["Waste_year"]})
df_province = pd.DataFrame(all_province)
top5 = (df_province.groupby("Province")["Waste"].mean().sort_values(ascending=False).head(5).index)
df_top5 = df_province[df_province["Province"].isin(top5)]
plt.figure(figsize=(14,8))
ax = sns.lineplot(data=df_top5,x="Year",y="Waste",hue="Province",marker="o",linewidth=2.5)
legend = ax.get_legend()
for text in legend.get_texts():
    text.set_fontproperties(thai_font)
legend.get_title().set_fontproperties(thai_font)
plt.title("Top 5 Provinces with Highest Municipal Solid Waste (Average 2563–2567)",fontsize=18,weight="bold")
plt.ylabel("Waste (ton/year)")
plt.xlabel("Year")
for line in ax.lines:
    x = line.get_xdata()
    y = line.get_ydata()
    for i in range(len(x)):
        ax.text(x[i],y[i],f"{y[i]:,.0f}",ha="center",va="bottom",fontsize=9)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout()
# path7 = os.path.expanduser("~/Desktop/top5_waste_trend.png")
# plt.savefig(path7,dpi=300,bbox_inches="tight",facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 8 : Bottom 5 Provinces Waste Trend
# ---------------------------------------------------------
bottom5 = (df_province.groupby("Province")["Waste"].mean().sort_values(ascending=True).head(5).index)
df_bottom5 = df_province[df_province["Province"].isin(bottom5)]
plt.figure(figsize=(14,8))
ax = sns.lineplot(data=df_bottom5,x="Year",y="Waste",hue="Province",marker="o",linewidth=2.5)
legend = ax.get_legend()
for text in legend.get_texts():
    text.set_fontproperties(thai_font)
legend.get_title().set_fontproperties(thai_font)
plt.title("Top 5 Provinces with Lowest Municipal Solid Waste (Average 2563–2567)",fontsize=18,weight="bold")
plt.ylabel("Waste (ton/year)")
plt.xlabel("Year")
for line in ax.lines:
    x = line.get_xdata()
    y = line.get_ydata()
    for i in range(len(x)):
        ax.text(x[i],y[i],f"{y[i]:,.0f}",ha="center",va="bottom",fontsize=9)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout()
# path8 = os.path.expanduser("~/Desktop/bottom5_waste_trend.png")
# plt.savefig(path8, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 9 : Annual Waste Disposal Capacity by Province (ton/year)
# ---------------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(18, 8))
axes = axes.flatten()
minx, miny, maxx, maxy = gdf.total_bounds
for ax, (year, df_map) in zip(axes, map_dfs.items()):
    merged = gdf.merge(df_map, on="Province_EN", how="left")
    merged.plot(column="Waste_ton_per_year",cmap="OrRd",linewidth=0.5,edgecolor="black",vmin=vmin,vmax=vmax,ax=ax,legend=False)
    ax.set_xlim(minx, maxx)
    ax.set_ylim(miny, maxy)
    ax.set_aspect('equal')
    ax.set_title(f"{year}\nAnnual Waste Disposal Capacity (ton/year)", fontsize=14)
    ax.axis("off")
cax = fig.add_axes([0.92, 0.15, 0.02, 0.7])  
from matplotlib.cm import ScalarMappable
sm = ScalarMappable(cmap="OrRd")
sm.set_clim(vmin, vmax)
cbar = fig.colorbar(sm, cax=cax)
cbar.set_label("ton/year")
plt.tight_layout(rect=[0, 0, 0.9, 1])  
# path9 = os.path.expanduser("~/Desktop/annual_waste_disposal_capacity.png")
# plt.savefig(path9, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 10 : Residual Waste by Province (ton)
# ---------------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(18, 8))
axes = axes.flatten()
minx, miny, maxx, maxy = gdf.total_bounds
for ax, (year, df_map) in zip(axes, map_residual.items()):
    merged = gdf.merge(df_map, on="Province_EN", how="left")
    merged.plot(column="Residual_waste",cmap="OrRd",linewidth=0.5,edgecolor="black",vmin=vmin_res, vmax=vmax_res,ax=ax,legend=False )
    ax.set_xlim(minx, maxx)
    ax.set_ylim(miny, maxy)
    ax.set_aspect('equal')
    ax.set_title(f"{year}\nResidual Waste by Province (ton)", fontsize=14)
    ax.axis("off")
cax = fig.add_axes([0.92, 0.15, 0.02, 0.7])
from matplotlib.cm import ScalarMappable
sm_res = ScalarMappable(cmap="OrRd")
sm_res.set_clim(vmin_res, vmax_res)
cbar = fig.colorbar(sm_res, cax=cax)
cbar.set_label("ton")
plt.tight_layout(rect=[0, 0, 0.9, 1])
# path10 = os.path.expanduser("~/Desktop/residual_waste_map.png")
# plt.savefig(path10, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 11 : Top 5 Provinces: Residual Waste Trend (2563–2567)
# ---------------------------------------------------------
plt.figure(figsize=(12,7))
for prov in top5_provinces_th:
    df_plot = df_top5_trend[df_top5_trend["Province"] == prov].sort_values("Year")
    plt.plot(df_plot["Year"],df_plot["Residual_waste"],marker="o",label=prov)
plt.xticks([2563, 2564, 2565, 2566, 2567])
plt.title("Top 5 Provinces: Residual Waste Trend (2563–2567)", fontsize=14)
plt.xlabel("Year")
plt.ylabel("Residual Waste (ton)")
plt.legend(prop=thai_font)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.grid(True)
plt.tight_layout()
# path11 = os.path.expanduser("~/Desktop/residual_waste_trend_top5.png")
# plt.savefig(path11, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 12 : Bottom 5 Provinces: Residual Waste Trend (2563–2567)
# ---------------------------------------------------------
plt.figure(figsize=(12,7))
for prov in bottom5_provinces_th:
    df_plot = df_bottom5_trend[
        df_bottom5_trend["Province"] == prov
    ].sort_values("Year")
    plt.plot(df_plot["Year"],df_plot["Residual_waste"],marker="o",label=prov)
plt.title("Bottom 5 Provinces: Residual Waste Trend (2563–2567)", fontsize=14)
plt.xlabel("Year")
plt.ylabel("Residual Waste (ton)")
plt.xticks([2563, 2564, 2565, 2566, 2567])
plt.legend(prop=thai_font)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.grid(True)
plt.tight_layout()
# path12 = os.path.expanduser("~/Desktop/residual_waste_trend_bottom5.png")
# plt.savefig(path12, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 13 : Top 5 Provinces: Waste Disposal Capacity Trend (2563–2567)
# ---------------------------------------------------------
plt.figure(figsize=(12,7))
for prov in top5_provinces_wq:
    df_plot = df_top5_trend_wq[df_top5_trend_wq["Province"] == prov].sort_values("Year")
    plt.plot(df_plot["Year"],df_plot["Waste Quatities (ton/day)"],marker="o",linestyle="--",label=prov)
plt.title("Top 5 Provinces: Waste Disposal Capacity Trend (2563–2567)", fontsize=14)
plt.xlabel("Year")
plt.ylabel("Waste Capacity (ton/day)")
plt.xticks([2563, 2564, 2565, 2566, 2567])
plt.legend(prop=thai_font)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.grid(True)
plt.tight_layout()
# path13 = os.path.expanduser("~/Desktop/top5_disposal_capacity_trend.png")
# plt.savefig(path13, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 14 : Bottom 5 Provinces: Waste Disposal Capacity Trend (2563–2567)
# ---------------------------------------------------------
plt.figure(figsize=(12,7))
for prov in bottom5_provinces_wq:
    df_plot = df_bottom5_trend_wq[df_bottom5_trend_wq["Province"] == prov].sort_values("Year")
    plt.plot(df_plot["Year"],df_plot["Waste Quatities (ton/day)"],marker="o",linestyle="--",label=prov)
plt.title("Bottom 5 Provinces: Waste Disposal Capacity Trend (2563–2567)", fontsize=14)
plt.xlabel("Year")
plt.ylabel("Waste Capacity (ton/day)")
plt.xticks([2563, 2564, 2565, 2566, 2567])
plt.legend(prop=thai_font)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.grid(True)
plt.tight_layout()
# path14 = os.path.expanduser("~/Desktop/bottom5_disposal_capacity_trend.png")
# plt.savefig(path14, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 15 : Residual Waste vs Disposal Sites (2563–2567)
# ---------------------------------------------------------
fig, ax1 = plt.subplots(figsize=(12,7))
ax1.plot(df_compare["Year"],df_compare["Residual_waste"],marker="o",linewidth=2,color="tab:pink",label="Residual Waste")
ax1.set_xlabel("Year")
ax1.set_ylabel("Residual Waste (ton)")
ax1.set_xticks([2563, 2564, 2565, 2566, 2567])
for x, y in zip(df_compare["Year"], df_compare["Residual_waste"]):
    ax1.annotate(f"{y:,.0f}",(x, y),textcoords="offset points",xytext=(0,8),  ha="center",fontsize=9)
ax2 = ax1.twinx()
ax2.plot(df_compare["Year"],df_compare["Disposal_sites"],marker="s",linewidth=2,color="tab:blue",label="Disposal Sites")
ax2.set_ylabel("Number of Disposal Sites")
for x, y in zip(df_compare["Year"], df_compare["Disposal_sites"]):
    ax2.annotate(f"{y}",(x, y),textcoords="offset points",xytext=(0,-12),  ha="center",fontsize=9)
lines_1, labels_1 = ax1.get_legend_handles_labels()
lines_2, labels_2 = ax2.get_legend_handles_labels()
ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper left")
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.title("Residual Waste vs Disposal Sites (2563–2567)")
plt.grid(True)
plt.tight_layout()
# path15 = os.path.expanduser("~/Desktop/dual_axis_waste_vs_sites.png")
# plt.savefig(path15, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 16 : GPP, Factory, Tourist vs Waste by Region (2566)
# ---------------------------------------------------------
def plot_dual(ax, x, y1, y2, label1, label2, color1):
    ax.plot(x, y1, marker="o", color=color1, label=label1)
    ax.set_ylabel(label1)
    y1_min = y1.min()
    y1_max = y1.max()
    ax.set_ylim(max(0, y1_min * 0.8), y1_max * 1.1)
    ax2 = ax.twinx()
    ax2.plot(x, y2, marker="s", color="pink", label=label2)
    ax2.set_ylabel(label2)
    y2_min = y2.min()
    y2_max = y2.max()
    ax2.set_ylim(max(0, y2_min * 0.8), y2_max * 1.1)
    corr = y1.corr(y2)
    ax.set_title(f"{label1} vs {label2} (r = {corr:.2f})")
    lines_1, labels_1 = ax.get_legend_handles_labels()
    lines_2, labels_2 = ax2.get_legend_handles_labels()
    ax.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper left")
fig, axes = plt.subplots(2, 2, figsize=(16,10))
plot_dual(axes[0,0],df_plot2566["Region2566"],df_plot2566["Population2566"],df_plot2566["Waste2566"], "Population","Waste",color1="blue")
plot_dual(axes[0,1],df_plot2566["Region2566"],df_plot2566["GPP2566"],df_plot2566["Waste2566"],"GPP","Waste",color1="green")
plot_dual(axes[1,0],df_plot2566["Region2566"],df_plot2566["Factory2566"],df_plot2566["Waste2566"],"Factory","Waste",color1="purple")
plot_dual(axes[1,1],df_plot2566["Region2566"],df_plot2566["Tourist2566"],df_plot2566["Waste2566"],"Tourist","Waste",color1="orange")
plt.tight_layout()
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ สำนักงานสภาพัฒนาการเศรษฐกิจและสังคมแห่งชาติ กระทรวงการท่องเที่ยวและกีฬา และสำนักบริหารการทะเบียน",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
# path16 = os.path.expanduser("~/Desktop/Waste_vs_Economic_Factors_by_Region_(2566).png")
# plt.savefig(path16, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Graph 17 : GPP, Factory, Tourist vs Waste by Region (2566)
# ---------------------------------------------------------
df_all_scatter = pd.DataFrame(all_data_scatter)
ymin_scatter = df_all_scatter["Residual"].min()
ymax_scatter = df_all_scatter["Residual"].max()
ymin_scatter *= 0.95
ymax_scatter *= 1.05
regions_scatter = ["North", "Northeast", "Central", "East", "West", "South"]
colors_scatter = ["red", "blue", "green", "orange", "purple", "brown"]
fig_scatter, axes_scatter = plt.subplots(2, 3, figsize=(18,10))
axes_scatter = axes_scatter.flatten()
for i, region in enumerate(regions_scatter):
    ax_scatter = axes_scatter[i]
    df_region_scatter = df_all_scatter[df_all_scatter["Region"] == region]
    ax_scatter.scatter(df_region_scatter["Year"],df_region_scatter["Residual"],s=df_region_scatter["Generated"]/20,color=colors_scatter[i],alpha=0.6)
    for _, row in df_region_scatter.iterrows():
        ax_scatter.text(row["Year"],row["Residual"],f"{row['Generated']:.0f}",fontsize=9,ha='center')
    ax_scatter.set_ylim(ymin_scatter, ymax_scatter)
    ax_scatter.set_title(region)
    ax_scatter.set_xticks(df_region_scatter["Year"])
    ax_scatter.set_xlabel("Year")
    ax_scatter.set_ylabel("Residual")
    ax_scatter.grid()
fig_scatter.suptitle("Residual Waste vs Waste Generated by Region (2563–2567)\nBubble Size = Waste Generated",fontsize=16)
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
plt.tight_layout(rect=[0, 0, 1, 0.95])
path17 = os.path.expanduser("~/Desktop/Residual_Waste_vs_Waste_Generated_by_Region_(2563–2567).png")
plt.savefig(path17, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()

# ---------------------------------------------------------
# Table 18 : Regional Waste Dashboard Comparison (2563 vs 2567)
# ---------------------------------------------------------
df_scatter_full = pd.DataFrame(all_data_scatter)
df_capacity = df_all.copy()
df_capacity["Province"] = (df_capacity["Province"].astype(str).str.replace("จังหวัด", "").str.replace("จ.", "").str.strip())
df_capacity["Region"] = df_capacity["Province"].map(region_map)
df_capacity_region = (df_capacity.groupby(["Year", "Region"])["Waste Quatities (ton/day)"].sum().reset_index().rename(columns={"Waste Quatities (ton/day)": "Capacity"}))
df_dashboard = pd.merge(df_scatter_full, df_capacity_region,on=["Year", "Region"],how="left")
df_dashboard["Utilization (%)"] = (df_dashboard["Generated"] / df_dashboard["Capacity"] * 100)
def get_status(x):
    if pd.isna(x):
        return "-"
    elif x < 80:
        return "OK"
    elif x < 100:
        return "Tight"
    else:
        return "Overload"
df_dashboard["Status"] = df_dashboard["Utilization (%)"].apply(get_status)
df_dashboard["Generated"] = df_dashboard["Generated"].round(0)
df_dashboard["Residual"] = df_dashboard["Residual"].round(0)
df_dashboard["Capacity"] = df_dashboard["Capacity"].round(0)
df_dashboard["Utilization (%)"] = df_dashboard["Utilization (%)"].round(1)
def draw_table(ax, df, year):
    df = df[df["Year"] == year].copy()
    df = df[["Region","Generated","Capacity","Residual","Utilization (%)","Status"]]
    max_val = df["Utilization (%)"].max()
    min_val = df["Utilization (%)"].min()
    cell_colors = []
    for i in range(len(df)):
        row_colors = []
        for col in df.columns:
            if col == "Utilization (%)":
                val = df.iloc[i][col]
                if val == max_val:
                    row_colors.append("salmon")       
                elif val == min_val:
                    row_colors.append("lightgreen")  
                else:
                    row_colors.append("white")
            else:
                row_colors.append("white")
        cell_colors.append(row_colors)
    ax.axis('off')
    table = ax.table(cellText=df.values,colLabels=df.columns,cellColours=cell_colors,loc='upper center',cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.5)
    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(weight='bold')
            cell.set_facecolor("#eeeeee")
    ax.set_title(f"Year {year}", fontsize=14)
fig, axes = plt.subplots(2, 1, figsize=(12,10))
draw_table(axes[0], df_dashboard, 2563)
draw_table(axes[1], df_dashboard, 2567)
plt.suptitle("Regional Waste Dashboard Comparison",fontsize=16)
plt.tight_layout(rect=[0, 0, 1, 0.95])
path18 = os.path.expanduser("~/Desktop/Regional_Waste_Dashboard_Comparison.png")
plt.savefig(path18, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()


# ---------------------------------------------------------
# Graph 19 : Thailand Waste Management Proportion (2563–2567)
# ---------------------------------------------------------
df = df_legal.merge(df_illegal, on="Year").merge(df_recovery, on="Year")
df["Total"] = df["Legal"] + df["Illegal"] + df["Recovered"]
df["Legal_%"] = df["Legal"] / df["Total"] * 100
df["Illegal_%"] = df["Illegal"] / df["Total"] * 100
df["Recovered_%"] = df["Recovered"] / df["Total"] * 100
years = df["Year"].astype(int)  
legal = df["Legal_%"]
illegal = df["Illegal_%"]
recovered = df["Recovered_%"]
sns.set_style("whitegrid")  
plt.figure(figsize=(10,6))
bars_legal = plt.bar(years, legal, label="Legal Disposal", color="#4c72b0")      
bars_illegal = plt.bar(years, illegal, bottom=legal, label="Illegal Disposal", color="#ff7f0e") 
bars_recovered = plt.bar(years, recovered, bottom=legal+illegal, label="Recovered", color="#55a868") 
for i in range(len(years)):
    plt.text(years[i], legal[i]/2, f"{legal[i]:.1f}%", ha="center", va="center", color="white", fontsize=10)
    plt.text(years[i], legal[i]+illegal[i]/2, f"{illegal[i]:.1f}%", ha="center", va="center", color="white", fontsize=10)
    plt.text(years[i], legal[i]+illegal[i]+recovered[i]/2, f"{recovered[i]:.1f}%", ha="center", va="center", color="white", fontsize=10)
plt.xticks(years)  
plt.ylim(0, 100)
plt.ylabel("Percentage (%)")
plt.xlabel("Year")
plt.title("Thailand Waste Management Proportion (%)", fontsize=14, fontweight='bold')
plt.legend(frameon=True, framealpha=0.9)
plt.gca().set_facecolor("#e6f2ff")  
plt.grid(color='white', linestyle='--', linewidth=1)  
plt.tight_layout()
plt.figtext(0.99,0.01,"ที่มา: กรมควบคุมมลพิษ",ha="right",fontsize=10,style="italic",fontproperties=thai_font)
path19 = os.path.expanduser("~/Desktop/Thailand_Waste_Management_Proportion.png")
plt.savefig(path19, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()