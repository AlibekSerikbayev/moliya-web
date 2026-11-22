import io
import os
from datetime import datetime, date

import gspread
from google.oauth2.service_account import Credentials
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv

# ==================== SAHIFA SOZLAMALARI ====================
st.set_page_config(
    page_title="Moliya Boshqaruvi",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

load_dotenv()

# ==================== MOBIL VA DESKTOP TARTIBLI DIZAYN (CSS) ====================
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none !important;
    }

    .block-container {
        padding-top: 0.9rem !important;
        padding-bottom: 2.2rem !important;
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
        max-width: 1100px;
    }

    /* Telefonda inputga bosganda ekran zoom bo'lib ketmasligi uchun 16px */
    input, textarea, select, div[data-baseweb="select"] > div {
        font-size: 16px !important;
        min-height: 44px !important;
    }

    /* Tugmalarni tekis va qulay qilish */
    .stButton > button,
    .stDownloadButton > button,
    div[data-testid="stFormSubmitButton"] > button {
        width: 100% !important;
        min-height: 44px !important;
        height: 44px !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        padding: 0 10px !important;
        white-space: nowrap !important;
    }

    /* Kirim/Chiqim radio tugmalari yonma-yon */
    div[role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        gap: 10px !important;
    }
    div[role="radiogroup"] > label {
        flex: 1 !important;
        background: rgba(255, 255, 255, 0.06) !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
        border-radius: 10px !important;
        padding: 8px 12px !important;
        justify-content: center !important;
        cursor: pointer !important;
    }

    /* Tepada ixcham sarlavha va amallar qatori */
    .top-header-box {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 8px;
        margin-bottom: 10px;
    }
    .app-main-title {
        font-size: 1.45rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.2;
    }

    /* KPI kartalar */
    .kpi-wrapper {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 10px;
        margin-bottom: 12px;
    }
    .kpi-card {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        padding: 11px 13px;
    }
    .kpi-label {
        font-size: 12px;
        opacity: 0.8;
        margin-bottom: 3px;
    }
    .kpi-value {
        font-size: 19px;
        font-weight: 700;
    }
    .kpi-bal { color: #38BDF8; }
    .kpi-in  { color: #10B981; }
    .kpi-out { color: #F87171; }

    /* Telefonda (<= 640px) KPI va menyuni chiroyli tartiblash */
    @media (max-width: 640px) {
        .app-main-title {
            font-size: 1.25rem;
        }
        .kpi-wrapper {
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }
        .kpi-card.main-bal {
            grid-column: span 2;
            text-align: center;
        }
        .kpi-value {
            font-size: 17px;
        }
        h2, h3 {
            font-size: 1.12rem !important;
        }
    }

    /* Tranzaksiya kartochkalari */
    .tx-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 10px;
        padding: 10px 12px;
        margin-bottom: 8px;
    }
    .tx-left {
        display: flex;
        flex-direction: column;
        min-width: 0;
    }
    .tx-title {
        font-weight: 600;
        font-size: 15px;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .tx-meta {
        font-size: 12px;
        opacity: 0.75;
    }
    .tx-amt-in {
        color: #10B981;
        font-weight: 700;
        font-size: 14px;
        white-space: nowrap;
    }
    .tx-amt-out {
        color: #F87171;
        font-weight: 700;
        font-size: 14px;
        white-space: nowrap;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def refresh_data():
    st.cache_data.clear()


def has_secrets_file() -> bool:
    local_secret = os.path.join(os.getcwd(), ".streamlit", "secrets.toml")
    home_secret = os.path.join(os.path.expanduser("~"), ".streamlit", "secrets.toml")
    return os.path.exists(local_secret) or os.path.exists(home_secret)


def get_secret(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    if has_secrets_file():
        try:
            if key in st.secrets:
                return str(st.secrets[key]).strip()
        except Exception:
            pass
    val = os.getenv(key)
    return val.strip() if val and val.strip() else default


SPREADSHEET_ID = get_secret("SPREADSHEET_ID")
APP_PASSWORD = get_secret("APP_PASSWORD", "12345")
USER_ID = int(get_secret("DEFAULT_USER_ID", "1"))

if not SPREADSHEET_ID:
    st.error("❌ SPREADSHEET_ID topilmadi! Streamlit Secrets yoki .env faylini tekshiring.")
    st.stop()

# ==================== PAROL BILAN HIMOYA (LOGIN) ====================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("## 🔐 Shaxsiy Moliya Paneli")
    with st.form("login_form"):
        pwd = st.text_input("Parolni kiriting:", type="password")
        submitted = st.form_submit_button("Kirish")
        if submitted:
            if pwd == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Parol noto'g'ri!")
    st.stop()


# ==================== GOOGLE SHEETS ULANISH ====================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@st.cache_resource
def get_spreadsheet():
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
            gc = gspread.authorize(creds)
            return gc.open_by_key(SPREADSHEET_ID)
    except Exception:
        pass

    creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


sh = get_spreadsheet()


def to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    cleaned = str(val).replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def next_id(df: pd.DataFrame) -> int:
    if df.empty or "Id" not in df.columns:
        return 1
    ids = pd.to_numeric(df["Id"], errors="coerce").dropna()
    return int(ids.max()) + 1 if not ids.empty else 1


@st.cache_data(ttl=15)
def load_sheet_df(sheet_name: str) -> pd.DataFrame:
    ws = sh.worksheet(sheet_name)
    rows = ws.get_all_values()
    if len(rows) < 2:
        headers = [h.strip() for h in rows[0]] if rows else []
        return pd.DataFrame(columns=headers)
    headers = [h.strip() for h in rows[0]]
    data = [[cell.strip() for cell in r] for r in rows[1:] if any(c.strip() for c in r)]
    return pd.DataFrame(data, columns=headers)


# ==================== MA'LUMOTLARNI YUKLASH ====================
df_cats = load_sheet_df("Categories")
df_trans = load_sheet_df("Transactions")
df_budgets = load_sheet_df("Budgets")
df_goals = load_sheet_df("Goals")

if not df_cats.empty:
    df_cats = df_cats[df_cats["UserId"].astype(str) == str(USER_ID)]
if not df_trans.empty:
    df_trans = df_trans[df_trans["UserId"].astype(str) == str(USER_ID)].copy()
    df_trans["AmountNum"] = df_trans["Amount"].apply(to_float)
if not df_budgets.empty:
    df_budgets = df_budgets[df_budgets["UserId"].astype(str) == str(USER_ID)].copy()
    df_budgets["LimitNum"] = df_budgets["LimitAmount"].apply(to_float)
if not df_goals.empty:
    df_goals = df_goals[df_goals["UserId"].astype(str) == str(USER_ID)].copy()
    df_goals["TargetNum"] = df_goals["TargetAmount"].apply(to_float)
    df_goals["CurrentNum"] = df_goals["CurrentAmount"].apply(to_float)


# ==================== YUQORI SARLAVHA VA TARTIBLI MENYU ====================
top_row = st.columns([6, 2, 2])
with top_row[0]:
    st.markdown('<div class="app-main-title">💰 Shaxsiy Moliya</div>', unsafe_allow_html=True)
with top_row[1]:
    if st.button("🔄", key="btn_refresh", help="Yangilash"):
        refresh_data()
        st.rerun()
with top_row[2]:
    if st.button("🔒", key="btn_logout", help="Chiqish"):
        st.session_state.authenticated = False
        st.rerun()

total_in = df_trans[df_trans["Type"].str.lower() == "kirim"]["AmountNum"].sum() if not df_trans.empty else 0.0
total_out = df_trans[df_trans["Type"].str.lower() == "chiqim"]["AmountNum"].sum() if not df_trans.empty else 0.0
balance = total_in - total_out

st.markdown(
    f"""
    <div class="kpi-wrapper">
        <div class="kpi-card main-bal">
            <div class="kpi-label">🏦 Joriy Balans</div>
            <div class="kpi-value kpi-bal">{balance:,.0f} UZS</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">📈 Jami Kirim</div>
            <div class="kpi-value kpi-in">+{total_in:,.0f}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">📉 Jami Chiqim</div>
            <div class="kpi-value kpi-out">-{total_out:,.0f}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if "page" not in st.session_state:
    st.session_state.page = "kiritish"

# Tepada 2 ta bir xil va muvozanatli qator: 3 ta + 2 ta tugma
menu_row_1 = st.columns([1, 1, 1])
with menu_row_1[0]:
    label_1 = "● ➕ Kiritish" if st.session_state.page == "kiritish" else "➕ Kiritish"
    if st.button(label_1, key="nav_kiritish"):
        st.session_state.page = "kiritish"
        st.rerun()
with menu_row_1[1]:
    label_2 = "● 📊 Byudjet" if st.session_state.page == "byudjet" else "📊 Byudjet"
    if st.button(label_2, key="nav_byudjet"):
        st.session_state.page = "byudjet"
        st.rerun()
with menu_row_1[2]:
    label_3 = "● 🎯 Maqsadlar" if st.session_state.page == "maqsadlar" else "🎯 Maqsadlar"
    if st.button(label_3, key="nav_maqsadlar"):
        st.session_state.page = "maqsadlar"
        st.rerun()

menu_row_2 = st.columns([1, 1])
with menu_row_2[0]:
    label_4 = "● ⚙️ Kategoriyalar" if st.session_state.page == "kategoriyalar" else "⚙️ Kategoriyalar"
    if st.button(label_4, key="nav_kategoriyalar"):
        st.session_state.page = "kategoriyalar"
        st.rerun()
with menu_row_2[1]:
    label_5 = "● 📋 Tarix va Excel" if st.session_state.page == "tarix" else "📋 Tarix va Excel"
    if st.button(label_5, key="nav_tarix"):
        st.session_state.page = "tarix"
        st.rerun()

st.markdown("---")


# ==================== 1. KIRIM / CHIQIM VA DIAGRAMMA ====================
if st.session_state.page == "kiritish":
    left_col, right_col = st.columns([1, 1.3])

    with left_col:
        st.subheader("➕ Yangi operatsiya")
        t_type = st.radio("Amal turi:", ["Chiqim", "Kirim"])

        cat_options = []
        if not df_cats.empty:
            cat_options = df_cats[df_cats["Type"].str.lower() == t_type.lower()]["CategoryName"].tolist()
        if not cat_options:
            cat_options = ["Boshqa"]

        with st.form("add_trans_form", clear_on_submit=True):
            selected_cat = st.selectbox("Kategoriya:", cat_options)
            title = st.text_input("Nomi (izoh):", placeholder="Masalan: Tushlik, Taksi, Oylik")
            amount = st.number_input("Summa (UZS):", min_value=0.0, step=1000.0, format="%.0f")
            trans_date = st.date_input("Sana:", value=date.today())
            save_btn = st.form_submit_button("✅ Saqlash")

            if save_btn:
                if not title.strip() or amount <= 0:
                    st.warning("Iltimos, nom va 0 dan katta summani kiriting!")
                else:
                    ws_t = sh.worksheet("Transactions")
                    all_t = load_sheet_df("Transactions")
                    new_id = next_id(all_t)
                    ws_t.append_row(
                        [new_id, USER_ID, title.strip(), amount, t_type, selected_cat, str(trans_date)],
                        value_input_option="USER_ENTERED",
                    )
                    refresh_data()
                    st.success(f"{t_type} saqlandi!")
                    st.rerun()

    with right_col:
        st.subheader("📊 Xarajatlar taqsimoti")
        if not df_trans.empty:
            exp_df = df_trans[df_trans["Type"].str.lower() == "chiqim"]
            if not exp_df.empty:
                grouped = exp_df.groupby("Category", as_index=False)["AmountNum"].sum()
                fig = px.pie(
                    grouped,
                    names="Category",
                    values="AmountNum",
                    hole=0.45,
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig.update_traces(textposition="inside", textinfo="percent")
                fig.update_layout(
                    margin=dict(t=10, b=60, l=10, r=10),
                    height=330,
                    legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5),
                    dragmode=False,
                )
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
            else:
                st.info("Diagramma uchun kamida bitta chiqim kiriting.")
        else:
            st.info("Hozircha tranzaksiyalar yo'q.")


# ==================== 2. BYUDJET (LIMITLAR) ====================
elif st.session_state.page == "byudjet":
    b_col1, b_col2 = st.columns([1, 1.4])

    with b_col1:
        st.subheader("📊 Oylik limit belgilash")
        exp_cats = []
        if not df_cats.empty:
            exp_cats = df_cats[df_cats["Type"].str.lower() == "chiqim"]["CategoryName"].tolist()
        if not exp_cats:
            exp_cats = ["Oziq-ovqat", "Transport"]

        with st.form("budget_form", clear_on_submit=True):
            b_cat = st.selectbox("Chiqim kategoriyasi:", exp_cats)
            b_limit = st.number_input("Oylik limit (UZS):", min_value=0.0, step=50000.0, format="%.0f")
            b_submit = st.form_submit_button("💾 Limitni saqlash")

            if b_submit and b_limit > 0:
                ws_b = sh.worksheet("Budgets")
                all_rows = ws_b.get_all_values()
                updated = False
                for idx, row in enumerate(all_rows[1:], start=2):
                    if len(row) >= 3 and row[1].strip() == str(USER_ID) and row[2].strip() == b_cat:
                        ws_b.update_cell(idx, 4, b_limit)
                        updated = True
                        break
                if not updated:
                    all_b = load_sheet_df("Budgets")
                    new_id = next_id(all_b)
                    r_num = len(all_rows) + 1
                    spent_f = f'=SUMIFS(Transactions!$D$2:$D, Transactions!$B$2:$B, B{r_num}, Transactions!$F$2:$F, C{r_num}, Transactions!$E$2:$E, "Chiqim")'
                    rem_f = f"=D{r_num}-E{r_num}"
                    ws_b.append_row([new_id, USER_ID, b_cat, b_limit, spent_f, rem_f], value_input_option="USER_ENTERED")
                refresh_data()
                st.success(f"**{b_cat}** limiti saqlandi!")
                st.rerun()

    with b_col2:
        st.subheader("Limitlar holati")
        if df_budgets.empty:
            st.info("Hali byudjet limitlari belgilanmagan.")
        else:
            exp_df = df_trans[df_trans["Type"].str.lower() == "chiqim"] if not df_trans.empty else pd.DataFrame()
            for _, row in df_budgets.iterrows():
                cat = row["Category"]
                lim = row["LimitNum"]
                spent = exp_df[exp_df["Category"] == cat]["AmountNum"].sum() if not exp_df.empty else 0.0
                rem = lim - spent
                pct = min(spent / lim, 1.0) if lim > 0 else 0.0

                st.markdown(f"**{cat}** — Limit: `{lim:,.0f}` | Sarflandi: `{spent:,.0f}` | Qoldi: `{rem:,.0f} UZS`")
                st.progress(pct)
                if spent > lim:
                    st.error(f"🚨 Limitdan {abs(rem):,.0f} UZS oshib ketdi!")


# ==================== 3. MAQSADLAR VA JAMG'ARMA ====================
elif st.session_state.page == "maqsadlar":
    g_col1, g_col2 = st.columns([1, 1.4])

    with g_col1:
        st.subheader("🎯 Yangi maqsad")
        with st.form("new_goal_form", clear_on_submit=True):
            g_title = st.text_input("Maqsad nomi:", placeholder="Masalan: Yangi telefon")
            g_target = st.number_input("Kerakli summa (UZS):", min_value=0.0, step=100000.0, format="%.0f")
            g_deadline = st.date_input("Muddat:")
            if st.form_submit_button("🎯 Maqsad yaratish"):
                if g_title.strip() and g_target > 0:
                    ws_g = sh.worksheet("Goals")
                    all_rows = ws_g.get_all_values()
                    new_id = next_id(load_sheet_df("Goals"))
                    r = len(all_rows) + 1
                    ws_g.append_row(
                        [
                            new_id,
                            USER_ID,
                            g_title.strip(),
                            g_target,
                            0,
                            f"=IF(D{r}>0, E{r}/D{r}, 0)",
                            str(g_deadline),
                            f'=IF(E{r}>=D{r}, "Bajarildi", "Jarayonda")',
                        ],
                        value_input_option="USER_ENTERED",
                    )
                    refresh_data()
                    st.success("Maqsad yaratildi!")
                    st.rerun()

        if not df_goals.empty:
            st.subheader("💰 Pul jamg'arish")
            with st.form("save_money_form", clear_on_submit=True):
                goal_map = {f"{r['Title']} (Id:{r['Id']})": r["Id"] for _, r in df_goals.iterrows()}
                chosen_label = st.selectbox("Maqsad:", list(goal_map.keys()))
                add_money = st.number_input("Summa (UZS):", min_value=0.0, step=50000.0, format="%.0f")
                if st.form_submit_button("➕ Jamg'armaga qo'shish"):
                    if add_money > 0:
                        gid = str(goal_map[chosen_label])
                        ws_g = sh.worksheet("Goals")
                        all_rows = ws_g.get_all_values()
                        for idx, row in enumerate(all_rows[1:], start=2):
                            if len(row) >= 5 and row[0].strip() == gid:
                                new_curr = to_float(row[4]) + add_money
                                ws_g.update_cell(idx, 5, new_curr)
                                break
                        refresh_data()
                        st.success("Jamg'arma qo'shildi!")
                        st.rerun()

    with g_col2:
        st.subheader("Maqsadlar holati")
        if df_goals.empty:
            st.info("Hozircha maqsadlar yo'q.")
        else:
            for _, g in df_goals.iterrows():
                t_amt = g["TargetNum"]
                c_amt = g["CurrentNum"]
                rem = max(0.0, t_amt - c_amt)
                pct = min(c_amt / t_amt, 1.0) if t_amt > 0 else 0.0
                badge = "✅ Bajarildi" if c_amt >= t_amt else "⏳ Jarayonda"
                st.markdown(
                    f"**🎯 {g['Title']}** ({badge})  \n"
                    f"Jamg'arildi: `{c_amt:,.0f}` / `{t_amt:,.0f} UZS` (**{pct*100:.1f}%**) | Qoldi: `{rem:,.0f} UZS`"
                )
                st.progress(pct)


# ==================== 4. KATEGORIYALAR ====================
elif st.session_state.page == "kategoriyalar":
    c_col1, c_col2 = st.columns([1, 1.4])
    with c_col1:
        st.subheader("⚙️ Yangi kategoriya")
        with st.form("cat_form", clear_on_submit=True):
            c_type = st.selectbox("Turi:", ["Chiqim", "Kirim"])
            c_name = st.text_input("Kategoriya nomi:", placeholder="Masalan: Taksi, Kommunal")
            if st.form_submit_button("➕ Qo'shish"):
                if c_name.strip():
                    ws_c = sh.worksheet("Categories")
                    new_id = next_id(load_sheet_df("Categories"))
                    ws_c.append_row([new_id, USER_ID, c_name.strip(), c_type], value_input_option="USER_ENTERED")
                    refresh_data()
                    st.success(f"'{c_name}' qo'shildi!")
                    st.rerun()

    with c_col2:
        st.subheader("Mavjud kategoriyalar")
        if not df_cats.empty:
            for _, c in df_cats.iterrows():
                icon = "🟢" if c["Type"].lower() == "kirim" else "🔴"
                st.markdown(f"{icon} **{c['CategoryName']}** — _{c['Type']}_")


# ==================== 5. TARIX VA EXCEL ====================
elif st.session_state.page == "tarix":
    st.subheader("📋 Operatsiyalar tarixi")
    if df_trans.empty:
        st.info("Tranzaksiyalar ro'yxati bo'sh.")
    else:
        wb = openpyxl.Workbook()
        h_fill = PatternFill(start_color="39608F", end_color="39608F", fill_type="solid")
        h_font = Font(color="FFFFFF", bold=True)

        ws1 = wb.active
        ws1.title = "Tranzaksiyalar"
        ws1.append(["ID", "Nomi", "Summa", "Turi", "Kategoriya", "Sana"])
        for _, r in df_trans.iterrows():
            ws1.append([r["Id"], r["Title"], r["AmountNum"], r["Type"], r["Category"], r["Date"]])

        ws2 = wb.create_sheet("Byudjetlar")
        ws2.append(["Kategoriya", "Limit Summa"])
        if not df_budgets.empty:
            for _, r in df_budgets.iterrows():
                ws2.append([r["Category"], r["LimitNum"]])

        ws3 = wb.create_sheet("Maqsadlar")
        ws3.append(["Maqsad", "Kerakli Summa", "Jamg'arilgan", "Qolgan", "Muddat"])
        if not df_goals.empty:
            for _, r in df_goals.iterrows():
                ws3.append([r["Title"], r["TargetNum"], r["CurrentNum"], max(0.0, r["TargetNum"] - r["CurrentNum"]), r.get("Deadline", "")])

        for ws in [ws1, ws2, ws3]:
            for cell in ws[1]:
                cell.fill = h_fill
                cell.font = h_font

        excel_buf = io.BytesIO()
        wb.save(excel_buf)
        excel_buf.seek(0)

        st.download_button(
            label="📥 Excel Hisobotni Yuklab Olish (.xlsx)",
            data=excel_buf,
            file_name=f"Moliya_Hisobot_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        filter_type = st.selectbox("Filtr:", ["Barchasi", "Kirim", "Chiqim"])
        view_df = df_trans.copy()
        if filter_type != "Barchasi":
            view_df = view_df[view_df["Type"].str.lower() == filter_type.lower()]

        for _, r in view_df.iloc[::-1].iterrows():
            is_in = r["Type"].lower() == "kirim"
            sign = "+" if is_in else "-"
            cls = "tx-amt-in" if is_in else "tx-amt-out"
            st.markdown(
                f"""
                <div class="tx-item">
                    <div class="tx-left">
                        <span class="tx-title">{r['Title']}</span>
                        <span class="tx-meta">📂 {r['Category']} • 📅 {r['Date']}</span>
                    </div>
                    <div class="{cls}">{sign}{r['AmountNum']:,.0f} UZS</div>
                </div>
                """,
                unsafe_allow_html=True,
            )