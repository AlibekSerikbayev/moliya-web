

# import io
# import os
# from datetime import datetime, date

# import gspread
# from google.oauth2.service_account import Credentials
# import openpyxl
# from openpyxl.styles import Font, PatternFill
# import pandas as pd
# import plotly.express as px
# import streamlit as st
# from dotenv import load_dotenv

# # ==================== SAHIFA SOZLAMALARI ====================
# st.set_page_config(
#     page_title="Moliya Boshqaruvi",
#     page_icon="💰",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# load_dotenv()

# # ==================== MOBIL VA DESKTOP TO'LIQ KENGLIK DIZAYNI (CSS) ====================
# st.markdown(
#     """
#     <style>
#     section[data-testid="stSidebar"], [data-testid="collapsedControl"] {
#         display: none !important;
#     }

#     .block-container {
#         padding-top: 2.3rem !important;
#         padding-bottom: 2.5rem !important;
#         padding-left: 1rem !important;
#         padding-right: 1rem !important;
#         max-width: 1100px;
#     }

#     /* Telefonda inputga bosganda ekran zoom bo'lib ketmasligi uchun 16px */
#     input, textarea, select, div[data-baseweb="select"] > div {
#         font-size: 16px !important;
#         min-height: 44px !important;
#     }

#     /* Barcha tugma konteynerlari va tugmalarni 100% ekran oxirigacha cho'zish */
#     div[data-testid="stButton"],
#     div[data-testid="stDownloadButton"],
#     div[data-testid="stFormSubmitButton"],
#     .stButton,
#     .stDownloadButton {
#         width: 100% !important;
#         display: block !important;
#     }

#     div[data-testid="stButton"] > button,
#     div[data-testid="stDownloadButton"] > button,
#     div[data-testid="stFormSubmitButton"] > button,
#     .stButton > button,
#     .stDownloadButton > button {
#         width: 100% !important;
#         max-width: 100% !important;
#         min-height: 46px !important;
#         height: 46px !important;
#         font-size: 15px !important;
#         font-weight: 600 !important;
#         border-radius: 10px !important;
#         padding: 0 14px !important;
#         display: flex !important;
#         align-items: center !important;
#         justify-content: center !important;
#         white-space: nowrap !important;
#     }

#     /* Kirim/Chiqim radio tugmalari yonma-yon to'liq kenglikda */
#     div[role="radiogroup"] {
#         display: flex !important;
#         flex-direction: row !important;
#         gap: 10px !important;
#         width: 100% !important;
#     }
#     div[role="radiogroup"] > label {
#         flex: 1 !important;
#         background: rgba(255, 255, 255, 0.06) !important;
#         border: 1px solid rgba(255, 255, 255, 0.15) !important;
#         border-radius: 10px !important;
#         padding: 10px 12px !important;
#         justify-content: center !important;
#         cursor: pointer !important;
#     }

#     .app-main-title {
#         font-size: 1.5rem;
#         font-weight: 700;
#         margin: 0 0 8px 0;
#         line-height: 46px;
#     }

#     /* KPI kartalar */
#     .kpi-wrapper {
#         display: grid;
#         grid-template-columns: repeat(3, 1fr);
#         gap: 12px;
#         margin-bottom: 14px;
#         width: 100%;
#     }
#     .kpi-card {
#         background: rgba(255, 255, 255, 0.05);
#         border: 1px solid rgba(255, 255, 255, 0.12);
#         border-radius: 12px;
#         padding: 12px 14px;
#     }
#     .kpi-label {
#         font-size: 12px;
#         opacity: 0.8;
#         margin-bottom: 4px;
#     }
#     .kpi-value {
#         font-size: 20px;
#         font-weight: 700;
#     }
#     .kpi-bal { color: #38BDF8; }
#     .kpi-in  { color: #10B981; }
#     .kpi-out { color: #F87171; }

#     /* Telefonda (<= 640px) KPI va tugmalar */
#     @media (max-width: 640px) {
#         .app-main-title {
#             font-size: 1.3rem;
#             line-height: 1.3;
#         }
#         .kpi-wrapper {
#             grid-template-columns: 1fr 1fr;
#             gap: 8px;
#         }
#         .kpi-card.main-bal {
#             grid-column: span 2;
#             text-align: center;
#         }
#         .kpi-value {
#             font-size: 17px;
#         }
#         h2, h3 {
#             font-size: 1.12rem !important;
#         }
#     }

#     /* Tranzaksiya kartochkalari */
#     .tx-item {
#         display: flex;
#         justify-content: space-between;
#         align-items: center;
#         gap: 10px;
#         background: rgba(255, 255, 255, 0.04);
#         border: 1px solid rgba(255, 255, 255, 0.08);
#         border-radius: 10px;
#         padding: 10px 12px;
#         margin-bottom: 8px;
#         width: 100%;
#     }
#     .tx-left {
#         display: flex;
#         flex-direction: column;
#         min-width: 0;
#     }
#     .tx-title {
#         font-weight: 600;
#         font-size: 15px;
#         overflow: hidden;
#         text-overflow: ellipsis;
#     }
#     .tx-meta {
#         font-size: 12px;
#         opacity: 0.75;
#     }
#     .tx-amt-in {
#         color: #10B981;
#         font-weight: 700;
#         font-size: 14px;
#         white-space: nowrap;
#     }
#     .tx-amt-out {
#         color: #F87171;
#         font-weight: 700;
#         font-size: 14px;
#         white-space: nowrap;
#     }
#     </style>
#     """,
#     unsafe_allow_html=True,
# )


# def refresh_data():
#     st.cache_data.clear()


# def has_secrets_file() -> bool:
#     local_secret = os.path.join(os.getcwd(), ".streamlit", "secrets.toml")
#     home_secret = os.path.join(os.path.expanduser("~"), ".streamlit", "secrets.toml")
#     return os.path.exists(local_secret) or os.path.exists(home_secret)


# def get_secret(key: str, default: str = "") -> str:
#     try:
#         if key in st.secrets:
#             return str(st.secrets[key]).strip()
#     except Exception:
#         pass
#     if has_secrets_file():
#         try:
#             if key in st.secrets:
#                 return str(st.secrets[key]).strip()
#         except Exception:
#             pass
#     val = os.getenv(key)
#     return val.strip() if val and val.strip() else default


# SPREADSHEET_ID = get_secret("SPREADSHEET_ID")
# APP_PASSWORD = get_secret("APP_PASSWORD", "12345")
# USER_ID = int(get_secret("DEFAULT_USER_ID", "1"))

# if not SPREADSHEET_ID:
#     st.error("❌ SPREADSHEET_ID topilmadi! Streamlit Secrets yoki .env faylini tekshiring.")
#     st.stop()

# # ==================== PAROL BILAN HIMOYA (LOGIN) ====================
# if "authenticated" not in st.session_state:
#     st.session_state.authenticated = False

# if not st.session_state.authenticated:
#     st.markdown("## 🔐 Shaxsiy Moliya Paneli")
#     with st.form("login_form"):
#         pwd = st.text_input("Parolni kiriting:", type="password")
#         submitted = st.form_submit_button("Kirish", use_container_width=True)
#         if submitted:
#             if pwd == APP_PASSWORD:
#                 st.session_state.authenticated = True
#                 st.rerun()
#             else:
#                 st.error("❌ Parol noto'g'ri!")
#     st.stop()


# # ==================== GOOGLE SHEETS ULANISH ====================
# SCOPES = [
#     "https://www.googleapis.com/auth/spreadsheets",
#     "https://www.googleapis.com/auth/drive",
# ]


# @st.cache_resource
# def get_spreadsheet():
#     try:
#         if "gcp_service_account" in st.secrets:
#             creds_dict = dict(st.secrets["gcp_service_account"])
#             creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
#             gc = gspread.authorize(creds)
#             return gc.open_by_key(SPREADSHEET_ID)
#     except Exception:
#         pass

#     creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
#     gc = gspread.authorize(creds)
#     return gc.open_by_key(SPREADSHEET_ID)


# sh = get_spreadsheet()


# def to_float(val) -> float:
#     if val is None or val == "":
#         return 0.0
#     cleaned = str(val).replace(" ", "").replace("\xa0", "").replace(",", ".")
#     try:
#         return float(cleaned)
#     except ValueError:
#         return 0.0


# def next_id(df: pd.DataFrame) -> int:
#     if df.empty or "Id" not in df.columns:
#         return 1
#     ids = pd.to_numeric(df["Id"], errors="coerce").dropna()
#     return int(ids.max()) + 1 if not ids.empty else 1


# @st.cache_data(ttl=15)
# def load_sheet_df(sheet_name: str) -> pd.DataFrame:
#     ws = sh.worksheet(sheet_name)
#     rows = ws.get_all_values()
#     if len(rows) < 2:
#         headers = [h.strip() for h in rows[0]] if rows else []
#         return pd.DataFrame(columns=headers)
#     headers = [h.strip() for h in rows[0]]
#     data = [[cell.strip() for cell in r] for r in rows[1:] if any(c.strip() for c in r)]
#     return pd.DataFrame(data, columns=headers)


# # ==================== MA'LUMOTLARNI YUKLASH ====================
# df_cats = load_sheet_df("Categories")
# df_trans = load_sheet_df("Transactions")
# df_budgets = load_sheet_df("Budgets")
# df_goals = load_sheet_df("Goals")

# if not df_cats.empty:
#     df_cats = df_cats[df_cats["UserId"].astype(str) == str(USER_ID)]
# if not df_trans.empty:
#     df_trans = df_trans[df_trans["UserId"].astype(str) == str(USER_ID)].copy()
#     df_trans["AmountNum"] = df_trans["Amount"].apply(to_float)
# if not df_budgets.empty:
#     df_budgets = df_budgets[df_budgets["UserId"].astype(str) == str(USER_ID)].copy()
#     df_budgets["LimitNum"] = df_budgets["LimitAmount"].apply(to_float)
# if not df_goals.empty:
#     df_goals = df_goals[df_goals["UserId"].astype(str) == str(USER_ID)].copy()
#     df_goals["TargetNum"] = df_goals["TargetAmount"].apply(to_float)
#     df_goals["CurrentNum"] = df_goals["CurrentAmount"].apply(to_float)


# # ==================== YUQORI SARLAVHA VA TUGMALAR ====================
# top_row = st.columns([6, 2, 2])
# with top_row[0]:
#     st.markdown('<div class="app-main-title">💰 Shaxsiy Moliya</div>', unsafe_allow_html=True)
# with top_row[1]:
#     if st.button("🔄 Yangilash", key="btn_refresh", use_container_width=True):
#         refresh_data()
#         st.rerun()
# with top_row[2]:
#     if st.button("🔒 Chiqish", key="btn_logout", use_container_width=True):
#         st.session_state.authenticated = False
#         st.rerun()

# total_in = df_trans[df_trans["Type"].str.lower() == "kirim"]["AmountNum"].sum() if not df_trans.empty else 0.0
# total_out = df_trans[df_trans["Type"].str.lower() == "chiqim"]["AmountNum"].sum() if not df_trans.empty else 0.0
# balance = total_in - total_out

# st.markdown(
#     f"""
#     <div class="kpi-wrapper">
#         <div class="kpi-card main-bal">
#             <div class="kpi-label">🏦 Joriy Balans</div>
#             <div class="kpi-value kpi-bal">{balance:,.0f} UZS</div>
#         </div>
#         <div class="kpi-card">
#             <div class="kpi-label">📈 Jami Kirim</div>
#             <div class="kpi-value kpi-in">+{total_in:,.0f}</div>
#         </div>
#         <div class="kpi-card">
#             <div class="kpi-label">📉 Jami Chiqim</div>
#             <div class="kpi-value kpi-out">-{total_out:,.0f}</div>
#         </div>
#     </div>
#     """,
#     unsafe_allow_html=True,
# )

# if "page" not in st.session_state:
#     st.session_state.page = "kiritish"

# # Ekran oxirigacha to'liq cho'zilgan menyu tugmalari (3 ta + 2 ta)
# menu_row_1 = st.columns([1, 1, 1])
# with menu_row_1[0]:
#     label_1 = "● ➕ Kiritish" if st.session_state.page == "kiritish" else "➕ Kiritish"
#     if st.button(label_1, key="nav_kiritish", use_container_width=True):
#         st.session_state.page = "kiritish"
#         st.rerun()
# with menu_row_1[1]:
#     label_2 = "● 📊 Byudjet" if st.session_state.page == "byudjet" else "📊 Byudjet"
#     if st.button(label_2, key="nav_byudjet", use_container_width=True):
#         st.session_state.page = "byudjet"
#         st.rerun()
# with menu_row_1[2]:
#     label_3 = "● 🎯 Maqsadlar" if st.session_state.page == "maqsadlar" else "🎯 Maqsadlar"
#     if st.button(label_3, key="nav_maqsadlar", use_container_width=True):
#         st.session_state.page = "maqsadlar"
#         st.rerun()

# menu_row_2 = st.columns([1, 1])
# with menu_row_2[0]:
#     label_4 = "● ⚙️ Kategoriyalar" if st.session_state.page == "kategoriyalar" else "⚙️ Kategoriyalar"
#     if st.button(label_4, key="nav_kategoriyalar", use_container_width=True):
#         st.session_state.page = "kategoriyalar"
#         st.rerun()
# with menu_row_2[1]:
#     label_5 = "● 📋 Tarix va Excel" if st.session_state.page == "tarix" else "📋 Tarix va Excel"
#     if st.button(label_5, key="nav_tarix", use_container_width=True):
#         st.session_state.page = "tarix"
#         st.rerun()

# st.markdown("---")


# # ==================== 1. KIRIM / CHIQIM VA DIAGRAMMA ====================
# if st.session_state.page == "kiritish":
#     left_col, right_col = st.columns([1, 1.3])

#     with left_col:
#         st.subheader("➕ Yangi operatsiya")
#         t_type = st.radio("Amal turi:", ["Chiqim", "Kirim"])

#         cat_options = []
#         if not df_cats.empty:
#             cat_options = df_cats[df_cats["Type"].str.lower() == t_type.lower()]["CategoryName"].tolist()
#         if not cat_options:
#             cat_options = ["Boshqa"]

#         with st.form("add_trans_form", clear_on_submit=True):
#             selected_cat = st.selectbox("Kategoriya:", cat_options)
#             title = st.text_input("Nomi (izoh):", placeholder="Masalan: Tushlik, Taksi, Oylik")
#             amount = st.number_input("Summa (UZS):", min_value=0.0, step=1000.0, format="%.0f")
#             trans_date = st.date_input("Sana:", value=date.today())
#             save_btn = st.form_submit_button("✅ Saqlash", use_container_width=True)

#             if save_btn:
#                 if not title.strip() or amount <= 0:
#                     st.warning("Iltimos, nom va 0 dan katta summani kiriting!")
#                 else:
#                     ws_t = sh.worksheet("Transactions")
#                     all_t = load_sheet_df("Transactions")
#                     new_id = next_id(all_t)
#                     ws_t.append_row(
#                         [new_id, USER_ID, title.strip(), amount, t_type, selected_cat, str(trans_date)],
#                         value_input_option="USER_ENTERED",
#                     )
#                     refresh_data()
#                     st.success(f"{t_type} saqlandi!")
#                     st.rerun()

#     with right_col:
#         st.subheader("📊 Xarajatlar taqsimoti")
#         if not df_trans.empty:
#             exp_df = df_trans[df_trans["Type"].str.lower() == "chiqim"]
#             if not exp_df.empty:
#                 grouped = exp_df.groupby("Category", as_index=False)["AmountNum"].sum()
#                 fig = px.pie(
#                     grouped,
#                     names="Category",
#                     values="AmountNum",
#                     hole=0.45,
#                     color_discrete_sequence=px.colors.qualitative.Set2,
#                 )
#                 fig.update_traces(textposition="inside", textinfo="percent")
#                 fig.update_layout(
#                     margin=dict(t=10, b=60, l=10, r=10),
#                     height=330,
#                     legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5),
#                     dragmode=False,
#                 )
#                 st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
#             else:
#                 st.info("Diagramma uchun kamida bitta chiqim kiriting.")
#         else:
#             st.info("Hozircha tranzaksiyalar yo'q.")


# # ==================== 2. BYUDJET (LIMITLAR) ====================
# elif st.session_state.page == "byudjet":
#     b_col1, b_col2 = st.columns([1, 1.4])

#     with b_col1:
#         st.subheader("📊 Oylik limit belgilash")
#         exp_cats = []
#         if not df_cats.empty:
#             exp_cats = df_cats[df_cats["Type"].str.lower() == "chiqim"]["CategoryName"].tolist()
#         if not exp_cats:
#             exp_cats = ["Oziq-ovqat", "Transport"]

#         with st.form("budget_form", clear_on_submit=True):
#             b_cat = st.selectbox("Chiqim kategoriyasi:", exp_cats)
#             b_limit = st.number_input("Oylik limit (UZS):", min_value=0.0, step=50000.0, format="%.0f")
#             b_submit = st.form_submit_button("💾 Limitni saqlash", use_container_width=True)

#             if b_submit and b_limit > 0:
#                 ws_b = sh.worksheet("Budgets")
#                 all_rows = ws_b.get_all_values()
#                 updated = False
#                 for idx, row in enumerate(all_rows[1:], start=2):
#                     if len(row) >= 3 and row[1].strip() == str(USER_ID) and row[2].strip() == b_cat:
#                         ws_b.update_cell(idx, 4, b_limit)
#                         updated = True
#                         break
#                 if not updated:
#                     all_b = load_sheet_df("Budgets")
#                     new_id = next_id(all_b)
#                     r_num = len(all_rows) + 1
#                     spent_f = f'=SUMIFS(Transactions!$D$2:$D, Transactions!$B$2:$B, B{r_num}, Transactions!$F$2:$F, C{r_num}, Transactions!$E$2:$E, "Chiqim")'
#                     rem_f = f"=D{r_num}-E{r_num}"
#                     ws_b.append_row([new_id, USER_ID, b_cat, b_limit, spent_f, rem_f], value_input_option="USER_ENTERED")
#                 refresh_data()
#                 st.success(f"**{b_cat}** limiti saqlandi!")
#                 st.rerun()

#     with b_col2:
#         st.subheader("Limitlar holati")
#         if df_budgets.empty:
#             st.info("Hali byudjet limitlari belgilanmagan.")
#         else:
#             exp_df = df_trans[df_trans["Type"].str.lower() == "chiqim"] if not df_trans.empty else pd.DataFrame()
#             for _, row in df_budgets.iterrows():
#                 cat = row["Category"]
#                 lim = row["LimitNum"]
#                 spent = exp_df[exp_df["Category"] == cat]["AmountNum"].sum() if not exp_df.empty else 0.0
#                 rem = lim - spent
#                 pct = min(spent / lim, 1.0) if lim > 0 else 0.0

#                 st.markdown(f"**{cat}** — Limit: `{lim:,.0f}` | Sarflandi: `{spent:,.0f}` | Qoldi: `{rem:,.0f} UZS`")
#                 st.progress(pct)
#                 if spent > lim:
#                     st.error(f"🚨 Limitdan {abs(rem):,.0f} UZS oshib ketdi!")


# # ==================== 3. MAQSADLAR VA JAMG'ARMA ====================
# elif st.session_state.page == "maqsadlar":
#     g_col1, g_col2 = st.columns([1, 1.4])

#     with g_col1:
#         st.subheader("🎯 Yangi maqsad")
#         with st.form("new_goal_form", clear_on_submit=True):
#             g_title = st.text_input("Maqsad nomi:", placeholder="Masalan: Yangi telefon")
#             g_target = st.number_input("Kerakli summa (UZS):", min_value=0.0, step=100000.0, format="%.0f")
#             g_deadline = st.date_input("Muddat:")
#             if st.form_submit_button("🎯 Maqsad yaratish", use_container_width=True):
#                 if g_title.strip() and g_target > 0:
#                     ws_g = sh.worksheet("Goals")
#                     all_rows = ws_g.get_all_values()
#                     new_id = next_id(load_sheet_df("Goals"))
#                     r = len(all_rows) + 1
#                     ws_g.append_row(
#                         [
#                             new_id,
#                             USER_ID,
#                             g_title.strip(),
#                             g_target,
#                             0,
#                             f"=IF(D{r}>0, E{r}/D{r}, 0)",
#                             str(g_deadline),
#                             f'=IF(E{r}>=D{r}, "Bajarildi", "Jarayonda")',
#                         ],
#                         value_input_option="USER_ENTERED",
#                     )
#                     refresh_data()
#                     st.success("Maqsad yaratildi!")
#                     st.rerun()

#         if not df_goals.empty:
#             st.subheader("💰 Pul jamg'arish")
#             with st.form("save_money_form", clear_on_submit=True):
#                 goal_map = {f"{r['Title']} (Id:{r['Id']})": r["Id"] for _, r in df_goals.iterrows()}
#                 chosen_label = st.selectbox("Maqsad:", list(goal_map.keys()))
#                 add_money = st.number_input("Summa (UZS):", min_value=0.0, step=50000.0, format="%.0f")
#                 if st.form_submit_button("➕ Jamg'armaga qo'shish", use_container_width=True):
#                     if add_money > 0:
#                         gid = str(goal_map[chosen_label])
#                         ws_g = sh.worksheet("Goals")
#                         all_rows = ws_g.get_all_values()
#                         for idx, row in enumerate(all_rows[1:], start=2):
#                             if len(row) >= 5 and row[0].strip() == gid:
#                                 new_curr = to_float(row[4]) + add_money
#                                 ws_g.update_cell(idx, 5, new_curr)
#                                 break
#                         refresh_data()
#                         st.success("Jamg'arma qo'shildi!")
#                         st.rerun()

#     with g_col2:
#         st.subheader("Maqsadlar holati")
#         if df_goals.empty:
#             st.info("Hozircha maqsadlar yo'q.")
#         else:
#             for _, g in df_goals.iterrows():
#                 t_amt = g["TargetNum"]
#                 c_amt = g["CurrentNum"]
#                 rem = max(0.0, t_amt - c_amt)
#                 pct = min(c_amt / t_amt, 1.0) if t_amt > 0 else 0.0
#                 badge = "✅ Bajarildi" if c_amt >= t_amt else "⏳ Jarayonda"
#                 st.markdown(
#                     f"**🎯 {g['Title']}** ({badge})  \n"
#                     f"Jamg'arildi: `{c_amt:,.0f}` / `{t_amt:,.0f} UZS` (**{pct*100:.1f}%**) | Qoldi: `{rem:,.0f} UZS`"
#                 )
#                 st.progress(pct)


# # ==================== 4. KATEGORIYALAR ====================
# elif st.session_state.page == "kategoriyalar":
#     c_col1, c_col2 = st.columns([1, 1.4])
#     with c_col1:
#         st.subheader("⚙️ Yangi kategoriya")
#         with st.form("cat_form", clear_on_submit=True):
#             c_type = st.selectbox("Turi:", ["Chiqim", "Kirim"])
#             c_name = st.text_input("Kategoriya nomi:", placeholder="Masalan: Taksi, Kommunal")
#             if st.form_submit_button("➕ Qo'shish", use_container_width=True):
#                 if c_name.strip():
#                     ws_c = sh.worksheet("Categories")
#                     new_id = next_id(load_sheet_df("Categories"))
#                     ws_c.append_row([new_id, USER_ID, c_name.strip(), c_type], value_input_option="USER_ENTERED")
#                     refresh_data()
#                     st.success(f"'{c_name}' qo'shildi!")
#                     st.rerun()

#     with c_col2:
#         st.subheader("Mavjud kategoriyalar")
#         if not df_cats.empty:
#             for _, c in df_cats.iterrows():
#                 icon = "🟢" if c["Type"].lower() == "kirim" else "🔴"
#                 st.markdown(f"{icon} **{c['CategoryName']}** — _{c['Type']}_")


# # ==================== 5. TARIX VA EXCEL ====================
# elif st.session_state.page == "tarix":
#     st.subheader("📋 Operatsiyalar tarixi")
#     if df_trans.empty:
#         st.info("Tranzaksiyalar ro'yxati bo'sh.")
#     else:
#         wb = openpyxl.Workbook()
#         h_fill = PatternFill(start_color="39608F", end_color="39608F", fill_type="solid")
#         h_font = Font(color="FFFFFF", bold=True)

#         ws1 = wb.active
#         ws1.title = "Tranzaksiyalar"
#         ws1.append(["ID", "Nomi", "Summa", "Turi", "Kategoriya", "Sana"])
#         for _, r in df_trans.iterrows():
#             ws1.append([r["Id"], r["Title"], r["AmountNum"], r["Type"], r["Category"], r["Date"]])

#         ws2 = wb.create_sheet("Byudjetlar")
#         ws2.append(["Kategoriya", "Limit Summa"])
#         if not df_budgets.empty:
#             for _, r in df_budgets.iterrows():
#                 ws2.append([r["Category"], r["LimitNum"]])

#         ws3 = wb.create_sheet("Maqsadlar")
#         ws3.append(["Maqsad", "Kerakli Summa", "Jamg'arilgan", "Qolgan", "Muddat"])
#         if not df_goals.empty:
#             for _, r in df_goals.iterrows():
#                 ws3.append([r["Title"], r["TargetNum"], r["CurrentNum"], max(0.0, r["TargetNum"] - r["CurrentNum"]), r.get("Deadline", "")])

#         for ws in [ws1, ws2, ws3]:
#             for cell in ws[1]:
#                 cell.fill = h_fill
#                 cell.font = h_font

#         excel_buf = io.BytesIO()
#         wb.save(excel_buf)
#         excel_buf.seek(0)

#         st.download_button(
#             label="📥 Excel Hisobotni Yuklab Olish (.xlsx)",
#             data=excel_buf,
#             file_name=f"Moliya_Hisobot_{datetime.now().strftime('%Y%m%d')}.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             use_container_width=True,
#         )

#         filter_type = st.selectbox("Filtr:", ["Barchasi", "Kirim", "Chiqim"])
#         view_df = df_trans.copy()
#         if filter_type != "Barchasi":
#             view_df = view_df[view_df["Type"].str.lower() == filter_type.lower()]

#         for _, r in view_df.iloc[::-1].iterrows():
#             is_in = r["Type"].lower() == "kirim"
#             sign = "+" if is_in else "-"
#             cls = "tx-amt-in" if is_in else "tx-amt-out"
#             st.markdown(
#                 f"""
#                 <div class="tx-item">
#                     <div class="tx-left">
#                         <span class="tx-title">{r['Title']}</span>
#                         <span class="tx-meta">📂 {r['Category']} • 📅 {r['Date']}</span>
#                     </div>
#                     <div class="{cls}">{sign}{r['AmountNum']:,.0f} UZS</div>
#                 </div>
#                 """,
#                 unsafe_allow_html=True,
#             )


import io
import os
from datetime import datetime, date, timedelta

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

# ==================== MOBIL VA DESKTOP TO'LIQ KENGLIK DIZAYNI (CSS) ====================
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none !important;
    }

    .block-container {
        padding-top: 2.3rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 1150px;
    }

    input, textarea, select, div[data-baseweb="select"] > div {
        font-size: 16px !important;
        min-height: 44px !important;
    }

    div[data-testid="stButton"],
    div[data-testid="stDownloadButton"],
    div[data-testid="stFormSubmitButton"],
    .stButton,
    .stDownloadButton {
        width: 100% !important;
        display: block !important;
    }

    div[data-testid="stButton"] > button,
    div[data-testid="stDownloadButton"] > button,
    div[data-testid="stFormSubmitButton"] > button,
    .stButton > button,
    .stDownloadButton > button {
        width: 100% !important;
        max-width: 100% !important;
        min-height: 44px !important;
        height: 44px !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
        padding: 0 10px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        white-space: nowrap !important;
    }

    div[role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        gap: 10px !important;
        width: 100% !important;
    }
    div[role="radiogroup"] > label {
        flex: 1 !important;
        background: rgba(255, 255, 255, 0.06) !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
        border-radius: 10px !important;
        padding: 10px 12px !important;
        justify-content: center !important;
        cursor: pointer !important;
    }

    .app-main-title {
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0 0 8px 0;
        line-height: 44px;
    }

    .kpi-wrapper {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 12px;
        margin-bottom: 14px;
        width: 100%;
    }
    .kpi-card {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        padding: 12px 14px;
    }
    .kpi-label {
        font-size: 12px;
        opacity: 0.8;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 20px;
        font-weight: 700;
    }
    .kpi-bal { color: #38BDF8; }
    .kpi-in  { color: #10B981; }
    .kpi-out { color: #F87171; }

    .analytics-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 10px;
        margin-bottom: 14px;
        width: 100%;
    }
    .stat-box {
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 11px 12px;
    }
    .stat-title {
        font-size: 12px;
        opacity: 0.78;
        margin-bottom: 4px;
    }
    .stat-num {
        font-size: 16px;
        font-weight: 700;
    }
    .stat-sub {
        font-size: 11px;
        opacity: 0.7;
        margin-top: 2px;
    }

    .asset-card {
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 10px;
    }
    .pos-green { color: #10B981; font-weight: 700; }
    .pos-red   { color: #F87171; font-weight: 700; }
    .pos-gold  { color: #FBBF24; font-weight: 700; }

    @media (max-width: 640px) {
        .app-main-title {
            font-size: 1.3rem;
            line-height: 1.3;
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
        .analytics-grid {
            grid-template-columns: 1fr;
            gap: 8px;
        }
        h2, h3 {
            font-size: 1.12rem !important;
        }
    }

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
        width: 100%;
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
        submitted = st.form_submit_button("Kirish", use_container_width=True)
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

EXTRA_SCHEMAS = {
    "Credits": [
        "Id", "UserId", "Title", "Principal", "AnnualRate",
        "Months", "MonthlyPayment", "PaidAmount", "RemainingPrincipal", "StartDate",
    ],
    "IslamicDeposits": [
        "Id", "UserId", "Title", "ContractType", "Principal",
        "MonthlyContribution", "ExpectedAnnualRate", "ActualProfit", "TermMonths", "StartDate",
    ],
    "Stocks": [
        "Id", "UserId", "Ticker", "CompanyName", "Shares",
        "BuyPrice", "CurrentPrice", "Dividend",
    ],
    "Bonds": [
        "Id", "UserId", "Title", "Quantity", "FaceValue",
        "BuyPrice", "CouponRate", "CouponReceived", "MaturityDate",
    ],
}


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


def ensure_extra_sheets():
    try:
        existing = {ws.title: ws for ws in sh.worksheets()}
        for name, headers in EXTRA_SCHEMAS.items():
            if name not in existing:
                ws = sh.add_worksheet(title=name, rows=500, cols=len(headers))
                ws.append_row(headers, value_input_option="USER_ENTERED")
    except Exception:
        pass


ensure_extra_sheets()


def to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    cleaned = str(val).replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_date_safe(val):
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip()[:10]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def calc_annuity(principal: float, annual_rate: float, months: int) -> float:
    if principal <= 0 or months <= 0:
        return 0.0
    if annual_rate <= 0:
        return principal / months
    r = (annual_rate / 100.0) / 12.0
    factor = (1.0 + r) ** months
    return principal * (r * factor) / (factor - 1.0)


def calc_islamic_projected_profit(principal: float, monthly_add: float, annual_rate: float, months: int) -> float:
    """Boshlang'ich mablag' va oyma-oy qo'shimcha badallar asosida taxminiy kutilma foydani hisoblaydi."""
    if months <= 0 or annual_rate <= 0:
        return 0.0
    m_rate = (annual_rate / 100.0) / 12.0
    # Boshlang'ich mablag' daromadi
    base_profit = principal * m_rate * months
    # Har oy oxirida/boshida qo'shiladigan oylik badalning o'rtacha muddatdagi daromadi
    contrib_profit = monthly_add * m_rate * (months * (months + 1) / 2.0)
    return base_profit + contrib_profit


def filter_by_period(df: pd.DataFrame, period: str) -> pd.DataFrame:
    if df.empty or period == "Barchasi":
        return df
    today = date.today()
    if period == "1 kunlik":
        start_d = today
    elif period == "1 haftalik":
        start_d = today - timedelta(days=6)
    elif period == "1 oylik":
        start_d = today - timedelta(days=29)
    elif period == "1 yillik":
        start_d = today - timedelta(days=364)
    else:
        return df

    parsed = df["Date"].apply(parse_date_safe)
    mask = parsed.apply(lambda d: d is not None and start_d <= d <= today)
    return df[mask].copy()


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
        headers = [h.strip() for h in rows[0]] if rows else EXTRA_SCHEMAS.get(sheet_name, [])
        return pd.DataFrame(columns=headers)
    headers = [h.strip() for h in rows[0]]
    data = [[cell.strip() for cell in r] for r in rows[1:] if any(c.strip() for c in r)]
    return pd.DataFrame(data, columns=headers)


# ==================== MA'LUMOTLARNI YUKLASH ====================
df_cats = load_sheet_df("Categories")
df_trans = load_sheet_df("Transactions")
df_budgets = load_sheet_df("Budgets")
df_goals = load_sheet_df("Goals")
df_credits = load_sheet_df("Credits")
df_deposits = load_sheet_df("IslamicDeposits")
df_stocks = load_sheet_df("Stocks")
df_bonds = load_sheet_df("Bonds")

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
if not df_credits.empty:
    df_credits = df_credits[df_credits["UserId"].astype(str) == str(USER_ID)].copy()
if not df_deposits.empty:
    df_deposits = df_deposits[df_deposits["UserId"].astype(str) == str(USER_ID)].copy()
if not df_stocks.empty:
    df_stocks = df_stocks[df_stocks["UserId"].astype(str) == str(USER_ID)].copy()
if not df_bonds.empty:
    df_bonds = df_bonds[df_bonds["UserId"].astype(str) == str(USER_ID)].copy()


# ==================== YUQORI SARLAVHA VA TUGMALAR ====================
top_row = st.columns([6, 2, 2])
with top_row[0]:
    st.markdown('<div class="app-main-title">💰 Shaxsiy Moliya</div>', unsafe_allow_html=True)
with top_row[1]:
    if st.button("🔄 Yangilash", key="btn_refresh", use_container_width=True):
        refresh_data()
        st.rerun()
with top_row[2]:
    if st.button("🔒 Chiqish", key="btn_logout", use_container_width=True):
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

NAV_BUTTONS = [
    ("kiritish", "➕ Kiritish"),
    ("byudjet", "📊 Byudjet"),
    ("maqsadlar", "🎯 Maqsadlar"),
    ("kredit", "💳 Kredit"),
    ("islomiy", "🕌 Islomiy Depozit"),
    ("aksiya", "📈 Aksiya"),
    ("obligatsiya", "📜 Obligatsiya"),
    ("kategoriyalar", "⚙️ Kategoriyalar"),
    ("tarix", "📋 Hisobot va Tahlil"),
]

for r_idx in range(0, len(NAV_BUTTONS), 3):
    row_cols = st.columns(3)
    for c_idx, (p_key, p_label) in enumerate(NAV_BUTTONS[r_idx : r_idx + 3]):
        with row_cols[c_idx]:
            btn_lbl = f"● {p_label}" if st.session_state.page == p_key else p_label
            if st.button(btn_lbl, key=f"nav_{p_key}", use_container_width=True):
                st.session_state.page = p_key
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
            save_btn = st.form_submit_button("✅ Saqlash", use_container_width=True)

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
            b_submit = st.form_submit_button("💾 Limitni saqlash", use_container_width=True)

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
            if st.form_submit_button("🎯 Maqsad yaratish", use_container_width=True):
                if g_title.strip() and g_target > 0:
                    ws_g = sh.worksheet("Goals")
                    all_rows = ws_g.get_all_values()
                    new_id = next_id(load_sheet_df("Goals"))
                    r = len(all_rows) + 1
                    ws_g.append_row(
                        [
                            new_id, USER_ID, g_title.strip(), g_target, 0,
                            f"=IF(D{r}>0, E{r}/D{r}, 0)", str(g_deadline),
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
                if st.form_submit_button("➕ Jamg'armaga qo'shish", use_container_width=True):
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


# ==================== 4. KREDIT (HISOBLASH VA TO'LASH) ====================
elif st.session_state.page == "kredit":
    cr_col1, cr_col2 = st.columns([1, 1.4])

    with cr_col1:
        st.subheader("💳 Kredit kalkulyatori va qo'shish")
        with st.form("add_credit_form", clear_on_submit=True):
            cr_title = st.text_input("Kredit nomi:", placeholder="Masalan: Avtokredit, Ipoteka")
            cr_principal = st.number_input("Kredit summasi (UZS):", min_value=0.0, step=1000000.0, format="%.0f")
            cr_rate = st.number_input("Yillik foiz stavkasi (%):", min_value=0.0, max_value=100.0, step=0.5, format="%.1f")
            cr_months = st.number_input("Muddati (oy):", min_value=1, max_value=360, value=12, step=1)
            cr_start = st.date_input("Olingan sana:", value=date.today())

            preview_pmt = calc_annuity(cr_principal, cr_rate, int(cr_months))
            st.caption(f"Taxminiy oylik to'lov: **{preview_pmt:,.0f} UZS**")

            if st.form_submit_button("💾 Kreditni ro'yxatga olish", use_container_width=True):
                if cr_title.strip() and cr_principal > 0:
                    monthly_pmt = calc_annuity(cr_principal, cr_rate, int(cr_months))
                    ws_cr = sh.worksheet("Credits")
                    new_id = next_id(load_sheet_df("Credits"))
                    ws_cr.append_row(
                        [
                            new_id, USER_ID, cr_title.strip(), cr_principal, cr_rate,
                            int(cr_months), round(monthly_pmt, 2), 0.0, cr_principal, str(cr_start),
                        ],
                        value_input_option="USER_ENTERED",
                    )
                    refresh_data()
                    st.success("Kredit qo'shildi!")
                    st.rerun()
                else:
                    st.warning("Kredit nomi va summasini kiriting!")

        if not df_credits.empty:
            st.subheader("💵 Kredit to'lovini kiritish")
            with st.form("pay_credit_form", clear_on_submit=True):
                cr_map = {
                    f"{r['Title']} (Qoldiq: {to_float(r['RemainingPrincipal']):,.0f})": r["Id"]
                    for _, r in df_credits.iterrows()
                }
                chosen_cr = st.selectbox("Kreditni tanlang:", list(cr_map.keys()))
                pay_amt = st.number_input("To'lov summasi (UZS):", min_value=0.0, step=100000.0, format="%.0f")
                record_tx = st.checkbox("Chiqim sifatida Tranzaksiyalarga ham yozish", value=True)

                if st.form_submit_button("✅ To'lovni amalga oshirish", use_container_width=True):
                    if pay_amt > 0:
                        cid = str(cr_map[chosen_cr])
                        ws_cr = sh.worksheet("Credits")
                        all_rows = ws_cr.get_all_values()
                        cr_name = ""
                        for idx, row in enumerate(all_rows[1:], start=2):
                            if len(row) >= 9 and row[0].strip() == cid:
                                cr_name = row[2].strip()
                                new_paid = to_float(row[7]) + pay_amt
                                new_rem = max(0.0, to_float(row[8]) - pay_amt)
                                ws_cr.update_cell(idx, 8, new_paid)
                                ws_cr.update_cell(idx, 9, new_rem)
                                break

                        if record_tx:
                            ws_t = sh.worksheet("Transactions")
                            tid = next_id(load_sheet_df("Transactions"))
                            ws_t.append_row(
                                [tid, USER_ID, f"Kredit to'lovi: {cr_name}", pay_amt, "Chiqim", "Kredit", str(date.today())],
                                value_input_option="USER_ENTERED",
                            )

                        refresh_data()
                        st.success("Kredit to'lovi saqlandi!")
                        st.rerun()

    with cr_col2:
        st.subheader("Kreditlar hisob-kitobi")
        if df_credits.empty:
            st.info("Hozircha kreditlar kiritilmagan.")
        else:
            tot_rem = sum(to_float(r["RemainingPrincipal"]) for _, r in df_credits.iterrows())
            tot_paid = sum(to_float(r["PaidAmount"]) for _, r in df_credits.iterrows())
            st.markdown(f"**Jami qoldiq qarz:** `{tot_rem:,.0f} UZS` | **Jami to'langan:** `{tot_paid:,.0f} UZS`")

            for _, r in df_credits.iterrows():
                princ = to_float(r["Principal"])
                paid = to_float(r["PaidAmount"])
                rem = to_float(r["RemainingPrincipal"])
                m_pmt = to_float(r["MonthlyPayment"])
                rate = to_float(r["AnnualRate"])
                months = int(to_float(r["Months"]))
                total_plan = m_pmt * months
                overpay = max(0.0, total_plan - princ)
                pct = min(paid / princ, 1.0) if princ > 0 else 0.0

                st.markdown(
                    f"""
                    <div class="asset-card">
                        <strong>💳 {r['Title']}</strong> ({rate:.1f}% yillik, {months} oy)<br>
                        Asosiy summa: <code>{princ:,.0f} UZS</code> • Oylik to'lov: <code>{m_pmt:,.0f} UZS</code><br>
                        Jami rejali to'lov: <code>{total_plan:,.0f} UZS</code> (Foiz/ustama: <code>{overpay:,.0f} UZS</code>)<br>
                        To'langan: <span class="pos-green">{paid:,.0f} UZS</span> • Qoldiq: <span class="pos-red">{rem:,.0f} UZS</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.progress(pct)


# ==================== 5. ISLOMIY BANK DEPOZITI (OYLIK QO'SHIMCHA BADAL BILAN) ====================
elif st.session_state.page == "islomiy":
    id_col1, id_col2 = st.columns([1, 1.4])

    with id_col1:
        st.subheader("🕌 Yangi islomiy depozit")
        with st.form("add_islamic_form", clear_on_submit=True):
            dep_title = st.text_input("Depozit / Bank nomi:", placeholder="Masalan: Islomiy omonat (Mudoraba)")
            contract_type = st.selectbox("Shartnoma turi:", ["Mudoraba", "Vakala", "Boshqa islomiy omonat"])
            dep_principal = st.number_input("Boshlang'ich mablag' (UZS):", min_value=0.0, step=500000.0, format="%.0f")
            monthly_add = st.number_input("Oyma-oy qo'shimcha pul (UZS):", min_value=0.0, step=100000.0, format="%.0f")
            exp_rate = st.number_input("Kutilayotgan yillik daromad (%):", min_value=0.0, max_value=100.0, step=0.5, format="%.1f")
            term_m = st.number_input("Muddat (oy):", min_value=1, max_value=120, value=12, step=1)
            dep_start = st.date_input("Ochilgan sana:", value=date.today())

            proj_total_contrib = dep_principal + (monthly_add * int(term_m))
            proj_profit = calc_islamic_projected_profit(dep_principal, monthly_add, exp_rate, int(term_m))
            st.caption(
                f"Muddat oxirida jami badal: **{proj_total_contrib:,.0f} UZS** | "
                f"Taxminiy kutilma foyda: **{proj_profit:,.0f} UZS**"
            )

            if st.form_submit_button("💾 Depozitni saqlash", use_container_width=True):
                if dep_title.strip() and (dep_principal > 0 or monthly_add > 0):
                    ws_d = sh.worksheet("IslamicDeposits")
                    new_id = next_id(load_sheet_df("IslamicDeposits"))
                    ws_d.append_row(
                        [
                            new_id, USER_ID, dep_title.strip(), contract_type,
                            dep_principal, monthly_add, exp_rate, 0.0, int(term_m), str(dep_start),
                        ],
                        value_input_option="USER_ENTERED",
                    )
                    refresh_data()
                    st.success("Islomiy depozit qo'shildi!")
                    st.rerun()
                else:
                    st.warning("Depozit nomi va boshlang'ich yoki oylik summani kiriting!")

        if not df_deposits.empty:
            st.subheader("➕ Depozitga pul qo'shish yoki foyda kiritish")
            with st.form("topup_islamic_form", clear_on_submit=True):
                d_map = {
                    f"{r['Title']} [Asosiy: {to_float(r['Principal']):,.0f}]": r["Id"]
                    for _, r in df_deposits.iterrows()
                }
                chosen_d = st.selectbox("Depozitni tanlang:", list(d_map.keys()))
                action_mode = st.radio("Amal turi:", ["Oylik / qo'shimcha pul qo'shish", "Haqiqiy foyda kiritish"])
                amt_val = st.number_input("Summa (UZS):", min_value=0.0, step=100000.0, format="%.0f")
                sync_tx = st.checkbox("Tranzaksiyalar ro'yxatiga ham qayd etish", value=False)

                if st.form_submit_button("✅ Saqlash", use_container_width=True):
                    if amt_val > 0:
                        did = str(d_map[chosen_d])
                        ws_d = sh.worksheet("IslamicDeposits")
                        all_rows = ws_d.get_all_values()
                        d_title = ""
                        for idx, row in enumerate(all_rows[1:], start=2):
                            if len(row) >= 8 and row[0].strip() == did:
                                d_title = row[2].strip()
                                if action_mode.startswith("Oylik"):
                                    # Ustun 5 = Principal (0-indeksda 4)
                                    new_p = to_float(row[4]) + amt_val
                                    ws_d.update_cell(idx, 5, new_p)
                                else:
                                    # Ustun 8 = ActualProfit (0-indeksda 7)
                                    new_pr = to_float(row[7]) + amt_val
                                    ws_d.update_cell(idx, 8, new_pr)
                                break

                        if sync_tx:
                            ws_t = sh.worksheet("Transactions")
                            tid = next_id(load_sheet_df("Transactions"))
                            if action_mode.startswith("Oylik"):
                                ws_t.append_row(
                                    [tid, USER_ID, f"Depozitga qo'shimcha: {d_title}", amt_val, "Chiqim", "Depozit", str(date.today())],
                                    value_input_option="USER_ENTERED",
                                )
                            else:
                                ws_t.append_row(
                                    [tid, USER_ID, f"Islomiy depozit foydasi: {d_title}", amt_val, "Kirim", "Depozit", str(date.today())],
                                    value_input_option="USER_ENTERED",
                                )

                        refresh_data()
                        st.success("Ma'lumot yangilandi!")
                        st.rerun()

    with id_col2:
        st.subheader("Islomiy depozitlar tahlili")
        if df_deposits.empty:
            st.info("Hozircha islomiy depozitlar kiritilmagan.")
        else:
            tot_dep = sum(to_float(r["Principal"]) for _, r in df_deposits.iterrows())
            tot_m_add = sum(to_float(r.get("MonthlyContribution", 0)) for _, r in df_deposits.iterrows())
            tot_prof = sum(to_float(r["ActualProfit"]) for _, r in df_deposits.iterrows())
            st.markdown(
                f"**Jami asosiy mablag':** `{tot_dep:,.0f} UZS` • "
                f"**Oylik qo'shimchalar yig'indisi:** `{tot_m_add:,.0f} UZS/oy` • "
                f"**Jami foyda:** `{tot_prof:,.0f} UZS`"
            )

            for _, r in df_deposits.iterrows():
                princ = to_float(r["Principal"])
                m_add = to_float(r.get("MonthlyContribution", 0))
                prof = to_float(r["ActualProfit"])
                rate = to_float(r["ExpectedAnnualRate"])
                months = int(to_float(r["TermMonths"]))
                proj_pr = calc_islamic_projected_profit(princ, m_add, rate, months)
                total_end_cap = princ + (m_add * months) + prof
                fact_pct = (prof / princ * 100.0) if princ > 0 else 0.0

                st.markdown(
                    f"""
                    <div class="asset-card">
                        <strong>🕌 {r['Title']}</strong> — <code>{r['ContractType']}</code><br>
                        Joriy asosiy: <code>{princ:,.0f} UZS</code> • Oyma-oy qo'shimcha: <span class="pos-gold">+{m_add:,.0f} UZS/oy</span><br>
                        Muddat: <code>{months} oy</code> • Kutilma foyda ({rate:.1f}%): <code>{proj_pr:,.0f} UZS</code><br>
                        Olingan foyda: <span class="pos-green">+{prof:,.0f} UZS</span> ({fact_pct:.2f}%) • Taxminiy yakuniy jamg'arma: <code>{total_end_cap:,.0f} UZS</code>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


# ==================== 6. AKSIYA PORTFELI ====================
elif st.session_state.page == "aksiya":
    st_col1, st_col2 = st.columns([1, 1.4])

    with st_col1:
        st.subheader("📈 Yangi aksiya qo'shish")
        with st.form("add_stock_form", clear_on_submit=True):
            ticker = st.text_input("Ticker / Belgisi:", placeholder="Masalan: AAPL, UZMK, TSLA")
            comp_name = st.text_input("Kompaniya nomi:", placeholder="Masalan: Apple Inc.")
            shares = st.number_input("Aksiyalar soni (dona):", min_value=0.0, step=1.0, format="%.2f")
            buy_p = st.number_input("O'rtacha xarid narxi (UZS):", min_value=0.0, step=1000.0, format="%.0f")
            curr_p = st.number_input("Joriy bozor narxi (UZS):", min_value=0.0, step=1000.0, format="%.0f")
            div_amt = st.number_input("Olingan dividend (UZS):", min_value=0.0, step=10000.0, format="%.0f")

            if st.form_submit_button("💾 Aksiyani saqlash", use_container_width=True):
                if ticker.strip() and shares > 0 and buy_p > 0:
                    ws_s = sh.worksheet("Stocks")
                    new_id = next_id(load_sheet_df("Stocks"))
                    final_curr = curr_p if curr_p > 0 else buy_p
                    ws_s.append_row(
                        [
                            new_id, USER_ID, ticker.strip().upper(),
                            comp_name.strip() or ticker.strip().upper(),
                            shares, buy_p, final_curr, div_amt,
                        ],
                        value_input_option="USER_ENTERED",
                    )
                    refresh_data()
                    st.success("Aksiya portfelga qo'shildi!")
                    st.rerun()
                else:
                    st.warning("Ticker, aksiya soni va xarid narxini kiriting!")

        if not df_stocks.empty:
            st.subheader("🔄 Narx yoki dividendni yangilash")
            with st.form("upd_stock_form", clear_on_submit=True):
                s_map = {
                    f"{r['Ticker']} — {r['CompanyName']} (Id:{r['Id']})": r["Id"]
                    for _, r in df_stocks.iterrows()
                }
                chosen_s = st.selectbox("Aksiyani tanlang:", list(s_map.keys()))
                new_curr_p = st.number_input("Yangi bozor narxi (UZS):", min_value=0.0, step=1000.0, format="%.0f")
                add_div = st.number_input("Qo'shimcha dividend (UZS):", min_value=0.0, step=10000.0, format="%.0f")

                if st.form_submit_button("💾 Yangilash", use_container_width=True):
                    sid = str(s_map[chosen_s])
                    ws_s = sh.worksheet("Stocks")
                    all_rows = ws_s.get_all_values()
                    for idx, row in enumerate(all_rows[1:], start=2):
                        if len(row) >= 8 and row[0].strip() == sid:
                            if new_curr_p > 0:
                                ws_s.update_cell(idx, 7, new_curr_p)
                            if add_div > 0:
                                ws_s.update_cell(idx, 8, to_float(row[7]) + add_div)
                            break
                    refresh_data()
                    st.success("Aksiya ma'lumotlari yangilandi!")
                    st.rerun()

    with st_col2:
        st.subheader("Aksiyalar hisob-kitobi")
        if df_stocks.empty:
            st.info("Hozircha aksiyalar kiritilmagan.")
        else:
            tot_cost = 0.0
            tot_val = 0.0
            tot_pnl = 0.0

            for _, r in df_stocks.iterrows():
                sh_cnt = to_float(r["Shares"])
                b_price = to_float(r["BuyPrice"])
                c_price = to_float(r["CurrentPrice"])
                div_val = to_float(r["Dividend"])

                cost = sh_cnt * b_price
                val = sh_cnt * c_price
                pnl = (val - cost) + div_val
                pnl_pct = (pnl / cost * 100.0) if cost > 0 else 0.0

                tot_cost += cost
                tot_val += val
                tot_pnl += pnl

                pnl_cls = "pos-green" if pnl >= 0 else "pos-red"
                sign = "+" if pnl >= 0 else ""

                st.markdown(
                    f"""
                    <div class="asset-card">
                        <strong>📈 {r['Ticker']}</strong> — {r['CompanyName']}<br>
                        Soni: <code>{sh_cnt:,.2f}</code> • Xarid: <code>{b_price:,.0f}</code> • Joriy: <code>{c_price:,.0f}</code><br>
                        Jami qiymat: <code>{val:,.0f} UZS</code> • Dividend: <code>{div_val:,.0f} UZS</code><br>
                        Foyda/Zarar: <span class="{pnl_cls}">{sign}{pnl:,.0f} UZS ({sign}{pnl_pct:.1f}%)</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            tot_cls = "pos-green" if tot_pnl >= 0 else "pos-red"
            tot_sign = "+" if tot_pnl >= 0 else ""
            st.markdown(
                f"**Portfel qiymati:** `{tot_val:,.0f} UZS` (Xarid: `{tot_cost:,.0f}`) | "
                f"**Umumiy natija:** <span class='{tot_cls}'>{tot_sign}{tot_pnl:,.0f} UZS</span>",
                unsafe_allow_html=True,
            )


# ==================== 7. OBLIGATSIYALAR ====================
elif st.session_state.page == "obligatsiya":
    bo_col1, bo_col2 = st.columns([1, 1.4])

    with bo_col1:
        st.subheader("📜 Yangi obligatsiya")
        with st.form("add_bond_form", clear_on_submit=True):
            b_title = st.text_input("Obligatsiya nomi / emitent:", placeholder="Masalan: Davlat / Korporativ obligatsiya")
            b_qty = st.number_input("Donalar soni:", min_value=1, step=1, value=1)
            b_face = st.number_input("Nominal qiymati (1 dona, UZS):", min_value=0.0, step=100000.0, format="%.0f")
            b_buy = st.number_input("Xarid narxi (1 dona, UZS):", min_value=0.0, step=100000.0, format="%.0f")
            b_coupon = st.number_input("Yillik kupon stavkasi (%):", min_value=0.0, max_value=100.0, step=0.5, format="%.1f")
            b_mat = st.date_input("So'ndirilish sanasi:")

            est_annual_coupon = int(b_qty) * b_face * (b_coupon / 100.0)
            st.caption(f"Taxminiy yillik kupon daromadi: **{est_annual_coupon:,.0f} UZS**")

            if st.form_submit_button("💾 Obligatsiyani saqlash", use_container_width=True):
                if b_title.strip() and b_face > 0 and b_buy > 0:
                    ws_bo = sh.worksheet("Bonds")
                    new_id = next_id(load_sheet_df("Bonds"))
                    ws_bo.append_row(
                        [
                            new_id, USER_ID, b_title.strip(), int(b_qty),
                            b_face, b_buy, b_coupon, 0.0, str(b_mat),
                        ],
                        value_input_option="USER_ENTERED",
                    )
                    refresh_data()
                    st.success("Obligatsiya qo'shildi!")
                    st.rerun()
                else:
                    st.warning("Obligatsiya nomi, nominal va xarid narxini kiriting!")

        if not df_bonds.empty:
            st.subheader("💵 Kupon to'lovini kiritish")
            with st.form("coupon_bond_form", clear_on_submit=True):
                bo_map = {f"{r['Title']} (Id:{r['Id']})": r["Id"] for _, r in df_bonds.iterrows()}
                chosen_bo = st.selectbox("Obligatsiyani tanlang:", list(bo_map.keys()))
                coup_amt = st.number_input("Olingan kupon summasi (UZS):", min_value=0.0, step=50000.0, format="%.0f")

                if st.form_submit_button("➕ Kuponni qo'shish", use_container_width=True):
                    if coup_amt > 0:
                        boid = str(bo_map[chosen_bo])
                        ws_bo = sh.worksheet("Bonds")
                        all_rows = ws_bo.get_all_values()
                        for idx, row in enumerate(all_rows[1:], start=2):
                            if len(row) >= 8 and row[0].strip() == boid:
                                ws_bo.update_cell(idx, 8, to_float(row[7]) + coup_amt)
                                break
                        refresh_data()
                        st.success("Kupon daromadi qo'shildi!")
                        st.rerun()

    with bo_col2:
        st.subheader("Obligatsiyalar hisob-kitobi")
        if df_bonds.empty:
            st.info("Hozircha obligatsiyalar kiritilmagan.")
        else:
            tot_nom = 0.0
            tot_coup = 0.0
            for _, r in df_bonds.iterrows():
                qty = int(to_float(r["Quantity"]))
                face = to_float(r["FaceValue"])
                buy_p = to_float(r["BuyPrice"])
                c_rate = to_float(r["CouponRate"])
                c_rec = to_float(r["CouponReceived"])

                nom_total = qty * face
                buy_total = qty * buy_p
                annual_coup = nom_total * (c_rate / 100.0)
                total_gain = (nom_total - buy_total) + c_rec

                tot_nom += nom_total
                tot_coup += c_rec

                g_cls = "pos-green" if total_gain >= 0 else "pos-red"
                g_sign = "+" if total_gain >= 0 else ""

                st.markdown(
                    f"""
                    <div class="asset-card">
                        <strong>📜 {r['Title']}</strong> ({qty} dona)<br>
                        Nominal: <code>{nom_total:,.0f} UZS</code> • Xarid: <code>{buy_total:,.0f} UZS</code><br>
                        Kupon: <code>{c_rate:.1f}%</code> (Yillik: <code>{annual_coup:,.0f} UZS</code>) • Olingan kupon: <span class="pos-green">+{c_rec:,.0f} UZS</span><br>
                        Jami foyda/diskont: <span class="{g_cls}">{g_sign}{total_gain:,.0f} UZS</span> • Muddat: <code>{r.get('MaturityDate', '-')}</code>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown(f"**Jami nominal:** `{tot_nom:,.0f} UZS` | **Jami olingan kupon:** `{tot_coup:,.0f} UZS`")


# ==================== 8. KATEGORIYALAR ====================
elif st.session_state.page == "kategoriyalar":
    c_col1, c_col2 = st.columns([1, 1.4])
    with c_col1:
        st.subheader("⚙️ Yangi kategoriya")
        with st.form("cat_form", clear_on_submit=True):
            c_type = st.selectbox("Turi:", ["Chiqim", "Kirim"])
            c_name = st.text_input("Kategoriya nomi:", placeholder="Masalan: Taksi, Kredit, Dividend")
            if st.form_submit_button("➕ Qo'shish", use_container_width=True):
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


# ==================== 9. HISOBOT (1 KUN, HAFTA, OY, YIL) VA XARAJATLAR TAHLILI ====================
elif st.session_state.page == "tarix":
    st.subheader("📋 Davrlar bo'yicha hisobot va xarajatlar tahlili")

    period_options = ["1 kunlik", "1 haftalik", "1 oylik", "1 yillik", "Barchasi"]
    selected_period = st.selectbox("Hisobot davrini tanlang:", period_options, index=2)

    p_df = filter_by_period(df_trans, selected_period) if not df_trans.empty else pd.DataFrame()
    p_in = p_df[p_df["Type"].str.lower() == "kirim"]["AmountNum"].sum() if not p_df.empty else 0.0
    p_out = p_df[p_df["Type"].str.lower() == "chiqim"]["AmountNum"].sum() if not p_df.empty else 0.0
    p_net = p_in - p_out

    st.markdown(
        f"**{selected_period} hisobot:** Kirim: <span class='pos-green'>+{p_in:,.0f} UZS</span> • "
        f"Chiqim: <span class='pos-red'>-{p_out:,.0f} UZS</span> • "
        f"Sof farq: <code>{p_net:+,.0f} UZS</code>",
        unsafe_allow_html=True,
    )

    # ---- XARAJATLAR TAHLILI: ENG QIMMAT, O'RTACHA, ENG KAM ----
    st.markdown("### 🔍 Xarajatlar tahlili (Eng qimmat / O'rtacha / Eng kam)")
    exp_p = p_df[p_df["Type"].str.lower() == "chiqim"].copy() if not p_df.empty else pd.DataFrame()

    if exp_p.empty:
        st.info(f"'{selected_period}' davrida chiqimlar topilmadi.")
    else:
        max_row = exp_p.loc[exp_p["AmountNum"].idxmax()]
        min_row = exp_p.loc[exp_p["AmountNum"].idxmin()]
        avg_val = exp_p["AmountNum"].mean()
        exp_count = len(exp_p)

        st.markdown(
            f"""
            <div class="analytics-grid">
                <div class="stat-box">
                    <div class="stat-title">🔥 Eng qimmat xarajat</div>
                    <div class="stat-num pos-red">{max_row['AmountNum']:,.0f} UZS</div>
                    <div class="stat-sub">{max_row['Title']} ({max_row['Category']}) • {max_row['Date']}</div>
                </div>
                <div class="stat-box">
                    <div class="stat-title">⚖️ O'rtacha xarajat</div>
                    <div class="stat-num pos-gold">{avg_val:,.0f} UZS</div>
                    <div class="stat-sub">{exp_count} ta xarajat bo'yicha o'rtacha</div>
                </div>
                <div class="stat-box">
                    <div class="stat-title">🧊 Eng kam xarajat</div>
                    <div class="stat-num pos-green">{min_row['AmountNum']:,.0f} UZS</div>
                    <div class="stat-sub">{min_row['Title']} ({min_row['Category']}) • {min_row['Date']}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ---- EXCEL HISOBOT (BARCHA DAVRLAR VA TAHLIL BILAN) ----
    wb = openpyxl.Workbook()
    h_fill = PatternFill(start_color="39608F", end_color="39608F", fill_type="solid")
    h_font = Font(color="FFFFFF", bold=True)

    ws_ana = wb.active
    ws_ana.title = "Davrlar va Tahlil"
    ws_ana.append(["Davr", "Jami Kirim", "Jami Chiqim", "Sof Farq", "Eng Qimmat Xarajat", "O'rtacha Xarajat", "Eng Kam Xarajat"])
    for prd in ["1 kunlik", "1 haftalik", "1 oylik", "1 yillik", "Barchasi"]:
        sub = filter_by_period(df_trans, prd) if not df_trans.empty else pd.DataFrame()
        c_in = sub[sub["Type"].str.lower() == "kirim"]["AmountNum"].sum() if not sub.empty else 0.0
        c_out_df = sub[sub["Type"].str.lower() == "chiqim"] if not sub.empty else pd.DataFrame()
        c_out = c_out_df["AmountNum"].sum() if not c_out_df.empty else 0.0
        c_max = c_out_df["AmountNum"].max() if not c_out_df.empty else 0.0
        c_avg = round(c_out_df["AmountNum"].mean(), 2) if not c_out_df.empty else 0.0
        c_min = c_out_df["AmountNum"].min() if not c_out_df.empty else 0.0
        ws_ana.append([prd, c_in, c_out, c_in - c_out, c_max, c_avg, c_min])

    ws1 = wb.create_sheet("Tranzaksiyalar")
    ws1.append(["ID", "Nomi", "Summa", "Turi", "Kategoriya", "Sana"])
    if not df_trans.empty:
        for _, r in df_trans.iterrows():
            ws1.append([r["Id"], r["Title"], r["AmountNum"], r["Type"], r["Category"], r["Date"]])

    ws2 = wb.create_sheet("Kreditlar")
    ws2.append(["Nomi", "Asosiy", "Foiz %", "Oy", "Oylik To'lov", "To'langan", "Qoldiq"])
    if not df_credits.empty:
        for _, r in df_credits.iterrows():
            ws2.append([
                r["Title"], to_float(r["Principal"]), to_float(r["AnnualRate"]),
                int(to_float(r["Months"])), to_float(r["MonthlyPayment"]),
                to_float(r["PaidAmount"]), to_float(r["RemainingPrincipal"]),
            ])

    ws3 = wb.create_sheet("Islomiy Depozit")
    ws3.append(["Nomi", "Shartnoma", "Asosiy", "Oylik Qo'shimcha", "Kutilma %", "Olingan Foyda", "Muddat"])
    if not df_deposits.empty:
        for _, r in df_deposits.iterrows():
            ws3.append([
                r["Title"], r["ContractType"], to_float(r["Principal"]),
                to_float(r.get("MonthlyContribution", 0)), to_float(r["ExpectedAnnualRate"]),
                to_float(r["ActualProfit"]), int(to_float(r["TermMonths"])),
            ])

    ws4 = wb.create_sheet("Aksiyalar")
    ws4.append(["Ticker", "Kompaniya", "Soni", "Xarid Narxi", "Joriy Narx", "Dividend", "Foyda/Zarar"])
    if not df_stocks.empty:
        for _, r in df_stocks.iterrows():
            sh_c = to_float(r["Shares"])
            bp = to_float(r["BuyPrice"])
            cp = to_float(r["CurrentPrice"])
            dv = to_float(r["Dividend"])
            pnl = (sh_c * cp - sh_c * bp) + dv
            ws4.append([r["Ticker"], r["CompanyName"], sh_c, bp, cp, dv, pnl])

    ws5 = wb.create_sheet("Obligatsiyalar")
    ws5.append(["Nomi", "Soni", "Nominal", "Xarid Narxi", "Kupon %", "Olingan Kupon", "Muddat"])
    if not df_bonds.empty:
        for _, r in df_bonds.iterrows():
            ws5.append([
                r["Title"], int(to_float(r["Quantity"])), to_float(r["FaceValue"]),
                to_float(r["BuyPrice"]), to_float(r["CouponRate"]), to_float(r["CouponReceived"]),
                r.get("MaturityDate", ""),
            ])

    for ws in [ws_ana, ws1, ws2, ws3, ws4, ws5]:
        for cell in ws[1]:
            cell.fill = h_fill
            cell.font = h_font

    excel_buf = io.BytesIO()
    wb.save(excel_buf)
    excel_buf.seek(0)

    st.download_button(
        label="📥 To'liq Excel Hisobotni Yuklab Olish (.xlsx)",
        data=excel_buf,
        file_name=f"Moliya_Hisobot_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if not p_df.empty:
        filter_type = st.selectbox("Tur bo'yicha filtr:", ["Barchasi", "Kirim", "Chiqim"])
        view_df = p_df.copy()
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