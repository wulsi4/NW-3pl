import streamlit as st
import pandas as pd
import sqlite3
import io
import random

st.set_page_config(page_title="3PL Nord Wheel", layout="wide")

# --- КАСТОМНЫЙ CSS ДЛЯ БЕЛО-СИНЕЙ СТИЛИСТИКИ И КРУГЛЫХ КНОПОК ---
st.markdown("""
    <style>
    .stApp {
        background-color: #f8fafc;
    }
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
        padding-top: 20px;
    }
    div.stButton > button, div.stFormSubmitButton > button, .stDownloadButton > button {
        border-radius: 30px !important;
        background-color: #0284c7 !important;
        color: white !important;
        font-weight: 600;
        border: none;
        padding: 0.6rem 1.2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 132, 199, 0.2);
        transition: all 0.3s ease;
        width: 100%;
    }
    div.stButton > button:hover, div.stFormSubmitButton > button:hover, .stDownloadButton > button:hover {
        background-color: #0369a1 !important;
        box-shadow: 0 6px 8px -1px rgba(3, 105, 161, 0.3);
    }
    h1, h2, h3 {
        color: #0f172a !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ SQLite ---
def get_connection():
    conn = sqlite3.connect("warehouse.db", check_same_thread=False)
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        tariff_A REAL,
        tariff_B REAL
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS locations (
        address TEXT PRIMARY KEY,
        zone TEXT,
        status TEXT DEFAULT 'FREE'
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallets (
        lpn TEXT PRIMARY KEY,
        client TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallet_locations (
        lpn TEXT,
        address TEXT,
        FOREIGN KEY(lpn) REFERENCES pallets(lpn),
        FOREIGN KEY(address) REFERENCES locations(address),
        PRIMARY KEY(lpn, address)
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallet_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        lpn TEXT,
        sku TEXT,
        item_name TEXT,
        qty INTEGER,
        FOREIGN KEY(lpn) REFERENCES pallets(lpn)
    )''')
    
    conn.commit()
    
    cursor.execute("SELECT COUNT(*) FROM locations")
    if cursor.fetchone()[0] == 0:
        locs = []
        zones_a_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M']
        # Упрощенная структура Зоны А: Линия, Секция (1-6), Ярус (1-5) без лишней глубины
        for aisle in zones_a_letters:
            for sec in range(1, 7):
                for tier in range(1, 6):
                    addr = f"A-{aisle}-{sec:02d}-{tier}"
                    locs.append((addr, "A", "FREE"))
        # Зона B: 150 паллетомест
        for i in range(1, 151):
            addr = f"B-{i:03d}"
            locs.append((addr, "B", "FREE"))
        cursor.executemany("INSERT OR IGNORE INTO locations (address, zone, status) VALUES (?, ?, ?)", locs)
        conn.commit()
    conn.close()

init_db()

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    processed_data = output.getvalue()
    return processed_data

if "page" not in st.session_state:
    st.session_state.page = "clients"

# --- ШАПКА ПРИЛОЖЕНИЯ ---
st.markdown("<h1 style='text-align: center; color: #0284c7;'>🛞 3PL Nord Wheel</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748b; margin-top: -15px;'>Система ответственного хранения грузов и биллинга</p>", unsafe_allow_html=True)
st.markdown("---")

# --- ЛЕВАЯ ПАНЕЛЬ С КНОПКАМИ И ЛОГОТИПОМ ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        st.markdown("<h2 style='text-align: center; color: #0284c7;'>🛞 Nord Wheel</h2>", unsafe_allow_html=True)

    st.markdown("<p style='text-align: center; font-size: 12px; color: #64748b;'>Логистический комплекс</p>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### Меню управления")

    if st.button("👥 Клиенты и тарифы", use_container_width=True):
        st.session_state.page = "clients"
    if st.button("🗺️ Карта склада и ячейки", use_container_width=True):
        st.session_state.page = "map"
    if st.button("📥 Приемка и размещение", use_container_width=True):
        st.session_state.page = "inbound"
    if st.button("🖨️ Печать листов А4 и Отчеты", use_container_width=True):
        st.session_state.page = "reports"
    if st.button("📊 Биллинг (Снапшот)", use_container_width=True):
        st.session_state.page = "billing"

conn = get_connection()

# --- РАЗДЕЛ 1: КЛИЕНТЫ И ТАРИФЫ ---
if st.session_state.page == "clients":
    st.header("👥 Регистрация и карточки поклажедателей")
    
    with st.expander("➕ Добавить нового клиента"):
        with st.form("add_client_form"):
            c_name = st.text_input("Название компании")
            t_a = st.number_input("Тариф Зона А (руб/паллето-день)", value=30.0)
            t_b = st.number_input("Тариф Зона B (руб/паллето-день)", value=20.0)
            submitted = st.form_submit_button("Зарегистрировать клиента")
            if submitted and c_name:
                try:
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO clients (name, tariff_A, tariff_B) VALUES (?, ?, ?)", (c_name, t_a, t_b))
                    conn.commit()
                    st.success(f"Клиент '{c_name}' успешно зарегистрирован!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Клиент с таким названием уже существует!")

    st.subheader("Сводная таблица контрагентов")
    clients_df = pd.read_sql("SELECT * FROM clients", conn)
    if not clients_df.empty:
        summary_data = []
        cursor = conn.cursor()
        for idx, row in clients_df.iterrows():
            c_name = row['name']
            cursor.execute("SELECT COUNT(DISTINCT lpn) FROM pallets WHERE client = ?", (c_name,))
            pallets_count = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(DISTINCT pl.address) FROM pallets p JOIN pallet_locations pl ON p.lpn = pl.lpn WHERE p.client = ?", (c_name,))
            cells_count = cursor.fetchone()[0]
            
            summary_data.append({
                "Клиент": c_name,
                "Тариф А (руб)": row['tariff_A'],
                "Тариф B (руб)": row['tariff_B'],
                "Всего паллет": pallets_count,
                "Занято ячеек": cells_count
            })
        
        sum_df = pd.DataFrame(summary_data)
        st.dataframe(sum_df, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🔍 Карточка клиента (Проваливание в детальную информацию)")
        selected_client_card = st.selectbox("Выберите клиента для просмотра:", clients_df['name'].tolist())
        
        if selected_client_card:
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                st.markdown(f"**Компания:** {selected_client_card}")
                c_tariffs = clients_df[clients_df['name'] == selected_client_card].iloc[0]
                st.write(f"• Тариф Зона А: {c_tariffs['tariff_A']} руб./день")
                st.write(f"• Тариф Зона B: {c_tariffs['tariff_B']} руб./день")
            
            with col_c2:
                st.markdown(f"**Размещение на складе:**")
            
            client_pallets_df = pd.read_sql("""
                SELECT p.lpn as 'LPN Паллеты', GROUP_CONCAT(pl.address) as 'Ячейки'
                FROM pallets p
                LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn
                WHERE p.client = ?
                GROUP BY p.lpn
            """, conn, params=(selected_client_card,))
            
            if not client_pallets_df.empty:
                st.dataframe(client_pallets_df, use_container_width=True)
                
                st.markdown("**Позиционная номенклатура на паллетах клиента:**")
                client_items_df = pd.read_sql("""
                    SELECT pi.lpn as 'LPN', pi.sku as 'SKU / Артикул', pi.item_name as 'Наименование', pi.qty as 'Количество'
                    FROM pallet_items pi
                    JOIN pallets p ON pi.lpn = p.lpn
                    WHERE p.client = ?
                """, conn, params=(selected_client_card,))
                if not client_items_df.empty:
                    st.dataframe(client_items_df, use_container_width=True)
                else:
                    st.info("Номенклатура для паллет этого клиента не заполнена.")
            else:
                st.info("У данного клиента нет активных паллет на складе.")
    else:
        st.info("Список клиентов пуст.")

# --- РАЗДЕЛ 2: КАРТА СКЛАДА С ПОИСКОМ ПО КЛИЕНТАМ ---
elif st.session_state.page == "map":
    st.header("🗺️ Интерактивная карта и поиск ячеек по клиентам")
    
    clients_list = ["Все клиенты"] + pd.read_sql("SELECT name FROM clients", conn)["name"].tolist()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        client_map_filter = st.selectbox("Фильтр по клиенту:", clients_list)
    with col2:
        zone_filter = st.selectbox("Фильтр по зоне", ["Все", "Зона А (Стеллажи)", "Зона B (2-й этаж)"])
    with col3:
        status_filter = st.selectbox("Статус ячейки", ["Все", "Свободные (FREE)", "Занятые (OCCUPIED)"])
    
    if client_map_filter != "Все клиенты":
        query = f"""
            SELECT l.address, l.zone, l.status, p.lpn, p.client 
            FROM locations l
            JOIN pallet_locations pl ON l.address = pl.address
            JOIN pallets p ON pl.lpn = p.lpn
            WHERE p.client = '{client_map_filter}'
        """
        if zone_filter == "Зона А (Стеллажи)":
            query += " AND l.zone = 'A'"
        elif zone_filter == "Зона B (2-й этаж)":
            query += " AND l.zone = 'B'"
    else:
        query = "SELECT address, zone, status, '' as lpn, '' as client FROM locations WHERE 1=1"
        if zone_filter == "Зона А (Стеллажи)":
            query += " AND zone = 'A'"
        elif zone_filter == "Зона B (2-й этаж)":
            query += " AND zone = 'B'"
            
        if status_filter == "Свободные (FREE)":
            query += " AND status = 'FREE'"
        elif status_filter == "Занятые (OCCUPIED)":
            query += " AND status = 'OCCUPIED'"
        
    loc_df = pd.read_sql(query, conn)
    search_query = st.text_input("Поиск по конкретному адресу ячейки (например, A-A-01-1 или B-042)")
    if search_query:
        loc_df = loc_df[loc_df["address"].str.contains(search_query, case=False)]

    st.metric("Найдено ячеек", len(loc_df))
    st.dataframe(loc_df.head(150), use_container_width=True)

# --- РАЗДЕЛ 3: ПРИЕМКА И КАЛЕНДАРНЫЙ ВЫБОР ЯЧЕЕК ---
elif st.session_state.page == "inbound":
    st.header("📥 Приемка паллеты и выбор ячейки в календарном виде")
    st.write("Выберите свободные ячейки в визуальной матрице (как в календаре), внесите состав и закрепите за паллетой.")
    
    clients_list = pd.read_sql("SELECT name FROM clients", conn)["name"].tolist()
    
    if not clients_list:
        st.warning("Сначала добавьте хотя бы одного клиента в разделе «Клиенты и тарифы».")
    else:
        with st.form("calendar_inbound_form"):
            selected_client = st.selectbox("Поклажедатель", clients_list)
            lpn_code = st.text_input("Номер/LPN паллеты", value=f"LPN-{random.randint(1000, 9999)}")
            
            st.markdown("---")
            st.markdown("### 📅 Календарный выбор ячеек склада")
            
            sel_zone = st.radio("Зона размещения:", ["Зона А (Стеллажи)", "Зона B (2-й этаж)"], horizontal=True)
            
            selected_cells = []
            cursor = conn.cursor()
            
            if sel_zone == "Зона А (Стеллажи)":
                c_a, c_b = st.columns(2)
                with c_a:
                    chosen_aisle = st.selectbox("Линия:", ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M'])
                with c_b:
                    chosen_sec = st.selectbox("Секция:", [1, 2, 3, 4, 5, 6])
                
                st.markdown("**Сетка ярусов (выберите ячейку кликом):**")
                # Выводим матрицу 5 ярусов (по 1 ячейке на ярус для этой секции)
                for tier in range(5, 0, -1):
                    addr = f"A-{chosen_aisle}-{chosen_sec:02d}-{tier}"
                    cursor.execute("SELECT status FROM locations WHERE address = ?", (addr,))
                    res = cursor.fetchone()
                    status = res[0] if res else 'FREE'
                    
                    if status == 'FREE':
                        if st.checkbox(f"Ярус {tier} — Адрес: {addr}", key=f"grid_cell_{addr}"):
                            selected_cells.append(addr)
                    else:
                        st.markdown(f"<span style='color: #991b1b; background-color: #fee2e2; padding: 4px 8px; border-radius: 4px;'>Ярус {tier} ({addr}) — <b>Занята</b></span>", unsafe_allow_html=True)
            else:
                st.markdown("**Матрица паллетомест 2-го этажа (150 мест):**")
                cursor.execute("SELECT address, status FROM locations WHERE zone = 'B' ORDER BY address")
                b_locs = cursor.fetchall()
                
                # Выводим сеткой по 10 штук в ряд (календарный вид)
                for i in range(0, len(b_locs), 10):
                    cols = st.columns(10)
                    chunk = b_locs[i:i+10]
                    for idx, (addr, status) in enumerate(chunk):
                        with cols[idx]:
                            short_n = addr.split('-')[1]
                            if status == 'FREE':
                                if st.checkbox(f"{short_n}", key=f"b_cell_{addr}"):
                                    selected_cells.append(addr)
                            else:
                                st.markdown(f"<div style='background-color:#fee2e2; color:#991b1b; padding:4px; text-align:center; font-size:10px; border-radius:3px;'>{short_n}</div>", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 📦 Позиционная номенклатура на паллете (для будущей сборки)")
            
            col_i1, col_i2, col_i3 = st.columns(3)
            with col_i1:
                sku1 = st.text_input("SKU / Артикул 1", "SKU-001")
                name1 = st.text_input("Наименование 1", "Товар А")
                qty1 = st.number_input("Кол-во 1", min_value=0, value=10)
            with col_i2:
                sku2 = st.text_input("SKU / Артикул 2 (опц.)")
                name2 = st.text_input("Наименование 2 (опц.)")
                qty2 = st.number_input("Кол-во 2", min_value=0, value=0)
            with col_i3:
                sku3 = st.text_input("SKU / Артикул 3 (опц.)")
                name3 = st.text_input("Наименование 3 (опц.)")
                qty3 = st.number_input("Кол-во 3", min_value=0, value=0)
                
            submit_calendar = st.form_submit_button("Оприходовать и занять ячейку")
            
            if submit_calendar:
                if not selected_cells:
                    st.error("Выберите хотя бы одну свободную ячейку в матрице!")
                else:
                    try:
                        cursor.execute("INSERT INTO pallets (lpn, client) VALUES (?, ?)", (lpn_code, selected_client))
                        
                        for cell in selected_cells:
                            cursor.execute("INSERT INTO pallet_locations (lpn, address) VALUES (?, ?)", (lpn_code, cell))
                            cursor.execute("UPDATE locations SET status = 'OCCUPIED' WHERE address = ?", (cell,))
                            
                        items_to_add = [(lpn_code, sku1, name1, qty1)]
                        if sku2 and qty2 > 0:
                            items_to_add.append((lpn_code, sku2, name2, qty2))
                        if sku3 and qty3 > 0:
                            items_to_add.append((lpn_code, sku3, name3, qty3))
                            
                        cursor.executemany("INSERT INTO pallet_items (lpn, sku, item_name, qty) VALUES (?, ?, ?, ?)", items_to_add)
                        
                        conn.commit()
                        st.success(f"Паллета {lpn_code} успешно закреплена за ячейками: {', '.join(selected_cells)}!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Паллета с таким LPN уже зарегистрирована на складе!")

    st.subheader("Текущие паллеты на складе")
    pallets_df = pd.read_sql("""
        SELECT p.lpn as 'LPN', p.client as 'Клиент', GROUP_CONCAT(pl.address) as 'Ячейки' 
        FROM pallets p 
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn 
        GROUP BY p.lpn
    """, conn)
    
    if not pallets_df.empty:
        st.dataframe(pallets_df, use_container_width=True)
    else:
        st.info("Палет на складе пока нет.")

# --- РАЗДЕЛ 4: ПЕЧАТЬ ЛИСТОВ А4 И ОТЧЕТЫ ---
elif st.session_state.page == "reports":
    st.header("🖨️ Паллетные листы А4 с попозиционной номенклатурой")
    
    pallets_list = pd.read_sql("SELECT lpn FROM pallets", conn)["lpn"].tolist()
    if not pallets_list:
        st.info("Нет созданных паллет.")
    else:
        p_lpn = st.selectbox("Выберите паллету для печати A4 листа", pallets_list)
        
        cursor = conn.cursor()
        cursor.execute("SELECT client FROM pallets WHERE lpn = ?", (p_lpn,))
        p_data = cursor.fetchone()
        cursor.execute("SELECT address FROM pallet_locations WHERE lpn = ?", (p_lpn,))
        locs_data = [row[0] for row in cursor.fetchall()]
        
        items_data = pd.read_sql("SELECT sku as 'Артикул', item_name as 'Наименование', qty as 'Количество' FROM pallet_items WHERE lpn = ?", conn, params=(p_lpn,))
        
        if p_data:
            st.markdown("---")
            st.markdown(f"<h2 style='text-align: center; color: #0284c7;'>📄 ПАЛЛЕТНЫЙ ЛИСТ (A4) — 3PL Nord Wheel</h2>", unsafe_allow_html=True)
            st.markdown(f"### **Клиент:** {p_data[0]}")
            st.markdown(f"<h1>ID ПАЛЛЕТЫ: {p_lpn}</h1>", unsafe_allow_html=True)
            st.markdown(f"**Ячейки размещения:** {', '.join(locs_data)}")
            st.markdown("#### Позиционный состав груза (для сборки):")
            st.dataframe(items_data, use_container_width=True)
            st.markdown("---")
            st.caption("Штрихкод для сканирования ТСД: [ |||||||||||||||||||||||||| ]")
            st.button("🖨️ Печать листа (PDF/A4)")
            
    st.subheader("Экспорт всех данных в Excel")
    full_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'Номер паллеты', p.client as 'Клиент', GROUP_CONCAT(pl.address) as 'Ячейки' 
        FROM pallets p 
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn 
        GROUP BY p.lpn
    """, conn)
    
    if not full_pallets_df.empty:
        excel_data = convert_df_to_excel(full_pallets_df)
        st.download_button(
            label="📥 Скачать общий отчет в Excel",
            data=excel_data,
            file_name="NordWheel_warehouse_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- РАЗДЕЛ 5: БИЛЛИНГ И СНАПШОТ ---
elif st.session_state.page == "billing":
    st.header("📊 Автоматический расчет хранения (Снапшот остатков)")
    st.write("Расчет строится на основе фактически занятых ячеек (паллето-мест) с учетом индивидуальных тарифов.")
    
    if st.button("Сделать срез (Snapshot) и рассчитать счета"):
        billing_query = """
            SELECT p.client, l.zone, COUNT(DISTINCT pl.address) as occupied_slots
            FROM pallets p
            JOIN pallet_locations pl ON p.lpn = pl.lpn
            JOIN locations l ON pl.address = l.address
            GROUP BY p.client, l.zone
        """
        snapshot_df = pd.read_sql(billing_query, conn)
        clients_df = pd.read_sql("SELECT * FROM clients", conn)
        
        if not snapshot_df.empty and not clients_df.empty:
            client_tariffs = clients_df.set_index("name").to_dict(orient="index")
            
            billing_results = []
            clients_in_snap = snapshot_df["client"].unique()
            
            for c_name in clients_in_snap:
                t_a = client_tariffs.get(c_name, {}).get("tariff_A", 30.0)
                t_b = client_tariffs.get(c_name, {}).get("tariff_B", 20.0)
                
                slots_a = snapshot_df[(snapshot_df["client"] == c_name) & (snapshot_df["zone"] == "A")]["occupied_slots"].sum()
                slots_b = snapshot_df[(snapshot_df["client"] == c_name) & (snapshot_df["zone"] == "B")]["occupied_slots"].sum()
                
                cost_a = slots_a * t_a
                cost_b = slots_b * t_b
                total = cost_a + cost_b
                
                billing_results.append({
                    "Клиент": c_name,
                    "Ячеек в Зоне А": int(slots_a),
                    "Ячеек в Зоне B": int(slots_b),
                    "Сумма за Зону А (руб.)": cost_a,
                    "Сумма за Зону B (руб.)": cost_b,
                    "ИТОГО К ОПЛАТЕ (руб.)": total
                })
                
            res_df = pd.DataFrame(billing_results)
            st.dataframe(res_df, use_container_width=True)
            
            billing_excel = convert_df_to_excel(res_df)
            st.download_button(
                label="📥 Скачать счет/акт биллинга в Excel",
                data=billing_excel,
                file_name="NordWheel_billing_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Нет данных для расчета (склад пуст или не заведены тарифы клиентов).")

conn.close()
