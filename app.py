import streamlit as st
import pandas as pd
import sqlite3
import io
import random

st.set_page_config(page_title="3PL Nord Wheel", layout="wide")

# --- КАСТОМНЫЙ CSS ДЛЯ БЕЛО-СИНЕ-КРАСНОЙ СТИЛИСТИКИ И КРУГЛЫХ КНОПОК ---
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
        tariff_A REAL DEFAULT 30.0,
        tariff_B REAL DEFAULT 20.0
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
        for aisle in zones_a_letters:
            for sec in range(1, 7):
                for tier in range(1, 6):
                    addr = f"A-{aisle}-{sec:02d}-{tier}"
                    locs.append((addr, "A", "FREE"))
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
if "selected_client" not in st.session_state:
    st.session_state.selected_client = None

# --- ШАПКА ПРИЛОЖЕНИЯ В БЕЛО-СИНЕ-КРАСНЫХ ТОНАХ ---
st.markdown("""
    <div style='background: linear-gradient(135deg, #1d4ed8 0%, #0284c7 50%, #dc2626 100%); padding: 25px; border-radius: 16px; text-align: center; color: white; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); margin-bottom: 25px;'>
        <h1 style='color: white; margin: 0; font-size: 34px; font-weight: 800; text-shadow: 0 2px 4px rgba(0,0,0,0.2);'>🛞 3PL Nord Wheel</h1>
        <p style='margin: 8px 0 0 0; font-size: 15px; color: #f1f5f9; font-weight: 500;'>Профессиональная система ответственного хранения грузов и биллинга</p>
    </div>
""", unsafe_allow_html=True)

# --- ЛЕВАЯ ПАНЕЛЬ С ЛОГОТИПОМ И КНОПКАМИ ---
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
        st.session_state.selected_client = None
    if st.button("🗺️ Карта склада и ячейки", use_container_width=True):
        st.session_state.page = "map"
    if st.button("📥 Приемка партии (1С стиль)", use_container_width=True):
        st.session_state.page = "inbound"
    if st.button("🖨️ Печать листов А4 и Отчеты", use_container_width=True):
        st.session_state.page = "reports"
    if st.button("📊 Биллинг (Снапшот)", use_container_width=True):
        st.session_state.page = "billing"

conn = get_connection()

# --- РАЗДЕЛ 1: КЛИЕНТЫ И ТАРИФЫ (И СТРАНИЦА КАРТОЧКИ КЛИЕНТА) ---
if st.session_state.page == "clients":
    st.header("👥 Регистрация и список поклажедателей")
    
    with st.expander("➕ Добавить нового клиента"):
        with st.form("add_client_form"):
            c_name = st.text_input("Название компании")
            submitted = st.form_submit_button("Зарегистрировать клиента")
            if submitted and c_name:
                try:
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO clients (name) VALUES (?)", (c_name,))
                    conn.commit()
                    st.success(f"Клиент '{c_name}' успешно зарегистрирован!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Клиент с таким названием уже существует!")

    st.subheader("Список контрагентов")
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
                "Всего паллет": pallets_count,
                "Занято ячеек": cells_count
            })
        
        sum_df = pd.DataFrame(summary_data)
        st.dataframe(sum_df, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📂 Открыть карточку клиента")
        selected_client_card = st.selectbox("Выберите клиента для перехода в карточку:", clients_df['name'].tolist())
        if st.button("Перейти в карточку клиента"):
            st.session_state.selected_client = selected_client_card
            st.session_state.page = "client_detail"
            st.rerun()
    else:
        st.info("Список клиентов пуст.")

elif st.session_state.page == "client_detail":
    c_name = st.session_state.selected_client
    if st.button("← Назад к списку клиентов"):
        st.session_state.page = "clients"
        st.session_state.selected_client = None
        st.rerun()
        
    st.header(f"🏢 Карточка клиента: {c_name}")
    
    cursor = conn.cursor()
    cursor.execute("SELECT tariff_A, tariff_B FROM clients WHERE name = ?", (c_name,))
    t_data = cursor.fetchone()
    if t_data:
        st.write(f"• Индивидуальный тариф Зона А: {t_data[0]} руб./день")
        st.write(f"• Индивидуальный тариф Зона B: {t_data[1]} руб./день")
        
    st.markdown("#### Активные паллеты и ячейки:")
    client_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'LPN Паллеты', GROUP_CONCAT(pl.address) as 'Ячейки'
        FROM pallets p
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn
        WHERE p.client = ?
        GROUP BY p.lpn
    """, conn, params=(c_name,))
    
    if not client_pallets_df.empty:
        st.dataframe(client_pallets_df, use_container_width=True)
        
        st.markdown("#### Детализированная номенклатура товаров:")
        client_items_df = pd.read_sql("""
            SELECT pi.lpn as 'LPN', pi.sku as 'SKU / Артикул', pi.item_name as 'Наименование', pi.qty as 'Количество'
            FROM pallet_items pi
            JOIN pallets p ON pi.lpn = p.lpn
            WHERE p.client = ?
        """, conn, params=(c_name,))
        if not client_items_df.empty:
            st.dataframe(client_items_df, use_container_width=True)
        else:
            st.info("Позиционная номенклатура не заполнена.")
    else:
        st.info("У данного клиента нет активных паллет на складе.")

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

# --- РАЗДЕЛ 3: ПРИЕМКА ПАРТИИ В СТИЛЕ 1С (ПАРТИОННЫЙ ВВОД С ТАБЛИЧНОЙ ЧАСТЬЮ) ---
elif st.session_state.page == "inbound":
    st.header("📥 Документ приемки партии товаров (Стиль 1С)")
    st.write("Сформируйте партию целиком, добавляя паллеты кнопкой, указывая для каждой адрес и попозиционную номенклатуру.")
    
    clients_list = pd.read_sql("SELECT name FROM clients", conn)["name"].tolist()
    
    if not clients_list:
        st.warning("Сначала добавьте хотя бы одного клиента в разделе «Клиенты и тарифы».")
    else:
        selected_client = st.selectbox("Поклажедатель", clients_list)
        batch_number = st.text_input("Номер приходной накладной / партии", value=f"Партия-{random.randint(100, 999)}")
        
        if "batch_pallets" not in st.session_state:
            st.session_state.batch_pallets = []
            
        st.markdown("---")
        st.subheader("Табличная часть: Паллеты в партии")
        
        if st.button("➕ Добавить еще паллет в партию"):
            st.session_state.batch_pallets.append({
                "lpn": f"LPN-{random.randint(1000, 9999)}",
                "cell": "",
                "items": [{"sku": "SKU-001", "name": "Товар", "qty": 10}]
            })
            st.rerun()
            
        cursor = conn.cursor()
        
        # Получаем список свободных ячеек
        cursor.execute("SELECT address FROM locations WHERE status = 'FREE'")
        free_cells = [row[0] for row in cursor.fetchall()]
        
        for idx, pal in enumerate(st.session_state.batch_pallets):
            with st.container():
                st.markdown(f"**Паллета #{idx+1}**")
                col_p1, col_p2 = st.columns([1, 2])
                with col_p1:
                    pal["lpn"] = st.text_input(f"Номер LPN #{idx+1}", value=pal["lpn"], key=f"lpn_{idx}")
                with col_p2:
                    pal["cell"] = st.selectbox(f"Ячейка размещения #{idx+1}", free_cells if free_cells else ["Нет свободных"], key=f"cell_{idx}")
                
                st.write("Состав позиций на этой паллете:")
                for item_idx, itm in enumerate(pal["items"]):
                    ic1, ic2, ic3, ic4 = st.columns([2, 3, 1, 1])
                    with ic1:
                        itm["sku"] = st.text_input("Артикул", value=itm["sku"], key=f"sku_{idx}_{item_idx}")
                    with ic2:
                        itm["name"] = st.text_input("Наименование", value=itm["name"], key=f"name_{idx}_{item_idx}")
                    with ic3:
                        itm["qty"] = st.number_input("Кол-во", value=itm["qty"], min_value=1, key=f"qty_{idx}_{item_idx}")
                    with ic4:
                        if st.button("🗑️ Удали", key=f"del_item_{idx}_{item_idx}"):
                            pal["items"].pop(item_idx)
                            st.rerun()
                            
                if st.button("➕ Добавить позицию на паллету", key=f"add_item_{idx}"):
                    pal["items"].append({"sku": "SKU-002", "name": "Еще товар", "qty": 5})
                    st.rerun()
                    
                if st.button(f"❌ Удалить паллету #{idx+1} из партии", key=f"del_pal_{idx}"):
                    st.session_state.batch_pallets.pop(idx)
                    st.rerun()
                st.markdown("---")
                
        if st.button("💾 Провести приход всей партии"):
            if not st.session_state.batch_pallets:
                st.error("Добавьте хотя бы одну паллету в партию!")
            else:
                try:
                    for pal in st.session_state.batch_pallets:
                        lpn = pal["lpn"]
                        cell = pal["cell"]
                        
                        cursor.execute("INSERT INTO pallets (lpn, client) VALUES (?, ?)", (lpn, selected_client))
                        cursor.execute("INSERT INTO pallet_locations (lpn, address) VALUES (?, ?)", (lpn, cell))
                        cursor.execute("UPDATE locations SET status = 'OCCUPIED' WHERE address = ?", (cell,))
                        
                        for itm in pal["items"]:
                            cursor.execute("INSERT INTO pallet_items (lpn, sku, item_name, qty) VALUES (?, ?, ?, ?)", 
                                           (lpn, itm["sku"], itm["name"], itm["qty"]))
                            
                    conn.commit()
                    st.success(f"Партия '{batch_number}' успешно проведена! Оприходовано паллет: {len(st.session_state.batch_pallets)}.")
                    st.session_state.batch_pallets = []
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка при проведении: {e}")

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
