import streamlit as st
import pandas as pd
import sqlite3
import io
import random
import base64
from datetime import date

st.set_page_config(page_title="3PL Nord Wheel", layout="wide")

# --- ФУНКЦИЯ ДЛЯ КОДИРОВАНИЯ ЛОГОТИПА В BASE64 ---
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except:
        return ""

# --- КАСТОМНЫЙ CSS ДЛЯ СТИЛИСТИКИ И КНОПОК ---
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

# --- ИНИЦИАЛИЗАЦИЯ И МИГРАЦИЯ БАЗЫ ДАННЫХ SQLite ---
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
    
    # Безохпасное добавление новых колонок для существующих баз данных
    for col, col_type in [("batch_name", "TEXT"), ("arrival_date", "TEXT"), ("status", "TEXT")]:
        try:
            cursor.execute(f"ALTER TABLE pallets ADD COLUMN {col} {col_type}")
        except sqlite3.OperationalError:
            pass # Колонка уже существует
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallet_locations (
        lpn TEXT,
        address TEXT,
        PRIMARY KEY(lpn, address)
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallet_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        lpn TEXT,
        sku TEXT,
        item_name TEXT,
        qty INTEGER
    )''')
    
    conn.commit()
    
    cursor.execute("SELECT COUNT(*) FROM locations")
    if cursor.fetchone()[0] == 0:
        locs = []
        zones_a_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M']
        for aisle in zones_a_letters:
            for sec in range(1, 7):
                for tier in range(1, 6):
                    for pos in range(1, 4):
                        addr = f"{aisle}-{sec}-{tier}-{pos}"
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

# --- ШАПКА ПРИЛОЖЕНИЯ С МИНИ-ЛОГОТИПОМ ---
logo_base64 = get_base64_image("logo.png")
logo_html = f"<img src='data:image/png;base64,{logo_base64}' style='width: 45px; height: 45px; border-radius: 50%; object-fit: cover; vertical-align: middle; margin-right: 15px; border: 2px solid white;'/>" if logo_base64 else "🛞 "

st.markdown(f"""
    <div style='background: linear-gradient(135deg, #1d4ed8 0%, #0284c7 50%, #dc2626 100%); padding: 22px 25px; border-radius: 16px; display: flex; align-items: center; justify-content: center; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); margin-bottom: 25px;'>
        {logo_html}
        <div style='text-align: left;'>
            <h1 style='color: white; margin: 0; font-size: 30px; font-weight: 800; text-shadow: 0 2px 4px rgba(0,0,0,0.2);'>3PL Nord Wheel</h1>
            <p style='margin: 4px 0 0 0; font-size: 14px; color: #f1f5f9; font-weight: 500;'>Профессиональная система ответственного хранения грузов и биллинга</p>
        </div>
    </div>
""", unsafe_allow_html=True)

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
        st.session_state.selected_client = None
    if st.button("🗺️ Карта склада и ячейки", use_container_width=True):
        st.session_state.page = "map"
    if st.button("📥 Приход", use_container_width=True):
        st.session_state.page = "inbound"
    if st.button("🔄 Списание и редакция", use_container_width=True):
        st.session_state.page = "management"
    if st.button("🖨️ Печать листов А4 и Отчеты", use_container_width=True):
        st.session_state.page = "reports"
    if st.button("📊 Биллинг (Снапшот)", use_container_width=True):
        st.session_state.page = "billing"

conn = get_connection()

# --- РАЗДЕЛ 1: КЛИЕНТЫ И ТАРИФЫ ---
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
        if st.button("Открыть карточку клиента"):
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
        st.write(f"• Тариф Зона А: {t_data[0]} руб./день")
        st.write(f"• Тариф Зона B: {t_data[1]} руб./день")
        
    st.markdown("#### Активные паллеты и ячейки:")
    client_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'LPN Паллеты', p.batch_name as 'Партия', p.arrival_date as 'Дата', p.status as 'Статус', GROUP_CONCAT(pl.address) as 'Ячейки'
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
    search_query = st.text_input("Поиск по конкретному адресу ячейки (например, A-1-1-1 или B-042)")
    if search_query:
        loc_df = loc_df[loc_df["address"].str.contains(search_query, case=False)]

    st.metric("Найдено ячеек", len(loc_df))
    st.dataframe(loc_df.head(150), use_container_width=True)

# --- РАЗДЕЛ 3: ПРИХОД ---
elif st.session_state.page == "inbound":
    st.header("📥 Документ прихода партии товаров")
    st.write("Сформируйте партию целиком. Если ячейка занята, появится предупреждение.")
    
    clients_list = pd.read_sql("SELECT name FROM clients", conn)["name"].tolist()
    
    if not clients_list:
        st.warning("Сначала добавьте хотя бы одного клиента в разделе «Клиенты и тарифы».")
    else:
        with st.container():
            st.markdown("### Шапка приходного документа")
            col_h1, col_h2 = st.columns(2)
            with col_h1:
                inb_client = st.selectbox("Клиент (Поклажедатель)", clients_list)
                inb_batch = st.text_input("Название всей партии / Накладной", value=f"Партия-{random.randint(100, 999)}")
            with col_h2:
                inb_date = st.date_input("Дата прихода", value=date.today())
                inb_status = st.selectbox("Статус документа", ["Активный", "Отложено (Черновик)"])
                inb_count = st.number_input("Общее количество паллет", min_value=1, max_value=100, value=2)

        if "wizard_step" not in st.session_state:
            st.session_state.wizard_step = 1

        if st.button("📄 Сгенерировать страницы паллет для заполнения"):
            st.session_state.wizard_pallets = []
            for i in range(inb_count):
                st.session_state.wizard_pallets.append({
                    "lpn": f"LPN-{random.randint(1000, 9999)}",
                    "cell": "",
                    "items": [{"sku": f"SKU-{i+1:03d}", "name": f"Товар партии {i+1}", "qty": 10}]
                })
            st.session_state.wizard_step = 2
            st.rerun()

        if st.session_state.get("wizard_step", 1) == 2 and "wizard_pallets" in st.session_state:
            st.markdown("---")
            st.subheader(f"Заполнение паллет для партии: {inb_batch} ({len(st.session_state.wizard_pallets)} шт.)")
            
            cursor = conn.cursor()
            cursor.execute("SELECT address, status FROM locations")
            all_cells_status = {row[0]: row[1] for row in cursor.fetchall()}
            all_cells_list = list(all_cells_status.keys())

            for idx, pal in enumerate(st.session_state.wizard_pallets):
                with st.expander(f"📦 Паллета #{idx+1} (LPN: {pal['lpn']})", expanded=(idx==0)):
                    col_p1, col_p2 = st.columns([1, 2])
                    with col_p1:
                        pal["lpn"] = st.text_input(f"Номер LPN #{idx+1}", value=pal["lpn"], key=f"wiz_lpn_{idx}")
                    with col_p2:
                        cell_options = []
                        for c in all_cells_list:
                            st_val = all_cells_status.get(c, 'FREE')
                            if st_val == 'OCCUPIED':
                                cell_options.append(f"🔴 [ЗАНЯТА] {c}")
                            else:
                                cell_options.append(f"🟢 [СВОБОДНА] {c}")
                        
                        selected_cell_display = st.selectbox(f"Ячейка размещения #{idx+1} (Формат А-1-1-1)", cell_options, key=f"wiz_cell_{idx}")
                        pal["cell"] = selected_cell_display.split("] ")[1] if "] " in selected_cell_display else selected_cell_display
                    
                    st.write("Позиции на паллете:")
                    for item_idx, itm in enumerate(pal["items"]):
                        ic1, ic2, ic3, ic4 = st.columns([2, 3, 1, 1])
                        with ic1:
                            itm["sku"] = st.text_input("Артикул", value=itm["sku"], key=f"wiz_sku_{idx}_{item_idx}")
                        with ic2:
                            itm["name"] = st.text_input("Наименование", value=itm["name"], key=f"wiz_name_{idx}_{item_idx}")
                        with ic3:
                            itm["qty"] = st.number_input("Кол-во", value=itm["qty"], min_value=1, key=f"wiz_qty_{idx}_{item_idx}")
                        with ic4:
                            if st.button("🗑️ Удали", key=f"wiz_del_item_{idx}_{item_idx}"):
                                pal["items"].pop(item_idx)
                                st.rerun()
                                
                    if st.button("➕ Добавить позицию", key=f"wiz_add_item_{idx}"):
                        pal["items"].append({"sku": "SKU-999", "name": "Новый товар", "qty": 5})
                        st.rerun()

            st.markdown("---")
            if st.button("💾 Провести и сохранить приходную партию"):
                try:
                    occupied_warnings = []
                    for pal in st.session_state.wizard_pallets:
                        c_stat = all_cells_status.get(pal["cell"], 'FREE')
                        if c_stat == 'OCCUPIED' and inb_status == "Активный":
                            occupied_warnings.append(pal["cell"])
                    
                    if occupied_warnings:
                        st.warning(f"⚠️ Предупреждение: Вы выбрали уже занятые ячейки: {', '.join(set(occupied_warnings))}. Размещение разрешено.")
                    
                    for pal in st.session_state.wizard_pallets:
                        lpn = pal["lpn"]
                        cell = pal["cell"]
                        
                        cursor.execute("INSERT OR REPLACE INTO pallets (lpn, client, batch_name, arrival_date, status) VALUES (?, ?, ?, ?, ?)", 
                                       (lpn, inb_client, inb_batch, str(inb_date), inb_status))
                        
                        if inb_status == "Активный" and cell:
                            cursor.execute("INSERT OR REPLACE INTO pallet_locations (lpn, address) VALUES (?, ?)", (lpn, cell))
                            cursor.execute("UPDATE locations SET status = 'OCCUPIED' WHERE address = ?", (cell,))
                        
                        for itm in pal["items"]:
                            cursor.execute("INSERT INTO pallet_items (lpn, sku, item_name, qty) VALUES (?, ?, ?, ?)", 
                                           (lpn, itm["sku"], itm["name"], itm["qty"]))
                            
                    conn.commit()
                    st.success(f"Приход партии '{inb_batch}' успешно сохранен! Статус: {inb_status}.")
                    st.session_state.wizard_step = 1
                    st.session_state.wizard_pallets = []
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка при сохранении: {e}")

    st.subheader("Текущие паллеты на складе")
    pallets_df = pd.read_sql("""
        SELECT p.lpn as 'LPN', p.client as 'Клиент', p.batch_name as 'Партия', p.arrival_date as 'Дата', p.status as 'Статус', GROUP_CONCAT(pl.address) as 'Ячейки' 
        FROM pallets p 
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn 
        GROUP BY p.lpn
    """, conn)
    if not pallets_df.empty:
        st.dataframe(pallets_df, use_container_width=True)
    else:
        st.info("Палет на складе пока нет.")

# --- РАЗДЕЛ 4: СПИСАНИЕ И РЕДАКЦИЯ ПАЛЛЕТ ---
elif st.session_state.page == "management":
    st.header("🔄 Списание товара и редакция ячеек")
    st.write("Управляйте хранящимися паллетами: редактируйте состав или списывайте товар.")
    
    pallets_list = pd.read_sql("SELECT lpn FROM pallets", conn)["lpn"].tolist()
    
    if not pallets_list:
        st.info("На складе нет активных паллет для управления.")
    else:
        selected_lpn = st.selectbox("Выберите паллету (LPN) для управления:", pallets_list)
        
        cursor = conn.cursor()
        cursor.execute("SELECT client, batch_name FROM pallets WHERE lpn = ?", (selected_lpn,))
        p_info = cursor.fetchone()
        cursor.execute("SELECT address FROM pallet_locations WHERE lpn = ?", (selected_lpn,))
        p_cells = [row[0] for row in cursor.fetchall()]
        
        st.markdown(f"**Клиент:** {p_info[0] if p_info else 'Не найден'}")
        st.markdown(f"**Партия:** {p_info[1] if p_info else 'Не указана'}")
        st.markdown(f"**Занимаемые ячейки:** {', '.join(p_cells)}")
        
        items_df = pd.read_sql("SELECT id, sku, item_name, qty FROM pallet_items WHERE lpn = ?", conn, params=(selected_lpn,))
        st.markdown("#### Позиции на паллете:")
        st.dataframe(items_df, use_container_width=True)
        
        col_m1, col_m2 = st.columns(2)
        
        with col_m1:
            st.markdown("### 🗑️ Списание паллеты")
            if st.button("🔴 Списать паллету полностью"):
                try:
                    for cell in p_cells:
                        cursor.execute("SELECT COUNT(*) FROM pallet_locations WHERE address = ? AND lpn != ?", (cell, selected_lpn))
                        other_pallets_count = cursor.fetchone()[0]
                        if other_pallets_count == 0:
                            cursor.execute("UPDATE locations SET status = 'FREE' WHERE address = ?", (cell,))
                    
                    cursor.execute("DELETE FROM pallet_locations WHERE lpn = ?", (selected_lpn,))
                    cursor.execute("DELETE FROM pallet_items WHERE lpn = ?", (selected_lpn,))
                    cursor.execute("DELETE FROM pallets WHERE lpn = ?", (selected_lpn,))
                    
                    conn.commit()
                    st.success(f"Паллета {selected_lpn} успешно списана!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка при списании: {e}")
                    
        with col_m2:
            st.markdown("### ✏️ Редактирование состава паллеты")
            with st.form("edit_item_form"):
                edit_sku = st.text_input("Новый Артикул / SKU")
                edit_name = st.text_input("Новое Наименование")
                edit_qty = st.number_input("Количество", min_value=1, value=1)
                submit_edit = st.form_submit_button("Добавить позицию на паллету")
                if submit_edit and edit_sku:
                    cursor.execute("INSERT INTO pallet_items (lpn, sku, item_name, qty) VALUES (?, ?, ?, ?)", 
                                   (selected_lpn, edit_sku, edit_name, edit_qty))
                    conn.commit()
                    st.success("Позиция успешно добавлена на паллету!")
                    st.rerun()

# --- РАЗДЕЛ 5: ПЕЧАТЬ ЛИСТОВ А4 И ОТЧЕТЫ ---
elif st.session_state.page == "reports":
    st.header("🖨️ Паллетные листы А4 с попозиционной номенклатурой")
    
    pallets_list = pd.read_sql("SELECT lpn FROM pallets", conn)["lpn"].tolist()
    if not pallets_list:
        st.info("Нет созданных паллет.")
    else:
        p_lpn = st.selectbox("Выберите паллету для печати A4 листа", pallets_list)
        
        cursor = conn.cursor()
        cursor.execute("SELECT client, batch_name, arrival_date FROM pallets WHERE lpn = ?", (p_lpn,))
        p_data = cursor.fetchone()
        cursor.execute("SELECT address FROM pallet_locations WHERE lpn = ?", (p_lpn,))
        locs_data = [row[0] for row in cursor.fetchall()]
        
        items_data = pd.read_sql("SELECT sku as 'Артикул', item_name as 'Наименование', qty as 'Количество' FROM pallet_items WHERE lpn = ?", conn, params=(p_lpn,))
        
        if p_data:
            st.markdown("---")
            st.markdown(f"<h2 style='text-align: center; color: #0284c7;'>📄 ПАЛЛЕТНЫЙ ЛИСТ (A4) — 3PL Nord Wheel</h2>", unsafe_allow_html=True)
            st.markdown(f"### **Клиент:** {p_data[0]} | **Партия:** {p_data[1]} | **Дата:** {p_data[2]}")
            st.markdown(f"<h1>ID ПАЛЛЕТЫ: {p_lpn}</h1>", unsafe_allow_html=True)
            st.markdown(f"**Ячейки размещения:** {', '.join(locs_data)}")
            st.markdown("#### Позиционный состав груза (для сборки):")
            st.dataframe(items_data, use_container_width=True)
            st.markdown("---")
            st.caption("Штрихкод для сканирования ТСД: [ |||||||||||||||||||||||||| ]")
            st.button("🖨️ Печать листа (PDF/A4)")
            
    st.subheader("Экспорт всех данных в Excel")
    full_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'Номер паллеты', p.client as 'Клиент', p.batch_name as 'Партия', p.arrival_date as 'Дата прихода', GROUP_CONCAT(pl.address) as 'Ячейки' 
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

# --- РАЗДЕЛ 6: БИЛЛИНГ И СНАПШОТ ---
elif st.session_state.page == "billing":
    st.header("📊 Автоматический расчет хранения (Снапшот остатков)")
    st.write("Расчет строится на основе фактически занятых ячеек с учетом тарифов.")
    
    if st.button("Сделать срез (Snapshot) и рассчитать счета"):
        billing_query = """
            SELECT p.client, l.zone, COUNT(DISTINCT pl.address) as occupied_slots
            FROM pallets p
            JOIN pallet_locations pl ON p.lpn = pl.lpn
            JOIN locations l ON pl.address = l.address
            WHERE p.status = 'Активный'
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
            st.info("Нет данных для расчета (склад пуст или нет активных паллет).")

conn.close()
