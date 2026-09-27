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
        tariff_B REAL DEFAULT 20.0,
        phone TEXT DEFAULT '',
        email TEXT DEFAULT ''
    )''')
    
    # Миграция колонок для существующих таблиц клиентов
    for col, col_type in [("phone", "TEXT"), ("email", "TEXT"), ("tariff_A", "REAL DEFAULT 30.0"), ("tariff_B", "REAL DEFAULT 20.0")]:
        try:
            cursor.execute(f"ALTER TABLE clients ADD COLUMN {col} {col_type}")
        except sqlite3.OperationalError:
            pass

    # Таблица для дополнительных тарифов на услуги
    cursor.execute('''CREATE TABLE IF NOT EXISTS client_tariffs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_name TEXT,
        service_name TEXT,
        price REAL,
        FOREIGN KEY(client_name) REFERENCES clients(name)
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS locations (
        address TEXT PRIMARY KEY,
        zone TEXT,
        status TEXT DEFAULT 'Свободна'
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallets (
        lpn TEXT PRIMARY KEY,
        client TEXT
    )''')
    
    for col, col_type in [("batch_name", "TEXT"), ("arrival_date", "TEXT"), ("status", "TEXT")]:
        try:
            cursor.execute(f"ALTER TABLE pallets ADD COLUMN {col} {col_type}")
        except sqlite3.OperationalError:
            pass
    
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
                        locs.append((addr, "A", "Свободна"))
        for i in range(1, 151):
            addr = f"B-{i:03d}"
            locs.append((addr, "B", "Свободна"))
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
            c_phone = st.text_input("Телефон")
            c_email = st.text_input("Электронная почта")
            submitted = st.form_submit_button("Зарегистрировать клиента")
            if submitted and c_name:
                try:
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO clients (name, phone, email) VALUES (?, ?, ?)", (c_name, c_phone, c_email))
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
                "Телефон": row['phone'],
                "Email": row['email'],
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
    cursor.execute("SELECT tariff_A, tariff_B, phone, email FROM clients WHERE name = ?", (c_name,))
    c_info = cursor.fetchone()
    
    # Блок редактирования данных клиента
    with st.expander("✏️ Редактировать данные клиента и тарифы хранения"):
        with st.form("edit_client_form"):
            new_name = st.text_input("Название компании", value=c_name)
            new_phone = st.text_input("Телефон", value=c_info[2] if c_info and c_info[2] else "")
            new_email = st.text_input("Email", value=c_info[3] if c_info and c_info[3] else "")
            new_t_a = st.number_input("Тариф Зона А (руб/день)", value=c_info[0] if c_info else 30.0)
            new_t_b = st.number_input("Тариф Зона B (руб/день)", value=c_info[1] if c_info else 20.0)
            
            submit_edit = st.form_submit_button("Сохранить изменения")
            if submit_edit:
                try:
                    cursor.execute("UPDATE clients SET name = ?, phone = ?, email = ?, tariff_A = ?, tariff_B = ? WHERE name = ?", 
                                   (new_name, new_phone, new_email, new_t_a, new_t_b, c_name))
                    if new_name != c_name:
                        cursor.execute("UPDATE pallets SET client = ? WHERE client = ?", (new_name, c_name))
                        cursor.execute("UPDATE client_tariffs SET client_name = ? WHERE client_name = ?", (new_name, c_name))
                        st.session_state.selected_client = new_name
                    conn.commit()
                    st.success("Данные клиента успешно обновлены!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка при обновлении: {e}")

    st.markdown(f"📞 **Телефон:** {c_info[2] if c_info and c_info[2] else 'Не указан'}")
    st.markdown(f"📧 **Email:** {c_info[3] if c_info and c_info[3] else 'Не указан'}")
    st.markdown(f"• **Тариф Зона А (хранение):** {c_info[0] if c_info else 30.0} руб./день")
    st.markdown(f"• **Тариф Зона B (хранение):** {c_info[1] if c_info else 20.0} руб./день")
    
    # Секция дополнительных тарифов на услуги
    st.markdown("---")
    st.markdown("#### 📋 Тарифы на дополнительные услуги:")
    services_df = pd.read_sql("SELECT id, service_name as 'Услуга', price as 'Стоимость (руб.)' FROM client_tariffs WHERE client_name = ?", conn, params=(st.session_state.selected_client,))
    if not services_df.empty:
        st.dataframe(services_df.drop(columns=['id']), use_container_width=True)
    else:
        st.info("Дополнительные тарифы на услуги не заведены.")
        
    with st.form("add_service_tariff_form"):
        st.markdown("**Добавить тариф на новый вид услуги:**")
        col_s1, col_s2 = st.columns([2, 1])
        with col_s1:
            s_name = st.text_input("Название услуги (например: Погрузка, Упаковка, Маркировка)")
        with col_s2:
            s_price = st.number_input("Стоимость (руб.)", min_value=0.0, value=100.0)
        sub_serv = st.form_submit_button("Добавить услугу")
        if sub_serv and s_name:
            cursor.execute("INSERT INTO client_tariffs (client_name, service_name, price) VALUES (?, ?, ?)", (st.session_state.selected_client, s_name, s_price))
            conn.commit()
            st.success(f"Услуга '{s_name}' успешно добавлена!")
            st.rerun()

    st.markdown("---")
    st.markdown("#### 📦 Активные паллеты и ячейки клиента:")
    client_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'Название паллета', p.batch_name as 'Партия', p.arrival_date as 'Дата', p.status as 'Статус', GROUP_CONCAT(pl.address) as 'Ячейки'
        FROM pallets p
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn
        WHERE p.client = ?
        GROUP BY p.lpn
    """, conn, params=(c_name,))
    
    if not client_pallets_df.empty:
        st.dataframe(client_pallets_df, use_container_width=True)
        
        st.markdown("#### Детализированная номенклатура товаров:")
        client_items_df = pd.read_sql("""
            SELECT pi.lpn as 'Название паллета', pi.sku as 'SKU / Артикул', pi.item_name as 'Наименование', pi.qty as 'Количество'
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

# --- РАЗДЕЛ 2: КАРТА СКЛАДА ---
elif st.session_state.page == "map":
    st.header("🗺️ Интерактивная карта и поиск ячеек")
    
    clients_list = ["Все клиенты"] + pd.read_sql("SELECT name FROM clients", conn)["name"].tolist()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        client_map_filter = st.selectbox("Фильтр по клиенту:", clients_list)
    with col2:
        zone_filter = st.selectbox("Фильтр по зоне", ["Все", "Зона А (Стеллажи)", "Зона B (2-й этаж)"])
    with col3:
        status_filter = st.selectbox("Статус ячейки", ["Все", "Свободна", "Занята"])
    
    query = """
        SELECT l.address as 'Адрес', l.zone as 'Зона', l.status as 'Статус', 
               COALESCE(p.lpn, '-') as 'Название паллета', 
               COALESCE(p.client, '-') as 'Клиент', 
               COALESCE(p.batch_name, '-') as 'Партия'
        FROM locations l
        LEFT JOIN pallet_locations pl ON l.address = pl.address
        LEFT JOIN pallets p ON pl.lpn = p.lpn
        WHERE 1=1
    """
    
    if client_map_filter != "Все клиенты":
        query += f" AND p.client = '{client_map_filter}'"
        
    if zone_filter == "Зона А (Стеллажи)":
        query += " AND l.zone = 'A'"
    elif zone_filter == "Зона B (2-й этаж)":
        query += " AND l.zone = 'B'"
        
    if status_filter == "Свободна":
        query += " AND l.status = 'Свободна'"
    elif status_filter == "Занята":
        query += " AND l.status = 'Занята'"
        
    loc_df = pd.read_sql(query, conn).drop_duplicates(subset=['Адрес'])
    
    search_query = st.text_input("Поиск по адресу ячейки (например, A-1-1-1 или B-042)")
    
    has_filter = (client_map_filter != "Все клиенты" or zone_filter != "Все" or status_filter != "Все" or bool(search_query))
    
    if not has_filter:
        st.info("💡 Введите поисковый запрос по адресу или выберите фильтры выше, чтобы отобразить ячейки склада.")
    else:
        if search_query:
            loc_df = loc_df[loc_df["Адрес"].str.contains(search_query, case=False)]
        st.metric("Найдено ячеек", len(loc_df))
        st.dataframe(loc_df, use_container_width=True)

# --- РАЗДЕЛ 3: ПРИХОД ---
elif st.session_state.page == "inbound":
    st.header("📥 Документ прихода партии товаров")
    st.write("Сформируйте партию целиком. Название паллетов формируется автоматически на основе названия партии.")
    
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
                    "pallet_name": f"{inb_batch} - Паллета {i+1}",
                    "cell": "",
                    "items": [{"sku": f"SKU-{i+1:03d}", "name": f"Товар {i+1}", "qty": 10}]
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
                with st.expander(f"📦 Паллета #{idx+1} ({pal['pallet_name']})", expanded=(idx==0)):
                    col_p1, col_p2 = st.columns([1, 2])
                    with col_p1:
                        pal["pallet_name"] = st.text_input(f"Название паллета #{idx+1}", value=pal["pallet_name"], key=f"wiz_lpn_{idx}")
                    with col_p2:
                        cell_options = []
                        for c in all_cells_list:
                            st_val = all_cells_status.get(c, 'Свободна')
                            if st_val == 'Занята':
                                cell_options.append(f"🔴 [Занята] {c}")
                            else:
                                cell_options.append(f"🟢 [Свободна] {c}")
                        
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
                        c_stat = all_cells_status.get(pal["cell"], 'Свободна')
                        if c_stat == 'Занята' and inb_status == "Активный":
                            occupied_warnings.append(pal["cell"])
                    
                    if occupied_warnings:
                        st.warning(f"⚠️ Предупреждение: Вы выбрали уже занятые ячейки: {', '.join(set(occupied_warnings))}. Размещение разрешено.")
                    
                    for pal in st.session_state.wizard_pallets:
                        lpn = pal["pallet_name"]
                        cell = pal["cell"]
                        
                        cursor.execute("INSERT OR REPLACE INTO pallets (lpn, client, batch_name, arrival_date, status) VALUES (?, ?, ?, ?, ?)", 
                                       (lpn, inb_client, inb_batch, str(inb_date), inb_status))
                        
                        if inb_status == "Активный" and cell:
                            cursor.execute("INSERT OR REPLACE INTO pallet_locations (lpn, address) VALUES (?, ?)", (lpn, cell))
                            cursor.execute("UPDATE locations SET status = 'Занята' WHERE address = ?", (cell,))
                        
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
        SELECT p.lpn as 'Название паллета', p.client as 'Клиент', p.batch_name as 'Партия', p.arrival_date as 'Дата', p.status as 'Статус', GROUP_CONCAT(pl.address) as 'Ячейки' 
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
        selected_lpn = st.selectbox("Выберите паллету для управления:", pallets_list)
        
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
                            cursor.execute("UPDATE locations SET status = 'Свободна' WHERE address = ?", (cell,))
                    
                    cursor.execute("DELETE FROM pallet_locations WHERE lpn = ?", (selected_lpn,))
                    cursor.execute("DELETE FROM pallet_items WHERE lpn = ?", (selected_lpn,))
                    cursor.execute("DELETE FROM pallets WHERE lpn = ?", (selected_lpn,))
                    
                    conn.commit()
                    st.success(f"Паллета '{selected_lpn}' успешно списана!")
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
            st.markdown(f"<h1>НАЗВАНИЕ ПАЛЛЕТА: {p_lpn}</h1>", unsafe_allow_html=True)
            st.markdown(f"**Ячейки размещения:** {', '.join(locs_data)}")
            st.markdown("#### Позиционный состав груза (для сборки):")
            st.dataframe(items_data, use_container_width=True)
            st.markdown("---")
            st.caption("Штрихкод для сканирования ТСД: [ |||||||||||||||||||||||||| ]")
            st.button("🖨️ Печать листа (PDF/A4)")
            
    st.subheader("Экспорт всех данных в Excel")
    full_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'Название паллета', p.client as 'Клиент', p.batch_name as 'Партия', p.arrival_date as 'Дата прихода', GROUP_CONCAT(pl.address) as 'Ячейки' 
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
