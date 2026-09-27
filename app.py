import streamlit as st
import pandas as pd
import sqlite3
import io
import random

st.set_page_config(page_title="3PL Warehouse Management System", layout="wide")

# --- ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ SQLite ---
def get_connection():
    conn = sqlite3.connect("warehouse.db", check_same_thread=False)
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Таблица клиентов
    cursor.execute('''CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        tariff_A REAL,
        tariff_B REAL
    )''')
    
    # Таблица ячеек склада
    cursor.execute('''CREATE TABLE IF NOT EXISTS locations (
        address TEXT PRIMARY KEY,
        zone TEXT,
        status TEXT DEFAULT 'FREE'
    )''')
    
    # Таблица паллет
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallets (
        lpn TEXT PRIMARY KEY,
        client TEXT,
        nomenclature TEXT
    )''')
    
    # Таблица связи паллет и ячеек (поддержка мульти-ячеек / негабарита)
    cursor.execute('''CREATE TABLE IF NOT EXISTS pallet_locations (
        lpn TEXT,
        address TEXT,
        FOREIGN KEY(lpn) REFERENCES pallets(lpn),
        FOREIGN KEY(address) REFERENCES locations(address),
        PRIMARY KEY(lpn, address)
    )''')
    
    conn.commit()
    
    # Автозаполнение ячеек, если таблица пуста
    cursor.execute("SELECT COUNT(*) FROM locations")
    if cursor.fetchone()[0] == 0:
        locs = []
        zones_a_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M']
        for aisle in zones_a_letters:
            for sec in range(1, 7):
                for tier in range(1, 6):
                    for pos in range(1, 4):
                        addr = f"A-{aisle}-{sec:02d}-{tier}-{pos}"
                        locs.append((addr, "A", "FREE"))
        for i in range(1, 151):
            addr = f"B-{i:03d}"
            locs.append((addr, "B", "FREE"))
        cursor.executemany("INSERT OR IGNORE INTO locations (address, zone, status) VALUES (?, ?, ?)", locs)
        conn.commit()
    conn.close()

init_db()

# Функция для конвертации DataFrame в Excel для скачивания
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    processed_data = output.getvalue()
    return processed_data

# --- ИНТЕРФЕЙС ПРИЛОЖЕНИЯ ---
st.title("📦 3PL Склад: Управление хранением и биллинг")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1. Клиенты и тарифы", 
    "2. Карта склада и ячейки", 
    "3. Приемка и размещение", 
    "4. Печать листов А4 и Отчеты",
    "5. Биллинг (Снапшот)"
])

conn = get_connection()

# --- ТАБ 1: КЛИЕНТЫ И ТАРИФЫ ---
with tab1:
    st.header("Регистрация поклажедателей")
    with st.form("add_client_form"):
        c_name = st.text_input("Название компании")
        t_a = st.number_input("Тариф Зона А (руб/паллето-день)", value=30.0)
        t_b = st.number_input("Тариф Зона B (руб/паллето-день)", value=20.0)
        submitted = st.form_submit_button("Добавить клиента")
        if submitted and c_name:
            try:
                cursor = conn.cursor()
                cursor.execute("INSERT INTO clients (name, tariff_A, tariff_B) VALUES (?, ?, ?)", (c_name, t_a, t_b))
                conn.commit()
                st.success(f"Клиент '{c_name}' успешно зарегистрирован!")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("Клиент с таким названием уже существует!")

    st.subheader("Список активных клиентов")
    clients_df = pd.read_sql("SELECT * FROM clients", conn)
    st.dataframe(clients_df, use_container_width=True)

# --- ТАБ 2: КАРТА СКЛАДА ---
with tab2:
    st.header("Интерактивная карта и статус ячеек")
    
    col1, col2 = st.columns(2)
    with col1:
        zone_filter = st.selectbox("Фильтр по зоне", ["Все", "Зона А (Стеллажи)", "Зона B (2-й этаж)"])
    with col2:
        status_filter = st.selectbox("Статус ячейки", ["Все", "Свободные (FREE)", "Занятые (OCCUPIED)"])
    
    query = "SELECT * FROM locations WHERE 1=1"
    if zone_filter == "Зона А (Стеллажи)":
        query += " AND zone = 'A'"
    elif zone_filter == "Зона B (2-й этаж)":
        query += " AND zone = 'B'"
        
    if status_filter == "Свободные (FREE)":
        query += " AND status = 'FREE'"
    elif status_filter == "Занятые (OCCUPIED)":
        query += " AND status = 'OCCUPIED'"
        
    loc_df = pd.read_sql(query, conn)
    search_query = st.text_input("Поиск по адресу ячейки (например, A-C-04 или B-042)")
    if search_query:
        loc_df = loc_df[loc_df["address"].str.contains(search_query, case=False)]

    st.metric("Найдено ячеек по фильтру", len(loc_df))
    st.dataframe(loc_df.head(100), use_container_width=True)

# --- ТАБ 3: ПРИЕМКА И РАЗМЕЩЕНИЕ ---
with tab3:
    st.header("Регистрация прихода и размещение паллет")
    
    clients_list = pd.read_sql("SELECT name FROM clients", conn)["name"].tolist()
    
    if not clients_list:
        st.warning("Сначала добавьте хотя бы одного клиента в Табе 1.")
    else:
        with st.form("inbound_form"):
            selected_client = st.selectbox("Поклажедатель", clients_list)
            pallet_num = st.text_input("Номер/Идентификатор паллеты (LPN)", value=f"LPN-{random.randint(1000, 9999)}")
            nomenclature = st.text_area("Состав номенклатуры", "Артикул, наименование, количество коробок")
            
            # Доступные свободные ячейки
            free_locs = pd.read_sql("SELECT address FROM locations WHERE status = 'FREE'", conn)["address"].tolist()
            assigned_locs = st.multiselect("Выберите ячейки для размещения (можно выбрать несколько для негабарита)", free_locs)
            
            submit_inbound = st.form_submit_button("Оприходовать и разместить")
            
            if submit_inbound:
                if not assigned_locs:
                    st.error("Выберите хотя бы одну ячейку!")
                else:
                    try:
                        cursor = conn.cursor()
                        # Сохраняем паллету
                        cursor.execute("INSERT INTO pallets (lpn, client, nomenclature) VALUES (?, ?, ?)", 
                                       (pallet_num, selected_client, nomenclature))
                        
                        # Привязываем ячейки
                        for loc in assigned_locs:
                            cursor.execute("INSERT INTO pallet_locations (lpn, address) VALUES (?, ?)", (pallet_num, loc))
                            cursor.execute("UPDATE locations SET status = 'OCCUPIED' WHERE address = ?", (loc,))
                            
                        conn.commit()
                        st.success(f"Паллета {pallet_num} оприходована и закреплена за ячейками: {', '.join(assigned_locs)}!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Паллета с таким LPN уже существует на складе!")

    st.subheader("Текущие паллеты на складе")
    pallets_df = pd.read_sql("""
        SELECT p.lpn, p.client, p.nomenclature, GROUP_CONCAT(pl.address) as locations 
        FROM pallets p 
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn 
        GROUP BY p.lpn
    """, conn)
    
    if not pallets_df.empty:
        search_pallets = st.text_input("Поиск по паллетам (LPN, клиент или состав)")
        if search_pallets:
            pallets_df = pallets_df[
                pallets_df['lpn'].str.contains(search_pallets, case=False) |
                pallets_df['client'].str.contains(search_pallets, case=False) |
                pallets_df['nomenclature'].str.contains(search_pallets, case=False)
            ]
        st.dataframe(pallets_df, use_container_width=True)
    else:
        st.info("Палет на складе пока нет.")

# --- ТАБ 4: ПЕЧАТЬ ЛИСТОВ А4 И ОТЧЕТЫ ---
with tab4:
    st.header("Генерация паллетных листов А4 и выгрузка данных")
    
    pallets_list = pd.read_sql("SELECT lpn FROM pallets", conn)["lpn"].tolist()
    if not pallets_list:
        st.info("Нет созданных паллет.")
    else:
        p_lpn = st.selectbox("Выберите паллету для печати A4 листа", pallets_list)
        
        cursor = conn.cursor()
        cursor.execute("SELECT client, nomenclature FROM pallets WHERE lpn = ?", (p_lpn,))
        p_data = cursor.fetchone()
        cursor.execute("SELECT address FROM pallet_locations WHERE lpn = ?", (p_lpn,))
        locs_data = [row[0] for row in cursor.fetchall()]
        
        if p_data:
            st.markdown("---")
            st.markdown(f"<h2 style='text-align: center;'>📄 ПАЛЛЕТНЫЙ ЛИСТ (A4)</h2>", unsafe_allow_html=True)
            st.markdown(f"### **Клиент:** {p_data[0]}")
            st.markdown(f"<h1>ID ПАЛЛЕТЫ: {p_lpn}</h1>", unsafe_allow_html=True)
            st.markdown(f"**Ячейки размещения:** {', '.join(locs_data)}")
            st.markdown("#### Состав груза:")
            st.write(p_data[1])
            st.markdown("---")
            st.caption("Штрихкод для сканирования ТСД: [ |||||||||||||||||||||||||| ]")
            st.button("🖨️ Печать листа (PDF/A4)")
            
    st.subheader("Экспорт всех паллет в Excel")
    full_pallets_df = pd.read_sql("""
        SELECT p.lpn as 'Номер паллеты', p.client as 'Клиент', p.nomenclature as 'Состав', GROUP_CONCAT(pl.address) as 'Ячейки' 
        FROM pallets p 
        LEFT JOIN pallet_locations pl ON p.lpn = pl.lpn 
        GROUP BY p.lpn
    """, conn)
    
    if not full_pallets_df.empty:
        excel_data = convert_df_to_excel(full_pallets_df)
        st.download_button(
            label="📥 Скачать отчет по паллетам в Excel",
            data=excel_data,
            file_name="warehouse_pallets_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- ТАБ 5: БИЛЛИНГ И СНАПШОТ ---
with tab5:
    st.header("Автоматический расчет хранения (Снапшот остатков)")
    st.write("Расчет строится на основе фактически занятых ячеек (паллето-мест) с учетом негабарита и индивидуальных тарифов.")
    
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
            
            # Кнопка скачивания биллинга в Excel
            billing_excel = convert_df_to_excel(res_df)
            st.download_button(
                label="📥 Скачать счет/акт биллинга в Excel",
                data=billing_excel,
                file_name="warehouse_billing_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Нет данных для расчета (склад пуст или не заведены тарифы клиентов).")

conn.close()
