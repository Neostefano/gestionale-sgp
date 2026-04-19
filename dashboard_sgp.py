import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import msal
import datetime
from streamlit_option_menu import option_menu

# --- 1. CONFIGURAZIONE ---
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
PERCORSO_PERSONALE = "stefano_pini_sgp-consulting_it"
SITE_HOSTNAME = "sgpconsultingstp.sharepoint.com"
SITE_PATH = "/sites/Server"
LISTA_TIMESHEET = "Timesheet"
LISTA_PIANIFICAZIONE = "Pianificazione"
LISTA_COMMESSE = "Commesse"
LISTA_SICUREZZA = "Sicurezza"
LISTA_PROGETTAZIONE = "Progettazione"

PASSWORD_ADMIN = st.secrets["PASSWORD_ADMIN"]
PASSWORD_USER = st.secrets["PASSWORD_USER"]

st.set_page_config(page_title="SGP Consulting - Dashboard", page_icon="⚙️", layout="wide")

# --- STILE GRAFICO SGP ---
st.markdown("""
    <style>
    h1, h2, h3 { color: #00AEEF !important; }
    .stButton>button { background-color: #00AEEF; color: white; border: none; border-radius: 5px; font-weight: bold; }
    .stButton>button:hover { background-color: #007bb5; color: white; }
    [data-testid="stSidebar"] { background-color: #f8f9fa; }
    </style>
""", unsafe_allow_html=True)

def get_ms_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    return app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"]).get("access_token")
import streamlit as st

# --- 1. RECUPERO PASSWORD DAI SECRETS ---
P_ADMIN = st.secrets["PASSWORD_ADMIN"]
P_USER = st.secrets["PASSWORD_USER"]

# --- 2. FUNZIONE DI LOGIN ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.role = None

if not st.session_state.authenticated:
    st.title("Accesso Gestionale SGP")
    password = st.text_input("Inserisci la password di accesso", type="password")
    
    if st.button("Entra"):
        if password == P_ADMIN:
            st.session_state.authenticated = True
            st.session_state.role = "admin"
            st.rerun()
        elif password == P_USER:
            st.session_state.authenticated = True
            st.session_state.role = "user"
            st.rerun()
        else:
            st.error("Password errata")
    st.stop() # Blocca l'esecuzione qui finché non sono loggati

# --- 3. CONFIGURAZIONE MENU IN BASE AL RUOLO ---
# Se sei Admin vedi tutto, se sei User togliamo le voci sensibili
if st.session_state.role == "admin":
    menu_options = ["Home", "Analisi Ore", "Pianificazione", "Sicurezza", "Contabilità", "Progettazione"]
    menu_icons = ["house", "clock", "calendar", "shield-lock", "currency-euro", "pencil"]
else:
    # Cristiano e Giuditta non vedono "Contabilità" e "Analisi Ore" (se contiene dati sensibili)
    menu_options = ["Home", "Pianificazione", "Sicurezza", "Progettazione"]
    menu_icons = ["house", "calendar", "shield-lock", "pencil"]

# Qui inserisci il tuo componente menu (es. option_menu)
with st.sidebar:
    st.image("IMG_0002.jpeg", width=200)
    selected = option_menu(
        "Menu SGP", 
        menu_options, 
        icons=menu_icons, 
        menu_icon="cast", 
        default_index=0
    )

# --- 4. PROTEZIONE DELLE PAGINE ---
if selected == "Contabilità":
    if st.session_state.role == "admin":
        st.title("Sezione Contabile (Riservata)")
        # Inserisci qui il tuo codice per i costi
    else:
        st.error("Non hai i permessi per visualizzare questa sezione.")

# ... continua con le altre sezioni ...
# ==========================================
# 🪄 IL MOTORE GENERATORE DI CARTELLE
# ==========================================
def esegui_creazione_cartelle(headers, step_cartelle):
    try:
        url_sito = f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_PATH}"
        sid = requests.get(url_sito, headers=headers).json().get('id')
        did = requests.get(f"https://graph.microsoft.com/v1.0/sites/{sid}/drive", headers=headers).json().get('id')
        
        def get_or_create(parent_path, name):
            get_url = f"https://graph.microsoft.com/v1.0/drives/{did}/{parent_path}/{name}"
            if requests.get(get_url, headers=headers).status_code == 200: 
                return True
            post_url = f"https://graph.microsoft.com/v1.0/drives/{did}/{parent_path}:/children"
            payload = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
            res = requests.post(post_url, headers=headers, json=payload)
            return res.status_code in [200, 201]

        for step in step_cartelle:
            if not get_or_create(step['parent'], step['name']): return False
        return True
    except Exception as e:
        st.error(f"❌ Errore API Cartelle: {e}")
        return False

def crea_cartelle_madre(codice, descrizione):
    headers = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
    nome_completo = f"{str(codice).strip()} {str(descrizione).strip()}" if str(descrizione).strip() else str(codice).strip()
    steps = [{"parent": "root:/Clienti", "name": nome_completo}]
    for sub in ["specifiche cliente", "offerta", "Sal", "documenti uscita", "documenti ingresso"]:
        steps.append({"parent": f"root:/Clienti/{nome_completo}", "name": sub})
    return esegui_creazione_cartelle(headers, steps)

def crea_cartelle_figlio(cartella_madre, cartella_figlio):
    headers = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
    steps = [
        {"parent": "root:/Clienti", "name": cartella_madre},
        {"parent": f"root:/Clienti/{cartella_madre}", "name": "Cantieri"},
        {"parent": f"root:/Clienti/{cartella_madre}/Cantieri", "name": cartella_figlio}
    ]
    for sub in ["PSC", "Incarico", "Verbali", "Specifiche"]:
        steps.append({"parent": f"root:/Clienti/{cartella_madre}/Cantieri/{cartella_figlio}", "name": sub})
    return esegui_creazione_cartelle(headers, steps)

# --- FUNZIONI DI LETTURA DATI (FETCH) ---

def fetch_timesheet():
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}
    url_site = f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}"
    site_id = requests.get(url_site, headers=headers).json().get("id")
    data = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LISTA_TIMESHEET}/items?expand=fields", headers=headers).json()
    
    rows = []
    if "value" in data:
        for item in data["value"]:
            if "fields" in item:
                f = item["fields"]
                rows.append({
                    "ID": item.get("id"),
                    "Collaboratore": f.get("Collaboratore", "N.D."), 
                    "Commessa": f.get("Commessa", "N.D."), 
                    "Ore": f.get("OreLavorate", 0), 
                    "Data": f.get("Created", "")[:10] if f.get("Created") else ""
                })
                
    df = pd.DataFrame(rows)
    if not df.empty: 
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    return df

def fetch_pianificazione():
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}
    url_site = f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}"
    site_id = requests.get(url_site, headers=headers).json().get("id")
    data = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LISTA_PIANIFICAZIONE}/items?expand=fields", headers=headers).json()
    
    rows = []
    if "value" in data:
        for item in data["value"]:
            if "fields" in item:
                f = item["fields"]
                rows.append({
                    "ID": item.get("id"), 
                    "Giorno": f.get("Data", "")[:10] if f.get("Data") else "", 
                    "Collaboratore": f.get("Title", "N.D."), 
                    "Cantiere": f.get("Commessa", "N.D."), 
                    "Istruzioni": f.get("Note", "")
                })
    df = pd.DataFrame(rows)
    if not df.empty: 
        df['Giorno'] = pd.to_datetime(df['Giorno'], errors='coerce')
    return df

def fetch_commesse():
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}
    url_site = f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}"
    site_id = requests.get(url_site, headers=headers).json().get("id")
    data = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LISTA_COMMESSE}/items?expand=fields", headers=headers).json()
    
    rows = []
    if "value" in data:
        for item in data["value"]:
            if "fields" in item:
                f = item["fields"]
                rows.append({
                    "ID": item.get("id"),
                    "Codice": f.get("Title", "N.D."),
                    "Descrizione": f.get("Descrizione", ""),
                    "Stato": f.get("Stato", "Attiva"),
                    "Fatturazione": f.get("Stato_Fatturazione", "Da Fatturare"),
                    "Avanzamento %": f.get("Avanzamento", 0),
                    "SAL Emessi": f.get("SalEmessi", 0),
                    "Importo Totale": f.get("Importo_x0020_Totale", 0.0),
                    "Scadenza": f.get("Scadenza", "")[:10] if f.get("Scadenza") else ""
                })
    df = pd.DataFrame(rows)
    if not df.empty:
        df['Scadenza_DT'] = pd.to_datetime(df['Scadenza'], errors='coerce')
        df['Anno Apertura'] = df['Codice'].apply(lambda x: int("20" + str(x)[:2]) if pd.notnull(x) and len(str(x)) >= 2 and str(x)[:2].isdigit() else None)
    return df

def fetch_sicurezza():
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}
    url_site = f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}"
    site_id = requests.get(url_site, headers=headers).json().get("id")
    cols = ["ID", "Commessa_Madre", "Sotto_Commessa", "Descrizione", "Stato", "Stato_PSC", "Riunione_Fatta", "Data_Apertura", "KM"]
    try:
        data = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LISTA_SICUREZZA}/items?expand=fields", headers=headers).json()
        rows = []
        if "value" in data:
            for item in data["value"]:
                if "fields" in item:
                    f = item["fields"]
                    rows.append({
                        "ID": item["id"],
                        "Commessa_Madre": f.get("Commessa", "N.D."),
                        "Sotto_Commessa": f.get("Sotto_x002d_Commessa_x0020__x002", "N.D."),
                        "Descrizione": f.get("DescrizioneCantiere", ""),
                        "Stato": f.get("Stato", "Attiva"), 
                        "Stato_PSC": f.get("StatoPSC", "Da Redigere"),
                        "Riunione_Fatta": f.get("RiunionediCoordinamento", False),
                        "Data_Apertura": f.get("DataScadenza", "")[:10] if f.get("DataScadenza") else "",
                        "KM": float(f.get("Lunghezza_KM", 0.0))
                    })
        df = pd.DataFrame(rows)
        return df if not df.empty else pd.DataFrame(columns=cols)
    except: 
        return pd.DataFrame(columns=cols)

def fetch_progettazione():
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}
    url_site = f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}"
    site_id = requests.get(url_site, headers=headers).json().get("id")
    cols = ["ID", "Commessa", "Documento", "Stato", "Assegnato"]
    try:
        data = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{LISTA_PROGETTAZIONE}/items?expand=fields", headers=headers).json()
        rows = []
        if "value" in data:
            for item in data["value"]:
                if "fields" in item:
                    f = item["fields"]
                    rows.append({
                        "ID": item["id"],
                        "Commessa": f.get("Commessa", ""),
                        "Documento": f.get("Title", ""),
                        "Stato": f.get("Stato_Doc", "Da Fare"),
                        "Assegnato": f.get("Assegnato", "")
                    })
        df = pd.DataFrame(rows)
        return df if not df.empty else pd.DataFrame(columns=cols)
    except: 
        return pd.DataFrame(columns=cols)

# --- INIZIALIZZAZIONE SESSIONE (FIX ERRORI ATTRIBUTEERROR) ---
if "df_comm" not in st.session_state: st.session_state.df_comm = fetch_commesse()
if "df_sic" not in st.session_state: st.session_state.df_sic = fetch_sicurezza()
if "df_prog" not in st.session_state: st.session_state.df_prog = fetch_progettazione()
if "df_timesheet" not in st.session_state: st.session_state.df_timesheet = None
if "df_plan" not in st.session_state: st.session_state.df_plan = None
if "admin_auth" not in st.session_state: st.session_state.admin_auth = False

# --- 2. NAVIGAZIONE LATERALE ---
with st.sidebar:
    try: st.image("IMG_0002.jpeg", use_container_width=True)
    except Exception: st.title("⚙️ SGP Consulting")
    st.markdown("<br>", unsafe_allow_html=True)
    selected = option_menu(
        menu_title=None, 
        options=["Gestione Commesse", "Analisi Ore", "Pianificazione", "Sicurezza Cantieri", "Progettazione"], 
        icons=["shield-lock", "clock", "calendar-check", "cone-striped", "rulers"], 
        default_index=0, 
        styles={
            "container": {"padding": "5!important", "background-color": "transparent"}, 
            "icon": {"color": "#00AEEF", "font-size": "20px"}, 
            "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#e0fbfc"}, 
            "nav-link-selected": {"background-color": "#00AEEF", "color": "white", "font-weight": "bold"}
        }
    )
    st.markdown("---")

# ==========================================
# 1. LOGICA PAGINA: GESTIONE COMMESSE (PADRE)
# ==========================================
if selected == "Gestione Commesse":
    if not st.session_state.admin_auth:
        st.header("🔐 Accesso Riservato")
        pwd = st.text_input("Password Amministratore", type="password")
        if st.button("Sblocca Dashboard"):
            if pwd == PASSWORD_ADMIN:
                st.session_state.admin_auth = True
                st.rerun()
            else: 
                st.error("❌ Password errata.")
    else:
        col_logo, col_testo = st.columns([0.6, 9.4])
        with col_logo:
            try: st.image("IMG_0002.jpeg", width=55)
            except: pass
        with col_testo: 
            st.header("Gestione Commesse SGP (Livello Padre)")
            
        if st.button("🔄 Sincronizza con SharePoint"): 
            with st.spinner("Aggiornamento dati in corso..."):
                st.session_state.df_comm = fetch_commesse()
                
        df_c = st.session_state.df_comm
        if df_c is not None and not df_c.empty:
            with st.expander("🔍 Strumenti di Filtro e Ricerca", expanded=True):
                c_f1, c_f2, c_f3, c_f4 = st.columns(4)
                
                stati_disponibili = sorted(df_c['Stato'].unique().tolist())
                f_stato = c_f1.multiselect("Stato Operativo:", stati_disponibili, default=["Attiva"] if "Attiva" in stati_disponibili else None)
                
                fatt_disponibili = sorted(df_c['Fatturazione'].unique().tolist())
                f_fatt = c_f2.multiselect("Stato Fatturazione:", fatt_disponibili)
                
                anni_apertura_disp = sorted(df_c['Anno Apertura'].dropna().unique().astype(int).tolist(), reverse=True)
                f_anno_ap = c_f3.multiselect("Anno Apertura:", anni_apertura_disp)
                
                anni_scadenza_disp = sorted(df_c['Scadenza_DT'].dt.year.dropna().unique().astype(int).tolist(), reverse=True)
                f_anno_scad = c_f4.multiselect("Anno Scadenza:", anni_scadenza_disp)
                
                f_testo = st.text_input("Ricerca libera (Codice o Descrizione):")

            df_display = df_c.copy()
            if f_stato: df_display = df_display[df_display['Stato'].isin(f_stato)]
            if f_fatt: df_display = df_display[df_display['Fatturazione'].isin(f_fatt)]
            if f_anno_ap: df_display = df_display[df_display['Anno Apertura'].isin(f_anno_ap)]
            if f_anno_scad: df_display = df_display[df_display['Scadenza_DT'].dt.year.isin(f_anno_scad)]
            if f_testo: 
                df_display = df_display[df_display['Codice'].str.contains(f_testo, case=False, na=False) | 
                                        df_display['Descrizione'].str.contains(f_testo, case=False, na=False)]

            m1, m2, m3 = st.columns(3)
            m1.metric("Voci in elenco", len(df_display))
            m2.metric("Budget Tot. Filtrato", f"€ {df_display['Importo Totale'].sum():,.2f}")
            m3.metric("Avanzamento Medio", f"{df_display['Avanzamento %'].mean():.1f}%" if not df_display.empty else "0%")

            st.dataframe(df_display.drop(columns=['ID', 'Scadenza_DT', 'Anno Apertura'], errors='ignore').sort_values(by="Codice"), use_container_width=True)
            
            st.markdown("---")
            tab_add, tab_edit, tab_mass = st.tabs(["➕ Nuova Commessa", "✏️ Modifica / Cartelle", "🚀 Importazione Massiva"])
            
            with tab_add:
                with st.form("new_comm"):
                    c_form1, c_form2 = st.columns(2)
                    nc = c_form1.text_input("Codice Commessa (es. 25_028_TEC)*")
                    nd = c_form2.text_input("Descrizione")
                    c_form3, c_form4, c_form5 = st.columns(3)
                    ns = c_form3.selectbox("Stato", ["Attiva", "Chiusa", "Sospesa"])
                    nf = c_form4.selectbox("Fatturazione", ["Da Fatturare", "In Corso", "Saldata"])
                    nsc = c_form5.date_input("Scadenza")
                    c_form6, c_form7, c_form8 = st.columns(3)
                    ni = c_form6.number_input("Importo Totale (€)", min_value=0.0)
                    na = c_form7.slider("Avanzamento (%)", 0, 100, 0)
                    nsal = c_form8.number_input("SAL Emessi", min_value=0)
                    
                    col_btn1, col_btn2 = st.columns(2)
                    btn_tutto = col_btn1.form_submit_button("💾 Salva e Genera Cartelle", type="primary")
                    btn_dati = col_btn2.form_submit_button("📝 Salva Solo Dati (Archivio)")
                    
                    if btn_tutto or btn_dati:
                        if nc:
                            h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                            sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                            payload = {"fields": {"Title": nc, "Descrizione": nd, "Stato": ns, "Stato_Fatturazione": nf, "Importo_x0020_Totale": ni, "Avanzamento": na, "SalEmessi": nsal, "Scadenza": nsc.strftime("%Y-%m-%d")}}
                            requests.post(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_COMMESSE}/items", headers=h, json=payload)
                            if btn_tutto:
                                with st.spinner("Creazione cartelle..."): crea_cartelle_madre(nc, nd)
                            st.success("✅ Salvataggio completato!"); st.session_state.df_comm = fetch_commesse(); st.rerun()

            with tab_edit:
                opzioni_c = dict(zip(df_c['ID'], df_c['Codice'] + " - " + df_c['Descrizione']))
                sel_c = st.selectbox("Seleziona commessa:", [""] + list(opzioni_c.keys()), format_func=lambda x: "Scegli..." if x == "" else opzioni_c[x])
                if sel_c != "":
                    r = df_c[df_c['ID'] == sel_c].iloc[0]
                    with st.form("edit_comm"):
                        ce1, ce2, ce3 = st.columns(3)
                        es = ce1.selectbox("Stato", ["Attiva", "Chiusa", "Sospesa"], index=["Attiva", "Chiusa", "Sospesa"].index(r['Stato']))
                        ef = ce2.selectbox("Fatturazione", ["Da Fatturare", "In Corso", "Saldata"], index=["Da Fatturare", "In Corso", "Saldata"].index(r['Fatturazione']))
                        ea = ce3.slider("Avanzamento (%)", 0, 100, int(r['Avanzamento %']))
                        b1, b2, b3 = st.columns(3)
                        if b1.form_submit_button("🔄 Aggiorna Dati"):
                            h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                            sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                            requests.patch(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_COMMESSE}/items/{sel_c}", headers=h, json={"fields": {"Stato": es, "Stato_Fatturazione": ef, "Avanzamento": ea}})
                            st.success("Dati aggiornati!"); st.session_state.df_comm = fetch_commesse(); st.rerun()
                        if b2.form_submit_button("📁 Crea Cartelle (Retroattivo)", type="primary"):
                            crea_cartelle_madre(r['Codice'], r['Descrizione'])
                            st.success("Cartelle generate!"); 
                        if b3.form_submit_button("🗑️ Elimina"):
                            h = {"Authorization": f"Bearer {get_ms_token()}"}
                            sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                            requests.delete(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_COMMESSE}/items/{sel_c}", headers=h)
                            st.warning("Eliminata!"); st.session_state.df_comm = fetch_commesse(); st.rerun()

            with tab_mass:
                st.write("Incolla qui l'elenco delle cartelle copiate dal Finder del Mac.")
                txt_mass = st.text_area("Elenco (es: 25_001_STP Progetto...):", height=200)
                if st.button("🚀 Avvia Importazione Automatica"):
                    if txt_mass:
                        h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                        sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                        righe = [r.strip() for r in txt_mass.split('\n') if r.strip()]
                        for r_mass in righe:
                            p_mass = r_mass.split(" ", 1)
                            cod_m = p_mass[0]
                            des_m = p_mass[1] if len(p_mass) > 1 else ""
                            requests.post(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_COMMESSE}/items", headers=h, json={"fields": {"Title": cod_m, "Descrizione": des_m, "Stato": "Attiva"}})
                        st.success(f"Caricate {len(righe)} commesse!"); st.session_state.df_comm = fetch_commesse(); st.rerun()

# ==========================================
# 2. ANALISI ORE
# ==========================================
elif selected == "Analisi Ore":
    st.header("⏱️ Controllo Ore Collaboratori")
    
    if st.button("🔄 Scarica Ore da SharePoint"): 
        with st.spinner("Sincronizzazione in corso..."):
            st.session_state.df_timesheet = fetch_timesheet()
            
    df_t = st.session_state.df_timesheet
    if df_t is not None and not df_t.empty:
        # --- FILTRI ---
        with st.expander("🔍 Filtri Avanzati Analisi", expanded=True):
            col_f1, col_f2 = st.columns(2)
            f_coll = col_f1.multiselect("Filtra Collaboratore:", sorted(df_t['Collaboratore'].unique().tolist()))
            f_comm = col_f2.multiselect("Filtra Commessa:", sorted(df_t['Commessa'].unique().tolist()))
            
        df_t_filt = df_t.copy()
        if f_coll: df_t_filt = df_t_filt[df_t_filt['Collaboratore'].isin(f_coll)]
        if f_comm: df_t_filt = df_t_filt[df_t_filt['Commessa'].isin(f_comm)]
        
        st.markdown("---")
        
        # --- TABELLA RIASSUNTIVA SOMMA ORE (Sostituisce il grafico) ---
        st.subheader("📊 Riepilogo Somma Ore")
        if not df_t_filt.empty:
            # Calcola la somma per Commessa e Collaboratore
            df_summary = df_t_filt.groupby(['Commessa', 'Collaboratore'])['Ore'].sum().reset_index()
            st.dataframe(df_summary, use_container_width=True)
            
        st.markdown("---")
        
        # --- TABS: DETTAGLIO E CORREZIONE ERRORI ---
        tab_dettaglio, tab_modifica = st.tabs(["📋 Dettaglio Registrazioni", "✏️ Modifica / Correggi Ore"])
        
        with tab_dettaglio:
            st.dataframe(df_t_filt.drop(columns=['ID'], errors='ignore').sort_values(by="Data", ascending=False), use_container_width=True)
            
        with tab_modifica:
            if 'ID' in df_t_filt.columns:
                # Creiamo una tendina per capire bene quale riga stiamo modificando
                opzioni_ore = dict(zip(df_t_filt['ID'], df_t_filt['Data'].astype(str) + " | " + df_t_filt['Collaboratore'] + " | " + df_t_filt['Commessa'] + " (" + df_t_filt['Ore'].astype(str) + "h)"))
                id_ore = st.selectbox("Seleziona la registrazione da correggere:", [""] + list(opzioni_ore.keys()), format_func=lambda x: "Scegli..." if x == "" else opzioni_ore[x])
                
                if id_ore != "":
                    riga_ore = df_t_filt[df_t_filt['ID'] == id_ore].iloc[0]
                    with st.form("form_edit_ore"):
                        nuove_ore = st.number_input("Correggi le Ore Lavorate:", value=float(riga_ore['Ore']), step=0.5)
                        
                        if st.form_submit_button("🔄 Salva Correzione"):
                            token = get_ms_token()
                            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
                            url_site = f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}"
                            sid = requests.get(url_site, headers=headers).json().get("id")
                            
                            # Aggiorna su SharePoint
                            requests.patch(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_TIMESHEET}/items/{id_ore}", headers=headers, json={"fields": {"OreLavorate": nuove_ore}})
                            st.success("✅ Ore aggiornate con successo!")
                            st.session_state.df_timesheet = fetch_timesheet()
                            st.rerun()
            else:
                st.warning("⚠️ Ricarica le ore da SharePoint (tasto blu in alto) per abilitare la modifica.")
    else:
        st.info("ℹ️ Nessuna ora registrata trovata. Assicurati di scaricare i dati.")
# ==========================================
# 3. PIANIFICAZIONE
# ==========================================
elif selected == "Pianificazione":
    st.header("📅 Pianificazione Team")
    
    # Preparo la lista delle commesse attive per la tendina
    lista_c_attive = ["Nessuna Commessa Attiva"]
    if st.session_state.df_comm is not None and not st.session_state.df_comm.empty:
        attive = st.session_state.df_comm[st.session_state.df_comm['Stato'] == 'Attiva']['Codice'].tolist()
        if attive: lista_c_attive = sorted(attive)
            
    with st.container(border=True):
        st.subheader("🚀 Assegna un nuovo Incarico")
        col_p1, col_p2 = st.columns(2)
        ragazzo = col_p1.selectbox("👤 Collaboratore:", ["Cristiano", "Stefano", "Giuditta", "Gianluca"])
        com_scelta = col_p2.selectbox("🏗️ Su quale Commessa:", lista_c_attive)
        
        data_p = st.date_input("📆 Giorno lavorativo:", datetime.date.today())
        note_p = st.text_area("📝 Istruzioni operative:")
        
        if st.button("Invia Pianificazione", type="primary"):
            h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
            sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
            payload = {"fields": {"Title": ragazzo, "Commessa": com_scelta, "Data": data_p.strftime("%Y-%m-%d"), "Note": note_p}}
            requests.post(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_PIANIFICAZIONE}/items", headers=h, json=payload)
            st.success("✅ Incarico salvato correttamente!")
            st.session_state.df_plan = fetch_pianificazione()
            st.rerun()

    st.markdown("---")
    st.subheader("🗓️ Visualizzazione Impegni")
    
    if st.button("🔄 Aggiorna Calendari") or st.session_state.df_plan is None: 
        with st.spinner("Caricamento dati..."):
            st.session_state.df_plan = fetch_pianificazione()
    
    df_p = st.session_state.df_plan
    if df_p is not None and not df_p.empty:
        import calendar
        
        tab_pers, tab_team, tab_list = st.tabs(["👤 Calendario Personale (Mese)", "👥 Calendario Team (Settimana)", "📋 Lista Dettagliata Note"])
        
        # ==========================================
        # TAB 1: CALENDARIO PERSONALE (STILE FOTO)
        # ==========================================
        with tab_pers:
            cp1, cp2, cp3 = st.columns(3)
            ragazzi_disp = sorted(df_p['Collaboratore'].unique().tolist())
            rag_sel = cp1.selectbox("Seleziona Collaboratore:", ragazzi_disp if ragazzi_disp else ["Nessuno"])
            
            anni = sorted(df_p['Giorno'].dt.year.dropna().unique().astype(int).tolist(), reverse=True)
            anno_sel = cp2.selectbox("Seleziona Anno:", anni if anni else [datetime.date.today().year], key="anno_pers")
            
            nomi_mesi = {1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio", 6: "Giugno", 7: "Luglio", 8: "Agosto", 9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre"}
            mese_sel = cp3.selectbox("Seleziona Mese:", list(range(1, 13)), index=datetime.date.today().month-1, format_func=lambda x: nomi_mesi[x], key="mese_pers")
            
            # Filtriamo i dati per il ragazzo e il mese scelto
            df_rag = df_p[(df_p['Collaboratore'] == rag_sel) & (df_p['Giorno'].dt.year == anno_sel) & (df_p['Giorno'].dt.month == mese_sel)]
            
            # Mappiamo i task sui giorni
            task_dict = {}
            for _, row in df_rag.iterrows():
                g = row['Giorno'].day
                if g not in task_dict: task_dict[g] = []
                task_dict[g].append(row['Cantiere'])
                
            st.markdown(f"#### {nomi_mesi[mese_sel]} {anno_sel} - {rag_sel}")
            
            # Disegniamo il calendario a griglia
            cal = calendar.monthcalendar(anno_sel, mese_sel)
            giorni_settimana = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"]
            
            # Intestazioni giorni
            cols = st.columns(7)
            for i, d in enumerate(giorni_settimana):
                cols[i].markdown(f"<div style='text-align:center; color:#00AEEF; font-weight:bold;'>{d}</div>", unsafe_allow_html=True)
            st.markdown("<hr style='margin: 0.5em 0;'>", unsafe_allow_html=True)
            
            # Settimane e giorni
            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    with cols[i]:
                        if day == 0:
                            st.write("") # Cella vuota se il mese non inizia/finisce qui
                        else:
                            st.markdown(f"**{day}**")
                            if day in task_dict:
                                for t in task_dict[day]:
                                    # Grafica del "blocchetto" per il task
                                    st.markdown(f"<div style='padding: 5px; background-color: #00AEEF; color: white; border-radius: 4px; font-size: 13px; margin-bottom: 2px; text-align: center;'>{t}</div>", unsafe_allow_html=True)
                            st.markdown("<div style='height: 40px; border-bottom: 1px solid #e0e0e0;'></div>", unsafe_allow_html=True) # Linea divisoria

        # ==========================================
        # TAB 2: CALENDARIO TEAM (SETTIMANALE)
        # ==========================================
        with tab_team:
            data_scelta = st.date_input("Scegli un giorno qualsiasi della settimana da controllare:", datetime.date.today())
            
            # Calcoliamo il Lunedì e la Domenica di quella settimana
            start_of_week = data_scelta - datetime.timedelta(days=data_scelta.weekday())
            dates_of_week = [start_of_week + datetime.timedelta(days=i) for i in range(7)]
            
            mask = df_p['Giorno'].dt.date.isin(dates_of_week)
            df_week = df_p[mask].copy()
            
            giorni_nomi_short = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
            all_cols_fmt = [f"{giorni_nomi_short[d.weekday()]} {d.strftime('%d/%m')}" for d in dates_of_week]
            
            st.markdown(f"#### Settimana dal {start_of_week.strftime('%d/%m/%Y')} al {dates_of_week[-1].strftime('%d/%m/%Y')}")
            
            if not df_week.empty:
                df_week['Giorno_Fmt'] = df_week['Giorno'].apply(lambda x: f"{giorni_nomi_short[x.weekday()]} {x.strftime('%d/%m')}")
                # Pivot per avere i ragazzi sulle righe e i giorni sulle colonne
                pivot_week = df_week.pivot_table(index='Collaboratore', columns='Giorno_Fmt', values='Cantiere', aggfunc=lambda x: ' | '.join(x)).fillna("")
                
                # Assicuriamoci che tutte e 7 le colonne ci siano sempre nell'ordine giusto
                for c in all_cols_fmt:
                    if c not in pivot_week.columns: pivot_week[c] = ""
                pivot_week = pivot_week[all_cols_fmt]
                
                st.dataframe(pivot_week, use_container_width=True)
            else:
                st.info("📭 Nessun incarico inserito per l'intero team in questa settimana.")
                # Mostriamo comunque la griglia vuota pronta
                df_vuoto = pd.DataFrame(index=["Cristiano", "Stefano", "Giuditta", "Gianluca"], columns=all_cols_fmt).fillna("")
                st.dataframe(df_vuoto, use_container_width=True)

        # ==========================================
        # TAB 3: LISTA E NOTE OPERATIVE
        # ==========================================
        with tab_list:
             st.dataframe(df_p.drop(columns=['ID'], errors='ignore').sort_values(by="Giorno", ascending=False), use_container_width=True)
             
    else:
        st.info("ℹ️ Nessun incarico pianificato trovato nell'archivio.")
        

# ==========================================
# 4. LOGICA PAGINA: SICUREZZA CANTIERI (FIGLI)
# ==========================================
elif selected == "Sicurezza Cantieri":
    st.header("Gestione Sicurezza Cantieri 🏗️")
    if st.button("🔄 Sincronizza Cantieri Sicurezza") or st.session_state.df_sic is None: 
        with st.spinner("Scaricamento dati da SharePoint..."):
            st.session_state.df_sic = fetch_sicurezza()
    
    df_s = st.session_state.df_sic
    if df_s is not None and not df_s.empty:
        # --- RIPRISTINO FILTRI E METRICHE ---
        with st.expander("🔍 Filtri e Ricerca Sotto-Commesse", expanded=True):
            f_col1, f_col2, f_col3, f_col4 = st.columns(4)
            stati_disp = sorted(df_s['Stato'].unique().tolist())
            st_c = f_col1.multiselect("Stato:", stati_disp, default=["Attiva"] if "Attiva" in stati_disp else None)
            f_m = f_col2.multiselect("Commessa Madre:", sorted(df_s['Commessa_Madre'].unique().tolist()))
            f_p = f_col3.multiselect("Stato PSC:", sorted(df_s['Stato_PSC'].unique().tolist()))
            f_q = f_col4.text_input("Ricerca libera (Codice/Descrizione):")

        # Applica i filtri scelti
        df_disp_s = df_s.copy()
        if st_c: df_disp_s = df_disp_s[df_disp_s['Stato'].isin(st_c)]
        if f_m: df_disp_s = df_disp_s[df_disp_s['Commessa_Madre'].isin(f_m)]
        if f_p: df_disp_s = df_disp_s[df_disp_s['Stato_PSC'].isin(f_p)]
        if f_q: 
            df_disp_s = df_disp_s[df_disp_s['Sotto_Commessa'].str.contains(f_q, case=False, na=False) | 
                                 df_disp_s['Descrizione'].str.contains(f_q, case=False, na=False)]

        # --- CALCOLO VALORE ECONOMICO (AGGIORNATO) ---
        if 'KM' in df_disp_s.columns:
            # Calcolo automatico: 430€ per Km + 180€ fissi a cantiere
            df_disp_s['Valore Stimato (€)'] = (df_disp_s['KM'] * 430) + 180
            
            # Mostra i box con i numeri (Metriche)
            m1, m2, m3 = st.columns(3)
            m1.metric("Cantieri Trovati", len(df_disp_s))
            m2.metric("KM Totali Scavo", f"{df_disp_s['KM'].sum():.2f} km")
            m3.metric("Valore Totale", f"€ {df_disp_s['Valore Stimato (€)'].sum():,.2f}")

        # Mostra tabella finale filtrata
        st.dataframe(df_disp_s.drop(columns=['ID'], errors='ignore').sort_values(by="Sotto_Commessa"), use_container_width=True)
    else:
        st.info("⚠️ Al momento non ci sono cantieri in archivio. Usa il tab 'Importazione'.")

    st.markdown("---")
    
    # --- TABS GESTIONE E IMPORTAZIONE ---
    t_add_s, t_edit_s, t_mass_s = st.tabs(["➕ Nuova Sotto-Commessa", "✏️ Gestione", "🚀 Importazione Massiva"])
    
    with t_add_s:
        with st.form("new_figlio"):
            c_f1, c_f2 = st.columns(2)
            madri_disp = sorted(st.session_state.df_comm['Codice'].tolist()) if not st.session_state.df_comm.empty else []
            m_scelta = c_f1.selectbox("Seleziona Commessa Madre*", madri_disp)
            s_cod = c_f2.text_input("Codice Sotto-Commessa (es. 01)*")
            
            s_desc = st.text_input("Descrizione Cantiere")
            
            c_f3, c_f4, c_f5 = st.columns(3)
            s_stato = c_f3.selectbox("Stato Cantiere", ["Attiva", "Chiusa", "Sospesa"])
            s_psc = c_f4.selectbox("Stato PSC", ["Da Redigere", "In Lavorazione", "Completato", "N/A"])
            s_km = c_f5.number_input("Lunghezza Scavo (KM)", min_value=0.0, step=0.1)
            
            # --- AGGIUNTO FLAG RIUNIONE NELLA CREAZIONE ---
            s_riunione = st.checkbox("Riunione di Coordinamento Fatta")
            
            col_btn_s1, col_btn_s2 = st.columns(2)
            btn_salva_cartelle_s = col_btn_s1.form_submit_button("💾 Salva e Genera Cartelle", type="primary")
            btn_salva_dati_s = col_btn_s2.form_submit_button("📝 Salva Solo Dati")
            
            if btn_salva_cartelle_s or btn_salva_dati_s:
                if m_scelta and s_cod:
                    h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                    sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                    
                    payload = {
                        "fields": {
                            "Commessa": m_scelta,
                            "Sotto_x002d_Commessa_x0020__x002": s_cod,
                            "DescrizioneCantiere": s_desc,
                            "Stato": s_stato,
                            "StatoPSC": s_psc,
                            "Lunghezza_KM": s_km,
                            "RiunionediCoordinamento": s_riunione
                        }
                    }
                    requests.post(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_SICUREZZA}/items", headers=h, json=payload)
                    
                    if btn_salva_cartelle_s:
                        with st.spinner("Creazione sottocartelle in corso..."):
                            desc_m = ""
                            if not st.session_state.df_comm.empty and m_scelta in st.session_state.df_comm['Codice'].values:
                                desc_m = st.session_state.df_comm[st.session_state.df_comm['Codice'] == m_scelta].iloc[0]['Descrizione']
                            
                            cartella_madre_nome = f"{str(m_scelta).strip()} {str(desc_m).strip()}" if desc_m else str(m_scelta).strip()
                            cartella_figlio_nome = f"{str(s_cod).strip()} {str(s_desc).strip()}" if s_desc else str(s_cod).strip()
                            
                            crea_cartelle_figlio(cartella_madre_nome, cartella_figlio_nome)
                    
                    st.success("✅ Cantiere figlio creato correttamente!")
                    st.session_state.df_sic = fetch_sicurezza()
                    st.rerun()
                else:
                    st.warning("⚠️ Compila i campi obbligatori (*).")

    with t_edit_s:
        if not df_s.empty:
            opzioni_s = dict(zip(df_s['ID'], df_s['Commessa_Madre'] + " - " + df_s['Sotto_Commessa'] + " (" + df_s['Descrizione'] + ")"))
            sel_id_s = st.selectbox("Seleziona Sotto-Commessa da gestire:", [""] + list(opzioni_s.keys()), format_func=lambda x: "Scegli..." if x == "" else opzioni_s[x])
            
            if sel_id_s != "":
                r_s = df_s[df_s['ID'] == sel_id_s].iloc[0]
                with st.form("edit_figlio"):
                    ce_s1, ce_s2, ce_s3 = st.columns(3)
                    
                    st_val = r_s['Stato'] if r_s['Stato'] in ["Attiva", "Chiusa", "Sospesa"] else "Attiva"
                    psc_val = r_s['Stato_PSC'] if r_s['Stato_PSC'] in ["Da Redigere", "In Lavorazione", "Completato", "N/A"] else "Da Redigere"
                    
                    es_stato = ce_s1.selectbox("Stato", ["Attiva", "Chiusa", "Sospesa"], index=["Attiva", "Chiusa", "Sospesa"].index(st_val))
                    es_psc = ce_s2.selectbox("Stato PSC", ["Da Redigere", "In Lavorazione", "Completato", "N/A"], index=["Da Redigere", "In Lavorazione", "Completato", "N/A"].index(psc_val))
                    es_km = ce_s3.number_input("KM", value=float(r_s.get('KM', 0.0)), step=0.1)
                    es_riunione = st.checkbox("Riunione di Coordinamento Fatta", value=bool(r_s.get('Riunione_Fatta', False)))
                    
                    b_s1, b_s2 = st.columns(2)
                    if b_s1.form_submit_button("🔄 Aggiorna Dati"):
                        h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                        sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                        requests.patch(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_SICUREZZA}/items/{sel_id_s}", headers=h, json={"fields": {"Stato": es_stato, "StatoPSC": es_psc, "Lunghezza_KM": es_km, "RiunionediCoordinamento": es_riunione}})
                        st.success("✅ Dati aggiornati!")
                        st.session_state.df_sic = fetch_sicurezza()
                        st.rerun()
                    
                    if b_s2.form_submit_button("📁 Genera Cartelle (Retroattivo)", type="primary"):
                        with st.spinner("Creazione in corso..."):
                            c_madre_cod = r_s['Commessa_Madre']
                            c_madre_desc = ""
                            if not st.session_state.df_comm.empty and c_madre_cod in st.session_state.df_comm['Codice'].values:
                                c_madre_desc = st.session_state.df_comm[st.session_state.df_comm['Codice'] == c_madre_cod].iloc[0]['Descrizione']
                            
                            n_madre = f"{str(c_madre_cod).strip()} {str(c_madre_desc).strip()}" if c_madre_desc else str(c_madre_cod).strip()
                            n_figlio = f"{str(r_s['Sotto_Commessa']).strip()} {str(r_s['Descrizione']).strip()}" if r_s['Descrizione'] else str(r_s['Sotto_Commessa']).strip()
                            crea_cartelle_figlio(n_madre, n_figlio)
                        st.success("✅ Sottocartelle generate su SharePoint!")

    with t_mass_s:
        st.write("Importa più cantieri associandoli a una Commessa Madre specifica.")
        madri = sorted(st.session_state.df_comm['Codice'].tolist()) if not st.session_state.df_comm.empty else []
        p_scelto = st.selectbox("Seleziona Padre per importazione:", madri)
        txt_s = st.text_area("Incolla elenco figli (es. 01_Cabina...):", height=150)
        
        if st.button("Avvia Importazione Figli", type="primary"):
            if txt_s and p_scelto:
                h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                sid_req = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h)
                
                if sid_req.status_code == 200:
                    sid = sid_req.json().get("id")
                    righe_pulite = [r.strip() for r in txt_s.split('\n') if r.strip()]
                    
                    with st.spinner(f"Importazione di {len(righe_pulite)} cantieri in corso..."):
                        for r_s in righe_pulite:
                            parti = r_s.split("_", 1)
                            codice_figlio = parti[0]
                            desc_figlio = parti[1] if len(parti)>1 else ""
                            
                            payload = {
                                "fields": {
                                    "Commessa": p_scelto, 
                                    "Sotto_x002d_Commessa_x0020__x002": codice_figlio, 
                                    "DescrizioneCantiere": desc_figlio, 
                                    "Stato": "Attiva"
                                }
                            }
                            requests.post(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_SICUREZZA}/items", headers=h, json=payload)
                            
                    st.success("✅ Importazione massiva completata!")
                    st.session_state.df_sic = fetch_sicurezza()
                    st.rerun()
# ==========================================
# 5. LOGICA PAGINA: PROGETTAZIONE 📐
# ==========================================
elif selected == "Progettazione":
    st.header("📐 Checklist Progettazione Documentale")
    if st.button("🔄 Aggiorna Lista") or st.session_state.df_prog is None: 
        with st.spinner("Scaricamento dati..."):
            st.session_state.df_prog = fetch_progettazione()
    
    df_p = st.session_state.df_prog
    madri = sorted(st.session_state.df_comm['Codice'].tolist()) if not st.session_state.df_comm.empty else []
    sel_m = st.selectbox("Seleziona Progetto da monitorare:", ["Tutti"] + madri)
    
    df_f = df_p.copy()
    if not df_f.empty and "Commessa" in df_f.columns:
        if sel_m != "Tutti":
            df_f = df_f[df_f['Commessa'] == sel_m]
        
        # --- PROPOSTA: CRUSCOTTO AVANZAMENTO E SEMAFORI ---
        if sel_m != "Tutti" and not df_f.empty:
            task_ok = len(df_f[df_f['Stato'].isin(["Completato", "Approvato"])])
            task_tot = len(df_f)
            perc = int((task_ok / task_tot) * 100)
            
            st.markdown(f"##### Avanzamento Documentale: {task_ok} su {task_tot} completati")
            st.progress(perc / 100.0)
            if perc == 100:
                st.success("✅ Tutta la documentazione di questo progetto è stata Approvata!")
        
        # Creazione dei semafori visivi per la tabella
        def semaforo(stato):
            if stato in ["Completato", "Approvato"]: return f"🟢 {stato}"
            elif stato == "Da revisionare": return f"🟡 {stato}"
            else: return f"🔴 {stato}"
            
        df_f['Stato Visivo'] = df_f['Stato'].apply(semaforo)
        
        # Mostriamo la tabella (altezza fissa a 300 per renderla bella larga)
        colonne_da_mostrare = ['Commessa', 'Documento', 'Stato Visivo', 'Assegnato']
        st.dataframe(df_f[colonne_da_mostrare], use_container_width=True, height=300)
    else:
        st.info("📭 Nessun documento presente per questa commessa. Aggiungine uno qui sotto.")

    st.markdown("---")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.subheader("➕ Aggiungi Task")
        with st.form("new_doc"):
            d_n = st.text_input("Nome Documento (es. Planimetria)")
            d_a = st.multiselect("Assegna a:", ["Cristiano", "Stefano", "Giuditta", "Gianluca"])
            if st.form_submit_button("Inserisci in Checklist"):
                if d_n and sel_m != "Tutti":
                    h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                    sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                    requests.post(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_PROGETTAZIONE}/items", headers=h, 
                                  json={"fields": {"Title": d_n, "Commessa": sel_m, "Stato_Doc": "Da Fare", "Assegnato": ", ".join(d_a)}})
                    st.success("✅ Documento aggiunto!")
                    st.session_state.df_prog = fetch_progettazione(); st.rerun()
                else: st.warning("⚠️ Seleziona una commessa specifica (non 'Tutti') prima di aggiungere un task.")

    with col_d2:
        st.subheader("✏️ Aggiorna Stato")
        if not df_f.empty:
            s_id = st.selectbox("Scegli task da aggiornare:", df_f['ID'].tolist(), format_func=lambda x: df_f[df_f['ID']==x]['Documento'].values[0])
            n_st = st.selectbox("Nuovo Stato:", ["Da Fare", "Completato", "Da revisionare", "Approvato"])
            if st.button("Salva Modifica"):
                h = {"Authorization": f"Bearer {get_ms_token()}", "Content-Type": "application/json"}
                sid = requests.get(f"https://graph.microsoft.com/v1.0/sites/sgpconsultingstp-my.sharepoint.com:/personal/{PERCORSO_PERSONALE}", headers=h).json()["id"]
                requests.patch(f"https://graph.microsoft.com/v1.0/sites/{sid}/lists/{LISTA_PROGETTAZIONE}/items/{s_id}", headers=h, json={"fields": {"Stato_Doc": n_st}})
                st.success("✅ Stato aggiornato!")
                st.session_state.df_prog = fetch_progettazione(); st.rerun()
