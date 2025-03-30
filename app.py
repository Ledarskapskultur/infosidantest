import streamlit as st
import uuid
import pandas as pd
import requests
from datetime import datetime
from io import BytesIO
from save_to_sharepoint import get_token, get_site_id

# --- Funktion: hämta fältnamn från SharePoint-lista ---
def visa_sharepoint_kolumner(token, site_id, list_name="kunddata1"):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/columns"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        kolumner = response.json().get("value", [])
        st.subheader("📓 Kolumner i Microsoft List")
        for kolumn in kolumner:
            st.write(f"**Display Name:** {kolumn.get('displayName')}  →  **Internal Name:** `{kolumn.get('name')}`")
    else:
        st.error("❌ Kunde inte hämta kolumner från SharePoint-listan.")

# --- Funktion: spara till Microsoft List ---
def append_kunddata_to_sharepoint(referensnummer, namn, telefon, mail, valda_kurser_df, token, site_id):
    list_name = "kunddata1"
    list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/items"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    for _, row in valda_kurser_df.iterrows():
        payload = {
            "fields": {
                "IDnr": referensnummer,
                "field_1": namn,
                "field_2": str(telefon),
                "field_3": mail,
                "field_4": datetime.today().strftime("%Y-%m-%d"),
                "field_5": int(row["Vecka"]),
                "field_6": row["Anläggning"],
                "field_7": row["Ort"],
                "field_8": row["Kursledare"],
                "field_9": float(row["Pris"])
            }
        }

        # Debug visas här – du kan kommentera bort dessa två rader om du inte vill se dem
        st.write("🔢 Payload till SharePoint:", payload)
        response = requests.post(list_url, headers=headers, json=payload)
        st.write(f"📡 Svar från SharePoint (kod {response.status_code}):")
        try:
            st.code(response.json())
        except Exception:
            st.warning("Kunde inte tolka svar som JSON")

        if response.status_code not in (200, 201):
            return False

    return True

# --- Grundinställningar ---
st.set_page_config(page_title="Kursbokning", layout="wide")

# --- SharePoint-inställningar ---
FILNAMN = "kurser aktiv.xlsx"
token = get_token()
site_id = get_site_id(token)

if token and site_id:
    if st.button("🔍 Visa kolumner i SharePoint-listan"):
        visa_sharepoint_kolumner(token, site_id)

# --- SIDOPANEL: Kontaktinformation ---
with st.sidebar:
    st.title("📞 Kontaktinformation")
    col1, col2 = st.columns(2)
    with col1:
        namn = st.text_input("Namn")
    with col2:
        telefon = st.text_input("Telefon")

    epost = st.text_input("E-postadress")

    col3, col4 = st.columns(2)
    with col3:
        vecka_input = st.text_input("Vecka (t.ex. 31, 32-34)")
    with col4:
        maxpris_input = st.text_input("Maxpris (kr)")

    plats_input = st.text_input("Plats")
    bokningsreferens = str(uuid.uuid4())[:8]
    st.text_input("Bokningsreferens (auto)", value=bokningsreferens, disabled=True)

# --- HUVUDDEL ---
st.title("📋 Kursbokningssystem")
st.markdown("Fyll i kontaktuppgifter i sidopanelen för att filtrera tillgängliga kurser.")

def parse_veckor(vecka_str):
    veckor = set()
    if vecka_str:
        delar = vecka_str.split(",")
        for d in delar:
            if "-" in d:
                start, slut = d.split("-")
                veckor.update(range(int(start), int(slut) + 1))
            else:
                veckor.add(int(d))
    return sorted(list(veckor))

if token and site_id:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{FILNAMN}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200 and (vecka_input or maxpris_input or plats_input):
        df = pd.read_excel(BytesIO(response.content))
        df = df.drop_duplicates(subset=["Pris", "Anläggning", "Kursledare"])

        veckor_valda = parse_veckor(vecka_input)
        if veckor_valda:
            df = df[df["Vecka"].isin(veckor_valda)]

        if plats_input:
            df = df[df["Ort"].astype(str).str.lower().str.contains(plats_input.lower())]

        if maxpris_input:
            try:
                gräns = float(maxpris_input)
                df = df[df["Pris"] <= gräns]
            except ValueError:
                st.warning("Maxpris måste vara ett numeriskt värde.")

        st.success(f"✅ {len(df)} kurser hittades")
        st.subheader("✅ Välj kurser")
        valda_kurser = []
        cols = st.columns(4)
        for i, row in enumerate(df.itertuples()):
            label = f"{row.Vecka} – {row.Pris} kr – {row.Anläggning} – {row.Kursledare}"
            if cols[i % 4].checkbox(label, key=f"kurs_{i}"):
                valda_kurser.append(row._asdict())

        if valda_kurser:
            st.markdown("---")
            st.subheader("📦 Valda kurser")
            valda_df = pd.DataFrame(valda_kurser)
            st.dataframe(valda_df[["Vecka", "Anläggning", "Ort", "Kursledare", "Pris"]])

            st.markdown("### 📨 Skicka bekräftelsemail")
            extra_mail = st.text_input("Extra e-postadress (valfritt)")
            meddelande = st.text_area("Meddelande till kunden (valfritt)")

            if st.button("📨 Skicka bekräftelsemail"):
                def skicka_mail(mottagare_lista, namn, telefon, mail, refnr, kurser_df, token, meddelande):
                    kurs_html = kurser_df[["Vecka", "Anläggning", "Ort", "Pris"]].to_html(index=False, border=0)
                    body = f"""
                    <html><body>
                    <p><strong>UGL Kurser</strong><br>Tack för din förfrågan!</p>
                    <p><strong>Referensnummer:</strong> {refnr}<br>
                    <strong>Namn:</strong> {namn}<br>
                    <strong>Telefon:</strong> {telefon}<br>
                    <strong>E-post:</strong> {mail}</p>
                    <p>{meddelande.replace('\\n', '<br>')}</p>
                    <p><strong>Valda kurser:</strong></p>{kurs_html}
                    </body></html>
                    """
                    email_data = {
                        "message": {
                            "subject": "UGL Kurser – Bekräftelse på din förfrågan",
                            "body": {"contentType": "HTML", "content": body},
                            "toRecipients": [{"emailAddress": {"address": e}} for e in mottagare_lista]
                        },
                        "saveToSentItems": "true"
                    }
                    headers = {
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json"
                    }
                    response = requests.post(
                        "https://graph.microsoft.com/v1.0/users/carl-fredrik@ledarskapskultur.se/sendMail",
                        headers=headers, json=email_data)
                    return response.status_code == 202

                mottagare = []
                if epost: mottagare.append(epost)
                if extra_mail: mottagare.append(extra_mail)

                if not mottagare:
                    st.warning("Ingen giltig e-postadress angiven.")
                else:
                    success = skicka_mail(mottagare, namn, telefon, epost, bokningsreferens, valda_df, token, meddelande)
                    if success:
                        st.success("✅ Bekräftelsemail skickat!")
                        saved = append_kunddata_to_sharepoint(bokningsreferens, namn, telefon, epost, valda_df, token, site_id)
                        if saved:
                            st.success("📅 Bokningsdata sparad till SharePoint!")
                        else:
                            st.error("❌ Det gick inte att spara bokningen till SharePoint.")
                    else:
                        st.error("❌ Det gick inte att skicka e-postmeddelandet.")
    else:
        st.info("Fyll i minst ett filter i sidopanelen (vecka, pris eller plats) för att visa tillgängliga kurser.")
else:
    st.error("❌ Kunde inte autentisera eller läsa från SharePoint.")
