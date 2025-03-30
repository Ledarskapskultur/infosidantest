
import streamlit as st
from save_to_sharepoint import get_token, get_site_id

st.title("SharePoint-koppling")

# Hämta känslig data från Streamlit Secrets
client_id = st.secrets["client_id"]
client_secret = st.secrets["client_secret"]
tenant_id = st.secrets["tenant_id"]
domain = st.secrets["domain"]
site_name = st.secrets["site_name"]

# Försök hämta token
try:
    token = get_token(client_id, client_secret, tenant_id)
    st.success("Access token hämtad!")

    # Försök hämta site ID
    site_id = get_site_id(token, domain, site_name)
    st.success(f"Site ID: {site_id}")

except Exception as e:
    st.error(f"Något gick fel: {e}")
