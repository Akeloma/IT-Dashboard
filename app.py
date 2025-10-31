import streamlit as st
from io import BytesIO

# === Page setup ===
st.set_page_config(page_title="IT Dashboard", page_icon="📊", layout="centered")
st.title("📊 IT Dashboards")

# === Sidebar navigation ===
page = st.sidebar.selectbox(
    "🔧 Select Dashboard",
    [
        "IT Cockpit",
        "Toxic Data Appender"
    ]
)

# === Shared uploader helper ===
def save_uploaded_file(uploaded_file, filename):
    with open(filename, "wb") as f:
        f.write(uploaded_file.read())
    return filename

# === Page Routing ===
if page == "IT Cockpit":
    import ITDashboard
    ITDashboard.main()

elif page == "Toxic Data Appender":
    import ToxicDashboard
    ToxicDashboard.main()
