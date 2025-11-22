import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from openai import OpenAI
import base64
from io import BytesIO
from PIL import Image # NEW LIBRARY FOR RESIZING

# 1. PAGE CONFIGURATION
st.set_page_config(
    page_title="Accessible Slides", 
    page_icon="♿", 
    layout="centered",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    h1, h2, h3 { color: #1E3A8A; font-weight: 800; }
    [data-testid="stSidebar"] { background-color: #F1F5F9; border-right: 1px solid #E2E8F0; }
    div.stButton > button:first-child { width: 100%; background-color: #2563EB; color: white; font-weight: bold; padding: 0.75rem; border-radius: 8px; border: none; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    div.stButton > button:first-child:hover { background-color: #1D4ED8; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if "license_valid" not in st.session_state: st.session_state["license_valid"] = False
if "org_name" not in st.session_state: st.session_state["org_name"] = ""

# --- LICENSE VERIFICATION ---
def verify_license(key):
    PRODUCT_ID = "RN-jQLgM_iudPwC-RnQZ6A==" 
    url = "https://api.gumroad.com/v2/licenses/verify"
    payload = {"product_id": PRODUCT_ID, "license_key": key.strip()}
    try:
        response = requests.post(url, data=payload)
        data = response.json()
        if data.get("success"): return True, data['purchase']['email']
        else: return False, None
    except: return False, None

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## ♿ Access")
    if st.session_state["license_valid"]:
        st.success("✅ License Active")
        st.caption(f"**Org:** {st.session_state['org_name']}")
        st.markdown("---")
        if st.button("Logout"):
            st.session_state["license_valid"] = False
            st.session_state["org_name"] = ""
            st.rerun()
    else:
        st.warning("🔒 Portal Locked")
        input_key = st.text_input("License Key", type="password")
        if st.button("Verify Access"):
            is_valid, org_email = verify_license(input_key)
            if is_valid:
                st.session_state["license_valid"] = True
                st.session_state["org_name"] = org_email
                st.rerun()
            else: st.error("Invalid Key")
        st.markdown("---")
        st.markdown("👉 **[Purchase Site License ($500)](https://accessibleslides.gumroad.com/l/accessibleslides_enterprise)**")

# --- MAIN CONTENT ---
st.title("Accessible Slides")
st.markdown("**Enterprise Edition** | Automated Section 508 Compliance")
st.markdown("---")

if not st.session_state["license_valid"]:
    c1, c2, c3 = st.columns([1,2,1])
    with c2:
        st.info("👋 **Welcome.** Please enter your Organization's License Key in the sidebar.")
        st.markdown("<div style='text-align: center; font-size: 60px;'>🔒</div>", unsafe_allow_html=True)
    st.stop()

with st.expander("⚙️ Configuration", expanded=True):
    st.caption("Data Privacy: Files are processed in-memory using your API key.")
    org_openai_key = st.text_input("OpenAI API Key (sk-...)", type="password")

st.markdown("### Upload Presentation")
uploaded_file = st.file_uploader("Drag and drop .pptx file", type=["pptx"], label_visibility="collapsed")

# --- NEW & IMPROVED AI LOGIC ---
def get_ai_desc(client, image_blob):
    try:
        # 1. OPTIMIZE IMAGE (Resize & Convert to JPEG)
        image = Image.open(BytesIO(image_blob))
        
        # Convert to RGB (Fixes CMYK/Alpha issues)
        if image.mode in ("RGBA", "P"):
            image = image.convert("RGB")

        # Resize if too big (Max 1024px) to prevent OpenAI Errors
        max_size = (1024, 1024)
        image.thumbnail(max_size)

        # Save to memory buffer as JPEG
        buffer = BytesIO()
        image.save(buffer, format="JPEG", quality=85)
        b64_image = base64.b64encode(buffer.getvalue()).decode('utf-8')

        # 2. SEND TO OPENAI
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": [
                {"type": "text", "text": "Generate a concise, objective alt text. Do not start with 'Image of'."},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
            ]}], max_tokens=100
        )
        return response.choices[0].message.content

    except Exception as e:
        # Return the ACTUAL error message so we know what's wrong
        return f"Error: {str(e)}"

if uploaded_file:
    st.write(f"📄 **File Selected:** {uploaded_file.name}")
    if not org_openai_key:
        st.error("⚠️ Please enter an API Key above.")
        st.stop()

    if st.button("✨ Auto-Generate Alt Text"):
        try:
            client = OpenAI(api_key=org_openai_key)
            client.models.list() 
        except:
            st.error("❌ Invalid OpenAI API Key.")
            st.stop()

        prs = Presentation(uploaded_file)
        slide_count = len(prs.slides)
        progress_text = "Scanning presentation..."
        my_bar = st.progress(0, text=progress_text)
        processed_images = 0
        errors = 0
        
        for i, slide in enumerate(prs.slides):
            my_bar.progress((i + 1) / slide_count, text=f"Scanning Slide {i+1} of {slide_count}...")
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    desc = get_ai_desc(client, shape.image.blob)
                    
                    # Count errors for metrics
                    if "Error:" in desc:
                        errors += 1
                    else:
                        processed_images += 1
                        
                    try:
                        shape._element.nvPicPr.cNvPr.set('descr', desc)
                        shape._element.nvPicPr.cNvPr.set('name', "Image")
                    except: pass
        
        my_bar.empty()
        st.success("✅ Processing Complete!")
        
        m1, m2, m3 = st.columns(3)
        with m1: st.metric("Slides Scanned", slide_count)
        with m2: st.metric("Images Fixed", processed_images)
        with m3: st.metric("Errors", errors)

        output = BytesIO()
        prs.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Accessible PPTX",
            data=output,
            file_name=f"Accessible_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )