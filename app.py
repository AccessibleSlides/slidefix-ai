import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from openai import OpenAI
import base64
from io import BytesIO

# 1. PAGE CONFIGURATION
st.set_page_config(
    page_title="Accessible Slides", 
    page_icon="♿", 
    layout="centered"
)

# Initialize Session State
if "license_valid" not in st.session_state:
    st.session_state["license_valid"] = False
if "org_name" not in st.session_state:
    st.session_state["org_name"] = ""

# --- LICENSE VERIFICATION FUNCTION ---
def verify_license(key):
    """
    Verifies the key against the specific Gumroad Product:
    'accessibleslides_enterprise'
    """
    # THIS MATCHES YOUR GUMROAD PERMALINK FROM STEP 1
    PRODUCT_PERMALINK = "accessibleslides_enterprise" 
    
    url = "https://api.gumroad.com/v2/licenses/verify"
    params = {
        "product_permalink": PRODUCT_PERMALINK,
        "license_key": key.strip()
    }
    
    try:
        response = requests.post(url, params=params)
        data = response.json()
        
        if data.get("success"):
            # Return True and the email (identifying the Org)
            return True, data['purchase']['email']
        else:
            return False, None
    except Exception:
        return False, None

# --- SIDEBAR: LOGIN & STATUS ---
with st.sidebar:
    st.header("🔐 License Status")
    
    if st.session_state["license_valid"]:
        st.success("✅ Enterprise Active")
        st.caption(f"Registered to:\n{st.session_state['org_name']}")
        
        if st.button("Logout"):
            st.session_state["license_valid"] = False
            st.session_state["org_name"] = ""
            st.rerun()
            
        st.divider()
        st.info("💡 Note: This tool runs on your Organization's OpenAI API Key.")
        st.markdown("[Manage OpenAI Keys](https://platform.openai.com/api-keys)")
        
    else:
        st.write("Please enter your Organization's License Key.")
        input_key = st.text_input("License Key", type="password", placeholder="XXXXXXXX-XXXXXXXX-...")
        
        if st.button("Verify Key"):
            is_valid, org_email = verify_license(input_key)
            if is_valid:
                st.session_state["license_valid"] = True
                st.session_state["org_name"] = org_email
                st.rerun()
            else:
                st.error("❌ Invalid or Refunded Key")
        
        st.divider()
        st.markdown("### Purchase Access")
        st.write("Valid for unlimited users within your organization.")
        st.markdown("👉 **[Buy Site License ($500)](https://accessibleslides.gumroad.com/l/accessibleslides_enterprise)**")

# --- MAIN APPLICATION ---

st.title("Accessible Slides")
st.caption("Automated Alt Text Generation for PowerPoint")

# 1. SECURITY GATE
if not st.session_state["license_valid"]:
    st.warning("🔒 Access Restricted. Please verify your Enterprise License in the sidebar to proceed.")
    st.stop() # Stops the code here if not logged in

# 2. API KEY INPUT (BYOK)
st.markdown("### 1. Configuration")
org_openai_key = st.text_input("Enter your OpenAI API Key (sk-...)", type="password")

if not org_openai_key:
    st.info("👋 Welcome! Please enter your API Key to enable the processing engine.")
    st.stop()

# 3. FILE UPLOADER
st.markdown("### 2. Upload Presentation")
uploaded_file = st.file_uploader("Choose a PowerPoint file (.pptx)", type=["pptx"])

# HELPER: AI VISION LOGIC
def get_ai_desc(client, image_blob):
    b64_image = base64.b64encode(image_blob).decode('utf-8')
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": [
                {"type": "text", "text": "Generate a concise, objective alt text. Do not start with 'Image of'."},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
            ]}], max_tokens=100
        )
        return response.choices[0].message.content
    except: return "Error generating text"

if uploaded_file:
    # 4. PROCESSING BUTTON
    if st.button("✨ Auto-Fix Presentation"):
        
        # VERIFY API KEY VALIDITY FIRST
        try:
            client = OpenAI(api_key=org_openai_key)
            client.models.list() # Simple call to check if key works
        except Exception:
            st.error("❌ The OpenAI API Key provided is invalid or has insufficient credits.")
            st.stop()

        prs = Presentation(uploaded_file)
        slide_count = len(prs.slides)
        
        progress_bar = st.progress(0)
        status = st.empty()
        processed_images = 0
        
        # LOOP THROUGH SLIDES
        for i, slide in enumerate(prs.slides):
            progress_bar.progress((i + 1) / slide_count)
            status.text(f"Scanning Slide {i+1} of {slide_count}...")
            
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    # Generate Text
                    desc = get_ai_desc(client, shape.image.blob)
                    try:
                        # Inject Text into PPTX XML
                        shape._element.nvPicPr.cNvPr.set('descr', desc)
                        shape._element.nvPicPr.cNvPr.set('name', "Image")
                        processed_images += 1
                    except: pass
        
        # SUCCESS & DOWNLOAD
        status.success(f"✅ Success! Fixed {processed_images} images across {slide_count} slides.")
        
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download Accessible PPTX",
            data=output,
            file_name=f"Accessible_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )