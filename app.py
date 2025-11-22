import streamlit as st

# 1. This MUST be the first command
st.set_page_config(page_title="SlideFix AI", page_icon="🚀", layout="centered")

st.title("🚀 SlideFix AI")

# 2. Secrets Connection
try:
    test_key = st.secrets["OPENAI_API_KEY"]
    test_pass = st.secrets["APP_PASSWORD"]
except FileNotFoundError:
    st.error("❌ ERROR: Secrets file not found.")
    st.stop()
except KeyError:
    st.error("❌ ERROR: Key missing in secrets.toml")
    st.stop()

# --- SECRETS ARE WORKING ---

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from openai import OpenAI
import base64
from io import BytesIO

# Initialize Session State
if "upload_count" not in st.session_state:
    st.session_state["upload_count"] = 0
if "is_pro" not in st.session_state:
    st.session_state["is_pro"] = False

# --- SIDEBAR (LOGIN) ---
with st.sidebar:
    st.header("🔐 Pro Access")
    password = st.text_input("Enter Access Code", type="password")
    
    if password == st.secrets["APP_PASSWORD"]:
        st.session_state["is_pro"] = True
        st.success("✅ Pro Mode Unlocked")
    elif password:
        st.error("❌ Incorrect Code")

    st.divider()
    if st.session_state["is_pro"]:
        st.write("**Plan:** Pro (Unlimited)")
    else:
        st.write("**Plan:** Free Demo")
        st.write(f"**Uploads Used:** {st.session_state['upload_count']}/3")
        st.write("**Limit:** Max 10 slides")
        
        # --- [LINK LOCATION 1: SIDEBAR] ---
        st.markdown("---")
        st.markdown("👉 **[Upgrade to Pro ($12)](https://accessibleslides.gumroad.com/l/fubrm)**") 

# --- MAIN APP LOGIC ---
st.markdown("### Automate PowerPoint Accessibility")
uploaded_file = st.file_uploader("Choose a PowerPoint file", type=["pptx"])

def get_ai_desc(client, image_blob):
    b64_image = base64.b64encode(image_blob).decode('utf-8')
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "Generate a concise, objective alt text. Do not start with 'Image of'."},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
                    ],
                }
            ],
            max_tokens=100,
        )
        return response.choices[0].message.content
    except Exception:
        return "Error generating text"

if uploaded_file:
    prs = Presentation(uploaded_file)
    slide_count = len(prs.slides)
    
    can_proceed = False

    if st.session_state["is_pro"]:
        can_proceed = True
    else:
        if st.session_state["upload_count"] >= 3:
            st.error("🚫 You have used your 3 free uploads.")
            # --- [LINK LOCATION 2: UPLOAD LIMIT ERROR] ---
            st.markdown("👉 **[Click here to Upgrade to Pro ($12)](https://accessibleslides.gumroad.com/l/fubrm)**")
            
        elif slide_count > 10:
            st.error(f"🚫 Free limit is 10 slides. This file has {slide_count}.")
            # --- [LINK LOCATION 3: SLIDE LIMIT ERROR] ---
            st.markdown("👉 **[Click here to Upgrade to Pro ($12)](https://accessibleslides.gumroad.com/l/fubrm)**")
        else:
            can_proceed = True

    if can_proceed:
        if st.button("✨ Fix Presentation Now"):
            client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
            progress_bar = st.progress(0)
            status = st.empty()
            processed_images = 0
            
            for i, slide in enumerate(prs.slides):
                progress_bar.progress((i + 1) / slide_count)
                status.text(f"Processing Slide {i+1}...")
                
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        new_alt = get_ai_desc(client, shape.image.blob)
                        try:
                            shape._element.nvPicPr.cNvPr.set('descr', new_alt)
                            shape._element.nvPicPr.cNvPr.set('name', "Image")
                            processed_images += 1
                        except AttributeError:
                            pass
            
            if not st.session_state["is_pro"]:
                st.session_state["upload_count"] += 1

            status.success(f"Done! {processed_images} images fixed.")
            
            output = BytesIO()
            prs.save(output)
            output.seek(0)
            
            st.download_button(
                label="📥 Download Accessible PPTX",
                data=output,
                file_name=f"Fixed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )