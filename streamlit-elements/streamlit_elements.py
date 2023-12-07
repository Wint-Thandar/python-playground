import streamlit as st
import time

# 0. session_state element
if "photo" not in st.session_state:
    st.session_state["photo"] = None

# 1. columns element
col1, col2, col3 = st.columns([1, 2, 1])

# 2. markdown element
col1.markdown(" ## Welcome to 000 ")
col1.markdown(" This is a test app of steamlit elements. ")

# 0.1 change session state
def change_photo_state():
    st.session_state["photo"] = "done"

# 3. file_uploader element
uploaded_photo = col2.file_uploader("Upload a photo", on_change = change_photo_state)

# 4. camera_input element
camera_photo = col2.camera_input("Take a photo", on_change = change_photo_state)

# 5. progress bar element
progress_bar = col2.progress(0)

# 5.1. progress bar loading
for percent_complete in range(100):
    time.sleep(0.05)
    progress_bar.progress(percent_complete + 1)

# 6. success element
col2.success("Photo uploaded successfully!")

# 7. metric element
col3.metric(label = "Temperature", value = "60 °C", delta = "3 °C") 

# 8. expander element
with st.expander("Click to read more"):
    st.write("This is a test of expander element.")
    
    if uploaded_photo is None:
        st.image(camera_photo)
    else:
        st.image(uploaded_photo)