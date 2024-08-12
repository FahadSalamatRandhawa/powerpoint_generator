from main import create_ppt

import streamlit as st
from pptx import Presentation
import io as io
import base64



st.title("AI PowerPoint Generator Tool")

# User inputs for slides
st.write("### Enter slide data:")
slides_data = []
num_slides = st.number_input("Number of slides", min_value=1, max_value=5, value=1, step=1)

for i in range(num_slides):
    st.write(f"#### Slide {i + 1}")
    title = st.text_input(f"Title for Slide {i + 1}")
    content = st.text_area(f"Content for Slide {i + 1}")
    slides_data.append({"title": title, "content": content})


# File paths
ppt_filename = "output_presentation.pptx"

# Generate PowerPoint and convert to PDF
if st.button("Generate PowerPoint"):
    create_ppt(slides_data, ppt_filename)
    
    # Show generated PowerPoint
    with open(ppt_filename, "rb") as ppt_file:
        st.download_button(
            label="Download PowerPoint",
            data=ppt_file,
            file_name=ppt_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        ppt_file.seek(0)
        ppt_data = ppt_file.read()

        # Display preview of the PowerPoint
        ppt_str = base64.b64encode(ppt_data).decode("utf-8")
        ppt_html = f'<iframe src="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{ppt_str}" width="100%" height="500px"></iframe>'
        st.write("### PowerPoint Preview:")
        st.markdown(ppt_html, unsafe_allow_html=True)