import streamlit as st
from tab1 import excel_extractor_tool
from tab2 import file_renamer_tool
from tab3 import white_to_transparent_tool
from tab4 import pdf_generator_tool
from tab5 import brand_renamer_tool
from tab6 import resize_with_transparent_canvas_tool
from tab7 import center_on_canvas_tool

# Configure Streamlit page
st.set_page_config(
    page_title="Brand Asset Management Tools​",
    page_icon="🔧",
    layout="wide"
)


def main():
    st.title("🔧 Brand Assets Tools")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "Extract Images from Excel",
        "Name assets by brand",
        "Canvas white to Transparent",
        "Asset Overview",
        "Name stims by block + create output files​",
        "Resize",
        "Place on Canvas"

    ])

    with tab1:
         excel_extractor_tool()
    with tab2:
         file_renamer_tool()  
    with tab3:
        white_to_transparent_tool()  
    with tab4:
         pdf_generator_tool() 
    with tab5:
         brand_renamer_tool() 
    with tab6:
        resize_with_transparent_canvas_tool()
    with tab7:
        center_on_canvas_tool()

if __name__ == "__main__":
    main()
