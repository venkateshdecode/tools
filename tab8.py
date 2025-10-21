import streamlit as st

def render():
    st.title("Tab 8 - Button Demo")

    # Create a button
    if st.button("Click me"):
        st.write("TEST")
