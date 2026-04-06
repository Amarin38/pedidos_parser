import streamlit as st
from excel_logic import Pedidos
from constants import SepararPorEnum

from db import DBBase, db_engine

def main():
    pedidos = Pedidos()

    st.set_page_config(page_title="Separar Pedidos", page_icon="📑")
    st.title("Separar Pedidos")

    if "zip_final" not in st.session_state:
        st.session_state.zip_final = None

    uploaded_files = st.file_uploader("Inserta los archivos", accept_multiple_files=True, type="txt")

    if uploaded_files is not None and uploaded_files != []:
        with st.container(width=500):
            aux1, centro, aux2 = st.columns([0.85,1,0.5])
            
            with centro:
                separar_por = st.selectbox(
                    label="Separar por:", 
                    options=list(SepararPorEnum), 
                    index=None, 
                    placeholder="-----"
                )
                
                if st.button(
                    label="Separar pedidos", 
                    type="primary", 
                    width=200, 
                    disabled=separar_por is None,
                ):
                    with st.spinner("Separando pedidos..."):
                        st.session_state.zip_final = pedidos.ejecutar_todo(uploaded_files, separar_por) # type: ignore



                if st.session_state.zip_final is not None:
                    st.success("Archivo generado.")
                    st.download_button(
                        label="Descargar pedidos.",
                        data=st.session_state.zip_final,
                        file_name=f"Pedidos separados {separar_por}.zip",
                        mime="application/zip",
                        on_click=lambda: st.session_state.update({"zip_final": None})
                    )
                

if __name__ == '__main__':
    main()
    DBBase.metadata.create_all(db_engine)