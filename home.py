import streamlit as st
from datetime import date

import constants as const
from constants import SepararPorEnum
from excel_logic import Pedidos, Codigos
from db import DBBase, db_engine

def main():
    pedidos = Pedidos()
    codigos = Codigos()

    st.set_page_config(page_title=const.PAGE_TITLE, page_icon="📑")
    st.title(const.PAGE_TITLE)

    if "zip_final" not in st.session_state:
        st.session_state.zip_final = None

    st.button(
                label=const.RECARGAR_LABEL,
                type="secondary",
                on_click=codigos.sacar_lista
            )
    
    uploaded_files = st.file_uploader(const.UPLOAD_TITLE, accept_multiple_files=True, type="txt")

    if uploaded_files is not None and uploaded_files != []:
        with st.container(width=const.CONTAINER_WIDTH):
            aux1, centro, aux2 = st.columns(const.MAIN_COLS)
            
            with centro:
                separar_por = st.selectbox(
                    label=const.SELECT_BOX_LABEL, 
                    options=list(SepararPorEnum), 
                    index=None, 
                    placeholder=const.PLACEHOLDER
                )
                
                if st.button(
                    label=const.SEPARAR_LABEL, 
                    type="primary", 
                    width=200, 
                    disabled=separar_por is None,
                ):
                    with st.spinner(const.SEPARAR_SPINNER):
                        st.session_state.zip_final = pedidos.ejecutar_todo(uploaded_files, separar_por) # type: ignore


                if st.session_state.zip_final is not None:
                    st.success(const.SUCCESS_FILE)
                    st.download_button(
                        label=const.DOWNLOAD_BTTN_LABEL,
                        data=st.session_state.zip_final,
                        file_name=f"Pedidos separados {separar_por} {date.today().strftime("%d-%m-%Y")}.zip",
                        mime="application/zip",
                        on_click=lambda: st.session_state.update({"zip_final": None})
                    )
                

if __name__ == '__main__':
    main()
    DBBase.metadata.create_all(db_engine)