import streamlit as st
import pandas as pd
from excel_logic import Pedidos, Codigos
from constants import SepararPorEnum


def main():
    pedidos = Pedidos()
    codigos = Codigos()

    st.set_page_config(page_title="Separar Pedidos", page_icon="📑")
    st.title("Separar Pedidos")

    st.button(
                label="Recargar códigos ⟳",
                type="secondary",
                on_click=codigos.ejecutar_todo
            )
    
    uploaded_files = st.file_uploader("Inserta los archivos", accept_multiple_files=True, type="txt")

    if uploaded_files is not None and uploaded_files != []:
        with st.container(width=500):
            aux1, centro, aux2 = st.columns([0.85,1,0.5])
            
            with centro:
                separar_por = st.selectbox("Separar por:", options=list(SepararPorEnum), index=None, placeholder="-----")
                
                st.button(
                    label="Separar", 
                    type="primary", 
                    width=200, 
                    disabled=separar_por is None,
                    on_click=pedidos.ejecutar_todo,
                    args=(uploaded_files, separar_por,)
                    )
                

if __name__ == '__main__':
    main()