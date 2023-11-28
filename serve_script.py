import streamlit as st

from the_excel_servant.main import try_with_xlwings

st.title("serve from excel demo")

number_1 = st.number_input("number1")
number_2 = st.number_input("number2")


the_result = try_with_xlwings(number_1, number_2)


st.write("test result is", the_result)
