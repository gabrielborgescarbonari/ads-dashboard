import os
from dotenv import load_dotenv

load_dotenv()


def get(key: str) -> str:
    try:
        import streamlit as st
        val = st.secrets.get(key)
        if val is not None:
            return str(val)
    except Exception:
        pass
    return os.getenv(key, "")
