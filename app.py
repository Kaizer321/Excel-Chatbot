import streamlit as st
import os
import time
from functions import *



def main():
    if 'theme' not in st.session_state:
        st.session_state.theme = 'light'

    st.set_page_config(page_title="AI Assistant", page_icon="🤖", layout="wide")

    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');

    :root {{
        --bg-color: {('#1a1a1a' if st.session_state.theme == 'dark' else '#f0f4f8')};
        --text-color: {('#f0f4f8' if st.session_state.theme == 'dark' else '#1a1a1a')};
        --sidebar-bg: {('#2c3e50' if st.session_state.theme == 'dark' else '#b2bfd9')};
        --sidebar-text: {('#f0f4f8' if st.session_state.theme == 'dark' else '#2c3e50')};
        --sidebar-hover: {('#34495e' if st.session_state.theme == 'dark' else '#e0e8f9')};
    }}

    body {{
        font-family: 'Roboto', sans-serif;
        background-color: var(--bg-color);
        color: var(--text-color);
        transition: all 0.3s ease;
    }}

    .main {{
        background: linear-gradient(135deg, var(--bg-color) 0%, var(--sidebar-bg) 100%);
        display: flex;
        flex-direction: column;
        align-items: center;
        padding: 20px;
        padding-bottom: 80px;
    }}

    .stApp {{
        max-width: 100%;
        margin: 0;
    }}

    .title {{
        color: var(--text-color);
        text-align: center;
        font-size: 2.5em;
        font-weight: 700;
        margin-bottom: 0.2em;
        padding-top: 1em;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }}

    .subtitle {{
        color: var(--text-color);
        text-align: center;
        font-size: 1.2em;
        font-weight: 300;
        margin-bottom: 2em;
    }}

    .chat-container {{
        background-color: transparent;
        padding: 20px;
        margin-bottom: 20px;
        max-height: calc(100vh - 300px);
        overflow-y: auto;
        width: 100%;
        max-width: 800px;
    }}

    .message {{
        padding: 10px 15px;
        border-radius: 20px;
        margin-bottom: 15px;
        max-width: 85%;
        line-height: 1.5;
        font-size: 0.9em;
        box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        position: relative;
    }}

    .user-message {{
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        margin-left: auto;
        border-bottom-right-radius: 0;
    }}

    .assistant-message {{
        background: linear-gradient(135deg, var(--sidebar-bg) 0%, var(--sidebar-hover) 100%);
        color: var(--text-color);
        margin-right: auto;
        border-bottom-left-radius: 0;
    }}

    .message-container {{
        display: flex;
        margin-bottom: 20px;
        align-items: flex-end;
    }}

    .user-container {{
        justify-content: flex-end;
    }}

    .assistant-container {{
        justify-content: flex-start;
    }}

    .icon {{
        width: 30px;
        height: 30px;
        border-radius: 50%;
        margin: 0 10px;
        box-shadow: 0 3px 10px rgba(0,0,0,0.2);
    }}

    .fixed-input {{
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: var(--bg-color);
        border-top: 1px solid var(--sidebar-bg);
        padding: 15px;
        box-shadow: 0 -2px 10px rgba(0, 0, 0, 0.1);
        z-index: 1000;
        display: flex;
        justify-content: center;
    }}

    .fixed-input > div {{
        width: 100%;
        max-width: 800px;
    }}

    .stChatInput {{
        position: fixed;
        bottom: 2%;
        left: 50%;
        transform: translateX(-40%);
        width: auto;
        background-color: transparent !important;
        box-sizing: border-box;
    }}

    .stFileUploader > div > div {{
        width: 100%;
    }}
                
    .stButton > button, .sidebar-button {{
        border-radius: 30px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 10px 20px;
        font-size: 0.9em;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }}

    .stButton > button:hover, .sidebar-button:hover {{
        transform: translateY(-3px);
        box-shadow: 0 8px 20px rgba(0,0,0,0.3);
    }}

    .sidebar-button {{
        width: 80%;
        margin: 20px auto;
        display: block;
    }}

    .upload-section {{
        padding: 20px;
        margin-bottom: 30px;
        width: 100%;
        max-width: 800px;
    }}

    .stAlert {{
        border-radius: 15px;
    }}

    .stSpinner > div {{
        border-top-color: #764ba2 !important;
    }}

    [data-testid="stSidebar"] {{
        width: 250px;
        background-color: var(--sidebar-bg);
        box-shadow: 2px 0 10px rgba(0,0,0,0.1);
        display: flex;
        flex-direction: column;
        align-items: center;
        padding: 20px;
        box-sizing: border-box;
        height: 100vh;
        position: fixed;
        top: 0;
        left: 0;
        z-index: 1000;
    }}

    [data-testid="stSidebar"] .sidebar-content {{
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: flex-start;
        height: 100%;
        text-align: center;
        padding-top: 20px;
        color: var(--sidebar-text);
    }}

    [data-testid="stSidebar"] .sidebar-content img {{
        width: 50px;
        height: 50px;
        border-radius: 50%;
        margin-right: 10px;
    }}

    [data-testid="stSidebar"] .sidebar-content h1 {{
        color: var(--sidebar-text);
        font-size: 1.5em;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        margin: 0;
    }}

    .input-container {{
        display: flex;
        align-items: center;
        gap: 10px;
    }}

    .input-container .stTextInput {{
        flex-grow: 1;
    }}

    .theme-switch {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 10px;
        background-color: var(--bg-color);
        border-radius: 20px;
        margin-top: 20px;
    }}

    .theme-switch button {{
        background-color: transparent;
        border: none;
        cursor: pointer;
        font-size: 20px;
    }}

    @media (max-width: 768px) {{
        .main {{
            padding-left: 0;
        }}
        .fixed-input {{
            padding: 10px;
        }}
        .stApp {{
            max-width: 100%;
        }}
        .title {{
            font-size: 2em;
            text-align: center;
        }}
        .subtitle {{
            font-size: 1em;
        }}
        .message {{
            font-size: 0.8em;
            max-width: 90%;
        }}
        .input-container .stTextInput > div > div > input {{
            padding: 8px 15px;
        }}
        .input-container .stButton > button {{
            padding: 8px 15px;
        }}
        .stTextInput > div > div > input {{
            padding: 8px 15px;
        }}
        .upload-section, .chat-container {{
            width: 95%;
        }}
        .stAlert {{
            font-size: 0.8em;
            max-width: 95%;
        }}
        .icon {{
            width: 25px;
            height: 25px;
        }}
        .chat-container, .upload-section {{
            width: 95%;
        }} 
        [data-testid="stSidebar"] {{
            position: fixed;
            left: -100%;
            top: 0;
            height: 100vh;
            z-index: 1001;
            transition: left 0.3s ease-in-out;
        }}
        [data-testid="stSidebar"].visible {{
            left: 0;
        }}
        .sidebar-toggle {{
            position: fixed;
            top: 10px;
            left: 10px;
            z-index: 1002;
            background-color: var(--sidebar-bg);
            border: none;
            border-radius: 5px;
            padding: 5px 10px;
            cursor: pointer;
        }}
    }}
    </style>

    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown('<div class="sidebar-content">', unsafe_allow_html=True)
        st.markdown('''
        <div style="display: flex; align-items: center; justify-content: center;">
            <img src="https://cdn-icons-png.flaticon.com/128/6231/6231457.png" width="50" style="margin-right: 10px;">
            <h1 style="margin: 0; color: var(--sidebar-text);">Sonic</h1>
        </div>
        ''', unsafe_allow_html=True)
        
        if st.button("New Chat", key="new_chat_button", use_container_width=True, type="primary"):
            start_new_chat()
        
        if st.button("Create Graph", key="create_graph_button", use_container_width=True, type="primary"):
            st.session_state.show_graph_options = True
        
        st.markdown(f'<span style="color: var(--sidebar-text);">Theme: {st.session_state.theme.capitalize()}</span>', unsafe_allow_html=True)
        if st.button("🌓 Dark" if st.session_state.theme == 'light' else "🌞 Light", key="theme_toggle", type="primary"):
            toggle_theme()
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<h1 class='title'>✨ AI Chat Companion ✨</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Explore your Excel documents with our intelligent assistant</p>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,6,1])
    with col2:
        st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
        excel_files = st.file_uploader("📊 Upload your Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if excel_files and 'assistant_initialized' not in st.session_state:
            file_locations = []
            for excel_file in excel_files:
                location = convert_excel_to_docx(excel_file)
                file_locations.append(location)
                st.success(f'📚 File {excel_file.name} has been uploaded and converted successfully.')

                # Store Excel data in session state
                st.session_state.excel_data = extract_data_from_excel(excel_file)

            file_ids = [saveFileOpenAI(location) for location in file_locations]
            assistant_id, vector_id = createAssistant(file_ids, "Excel Data Assistant")

            for location in file_locations:
                try:
                    os.remove(location)
                except PermissionError:
                    st.warning(f"⚠️ Could not delete temporary file: {location}. It will be deleted when the app is closed.")

            st.session_state.assistant_id = assistant_id
            st.session_state.vector_id = vector_id
            st.session_state.assistant_initialized = True
        st.markdown("</div>", unsafe_allow_html=True)

    if 'show_graph_options' in st.session_state and st.session_state.show_graph_options:
        st.markdown("<h2>Create Graph</h2>", unsafe_allow_html=True)
        create_graph()

    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []

    col1, col2, col3 = st.columns([1, 6, 1])
    with col2:
        st.markdown("<div class='chat-container'>", unsafe_allow_html=True)
        for message in st.session_state.chat_history:
            if message['role'] == 'user':
                st.markdown(f"""
                <div class='message-container user-container'>
                    <div class='message user-message'>
                        {message['content']}
                    </div>
                    <img class='icon' src='https://cdn-icons-png.flaticon.com/128/16683/16683419.png' alt='User Icon'>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class='message-container assistant-container'>
                    <img class='icon' src='https://cdn-icons-png.flaticon.com/128/8943/8943377.png' alt='Assistant Icon'>
                    <div class='message assistant-message'>
                        {message['content']}
                    </div>
                </div>
                """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 6, 1])
    with col2:
        st.markdown("<div class='input-container'>", unsafe_allow_html=True)
        user_input = st.chat_input("💬 Ask your question:", key="chat_input", max_chars=2000)
        st.markdown("</div>", unsafe_allow_html=True)
        
        if user_input:
            handle_input(user_input)
    

if __name__ == "__main__":
    main()