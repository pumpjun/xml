import streamlit as st
import xml.etree.ElementTree as ET
import io
import copy
import os
import zipfile  # 압축 해제를 위한 파이썬 기본 라이브러리 추가

st.set_page_config(page_title="염료 데이터 추출기", page_icon="🎨", layout="wide")

def clear_selection():
    st.session_state.selected_dyes = set()

if 'selected_dyes' not in st.session_state:
    st.session_state.selected_dyes = set()

st.title("🎨 단색 염료 데이터 추출기")
st.write("👈 **왼쪽 사이드바**에서 사용할 DB를 선택하고 염료를 검색하여 추출하세요.")

# 1. 깃허브에 업로드할 ZIP 파일 목록으로 변경
AVAILABLE_FILES = {
        "Disperse Interlock": "Disperse Interlock.zip",
        "Disperse Woven": "Disperse Woven.zip",
        "Reactive": "Reactive.zip"
}

# ==========================================
# ⬅️ 왼쪽 사이드바 영역
# ==========================================
st.sidebar.header("📂 데이터베이스 선택")
selected_db_name = st.sidebar.selectbox(
    "사용할 염료 데이터베이스를 선택하세요:",
    options=list(AVAILABLE_FILES.keys()),
    on_change=clear_selection
)

ZIP_FILE_PATH = AVAILABLE_FILES[selected_db_name]
st.sidebar.markdown("---")

if os.path.exists(ZIP_FILE_PATH):
    try:
        # 2. ZIP 파일 읽기 및 압축 해제 (메모리 상에서 처리)
        with zipfile.ZipFile(ZIP_FILE_PATH, 'r') as z:
            # ZIP 파일 안에 있는 첫 번째 파일(XML)의 이름을 가져옵니다.
            xml_filename = z.namelist()[0]
            
            # 해당 XML 파일을 열어서 파싱합니다.
            with z.open(xml_filename) as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

        # 3. 염료 매핑 딕셔너리 생성 (PRODUCT_ID -> PRODUCT_NAME)
        dye_mapping = {}
        
        for elem in root.iter('Product'):
            pid_node = elem.find('PRODUCT_ID')
            pname_node = elem.find('PRODUCT_NAME')
            if pid_node is not None and pid_node.text:
                pid = pid_node.text.strip()
                pname = pname_node.text.strip() if pname_node is not None and pname_node.text else pid
                if pid != 'H2O':
                    dye_mapping[pid] = pname

        for elem in root.iter():
            if elem.tag in ['Product', 'Dyestuff', 'Calibration']:
                pid_node = elem.find('PRODUCT_ID')
                if pid_node is not None and pid_node.text:
                    pid = pid_node.text.strip()
                    if pid != 'H2O' and pid not in dye_mapping:
                        dye_mapping[pid] = pid 

        if not dye_mapping:
            st.error("염료를 찾을 수 없습니다. XML 파일 구조를 확인해주세요.")
        else:
            st.sidebar.header("🔍 염료 검색 및 선택")
            
            search_query = st.sidebar.text_input("Full Name 검색어 입력 (예: APEX, ECO 등)", "")
            
            filtered_pids = [pid for pid, pname in dye_mapping.items() if search_query.lower() in pname.lower()]
            
            st.sidebar.markdown(f"**검색 결과: {len(filtered_pids)}개**")

            col1, col2 = st.sidebar.columns(2)
            if col1.button("✅ 전체 선택"):
                for pid in filtered_pids:
                    st.session_state.selected_dyes.add(pid)
            if col2.button("❌ 전체 해제"):
                for pid in filtered_pids:
                    st.session_state.selected_dyes.discard(pid)

            st.sidebar.markdown("---")
            
            st.sidebar.write(f"👇 **[{selected_db_name}] 염료 목록**")
            for pid in filtered_pids:
                pname = dye_mapping[pid]
                
                is_checked = st.sidebar.checkbox(
                    pname, 
                    value=(pid in st.session_state.selected_dyes), 
                    key=f"chk_{pid}"
                )
                
                if is_checked:
                    st.session_state.selected_dyes.add(pid)
                else:
                    st.session_state.selected_dyes.discard(pid)

            # ==========================================
            # ➡️ 메인 화면 영역
            # ==========================================
            st.success(f"**{selected_db_name}** 데이터베이스 로드 완료! 전체 {len(dye_mapping)}개의 염료 중 현재 **{len(st.session_state.selected_dyes)}**개를 선택했습니다.")
            
            if st.session_state.selected_dyes:
                with st.expander("📌 현재 선택된 염료 목록 보기 (클릭하여 펼치기)", expanded=True):
                    sorted_selected = sorted(list(st.session_state.selected_dyes), key=lambda x: dye_mapping[x])
                    for selected_pid in sorted_selected:
                        st.write(f"- **{dye_mapping[selected_pid]}** <span style='color:gray; font-size:0.8em;'>(내부 ID: {selected_pid})</span>", unsafe_allow_html=True)
            
            st.markdown("---")

            if st.button("🚀 선택한 염료로 XML 추출하기", use_container_width=True):
                if not st.session_state.selected_dyes:
                    st.warning("왼쪽 사이드바에서 먼저 하나 이상의 염료를 체크해주세요.")
                else:
                    with st.spinner("XML 파일을 생성 중입니다..."):
                        new_root = copy.deepcopy(root)
                        elements_to_remove = []

                        for parent in new_root.iter():
                            for child in list(parent):
                                if child.tag in ['Product', 'Dyestuff', 'Calibration']:
                                    pid_node = child.find('PRODUCT_ID')
                                    if pid_node is not None and pid_node.text:
                                        pid = pid_node.text.strip()
                                        
                                        if pid not in st.session_state.selected_dyes and pid != 'H2O':
                                            elements_to_remove.append((parent, child))

                        for parent, child in elements_to_remove:
                            if child in parent:
                                parent.remove(child)

                        new_tree = ET.ElementTree(new_root)
                        xml_buffer = io.BytesIO()
                        new_tree.write(xml_buffer, encoding='ISO-8859-1', xml_declaration=True)
                        xml_buffer.seek(0)

                    export_filename = f"Filtered_{selected_db_name}.xml"

                    st.download_button(
                        label=f"📥 {export_filename} 다운로드",
                        data=xml_buffer,
                        file_name=export_filename,
                        mime="application/xml",
                        type="primary"
                    )

    except zipfile.BadZipFile:
        st.error("ZIP 파일이 손상되었거나 올바른 압축 파일이 아닙니다.")
    except Exception as e:
        st.error(f"데이터베이스를 처리하는 중 오류가 발생했습니다: {e}")
else:
    st.error(f"⚠️ 서버에서 '{ZIP_FILE_PATH}' 파일을 찾을 수 없습니다. GitHub 저장소에 ZIP 파일이 정상적으로 업로드되었는지 확인해주세요.")