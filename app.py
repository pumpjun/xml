import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import copy
import os
import zipfile

st.set_page_config(page_title="염료 데이터 추출기", page_icon="🎨", layout="wide")

def clear_selection():
    st.session_state.selected_dyes = set()

if 'selected_dyes' not in st.session_state:
    st.session_state.selected_dyes = set()

st.title("🎨 단색 염료 데이터 추출기")
st.write("👈 **왼쪽 사이드바**에서 사용할 DB를 선택하고 염료를 검색하여 추출하세요.")

# 1. 깃허브에 업로드할 ZIP 파일 목록
AVAILABLE_FILES = {
        "Disperse Interlock": "Disperse Interlock.zip",
        "Disperse Woven": "Disperse Woven.zip",
        "Reactive": "Reactive.zip"
}

# 엑셀 매핑 및 순서 불러오기 함수
@st.cache_data
def load_excel_mapping(db_name):
    mapping = {}
    order_map = {} # 엑셀 행 순서를 기억할 딕셔너리 추가
    try:
        if db_name == "Reactive" and os.path.exists("dye_list.xlsx"):
            df = pd.read_excel("dye_list.xlsx", header=None)
            for idx, row in df.iterrows():
                if len(row) > 2 and pd.notna(row[1]) and pd.notna(row[2]):
                    xml_name = str(row[1]).strip()
                    orig_name = str(row[2]).strip()
                    mapping[xml_name] = orig_name
                    if orig_name not in order_map:
                        order_map[orig_name] = idx # 행 번호(idx)를 순서 값으로 저장

        elif db_name == "Disperse Interlock" and os.path.exists("dis_dye_list.xlsx"):
            df = pd.read_excel("dis_dye_list.xlsx", header=None)
            for idx, row in df.iterrows():
                if len(row) > 3 and pd.notna(row[1]) and pd.notna(row[3]):
                    xml_name = str(row[1]).strip()
                    orig_name = str(row[3]).strip()
                    mapping[xml_name] = orig_name
                    if orig_name not in order_map:
                        order_map[orig_name] = idx

        elif db_name == "Disperse Woven" and os.path.exists("dis_dye_list.xlsx"):
            df = pd.read_excel("dis_dye_list.xlsx", header=None)
            for idx, row in df.iterrows():
                if len(row) > 3 and pd.notna(row[2]) and pd.notna(row[3]):
                    xml_name = str(row[2]).strip()
                    orig_name = str(row[3]).strip()
                    mapping[xml_name] = orig_name
                    if orig_name not in order_map:
                        order_map[orig_name] = idx
    except Exception as e:
        st.sidebar.warning(f"⚠️ 엑셀 매핑 파일 읽기 오류: {e}")
        
    return mapping, order_map

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
        # 엑셀 매핑 데이터와 '순서 데이터(order_map)'를 함께 불러옴
        excel_name_map, excel_order_map = load_excel_mapping(selected_db_name)

        # 2. ZIP 파일 읽기 및 압축 해제
        with zipfile.ZipFile(ZIP_FILE_PATH, 'r') as z:
            xml_filename = z.namelist()[0]
            with z.open(xml_filename) as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

        # 3. 염료 매핑 딕셔너리 생성 (PRODUCT_ID -> 표시할 오리지널 이름)
        dye_mapping = {}
        
        for elem in root.iter('Product'):
            pid_node = elem.find('PRODUCT_ID')
            pname_node = elem.find('PRODUCT_NAME')
            if pid_node is not None and pid_node.text:
                pid = pid_node.text.strip()
                pname = pname_node.text.strip() if pname_node is not None and pname_node.text else pid
                if pid != 'H2O':
                    display_name = excel_name_map.get(pname, excel_name_map.get(pid, pname))
                    dye_mapping[pid] = display_name

        # Product 태그에는 없지만 다른 곳에 있는 염료 스캔
        for elem in root.iter():
            if elem.tag in ['Product', 'Dyestuff', 'Calibration']:
                pid_node = elem.find('PRODUCT_ID')
                if pid_node is not None and pid_node.text:
                    pid = pid_node.text.strip()
                    if pid != 'H2O' and pid not in dye_mapping:
                        display_name = excel_name_map.get(pid, pid)
                        dye_mapping[pid] = display_name 

        if not dye_mapping:
            st.error("염료를 찾을 수 없습니다. XML 파일 구조를 확인해주세요.")
        else:
            st.sidebar.header("🔍 염료 검색 및 선택")
            search_query = st.sidebar.text_input("오리지널 염료명 검색 (예: APEX, ECO 등)", "")
            
            # 검색어로 1차 필터링
            filtered_pids = [pid for pid, display_name in dye_mapping.items() if search_query.lower() in display_name.lower()]
            
            # 💡 핵심 로직: 엑셀 행 순서(excel_order_map)에 맞게 정렬 (엑셀에 없는 항목은 맨 뒤로 배치)
            filtered_pids.sort(key=lambda pid: (excel_order_map.get(dye_mapping[pid], float('inf')), dye_mapping[pid]))
            
            st.sidebar.markdown(f"**검색 결과: {len(filtered_pids)}개**")

            col1, col2 = st.sidebar.columns(2)
            if col1.button("✅ 전체 선택"):
                for pid in filtered_pids:
                    st.session_state.selected_dyes.add(pid)
            if col2.button("❌ 전체 해제"):
                for pid in filtered_pids:
                    st.session_state.selected_dyes.discard(pid)

            st.sidebar.markdown("---")
            
            st.sidebar.write(f"👇 **[{selected_db_name}] 염료 목록 (엑셀 순서 정렬)**")
            for pid in filtered_pids:
                display_name = dye_mapping[pid]
                
                is_checked = st.sidebar.checkbox(
                    display_name, 
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
                    # 💡 메인 화면의 선택된 리스트도 엑셀 순서에 맞게 정렬하여 표시
                    sorted_selected = sorted(
                        list(st.session_state.selected_dyes), 
                        key=lambda x: (excel_order_map.get(dye_mapping[x], float('inf')), dye_mapping[x])
                    )
                    
                    for selected_pid in sorted_selected:
                        st.write(f"- **{dye_mapping[selected_pid]}** <span style='color:gray; font-size:0.8em;'>(내부 XML명: {selected_pid})</span>", unsafe_allow_html=True)
            
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
    st.error(f"⚠️ 서버에서 '{ZIP_FILE_PATH}' 파일을 찾을 수 없습니다. GitHub 저장소에 ZIP 파일이 업로드되었는지 확인해주세요.")