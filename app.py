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
    order_map = {} 
    try:
        if db_name == "Reactive" and os.path.exists("dye_list.xlsx"):
            df = pd.read_excel("dye_list.xlsx", header=None)
            for idx, row in df.iterrows():
                if len(row) > 2 and pd.notna(row[1]) and pd.notna(row[2]):
                    xml_name = str(row[1]).strip()
                    orig_name = str(row[2]).strip()
                    mapping[xml_name] = orig_name
                    if orig_name not in order_map:
                        order_map[orig_name] = idx

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

# 💡 Datacolor QTX 포맷 파일 생성 함수
def generate_qtx_files(selected_pids, root, dye_mapping):
    qtx_dict = {} 
    
    # 1. 샘플 데이터(스펙트럼) 스캔
    samples_data = {}
    for sample in root.iter('Sample'):
        s_id_node = sample.find('SAMPLEID')
        s_name_node = sample.find('NAME')
        if s_id_node is not None and s_name_node is not None:
            s_id = s_id_node.text.strip()
            s_name = s_name_node.text.strip()
            
            spectrum_dict = {}
            for spec in sample.iter('Spectrum'):
                wl = spec.find('WAVELENGTH')
                val = spec.find('SPECTRUMVALUE')
                if wl is not None and val is not None:
                    # XML의 반사율 값(0~1)을 QTX 표준 백분율(0~100)로 변환
                    spectrum_dict[int(wl.text.strip())] = float(val.text.strip()) * 100
            
            samples_data[s_id] = {'name': s_name, 'spectrum': spectrum_dict}

    # 2. 선택된 염료 각각에 대해 별도의 QTX 내용 작성
    for pid in selected_pids:
        qtx_lines = ["[VERSION]", "QTX=QTX 1.0", "", "[DATAMETRIC]", "DATAMETRIC=1", ""]
        
        calib_node = None
        for calib in root.iter('Calibration'):
            pid_node = calib.find('PRODUCT_ID')
            if pid_node is not None and pid_node.text.strip() == pid:
                calib_node = calib
                break
                
        if calib_node is not None:
            series_list = []
            for serie in calib_node.iter('CalibrationSerie'):
                sample_id_node = serie.find('SAMPLEID')
                if sample_id_node is not None:
                    s_id = sample_id_node.text.strip()
                    if s_id in samples_data:
                        series_list.append(samples_data[s_id])
            
            if not series_list:
                continue
                
            # 측정 파장 범위 및 포인트 개수 자동 계산
            first_sample_spec = series_list[0]['spectrum']
            if not first_sample_spec:
                continue
                
            wls = sorted(first_sample_spec.keys())
            min_wl = min(wls)
            max_wl = max(wls)
            num_pts = ((max_wl - min_wl) // 10) + 1
            
            # 첫 번째 샘플을 Standard로 지정
            std_sample = series_list[0]
            std_name = std_sample['name']
            
            qtx_lines.append("[STANDARD_DATA 0]")
            qtx_lines.append(f"STD_NAME={std_name}")
            qtx_lines.append("STD_DATETIME=0")
            qtx_lines.append(f"STD_REFLPOINTS={num_pts}")
            qtx_lines.append("STD_REFLINTERVAL=10")
            qtx_lines.append(f"STD_REFLLOW={min_wl}")
            qtx_lines.append("STD_VIEWING=SCI")
            
            std_r_vals = [f"{std_sample['spectrum'].get(wl, 0.0):.6f}" for wl in range(min_wl, max_wl + 10, 10)]
            qtx_lines.append(f"STD_R= {', '.join(std_r_vals)}")
            qtx_lines.append("")
            
            # 나머지 샘플들을 Batch로 지정
            for i, bat_sample in enumerate(series_list[1:]):
                qtx_lines.append(f"[BATCH_DATA {i}]")
                qtx_lines.append(f"STD_NAME={std_name}")
                qtx_lines.append(f"BAT_NAME={bat_sample['name']}")
                qtx_lines.append("BAT_DATETIME=0")
                qtx_lines.append("BAT_PF_JUDGE=1")
                qtx_lines.append(f"BAT_REFLPOINTS={num_pts}")
                qtx_lines.append("BAT_REFLINTERVAL=10")
                qtx_lines.append(f"BAT_REFLLOW={min_wl}")
                qtx_lines.append("BAT_VIEWING=SCI")
                
                bat_r_vals = [f"{bat_sample['spectrum'].get(wl, 0.0):.6f}" for wl in range(min_wl, max_wl + 10, 10)]
                qtx_lines.append(f"BAT_R= {', '.join(bat_r_vals)}")
                qtx_lines.append("")
            
            # 파일명에 쓸 수 없는 특수문자 제거 후 딕셔너리에 저장
            display_name = dye_mapping.get(pid, pid)
            safe_filename = "".join(c for c in display_name if c not in r'\/:*?"<>|')
            qtx_dict[f"{safe_filename}.qtx"] = "\n".join(qtx_lines)
            
    return qtx_dict

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
        excel_name_map, excel_order_map = load_excel_mapping(selected_db_name)

        with zipfile.ZipFile(ZIP_FILE_PATH, 'r') as z:
            xml_filename = z.namelist()[0]
            with z.open(xml_filename) as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

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
            
            filtered_pids = [pid for pid, display_name in dye_mapping.items() if search_query.lower() in display_name.lower()]
            filtered_pids.sort(key=lambda pid: (excel_order_map.get(dye_mapping[pid], float('inf')), dye_mapping[pid]))
            
            st.sidebar.markdown(f"**검색 결과: {len(filtered_pids)}개**")

            def select_all_filtered(pids):
                for pid in pids:
                    st.session_state.selected_dyes.add(pid)
                    st.session_state[f"chk_{pid}"] = True

            def deselect_all_filtered(pids):
                for pid in pids:
                    st.session_state.selected_dyes.discard(pid)
                    st.session_state[f"chk_{pid}"] = False

            def toggle_dye(pid):
                if st.session_state.get(f"chk_{pid}", False):
                    st.session_state.selected_dyes.add(pid)
                else:
                    st.session_state.selected_dyes.discard(pid)

            col1, col2 = st.sidebar.columns(2)
            col1.button("✅ 전체 선택", on_click=select_all_filtered, args=(filtered_pids,))
            col2.button("❌ 전체 해제", on_click=deselect_all_filtered, args=(filtered_pids,))

            st.sidebar.markdown("---")
            
            st.sidebar.write(f"👇 **[{selected_db_name}] 염료 목록 (엑셀 순서 정렬)**")
            for pid in filtered_pids:
                display_name = dye_mapping[pid]
                
                if f"chk_{pid}" not in st.session_state:
                    st.session_state[f"chk_{pid}"] = (pid in st.session_state.selected_dyes)

                st.sidebar.checkbox(
                    display_name, 
                    key=f"chk_{pid}",
                    on_change=toggle_dye,
                    args=(pid,)
                )

            # ==========================================
            # ➡️ 메인 화면 영역
            # ==========================================
            st.success(f"**{selected_db_name}** 데이터베이스 로드 완료! 전체 {len(dye_mapping)}개의 염료 중 현재 **{len(st.session_state.selected_dyes)}**개를 선택했습니다.")
            
            if st.session_state.selected_dyes:
                with st.expander("📌 현재 선택된 염료 목록 보기 (클릭하여 펼치기)", expanded=True):
                    sorted_selected = sorted(
                        list(st.session_state.selected_dyes), 
                        key=lambda x: (excel_order_map.get(dye_mapping[x], float('inf')), dye_mapping[x])
                    )
                    
                    for selected_pid in sorted_selected:
                        st.write(f"- **{dye_mapping[selected_pid]}** <span style='color:gray; font-size:0.8em;'>(내부 XML명: {selected_pid})</span>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            col_btn1, col_btn2 = st.columns(2)

            # XML 다운로드 로직
            with col_btn1:
                if st.button("🚀 XML 파일로 추출하기", use_container_width=True):
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

                        st.download_button(
                            label=f"📥 Filtered_{selected_db_name}.xml 다운로드",
                            data=xml_buffer,
                            file_name=f"Filtered_{selected_db_name}.xml",
                            mime="application/xml",
                            type="primary",
                            use_container_width=True
                        )

            # QTX 다중 파일(ZIP) 다운로드 로직
            with col_btn2:
                if st.button("📊 QTX 파일로 추출하기", use_container_width=True):
                    if not st.session_state.selected_dyes:
                        st.warning("왼쪽 사이드바에서 먼저 하나 이상의 염료를 체크해주세요.")
                    else:
                        with st.spinner("개별 QTX 파일을 생성 및 압축 중입니다..."):
                            qtx_files_dict = generate_qtx_files(st.session_state.selected_dyes, root, dye_mapping)
                            
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for filename, content in qtx_files_dict.items():
                                    zf.writestr(filename, content)
                            
                            zip_buffer.seek(0)
                        
                        st.download_button(
                            label=f"📥 {selected_db_name}_QTX.zip 다운로드",
                            data=zip_buffer,
                            file_name=f"{selected_db_name}_QTX.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True
                        )

    except zipfile.BadZipFile:
        st.error("ZIP 파일이 손상되었거나 올바른 압축 파일이 아닙니다.")
    except Exception as e:
        st.error(f"데이터베이스를 처리하는 중 오류가 발생했습니다: {e}")
else:
    st.error(f"⚠️ 서버에서 '{ZIP_FILE_PATH}' 파일을 찾을 수 없습니다. GitHub 저장소에 ZIP 파일이 업로드되었는지 확인해주세요.")