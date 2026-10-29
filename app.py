import streamlit as st
import pandas as pd
import json
import io
import re
from openai import OpenAI
from thefuzz import process, fuzz

st.set_page_config(page_title="선물세트 엑셀 자동 매핑기", layout="wide")
st.title("🎁 선물세트 엑셀 파일 기반 자동 매핑 프로그램")
st.write("마스터 엑셀 매핑, 퍼센트/총 price 및 **AE~AP열(RITC 매핑 & s_bundle=0)**까지 한 번에 산출합니다.")

with st.sidebar:
    st.header("⚙️ 설정")
    api_key = st.text_input("OpenAI API Key를 입력하세요", type="password")

def standardize_sheet_columns(df):
    col_map = {}
    for col in df.columns:
        clean_col = str(col).strip().lower().replace(" ", "").replace("_", "")
        if clean_col in ["company", "회사", "제조사", "업체명"]:
            col_map[col] = "Company"
        elif clean_col in ["category", "cat", "카테고리", "분류"]:
            col_map[col] = "Category"
        elif clean_col in ["brands", "brand", "상품명", "품목명", "제품명", "단품명"]:
            col_map[col] = "Brands"
        elif clean_col in ["volume", "vol", "용량", "중량", "규격"]:
            col_map[col] = "Volume"
        elif clean_col in ["price", "개당가격", "단가", "가격", "공급가", "원가", "금액", "unitprice"]:
            col_map[col] = "price"
        elif clean_col in ["v1"]:
            col_map[col] = "v1"
        elif clean_col in ["v2"]:
            col_map[col] = "v2"
        elif clean_col in ["type", "v3label"]:
            col_map[col] = "V3_LABEL"
        else:
            col_map[col] = str(col).strip()
            
    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def clean_concat_val(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s.lower() == "nan":
        return ""
    return s

def normalize_volume(val):
    if pd.isna(val) or val == "":
        return ""
    val_str = str(val).strip().lower()
    return re.sub(r'\s+', '', val_str)

# 용량/중량에서 순수 숫자 값만 추출 (500ml -> 500, 500g -> 500)
def extract_num_volume(val):
    if pd.isna(val) or val == "":
        return ""
    v_str = str(val).strip().lower().replace(" ", "")
    m = re.search(r'(\d+(?:\.\d+)?)', v_str)
    return m.group(1) if m else ""

def safe_parse_json(text):
    text = text.strip()
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
        elif isinstance(data, list):
            return {"items": data}
    except Exception:
        pass

    json_match = re.search(r'(\{.*\})', text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except Exception:
            pass
    raise ValueError(f"JSON 파싱 실패: {text[:100]}...")

def calculate_percentages_and_total_price(df):
    if df.empty or "price" not in df.columns or "수량" not in df.columns:
        df["퍼센트"] = 0
        df["총 price"] = 0
        return df

    df["_num_price"] = pd.to_numeric(df["price"], errors="coerce").fillna(0)
    df["_num_qty"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0)
    df["_amount"] = df["_num_price"] * df["_num_qty"]

    group_keys = df["Barcode"].astype(str) + "___" + df["세트명"].astype(str)
    df["_group"] = group_keys

    result_dfs = []
    for _, group in df.groupby("_group", sort=False):
        total_amount = group["_amount"].sum()
        group["총 price"] = int(round(total_amount))

        if total_amount <= 0:
            group["퍼센트"] = 0
        else:
            exact_pcts = (group["_amount"] / total_amount) * 100
            int_pcts = exact_pcts.round().astype(int)
            diff = 100 - int_pcts.sum()

            if diff != 0 and len(group) > 0:
                remainders = (exact_pcts - int_pcts).abs()
                target_idx = remainders.idxmax()
                int_pcts.loc[target_idx] += diff

            group["퍼센트"] = int_pcts

        result_dfs.append(group)

    final_res = pd.concat(result_dfs, ignore_index=True)
    final_res = final_res.drop(columns=["_num_price", "_num_qty", "_amount", "_group"])
    return final_res

def build_description(row):
    parts = []
    company = str(row.get("Company", "")).strip()
    if company and company.lower() != "nan":
        parts.append(company)

    brand = str(row.get("Brands", "")).strip()
    if brand and brand.lower() != "nan":
        parts.append(brand)

    for i in range(3, 11):
        lbl = str(row.get(f"V{i}_LABEL", "")).strip()
        if lbl and lbl.lower() != "nan":
            parts.append(lbl)

    vol = str(row.get("Volume", "")).strip()
    if vol and vol.lower() != "nan":
        parts.append(vol)

    return " ".join(parts)

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 마스터 엑셀 업로드")
    uploaded_master = st.file_uploader("코드 매핑 마스터 엑셀 (.xlsx)", type=["xlsx"], key="master")
    master_df = None
    ritc_lookup = {}

    if uploaded_master:
        try:
            excel_file = pd.ExcelFile(uploaded_master, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            combined_dfs = []
            
            for sname in sheet_names:
                if sname.strip().upper() == "RITC":
                    df_ritc = pd.read_excel(excel_file, sheet_name=sname, header=None, dtype=str).fillna("")
                    for r_idx, row in df_ritc.iterrows():
                        if r_idx == 0 and ("ref" in str(row[0]).lower() or "item" in str(row[0]).lower()):
                            continue
                        ref_item_code = clean_concat_val(row[0])
                        bm_concat_vals = [clean_concat_val(row[c]) for c in range(1, min(13, len(row)))]
                        concat_key = "".join(bm_concat_vals)
                        if concat_key:
                            ritc_lookup[concat_key] = ref_item_code
                    continue

                df_sheet = pd.read_excel(excel_file, sheet_name=sname, dtype=str)
                if not df_sheet.empty:
                    df_sheet = standardize_sheet_columns(df_sheet)
                    df_sheet["_출처시트"] = sname
                    combined_dfs.append(df_sheet)

            if combined_dfs:
                master_df = pd.concat(combined_dfs, ignore_index=True).fillna("")
                st.success(f"마스터 품목 시트 총 {len(combined_dfs)}개 로드 완료 ({len(master_df)}개 행)")
            if ritc_lookup:
                st.info(f"✅ RITC 시트 확인 완료: 총 {len(ritc_lookup)}개 코드 룩업 테이블 구축됨")

        except Exception as e:
            st.error(f"마스터 엑셀 로드 오류: {e}")

with col2:
    st.subheader("2. 작업 대상 선물세트 엑셀 업로드")
    uploaded_target = st.file_uploader("작업 대상 엑셀 (.xlsx, .xls)", type=["xlsx", "xls"], key="target")
    target_df = None
    if uploaded_target:
        try:
            target_df = pd.read_excel(uploaded_target, header=None, dtype=str).fillna("")
            st.success(f"작업 대상 파일 로드 완료 (총 {len(target_df)}개 행)")
            with st.expander("작업 대상 데이터 미리보기"):
                st.dataframe(target_df.head(5))
        except Exception as e:
            st.error(f"작업 대상 엑셀 로드 오류: {e}")

if st.button("🚀 엑셀 매핑 및 확장 서식(RITC 포함) 일괄 산출 시작", type="primary"):
    if not api_key:
        st.error("OpenAI API Key를 입력해주세요.")
    elif master_df is None or target_df is None:
        st.error("마스터 엑셀과 작업 대상 엑셀을 모두 업로드해주세요.")
    else:
        brand_col = "Brands" if "Brands" in master_df.columns else master_df.columns[0]
        vol_col = "Volume" if "Volume" in master_df.columns else None
        comp_col = "Company" if "Company" in master_df.columns else None

        all_results_list = []
        client = OpenAI(api_key=api_key)
        progress_bar = st.progress(0)
        status_text = st.empty()

        total_rows = len(target_df)

        for idx, row in target_df.iterrows():
            barcode_val = str(row[0]).strip() if len(row) > 0 else ""
            set_name_val = str(row[1]).strip() if len(row) > 1 else ""
            composition_text = str(row[2]).strip() if len(row) > 2 else ""

            if idx == 0 and ("바코드" in barcode_val or "barcode" in barcode_val.lower()):
                continue

            if not composition_text or not set_name_val:
                continue

            status_text.text(f"[{idx+1}/{total_rows}] '{set_name_val}' ({barcode_val}) 분석 및 매핑 중...")

            try:
                prompt = f"""
                세트명: '{set_name_val}'
                구성품 텍스트: '{composition_text}'
                
                위 선물세트 정보에서 구성 단품들을 분석해줘.
                
                [지침]
                1. 세트의 메인 제조사/브랜드 키워드(company_keyword)를 1단어로 파악해줘 (예: "CJ", "백설", "동원", "사조", "청정원", "대상", "오설록", "오뚜기" 등).
                2. 쉼표(`,`)로 구분된 각 구성품의 정확한 명칭(name), 용량/중량(volume), 수량(qty)을 산출해줘.
                   - '카놀라유 500ml' -> volume: "500ml", qty: 1
                   - '올리고당 500ml' -> volume: "500ml", qty: 1
                   - '콩기름 500ml×2' -> volume: "500ml", qty: 2
                   - 'Calvo 참치 65g×3(해바라기유)×4' -> name: "Calvo 참치 해바라기유", volume: "65g", qty: 12
                
                반드시 아래 JSON 형식으로만 응답해:
                {{
                  "company_keyword": "CJ",
                  "items": [
                    {{"name": "백설 카놀라유", "volume": "500ml", "qty": 1}},
                    {{"name": "백설 요리올리고당", "volume": "700g", "qty": 1}}
                  ]
                }}
                """

                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}],
                    response_format={"type": "json_object"},
                    temperature=0.0
                )

                parsed_res = safe_parse_json(response.choices[0].message.content)
                company_kw = parsed_res.get("company_keyword", "").strip()
                extracted_items = parsed_res.get("items", [])

                target_master = master_df
                if comp_col and company_kw:
                    filtered_df = master_df[master_df[comp_col].astype(str).str.contains(company_kw, case=False, na=False)]
                    if not filtered_df.empty:
                        target_master = filtered_df
                    else:
                        filtered_df = master_df[master_df[brand_col].astype(str).str.contains(company_kw, case=False, na=False)]
                        if not filtered_df.empty:
                            target_master = filtered_df

                matched_rows = []
                used_indices_in_set = set()

                for item_no, item in enumerate(extracted_items, 1):
                    item_name = item.get("name", "")
                    item_vol = normalize_volume(item.get("volume", ""))
                    item_vol_num = extract_num_volume(item_vol)  # 순수 숫자 (예: '500')
                    item_qty = item.get("qty", 1)

                    available_df = target_master[~target_master.index.isin(used_indices_in_set)]
                    if available_df.empty:
                        available_df = target_master

                    matched_idx = None

                    # 1순위: 품목명 유사도 사전 필터링 (엉뚱한 제품인 올리브오일 매칭 사전 차단)
                    all_names = available_df[brand_col].astype(str).tolist()
                    name_scores = [fuzz.token_set_ratio(item_name, name) for name in all_names]
                    max_score = max(name_scores) if name_scores else 0
                    
                    # 품목명이 일정 수준 이상 유사한 행들만 1차 후보군으로 압축 (올리고당 vs 올리브유 분리)
                    if max_score >= 60:
                        high_sim_indices = [available_df.index[i] for i, s in enumerate(name_scores) if s >= (max_score - 15)]
                        name_filtered_df = available_df.loc[high_sim_indices]
                    else:
                        name_filtered_df = available_df

                    # 2순위: 용량/중량 매칭 (ml <-> g 상호 호환 적용)
                    if vol_col and item_vol:
                        norm_vols = name_filtered_df[vol_col].apply(normalize_volume)
                        num_vols = name_filtered_df[vol_col].apply(extract_num_volume)

                        # A. 완전 일치 ('500ml' == '500ml')
                        vol_matches = name_filtered_df[norm_vols == item_vol]

                        # B. 단위 호환 일치 ('500ml'와 '500g'의 숫자 '500' 일치)
                        if vol_matches.empty and item_vol_num:
                            vol_matches = name_filtered_df[num_vols == item_vol_num]

                        # C. 앞뒤 숫자가 붙지 않은 정규식 일치 (450ml 제외)
                        if vol_matches.empty:
                            clean_v = re.escape(item_vol)
                            vol_pattern = rf'(?<!\d){clean_v}(?!\d)'
                            vol_matches = name_filtered_df[norm_vols.str.contains(vol_pattern, regex=True, na=False)]

                        if not vol_matches.empty:
                            sub_names = vol_matches[brand_col].tolist()
                            best_name, _ = process.extractOne(item_name, sub_names, scorer=fuzz.token_set_ratio)
                            matched_idx = vol_matches[vol_matches[brand_col] == best_name].index[0]

                    # 3순위: 전체 유사도 매칭 (용량이 없거나 매칭되지 않은 경우)
                    if matched_idx is None:
                        if vol_col:
                            lookup_keys = (
                                available_df[brand_col].astype(str) + " " + available_df[vol_col].apply(normalize_volume)
                            ).tolist()
                            query_key = f"{item_name} {item_vol}".strip()
                        else:
                            lookup_keys = available_df[brand_col].astype(str).tolist()
                            query_key = item_name

                        best_key, _ = process.extractOne(query_key, lookup_keys, scorer=fuzz.token_set_ratio)
                        best_match_idx = lookup_keys.index(best_key)
                        matched_idx = available_df.index[best_match_idx]

                    used_indices_in_set.add(matched_idx)

                    matched_row = master_df.loc[matched_idx].to_dict()
                    matched_row["Barcode"] = barcode_val
                    matched_row["세트명"] = set_name_val
                    matched_row["수량"] = item_qty
                    matched_row["itemno"] = item_no

                    if "price" in matched_row:
                        p_val = str(matched_row["price"]).strip()
                        if p_val.endswith(".0"):
                            matched_row["price"] = p_val[:-2]

                    matched_rows.append(matched_row)

                all_results_list.extend(matched_rows)

            except Exception as e:
                st.error(f"'{set_name_val}' 처리 중 오류: {e}")

            progress_bar.progress((idx + 1) / total_rows)

        if all_results_list:
            raw_result_df = pd.DataFrame(all_results_list)

            # 1. 퍼센트 및 총 price 계산
            mid_df = calculate_percentages_and_total_price(raw_result_df)

            # 2. A열 ~ AD열 기본 컬럼 구성
            base_columns = [
                "Barcode", "세트명", "Category", "Company", "v1", "Brands", "v2", "Volume",
                "V3", "V3_LABEL", "V4", "V4_LABEL", "V5", "V5_LABEL",
                "V6", "V6_LABEL", "V7", "V7_LABEL", "V8", "V8_LABEL",
                "V9", "V9_LABEL", "V10", "V10_LABEL", "U11", "price", "수량", "itemno", "퍼센트", "총 price"
            ]

            for col in base_columns:
                if col not in mid_df.columns:
                    mid_df[col] = ""

            base_df = mid_df[base_columns].copy()

            # 3. RITC 키 계산 함수
            def compute_ritc(row):
                if not ritc_lookup:
                    return ""
                c_val = clean_concat_val(row.get("Category", ""))
                e_val = clean_concat_val(row.get("v1", ""))
                g_val = clean_concat_val(row.get("v2", ""))
                i_val = clean_concat_val(row.get("V3", ""))
                k_val = clean_concat_val(row.get("V4", ""))
                m_val = clean_concat_val(row.get("V5", ""))
                o_val = clean_concat_val(row.get("V6", ""))
                q_val = clean_concat_val(row.get("V7", ""))
                y_val = clean_concat_val(row.get("U11", ""))
                
                concat_key = f"{c_val}{e_val}{g_val}{i_val}{k_val}{m_val}{o_val}{q_val}{y_val}"
                return ritc_lookup.get(concat_key, "")

            # 4. AE열 ~ AP열 확장 컬럼 구성
            ext_df = pd.DataFrame()
            ext_df["Barcode"] = base_df["Barcode"]
            ext_df["Startdate"] = ""
            ext_df["Description"] = mid_df.apply(build_description, axis=1)
            ext_df["Shortdesc"] = base_df["세트명"]
            ext_df["Category"] = base_df["Category"]
            ext_df["RITC"] = base_df.apply(compute_ritc, axis=1)
            ext_df["Comments"] = base_df["총 price"]
            ext_df["EndDate"] = ""
            ext_df["itemno"] = base_df["itemno"]
            ext_df["price_percentage"] = base_df["퍼센트"]
            ext_df["no_in_pack"] = base_df["수량"]
            ext_df["s_bundle"] = 0

            excel_headers = list(base_df.columns) + list(ext_df.columns)
            excel_values = pd.concat([base_df, ext_df], axis=1).values.tolist()

            preview_df = pd.concat([base_df, ext_df], axis=1).copy()
            preview_df.columns = [
                f"{col}_(ERP)" if i >= len(base_columns) and col in ["Barcode", "Category", "itemno"] else col
                for i, col in enumerate(preview_df.columns)
            ]

            status_text.text("✅ 전체 선물세트 매핑 및 RITC 산출이 완료되었습니다!")
            st.success(f"총 {len(excel_values)}개 품목 매핑 완료 (A열 ~ AP열 전체 완성)")
            st.dataframe(preview_df.head(10), use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                wb = writer.book
                ws = wb.create_sheet(title='매핑결과')
                ws.append(excel_headers)
                for row_data in excel_values:
                    ws.append([None if pd.isna(x) else x for x in row_data])
                if 'Sheet' in wb.sheetnames:
                    del wb['Sheet']

            excel_data = output.getvalue()

            st.download_button(
                label="📥 최종 완료 엑셀 다운로드 (.xlsx)",
                data=excel_data,
                file_name="선물세트_전체완성_RITC매핑결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )