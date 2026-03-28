import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import gspread

import re

def safe_money_int_cast(val):
    """
    金額文字列（カンマ、￥、小数点あり）を安全に整数に変換する。
    """
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    try:
        # 文字列クリーンアップ
        s = str(val).replace(",", "").replace("￥", "").replace("¥", "").strip()
        if not s:
            return 0
        # 小数点が含まれる場合は一旦floatにしてからintにする
        if "." in s:
            return int(float(s))
        return int(s)
    except Exception:
        return 0

def _get_year_month(ym_str):
    s = str(ym_str).strip()
    if not s:
        return (9999, 12)
    m = re.search(r"(\d{4})[年/\.\-](\d{1,2})", s)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m2 = re.search(r"(\d{2})[年/\.\-](\d{1,2})", s)
    if m2:
        return (2000 + int(m2.group(1)), int(m2.group(2)))
    return (9999, 12)

def ensure_id_column_and_formula(ws_pay):
    """
    支払管理シートのA列に『ID』列を追加し、固定費・変動費の各行に
    一意識別キーの数式をセットする。
    """
    from app import safe_gspread_call
    try:
        # シート全データを取得してヘッダー行を特定 (Ver 4.26.3)
        cells = safe_gspread_call(ws_pay.get_all_values)
        if not cells: return

        h_row_idx = -1
        ozukai_row = -1
        for i_r, r_v in enumerate(cells):
            if h_row_idx == -1 and r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
                h_row_idx = i_r
            # A列(index 0)またはB列(index 1)を検索
            search_str = ""
            if r_v:
                search_str = "".join([str(c) for c in r_v[:2]])
            if "小遣い" in search_str:
                ozukai_row = i_r + 1
                break
                
        if h_row_idx == -1: h_row_idx = 6 # フォールバック
        
        # 境界の設定 (Ver 4.27.2: 70行目フォールバック)
        boundary_row = ozukai_row if ozukai_row != -1 else 70

        actual_headers = cells[h_row_idx]
        is_already_id = False
        if len(actual_headers) > 0:
            h0 = str(actual_headers[0]).strip().lower()
            if h0 == "id" or h0 == "key":
                is_already_id = True
        
        if not is_already_id:
            safe_gspread_call(ws_pay.insert_cols, [["ID"]], 1)
            # 挿入後は cells をリロードした方が安全だが、ここでは簡易的に続行
            cells = safe_gspread_call(ws_pay.get_all_values)
        
        formulas = []
        # ヘッダーの次行から開始
        # ヘッダーの次行から開始、小遣い行の手前まで
        for i in range(h_row_idx + 1, len(cells)):
            if i + 1 >= boundary_row:
                break
            row = cells[i]
            r_idx = i + 1
            dai = str(row[1]).strip() if len(row) > 1 else ""
            k1 = str(row[4]).strip() if len(row) > 4 else ""
            
            if "固定費" in dai:
                if "クレジットカード" in k1:
                    formula = f"=B{r_idx}&F{r_idx}"
                else:
                    formula = f"=B{r_idx}&E{r_idx}"
                formulas.append([formula])
            elif "変動費" in dai:
                formula = f"=F{r_idx}"
                formulas.append([formula])
            else:
                formulas.append([""])
        
        if formulas:
            start_row = h_row_idx + 2
            end_row = h_row_idx + 1 + len(formulas)
            safe_gspread_call(ws_pay.update, values=formulas, range_name=f"A{start_row}:A{end_row}", value_input_option='USER_ENTERED')
            
    except Exception as e:
        print(f"Error in ensure_id_column_and_formula: {e}")

def _generate_target_months():
    # 2026.1月 to 2036.12月
    months = []
    for y in range(2026, 2037):
        for m in range(1, 13):
            months.append(f"{y}.{m}月")
    return months

# 文字列が数式形式 (="値") の場合に中身を取り出す補助関数
def _clean_val(v):
    s = str(v).strip()
    if s.startswith("="):
        s = s[1:].strip()
        if s.startswith('"') and s.endswith('"'):
            s = s[1:-1].strip()
    return s

# 強力な正規化（数値の整数化、全角半角の統一、空白除去）
def _normalize(s):
    if s is None: return ""
    s = str(s).strip()
    # 数値形式の正規化 (1.0 -> 1)
    try:
        f_val = float(s)
        if f_val == int(f_val):
            s = str(int(f_val))
    except:
        pass
    import unicodedata
    s_norm = unicodedata.normalize('NFKC', s)
    return "".join(s_norm.split())

def _find_val(d, keywords, exclude=[]):
    for k, v in d.items():
        clean_k = str(k).replace("\n", "").replace(" ", "").replace("　", "").strip()
        for kw in keywords:
            if kw in clean_k:
                if any(ex in clean_k for ex in exclude):
                    continue
                return v
    return ""

def execute_expansion(username, mode="NEW", start_ym=None):
    """
    mode: "NEW", "RE_EXECUTE", "NEXT_MONTH"
    """
    bk_ws = None
        
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    if not client:
        return False, "Google Docsへの接続に失敗しました。"
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
    except Exception as e:
        return False, f"支払管理シート({sheet_name})が見つかりません。"
        
    try:
        ws_master = ss.worksheet("固定費マスター")
        ws_pay = ss.worksheet("支払管理")
    except Exception as e:
        return False, f"「固定費マスター」または「支払管理」シートが見つかりません。"
        
    try:
        master_data = safe_gspread_call(ws_master.get_all_records)
    except Exception as e:
        # Fallback to get_all_values if header is weird
        master_raw = safe_gspread_call(ws_master.get_all_values)
        if not master_raw or len(master_raw) < 2:
            return False, "固定費マスターにデータがありません。"
        headers = master_raw[0]
        master_data = []
        for row in master_raw[1:]:
            row_dict = {}
            for i, h in enumerate(headers):
                if i < len(row):
                    row_dict[h] = row[i]
                else:
                    row_dict[h] = ""
            master_data.append(row_dict)

    # 数式を維持するために FORMULA レンダリングオプションで取得
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    
    # Extract headers from pay sheet (row 7 -> index 6)
    if len(pay_raw) < 7:
        return False, "「支払管理」のフォーマットが正しくありません（7行目にヘッダーが必要です）。"
        
    # ヘッダー行を「ID」が含まれる行として動的に特定 (Ver 4.26.2)
    h_row_idx = -1
    for i_r, r_v in enumerate(pay_raw):
        if r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
            h_row_idx = i_r
            break
    if h_row_idx == -1: h_row_idx = 6 # フォールバック
    pay_headers = pay_raw[h_row_idx]
    # --- ヘッダー・最終月の超堅牢取得 (Ver 5.3.0) ---
    pay_formatted = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMATTED_VALUE')
    actual_headers = pay_formatted[h_row_idx] if len(pay_formatted) > h_row_idx else []
    
    # 数式版と表示値版の長い方を基準にする
    header_len = max(len(pay_headers), len(actual_headers))
    actual_h_ids = []
    sheet_months = []
    
    # --- カレンダー背景行 (Row 4, 5, 6) を取得 (Ver 5.4.0) ---
    # index: Row 1->0, Row 5->4, Row 6->5
    y_row = pay_formatted[4] if len(pay_formatted) > 4 else []
    m_row = pay_formatted[5] if len(pay_formatted) > 5 else []
    
    for i in range(header_len):
        h_formula = _clean_val(pay_headers[i]) if i < len(pay_headers) else ""
        h_formatted = _clean_val(actual_headers[i]) if i < len(actual_headers) else ""
        
        # 年月の走査 (表示値 > 数式 > カレンダー行 の順で判定)
        y_m = _get_year_month(h_formatted)
        if y_m == (9999, 12):
            y_m = _get_year_month(h_formula)
        
        # --- カレンダー行からのバックアップ取得 (Ver 5.4.0) ---
        detected_ym_str = ""
        if y_m == (9999, 12):
            y_val = str(y_row[i]).strip() if i < len(y_row) else ""
            m_val = str(m_row[i]).strip() if i < len(m_row) else ""
            # 西暦4桁と月(1-12)を抽出
            yy = re.search(r"(\d{4})", y_val)
            mm_found = re.search(r"(\d{1,2})", m_val)
            if yy and mm_found:
                y_m = (int(yy.group(1)), int(mm_found.group(1)))
                detected_ym_str = f"{y_m[0]}.{y_m[1]}月"
        
        # 表示値ベースの正規化 (識別ID用) (Ver 5.4.2: 重複追加を完全に排除)
        norm_val = ""
        if h_formatted: norm_val = _normalize(h_formatted)
        elif h_formula: norm_val = _normalize(h_formula)
        elif detected_ym_str: norm_val = _normalize(detected_ym_str)
        actual_h_ids.append(norm_val)
        
        if y_m != (9999, 12):
            col_letter = chr(ord('A') + i) if i < 26 else chr(ord('A') + i//26 - 1) + chr(ord('A') + i%26)
            sheet_months.append((y_m, col_letter))
            
    if sheet_months:
        last_item = max(sheet_months, key=lambda x: x[0])
        sheet_last_ym = last_item[0]
        detected_col = last_item[1]
    else:
        sheet_last_ym = (2036, 12)
        detected_col = "末尾"
        
    st.toast(f"展開・最終月を検知: {sheet_last_ym[0]}.{sheet_last_ym[1]}月 ({detected_col}列)")
    month_cols = _generate_target_months()

    # --- 境界列判定（split_col_idx）の統合 (Ver 5.4.2) ---
    split_col_idx = 8 # デフォルト I列
    target_ym_for_split = start_ym if mode == "NEXT_MONTH" and start_ym else month_cols[0]
    norm_target = _normalize(target_ym_for_split)
    try:
        split_col_idx = actual_h_ids.index(norm_target)
    except:
        # あいまい検索
        for i, ach in enumerate(actual_h_ids):
            if norm_target.replace("月","") == ach.replace("月","") and ach != "":
                split_col_idx = i
                break
        
    ozukai_row = -1
    for i_r, r_v in enumerate(pay_raw):
        search_str = "".join([str(c) for c in r_v[:2]]) if r_v else ""
        if "小遣い" in search_str:
            ozukai_row = i_r + 1
            break
    boundary_row = ozukai_row if ozukai_row != -1 else 70


    # Base columns before months
    base_cols = ["大分類", "変動or固定", "有限or無限", "科目１", "科目２", "Sno", "科目詳細"]
    # Prepare old data if mode == NEXT_MONTH
    old_data_map = {}
    if mode == "NEXT_MONTH":
        # Find index of current month in month_cols
        current_ym_norm = _normalize(start_ym) if start_ym else ""
        if not current_ym_norm:
            now = datetime.now()
            current_ym_norm = f"{now.year}.{now.month}月"
            
        try:
            current_month_idx = month_cols.index(current_ym_norm)
        except:
            current_month_idx = 0
            
        # Determine the physical column index for the start month in the existing sheet
        # actual_headers を使って判定を行う
        try:
            split_col_idx = actual_h_ids.index(current_ym_norm)
        except:
            # Fallback
            try:
                # '月'の有無を許容して再試行
                alt_ym = current_ym_norm.replace("月", "")
                split_col_idx = next(i for i, h in enumerate(actual_h_ids) if h.replace("月", "") == alt_ym)
            except:
                split_col_idx = 7 
            
        # Read old rows and store protectable prefix (all cells to the left of target month)
        for row in pay_raw[h_row_idx + 1:]:
            k1_idx = next((i for i, h in enumerate(actual_h_ids) if "科目1" in h or "科目１" in h or "固定支払1" in h or "固定支払１" in h), 3)
            if len(row) <= k1_idx or not _normalize(row[k1_idx]) or "計" in _normalize(row[k1_idx]):
                continue
                
            k2_idx = next((i for i, h in enumerate(actual_h_ids) if "科目2" in h or "科目２" in h or "固定支払2" in h or "固定支払２" in h), 4)
            sno_idx = next((i for i, h in enumerate(actual_h_ids) if "Sno" in h or "seq" in h.lower()), 5)
            det_idx = next((i for i, h in enumerate(actual_h_ids) if "詳細" in h or "明細" in h), 6)
            fixed_idx = next((i for i, h in enumerate(actual_h_ids) if "変動" in h or ("固定" in h and "支払" not in h)), 1)
            finite_idx = next((i for i, h in enumerate(actual_h_ids) if "有限" in h or "無限" in h), 2)
            
            # Extract values for matching
            k1 = _normalize(_clean_val(row[k1_idx]))
            k2 = _normalize(_clean_val(row[k2_idx])) if k2_idx < len(row) else ""
            det = _normalize(_clean_val(row[det_idx])) if det_idx < len(row) else ""
            
            # Simplified 3-point identification key (Category1, Category2, Detail)
            # This is more robust against sno or finite/infinite changes.
            key = f"{k1}_{k2}_{det}"
            if key not in old_data_map:
                old_data_map[key] = []
            
            # Physical prefix protection: content of ALL columns before start month
            prefix_data = row[:split_col_idx] if split_col_idx < len(row) else row
            old_data_map[key].append(prefix_data)

    def _dict_to_row(d, prefix=None):
        r = []
        for i, h in enumerate(pay_headers):
            # Physical protection: If this index is before the split point, use the raw prefix data
            if prefix and i < len(prefix):
                r.append(prefix[i])
                continue
                
            # Otherwise, use normal logic (for columns from start_ym onwards)
            ach = actual_h_ids[i]
            if "科目1" in ach or "科目１" in ach or "固定支払1" in ach or "固定支払１" in ach: val = d.get("科目１", "")
            elif "科目2" in ach or "科目２" in ach or "固定支払2" in ach or "固定支払２" in ach: val = d.get("科目２", "")
            elif "変動" in ach or ("固定" in ach and "支払" not in ach): val = d.get("変動or固定", "")
            elif "有限" in ach or "無限" in ach: val = d.get("有限or無限", "")
            elif "Sno" in ach or "seq" in ach.lower(): val = d.get("Sno", "")
            elif "詳細" in ach or "明細" in ach: val = d.get("科目詳細", "")
            elif "大分類" in ach: val = d.get("大分類", "")
            elif ach in month_cols:
                val = d.get(ach, "")
            elif (ach + "月") in month_cols:
                val = d.get(ach + "月", "")
            elif ach.replace("月", "") in [m.replace("月", "") for m in month_cols]:
                # Find matched month item
                match_m = next((m for m in month_cols if m.replace("月", "") == ach.replace("月", "")), None)
                val = d.get(match_m, "")
            else:
                # 完了フラグ列の特定: 「月名の右隣の列」を最優先条件とする（位置ベース）
                is_flag_col = False
                prev_target_m = None
                if i > 0:
                    potential_prev_m = actual_h_ids[i-1]
                    if potential_prev_m in month_cols or (potential_prev_m + "月") in month_cols:
                        is_flag_col = True
                        prev_target_m = potential_prev_m if potential_prev_m in month_cols else (potential_prev_m + "月")
                
                if is_flag_col:
                    val = d.get(f"{prev_target_m}_flag", "")
                else:
                    # それ以外はフォールバック（大分類名等）
                    val = d.get(ach, d.get(h, ""))
            r.append(val)
        return r

    # --- 1. User_Masterから生年月日の取得とE2セル設定 ---
    try:
        k_ss = client.open("Kakeibo_Data")
        u_ws = k_ss.worksheet("User_Master")
        u_data = safe_gspread_call(u_ws.get_all_records)
        u_rec = next((u for u in u_data if u.get("username") == username), None)
        if u_rec:
            # yyyy-mm-dd or yyyy/mm/dd -> yyyymmdd
            b_str = str(u_rec.get("birthdate", "")).replace("-", "").replace("/", "").strip()
            if len(b_str) >= 8:
                b_val = b_str[:8]
                # ユーザーの要望により E2 から F2 に変更
                safe_gspread_call(ws_pay.update_acell, 'F2', b_val)
    except Exception as e:
        print(f"Birthdate update error: {e}")

    # --- 2. 支払管理シートの恒久バックアップ (固定名) ---
    try:
        bk_name = "支払管理BK"
        try:
            old_bk = ss.worksheet(bk_name)
            safe_gspread_call(ss.del_worksheet, old_bk)
        except Exception:
            pass # 既存のバックアップが存在しない場合は無視
            
        bk_ws = safe_gspread_call(ss.duplicate_sheet, ws_pay.id, new_sheet_name=bk_name)
    except Exception as e:
        print(f"Backup error: {e}")

    # --- 3. 固定費エリアのクリア (小遣い行の保護) ---
    try:
        # A8から境界の手前までを一括クリア（行削除は行わない）
        if boundary_row > 8:
            safe_gspread_call(ws_pay.batch_clear, [f"A8:ZZ{boundary_row - 1}"])
    except Exception as e:
        print(f"Clear fixed cost area error: {e}")

    try:
        new_rows_data = []
        budget_map = {}
    
        # Process master data
        category_groups = {"クレジットカード": [], "口座引落": [], "銀行振込": []}
        key_usage_counters = {} # key -> count
        
        sno = 1
        for m_rec in master_data:
            # Robust key matching
            k1 = _normalize(_clean_val(_find_val(m_rec, ["科目1", "科目１", "固定支払1", "固定支払１"])))
            if not k1:
                k1 = _normalize(_clean_val(m_rec.get("科目１", m_rec.get("固定支払１", ""))))
                
            k2 = _normalize(_clean_val(_find_val(m_rec, ["科目2", "科目２", "固定支払2", "固定支払２"])))
            is_finite_str = _normalize(_clean_val(_find_val(m_rec, ["有限", "無限"])))
            detail = _normalize(_clean_val(_find_val(m_rec, ["詳細", "明細"])))
            sno_val = _normalize(_clean_val(_find_val(m_rec, ["Sno", "seq"])))
            fixed_var = _normalize(_clean_val(_find_val(m_rec, ["変動", "固定"], exclude=["支払"])))
            
            amt_str = str(_find_val(m_rec, ["支払額", "金額"], exclude=["最終月額", "最終"])).replace(",", "").replace("¥", "").replace("￥", "")
            amt = safe_money_int_cast(amt_str)
            
            pay_month_freq = str(_find_val(m_rec, ["支払月", "頻度"])).strip()
            pay_year_freq = str(_find_val(m_rec, ["支払年", "年"], exclude=["月", "開始", "完済", "終了", "完了"])).strip()
            
            final_amt_str = str(_find_val(m_rec, ["最終月額"])).replace(",", "").replace("¥", "").replace("￥", "").strip()
            final_amt = safe_money_int_cast(final_amt_str) if final_amt_str else amt
            
            fee_str = str(_find_val(m_rec, ["振込手数料", "手数料"])).replace(",", "").replace("¥", "").replace("￥", "")
            fee = safe_money_int_cast(fee_str)
            
            start_m_str = str(_find_val(m_rec, ["開始"])).strip()
            end_m_str = str(_find_val(m_rec, ["完済", "終了", "完了"])).strip()
            
            if not k1:
                continue

            # --- 小遣い判定 (Ver 5.4.5: メインループ内で集計し、通常出力行からは除外) ---
            is_ozukai_budget = False
            if "小遣い" in (k1 + k2 + detail):
                is_ozukai_budget = True
                
            is_finite = ("有限" in is_finite_str)
            sy, sm = _get_year_month(start_m_str) if start_m_str else (0,0)
            ey, em = _get_year_month(end_m_str) if is_finite and end_m_str else (9999,12)
            
            # Simplified identifying key
            key = f"{k1}_{k2}_{detail}"
            
            # 同一キーの複数行対応: 出現順に old_data_map から取得
            key_usage_idx = key_usage_counters.get(key, 0)
            protected_prefix_array = None
            if mode == "NEXT_MONTH" and key in old_data_map:
                if key_usage_idx < len(old_data_map[key]):
                    protected_prefix_array = old_data_map[key][key_usage_idx]
                    key_usage_counters[key] = key_usage_idx + 1
            
            # Build row dict
            row_dict = {
                "大分類": "固定費",
                "変動or固定": fixed_var,
                "有限or無限": is_finite_str,
                "科目１": k1,
                "科目２": k2,
                "Sno": sno_val,
                "科目詳細": detail,
                "protected_prefix": protected_prefix_array # Physical cells to copy
            }
            
            # Calculate for each month
            # --- 実行モードに応じた開始月の決定 (Ver 5.2.0) ---
            if mode == "NEXT_MONTH" and start_ym:
                uy, um = _get_year_month(start_ym)
                # マスター開始月とユーザー選択月のうち、遅い方を採用
                if (uy > sy) or (uy == sy and um > sm):
                    eff_sy, eff_sm = uy, um
                else:
                    eff_sy, eff_sm = sy, sm
            else:
                eff_sy, eff_sm = sy, sm

            # --- 有限・無限に応じた終了月の決定 (Ver 5.2.0) ---
            eff_ey, eff_em = (ey, em) if is_finite else sheet_last_ym

            for mc in month_cols:
                my, mm = [int(x.replace("月","")) for x in mc.split(".")]
                
                # Logic: If month column found in sheet headers, and its index < split_col_idx, it should be protected.
                # But the dict_to_row function will handle mergingprotected_prefix.
                # We only calculate for the "new" range (>= split_col_idx).
                
                # Basic values for dict-to-row fallback
                    
                val = ""
                # Check Active duration
                if (my > eff_sy) or (my == eff_sy and mm >= eff_sm):
                    if (my < eff_ey) or (my == eff_ey and mm <= eff_em):
                        # It's active
                        
                        # Check Frequency (Ver 4.20.0: Support 偶数月/奇数月)
                        is_pay_month = False
                        f_clean = _normalize(pay_month_freq)
                        if "毎月" in f_clean:
                            is_pay_month = True
                        elif "偶数月" in f_clean:
                            is_pay_month = (mm % 2 == 0)
                        elif "奇数月" in f_clean:
                            is_pay_month = (mm % 2 != 0)
                        else:
                            # expected "9月" or similar
                            if str(mm) in f_clean:
                                is_pay_month = True
                                
                        # --- 支払年の判定を追加 (Ver 4.20.0) ---
                        if is_pay_month and pay_year_freq:
                            y_clean = _normalize(pay_year_freq)
                            if "偶数年" in y_clean:
                                if my % 2 != 0: is_pay_month = False
                            elif "奇数年" in y_clean:
                                if my % 2 == 0: is_pay_month = False
                        # --------------------------------------
                                
                        if is_pay_month:
                            # Is it the very last month?
                            if is_finite and my == ey and mm == em:
                                v = final_amt
                            else:
                                v = amt
                                
                            # Add fee
                            v += fee
                            
                            if is_ozukai_budget:
                                budget_map[mc] = budget_map.get(mc, 0) + v
                            else:
                                row_dict[mc] = v
                            
                if not is_ozukai_budget and mc not in row_dict:
                    row_dict[mc] = val
            
            # 小遣い予算行は row_dict としては登録しない（後で1行にまとめて更新するため）
            if is_ozukai_budget:
                continue

            # Group by 科目1
            if k1 not in category_groups:
                category_groups[k1] = []
            category_groups[k1].append(row_dict)
    
        # Now assemble the final sheet data
        # Ensure standard order: クレジットカード -> 口座引落 -> 銀行振込
        target_k1_order = ["クレジットカード", "口座引落", "銀行振込"]
        
        # Build array for batch update (Row 8 onwards)
        final_sheet_array = []
        # Ver 5.4.3: header_len は冒頭で計算した「カレンダー最大長」を維持し、ここで再定義（短縮）しない
    
        def dict_to_row(d):
            r = []
            prefix = d.get("protected_prefix") # List of original cells if in NEXT_MONTH protection
            
            for i in range(header_len):
                if i == 0: continue # Skip Column A (ID column)
                
                h = pay_headers[i] if i < len(pay_headers) else ""
                
                # Physical protection: If this index is before the split point, use the raw prefix data
                if prefix and i < len(prefix):
                    r.append(prefix[i])
                    continue
                    
                # Otherwise, use normal logic (for columns from start_ym onwards)
                ach = actual_h_ids[i]
                if "科目1" in ach or "科目１" in ach or "固定支払1" in ach or "固定支払１" in ach: val = d.get("科目１", "")
                elif "科目2" in ach or "科目２" in ach or "固定支払2" in ach or "固定支払２" in ach: val = d.get("科目２", "")
                elif "変動" in ach or ("固定" in ach and "支払" not in ach): val = d.get("変動or固定", "")
                elif "有限" in ach or "無限" in ach: val = d.get("有限or無限", "")
                elif "Sno" in ach or "seq" in ach.lower(): val = d.get("Sno", "")
                elif "詳細" in ach or "明細" in ach: val = d.get("科目詳細", "")
                elif "大分類" in ach: val = d.get("大分類", "")
                elif ach in month_cols:
                    val = d.get(ach, "")
                elif (ach + "月") in month_cols:
                    val = d.get(ach + "月", "")
                elif ach.replace("月", "") in [m.replace("月", "") for m in month_cols]:
                    # Find matched month item
                    match_m = next((m for m in month_cols if m.replace("月", "") == ach.replace("月", "")), None)
                    val = d.get(match_m, "")
                else:
                    # 完了フラグ列の特定: 「月名の右隣の列」を最優先条件とする（位置ベース）
                    is_flag_col = False
                    prev_target_m = None
                    if i > 0:
                        potential_prev_m = actual_h_ids[i-1]
                        if potential_prev_m in month_cols or (potential_prev_m + "月") in month_cols:
                            is_flag_col = True
                            prev_target_m = potential_prev_m if potential_prev_m in month_cols else (potential_prev_m + "月")
                    
                    if is_flag_col:
                        val = d.get(f"{prev_target_m}_flag", "")
                    else:
                        # それ以外はフォールバック（大分類名等）
                        val = d.get(ach, d.get(h, ""))
                r.append(val)
            return r
            
        start_row_idx = 8 # 1-based, after row 7 headers
        current_row_num = start_row_idx
        
        # Formula ranges memory
        group_ranges = []
        
        # Track rows for '変動' highlighting
        variable_rows = []
        
        # Format requests for borders
        sheet_id = ws_pay.id
        format_requests = []
        
        # 1. Clear previous borders from row 8 downwards
        format_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 7,
                    "endRowIndex": 500,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "top": {"style": "NONE"},
                "bottom": {"style": "NONE"},
                "left": {"style": "NONE"},
                "right": {"style": "NONE"},
                "innerHorizontal": {"style": "NONE"},
                "innerVertical": {"style": "NONE"}
            }
        })
        
        # 標準の順序（カード、引落、振込）を優先し、それ以外の新規カテゴリもすべて含める
        all_categories = target_k1_order + [k for k in category_groups.keys() if k not in target_k1_order]
        
        for tk in all_categories:
            group_rows = category_groups.get(tk, [])
            if not group_rows and tk not in target_k1_order:
                # 標準カテゴリ以外が空の場合はスキップ
                continue
                
            # --- 新仕様: カテゴリごとの出力行数の完全固定化 ---
            pad_limit = -1
            if tk == "クレジットカード":
                pad_limit = 15
            elif tk == "口座引落":
                pad_limit = 10
            elif tk == "銀行振込":
                pad_limit = 10
                
            # 上限を超えた場合は表示崩れを防ぐため切り捨てる
            if pad_limit != -1 and len(group_rows) > pad_limit:
                group_rows = group_rows[:pad_limit]

            g_start = current_row_num
            sno = 1
            
            last_var = ""
            last_inf = ""
            last_k2 = ""
            
            for gr in group_rows:
                gr["Sno"] = str(sno)
                sno += 1
                last_var = gr["変動or固定"]
                last_inf = gr["有限or無限"]
                last_k2 = gr["科目２"]
                
                final_sheet_array.append(dict_to_row(gr))
                current_row_num += 1
                
            # 指定された上限（pad_limit）に到達するまで空行動的パディング
            pad_amount = 3 # 指定外のカテゴリ用のデフォルト値
            if pad_limit != -1:
                pad_amount = pad_limit - len(group_rows)
                
            for _ in range(pad_amount):
                empty_dict = {
                    "大分類": "固定費",
                    "変動or固定": last_var,
                    "有限or無限": last_inf,
                    "科目１": tk,
                    "科目２": last_k2,
                    "Sno": str(sno)
                }
                sno += 1
                r = dict_to_row(empty_dict)
                final_sheet_array.append(r)
                current_row_num += 1
                
            g_end = current_row_num - 1
            
            # Add 計 row
            # B列から開始するため header_len - 1
            subtotal_row = [""] * (header_len - 1)
            try:
                k1_idx = -1
                for i_h, h_v in enumerate(actual_h_ids):
                    if h_v == "科目1" or h_v == "科目１" or h_v == "固定支払1" or h_v == "固定支払１":
                        k1_idx = i_h
                        break
                if k1_idx != -1:
                    # B列開始なのでインデックスを調整 (-1) し、境界チェックを追加
                    if 0 <= k1_idx - 1 < len(subtotal_row):
                        subtotal_row[k1_idx - 1] = f"【{tk} 計】"
            except: pass
            
            # Add formulas for months
            for mc in month_cols:
                try:
                    # actual_h_ids からインデックスを特定
                    m_norm = _normalize(mc)
                    c_idx = -1
                    for i_h, h_v in enumerate(actual_h_ids):
                        if h_v == m_norm or h_v.replace("月","") == m_norm.replace("月",""):
                            c_idx = i_h
                            break
                    
                    if c_idx == -1: continue
                    
                    f_idx = c_idx + 1 # 完了F
                    col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                    flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                    if g_start <= g_end:
                        # 完了Fが空の行のみを合計
                        formula = f"=SUMIFS({col_letter}{g_start}:{col_letter}{g_end}, {flag_letter}{g_start}:{flag_letter}{g_end}, \"\")"
                        # B列開始なのでインデックスを調整 (-1) し、境界チェックを追加
                        if 0 <= c_idx - 1 < len(subtotal_row):
                            subtotal_row[c_idx - 1] = formula
                except: pass
                
            final_sheet_array.append(subtotal_row)
            
            # Add thick border format request for this group
            format_requests.append({
                "updateBorders": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": g_start - 1, # 0-based
                        "endRowIndex": current_row_num, # 0-based exclusive (includes the subtotal row)
                        "startColumnIndex": 0,
                        "endColumnIndex": header_len
                    },
                    "top": {"style": "SOLID_MEDIUM"},
                    "bottom": {"style": "SOLID_MEDIUM"},
                    "left": {"style": "SOLID_MEDIUM"},
                    "right": {"style": "SOLID_MEDIUM"}
                }
            })
            
            group_ranges.append(current_row_num) # save the row number of the subtotal
            
            # Add background color and bold for subtotal row
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": current_row_num - 1, # 0-based
                        "endRowIndex": current_row_num,
                        "startColumnIndex": 0,
                        "endColumnIndex": header_len
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 0.95,
                                "blue": 0.8
                            },
                            "textFormat": {
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat.bold)"
                }
            })
            
            current_row_num += 1
            
        # Add Grand Total Row
        # B列から開始するため header_len - 1
        grand_total_row = [""] * (header_len - 1)
        try:
            # ヘッダー項目をクリーンアップして検索 (="科目１" 等の数式対応)
            clean_pay_headers = [_clean_val(h) for h in pay_headers]
            k1_idx = clean_pay_headers.index("科目１")
            
            # B列開始なのでインデックスを調整 (-1) し、境界チェックを追加
            if 0 < len(grand_total_row): grand_total_row[0] = "固定費合計"
            if 0 <= k1_idx - 1 < len(grand_total_row):
                grand_total_row[k1_idx - 1] = "固定費合計"
            
            # 科目明細列（G列付近）にも念のため文言を入れる
            k_detail_idx = -1
            for h_i, h_val in enumerate(pay_headers):
                h_clean = _clean_v(h_val).strip()
                if "科目明細" in h_clean or "詳細" in h_clean: 
                    k_detail_idx = h_i
                    break
            if k_detail_idx != -1:
                if 0 <= k_detail_idx - 1 < len(grand_total_row):
                    grand_total_row[k_detail_idx - 1] = "固定費合計"
                
        except: pass
        
        for mc in month_cols:
            try:
                # actual_h_ids からインデックスを特定
                m_norm = _normalize(mc)
                c_idx = -1
                for i_h, h_v in enumerate(actual_h_ids):
                    if h_v == m_norm or h_v.replace("月","") == m_norm.replace("月",""):
                        c_idx = i_h
                        break
                
                if c_idx == -1: continue
                
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                if group_ranges:
                    # 各サブグループの「計」行そのものがすでにSUMIFSになっているので、
                    # グランドトータルは単純にそれ。を合計しても良いが、完了Fが1の「計」行はないはず。
                    # B列開始なのでインデックスを調整 (-1) し、境界チェックを追加
                    if group_ranges:
                        cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in group_ranges]
                        formula = f"=SUM({','.join(cells)})"
                        if 0 <= c_idx - 1 < len(grand_total_row):
                            grand_total_row[c_idx - 1] = formula
            except: pass
            
        # 合計行のA〜G列（またはheader_len）を結合して中央揃え
        try:
            k_detail_idx = -1
            for h_i, h_val in enumerate(pay_headers):
                h_clean = _clean_v(h_val).strip()
                if "科目明細" in h_clean or "詳細" in h_clean: k_detail_idx = h_i
            
            # もし科目明細が見つからなければ、G列（インデックス6）までを対象とする
            merge_end = k_detail_idx + 1 if k_detail_idx != -1 else 7
            
            # 既存の結合を解除（エラー防止）- データエリア全体の可能性のある範囲
            format_requests.append({
                "unmergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 7,  # Row 8 onwards
                        "endRowIndex": ws_pay.row_count,
                        "startColumnIndex": 0,
                        "endColumnIndex": header_len
                    }
                }
            })
            
            format_requests.append({
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": current_row_num - 1,
                        "endRowIndex": current_row_num,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    }
                }
            })
            # 中央揃え
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": current_row_num - 1,
                        "endRowIndex": current_row_num,
                        "startColumnIndex": 0,
                        "endColumnIndex": merge_end
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            })
        except: pass
        
        final_sheet_array.append(grand_total_row)
        
        # Add background color and bold for grand total row (Fixed Cost Total - Blue Reverse)
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": current_row_num - 1,
                    "endRowIndex": current_row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.0,
                            "green": 0.0,
                            "blue": 1.0
                        },
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat.bold,textFormat.foregroundColor)"
            }
        })
        total_data_end = current_row_num
        
        # Base borders request for ALL data (insert at index 1 so it runs after ANY clears and before THICK borders)
        base_border_request = {
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 7, # 8行目から
                    "endRowIndex": 1000, # 常に1000行目まで書式を維持（空行でも数式や書式が入るように）
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "top": {"style": "SOLID"},
                "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"},
                "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"},
                "innerVertical": {"style": "SOLID"}
            }
        }
        format_requests.insert(1, base_border_request)
        
        # Add thick borders for each year (12 month columns)
        try:
            years = sorted(list(set([mc.split(".")[0] for mc in month_cols])))
            for y in years:
                start_m = f"{y}.1月"
                end_m = f"{y}.12月"
                if start_m in pay_headers and end_m in pay_headers:
                    sc = pay_headers.index(start_m)
                    ec = pay_headers.index(end_m)
                    format_requests.append({
                        "updateBorders": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 6, # include header
                                "endRowIndex": total_data_end,
                                "startColumnIndex": sc,
                                "endColumnIndex": ec + 1 # exclusive
                            },
                            "top": {"style": "SOLID_MEDIUM"},
                            "bottom": {"style": "SOLID_MEDIUM"},
                            "left": {"style": "SOLID_MEDIUM"},
                            "right": {"style": "SOLID_MEDIUM"}
                        }
                    })
        except:
            pass
            
        # Format numbers with commas (#,##0)
        try:
            first_month = month_cols[0]
            if first_month in pay_headers:
                start_m_col = pay_headers.index(first_month)
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 7, # data only
                            "endRowIndex": 1000, # Extend to cover potential future rows
                            "startColumnIndex": start_m_col,
                            "endColumnIndex": header_len
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "numberFormat": {
                                    "type": "NUMBER",
                                    "pattern": "#,##0"
                                }
                            }
                        },
                        "fields": "userEnteredFormat.numberFormat"
                    }
                })
        except:
            pass
            
        # 完了Fの左側を破線にする
        try:
            for c_idx, h_name in enumerate(pay_headers):
                clean_h = str(h_name).replace("\n", "").replace(" ", "").strip()
                if "完了" in clean_h:
                    format_requests.append({
                        "updateBorders": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 7, 
                                "endRowIndex": 1000,
                                                                "endColumnIndex": c_idx + 1
                            },
                            "left": {"style": "DASHED"}
                        }
                    })
        except: pass
            
        # 最終的なシート配列を全スキャンして「変動」の行を特定する (Ver 4.15.0 厳密化)
        variable_rows = []
        try:
            # 「変動or固定」列のインデックスを特定
            v_idx = -1
            for i, h in enumerate(pay_headers):
                ch = _normalize(_clean_val(h))
                if "変動" in ch or ("固定" in ch and "支払" not in ch):
                    v_idx = i
                    break
            
            if v_idx != -1:
                for r_idx, row in enumerate(final_sheet_array):
                    if v_idx < len(row):
                        row_val = _normalize(_clean_val(row[v_idx]))
                        if "変動" in row_val:
                            variable_rows.append(r_idx)
        except Exception as e:
            print(f"Error in variable_rows scan: {e}")
            
        # Format rows with "変動" to have yellow background in cols B to G (index 1 to 7)
        for r_idx in variable_rows:
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": r_idx + 7, # 0-indexed row number (Row 8 is index 7)
                        "endRowIndex": r_idx + 8,
                        "startColumnIndex": 1,
                        "endColumnIndex": 7
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 0.4  # Yellow
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
        
        # Write back to sheet and apply formats
        # 空間の確保と書き込み (Ver 4.27.0: 小遣い行を押し出す)
        data_row_count = len(final_sheet_array)
        available_space = boundary_row - 8 # Row 8 to boundary_row-1
        
        if data_row_count > available_space:
            # 不足分を境界行（小遣い行）に挿入して押し出す
            diff = data_row_count - available_space
            safe_gspread_call(ws_pay.insert_rows, [[""] * header_len] * diff, row=boundary_row)
        elif ozukai_row == -1:
            # 小遣い行がない場合の従来の挙動 (全件更新時の行数確保)
            current_rows = ws_pay.row_count
            needed_total = 7 + data_row_count
            if needed_total > current_rows:
                safe_gspread_call(ws_pay.add_rows, needed_total - current_rows)
            
        # データ書き込み開始位置を B8 (作成開始位置) に変更
        safe_gspread_call(ws_pay.update, values=final_sheet_array, range_name="B8", value_input_option='USER_ENTERED')
        
        # --- 小遣い予算の書き込み (統合Ver 5.4.5) ---
        try:
            ozukai_budget_row_idx = -1
            target_det_idx = next((i for i, h in enumerate(actual_headers) if "詳細" in _clean_val(h).strip() or "明細" in _clean_val(h).strip()), -1)
            
            if target_det_idx != -1:
                final_pay_cells = safe_gspread_call(ws_pay.get_all_values)
                for i, row in enumerate(final_pay_cells):
                    if target_det_idx < len(row) and _normalize(_clean_val(row[target_det_idx])) == "小遣い予算":
                        ozukai_budget_row_idx = i + 1
                        break
                        
            if ozukai_budget_row_idx != -1 and 'budget_map' in locals() and budget_map:
                base_row = final_pay_cells[ozukai_budget_row_idx - 1]
                # B列からヘッダー長までのデータを作成
                update_vals = [base_row[i] if i < len(base_row) else "" for i in range(1, header_len)]
                for i in range(1, header_len):
                    ach = actual_h_ids[i] if i < len(actual_h_ids) else ""
                    if ach in budget_map:
                        update_vals[i-1] = budget_map[ach]
                safe_gspread_call(ws_pay.update, values=[update_vals], range_name=f"B{ozukai_budget_row_idx}", value_input_option='USER_ENTERED')
        except Exception as e:
            print(f"Ozukai Budget Expansion Error: {e}")

        # 書式を一括適用
        safe_gspread_call(ss.batch_update, {"requests": format_requests})

        # --- 新仕様：バックアップシートは削除せず、常に最新を維持して保管 ---
        pass

        # --- 正常終了時の更新日時記録 (F4: 固定費データ展開) ---
        try:
            current_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            safe_gspread_call(ws_pay.update, values=[[current_now]], range_name="F4", value_input_option='USER_ENTERED')
        except:
            pass

        # --- A列(ID列)の保守と数式設定 ---
        ensure_id_column_and_formula(ws_pay)
        # --- 自動で変動費データ更新を実施 (Ver 4.17.0 追加仕様) ---
        if mode == "NEXT_MONTH":
            # 内部で execute_variable_cost_update を呼び出す (バックアップ重複回避)
            v_success, v_msg = execute_variable_cost_update(username, start_ym, skip_backup=True)
            if not v_success:
                return True, f"固定費展開は成功しましたが、変動費更新でエラーが発生しました: {v_msg}"
            return True, "固定費展開および変動費データ更新が正常に完了しました！"

        return True, "データ展開に成功しました！"
    except Exception as e:
        return False, f"書き込みエラー: {e}"
    finally:
        # バックアップデータの明示的な破棄 (Ver 4.17.0 追加仕様)
        try:
            if 'pay_raw' in locals(): del pay_raw
            if 'old_data_map' in locals(): del old_data_map
        except: pass

def show_open_management_sheet():
    """支払管理シートを確認する UI"""
    st.markdown("<h2 style='font-size: 1.75rem !important;'>📊 支払管理シートを確認</h2>", unsafe_allow_html=True)
    
    username = st.session_state.get("username", "")
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        try:
            ws = ss.worksheet("支払管理")
            target_url = getattr(ws, 'url', f"{ss.url}#gid={ws.id}")
        except:
            target_url = ss.url
    except Exception as e:
        st.warning(f"現在、あなた（{username}）専用の支払管理シートは見つかりません。")
        st.info("「支払管理シート新規作成」メニューからシートを発行してください。")
        return
        
    st.info("この画面では、月々の支払予定を一覧管理する『支払管理』シートを確認できます。口座引落日などの条件に合わせて完了フラグが自動更新され、家計全体の収支見通しを立てるのに役立ちます。")
    
    st.link_button("🌐 開く", url=target_url, type="primary", use_container_width=True)

def show_fixed_cost_data_expansion():
    """固定費データ展開 UI"""
    st.markdown("## 🛠️ 固定費データ展開")
    st.info("「固定費マスター」の情報をもとに、「支払管理」シートに月別のデータ（2036年12月まで）を展開・更新します。")
    
    username = st.session_state.get("username", "")
    
    from app import get_gspread_client, safe_gspread_call
    
    # 状態確認
    client = get_gspread_client()
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        ws_pay = ss.worksheet("支払管理")
        pay_raw = safe_gspread_call(ws_pay.get_all_values)
    except Exception as e:
        st.warning(f"現在、あなた（{username}）専用の支払管理シート、または必要なシートが見つかりません。")
        st.info("「支払管理シート新規作成」メニューからシートを発行してください。")
        return
        
    # 展開済みか判定
    # 8行目以降に何らかのデータがあるか
    is_expanded = False
    
    # ヘッダー行を「ID」が含まれる行として動的に特定 (Ver 4.26.3)
    h_row_idx = -1
    for i_r, r_v in enumerate(pay_raw):
        if r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
            h_row_idx = i_r
            break
    if h_row_idx == -1: h_row_idx = 6 # フォールバック

    if len(pay_raw) > h_row_idx + 1:
        # Check if rows after header have actual data
        for row in pay_raw[h_row_idx + 1:]:
            # check any cell has value
            if any(cell.strip() for cell in row):
                is_expanded = True
                break
                
    if is_expanded:
        st.warning(f"「{sheet_name}」は既にデータ展開済です。下記をご確認下さい。")
        
        # 確認状態を管理
        if "fce_confirm_re_execute" not in st.session_state:
            st.session_state["fce_confirm_re_execute"] = False
        if "fce_confirm_next_month" not in st.session_state:
            st.session_state["fce_confirm_next_month"] = False

        if st.session_state["fce_confirm_re_execute"]:
            st.warning("⚠️ **既に設定済のデータも全て再作成されてしまいます。よろしいでしょうか？**")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("✅ はい（実行）", type="primary", use_container_width=True):
                    st.session_state["fce_action"] = "RE_EXECUTE"
                    st.session_state["fce_confirm_re_execute"] = False
                    st.rerun()
            with c2:
                if st.button("❌ キャンセル", use_container_width=True):
                    st.session_state["fce_confirm_re_execute"] = False
                    st.rerun()
        elif st.session_state["fce_confirm_next_month"]:
            st.info("📅 **展開を開始する年月を選択してください（選択した月以降が更新されます）。**")
            # 選択肢の生成
            from dateutil.relativedelta import relativedelta
            now_dt = datetime.now()
            next_month_dt = now_dt + relativedelta(months=1)
            
            years = [y for y in range(2026, 2037) if y >= next_month_dt.year]
            months_all = [m for m in range(1, 13)]
            
            col_y, col_m = st.columns(2)
            with col_y:
                sel_y = st.selectbox("開始年", years, index=0)
            with col_m:
                # 選択された年が翌月の年と同じなら、翌月以降に制限
                if sel_y == next_month_dt.year:
                    valid_months = [m for m in months_all if m >= next_month_dt.month]
                else:
                    valid_months = months_all
                sel_m = st.selectbox("開始月", valid_months, index=0)
            
            target_ym = f"{sel_y}.{sel_m}月"
            st.write(f"展開・更新の適用開始月: **{target_ym}**")
            st.caption(f"※{target_ym}より前の月のデータ（金額および完了フラグ）は絶対に変更されません。")
            
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🚀 確定して実行", type="primary", use_container_width=True):
                    st.session_state["fce_action"] = "NEXT_MONTH"
                    st.session_state["fce_start_ym"] = target_ym
                    st.session_state["fce_confirm_next_month"] = False
                    st.rerun()
            with c2:
                if st.button("❌ キャンセル", use_container_width=True, key="cancel_nm"):
                    st.session_state["fce_confirm_next_month"] = False
                    st.rerun()
        else:
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("再実行", type="primary", use_container_width=True):
                    st.session_state["fce_confirm_re_execute"] = True
                    st.rerun()
            with col2:
                if st.button("翌月以降", use_container_width=True):
                    st.session_state["fce_confirm_next_month"] = True
                    st.rerun()
            with col3:
                if st.button("キャンセル", use_container_width=True):
                    st.session_state["fce_action"] = "CANCEL"
                    st.rerun()
                
        action = st.session_state.get("fce_action")
        if action in ["RE_EXECUTE", "NEXT_MONTH"]:
            with st.spinner("データ展開中...（数秒〜数十秒かかります）"):
                # NEXT_MONTH の場合は開始年月を渡す
                start_ym = st.session_state.get("fce_start_ym") if action == "NEXT_MONTH" else None
                success, msg = execute_expansion(username, mode=action, start_ym=start_ym)
                if success:
                    # 続けて変動費データ更新を実行 (バックアップ重複回避)
                    with st.spinner("続けて変動費（クレジットカード利用等）を集計中..."):
                        v_success, v_msg = execute_variable_cost_update(username, skip_backup=True)
                        if v_success:
                            st.success("固定費データの展開と変動費の集計が完了しました！")
                        else:
                            st.warning(f"固定費の展開は完了しましたが、変動費の更新に失敗しました: {v_msg}")
                    st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")
                    st.session_state["fce_action"] = None
                else:
                    st.error(msg)
                    
    else:
        if st.button("新規データ展開", type="primary"):
            with st.spinner("データ展開中...（数秒〜数十秒かかります）"):
                success, msg = execute_expansion(username, mode="NEW")
                if success:
                    # 続けて変動費データ更新を実行 (バックアップ重複回避)
                    with st.spinner("続けて変動費（クレジットカード利用等）を集計中..."):
                        v_success, v_msg = execute_variable_cost_update(username, skip_backup=True)
                        if v_success:
                            st.success("固定費データの展開と変動費の集計が完了しました！")
                        else:
                            st.warning(f"固定費の展開は完了しましたが、変動費の更新に失敗しました: {v_msg}")
                    st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")
                else:
                    st.error(msg)

def execute_variable_cost_update(username, start_ym=None, skip_backup=False):
    from app import get_gspread_client, safe_gspread_call, get_payment_methods, get_sheet, TRANSACTIONS_WORKSHEET_NAME
    import calendar
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    
    client = get_gspread_client()
    if not client:
        return False, "Google Docsへの接続に失敗しました。"
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        ws_pay = ss.worksheet("支払管理")
    except Exception as e:
        return False, f"支払管理シート({sheet_name})が見つかりません。先に「支払管理シート新規作成」を実行してください。"
        
    # --- 変動費データ更新 単独実行時のバックアップ作成 ---
    if not skip_backup:
        try:
            bk_name = "支払管理BK"
            try:
                old_bk = ss.worksheet(bk_name)
                safe_gspread_call(ss.del_worksheet, old_bk)
            except Exception:
                pass # 既存の同名バックアップが存在しない場合は無視
            
            # バックアップシートの作成と恒久保存
            safe_gspread_call(ss.duplicate_sheet, ws_pay.id, new_sheet_name=bk_name)
        except Exception as e:
            print(f"Variable Cost Update Backup error: {e}")

    # 数式を維持するために FORMULA レンダリングオプションで取得
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    if len(pay_raw) < 7:
        return False, "「支払管理」のフォーマットが正しくありません。"
        
    # ヘッダー行を「ID」が含まれる行として動的に特定 (Ver 4.26.2)
    h_row_idx = -1
    ozukai_row = -1
    for i_r, r_v in enumerate(pay_raw):
        if h_row_idx == -1 and r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
            h_row_idx = i_r
        search_str = "".join([str(c) for c in r_v[:2]]) if r_v else ""
        if "小遣い" in search_str:
            ozukai_row = i_r + 1
            break
            
    if h_row_idx == -1: h_row_idx = 6 # フォールバック
    boundary_row = ozukai_row if ozukai_row != -1 else 70
    # --- ヘッダー・最終月の超堅牢取得 (Ver 5.4.1/移植) ---
    pay_formatted = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMATTED_VALUE')
    actual_headers = pay_formatted[h_row_idx] if len(pay_formatted) > h_row_idx else []
    pay_headers = pay_raw[h_row_idx]
    
    header_len = max(len(pay_headers), len(actual_headers))
    actual_h_ids = []
    y_row = pay_formatted[4] if len(pay_formatted) > 4 else []
    m_row = pay_formatted[5] if len(pay_formatted) > 5 else []
    
    for i in range(header_len):
        h_formula = _clean_val(pay_headers[i]) if i < len(pay_headers) else ""
        h_formatted = _clean_val(actual_headers[i]) if i < len(actual_headers) else ""
        
        # 検知ロジック
        y_m = _get_year_month(h_formatted)
        if y_m == (9999, 12): y_m = _get_year_month(h_formula)
        
        detected_ym_str = ""
        if y_m == (9999, 12):
            y_v = str(y_row[i]).strip() if i < len(y_row) else ""
            m_v = str(m_row[i]).strip() if i < len(m_row) else ""
            yy = re.search(r"(\d{4})", y_v)
            mm_f = re.search(r"(\d{1,2})", m_v)
            if yy and mm_f:
                detected_ym_str = f"{yy.group(1)}.{mm_f.group(1)}月"
        
        if h_formatted: norm_v = _normalize(h_formatted)
        elif h_formula: norm_v = _normalize(h_formula)
        elif detected_ym_str: norm_v = _normalize(detected_ym_str)
        else: norm_v = ""
        actual_h_ids.append(norm_v)
    
    # 科目１列のインデックスを探す
    k1_idx = -1
    for i, h in enumerate(actual_h_ids):
        if "科目1" in h or "科目１" in h or "固定支払1" in h or "固定支払１" in h:
            k1_idx = i
            break
            
    if k1_idx == -1:
        return False, "ヘッダーから「科目１」列が見つかりません。"
        
    # 行をスキャンして「固定費合計」または「【合計】」を探す (大分類～科目詳細のどこにあっても見つける)
    total_row_idx = -1
    for i, row in enumerate(pay_raw):
        if i <= h_row_idx: continue
        row_str = "".join([str(c) for c in row[:7]]) # 最初の数列を結合して検索
        if "固定費合計" in row_str or "【合計】" in row_str:
            total_row_idx = i
            break
            
    # 見つからない場合のフォールバック：上から順に SUM(IF... 数式がある最初の行を探す (ラベルが消えている場合への対策)
    if total_row_idx == -1:
        for i in range(h_row_idx + 1, len(pay_raw)):
            row = pay_raw[i]
            # 月カラム（通常はインデックス 7 以降）のどこかにグランド合計用の数式が入っているかチェック
            row_content = "".join([str(c) for c in row])
            if "=SUM(IF(" in row_content or "=SUMIFS(" in row_content:
                # 明細行のサブ計ではなく、複数の SUM かつ IF( が含まれるグランド合計っぽいもの
                if row_content.count("SUM(") >= 1 and row_content.count("IF(") > 1:
                    total_row_idx = i
                    break

    if total_row_idx == -1:
        return False, "「固定費合計」または「【合計】」行が見つかりません。先に「固定費データ展開」を実行してください。"
        
    # 既存の固定費エリアのサブ合計行とグランド合計行を区別して収集 (Ver 5.0.0 循環参照防止)
    fixed_subtotals = []
    subtotal_row_nums = []
    grand_total_row_num = -1
    group_start = h_row_idx + 2
    for i, row in enumerate(pay_raw[:total_row_idx + 1]):
        if i <= h_row_idx: continue
        r_k1 = str(row[k1_idx]).strip() if k1_idx < len(row) else ""
        if ("【" in r_k1 and "計】" in r_k1):
            fixed_subtotals.append((i + 1, group_start, i))
            subtotal_row_nums.append(i + 1)
            group_start = i + 2
        elif "【合計】" in r_k1 or "固定費合計" in r_k1:
            grand_total_row_num = i + 1
            
    from fixed_cost_expansion import _generate_target_months
    # 対象月カラムの抽出
    month_cols = _generate_target_months()
    
    # 新しい挿入開始行 (スプレッドシートの行番号は 1-based)
    # Ver 4.27.4: 固定費合計の次から開始 (+2)
    start_row_num = total_row_idx + 2
    
    try:
        methods = safe_gspread_call(get_payment_methods, username)
        cc_methods = [m for m in methods if m.get("is_credit_card", False) or m.get("type") == "クレジットカード"]
    except Exception as e:
        return False, f"支払方法マスターの取得に失敗しました: {e}"
        
    if not cc_methods:
        return False, "クレジットカードが1件も登録されていません。「支払方法マスター」を確認してください。"
        
    try:
        tx_sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
        all_txs = safe_gspread_call(tx_sheet.get_all_records)
        user_txs = [tx for tx in all_txs if str(tx.get("username", "")).lower() == username.lower()]
    except Exception as e:
        return False, f"取引履歴の取得に失敗しました: {e}"
        
    def _dict_to_row(d):
        r = []
        for i in range(header_len):
            h = actual_headers[i] if i < len(actual_headers) else ""
            
            ach = actual_h_ids[i]
            if "科目1" in ach or "科目１" in ach or "固定支払1" in ach or "固定支払１" in ach: val = d.get("科目１", "")
            elif "科目2" in ach or "科目２" in ach or "固定支払2" in ach or "固定支払２" in ach: val = d.get("科目２", "")
            elif "変動" in ach or ("固定" in ach and "支払" not in ach): val = d.get("変動or固定", "")
            elif "有限" in ach or "無限" in ach: val = d.get("有限or無限", "")
            elif "Sno" in ach or "seq" in ach.lower(): val = d.get("Sno", "")
            elif "詳細" in ach or "明細" in ach: val = d.get("科目詳細", "")
            elif "大分類" in ach: val = d.get("大分類", "")
            elif ach in month_cols:
                val = d.get(ach, "")
            elif (ach + "月") in month_cols:
                val = d.get(ach + "月", "")
            elif ach.replace("月", "") in [m.replace("月", "") for m in month_cols]:
                match_m = next((m for m in month_cols if m.replace("月", "") == ach.replace("月", "")), None)
                val = d.get(match_m, "")
            else:
                # 完了フラグ等のフォールバック
                val = d.get(ach, d.get(_clean_val(h).strip(), ""))
            r.append(val)
        return r

    cc_rows_array = []
    current_row_num = start_row_num
    var_start = current_row_num
    var_sno = 1
    
    # 事前準備：pay_raw内から各クレジットカードごとの固定費行(行番号, 1-based)を抽出
    # ヘッダーより下（7行目以降）から合計行まで
    fc_payment_rows = {cc.get("name", ""): [] for cc in cc_methods}
    # 他の支払情報 (口座引落, 銀行振込) の行を抽出
    other_pay_rows = []
    
    for i, row in enumerate(pay_raw[:total_row_idx]):
        if i <= h_row_idx: continue
        try:
            # ヘッダー検索も表示値ベースのインデックスを使用
            k1_h_idx = -1
            k2_h_idx = -1
            for h_i, h_val in enumerate(actual_headers):
                h_clean = _clean_val(h_val).strip()
                if h_clean == "科目１": k1_h_idx = h_i
                if h_clean == "科目２": k2_h_idx = h_i
            
            if k1_h_idx != -1 and k2_h_idx != -1:
                r_k1 = _clean_val(row[k1_h_idx]).strip()
                r_k2 = _clean_val(row[k2_h_idx]).strip()
                if r_k1 == "クレジットカード" and r_k2 in fc_payment_rows:
                    fc_payment_rows[r_k2].append(i + 1)
                elif r_k1 in ["口座引落", "銀行振込"]:
                    if r_k2:
                        other_pay_rows.append(i + 1)
        except: pass
        
    var_cost_rows = {}
    cc_monthly_amounts = {}
    
    for cc in cc_methods:
        cc_name = cc.get("name", "")
        closing_str = str(cc.get("closing_date", ""))
        pay_month_str = str(cc.get("payment_month", ""))
        pay_date_str = str(cc.get("payment_date", ""))
        
        # 支払日の数値抽出と文言（翌月/当月）の判定
        try:
            p_day_match = re.search(r"\d+", str(pay_date_str))
            p_day_val = int(p_day_match.group()) if p_day_match else 27
        except: p_day_val = 27
        timing_label = "翌月" if p_day_val <= 20 else "当月"
        payment_desc = f"支払日は{timing_label}の{pay_date_str}となります。"

        cc_row_dict = {
            "大分類": "変動費",
            "変動or固定": "",
            "有限or無限": "",
            "科目１": "クレジットカード",
            "科目２": cc_name,
            "科目詳細": payment_desc,
            "Sno": str(var_sno)
        }
        var_sno += 1
        
        for mc in month_cols:
            try:
                my, mm = [int(x.replace("月","")) for x in mc.split(".")]
                base_date = datetime(my, mm, 1)
                
                # 表示ルールの適用: 20日基準で表示年月をシフト
                # 20日以前なら：支払月 ＝ 表示月
                # 21日以降なら：支払月 ＝ 表示月 － 1ヶ月 (翌月の21日以降なら翌月の年月に表示、の逆算)
                try:
                    p_day_m = re.search(r"\d+", str(pay_date_str))
                    p_day = int(p_day_m.group()) if p_day_m else 27
                except:
                    p_day = 27

                if p_day <= 20:
                    # 支払月の前月に表示する（例: 2/10払 → 1月カラム） => 支払月 ＝ 表示月 ＋ 1ヶ月
                    target_pay_date = base_date + relativedelta(months=1)
                else:
                    # 支払月と同じ年月に表示する（例: 2/27払 → 2月カラム） => 支払月 ＝ 表示月
                    target_pay_date = base_date
                
                from app import calculate_credit_card_periods
                periods = calculate_credit_card_periods(target_pay_date, closing_str, pay_month_str, pay_date_str)
                
                if periods and len(periods) > 0:
                    s_date = periods[0]["start"]
                    e_date = periods[0]["end"]
                    
                    month_sum = 0
                    norm_cc_name = _normalize(_clean_val(cc_name))
                    for tx in user_txs:
                        tx_pm = _normalize(_clean_val(tx.get("payment_method", "")))
                        if tx_pm == norm_cc_name:
                            # 内税は集計に含まない
                            if tx.get("category") == "消費税（内税）":
                                continue
                            tx_date_str = str(tx.get("date", ""))
                            if tx_date_str:
                                try:
                                    tx_date = datetime.strptime(tx_date_str, "%Y-%m-%d").date()
                                    if s_date <= tx_date <= e_date:
                                        amt = str(tx.get("amount", "0")).replace(",", "").replace("¥", "").replace("￥", "")
                                        if amt is not None and str(amt).strip() != "":
                                            month_sum += safe_money_int_cast(amt)
                                except:
                                    pass
                    cc_row_dict[mc] = month_sum if month_sum > 0 else ""
                else:
                    cc_row_dict[mc] = ""
            except Exception as e:
                cc_row_dict[mc] = ""

        var_cost_rows[cc_name] = current_row_num
        cc_monthly_amounts[cc_name] = { mc: cc_row_dict.get(mc, "") for mc in month_cols }
        row_arr = _dict_to_row(cc_row_dict)
        
        # 月次データとフラグの転記
        for mc in month_cols:
            try:
                # actual_headers を使ってインデックスを取得
                c_idx = -1
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)) == mc:
                        c_idx = i_h
                        break
                if c_idx == -1:
                    # '月'なしでも試行
                    alt_m = mc.replace("月", "")
                    for i_h, h_v in enumerate(actual_headers):
                        if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                            c_idx = i_h
                            break
                
                if c_idx == -1: continue # 見つからない場合はスキップ
                f_idx = c_idx + 1
                
                # 金額設定 - 正しいインデックスを使用 (Ver 4.27.4: -1を削除)
                amt = cc_row_dict.get(mc, "")
                if 0 <= c_idx < len(row_arr):
                    row_arr[c_idx] = amt
                
                # 支払日の判定と自動フラグ設定 (変動費・カード個別行)
                try:
                    p_day_m = re.search(r"\d+", str(pay_date_str))
                    p_day = int(p_day_m.group()) if p_day_m else 27
                    
                    # 20日基準の表示年月シフトルールに従って支払日を特定
                    y, m = _get_year_month(mc)
                    if y != 9999:
                        base_date_dt = datetime(y, m, 1)
                        if p_day <= 20:
                            target_pay_dt = base_date_dt + relativedelta(months=1)
                        else:
                            target_pay_dt = base_date_dt
                        
                        from app import calculate_credit_card_periods
                        t_periods = calculate_credit_card_periods(target_pay_dt, closing_str, pay_month_str, pay_date_str)
                        if t_periods:
                            actual_due_date = t_periods[0]["pay_date"]
                            if datetime.now().date() >= actual_due_date:
                                # 正しいインデックスを使用 (Ver 4.27.4: -1を削除)
                                if 0 <= f_idx < len(row_arr):
                                    row_arr[f_idx] = "1"
                except: pass
            except: pass
            
        cc_rows_array.append(row_arr)
        current_row_num += 1
        
    # 変動費合計行を追加 (ピンク)
    # Ver 4.28.2: 表示内容を「変動費_支払残_合計額」に変更
    var_total_row = [""] * header_len
    try:
        if 0 < len(var_total_row):
            var_total_row[1] = "変動費_支払残_合計額" # B列位置
        # A列(ID列)は空欄（画像2に合わせる）
        if 0 < len(var_total_row):
            var_total_row[0] = ""
    except: pass
    
    for mc in month_cols:
        try:
            # actual_headers を使ってインデックスを取得
            c_idx = -1
            for i_h, h_v in enumerate(actual_headers):
                if _normalize(_clean_val(h_v)) == mc:
                    c_idx = i_h
                    break
            if c_idx == -1:
                alt_m = mc.replace("月", "")
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                        c_idx = i_h
                        break
            if c_idx == -1: continue
            
            f_idx = c_idx + 1
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            if var_start <= current_row_num - 1:
                # 完了Fが空の行のみを合計
                formula = f"=SUMIFS({col_letter}{var_start}:{col_letter}{current_row_num - 1}, {flag_letter}{var_start}:{flag_letter}{current_row_num - 1}, \"\")"
                # 正しいインデックスを使用 (Ver 4.27.4: -1を削除)
                if 0 <= c_idx < len(var_total_row):
                    var_total_row[c_idx] = formula
        except: pass
        
    cc_rows_array.append(var_total_row)
    current_row_num += 1
    
    # === クレジットカード総合計ブロック ===
    summary_start_row = current_row_num
    card_total_rows = []
    
    summary_data_start = current_row_num
    
    # 最初の行にだけ「クレジットカード合計」を表示。他は空にする。
    is_first_summary_row = True
    
    for cc in cc_methods:
        cc_name = cc.get("name", "")
        pay_month_str = str(cc.get("payment_month", ""))
        pay_date_str = str(cc.get("payment_date", ""))
        closing_str = str(cc.get("closing_date", ""))
        
        # 1. 固定費分 (Ver 4.28.1: B-E列は後で結合されるため、ラベルを「クレジットカード 支払残額」に統一)
        fc_dict = {
            "大分類": "クレジットカード 支払残額", 
            "科目１": "クレジットカード 支払残額", 
            "科目２": cc_name, 
            "科目詳細": "固定費分"
        }
        r_fc = _dict_to_row(fc_dict)
        for mc in month_cols:
            try:
                # actual_headers を使ってインデックスを取得
                c_idx = -1
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)) == mc:
                        c_idx = i_h
                        break
                if c_idx == -1:
                    alt_m = mc.replace("月", "")
                    for i_h, h_v in enumerate(actual_headers):
                        if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                            c_idx = i_h
                            break
                if c_idx == -1: continue
                
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                if fc_payment_rows.get(cc_name):
                    cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in fc_payment_rows[cc_name]]
                    # B列開始なのでインデックスを調整し、境界チェックを追加
                    if 0 <= c_idx < len(r_fc):
                        r_fc[c_idx] = f"=SUM({','.join(cells)})"
                else:
                    if 0 <= c_idx < len(r_fc):
                        r_fc[c_idx] = 0
                
                # 支払日の判定と自動フラグ設定 (固定費分)
                try:
                    y, m = _get_year_month(mc)
                    if y != 9999:
                        p_day_m = re.search(r"\d+", str(pay_date_str))
                        p_day = int(p_day_m.group()) if p_day_m else 27
                        
                        base_date_dt = datetime(y, m, 1)
                        if p_day <= 20:
                            target_pay_dt = base_date_dt + relativedelta(months=1)
                        else:
                            target_pay_dt = base_date_dt
                            
                        from app import calculate_credit_card_periods
                        t_periods = calculate_credit_card_periods(target_pay_dt, closing_str, pay_month_str, pay_date_str)
                        if t_periods:
                            actual_due_date = t_periods[0]["pay_date"]
                            if datetime.now().date() >= actual_due_date:
                                # B列開始なのでインデックスを調整し、境界チェックを追加
                                if 0 <= f_idx < len(r_fc):
                                    r_fc[f_idx] = 1
                except: pass
            except: pass
        cc_rows_array.append(r_fc)
        current_row_num += 1
        
        # 2. 変動費分 (Ver 4.28.1: ラベル更新)
        vc_row_idx = var_cost_rows.get(cc_name)
        vc_dict = {
            "大分類": "クレジットカード 支払残額", 
            "科目１": "クレジットカード 支払残額", 
            "科目２": cc_name, 
            "科目詳細": "変動費分"
        }
        r_vc = _dict_to_row(vc_dict)
        for mc in month_cols:
            try:
                # actual_headers を使ってインデックスを取得
                c_idx = -1
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)) == mc:
                        c_idx = i_h
                        break
                if c_idx == -1:
                    alt_m = mc.replace("月", "")
                    for i_h, h_v in enumerate(actual_headers):
                        if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                            c_idx = i_h
                            break
                if c_idx == -1: continue
                
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                if vc_row_idx:
                    # 詳細行の完了Fが空の場合のみ表示
                    # B列開始なのでインデックスを調整し、境界チェックを追加
                    if 0 <= c_idx < len(r_vc):
                        r_vc[c_idx] = f"=IF({flag_letter}{vc_row_idx}=\"\", {col_letter}{vc_row_idx}, 0)"
                else:
                    if 0 <= c_idx < len(r_vc):
                        r_vc[c_idx] = 0

                # 支払日の判定と自動フラグ設定 (変動費分)
                try:
                    y, m = _get_year_month(mc)
                    if y != 9999:
                        month_offset = 1 
                        if "当月" in str(pay_month_str): month_offset = 0
                        elif "翌々月" in str(pay_month_str): month_offset = 2
                        target_base_date = datetime(y, m, 1) + relativedelta(months=month_offset)
                        day_match = re.search(r"\d+", str(pay_date_str))
                        if day_match:
                            day_val = int(day_match.group())
                            actual_due_date = target_base_date + relativedelta(day=day_val)
                            if datetime.now().date() >= actual_due_date.date():
                                # B列開始なのでインデックスを調整し、境界チェックを追加
                                if 0 <= f_idx < len(r_vc):
                                    r_vc[f_idx] = 1
                except: pass
            except: pass
        cc_rows_array.append(r_vc)
        current_row_num += 1
        
        # 3. 合計 (固定費分+変動費分)
        # 支払日の文言判定
        try:
            p_day_sum_m = re.search(r"\d+", str(pay_date_str))
            p_day_sum_v = int(p_day_sum_m.group()) if p_day_sum_m else 27
        except: p_day_sum_v = 27
        timing_sum_label = "翌月" if p_day_sum_v <= 20 else "当月"
        payment_sum_desc = f"支払日は{timing_sum_label}の{pay_date_str}となります。"

        sum_dict = {"大分類": "クレジットカード 支払残額", "科目１": "クレジットカード 支払残額", "科目２": cc_name, "科目詳細": payment_sum_desc}
        r_sum = _dict_to_row(sum_dict)
        for mc in month_cols:
            try:
                # actual_headers を使ってインデックスを取得
                c_idx = -1
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)) == mc:
                        c_idx = i_h
                        break
                if c_idx == -1:
                    alt_m = mc.replace("月", "")
                    for i_h, h_v in enumerate(actual_headers):
                        if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                            c_idx = i_h
                            break
                if c_idx == -1: continue
                
                f_idx = c_idx + 1
                col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                
                fc_row = current_row_num - 2
                vc_row = current_row_num - 1
                # 各内訳行（固定費分、変動費分）の完了Fをチェックして合計する形式
                formula = f"SUMIFS({col_letter}{fc_row}:{col_letter}{vc_row}, {flag_letter}{fc_row}:{flag_letter}{vc_row}, \"\")"
                # 正しいインデックスを使用 (Ver 4.27.4: -1を削除)
                if 0 <= c_idx < len(r_sum):
                    r_sum[c_idx] = "=" + formula
            except: pass
        cc_rows_array.append(r_sum)
        card_total_rows.append(current_row_num)
        current_row_num += 1
        
    # クレジットカード支払残_合計額 (Orange)
    # Ver 4.28.2: B列以降に配備
    grand_dict = {
        "大分類": "クレジットカード支払残_合計額", 
        "科目１": "クレジットカード支払残_合計額"
    }
    r_grand = _dict_to_row(grand_dict)
    # A列(ID列)は空欄
    r_grand[0] = ""
    for mc in month_cols:
        try:
            # actual_headers を使ってインデックスを取得
            c_idx = -1
            for i_h, h_v in enumerate(actual_headers):
                if _normalize(_clean_val(h_v)) == mc:
                    c_idx = i_h
                    break
            if c_idx == -1:
                alt_m = mc.replace("月", "")
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                        c_idx = i_h
                        break
            if c_idx == -1: continue
            
            f_idx = c_idx + 1
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            if card_total_rows:
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in card_total_rows]
                # 正しいインデックスを使用 (Ver 4.27.4: -1を削除)
                if 0 <= c_idx < len(r_grand):
                    r_grand[c_idx] = f"=SUM({','.join(cells)})"
            else:
                if 0 <= c_idx < len(r_grand):
                    r_grand[c_idx] = 0
        except: pass
    cc_rows_array.append(r_grand)
    current_row_num += 1
    
    # === 支払合計 (Ver 4.28.2: B列以降) ===
    pay_total_dict = {"大分類": "固定費＆変動費_支払残_合計額", "科目１": "固定費＆変動費_支払残_合計額"}
    r_pay_total = _dict_to_row(pay_total_dict)
    # A列(ID列)は空欄
    r_pay_total[0] = ""
    
    for mc in month_cols:
        try:
            # actual_headers を使ってインデックスを取得
            c_idx = -1
            for i_h, h_v in enumerate(actual_headers):
                if _normalize(_clean_val(h_v)) == mc:
                    c_idx = i_h
                    break
            if c_idx == -1:
                alt_m = mc.replace("月", "")
                for i_h, h_v in enumerate(actual_headers):
                    if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                        c_idx = i_h
                        break
            if c_idx == -1: continue

            f_idx = c_idx + 1 # 完了Fのインデックス (通常は月の右隣)
            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
            
            formula_parts = []
            
            # 1. 口座引落・銀行振込
            if other_pay_rows:
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in other_pay_rows]
                formula_parts.append(f"SUM({','.join(cells)})")
                
            # 2. クレジットカード合計行
            if card_total_rows:
                cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in card_total_rows]
                formula_parts.append(f"SUM({','.join(cells)})")
            
            if formula_parts:
                # 正しいインデックスを使用 (Ver 4.27.4: -1を削除)
                if 0 <= c_idx < len(r_pay_total):
                    r_pay_total[c_idx] = "=" + "+".join(formula_parts)
            else:
                if 0 <= c_idx < len(r_pay_total):
                    r_pay_total[c_idx] = 0
        except: pass
    cc_rows_array.append(r_pay_total)
    current_row_num += 1
    
    summary_end_row = current_row_num - 1
    
    header_len = len(actual_headers)
    
    # Append the variable cost area array
    try:
        # 古い変動費（合計より下、小遣いより上）の値をクリア（書式は維持）
        current_rows = ws_pay.row_count
        if start_row_num <= current_rows:
            clear_end_row = min(boundary_row - 1, current_rows)
            if start_row_num <= clear_end_row:
                safe_gspread_call(ws_pay.batch_clear, [f"B{start_row_num}:ZZ{clear_end_row}"])

        # 空間の確保と書き込み (小遣い行を押し出す)
        data_count = len(cc_rows_array)
        available_space = boundary_row - start_row_num
        
        if data_count > available_space:
            diff = data_count - available_space
            safe_gspread_call(ws_pay.insert_rows, [[""] * header_len] * diff, row=boundary_row)
            
        # 新しい変動費データを書き込み
        # A列をスキップしてB列から書き込み (Ver 4.27.4: r[1:] で A列を除外)
        safe_gspread_call(ws_pay.update, values=[r[1:] for r in cc_rows_array], range_name=f"B{start_row_num}", value_input_option='USER_ENTERED')
        
        # 既存の固定費サブ合計行も新方式の数式（完了F対応）に更新する
        if fixed_subtotals:
            update_data = [] # List of {'range': ..., 'values': [[...]]}
            # 1. 各サブセクションの集計行を更新
            for row_num, s_row, e_row in fixed_subtotals:
                row_vals = [""] * header_len
                for mc in month_cols:
                    c_idx = -1
                    for i_h, h_v in enumerate(actual_headers):
                        if _normalize(_clean_val(h_v)) == mc:
                            c_idx = i_h
                            break
                    if c_idx == -1:
                        alt_m = mc.replace("月", "")
                        for i_h, h_v in enumerate(actual_headers):
                            if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                                c_idx = i_h
                                break
                    if c_idx != -1:
                        f_idx = c_idx + 1
                        col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                        flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                        formula = f"=SUMIFS({col_letter}{s_row}:{col_letter}{e_row}, {flag_letter}{s_row}:{flag_letter}{e_row}, \"\")"
                        if 0 <= c_idx < len(row_vals): row_vals[c_idx] = formula
                
                # 月カラム以外のメタデータ保持
                original_row = pay_raw[row_num - 1]
                for i in range(len(actual_headers)):
                    if i < len(row_vals) and i < len(original_row): row_vals[i] = original_row[i]

                update_data.append({'range': f"A{row_num}", 'values': [row_vals]})

            # 2. グランド合計（固定費合計）行の更新 (Ver 5.0.1 循環参照回避ロジック)
            if grand_total_row_num != -1:
                row_vals = [""] * header_len
                for mc in month_cols:
                    c_idx = -1
                    for i_h, h_v in enumerate(actual_headers):
                        if _normalize(_clean_val(h_v)) == mc:
                            c_idx = i_h
                            break
                    if c_idx == -1:
                        alt_m = mc.replace("月", "")
                        for i_h, h_v in enumerate(actual_headers):
                            if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                                c_idx = i_h
                                break
                    if c_idx != -1:
                        f_idx = c_idx + 1
                        col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                        flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                        if subtotal_row_nums:
                            inner_cells = [f"IF({flag_letter}{r}=\"\", {col_letter}{r}, 0)" for r in subtotal_row_nums]
                            formula = f"=SUM({','.join(inner_cells)})"
                            if 0 <= c_idx < len(row_vals): row_vals[c_idx] = formula

                # メタデータ保持
                original_row = pay_raw[grand_total_row_num - 1]
                for i in range(len(actual_headers)):
                    if i < len(row_vals) and i < len(original_row): row_vals[i] = original_row[i]
                
                # ラベル補正 (Ver 4.28.2: B列以降)
                for j in range(7):
                    if j < len(row_vals) and ("【合計】" in str(row_vals[j]) or "固定費合計" in str(row_vals[j]) or "_支払残_" in str(row_vals[j])):
                        row_vals[j] = "固定費_支払残_合計額"
                        # A列は空のまま（後でB-H結合するため）
                        row_vals[0] = ""
                        break

                update_data.append({'range': f"A{grand_total_row_num}", 'values': [row_vals]})

            # --- 追加: 固定費エリアも含めた全シートの自動フラグ更新 ---
            scan_update_data = []
            try:
                # 固定費マスターから引落日情報を取得
                ws_master = ss.worksheet("固定費マスター")
                master_records = safe_gspread_call(ws_master.get_all_records)
                
                # (科目1, 科目2, 科目明細) -> (引落日, 支払月/頻度) のマップを作成
                def _get_master_k1(m): return str(_find_val(m, ["科目1", "科目１", "固定支払1", "固定支払１"])).strip()
                def _get_master_k2(m): return str(_find_val(m, ["科目2", "科目２", "固定支払2", "固定支払２"])).strip()
                def _get_master_detail(m): return str(_find_val(m, ["詳細", "明細"])).strip()
                def _get_master_due_date(m): return str(_find_val(m, ["口座引落日", "引落日"])).strip()
                def _get_master_pay_month(m): return str(_find_val(m, ["支払月", "頻度"])).strip()

                bank_withdrawal_map = {}
                for m in master_records:
                    mk1 = _get_master_k1(m)
                    if mk1 == "口座引落":
                        key = (mk1, _get_master_k2(m), _get_master_detail(m))
                        bank_withdrawal_map[key] = {
                            "due_date": _get_master_due_date(m),
                            "pay_month": _get_master_pay_month(m)
                        }

                # 科目１、科目２、大分類、詳細のインデックス特定
                idx_k1 = -1
                idx_k2 = -1
                idx_dai = -1
                idx_det = -1
                for i, h in enumerate(actual_headers):
                    h_c = _clean_val(h).strip()
                    if h_c == "科目１": idx_k1 = i
                    if h_c == "科目２": idx_k2 = i
                    if h_c == "大分類": idx_dai = i
                    if "詳細" in h_c or "明細" in h_c: idx_det = i

                # 8行目から、変動費エリアの手前までを走査
                for r_idx in range(h_row_idx + 1, start_row_num - 1):
                    if r_idx >= len(pay_raw): break
                    row = pay_raw[r_idx]
                    
                    is_cc = False
                    is_bw = False
                    current_info = None
                    
                    # 科目１または大分類で判定
                    r_k1 = str(row[idx_k1]).strip() if idx_k1 != -1 else ""
                    r_k2 = str(row[idx_k2]).strip() if idx_k2 != -1 else ""
                    r_det = str(row[idx_det]).strip() if idx_det != -1 else ""
                    
                    if r_k1 == "クレジットカード": is_cc = True
                    elif r_k1 == "口座引落": is_bw = True
                    
                    if is_cc:
                        for cc in cc_methods:
                            if cc.get("name") == r_k2:
                                current_info = {
                                    "p_month": str(cc.get("payment_month", "")),
                                    "p_date": str(cc.get("payment_date", ""))
                                }
                                break
                    elif is_bw:
                        bw_key = ("口座引落", r_k2, r_det)
                        if bw_key in bank_withdrawal_map:
                            info = bank_withdrawal_map[bw_key]
                            current_info = {
                                "p_month": info["pay_month"],
                                "p_date": info["due_date"]
                            }
                    
                    if current_info:
                        row_to_update = list(row)
                        row_changed = False
                        p_month = current_info["p_month"]
                        p_date = current_info["p_date"]
                        
                        for mc in month_cols:
                            try:
                                # actual_headers を使ってインデックスを取得
                                c_idx = -1
                                for i_h, h_v in enumerate(actual_headers):
                                    if _normalize(_clean_val(h_v)) == mc:
                                        c_idx = i_h
                                        break
                                if c_idx == -1:
                                    alt_m = mc.replace("月", "")
                                    for i_h, h_v in enumerate(actual_headers):
                                        if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                                            c_idx = i_h
                                            break
                                if c_idx == -1: continue
                                
                                f_idx = c_idx + 1
                                if f_idx >= len(row_to_update): continue
                                
                                if str(row_to_update[f_idx]).strip() == "1": continue
                                
                                # 支払日判定
                                y, m = _get_year_month(mc)
                                if y != 9999:
                                    if is_bw:
                                        # 口座引落の場合：引落日の文言を優先的に判定
                                        if "翌月" in p_date: off = 1
                                        elif "当月" in p_date: off = 0
                                        elif "前月" in p_date: off = -1
                                        else:
                                            # 文言がない場合はデフォルトで当月判定（または頻度から推測）
                                            off = 0
                                    else:
                                        # クレジットカードの場合
                                        off = 1
                                        if "当月" in p_month: off = 0
                                        elif "翌々月" in p_month: off = 2
                                    
                                    base_dt = datetime(y, m, 1) + relativedelta(months=off)
                                    d_m = re.search(r"\d+", p_date)
                                    if d_m:
                                        d_v = int(d_m.group())
                                        due_dt = base_dt + relativedelta(day=d_v)
                                        if datetime.now().date() >= due_dt.date():
                                            row_to_update[f_idx] = "1"
                                            row_changed = True
                            except: pass
                        
                        if row_changed:
                            scan_update_data.append({
                                'range': f"B{r_idx + 1}",
                                'values': [row_to_update[1:]] # A列以外を書き込む
                            })
            except Exception as e:
                print(f"Error in full sheet scan: {e}")

            if scan_update_data:
                update_data.extend(scan_update_data)
            
            if update_data:
                # batch_update を使用して効率的に更新
                safe_gspread_call(ws_pay.batch_update, update_data, value_input_option='USER_ENTERED')

        # フォーマット適用 (罫線など)
        format_requests = []
        sheet_id = ws_pay.id
        
        # 1. Base borders for the new rows
        format_requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row_num - 1,
                    "endRowIndex": current_row_num - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": header_len
                },
                "top": {"style": "SOLID"}, "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"}, "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"}, "innerVertical": {"style": "SOLID"}
            }
        })
        
        # 1.1 完了Fの左側を破線にする
        try:
            for c_idx, h_name in enumerate(actual_headers):
                clean_h = str(h_name).replace("\n", "").replace(" ", "").strip()
                if "完了" in clean_h:
                    format_requests.append({
                        "updateBorders": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 7,
                                "endRowIndex": 1000,
                                "startColumnIndex": c_idx,
                                "endColumnIndex": c_idx + 1
                            },
                            "left": {"style": "DASHED"}
                        }
                    })
        except: pass

        # --- 合計行・集計行の描画（各行個別に解除してから結合） ---

        # A) 固定費_支払残_合計額 (Blue) B-H合併
        if grand_total_row_num != -1:
            format_requests.append({
                "unmergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": grand_total_row_num - 1, "endRowIndex": grand_total_row_num,
                        "startColumnIndex": 0, "endColumnIndex": 256
                    }
                }
            })
            format_requests.append({
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": grand_total_row_num - 1, "endRowIndex": grand_total_row_num,
                        "startColumnIndex": 1, "endColumnIndex": 8
                    }
                }
            })
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": grand_total_row_num - 1, "endRowIndex": grand_total_row_num,
                        "startColumnIndex": 1, "endColumnIndex": 8
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0},
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 12}
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat.bold,textFormat.foregroundColor,textFormat.fontSize)"
                }
            })

        # B) 変動費_支払残_合計額 (Green) B-H合併
        var_total_phys_row = summary_start_row - 1
        format_requests.append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": var_total_phys_row - 1, "endRowIndex": var_total_phys_row,
                    "startColumnIndex": 0, "endColumnIndex": 256
                }
            }
        })
        format_requests.append({
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": var_total_phys_row - 1, "endRowIndex": var_total_phys_row,
                    "startColumnIndex": 1, "endColumnIndex": 8
                }
            }
        })
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": var_total_phys_row - 1, "endRowIndex": var_total_phys_row,
                    "startColumnIndex": 1, "endColumnIndex": 8
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.0, "green": 1.0, "blue": 0.0},
                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "fontSize": 12}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat.bold,textFormat.foregroundColor,textFormat.fontSize)"
            }
        })

        # C) クレジットカード 支払残額 (B-E列 縦結合)
        if summary_end_row - 2 >= summary_data_start - 1:
            format_requests.append({
                "unmergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_data_start - 1, "endRowIndex": summary_end_row - 2,
                        "startColumnIndex": 0, "endColumnIndex": 256
                    }
                }
            })
            format_requests.append({
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_data_start - 1, "endRowIndex": summary_end_row - 2,
                        "startColumnIndex": 1, "endColumnIndex": 5
                    }
                }
            })
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": summary_data_start - 1, "endRowIndex": summary_end_row - 2,
                        "startColumnIndex": 1, "endColumnIndex": 5
                    },
                    "cell": {
                        "userEnteredFormat": { "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "textFormat": {"bold": True} }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat.bold)"
                }
            })

        # D) クレジットカード支払残_合計額 (Orange) B-H合併
        format_requests.append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_end_row - 2, "endRowIndex": summary_end_row - 1,
                    "startColumnIndex": 0, "endColumnIndex": 256
                }
            }
        })
        format_requests.append({
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_end_row - 2, "endRowIndex": summary_end_row - 1,
                    "startColumnIndex": 1, "endColumnIndex": 8
                }
            }
        })
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_end_row - 2, "endRowIndex": summary_end_row - 1,
                    "startColumnIndex": 1, "endColumnIndex": 8
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 1.0, "green": 0.65, "blue": 0.0},
                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "fontSize": 12}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat.bold,textFormat.foregroundColor,textFormat.fontSize)"
            }
        })

        # E) 固定費＆変動費_支払残_合計額 (Red) B-H合併
        format_requests.append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_end_row - 1, "endRowIndex": summary_end_row,
                    "startColumnIndex": 0, "endColumnIndex": 256
                }
            }
        })
        format_requests.append({
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_end_row - 1, "endRowIndex": summary_end_row,
                    "startColumnIndex": 1, "endColumnIndex": 8
                }
            }
        })
        format_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": summary_end_row - 1, "endRowIndex": summary_end_row,
                    "startColumnIndex": 1, "endColumnIndex": 8
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0},
                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 12}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat.bold,textFormat.foregroundColor,textFormat.fontSize)"
            }
        })

        # --- 追加の全体書式設定（年単位の太枠、カンマ区切り設定など） ---
        try:
            years = sorted(list(set([mc.split(".")[0] for mc in month_cols])))
            for y in years:
                start_m = f"{y}.1月"
                end_m = f"{y}.12月"
                if start_m in actual_headers and end_m in actual_headers:
                    sc = actual_headers.index(start_m)
                    ec = actual_headers.index(end_m)
                    format_requests.append({
                        "updateBorders": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 7, "endRowIndex": 1000,
                                "startColumnIndex": sc, "endColumnIndex": ec + 1
                            },
                            "left": {"style": "SOLID_MEDIUM"}, "right": {"style": "SOLID_MEDIUM"}
                        }
                    })
            
            # 数値のカンマ区切りおよびマイナス赤字設定
            # 8行目以降（インデックス7〜）のすべての列を対象に一括適用
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 7, "endRowIndex": 1000,
                        "startColumnIndex": 0, "endColumnIndex": header_len
                    },
                    "cell": { "userEnteredFormat": { "numberFormat": { "type": "NUMBER", "pattern": "#,##0;[Red]-#,##0;" } } },
                        "fields": "userEnteredFormat.numberFormat"
                    }
                })
        except: pass

        if format_requests:
            safe_gspread_call(ss.batch_update, {"requests": format_requests})

        # --- 更新日時記録 (F5) ---
        try:
            current_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            safe_gspread_call(ws_pay.update, values=[[current_now]], range_name="F5", value_input_option='USER_ENTERED')
        except: pass

        # --- A列(ID列)の保守と数式設定 ---
        ensure_id_column_and_formula(ws_pay)

        # --- 追加仕様: 支払い方法の名前と月別集計金額をH72から順に設定 ---
        try:
            from app import get_payment_methods
            pm_list = get_payment_methods(username)
            if pm_list:
                valid_pms = [m for m in pm_list if str(m.get("name", "")).strip()]
                if valid_pms:
                    # H列(index 7)を起点とした2D配列を作成
                    grid_width = header_len - 7 if header_len > 7 else 1
                    update_grid = []
                    
                    # クレジットカード以外用の月別集計辞書を transactions から作成
                    tx_monthly_amounts = {}
                    for tx in user_txs:
                        t_cat = _normalize(_clean_val(tx.get("category", "")))
                        if "内税" in t_cat: continue
                        
                        t_pm = _normalize(_clean_val(tx.get("payment_method", "")))
                        if not t_pm: continue
                        
                        t_date_str = str(tx.get("date", ""))
                        if not t_date_str: continue
                        
                        try:
                            dt = datetime.strptime(t_date_str, "%Y-%m-%d")
                            t_mc = f"{dt.year}.{dt.month}月"
                            
                            if t_mc in month_cols:
                                amt_val = str(tx.get("amount", "0")).replace(",", "").replace("¥", "").replace("￥", "")
                                amt = safe_money_int_cast(amt_val)
                                if amt > 0:
                                    if t_pm not in tx_monthly_amounts: tx_monthly_amounts[t_pm] = {}
                                    tx_monthly_amounts[t_pm][t_mc] = tx_monthly_amounts[t_pm].get(t_mc, 0) + amt
                        except: pass
                        
                    for m in valid_pms:
                        pm_name = str(m.get("name", "")).strip()
                        pm_type = str(m.get("type", "")).strip()
                        is_cc = m.get("is_credit_card", False)
                        norm_pm = _normalize(_clean_val(pm_name))
                        
                        row_data = [""] * grid_width
                        row_data[0] = pm_name # H列にセット
                        
                        for mc in month_cols:
                            # actual_headers上での対象月インデックスを検索
                            c_idx = -1
                            for i_h, h_v in enumerate(actual_headers):
                                if _normalize(_clean_val(h_v)) == mc:
                                    c_idx = i_h
                                    break
                            if c_idx == -1:
                                alt_m = mc.replace("月", "")
                                for i_h, h_v in enumerate(actual_headers):
                                    if _normalize(_clean_val(h_v)).replace("月", "") == alt_m:
                                        c_idx = i_h
                                        break
                                        
                            if c_idx >= 7 and (c_idx - 7) < grid_width:
                                val = ""
                                if is_cc or pm_type == "クレジットカード":
                                    val = cc_monthly_amounts.get(pm_name, {}).get(mc, "")
                                else:
                                    sum_amt = tx_monthly_amounts.get(norm_pm, {}).get(mc, 0)
                                    if sum_amt > 0: val = sum_amt
                                row_data[c_idx - 7] = val
                                
                        update_grid.append(row_data)
                        
                    if update_grid:
                        # 既存の記述が残らないようにH72から該当行のZZ列までをクリア
                        end_row = 72 + len(update_grid) - 1
                        safe_gspread_call(ws_pay.batch_clear, [f"H72:ZZ{end_row}"])
                        
                        safe_gspread_call(ws_pay.update, values=update_grid, range_name="H72", value_input_option='USER_ENTERED')
        
        except Exception as e:
            print(f"H72 Payment Methods Tracking update error: {e}")

        # --- 新仕様: 「収入」シートからのデータ自動連係マッピング ---
        try:
            try:
                ws_income = ss.worksheet("収入")
            except:
                ws_income = None
            
            if ws_income:
                income_raw = safe_gspread_call(ws_income.get_all_values, value_render_option='UNFORMATTED_VALUE')
                if len(income_raw) > 7: # 8行目があるか
                    income_headers = income_raw[7] # 8行目 (0-based index 7)
                    income_data_map = {}
                    
                    # 収入シートの全行からB列(インデックス1)をキーにデータを辞書化
                    for r_i, r_v in enumerate(income_raw):
                        if r_i <= 7 or not r_v: continue
                        i_name = str(r_v[1]).strip() if len(r_v) > 1 else ""
                        if not i_name: continue
                        
                        i_map = {}
                        # 収入シートの各月と金額をマッピング
                        for col_idx, h_str in enumerate(income_headers):
                            month_str = _normalize(_clean_val(str(h_str)))
                            if "月" in month_str and col_idx < len(r_v):
                                val = r_v[col_idx]
                                if val not in [None, ""]:
                                    i_map[month_str] = val
                        if i_map:
                            income_data_map[i_name] = i_map
                            
                    if income_data_map:
                        # 支払管理の89行目〜94行目を走査してマッピング (0-based index 88 to 93)
                        for target_r in range(88, 94):
                            # H列(インデックス7番)の収入項目名を取得
                            if target_r < len(pay_formatted):
                                target_row_vals = pay_formatted[target_r]
                                p_name = str(target_row_vals[7]).strip() if len(target_row_vals) > 7 else ""
                                
                                if p_name in income_data_map:
                                    matched_map = income_data_map[p_name]
                                    
                                    # 現在の支払管理シートの該当行データをベースにする (pay_raw利用で数式など維持)
                                    base_row = list(pay_raw[target_r]) if target_r < len(pay_raw) else []
                                    if len(base_row) < header_len:
                                        base_row.extend([""] * (header_len - len(base_row)))
                                        
                                    has_update = False
                                    # I列(インデックス8番)以降の対象月のみを上書き
                                    for c_idx in range(8, header_len):
                                        h_str = actual_headers[c_idx] if c_idx < len(actual_headers) else ""
                                        clean_h = _normalize(_clean_val(str(h_str)))
                                        alt_h = clean_h.replace("月", "") + "月"
                                        
                                        if clean_h in matched_map:
                                            base_row[c_idx] = matched_map[clean_h]
                                            has_update = True
                                        elif alt_h in matched_map:
                                            base_row[c_idx] = matched_map[alt_h]
                                            has_update = True
                                            
                                    if has_update:
                                        # 行全体をA列から書き戻す
                                        safe_gspread_call(ws_pay.update, values=[base_row], range_name=f"A{target_r + 1}", value_input_option='USER_ENTERED')
        except Exception as e:
            print(f"Income Sync error: {e}")

        return True, "変動費データの更新が完了しました！"
    except Exception as e:
        return False, f"書き込みエラー: {e}"

def show_variable_cost_update():
    """変動費データ更新 UI"""
    import streamlit as st
    st.markdown("## 🔄 変動費データ更新")
    st.info("「支払方法マスター」に登録されたクレジットカードと、各レシート取引の履歴をもとに、変動費エリアを自動集計して「支払管理」シートを更新します。")
    
    username = st.session_state.get("username", "")
    from app import get_gspread_client
    client = get_gspread_client()
    
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
    except Exception as e:
        st.warning("現在、あなた専用の支払管理シートが見つかりません。先に「支払管理シート新規作成」を行ってください。")
        return
        
    if st.button("更新を実行する", type="primary"):
        with st.spinner("クレジットカードの利用履歴を集計中...（完了まで数秒〜数十秒かかります）"):
            success, msg = execute_variable_cost_update(username)
            if success:
                st.success(msg)
                st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")
            else:
                st.error(msg)
