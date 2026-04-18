import streamlit as st
import pandas as pd
from datetime import datetime, timezone, timedelta
JST = timezone(timedelta(hours=+9), 'JST')
from dateutil.relativedelta import relativedelta
import gspread
import jpholiday
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
    
    # 【Ver 6.2.0】環境（ロケール）による日付形式の差異を吸収するため、Pandasの柔軟な解析を導入
    try:
        # 日本語が含まれる場合は置換してからパース
        s_clean = s.replace("年","/").replace("月","/").replace(".","/")
        if re.search(r"\d{4}", s_clean):
            dt = pd.to_datetime(s_clean, errors='coerce')
            if pd.notnull(dt) and dt.year > 1900:
                return (dt.year, dt.month)
    except:
        pass

    # フォールバック: 正規表現による抽出 (従来どおり)
    # YYYY/M, YYYY/MM, YYYY.M, YYYY.MM, YYYY年M月 などに対応
    m = re.search(r"(\d{4})[年/\.\-](\d{1,2})", s)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    # US形式 (M/D/YYYY) への対応
    m_us = re.search(r"(\d{1,2})[/\.\-](\d{1,2})[/\.\-](\d{4})", s)
    if m_us:
        return (int(m_us.group(3)), int(m_us.group(1)))
        
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
    
    # 日付形式の揺れを吸収 (2026/03 -> 2026.3月)
    # 日本語の「月」がついている場合は除去して判定、最後に統一形式にする
    y_m = _get_year_month(s)
    if y_m != (9999, 12):
        return f"{y_m[0]}.{y_m[1]}月"

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

def createBackup(ss):
    from app import safe_gspread_call
    try:
        bk_name = "支払管理BK"
        try:
            old_bk = ss.worksheet(bk_name)
            safe_gspread_call(ss.del_worksheet, old_bk)
        except Exception:
            pass # 既存の同名バックアップが存在しない場合は無視
            
        ws_pay = next((ws for ws in ss.worksheets() if ws.title.strip() == "支払管理"), None)
        if ws_pay:
            safe_gspread_call(ws_pay.duplicate, new_sheet_name=bk_name)
        return True
    except Exception as e:
        print(f"Backup Error: {e}")
        return False


def execute_expansion(username, mode="NEW", start_ym=None):
    from app import get_gspread_client, safe_gspread_call, get_payment_methods
    from fixed_cost_expansion import _find_val, _clean_val, _normalize, _get_year_month
    import re
    from dateutil.relativedelta import relativedelta
    
    client = get_gspread_client()
    if not client: return False, "Google Drive APIに接続できません。"
    
    try:
        ss = client.open(f"{username}_支払管理")
        ws_master = ss.worksheet("固定費マスター")
        ws_pay = next((ws for ws in ss.worksheets() if ws.title.strip() == "支払管理"), None)
        if not ws_pay: return False, "支払管理シートが見つかりません。"
        
        # --- A. 固定費データ展開の事前バックアップ ---
        createBackup(ss)

        # --- プロフィール連動（生年月日・現在日時）の設定 ---
        try:
            from app import get_sheet, USER_MASTER_WORKSHEET_NAME
            from datetime import datetime
            
            # 1. 生年月日の取得 (User_Masterシートから)
            # User_Masterは共有シート Kakeibo_Data にあるため get_sheet を使用
            master_sheet = get_sheet(USER_MASTER_WORKSHEET_NAME)
            user_records = safe_gspread_call(master_sheet.get_all_records)
            user_profile = next((r for r in user_records if str(r.get("username", "")).lower() == username.lower()), None)
            
            if user_profile:
                birth_raw = str(user_profile.get("birthdate", "")).strip()
                # 数字のみを抽出して YYYYMMDD に整形 (例: 1958/12/12 -> 19581212)
                birth_clean = "".join(filter(str.isdigit, birth_raw))
                if len(birth_clean) >= 8:
                    # 値のみを書き込み（書式は維持される）
                    safe_gspread_call(ws_pay.update_acell, "F2", birth_clean[:8])
            
            # 2. 現在日時の設定 (yyyy-mm-dd hh:mm:ss)
            now_str = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
            safe_gspread_call(ws_pay.update_acell, "F4", now_str)
            
        except Exception as profile_err:
            # プロフィール連携の失敗はメイン処理を止めないようログ出力に留める
            print(f"Profile Sync Error: {profile_err}")
    except Exception as e:
        return False, f"シート読み込みエラー: {e}"
        
    master_raw = safe_gspread_call(ws_master.get_all_values)
    if not master_raw or len(master_raw) < 2:
        return True, "固定費マスターにデータがないため、展開をスキップしました。"
    headers = master_raw[0]
    master_data = [dict(zip(headers, r + [""]*(len(headers)-len(r)))) for r in master_raw[1:]]
    
    # GET PAYMENT METHODS FOR OFFSET CALCULATION
    methods = safe_gspread_call(get_payment_methods, username)
    
    # GET PAYMENT SHEET DATA
    pay_formatted = safe_gspread_call(ws_pay.get_all_values, value_render_option="FORMATTED_VALUE")
    
    h_row_idx = -1
    for i_r, r_v in enumerate(pay_formatted):
        if r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
            h_row_idx = i_r
            break
    if h_row_idx == -1: h_row_idx = 6
    
    actual_headers = pay_formatted[h_row_idx]
    
    # GET COLUMNS BY MONTH
    month_cols = []
    for i, h in enumerate(actual_headers):
        h_clean = _clean_val(h).strip()
        y_m = _get_year_month(h_clean)
        if y_m != (9999, 12):
            month_cols.append({"col_idx": i, "year": y_m[0], "month": y_m[1], "ym": y_m[0]*100 + y_m[1]})
            
    # FIND INDICES
    cat1_idx = next((i for i, h in enumerate(actual_headers) if "大分類" in h or "カテゴリ1" in h), -1)
    k1_idx = next((i for i, h in enumerate(actual_headers) if "科目1" in h or "科目１" in h or "固定支払1" in h or "固定支払１" in h), -1)
    v_f_idx = next((i for i, h in enumerate(actual_headers) if "変動" in h and "固定" in h), -1)
    f_i_idx = next((i for i, h in enumerate(actual_headers) if "有限" in h and "無限" in h), -1)
    k2_idx = next((i for i, h in enumerate(actual_headers) if "科目2" in h or "科目２" in h), -1)
    sno_idx = next((i for i, h in enumerate(actual_headers) if "Sno" in h or "seq" in h.lower()), -1)
    det_idx = next((i for i, h in enumerate(actual_headers) if "詳細" in h or "明細" in h), -1)
    
    if det_idx == -1: return False, "支払管理シートに「科目明細」列が見つかりません。"
    
    # Tracking allocated empty rows to avoid putting multiple master items on the same empty row
    allocated_empty_rows = set()
    
    requests = []
    
    for m_rec in master_data:
        m_detail = _clean_val(_find_val(m_rec, ["詳細", "明細"])).strip()
        m_k1 = _clean_val(_find_val(m_rec, ["科目1", "科目１", "固定支払1", "固定支払１"])).strip()
        m_sno = _clean_val(_find_val(m_rec, ["Sno", "seq"])).strip()
        
        if not m_detail or not m_k1: continue
        
        # FIND TARGET ROW IN PAY SHEET
        # Priority 1: Exact text match in Detail column
        target_r_idx = -1
        for r_i, r_v in enumerate(pay_formatted):
            if r_i <= h_row_idx: continue
            if len(r_v) > det_idx and _clean_val(r_v[det_idx]).strip() == m_detail:
                target_r_idx = r_i
                break
                
        # Priority 2: Fallback to matching K1 + matching Sno, ONLY IF DET is EMPTY
        if target_r_idx == -1 and k1_idx != -1 and sno_idx != -1:
            for r_i, r_v in enumerate(pay_formatted):
                if r_i <= h_row_idx or r_i in allocated_empty_rows: continue
                if len(r_v) > k1_idx and _clean_val(r_v[k1_idx]).strip() == m_k1:
                    if len(r_v) > sno_idx and _clean_val(r_v[sno_idx]).strip() == m_sno:
                        r_det = _clean_val(r_v[det_idx]).strip() if len(r_v) > det_idx else ""
                        if not r_det:
                            target_r_idx = r_i
                            allocated_empty_rows.add(r_i)
                            break
                            
        # Priority 3: Fallback to matched K1 and first EMPTY Det slot
        if target_r_idx == -1 and k1_idx != -1:
            for r_i, r_v in enumerate(pay_formatted):
                if r_i <= h_row_idx or r_i in allocated_empty_rows: continue
                if len(r_v) > k1_idx and _clean_val(r_v[k1_idx]).strip() == m_k1:
                    r_det = _clean_val(r_v[det_idx]).strip() if len(r_v) > det_idx else ""
                    if not r_det:
                        target_r_idx = r_i
                        allocated_empty_rows.add(r_i)
                        break
                
        if target_r_idx == -1:
            continue # NO AVAILABLE ROW - SKIP
            
        # Populate metadata (Category, Subject2, V/F, F/I) and Detail string itself
        m_cat1 = _clean_val(_find_val(m_rec, ["大分類", "カテゴリ1"])).strip()
        m_vf = _clean_val(_find_val(m_rec, ["変動or固定"])).strip()
        m_fi = _clean_val(_find_val(m_rec, ["有限or無限"])).strip()
        m_k2 = _clean_val(_find_val(m_rec, ["科目2", "科目２"])).strip()

        meta_updates = []
        if cat1_idx != -1 and m_cat1: meta_updates.append((cat1_idx, m_cat1))
        if v_f_idx != -1 and m_vf:   meta_updates.append((v_f_idx, m_vf))
        if f_i_idx != -1 and m_fi:   meta_updates.append((f_i_idx, m_fi))
        if k2_idx != -1 and m_k2:    meta_updates.append((k2_idx, m_k2))
        
        # Detail column
        meta_updates.append((det_idx, m_detail))

        for idx, val in meta_updates:
            # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
            if (target_r_idx + 1) in [1, 77]:
                continue
            if (target_r_idx + 1) in [54, 57, 60] and idx >= 6:
                continue
            current_val = pay_formatted[target_r_idx][idx] if idx < len(pay_formatted[target_r_idx]) else ""
            if str(current_val).strip() != val:
                # Calculate column letter (A, B, ..., Z, AA, AB, ...)
                if idx < 26:
                    col_letter = chr(ord("A") + idx)
                else:
                    col_letter = chr(ord("A") + idx // 26 - 1) + chr(ord("A") + idx % 26)
                
                requests.append({
                    "range": f"支払管理!{col_letter}{target_r_idx + 1}",
                    "values": [[val]]
                })
            
        # Values extraction
        amt_str = str(_find_val(m_rec, ["支払額", "金額"], exclude=["最終月額", "最終"])).replace(",", "").replace("¥", "").replace("￥", "")
        amt = amt_str.strip()
        final_amt_str = str(_find_val(m_rec, ["最終月額"])).replace(",", "").replace("¥", "").replace("￥", "").strip()
        final_amt = final_amt_str if final_amt_str else amt
        
        start_m_str = str(_find_val(m_rec, ["開始"])).strip()
        end_m_str = str(_find_val(m_rec, ["完済", "終了", "完了"])).strip()
        is_finite_str = str(_find_val(m_rec, ["有限", "無限"]))
        pay_month_freq = str(_find_val(m_rec, ["支払月", "頻度"])).strip()
        
        is_finite = "有限" in is_finite_str
        sy, sm = _get_year_month(start_m_str) if start_m_str else (0,0)
        ey, em = _get_year_month(end_m_str) if (is_finite and end_m_str) else (9999,12)
        start_ym_val = sy * 100 + sm
        end_ym_val = ey * 100 + em
        
        # --- クレジットカード支払日に基づくオフセット（シフト）判定 ---
        offset = 0
        if _normalize(m_k1) == "クレジットカード":
            # 支払方法マスターから該当カードの設定を探す（正規化して比較）
            m_k2_norm = _normalize(m_k2)
            method = next((m for m in methods if _normalize(m.get("name", "")) == m_k2_norm), None)
            
            if method:
                c_date = _normalize(method.get("closing_date", ""))
                p_month = _normalize(method.get("payment_month", ""))
                p_date_str = _normalize(method.get("payment_date", ""))
                
                # 支払日の数値抽出と判定
                p_day_match = re.search(r"\d+", p_date_str)
                p_day_val = int(p_day_match.group()) if p_day_match else 0
                is_pay_late = (p_day_val >= 20 or "末日" in p_date_str or "月末" in p_date_str)
                
                # 【Ver 6.0.0 変更】判定条件を20日基準（is_pay_late）のみに簡素化・統一
                if is_pay_late:
                    offset = 1
                
                # ユーザー確認用ログの出力 (Ver 6.1.0 診断強化)
                log_msg = f"  - 💳 カード: **{m_k2}** (判定用支払日: `{p_day_val}`, 締日: `{c_date}`, 支払月: `{p_month}`, 支払日内容: `{p_date_str}`)"
                if offset > 0:
                    st.write(f"{log_msg} → ⚡ **1ヶ月シフト適用**")
                else:
                    st.write(f"{log_msg} → シフトなし")

        # Months check
        for m_col in month_cols:
            c_idx = m_col["col_idx"]
            col_ym = m_col["ym"]
            c_m = m_col["month"]
            
            # UIから指定された開始月以前のデータ更新はスキップする
            if start_ym and col_ym < start_ym:
                continue
                
            # オフセット適用後の「利用月」ベースで判定を行う
            # 展開先の月(col_ym)からオフセット分を引いた月を利用月とする
            usage_date = datetime(m_col["year"], m_col["month"], 1) - relativedelta(months=offset)
            usage_ym = usage_date.year * 100 + usage_date.month
            usage_m = usage_date.month

            val_to_set = ""
            if usage_ym >= start_ym_val and (not is_finite or usage_ym <= end_ym_val):
                months_targeted = []
                if "偶数" in pay_month_freq:
                    months_targeted = [2, 4, 6, 8, 10, 12]
                elif "奇数" in pay_month_freq:
                    months_targeted = [1, 3, 5, 7, 9, 11]
                else:
                    mm = re.findall(r"\d+", pay_month_freq)
                    if mm:
                        months_targeted = [int(x) for x in mm]
                    else:
                        months_targeted = [usage_m] # 利用月の月で判定
                
                if usage_m in months_targeted:
                    if is_finite and usage_ym == end_ym_val:
                        val_to_set = final_amt
                    else:
                        val_to_set = amt
            
            current_val = pay_formatted[target_r_idx][c_idx] if c_idx < len(pay_formatted[target_r_idx]) else ""
            normalized_current = str(current_val).replace(",", "").strip()
            
            # 条件不一致月も明示的にクリア（"" をセット）するための判定
            if str(val_to_set) != normalized_current:
                # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
                if (target_r_idx + 1) in [1, 77]:
                    continue
                if (target_r_idx + 1) in [54, 57, 60] and c_idx >= 6:
                    continue
                col_letter = chr(ord("A") + c_idx) if c_idx < 26 else chr(ord("A") + c_idx//26 - 1) + chr(ord("A") + c_idx%26)
                cell_name = f"{col_letter}{target_r_idx + 1}"
                
                # 数値の場合は int 変換、空文字や非数値の場合はそのままセット
                final_val = int(val_to_set) if str(val_to_set).isdigit() else val_to_set
                requests.append({
                    "range": f"支払管理!{cell_name}",
                    "values": [[final_val]]
                })
                
    # --- 【追加仕様】小遣い予算（78行目）の反映 ---
    ozukai_record = None
    for m in master_data:
        k1 = _clean_val(_find_val(m, ["科目1", "科目１", "固定支払1", "固定支払１"])).strip()
        if k1 == "小遣い予算":
            ozukai_record = m
            break
            
    if ozukai_record:
        ozukai_amt_str = str(_find_val(ozukai_record, ["金額", "支払額"])).replace(",", "").replace("¥", "").replace("￥", "").strip()
        ozukai_start_str = str(_find_val(ozukai_record, ["開始", "開始月"])).strip()
        
        s_y, s_m = _get_year_month(ozukai_start_str) if ozukai_start_str else (0, 0)
        start_ym_val = s_y * 100 + s_m
        
        target_r_idx = 77  # 78行目 (0-indexed)
        
        for m_col in month_cols:
            c_idx = m_col["col_idx"]
            col_ym = m_col["ym"]
            
            if start_ym and col_ym < start_ym:
                continue
                
            val_to_set = ""
            if col_ym >= start_ym_val and ozukai_amt_str:
                val_to_set = ozukai_amt_str
            
            # 78行目の現在の値を取得
            current_val = pay_formatted[target_r_idx][c_idx] if target_r_idx < len(pay_formatted) and c_idx < len(pay_formatted[target_r_idx]) else ""
            normalized_current = str(current_val).replace(",", "").strip()
            
            if str(val_to_set) != normalized_current:
                # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
                if (target_r_idx + 1) in [1, 77]:
                    continue
                if (target_r_idx + 1) in [54, 57, 60] and c_idx >= 6:
                    continue
                col_letter = chr(ord("A") + c_idx) if c_idx < 26 else chr(ord("A") + c_idx//26 - 1) + chr(ord("A") + c_idx%26)
                cell_name = f"{col_letter}{target_r_idx + 1}"
                
                final_val = int(val_to_set) if str(val_to_set).isdigit() else val_to_set
                requests.append({
                    "range": f"支払管理!{cell_name}",
                    "values": [[final_val]]
                })

    if requests:
        safe_gspread_call(ss.values_batch_update, {"valueInputOption": "USER_ENTERED", "data": requests})
        
    return True, "データ展開に成功しました！"
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
    import streamlit as st
    st.markdown("## 🛠️ 固定費データ展開")
    st.info("「固定費マスター」の情報をもとに、「支払管理」シートに月別のデータを「値のみ」で安全に展開・追加します。")
    username = st.session_state.get("username", "")
    from app import get_gspread_client, safe_gspread_call
    client = get_gspread_client()
    if not client:
        st.error("Google Drive APIに接続できません。")
        return
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
    except Exception:
        st.warning("現在、あなた専用の支払管理シートが見つかりません。先に『支払管理シート新規作成』を行ってください。")
        return
    
    import datetime
    from dateutil.relativedelta import relativedelta
    now = datetime.datetime.now(JST)
    month_options = [f"{now.year}/{now.month:02d}（当月）"]
    for i in range(1, 6):
        dt = now + relativedelta(months=i)
        month_options.append(f"{dt.year}/{dt.month:02d}")

    st.markdown("### 実行オプション")
    st.markdown("#### A: 全期間の再設定")

    if "confirm_expansion" not in st.session_state:
        st.session_state.confirm_expansion = False

    if not st.session_state.confirm_expansion:
        if st.button("🚀 新規データ展開 or 再作成", type="primary", use_container_width=True):
            st.session_state.confirm_expansion = True
            st.rerun()

    if st.session_state.confirm_expansion:
        st.warning("⚠️ 既にデータ展開されている場合は全て初期化します。本当によろしいですか？")
        col1, col2 = st.columns(2)
        with col1:
            confirm = st.button("はい、実行します", type="primary", use_container_width=True)
        with col2:
            cancel = st.button("キャンセル（いいえ）", use_container_width=True)

        if confirm:
            st.session_state.confirm_expansion = False
            with st.spinner("データ展開中...（数秒〜数十秒かかります）"):
                success, msg = execute_expansion(username)
                if success:
                    with st.spinner("続けて変動費（クレジットカード利用等）を集計中..."):
                        v_success, v_msg = execute_variable_cost_update(username, skip_backup=True)
                        if v_success:
                            st.success("固定費データの展開と変動費の集計が完了しました！")
                        else:
                            st.warning(f"固定費展開は完了しましたが、変動費更新でエラーが発生しました: {v_msg}")
                    if ss:
                        st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")
                else:
                    st.error(msg)
        elif cancel:
            st.session_state.confirm_expansion = False
            st.rerun()

    st.markdown("---")
    st.markdown("#### B: 指定月以降のデータ展開（既存データ保持）")
    st.info("選択した月以降のデータのみを更新し、過去の確定データは保護します。")
    
    colA, colB = st.columns([1, 1])
    with colA:
        selected_month_str = st.selectbox("更新を開始する月を選択", options=month_options, label_visibility="collapsed")
    with colB:
        if st.button("🔄 翌月以降データ展開", use_container_width=True):
            clean_str = selected_month_str.replace("（当月）", "")
            y, m = clean_str.split("/")
            s_ym = int(y) * 100 + int(m)
            with st.spinner(f"{clean_str} 以降のデータを展開中..."):
                success, msg = execute_expansion(username, start_ym=s_ym)
                if success:
                    with st.spinner("続けて変動費を集計中..."):
                        v_success, v_msg = execute_variable_cost_update(username, start_ym=s_ym, skip_backup=True)
                        if v_success:
                            st.success(f"{clean_str} 以降の固定費展開と変動費更新が完了しました！")
                        else:
                            st.warning(f"固定費展開は完了しましたが、変動費更新でエラーが発生しました: {v_msg}")
                    if ss:
                        st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")
                else:
                    st.error(msg)
def execute_variable_cost_update(username, start_ym=None, skip_backup=False):
    from app import get_gspread_client, safe_gspread_call, get_payment_methods, get_sheet, TRANSACTIONS_WORKSHEET_NAME
    import calendar
    from datetime import datetime, timezone, timedelta
    from dateutil.relativedelta import relativedelta
    
    client = get_gspread_client()
    if not client:
        return False, "Google Docsへの接続に失敗しました。"
        
    sheet_name = f"{username}_支払管理"
    try:
        ss = client.open(sheet_name)
        # Handle possible trailing or leading spaces in the tab name
        ws_pay = next((ws for ws in ss.worksheets() if ws.title.strip() == "支払管理"), None)
        if not ws_pay:
            raise Exception("支払管理シートが見つかりません。")
            
        if not ws_pay:
            raise Exception("支払管理シートが見つかりません。")
            
    except Exception as e:
        return False, f"支払管理シート({sheet_name})が見つかりません。先に「支払管理シート新規作成」を実行してください。"
        
    # --- B. 変動費データ更新 単独実行時のバックアップ作成 ---
    if not skip_backup:
        createBackup(ss)

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
    st.write("🔍 変動費更新のためのヘッダー解析中...")
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
        return True, "クレジットカードが登録されていないため、変動費の集計をスキップしました。"
        
    st.write("📊 取引履歴を取得し、各月の支払額を集計中...")
    try:
        tx_sheet = get_sheet(TRANSACTIONS_WORKSHEET_NAME)
        all_txs = safe_gspread_call(tx_sheet.get_all_records)
        user_txs = [tx for tx in all_txs if str(tx.get("username", "")).lower() == username.lower()]
        if not user_txs:
            return True, "対象となる取引データがないため、変動費の更新をスキップしました。"
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

    st.write("📊 クレジットカード支払情報を集計中...")
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
        timing_label = "翌月" if p_day_val < 20 else "当月"
        # 【Ver 6.0.0】is_pay_late をこのスコープで正しく定義
        is_pay_late = (p_day_val >= 20 or "末日" in str(pay_date_str) or "月末" in str(pay_date_str))
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
                # 【Ver 6.0.0】堅牢な解析への切り替え (点やスラッシュの違いに左右されない)
                my, mm = _get_year_month(_clean_val(mc).strip())
                if my == 9999:
                    continue
                base_date = datetime(my, mm, 1)
                
                # 表示ルールの適用: 20日基準で表示年月をシフト
                # 20日未満なら：支払月（表示月） ＝ 利用月
                # 20日以降なら：支払月（表示月） ＝ 利用月 ＋ 1ヶ月  （利用月 ＝ 表示月 － 1ヶ月）
                try:
                    p_day_m = re.search(r"\d+", str(pay_date_str))
                    p_day = int(p_day_m.group()) if p_day_m else 0
                except:
                    p_day = 0
                
                # 改訂ルール#8：
                # 通常：利用月 ＝ 支払月（表示月）
                # 遅い（20日以降）：利用月 ＝ 支払月（表示月） － 1ヶ月
                # 【Ver 6.0.0】未定義だった is_pay_late を使用
                if not is_pay_late:
                    target_pay_date = base_date + relativedelta(months=1)
                else:
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
                    amt_val = str(amt).strip().replace(",", "").replace("¥", "").replace("￥", "")
                    has_valid_amount = False
                    if amt_val and amt_val != "-":
                        try:
                            if float(amt_val) > 0:
                                has_valid_amount = True
                        except: pass

                    is_past_due = False
                    y, m = _get_year_month(mc)
                    if y != 9999 and has_valid_amount:
                        off = 1
                        if "当月" in str(pay_month_str): off = 0
                        elif "翌々月" in str(pay_month_str): off = 2
                        
                        # 【例外ルール】変動費カード(50-52行)で「翌月20日以降払」の場合は当月(0)とみなす
                        d_m_check = re.search(r"\d+", str(pay_date_str))
                        p_day_val = int(d_m_check.group()) if d_m_check else 0
                        is_late = (p_day_val >= 20 or "末日" in str(pay_date_str) or "月末" in str(pay_date_str))
                        if off == 1 and is_late:
                            off = 0
                                
                        base_dt = datetime(y, m, 1) + relativedelta(months=off)
                        due_dt = base_dt + relativedelta(day=p_day_val)
                        
                        # 営業日チェックと翌営業日へのスライド
                        while due_dt.weekday() >= 5 or jpholiday.is_holiday(due_dt.date()):
                            due_dt += timedelta(days=1)
                            
                        if datetime.now(JST).date() >= due_dt.date():
                            is_past_due = True

                    if has_valid_amount and is_past_due:
                        if 0 <= f_idx < len(row_arr):
                            row_arr[f_idx] = "1"
                    else:
                        if 0 <= f_idx < len(row_arr):
                            row_arr[f_idx] = ""
                except Exception as ex: 
                    print(f"Error in var flag: {ex}")
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
        
    # pass # FIXED_FORMAT_MAINTENANCE: Do not append dynamic var_total_row
    # cc_rows_array.append(var_total_row)
    # current_row_num += 1
    
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
                        if p_day < 20:
                            target_pay_dt = base_date_dt + relativedelta(months=1)
                        else:
                            target_pay_dt = base_date_dt
                            
                        from app import calculate_credit_card_periods
                        t_periods = calculate_credit_card_periods(target_pay_dt, closing_str, pay_month_str, pay_date_str)
                        if t_periods:
                            actual_due_date = t_periods[0]["pay_date"]
                            if datetime.now(JST).date() >= actual_due_date:
                                # B列開始なのでインデックスを調整し、境界チェックを追加
                                if 0 <= f_idx < len(r_fc):
                                    r_fc[f_idx] = 1
                except: pass
            except: pass
        # pass # FIXED_FORMAT_MAINTENANCE: Do not append dynamic r_fc
        # cc_rows_array.append(r_fc)
        # current_row_num += 1
        
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
                            if datetime.now(JST).date() >= actual_due_date.date():
                                # B列開始なのでインデックスを調整し、境界チェックを追加
                                if 0 <= f_idx < len(r_vc):
                                    r_vc[f_idx] = 1
                except: pass
            except: pass
        # pass # FIXED_FORMAT_MAINTENANCE: Do not append dynamic r_vc
        # cc_rows_array.append(r_vc)
        # current_row_num += 1
        
        # 3. 合計 (固定費分+変動費分)
        # 支払日の文言判定
        try:
            p_day_sum_m = re.search(r"\d+", str(pay_date_str))
            p_day_sum_v = int(p_day_sum_m.group()) if p_day_sum_m else 27
        except: p_day_sum_v = 27
        timing_sum_label = "翌月" if p_day_sum_v < 20 else "当月"
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
        # pass # FIXED_FORMAT_MAINTENANCE: Do not append dynamic r_sum
        # cc_rows_array.append(r_sum)
        # card_total_rows.append(current_row_num)
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
    # pass # FIXED_FORMAT_MAINTENANCE: Do not append dynamic r_grand
    # cc_rows_array.append(r_grand)
    # current_row_num += 1
    
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
    # pass # FIXED_FORMAT_MAINTENANCE: Do not append dynamic r_pay_total
    # cc_rows_array.append(r_pay_total)
    # current_row_num += 1
    
    summary_end_row = current_row_num - 1
    
    header_len = len(actual_headers)
    
    # Append the variable cost area array
    try:
        # 古い変動費（合計より下、小遣いより上）の値をクリア（書式は維持）
        current_rows = ws_pay.row_count
        if start_row_num <= current_rows:
            clear_end_row = min(boundary_row - 1, current_rows)
            pass # FIXED_FORMAT_MAINTENANCE: Do not completely clear B:ZZ bounds.
            # if start_row_num <= clear_end_row:
            #     safe_gspread_call(ws_pay.batch_clear, [f"B{start_row_num}:ZZ{clear_end_row}"])

        # 空間の確保と書き込み (小遣い行を押し出す)
        data_count = len(cc_rows_array)
        available_space = boundary_row - start_row_num
        
        if data_count > available_space:
            diff = data_count - available_space
        pass # FIXED_FORMAT_REMOVED_NO_INSERTIONS
            
        # 新しい変動費データを書き込み
        st.write("💾 変動費集計結果をシートに書き込み中...")

        # --- 追加仕様: 50-52行のG列(index 6)とF列(index 5)の対応マップ作成 ---
        g_to_f_map = {}
        for r_idx_map in [49, 50, 51]: # 50, 51, 52行目 (0-indexed)
            if r_idx_map < len(pay_raw):
                row_map = pay_raw[r_idx_map]
                if len(row_map) > 6:
                    g_val = str(row_map[6]).strip() # G列
                    f_val = str(row_map[5]).strip() # F列
                    if g_val:
                        g_to_f_map[g_val] = f_val
        
        # 50行から52行のE列(インデックス4)とG列(インデックス6)への値の設定をしないように既存の値を復元
        # 54行から63行のH列(インデックス7)への値の設定をしないように既存の値を復元
        for i, r in enumerate(cc_rows_array):
            target_r_idx = start_row_num - 1 + i
            sheet_row_num = target_r_idx + 1
            if target_r_idx < len(pay_raw):
                raw_row = pay_raw[target_r_idx]
                
                # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
                if sheet_row_num in [1, 77]:
                    cc_rows_array[i] = raw_row
                    continue
                if sheet_row_num in [54, 57, 60]:
                    # G列(index 6)以降を元の値で復元（更新させない）
                    for col_idx in range(6, len(r)):
                        if col_idx < len(raw_row):
                            r[col_idx] = raw_row[col_idx]

                if sheet_row_num in [50, 51, 52]:
                    if len(raw_row) > 4 and len(r) > 4:
                        r[4] = raw_row[4]
                    if len(raw_row) > 6 and len(r) > 6:
                        r[6] = raw_row[6]
                
                if 54 <= sheet_row_num <= 62:
                    # G列 (index 6, Sno) を元の値から復元（マッピングのキーとして使用するため、全行で復元）
                    if len(raw_row) > 6 and len(r) > 6:
                        r[6] = raw_row[6]
                    
                    # --- 追加仕様: G列の値に基づいて50-52行のF列の値を設定 ---
                    if len(r) > 6:
                        current_g = str(r[6]).strip()
                        if current_g in g_to_f_map:
                            if len(r) > 5:
                                r[5] = g_to_f_map[current_g]

                    # H列 (index 7, 科目明細) は既存の値を維持
                    if len(raw_row) > 7 and len(r) > 7:
                        r[7] = raw_row[7]
                    
                    # --- 【追加仕様】クレジットカード内訳集計ロジック ---
                    # H列に含まれる「固定費」または「変動費」キーワードを特定
                    h_val = str(r[7]).strip()
                    cost_type_keyword = ""
                    if "固定費" in h_val: cost_type_keyword = "固定費"
                    elif "変動費" in h_val: cost_type_keyword = "変動費"
                    
                    # F列(Index 5)のカード名を取得
                    card_name = str(r[5]).strip()
                    
                    if cost_type_keyword and card_name:
                        # 8行〜52行 (Index 7〜51) を走査して集計
                        # 既に取得済みの pay_formatted (書式付き値) を使用
                        source_data = pay_formatted[7:52] if len(pay_formatted) > 7 else []
                        
                        for m_col in month_cols:
                            c_idx = m_col["col_idx"]
                            total_sum = 0
                            for src_row in source_data:
                                if len(src_row) > 3:
                                    src_cat = str(src_row[0]).strip() # A列 (大分類)
                                    src_k2 = str(src_row[3]).strip()  # D列 (科目２)
                                    if cost_type_keyword in src_cat and src_k2 == card_name:
                                        # 金額を取得して加算
                                        amt_val = str(src_row[c_idx]).strip()
                                        total_sum += safe_money_int_cast(amt_val)
                            
                            # 集計値をセット (0の場合は空文字列にして「値のみ」で上書き)
                            if 0 <= c_idx < len(r):
                                r[c_idx] = total_sum if total_sum > 0 else ""
                
                if sheet_row_num == 63:
                    if len(raw_row) > 7 and len(r) > 7:
                        r[7] = raw_row[7]
        
        # A列をスキップしてB列から書き込み (Ver 4.27.4: r[1:] で A列を除外)
        safe_gspread_call(ws_pay.update, values=[r[1:] for r in cc_rows_array], range_name=f"B{start_row_num}", value_input_option='USER_ENTERED')
        
        # 既存の固定費サブ合計行も新方式の数式（完了F対応）に更新する
        if fixed_subtotals:
            update_data = [] # List of {'range': ..., 'values': [[...]]}
            # 1. 各サブセクションの集計行を更新
            for row_num, s_row, e_row in fixed_subtotals:
                if 54 <= row_num <= 63:
                    continue # FIXED_FORMAT_MAINTENANCE: Do not touch total rows here
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

                # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
                if row_num in [1, 77]:
                    continue
                if row_num in [54, 57, 60]:
                    # 月次データのカラム(c_idx >= 6)を構築する際に元の値を維持する必要があるが、
                    # ここでは formula を生成しているので、もし row_num が 54, 57, 60 なら
                    # c_idx >= 6 の更新を個別にスキップする。
                    pass 
                
                if row_num not in [1, 77]:
                    update_data.append({'range': f"A{row_num}", 'values': [row_vals]})

            # 2. グランド合計（固定費合計）行の更新 (Ver 5.0.1 循環参照回避ロジック)
            if grand_total_row_num != -1 and not (54 <= grand_total_row_num <= 63):
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
                            formula = f"=SUM({','.join(inner_cells)})"
                            if 0 <= c_idx < len(row_vals):
                                # 54, 57, 60行目のG列以降は更新しない
                                if row_num in [54, 57, 60] and c_idx >= 6:
                                    if c_idx < len(original_row):
                                        row_vals[c_idx] = original_row[c_idx]
                                else:
                                    row_vals[c_idx] = formula

                # メタデータ保持
                original_row = pay_raw[grand_total_row_num - 1]
                for i in range(len(actual_headers)):
                    if i < len(row_vals) and i < len(original_row): row_vals[i] = original_row[i]
                
                for j in range(7):
                    if j < len(row_vals) and ("【合計】" in str(row_vals[j]) or "固定費合計" in str(row_vals[j]) or "_支払残_" in str(row_vals[j])):
                        row_vals[j] = "固定費_支払残_合計額"
                        # A列は空のまま（後でB-H結合するため）
                        row_vals[0] = ""
                        break

                # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
                if grand_total_row_num not in [1, 77]:
                    # grand_total_row_num が 54, 57, 60 の場合の個別制御は上記ループ内(c_idx)で実施
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
                    
                    # --- 【追加仕様】内訳エリア（54行〜62行）は完了F更新対象外 ---
                    # ユーザー指定により 54, 57, 60行目も個別に保護
                    if 54 <= r_idx + 1 <= 62:
                        continue
                    
                    is_cc = False
                    is_bw = False
                    current_info = None
                    
                    # 科目１または大分類で判定
                    r_k1 = str(row[idx_k1]).strip() if idx_k1 != -1 else ""
                    r_k2 = str(row[idx_k2]).strip() if idx_k2 != -1 else ""
                    r_det = str(row[idx_det]).strip() if idx_det != -1 else ""
                    r_dai = str(row[idx_dai]).strip() if idx_dai != -1 else ""
                    
                    if r_k1 == "クレジットカード" or r_dai == "変動費": is_cc = True
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
                                
                                # UIから指定された開始月以前のフラグ更新はスキップする
                                y_mc, m_mc = _get_year_month(mc)
                                col_ym_val = y_mc * 100 + m_mc
                                if start_ym and col_ym_val < start_ym:
                                    continue
                                
                                f_idx = c_idx + 1
                                if f_idx >= len(row_to_update): continue
                                
                                current_flag = str(row_to_update[f_idx]).strip()
                                
                                # FORMULA文字列を避けるため、画面表示状態(pay_formatted)を取得して判定する
                                fmt_val = ""
                                if r_idx < len(pay_formatted) and c_idx < len(pay_formatted[r_idx]):
                                    fmt_val = str(pay_formatted[r_idx][c_idx]).strip().replace(",", "").replace("¥", "").replace("￥", "")
                                
                                # 1. 金額チェック: 空欄ではなく、> 0 の場合のみ True
                                has_valid_amount = False
                                if fmt_val and fmt_val != "-":
                                    try:
                                        if float(fmt_val) > 0:
                                            has_valid_amount = True
                                    except: pass
                                
                                # 2. 日付チェックの準備
                                is_past_due = False
                                y, m = _get_year_month(mc)
                                if y != 9999 and has_valid_amount:
                                    if is_bw:
                                        if "翌月" in p_date: off = 1
                                        elif "当月" in p_date: off = 0
                                        elif "前月" in p_date: off = -1
                                        else: off = 0
                                    else:
                                        off = 1
                                        if "当月" in p_month: off = 0
                                        elif "翌々月" in p_month: off = 2
                                        
                                        # 【重要】変動Ａまたは固定のクレジットカードで、20日以降は「遅」とみなす (Ver 5.9.2)
                                        is_unified_target = (r_dai == "変動費") or (r_dai == "固定費" and r_k1 == "クレジットカード")
                                        if is_unified_target and off == 1:
                                            d_m_check = re.search(r"\d+", p_date)
                                            if d_m_check and int(d_m_check.group()) >= 20:
                                                off = 0
                                                
                                    base_dt = datetime(y, m, 1) + relativedelta(months=off)
                                    d_m = re.search(r"\d+", p_date)
                                    if d_m:
                                        d_v = int(d_m.group())
                                        try:
                                            due_dt = base_dt + relativedelta(day=d_v)
                                            # 営業日チェックと翌営業日へのスライド
                                            while due_dt.weekday() >= 5 or jpholiday.is_holiday(due_dt.date()):
                                                due_dt += timedelta(days=1)
                                                
                                            if datetime.now(JST).date() >= due_dt.date():
                                                is_past_due = True
                                        except: pass

                                # 最終フラグの決定
                                new_flag = "1" if (has_valid_amount and is_past_due) else ""
                                
                                if current_flag != new_flag:
                                    row_to_update[f_idx] = new_flag
                                    row_changed = True
                            except: pass
                        
                        if row_changed:
                            # 【追加仕様】1, 77行目は更新不要。54, 57, 60行目はG列(index 6)以降は更新不要
                            if (r_idx + 1) in [1, 77]:
                                continue
                            
                            # フラグ更新対象がG列以降(f_idx >= 6)かつ対象行ならスキップ
                            # ただしここでは row_to_update 全体を書き込んでいるので、
                            # 対象行の場合はG列以降を元の値に戻す。
                            if (r_idx + 1) in [54, 57, 60]:
                                # pay_raw から元の値(G列以降)を復元
                                if r_idx < len(pay_raw):
                                    orig_row = pay_raw[r_idx]
                                    for col_idx in range(6, len(row_to_update)):
                                        if col_idx < len(orig_row):
                                            row_to_update[col_idx] = orig_row[col_idx]
                            
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
        
        # --- 既存の結合情報を取得 (APIError 400 回復用) ---
        # 構造変更後の最新情報を取得
        existing_merges = []
        try:
            remote_meta = safe_gspread_call(ss.fetch_sheet_metadata)
            current_sheet_info = next((s for s in remote_meta['sheets'] if s['properties']['title'].strip() == "支払管理"), None)
            if current_sheet_info:
                existing_merges = current_sheet_info.get('merges', [])
        except: pass
        
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

        # --- 合計行・集計行の描画 ---
        st.write("📊 クレジットカード内訳・収入データの自動連動および機能保守を実行中...")
        
        # 💡 [Ver 5.9.0] クレジットカード内訳（54-62行）の自動集計および収入項目の動的同期ロジック実装
        if 'existing_merges' in locals() and existing_merges:
            try:
                # 調査・解除対象範囲: 固定費合計 row 以降、全体エリア
                unmerge_start = min(grand_total_row_num - 1 if grand_total_row_num != -1 else 9999, summary_start_row - 1)
                unmerge_end = boundary_row # 小遣い行まで含めて念のため広めに解除
                
                for m_range in existing_merges:
                    m_s = m_range.get('startRowIndex', 0)
                    m_e = m_range.get('endRowIndex', 0)
                    if not (m_e <= unmerge_start or m_s >= unmerge_end):
                        format_requests.append({"unmergeCells": {"range": m_range}})
            except: pass
        # A) 固定費_支払残_合計額 (Blue) B-H合併
        if grand_total_row_num != -1:
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
            pass # FIXED_FORMAT_MAINTENANCE: No formatting or merging allowed.
            # safe_gspread_call(ss.batch_update, {"requests": format_requests})

        # --- 更新日時記録 (F5) ---
        try:
            current_now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
            safe_gspread_call(ws_pay.update, values=[[current_now]], range_name="F5", value_input_option='USER_ENTERED')
        except: pass

        # --- A列(ID列)の保守と数式設定 ---
        ensure_id_column_and_formula(ws_pay)

        # --- 追加仕様: 支払い方法の名前と月別集計金額をH79から順に設定（最大10件） ---
        try:
            from app import get_payment_methods
            pm_list = get_payment_methods(username)
            if pm_list:
                valid_pms = [m for m in pm_list if str(m.get("name", "")).strip()][:10]
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
                                if amt != 0:
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
                                    if sum_amt != 0: val = sum_amt
                                row_data[c_idx - 7] = val
                                
                        update_grid.append(row_data)
                        
                    if update_grid:
                        # 既存の記述が残らないようにH79から該当行のZZ列までをクリア
                        end_row = 79 + len(update_grid) - 1
                        safe_gspread_call(ws_pay.batch_clear, [f"H79:ZZ{end_row}"])
                        
                        safe_gspread_call(ws_pay.update, values=update_grid, range_name="H79", value_input_option='USER_ENTERED')
        
        except Exception as e:
            print(f"H79 Payment Methods Tracking update error: {e}")

        # --- 【追加仕様】「収入」データ連動ロジック（H96〜H100への数式セット） ---
        try:
            ws_income = None
            try:
                ws_income = ss.worksheet("収入")
            except: pass
            
            if ws_income:
                # 文字列として安全に比較するため FORMATTED_VALUE を使用
                income_raw = safe_gspread_call(ws_income.get_all_values, value_render_option='FORMATTED_VALUE')
                if len(income_raw) > 7: # 8行目(index 7)がヘッダー
                    income_headers = income_raw[7]
                    
                    # 1. 収入シートの行番号マッピング（B列）
                    income_row_map = {}
                    for r_i, r_v in enumerate(income_raw):
                        if r_i <= 7 or not r_v: continue
                        i_name = str(r_v[1]).strip() if len(r_v) > 1 else ""
                        if i_name:
                            income_row_map[i_name] = r_i + 1 # 1-based row number
                            
                    # 2. 収入シートの列アルファベットマッピング（8行目ヘッダー）
                    income_col_map = {}
                    for c_idx, h_str in enumerate(income_headers):
                        clean_h = _normalize(_clean_val(str(h_str)))
                        if "月" in clean_h:
                            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                            income_col_map[clean_h] = col_letter
                            
                    # 3. 支払管理シートの照合と数式書き込み (対象行を動的にスキャン)
                    # 以前は96-100行目に固定されていましたが、レイアウトの差異に対応するため全体から検索します
                    for target_r in range(len(pay_formatted)):
                        if target_r < len(pay_formatted):
                            target_row_vals = pay_formatted[target_r]
                            # H列はインデックス7
                            p_name = str(target_row_vals[7]).strip() if len(target_row_vals) > 7 else ""
                            
                            if p_name and p_name in income_row_map:
                                inc_row_num = income_row_map[p_name]
                                
                                # シートの行全体をコピー（元の数式や値を維持）
                                base_row = list(pay_raw[target_r]) if target_r < len(pay_raw) else []
                                if len(base_row) < header_len:
                                    base_row.extend([""] * (header_len - len(base_row)))
                                    
                                has_update = False
                                # I列(インデックス8)以降の対象年月列を探して数式をセット
                                for c_idx in range(8, header_len):
                                    pay_h_str = actual_headers[c_idx] if c_idx < len(actual_headers) else ""
                                    pay_clean_h = _normalize(_clean_val(str(pay_h_str)))
                                    pay_alt_h = pay_clean_h.replace("月", "") + "月"
                                    
                                    # 支払管理の対象年月が、収入シートのヘッダーにあるか
                                    target_col_letter = income_col_map.get(pay_clean_h) or income_col_map.get(pay_alt_h)
                                    
                                    if target_col_letter:
                                        formula_str = f"='収入'!{target_col_letter}{inc_row_num}"
                                        # すでに同じ数式が入っていなければ更新
                                        if str(base_row[c_idx]).strip() != formula_str:
                                            base_row[c_idx] = formula_str
                                            has_update = True
                                            
                                if has_update:
                                    # A列から行全体を上書き(数式として解釈させるため USER_ENTERED を指定)
                                    safe_gspread_call(ws_pay.update, values=[base_row], range_name=f"A{target_r + 1}", value_input_option='USER_ENTERED')
        except Exception as e:
            print(f"Income Sync formula mapping error: {e}")

        # --- 【追加仕様】クレジットカード内訳：カード名の自動同期（54行〜63行） ---
        try:
            # ターゲット範囲の探索用シートデータ（最新版を取得）
            pay_fresh = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
            ws_title = ws_pay.title

            # 1. 50-52行目のG列(index 6)とF列(index 5)の最新対応マップを生成
            # マスター名ではなく、50-52行目の現在の表示名（空白含む）を同期させる (Ver 6.2.25)
            g_to_f_map_fresh = {}
            for r_idx_map in [49, 50, 51]: # 50, 51, 52行目 (0-indexed)
                if r_idx_map < len(pay_fresh):
                    row_map = pay_fresh[r_idx_map]
                    if len(row_map) > 6:
                        g_v = str(row_map[6]).strip()
                        f_v = str(row_map[5]).strip()
                        if g_v:
                            g_to_f_map_fresh[g_v] = f_v
            
            # 2. ターゲット範囲（54行〜63行目）をマッピング
            f_requests = []
            for sheet_row in range(54, 64): # 54行目〜63行目 (1-indexed)
                idx = sheet_row - 1
                if idx < len(pay_fresh):
                    r_raw = pay_fresh[idx]
                    # ターゲット行（内訳エリア）では G列(index 6) に ID が入っている
                    g_val = str(r_raw[6]).strip() if len(r_raw) > 6 else ""
                    h_val = str(r_raw[7]).strip() if len(r_raw) > 7 else ""
                    
                    # 50-52行目から取得した最新のマップを使用
                    target_f_val = g_to_f_map_fresh.get(g_val, None)
                    
                    if g_val and target_f_val is not None:
                        # 書き込み先は F列(index 5)
                        current_f_val = str(r_raw[5]).strip() if len(r_raw) > 5 else ""
                        
                        # カード名に変更がある場合のみ更新リクエストに追加
                        if current_f_val != target_f_val:
                            f_requests.append({
                                "range": f"{ws_title}!F{sheet_row}",
                                "values": [[target_f_val]]
                            })
                            
                    # --- 3. H列の内容(変動費分/固定費分など)に基づく月次データの展開 ---
                    # F列の値（カード名）を取得（上記で確定した値、または現在のセル値）
                    f_val = target_f_val if (g_val and g_val in g_to_f_map_fresh) else (str(r_raw[5]).strip() if len(r_raw) > 5 else "")
                    
                    if f_val and h_val:
                        has_row_update = False
                        
                        # 固定費行の取得 (存在する場合のみ)
                        fc_rows = []
                        if "fc_payment_rows" in locals() and f_val in fc_payment_rows:
                            fc_rows = fc_payment_rows[f_val]
                        
                        # --- 【ユーザー追加仕様】54行〜63行は「変動費分」のみI列以降を更新する（例外はスキップ） ---
                        if 54 <= sheet_row <= 63 and "変動" not in h_val:
                            continue
                            
                        for c_idx in range(8, len(actual_headers)):
                            h_str = actual_headers[c_idx] if c_idx < len(actual_headers) else ""
                            clean_h = _normalize(_clean_val(str(h_str)))
                            alt_h = clean_h.replace("月", "") + "月"
                            
                            # 完了フラグ列は更新対象外としてスキップ
                            if "完了" in clean_h:
                                continue
                                
                            val = ""
                            f_idx = c_idx + 1 # 完了フラグ列インデックス
                            col_letter = chr(ord('A') + c_idx) if c_idx < 26 else chr(ord('A') + c_idx//26 - 1) + chr(ord('A') + c_idx%26)
                            flag_letter = chr(ord('A') + f_idx) if f_idx < 26 else chr(ord('A') + f_idx//26 - 1) + chr(ord('A') + f_idx%26)
                            
                            # H列の内容に応じて出し分け (数値または数式を転記)
                            if "変動" in h_val:
                                # 変動費分
                                val = cc_monthly_amounts.get(f_val, {}).get(clean_h)
                                if val is None or val == "":
                                    val = cc_monthly_amounts.get(f_val, {}).get(alt_h, 0)
                                if val == "": val = 0
                            elif "固定" in h_val:
                                # 固定費分
                                if fc_rows:
                                    cells = [f'IF({flag_letter}{r}="", {col_letter}{r}, 0)' for r in fc_rows]
                                    val = f"=SUM({','.join(cells)})"
                                else:
                                    val = 0
                            elif "計" in h_val or "残" in h_val or "合算" in h_val:
                                # 変動費(値) + 固定費(数式) の合算
                                vc_val = cc_monthly_amounts.get(f_val, {}).get(clean_h)
                                if vc_val is None or vc_val == "":
                                    vc_val = cc_monthly_amounts.get(f_val, {}).get(alt_h, 0)
                                if vc_val == "": vc_val = 0
                                
                                if fc_rows:
                                    fc_formula = "SUM(" + ",".join([f'IF({flag_letter}{r}="", {col_letter}{r}, 0)' for r in fc_rows]) + ")"
                                    val = f"={vc_val} + {fc_formula}"
                                else:
                                    val = vc_val
                                    
                            if val != "":
                                f_requests.append({
                                    "range": f"{ws_title}!{col_letter}{sheet_row}",
                                    "values": [[val]]
                                })
                                has_row_update = True
                            
            if f_requests:
                safe_gspread_call(ss.values_batch_update, {"valueInputOption": "USER_ENTERED", "data": f_requests})
        except Exception as e:
            print(f"Row 54-63 mapping error: {e}")

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
    except Exception:
        st.warning("現在、あなた専用の支払管理シートが見つかりません。先に『支払管理シート新規作成』を行ってください。")
        return
        
    if st.button("更新を実行する", type="primary"):
        with st.spinner("クレジットカードの利用履歴を集計中...（完了まで数秒〜数十秒かかります）"):
            success, msg = execute_variable_cost_update(username)
            if success:
                st.success(msg)
                st.markdown(f"**🔗 [支払管理シートを確認]({ss.url})**")
            else:
                st.error(msg)