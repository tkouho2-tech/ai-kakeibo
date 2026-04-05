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
        
    st.write("🔍 シート情報を取得中...")
    try:
        ws_master = ss.worksheet("固定費マスター")
        ws_pay = ss.worksheet("支払管理")
    except Exception as e:
        return False, f"「固定費マスター」または「支払管理」シートが見つかりません。"
        
    st.write("📖 固定費マスターを読み込んでいます...")
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

    st.write("📊 支払管理シートのヘッダーを解析中...")
    # 数式を維持するために FORMULA レンダリングオプションで取得
    pay_raw = safe_gspread_call(ws_pay.get_all_values, value_render_option='FORMULA')
    
    # Extract headers from pay sheet (row 7 -> index 6)
    if len(pay_raw) < 7:
        return False, "「支払管理」のフォーマットが正しくありません（7行目にヘッダーが必要です）。"
        
    # ヘッダー行を「ID」が含まれる行として動的に特定
    h_row_idx = -1
    for i_r, r_v in enumerate(pay_raw):
        if r_v and str(r_v[0]).strip().lower() in ["id", "key"]:
            h_row_idx = i_r
            break
    if h_row_idx == -1: h_row_idx = 6 # フォールバック
    pay_headers = pay_raw[h_row_idx]
    
    # --- ヘッダー・最終月の超堅牢取得 ---
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
            now = datetime.now(JST)
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
        
        st.write(f"🔄 固定費マスターからデータを抽出中... (全 {len(master_data)} 件)")
        sno = 1
        for i_m, m_rec in enumerate(master_data):
            if i_m % 10 == 0 and i_m > 0:
                st.write(f"  ... {i_m} 件目まで処理済み")
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