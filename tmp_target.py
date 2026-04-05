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
        st.write(f"💾 スプレッドシートへ書き込み中... ({len(final_sheet_array)} 行)")
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
            current_now = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
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
        # Handle possible trailing or leading spaces in the tab name
        ws_pay = next((ws for ws in ss.worksheets() if ws.title.strip() == "支払管理"), None)
        if not ws_pay:
            raise Exception("支払管理シートが見つかりません。")
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
            now_dt = datetime.now(JST)
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
        return False, "クレジットカードが1件も登録されていません。「支払方法マスター」を確認してください。"
        
    st.write("📊 取引履歴を取得し、各月の支払額を集計中...")
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