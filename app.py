import streamlit as st
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. ページの設定
st.set_page_config(page_title="冷凍冷蔵設備 負荷計算ツール", layout="wide")
st.title("❄️ 冷凍冷蔵設備 負荷計算ツール")

# 主要食品の物性値マスターデータ
FOOD_PROPERTIES = {
    "牛肉・豚肉": {"tf": -1.7, "c1": 3.20, "c2": 1.70, "qf": 230.0},
    "鮮魚": {"tf": -1.1, "c1": 3.60, "c2": 2.00, "qf": 270.0},
    "野菜（ほうれん草等）": {"tf": -0.8, "c1": 3.90, "c2": 2.10, "qf": 300.0},
    "水（参考値）": {"tf": 0.0, "c1": 4.19, "c2": 2.05, "qf": 333.5},
}

# 2. 画面を2列に分割（入力エリアと計算結果エリア）
col_input, col_result = st.columns([1, 1])

# --- 左側：入力エリア ---
with col_input:
    st.header("📋 入力条件設定")
    
    # ① 部屋寸法・基本設定
    with st.expander("① 部屋情報・基本設定", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            width = st.number_input("部屋の幅 (m)", value=5.4, step=0.1)
        with c2:
            length = st.number_input("部屋の奥行 (m)", value=3.6, step=0.1)
        with c3:
            height = st.number_input("部屋の高さ (m)", value=2.4, step=0.1)
            
        t_int = st.number_input("庫内温度 T_I (℃)", value=-25.0, step=1.0)
        margin_rate = st.number_input("安全率（余裕率）", value=1.25, step=0.05)
        
        volume = width * length * height
        floor_area = width * length
        capacity_ton = volume * 0.40102
        st.info(f"💡 計算値：床面積 {floor_area:.1f} m² / 内容積 {volume:.1f} m³ / 収容能力 {capacity_ton:.1f} トン")

    # ② 各部断熱・周囲温度設定
    with st.expander("② 各部断熱・周囲温度設定", expanded=True):
        panel_type = st.selectbox("パネル厚み選択", ["冷凍 100mm (K=0.21)", "冷蔵 50mm (K=0.42)", "手入力"])
        k_val = 0.21 if "100mm" in panel_type else (0.42 if "50mm" in panel_type else 0.21)
        
        st.caption("各面の外気・周囲温度 (℃)")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            t_ceiling = st.number_input("天井周囲温度", value=40.0, step=1.0)
            t_floor = st.number_input("床下周囲温度", value=25.0, step=1.0)
        with cc2:
            t_wall_f = st.number_input("壁（正面）温度", value=33.0, step=1.0)
            t_wall_b = st.number_input("壁（後面）温度", value=33.0, step=1.0)
        with cc3:
            t_wall_r = st.number_input("壁（右面）温度", value=33.0, step=1.0)
            t_wall_l = st.number_input("壁（左面）温度", value=33.0, step=1.0)
            
        if panel_type == "手入力":
            k_val = st.number_input("熱通過率 K値 [W/(m²・K)]", value=0.2100, format="%.4f")

    # ③ 【入庫情報】設定
    with st.expander("③ 入庫情報・食品物性設定", expanded=True):
        st.subheader("入庫品の温度・比熱設定")
        
        t_in = st.number_input("入庫品の温度 (℃)", value=5.0, step=1.0)
        food_mode = st.radio("食品比熱の入力設定", ["主要食品から選ぶ（自動入力）", "手動入力"], horizontal=True)
        
        if food_mode == "主要食品から選ぶ（自動入力）":
            selected_food = st.selectbox("主要な食品を選択してください", list(FOOD_PROPERTIES.keys()))
            base_tf = FOOD_PROPERTIES[selected_food]["tf"]
            base_c1 = FOOD_PROPERTIES[selected_food]["c1"]
            base_c2 = FOOD_PROPERTIES[selected_food]["c2"]
            base_qf = FOOD_PROPERTIES[selected_food]["qf"]
        else:
            selected_food = "カスタム"
            base_tf = -2.0
            base_c1 = 3.34880
            base_c2 = 2.09300
            base_qf = 209.3

        if t_in > base_tf:
            st.markdown("*※入庫品が冷蔵帯（凍結点以上）のため、凍結潜熱・前後の比熱を表示します*")
            fc1, fc2 = st.columns(2)
            with fc1:
                t_f = st.number_input("入庫品の凍結点 (℃)", value=base_tf, disabled=(food_mode=="主要食品から選ぶ（自動入力）"))
                c1 = st.number_input("凍結前状態の比熱 [kJ/(kg・K)]", value=base_c1, format="%.5f", disabled=(food_mode=="主要食品から選ぶ（自動入力）"))
            with fc2:
                q_f = st.number_input("入庫品の凍結潜熱 [kJ/kg]", value=base_qf, disabled=(food_mode=="主要食品から選ぶ（自動入力）"))
                c2 = st.number_input("凍結後状態の比熱 [kJ/(kg・K)]", value=base_c2, format="%.5f", disabled=(food_mode=="主要食品から選ぶ（自動入力）"))
        else:
            st.markdown("*※入庫品はすでに冷凍状態のため、凍結後の比熱のみを使用します*")
            c2 = st.number_input("凍結後状態の比熱 [kJ/(kg・K)]", value=base_c2, format="%.5f", disabled=(food_mode=="主要食品から選ぶ（自動入力）"))
            t_f, c1, q_f = base_tf, base_c1, base_qf

        st.markdown("---")
        st.subheader("入庫量・重量の設定")
        q4_mode = st.radio("入庫品重量(W) の計算方法", ["収容率・収容量から算出（自動）", "直接入力（手動）"], horizontal=True)
        
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            storage_rate = st.number_input("収容率 (%)", value=60.0, step=5.0)
            turnover_rate = st.number_input("入出庫率 (%)", value=33.0, step=1.0)
        with col_w2:
            calc_w = capacity_ton * 1000 * (storage_rate / 100.0) * (turnover_rate / 100.0)
            if q4_mode == "収容率・収容量から算出（自動）":
                w_weight = st.number_input("入庫品重量 W (kg)", value=float(round(calc_w, 1)), disabled=True)
            else:
                w_weight = st.number_input("入庫品重量 W (kg) [手動入力]", value=4621.0, step=100.0)

        h_cooling = st.number_input("入庫物冷却時間 H (時間)", value=24.0, step=1.0)

    # ④ 各種諸負荷の設定
    with st.expander("④ 負荷情報（自動・手動設定）", expanded=False):
        st.subheader("《換気負荷》 (Q2)")
        q2_mode = st.radio("換気回数(N) の設定", ["自動", "手動"], horizontal=True, key="q2_m")
        default_n = 15.0 if volume < 20 else (10.9 if volume < 50 else (7.5 if volume < 100 else 5.0))
        v_cycles = st.number_input("換気回数 N (回/24h)", value=default_n, disabled=(q2_mode=="自動"))
        e_vent = st.number_input("換気熱量 E (W/m³)", value=34.5, step=0.1)
        
        st.markdown("---")
        st.subheader("《作業員負荷》 (Q3)")
        q3_mode = st.radio("作業人数 の設定", ["自動", "手動"], horizontal=True, key="q3_m")
        calc_p_count = max(0.1, round(volume / 250.0, 1))
        p_count = st.number_input("作業人数 D (人)", value=float(calc_p_count), disabled=(q3_mode=="自動"), format="%.1f")
        p_heat = st.number_input("作業員発生熱量 F (W/人)", value=410.0, step=10.0)
        p_hours = st.number_input("作業時間 H (時間)", value=3.0, step=0.5)
        
        st.markdown("---")
        st.subheader("《電灯負荷》 (Q5)")
        q5_mode = st.radio("電灯総負荷 の設定", ["自動", "手動"], horizontal=True, key="q5_m")
        calc_light_kw = max(1.0, round(volume / 40.0, 1)) * 0.1
        light_kw = st.number_input("電灯総負荷 (kW)", value=float(calc_light_kw), disabled=(q5_mode=="自動"), format="%.2f")
        light_hours = st.number_input("照明時間 H (時間)", value=3.0, step=0.5)
        
        st.markdown("---")
        st.subheader("《その他オプション負荷》")
        use_forklift = st.checkbox("フォークリフトの負荷 (Q6) を含める")
        Q6 = (st.number_input("動力(kW)", value=0.0) * st.number_input("台数", value=0) * st.number_input("運転時間(h)", value=0.0) * (1.0 / 24.0)) if use_forklift else 0.0


# --- 右側：計算結果エリア ---
with col_result:
    st.header("📊 負荷計算結果")

    # 各面の面積計算とQ1算出
    a_ceiling = width * length
    a_floor = width * length
    a_wall_f = width * height
    a_wall_b = width * height
    a_wall_r = length * height
    a_wall_l = length * height

    q1_c = a_ceiling * k_val * (t_ceiling - t_int) / 1000.0
    q1_f = a_floor * k_val * (t_floor - t_int) / 1000.0
    q1_wf = a_wall_f * k_val * (t_wall_f - t_int) / 1000.0
    q1_wb = a_wall_b * k_val * (t_wall_b - t_int) / 1000.0
    q1_wr = a_wall_r * k_val * (t_wall_r - t_int) / 1000.0
    q1_wl = a_wall_l * k_val * (t_wall_l - t_int) / 1000.0
    Q1 = q1_c + q1_f + q1_wf + q1_wb + q1_wr + q1_wl

    # 各諸負荷の集計
    Q2 = volume * e_vent * v_cycles * (1.0 / 24.0) / 1000.0
    Q3 = p_heat * p_count * p_hours * (1.0 / 24.0) / 1000.0
    Q5 = light_kw * light_hours * (1.0 / 24.0)

    # 冷却負荷 Q4 の計算ロジック
    if t_in > t_f:
        q_sens1 = w_weight * c1 * (t_in - t_f)
        q_latent = w_weight * q_f
        q_sens2 = w_weight * c2 * (t_f - t_int)
        total_j = q_sens1 + q_latent + q_sens2
        Q4 = total_j * (1.0 / 3600.0) * (1.0 / h_cooling)
    else:
        Q4 = w_weight * c2 * (t_in - t_int) * (1.0 / 3600.0) * (1.0 / h_cooling)

    # 最終的な合計負荷 Q の算出（除霜時間2h補正、安全率適用）
    sum_Q = Q1 + Q2 + Q3 + Q4 + Q5 + Q6
    defrost_hours = 2.0
    final_Q = (sum_Q * 24.0 / (24.0 - defrost_hours)) * margin_rate

    # 結果表示
    st.success(f"### 🚀 必要とされる冷却能力 (Q) : {final_Q:.2f} kW")
    st.metric(label="冷凍トン換算", value=f"{(final_Q / 3.517):.2f} USRT")

    st.markdown("---")
    st.subheader("🔍 各負荷の明細")
    
    # 計算式を項目名に内包させたテーブル表示
    st.table({
        "負荷項目": [
            "《壁等からの侵入熱》 (Q1)  【 式：AXKX(TO-TI)×1/1000 】",
            "《換気による負荷》 (Q2)  【 式：((VXEXN)×1/24)/1000 】",
            "《作業員による負荷》 (Q3)  【 式：FXDXHX1/24 】",
            "《冷却負荷》 (Q4)  【 式：WXCX(TA-TI)×2.778×10⁻⁴×1/H 】",
            "《電灯の負荷》 (Q5)  【 式：QLXWXHX1/24 】",
            "合計 (単純総和)"
        ],
        "計算値 (kW)": [f"{Q1:.2f}", f"{Q2:.2f}", f"{Q3:.2f}", f"{Q4:.2f}", f"{Q5:.2f}", f"{sum_Q:.2f}"]
    })

    # =================================================================
    # ★新機能：三菱風レイアウトのExcelファイルを自動生成する関数
    # =================================================================
    def generate_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "負荷計算書"
        
        # グリッド線を表示する設定
        ws.views.sheetView[0].showGridLines = True
        
        # スタイル定義
        font_title = Font(name="MS ゴシック", size=16, bold=True)
        font_header = Font(name="MS ゴシック", size=11, bold=True, color="FFFFFF")
        font_bold = Font(name="MS ゴシック", size=11, bold=True)
        font_normal = Font(name="MS ゴシック", size=11)
        
        fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        fill_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        
        thin_side = Side(border_style="thin", color="000000")
        double_side = Side(border_style="double", color="000000")
        border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        border_total = Border(top=thin_side, bottom=double_side)
        
        # 1. タイトル
        ws["A1"] = "***** 冷蔵庫の負荷計算書 *****"
        ws["A1"].font = font_title
        
        # 2. 条件設定エリア
        ws["A3"] = "【基本条件設定】"
        ws["A3"].font = font_bold
        
        conditions = [
            ("部屋寸法", f"{width:.1f} m × {length:.1f} m × {height:.1f} m", "庫内温度 T_I", f"{t_int:.1f} ℃"),
            ("内容積 V", f"{volume:.1f} m³", "外気温度 T_O", f"{t_ceiling:.1f} ℃ (天井基準)"),
            ("収容能力", f"{capacity_ton:.1f} トン", "安全率（余裕率）", f"{margin_rate:.2f}"),
            ("入庫品温度 T_A", f"{t_in:.1f} ℃", "入庫品重量 W", f"{w_weight:.1f} kg")
        ]
        
        row_idx = 4
        for c_data in conditions:
            ws.cell(row=row_idx, column=1, value=c_data[0]).font = font_normal
            ws.cell(row=row_idx, column=2, value=c_data[1]).font = font_bold
            ws.cell(row=row_idx, column=4, value=c_data[2]).font = font_normal
            ws.cell(row=row_idx, column=5, value=c_data[3]).font = font_bold
            row_idx += 1
            
        # 3. 負荷計算結果テーブル（Book1.pdfヘッダーの再現）
        row_idx += 2
        ws.cell(row=row_idx, column=1, value="負荷項目").font = font_header
        ws.cell(row=row_idx, column=1).fill = fill_navy
        ws.cell(row=row_idx, column=2, value="計算基本公式").font = font_header
        ws.cell(row=row_idx, column=2).fill = fill_navy
        ws.cell(row=row_idx, column=3, value="計算値 (kW)").font = font_header
        ws.cell(row=row_idx, column=3).fill = fill_navy
        
        results_data = [
            ("① 壁等からの侵入熱 (Q1)", "A × K × (TO - TI) × 1/1000", Q1),
            ("② 換気による負荷 (Q2)", "((V × E × N) × 1/24) / 1000", Q2),
            ("③ 作業員による負荷 (Q3)", "F × D × H × 1/24", Q3),
            ("④ 物品冷却負荷 (Q4)", "W × C × (TA - TI) × 2.778×10⁻⁴ × 1/H", Q4),
            ("⑤ 電灯の負荷 (Q5)", "QL × W × H × 1/24", Q5),
            ("⑥ その他オプション負荷 (Q6)", "チェックスイッチ連動動力負荷", Q6),
            ("■ 諸負荷単純合計 (sum_Q)", "Q1 ＋ Q2 ＋ Q3 ＋ Q4 ＋ Q5 ＋ Q6", sum_Q)
        ]
        
        start_row = row_idx + 1
        for item, formula, val in results_data:
            row_idx += 1
            ws.cell(row=row_idx, column=1, value=item).font = font_normal
            ws.cell(row=row_idx, column=2, value=formula).font = font_normal
            
            val_cell = ws.cell(row=row_idx, column=3, value=round(val, 2))
            val_cell.font = font_bold
            val_cell.alignment = Alignment(horizontal="right")
            
            # 罫線と装飾
            for col in range(1, 4):
                ws.cell(row=row_idx, column=col).border = border_box
            if "合計" in item:
                for col in range(1, 4):
                    ws.cell(row=row_idx, column=col).fill = fill_light
                    
        # 4. 最終総合負荷 Q (運転時間補正・余裕率適用後)
        row_idx += 2
        ws.cell(row=row_idx, column=1, value="🚀 必要とされる総冷却能力 (Q)").font = font_bold
        final_cell = ws.cell(row=row_idx, column=3, value=round(final_Q, 2))
        final_cell.font = Font(name="MS ゴシック", size=12, bold=True, color="FF0000")
        final_cell.alignment = Alignment(horizontal="right")
        final_cell.border = border_total
        
        ws.cell(row=row_idx+1, column=1, value="（参考）冷凍トン換算能力").font = font_normal
        ton_cell = ws.cell(row=row_idx+1, column=3, value=round(final_Q / 3.517, 2))
        ton_cell.font = font_bold
        ton_cell.alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx+1, column=4, value="USRT").font = font_normal

        # 列幅の自動調整
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 42
        ws.column_dimensions["C"].width = 15
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 25
        
        # バイナリストリームに保存して返す
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # Excelダウンロードボタンの設置
    st.markdown("---")
    excel_file = generate_excel()
    st.download_button(
        label="📥 計算結果をExcelファイルに出力する",
        data=excel_file,
        file_name="冷凍冷蔵設備_負荷計算書.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )