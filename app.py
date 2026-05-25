import streamlit as st
import io
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- ページの設定 ---
st.set_page_config(page_title="設備設計熱計算総合ツール", layout="wide")

# ==========================================================================================
# ★ 会社ロゴの挿入（ファイルの存在チェック付き）
# ==========================================================================================
logo_filename = "company_logo.png"

if os.path.exists(logo_filename):
    st.image(logo_filename, width=250)
else:
    st.caption("🏢 [ここに company_logo.png を配置するとロゴ画像に切り替わります]")

st.title("🏢 設備設計熱計算総合ツール")

# --- トップタブの設置 ---
tab_reito, tab_kucho = st.tabs(["❄️ 冷凍冷蔵設備 負荷計算", "🍃 空調設備 負荷計算・選定"])


# ==========================================================================================
# ★ タブ1：冷凍冷蔵設備 負荷計算
# ==========================================================================================
with tab_reito:
    FOOD_PROPERTIES = {
        "牛肉・豚肉": {"tf": -1.7, "c1": 3.20, "c2": 1.70, "qf": 230.0},
        "鮮魚": {"tf": -1.1, "c1": 3.60, "c2": 2.00, "qf": 270.0},
        "野菜（ほうれん草等）": {"tf": -0.8, "c1": 3.90, "c2": 2.10, "qf": 300.0},
        "水（参考値）": {"tf": 0.0, "c1": 4.19, "c2": 2.05, "qf": 333.5},
    }

    col_input, col_result = st.columns([1, 1])

    with col_input:
        st.header("📋 冷凍冷蔵 入力条件設定")
        
        with st.expander("① 部屋情報・基本設定", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1: width = st.number_input("部屋の幅 (m)", value=5.4, step=0.1, key="r_w")
            with c2: length = st.number_input("部屋の奥行 (m)", value=3.6, step=0.1, key="r_l")
            with c3: height = st.number_input("部屋の高さ (m)", value=2.4, step=0.1, key="r_h")
                
            t_int = st.number_input("庫内温度 T_I (℃)", value=-25.0, step=1.0, key="r_ti")
            margin_rate = st.number_input("安全率（余裕率）", value=1.25, step=0.05, key="r_m")
            
            volume = width * length * height
            floor_area = width * length
            capacity_ton = volume * 0.40102
            st.info(f"💡 計算値：床面積 {floor_area:.1f} m² / 内容積 {volume:.1f} m³ / 収容能力 {capacity_ton:.1f} トン")

        with st.expander("② 各部断熱・周囲温度設定", expanded=False):
            panel_type = st.selectbox("パネル厚み選択", ["冷凍 100mm (K=0.21)", "冷蔵 50mm (K=0.42)", "手入力"], key="r_p")
            k_val = 0.21 if "100mm" in panel_type else (0.42 if "50mm" in panel_type else 0.21)
            
            st.caption("各面の外気・周囲温度 (℃)")
            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                t_ceiling = st.number_input("天井周囲温度", value=40.0, step=1.0, key="r_tc")
                t_floor = st.number_input("床下周囲温度", value=25.0, step=1.0, key="r_tf")
            with cc2:
                t_wall_f = st.number_input("壁（正面）温度", value=33.0, step=1.0, key="r_t_wf")
                t_wall_b = st.number_input("壁（後面）温度", value=33.0, step=1.0, key="r_t_wb")
            with cc3:
                t_wall_r = st.number_input("壁（右面）温度", value=33.0, step=1.0, key="r_t_wr")
                t_wall_l = st.number_input("壁（左面）温度", value=33.0, step=1.0, key="r_t_wl")
                
            if panel_type == "手入力":
                k_val = st.number_input("熱通過率 K値 [W/(m²・K)]", value=0.2100, format="%.4f", key="r_k")

        with st.expander("③ 入庫情報・食品物性設定", expanded=True):
            st.subheader("入庫品の温度・比熱設定")
            t_in = st.number_input("入庫品の温度 (℃)", value=5.0, step=1.0, key="r_tin")
            food_mode = st.radio("食品比熱の入力設定", ["主要食品から選ぶ（自動入力）", "手動入力"], horizontal=True, key="r_fmode")
            
            if food_mode == "主要食品から選ぶ（自動入力）":
                selected_food = st.selectbox("主要な食品を選択してください", list(FOOD_PROPERTIES.keys()), key="r_food")
                base_tf, base_c1, base_c2, base_qf = FOOD_PROPERTIES[selected_food]["tf"], FOOD_PROPERTIES[selected_food]["c1"], FOOD_PROPERTIES[selected_food]["c2"], FOOD_PROPERTIES[selected_food]["qf"]
            else:
                base_tf, base_c1, base_c2, base_qf = -2.0, 3.34880, 2.09300, 209.3

            if t_in > base_tf:
                fc1, fc2 = st.columns(2)
                with fc1:
                    t_f = st.number_input("入庫品の凍結点 (℃)", value=base_tf, disabled=(food_mode=="主要食品から選ぶ（自動入力）"), key="r_v_tf")
                    c1 = st.number_input("凍結前状態の比熱 [kJ/(kg・K)]", value=base_c1, format="%.5f", disabled=(food_mode=="主要食品から選ぶ（自動入力）"), key="r_v_c1")
                with fc2:
                    q_f = st.number_input("入庫品の凍結潜熱 [kJ/kg]", value=base_qf, disabled=(food_mode=="主要食品から選ぶ（自動入力）"), key="r_v_qf")
                    c2 = st.number_input("凍結後状態の比熱 [kJ/(kg・K)]", value=base_c2, format="%.5f", disabled=(food_mode=="主要食品から選ぶ（自動入力）"), key="r_v_c2")
            else:
                c2 = st.number_input("凍結後状態の比熱 [kJ/(kg・K)]", value=base_c2, format="%.5f", disabled=(food_mode=="主要食品から選ぶ（自動入力）"), key="r_v_c2_fr")
                t_f, c1, q_f = base_tf, base_c1, base_qf

            st.markdown("---")
            st.subheader("入庫量・重量の設定")
            q4_mode = st.radio("入庫品重量(W) の計算方法", ["収容率・収容量から算出（自動）", "直接入力（手動）"], horizontal=True, key="r_wmode")
            
            col_w1, col_w2 = st.columns(2)
            with col_w1:
                storage_rate = st.number_input("収容率 (%)", value=60.0, step=5.0, key="r_sr")
                turnover_rate = st.number_input("入出庫率 (%)", value=33.0, step=1.0, key="r_tr")
            with col_w2:
                calc_w = capacity_ton * 1000 * (storage_rate / 100.0) * (turnover_rate / 100.0)
                w_weight = st.number_input("入庫品重量 W (kg)", value=float(round(calc_w, 1)), disabled=(q4_mode=="収容率・収容量から算出（自動）"), key="r_w_wt")

            h_cooling = st.number_input("入庫物冷却時間 H (時間)", value=24.0, step=1.0, key="r_hc")

        with st.expander("④ 負荷情報（自動・手動設定）", expanded=False):
            st.subheader("《換気負荷》 (Q2)")
            q2_mode = st.radio("換気回数(N) の設定", ["自動", "手動"], horizontal=True, key="r_q2m")
            default_n = 15.0 if volume < 20 else (10.9 if volume < 50 else (7.5 if volume < 100 else 5.0))
            v_cycles = st.number_input("換気回数 N (回/24h)", value=default_n, disabled=(q2_mode=="自動"), key="r_vc")
            e_vent = st.number_input("換気熱量 E (W/m³)", value=34.5, step=0.1, key="r_ev")
            
            st.markdown("---")
            st.subheader("《作業員負荷》 (Q3)")
            q3_mode = st.radio("作業人数 の設定", ["自動", "手動"], horizontal=True, key="r_q3m")
            calc_p_count = max(0.1, round(volume / 250.0, 1))
            p_count = st.number_input("作業人数 D (人)", value=float(calc_p_count), disabled=(q3_mode=="自動"), format="%.1f", key="r_pc")
            p_heat = st.number_input("作業員発生熱量 F (W/人)", value=410.0, step=10.0, key="r_ph")
            p_hours = st.number_input("作業時間 H (時間)", value=3.0, step=0.5, key="r_p_hr")
            
            st.markdown("---")
            st.subheader("《電灯負荷》 (Q5)")
            q5_mode = st.radio("電灯総負荷 の設定", ["自動", "手動"], horizontal=True, key="r_q5m")
            calc_light_kw = max(1.0, round(volume / 40.0, 1)) * 0.1
            light_kw = st.number_input("電灯総負荷 (kW)", value=float(calc_light_kw), disabled=(q5_mode=="自動"), format="%.2f", key="r_lkw")
            light_hours = st.number_input("照明時間 H (時間)", value=3.0, step=0.5, key="r_lhr")
            
            st.markdown("---")
            st.subheader("《その他オプション負荷》")
            use_forklift = st.checkbox("フォークリフトの負荷 (Q6) を含める", key="r_defl")
            Q6 = (st.number_input("動力(kW)", value=0.0, key="r_f_kw") * st.number_input("台数", value=0, key="r_f_un") * st.number_input("運転時間(h)", value=0.0, key="r_f_hr") * (1.0 / 24.0)) if use_forklift else 0.0

    with col_result:
        st.header("📊 冷凍冷蔵 負荷計算結果")
        Q1 = ( (width * length * k_val * (t_ceiling - t_int)) + (width * length * k_val * (t_floor - t_int)) + (width * height * k_val * (t_wall_f - t_int)) + (width * height * k_val * (t_wall_b - t_int)) + (length * height * k_val * (t_wall_r - t_int)) + (length * height * k_val * (t_wall_l - t_int)) ) / 1000.0
        Q2 = volume * e_vent * v_cycles * (1.0 / 24.0) / 1000.0
        Q3 = p_heat * p_count * p_hours * (1.0 / 24.0) / 1000.0
        Q5 = light_kw * light_hours * (1.0 / 24.0)

        if t_in > t_f:
            q_sens1 = w_weight * c1 * (t_in - t_f)
            q_latent = w_weight * q_f
            q_sens2 = w_weight * c2 * (t_f - t_int)
            Q4 = (q_sens1 + q_latent + q_sens2) * (1.0 / 3600.0) * (1.0 / h_cooling)
        else:
            Q4 = w_weight * c2 * (t_in - t_int) * (1.0 / 3600.0) * (1.0 / h_cooling)

        sum_Q = Q1 + Q2 + Q3 + Q4 + Q5 + Q6
        defrost_hours = 2.0
        final_Q = (sum_Q * 24.0 / (24.0 - defrost_hours)) * margin_rate

        st.success(f"### 🚀 必要とされる冷却能力 (Q) : {final_Q:.2f} kW")
        st.metric(label="冷凍トン換算", value=f"{(final_Q / 3.517):.2f} USRT")
        st.markdown("---")
        st.subheader("🔍 各負荷の明細")
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
        # ★ 新機能：冷却能力(Q)の計算式プロセスの明記
        # =================================================================
        st.info(f"💡 **必要とされる冷却能力 (Q) の計算プロセス**\n\n"
                f"**【計算式】 (諸負荷単純合計 × 24 / (24 - 霜取時間)) × 安全率**\n\n"
                f"＝ ({sum_Q:.2f} kW × 24 / (24 - {defrost_hours:.1f} h)) × {margin_rate:.2f} \n\n"
                f"＝ **{final_Q:.2f} kW**")


# ==========================================================================================
# ★ タブ2：空調設備 負荷計算・選定
# ==========================================================================================
with tab_kucho:
    AIR_DENSITY_MASTER = {
        "一般事務室 (150 W/m²)": 150,
        "会議室・食堂 (250 W/m²)": 250,
        "店舗・物販 (200 W/m²)": 200,
        "サーバー室・電算室 (400 W/m²)": 400,
        "工場・軽作業スペース (300 W/m²)": 300
    }

    col_k_in, col_k_res = st.columns([1, 1])

    with col_k_in:
        st.header("📋 空調 入力条件設定")
        
        calc_method = st.radio(
            "計算方式を選択してください",
            ["パターンA（超簡易：面積換算）", "パターンB（実務簡易計算）"],
            horizontal=True
        )

        with st.expander("① 部屋寸法・基本設定", expanded=True):
            ck1, ck2, ck3 = st.columns(3)
            with ck1: k_width = st.number_input("部屋の幅 (m)", value=10.0, step=0.5, key="k_w")
            with ck2: k_length = st.number_input("部屋の奥行 (m)", value=8.0, step=0.5, key="k_l")
            with ck3: k_height = st.number_input("部屋の高さ (m)", value=2.7, step=0.1, key="k_h")
            
            k_area = k_width * k_length
            k_volume = k_area * k_height
            st.info(f"💡 計算値：床面積 {k_area:.1f} m² / 容積 {k_volume:.1f} m³")

        if calc_method == "パターンA（超簡易：面積換算）":
            with st.expander("② 面積換算用の条件設定", expanded=True):
                usage_type = st.selectbox("部屋の用途（負荷密度）", list(AIR_DENSITY_MASTER.keys()))
                load_density = AIR_DENSITY_MASTER[usage_type]
                k_margin = st.number_input("安全率（余裕率）", value=1.10, step=0.05, key="k_margin_a")

        else:
            with st.expander("② 構造体侵入熱の設定（ガラス・壁面）", expanded=True):
                st.caption("外気温と室内設計温度")
                kb1, kb2 = st.columns(2)
                with kb1: t_ext_k = st.number_input("設計外気温度 (℃)", value=35.0, step=0.5, key="k_te")
                with kb2: t_int_k = st.number_input("室内設定温度 (℃)", value=26.0, step=0.5, key="k_ti_ac")
                
                delta_t = t_ext_k - t_int_k
                
                st.markdown("---")
                st.caption("ガラス窓の条件")
                k_glass_area = st.number_input("ガラス窓の総面積 (m²)", value=12.0, step=1.0)
                
                k_glass_heat = st.number_input(
                    "窓面の日射・通過熱量係数 (W/m²)", 
                    value=350.0, 
                    step=10.0, 
                    help="""【窓面係数の実務目安】
                    
                    ■ 方位・日射の影響（ブラインドありの場合）
                    ・東面 / 西面（西日が当たる）：350 〜 450 W/m²
                    ・南面（昼間に当たる）     ：250 〜 350 W/m²
                    ・北面（日陰・通過熱のみ） ：120 〜 180 W/m²
                    
                    ■ ガラスの種類による補正
                    ・一般的な単板ガラス  ：上記目安通り
                    ・複層（ペア）ガラス   ：上記目安より -50 W/m²
                    ・遮熱Low-Eガラス    ：上記目安より -100 W/m²"""
                )
                
                st.caption("外壁・天井の条件")
                k_wall_area = (2 * (k_width + k_length) * k_height) - k_glass_area + k_area 
                
                k_wall_u = st.number_input(
                    "外壁・天井の平均熱貫流率 U値 [W/(m²・K)]", 
                    value=0.85, 
                    step=0.05,
                    help="""【構造体U値（熱貫流率）の実務目安】
                    
                    ■ 外壁の目安
                    ・コンクリート打ちっぱなし（断熱なし）：3.0 〜 4.0 W/(m²・K)
                    ・一般的なALCパネル（厚100mm）     ：1.5 〜 1.8 W/(m²・K)
                    ・内断熱あり（グラスウール等50mm）  ：0.6 〜 0.9 W/(m²・K)
                    ・高断熱ビル構造                    ：0.3 〜 0.5 W/(m²・K)
                    
                    ■ 天井・屋根の目安
                    ・折板屋根（断熱なし・直下天井）    ：5.0 〜 6.0 W/(m²・K)
                    ・屋上コンクリート（断熱あり）      ：0.7 〜 1.0 W/(m²・K)
                    ・グラスウール敷込ありの二重天井     ：0.5 〜 0.8 W/(m²・K)"""
                )

            with st.expander("③ 内部発熱の設定（人間・照明・機器）", expanded=True):
                st.caption("室内の人員発熱")
                k_people = st.number_input("在室最大人数 (人)", value=12, step=1)
                k_people_heat = st.number_input("人間1人あたりの発生熱量 (W/人)", value=120.0, step=5.0, help="一般事務等で約110〜130W")
                
                st.markdown("---")
                st.caption("照明・OA機器発熱")
                k_light_w = st.number_input("電灯・照明総消費電力 (W)", value=600, step=50)
                k_oa_w = st.number_input("OA機器・その他熱源総出力 (W)", value=1500, step=100)

            with st.expander("④ 外気処理負荷の設定（換気）", expanded=True):
                k_vent_mode = st.radio("換気量の設定", ["自動（1人あたり30m³/h）", "手動入力"], horizontal=True)
                if k_vent_mode == "自動（1人あたり30m³/h）":
                    k_vent_vol = st.number_input("必要換気量 (m³/h)", value=float(k_people * 30), disabled=True)
                else:
                    k_vent_vol = st.number_input("必要換気量 (m³/h) [手動]", value=360.0, step=10.0)
                    
                k_vent_enthalpy = st.number_input("外気処理の熱量係数 (W / (m³/h))", value=11.5, step=0.5, key="k_vent_ent")
                k_margin = st.number_input("安全率（余裕率）", value=1.15, step=0.05, key="k_margin_b")

    with col_k_res:
        st.header("📊 空調 負荷計算・選定結果")
        
        if calc_method == "パターンA（超簡易：面積換算）":
            q_k_total_w = k_area * load_density
            q_k_final_kw = (q_k_total_w / 1000.0) * k_margin
            st.success(f"### 🚀 必要冷房能力 : {q_k_final_kw:.2f} kW")
            st.table({
                "計算ステップ": ["床面積 (m²)", "用途別 負荷密度 (W/m²)", "安全率（余裕率）", "最終必要空調能力 (kW)"],
                "数値 / 公式": [f"{k_area:.1f}", f"{load_density}", f"{k_margin:.2f}", "【式：(床面積×負荷密度/1000)×安全率】"]
            })
            
        else:
            q_glass = k_glass_area * k_glass_heat
            q_wall = k_wall_area * k_wall_u * delta_t
            q_structure_k = q_glass + q_wall
            
            q_human = k_people * k_people_heat
            q_internal_k = q_human + k_light_w + k_oa_w
            
            q_vent_k = k_vent_vol * k_vent_enthalpy
            
            sum_q_k = q_structure_k + q_internal_k + q_vent_k
            q_k_final_kw = (sum_q_k / 1000.0) * k_margin
            
            st.success(f"### 🚀 必要冷房能力 : {q_k_final_kw:.2f} kW")
            st.subheader("🔍 冷房負荷の明細")
            st.table({
                "負荷項目": [
                    "① 構造体侵入熱 (窓面日射 ＋ 壁・天井熱通過) 【式：(Ag×Qg)＋(Aw×U×ΔT)】",
                    "② 室内内部発熱 (人間 ＋ 照明 ＋ OA機器) 【式：(人×F)＋照明W＋機器W】",
                    "③ 外気処理換気負荷 【式：換気量(m³/h) × エンタルピー係数】",
                    "合計 (単純総和)"
                ],
                "計算値 (kW)": [
                    f"{(q_structure_k/1000.0):.2f}",
                    f"{(q_internal_k/1000.0):.2f}",
                    f"{(q_vent_k/1000.0):.2f}",
                    f"{(sum_q_k/1000.0):.2f}"
                ]
            })

        required_hp = q_k_final_kw / 2.8
        st.info(f"💡 **推奨されるエアコンの容量（目安）： {required_hp:.1f} 馬力**")

        def generate_kucho_excel():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "空調負荷計算書"
            ws.views.sheetView[0].showGridLines = True
            
            font_title = Font(name="MS ゴシック", size=14, bold=True)
            font_header = Font(name="MS ゴシック", size=11, bold=True, color="FFFFFF")
            font_bold = Font(name="MS ゴシック", size=11, bold=True)
            font_normal = Font(name="MS ゴシック", size=11)
            fill_green = PatternFill(start_color="385723", end_color="385723", fill_type="solid")
            thin_side = Side(border_style="thin", color="000000")
            border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            
            ws["A1"] = "***** 空調設備 熱負荷計算書 *****"
            ws["A1"].font = font_title
            
            ws["A3"] = "【計算基本条件】"
            ws["A3"].font = font_bold
            ws["A4"] = "部屋寸法"; ws["B4"] = f"{k_width:.1f}m × {k_length:.1f}m × {k_height:.1f}m"
            ws["A5"] = "床面積"; ws["B5"] = f"{k_area:.1f} m²"
            ws["A6"] = "計算方式"; ws["B6"] = calc_method
            
            ws.cell(row=8, column=1, value="負荷内訳項目").font = font_header
            ws.cell(row=8, column=2, value="計算値 (kW)").font = font_header
            
            if calc_method == "パターンA（超簡易：面積換算）":
                ws.cell(row=9, column=1, value=f"床面積換算負荷 ({usage_type})").font = font_normal
                ws.cell(row=9, column=2, value=round(q_k_total_w/1000.0, 2)).font = font_bold
                ws.cell(row=10, column=1, value="安全率適用後 最終空調能力").font = font_bold
                ws.cell(row=10, column=2, value=round(q_k_final_kw, 2)).font = font_bold
            else:
                ws.cell(row=9, column=1, value="① 構造体侵入熱 (Q_structure)").font = font_normal
                ws.cell(row=9, column=2, value=round(q_structure_k/1000.0, 2)).font = font_bold
                ws.cell(row=10, column=1, value="② 室内内部発熱 (Q_internal)").font = font_normal
                ws.cell(row=10, column=2, value=round(q_internal_k/1000.0, 2)).font = font_bold
                ws.cell(row=11, column=1, value="③ 外気処理換気負荷 (Q_vent)").font = font_normal
                ws.cell(row=11, column=2, value=round(q_vent_k/1000.0, 2)).font = font_bold
                ws.cell(row=12, column=1, value="🚀 必要総空調能力 (安全率含む)").font = font_bold
                ws.cell(row=12, column=2, value=round(q_k_final_kw, 2)).font = font_bold
            
            ws.column_dimensions["A"].width = 45
            ws.column_dimensions["B"].width = 25
            
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output

        st.markdown("---")
        k_excel = generate_kucho_excel()
        st.download_button(
            label="📥 空調の計算結果をExcel出力する",
            data=k_excel,
            file_name="空調設備_負荷計算書.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
