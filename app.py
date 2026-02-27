import streamlit as st
import pandas as pd
import io

# ===========================
# 页面配置
# ===========================
st.set_page_config(page_title="储能调度模拟 (含效率)", layout="wide")
st.title("🔋 风电储能调度模拟器 (含充放电效率)")
st.markdown("""
本工具基于 **逐小时仿真** 逻辑，考虑了 **充放电效率 (&eta;)** 对 SOC 和功率的影响。
- **充电时**：存入能量 = 实际充电功率 &times; &eta;
- **放电时**：消耗能量 = 放电功率 / &eta;
""")

# ===========================
# 侧边栏：参数设置
# ===========================
st.sidebar.header("⚙️ 参数设置")

P_es_max = st.sidebar.number_input(
    "储能最大充/放电功率 (kW)",
    min_value=0.0,
    value=20000.0,
    step=500.0,
    help="对应 P_es_max"
)

E_cap = st.sidebar.number_input(
    "储能总容量 (kWh)",
    min_value=0.0,
    value=72000.0,
    step=1000.0,
    help="对应 E_cap"
)

initial_soc = st.sidebar.number_input(
    "初始 SOC (kWh)",
    min_value=0.0,
    value=0.0,
    step=100.0,
    help="对应 initial_soc"
)

eta = st.sidebar.number_input(
    "充放电效率 (&eta;)",
    min_value=0.1,
    max_value=1.0,
    value=0.93,
    step=0.01,
    help="对应 eta (0-1之间)"
)

# ===========================
# 文件上传
# ===========================
uploaded_file = st.file_uploader("📂 上传数据文件 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 读取数据
        df = pd.read_excel(uploaded_file)

        # 检查必要列
        required_cols = ['wind_mw', 'grid_limit']
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ 错误：输入文件必须包含 {required_cols} 两列！")
            st.stop()

        # 添加小时序号
        if 'hour' not in df.columns:
            df['hour'] = range(1, len(df) + 1)

        st.success(f"✅ 数据加载成功！共 {len(df)} 条记录。")

        with st.expander("查看原始数据预览"):
            st.dataframe(df.head())

        # ===========================
        # 核心仿真逻辑 (含效率)
        # ===========================
        soc_current = initial_soc

        # 预分配列表以提高性能
        list_surplus = []
        list_es_chargeable = []
        list_actual_charge = []
        list_curtailment = []
        list_discharge = []
        list_soc = []

        total_rows = len(df)
        progress_bar = st.progress(0)

        for i in range(total_rows):
            wind = df.loc[i, 'wind_mw']
            P_grid_limit = df.loc[i, 'grid_limit']

            surplus = max(0, wind - P_grid_limit)

            if surplus > 0:
                # === 充电阶段 ===
                # 最大可接受充电功率（受设备和容量限制）
                # 注意：这里要除以 eta，因为 actual_charge * eta 才是存入的电量
                # 剩余可用容量为 (E_cap - soc_current)
                # 所以最大允许输入功率 = 剩余容量 / eta
                remaining_capacity = E_cap - soc_current
                if remaining_capacity < 0: remaining_capacity = 0  # 防止过充计算错误

                chargeable_power = min(P_es_max, remaining_capacity / eta)

                actual_charge = min(surplus, chargeable_power)
                curtailment = surplus - actual_charge
                discharge = 0.0

                # SOC 更新：存入的能量 = actual_charge * eta
                soc_next = soc_current + actual_charge * eta
                es_chargeable = chargeable_power

            else:
                # === 放电阶段 ===
                discharge_allow = P_grid_limit - wind  # 电网还能接受多少补充电力

                # 最大可放电功率限制
                # 1. 设备限制: P_es_max
                # 2. 通道限制: discharge_allow
                # 3. 电量限制: 此时放电功率为 P_out, 消耗电量为 P_out / eta.
                #    所以 P_out / eta <= soc_current  =>  P_out <= soc_current * eta
                max_discharge_by_soc = soc_current * eta

                discharge = min(P_es_max, discharge_allow, max_discharge_by_soc)
                discharge = max(0, discharge)  # 确保非负

                actual_charge = 0.0
                curtailment = 0.0

                # SOC 更新：消耗的能量 = discharge / eta
                energy_consumed = discharge / eta
                soc_next = soc_current - energy_consumed
                soc_next = max(0, soc_next)  # 强制 >= 0，防止浮点误差导致负值

                es_chargeable = 0.0

            # 记录结果
            list_surplus.append(surplus)
            list_es_chargeable.append(es_chargeable)
            list_actual_charge.append(actual_charge)
            list_curtailment.append(curtailment)
            list_discharge.append(discharge)
            list_soc.append(soc_next)

            soc_current = soc_next

            if i % 1000 == 0:
                progress_bar.progress((i + 1) / total_rows)

        progress_bar.progress(1.0)

        # 写入 DataFrame
        df['surplus'] = list_surplus
        df['es_chargeable'] = list_es_chargeable
        df['actual_charge'] = list_actual_charge
        df['curtailment'] = list_curtailment
        df['discharge'] = list_discharge
        df['soc'] = list_soc

        # ===========================
        # 统计计算
        # ===========================
        total_curtailment_kwh = df['curtailment'].sum()
        total_wind_energy_kwh = df['wind_mw'].sum()
        final_soc = df['soc'].iloc[-1]

        curtailment_rate = (total_curtailment_kwh / total_wind_energy_kwh * 100) if total_wind_energy_kwh > 0 else 0.0

        # ===========================
        # 结果展示
        # ===========================
        st.divider()
        st.subheader(f"📈 全年弃风统计结果 (含充放电效率 &eta;={eta})")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(label="总弃风电量 (kWh)", value=f"{total_curtailment_kwh:,.2f}")
        with col2:
            st.metric(label="总风电发电量 (kWh)", value=f"{total_wind_energy_kwh:,.2f}")
        with col3:
            st.metric(label="弃风率", value=f"{curtailment_rate:.2f}%", delta_color="inverse")
        with col4:
            st.metric(label="储能最终电量 (kWh)", value=f"{final_soc:,.2f}")

        st.divider()

        # ===========================
        # 数据导出与预览
        # ===========================
        st.subheader("📊 详细数据预览与导出")

        column_mapping = {
            'hour': '小时',
            'wind_mw': '风电出力 (kW)',
            'grid_limit': '电网限值 (kW)',
            'surplus': '弃风余量 (kW)',
            'es_chargeable': '储能可充电功率 (kW)',
            'actual_charge': '实际充电功率 (kW)',
            'curtailment': '最终弃风量 (kW)',
            'soc': '储能电量 (kWh)',
            'discharge': '放电功率 (kW)'
        }

        result_columns = [col for col in column_mapping.keys() if col in df.columns]
        output_df = df[result_columns].rename(columns=column_mapping)

        st.dataframe(output_df.head(10), use_container_width=True)

        # 生成 Excel 文件
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            output_df.to_excel(writer, index=False, sheet_name='调度结果_含效率')

            # 格式化
            workbook = writer.book
            worksheet = writer.sheets['调度结果_含效率']
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})

            # 简单应用格式到所有数字列 (从第2列开始，第1列是小时)
            for col_num in range(1, len(output_df.columns)):
                worksheet.set_column(col_num, col_num, 15, money_fmt)
            worksheet.set_column(0, 0, 10)  # 小时列宽

        buffer.seek(0)

        st.download_button(
            label="📥 下载处理后的 Excel 数据",
            data=buffer,
            file_name=f"储能调度模拟结果_含效率_{eta}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"发生错误：{str(e)}")
        st.info("请检查文件格式是否正确，是否包含必需的列。")

else:
    st.info("👆 请在上方上传 Excel 文件以开始模拟。")
    st.markdown("""
    ### 📝 数据模板要求
    上传的 Excel 文件必须包含以下两列：
    - `wind_mw`: 风电出力 (单位需与功率参数一致，通常为 kW)
    - `grid_limit`: 电网传输限值 (单位需与功率参数一致，通常为 kW)
    """)