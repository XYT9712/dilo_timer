# -*- coding: utf-8 -*-
"""
IE DILO 计时工具 - Streamlit 后端版（方案B）
========================================
核心特性：
  1. 动作编码 N1~N100，说明 NT1~NT100（可自定义添加）
  2. 序号列（第一列，按行自动计数）+ 备注列（最后一列）
  3. 备注编辑：选行号 -> 显示动作编码 -> 输入备注 -> 更新
  4. 计时：开始 / 停止 / 重置（重置二次确认，防误触）
  5. 切换编码自动保存上一段记录
  6. 【关键】每录一条记录，自动追加写入同一个 Excel 文件，不丢数据
  7. 累计总时长 + 记录条数统计

依赖安装：
  pip install streamlit pandas openpyxl
  pip install streamlit-autorefresh        # 可选：计时器页面自动刷新

运行：
  streamlit run dilo_timer.py
  python -m streamlit run dilo_timer.py
"""

import streamlit as st
import pandas as pd
import time
import os
from datetime import datetime

st.set_page_config(page_title="IE DILO 计时工具", page_icon="⏱️", layout="wide")

# ============ 配置 ============
EXCEL_FILE = "DILO记录.xlsx"      # 所有记录追加到同一个文件
CODE_FILE  = "DILO编码表.xlsx"    # 编码对照表（可自定义）
COLUMNS    = ["动作编码", "动作说明", "时长(秒)", "时长(时分秒)", "记录时间", "备注"]

# ============ 工具函数 ============
def pad(n):
    return str(int(n)).zfill(2)

def fmt_time(sec):
    """秒 -> HH:MM:SS:CC"""
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    c = int((sec - int(sec)) * 100)
    return f"{pad(h)}:{pad(m)}:{pad(s)}:{pad(c)}"

def now_str():
    now = datetime.now()
    return f"{now.year}-{pad(now.month)}-{pad(now.day)} {pad(now.hour)}:{pad(now.minute)}:{pad(now.second)}"

# ============ 数据加载 / 保存 ============
def load_records():
    """从 Excel 读取历史记录（记录自动累积在同一个文件）"""
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
            # 补齐列，避免旧文件缺列
            for c in COLUMNS:
                if c not in df.columns:
                    df[c] = ""
            return df[COLUMNS]
        except Exception as e:
            st.warning(f"读取记录文件失败：{e}")
    return pd.DataFrame(columns=COLUMNS)

def save_records(df):
    """把完整记录写回同一个 Excel 文件"""
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

def load_code_table():
    """编码表：优先读文件，否则用默认 N1~N100"""
    if os.path.exists(CODE_FILE):
        try:
            cdf = pd.read_excel(CODE_FILE, engine="openpyxl")
            return {str(k): str(v) for k, v in zip(cdf["动作编码"], cdf["动作说明"])}
        except Exception:
            pass
    return {f"N{i}": f"NT{i}" for i in range(1, 101)}

def save_code_table(ct):
    cdf = pd.DataFrame([{"动作编码": k, "动作说明": v} for k, v in ct.items()])
    cdf.to_excel(CODE_FILE, index=False, engine="openpyxl")

# ============ 会话状态初始化 ============
def init_state():
    if "records" not in st.session_state:
        st.session_state.records = load_records()
    if "code_table" not in st.session_state:
        st.session_state.code_table = load_code_table()
    if "running" not in st.session_state:
        st.session_state.running = False
    if "start_ts" not in st.session_state:
        st.session_state.start_ts = 0
    if "seg_elapsed" not in st.session_state:
        st.session_state.seg_elapsed = 0.0          # 暂停/已累计段时长
    if "selected_code" not in st.session_state:
        st.session_state.selected_code = "N1"
    if "last_code" not in st.session_state:
        st.session_state.last_code = "N1"
    if "reset_confirm" not in st.session_state:
        st.session_state.reset_confirm = False

init_state()

# ============ 记录操作 ============
def add_record(code, desc, duration_sec):
    """追加一条记录并自动写回 Excel（同一个文档）"""
    rec = {
        "动作编码": code,
        "动作说明": desc,
        "时长(秒)": round(duration_sec, 2),
        "时长(时分秒)": fmt_time(duration_sec),
        "记录时间": now_str(),
        "备注": "",
    }
    df = st.session_state.records
    df = pd.concat([df, pd.DataFrame([rec])], ignore_index=True)
    st.session_state.records = df
    save_records(df)   # 关键：每次自动写入同一个 Excel 文件
    return df

def update_remark(row_index, remark):
    df = st.session_state.records
    df.loc[row_index, "备注"] = remark
    st.session_state.records = df
    save_records(df)

# ============ 页面 UI ============
st.title("⏱️ IE DILO 计时工具")

# 顶部统计
total_sec = st.session_state.records["时长(秒)"].astype(float).sum() if len(st.session_state.records) else 0.0
total_sec = float(total_sec)
st.metric("累计总时长", fmt_time(total_sec), f"记录条数：{len(st.session_state.records)}")

# ---------- 第一行：编码选择 ----------
col_c, col_d = st.columns([1, 2])
with col_c:
    code_options = list(st.session_state.code_table.keys())
    sel_code = st.selectbox(
        "动作编码",
        code_options,
        index=code_options.index(st.session_state.selected_code)
        if st.session_state.selected_code in code_options else 0,
        key="sel_code_box",
    )
    # 切换编码：如果正在计时，自动保存上一段
    if sel_code != st.session_state.last_code:
        if st.session_state.running:
            prev_code = st.session_state.last_code
            prev_desc = st.session_state.code_table.get(prev_code, "")
            now = time.time()
            dur = (now - st.session_state.start_ts) + st.session_state.seg_elapsed
            add_record(prev_code, prev_desc, dur)
            st.session_state.seg_elapsed = 0.0
            st.success(f"已自动保存记录 {prev_code}，时长 {fmt_time(dur)}")
        st.session_state.last_code = sel_code
        st.session_state.selected_code = sel_code
with col_d:
    st.text_input("动作说明（只读）", value=st.session_state.code_table.get(sel_code, ""),
                  disabled=True)

# ---------- 新增/覆盖编码 ----------
with st.expander("➕ 新增 / 覆盖动作编码"):
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        new_code = st.text_input("动作编码", key="new_code")
    with c2:
        new_desc = st.text_input("动作说明", key="new_desc")
    with c3:
        st.write("")
        st.write("")
        if st.button("添加", use_container_width=True):
            nc = new_code.strip()
            nd = new_desc.strip()
            if nc and nd:
                st.session_state.code_table[nc] = nd
                save_code_table(st.session_state.code_table)
                st.success(f"已更新编码：{nc} -> {nd}")
            else:
                st.warning("请输入动作编码和动作说明")

# ---------- 计时区 ----------
col_a, col_b, col_c2, col_d2 = st.columns(4)
with col_a:
    if st.button("▶ 开始", use_container_width=True, type="primary"):
        if not st.session_state.running:
            st.session_state.running = True
            st.session_state.start_ts = time.time()
            st.session_state.seg_elapsed = 0.0
            st.session_state.last_code = sel_code
            st.rerun()
with col_b:
    if st.button("⏸ 停止并保存", use_container_width=True):
        if st.session_state.running:
            now = time.time()
            dur = (now - st.session_state.start_ts) + st.session_state.seg_elapsed
            add_record(st.session_state.selected_code,
                       st.session_state.code_table.get(st.session_state.selected_code, ""),
                       dur)
            st.session_state.running = False
            st.session_state.seg_elapsed = 0.0
            st.success(f"已停止并保存，时长 {fmt_time(dur)}")
            st.rerun()
        else:
            st.warning("计时未启动")
with col_c2:
    # 重置：二次确认（防误触）
    if st.session_state.reset_confirm:
        if st.button("⚠ 再点一次确认重置", use_container_width=True, type="secondary"):
            st.session_state.running = False
            st.session_state.seg_elapsed = 0.0
            st.session_state.start_ts = 0
            st.session_state.reset_confirm = False
            st.info("计时器已重置（历史记录保留）")
            st.rerun()
    else:
        if st.button("↺ 重置计时器", use_container_width=True):
            st.session_state.reset_confirm = True
            st.rerun()
with col_d2:
    if st.button("📊 查看数据", use_container_width=True):
        st.session_state.show_data = not st.session_state.get("show_data", False)

# 当前计时段实时显示（依赖 autorefresh；不装包时页面刷新也会更新）
timer_placeholder = st.empty()
if st.session_state.running:
    try:
        from streamlit_autorefresh import st_autorefresh
        st_autorefresh(interval=200, key="dilo_timer_refresh")
    except Exception:
        pass  # 未安装 autorefresh，刷新页面时更新时间
    cur = (time.time() - st.session_state.start_ts) + st.session_state.seg_elapsed
    timer_placeholder.markdown(
        f"### 当前计时：<span style='font-size:42px;color:#2196F3;'>{fmt_time(cur)}</span>",
        unsafe_allow_html=True)
else:
    timer_placeholder.markdown(f"### 当前计时：<span style='font-size:42px;color:#888;'>00:00:00:00</span>",
                               unsafe_allow_html=True)

# ---------- 备注编辑区 ----------
with st.expander("✏️ 备注编辑（选择数据行后修改备注）", expanded=False):
    if len(st.session_state.records) > 0:
        total_n = len(st.session_state.records)
        row_labels = [f"第{i+1}行（{r['动作编码']}）"
                      for i, r in st.session_state.records.iterrows()]
        sel_row = st.selectbox("选择数据行", range(total_n), format_func=lambda i: row_labels[i], key="sel_row")
        sel_code_disp = st.session_state.records.loc[sel_row, "动作编码"]
        cur_remark = st.session_state.records.loc[sel_row, "备注"]
        c1, c2 = st.columns([1, 3])
        with c1:
            st.text_input("动作编码（只读）", value=sel_code_disp, disabled=True)
        with c2:
            new_remark = st.text_input("输入备注", value=str(cur_remark) if cur_remark is not None else "", key="remark_input")
        if st.button("✅ 更新备注"):
            update_remark(sel_row, new_remark)
            st.success(f"已更新第 {sel_row+1} 行备注")
            st.rerun()
    else:
        st.info("暂无记录")

# ---------- 数据表格 ----------
if st.session_state.get("show_data", False):
    st.subheader("📋 记录表格")
    if len(st.session_state.records) > 0:
        disp = st.session_state.records.copy()
        disp.insert(0, "序号", range(1, len(disp) + 1))
        st.dataframe(disp, use_container_width=True)
    else:
        st.info("暂无记录")

# ---------- 下载 / 清空 ----------
col_down, col_clear = st.columns(2)
with col_down:
    if st.button("📥 下载Excel", use_container_width=True):
        if len(st.session_state.records) > 0:
            exp = st.session_state.records.copy()
            exp.insert(0, "序号", range(1, len(exp) + 1))
            csv_data = exp.to_csv(index=False).encode("utf-8-sig")
            st.download_button("点击下载 CSV（可用Excel打开）", csv_data,
                               f"DILO计时记录_{datetime.now().strftime('%Y-%m-%d')}.csv",
                               "text/csv")
        else:
            st.warning("暂无数据")
with col_clear:
    if st.button("🗑 清空全部数据", use_container_width=True):
        st.session_state.records = pd.DataFrame(columns=COLUMNS)
        save_records(st.session_state.records)
        st.session_state.show_data = False
        st.info("已清空全部记录")
        st.rerun()

st.caption("提示：记录已自动追加写入本地 Excel 文件，关掉网页或刷新都不会丢数据。")
