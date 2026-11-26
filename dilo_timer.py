# -*- coding: utf-8 -*-
"""DILO activity timer implemented with Streamlit."""

from __future__ import annotations

from datetime import datetime
from io import BytesIO
import time

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import streamlit as st


st.set_page_config(
    page_title="DILO | 活动计时器",
    page_icon="⏱️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

DEFAULT_CODE_TABLE = {f"N{i}": f"NT{i}" for i in range(1, 101)}
RECORD_COLUMNS = [
    "序号",
    "动作编码",
    "动作说明",
    "时长(秒)",
    "时长(时分秒)",
    "记录时间",
    "备注",
]

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Manrope:wght@400;500;600;700;800&family=Noto+Sans+SC:wght@400;500;600;700;800&display=swap');
    :root {
        --ink: #172533;
        --muted: #718091;
        --blue: #1769aa;
        --blue-dark: #12558a;
        --line: #e5ebf1;
        --surface: #ffffff;
        --canvas: #f3f6f9;
    }
    .stApp {
        background:
            radial-gradient(ellipse at 12% 0%, rgba(52, 140, 207, .10), transparent 34rem),
            var(--canvas);
        color: var(--ink);
        font-family: 'Manrope', 'Noto Sans SC', sans-serif;
    }
    .block-container { max-width: 860px; padding: 2rem 1.15rem 3rem; }
    header[data-testid="stHeader"] { background: transparent; }
    h1, h2, h3, p, label { font-family: 'Manrope', 'Noto Sans SC', sans-serif; }
    .hero {
        display:flex; align-items:center; gap:14px; margin: 0 0 18px;
    }
    .hero-mark {
        width:48px; height:48px; display:grid; place-items:center;
        border-radius:15px; color:#fff; font-size:24px;
        background:linear-gradient(145deg,#3192d1,#12558a);
        box-shadow:0 8px 20px rgba(23,105,170,.20);
    }
    .hero-title { font-size:25px; line-height:1.15; font-weight:800; letter-spacing:-.5px; }
    .hero-subtitle { margin-top:5px; color:var(--muted); font-size:13px; }
    .timer-card {
        position:relative; overflow:hidden; padding:22px 18px 20px;
        border:1px solid rgba(220,230,239,.9); border-radius:22px;
        background:linear-gradient(145deg,#fff,#f8fbfd);
        box-shadow:0 12px 32px rgba(32,61,83,.07);
    }
    .timer-kicker { color:var(--muted); text-align:center; font-size:12px; font-weight:700; letter-spacing:1.6px; text-transform:uppercase; }
    .timer-readout {
        margin:9px 0 14px; color:#1769aa; text-align:center;
        font:500 clamp(2.15rem, 11vw, 4.3rem)/1.1 'DM Mono', monospace;
        font-variant-numeric:tabular-nums; letter-spacing:-2px; white-space:nowrap;
    }
    .status-row { display:flex; align-items:center; justify-content:center; gap:8px; color:var(--muted); font-size:12px; }
    .status-dot { width:8px; height:8px; border-radius:50%; background:#9aa8b5; }
    .status-dot.running { background:#28a36a; box-shadow:0 0 0 4px rgba(40,163,106,.12); }
    .metric-strip {
        display:flex; justify-content:center; gap:0; margin-top:16px;
        border-top:1px solid var(--line); padding-top:14px;
    }
    .metric-item { text-align:center; padding:0 18px; }
    .metric-item + .metric-item { border-left:1px solid var(--line); }
    .metric-label { color:var(--muted); font-size:11px; }
    .metric-value { margin-top:3px; color:var(--ink); font-size:15px; font-weight:700; font-variant-numeric:tabular-nums; }
    .section-label {
        margin:23px 0 9px; color:#516173; font-size:12px;
        font-weight:800; letter-spacing:.6px;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-color:var(--line); border-radius:16px; background:var(--surface);
    }
    div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button {
        min-height:44px; border-radius:11px; font-weight:700;
        transition:transform .12s ease, box-shadow .12s ease;
    }
    div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover {
        transform:translateY(-1px); box-shadow:0 5px 14px rgba(23,105,170,.12);
    }
    div[data-testid="stTextInput"] input, div[data-testid="stSelectbox"] [role="combobox"] {
        border-radius:10px;
    }
    div[data-testid="stAlert"] { border-radius:12px; }
    div[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
    .fine-print { margin-top:15px; color:var(--muted); font-size:11px; text-align:center; }
    @media (max-width: 600px) {
        .block-container { padding:1.2rem .8rem 2rem; }
        .hero-title { font-size:22px; }
        .timer-card { padding:19px 11px 16px; border-radius:18px; }
        .metric-item { padding:0 10px; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def fmt_time(seconds: float) -> str:
    """Format seconds as HH:MM:SS:CC, matching the original timer."""
    seconds = max(0.0, float(seconds))
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    whole_seconds = int(seconds % 60)
    centiseconds = int((seconds - int(seconds)) * 100)
    return f"{hours:02d}:{minutes:02d}:{whole_seconds:02d}:{centiseconds:02d}"


def init_state() -> None:
    defaults = {
        "records": [],
        "code_table": dict(DEFAULT_CODE_TABLE),
        "selected_code": "N1",
        "last_selected_code": "N1",
        "running": False,
        "start_ts": 0.0,
        "stopped_elapsed": 0.0,
        "quick_remark": "",
        "reset_confirm_until": 0.0,
        "clear_confirm_until": 0.0,
        "clear_password_visible": False,
        "show_records": False,
        "selected_record": 0,
        "notice": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if "action_description" not in st.session_state:
        st.session_state.action_description = st.session_state.code_table.get(
            st.session_state.selected_code, ""
        )


def notify(message: str, kind: str = "info") -> None:
    st.session_state.notice = (kind, message)


def record_segment(code: str, duration: float, remark: str = "") -> None:
    st.session_state.records.append(
        {
            "动作编码": code,
            "动作说明": st.session_state.code_table.get(code, ""),
            "时长(秒)": round(duration, 2),
            "时长(时分秒)": fmt_time(duration),
            "记录时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "备注": remark.strip(),
        }
    )


def on_code_change() -> None:
    new_code = st.session_state.selected_code
    previous_code = st.session_state.get("last_selected_code", new_code)
    if new_code != previous_code and st.session_state.running:
        elapsed = max(0.0, time.time() - st.session_state.start_ts)
        record_segment(previous_code, elapsed, st.session_state.quick_remark)
        st.session_state.start_ts = time.time()
        st.session_state.stopped_elapsed = 0.0
        notify(f"已自动保存 {previous_code} 的记录，时长 {fmt_time(elapsed)}", "success")
    if new_code != previous_code:
        st.session_state.quick_remark = ""
        st.session_state.action_description = st.session_state.code_table.get(new_code, "")
    st.session_state.last_selected_code = new_code


def add_or_update_code() -> None:
    code = st.session_state.new_code.strip()
    description = st.session_state.new_description.strip()
    if not code or not description:
        notify("请输入动作编码和动作说明。", "warning")
        return
    st.session_state.code_table[code] = description
    st.session_state.new_code = ""
    st.session_state.new_description = ""
    notify(f"已成功添加/更新编码：{code} → {description}", "success")


def update_description() -> None:
    code = st.session_state.selected_code
    description = st.session_state.action_description.strip()
    if not code:
        notify("没有可更新的动作编码。", "warning")
    elif not description:
        notify("动作说明不能为空。", "warning")
    else:
        old_description = st.session_state.code_table.get(code, "")
        st.session_state.code_table[code] = description
        notify(f"已更新 {code} 的动作说明：{old_description} → {description}", "success")


def load_selected_remark() -> None:
    index = st.session_state.selected_record
    if 0 <= index < len(st.session_state.records):
        st.session_state.remark_editor = st.session_state.records[index].get("备注", "")


def update_remark() -> None:
    index = st.session_state.selected_record
    if not 0 <= index < len(st.session_state.records):
        notify("请先选择要编辑的记录行。", "warning")
        return
    remark = st.session_state.remark_editor.strip()
    st.session_state.records[index]["备注"] = remark
    notify(f"已更新第 {index + 1} 行备注：{remark or '（空）'}", "success")


def start_timer() -> None:
    code = st.session_state.selected_code
    if st.session_state.running:
        notify("计时已经在运行中。", "warning")
    elif not code or not st.session_state.code_table.get(code):
        notify("请先选择有效的动作编码。", "warning")
    else:
        st.session_state.running = True
        st.session_state.start_ts = time.time()
        st.session_state.stopped_elapsed = 0.0
        st.session_state.last_selected_code = code
        notify(f"计时已启动，当前编码：{code}", "success")


def stop_timer() -> None:
    if not st.session_state.running:
        notify("计时未启动。", "warning")
        return
    elapsed = max(0.0, time.time() - st.session_state.start_ts)
    code = st.session_state.get("last_selected_code", st.session_state.selected_code)
    record_segment(code, elapsed, st.session_state.quick_remark)
    st.session_state.running = False
    st.session_state.stopped_elapsed = elapsed
    st.session_state.start_ts = 0.0
    st.session_state.quick_remark = ""
    notify(f"已停止并保存 {code} 的记录，时长 {fmt_time(elapsed)}", "success")


def reset_timer() -> None:
    now = time.time()
    if st.session_state.reset_confirm_until >= now:
        st.session_state.running = False
        st.session_state.start_ts = 0.0
        st.session_state.stopped_elapsed = 0.0
        st.session_state.reset_confirm_until = 0.0
        st.session_state.quick_remark = ""
        notify("计时器已重置（历史记录保留）。", "success")
    else:
        st.session_state.reset_confirm_until = now + 5
        notify("⚠️ 请在 5 秒内再次点击确认重置；历史记录会保留。", "warning")


def toggle_records() -> None:
    st.session_state.show_records = not st.session_state.show_records
    if not st.session_state.records and st.session_state.show_records:
        notify("暂无记录。", "info")


def request_clear() -> None:
    now = time.time()
    if st.session_state.clear_confirm_until >= now:
        st.session_state.clear_password_visible = True
        st.session_state.clear_confirm_until = 0.0
        notify("请输入密码 123 以确认清除全部数据。", "warning")
    else:
        st.session_state.clear_confirm_until = now + 5
        st.session_state.clear_password_visible = False
        notify("⚠️ 请在 5 秒内再次点击；随后还需输入密码。", "warning")


def confirm_clear() -> None:
    if st.session_state.clear_password.strip() != "123":
        st.session_state.clear_password = ""
        notify("密码错误，无法清除数据。", "error")
        return
    st.session_state.records = []
    st.session_state.code_table = {}
    st.session_state.running = False
    st.session_state.start_ts = 0.0
    st.session_state.stopped_elapsed = 0.0
    st.session_state.quick_remark = ""
    st.session_state.selected_code = ""
    st.session_state.last_selected_code = ""
    st.session_state.action_description = ""
    st.session_state.clear_password = ""
    st.session_state.clear_password_visible = False
    st.session_state.show_records = False
    st.session_state.selected_record = 0
    st.session_state.notice = ("success", "✅ 全部缓存数据已清除（编码表和计时记录）。")


def cancel_clear() -> None:
    st.session_state.clear_password = ""
    st.session_state.clear_password_visible = False
    st.session_state.clear_confirm_until = 0.0
    notify("已取消清除操作。", "info")


def make_excel() -> bytes:
    workbook = Workbook()
    records_sheet = workbook.active
    records_sheet.title = "计时记录"
    records_sheet.append(RECORD_COLUMNS)
    for index, record in enumerate(st.session_state.records, start=1):
        records_sheet.append(
            [
                index,
                record.get("动作编码", ""),
                record.get("动作说明", ""),
                record.get("时长(秒)", 0),
                record.get("时长(时分秒)", ""),
                record.get("记录时间", ""),
                record.get("备注", ""),
            ]
        )
    code_sheet = workbook.create_sheet("编码对照表")
    code_sheet.append(["动作编码", "动作说明"])
    for code, description in st.session_state.code_table.items():
        code_sheet.append([code, description])

    for sheet, widths in (
        (records_sheet, [8, 14, 24, 14, 20, 22, 32]),
        (code_sheet, [18, 30]),
    ):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1769AA")
            cell.alignment = Alignment(horizontal="center")
        for column, width in enumerate(widths, start=1):
            sheet.column_dimensions[chr(64 + column)].width = width
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def render_timer() -> None:
    running = st.session_state.running
    elapsed = (
        max(0.0, time.time() - st.session_state.start_ts)
        if running
        else st.session_state.stopped_elapsed
    )
    start_ms = int((time.time() - elapsed) * 1000)
    state_text = "正在计时" if running else "已暂停"
    dot_class = "running" if running else ""
    timer_color = "#1769aa" if running else "#7b8996"
    st.html(
        f"""
        <div class="timer-card">
          <div class="timer-kicker">当前时段</div>
          <div id="dilo-clock" class="timer-readout" style="color:{timer_color}">{fmt_time(elapsed)}</div>
          <div class="status-row"><span class="status-dot {dot_class}"></span><span>{state_text}</span></div>
        </div>
        <script>
        (() => {{
          const el = document.getElementById("dilo-clock");
          const running = {str(running).lower()};
          const start = {start_ms};
          const pad = (n) => String(Math.floor(n)).padStart(2, "0");
          const format = (seconds) => {{
            const h = Math.floor(seconds / 3600);
            const m = Math.floor((seconds % 3600) / 60);
            const s = Math.floor(seconds % 60);
            const c = Math.floor((seconds - Math.floor(seconds)) * 100);
            return `${{pad(h)}}:${{pad(m)}}:${{pad(s)}}:${{pad(c)}}`;
          }};
          function tick() {{
            if (!running || !el) return;
            el.textContent = format(Math.max(0, (Date.now() - start) / 1000));
            requestAnimationFrame(tick);
          }}
          if (running) tick();
        }})();
        </script>
        """,
        unsafe_allow_javascript=True,
    )


init_state()
code_options = list(st.session_state.code_table)
if code_options and st.session_state.selected_code not in code_options:
    st.session_state.selected_code = code_options[0]
    st.session_state.last_selected_code = code_options[0]
    st.session_state.action_description = st.session_state.code_table[code_options[0]]

st.markdown(
    """
    <div class="hero">
      <div class="hero-mark">⏱</div>
      <div>
        <div class="hero-title">DILO 活动计时器</div>
        <div class="hero-subtitle">记录、整理并导出每一段工作活动</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

total_seconds = sum(float(record.get("时长(秒)", 0) or 0) for record in st.session_state.records)
st.markdown(
    f"""
    <div class="timer-card" style="padding-bottom:17px">
      <div class="timer-kicker">DILO TIMER</div>
      <div class="metric-strip">
        <div class="metric-item"><div class="metric-label">累计总时长</div><div class="metric-value">{fmt_time(total_seconds)}</div></div>
        <div class="metric-item"><div class="metric-label">记录条数</div><div class="metric-value">{len(st.session_state.records)}</div></div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)
render_timer()

if st.session_state.notice:
    notice_kind, notice_text = st.session_state.notice
    getattr(st, notice_kind if notice_kind in {"success", "warning", "error", "info"} else "info")(
        notice_text
    )
    st.session_state.notice = None

st.markdown('<div class="section-label">当前动作</div>', unsafe_allow_html=True)
code_col, desc_col = st.columns([1, 1.8], gap="medium")
with code_col:
    if code_options:
        selected_code = st.selectbox(
            "动作编码",
            code_options,
            key="selected_code",
            on_change=on_code_change,
            label_visibility="collapsed",
        )
    else:
        selected_code = ""
        st.selectbox("动作编码", ["无可用编码"], disabled=True, label_visibility="collapsed")
with desc_col:
    st.text_input("动作说明", key="action_description", placeholder="动作说明")
st.text_input(
    "快速备注",
    key="quick_remark",
    placeholder="快速备注（切换或停止编码时自动保存到本段记录）",
    help="填写后，在切换动作编码或停止计时保存时会附加到当前记录。",
)
st.button(
    "更新当前编码的动作说明",
    on_click=update_description,
    disabled=not bool(selected_code),
    use_container_width=True,
)

with st.expander("＋ 添加 / 覆盖动作编码", expanded=False):
    with st.form("add_code_form", clear_on_submit=False):
        add_code_col, add_desc_col = st.columns([1, 2])
        with add_code_col:
            st.text_input("新动作编码", key="new_code", placeholder="例如：N101")
        with add_desc_col:
            st.text_input("动作说明", key="new_description", placeholder="例如：整理工作区域")
        st.form_submit_button("添加或更新编码", on_click=add_or_update_code, use_container_width=True)

st.markdown('<div class="section-label">计时控制</div>', unsafe_allow_html=True)
start_col, stop_col, reset_col, data_col = st.columns(4, gap="small")
with start_col:
    st.button(
        "▶ 开始",
        on_click=start_timer,
        type="primary",
        use_container_width=True,
        disabled=not bool(selected_code),
    )
with stop_col:
    st.button("■ 停止并保存", on_click=stop_timer, use_container_width=True)
with reset_col:
    reset_pending = st.session_state.reset_confirm_until >= time.time()
    st.button(
        "确认重置" if reset_pending else "↺ 重置计时器",
        on_click=reset_timer,
        use_container_width=True,
        type="secondary" if reset_pending else "tertiary",
    )
with data_col:
    st.button(
        "隐藏数据" if st.session_state.show_records else "查看数据",
        on_click=toggle_records,
        use_container_width=True,
    )

if st.session_state.show_records:
    st.markdown('<div class="section-label">记录表格</div>', unsafe_allow_html=True)
    if st.session_state.records:
        table_rows = [
            {
                "序号": index,
                **record,
            }
            for index, record in enumerate(st.session_state.records, start=1)
        ]
        st.dataframe(table_rows, use_container_width=True, hide_index=True)
    else:
        st.info("暂无记录")

with st.expander("✎ 编辑记录备注", expanded=False):
    if st.session_state.records:
        record_labels = [
            f"第 {index} 行 · {record.get('动作编码', '')} · {record.get('记录时间', '')}"
            for index, record in enumerate(st.session_state.records)
        ]
        st.selectbox(
            "选择数据行",
            options=range(len(record_labels)),
            format_func=lambda index: record_labels[index],
            key="selected_record",
            on_change=load_selected_remark,
        )
        selected_record = st.session_state.records[st.session_state.selected_record]
        st.text_input("动作编码（只读）", value=selected_record.get("动作编码", ""), disabled=True)
        st.text_input("输入备注", key="remark_editor", placeholder="为所选记录添加备注")
        st.button("更新备注", on_click=update_remark, use_container_width=True)
    else:
        st.info("暂无记录")

st.markdown('<div class="section-label">导出与数据管理</div>', unsafe_allow_html=True)
export_col, clear_col = st.columns(2, gap="medium")
with export_col:
    st.download_button(
        "下载 Excel 工作簿",
        data=make_excel(),
        file_name=f"DILO计时记录_{datetime.now():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not bool(st.session_state.records),
        use_container_width=True,
        help="包含「计时记录」和「编码对照表」两个工作表。",
    )
with clear_col:
    clear_pending = st.session_state.clear_confirm_until >= time.time()
    st.button(
        "再次点击确认清除" if clear_pending else "清除缓存数据",
        on_click=request_clear,
        use_container_width=True,
        type="secondary" if clear_pending or st.session_state.clear_password_visible else "tertiary",
    )

if st.session_state.clear_password_visible:
    with st.container(border=True):
        st.warning("此操作会清空当前会话的编码表和计时记录，且不可恢复。请输入密码确认。")
        st.text_input("清除密码", key="clear_password", type="password", placeholder="请输入密码（123）")
        confirm_col, cancel_col = st.columns(2)
        with confirm_col:
            st.button("确认清除", on_click=confirm_clear, type="primary", use_container_width=True)
        with cancel_col:
            st.button("取消", on_click=cancel_clear, use_container_width=True)

st.markdown(
    '<div class="fine-print">当前数据保存在此浏览器会话的 Streamlit 会话状态中；请定期下载 Excel 备份。</div>',
    unsafe_allow_html=True,
)
