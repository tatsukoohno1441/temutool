import streamlit as st
import tempfile, os, uuid, subprocess, sys, inspect
from pathlib import Path

# -----------------------------------------------------------------------------
# 尝试以「模块」形式加载原脚本；若失败就回退到 subprocess
# -----------------------------------------------------------------------------
try:
    import build_report_v3
except ImportError:
    build_report_v3 = None

try:
    import make_shipping_csv_v2
except ImportError:
    make_shipping_csv_v2 = None

st.set_page_config(page_title="Temu 订单处理", page_icon="🧡", layout="centered")

# -----------------------------------------------------------------------------
# 样式注入（圆角彩色按钮 + 边框）
# -----------------------------------------------------------------------------
CSS = """
<style>
.blue-btn > button {background-color:#46b6ff;color:#fff;border:none;border-radius:24px;height:48px;width:230px;font-size:18px;font-weight:600;cursor:pointer;}
.green-btn > button {background-color:#45c46b;color:#fff;border:none;border-radius:24px;height:48px;width:230px;font-size:18px;font-weight:600;cursor:pointer;}
.filebox .stUploadDropzone {border:2px solid #000;border-radius:6px;height:90px;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

st.title("Temu 发货助手 🧡")

# -----------------------------------------------------------------------------
# 工具函数
# -----------------------------------------------------------------------------

def save_upload(uploaded_file):
    """把 UploadedFile 写到临时磁盘并返回路径"""
    suffix = Path(uploaded_file.name).suffix or ""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()
    return tmp.name

def popen_script(script_name: str, *args):
    """后备方案：创建子进程执行原脚本"""
    subprocess.run([sys.executable, script_name, *args], check=True)

def call_script(module, script_name: str, arg_paths: list[str]):
    """优先用模块 main()，若签名不匹配则退到 subprocess"""
    if module is None or not hasattr(module, "main"):
        popen_script(script_name, *arg_paths)
        return

    sig = inspect.signature(module.main)
    try:
        if len(sig.parameters) == 0:  # 旧脚本：main() 取 sys.argv
            old_argv = sys.argv.copy()
            sys.argv = [script_name, *arg_paths]
            module.main()
            sys.argv = old_argv
        else:
            module.main(*arg_paths)
    except TypeError:
        # 参数个数对不上时再降级
        popen_script(script_name, *arg_paths)

# -----------------------------------------------------------------------------
# ① 生成检货/发货单
# -----------------------------------------------------------------------------

st.header("1️⃣ 生成发货单 (检货单)")
order_file = st.file_uploader("选择 Temu 订单 Excel / CSV", type=["xlsx", "csv"], key="order")

if st.button("生成检货单", type="primary", key="btn-pick", help="根据原始订单生成发货 Excel") and order_file:
    with st.spinner("正在生成发货单，请稍候…"):
        in_path = save_upload(order_file)
        out_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}_report.xlsx")
        try:
            call_script(build_report_v3, "build_report_v3.py", [in_path, out_path])
        except Exception as e:
            st.error(f"生成失败：{e}")
        else:
            st.session_state.pick_download = out_path
            st.success("完成啦！点击下方按钮下载 ✅")

if "pick_download" in st.session_state:
    st.download_button("下载发货单", open(st.session_state.pick_download, "rb"), file_name="temu_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl-pick")

# -----------------------------------------------------------------------------
# 分割线
# -----------------------------------------------------------------------------

st.markdown("---")

# -----------------------------------------------------------------------------
# ② 生成 Yamato CSV
# -----------------------------------------------------------------------------

st.header("2️⃣ 生成 Yamato CSV 文件")
orig_csv = st.file_uploader("原始 Temu CSV", type="csv", key="orig")
pick_excel = st.file_uploader("发货单 / 检货单 Excel", type="xlsx", key="pick")

if st.button("生成 Yamato CSV", key="btn-ship"):
    if not (orig_csv and pick_excel):
        st.warning("请同时上传两个文件哦 🥺")
    else:
        with st.spinner("正在生成 Yamato CSV…"):
            csv_path = save_upload(orig_csv)
            excel_path = save_upload(pick_excel)
            out_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}_yamato.csv")
            try:
                call_script(make_shipping_csv_v2, "make_shipping_csv_v2.py", [csv_path, excel_path, out_path])
            except Exception as e:
                st.error(f"生成失败：{e}")
            else:
                st.session_state.ship_download = out_path
                st.success("Yamato CSV 已生成 ✅，点击下方按钮下载")

if "ship_download" in st.session_state:
    st.download_button("下载 Yamato CSV", open(st.session_state.ship_download, "rb"), file_name="yamato.csv", mime="text/csv", key="dl-ship")
