import hashlib
import io
import json
import os
import re
import socket
import time
from datetime import datetime
import pandas as pd
import routeros_api
import streamlit as st
import xlsxwriter

USERS_FILE = "users_db.json"
CONFIG_FILE = "config.json"
TOWERS_FILE = "towers_db.json"
LOG_DIR = "logs"

st.set_page_config(
    page_title="نظام إدارة ومراقبة زبائن البرودباند",
    page_icon="📡",
    layout="wide",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&display=swap');
    html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; }
    .block-container { max-width: 1300px; padding-top: 2rem; padding-bottom: 2rem; margin: auto; }
    .stTextInput input, .stSelectbox, .stNumberInput { text-align: right; border-radius: 10px !important; }
    [data-testid="stSidebar"] { text-align: right; background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%); border-left: 1px solid #334155; }
    .login-card-container { background: linear-gradient(145deg, #1e293b 0%, #0f172a 100%); border: 1px solid #334155; border-radius: 20px; padding: 40px; box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3); direction: rtl; max-width: 500px; margin: 40px auto; }
    .login-title { font-size: 1.8rem; font-weight: 900; color: #f8fafc; text-align: center; margin-bottom: 10px; }
    .login-subtitle { font-size: 0.95rem; color: #94a3b8; text-align: center; margin-bottom: 30px; }
    .metric-card { background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border: 1px solid #334155; border-radius: 14px; padding: 20px; text-align: center; box-shadow: 0 6px 20px rgba(0, 0, 0, 0.2); }
    .metric-val { font-size: 1.8rem; font-weight: 900; color: #38bdf8; margin-top: 8px; }
    .metric-lbl { font-size: 0.95rem; color: #94a3b8; font-weight: 500; }
    .modern-card { background: #1e293b; border: 1px solid #334155; border-radius: 16px; padding: 22px; margin-bottom: 20px; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2); direction: rtl; }
    .card-title-bar { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; padding-bottom: 14px; margin-bottom: 15px; }
    .client-name { font-size: 1.25rem; font-weight: 700; color: #f8fafc; }
    .badge-online { background: rgba(34, 197, 94, 0.15); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.3); padding: 5px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700; }
    .badge-offline { background: rgba(148, 163, 184, 0.15); color: #94a3b8; border: 1px solid rgba(148, 163, 184, 0.3); padding: 5px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700; }
    .badge-disabled { background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); padding: 5px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700; }
    .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; font-size: 0.95rem; margin-bottom: 18px; }
    .info-item { color: #cbd5e1; }
    .speed-box { background: #0f172a; border-radius: 12px; padding: 14px; display: flex; justify-content: space-around; align-items: center; border: 1px solid #334155; }
    .speed-rx { color: #38bdf8; font-weight: 700; font-size: 1.05rem; }
    .speed-tx { color: #f59e0b; font-weight: 700; font-size: 1.05rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

def make_hash(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def check_hash(password, hashed_text):
    return make_hash(password) == hashed_text

def log_action(username, action_desc):
    if not username:
        return
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    today = datetime.now().strftime("%Y-%m-%d")
    log_filename = os.path.join(LOG_DIR, f"log_{today}.json")
    logs = []
    if os.path.exists(log_filename):
        try:
            with open(log_filename, "r", encoding="utf-8") as f:
                logs = json.load(f)
        except:
            logs = []
    logs.append({"time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "user": username, "action": action_desc})
    try:
        with open(log_filename, "w", encoding="utf-8") as f:
            json.dump(logs, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"خطأ في حفظ السجل: {e}")

def get_all_logs():
    all_logs = []
    if os.path.exists(LOG_DIR):
        files = sorted([f for f in os.listdir(LOG_DIR) if f.endswith('.json')], reverse=True)
        for file in files:
            file_path = os.path.join(LOG_DIR, file)
            with open(file_path, "r", encoding="utf-8") as f:
                try:
                    all_logs.extend(json.load(f))
                except:
                    continue
    return sorted(all_logs, key=lambda x: x['time'], reverse=True)

def extract_speed_mb(profile_name):
    if not profile_name:
        return 0.0
    match = re.search(r"(\d+(?:\.\d+)?)", str(profile_name))
    if match:
        try: return float(match.group(1))
        except: return 0.0
    return 0.0

def encode_comment(comment_text):
    if not comment_text:
        return ""
    try:
        encoded_bytes = comment_text.encode("cp1256", errors="ignore")
        return "".join([f"\\x{b:02x}" for b in encoded_bytes])
    except:
        return comment_text

def decode_comment(comment_val):
    if not comment_val:
        return "لا يوجد تعليق"
    try:
        if isinstance(comment_val, str) and "\\x" in comment_val:
            hex_parts = comment_val.split("\\x")
            byte_arr = bytearray()
            for p in hex_parts:
                if p:
                    try:
                        byte_arr.append(int(p[:2], 16))
                        if len(p) > 2:
                            for char in p[2:]:
                                byte_arr.append(ord(char))
                    except:
                        pass
            return byte_arr.decode("cp1256", errors="ignore")
        if isinstance(comment_val, bytes):
            return comment_val.decode("cp1256", errors="ignore")
        return str(comment_val)
    except:
        return str(comment_val)

def format_speed(bps_str):
    try:
        bps = float(bps_str)
        mbps = bps / 1_000_000
        return f"{mbps:.2f} Mbps", bps
    except:
        return "0.00 Mbps", 0

def load_mikrotik_config():
    if not os.path.exists(CONFIG_FILE):
        default_config = {"host": "192.168.88.1", "user": "admin", "password": "admin", "port": 8728}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4)
        return default_config
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_mikrotik_config(new_conf):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(new_conf, f, indent=4)

def load_users():
    if not os.path.exists(USERS_FILE):
        default_users = {"admin": {"password": make_hash("admin"), "role": "مدير النظام", "disabled": False, "last_active": ""}}
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(default_users, f, ensure_ascii=False, indent=4)
        return default_users
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        users = json.load(f)
        for u in users:
            if "disabled" not in users[u]: users[u]["disabled"] = False
            if "last_active" not in users[u]: users[u]["last_active"] = ""
        return users

def save_users(users_data):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users_data, f, ensure_ascii=False, indent=4)

def load_towers():
    if not os.path.exists(TOWERS_FILE):
        default_towers = {
            "tower_1": {"name": "البرج الرئيسي", "ip": "192.168.88.1"}
        }
        with open(TOWERS_FILE, "w", encoding="utf-8") as f:
            json.dump(default_towers, f, ensure_ascii=False, indent=4)
        return default_towers
    try:
        with open(TOWERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_towers(towers_data):
    with open(TOWERS_FILE, "w", encoding="utf-8") as f:
        json.dump(towers_data, f, ensure_ascii=False, indent=4)

def check_tower_status(api, ip):
    clean_ip = str(ip).strip()
    if not clean_ip:
        return False, "طافي"

    try:
        ping_res = api.get_resource('/ping').call({
            'address': clean_ip,
            'count': '1'
        })
        if ping_res:
            item = ping_res[0]
            norm_item = {}
            for k, v in item.items():
                k_str = k.decode('utf-8') if isinstance(k, bytes) else str(k)
                v_str = v.decode('utf-8') if isinstance(v, bytes) else str(v)
                norm_item[k_str] = v_str

            status = str(norm_item.get('status', '')).lower()
            time_rtt = norm_item.get('time') or norm_item.get('avg-rtt')
            received = norm_item.get('received', '0')

            if status not in ['timeout', 'no-route', 'error', 'host-unreachable'] and (time_rtt or received != '0'):
                rtt = time_rtt if time_rtt else "<1ms"
                return True, f"يعمل ({rtt})"
    except Exception:
        pass

    common_ports = [80, 8080, 8291, 443, 22]
    for port in common_ports:
        try:
            start_t = time.time()
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(0.6)
            result = sock.connect_ex((clean_ip, port))
            sock.close()
            if result == 0:
                elapsed_ms = int((time.time() - start_t) * 1000)
                return True, f"يعمل ({elapsed_ms}ms)"
        except Exception:
            pass

    return False, "طافي"

def get_api_connection(conf):
    return routeros_api.RouterOsApiPool(
        host=conf["host"],
        username=conf["user"],
        password=conf["password"],
        port=int(conf.get("port", 8728)),
        plaintext_login=True,
    )

def disconnect_ppp_user(api, username):
    try:
        active_resource = api.get_resource("ppp/active")
        actives = active_resource.get(name=username)
        for act in actives:
            act_id = act.get(".id") or act.get("id")
            if act_id: active_resource.remove(id=act_id)
        return True
    except Exception as e:
        st.error(f"خطأ أثناء قطع الاتصال: {e}")
        return False

def get_secret_id_by_name(api, username):
    try:
        secret_resource = api.get_resource("ppp/secret")
        user = secret_resource.get(name=username)
        if user: return user[0].get(".id") or user[0].get("id")
    except Exception:
        pass
    return None

def toggle_ppp_secret(api, username, disable=True, secret_id=None):
    try:
        secret_resource = api.get_resource("ppp/secret")
        target_id = secret_id or get_secret_id_by_name(api, username)
        if target_id:
            secret_resource.set(id=target_id, disabled="yes" if disable else "no")
            return True
        else:
            st.error(f"لم يتم العثور على المعرف للمستخدم {username}")
            return False
    except Exception as e:
        st.error(f"خطأ أثناء تعديل حالة الحساب: {e}")
        return False

def build_styled_excel(df_users, df_profiles):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    title_fmt = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#38BDF8", "bg_color": "#0F172A", "align": "center", "valign": "vcenter"})
    sub_title_fmt = workbook.add_format({"font_size": 10, "font_color": "#94A3B8", "bg_color": "#0F172A", "align": "center", "valign": "vcenter"})
    header_fmt = workbook.add_format({"bold": True, "font_size": 11, "font_color": "#FFFFFF", "bg_color": "#1E293B", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#334155"})
    cell_fmt = workbook.add_format({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#E2E8F0"})
    num_fmt = workbook.add_format({"font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#E2E8F0", "num_format": "#,##0.0"})
    total_fmt = workbook.add_format({"bold": True, "font_size": 11, "font_color": "#0F172A", "bg_color": "#E2E8F0", "align": "center", "valign": "vcenter", "border": 2, "border_color": "#94A3B8", "num_format": "#,##0.0"})
    status_online = workbook.add_format({"bold": True, "font_color": "#15803D", "bg_color": "#DCFCE7", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#BBF7D0"})
    status_disabled = workbook.add_format({"bold": True, "font_color": "#B91C1C", "bg_color": "#FEE2E2", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#FECACA"})

    ws1 = workbook.add_worksheet("المستخدمين")
    ws1.right_to_left()
    ws1.merge_range("A1:H1", "📡 كشف شبكة المشتركين - نظام إدارة البرودباند", title_fmt)
    ws1.merge_range("A2:H2", f"تاريخ الاستخراج: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", sub_title_fmt)
    ws1.set_row(0, 25)
    ws1.set_row(1, 18)
    headers = list(df_users.columns)
    ws1.set_row(3, 22)
    for col_idx, header in enumerate(headers):
        ws1.write(3, col_idx, header, header_fmt)

    for row_idx, row_data in enumerate(df_users.values, start=4):
        ws1.set_row(row_idx, 20)
        for col_idx, val in enumerate(row_data):
            val_str = str(val)
            if headers[col_idx] == "الحالة":
                if val_str == "متصل": ws1.write(row_idx, col_idx, val_str, status_online)
                elif val_str == "معطل": ws1.write(row_idx, col_idx, val_str, status_disabled)
                else: ws1.write(row_idx, col_idx, val_str, cell_fmt)
            elif headers[col_idx] == "السرعة (Mbps)":
                ws1.write(row_idx, col_idx, val, num_fmt)
            else:
                ws1.write(row_idx, col_idx, val_str, cell_fmt)

    tot_row_ws1 = 4 + len(df_users)
    last_user_excel_row = 4 + len(df_users)
    ws1.set_row(tot_row_ws1, 22)
    ws1.write(tot_row_ws1, 0, "إجمالي المشتركين ومبيعات السرعة الكلية", total_fmt)
    ws1.write_formula(tot_row_ws1, 1, f"=COUNTA(B5:B{last_user_excel_row})", total_fmt)
    for c in range(2, len(headers) - 1):
        ws1.write(tot_row_ws1, c, "-", total_fmt)
    speed_col_letter = chr(ord('A') + headers.index("السرعة (Mbps)"))
    ws1.write_formula(tot_row_ws1, headers.index("السرعة (Mbps)"), f"=SUM({speed_col_letter}5:{speed_col_letter}{last_user_excel_row})", total_fmt)

    for col_idx, header in enumerate(headers):
        max_len = max(df_users[header].astype(str).map(len).max(), len(header)) + 5
        ws1.set_column(col_idx, col_idx, max(max_len, 16))

    ws2 = workbook.add_worksheet("جرد البروفايلات")
    ws2.right_to_left()
    ws2.merge_range("A1:D1", "📊 ملخص جرد بروفايلات السرعة وتجميع السعات (MB)", title_fmt)
    ws2.merge_range("A2:D2", f"تاريخ الاستخراج: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", sub_title_fmt)
    ws2.set_row(0, 25)
    ws2.set_row(1, 18)
    prof_headers = ["البروفايل", "السرعة (Mbps)", "عدد المشتركين", "إجمالي السعة (Mbps)"]
    ws2.set_row(3, 22)
    for col_idx, header in enumerate(prof_headers):
        ws2.write(3, col_idx, header, header_fmt)

    for idx, row_data in enumerate(df_profiles.values, start=4):
        excel_row_num = idx + 1
        ws2.set_row(idx, 20)
        ws2.write(idx, 0, row_data[0], cell_fmt)
        ws2.write(idx, 1, row_data[1], num_fmt)
        ws2.write(idx, 2, row_data[2], cell_fmt)
        ws2.write_formula(idx, 3, f"=B{excel_row_num}*C{excel_row_num}", num_fmt)

    tot_row_ws2 = 4 + len(df_profiles)
    last_prof_excel_row = 4 + len(df_profiles)
    ws2.set_row(tot_row_ws2, 22)
    ws2.write(tot_row_ws2, 0, "المجموع الكلي للشبكة", total_fmt)
    ws2.write(tot_row_ws2, 1, "-", total_fmt)
    ws2.write_formula(tot_row_ws2, 2, f"=SUM(C5:C{last_prof_excel_row})", total_fmt)
    ws2.write_formula(tot_row_ws2, 3, f"=SUM(D5:D{last_prof_excel_row})", total_fmt)

    for col_idx, header in enumerate(prof_headers):
        ws2.set_column(col_idx, col_idx, 22)

    workbook.close()
    return output.getvalue()

if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "username" not in st.session_state: st.session_state.username = ""
if "role" not in st.session_state: st.session_state.role = ""
if "last_active_menu" not in st.session_state: st.session_state.last_active_menu = ""

users_db = load_users()
mikrotik_conf = load_mikrotik_config()

if not st.session_state.logged_in and "session_user" in st.query_params:
    saved_user = st.query_params["session_user"]
    if saved_user in users_db and not users_db[saved_user].get("disabled", False):
        st.session_state.logged_in = True
        st.session_state.username = saved_user
        st.session_state.role = users_db[saved_user]["role"]
        users_db[saved_user]["last_active"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_users(users_db)
        log_action(saved_user, f"استعادة الجلسة المحفوظة عبر زر 'تذكرني' (الصلاحية: {st.session_state.role})")

if not st.session_state.logged_in:
    st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.markdown("""<div class="login-card-container"><div class="login-title">🔐 تسجيل الدخول</div><div class="login-subtitle">نظام إدارة ومراقبة زبائن البرودباند</div>""", unsafe_allow_html=True)
        username_input = st.text_input("اسم المستخدم", key="login_user", placeholder="أدخل اسم المستخدم...")
        password_input = st.text_input("كلمة المرور", type="password", key="login_pass", placeholder="أدخل كلمة المرور...")
        remember_me = st.checkbox("تذكرني", value=True, key="login_remember")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("🚀 دخول النظام", type="primary", use_container_width=True):
            if username_input in users_db:
                if users_db[username_input].get("disabled", False):
                    log_action(username_input, "محاولة دخول فاشلة (حساب معطل)")
                    st.error("❌ هذا الحساب معطل.")
                elif check_hash(password_input, users_db[username_input]["password"]):
                    st.session_state.logged_in = True
                    st.session_state.username = username_input
                    st.session_state.role = users_db[username_input]["role"]
                    users_db[username_input]["last_active"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    save_users(users_db)
                    log_action(username_input, f"تسجيل الدخول إلى النظام (الصلاحية: {st.session_state.role})")
                    if remember_me:
                        st.query_params["session_user"] = username_input
                    else:
                        if "session_user" in st.query_params:
                            del st.query_params["session_user"]
                    st.rerun()
                else:
                    log_action(username_input, "محاولة دخول فاشلة (كلمة المرور خاطئة)")
                    st.error("⚠️ اسم المستخدم أو كلمة المرور غير صحيحة")
            else:
                log_action(username_input, "محاولة دخول فاشلة (اسم المستخدم غير موجود)")
                st.error("⚠️ اسم المستخدم أو كلمة المرور غير صحيحة")
        st.markdown("</div>", unsafe_allow_html=True)
else:
    if st.session_state.username in users_db:
        users_db[st.session_state.username]["last_active"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_users(users_db)

    st.sidebar.markdown(f"### 👤 {st.session_state.username}")
    st.sidebar.markdown(f"🛡️ الصلاحية: **{st.session_state.role}**")
    st.sidebar.markdown("---")

    is_admin = (st.session_state.role == "مدير النظام")

    if is_admin:
        menu_options = [
            "لوحة المراقبة الرئيسية",
            "➕ إضافة مشترك جديد",
            "🌐 مراقبة كافة منافذ Ethernet والترافك",
            "📡 مراقبة الأبراج والروابط",
            "📡 اكتشاف قطع الشبكة (IP Neighbors)",
            "📊 تصدير التقارير (Excel)",
            "👥 إدارة طاقم العمل وفريق الدعم",
            "⚙️ إعدادات المايكروتك"
        ]
        admin_menu = st.sidebar.selectbox("📋 اختر الشاشة المطلوبة:", menu_options)
        st.sidebar.markdown("---")
        if admin_menu == "لوحة المراقبة الرئيسية":
            auto_refresh = st.sidebar.checkbox("🔄 تحديث السرعات التلقائي", value=False)
            refresh_rate = st.sidebar.slider("معدل التحديث (ثواني)", min_value=3, max_value=15, value=5)
    else:
        admin_menu = "لوحة المراقبة الرئيسية"
        auto_refresh = st.sidebar.checkbox("🔄 تحديث السرعات التلقائي", value=False)
        refresh_rate = st.sidebar.slider("معدل التحديث (ثواني)", min_value=3, max_value=15, value=5)

    # تسجيل حركة الانتقال والتصفح بين التبويبات بدون تكرار تلقائي
    if st.session_state.last_active_menu != admin_menu:
        st.session_state.last_active_menu = admin_menu
        log_action(st.session_state.username, f"الانتقال إلى تبويب: '{admin_menu}'")

    if st.sidebar.button("🚪 تسجيل الخروج", type="secondary", use_container_width=True):
        log_action(st.session_state.username, "تسجيل الخروج من النظام")
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.role = ""
        st.session_state.last_active_menu = ""
        if "session_user" in st.query_params: del st.query_params["session_user"]
        st.rerun()

    # --- 1. لوحة المراقبة الرئيسية ---
    if admin_menu == "لوحة المراقبة الرئيسية":
        st.title("📡 لوحة متابعة أعطال وزبائن البرودباند")
        st.markdown("💡 يمكنك النقر مباشرة على عنوان الآيبي الخاص بالمشترك لفتح واجهة الراوتر الخاص به.")
        st.markdown("---")
        try:
            connection = get_api_connection(mikrotik_conf)
            api = connection.get_api()
            
            ppp_secrets = api.get_resource("ppp/secret").get()
            active_users = api.get_resource("ppp/active").get()
            active_dict = {act.get("name"): act for act in active_users}
            formatted_data = []
            active_count = 0

            for sec in ppp_secrets:
                username = sec.get("name", "غير معروف")
                sec_id = sec.get(".id") or sec.get("id", "")
                is_disabled = sec.get("disabled", "false") == "true"
                raw_comment = sec.get("comment", "")
                decoded_comment = decode_comment(raw_comment)
                last_logged_out = sec.get("last-logged-out", "غير متوفر")
                active_info = active_dict.get(username)
                is_online = active_info is not None

                if is_online and not is_disabled: active_count += 1
                if is_disabled: status_text, badge_class = "حساب معطل", "badge-disabled"
                elif is_online: status_text, badge_class = "متصل الآن 🟢", "badge-online"
                else: status_text, badge_class = "غير متصل ⚪", "badge-offline"

                secret_ip = sec.get("remote-address", "").strip()
                if secret_ip and secret_ip != "0.0.0.0":
                    client_ip = secret_ip
                else:
                    client_ip = active_info.get("address", "-") if active_info else "-"

                formatted_data.append({
                    "secret_id": sec_id, "username": username, "comment": decoded_comment,
                    "disabled": is_disabled, "is_online": is_online, "status_text": status_text,
                    "badge_class": badge_class, "address": client_ip,
                    "mac": active_info.get("caller-id", active_info.get("mac-address", "-")) if active_info else "-",
                    "uptime": active_info.get("uptime", "-") if active_info else "-", "last_logged_out": last_logged_out
                })

            if is_admin:
                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1: st.markdown(f'<div class="metric-card"><div class="metric-lbl">المشتركين النشطين</div><div class="metric-val">🟢 {active_count}</div></div>', unsafe_allow_html=True)
                with col_m2: st.markdown(f'<div class="metric-card"><div class="metric-lbl">حالة الاتصال</div><div class="metric-val" style="color: #4ade80;">متصل ✅</div></div>', unsafe_allow_html=True)
                with col_m3: st.markdown(f'<div class="metric-card"><div class="metric-lbl">إجمالي الحسابات</div><div class="metric-val" style="color: #f59e0b;">{len(ppp_secrets)}</div></div>', unsafe_allow_html=True)
                st.write(""); st.write("")

            if "search_term" not in st.session_state: st.session_state.search_term = ""

            col_s1, col_s2, col_s3 = st.columns([3, 1, 1])
            with col_s1: search_input = st.text_input("🔍 ابحث باسم المشترك أو الكومنت:", value=st.session_state.search_term)
            with col_s2:
                st.write("")
                if st.button("🔍 بحث", type="primary", use_container_width=True):
                    st.session_state.search_term = search_input
                    if search_input:
                        log_action(st.session_state.username, f"البحث عن مشترك بكلمة: '{search_input}'")
                    st.rerun()
            with col_s3:
                st.write("")
                if st.button("🔄 إعادة ضبط", use_container_width=True):
                    st.session_state.search_term = ""
                    st.rerun()

            search_query = st.session_state.search_term or search_input
            if search_query:
                filtered_data = [item for item in formatted_data if search_query.lower() in item["username"].lower() or search_query.lower() in item["comment"].lower()]
                if filtered_data:
                    st.markdown("---")
                    for i in range(0, len(filtered_data), 2):
                        cols = st.columns(2)
                        for col_idx, user_idx in enumerate([i, i + 1]):
                            if user_idx < len(filtered_data):
                                user = filtered_data[user_idx]
                                matched_tx_str, matched_rx_str = "0.00 Mbps", "0.00 Mbps"
                                if user["is_online"]:
                                    possible_inf_names = [f"<pppoe-{user['username']}>", f"<ppoe-{user['username']}>", user["username"]]
                                    for inf_key in possible_inf_names:
                                        try:
                                            traffic = api.get_resource("interface").call("monitor-traffic", {"interface": inf_key, "once": ""})
                                            if traffic:
                                                rx_fmt, _ = format_speed(traffic[0].get("tx-bits-per-second", "0"))
                                                tx_fmt, _ = format_speed(traffic[0].get("rx-bits-per-second", "0"))
                                                matched_rx_str, matched_tx_str = rx_fmt, tx_fmt
                                                break
                                        except: pass

                                with cols[col_idx]:
                                    if is_admin:
                                        btn_c1, btn_c2 = st.columns(2)
                                        if btn_c1.button("🔌 فصل", key=f"card_kick_{user['username']}_{user_idx}", use_container_width=True, disabled=not user["is_online"]):
                                            if disconnect_ppp_user(api, user["username"]):
                                                log_action(st.session_state.username, f"فصل اتصال المشترك: '{user['username']}'")
                                                st.success("تم الفصل"); st.rerun()
                                        if user["disabled"]:
                                            if btn_c2.button("✅ تفعيل", key=f"card_enable_{user['username']}_{user_idx}", use_container_width=True):
                                                if toggle_ppp_secret(api, user["username"], disable=False, secret_id=user["secret_id"]):
                                                    log_action(st.session_state.username, f"تفعيل حساب المشترك: '{user['username']}'")
                                                    st.success("تم التفعيل"); st.rerun()
                                        else:
                                            if btn_c2.button("🚫 تعطيل", key=f"card_disable_{user['username']}_{user_idx}", use_container_width=True):
                                                if toggle_ppp_secret(api, user["username"], disable=True, secret_id=user["secret_id"]):
                                                    disconnect_ppp_user(api, user["username"])
                                                    log_action(st.session_state.username, f"تعطيل حساب المشترك وفصله: '{user['username']}'")
                                                    st.warning("تم التعطيل"); st.rerun()

                                    status_info_html = f'<div class="info-item">⏱️ المدة: <b dir="ltr" style="color:#4ade80;">{user["uptime"]}</b></div>' if user["is_online"] else f'<div class="info-item">🕒 آخر خروج: <b dir="ltr" style="color:#94a3b8;">{user["last_logged_out"]}</b></div>'
                                    ip_link = f'<a href="http://{user["address"]}" target="_blank" style="color:#38bdf8; text-decoration: underline;">{user["address"]} 🔗</a>' if user["address"] != "-" else '<b dir="ltr" style="color:#38bdf8;">-</b>'

                                    st.markdown(f"""
                                        <div class="modern-card">
                                            <div class="card-title-bar">
                                                <span class="client-name">💬 {user['comment']}</span>
                                                <span class="{user['badge_class']}">{user['status_text']}</span>
                                            </div>
                                            <div class="info-grid">
                                                <div class="info-item">👤 المستخدم: <b dir="ltr">{user['username']}</b></div>
                                                <div class="info-item">🌐 الآيبي: {ip_link}</div>
                                                <div class="info-item">💻 الماك: <b dir="ltr" style="color:#38bdf8;">{user['mac']}</b></div>
                                                {status_info_html}
                                            </div>
                                            <div class="speed-box">
                                                <div class="speed-rx" dir="ltr">📥 تنزيل: {matched_rx_str}</div>
                                                <div class="speed-tx" dir="ltr">📤 رفع: {matched_tx_str}</div>
                                            </div>
                                        </div>
                                    """, unsafe_allow_html=True)
                else: st.info("لا توجد نتائج مطابقة للبحث.")
            else: st.info("💡 قم بكتابة اسم المشترك أو الكومنت في خانة البحث أعلاه لعرض النتائج.")

            connection.disconnect()
            if auto_refresh:
                time.sleep(refresh_rate)
                st.rerun()
        except Exception as e:
            st.error(f"خطأ في الاتصال بالمايكروتك: {e}")

    # --- باقي الشاشات الخاصة بمدير النظام فقط ---
    elif is_admin:
        if admin_menu == "➕ إضافة مشترك جديد":
            st.title("➕ إضافة مشترك برودباند جديد")
            st.markdown("---")
            try:
                connection = get_api_connection(mikrotik_conf)
                api = connection.get_api()
                profiles_res = api.get_resource("ppp/profile").get()
                profile_names = [p.get("name") for p in profiles_res] if profiles_res else ["default"]

                with st.form("add_user_form"):
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        new_username = st.text_input("اسم المستخدم (PPP Username)")
                        new_password = st.text_input("كلمة المرور (Password)")
                    with col_f2:
                        selected_profile = st.selectbox("اختر البروفايل (Profile)", profile_names)
                        new_comment = st.text_input("اسم المشترك / ملاحظة (باللغة العربية)")

                    if st.form_submit_button("💾 حفظ وإضافة المشترك", type="primary", use_container_width=True):
                        if new_username and new_password:
                            try:
                                formatted_comment = encode_comment(new_comment)
                                api.get_resource("ppp/secret").add(name=new_username, password=new_password, profile=selected_profile, service="pppoe", comment=formatted_comment)
                                log_action(st.session_state.username, f"إضافة مشترك جديد: '{new_username}' البروفايل: '{selected_profile}' الملاحظة: '{new_comment}'")
                                st.success(f"🎉 تم إنشاء المشترك ({new_username}) بنجاح!")
                            except Exception as create_err: st.error(f"خطأ: {create_err}")
                        else: st.warning("أدخل اسم المستخدم وكلمة المرور.")
                connection.disconnect()
            except Exception as e: st.error(f"خطأ الاتصال: {e}")

        elif admin_menu == "🌐 مراقبة كافة منافذ Ethernet والترافك":
            st.title("🌐 لوحة مراقبة منافذ Ethernet وقراءة سرعة الكابل (Rate) والترافك")
            st.markdown("💡 تعرض هذه اللوحة منافذ الـ Ethernet فقط، مع سرعة ربط الكابل الحية الدقيقة (Link Rate / Status) عبر استدعاء الأداة المباشرة لكل منفذ.")
            st.markdown("---")
            
            auto_refresh_if = st.checkbox("🔄 تحديث تلقائي للشاشات الحية", value=False)
            refresh_rate_if = st.slider("معدل التحديث (ثواني)", min_value=3, max_value=15, value=5)

            try:
                connection = get_api_connection(mikrotik_conf)
                api = connection.get_api()
                all_interfaces = api.get_resource("interface").get()
                ethernet_interfaces = [inf for inf in all_interfaces if "ether" in inf.get("name", "").lower() or inf.get("type", "").lower() == "ether"]
                
                if ethernet_interfaces:
                    for inf in ethernet_interfaces:
                        inf_name = inf.get("name", "غير معروف")
                        inf_id = inf.get(".id") or inf.get("id", "")
                        disabled = inf.get("disabled", "false") == "true"
                        is_running = inf.get("running", "false") == "true"
                        
                        link_rate = "غير متصل"
                        link_status = "no-link"

                        if not disabled:
                            try:
                                monitor_res = api.get_resource("interface/ethernet").call("monitor", {"numbers": inf_id, "once": ""})
                                if monitor_res:
                                    m_data = monitor_res[0]
                                    link_status = m_data.get("status", "no-link")
                                    if link_status == "link-ok":
                                        link_rate = m_data.get("rate", "غير متحدد")
                                    else:
                                        link_rate = "Link Down"
                            except Exception:
                                link_rate = inf.get("rate", "غير متوفر")

                        rx_fmt, tx_fmt = "0.00 Mbps", "0.00 Mbps"
                        try:
                            traffic = api.get_resource("interface").call("monitor-traffic", {"interface": inf_name, "once": ""})
                            if traffic:
                                rx_fmt, _ = format_speed(traffic[0].get("tx-bits-per-second", "0"))
                                tx_fmt, _ = format_speed(traffic[0].get("rx-bits-per-second", "0"))
                        except:
                            pass

                        if disabled:
                            badge_status = "🔴 معطل (Disabled)"
                            link_rate = "معطل"
                        elif link_status == "link-ok" or is_running:
                            badge_status = "🟢 متصل وفعال (Running)"
                        else:
                            badge_status = "⚪ مقطوع / لا يوجد كابل (Link Down)"
                            link_rate = "Link Down"

                        st.markdown(f"""
                            <div class="modern-card">
                                <div class="card-title-bar">
                                    <span class="client-name">🔌 منفذ إيثرنت: <b dir="ltr">{inf_name}</b></span>
                                    <span style="font-size: 0.9rem; font-weight: 700;">{badge_status}</span>
                                </div>
                                <div class="info-grid">
                                    <div class="info-item">⚡ سرعة الكابل (Rate / Status): <b dir="ltr" style="color:#38bdf8;">{link_rate}</b></div>
                                    <div class="info-item">📦 الحزم (Rx/Tx): <b dir="ltr">{inf.get('rx-packets', 0)} / {inf.get('tx-packets', 0)}</b></div>
                                </div>
                                <div class="speed-box">
                                    <div class="speed-rx" dir="ltr">📥 تنزيل (Download): {rx_fmt}</div>
                                    <div class="speed-tx" dir="ltr">📤 رفع (Upload): {tx_fmt}</div>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                else:
                    st.info("لا توجد منافذ إيثرنت (Ethernet) معرفة في الراوتر.")
                
                connection.disconnect()

                if auto_refresh_if:
                    time.sleep(refresh_rate_if)
                    st.rerun()
            except Exception as e:
                st.error(f"خطأ أثناء جلب بيانات منافذ الإيثرنت: {e}")

        # --- شاشة مراقبة الأبراج والروابط (مع البحث، الفلترة، والتعداد) ---
        elif admin_menu == "📡 مراقبة الأبراج والروابط":
            st.title("📡 التبويب المستقل - إدارة ومراقبة الأبراج والروابط")
            st.markdown("💡 يمكنك إضافة، تعديل وحذف الأبراج مع إمكانية البحث والفلترة والتعداد اللحظي للروابط.")
            st.markdown("---")

            col_ref1, col_ref2 = st.columns([2, 3])
            with col_ref1:
                auto_refresh_towers = st.checkbox("🔄 تفعيل التحديث التلقائي للحالة", value=False)
            with col_ref2:
                tower_refresh_rate = st.slider("⏱️ اختر مدة التحديث التلقائي (ثواني):", min_value=3, max_value=30, value=15, step=1)

            st.markdown("---")

            with st.expander("➕ إضافة برج أو لينك جديد للشبكة", expanded=False):
                with st.form("add_tower_form"):
                    c_t1, c_t2, c_t3 = st.columns([2, 2, 1])
                    with c_t1:
                        tower_name_input = st.text_input("اسم البرج / اللينك (مثال: برج المشتل الرئيسي)")
                    with c_t2:
                        tower_ip_input = st.text_input("عنوان IP (مثال: 192.168.10.1)")
                    with c_t3:
                        st.write("")
                        st.write("")
                        submitted_tower = st.form_submit_button("حفظ وإضافة", type="primary", use_container_width=True)

                    if submitted_tower:
                        if tower_name_input and tower_ip_input:
                            towers_dict = load_towers()
                            new_key = f"tower_{int(time.time())}"
                            towers_dict[new_key] = {"name": tower_name_input.strip(), "ip": tower_ip_input.strip()}
                            save_towers(towers_dict)
                            log_action(st.session_state.username, f"إضافة برج/لينك جديد: '{tower_name_input}' ({tower_ip_input})")
                            st.success(f"تم إضافة البرج ({tower_name_input}) بنجاح!")
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            st.warning("يرجى إدخال اسم البرج وعنوان الآيبي بشكل صحيح.")

            try:
                connection = get_api_connection(mikrotik_conf)
                api = connection.get_api()
                towers_dict = load_towers()

                if towers_dict:
                    search_tower = st.text_input("🔍 ابحث في الأبراج (بالاسم أو عنوان IP):", value="")
                    if search_tower:
                        log_action(st.session_state.username, f"البحث في شاشة الأبراج بكلمة: '{search_tower}'")

                    filtered_towers = {}
                    for t_key, t_info in towers_dict.items():
                        if not search_tower or search_tower.lower() in t_info["name"].lower() or search_tower.lower() in t_info["ip"].lower():
                            filtered_towers[t_key] = t_info

                    total_towers_count = len(towers_dict)
                    filtered_towers_count = len(filtered_towers)
                    st.markdown(f"📊 **التعداد:** عرض **{filtered_towers_count}** برج مطابق (من إجمالي **{total_towers_count}** برج مسجل).")
                    st.write("")

                    if filtered_towers:
                        t_cols = st.columns(min(len(filtered_towers), 3) if len(filtered_towers) > 0 else 1)
                        idx = 0
                        for t_key, t_info in list(filtered_towers.items()):
                            card_num = idx + 1
                            is_ok, rtt_detail = check_tower_status(api, t_info["ip"])
                            badge_style = "background: rgba(34, 197, 94, 0.15); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.3);" if is_ok else "background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3);"
                            status_text = "🟢 يعمل" if is_ok else "🔴 طافي"

                            with t_cols[idx % len(t_cols)]:
                                st.markdown(f"""
                                    <div class="modern-card">
                                        <div class="card-title-bar" style="margin-bottom: 8px; padding-bottom: 8px;">
                                            <span class="client-name"><span style="color: #f59e0b; margin-left: 8px;">#{card_num}</span>📡 {t_info['name']}</span>
                                            <span class="{badge_style}" style="padding: 3px 10px; border-radius: 12px; font-size: 0.8rem;">{status_text}</span>
                                        </div>
                                        <div style="font-size: 0.95rem; color: #cbd5e1; margin-bottom: 12px;">
                                            🌐 الآيبي: <a href="http://{t_info['ip']}" target="_blank" style="color:#38bdf8; text-decoration: underline;" dir="ltr">{t_info['ip']}</a><br>
                                            ⏱️ الحالة: <b dir="ltr" style="color:#4ade80;">{rtt_detail}</b>
                                        </div>
                                """, unsafe_allow_html=True)
                                
                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1:
                                    with st.popover("✏️ تعديل", use_container_width=True):
                                        st.markdown(f"**تعديل بيانات {t_info['name']}**")
                                        edit_name = st.text_input("اسم البرج", value=t_info['name'], key=f"e_name_{t_key}")
                                        edit_ip = st.text_input("عنوان IP", value=t_info['ip'], key=f"e_ip_{t_key}")
                                        if st.button("💾 حفظ التعديل", key=f"save_e_{t_key}", type="primary", use_container_width=True):
                                            towers_dict[t_key] = {"name": edit_name.strip(), "ip": edit_ip.strip()}
                                            save_towers(towers_dict)
                                            log_action(st.session_state.username, f"تعديل بيانات البرج: '{edit_name}' ({edit_ip})")
                                            st.success("تم الحفظ!")
                                            st.rerun()

                                with col_btn2:
                                    if st.button("🗑️ حذف", key=f"del_t_{t_key}", use_container_width=True):
                                        log_action(st.session_state.username, f"حذف البرج/اللينك: '{t_info['name']}' ({t_info['ip']})")
                                        del towers_dict[t_key]
                                        save_towers(towers_dict)
                                        st.success("تم الحذف بنجاح")
                                        st.rerun()

                                st.markdown("</div>", unsafe_allow_html=True)
                            idx += 1
                    else:
                        st.info("❌ لا توجد أبراج مطابقة لكلمة البحث الحالية.")
                else:
                    st.info("لا توجد أبراج أو روابط مضافة حتى الآن. قم بإضافة برج جديد من النموذج أعلاه.")

                connection.disconnect()

                if auto_refresh_towers:
                    time.sleep(tower_refresh_rate)
                    st.rerun()

            except Exception as e:
                st.error(f"خطأ أثناء فحص الأبراج: {e}")

        # --- شاشة اكتشاف قطع الشبكة (IP Neighbors) على شكل بطاقات تفاعلية ---
        elif admin_menu == "📡 اكتشاف قطع الشبكة (IP Neighbors)":
            st.title("📡 اكتشاف قطع الشبكة وأجهزة الجيران (IP Neighbors)")
            st.markdown("💡 تعرض هذه الشاشة الأجهزة المكتشفة تلقائياً عبر بروتوكول IP Neighbors على شكل **بطاقات تفاعلية** مع إمكانية البحث والفلترة والتعداد اللحظي.")
            st.markdown("---")

            col_nbr1, col_nbr2 = st.columns([2, 3])
            with col_nbr1:
                auto_refresh_nbr = st.checkbox("🔄 تفعيل التحديث التلقائي للجيران", value=False)
            with col_nbr2:
                nbr_refresh_rate = st.slider("⏱️ مدة التحديث التلقائي (ثواني):", min_value=3, max_value=30, value=5, step=1)

            st.markdown("---")

            try:
                connection = get_api_connection(mikrotik_conf)
                api = connection.get_api()
                neighbors_res = api.get_resource("ip/neighbor").get()
                connection.disconnect()

                if neighbors_res:
                    processed_neighbors = []
                    for n in neighbors_res:
                        norm_n = {}
                        for k, v in n.items():
                            k_str = k.decode('utf-8') if isinstance(k, bytes) else str(k)
                            v_str = v.decode('utf-8') if isinstance(v, bytes) else str(v)
                            norm_n[k_str] = v_str
                        processed_neighbors.append({
                            "name": norm_n.get("identity", "غير معروف"),
                            "platform": norm_n.get("platform", norm_n.get("board-name", "غير محدد")),
                            "ip": norm_n.get("address", "غير متوفر"),
                            "mac": norm_n.get("mac-address", "-"),
                            "interface": norm_n.get("interface", "-")
                        })

                    search_nbr = st.text_input("🔍 ابحث في القطع (بالاسم، عنوان IP، الماك، أو المنفذ):", value="")
                    if search_nbr:
                        log_action(st.session_state.username, f"البحث في شاشة استكشاف القطع بكلمة: '{search_nbr}'")

                    if search_nbr:
                        filtered_nbrs = [
                            item for item in processed_neighbors 
                            if search_nbr.lower() in item["name"].lower() or 
                               search_nbr.lower() in item["ip"].lower() or 
                               search_nbr.lower() in item["platform"].lower() or 
                               search_nbr.lower() in item["mac"].lower() or 
                               search_nbr.lower() in item["interface"].lower()
                        ]
                    else:
                        filtered_nbrs = processed_neighbors

                    total_count = len(processed_neighbors)
                    filtered_count = len(filtered_nbrs)
                    st.markdown(f"📊 **التعداد:** تم العثور على **{filtered_count}** قطعة مطابقة (من إجمالي **{total_count}** جهاز متصل).")
                    st.write("")

                    if filtered_nbrs:
                        for i in range(0, len(filtered_nbrs), 2):
                            cols = st.columns(2)
                            for col_idx, dev_idx in enumerate([i, i + 1]):
                                if dev_idx < len(filtered_nbrs):
                                    dev = filtered_nbrs[dev_idx]
                                    dev_ip = dev["ip"]
                                    dev_name = dev["name"]
                                    dev_platform = dev["platform"]
                                    dev_mac = dev["mac"]
                                    dev_interface = dev["interface"]
                                    idx_num = dev_idx + 1

                                    ip_html = f'<a href="http://{dev_ip}" target="_blank" style="color:#38bdf8; text-decoration: underline;" dir="ltr">{dev_ip} 🔗</a>' if dev_ip != "غير متوفر" else '<b dir="ltr" style="color:#38bdf8;">-</b>'

                                    with cols[col_idx]:
                                        st.markdown(f"""
                                            <div class="modern-card">
                                                <div class="card-title-bar">
                                                    <span class="client-name"><span style="color: #f59e0b; margin-left: 8px;">#{idx_num}</span>📡 {dev_name}</span>
                                                    <span style="background: rgba(56, 189, 248, 0.15); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.3); padding: 5px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700;">{dev_platform}</span>
                                                </div>
                                                <div class="info-grid">
                                                    <div class="info-item">🌐 الآيبي: {ip_html}</div>
                                                    <div class="info-item">🔌 المنفذ: <b style="color:#f59e0b;">{dev_interface}</b></div>
                                                    <div class="info-item" style="grid-column: span 2;">💻 الماك: <b dir="ltr" style="color:#38bdf8;">{dev_mac}</b></div>
                                                </div>
                                            </div>
                                        """, unsafe_allow_html=True)
                    else:
                        st.info("❌ لا توجد قطع أو أجهزة مطابقة لكلمة البحث الحالية.")
                else:
                    st.info("لا توجد أجهزة مكتشفة حالياً عبر IP Neighbors في المايكروتك.")

                if auto_refresh_nbr:
                    time.sleep(nbr_refresh_rate)
                    st.rerun()

            except Exception as e:
                st.error(f"خطأ أثناء جلب بيانات IP Neighbors: {e}")

        elif admin_menu == "📊 تصدير التقارير (Excel)":
            st.title("📊 مركز التقارير وتحليل بيانات الزبائن")
            st.markdown("---")
            try:
                connection = get_api_connection(mikrotik_conf)
                api = connection.get_api()
                ppp_secrets = api.get_resource("ppp/secret").get()
                active_users = api.get_resource("ppp/active").get()
                active_dict = {act.get("name"): act for act in active_users}
                report_list = []

                for sec in ppp_secrets:
                    uname = sec.get("name", "")
                    is_disabled = sec.get("disabled", "false") == "true"
                    raw_comment = sec.get("comment", "")
                    decoded_comment = decode_comment(raw_comment)
                    prof_name = sec.get("profile", "افتراضي")
                    act_info = active_dict.get(uname)
                    is_online = act_info is not None

                    if is_disabled: status = "معطل"
                    elif is_online: status = "متصل"
                    else: status = "غير متصل"

                    secret_ip = sec.get("remote-address", "").strip()
                    if secret_ip and secret_ip != "0.0.0.0":
                        client_ip = secret_ip
                    else:
                        client_ip = act_info.get("address", "-") if active_info else "-"

                    user_speed_mb = extract_speed_mb(prof_name)
                    report_list.append({
                        "اسم المشترك / التعليق": decoded_comment, "اسم المستخدم (User)": uname, "الحالة": status,
                        "عنوان IP": client_ip, "عنوان MAC": act_info.get("caller-id", "-") if active_info else "-",
                        "مدة الاتصال": act_info.get("uptime", "-") if active_info else "-", "البروفايل": prof_name, "السرعة (Mbps)": user_speed_mb,
                    })
                connection.disconnect()

                df_users = pd.DataFrame(report_list)
                df_profiles = df_users["البروفايل"].value_counts().reset_index()
                df_profiles.columns = ["البروفايل", "عدد المشتركين"]
                df_profiles["السرعة (Mbps)"] = df_profiles["البروفايل"].apply(extract_speed_mb)
                df_profiles["إجمالي السعة (Mbps)"] = df_profiles["السرعة (Mbps)"] * df_profiles["عدد المشتركين"]

                excel_data = build_styled_excel(df_users, df_profiles)
                if st.download_button(
                    label="📥 تنزيل التقرير المنسق (Excel)", data=excel_data,
                    file_name=f"subscribers_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True
                ):
                    log_action(st.session_state.username, "تصدير وتنزيل تقرير Excel شامل للشبكة")
            except Exception as e: st.error(f"خطأ: {e}")

        elif admin_menu == "👥 إدارة طاقم العمل وفريق الدعم":
            st.title("👥 إدارة مستخدمي النظام وسجل الحركات")
            st.markdown("---")
            all_sys_users = load_users()
            tab_manage, tab_logs = st.tabs(["👥 إدارة المستخدمين", "📜 سجل حركات الفريق"])
            with tab_manage:
                with st.expander("➕ إضافة عضو جديد للنظام"):
                    with st.form("add_staff_form"):
                        c1, c2, c3 = st.columns(3)
                        with c1: new_st_user = st.text_input("اسم المستخدم")
                        with c2: new_st_pass = st.text_input("كلمة المرور", type="password")
                        with c3: new_st_role = st.selectbox("الصلاحية", ["طاقم فني", "مدير النظام"])
                        if st.form_submit_button("إضافة الحساب", type="primary", use_container_width=True):
                            if new_st_user and new_st_pass:
                                all_sys_users[new_st_user] = {"password": make_hash(new_st_pass), "role": new_st_role, "disabled": False, "last_active": ""}
                                save_users(all_sys_users)
                                log_action(st.session_state.username, f"إضافة حساب مستخدم جديد للنظام: '{new_st_user}' بصلاحية '{new_st_role}'")
                                st.success("تم الحفظ بنجاح"); st.rerun()

                for uname, udata in list(all_sys_users.items()):
                    cols = st.columns([2, 2, 2, 2, 1])
                    with cols[0]: st.markdown(f"**👤 {uname}**")
                    with cols[1]: st.markdown(f"🛡️ {udata.get('role', '')}")
                    with cols[2]:
                        new_p = st.text_input(f"باسورد لـ {uname}", type="password", key=f"pass_{uname}")
                        if st.button("تحديث", key=f"btn_p_{uname}"):
                            if new_p:
                                all_sys_users[uname]["password"] = make_hash(new_p)
                                save_users(all_sys_users)
                                log_action(st.session_state.username, f"تغيير كلمة المرور للمستخدم: '{uname}'")
                                st.success("تم التحديث"); st.rerun()
                    with cols[3]:
                        if uname != "admin":
                            if udata.get("disabled", False):
                                if st.button("تفعيل", key=f"tog_{uname}"):
                                    all_sys_users[uname]["disabled"] = False
                                    save_users(all_sys_users)
                                    log_action(st.session_state.username, f"تفعيل حساب مستخدم النظام: '{uname}'")
                                    st.rerun()
                            else:
                                if st.button("تعطيل", key=f"tog_{uname}"):
                                    all_sys_users[uname]["disabled"] = True
                                    save_users(all_sys_users)
                                    log_action(st.session_state.username, f"تعطيل حساب مستخدم النظام: '{uname}'")
                                    st.rerun()
                    with cols[4]:
                        if uname != "admin" and st.button("حذف", key=f"del_{uname}"):
                            log_action(st.session_state.username, f"حذف حساب مستخدم النظام: '{uname}'")
                            del all_sys_users[uname]
                            save_users(all_sys_users)
                            st.rerun()
                    st.markdown("---")
            with tab_logs:
                log_data = get_all_logs()
                if log_data:
                    search_log = st.text_input("🔍 تصفية سجل الحركات (حسب اسم المستخدم أو الحركة):", value="")
                    if search_log:
                        log_data = [l for l in log_data if search_log.lower() in l['user'].lower() or search_log.lower() in l['action'].lower()]

                    st.markdown(f"📋 **إجمالي الحركات المسجلة المعروضة:** **{len(log_data)}** حركة (عرض كافة الحركات بدون استثناء).")
                    st.write("")
                    df_logs = pd.DataFrame(log_data)
                    df_logs.columns = ["التاريخ والوقت", "اسم المستخدم", "تفاصيل الحركة والإجراء"]
                    st.dataframe(df_logs, use_container_width=True, hide_index=True)
                else: 
                    st.info("لا توجد حركات مسجلة.")

        elif admin_menu == "⚙️ إعدادات المايكروتك":
            st.title("⚙️ تعديل بيانات الاتصال بالمايكروتك")
            st.markdown("---")
            current_conf = load_mikrotik_config()
            with st.form("edit_mikrotik_config_file"):
                new_host = st.text_input("عنوان IP للراوتر", value=current_conf.get("host", ""))
                new_user = st.text_input("اسم مستخدم API", value=current_conf.get("user", ""))
                new_port = st.number_input("منفذ API", value=int(current_conf.get("port", 8728)), step=1)
                new_pass = st.text_input("كلمة المرور", value=current_conf.get("password", ""), type="password")

                if st.form_submit_button("💾 حفظ الإعدادات", type="primary", use_container_width=True):
                    updated_conf = {"host": new_host.strip(), "user": new_user.strip(), "password": new_pass.strip(), "port": int(new_port)}
                    save_mikrotik_config(updated_conf)
                    log_action(st.session_state.username, f"تعديل إعدادات المايكروتك (IP: {new_host}, Port: {new_port})")
                    st.success("تم حفظ الإعدادات بنجاح!")
                    time.sleep(1); st.rerun()