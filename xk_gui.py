import base64
import json
import os
import random
import re
import threading
import time
from datetime import datetime
from typing import Any, Dict, List

import requests
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad


PREFS_FILE = "xk_gui_prefs.json"
DEFAULT_XK_HOST = "https://xk.neusoft.edu.cn"
DEFAULT_OPENAI_BASE_URL = ""
DEFAULT_OPENAI_MODEL = ""
DEFAULT_RETRY_TIMES = 3
DEFAULT_TEACHING_CLASS_TYPE = "XGKC"
DEFAULT_PAGE_SIZE = 100
DEFAULT_OUTPUT_FILE = "xk_courses.json"
DEFAULT_CAPTCHA_MODE = "本地OCR"
DEFAULT_AES_KEY = "MWMqg2tPcDkxcm11"
DEFAULT_LIST_PAGE_INTERVAL_MS = 1800
DEFAULT_LIST_403_RETRY = 10

_DDDD_OCR: Any = None


def _load_prefs(file_path: str = PREFS_FILE) -> Dict[str, Any]:
    if not os.path.exists(file_path):
        return {}
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _save_prefs(data: Dict[str, Any], file_path: str = PREFS_FILE) -> None:
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _get_dddd_ocr() -> Any:
    global _DDDD_OCR
    if _DDDD_OCR is not None:
        return _DDDD_OCR
    try:
        import ddddocr  # type: ignore
        if hasattr(ddddocr, "DdddOcr"):
            _DDDD_OCR = ddddocr.DdddOcr(show_ad=False)
            return _DDDD_OCR
    except Exception as exc:
        pass

    try:
        from ddddocr.core.ocr_engine import OCREngine  # type: ignore
        _DDDD_OCR = OCREngine()
    except Exception as exc:
        raise RuntimeError("ddddocr 初始化失败，请检查版本兼容或重装 ddddocr") from exc

    return _DDDD_OCR


def _extract_aes_key(html: str) -> str:
    match = re.search(r'loginVue\.loginForm\.aesKey\s*=\s*"([^"]+)"', html)
    if not match:
        raise RuntimeError("未在 index.html 中找到 aesKey")
    return match.group(1)


def _encrypt_password(plain_password: str, aes_key: str) -> str:
    key_bytes = aes_key.encode("utf-8")
    if len(key_bytes) not in (16, 24, 32):
        raise RuntimeError(f"aesKey 长度异常: {len(key_bytes)}，期望 16/24/32")
    cipher = AES.new(key_bytes, AES.MODE_ECB)
    encrypted = cipher.encrypt(pad(plain_password.encode("utf-8"), AES.block_size, style="pkcs7"))
    return base64.b64encode(encrypted).decode("utf-8")


def _solve_captcha_with_openai(
    captcha_data_url: str,
    openai_api_key: str,
    openai_base_url: str,
    openai_model: str,
) -> str:
    payload: Dict[str, Any] = {
        "model": openai_model,
        "temperature": 0,
        "max_tokens": 20,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "识别这张验证码图片中的字符，只返回验证码本身，不要任何解释。",
                    },
                    {
                        "type": "image_url",
                        "image_url": {"url": captcha_data_url},
                    },
                ],
            }
        ],
    }

    response = requests.post(
        f"{openai_base_url}/chat/completions",
        headers={
            "Authorization": f"Bearer {openai_api_key}",
            "Content-Type": "application/json",
        },
        data=json.dumps(payload, ensure_ascii=False),
        timeout=30,
    )
    response.raise_for_status()
    content = response.json()["choices"][0]["message"]["content"]
    text = re.sub(r"\s+", "", str(content)).strip()
    text = re.sub(r"[^0-9A-Za-z]", "", text)
    if not text:
        raise RuntimeError(f"OpenAI 返回内容无法解析验证码: {content!r}")
    return text[:8]


def _solve_captcha_with_ddddocr(captcha_data_url: str) -> str:
    ocr = _get_dddd_ocr()
    raw_b64 = _get_data_url_base64(captcha_data_url)
    image_bytes = base64.b64decode(raw_b64)
    if hasattr(ocr, "classification"):
        text = str(ocr.classification(image_bytes) or "")
    elif hasattr(ocr, "predict"):
        text = str(ocr.predict(image_bytes) or "")
    else:
        raise RuntimeError("ddddocr 对象不支持 classification/predict")
    text = re.sub(r"\s+", "", text).strip()
    text = re.sub(r"[^0-9A-Za-z]", "", text)
    if not text:
        raise RuntimeError(f"ddddocr 返回内容无法解析验证码: {text!r}")
    return text[:8]


def _is_captcha_error(login_json: Dict[str, Any]) -> bool:
    msg = str(login_json.get("msg") or login_json.get("message") or "")
    return ("验证码" in msg) or ("captcha" in msg.lower())


def _is_time_conflict_unselectable(course: Dict[str, Any]) -> bool:
    conflict_desc = str(course.get("conflictDesc") or "").strip()
    sfkt = str(course.get("SFKT") or "").strip()
    if not conflict_desc:
        return False
    if sfkt == "0":
        return True
    return "冲突" in conflict_desc


def _has_conflict(course: Dict[str, Any]) -> bool:
    conflict_desc = str(course.get("conflictDesc") or "").strip()
    if conflict_desc and "冲突" in conflict_desc:
        return True
    sfct = str(course.get("SFCT") or "").strip()
    return sfct == "1"


def _is_selectable(course: Dict[str, Any]) -> bool:
    sfkt = str(course.get("SFKT") or "").strip()
    if _has_conflict(course):
        return False
    return sfkt == "1"


def _selectable_flag(course: Dict[str, Any]) -> str:
    return "1" if _is_selectable(course) else "0"


def _get_teaching_place(course: Dict[str, Any]) -> str:
    return str(
        course.get("YPSJDD")
        or course.get("teachingPlace")
        or course.get("teachingPlaceHide")
        or ""
    ).strip()


def _get_data_url_base64(data_url: str) -> str:
    if "," not in data_url:
        return data_url
    return data_url.split(",", 1)[1]


def _ask_captcha_manually(root: tk.Tk, captcha_data_url: str) -> str:
    raw_b64 = _get_data_url_base64(captcha_data_url)

    dialog = tk.Toplevel(root)
    dialog.title("请输入验证码")
    dialog.geometry("340x220")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()

    ttk.Label(dialog, text="请根据图片输入验证码：").pack(pady=(12, 8))

    photo = tk.PhotoImage(data=raw_b64)
    image_label = ttk.Label(dialog, image=photo)
    image_label.pack(pady=(0, 8))

    value_var = tk.StringVar()
    entry = ttk.Entry(dialog, textvariable=value_var, width=20)
    entry.pack()
    entry.focus_set()

    result: Dict[str, str] = {"value": ""}

    def on_ok() -> None:
        result["value"] = value_var.get().strip()
        dialog.destroy()

    def on_cancel() -> None:
        result["value"] = ""
        dialog.destroy()

    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=12)
    ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=6)
    ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=6)

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    root.wait_window(dialog)

    code = re.sub(r"[^0-9A-Za-z]", "", result["value"]).strip()
    if not code:
        raise RuntimeError("未输入验证码")
    return code


def _heartbeat(session: requests.Session, base_api: str) -> None:
    resp = session.post(f"{base_api}/web/now", timeout=10)
    resp.raise_for_status()


def _select_course_once(
    session: requests.Session,
    base_host: str,
    base_api: str,
    batch_id: str,
    clazz_type: str,
    clazz_id: str,
    secret_val: str,
) -> Dict[str, Any]:
    resp = session.post(
        f"{base_api}/elective/clazz/add",
        data={
            "clazzType": clazz_type,
            "clazzId": clazz_id,
            "secretVal": secret_val,
        },
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "batchId": batch_id,
            "Referer": f"{base_api}/elective/grablessons?batchId={batch_id}",
            "Origin": base_host,
        },
        timeout=12,
    )
    resp.raise_for_status()
    return resp.json()


def _fetch_all_courses(
    session: requests.Session,
    base_host: str,
    base_api: str,
    token: str,
    batch_id: str,
    teaching_class_type: str,
    campus: str,
    page_size: int,
    page_interval_ms: int = DEFAULT_LIST_PAGE_INTERVAL_MS,
    retry_on_403: int = DEFAULT_LIST_403_RETRY,
) -> Dict[str, Any]:
    session.headers["Authorization"] = token

    rows: List[Dict[str, Any]] = []
    page_number = 1
    total = None
    rate_limit_markers = ("请求过快", "请登录后再试", "过于频繁")

    while True:
        payload = {
            "teachingClassType": teaching_class_type,
            "pageNumber": page_number,
            "pageSize": page_size,
            "orderBy": "",
            "campus": campus,
        }
        body: Dict[str, Any] = {}
        last_error = None
        for attempt in range(retry_on_403 + 1):
            resp = session.post(
                f"{base_api}/elective/clazz/list",
                headers={
                    "Content-Type": "application/json;charset=UTF-8",
                    "batchId": batch_id,
                    "Referer": f"{base_api}/elective/grablessons?batchId={batch_id}",
                    "Origin": base_host,
                },
                data=json.dumps(payload, ensure_ascii=False),
                timeout=20,
            )
            resp.raise_for_status()
            body = resp.json()
            if body.get("code") == 200:
                break

            msg_text = str(body.get("msg") or "")
            is_too_fast = str(body.get("code")) == "403" and any(marker in msg_text for marker in rate_limit_markers)
            if is_too_fast and attempt < retry_on_403:
                try:
                    _heartbeat(session, base_api)
                except Exception:
                    pass

                base_wait = max(1.5, page_interval_ms / 1000.0)
                sleep_seconds = base_wait * (attempt + 1) + random.uniform(0.2, 0.8)
                time.sleep(sleep_seconds)
                continue

            last_error = body
            break

        if body.get("code") != 200:
            raise RuntimeError(f"课程列表接口失败(第{page_number}页): {last_error or body}")

        data = body.get("data") or {}
        page_rows = data.get("rows") or []
        if total is None:
            total = int(data.get("total") or 0)

        if not page_rows:
            break

        rows.extend(page_rows)
        if len(rows) >= total:
            break

        if page_interval_ms > 0:
            time.sleep((page_interval_ms / 1000.0) + random.uniform(0.03, 0.15))
        page_number += 1

    return {
        "batch_id": batch_id,
        "total": total or 0,
        "fetched": len(rows),
        "rows": rows,
    }


class XKGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("选课课程抓取")
        self.root.geometry("1100x700")

        self.prefs = _load_prefs()
        self.all_rows: List[Dict[str, Any]] = []
        self.view_rows: List[Dict[str, Any]] = []
        self.last_result: Dict[str, Any] = {}
        self.current_batch_id = ""
        self.current_session: requests.Session | None = None
        self.current_base_host = ""
        self.current_base_api = ""
        self.current_clazz_type = "XGKC"
        self.rob_thread: threading.Thread | None = None
        self.rob_pause_event = threading.Event()
        self.rob_stop_event = threading.Event()
        self.rob_running = False

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="学号:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.username_var = tk.StringVar(value=str(self.prefs.get("username") or ""))
        ttk.Entry(top, textvariable=self.username_var, width=24).grid(row=0, column=1, sticky=tk.W)

        ttk.Label(top, text="密码:").grid(row=0, column=2, sticky=tk.W, padx=(12, 6))
        self.password_var = tk.StringVar(value=str(self.prefs.get("password") or ""))
        ttk.Entry(top, textvariable=self.password_var, width=24, show="*").grid(row=0, column=3, sticky=tk.W)

        ttk.Label(top, text="验证码方式:").grid(row=0, column=4, sticky=tk.W, padx=(16, 6))
        self.captcha_mode_var = tk.StringVar(value=str(self.prefs.get("captcha_mode") or DEFAULT_CAPTCHA_MODE))
        ttk.Combobox(
            top,
            textvariable=self.captcha_mode_var,
            state="readonly",
            width=10,
            values=("本地OCR", "AI识别", "手动输入"),
        ).grid(row=0, column=5, sticky=tk.W)

        self.selectable_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top,
            text="仅显示可选课程",
            variable=self.selectable_only_var,
            command=self.on_filter_changed,
        ).grid(
            row=0, column=6, padx=(16, 0), sticky=tk.W
        )

        self.fetch_btn = ttk.Button(top, text="登录并抓取本轮课程", command=self.on_fetch)
        self.fetch_btn.grid(row=0, column=7, padx=(16, 0), sticky=tk.W)

        api = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        api.pack(fill=tk.X)

        ttk.Label(api, text="API Key:").grid(row=0, column=0, sticky=tk.W)
        self.api_key_var = tk.StringVar(value=str(self.prefs.get("api_key") or ""))
        ttk.Entry(api, textvariable=self.api_key_var, width=40, show="*").grid(row=0, column=1, padx=(6, 12), sticky=tk.W)

        ttk.Label(api, text="Base URL:").grid(row=0, column=2, sticky=tk.W)
        self.api_base_url_var = tk.StringVar(value=str(self.prefs.get("api_base_url") or DEFAULT_OPENAI_BASE_URL))
        ttk.Entry(api, textvariable=self.api_base_url_var, width=34).grid(row=0, column=3, padx=(6, 12), sticky=tk.W)

        ttk.Label(api, text="Model:").grid(row=0, column=4, sticky=tk.W)
        self.api_model_var = tk.StringVar(value=str(self.prefs.get("api_model") or DEFAULT_OPENAI_MODEL))
        ttk.Entry(api, textvariable=self.api_model_var, width=16).grid(row=0, column=5, padx=(6, 0), sticky=tk.W)

        ttk.Label(api, text="爬课间隔(ms):").grid(row=1, column=0, sticky=tk.W, pady=(6, 0))
        self.list_interval_ms_var = tk.StringVar(value=str(self.prefs.get("list_interval_ms") or DEFAULT_LIST_PAGE_INTERVAL_MS))
        ttk.Entry(api, textvariable=self.list_interval_ms_var, width=10).grid(row=1, column=1, padx=(6, 12), pady=(6, 0), sticky=tk.W)

        ttk.Label(api, text="403重试次数:").grid(row=1, column=2, sticky=tk.W, pady=(6, 0))
        self.list_403_retry_var = tk.StringVar(value=str(self.prefs.get("list_403_retry") or DEFAULT_LIST_403_RETRY))
        ttk.Entry(api, textvariable=self.list_403_retry_var, width=10).grid(row=1, column=3, padx=(6, 12), pady=(6, 0), sticky=tk.W)

        rob = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        rob.pack(fill=tk.X)

        ttk.Label(rob, text="开始方式:").grid(row=0, column=0, sticky=tk.W)
        self.start_mode_var = tk.StringVar(value="立即开始")
        self.start_mode_combo = ttk.Combobox(
            rob,
            textvariable=self.start_mode_var,
            state="readonly",
            width=10,
            values=("立即开始", "延时开始", "指定时刻"),
        )
        self.start_mode_combo.grid(row=0, column=1, padx=(6, 8), sticky=tk.W)
        self.start_mode_combo.bind("<<ComboboxSelected>>", self._on_start_mode_changed)

        self.delay_group = ttk.Frame(rob)
        self.delay_group.grid(row=0, column=2, columnspan=2, sticky=tk.W)
        ttk.Label(self.delay_group, text="延时(s):").pack(side=tk.LEFT)
        self.start_after_seconds_var = tk.StringVar(value="0")
        ttk.Spinbox(
            self.delay_group,
            from_=0,
            to=3600,
            textvariable=self.start_after_seconds_var,
            width=6,
        ).pack(side=tk.LEFT, padx=(6, 12))

        now = datetime.now()
        self.time_group = ttk.Frame(rob)
        self.time_group.grid(row=0, column=4, columnspan=6, sticky=tk.W)
        ttk.Label(self.time_group, text="指定时刻:").pack(side=tk.LEFT)
        self.start_hour_var = tk.StringVar(value=f"{now.hour:02d}")
        self.start_minute_var = tk.StringVar(value=f"{now.minute:02d}")
        self.start_second_var = tk.StringVar(value=f"{now.second:02d}")
        ttk.Combobox(
            self.time_group,
            textvariable=self.start_hour_var,
            state="readonly",
            width=3,
            values=tuple(f"{i:02d}" for i in range(24)),
        ).pack(side=tk.LEFT, padx=(6, 2))
        ttk.Label(self.time_group, text=":").pack(side=tk.LEFT)
        ttk.Combobox(
            self.time_group,
            textvariable=self.start_minute_var,
            state="readonly",
            width=3,
            values=tuple(f"{i:02d}" for i in range(60)),
        ).pack(side=tk.LEFT, padx=(2, 2))
        ttk.Label(self.time_group, text=":").pack(side=tk.LEFT)
        ttk.Combobox(
            self.time_group,
            textvariable=self.start_second_var,
            state="readonly",
            width=3,
            values=tuple(f"{i:02d}" for i in range(60)),
        ).pack(side=tk.LEFT, padx=(2, 12))

        ttk.Label(rob, text="点击间隔(ms):").grid(row=0, column=10, sticky=tk.W)
        self.rob_delay_ms_var = tk.StringVar(value="250")
        ttk.Entry(rob, textvariable=self.rob_delay_ms_var, width=8).grid(row=0, column=11, padx=(6, 12), sticky=tk.W)

        ttk.Label(rob, text="点击次数:").grid(row=0, column=12, sticky=tk.W)
        self.rob_click_times_var = tk.StringVar(value="50")
        ttk.Entry(rob, textvariable=self.rob_click_times_var, width=8).grid(row=0, column=13, padx=(6, 12), sticky=tk.W)

        self.keep_alive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(rob, text="保活防踢", variable=self.keep_alive_var).grid(row=0, column=14, padx=(0, 8), sticky=tk.W)

        ttk.Label(rob, text="保活间隔(s):").grid(row=0, column=15, sticky=tk.W)
        self.keep_alive_interval_var = tk.StringVar(value="25")
        ttk.Entry(rob, textvariable=self.keep_alive_interval_var, width=6).grid(row=0, column=16, padx=(6, 12), sticky=tk.W)

        self.rob_btn = ttk.Button(rob, text="抢选中课程", command=self.on_rob_course, state=tk.DISABLED)
        self.rob_btn.grid(row=0, column=17, sticky=tk.W)

        self.pause_btn = ttk.Button(rob, text="暂停抢课", command=self.on_toggle_pause, state=tk.DISABLED)
        self.pause_btn.grid(row=0, column=18, padx=(8, 0), sticky=tk.W)

        self.stop_btn = ttk.Button(rob, text="终止抢课", command=self.on_stop_rob, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=19, padx=(8, 0), sticky=tk.W)

        self._on_start_mode_changed()

        self.status_var = tk.StringVar(value="请先输入学号和密码")
        ttk.Label(self.root, textvariable=self.status_var, foreground="#1f4e79").pack(fill=tk.X, padx=10)

        columns = (
            "KCH",
            "KCM",
            "KXH",
            "SKJS",
            "XF",
            "XGXKLB",
            "classCapacity",
            "numberOfSelected",
            "teachingPlace",
            "SFKT",
        )
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=26, selectmode="extended")
        headers = {
            "KCH": "课程号",
            "KCM": "课程名",
            "KXH": "课序号",
            "SKJS": "教师",
            "XF": "学分",
            "XGXKLB": "通识选修课类别",
            "classCapacity": "课容量",
            "numberOfSelected": "已选人数",
            "teachingPlace": "上课时间地点",
            "SFKT": "可选(1是0否)",
        }
        widths = {
            "KCH": 130,
            "KCM": 220,
            "KXH": 80,
            "SKJS": 140,
            "XF": 70,
            "XGXKLB": 140,
            "classCapacity": 90,
            "numberOfSelected": 90,
            "teachingPlace": 560,
            "SFKT": 120,
        }
        for col in columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, width=widths[col], anchor=tk.W)

        yscroll = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        xscroll = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(10, 0))
        yscroll.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X, padx=(10, 10), pady=(0, 10))

    def _pick_batch_id(self, batch_list: List[Dict[str, Any]]) -> str:
        if not batch_list:
            raise RuntimeError("未获取到选课轮次")

        if len(batch_list) == 1:
            return str(batch_list[0].get("code", ""))

        lines = []
        default_index = 1
        for i, item in enumerate(batch_list, start=1):
            can_select = str(item.get("canSelect", ""))
            name = str(item.get("name", ""))
            term = str(item.get("schoolTermName", ""))
            begin_time = str(item.get("beginTime", ""))
            end_time = str(item.get("endTime", ""))
            lines.append(f"{i}. {name} | {term} | 可选:{can_select} | {begin_time} ~ {end_time}")
            if can_select == "1":
                default_index = i

        msg = "请选择本次抓取的选课轮次（输入序号）:\n\n" + "\n".join(lines)
        selected = simpledialog.askinteger(
            "选择选课轮次",
            msg,
            initialvalue=default_index,
            minvalue=1,
            maxvalue=len(batch_list),
            parent=self.root,
        )
        if selected is None:
            raise RuntimeError("你已取消轮次选择")

        return str(batch_list[selected - 1].get("code", ""))

    def _clear_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _render_courses(self, courses: List[Dict[str, Any]]) -> None:
        self._clear_tree()
        for row in courses:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row.get("KCH", ""),
                    row.get("KCM", ""),
                    row.get("KXH", ""),
                    row.get("SKJS", ""),
                    row.get("XF", ""),
                    row.get("XGXKLB", ""),
                    row.get("classCapacity", row.get("KRL", "")),
                    row.get("numberOfSelected", row.get("YXRS", "")),
                    _get_teaching_place(row),
                    _selectable_flag(row),
                ),
            )

    def _apply_filter_and_render(self) -> None:
        if self.selectable_only_var.get():
            show_rows = [r for r in self.all_rows if _is_selectable(r)]
        else:
            show_rows = self.all_rows

        self.view_rows = show_rows

        self._render_courses(show_rows)

        if self.last_result:
            batch_id = self.last_result.get("batch_id", "")
            fetched = self.last_result.get("fetched", 0)
            total = self.last_result.get("total", 0)
            self.status_var.set(
                f"完成：轮次 {batch_id}，抓取 {fetched}/{total}，显示 {len(show_rows)} 条"
            )

    def on_filter_changed(self) -> None:
        if not self.all_rows:
            return
        self._apply_filter_and_render()

    def _on_start_mode_changed(self, event: Any = None) -> None:
        mode = self.start_mode_var.get().strip()
        if mode == "立即开始":
            self.delay_group.grid_remove()
            self.time_group.grid_remove()
        elif mode == "延时开始":
            self.delay_group.grid()
            self.time_group.grid_remove()
        else:
            self.delay_group.grid_remove()
            self.time_group.grid()

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))

    def _persist_prefs(self) -> None:
        self.prefs.update(
            {
                "username": self.username_var.get().strip(),
                "password": self.password_var.get().strip(),
                "api_key": self.api_key_var.get().strip(),
                "api_base_url": self.api_base_url_var.get().strip(),
                "api_model": self.api_model_var.get().strip(),
                "captcha_mode": self.captcha_mode_var.get().strip() or DEFAULT_CAPTCHA_MODE,
                "list_interval_ms": self.list_interval_ms_var.get().strip(),
                "list_403_retry": self.list_403_retry_var.get().strip(),
            }
        )
        _save_prefs(self.prefs)

    def _ui_enable_after_rob(self) -> None:
        self.fetch_btn.config(state=tk.NORMAL)
        self.rob_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(text="暂停抢课")
        self.rob_pause_event.clear()
        self.rob_stop_event.clear()
        self.rob_running = False

    def on_fetch(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username or not password:
            messagebox.showerror("输入错误", "请输入学号和密码")
            return

        self._persist_prefs()

        api_key = self.api_key_var.get().strip()
        base_url = self.api_base_url_var.get().strip().rstrip("/")
        model = self.api_model_var.get().strip()
        captcha_mode = self.captcha_mode_var.get().strip() or DEFAULT_CAPTCHA_MODE
        if captcha_mode == "AI识别":
            if not api_key:
                messagebox.showerror("输入错误", "AI识别模式下 API Key 不能为空")
                return
            if not base_url:
                messagebox.showerror("输入错误", "AI识别模式下 Base URL 不能为空")
                return
            if not model:
                messagebox.showerror("输入错误", "AI识别模式下 Model 不能为空")
                return

        base_host = (str(self.prefs.get("xk_host") or os.getenv("XK_HOST") or DEFAULT_XK_HOST)).strip().rstrip("/")
        if not re.match(r"^https?://", base_host, re.IGNORECASE):
            messagebox.showerror("参数错误", "选课系统地址必须以 http:// 或 https:// 开头")
            return
        base_api = f"{base_host}/xsxk"
        base_profile = f"{base_api}/profile"
        retry_times = DEFAULT_RETRY_TIMES
        teaching_class_type = DEFAULT_TEACHING_CLASS_TYPE
        page_size = DEFAULT_PAGE_SIZE
        output_file = DEFAULT_OUTPUT_FILE
        try:
            list_interval_ms = int(self.list_interval_ms_var.get().strip())
            list_403_retry = int(self.list_403_retry_var.get().strip())
        except ValueError:
            messagebox.showerror("参数错误", "爬课间隔和403重试次数必须是整数")
            return
        if list_interval_ms < 0 or list_403_retry < 0:
            messagebox.showerror("参数错误", "爬课间隔和403重试次数必须 >= 0")
            return

        effective_list_interval_ms = max(1500, list_interval_ms)

        self.fetch_btn.config(state=tk.DISABLED)
        self.status_var.set("正在登录并识别验证码...")
        self.root.update_idletasks()

        try:
            session = requests.Session()
            session.headers.update(
                {
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/145.0.0.0 Safari/537.36"
                    ),
                    "Accept": "application/json, text/plain, */*",
                    "Referer": f"{base_profile}/index.html",
                    "Origin": base_host,
                }
            )

            aes_key = DEFAULT_AES_KEY

            encrypted_password = _encrypt_password(password, aes_key)

            login_json: Dict[str, Any] = {}
            for attempt in range(1, retry_times + 1):
                captcha_resp = session.post(f"{base_api}/auth/captcha", timeout=20)
                captcha_resp.raise_for_status()
                captcha_json = captcha_resp.json()
                if captcha_json.get("code") != 200:
                    raise RuntimeError(f"验证码接口失败: {captcha_json}")

                captcha_data = captcha_json.get("data") or {}
                captcha_data_url = str(captcha_data.get("captcha") or "")
                uuid = str(captcha_data.get("uuid") or "")
                captcha_type = str(captcha_data.get("type") or "")
                if not captcha_data_url or not uuid:
                    raise RuntimeError(f"验证码返回不完整: {captcha_json}")
                if captcha_type != "1":
                    raise RuntimeError(f"当前验证码类型为 {captcha_type}（仅支持文本验证码 type=1）")

                if captcha_mode == "AI识别":
                    captcha_code = _solve_captcha_with_openai(
                        captcha_data_url=captcha_data_url,
                        openai_api_key=api_key,
                        openai_base_url=base_url,
                        openai_model=model,
                    )
                elif captcha_mode == "本地OCR":
                    captcha_code = _solve_captcha_with_ddddocr(captcha_data_url)
                else:
                    captcha_code = _ask_captcha_manually(self.root, captcha_data_url)
                self.status_var.set(f"验证码识别成功({attempt}/{retry_times}): {captcha_code}，正在登录...")
                self.root.update_idletasks()

                login_resp = session.post(
                    f"{base_api}/auth/login",
                    data={
                        "loginname": username,
                        "password": encrypted_password,
                        "captcha": captcha_code,
                        "uuid": uuid,
                    },
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    timeout=20,
                )
                login_resp.raise_for_status()
                login_json = login_resp.json()

                if login_json.get("code") == 200:
                    break

                if _is_captcha_error(login_json) and attempt < retry_times:
                    self.status_var.set(f"验证码错误，自动重试 {attempt + 1}/{retry_times}...")
                    self.root.update_idletasks()
                    continue

                raise RuntimeError(f"登录失败: {login_json}")

            if login_json.get("code") != 200:
                raise RuntimeError(f"登录失败: {login_json}")

            token = str((login_json.get("data") or {}).get("token") or "")
            if not token:
                raise RuntimeError("登录成功但未拿到 token")

            batch_list = ((login_json.get("data") or {}).get("student") or {}).get("electiveBatchList") or []
            batch_id = self._pick_batch_id(batch_list)
            if not batch_id:
                raise RuntimeError("选课轮次为空")

            campus = str(((login_json.get("data") or {}).get("student") or {}).get("campus") or "1")
            self.status_var.set("正在抓取课程列表...")
            self.root.update_idletasks()

            result = _fetch_all_courses(
                session=session,
                base_host=base_host,
                base_api=base_api,
                token=token,
                batch_id=batch_id,
                teaching_class_type=teaching_class_type,
                campus=campus,
                page_size=page_size,
                page_interval_ms=effective_list_interval_ms,
                retry_on_403=list_403_retry,
            )
            self.all_rows = result["rows"]
            self.last_result = result
            self.current_batch_id = batch_id
            self.current_session = session
            self.current_base_host = base_host
            self.current_base_api = base_api
            self.current_clazz_type = teaching_class_type

            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            self._apply_filter_and_render()
            self.status_var.set(self.status_var.get() + f"，已保存 {output_file}")
            self.rob_btn.config(state=tk.NORMAL)
            self.pause_btn.config(state=tk.DISABLED)

        except Exception as exc:
            self.status_var.set("执行失败")
            messagebox.showerror("错误", str(exc))
        finally:
            self.fetch_btn.config(state=tk.NORMAL)

    def on_rob_course(self) -> None:
        if self.rob_running:
            messagebox.showwarning("抢课中", "当前已有抢课任务在执行")
            return

        if not self.current_session or not self.current_batch_id or not self.current_base_api:
            messagebox.showerror("未准备好", "请先登录并抓取课程")
            return

        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("未选择课程", "请先在表格中选择至少一门课程")
            return

        selected_courses: List[Dict[str, Any]] = []
        invalid_course_names: List[str] = []
        for item_id in selected:
            try:
                row_index = self.tree.index(item_id)
                course = self.view_rows[row_index]
            except Exception:
                continue

            clazz_id = str(course.get("JXBID") or "").strip()
            secret_val = str(course.get("secretVal") or "").strip()
            if not clazz_id or not secret_val:
                invalid_course_names.append(str(course.get("KCM") or course.get("KXH") or "未知课程"))
                continue

            selected_courses.append(course)

        if not selected_courses:
            messagebox.showerror("数据缺失", "所选课程都缺少 JXBID 或 secretVal，无法抢课")
            return

        if invalid_course_names:
            preview = "，".join(invalid_course_names[:3])
            suffix = "..." if len(invalid_course_names) > 3 else ""
            self.status_var.set(f"已跳过 {len(invalid_course_names)} 门数据不完整课程：{preview}{suffix}")

        try:
            delay_ms = int(self.rob_delay_ms_var.get().strip())
            click_times = int(self.rob_click_times_var.get().strip())
            keep_alive_interval = int(self.keep_alive_interval_var.get().strip())
        except ValueError:
            messagebox.showerror("参数错误", "延迟/点击次数/保活间隔必须是整数")
            return

        if delay_ms < 0 or click_times < 1 or keep_alive_interval < 1:
            messagebox.showerror("参数错误", "延迟>=0，点击次数>=1，保活间隔>=1")
            return

        start_mode = self.start_mode_var.get().strip()
        start_at = None
        if start_mode == "延时开始":
            try:
                start_after_seconds = int(self.start_after_seconds_var.get().strip())
            except ValueError:
                messagebox.showerror("参数错误", "延时秒数必须是整数")
                return
            if start_after_seconds < 0:
                messagebox.showerror("参数错误", "延时秒数必须 >= 0")
                return
            start_at = datetime.fromtimestamp(time.time() + start_after_seconds)
        elif start_mode == "指定时刻":
            try:
                hour = int(self.start_hour_var.get())
                minute = int(self.start_minute_var.get())
                second = int(self.start_second_var.get())
                now = datetime.now()
                start_at = now.replace(hour=hour, minute=minute, second=second, microsecond=0)
                if start_at <= now:
                    start_at = datetime.fromtimestamp(start_at.timestamp() + 86400)
            except ValueError:
                messagebox.showerror("参数错误", "指定时刻选择无效")
                return

        self.fetch_btn.config(state=tk.DISABLED)
        self.rob_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.pause_btn.config(text="暂停抢课")
        self.rob_pause_event.clear()
        self.rob_stop_event.clear()
        self.rob_running = True

        self.rob_thread = threading.Thread(
            target=self._rob_worker,
            args=(
                selected_courses,
                delay_ms,
                click_times,
                keep_alive_interval,
                start_at,
            ),
            daemon=True,
        )
        self.rob_thread.start()

    def on_toggle_pause(self) -> None:
        if not self.rob_running:
            return
        if self.rob_pause_event.is_set():
            self.rob_pause_event.clear() 
            self.pause_btn.config(text="暂停抢课")
            self.status_var.set("已继续抢课")
        else:
            self.rob_pause_event.set()
            self.pause_btn.config(text="继续抢课")
            self.status_var.set("已暂停抢课")

    def on_stop_rob(self) -> None:
        if not self.rob_running:
            return
        self.rob_stop_event.set()
        self.rob_pause_event.clear()
        self.pause_btn.config(text="暂停抢课")
        self.status_var.set("正在终止抢课...")

    def _rob_worker(
        self,
        courses: List[Dict[str, Any]],
        delay_ms: int,
        click_times: int,
        keep_alive_interval: int,
        start_at: datetime | None,
    ) -> None:
        session = self.current_session
        base_api = self.current_base_api
        base_host = self.current_base_host
        if session is None or not base_api or not base_host:
            self.root.after(0, lambda: messagebox.showerror("状态错误", "会话已失效，请重新登录抓取后再抢课"))
            self.root.after(0, self._ui_enable_after_rob)
            return

        if not courses:
            self.root.after(0, lambda: messagebox.showerror("数据缺失", "没有可抢的课程"))
            self.root.after(0, self._ui_enable_after_rob)
            return

        try:
            last_keep_alive = 0.0

            while start_at and datetime.now() < start_at:
                if self.rob_stop_event.is_set():
                    self._set_status("抢课已终止")
                    return

                while self.rob_pause_event.is_set():
                    if self.rob_stop_event.is_set():
                        self._set_status("抢课已终止")
                        return
                    self._set_status("抢课已暂停")
                    time.sleep(0.2)

                now_ts = time.time()
                if self.keep_alive_var.get() and (now_ts - last_keep_alive >= keep_alive_interval):
                    try:
                        _heartbeat(session, base_api)
                    except Exception:
                        pass
                    last_keep_alive = now_ts

                wait_seconds = (start_at - datetime.now()).total_seconds()
                self._set_status(f"等待开始时间... 剩余 {max(0, int(wait_seconds))} 秒")
                time.sleep(0.2)

            final_resp: Dict[str, Any] = {}
            success_courses: List[str] = []
            pending_courses = list(courses)

            for i in range(1, click_times + 1):
                if self.rob_stop_event.is_set():
                    self._set_status("抢课已终止")
                    return

                while self.rob_pause_event.is_set():
                    if self.rob_stop_event.is_set():
                        self._set_status("抢课已终止")
                        return
                    self._set_status("抢课已暂停")
                    time.sleep(0.2)

                now_ts = time.time()
                if self.keep_alive_var.get() and (now_ts - last_keep_alive >= keep_alive_interval):
                    try:
                        _heartbeat(session, base_api)
                        self._set_status(f"已保活，轮次 {i}/{click_times}，待成功 {len(pending_courses)} 门...")
                    except Exception:
                        self._set_status(f"保活失败，继续轮次 {i}/{click_times}，待成功 {len(pending_courses)} 门...")
                    last_keep_alive = now_ts

                next_pending: List[Dict[str, Any]] = []
                for course in pending_courses:
                    clazz_id = str(course.get("JXBID") or "").strip()
                    secret_val = str(course.get("secretVal") or "").strip()
                    course_name = str(course.get("KCM") or course.get("KXH") or "未知课程")

                    if not clazz_id or not secret_val:
                        next_pending.append(course)
                        continue

                    resp = _select_course_once(
                        session=session,
                        base_host=base_host,
                        base_api=base_api,
                        batch_id=self.current_batch_id,
                        clazz_type=self.current_clazz_type,
                        clazz_id=clazz_id,
                        secret_val=secret_val,
                    )
                    final_resp = resp
                    code = resp.get("code")
                    msg = str(resp.get("msg") or "")
                    self._set_status(
                        f"轮次 {i}/{click_times} | {course_name}: code={code}, msg={msg}"
                    )

                    if code == 200 or ("成功" in msg):
                        success_courses.append(course_name)
                    else:
                        next_pending.append(course)

                    if self.rob_stop_event.is_set():
                        self._set_status("抢课已终止")
                        return

                pending_courses = next_pending
                if not pending_courses:
                    break

                if delay_ms > 0 and i < click_times:
                    sleep_left = delay_ms / 1000.0
                    while sleep_left > 0:
                        if self.rob_stop_event.is_set():
                            self._set_status("抢课已终止")
                            return
                        step = min(0.1, sleep_left)
                        time.sleep(step)
                        sleep_left -= step

            fail_courses = [str(c.get("KCM") or c.get("KXH") or "未知课程") for c in pending_courses]
            success_text = "、".join(success_courses) if success_courses else "无"
            fail_text = "、".join(fail_courses) if fail_courses else "无"

            if not fail_courses and success_courses:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "抢课结果",
                        f"多选抢课完成，全部成功。\n成功课程：{success_text}",
                    ),
                )
            else:
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "抢课结束",
                        (
                            f"已执行 {click_times} 轮。\n"
                            f"成功课程：{success_text}\n"
                            f"未成功课程：{fail_text}\n"
                            f"最后返回：{final_resp}"
                        ),
                    ),
                )

        except Exception as exc:
            self.root.after(0, lambda: messagebox.showerror("抢课失败", str(exc)))
        finally:
            self.root.after(0, self._ui_enable_after_rob)


def main() -> None:
    root = tk.Tk()
    app = XKGuiApp(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
