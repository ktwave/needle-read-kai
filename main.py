import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import pyperclip
import threading
import time
import os
import json
from PIL import Image, ImageTk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERSION = "1.1.0"
CONFIG_PATH = os.path.join(BASE_DIR, "app_config.json")


class NeedleReaderKai:
    MATCH_THRESHOLD = 0.85
    QR_MATCH_THRESHOLD = 0.85
    DEBUG_PREVIEW_RCLICK_SAVE = True
    PREVIEW_COLS = 5
    PREVIEW_ROWS = 2
    PREVIEW_CELL_PX = 60

    def __init__(self, root):
        self.root = root
        self.root.title(f"Needle Reader -改- v{VERSION}")
        self.root.geometry("560x310")
        self.root.minsize(560, 310)

        self.monitoring = False
        self.detected_values = []
        self.title_templates = {}
        self.qr_templates = {}
        self.last_detection_time = 0
        self.interval = 1.0
        self.mode_var = tk.StringVar(value="title")
        self.interval_var = tk.StringVar()
        self.paste_gen7_var = tk.BooleanVar()
        self.mode_props = {
            "title": {"interval": 1.0, "paste_gen7": False},
            "qr": {"interval": 0.5, "paste_gen7": False},
        }

        self.load_templates()
        self.load_config()
        self.setup_ui()
        self.apply_mode_to_ui()

    def _load_template_group(self, image_dir):
        templates = {}
        if not os.path.exists(image_dir):
            return templates
        for i in range(17):
            for suffix in ["a", "b"]:
                filename = f"{i}_{suffix}.png"
                filepath = os.path.join(image_dir, filename)
                if not os.path.exists(filepath):
                    print(f"Warning: {filename} not found in {image_dir}")
                    continue
                template = cv2.imread(filepath, cv2.IMREAD_GRAYSCALE)
                if template is not None:
                    templates[f"{i}_{suffix}"] = template
        return templates

    def load_templates(self):
        title_dir = os.path.join(BASE_DIR, "resources", "images", "title")
        qr_dir = os.path.join(BASE_DIR, "resources", "images", "qr")
        self.title_templates = self._load_template_group(title_dir)
        self.qr_templates = self._load_template_group(qr_dir)
        if not self.title_templates:
            messagebox.showwarning("Warning", f"タイトル画像が見つかりません: {title_dir}")
        if not self.qr_templates:
            messagebox.showwarning("Warning", f"QR画像が見つかりません: {qr_dir}")

    def load_config(self):
        if not os.path.exists(CONFIG_PATH):
            return
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as e:
            print(f"[config] load failed: {e}")
            return
        mode = raw.get("mode")
        if mode in ("title", "qr"):
            pass
        raw_props = raw.get("mode_props", {})
        for mode_name in ("title", "qr"):
            props = raw_props.get(mode_name, {})
            if "interval" in props:
                try:
                    iv = float(props["interval"])
                    if iv > 0:
                        self.mode_props[mode_name]["interval"] = iv
                except Exception:
                    pass
            if "paste_gen7" in props:
                self.mode_props[mode_name]["paste_gen7"] = bool(props["paste_gen7"])

    def save_config(self):
        payload = {"mode": self.mode_var.get(), "mode_props": self.mode_props}
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[config] save failed: {e}")

    def setup_ui(self):
        compact = "Compact.TButton"
        try:
            ttk.Style().configure(compact, padding=(6, 2))
        except tk.TclError:
            compact = None

        outer = ttk.Frame(self.root, padding=6)
        outer.pack(fill=tk.BOTH, expand=True)

        row1 = ttk.Frame(outer)
        row1.pack(fill=tk.X, pady=(0, 2))
        b_kw = {"style": compact} if compact else {}
        self.start_btn = ttk.Button(row1, text="監視開始", command=self.start_monitoring, **b_kw)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.stop_btn = ttk.Button(
            row1, text="監視停止", command=self.stop_monitoring, state=tk.DISABLED, **b_kw
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.status_label = ttk.Label(row1, text="ステータス: 停止中")
        self.status_label.pack(side=tk.RIGHT)

        row2 = ttk.Frame(outer)
        row2.pack(fill=tk.X, pady=(0, 4))
        self.title_link_chk = ttk.Checkbutton(
            row2,
            text="停止時にGen7 Main RNG Toolへ連携",
            variable=self.paste_gen7_var,
            command=self._on_mode_option_changed,
        )
        self.title_link_chk.pack(side=tk.LEFT)
        self.qr_output_btn = ttk.Button(row2, text="出力", command=self.output_qr_to_gen7, **b_kw)
        interval_box = ttk.Frame(row2)
        interval_box.pack(side=tk.RIGHT)
        ttk.Label(interval_box, text="検知インターバル(秒)").pack(side=tk.LEFT, padx=(0, 4))
        self.interval_entry = ttk.Entry(interval_box, width=6, textvariable=self.interval_var)
        self.interval_entry.pack(side=tk.LEFT)
        self.interval_entry.bind("<FocusOut>", lambda _e: self._apply_interval_from_ui(show_error=False))
        self.interval_entry.bind("<Return>", lambda _e: self._apply_interval_from_ui(show_error=True))

        row3 = ttk.Frame(outer)
        row3.pack(fill=tk.X, pady=(0, 4))
        mode_box = ttk.LabelFrame(row3, text="モード", padding=4)
        mode_box.pack(side=tk.LEFT)
        self.mode_title_rb = ttk.Radiobutton(
            mode_box, text="タイトル", value="title", variable=self.mode_var, command=self.on_mode_changed
        )
        self.mode_title_rb.pack(side=tk.LEFT, padx=(0, 8))
        self.mode_qr_rb = ttk.Radiobutton(
            mode_box, text="QR", value="qr", variable=self.mode_var, command=self.on_mode_changed
        )
        self.mode_qr_rb.pack(side=tk.LEFT)

        results_frame = ttk.Frame(outer)
        results_frame.pack(fill=tk.X, expand=False, pady=(0, 2))

        cw = self.PREVIEW_COLS * self.PREVIEW_CELL_PX + 8
        ch = self.PREVIEW_ROWS * self.PREVIEW_CELL_PX + 8
        self.preview_frame = ttk.LabelFrame(results_frame, text="直近10件の検知画像", padding=4)
        self.preview_frame.pack(side=tk.RIGHT, padx=(6, 0), anchor=tk.N)

        self.preview_canvas = tk.Canvas(
            self.preview_frame,
            width=cw,
            height=ch,
            bg="gray",
            highlightthickness=0,
        )
        self.preview_canvas.pack()
        self.preview_images = []
        self.preview_snapshots = []

        if self.DEBUG_PREVIEW_RCLICK_SAVE:
            self.preview_canvas.bind("<Button-3>", self._on_preview_right_click)

        self.log_font = tkfont.Font(root=self.root, family="Segoe UI", size=9)
        self.log_frame = ttk.LabelFrame(results_frame, text="出力結果テキスト", padding=4)
        self.log_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 0))

        self.log_text = tk.Text(self.log_frame, height=3, wrap=tk.NONE, font=self.log_font)
        log_xscroll = ttk.Scrollbar(self.log_frame, orient=tk.HORIZONTAL, command=self.log_text.xview)
        self.log_text.configure(xscrollcommand=log_xscroll.set)
        self.log_text.pack(side=tk.TOP, fill=tk.X, expand=False)
        log_xscroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.root.after_idle(self._sync_log_height_to_preview)

        action_frame = ttk.Frame(outer)
        action_frame.pack(fill=tk.X, pady=(2, 0))
        ttk.Button(action_frame, text="出力結果コピー", command=self.copy_results, **b_kw).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(action_frame, text="出力結果クリア", command=self.clear_results, **b_kw).pack(
            side=tk.LEFT
        )
        self._apply_mode_visibility()

    def get_active_props(self):
        return self.mode_props[self.mode_var.get()]

    def apply_mode_to_ui(self):
        props = self.get_active_props()
        self.interval = float(props["interval"])
        self.interval_var.set(f"{self.interval:g}")
        self.paste_gen7_var.set(bool(props["paste_gen7"]))

    def _persist_active_props_from_ui(self):
        props = self.get_active_props()
        props["interval"] = float(self.interval)
        props["paste_gen7"] = bool(self.paste_gen7_var.get())
        self.save_config()

    def on_mode_changed(self):
        self.apply_mode_to_ui()
        self._apply_mode_visibility()
        self.save_config()

    def _on_mode_option_changed(self):
        self._persist_active_props_from_ui()

    def _apply_mode_visibility(self):
        mode = self.mode_var.get()
        if mode == "title":
            if not self.title_link_chk.winfo_ismapped():
                self.title_link_chk.pack(side=tk.LEFT)
            if self.qr_output_btn.winfo_ismapped():
                self.qr_output_btn.pack_forget()
        else:
            if self.title_link_chk.winfo_ismapped():
                self.title_link_chk.pack_forget()
            if not self.qr_output_btn.winfo_ismapped():
                self.qr_output_btn.pack(side=tk.LEFT)

    def _sync_log_height_to_preview(self):
        line_px = max(10, self.log_font.metrics("linespace"))
        try:
            for _ in range(5):
                self.root.update_idletasks()
                target = self.preview_frame.winfo_reqheight()
                chrome = self.log_frame.winfo_reqheight() - self.log_text.winfo_reqheight()
                inner = max(line_px, target - chrome)
                lines = max(2, round(inner / line_px))
                self.log_text.configure(height=lines)
                self.root.update_idletasks()
                if abs(self.log_frame.winfo_reqheight() - target) <= 3:
                    break
        except tk.TclError:
            pass

    def start_monitoring(self):
        if not self._apply_interval_from_ui(show_error=True):
            return
        self.monitoring = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.mode_title_rb.config(state=tk.DISABLED)
        self.mode_qr_rb.config(state=tk.DISABLED)
        self.status_label.config(text="ステータス: 監視中")
        self.clear_results()
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()

    def _apply_interval_from_ui(self, show_error):
        raw = self.interval_var.get().strip()
        try:
            value = float(raw)
            if value <= 0:
                raise ValueError
        except ValueError:
            fallback = self.get_active_props()["interval"]
            self.interval = float(fallback)
            self.interval_var.set(f"{self.interval:g}")
            if show_error:
                messagebox.showwarning(
                    "入力エラー",
                    "インターバルは 0 より大きい数値で入力してください。前の値に戻しました。",
                )
            return False
        self.interval = value
        self.interval_var.set(f"{value:g}")
        self._persist_active_props_from_ui()
        return True

    def stop_monitoring(self):
        self.monitoring = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.mode_title_rb.config(state=tk.NORMAL)
        self.mode_qr_rb.config(state=tk.NORMAL)
        self.status_label.config(text="ステータス: 停止中")
        if self.mode_var.get() == "title":
            if self.paste_gen7_var.get():
                if not self.paste_to_gen7_tool():
                    self.copy_results()
            else:
                self.copy_results()
        self._persist_active_props_from_ui()

    def clear_results(self):
        self.detected_values = []
        self.log_text.delete(1.0, tk.END)
        self.preview_canvas.delete("all")
        self.preview_images = []
        self.preview_snapshots = []

    def _on_preview_right_click(self, event):
        if not self.preview_snapshots:
            return
        cid = self.preview_canvas.find_closest(event.x, event.y)
        item_id = cid[0] if isinstance(cid, (list, tuple)) else cid
        if not item_id:
            return
        tags = self.preview_canvas.gettags(item_id)
        idx = None
        for t in tags:
            if t.startswith("slot") and t[4:].isdigit():
                idx = int(t[4:])
                break
        if idx is None or idx < 0 or idx >= len(self.preview_snapshots):
            return

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(
            label="このキャプチャを保存 (自動名・_debug_preview_saves/)",
            command=lambda i=idx: self._debug_save_preview_auto(i),
        )
        menu.add_command(
            label="名前を付けて保存…",
            command=lambda i=idx: self._debug_save_preview_dialog(i),
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _debug_save_preview_auto(self, index):
        im = self.preview_snapshots[index]
        out_dir = os.path.join(BASE_DIR, "_debug_preview_saves")
        os.makedirs(out_dir, exist_ok=True)
        name = time.strftime("capture_%Y%m%d_%H%M%S") + f"_slot{index}.png"
        path = os.path.join(out_dir, name)
        im.copy().save(path, "PNG")
        print(f"[preview save] {path}")

    def _debug_save_preview_dialog(self, index):
        im = self.preview_snapshots[index]
        initial = time.strftime("capture_%Y%m%d_%H%M%S") + f"_slot{index}.png"
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="キャプチャを保存",
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("すべて", "*.*")],
            initialfile=initial,
        )
        if path:
            im.copy().save(path, "PNG")
            print(f"[preview save] {path}")

    def _score_all_templates(self, frame, templates):
        scores = []
        fh, fw = frame.shape[:2]
        for name, template in templates.items():
            th, tw = template.shape[:2]
            if th > fh or tw > fw:
                continue
            try:
                res = cv2.matchTemplate(frame, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, _ = cv2.minMaxLoc(res)
            except cv2.error:
                continue
            scores.append((name, float(max_val)))
        scores.sort(key=lambda x: x[1], reverse=True)
        return scores

    def copy_results(self):
        text = ",".join(map(str, self.detected_values))
        pyperclip.copy(text)

    def _find_target_windows(self):
        candidates = gw.getWindowsWithTitle("new 3pairs")
        return [w for w in candidates if w.title and w.title.startswith("new 3pairs")]

    def _find_gen7_window(self):
        for w in gw.getWindowsWithTitle("Gen7"):
            if not w.title:
                continue
            if "Main RNG" in w.title and "Gen7" in w.title:
                return w
        for w in gw.getWindowsWithTitle("Main RNG"):
            if w.title and "Gen7" in w.title:
                return w
        return None

    @staticmethod
    def _hwnd_int(win):
        h = getattr(win, "_hWnd", None)
        if h is None:
            return None
        try:
            return int(h)
        except (TypeError, ValueError):
            try:
                return int(h.value)
            except Exception:
                return None

    @staticmethod
    def _gen7_uia_root(hwnd):
        from pywinauto import Desktop

        root = Desktop(backend="uia").window(handle=hwnd)
        root.wait("exists", timeout=5)
        return root

    @staticmethod
    def _gen7_find_uia(root, name, control_type=None):
        variants = []
        if control_type:
            variants.append({"auto_id": name, "control_type": control_type})
            variants.append({"title": name, "control_type": control_type})
        variants.append({"auto_id": name})
        variants.append({"title": name})
        last_err = None
        for kw in variants:
            try:
                spec = root.child_window(**kw)
                spec.wait("exists", timeout=2)
                return spec.wrapper_object()
            except Exception as e:
                last_err = e
                continue
        print(f"[Gen7 UIA] コントロール '{name}' が見つかりません: {last_err}")
        return None

    def _gen7_stop_sequence(self, win, text):
        hwnd = self._hwnd_int(win)
        if not hwnd:
            print("[Gen7 UIA] HWND を取得できませんでした")
            return False
        try:
            root = self._gen7_uia_root(hwnd)

            rb_save = self._gen7_find_uia(root, "RB_SaveScreen", "RadioButton")
            if rb_save is None:
                return False
            rb_save.click()
            time.sleep(0.05)

            rb_input = self._gen7_find_uia(root, "StartClockInput", "RadioButton")
            if rb_input is None:
                return False
            rb_input.click()
            time.sleep(0.05)

            clock = self._gen7_find_uia(root, "Clock_List", "Edit")
            if clock is None:
                return False
            if hasattr(clock, "set_edit_text"):
                clock.set_edit_text("")
                clock.set_edit_text(text)
            elif hasattr(clock, "set_value"):
                clock.set_value("")
                clock.set_value(text)
            else:
                print("[Gen7 UIA] Clock_List に set_edit_text / set_value がありません")
                return False
            time.sleep(0.05)

            btn = self._gen7_find_uia(root, "B_Search", "Button")
            if btn is None:
                return False
            btn.click()
            return True
        except Exception as e:
            print(f"[Gen7 UIA] 連携失敗: {e}")
            return False

    def paste_to_gen7_tool(self):
        self.copy_results()
        text = ",".join(map(str, self.detected_values))
        win = self._find_gen7_window()
        if not win:
            return False
        return self._gen7_stop_sequence(win, text)

    def _gen7_output_sequence(self, win, text):
        hwnd = self._hwnd_int(win)
        if not hwnd:
            print("[Gen7 UIA] HWND を取得できませんでした")
            return False
        try:
            root = self._gen7_uia_root(hwnd)
            clock = self._gen7_find_uia(root, "Clock_List", "Edit")
            if clock is None:
                return False
            if hasattr(clock, "set_edit_text"):
                clock.set_edit_text("")
                clock.set_edit_text(text)
            elif hasattr(clock, "set_value"):
                clock.set_value("")
                clock.set_value(text)
            else:
                print("[Gen7 UIA] Clock_List に set_edit_text / set_value がありません")
                return False
            btn = self._gen7_find_uia(root, "B_Search", "Button")
            if btn is None:
                return False
            btn.click()
            return True
        except Exception as e:
            print(f"[Gen7 UIA] 出力失敗: {e}")
            return False

    def output_qr_to_gen7(self):
        self.copy_results()
        text = ",".join(map(str, self.detected_values))
        win = self._find_gen7_window()
        if not win:
            messagebox.showwarning("出力失敗", "Gen7 Main RNG Tool が見つかりません。")
            return
        if not self._gen7_output_sequence(win, text):
            messagebox.showwarning("出力失敗", "針リストへの出力に失敗しました。")

    def monitor_loop(self):
        while self.monitoring:
            windows = self._find_target_windows()
            if not windows:
                time.sleep(1)
                continue

            target_window = windows[0]

            if target_window.isMinimized:
                time.sleep(1)
                continue

            left, top = target_window.left, target_window.top
            title_x = left + 165
            title_y = top + 205
            mode = self.mode_var.get()

            try:
                if mode == "qr":
                    qr_x = title_x + 521
                    qr_y = title_y + 380
                    screenshot = pyautogui.screenshot(region=(qr_x, qr_y, 28, 28))
                else:
                    screenshot = pyautogui.screenshot(region=(title_x, title_y, 64, 64))
                frame = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
            except Exception as e:
                print(f"[capture error] {e}")
                time.sleep(1)
                continue

            current_time = time.time()
            dt_since_last = current_time - self.last_detection_time

            if mode == "qr":
                if dt_since_last > self.interval:
                    scores = self._score_all_templates(frame, self.qr_templates)
                    best_match = scores[0][0] if scores else None
                    best_val = scores[0][1] if scores else -1.0
                    if best_match and best_val > self.QR_MATCH_THRESHOLD:
                        value = best_match.split("_")[0]
                        self.process_detection(value, screenshot)
                        print(f"Detected: {value} (Score: {best_val:.3f})")
                    self.last_detection_time = current_time
            else:
                scores = self._score_all_templates(frame, self.title_templates)
                best_match = scores[0][0] if scores else None
                best_val = scores[0][1] if scores else -1.0
                th = self.MATCH_THRESHOLD
                if best_match and best_val > th and dt_since_last > self.interval:
                    value = best_match.split("_")[0]
                    self.process_detection(value, screenshot)
                    self.last_detection_time = current_time
                    print(f"Detected: {value} (Score: {best_val:.3f})")

            time.sleep(0.05)

    def process_detection(self, value, screenshot):
        self.root.after(0, self._update_ui, value, screenshot)

    def process_preview_only(self, screenshot):
        self.root.after(0, self._update_preview_only, screenshot)

    def _update_ui(self, value, screenshot):
        self.detected_values.append(value)
        line = ",".join(map(str, self.detected_values))
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(1.0, line)
        self.log_text.see(tk.END)
        self._append_preview(screenshot)

    def _update_preview_only(self, screenshot):
        self._append_preview(screenshot)

    def _append_preview(self, screenshot):
        img = screenshot.resize((48, 48), Image.Resampling.LANCZOS)
        tk_img = ImageTk.PhotoImage(img)

        self.preview_images.append(tk_img)
        if self.DEBUG_PREVIEW_RCLICK_SAVE:
            self.preview_snapshots.append(screenshot.copy())
        if len(self.preview_images) > 10:
            self.preview_images.pop(0)
            if self.DEBUG_PREVIEW_RCLICK_SAVE:
                self.preview_snapshots.pop(0)

        self.render_previews()

    def render_previews(self):
        self.preview_canvas.delete("all")
        cols = self.PREVIEW_COLS
        cell = self.PREVIEW_CELL_PX
        half = cell // 2
        for i, img in enumerate(self.preview_images):
            col = i % cols
            row = i // cols
            self.preview_canvas.create_image(
                col * cell + half,
                row * cell + half,
                image=img,
                tags=("pvimg", f"slot{i}"),
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = NeedleReaderKai(root)
    root.mainloop()
