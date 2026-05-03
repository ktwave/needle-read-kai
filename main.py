import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import pyperclip
import threading
import time
import os
from PIL import Image, ImageTk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERSION = "1.0.0"

class NeedleReaderKai:
    MATCH_THRESHOLD = 0.85
    # デバッグ用: プレビュー右クリック保存。無効化するなら False にするか、下記メソッドと bind を削除。
    DEBUG_PREVIEW_RCLICK_SAVE = True

    def __init__(self, root):
        self.root = root
        self.root.title(f"Needle Reader -改- v{VERSION}")
        self.root.geometry("720x520")

        self.monitoring = False
        self.detected_values = []
        self.templates = {}
        self.last_detection_time = 0
        self.interval = 1.0  # 検知後 1 秒は次の検知を行わない

        self.load_templates()
        self.setup_ui()

    def load_templates(self):
        """resources/images/ からテンプレート画像をロードする"""
        image_dir = os.path.join(BASE_DIR, "resources", "images")
        if not os.path.exists(image_dir):
            messagebox.showerror("Error", f"Directory not found: {image_dir}")
            return

        for i in range(17):
            for suffix in ['a', 'b']:
                filename = f"{i}_{suffix}.png"
                filepath = os.path.join(image_dir, filename)
                if os.path.exists(filepath):
                    # グレースケールで読み込み
                    template = cv2.imread(filepath, cv2.IMREAD_GRAYSCALE)
                    if template is not None:
                        self.templates[f"{i}_{suffix}"] = template
                else:
                    print(f"Warning: {filename} not found.")

    def setup_ui(self):
        # 制御フレーム
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(control_frame, text="監視開始", command=self.start_monitoring)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(control_frame, text="監視停止", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.paste_gen7_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            control_frame,
            text="停止時に Gen7 の針リスト (Clock_List) へ出力",
            variable=self.paste_gen7_var,
        ).pack(side=tk.LEFT, padx=10)

        self.status_label = ttk.Label(control_frame, text="ステータス: 停止中")
        self.status_label.pack(side=tk.RIGHT, padx=5)

        # 結果表示フレーム
        results_frame = ttk.Frame(self.root, padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True)

        # 左側: テキストログ
        log_frame = ttk.LabelFrame(results_frame, text="出力結果テキスト")
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.log_text = tk.Text(log_frame, width=20, height=3, wrap=tk.NONE)
        log_xscroll = ttk.Scrollbar(log_frame, orient=tk.HORIZONTAL, command=self.log_text.xview)
        self.log_text.configure(xscrollcommand=log_xscroll.set)
        self.log_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        log_xscroll.pack(side=tk.BOTTOM, fill=tk.X)

        # 右側: 画像プレビュー
        preview_frame = ttk.LabelFrame(results_frame, text="直近10件の検知画像")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        self.preview_canvas = tk.Canvas(preview_frame, bg="gray")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_images = []  # PhotoImage 参照（表示用）
        self.preview_snapshots = []  # DEBUG: 元 64x64 PIL（保存用）

        if self.DEBUG_PREVIEW_RCLICK_SAVE:
            self.preview_canvas.bind("<Button-3>", self._on_preview_right_click)

        # 操作フレーム
        action_frame = ttk.Frame(self.root, padding="10")
        action_frame.pack(fill=tk.X)

        ttk.Button(action_frame, text="出力結果コピー", command=self.copy_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="出力結果クリア", command=self.clear_results).pack(side=tk.LEFT, padx=5)

    def start_monitoring(self):
        self.monitoring = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_label.config(text="ステータス: 監視中")
        self.clear_results()
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.monitoring = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_label.config(text="ステータス: 停止中")
        if self.paste_gen7_var.get():
            if not self.paste_to_gen7_tool():
                self.copy_results()
        else:
            self.copy_results()

    def clear_results(self):
        self.detected_values = []
        self.log_text.delete(1.0, tk.END)
        self.preview_canvas.delete("all")
        self.preview_images = []
        self.preview_snapshots = []

    # --- DEBUG: プレビュー右クリック保存（不要なら DEBUG_PREVIEW_RCLICK_SAVE=False と本ブロック削除）---
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

    # --- /DEBUG ---

    def _score_all_templates(self, frame):
        """各テンプレートの TM_CCOEFF_NORMED 最大値を計算し、(name, score) の降順リストを返す"""
        scores = []
        fh, fw = frame.shape[:2]
        for name, template in self.templates.items():
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
        """タイトルが new 3pairs で始まるウィンドウのみ（部分一致検索で負荷を抑える）"""
        candidates = gw.getWindowsWithTitle("new 3pairs")
        return [w for w in candidates if w.title and w.title.startswith("new 3pairs")]

    def _find_gen7_window(self):
        """Gen7 Main RNG Tool 向け: タイトルに Gen7 と Main RNG を含むウィンドウ"""
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
                return int(h.value)  # ctypes
            except Exception:
                return None

    def _gen7_find_clock_list_uia(self, hwnd):
        """UI Automation で Name / AutomationId が Clock_List の編集コントロールを返す"""
        from pywinauto import Desktop

        root = Desktop(backend="uia").window(handle=hwnd)
        root.wait("exists", timeout=5)
        candidates = (
            dict(auto_id="Clock_List", control_type="Edit"),
            dict(title="Clock_List", control_type="Edit"),
            dict(auto_id="Clock_List"),
            dict(title="Clock_List"),
        )
        last_err = None
        for kw in candidates:
            try:
                ctrl = root.child_window(**kw)
                ctrl.wait("exists", timeout=2)
                return ctrl.wrapper_object()
            except Exception as e:
                last_err = e
                continue
        if last_err:
            print(f"[Gen7 UIA] Clock_List が見つかりません: {last_err}")
        return None

    def _gen7_set_clock_list_text(self, win, text):
        """針リストを空にしてから出力文字列を設定（フォーカス不要）"""
        hwnd = self._hwnd_int(win)
        if not hwnd:
            print("[Gen7 UIA] HWND を取得できませんでした")
            return False
        try:
            wrap = self._gen7_find_clock_list_uia(hwnd)
            if wrap is None:
                return False
            if hasattr(wrap, "set_edit_text"):
                wrap.set_edit_text("")
                wrap.set_edit_text(text)
            elif hasattr(wrap, "set_value"):
                wrap.set_value("")
                wrap.set_value(text)
            else:
                print("[Gen7 UIA] set_edit_text / set_value が使えません")
                return False
            return True
        except Exception as e:
            print(f"[Gen7 UIA] 書き込み失敗: {e}")
            return False

    def paste_to_gen7_tool(self):
        """クリップボードへコピーし、Gen7 の Clock_List に UIA で直接反映する"""
        self.copy_results()
        text = ",".join(map(str, self.detected_values))
        win = self._find_gen7_window()
        if not win:
            return False
        return self._gen7_set_clock_list_text(win, text)

    def monitor_loop(self):
        while self.monitoring:
            # 1. ターゲットウィンドウを探す
            windows = self._find_target_windows()
            if not windows:
                time.sleep(1)
                continue

            target_window = windows[0]

            # ウィンドウが最小化されている場合はスキップ
            if target_window.isMinimized:
                time.sleep(1)
                continue

            # 2. 指定範囲をキャプチャ (159, 205) から 64x64
            # ウィンドウの左上座標を取得
            left, top = target_window.left, target_window.top
            capture_x = left + 165
            capture_y = top + 205
            
            # 画面外に出ないようにチェック（簡易的）
            try:
                screenshot = pyautogui.screenshot(region=(capture_x, capture_y, 64, 64))
                frame = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
            except Exception as e:
                print(f"[capture error] {e}")
                time.sleep(1)
                continue

            scores = self._score_all_templates(frame)
            best_match = scores[0][0] if scores else None
            best_val = scores[0][1] if scores else -1.0
            th = self.MATCH_THRESHOLD

            current_time = time.time()
            dt_since_last = current_time - self.last_detection_time

            if best_match and best_val > th:
                if dt_since_last > self.interval:
                    value = best_match.split("_")[0]
                    self.process_detection(value, screenshot)
                    self.last_detection_time = current_time
                    print(f"Detected: {value} (Score: {best_val:.3f})")

            time.sleep(0.1) # CPU負荷軽減

    def process_detection(self, value, screenshot):
        # UI更新はメインスレッドで行う
        self.root.after(0, self._update_ui, value, screenshot)

    def _update_ui(self, value, screenshot):
        self.detected_values.append(value)
        line = ",".join(map(str, self.detected_values))
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(1.0, line)
        self.log_text.see(tk.END)
        
        # プレビュー更新
        # 64x64を少しリサイズして表示
        img = screenshot.resize((48, 48), Image.Resampling.LANCZOS)
        tk_img = ImageTk.PhotoImage(img)
        
        # 参照を保持してガベージコレクションを防ぐ
        self.preview_images.insert(0, tk_img)
        if self.DEBUG_PREVIEW_RCLICK_SAVE:
            self.preview_snapshots.insert(0, screenshot.copy())
        if len(self.preview_images) > 10:
            self.preview_images.pop()
            if self.DEBUG_PREVIEW_RCLICK_SAVE:
                self.preview_snapshots.pop()

        self.render_previews()

    def render_previews(self):
        self.preview_canvas.delete("all")
        cols = 2
        for i, img in enumerate(self.preview_images):
            col = i % cols
            row = i // cols
            self.preview_canvas.create_image(
                col * 60 + 30,
                row * 60 + 30,
                image=img,
                tags=("pvimg", f"slot{i}"),
            )

if __name__ == "__main__":
    root = tk.Tk()
    app = NeedleReaderKai(root)
    root.mainloop()
