import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pyautogui
from PIL import Image, ImageTk, ImageEnhance, ImageFilter
import time
import threading
from pynput import keyboard
import io
import os
import platform

# macOSでの動作改善のためpyautoguiの設定を調整
pyautogui.FAILSAFE = False  # フェイルセーフを無効化（マウスを左上角に移動してもエラーにならない）
pyautogui.PAUSE = 0.1  # 操作間隔を短縮

class ImageEditorWindow(tk.Toplevel):
    def __init__(self, master, images, app_instance):
        super().__init__(master)
        self.app = app_instance
        self.original_images = images 
        self.transient(master)
        self.title("画像編集")
        self.geometry("800x700")
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        # Data
        self.image_items = [] # Holds dicts of {"thumb": PhotoImage, "original": PIL.Image, "label": widget}
        self.current_selection_index = 0
        
        # --- Layout ---
        # Section 3: Buttons (packed to the bottom first)
        button_frame = ttk.Frame(self)
        button_frame.pack(side="bottom", fill="x", pady=5, padx=10)
        ttk.Button(button_frame, text="選択した画像を削除", command=self.delete_selected).pack(side="left", padx=5)
        ttk.Button(button_frame, text="PDFとして保存", command=self.save_to_pdf).pack(side="right", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.cancel).pack(side="right")

        # Section 2: Thumbnails
        thumb_container = ttk.Frame(self)
        thumb_container.pack(side="bottom", fill="x", pady=5)
        
        self.thumb_canvas = tk.Canvas(thumb_container)
        x_scrollbar = ttk.Scrollbar(thumb_container, orient="horizontal", command=self.thumb_canvas.xview)
        self.thumb_canvas.configure(xscrollcommand=x_scrollbar.set)
        
        x_scrollbar.pack(side="bottom", fill="x")
        self.thumb_canvas.pack(side="top", fill="x", expand=True)
        
        self.thumb_frame = ttk.Frame(self.thumb_canvas)
        self.thumb_canvas.create_window((0, 0), window=self.thumb_frame, anchor="nw")
        self.thumb_frame.bind("<Configure>", lambda e: self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all")))

        # Section 1: Preview
        self.preview_frame = ttk.Frame(self)
        self.preview_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        self.preview_label = ttk.Label(self.preview_frame, anchor="center")
        self.preview_label.pack(fill="both", expand=True)

        # --- Bindings ---
        self.bind("<Left>", self.select_previous)
        self.bind("<Right>", self.select_next)
        self.bind("<MouseWheel>", self.on_mouse_wheel)

        # --- Load Images ---
        self.load_images()
        if self.image_items:
            self.select_image(0)

    def load_images(self):
        for widget in self.thumb_frame.winfo_children():
            widget.destroy()
        self.image_items = []

        for i, img in enumerate(self.original_images):
            thumb_img = img.copy()
            thumb_img.thumbnail((120, 120))
            thumb_photo = ImageTk.PhotoImage(thumb_img)
            
            label = ttk.Label(self.thumb_frame, image=thumb_photo, text=f" {i+1} ", compound="bottom", cursor="hand2")
            label.image = thumb_photo # Keep a reference
            label.pack(side="left", padx=5, pady=5)
            
            item_data = {"thumb": thumb_photo, "original": img, "label": label}
            self.image_items.append(item_data)
            
            label.bind("<Button-1>", lambda e, index=i: self.select_image(index))

    def select_image(self, index):
        if not (0 <= index < len(self.image_items)):
            return

        self.current_selection_index = index

        # Update preview
        original_img = self.image_items[index]["original"]
        
        self.update_idletasks()
        frame_w = self.preview_frame.winfo_width()
        frame_h = self.preview_frame.winfo_height()
        img_w, img_h = original_img.size
        
        if frame_w <= 1 or frame_h <= 1:
            self.after(50, lambda: self.select_image(index))
            return

        scale = min(frame_w / img_w, frame_h / img_h)
        new_w, new_h = int(img_w * scale), int(img_h * scale)
        
        # 高品質リサイズアルゴリズムを使用
        resized_img = original_img.resize((new_w, new_h), Image.LANCZOS)
        preview_photo = ImageTk.PhotoImage(resized_img)
        
        self.preview_label.config(image=preview_photo)
        self.preview_label.image = preview_photo

        # Update thumbnail selection highlight
        for i, item in enumerate(self.image_items):
            if i == index:
                item["label"].config(relief="solid", borderwidth=2)
            else:
                item["label"].config(relief="flat")

        # Auto-scroll logic
        self.thumb_canvas.update_idletasks()
        selected_label = self.image_items[index]["label"]
        
        canvas_width = self.thumb_canvas.winfo_width()
        scroll_region = self.thumb_canvas.bbox("all")
        if not scroll_region: return
        
        total_width = scroll_region[2]
        if total_width <= canvas_width: return # No need to scroll

        label_x = selected_label.winfo_x()
        label_width = selected_label.winfo_width()
        
        start_frac = label_x / total_width
        end_frac = (label_x + label_width) / total_width
        
        view_start = self.thumb_canvas.xview()[0]
        view_end = self.thumb_canvas.xview()[1]

        if start_frac < view_start:
            self.thumb_canvas.xview_moveto(start_frac)
        elif end_frac > view_end:
            self.thumb_canvas.xview_moveto(end_frac - (view_end - view_start))

    def select_next(self, event=None):
        self.select_image(self.current_selection_index + 1)

    def select_previous(self, event=None):
        self.select_image(self.current_selection_index - 1)

    def on_mouse_wheel(self, event):
        if event.delta > 0:
            self.select_previous()
        else:
            self.select_next()

    def delete_selected(self):
        if not self.image_items or not (0 <= self.current_selection_index < len(self.image_items)):
            return
        
        self.original_images.pop(self.current_selection_index)
        
        if self.current_selection_index >= len(self.original_images):
            self.current_selection_index = len(self.original_images) - 1

        self.load_images()
        
        if not self.image_items:
            self.preview_label.config(image=None, text="No images")
        else:
            self.select_image(self.current_selection_index)

    def save_to_pdf(self):
        if not self.original_images:
            messagebox.showwarning("No Images", "There are no images to save.")
            return
        self.app.save_pdf(self.original_images)
        self.destroy()

    def cancel(self):
        self.destroy()
        self.app.status_label.config(text="PDF化をキャンセルしました。")


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("電子書籍スクリーンショット")
        self.master.geometry("400x550")
        self.pack(fill=tk.BOTH, expand=True)
        self.screenshot_area = None
        self.click_position = None
        self.is_running = False
        self.thread = None
        self.images = []
        self.create_widgets()
        self.listener = keyboard.Listener(on_press=self.on_key_press)
        self.listener.start()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_key_press(self, key):
        if key == keyboard.Key.esc and self.is_running:
            self.master.after(0, self.emergency_stop)

    def on_closing(self):
        self.listener.stop()
        self.master.destroy()

    def emergency_stop(self):
        if self.is_running:
            self.is_running = False
            messagebox.showinfo("緊急停止", "Escapeキーが押されました。処理を停止します。")

    def create_widgets(self):
        setting_frame = ttk.LabelFrame(self, text="設定")
        setting_frame.pack(pady=10, padx=10, fill="x")
        self.pdf_path = tk.StringVar()
        pdf_path_frame = ttk.Frame(setting_frame)
        pdf_path_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(pdf_path_frame, text="PDF出力パス:").pack(side=tk.LEFT)
        ttk.Entry(pdf_path_frame, textvariable=self.pdf_path).pack(side=tk.LEFT, expand=True, fill="x")
        ttk.Button(pdf_path_frame, text="選択", command=self.select_pdf_path).pack(side=tk.LEFT)
        wait_time_frame = ttk.Frame(setting_frame)
        wait_time_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(wait_time_frame, text="待機時間(秒):").pack(side=tk.LEFT)
        self.wait_time = tk.DoubleVar(value=0.5)
        ttk.Spinbox(wait_time_frame, from_=0.1, to=10.0, increment=0.1, textvariable=self.wait_time, width=5).pack(side=tk.LEFT)
        count_frame = ttk.Frame(setting_frame)
        count_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(count_frame, text="撮影枚数:").pack(side=tk.LEFT)
        self.screenshot_count = tk.IntVar(value=10)
        ttk.Spinbox(count_frame, from_=1, to=1000, textvariable=self.screenshot_count, width=5).pack(side=tk.LEFT)
        
        # 画質設定フレーム追加（最適化版）
        quality_frame = ttk.Frame(setting_frame)
        quality_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(quality_frame, text="PDF品質設定:").pack(side=tk.LEFT)
        self.quality_scale = tk.DoubleVar(value=1.5)
        quality_spinbox = ttk.Spinbox(quality_frame, from_=1.0, to=3.0, increment=0.5, 
                                     textvariable=self.quality_scale, width=5)
        quality_spinbox.pack(side=tk.LEFT, padx=(5, 10))
        ttk.Label(quality_frame, text="倍 (1.0=軽量, 1.5=推奨, 3.0=高品質)", 
                 font=('Segoe UI', 8)).pack(side=tk.LEFT)
        
        action_frame = ttk.LabelFrame(self, text="操作")
        action_frame.pack(pady=10, padx=10, fill="x")
        ttk.Button(action_frame, text="範囲選択", command=self.select_area).pack(fill="x", pady=5)
        ttk.Button(action_frame, text="クリック位置設定", command=self.set_click_position).pack(fill="x", pady=5)
        self.start_button = ttk.Button(action_frame, text="開始", command=self.start)
        self.start_button.pack(fill="x", pady=5)
        self.stop_button = ttk.Button(action_frame, text="停止", command=self.stop, state=tk.DISABLED)
        self.stop_button.pack(fill="x", pady=5)
        status_frame = ttk.LabelFrame(self, text="状態")
        status_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.status_label = ttk.Label(status_frame, text="待機中")
        self.status_label.pack(pady=5)
        self.page_count_label = ttk.Label(status_frame, text="撮影枚数: 0")
        self.page_count_label.pack(pady=5)
        self.settings_display = tk.Text(status_frame, height=8, width=45)
        self.settings_display.pack(pady=5, padx=5)
        self.settings_display.config(state=tk.DISABLED)
        self.update_settings_display()

    def select_pdf_path(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path.set(path)
            self.update_settings_display()

    def select_area(self):
        self.master.withdraw()
        self.area_selection_window = tk.Toplevel(self.master)
        
        # プラットフォーム固有のフルスクリーン設定
        if platform.system() == "Darwin":  # macOS
            # macOSでは overrideredirect と geometry を使用
            self.area_selection_window.overrideredirect(True)
            screen_width = self.area_selection_window.winfo_screenwidth()
            screen_height = self.area_selection_window.winfo_screenheight()
            self.area_selection_window.geometry(f"{screen_width}x{screen_height}+0+0")
            self.area_selection_window.attributes("-alpha", 0.3)
            self.area_selection_window.attributes("-topmost", True)
            self.area_selection_window.lift()
        else:  # Windows
            self.area_selection_window.attributes("-fullscreen", True)
            self.area_selection_window.attributes("-alpha", 0.3)
        
        self.area_selection_window.bind("<ButtonPress-1>", self.on_area_select_start)
        self.area_selection_window.bind("<B1-Motion>", self.on_area_select_drag)
        self.area_selection_window.bind("<ButtonRelease-1>", self.on_area_select_end)
        
        # Escapeキーでキャンセル
        self.area_selection_window.bind("<Escape>", self.cancel_area_selection)
        self.area_selection_window.focus_set()
        
        self.area_canvas = tk.Canvas(self.area_selection_window, cursor="crosshair", highlightthickness=0)
        self.area_canvas.pack(fill=tk.BOTH, expand=True)

    def cancel_area_selection(self, event=None):
        """範囲選択をキャンセル"""
        if hasattr(self, 'area_selection_window') and self.area_selection_window.winfo_exists():
            self.area_selection_window.destroy()
        self.master.deiconify()


    def on_area_select_start(self, event):
        self.area_start_x = event.x
        self.area_start_y = event.y
        self.area_rect = None

    def on_area_select_drag(self, event):
        if self.area_rect:
            self.area_canvas.delete(self.area_rect)
        self.area_rect = self.area_canvas.create_rectangle(self.area_start_x, self.area_start_y, event.x, event.y, outline='red', width=2)

    def on_area_select_end(self, event):
        self.area_end_x = event.x
        self.area_end_y = event.y
        self.area_selection_window.destroy()
        x1 = min(self.area_start_x, self.area_end_x)
        y1 = min(self.area_start_y, self.area_end_y)
        x2 = max(self.area_start_x, self.area_end_x)
        y2 = max(self.area_start_y, self.area_end_y)
        # 最高解像度でスクリーンショットを取得
        screenshot = self.capture_high_quality_screenshot(x1, y1, x2 - x1, y2 - y1)
        self.show_preview(screenshot, (x1, y1, x2, y2))

    def show_preview(self, screenshot, rect):
        preview_window = tk.Toplevel(self.master)
        preview_window.title("プレビュー")
        preview_window.resizable(False, False)
        MAX_WIDTH, MAX_HEIGHT = 800, 600
        preview_image = screenshot.copy()
        preview_image.thumbnail((MAX_WIDTH, MAX_HEIGHT), Image.LANCZOS)
        photo = ImageTk.PhotoImage(preview_image)
        
        canvas = tk.Canvas(preview_window, width=preview_image.width, height=preview_image.height)
        canvas.pack(padx=10, pady=10)
        canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        canvas.image = photo
        
        button_frame = ttk.Frame(preview_window)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        def confirm():
            self.screenshot_area = rect
            self.master.deiconify()
            self.update_settings_display()
            preview_window.destroy()
        
        def retry():
            preview_window.destroy()
            self.select_area()
            
        def cancel():
            self.master.deiconify()
            preview_window.destroy()
            
        ttk.Button(button_frame, text="確定", command=confirm).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="やり直す", command=retry).pack(side=tk.RIGHT)
        preview_window.protocol("WM_DELETE_WINDOW", cancel)
        preview_window.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() - preview_window.winfo_width()) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - preview_window.winfo_height()) // 2
        preview_window.geometry(f"+{x}+{y}")

    def set_click_position(self):
        self.master.withdraw()
        self.click_position_window = tk.Toplevel(self.master)
        
        # プラットフォーム固有のフルスクリーン設定
        if platform.system() == "Darwin":  # macOS
            # macOSでは overrideredirect と geometry を使用
            self.click_position_window.overrideredirect(True)
            screen_width = self.click_position_window.winfo_screenwidth()
            screen_height = self.click_position_window.winfo_screenheight()
            self.click_position_window.geometry(f"{screen_width}x{screen_height}+0+0")
            self.click_position_window.attributes("-alpha", 0.3)
            self.click_position_window.attributes("-topmost", True)
            self.click_position_window.lift()
        else:  # Windows
            self.click_position_window.attributes("-fullscreen", True)
            self.click_position_window.attributes("-alpha", 0.3)
        
        self.click_position_window.bind("<ButtonPress-1>", self.on_click_position_set)
        
        # Escapeキーでキャンセル
        self.click_position_window.bind("<Escape>", self.cancel_click_position)
        self.click_position_window.focus_set()
        
        self.click_position_canvas = tk.Canvas(self.click_position_window, cursor="crosshair", highlightthickness=0)
        self.click_position_canvas.pack(fill=tk.BOTH, expand=True)

    def cancel_click_position(self, event=None):
        """クリック位置設定をキャンセル"""
        if hasattr(self, 'click_position_window') and self.click_position_window.winfo_exists():
            self.click_position_window.destroy()
        self.master.deiconify()

    def on_click_position_set(self, event):
        self.click_position = (event.x, event.y)
        self.click_position_window.destroy()
        self.master.deiconify()
        self.update_settings_display()

    def start(self):
        if not self.pdf_path.get():
            messagebox.showerror("エラー", "PDF出力パスが設定されていません。")
            return
        if not self.screenshot_area:
            messagebox.showerror("エラー", "スクリーンショット範囲が設定されていません。")
            return
        if not self.click_position:
            messagebox.showerror("エラー", "クリック位置が設定されていません。")
            return
        if self.screenshot_count.get() <= 0:
            messagebox.showerror("エラー", "撮影枚数は1以上に設定してください。")
            return

        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_label.config(text="処理中...")
        self.page_count_label.config(text=f"撮影枚数: 0/{self.screenshot_count.get()}")
        self.images = []
        self.thread = threading.Thread(target=self.automation_thread)
        self.thread.start()

    def stop(self):
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="停止処理中...")
        
    def open_editor(self):
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="撮影完了。画像編集中...")
        editor = ImageEditorWindow(self.master, self.images, self)
        editor.grab_set()

    def save_pdf(self, images_to_save):
        if self.thread and self.thread.is_alive():
            self.thread.join()

        if images_to_save:
            try:
                print("従来の方法でPDF保存を実行...")
                
                # RGB変換（PDF用）- 3倍拡大して高解像度化
                converted_images = []
                for i, img in enumerate(images_to_save):
                    # まずRGBに変換
                    if img.mode != 'RGB':
                        rgb_img = img.convert('RGB')
                    else:
                        rgb_img = img
                    
                    # 画像を設定倍率で拡大（高品質リサイズ）
                    scale_factor = self.quality_scale.get()
                    original_size = rgb_img.size
                    enlarged_size = (int(original_size[0] * scale_factor), int(original_size[1] * scale_factor))
                    enlarged_img = rgb_img.resize(enlarged_size, Image.LANCZOS)
                    
                    # 読書用最適DPI情報を設定（100-150 DPI）
                    target_dpi = 100 if scale_factor <= 1.5 else 150
                    enlarged_img.info['dpi'] = (target_dpi, target_dpi)
                    converted_images.append(enlarged_img)
                    print(f"変換後画像 {i+1}: 元サイズ={original_size} -> {scale_factor}倍拡大 -> 拡大後サイズ={enlarged_img.size} (DPI: {target_dpi})")
                
                # 読書用最適化PDF保存（軽量で鮮明）
                converted_images[0].save(
                    self.pdf_path.get(),
                    save_all=True,
                    append_images=converted_images[1:] if len(converted_images) > 1 else [],
                    resolution=150.0,   # 読書用最適解像度
                    quality=85,         # 品質と容量のバランス
                    optimize=True,      # ファイルサイズ最適化
                    compress_level=6    # 適度な圧縮
                )
                
                self.status_label.config(text=f"読書用最適化PDF保存完了: {self.pdf_path.get()}")
                
            except Exception as e:
                print(f"PDF保存エラー詳細: {e}")
                messagebox.showerror("PDF保存エラー", f"エラー詳細:\n{str(e)}")
                self.status_label.config(text="PDF保存エラー")
        else:
            self.status_label.config(text="保存する画像がありません")

    def capture_high_quality_screenshot(self, x, y, width, height):
        """最高画質でスクリーンショットを取得"""
        try:
            if platform.system() == "Windows":
                # Windows の場合、DPI対応の高解像度スクリーンショットを試行
                try:
                    import win32gui
                    import win32ui
                    import win32con
                    from PIL import Image
                    
                    # デスクトップのデバイスコンテキストを取得
                    hdesktop = win32gui.GetDesktopWindow()
                    desktop_dc = win32gui.GetWindowDC(hdesktop)
                    img_dc = win32ui.CreateDCFromHandle(desktop_dc)
                    mem_dc = img_dc.CreateCompatibleDC()
                    
                    # ビットマップを作成
                    screenshot = win32ui.CreateBitmap()
                    screenshot.CreateCompatibleBitmap(img_dc, width, height)
                    mem_dc.SelectObject(screenshot)
                    
                    # 高品質でコピー（HALFTONE モードで品質向上）
                    mem_dc.SetStretchBltMode(win32con.HALFTONE)
                    mem_dc.BitBlt((0, 0), (width, height), img_dc, (x, y), win32con.SRCCOPY)
                    
                    # ビットマップデータを取得
                    bmpinfo = screenshot.GetInfo()
                    bmpstr = screenshot.GetBitmapBits(True)
                    
                    # PILイメージに変換
                    img = Image.frombuffer(
                        'RGB',
                        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                        bmpstr, 'raw', 'BGRX', 0, 1
                    )
                    
                    # リソースを解放
                    mem_dc.DeleteDC()
                    img_dc.DeleteDC()
                    win32gui.ReleaseDC(hdesktop, desktop_dc)
                    win32gui.DeleteObject(screenshot.GetHandle())
                    
                    print(f"Windows API高解像度スクリーンショット取得: {img.size}")
                    return img
                    
                except ImportError:
                    print("pywin32が利用できません。pyautoguiを使用します。")
                    pass
                except Exception as e:
                    print(f"Windows API スクリーンショット取得エラー: {e}")
                    pass
            
            elif platform.system() == "Darwin":  # macOS
                # macOS の場合、Quartz APIを使用した高解像度スクリーンショット
                try:
                    import Quartz
                    from Cocoa import NSBitmapImageRep
                    from PIL import Image
                    
                    # CGDisplayCreateImageForRect を使用してRetinaスケールに対応
                    region = Quartz.CGRectMake(x, y, width, height)
                    image_ref = Quartz.CGWindowListCreateImage(
                        region,
                        Quartz.kCGWindowListOptionOnScreenOnly,
                        Quartz.kCGNullWindowID,
                        Quartz.kCGWindowImageDefault
                    )
                    
                    if image_ref:
                        # CGImageをNSBitmapImageRepに変換
                        bitmap_rep = NSBitmapImageRep.alloc().initWithCGImage_(image_ref)
                        
                        # PNG データを取得
                        png_data = bitmap_rep.representationUsingType_properties_(
                            1,  # NSBitmapImageFileTypePNG
                            None
                        )
                        
                        # PIL Image に変換
                        img = Image.open(io.BytesIO(png_data.bytes()))
                        
                        print(f"macOS Quartz API高解像度スクリーンショット取得: {img.size}")
                        return img
                        
                except ImportError:
                    print("PyObjCが利用できません。pyautoguiを使用します。")
                    pass
                except Exception as e:
                    print(f"macOS Quartz API スクリーンショット取得エラー: {e}")
                    pass
        except:
            pass
        
        # フォールバック: pyautoguiを使用（最高品質設定）
        print("pyautoguiでスクリーンショット取得")
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        
        # スクリーンショット後の品質向上処理
        # 1. シャープネス強化（文字をより鮮明に）
        
        # 読書用画質調整（自然で読みやすい品質）
        contrast_enhancer = ImageEnhance.Contrast(screenshot)
        screenshot = contrast_enhancer.enhance(1.05)  # 軽微なコントラスト向上
        
        # シャープネス調整（自然な文字の鮮明さ）
        sharpness_enhancer = ImageEnhance.Sharpness(screenshot)
        screenshot = sharpness_enhancer.enhance(1.1)  # 控えめなシャープネス向上
        
        print(f"スクリーンショット品質向上処理完了: {screenshot.size}")
        return screenshot

    def automation_thread(self):
        max_pages = self.screenshot_count.get()
        
        # macOSの場合、最初にアプリ前面化処理を行う
        if platform.system() == "Darwin":  # macOS
            print("macOS: アプリケーション前面化処理開始")
            
            # 1枚目をスクリーンショット
            x1, y1, x2, y2 = self.screenshot_area
            screenshot = self.capture_high_quality_screenshot(x1, y1, x2 - x1, y2 - y1)
            self.images.append(screenshot)
            self.master.after(0, self.page_count_label.config, {"text": f"撮影枚数: 1/{max_pages}"})
            print("macOS: 1枚目スクリーンショット完了")
            
            # アプリ前面化のために画面をクリック（ページめくりしない位置）
            # クリック位置を一時的に画面中央に変更してアプリを前面に出す
            screen_width = self.master.winfo_screenwidth()
            screen_height = self.master.winfo_screenheight()
            temp_click_pos = (screen_width // 2, screen_height // 2)
            
            time.sleep(1.0)
            self.perform_app_focus_click(temp_click_pos)
            time.sleep(1.0)  # アプリ前面化待機
            print("macOS: アプリケーション前面化完了")
            
            # 2枚目以降のループ
            for page_count in range(2, max_pages + 1):
                if not self.is_running:
                    break
                
                # ページめくりクリック
                self.perform_click(self.click_position, is_first_click=False)
                time.sleep(self.wait_time.get())
                
                # スクリーンショット
                screenshot = self.capture_high_quality_screenshot(x1, y1, x2 - x1, y2 - y1)
                self.images.append(screenshot)
                self.master.after(0, self.page_count_label.config, {"text": f"撮影枚数: {page_count}/{max_pages}"})
        
        else:  # Windows など
            for page_count in range(1, max_pages + 1):
                if not self.is_running:
                    break

                x1, y1, x2, y2 = self.screenshot_area
                # 最高解像度でスクリーンショットを取得
                screenshot = self.capture_high_quality_screenshot(x1, y1, x2 - x1, y2 - y1)
                # PNG形式で高品質を保持（ロスレス圧縮）
                self.images.append(screenshot)
                
                self.master.after(0, self.page_count_label.config, {"text": f"撮影枚数: {page_count}/{max_pages}"})

                if page_count < max_pages:
                    wait = 1.0 if page_count == 1 else self.wait_time.get()
                    time.sleep(wait)
                    if self.is_running:
                        self.perform_click(self.click_position, is_first_click=False)
        
        self.master.after(0, self.open_editor)

    def perform_app_focus_click(self, position):
        """macOS用アプリ前面化専用クリック（ページめくりしない）"""
        if not position:
            print("アプリ前面化クリック位置が設定されていません")
            return
            
        x, y = position
        print(f"macOS: アプリ前面化クリック実行 - 座標({x}, {y})")
        
        try:
            # シンプルな1回クリックでアプリを前面に出す
            pyautogui.click(x, y)
            print("macOS: アプリ前面化クリック完了")
        except Exception as e:
            print(f"アプリ前面化クリックエラー: {e}")

    def perform_click(self, position, is_first_click=False):
        """ページめくり用クリック実行"""
        if not position:
            print("クリック位置が設定されていません")
            return
            
        x, y = position
        print(f"ページめくりクリック実行: 座標({x}, {y})")
        
        try:
            pyautogui.click(x, y)
            print("ページめくりクリック完了")
        except Exception as e:
            print(f"ページめくりクリックエラー: {e}")

    def update_settings_display(self):
        self.settings_display.config(state=tk.NORMAL)
        self.settings_display.delete(1.0, tk.END)
        self.settings_display.insert(tk.END, f"PDFパス: {self.pdf_path.get()}\n")
        self.settings_display.insert(tk.END, f"待機時間: {self.wait_time.get()}秒\n")
        self.settings_display.insert(tk.END, f"撮影枚数: {self.screenshot_count.get()}枚\n")
        self.settings_display.insert(tk.END, f"画質設定: {self.quality_scale.get()}倍拡大\n")
        
        if self.screenshot_area:
            self.settings_display.insert(tk.END, f"スクリーンショット範囲: {self.screenshot_area}\n")
        if self.click_position:
            self.settings_display.insert(tk.END, f"クリック位置: {self.click_position}\n")
        self.settings_display.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
