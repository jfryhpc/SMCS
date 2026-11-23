import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import os
import faulthandler
import audioop
import json
import traceback
import cv2
import mediapipe as mp
import pyautogui
from PIL import ImageFont, ImageDraw, Image, ImageTk
import numpy as np
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox

# 若程式是從 python.exe 啟動，隱藏黑色 CMD 視窗，保留純 GUI 顯示。
# 不直接呼叫 FreeConsole()，避免影響 Tkinter / 主程式啟動。 
def hide_console_window():
    if os.name != "nt":
        return
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass


hide_console_window()
import sys
import datetime
import time
import random
import winsound
import win32com.client
import pythoncom
import threading
import speech_recognition as sr
import pyaudio
import pygame
import webbrowser
from tkinter import scrolledtext
import tkinter.ttk as ttk 
from google import genai
from google.genai import types
from collections import deque
import queue
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import requests
import re
import math
import unicodedata
from urllib.parse import quote_plus

# 配置 Gemini API
GEMINI_MODEL = "gemini-3-flash-preview"
API_Key = "CWA-5A4A407C-FB57-4EB1-8D1D-F3295162DA89"
# =====================
# 常數與路徑
# =====================
BASE_DIR = os.path.dirname(
    os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
)
RESOURCE_DIR = getattr(sys, "_MEIPASS", BASE_DIR)
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
WINDOW_ICON_FILE = os.path.join(RESOURCE_DIR, "icon.ico")
GEMINI_API_KEY_FILE = os.path.join(BASE_DIR, "gemini_api_key.txt")
FONT_FILE = os.path.join(RESOURCE_DIR, "HuaKangTiFan-CuTi-1.otf")
THUMB, INDEX, MIDDLE, RING, PINKY = 4, 8, 12, 16, 20
PINCH_THRESHOLD = 0.035
CLICK_TIMEOUT = 0.5 


def load_gemini_api_key():
    environment_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if environment_key:
        return environment_key
    try:
        with open(GEMINI_API_KEY_FILE, "r", encoding="utf-8") as key_file:
            return key_file.read().strip()
    except FileNotFoundError:
        return "AQ.Ab8RN6LW7mvCHb6dL-Wfr5_cCoBCbGwJgtkp21h0k_8WRTsd9A"


GEMINI_API_KEY = load_gemini_api_key()
LANGUAGE_PRESET_FILES = {
    "國語 (華語)": "mandarin.json",
    "台文 (臺灣閩南語)": "taigi.json",
    "客語 (臺灣客家語)": "hakka.json",
    "English": "english.json",
}


def load_language_presets(base_dir):
    preset_dir = os.path.join(base_dir, "language_presets")
    loaded = {}

    for lang_name, filename in LANGUAGE_PRESET_FILES.items():
        preset_path = os.path.join(preset_dir, filename)
        try:
            with open(preset_path, "r", encoding="utf-8") as file:
                preset_data = json.load(file)
            if isinstance(preset_data, dict):
                loaded[lang_name] = preset_data
        except Exception as exc:
            raise RuntimeError(f"無法載入語言預設檔 {filename}: {exc}") from exc

    if not loaded:
        raise RuntimeError("沒有找到任何語言預設檔。")

    return loaded


I18N = load_language_presets(RESOURCE_DIR)

ENGLISH_RUNTIME_TRANSLATIONS = I18N["English"].get("RUNTIME_TRANSLATIONS", {})

DEFAULT_SETTINGS = {
    "language": "國語 (華語)",
    "user_name": "",
    "smooth_factor": 5.0,
    "speed_multiplier": 1.5,
    "camera_box_x_offset": 0,
    "camera_box_y_offset": 0,
    "camera_box_width": 40,
    "camera_box_height": 40,
    "voice_speed": 1,
    "tts_enabled": True,
    "dwell_threshold": 1,
    "email_enabled": True,
    "receiver_email": "",
    "ui_style_mode": "精簡模式 (純純九宮格)" ,      
    "font_size": "normal",
    "enable_trail": True,
    "trail_length": 20,
    "theme_mode": "自動 (隨時間)",
    "fall_detection_enabled": True,
    "camera_index": 0,
    "microphone_index": 0,
    "microphone_gain": 1.0,
    "speaker_index": None,
    "camera_width": 640,
    "camera_height": 480,
    "setup_completed": False,
}

class FallDetector:
    def __init__(self,app,min_detection_confidence=0.55,min_tracking_confidence=0.55):
        self.app=app
        self.mp_pose=mp.solutions.pose
        self.closed=False
        self.pose=self.mp_pose.Pose(static_image_mode=False,model_complexity=1,smooth_landmarks=True,enable_segmentation=False,min_detection_confidence=min_detection_confidence,min_tracking_confidence=min_tracking_confidence)
        self.enabled=True
        self.history=deque(maxlen=30)
        self.last_frame_time=time.time()
        self.FALL_SPEED_THRESHOLD=0.045
        self.HORIZONTAL_ANGLE_THRESHOLD=58
        self.HEIGHT_DROP_RATIO=0.40
        self.CANDIDATE_SECONDS=0.8
        self.CONFIRM_SECONDS=3.0
        self.FALL_SCORE_THRESHOLD=70
        self.fall_candidate=False
        self.candidate_start_time=None
        self.fall_confirming=False
        self.confirm_start_time=None
        self.alert_active=False
        self.alert_start_time=None
        self.last_alert_time=0
        self.ALERT_COOLDOWN=60
        self.notification_sent=False
        self.warning_window=None
        self.fall_title_label=None
        self.fall_message_label=None
        self.fall_countdown_label=None
        print("【FallDetector】跌倒偵測模組已初始化")
    def process(self,frame):
        if self.closed or self.pose is None:
            return frame, "跌倒偵測已關閉", 0, False
        if not self.enabled:
            return frame, "跌倒偵測關閉", 0, False
        if self.alert_active:
            return self.draw_alert_status(frame, "🚨 疑似跌倒｜等待長者確認")
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        try:
            if self.pose is None:
                return frame, "跌倒偵測已關閉", 0, False
            results = self.pose.process(rgb)
        except Exception as e:
            print(f"【FallDetector】Pose 處理錯誤：{e}")
            return frame, "跌倒偵測錯誤", 0, False
        if not results.pose_landmarks:
            self.reset_tracking()
            cv2.putText(frame,"Fall Monitor: No Person",(20,110),cv2.FONT_HERSHEY_SIMPLEX,0.65,(150,150,150),2)
            return frame,"跌倒偵測：未偵測到人體",0,False
        landmarks=results.pose_landmarks.landmark
        P=self.mp_pose.PoseLandmark
        left_shoulder=landmarks[P.LEFT_SHOULDER.value]
        right_shoulder=landmarks[P.RIGHT_SHOULDER.value]
        left_hip=landmarks[P.LEFT_HIP.value]
        right_hip=landmarks[P.RIGHT_HIP.value]
        left_ankle=landmarks[P.LEFT_ANKLE.value]
        right_ankle=landmarks[P.RIGHT_ANKLE.value]
        important_points=[left_shoulder,right_shoulder,left_hip,right_hip]
        visible_count=sum(1 for p in important_points if p.visibility>0.45)
        if visible_count<3:
            self.reset_tracking()
            return frame,"跌倒偵測：人體姿勢不完整",0,False
        shoulder_x=(left_shoulder.x+right_shoulder.x)/2
        shoulder_y=(left_shoulder.y+right_shoulder.y)/2
        hip_x=(left_hip.x+right_hip.x)/2
        hip_y=(left_hip.y+right_hip.y)/2
        ankle_y=(left_ankle.y+right_ankle.y)/2
        body_height=abs(ankle_y-shoulder_y)
        dx=hip_x-shoulder_x
        dy=hip_y-shoulder_y
        angle=math.degrees(math.atan2(abs(dx),abs(dy)))
        now=time.time()
        self.history.append({"time":now,"shoulder_y":shoulder_y,"hip_y":hip_y,"body_height":body_height,"angle":angle})
        fall_speed=0
        if len(self.history)>=2:
            previous=self.history[-2]
            dt=now-previous["time"]
            if dt>0:fall_speed=(hip_y-previous["hip_y"])/dt
        height_drop=False
        if len(self.history)>=8:
            old_height=self.history[-8]["body_height"]
            if old_height>0 and body_height/old_height<(1-self.HEIGHT_DROP_RATIO):height_drop=True
        score=0
        if fall_speed>self.FALL_SPEED_THRESHOLD:score+=30
        if angle>self.HORIZONTAL_ANGLE_THRESHOLD:score+=30
        if height_drop:score+=20
        if angle>65:score+=20
        try:mp.solutions.drawing_utils.draw_landmarks(frame,results.pose_landmarks,self.mp_pose.POSE_CONNECTIONS)
        except Exception:pass
        cv2.putText(frame,f"Fall Score: {score}",(20,80),cv2.FONT_HERSHEY_SIMPLEX,0.65,self.get_score_color(score),2)
        cv2.putText(frame,f"Angle: {int(angle)}",(20,110),cv2.FONT_HERSHEY_SIMPLEX,0.55,(255,255,255),2)
        if score>=self.FALL_SCORE_THRESHOLD:
            if not self.fall_candidate:
                self.fall_candidate=True
                self.candidate_start_time=now
                print(f"【FallDetector】疑似跌倒開始 score={score}")
            elif now-self.candidate_start_time>=self.CANDIDATE_SECONDS:
                self.start_fall_confirmation()
                return self.draw_alert_status(frame,"🚨 疑似跌倒！")
        else:
            if self.fall_candidate:
                if now-self.candidate_start_time>1.5:print("【FallDetector】跌倒候選取消")
                self.fall_candidate=False
                self.candidate_start_time=None
        return frame,f"跌倒監測中｜分數 {score}",score,False
    def start_fall_confirmation(self, force=False):
        now = time.time()
        if not force and (now - self.last_alert_time < self.ALERT_COOLDOWN or self.alert_active):
            return
        self.alert_active = True
        self.alert_start_time = now
        self.confirm_start_time = now
        self.notification_sent = False
        self.fall_candidate = False
        self.candidate_start_time = None
        self.last_alert_time = now
        print("🚨【FallDetector】進入跌倒確認階段")
        threading.Thread(target=self._fall_warning_worker, daemon=True).start()
    def _fall_warning_worker(self):
        try:
            self.app.play_system_sound("warn")
            self.app.tts_speak("系統偵測到您可能跌倒了，請問您還好嗎？")
            self.app.log_status("🚨 系統偵測到疑似跌倒，等待長者回應。")
        except Exception as e:print(f"【FallDetector】語音提醒錯誤：{e}")
        try:self.app.root.after(0,self.show_fall_warning_window)
        except Exception as e:print(f"【FallDetector】建立確認視窗錯誤：{e}")
    def show_fall_warning_window(self):
        if self.warning_window is not None and self.warning_window.winfo_exists():return
        win=tk.Toplevel(self.app.root)
        self.warning_window=win
        win.title(self.app.localize_runtime_text("🚨 疑似跌倒"))
        win.geometry("900x560")
        win.configure(bg="#FFF7ED")
        win.attributes("-topmost",True)
        win.grab_set()

        self.fall_title_label = tk.Label(win,text=self.app.localize_runtime_text("🚨 疑似偵測到跌倒"),font=("Microsoft JhengHei",30,"bold"),bg="#FFF7ED",fg="#B91C1C")
        self.fall_title_label.pack(pady=(35,15))

        self.fall_message_label = tk.Label(
            win,
            text=self.app.localize_runtime_text("系統偵測到您可能跌倒了。\n\n請確認您現在是否安全。"),
            font=("Microsoft JhengHei",21),
            bg="#FFF7ED",
            fg="#333333",
            justify="center",
            wraplength=620,
        )
        self.fall_message_label.pack(pady=10)

        self.fall_countdown_label = tk.Label(
            win,
            text=self.app.localize_runtime_text("若沒有回應，10 秒後將通知家屬。"),
            font=("Microsoft JhengHei",17,"bold"),
            bg="#FFF7ED",
            fg="#B45309",
            justify="center",
            wraplength=620,
        )
        self.fall_countdown_label.pack(pady=(0, 15))

        button_frame=tk.Frame(win,bg="#FFF7ED")
        button_frame.pack(fill="both",expand=True,padx=35,pady=20)
        tk.Button(
            button_frame, text=self.app.localize_runtime_text("👍 我沒事\n\n點這裡取消警報"),
            font=("Microsoft JhengHei",20,"bold"), bg="#16A34A", fg="white",
            activebackground="#15803D", activeforeground="white", bd=0,
            wraplength=330, justify="center", command=self.cancel_fall_alert
        ).pack(side="left", fill="both", expand=True, padx=10)
        tk.Button(
            button_frame, text=self.app.localize_runtime_text("🚨 我需要幫忙\n\n通知家人"),
            font=("Microsoft JhengHei",20,"bold"), bg="#DC2626", fg="white",
            activebackground="#B91C1C", activeforeground="white", bd=0,
            wraplength=330, justify="center", command=self.confirm_fall_help
        ).pack(side="right", fill="both", expand=True, padx=10)
        win.protocol("WM_DELETE_WINDOW",self.cancel_fall_alert)
        win.after(1000,self.check_fall_timeout)
    def update_fall_warning_text(self, remaining=None):
        if self.warning_window is None or not self.warning_window.winfo_exists():
            return
        if self.fall_message_label is not None:
            self.fall_message_label.configure(text=self.app.localize_runtime_text("系統偵測到您可能跌倒了。\n\n請確認您現在是否安全。"))
        if self.fall_countdown_label is not None:
            if remaining is None:
                remaining_text = "若沒有回應，10 秒後將通知家屬。"
            else:
                remaining_text = f"若沒有回應，{remaining} 秒後將通知家屬。"
            self.fall_countdown_label.configure(text=self.app.localize_runtime_text(remaining_text))
    def check_fall_timeout(self):
        if not self.alert_active:return
        elapsed=time.time()-self.alert_start_time
        if elapsed>=10:
            print("🚨【FallDetector】長者 10 秒沒有回應")
            self.confirm_fall_help()
            return
        if self.warning_window is not None and self.warning_window.winfo_exists():
            remaining=max(0,10-int(elapsed))
            self.update_fall_warning_text(remaining)
            self.warning_window.after(1000,self.check_fall_timeout)
    def cancel_fall_alert(self):
        print("【FallDetector】長者回報：我沒事")
        self.alert_active = False
        self.notification_sent = False
        self.last_alert_time = 0
        self.app.play_system_sound("execute")
        self.app.tts_speak("好的，確認您沒事了。")
        self.app.log_status("🟢 長者已回報平安，跌倒警報解除。")
        self.close_warning_window()
        self.reset_tracking()
    def confirm_fall_help(self):
        if not self.alert_active or self.notification_sent:return
        self.notification_sent=True
        print("🚨【FallDetector】長者需要協助")
        self.app.play_system_sound("warn")
        self.app.tts_speak("好的，正在通知家屬，請先保持冷靜。")
        self.app.log_status("🚨 長者回報需要協助，正在通知家屬。")
        threading.Thread(target=self.send_fall_email,daemon=True).start()
        self.close_warning_window()
        self.app.root.after(3000,self.finish_alert)
    def send_fall_email(self):
        try:
            today_str = datetime.date.today().strftime("%Y-%m-%d")
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            self.app.send_email_notification(
                today_str,
                current_time,
                streak_days=self.app.settings.get("streak_days", 1),
                is_sos=True,
                alert_title="疑似跌倒警報",
                alert_details=(
                    "系統偵測到長者可能發生跌倒，且長者未在確認時間內回報安全，或主動按下「我需要幫忙」。"
                ),
            )
            print(f"🚨【FallDetector】已使用 SOS Email 流程發送跌倒通知")
        except Exception as e:
            print(f"❌【FallDetector】Email 發送失敗：{e}")
    def finish_alert(self):
        self.alert_active=False
        self.notification_sent=False
        self.last_alert_time = 0
        self.close_warning_window()
        self.reset_tracking()
    def close_warning_window(self):
        try:
            if self.warning_window is not None and self.warning_window.winfo_exists():
                try:self.warning_window.grab_release()
                except Exception:pass
                self.warning_window.destroy()
        except Exception:pass
        self.warning_window=None
        self.fall_title_label=None
        self.fall_message_label=None
        self.fall_countdown_label=None
    def reset_tracking(self):
        self.history.clear()
        self.fall_candidate=False
        self.candidate_start_time=None
    def get_score_color(self,score):
        if score>=70:return (0,0,255)
        if score>=40:return (0,165,255)
        return (0,255,0)
    def draw_alert_status(self,frame,text):
        overlay=frame.copy()
        h,w=frame.shape[:2]
        cv2.rectangle(overlay,(0,0),(w,90),(0,0,180),-1)
        cv2.addWeighted(overlay,0.75,frame,0.25,0,frame)
        cv2.putText(frame,text,(20,55),cv2.FONT_HERSHEY_SIMPLEX,0.85,(255,255,255),2)
        return frame,text,100,True
    def close(self):
        self.closed = True
        self.enabled = False
        self.alert_active = False
        self.fall_candidate = False
        self.candidate_start_time = None
        try:
            if self.pose is not None:
                self.pose.close()
        except Exception:
            pass
        finally:
            self.pose = None
        self.close_warning_window()
        print("【FallDetector】已關閉")

class SmartCareSystem:
    def center_window(self, window, width, height):
        """將指定視窗置於目前螢幕中央。"""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    return {**DEFAULT_SETTINGS, **json.load(f)}
            except Exception:
                return DEFAULT_SETTINGS
        return DEFAULT_SETTINGS

    def save_settings(self, settings):
        with self._settings_io_lock:
            if getattr(self, "_settings_file_cleared", False):
                return
            try:
                with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                    json.dump(settings, f, ensure_ascii=False, indent=4)
            except Exception as e:
                print(f"儲存設定失敗: {e}")

    def bind_button_tween(self, widget):
        if getattr(widget, "_button_tween_state", None) is not None:
            return
        try:
            font = tkfont.Font(root=widget, font=widget.cget("font"))
            current_font_size = abs(int(font.actual()["size"]))
            padx = widget.tk.getint(widget.cget("padx"))
            pady = widget.tk.getint(widget.cget("pady"))
            border_width = widget.tk.getint(widget.cget("bd"))
        except (tk.TclError, TypeError, ValueError):
            return

        base_font_size = max(
            1, round(current_font_size / self._font_size_multiplier())
        )
        state = {
            "scale": 1.0,
            "after_id": None,
            "base_font_size": base_font_size,
            "font": font,
            "padx": padx,
            "pady": pady,
            "relief": widget.cget("relief"),
            "bd": border_width,
        }

        def animate_to(target_scale):
            if not widget.winfo_exists():
                return
            current = state["scale"]
            if state["after_id"] is not None:
                try:
                    widget.after_cancel(state["after_id"])
                except tk.TclError:
                    pass
            steps = 7
            duration_ms = 140

            def step(index):
                if not widget.winfo_exists():
                    return
                progress = index / steps
                eased = 1 - (1 - progress) ** 3
                scale = current + (target_scale - current) * eased
                state["scale"] = scale
                font_size = max(
                    1,
                    round(
                        state["base_font_size"]
                        * self._font_size_multiplier()
                        * scale
                    ),
                )
                state["font"].configure(size=font_size)
                widget.configure(
                    font=state["font"],
                    padx=round(state["padx"] * scale),
                    pady=round(state["pady"] * scale),
                    relief=(
                        "solid" if scale > 1.02 and state["relief"] == "raised"
                        else state["relief"]
                    ),
                    bd=state["bd"] + (1 if scale > 1.02 and state["bd"] > 0 else 0),
                )
                if index < steps:
                    state["after_id"] = widget.after(
                        max(1, duration_ms // steps), lambda: step(index + 1)
                    )
                else:
                    state["after_id"] = None

            step(0)

        state["animate"] = animate_to
        widget._button_tween_state = state

    def _on_button_tween_enter(self, event):
        widget = event.widget
        if not isinstance(widget, (tk.Button, tk.Checkbutton, tk.Radiobutton)):
            return
        self.bind_button_tween(widget)
        state = getattr(widget, "_button_tween_state", None)
        if state is not None:
            state["animate"](1.07)

    def _on_button_tween_leave(self, event):
        state = getattr(event.widget, "_button_tween_state", None)
        if state is not None:
            state["animate"](1.0)

    def transition_text(self, widget, text, configure=None):
        if not widget.winfo_exists():
            return
        current_text = str(widget.cget("text"))
        if not current_text:
            if configure is not None:
                configure()
            widget.configure(text=text)
            return
        if current_text == text:
            if configure is not None:
                configure()
            return

        token = getattr(widget, "_text_transition_token", 0) + 1
        widget._text_transition_token = token
        previous_job = getattr(widget, "_text_transition_after_id", None)
        if previous_job is not None:
            try:
                widget.after_cancel(previous_job)
            except tk.TclError:
                pass

        try:
            start_color = widget.winfo_rgb(widget.cget("fg"))
            start_bg = widget.winfo_rgb(widget.cget("bg"))
        except tk.TclError:
            if configure is not None:
                configure()
            widget.configure(text=text)
            return

        steps = 5

        def blend_color(start, end, amount):
            channels = [
                round(a + (b - a) * amount) // 257
                for a, b in zip(start, end)
            ]
            return "#" + "".join(f"{channel:02x}" for channel in channels)

        transition_state = {
            "phase": "out",
            "start_color": start_color,
            "target_color": start_bg,
        }

        def animate(index):
            if not widget.winfo_exists() or widget._text_transition_token != token:
                return
            progress = index / steps
            eased = progress * progress * (3 - 2 * progress)
            if transition_state["phase"] == "out":
                widget.configure(
                    fg=blend_color(transition_state["start_color"], transition_state["target_color"], eased)
                )
                if index < steps:
                    widget._text_transition_after_id = widget.after(18, lambda: animate(index + 1))
                    return
                widget.configure(text=text)
                faded_foreground = widget.winfo_rgb(widget.cget("fg"))
                original_foreground = transition_state["start_color"]
                if configure is not None:
                    configure()
                configured_foreground = widget.winfo_rgb(widget.cget("fg"))
                transition_state["phase"] = "in"
                transition_state["start_color"] = widget.winfo_rgb(widget.cget("bg"))
                transition_state["target_color"] = (
                    original_foreground
                    if configured_foreground == faded_foreground
                    else configured_foreground
                )
                widget.configure(fg=blend_color(
                    transition_state["start_color"], transition_state["target_color"], 0
                ))
                animate(0)
                return
            widget.configure(
                fg=blend_color(transition_state["start_color"], transition_state["target_color"], eased)
            )
            if index < steps:
                widget._text_transition_after_id = widget.after(18, lambda: animate(index + 1))
            else:
                widget._text_transition_after_id = None

        animate(0)

    def _font_size_multiplier(self):
        return {
            "small": 0.9,
            "normal": 1.0,
            "large": 1.15,
            "largest": 1.3,
        }.get(self.settings.get("font_size", "normal"), 1.0)

    def apply_font_size(self):
        """Resize existing Tk fonts and popup windows immediately."""
        scale = self._font_size_multiplier()

        if not hasattr(self, "_font_size_cache"):
            self._font_size_cache = {}

        def scaled_font(font_spec, cache_key):
            if cache_key not in self._font_size_cache:
                font = tkfont.Font(root=self.root, font=font_spec)
                self._font_size_cache[cache_key] = (
                    font, abs(int(font.actual()["size"]))
                )
            font, base_size = self._font_size_cache[cache_key]
            font.configure(size=max(1, round(base_size * scale)))
            return font

        def resize_widget_fonts(parent):
            for widget in parent.winfo_children():
                try:
                    tween_state = getattr(widget, "_button_tween_state", None)
                    if tween_state is not None:
                        tween_state["font"].configure(
                            size=max(
                                1,
                                round(
                                    tween_state["base_font_size"]
                                    * scale
                                    * tween_state["scale"]
                                ),
                            )
                        )
                        widget.configure(font=tween_state["font"])
                    else:
                        font_spec = widget.cget("font")
                        if font_spec:
                            font = scaled_font(font_spec, ("widget", str(widget)))
                            widget.configure(font=font)
                except tk.TclError:
                    pass
                if isinstance(widget, tk.Canvas):
                    try:
                        if not hasattr(widget, "_font_size_item_cache"):
                            widget._font_size_item_cache = {}
                        for item_id in widget.find_all():
                            if widget.type(item_id) != "text":
                                continue
                            font_spec = widget.itemcget(item_id, "font")
                            if font_spec:
                                font = scaled_font(
                                    font_spec, ("canvas", str(widget), item_id)
                                )
                                widget.itemconfigure(item_id, font=font)
                    except tk.TclError:
                        pass
                resize_widget_fonts(widget)

        resize_widget_fonts(self.root)
        if hasattr(self, "_settings_tab_font"):
            self._settings_tab_font.configure(
                size=max(1, round(12 * scale))
            )
        if hasattr(self, "font_file"):
            for attribute, base_size in (
                ("main_font", 36),
                ("boot_font", 40),
            ):
                try:
                    setattr(
                        self,
                        attribute,
                        ImageFont.truetype(
                            self.font_file, max(1, round(base_size * scale))
                        ),
                    )
                except OSError as exc:
                    print(f"【字型】套用介面字體大小失敗: {exc}")
        self._resize_popup_windows(scale)
        if hasattr(self, "network_status_label"):
            self._fit_network_status_label()
        if hasattr(self, "grid_feature_buttons"):
            self._fit_grid_feature_texts()

    def _fit_grid_feature_text(self, event_or_button):
        button = getattr(event_or_button, "widget", event_or_button)
        if not isinstance(button, tk.Button):
            return
        available_width = button.winfo_width()
        if available_width <= 1:
            return
        padding = (
            2 * int(button.cget("padx"))
            + 2 * int(button.cget("bd"))
            + 12
        )
        button.configure(
            wraplength=max(80, available_width - padding),
            justify="center",
        )

    def _fit_grid_feature_texts(self):
        for button in getattr(self, "grid_feature_buttons", ()):
            if button.winfo_exists():
                self._fit_grid_feature_text(button)

    def _resize_popup_windows(self, scale):
        """Scale popup geometry with the selected font size without leaving the screen."""
        if not hasattr(self, "root"):
            return

        popup_windows = []

        def collect_popups(parent):
            for child in parent.winfo_children():
                if isinstance(child, tk.Toplevel):
                    popup_windows.append(child)
                collect_popups(child)

        try:
            collect_popups(self.root)
        except tk.TclError:
            return

        for window in popup_windows:
            try:
                if not window.winfo_exists():
                    continue
                window.update_idletasks()
                if not hasattr(window, "_font_size_base_geometry"):
                    width = window.winfo_width()
                    height = window.winfo_height()
                    if width <= 1 or height <= 1:
                        geometry = re.match(
                            r"(\d+)x(\d+)([+-]\d+)([+-]\d+)",
                            window.geometry(),
                        )
                        if geometry:
                            width = int(geometry.group(1))
                            height = int(geometry.group(2))
                    if width <= 1 or height <= 1:
                        continue
                    window._font_size_base_geometry = (
                        width,
                        height,
                        window.winfo_x() + width // 2,
                        window.winfo_y() + height // 2,
                    )

                base_width, base_height, center_x, center_y = (
                    window._font_size_base_geometry
                )
                screen_width = window.winfo_screenwidth()
                screen_height = window.winfo_screenheight()
                width = min(
                    max(320, round(base_width * scale)),
                    max(320, screen_width - 60),
                )
                height = min(
                    max(240, round(base_height * scale)),
                    max(240, screen_height - 80),
                )
                x = min(max(0, center_x - width // 2), screen_width - width)
                y = min(max(0, center_y - height // 2), screen_height - height)
                window.geometry(f"{width}x{height}+{x}+{y}")
            except tk.TclError:
                continue

    def apply_language(self):
        """根據目前的設定同步更新主畫面的語言介面"""
        lang = self.settings.get("language", "國語 (華語)")
        lang_dict = I18N.get(lang, I18N["國語 (華語)"])
        user_name = str(self.settings.get("user_name", "")).strip()
        #  更新 Gemini API Prompt 語言
        self.update_gemini_prompt_language()
        if hasattr(self, "root"):
            self.root.title(lang_dict.get("main_title_header", "智慧多媒體操作平台"))

        #  更新主畫面左下角提示 [新增此段]
        if getattr(self, 'status_bar', None) is not None:
            self.transition_text(
                self.status_bar,
                lang_dict.get("status_ready", "系統狀態：準備就緒。將右手食指移到按鈕上不動即可點選。"),
            )
        if hasattr(self, 'lbl_title'):
            self.transition_text(
                self.lbl_title,
                lang_dict.get("main_title_header", "智慧多媒體操作平台"),
            )
        if hasattr(self, 'title_label'):
            title_text = lang_dict.get("main_title_sub", "您好！今天想做點什麼呢？")
            if user_name:
                if lang == "English":
                    title_text = f"Hello, {user_name}! What would you like to do today?"
                elif lang == "台文 (臺灣閩南語)":
                    title_text = f"{user_name}，{title_text}"
                else:
                    title_text = f"{user_name}，{title_text}"
            self.transition_text(self.title_label, title_text)

        #  更新頂部與狀態欄按鈕文字
        if hasattr(self, 'btn_switch_grid'):
            self.transition_text(
                self.btn_switch_grid,
                lang_dict.get("btn_grid_mode", "📱 顯示九宮格清單"),
            )
        if hasattr(self, 'btn_switch_media'):
            self.transition_text(
                self.btn_switch_media,
                lang_dict.get("btn_media_mode", "🎬 顯示左右媒體庫"),
            )
        if hasattr(self, 'mic_button') and not getattr(self, "is_listening", False):
            self.transition_text(
                self.mic_button,
                lang_dict.get("btn_mic_off", "🎤 點我說話 (關閉中)"),
            )
        if hasattr(self, 'mode_status_label'):
            prefix = lang_dict.get("mode_status_prefix", "📊 當前顯示樣式：")
            mode_name = self.localize_runtime_text(self.current_style_mode)
            self.transition_text(self.mode_status_label, f"{prefix}{mode_name}")
        if hasattr(self, 'btn_sign'):
            self.transition_text(
                self.btn_sign,
                self.localize_runtime_text("📋 補簽到/查看紀錄"),
            )
        if hasattr(self, 'lbl_weather') and self.settings.get("language") == "English":
            self.transition_text(
                self.lbl_weather,
                self.localize_runtime_text(self.lbl_weather.cget("text")),
            )
        if hasattr(self, 'lbl_weather') and self.settings.get("language") != "English":
            self.transition_text(self.lbl_weather, "☀ 天氣載入中...")
        if hasattr(self, 'status_bar'):
            self.transition_text(
                self.status_bar,
                lang_dict.get("status_ready", "系統狀態：準備就緒。"),
            )
        if hasattr(self, "network_status_label"):
            self.update_network_status_label()

        #  更新媒體庫輪播按鈕
        if hasattr(self, 'btn_media_prev'):
            self.transition_text(
                self.btn_media_prev,
                lang_dict.get("btn_prev", "◀ 上一個功能"),
            )
        if hasattr(self, 'btn_media_next'):
            self.transition_text(
                self.btn_media_next,
                lang_dict.get("btn_next", "下一個功能 ▶"),
            )

        #  更新 9 大功能選單清單
        new_texts = [
            lang_dict.get("btn_rehab", "做復健運動\n(觀看影片)"),
            lang_dict.get("btn_finger", "手指健康操\n(活化大腦)"),
            lang_dict.get("btn_scenery", "看懷舊風景\n(放鬆心情)"),
            lang_dict.get("btn_health_net", "衛教健康網\n(健康查詢)"),
            lang_dict.get("btn_game", "數字記憶挑戰"),
            lang_dict.get("btn_ai_chat", "AI 照護助理\n(對話聊天)"),
            lang_dict.get("btn_settings", "系統參數\n(設定選單)"),
            lang_dict.get("btn_restart_cam", "相機視窗\n(即時畫面)"),
            lang_dict.get("btn_exit", "關閉程式\n(退出系統)")
        ]

        if hasattr(self, 'all_system_features'):
            for idx, text in enumerate(new_texts):
                if idx < len(self.all_system_features):
                    self.all_system_features[idx]["text"] = text

            # 重新刷洗九宮格容器內各按鈕文字
            if hasattr(self, 'grid_container_frame'):
                for idx, child in enumerate(self.grid_container_frame.winfo_children()):
                    if isinstance(child, tk.Button) and idx < len(self.all_system_features):
                        self.transition_text(
                            child, self.all_system_features[idx]["text"]
                        )

            # 刷新媒體庫卡片
            if hasattr(self, 'update_media_gallery_display'):
                self.update_media_gallery_display()
            self._fit_grid_feature_texts()
    def update_gemini_prompt_language(self):
        """根據目前系統語言動態更新 Gemini 的 System Instruction 與對話實體"""
        lang = self.settings.get("language", "國語 (華語)")
        lang_dict = I18N.get(lang, I18N["國語 (華語)"])

        # 讀取該語言對應的系統指令
        sys_instruction = lang_dict.get(
            "gemini_sys_instruction",
            I18N["國語 (華語)"]["gemini_sys_instruction"]
        )

        # 重新建立 Config 與 Chat Session
        self.ai_config = types.GenerateContentConfig(
            system_instruction=sys_instruction,
            temperature=0.6,
        )

        self.gemini_chat = self.ai_client.chats.create(
            model=self.model_id,
            config=self.ai_config
        )
        print(f"【Gemini】Prompt 已更新為「{lang}」語言模式。")

    def _apply_window_icon_to(self, window):
        if not os.path.isfile(WINDOW_ICON_FILE):
            print(f"【視窗圖示】找不到圖示檔案：{WINDOW_ICON_FILE}")
            return
        try:
            window.iconbitmap(WINDOW_ICON_FILE)
        except tk.TclError as exc:
            print(f"【視窗圖示】套用圖示失敗：{exc}")

    def _apply_window_icon(self, event):
        if isinstance(event.widget, (tk.Tk, tk.Toplevel)):
            self._apply_window_icon_to(event.widget)

    def _load_icon_photo(self, size):
        try:
            with Image.open(WINDOW_ICON_FILE) as icon:
                image = icon.convert("RGBA")
                image.thumbnail((size, size), Image.Resampling.LANCZOS)
                return ImageTk.PhotoImage(image, master=self.root)
        except OSError as exc:
            print(f"【視窗圖示】無法載入圖示圖片：{exc}")
            return None

    def __init__(self, root):
        self.root = root
        self._crash_log_file = open(
            os.path.join(BASE_DIR, "crash.log"),
            "a",
            encoding="utf-8",
        )
        faulthandler.enable(file=self._crash_log_file, all_threads=True)
        self.root.report_callback_exception = self._report_tk_exception
        threading.excepthook = self._report_thread_exception
        self.root.bind_all("<Enter>", self._on_button_tween_enter, add="+")
        self.root.bind_all("<Leave>", self._on_button_tween_leave, add="+")
        self.root.bind_all("<Map>", self._apply_window_icon, add="+")
        self._apply_window_icon_to(self.root)
        self.root.title("智慧多媒體操作平台")
        self.root.geometry("1440x900")
        self.center_window(self.root, 1440, 900)
        self.root.resizable(True, True)
        self.root.configure(bg="#EBF8FF") # 💡 初始背景色
        # 🆘 雙手舉起 SOS 狀態變數
        self.sos_counter = 0
        self.SOS_FRAMES_REQUIRED = 15  # 連續 15 幀舉手即觸發 (約 0.5~1 秒)
        self.last_sos_time = 0
        self.SOS_COOLDOWN = 60         # 觸發後 60 秒冷卻，避免重複發送
        self._settings_io_lock = threading.Lock()
        # 載入設定參數
        self.settings = self.load_settings()
        startup_frame = tk.Frame(self.root, bg="#EBF8FF")
        startup_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        startup_text = (
            "Loading, please wait..."
            if self.settings.get("language") == "English"
            else "加載中，請稍候..."
        )
        self.startup_icon_image = self._load_icon_photo(48)
        self.startup_label = tk.Label(
            startup_frame,
            text=startup_text,
            image=self.startup_icon_image,
            compound=tk.LEFT,
            font=("Microsoft JhengHei", 22, "bold"),
            bg="#EBF8FF",
            fg="#1E3A5F",
            padx=12,
        )
        self.startup_label.pack(expand=True)
        startup_frame.lift()
        self.root.update_idletasks()
        self.scan_media_devices()
        if "history_dates" in self.settings:
            # 計算 7 天前的邊界日期字串 (含今天共7天，所以 days=6)
            seven_days_ago = (datetime.datetime.now() - datetime.timedelta(days=6)).strftime("%Y-%m-%d")
            
            # 僅保留大於等於 7 天前的紀錄
            self.settings["history_dates"] = [
                d for d in self.settings["history_dates"] 
                if d >= seven_days_ago
            ]
            
            # 清理完立刻存檔，保持 json 檔案乾淨
            self.save_settings(self.settings)
        # 多媒體與狀態控制變數
        self.is_paused = False
        self.cap = None
        self.camera_ready = False
        self._camera_init_generation = 0
        self.latest_camera_frame = None
        self.latest_camera_preview_frame = None
        self.camera_frame_lock = threading.Lock()
        self.camera_capture_lock = threading.Lock()
        self.camera_release_complete = threading.Event()
        self.camera_release_complete.set()
        self.camera_ui_queue = queue.Queue()
        self._camera_preview_lock = threading.Lock()
        self._camera_preview_stop_events = set()
        self._camera_preview_captures = {}
        self._shutting_down = False
        self.loop_active = True
        self.is_refreshing = False # 用於鎖定相機重啟狀態
        self.gesture_game_active = False
        self.gesture_game_window_active = False
        self.gesture_game_state = None
        self.modal_gesture_suspended = False
        t_len = int(self.settings.get("trail_length", 20))
        self.trail_points = deque(maxlen=t_len)
        self.target_county = "臺中市"
        self.last_earthquake_no = None  # 用來記錄上次處理過的地震編號，避免重複彈窗
        self.is_listening = False  # 標記目前麥克風是否處於開啟錄音狀態
        self.mic_button = None     # 儲存麥克風按鈕物件

        # 啟動背景執行緒，不卡死主執行緒 UI
        threading.Thread(target=self.auto_monitor_earthquake_loop, daemon=True).start()

        # 智慧滑鼠平滑緩衝與停留計時器 (Dwell Click)
        self.pos_buffer = deque(maxlen=int(self.settings.get("smooth_factor", 5)))
        self.last_cursor_pos = (0, 0)
        self.dwell_start_time = None
        self.right_click_gesture_active = False
        self.right_click_gesture_fired = False
        
        # 初始化音效與語音
        pygame.mixer.init()
        self.tts_lock = threading.Lock()
        
        # 串接 Gemini 
        if not GEMINI_API_KEY:
            raise RuntimeError(
                "請設定 GEMINI_API_KEY 環境變數，或在程式／exe 同資料夾"
                "建立 gemini_api_key.txt 並放入 Gemini API 金鑰。"
            )
        self.ai_client = genai.Client(api_key=GEMINI_API_KEY)
        self.model_id = GEMINI_MODEL

        # 2. 建立專屬溫柔護理師的系統指令設定 (System Instruction)
        self.ai_config = types.GenerateContentConfig(
            system_instruction=(
                "你是一位專門照顧長者的溫柔社區護理師。當長者或家屬詢問任何與健康、身體異狀或日常困擾有關的問題時，"
                "請根據長者具體的提問（例如：問忘記事情、問做惡夢、問情緒不穩），用最溫柔、極其簡單的語言解釋可能的原因與照護方式。"
                "【重要限制】：請針對長者當下的症狀進行親切的科學解釋（例如：解釋做惡夢可能是睡眠品質受到影響），"
                "絕對不要在每次回答都重複使用相同的比喻（如大腦時鐘、橡皮擦等固定詞彙），必須根據對話自然變換說法。"
                "字數嚴格控制在 50 個字以內。不論長者重覆問幾次，都要像第一次聽到一樣，不厭其煩地用不同的溫柔口吻安慰並給予安心感，千萬不要露出破綻。"
            ),
            temperature=0.6, 
        )

        # 3. 創立對話時綁定 config
        self.gemini_chat = self.ai_client.chats.create(
            model=self.model_id, 
            config=self.ai_config
        )
        # 初始化 MediaPipe
        try:
            self.mp_hands = mp.solutions.hands
            self.hands = self.mp_hands.Hands(
                max_num_hands=2,
                min_detection_confidence=0.7,
                min_tracking_confidence=0.7
            )
        except AttributeError:
            print("【錯誤】偵測到新版 MediaPipe。請在終端機執行 pip install mediapipe==0.10.14 來降級。")
            sys.exit(1)

        # --- 從 settings 讀取上次儲存的樣式設定 ---
        saved_mode = self.settings.get("ui_style_mode", "精簡模式 (純純九宮格)")
        if saved_mode in ["精簡模式 (純純九宮格)", "焦點媒體模式 (專注看影片)"]:
            self.current_style_mode = saved_mode
        else:
            self.current_style_mode = "精簡模式 (純純九宮格)"

        # 載入中文字型
        self.font_file = FONT_FILE
        try:
            self.main_font = ImageFont.truetype(FONT_FILE , 36)
            self.boot_font = ImageFont.truetype(FONT_FILE , 40)
            print(f"【字型】成功載入介面字型路徑: {FONT_FILE}")
        except Exception as e:
            self.main_font = ImageFont.load_default()
            self.boot_font = ImageFont.load_default()
            print(f"【字型】字型載入失敗，使用系統預設字型。錯誤: {e}")

        # 初始化語音辨識與指令
        self.init_speech_control()
        # 建立頂部狀態列
        self.setup_header()
        # 建立大字體九宮格主首頁
        self.create_grid_ui()
        self.apply_font_size()
        self.apply_language()
        startup_frame.destroy()
        self.root.update_idletasks()
        if not self.settings.get("setup_completed", False):
            self.open_setup_wizard(self.finish_startup)
        else:
            self.finish_startup()

    def _write_crash_log(self, source, exc_info):
        timestamp = datetime.datetime.now().isoformat(timespec="seconds")
        self._crash_log_file.write(f"\n[{timestamp}] {source}\n")
        traceback.print_exception(*exc_info, file=self._crash_log_file)
        self._crash_log_file.flush()

    def _report_tk_exception(self, exc_type, exc_value, exc_traceback):
        self._write_crash_log(
            "Unhandled Tk callback exception",
            (exc_type, exc_value, exc_traceback),
        )

    def _report_thread_exception(self, args):
        self._write_crash_log(
            f"Unhandled exception in thread {args.thread.name}",
            (args.exc_type, args.exc_value, args.exc_traceback),
        )

    def finish_startup(self):
        """設定精靈結束後才啟動相機、語音與每日簽到流程。"""
        if getattr(self, "_startup_finished", False):
            return
        self._startup_finished = True

        # 背景更新天氣
        self.update_weather()
        self.check_and_apply_theme()
        # 建立獨立相機監控視窗
        self.create_camera_window()

        # 相機視窗預設隱藏；視窗未開啟時不初始化攝影機。
        self.init_camera(is_boot=True)
        self.start_voice_listening()
        # 綁定視窗右上角 [X] 關閉按鈕
        self.root.protocol("WM_DELETE_WINDOW", self.safe_exit)
        self.tts_speak("系統已啟動。請用右手移動滑鼠，停留在按鈕上即可點選。")
        self.root.after(2000, self.check_daily_attendance_on_boot)

    def scan_media_devices(self):
        """Scan available audio and camera devices before building the main UI."""
        self.audio_input_devices = []
        self.audio_output_devices = []
        self.audio_device_error = None
        self.camera_indices = []

        audio_enumerator = None
        try:
            audio_enumerator = pyaudio.PyAudio()
            seen_input_names = set()
            seen_output_names = set()
            for device_index in range(audio_enumerator.get_device_count()):
                device_info = audio_enumerator.get_device_info_by_index(
                    device_index
                )
                device_name = str(
                    device_info.get("name", f"裝置 {device_index}")
                )
                normalized_name = " ".join(device_name.split()).casefold()
                if normalized_name.startswith("microsoft ") and (
                    "mapper" in normalized_name or "對應表" in normalized_name
                ):
                    continue
                device = (device_name, device_index)
                if (
                    int(device_info.get("maxInputChannels", 0)) > 0
                    and normalized_name not in seen_input_names
                ):
                    self.audio_input_devices.append(device)
                    seen_input_names.add(normalized_name)
                if (
                    int(device_info.get("maxOutputChannels", 0)) > 0
                    and normalized_name not in seen_output_names
                ):
                    self.audio_output_devices.append(device)
                    seen_output_names.add(normalized_name)
        except Exception as exc:
            self.audio_device_error = exc
            print(f"音訊裝置列舉失敗: {exc}")
        finally:
            if audio_enumerator is not None:
                audio_enumerator.terminate()

        for camera_index in range(10):
            camera_probe = None
            try:
                camera_probe = cv2.VideoCapture(
                    camera_index, cv2.CAP_DSHOW
                )
                if camera_probe.isOpened():
                    self.camera_indices.append(camera_index)
            except Exception as exc:
                print(f"攝影機裝置 {camera_index} 掃描失敗: {exc}")
            finally:
                if camera_probe is not None:
                    camera_probe.release()

        print(
            "裝置掃描完成："
            f"{len(self.audio_input_devices)} 個音訊輸入、"
            f"{len(self.audio_output_devices)} 個音訊輸出、"
            f"{len(self.camera_indices)} 個攝影機。"
        )

    def trigger_sos_alert(self):
        """觸發 SOS 警報與強制發送 Email 通知"""
        print("🚨 偵測到雙手舉起！觸發 SOS 緊急求救警報！")
        
        # 1. 播放警報聲
        threading.Thread(target=lambda: winsound.Beep(2000, 1000), daemon=True).start()
        
        # 2. 語音提示
        self.tts_speak("偵測到緊急求救手勢，已自動發送緊急通知給家屬！")

        # 3. 強制發送 SOS Email (is_sos=True 忽略開關限制)
        today_str = datetime.date.today().strftime("%Y-%m-%d")
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.send_email_notification(
            today_str, 
            current_time, 
            streak_days=self.settings.get("streak_days", 1), 
            is_sos=True
        )

    def setup_header(self):
        # 1. 頂部狀態列容器 (改存為 self.header_frame 方便隨主題更新背景色)
        self.header_frame = tk.Frame(self.root, bg="#2C3E50", height=80)
        self.header_frame.pack(fill=tk.X, side=tk.TOP)
        self.header_frame.pack_propagate(False)

        # 2. 左側：系統標題 (改存為 self.lbl_title 方便隨主題變色)
        self.header_icon_image = self._load_icon_photo(36)
        self.lbl_title = tk.Label(
            self.header_frame, 
            text="智慧多媒體操作平台", 
            image=self.header_icon_image,
            compound=tk.LEFT,
            font=("Microsoft JhengHei", 20, "bold"), 
            fg="white", 
            bg="#2C3E50",
            padx=8,
        )
        self.lbl_title.pack(side=tk.LEFT, padx=20, pady=15)

        # 3. 中央：天氣顯示標籤
        self.lbl_weather = tk.Label(
            self.header_frame, 
            text="☀ 天氣載入中...", 
            font=("Microsoft JhengHei", 16, "bold"), 
            fg="#F39C12", 
            bg="#2C3E50",
            anchor="e",
            justify="right"
        )
        self.lbl_weather.pack(expand=True, fill=tk.BOTH, padx=(10, 20), pady=15)
    # =====================
    # 介面佈局 (UI Layout)
    # =====================
    def check_and_apply_theme(self):
        """根據時間或使用者設定自動切換主題配色 (全自動遍歷所有 UI 元件)"""
        mode = self.settings.get("theme_mode", "自動 (隨時間)")
        current_hour = datetime.datetime.now().hour
        
        # 判斷是否為夜間 (晚間 6 點至早上 6 點)
        is_dark = mode == "黑暗模式" or (mode == "自動 (隨時間)" and (current_hour >= 18 or current_hour < 6))
        
        if is_dark:
            bg_main = "#1E293B"      # 主畫面背景
            bg_bar = "#0F172A"       # 工具列/狀態欄背景
            bg_entry = "#334155"     # 輸入框深色背景
            fg_text = "#F8FAFC"      # 文字亮白色
            title_color = "#38BDF8"  # 標題亮藍色
            cursor_color = "#FFFFFF" # 游標顏色
        else:
            bg_main = "#EBF8FF"      # 主畫面淺藍背景
            bg_bar = "#DBEAFE"       # 工具列/狀態欄背景
            bg_entry = "#FFFFFF"     # 輸入框白底
            fg_text = "#1E293B"      # 文字深色
            title_color = "#1E293B"  # 標題深色
            cursor_color = "#000000" # 游標顏色

        # 1. 設置根視窗背景
        self.root.configure(bg=bg_main)

        # 2. 遞迴函式：深入每個 Frame 尋找並更新所有子元件
        def update_widget_colors(parent):
            for child in parent.winfo_children():
                # 遇到輸入框 (Entry)
                if isinstance(child, tk.Entry):
                    child.configure(
                        bg=bg_entry, 
                        fg=fg_text, 
                        insertbackground=cursor_color,
                        highlightbackground=bg_bar,
                        bd=1
                    )
                # 遇到容器 Frame
                elif isinstance(child, tk.Frame):
                    # 如果該 Frame 包含 Entry 或按鈕列，套用 bar 背景；否則套用主背景
                    child.configure(bg=bg_main)
                    update_widget_colors(child) # 繼續往下尋找子元件
                # 遇到文字標籤 Label
                elif isinstance(child, tk.Label):
                    if (
                        hasattr(self, "network_status_label")
                        and child == self.network_status_label
                    ):
                        self.update_network_status_label()
                    # 避免動到時間標籤特有顏色
                    elif hasattr(self, 'time_label') and child == self.time_label:
                        child.configure(bg=bg_main, fg="#60A5FA" if is_dark else "#1E3A8A")
                    elif hasattr(self, 'title_label') and child == self.title_label:
                        child.configure(bg=bg_main, fg=title_color)
                    else:
                        child.configure(bg=bg_main, fg=fg_text)

        # 3. 對主視窗執行全自動掃描染色
        update_widget_colors(self.root)

        # 4. 針對狀態欄額外補強深色
        if hasattr(self, 'status_bar'):
            self.status_bar.configure(bg=bg_bar, fg="#93C5FD" if is_dark else "#1E3A8A")
    def show_voice_commands_dialog(self):
        """跳出操作說明視窗，並進行語音說明"""
        self.play_system_sound("click")
        self.tts_speak("以下是系統支援的操作清單")

        # 建立獨立彈出視窗
        cmd_win = tk.Toplevel(self.root)
        cmd_win.title(self.localize_runtime_text("🗣️ 系統操作指令清單"))
        cmd_win.geometry("760x720")
        cmd_win.configure(bg="#F8FAFC")
        cmd_win.transient(self.root)
        cmd_win.grab_set()

        # 標題欄 (已修正 pady)
        lbl_title = tk.Label(
            cmd_win, text=self.localize_runtime_text("🗣️ 常用操作指令說明"), 
            font=("Microsoft JhengHei", 16, "bold"), 
            bg="#1E293B", fg="white", pady=12
        )
        lbl_title.pack(fill="x")

        # 可滾動容器區域
        canvas = tk.Canvas(cmd_win, bg="#F8FAFC", highlightthickness=0)
        command_scroll_style = "Command.Vertical.TScrollbar"
        ttk.Style(cmd_win).configure(command_scroll_style, width=28)
        scrollbar = ttk.Scrollbar(
            cmd_win, orient="vertical", command=canvas.yview,
            style=command_scroll_style
        )
        scrollable_frame = tk.Frame(canvas, bg="#F8FAFC", padx=15, pady=10)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- 🖱️ 新增：滑鼠滾輪事件繫結 ---
        def _on_mousewheel(event):
            # Windows 滾輪數值為 event.delta (+120 向上, -120 向下)
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            # Linux/Mac 支援
            elif event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

        # 當滑鼠進入/離開彈出視窗時綁定/解綁全局滾輪，確保滾動滑暢
        cmd_win.bind_all("<MouseWheel>", _on_mousewheel)
        cmd_win.bind_all("<Button-4>", _on_mousewheel)
        cmd_win.bind_all("<Button-5>", _on_mousewheel)

        # 關閉視窗時解除全域滾輪繫結，避免影響主視窗
        def _on_close():
            cmd_win.unbind_all("<MouseWheel>")
            cmd_win.unbind_all("<Button-4>")
            cmd_win.unbind_all("<Button-5>")
            cmd_win.destroy()

        cmd_win.protocol("WM_DELETE_WINDOW", _on_close)

        # 指令清單分類資料
        commands_data = [
            ("👉手勢操作", [
                ("右手食指", "操控滑鼠游標移動,停留在按鈕上即可執行左鍵點選"),
                ("右手食指、中指伸直並停留", "執行滑鼠右鍵點選"),
                ("左手四指", "暫停右手手勢滑鼠游標控制權"),
                ("左手三指", "恢復右手手勢滑鼠游標控制權"),
                ("雙手舉起", "觸發 SOS 緊急求救警報,並自動發送 Email 通知家屬")
            ]),
            ("🏃‍♂️健康與復健", [
                ("做復健運動", "選擇適合活動部位的復健影片"),
                ("手指健康操", "播放長者手指運動影片"),
                ("衛教健康網", "開啟國民健康署衛教網站")
            ]),
            ("🧠娛樂與生活", [
                ("看懷舊風景", "播放台灣懷舊風景 4K 放鬆影片"),
                ("記憶小遊戲", "啟動黑猩猩記憶力認知小遊戲"),
                ("照護助理", "開啟 AI 智慧照護對話視窗"),
                ("今日天氣", "播報當前地區的氣溫與天氣狀況")
            ]),
            ("⚙️系統與管理", [
                ("系統參數", "開啟系統設定與參數調整視窗"),
                ("相機視窗", "開啟相機視窗並顯示即時畫面"),
                ("指令說明", "顯示本操作指令幫助清單"),
                ("關閉程式", "安全結束並退出系統")
            ])
        ]
        if self.settings.get("language") == "English":
            commands_data = [
                (self.localize_runtime_text(category), [
                    (self.localize_runtime_text(command), self.localize_runtime_text(description))
                    for command, description in commands
                ])
                for category, commands in commands_data
            ]
        is_english = self.settings.get("language") == "English"
        command_prefix = "Command" if is_english else "口令"
        function_prefix = "Function" if is_english else "功能"

        # 動態繪製指令卡片
        for category_title, items in commands_data:
            # 分類標頭
            lbl_cat = tk.Label(
                scrollable_frame, text=category_title, 
                font=("Microsoft JhengHei", 13, "bold"), 
                bg="#F8FAFC", fg="#1E293B"
            )
            lbl_cat.pack(anchor="w", pady=(10, 5))

            for cmd_text, desc in items:
                item_box = tk.Frame(scrollable_frame, bg="#FFFFFF", bd=1, relief="solid", padx=10, pady=8)
                item_box.pack(fill="x", pady=4, ipadx=5)

                item_prefix = (
                    ("Gesture" if is_english else "手勢")
                    if category_title == self.localize_runtime_text("👉手勢操作")
                    else command_prefix
                )
                lbl_cmd = tk.Label(
                    item_box, text=f"{item_prefix}: 「{cmd_text}」", 
                    font=("Microsoft JhengHei", 11, "bold"), 
                    bg="#FFFFFF", fg="#2563EB"
                )
                lbl_cmd.pack(anchor="w")

                lbl_desc = tk.Label(
                    item_box, text=f"💡 {function_prefix}: {desc}", 
                    font=("Microsoft JhengHei", 10), 
                    bg="#FFFFFF", fg="#475569"
                )
                lbl_desc.pack(anchor="w", pady=(2, 0))

        # 底部關閉按鈕
        btn_close = tk.Button(
            cmd_win, text="Close window" if is_english else "關閉視窗",
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#64748B", fg="white", activebackground="#475569", activeforeground="white",
            bd=0, relief="flat", cursor="hand2", command=_on_close
        )
        btn_close.pack(fill="x", side="bottom", padx=15, pady=10, ipady=6)
    def update_weather(self):
        """跨執行緒安全的背景天氣抓取邏輯 (已修正 Thread/MainLoop 異常)"""
        if not hasattr(self, 'weather_queue'):
            self.weather_queue = queue.Queue()

        def _fetch_task():
            url = f"https://opendata.cwa.gov.tw/api/v1/rest/datastore/F-C0032-001?Authorization={API_Key}&locationName={self.target_county}"
            try:
                response = requests.get(url, timeout=5)
                data = response.json()
                
                if data.get("success") == "true":
                    location_data = data["records"]["location"][0]
                    elements = location_data["weatherElement"]
                    
                    wx = elements[0]["time"][0]["parameter"]["parameterName"]
                    pop = elements[1]["time"][0]["parameter"]["parameterName"]
                    min_t = elements[2]["time"][0]["parameter"]["parameterName"]
                    max_t = elements[4]["time"][0]["parameter"]["parameterName"]

                    icon = "☀️"
                    if "雨" in wx:
                        icon = "🌧️"
                    elif "陰" in wx or "多雲" in wx:
                        icon = "⛅"

                    if self.settings.get("language") == "English":
                        county_name = self.localize_runtime_text(self.target_county)
                        weather_name = next(
                            (
                                translated
                                for source, translated in (
                                    ("多雲時晴", "Partly cloudy, occasionally clear"),
                                    ("晴時多雲", "Mostly clear, occasionally cloudy"),
                                    ("多雲", "Partly cloudy"),
                                    ("陰", "Cloudy"),
                                    ("雨", "Rain"),
                                    ("晴", "Clear"),
                                )
                                if source in wx
                            ),
                            wx,
                        )
                        weather_text = f"{icon} {county_name} {weather_name}  {min_t}°C ~ {max_t}°C  Rain chance {pop}%"
                    else:
                        weather_text = f"{icon} {self.target_county} {wx}  {min_t}°C ~ {max_t}°C  降雨 {pop}%"
                    self.weather_queue.put(("SUCCESS", weather_text))
                else:
                    self.weather_queue.put(("FAIL", "⚠️ 天氣資訊取得失敗"))
            except Exception as e:
                print(f"[Error] 天氣 API 讀取異常: {e}")
                self.weather_queue.put(("FAIL", "⚠️ 天氣連線異常"))

        def _check_queue():
            """由主執行緒 (Main Thread) 定期檢查佇列，安全更新 UI"""
            try:
                status, msg = self.weather_queue.get_nowait()
                # 只有在主執行緒中觸發 log_status 和 label config
                if hasattr(self, 'lbl_weather') and self.lbl_weather.winfo_exists():
                    self.transition_text(
                        self.lbl_weather, self.localize_runtime_text(msg)
                    )
                if status == "SUCCESS" and hasattr(self, 'log_status'):
                    self.log_status(msg)
            except queue.Empty:
                # 若無資料，繼續排程下一次監聽 (每 500ms 一次)
                if hasattr(self, 'root') and self.root.winfo_exists():
                    self.root.after(500, _check_queue)
            except Exception as err:
                print(f"[Error] UI 天氣更新異常: {err}")

        # 1. 啟動背景抓取執行緒
        threading.Thread(target=_fetch_task, daemon=True).start()
        
        # 2. 啟動主執行緒安全輪詢
        self.root.after(100, _check_queue)

        # 3. 每 1 小時 (3,600,000 毫秒) 自動重新呼叫一次 update_weather
        self.root.after(3600000, self.update_weather)
    def speak_today_weather(self):
        """提供給語音指令或按鈕點擊時呼叫"""
        if not hasattr(self, 'lbl_weather'):
            self.tts_speak("抱歉，天氣資訊尚未載入。")
            return
            
        weather_info = self.lbl_weather.cget("text")
        if "°C" in weather_info:
            # 使用正則表達式移除常見的 Unicode Emoji 圖示
            clean_info = re.sub(r'[\U00010000-\U0010ffff\u2600-\u26FF\u2700-\u27BF]', '', weather_info).strip()
            speech_text = f"報告長輩，現在{clean_info}"
            self.tts_speak(speech_text)
        else:
            self.tts_speak("抱歉，目前無法取得氣象資料。")
    def auto_monitor_earthquake_loop(self):
        """背景定期輪詢執行緒：檢查 1~2 天內最新地震 (放寬時間範圍並依地區與純數字門檻篩選)"""
        api = f"https://opendata.cwa.gov.tw/api/v1/rest/datastore/E-A0015-001?Authorization={API_Key}&limit=1&format=JSON"

        while getattr(self, "loop_active", True):
            if not self.settings.get("earthquake_enabled", True):
                time.sleep(10)
                continue
            try:
                r = requests.get(api, headers={'Accept': 'application/json'}, timeout=10)
                jdata = r.json()
                
                if jdata.get('success') == 'true' or jdata.get('success') is True:
                    eq_record = jdata['records']['Earthquake'][0]
                    eq_no = eq_record.get('EarthquakeNo')
                    info = eq_record['EarthquakeInfo']
                    
                    # 1. 轉型與取得當前時間（嚴格遵循 import datetime 格式）
                    origin_time_str = info['OriginTime'].split("+")[0].replace("T", " ")
                    eq_time = datetime.datetime.strptime(origin_time_str, "%Y-%m-%d %H:%M:%S")
                    current_time = datetime.datetime.now()
                    
                    # 計算時間差
                    time_diff = current_time - eq_time

                    # ⏳ 舊地震過濾：改為超過 48 小時（2天）或異常未來的報告直接略過
                    if time_diff > datetime.timedelta(hours=48) or time_diff < datetime.timedelta(seconds=0):
                        pass 
                    
                    # 2. 如果是 1~2 天內的新地震，且編號未重複，進入震度門檻檢查
                    elif eq_no != self.last_earthquake_no:
                        
                        # 讀取設定的最低觸發震度門檻（預設為 3）
                        try:
                            min_threshold = int(self.settings.get("min_alert_intensity", getattr(self, "min_alert_intensity", "3")))
                        except ValueError:
                            min_threshold = 3

                        # A. 提取居住地區的震度
                        local_intensity_str = "0"
                        shaking_areas = eq_record['Intensity']['ShakingArea']
                        for area in shaking_areas:
                            if self.target_county in area['CountyName'] or area['CountyName'] in self.target_county:
                                local_intensity_str = area['AreaIntensity']
                                break
                        
                        # B. 提取全國最大震度（當作居住地無感時的備用比對）
                        max_intensity_str = "0"
                        if 'AreaIntensity' in eq_record['Intensity']:
                            max_intensity_str = eq_record['Intensity']['AreaIntensity']
                        elif "最大震度" in eq_record['ReportContent']:
                            # 從報告內文中概抓最大震度字串
                            max_intensity_str = eq_record['ReportContent'].split("最大震度")[-1][:2]

                        # C. 將提取到的震度字串清洗為「單個純數字」
                        def extract_digit(intensity_str):
                            # 使用 re 尋找字串中的第一個數字 
                            match = re.search(r'[0-7]', str(intensity_str))
                            return int(match.group()) if match else 0

                        local_val = extract_digit(local_intensity_str)
                        max_val = extract_digit(max_intensity_str)

                        # D. 核心邏輯判斷：決定是否要觸發警報
                        should_alert = False
                        final_display_intensity = "3" # 預設顯示值

                        if local_val > 0:
                            # 居住地有感：若大於等於門檻，則觸發
                            if local_val >= min_threshold:
                                should_alert = True
                                final_display_intensity = str(local_val)
                        else:
                            # 居住地無感：若全國最大震度達到門檻，長者雖不會晃但可能受驚嚇，亦觸發提示
                            if max_val >= min_threshold:
                                should_alert = True
                                final_display_intensity = str(max_val)

                        # E. 通過門檻，更新歷史編號並送回主執行緒彈窗
                        if should_alert:
                            self.last_earthquake_no = eq_no  
                            
                            report_content = eq_record['ReportContent']
                            epicenter = info['Epicenter']['Location']
                            scale = f"芮氏規模 {info['EarthquakeMagnitude']['MagnitudeValue']}"

                            self.root.after(0, lambda: self.popup_earthquake_report_ui(
                                scale, final_display_intensity, origin_time_str, epicenter, report_content
                            ))
                        else:
                            # 未達門檻，僅過濾更新編號但不干擾長者
                            self.last_earthquake_no = eq_no
            
            except Exception as e:
                print(f"背景地震監測連線異常: {e}")
            time.sleep(10)

    def popup_earthquake_report_ui(self, scale, intensity, origin_time, epicenter, report_content):
        """去邊框、高齡無障礙大字體地震警告主視窗 (無閃爍版本)"""
        eq_win = tk.Toplevel(self.root)
        eq_win.title(self.localize_runtime_text("🚨 中央氣象署：最新地震有感報告 🚨"))
        eq_win.geometry("950x780")
        
        # 🎨 將純數字震度轉為 int 判斷分級（5 以上為高強度）
        try:
            int_level = int(intensity)
        except ValueError:
            int_level = 5
            
        is_high_alert = int_level >= 5
        
        # 使用溫和且固定的高對比配色，不進行動態閃爍
        base_bg = "#FEF2F2" if is_high_alert else "#FFFBEB"      # 紅底 vs 黃底
        text_color = "#991B1B" if is_high_alert else "#92400E"   # 紅字 vs 褐字
        beep_freq = 1500 if is_high_alert else 900               # 5 以上聲音更尖銳
        
        eq_win.configure(bg=base_bg) 
        eq_win.attributes("-topmost", True) 
        eq_win.grab_set() 

        # 語音播報整合
        def _speak():
            alert_msg = (
                f"Please pay attention. A felt earthquake just occurred. {report_content}. "
                "Please use the buttons to confirm that you are safe."
                if self.settings.get("language") == "English"
                else f"大哥大姐請注意，剛剛發生有感地震。{report_content}。請點擊畫面確認您是否安全。"
            )
            self.tts_speak(alert_msg)

        threading.Thread(target=_speak, daemon=True).start()

        # 🔊 背景連續警報音效執行緒（僅短暫提示，避免過度驚嚇）
        def _play_alarm():
            for _ in range(3): # 僅提示 3 次音效
                if not eq_win.winfo_exists(): break
                try:
                    winsound.Beep(beep_freq, 400) 
                    time.sleep(0.4) 
                except:
                    break

        threading.Thread(target=_play_alarm, daemon=True).start()

        # 介面大字體元件排列（靜態呈現，不閃爍）
        lbl_title = tk.Label(eq_win, text=self.localize_runtime_text("🚨 地震安全回報視窗 🚨"), font=("Microsoft JhengHei", 36, "bold"), bg=base_bg, fg=text_color, wraplength=850)
        lbl_title.pack(pady=20)

        intensity_text = f"Estimated intensity: [ {intensity} ]" if self.settings.get("language") == "English" else f"估計震度：【 {intensity} 】"
        lbl_intensity = tk.Label(eq_win, text=intensity_text, font=("Microsoft JhengHei", 38, "bold"), bg=base_bg, fg="#DC2626" if is_high_alert else "#EA580C")
        lbl_intensity.pack(pady=5)

        if self.settings.get("language") == "English":
            details_text = f"● Time: {origin_time}\n● Magnitude: {scale.replace('芮氏規模', 'Magnitude')}\n● Epicenter: {epicenter}"
        else:
            details_text = f"● 發生時間：{origin_time}\n● 地震規模：{scale}\n● 震央位置：{epicenter}"
        lbl_details = tk.Label(eq_win, text=details_text, font=("Microsoft JhengHei", 20, "bold"), bg=base_bg, fg=text_color, justify="left", anchor="w")
        lbl_details.pack(pady=10, padx=50)

        lbl_guide = tk.Label(eq_win, text=self.localize_runtime_text("💡 【安全確認】\n請檢查自身與周遭環境，點擊下方按鈕回報您的目前狀況。"), font=("Microsoft JhengHei", 18, "bold"), bg=base_bg, fg="#1E3A8A", wraplength=850)
        lbl_guide.pack(pady=10)

        # ========================================================
        # 🛡️ 是否安全的互動回報按鈕區（雙選項設計）
        # ========================================================
        btn_frame = tk.Frame(eq_win, bg=base_bg)
        btn_frame.pack(side="bottom", fill="x", padx=60, pady=30)

        def report_safe():
            """回報安全：關閉警報，並透過語音或系統日誌給予安心反饋"""
            self.play_system_sound("execute")
            self.tts_speak("太好了，已記錄您目前平安安全，請放心。")
            self.log_status("智慧醫護：長者已回報【目前安全】。")
            eq_win.destroy()

        def report_need_help():
            """回報需要協助：觸發警報提示並發送緊急通知給家屬"""
            self.play_system_sound("warn")
            self.tts_speak("已收到您的回報，系統正在通知家屬與醫護人員，請留在原地保持冷靜。")
            self.log_status("⚠️ 警告：長者回報【需要協助】！")
            
            self.send_email_notification(
                datetime.date.today().strftime("%Y-%m-%d"), 
                datetime.datetime.now().strftime("%H:%M:%S"), 
                self.settings.get("streak_days", 1)
            )
            eq_win.destroy()

        # 按鈕 1：我很安全
        btn_safe = tk.Button(
            btn_frame, text=self.localize_runtime_text("👍 我很安全 ｜ 點此回報平安"), font=("Microsoft JhengHei", 20, "bold"), 
            bg="#16A34A", fg="white", activebackground="#15803D", activeforeground="white",
            bd=0, relief="flat", cursor="hand2", command=report_safe
        )
        btn_safe.pack(side="top", fill="x", ipady=12, pady=10)

        # 按鈕 2：需要協助
        btn_help = tk.Button(
            btn_frame, text=self.localize_runtime_text("🚨 我需要協助 ｜ 通知家屬"), font=("Microsoft JhengHei", 20, "bold"), 
            bg="#DC2626", fg="white", activebackground="#B91C1C", activeforeground="white",
            bd=0, relief="flat", cursor="hand2", command=report_need_help
        )
        btn_help.pack(side="top", fill="x", ipady=12, pady=5)
        # ========================================================
    def generate_email_verification_code(self):
        return str(random.randint(100000, 999999))

    def send_verification_code_email(self, target_email, verification_code):
        sender_email = "smartcaremultimediasys@gmail.com"
        sender_password = "hugj atxd usci zpts"
        target_email = str(target_email or "").strip()

        if not target_email or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", target_email):
            return False, "請輸入有效的家屬收件信箱。"

        subject = "【智慧照護系統】家屬收件信箱驗證碼"
        body = (
            f"您好，\n\n"
            f"您的智慧照護系統驗證碼為：{verification_code}\n\n"
            f"此驗證碼 5 分鐘內有效，請在系統中輸入完成驗證。\n\n"
            f"若您未曾進行此設定，請忽略此郵件。"
        )

        msg = MIMEText(body, 'plain', 'utf-8')
        msg['From'] = Header(f"智慧照護系統助理 <{sender_email}>", 'utf-8')
        msg['To'] = Header(target_email, 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, [target_email], msg.as_string())
            server.quit()
            print(f"[Email Verify] 驗證碼已成功寄送至: {target_email}")
            return True, "驗證碼已寄出，請至信箱查看並輸入驗證碼。"
        except Exception as e:
            print(f"[Email Verify] 發送驗證碼失敗: {e}")
            return False, f"驗證碼發送失敗：{e}"

    def send_email_notification(self, today_str, current_time, streak_days=1, is_sos=False, alert_title=None, alert_details=None):
        """發送 Email 通知（當 is_sos=True 時必發 Email）"""

        # 若非 SOS 且家屬通知開關未開啟，則略過
        if not is_sos and not self.settings.get("email_enabled", True):
            print("ℹ️ 照護通知未開啟：根據家屬設定，不發送 Email。")
            return

        sender_email = str(self.settings.get("gmail_sender_email", "smartcaremultimediasys@gmail.com")).strip() or "smartcaremultimediasys@gmail.com"
        sender_password = str(self.settings.get("gmail_app_password", "hugj atxd usci zpts")).strip() or "hugj atxd usci zpts"
        receiver_email = str(self.settings.get("receiver_email", "")).strip()

        if not receiver_email:
            print("❌ 發送失敗：設定的家屬 Email 地址不可為空！")
            return

        def _send():
            if is_sos:
                if alert_title:
                    subject = f"🚨【{alert_title}】長輩緊急狀況通知 ({today_str} {current_time})"
                else:
                    subject = f"🚨【緊急求救 SOS】長輩發出緊急求救訊號！ ({today_str} {current_time})"

                body = (
                    f"⚠️ 緊急通知：\n\n"
                    f"{alert_details or '系統偵測到長輩發出緊急求救訊號。'}\n"
                    f"--------------------------------------------------\n"
                    f"📅 事件日期：{today_str}\n"
                    f"⏰ 事件時間：{current_time}\n"
                    f"--------------------------------------------------\n"
                    f"請儘快確認長輩安全狀況！"
                )
            else:
                subject = f"☀️【智慧照護日誌】長輩今日已完成健康簽到 ({today_str})"
                body = (
                    f"親愛的家屬您好：\n\n"
                    f"智慧多媒體照護系統偵測到長輩已安全起床，並順利完成晨間點名！\n"
                    f"--------------------------------------------------\n"
                    f"📅 簽到日期：{today_str}\n"
                    f"⏰ 簽到時間：{current_time}\n"
                    f"🔥 連續健康簽到：{streak_days} 天\n"
                    f"--------------------------------------------------\n"
                    f"目前長輩精神狀態良好，請您放心。"
                )

            msg = MIMEText(body, 'plain', 'utf-8')
            msg['From'] = Header(f"智慧照護系統助理 <{sender_email}>", 'utf-8')
            msg['To'] = Header(receiver_email, 'utf-8')
            msg['Subject'] = Header(subject, 'utf-8')

            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, [receiver_email], msg.as_string())
                server.quit()
                print(f"[{'SOS緊急' if is_sos else '日常'}] Email 已成功發送至: {receiver_email}")
            except Exception as e:
                print(f"Email 發送失敗: {e}")

        threading.Thread(target=_send, daemon=True).start()

    def check_daily_attendance_on_boot(self, force_open=False):
        """開機自動檢查今日是否已簽到（修正版：Unix時間戳記 + 7天視覺化日曆，完全去邊框設計）"""
        
        # 取得今天的日期物件與當前 Unix 時間戳記
        now_timestamp = datetime.datetime.now().timestamp()
        today = datetime.date.today()
        today_str = today.strftime("%Y-%m-%d")
        
        # 核心防禦：防止多個簽到視窗同時重複跳出
        if not hasattr(self, "attendance_window"):
            self.attendance_window = None

        if self.attendance_window is not None and self.attendance_window.winfo_exists():
            print("⚠️ 防禦機制攔截：簽到視窗已在畫面上，忽略重複觸發請求。")
            try:
                self.attendance_window.lift()
                self.attendance_window.focus_set()
            except:
                pass
            return

        # 從現有設定檔中讀取歷史簽到紀錄
        last_sign_in = self.settings.get("last_sign_in_date", 0)
        try:
            streak = max(0, int(self.settings.get("streak_days", 0)))
        except (TypeError, ValueError):
            streak = 0
        try:
            total_count = max(0, int(self.settings.get("total_sign_in_count", 0)))
        except (TypeError, ValueError):
            total_count = 0

        # 只保留有效日期，避免不完整或舊版設定造成日曆判定失敗。
        raw_history_dates = self.settings.get("history_dates", [])
        history_dates = []
        if isinstance(raw_history_dates, list):
            for value in raw_history_dates:
                if not isinstance(value, str):
                    continue
                try:
                    parsed_date = datetime.datetime.strptime(value, "%Y-%m-%d").date()
                except ValueError:
                    continue
                if parsed_date.strftime("%Y-%m-%d") == value:
                    history_dates.append(value)
        history_dates = sorted(set(history_dates))

        last_date_obj = None
        if isinstance(last_sign_in, (int, float)) and last_sign_in > 0:
            # 數值型態：直接從 Unix 時間戳記轉回日期物件
            try:
                last_date_obj = datetime.date.fromtimestamp(last_sign_in)
            except (OverflowError, OSError, ValueError):
                last_date_obj = None
        elif isinstance(last_sign_in, str) and last_sign_in.strip():
            # 相容舊版：若是字串則解析回來
            try:
                last_date_obj = datetime.datetime.strptime(last_sign_in, "%Y-%m-%d").date()
            except ValueError:
                last_date_obj = None

        # 若上次簽到欄位未同步，使用歷史清單中最近且不晚於今天的日期補回。
        history_date_objs = [
            datetime.datetime.strptime(value, "%Y-%m-%d").date()
            for value in history_dates
        ]
        latest_history_date = max(
            (value for value in history_date_objs if value <= today),
            default=None,
        )
        settings_need_sync = history_dates != raw_history_dates
        if latest_history_date is not None and (
            last_date_obj is None or latest_history_date > last_date_obj
        ):
            last_date_obj = latest_history_date
            self.settings["last_sign_in_date"] = datetime.datetime.combine(
                last_date_obj, datetime.time.min
            ).timestamp()
            settings_need_sync = True

        # 計算天數差
        if last_date_obj is not None:
            days_diff = (today - last_date_obj).days
            is_already_signed_in = (days_diff == 0 or today_str in history_dates)
        else:
            days_diff = -1
            is_already_signed_in = today_str in history_dates

        if is_already_signed_in and today_str not in history_dates:
            history_dates.append(today_str)
            self.settings["history_dates"] = history_dates
            settings_need_sync = True

        # 進度只能延續到今天或昨天；中斷時先同步歸零，待今日簽到後再從 1 開始。
        if not is_already_signed_in and days_diff != 1 and streak != 0:
            streak = 0
            self.settings["streak_days"] = streak
            settings_need_sync = True

        if self.settings.get("history_dates") != history_dates:
            self.settings["history_dates"] = history_dates
            settings_need_sync = True
        if settings_need_sync:
            self.save_settings(self.settings)

        # 如果不是手動強制打開，且今天已經自動簽到過了，就不彈出視窗
        if not force_open and is_already_signed_in:
            self.log_status(f"今日已完成簽到！累計連續簽到 {streak} 天。")
            return

        # ========================================================
        # 🔑 Unix 核心天數差判定（0天=重複開啟, 1天=連續, >1天=中斷重置）
        # ========================================================
        if is_already_signed_in:
            new_streak = streak
            streak_msg = f"您今天已經完成點名囉！目前連續健康天數：{streak} 天。🌟"
        elif days_diff == 1:
            # 剛好差 1 天，代表昨天有簽到
            new_streak = streak + 1
            streak_msg = f"太棒了！您已經連續簽到 {new_streak} 天囉！🌟"
        else:
            # 天數差大於 1 天（中斷）或無歷史紀錄
            new_streak = 1
            if last_date_obj is not None:
                streak_msg = "偵測到中斷簽到，別氣餒！讓我們重新開始累積健康天數吧！💪"
            else:
                streak_msg = "歡迎第一次使用！讓我們開始累積健康天數吧！💪"

        self.log_status("正在查看今日健康簽到狀態...")
        self.play_system_sound("execute")
        
        if is_already_signed_in:
            self.tts_speak("您今天已經完成健康點名簽到囉！")
        else:
            self.tts_speak("早安！請點擊畫面上的大愛心完成今日簽到！")

        # 建立簽到/紀錄視窗 (拉高視窗容納日曆)
        boot_win = tk.Toplevel(self.root)
        self.attendance_window = boot_win 
        
        boot_win.title("每日健康點點名")
        boot_win.geometry("850x820")
        boot_win.configure(bg="#FFFBEB") 
        
        tk.Label(
            boot_win, text="✨ 每日健康點點名 ✨",
            font=("Microsoft JhengHei", 32, "bold"), bg="#FFFBEB", fg="#B45309"
        ).pack(pady=15)

        # 1. 歷史紀錄看板 (同樣配合去邊框化，改用 flat 樣式)
        info_frame = tk.Frame(boot_win, bg="#FEF3C7", bd=0, relief="flat", padx=25, pady=10)
        info_frame.pack(fill="x", padx=60)

        display_last_date = last_date_obj.strftime("%Y-%m-%d") if last_date_obj else "無歷史紀錄（第一次簽到）"
        tk.Label(
            info_frame, text=f"📅 上次簽到日期：{display_last_date}", 
            font=("Microsoft JhengHei", 16), bg="#FEF3C7", fg="#78350F"
        ).pack(anchor="w", pady=2)
        
        tk.Label(
            info_frame, text=f"🔥 目前連續簽到：{streak} 天", 
            font=("Microsoft JhengHei", 16), bg="#FEF3C7", fg="#78350F"
        ).pack(anchor="w", pady=2)

        # ========================================================
        # 📅 🔑 視覺化「健康日曆看板」區域 (✨ 徹底去除邊框平滑設計)
        # ========================================================
        calendar_frame = tk.Frame(boot_win, bg="#FFFBEB", bd=0, relief="flat")
        calendar_frame.pack(fill="x", padx=60, pady=10)
        
        tk.Label(
            calendar_frame, text="📊 最近七天健康紀錄看板：", 
            font=("Microsoft JhengHei", 14, "bold"), bg="#FFFBEB", fg="#9A3412"
        ).pack(anchor="w", pady=5)
        
        days_box_frame = tk.Frame(calendar_frame, bg="#FFFBEB")
        days_box_frame.pack(fill="x")
        
        week_days_ch = ["一", "二", "三", "四", "五", "六", "日"]
        calendar_boxes = {} 

        # 繪製過去 6 天到今天的日曆格子
        for i in range(6, -1, -1):
            target_date = today - datetime.timedelta(days=i)
            target_date_str = target_date.strftime("%Y-%m-%d")
            
            # --- 🛡️ 核心修正：多重防禦判定 ---
            # 1. 檢查是否在歷史清單中
            in_history = target_date_str in history_dates
            
            # 2. 檢查今天是否已簽到
            is_today_signed = (target_date_str == today_str and is_already_signed_in)
            
            # 3. ✨ 保險回推機制：如果今天還沒簽到(is_already_signed_in為False)，但目前 streak >= 1
            #    代表從「昨天」倒推回去的 streak 天數內，都應該是已簽到狀態！
            is_in_streak = False
            if not is_already_signed_in and streak > 0:
                # 昨天是倒推 1 天，前天是 2 天... 直到包含昨天的連續天數
                if 1 <= (today - target_date).days <= streak:
                    is_in_streak = True
            elif is_already_signed_in and streak > 0:
                # 如果今天已經簽到了，代表從「今天」倒推回去的 streak 天數內都是已簽到
                if 0 <= (today - target_date).days < streak:
                    is_in_streak = True

            # 三個條件只要有一個成立，就代表有簽到！
            has_signed = in_history or is_today_signed or is_in_streak
            # ----------------------------------
            
            # 採用高對比扁平柔和彩塊
            bg_box = "#D1FAE5" if has_signed else "#E5E7EB" 
            fg_box = "#065F46" if has_signed else "#9CA3AF"
            
            box = tk.Frame(days_box_frame, bg=bg_box, bd=0, relief="flat")
            box.pack(side="left", expand=True, fill="both", padx=6, ipady=8) 
            
            lbl_date = tk.Label(
                box, text=f"周{week_days_ch[target_date.weekday()]}\n{target_date.month}/{target_date.day}",
                font=("Microsoft JhengHei", 12, "bold"), bg=bg_box, fg="#374151"
            )
            lbl_date.pack(pady=2)
            
            status_symbol = "♥" if has_signed else "·"
            lbl_status = tk.Label(
                box, text=status_symbol, font=("Microsoft JhengHei", 20, "bold"), bg=bg_box, fg=fg_box
            )
            lbl_status.pack(pady=2)
            
            # 儲存今天這一格的指標，用來做點擊後的即時動態刷新
            if target_date_str == today_str:
                calendar_boxes["frame"] = box
                calendar_boxes["lbl_date"] = lbl_date
                calendar_boxes["lbl_status"] = lbl_status
        # ========================================================

        # 宣告暫不簽到按鈕的全域變數指針，方便在成功時跨函式將它銷毀
        btn_skip_ref = [None]

        def close_window():
            self.play_system_sound("click")
            self.attendance_window = None 
            boot_win.destroy()       

        boot_win.protocol("WM_DELETE_WINDOW", close_window)

        # 互動操作主要容器
        action_frame = tk.Frame(boot_win, bg="#FFFBEB")
        action_frame.pack(fill="both", expand=True, pady=5, padx=60)

        # 核心互動與成功邏輯
        def do_sign_in():
            self.play_system_sound("execute")
            
            # 🛡️ 確保從實體 settings 安全同步
            if "history_dates" not in self.settings:
                self.settings["history_dates"] = []
            
            # 更新歷史日期清單
            if today_str not in self.settings["history_dates"]:
                self.settings["history_dates"].append(today_str)

            # ================= 🧹 7日紀錄舊資料清理邏輯 =================
            # 計算出 7 天前那天的日期字串 (格式: YYYY-MM-DD)
            seven_days_ago = (datetime.datetime.now() - datetime.timedelta(days=6)).strftime("%Y-%m-%d")
            
            # 💡 核心修正：只保留大於或等於 seven_days_ago 的日期，自動刪除舊紀錄
            self.settings["history_dates"] = [
                d for d in self.settings["history_dates"] 
                if d >= seven_days_ago
            ]
            # ==========================================================

            # 將 last_sign_in_date 改寫入 Unix 時間戳記 (now_timestamp)
            self.settings["last_sign_in_date"] = now_timestamp
            self.settings["streak_days"] = new_streak
            self.settings["total_sign_in_count"] = total_count + 1
            
            # 強制調用專屬的儲存函式，確保實體 json 檔案立刻寫入並同步更新
            self.save_settings(self.settings)

            # 【即時變色】讓日曆上「今天」那一格動態亮起大紅心
            if "frame" in calendar_boxes:
                calendar_boxes["frame"].configure(bg="#D1FAE5")
                calendar_boxes["lbl_date"].configure(bg="#D1FAE5")
                calendar_boxes["lbl_status"].configure(bg="#D1FAE5", text="♥", fg="#065F46")

            # 發送 Email 遠端照護通知
            current_time = time.strftime("%H:%M:%S")
            self.send_email_notification(today_str, current_time, new_streak)

            # 重新整理畫面顯示成功畫面
            show_success_ui()

        def show_success_ui():
            # 1. 🛠️ 關鍵修正：如果底部存有「暫不簽到」按鈕，立刻將它徹底從視窗上消滅
            if btn_skip_ref[0] is not None and btn_skip_ref[0].winfo_exists():
                btn_skip_ref[0].destroy()

            # 清空內部組件
            for widget in action_frame.winfo_children():
                widget.destroy()

            tk.Label(
                action_frame, text=f"🎉 今日已完成健康簽到！", 
                font=("Microsoft JhengHei", 24, "bold"), bg="#FFFBEB", fg="#16A34A"
            ).pack(pady=10)

            # 🛡️ 修正點 3：動態計算最新修復的連續簽到訊息文字，避免變數卡在未簽到前的狀態
            dynamic_streak_msg = f"太棒了！您已經連續簽到 {new_streak} 天囉！🌟" if new_streak > 1 else streak_msg
            tk.Label(
                action_frame, text=dynamic_streak_msg, 
                font=("Microsoft JhengHei", 18, "bold"), bg="#FFFBEB", fg="#1E3A8A"
            ).pack(pady=5)

            msg_ai = tk.Message(
                action_frame, text="護理師正在為您準備今日健康話語...", 
                font=("Microsoft JhengHei", 16, "bold"), bg="#FFFBEB", fg="#4B5563",
                width=700, aspect=400, justify="center"
            )
            msg_ai.pack(pady=10, fill="x")

            # 按鈕也一併扁平化處理 (bd=0, relief="flat")
            btn_close = tk.Button(
                action_frame, text="👍 我知道了，回首頁", font=("Microsoft JhengHei", 20, "bold"),
                bg="#10B981", fg="white", activebackground="#059669", activeforeground="white",
                bd=0, relief="flat", cursor="hand2", command=close_window
            )
            btn_close.pack(side="bottom", pady=10, ipady=5, fill="x", padx=100)

            def run_ai_greeting():
                lang = self.settings.get("language", "國語 (華語)")
                lang_dict = I18N.get(lang, I18N["國語 (華語)"])
                
                # 從字典取得動態語言格式化的 Prompt [修改]
                prompt_template = lang_dict.get(
                    "gemini_greeting_prompt", 
                    I18N["國語 (華語)"]["gemini_greeting_prompt"]
                )
                ai_query = prompt_template.format(streak=new_streak)

                try:
                    response = self.gemini_chat.send_message(message=ai_query)
                    answer = response.text
                    if boot_win.winfo_exists():
                        self.transition_text(
                            msg_ai,
                            f"護理師：\n{answer}",
                            configure=lambda: msg_ai.configure(fg="#1E3A8A"),
                        )
                        self.tts_speak(answer)
                except Exception as e:
                    print(f"Gemini 簽到問候失敗: {e}")
            
            threading.Thread(target=run_ai_greeting, daemon=True).start()

        # 分流渲染控制
        if is_already_signed_in:
            show_success_ui()
        else:
            # 未簽到狀態下：顯示愛心按鈕 (主動按鈕可保留 raised 立體感，或者需要的話也可以改成 flat 0 邊框)
            btn_heart = tk.Button(
                action_frame, text="♥\n點我點名簽到", font=("Microsoft JhengHei", 28, "bold"),
                bg="#EF4444", fg="white", activebackground="#DC2626", activeforeground="white",
                bd=0, relief="flat", cursor="hand2", command=do_sign_in
            )
            btn_heart.pack(expand=True, fill="both", padx=80, pady=10)

            # 未簽到狀態下：建立底部的「暫不簽到」按鈕 (去邊框扁平樣式)
            btn_skip = tk.Button(
                boot_win, text="❌ 暫不簽到，稍後再說", font=("Microsoft JhengHei", 14, "bold"),
                bg="#9CA3AF", fg="white", activebackground="#71717A", activeforeground="white",
                bd=0, relief="flat", cursor="hand2", command=close_window
            )
            btn_skip.pack(side="bottom", pady=15)
            btn_skip_ref[0] = btn_skip # 🔑 儲存指標

    def create_grid_ui(self):
        """建立符合高齡無障礙標準的【左上角樣式切換鈕 + 互斥式排版（立體九宮格 vs 獨佔左右媒體輪播看板，去邊框）】"""
        BG_COLOR = "#EBF8FF"
        
        # 程式預設為「純九宮格模式」
        if not hasattr(self, 'current_style_mode'):
            self.current_style_mode = "精簡模式 (純純九宮格)"
        
        # 主容器包覆整區
        self.main_frame = tk.Frame(self.root, bg=BG_COLOR, padx=40, pady=15)
        self.main_frame.pack(fill="both", expand=True)

        # 配置列與欄權重 (主畫面維持 3 欄，剛好對應 媒體庫的 左按鈕、中卡片、右按鈕)
        for i in range(3):
            self.main_frame.columnconfigure(i, weight=1, uniform="button_grid")
            
        # 限制特定控制列不參與均分計算
        self.main_frame.rowconfigure(0, weight=0, uniform="") # 頂部左鈕工具列
        self.main_frame.rowconfigure(1, weight=0, uniform="") # 網路狀態列
        self.main_frame.rowconfigure(2, weight=0, uniform="") # 時間列
        self.main_frame.rowconfigure(3, weight=0, uniform="") # 主標題列
        self.main_frame.rowconfigure(4, weight=0, uniform="") # 樣式狀態標籤列
        
        # 讓主容器的 row 5 填滿中腹（九宮格與媒體庫的共用渲染槽）
        self.main_frame.rowconfigure(5, weight=1, uniform="")

        self.main_frame.columnconfigure(0, weight=0) 
        self.main_frame.columnconfigure(1, weight=1)
        # ========================================================
        # 📦 統一管理系統「所有功能內容」資料集 (共 9 個功能)
        # ========================================================
        self.all_system_features = [
            {"text": "做復健運動\n(觀看影片)", "color": "#EF4444", "bg_light": "#FEE2E2", "fg_dark": "#991B1B", "cmd": self.show_rehab_video_picker},
            {"text": "手指健康操\n(鏡頭手勢挑戰)", "color": "#F97316", "bg_light": "#FFEDD5", "fg_dark": "#9A3412", "cmd": self.trigger_finger_gesture_game},
            {"text": "看懷舊風景\n(放鬆心情)", "color": "#3B82F6", "bg_light": "#DBEAFE", "fg_dark": "#1E40AF", "cmd": lambda: self.execute_medical_search("台灣懷舊風景 4K 放鬆影片")},
            {"text": "衛教健康網\n(健康查詢)", "color": "#10B981", "bg_light": "#D1FAE5", "fg_dark": "#065F46", "cmd": lambda: self.open_link("https://www.hpa.gov.tw/Pages/List.aspx?nodeid=366")},
            {"text": "數字記憶挑戰\n(腦力鍛鍊)", "color": "#8B5CF6", "bg_light": "#EDE9FE", "fg_dark": "#5B21B6", "cmd": self.trigger_chimp_game},
            {"text": "AI 照護助理\n(對話聊天)", "color": "#EC4899", "bg_light": "#FCE7F3", "fg_dark": "#9D174D", "cmd": self.trigger_ai_chat},
            {"text": "系統參數\n(設定選單)", "color": "#A5A5A5", "bg_light": "#F3F4F6", "fg_dark": "#374151", "cmd": self.open_settings_window},
            {"text": "相機視窗\n(即時畫面)", "color": "#06B6D4", "bg_light": "#E0F7FA", "fg_dark": "#006064", "cmd": self.show_camera_window},
            {"text": "關閉程式\n(退出系統)", "color": "#3A3A3A", "bg_light": "#E5E7EB", "fg_dark": "#111827", "cmd": self.safe_exit}
        ]
        self.current_media_index = 0  # 紀錄當前切換到第幾個功能
        # --------------------------------------------------------
        # 📱 容器 A：九宮格獨立容器
        # --------------------------------------------------------
        self.grid_container_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        for i in range(3):
            self.grid_container_frame.columnconfigure(i, weight=1, uniform="sub_grid")
            self.grid_container_frame.rowconfigure(i, weight=1, uniform="sub_grid")

        # 動態渲染九宮格按鈕
        self.grid_feature_buttons = []
        for idx, item in enumerate(self.all_system_features):
            r, c = idx // 3, idx % 3
            btn = tk.Button(
                self.grid_container_frame, text=item["text"], font=("Microsoft JhengHei", 20, "bold"),
                bg=item["color"], fg="white", activebackground=item["color"], activeforeground="white",
                bd=4, relief="raised", command=item["cmd"], cursor="hand2",
                justify="center", wraplength=320
            )
            btn.grid(row=r, column=c, padx=12, pady=8, sticky="nsew")
            btn.bind("<Configure>", self._fit_grid_feature_text, add="+")
            self.grid_feature_buttons.append(btn)

        # --------------------------------------------------------
        # 🎬 ✨ 容器 B：左右媒體庫容器 (左中右排版、每次只顯示一項、去邊框)
        # --------------------------------------------------------
        # 外框完全去邊框 bd=0, relief="flat"
        self.gallery_frame = tk.Frame(self.main_frame, bg=BG_COLOR, bd=0, relief="flat")
        
        # 分配 3 欄權重：左按鈕(25%)、中間卡片(50%)、右按鈕(25%)
        self.gallery_frame.columnconfigure(0, weight=1, uniform="gallery_cols")
        self.gallery_frame.columnconfigure(1, weight=2, uniform="gallery_cols")
        self.gallery_frame.columnconfigure(2, weight=1, uniform="gallery_cols")
        self.gallery_frame.rowconfigure(0, weight=1, uniform="gallery_rows")

        # 左側切換按鈕 ◀ (去邊框扁平化設計)
        self.btn_media_prev = tk.Button(
            self.gallery_frame, text="◀ 上一個功能", font=("Microsoft JhengHei", 20, "bold"),
            bg="#FFFFFF", fg="#1E293B", activebackground="#E2E8F0", activeforeground="#1E293B",
            bd=0, relief="flat", command=self.prev_media_item, cursor="hand2"
        )
        self.btn_media_prev.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        # 中間主體功能卡片 (去邊框扁平化設計，點擊直接觸發對應功能)
        self.btn_media_card = tk.Button(
            self.gallery_frame, text="", font=("Microsoft JhengHei", 26, "bold"),
            bd=0, relief="flat", cursor="hand2", command=self.click_media_card,
            wraplength=430, justify="center", anchor="center"
        )
        self.btn_media_card.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        # 右側切換按鈕 ▶ (去邊框扁平化設計)
        self.btn_media_next = tk.Button(
            self.gallery_frame, text="下一個功能 ▶", font=("Microsoft JhengHei", 20, "bold"),
            bg="#FFFFFF", fg="#1E293B", activebackground="#E2E8F0", activeforeground="#1E293B",
            bd=0, relief="flat", command=self.next_media_item, cursor="hand2"
        )
        self.btn_media_next.grid(row=0, column=2, padx=20, pady=20, sticky="nsew")

        # 側邊按鈕移入微特效
        def side_btn_enter(event, b): b.configure(bg="#E2E8F0")
        def side_btn_leave(event, b): b.configure(bg="#FFFFFF")
        self.btn_media_prev.bind("<Enter>", lambda e, b=self.btn_media_prev: side_btn_enter(e, b))
        self.btn_media_prev.bind("<Leave>", lambda e, b=self.btn_media_prev: side_btn_leave(e, b))
        self.btn_media_next.bind("<Enter>", lambda e, b=self.btn_media_next: side_btn_enter(e, b))
        self.btn_media_next.bind("<Leave>", lambda e, b=self.btn_media_next: side_btn_leave(e, b))

        # ========================================================
        # 🛠️ 【左上角獨立按鈕切換區】
        # ========================================================
        top_bar_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        top_bar_frame.grid(row=0, column=0, columnspan=3, sticky="ew", padx=15, pady=5)

        # 1a. 左側按鈕區
        top_left_frame = tk.Frame(top_bar_frame, bg=BG_COLOR)
        top_left_frame.pack(side="left")

        self.btn_switch_grid = tk.Button(
            top_left_frame, text="📱 顯示九宮格清單", font=("Microsoft JhengHei", 12, "bold"),
            bg="#10B981", fg="white", activebackground="#059669", activeforeground="white",
            bd=2, relief="raised", cursor="hand2",
            command=lambda: self.set_ui_style_mode("精簡模式 (純純九宮格)")
        )
        self.btn_switch_grid.pack(side="left", padx=(0, 10))

        self.btn_switch_media = tk.Button(
            top_left_frame, text="🎬 顯示左右媒體庫", font=("Microsoft JhengHei", 12, "bold"),
            bg="#3B82F6", fg="white", activebackground="#2563EB", activeforeground="white",
            bd=2, relief="raised", cursor="hand2",
            command=lambda: self.set_ui_style_mode("焦點媒體模式 (專注看影片)")
        )
        self.btn_switch_media.pack(side="left")

        # 1b. 右側按鈕區
        right_btn_frame = tk.Frame(top_bar_frame, bg=BG_COLOR)
        right_btn_frame.pack(side="right")

        self.btn_sign = tk.Button(
            right_btn_frame, text="📋 補簽到/查看紀錄", font=("Microsoft JhengHei", 12, "bold"),
            bg="#F39C12", fg="white", activebackground="#D68910", activeforeground="white",
            bd=2, relief="raised", cursor="hand2",
            command=lambda: self.check_daily_attendance_on_boot(force_open=True)
        )
        self.btn_sign.pack(side="right", padx=(5, 0))

        self.mic_button = tk.Button(
            right_btn_frame, text="🎤 點我說話 (關閉中)", font=("Microsoft JhengHei", 12, "bold"),
            bg="#5D6D7E", fg="white", activebackground="#34495E", activeforeground="white",
            bd=2, relief="raised", cursor="hand2",
            command=self.toggle_microphone
        )
        self.mic_button.pack(side="right", padx=5)

        self.network_status_font = tkfont.Font(
            root=self.root,
            family="Microsoft JhengHei",
            size=14,
            weight="bold",
        )
        self.network_status_label = tk.Label(
            self.main_frame,
            text="🌐 裝置網路：未連接",
            font=self.network_status_font,
            padx=12,
            pady=5,
            bd=1,
            relief="solid",
        )
        self.network_status_label.grid(
            row=1, column=0, columnspan=3, pady=(4, 0), sticky="n"
        )
        self.main_frame.bind(
            "<Configure>", self._schedule_network_status_fit, add="+"
        )
        self.start_network_status_monitor()

        # ========================================================
        # 1. 當日日期與時間的標籤 (維持原本的 grid)
        # ========================================================
        self.time_label = tk.Label(
            self.main_frame, text="", 
            font=("Microsoft JhengHei", 18, "bold"), bg=BG_COLOR, fg="#1E3A8A"
        )
        self.time_label.grid(row=2, column=0, columnspan=3, pady=(5, 5), sticky="n")
        self.update_live_time()

        # 2. 主標題
        self.title_label = tk.Label(
            self.main_frame, text="您好！今天想做點什麼呢？", 
            font=("Microsoft JhengHei", 28, "bold"), bg=BG_COLOR, fg="#1E293B",
            wraplength=1280, justify="center"
        )
        self.title_label.grid(row=3, column=0, columnspan=3, pady=(0, 5), sticky="n")

        # 當前樣式模式顯示標籤
        self.mode_status_label = tk.Label(
            self.main_frame, text=f"📊 當前顯示樣式：{self.current_style_mode}", 
            font=("Microsoft JhengHei", 14, "bold"), bg="#DBEAFE", fg="#1E40AF",
            padx=15, pady=5, bd=1, relief="solid"
        )
        self.mode_status_label.grid(row=4, column=0, columnspan=3, pady=(0, 10), sticky="n")

        # 初始化更新媒體庫單張字卡的顯示
        self.update_media_gallery_display()

        # 根據預設值「純九宮格模式」調整初始排版布局
        self.apply_current_style_layout()

        # 底部狀態列
        self.status_bar = tk.Label(
            self.root, 
            text="系統狀態：準備就緒。將右手食指移到按鈕上不動即可點選。",
            font=("Microsoft JhengHei", 16, "bold"), 
            bg="#DBEAFE", 
            fg="#1E3A8A", 
            anchor="w", 
            padx=20, 
            pady=10
        )
        self.status_bar.pack(fill="x", side="bottom")

    def _schedule_network_status_fit(self, event=None):
        if getattr(self, "_network_status_fit_after_id", None) is None:
            self._network_status_fit_after_id = self.root.after_idle(
                self._fit_network_status_label
            )

    def _fit_network_status_label(self):
        self._network_status_fit_after_id = None
        if not hasattr(self, "network_status_label"):
            return
        if not self.network_status_label.winfo_exists():
            return

        available_width = self.main_frame.winfo_width() - 80
        if available_width <= 1:
            return

        base_size = max(
            1, round(14 * self._font_size_multiplier())
        )
        font_size = base_size
        self.network_status_font.configure(size=font_size)
        horizontal_padding = (
            2 * int(self.network_status_label.cget("padx"))
            + 2 * int(self.network_status_label.cget("bd"))
        )
        max_text_width = max(1, available_width - horizontal_padding)
        while (
            font_size > 8
            and self.network_status_font.measure(
                self.network_status_label.cget("text")
            ) > max_text_width
        ):
            font_size -= 1
            self.network_status_font.configure(size=font_size)
        self.network_status_label.configure(
            font=self.network_status_font,
            wraplength=max_text_width,
        )

    def start_network_status_monitor(self):
        self.network_connected = False
        self.network_status_queue = queue.Queue()
        self.update_network_status_label()
        self._check_network_status()
        self.root.after(500, self._poll_network_status)

    def update_network_status_label(self):
        lang_dict = I18N.get(
            self.settings.get("language"), I18N["國語 (華語)"]
        )
        state_key = (
            "network_status_connected"
            if self.network_connected
            else "network_status_disconnected"
        )
        status_text = lang_dict.get(
            state_key,
            "已連線" if self.network_connected else "未連接",
        )
        prefix = lang_dict.get("network_status_prefix", "🌐 裝置網路：")
        background, foreground = (
            ("#DCFCE7", "#008000")
            if self.network_connected
            else ("#FEE2E2", "#FF0000")
        )
        self.network_status_label.configure(
            text=f"{prefix}{status_text}",
            bg=background,
            fg=foreground,
        )
        self._schedule_network_status_fit()

    def _check_network_status(self):
        def _probe():
            try:
                response = requests.get(
                    "https://www.google.com/generate_204",
                    timeout=4,
                )
                connected = response.status_code < 500
            except requests.RequestException:
                connected = False
            self.network_status_queue.put(connected)

        threading.Thread(target=_probe, daemon=True).start()
        self.root.after(30000, self._check_network_status)

    def _poll_network_status(self):
        try:
            self.network_connected = self.network_status_queue.get_nowait()
            self.update_network_status_label()
        except queue.Empty:
            pass

        if self.root.winfo_exists():
            self.root.after(500, self._poll_network_status)

    # ========================================================
    # 🔄 樣式與排版互斥切換邏輯
    # ========================================================
    def set_ui_style_mode(self, target_mode):
        """直接指定切換至特定的顯示樣式，並寫入設定檔記憶"""
        self.play_system_sound("click")
        self.current_style_mode = target_mode
        prefix = I18N.get(self.settings.get("language"), {}).get("mode_status_prefix", "📊 當前顯示樣式：")
        self.transition_text(
            self.mode_status_label,
            f"{prefix}{self.localize_runtime_text(self.current_style_mode)}",
        )
        
        # ✨ 修正：將當前的設定即時儲存至設定字典並存檔
        self.settings["ui_style_mode"] = self.current_style_mode
        self.save_settings(self.settings)
            
        self.apply_current_style_layout()
        self.tts_speak(f"主畫面已切換為：{self.current_style_mode}")

    def apply_current_style_layout(self):
        """互斥切換：顯示媒體庫時不顯示九宮格"""
        if self.current_style_mode == "焦點媒體模式 (專注看影片)":
            self.grid_container_frame.grid_remove() # 隱藏九宮格
            self.gallery_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=10) # 顯示左右媒體庫
        else:
            self.gallery_frame.grid_remove() # 隱藏左右媒體庫
            self.grid_container_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=10) # 顯示九宮格

    # ========================================================
    # 🎬 左右媒體庫交互控制核心
    # ========================================================
    def update_media_gallery_display(self):
        item = self.all_system_features[self.current_media_index]
        clean_text = item["text"].replace('\n', ' ')

        lang = self.settings.get("language", "國語 (華語)")
        hint_text = I18N.get(lang, I18N["國語 (華語)"]).get("card_click_hint", "👉【點選中間即可執行此功能】")

        display_text = f"{clean_text}\n\n{hint_text}"
        self.transition_text(
            self.btn_media_card,
            display_text,
            configure=lambda: self.btn_media_card.configure(
                bg=item["bg_light"], fg=item["fg_dark"],
                activebackground=item["bg_light"], activeforeground=item["fg_dark"],
                wraplength=430, justify="center", anchor="center"
            ),
        )
    def next_media_item(self):
        """切換到下一個系統功能項目"""
        self.current_media_index = (self.current_media_index + 1) % len(self.all_system_features)
        self.update_media_gallery_display()
        self.play_system_sound("click")
        
        # 取得不換行的名稱以供語音朗讀
        speak_name = self.all_system_features[self.current_media_index]["text"].replace('\n', ' ')
        self.tts_speak(f"已更換，當前功能是：{speak_name}")

    def prev_media_item(self):
        """切換到上一個系統功能項目"""
        self.current_media_index = (self.current_media_index - 1) % len(self.all_system_features)
        self.update_media_gallery_display()
        self.play_system_sound("click")
        
        speak_name = self.all_system_features[self.current_media_index]["text"].replace('\n', ' ')
        self.tts_speak(f"已更換，當前功能是：{speak_name}")

    def click_media_card(self):
        """點擊中央去邊框大字卡時，直接執行目前選中的系統功能"""
        item = self.all_system_features[self.current_media_index]
        self.play_system_sound("execute")
        item["cmd"]()  # 執行該功能綁定的 Function
    def update_live_time(self):
        weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
        now = datetime.datetime.now()
        current_week = weekdays[now.weekday()]
        if self.settings.get("language") == "English":
            current_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][now.weekday()]
            time_str = now.strftime(f"%Y-%m-%d {current_week} %H:%M:%S")
        else:
            time_str = now.strftime(f"%Y年%m月%d日 {current_week} %H:%M:%S")
        
        # 👈 每秒順便檢查並套用當前主題配色
        self.check_and_apply_theme()
        
        try:
            self.time_label.configure(text=time_str)
            self.root.after(1000, self.update_live_time)
        except:
            pass
    def trigger_chimp_game(self):
        """點擊「記憶小遊戲」的動作：彈出動態按鈕散落版的數字記憶記憶測驗"""
        self.log_status("數字記憶遊戲已啟動...")
        self.play_system_sound("execute")
        self.tts_speak("歡迎來到數字記憶挑戰。準備好後請按下開始遊戲。")

        # 1. 建立獨立遊戲頂層視窗
        game_win = tk.Toplevel(self.root)
        game_win.title(self.localize_runtime_text("數字記憶挑戰"))
        game_win.geometry("850x780")
        game_win.configure(bg="#F5F3FF")

        # 2. 遊戲狀態變數
        game_vars = {
            "level": 1,             # 目前關卡
            "next_target": 1,       # 長者下一個應該點擊的數字
            "numbers_on_board": [],    # 目前場上的數字清單
            "buttons_dict": {},     # 存放當前關卡按鈕的字典 (key為數字)
            "lives": 3,             # 共有三條命
            "game_over": False,     # 遊戲是否結束
            "round_times": [],
            "level_started_at": None,
            "timer_after_id": None,
        }

        # 3. UI 頂部資訊列
        lbl_title = tk.Label(
            game_win, text="數字記憶挑戰", 
            font=("Microsoft JhengHei", 26, "bold"), bg="#F5F3FF", fg="#6D28D9"
        )
        lbl_title.pack(pady=10)

        # 狀態列框架
        info_frame = tk.Frame(game_win, bg="#F5F3FF")
        info_frame.pack(fill="x", padx=40)

        lbl_level = tk.Label(
            info_frame, text=self.localize_runtime_text(f"目前難度：{game_vars['level'] + 2} 個數字"), 
            font=("Microsoft JhengHei", 18, "bold"), bg="#F5F3FF", fg="#4C1D95"
        )
        lbl_level.pack(side="left", padx=20)

        lbl_lives = tk.Label(
            info_frame, text=self.localize_runtime_text(f"剩餘機會：{ '❤' * game_vars['lives'] }"), 
            font=("Microsoft JhengHei", 18, "bold"), bg="#F5F3FF", fg="#DC2626"
        )
        lbl_lives.pack(side="right", padx=20)

        lbl_round_time = tk.Label(
            info_frame,
            text=self.localize_runtime_text("本關經過時間：0.0 秒"),
            font=("Microsoft JhengHei", 16, "bold"),
            bg="#F5F3FF",
            fg="#7C3AED",
        )
        lbl_round_time.pack(expand=True)

        # 核心遊戲面板：5x5 隱形網格容器 (底色為淡紫色)
        grid_frame = tk.Frame(game_win, bg="#E9D5FF", bd=3, relief="ridge")
        grid_frame.pack(pady=15, padx=40, fill="both", expand=True)

        # 預先配置 5x5 的行列權重，確保散落時每個格子大小依然均等
        for i in range(5):
            grid_frame.rowconfigure(i, weight=1)
            grid_frame.columnconfigure(i, weight=1)

        lbl_hint = tk.Label(
            game_win, text=self.localize_runtime_text("提示：請專心記住畫面上數字的位置..."), 
            font=("Microsoft JhengHei", 16, "bold"), bg="#F5F3FF", fg="#5B21B6"
        )
        lbl_hint.pack(pady=5)

        # 4. 遊戲核心邏輯
        def update_round_timer():
            if not game_win.winfo_exists() or game_vars["game_over"]:
                game_vars["timer_after_id"] = None
                return
            started_at = game_vars["level_started_at"]
            if started_at is None:
                game_vars["timer_after_id"] = None
                return

            elapsed = time.monotonic() - started_at
            lbl_round_time.configure(
                text=self.localize_runtime_text(
                    f"本關經過時間：{elapsed:.1f} 秒"
                )
            )
            game_vars["timer_after_id"] = game_win.after(
                100, update_round_timer
            )

        def start_new_round():
            """開始新的一輪：生成散落按鈕，並在展示期強制鎖定以防誤觸"""
            if not game_win.winfo_exists(): return
            if game_vars["game_over"]: return

            # 每次清空舊按鈕
            for btn in game_vars["buttons_dict"].values():
                btn.destroy()
            game_vars["buttons_dict"].clear()

            game_vars["next_target"] = 1
            lbl_level.configure(text=self.localize_runtime_text(f"目前難度：{game_vars['level'] + 2} 個數字"))
            lbl_lives.configure(text=self.localize_runtime_text(f"剩餘機會：{ '❤' * game_vars['lives'] }"))
            if game_vars["level_started_at"] is None:
                lbl_round_time.configure(
                    text=self.localize_runtime_text("本關經過時間：0.0 秒")
                )
            lbl_hint.configure(text=self.localize_runtime_text("👀 專心看！數字位置顯示中..."), fg="#1E3A8A")

            # 決定這局有幾顆按鈕
            num_count = min(2 + game_vars["level"], 15)
            target_numbers = list(range(1, num_count + 1))
            game_vars["numbers_on_board"] = target_numbers.copy()

            # 隨機抽取位置
            all_positions = [(r, c) for r in range(5) for c in range(5)]
            chosen_positions = random.sample(all_positions, num_count)

            # 依據隨機座標生成按鈕
            for num, (row, col) in zip(target_numbers, chosen_positions):
                btn = tk.Button(
                    grid_frame, text=str(num), font=("Microsoft JhengHei", 24, "bold"),
                    bg="white", fg="#2563EB", relief="raised", bd=4, activebackground="#DDBFFF",
                    command=lambda n=num: on_card_click(n)
                )
                
                # 💡 【關鍵防護一】：一生成就立刻將狀態設為 disabled，防止長者在 2.5 秒內按到！
                btn.configure(state="disabled")
                
                # 丟進對應的網格座標
                btn.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
                game_vars["buttons_dict"][num] = btn

            # 給使用者 2.5 秒的時間記位置，隨後隱藏並解鎖
            game_win.after(2500, hide_numbers)

        def hide_numbers():
            """將字體隱藏，並正式解除按鈕鎖定，開放使用者點擊"""
            if not game_win.winfo_exists(): return
            if game_vars["game_over"]: return
            if game_vars["next_target"] != 1: return 

            lbl_hint.configure(text=self.localize_runtime_text("🟩 數字已隱藏！請憑記憶從小到大點擊方塊！"), fg="#EA580C")
            if game_vars["level_started_at"] is None:
                game_vars["level_started_at"] = time.monotonic()
                update_round_timer()
            
            # 將存在的那幾顆按鈕抹白，並恢復成可點擊狀態
            for num, btn in game_vars["buttons_dict"].items():
                if btn.winfo_exists():
                    # 💡 【關鍵防護二】：抹白字體的同時，將狀態改回 normal，這時點擊才有效！
                    btn.configure(text=" ", bg="#E2E8F0", activebackground="#CBD5E1", state="normal")

        def on_card_click(clicked_num):
            """點擊判定"""
            if not game_win.winfo_exists() or game_vars["game_over"]: return
            
            target = game_vars["next_target"]
            btn = game_vars["buttons_dict"].get(clicked_num)
            if not btn: return

            if clicked_num == target:
                # 點對了：顯現數字並切換為綠色成功狀態
                btn.configure(text=str(clicked_num), bg="#DCFCE7", fg="#16A34A", state="disabled")
                self.play_system_sound("execute")
                game_vars["next_target"] += 1
                
                # 檢查是否整局清空
                if game_vars["next_target"] > max(game_vars["numbers_on_board"]):
                    if game_vars["level_started_at"] is not None:
                        round_elapsed = (
                            time.monotonic() - game_vars["level_started_at"]
                        )
                        game_vars["round_times"].append(round_elapsed)
                        lbl_round_time.configure(
                            text=self.localize_runtime_text(
                                f"本關經過時間：{round_elapsed:.1f} 秒"
                            )
                        )
                    game_vars["level_started_at"] = None
                    if game_vars["timer_after_id"] is not None:
                        game_win.after_cancel(game_vars["timer_after_id"])
                        game_vars["timer_after_id"] = None
                    lbl_hint.configure(text=self.localize_runtime_text("🎉 太厲害了！完全答對！難度提升！"), fg="#16A34A")
                    self.tts_speak("哇！全對！太厲害了，我們晉級下一關。")
                    
                    game_vars["level"] += 1
                    game_win.after(2000, start_new_round)
            else:
                # 點錯了（扣命機制）
                self.play_system_sound("warn")
                game_vars["lives"] -= 1
                lbl_lives.configure(text=self.localize_runtime_text(f"剩餘機會：{ '❤' * game_vars['lives'] }"))
                
                # 點錯懲罰：強行翻開場上所有答案
                for num, b in game_vars["buttons_dict"].items():
                    if b.winfo_exists():
                        b.configure(text=str(num), fg="#475569", state="disabled")
                        if num == target:
                            b.configure(bg="#FECACA", fg="#DC2626")

                if game_vars["lives"] > 0:
                    lbl_hint.configure(text=self.localize_runtime_text("❌ 答錯囉！扣一條命，兩秒後重新挑戰本關！"), fg="#DC2626")
                    self.tts_speak(f"猜錯囉，沒關係，我們再來一次。")
                    game_win.after(3000, start_new_round)
                else:
                    # 💡 【修改重點】：三條命扣光，清除網格並顯示成績結算
                    game_vars["game_over"] = True
                    if game_vars["level_started_at"] is not None:
                        round_elapsed = (
                            time.monotonic() - game_vars["level_started_at"]
                        )
                        lbl_round_time.configure(
                            text=self.localize_runtime_text(
                                f"本關經過時間：{round_elapsed:.1f} 秒"
                            )
                        )
                    if game_vars["timer_after_id"] is not None:
                        game_win.after_cancel(game_vars["timer_after_id"])
                        game_vars["timer_after_id"] = None
                    lbl_hint.configure(text=self.localize_runtime_text("🏁 挑戰結束！"), fg="#7F1D1D")
                    lbl_lives.configure(text=self.localize_runtime_text("剩餘機會：❌ 已用盡"))
                    
                    # 1. 計算最終成績 (通過的關卡數 = 當前 level - 1)
                    final_score = len(game_vars["round_times"])
                    average_round_time = (
                        sum(game_vars["round_times"]) / final_score
                        if final_score
                        else None
                    )
                    
                    # 2. 把 5x5 的網格框架清空，用來放結算畫面
                    for widget in grid_frame.winfo_children():
                        widget.destroy()
                    
                    # 3. 建立結算畫面容器 (滿版鋪在原本網格的位置)
                    score_panel = tk.Frame(grid_frame, bg="#FEE2E2")
                    score_panel.place(relx=0, rely=0, relwidth=1, relheight=1)
                    
                    # 4. 顯示得分文字
                    lbl_score_title = tk.Label(
                        score_panel, text=self.localize_runtime_text("🏆 今日鍛鍊成果 🏆"),
                        font=("Microsoft JhengHei", 28, "bold"), bg="#FEE2E2", fg="#991B1B"
                    )
                    lbl_score_title.pack(pady=30)
                    
                    lbl_score_num = tk.Label(
                        score_panel, text=self.localize_runtime_text(f"總共成功闖過： {final_score} 關"),
                        font=("Microsoft JhengHei", 36, "bold"), bg="#FEE2E2", fg="#DC2626"
                    )
                    lbl_score_num.pack(pady=15)

                    average_time_text = (
                        self.localize_runtime_text(
                            f"平均每關時間：{average_round_time:.1f} 秒"
                        )
                        if average_round_time is not None
                        else self.localize_runtime_text(
                            "平均每關時間：尚未完成任何關卡"
                        )
                    )
                    lbl_average_time = tk.Label(
                        score_panel,
                        text=average_time_text,
                        font=("Microsoft JhengHei", 24, "bold"),
                        bg="#FEE2E2",
                        fg="#7F1D1D",
                    )
                    lbl_average_time.pack(pady=10)
                    
                    # 5. 語音播報成績
                    self.tts_speak(f"挑戰結束。您今天總共成功闖過 {final_score} 關，再接再厲！要再玩一次嗎？")
                    
                    # 6. 核心：放置「再玩一次」大按鈕
                    btn_retry = tk.Button(
                        score_panel, text="🔄 挑戰重來 ｜ 再玩一次", font=("Microsoft JhengHei", 22, "bold"),
                        bg="#16A34A", fg="white", activebackground="#15803D", activeforeground="white",
                        padx=40, pady=15, relief="raised", bd=5, cursor="hand2",
                        command=restart_game
                    )
                    btn_retry.pack(pady=40)

        # 💡 【全新加入】：重玩遊戲的重置邏輯
        def restart_game():
            """將所有遊戲數據歸零，重新開始挑戰"""
            game_vars["level"] = 1
            game_vars["lives"] = 3
            game_vars["game_over"] = False
            game_vars["round_times"] = []
            game_vars["level_started_at"] = None
            
            # 清除結算畫面，重新呼叫第一局
            for widget in grid_frame.winfo_children():
                widget.destroy()
                
            self.play_system_sound("execute")
            self.tts_speak("好的，新挑戰開始，請看畫面！")
            start_new_round()

        # 5. 顯示起始畫面，避免遊戲在視窗開啟後自動開始
        start_panel = tk.Frame(grid_frame, bg="#EDE9FE")
        start_panel.place(relx=0, rely=0, relwidth=1, relheight=1)

        tk.Label(
            start_panel,
            text=self.localize_runtime_text("遊戲準備好了！"),
            font=("Microsoft JhengHei", 30, "bold"),
            bg="#EDE9FE",
            fg="#5B21B6"
        ).pack(pady=(120, 25))

        tk.Label(
            start_panel,
            text=self.localize_runtime_text("請按下開始遊戲，準備記住數字的位置。"),
            font=("Microsoft JhengHei", 18),
            bg="#EDE9FE",
            fg="#4C1D95",
            wraplength=600
        ).pack(pady=15)

        def begin_game():
            if not game_win.winfo_exists() or game_vars["game_over"]:
                return
            start_panel.destroy()
            self.play_system_sound("execute")
            self.tts_speak("請記住畫面上數字的位置，並從小到大依序點擊。")
            start_new_round()

        tk.Button(
            start_panel,
            text=self.localize_runtime_text("開始遊戲"),
            font=("Microsoft JhengHei", 24, "bold"),
            bg="#7C3AED",
            fg="white",
            activebackground="#6D28D9",
            activeforeground="white",
            padx=50,
            pady=18,
            relief="raised",
            bd=5,
            cursor="hand2",
            command=begin_game
        ).pack(pady=35)

        # 6. 底部的關閉退出按鈕
        def close_game():
            self.tts_speak("不玩了，休息一下吧。")
            game_win.destroy()

        exit_btn = tk.Button(
            game_win, text="❌ 結束遊戲 ｜ 返回首頁", font=("Microsoft JhengHei", 16, "bold"),
            bg="#EF4444", fg="white", activebackground="#DC2626", activeforeground="white",
            padx=30, pady=10, command=close_game, cursor="hand2"
        )
        exit_btn.pack(pady=15)

        self.localize_window_widgets(game_win)

    def trigger_finger_gesture_game(self):
        """開啟鏡頭手勢挑戰；遊戲期間獨佔手勢辨識，避免控制滑鼠或觸發快捷鍵。"""
        if self.gesture_game_window_active:
            return

        game_win = tk.Toplevel(self.root)
        self.gesture_game_window_active = True
        game_win.title(self.localize_runtime_text("鏡頭手勢互動挑戰"))
        game_win.geometry("760x760")
        game_win.minsize(760, 700)
        game_win.configure(bg="#FFF7ED")

        actions = [
            ("右手張開手掌", {"Right": "open"}),
            ("右手握拳", {"Right": "fist"}),
            ("右手比2指", {"Right": "two"}),
            ("右手比3指", {"Right": "three"}),
            ("右手比一", {"Right": "point"}),
            ("右手讚", {"Right": "thumbs_up"}),
            ("右手 OK 手勢", {"Right": "ok"}),
            ("右手比槍手勢", {"Right": "gun"}),
            ("左手張開手掌", {"Left": "open"}),
            ("左手握拳", {"Left": "fist"}),
            ("左手比2指", {"Left": "two"}),
            ("左手比3指", {"Left": "three"}),
            ("左手比一", {"Left": "point"}),
            ("左手讚", {"Left": "thumbs_up"}),
            ("左手 OK 手勢", {"Left": "ok"}),
            ("左手比槍手勢", {"Left": "gun"}),
            ("雙手張開手掌", {"Right": "open", "Left": "open"}),
            ("雙手握拳", {"Right": "fist", "Left": "fist"}),
            ("雙手比2指", {"Right": "two", "Left": "two"}),
            ("雙手比3指", {"Right": "three", "Left": "three"}),
            ("雙手比槍手勢", {"Right": "gun", "Left": "gun"}),
            ("雙手 OK 手勢", {"Right": "ok", "Left": "ok"}),
            ("雙手比一", {"Right": "point", "Left": "point"}),
            ("雙手交叉", {"__both__": "crossed"}),
            ("雙手圓圈", {"__both__": "circle"}),
        ]
        state = {
            "score": 0,
            "timeout_count": 0,
            "response_times": [],
            "current_streak": 0,
            "max_streak": 0,
            "round": 0,
            "target": None,
            "finished": False,
            "deadline": time.time() + 600,
            "last_match": None,
            "match_frames": 0,
            "await_release": False,
            "actions": actions,
            "previous_target": None,
            "round_deadline": None,
            "round_id": 0,
            "round_started_at": None,
            "timeout_handled_round": None,
            "duration_minutes": 10,
        }
        self.gesture_game_state = state

        tk.Label(
            game_win, text=self.localize_runtime_text("鏡頭手勢互動挑戰"),
            font=("Microsoft JhengHei", 28, "bold"), bg="#FFF7ED", fg="#C2410C"
        ).pack(pady=(22, 8))
        tk.Label(
            game_win,
            text=self.localize_runtime_text("請看提示，用左右手做出指定動作。遊戲期間不會操控滑鼠或觸發系統快捷手勢。"),
            font=("Microsoft JhengHei", 15), bg="#FFF7ED", fg="#7C2D12",
            wraplength=680, justify="center"
        ).pack(pady=4)

        info = tk.Frame(game_win, bg="#FFF7ED")
        info.pack(fill="x", padx=45, pady=16)
        score_label = tk.Label(info, text="", font=("Microsoft JhengHei", 18, "bold"), bg="#FFF7ED", fg="#166534")
        score_label.pack(side="left")
        time_label = tk.Label(info, text="", font=("Microsoft JhengHei", 18, "bold"), bg="#FFF7ED", fg="#9A3412")
        time_label.pack(side="right")
        stats_label = tk.Label(
            game_win, text="", font=("Microsoft JhengHei", 13, "bold"),
            bg="#FFF7ED", fg="#475569", wraplength=690, justify="center"
        )
        stats_label.pack(pady=(0, 8))
        state["stats_label"] = stats_label
        target_label = tk.Label(
            game_win, text="", font=("Microsoft JhengHei", 32, "bold"),
            bg="#FED7AA", fg="#9A3412", padx=25, pady=28
        )
        target_label.pack(fill="x", padx=45, pady=20)
        state["target_label"] = target_label
        state["score_label"] = score_label
        status_label = tk.Label(
            game_win, text=self.localize_runtime_text("按下開始遊戲後，請將手放在鏡頭前。"),
            font=("Microsoft JhengHei", 16), bg="#FFF7ED", fg="#475569",
            wraplength=650, justify="center"
        )
        status_label.pack(pady=8)
        state["status_label"] = status_label

        camera_hint = tk.Label(
            game_win,
            text=self.localize_runtime_text("開始前請先開啟相機視窗，確認手部能被鏡頭看見。"),
            font=("Microsoft JhengHei", 15, "bold"),
            bg="#FEF3C7",
            fg="#92400E",
            wraplength=650,
            justify="center",
            padx=15,
            pady=10,
        )
        camera_hint.pack(fill="x", padx=45, pady=(8, 4))

        def camera_is_ready():
            camera_ready = self._is_camera_open()
            camera_visible = (
                hasattr(self, "cam_win")
                and self.cam_win.winfo_exists()
                and self.cam_win.state() != "withdrawn"
            )
            return camera_ready and camera_visible

        def open_camera_for_game():
            if not self._is_camera_open():
                camera_button.configure(state="disabled")
                camera_hint.configure(
                    text=self.localize_runtime_text("相機重啟中..."),
                    bg="#FEF3C7",
                    fg="#92400E",
                )
                self.init_camera(is_boot=False)
                game_win.after(100, check_camera_for_game)
                return
            if hasattr(self, "cam_win") and self.cam_win.winfo_exists():
                self.cam_win.deiconify()
                self.cam_win.lift()
            camera_hint.configure(
                text=self.localize_runtime_text("相機已開啟，請確認雙手出現在鏡頭畫面中。"),
                bg="#DCFCE7",
                fg="#166534",
            )
            camera_button.configure(state="disabled")

        def check_camera_for_game():
            if not game_win.winfo_exists():
                return
            if camera_is_ready():
                if hasattr(self, "cam_win") and self.cam_win.winfo_exists():
                    self.cam_win.deiconify()
                    self.cam_win.lift()
                camera_hint.configure(
                    text=self.localize_runtime_text("相機已開啟，請確認雙手出現在鏡頭畫面中。"),
                    bg="#DCFCE7",
                    fg="#166534",
                )
                return
            if self.is_refreshing:
                game_win.after(100, check_camera_for_game)
                return
            camera_hint.configure(
                text=self.localize_runtime_text("相機尚未就緒，請檢查攝影機是否已連接。"),
                bg="#FEE2E2",
                fg="#991B1B",
            )
            camera_button.configure(state="normal")

        camera_button = tk.Button(
            game_win,
            text=self.localize_runtime_text("開啟相機視窗"),
            font=("Microsoft JhengHei", 17, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            bd=0,
            padx=28,
            pady=9,
            command=open_camera_for_game,
            cursor="hand2",
        )
        camera_button.pack(pady=(2, 8))

        duration_frame = tk.Frame(game_win, bg="#FFF7ED")
        duration_frame.pack(pady=(2, 5))
        tk.Label(
            duration_frame,
            text=self.localize_runtime_text("遊戲時間："),
            font=("Microsoft JhengHei", 16, "bold"),
            bg="#FFF7ED",
            fg="#7C2D12",
        ).pack(side="left", padx=(0, 8))
        duration_var = tk.StringVar(value="10")
        duration_combo = ttk.Combobox(
            duration_frame,
            textvariable=duration_var,
            values=["3", "5", "10"],
            state="readonly",
            width=5,
            font=("Microsoft JhengHei", 15),
        )
        duration_combo.pack(side="left")
        tk.Label(
            duration_frame,
            text=self.localize_runtime_text("分鐘"),
            font=("Microsoft JhengHei", 16, "bold"),
            bg="#FFF7ED",
            fg="#7C2D12",
        ).pack(side="left", padx=(8, 0))

        def refresh_game_stats():
            response_times = state["response_times"]
            average_response = (
                f"{sum(response_times) / len(response_times):.1f}"
                if response_times else self.localize_runtime_text("尚無")
            )
            stats_label.configure(
                text=(
                    f"{self.localize_runtime_text('超時次數：')}"
                    f"{state['timeout_count']}{self.localize_runtime_text('次')}　"
                    f"{self.localize_runtime_text('平均反應時間：')}"
                    f"{average_response}"
                    f"{self.localize_runtime_text('秒') if response_times else ''}　"
                    f"{self.localize_runtime_text('最高連續答對：')}"
                    f"{state['max_streak']}{self.localize_runtime_text('次')}"
                )
            )

        def save_game_stats():
            response_times = state["response_times"]
            self.settings["finger_game_stats"] = {
                "recorded_at": datetime.datetime.now().isoformat(timespec="seconds"),
                "timeout_count": state["timeout_count"],
                "average_response_seconds": (
                    round(sum(response_times) / len(response_times), 2)
                    if response_times else None
                ),
                "max_consecutive_correct": state["max_streak"],
                "correct_count": state["score"],
            }
            self.save_settings(self.settings)

        state["refresh_game_stats"] = refresh_game_stats

        def finish_game(reason):
            if state["finished"]:
                return
            state["finished"] = True
            self.gesture_game_active = False
            self.gesture_game_state = state
            save_game_stats()
            refresh_game_stats()
            target_label.configure(text=self.localize_runtime_text("挑戰結束"))
            status_label.configure(
                text=self.localize_runtime_text(f"本次完成 {state['score']} 個動作。{reason}"),
                fg="#166534"
            )
            start_button.configure(
                state="normal",
                text=self.localize_runtime_text("重新開始")
            )
            duration_combo.configure(state="readonly")
            self.play_system_sound("execute")
            self.tts_speak(f"挑戰結束，您完成了 {state['score']} 個動作。")

        def update_clock():
            if not game_win.winfo_exists() or state["finished"]:
                return
            remaining = max(0, int(state["deadline"] - time.time()))
            minutes, seconds = divmod(remaining, 60)
            time_label.configure(text=f"{self.localize_runtime_text('剩餘時間')} {minutes:02d}:{seconds:02d}")
            if remaining <= 0:
                finish_game(f"{state['duration_minutes']} 分鐘時間到了。")
            elif (
                state["round_deadline"] is not None
                and time.time() >= state["round_deadline"]
                and state["timeout_handled_round"] != state["round_id"]
            ):
                timeout_round_id = state["round_id"]
                state["timeout_handled_round"] = timeout_round_id
                state["timeout_count"] += 1
                state["current_streak"] = 0
                refresh_game_stats()
                status_label.configure(
                    text=self.localize_runtime_text("這個動作超過 20 秒，換下一個動作。"),
                    fg="#C2410C"
                )
                self.root.after(
                    1200,
                    lambda: self._advance_gesture_game_round(timeout_round_id)
                )
                game_win.after(250, update_clock)
            else:
                game_win.after(250, update_clock)

        def next_target():
            available_actions = [
                action for action in actions
                if action[0] != state["previous_target"]
            ] or actions
            state["target"] = random.choice(available_actions)
            state["previous_target"] = state["target"][0]
            state["round"] += 1
            state["round_id"] += 1
            state["round_deadline"] = min(time.time() + 20, state["deadline"])
            state["round_started_at"] = time.perf_counter()
            state["match_frames"] = 0
            state["await_release"] = False
            target_label.configure(text=self.localize_runtime_text(f"請做：{state['target'][0]}"))
            score_label.configure(text=self.localize_runtime_text(f"完成：{state['score']}"))
            status_label.configure(text=self.localize_runtime_text("保持動作約半秒，辨識成功後再放開。"), fg="#475569")

        def begin_game():
            if not camera_is_ready():
                should_open = messagebox.askyesno(
                    self.localize_runtime_text("需要開啟相機"),
                    self.localize_runtime_text("相機視窗目前未開啟，是否現在開啟相機？"),
                    parent=game_win
                )
                if not should_open:
                    status_label.configure(
                        text=self.localize_runtime_text("請先開啟相機視窗，才能開始手勢挑戰。"),
                        fg="#DC2626"
                    )
                    return
                open_camera_for_game()
                if not camera_is_ready():
                    return
            state["duration_minutes"] = int(duration_var.get())
            state["deadline"] = time.time() + state["duration_minutes"] * 60
            state["score"] = 0
            state["timeout_count"] = 0
            state["response_times"] = []
            state["current_streak"] = 0
            state["max_streak"] = 0
            state["finished"] = False
            state["previous_target"] = None
            state["match_frames"] = 0
            state["await_release"] = False
            state["round_deadline"] = None
            state["round_started_at"] = None
            state["timeout_handled_round"] = None
            refresh_game_stats()
            self.gesture_game_active = True
            start_button.configure(
                state="disabled",
                text=self.localize_runtime_text("開始遊戲")
            )
            duration_combo.configure(state="disabled")
            self.play_system_sound("execute")
            self.tts_speak("遊戲開始，請依照畫面提示做動作。")
            next_target()
            update_clock()

        def close_game():
            if self.gesture_game_active:
                save_game_stats()
            self.gesture_game_active = False
            self.gesture_game_window_active = False
            self.gesture_game_state = None
            if game_win.winfo_exists():
                game_win.destroy()

        start_button = tk.Button(
            game_win, text=self.localize_runtime_text("開始遊戲"),
            font=("Microsoft JhengHei", 22, "bold"), bg="#EA580C", fg="white",
            activebackground="#C2410C", activeforeground="white", bd=0,
            padx=40, pady=12, command=begin_game, cursor="hand2"
        )
        start_button.pack(pady=14)
        tk.Button(
            game_win, text=self.localize_runtime_text("結束遊戲｜返回首頁"),
            font=("Microsoft JhengHei", 15, "bold"), bg="#64748B", fg="white",
            activebackground="#475569", activeforeground="white", bd=0,
            padx=25, pady=8, command=close_game, cursor="hand2"
        ).pack(pady=8)
        game_win.protocol("WM_DELETE_WINDOW", close_game)
        self.localize_window_widgets(game_win)

    def _get_hand_gesture(self, hand_lms):
        """將單手 landmarks 分成遊戲需要的穩定動作。"""
        landmarks = hand_lms.landmark
        finger_up = {
            tip: landmarks[tip].y < landmarks[tip - 2].y - 0.015
            for tip in [INDEX, MIDDLE, RING, PINKY]
        }
        extended_fingers = [
            tip for tip, is_up in finger_up.items() if is_up
        ]
        finger_count = len(extended_fingers)
        wrist = landmarks[0]
        thumb_ip = landmarks[THUMB - 1]
        thumb_tip = landmarks[THUMB]
        wrist_to_thumb = math.hypot(
            thumb_tip.x - wrist.x,
            thumb_tip.y - wrist.y,
        )
        wrist_to_thumb_ip = math.hypot(
            thumb_ip.x - wrist.x,
            thumb_ip.y - wrist.y,
        )
        thumb_extended = wrist_to_thumb > wrist_to_thumb_ip * 1.25
        thumb_up = thumb_tip.y < thumb_ip.y - 0.04
        index_thumb_distance = math.hypot(
            landmarks[THUMB].x - landmarks[INDEX].x,
            landmarks[THUMB].y - landmarks[INDEX].y,
        )
        other_fingers_up = (
            finger_up[MIDDLE] or finger_up[RING] or finger_up[PINKY]
        )
        if index_thumb_distance < 0.075 and other_fingers_up:
            return "ok"
        if thumb_extended and finger_up[INDEX] and not other_fingers_up:
            return "gun"
        if thumb_up and finger_count == 0:
            return "thumbs_up"
        if finger_count >= 4:
            return "open"
        if finger_count == 3:
            return "three"
        if finger_count == 2:
            return "two"
        if finger_up[INDEX] and finger_count == 1:
            return "point"
        if finger_count == 0 and not thumb_extended and not thumb_up:
            return "fist"
        return "other"

    def _match_two_hand_gesture(self, detected, hand_landmarks):
        """判斷需要兩手配合的交叉與圓圈動作。"""
        if len(hand_landmarks) != 2:
            return False
        hands = {label: landmarks for label, landmarks in hand_landmarks.items()}
        if "crossed" in detected:
            return False
        if "Right" not in hands or "Left" not in hands:
            return False
        right_wrist = hands["Right"].landmark[0]
        left_wrist = hands["Left"].landmark[0]
        right_middle = hands["Right"].landmark[MIDDLE]
        left_middle = hands["Left"].landmark[MIDDLE]
        wrist_distance = math.hypot(
            right_wrist.x - left_wrist.x,
            right_wrist.y - left_wrist.y,
        )
        middle_distance = math.hypot(
            right_middle.x - left_middle.x,
            right_middle.y - left_middle.y,
        )
        target = self.gesture_game_state.get("target")
        special = target[1].get("__both__") if target else None
        if special == "crossed":
            return (
                (
                    (
                        right_wrist.x < left_wrist.x
                        and right_middle.x > left_middle.x
                    ) or (
                        right_wrist.x > left_wrist.x
                        and right_middle.x < left_middle.x
                    )
                )
                and wrist_distance < 0.45
            )
        if special == "circle":
            return (
                detected.get("Right") == "ok"
                and detected.get("Left") == "ok"
            ) or (
                wrist_distance < 0.3
                and middle_distance < 0.22
                and detected.get("Right") == "open"
                and detected.get("Left") == "open"
            )
        return False

    def _process_gesture_game(self, results):
        """在相機執行緒中判斷遊戲手勢；不進入一般滑鼠與快捷鍵流程。"""
        if not self.gesture_game_active or not self.gesture_game_state:
            return False
        state = self.gesture_game_state
        detected = {}
        hand_landmarks = {}
        if results.multi_hand_landmarks and results.multi_handedness:
            for hand_lms, handedness in zip(results.multi_hand_landmarks, results.multi_handedness):
                label = handedness.classification[0].label
                detected[label] = self._get_hand_gesture(hand_lms)
                hand_landmarks[label] = hand_lms
        target = state.get("target")
        special_target = target and target[1].get("__both__")
        matched = (
            self._match_two_hand_gesture(detected, hand_landmarks)
            if special_target
            else bool(target) and all(
                detected.get(hand) == gesture for hand, gesture in target[1].items()
            )
        )
        if matched and not state["await_release"]:
            state["match_frames"] += 1
            if state["match_frames"] >= 8:
                state["await_release"] = True
                round_id = state.get("round_id")
                self._post_camera_ui(
                    lambda: self._complete_gesture_game_round(round_id)
                )
        elif not matched:
            state["match_frames"] = 0
            state["await_release"] = False
        return True

    def _complete_gesture_game_round(self, round_id):
        if not self.gesture_game_active or not self.gesture_game_state:
            return
        state = self.gesture_game_state
        if (
            round_id != state.get("round_id")
            or state.get("timeout_handled_round") == round_id
            or (
                state.get("round_deadline") is not None
                and time.time() >= state["round_deadline"]
            )
        ):
            return
        state["score"] += 1
        state["current_streak"] += 1
        state["max_streak"] = max(state["max_streak"], state["current_streak"])
        if state.get("round_started_at") is not None:
            state["response_times"].append(
                max(0.0, time.perf_counter() - state["round_started_at"])
            )
        state["refresh_game_stats"]()
        self.play_system_sound("execute")
        self.root.after(250, self._advance_gesture_game_round)

    def _advance_gesture_game_round(self, expected_round_id=None):
        state = self.gesture_game_state
        if not self.gesture_game_active or not state:
            return
        if (
            expected_round_id is not None
            and expected_round_id != state.get("round_id")
        ):
            return
        available_actions = [
            action for action in state["actions"]
            if action[0] != state.get("previous_target")
        ] or state["actions"]
        target = random.choice(available_actions)
        state["target"] = target
        state["previous_target"] = target[0]
        state["round"] += 1
        state["round_id"] += 1
        state["round_deadline"] = min(time.time() + 20, state["deadline"])
        state["round_started_at"] = time.perf_counter()
        state["timeout_handled_round"] = None
        state["match_frames"] = 0
        state["await_release"] = False
        if state["target_label"].winfo_exists():
            state["target_label"].configure(text=self.localize_runtime_text(f"請做：{target[0]}"))
        if state["score_label"].winfo_exists():
            state["score_label"].configure(text=self.localize_runtime_text(f"完成：{state['score']}"))
        if state["status_label"].winfo_exists():
            state["status_label"].configure(
                text=self.localize_runtime_text("保持動作約半秒，辨識成功後再放開。"),
                fg="#475569"
            )

    def create_camera_window(self):
        """建立固定為 4:3 的鏡頭與手勢分析視窗。"""
        self.cam_win = tk.Toplevel(self.root)
        self.cam_win.title(self.localize_runtime_text("手勢操作監控畫面"))
        self.cam_win.geometry("640x480")
        self.cam_win.resizable(False, False)
        
        # 預設彈出在螢幕右下角，避免干擾中央九宮格操作（留出適當邊距防遮擋）
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.cam_win.geometry(f"640x480+{sw - 670}+{sh - 540}")

        self.video_label = tk.Label(self.cam_win, bg="#000000")
        self.video_label.pack(expand=True, fill="both")
        
        # 使用者手動關閉時，僅隱藏視窗而不銷毀程式
        self.cam_win.protocol("WM_DELETE_WINDOW", self.close_camera_window)
        self.cam_win.withdraw()
        self.root.after(16, self.update_camera_display)

    def show_camera_window(self):
        if self.is_refreshing:
            if hasattr(self, "cam_win") and self.cam_win.winfo_exists():
                self.cam_win.deiconify()
                self.cam_win.lift()
            return
        if not self._is_camera_open():
            self.init_camera(is_boot=False)
            return

        if self.cam_win.winfo_exists():
            self.cam_win.deiconify()
            self.cam_win.lift()

    def _is_camera_open(self):
        return self.camera_ready

    @staticmethod
    def _crop_frame_to_aspect(frame, target_width, target_height):
        frame_height, frame_width = frame.shape[:2]
        target_aspect = target_width / target_height
        frame_aspect = frame_width / frame_height

        if frame_aspect > target_aspect:
            cropped_width = int(frame_height * target_aspect)
            x_offset = (frame_width - cropped_width) // 2
            return frame[:, x_offset:x_offset + cropped_width]

        if frame_aspect < target_aspect:
            cropped_height = int(frame_width / target_aspect)
            y_offset = (frame_height - cropped_height) // 2
            return frame[y_offset:y_offset + cropped_height, :]

        return frame

    def update_camera_display(self):
        if not self.loop_active:
            return

        self._run_camera_ui_callbacks()
        active_grab = str(self.root.tk.call("grab", "current"))
        self.modal_gesture_suspended = bool(active_grab) and any(
            active_grab == str(window)
            or active_grab.startswith(f"{window}.")
            for window in self.root.winfo_children()
            if isinstance(window, tk.Toplevel) and window.winfo_exists()
        )
        if (
            self.cam_win.winfo_exists()
            and self.cam_win.state() != "withdrawn"
            and not self.is_refreshing
            and self._is_camera_open()
        ):
            with self.camera_frame_lock:
                frame = (
                    self.latest_camera_frame.copy()
                    if self.latest_camera_frame is not None else None
                )

            if frame is not None:
                display_width = max(1, self.video_label.winfo_width())
                display_height = max(1, self.video_label.winfo_height())
                frame = self._crop_frame_to_aspect(
                    frame, display_width, display_height
                )
                interpolation = (
                    cv2.INTER_AREA
                    if display_width < frame.shape[1]
                    or display_height < frame.shape[0]
                    else cv2.INTER_LINEAR
                )
                display_frame = cv2.resize(
                    frame,
                    (display_width, display_height),
                    interpolation=interpolation,
                )
                display_rgb = cv2.cvtColor(display_frame, cv2.COLOR_BGR2RGB)
                photo = ImageTk.PhotoImage(image=Image.fromarray(display_rgb))
                self.video_label.configure(image=photo, text="")
                self.video_label.image = photo

        self.root.after(16, self.update_camera_display)

    def _show_camera_placeholder(self, text):
        if not hasattr(self, "video_label") or not self.video_label.winfo_exists():
            return
        self.video_label.configure(
            image="",
            text=self.localize_runtime_text(text),
            font=("Microsoft JhengHei", 30, "bold"),
            fg="white",
            bg="#1F2937",
            compound="center",
            justify="center",
        )
        self.video_label.image = None

    def _post_camera_ui(self, callback):
        self.camera_ui_queue.put(callback)

    def _run_camera_ui_callbacks(self):
        while True:
            try:
                callback = self.camera_ui_queue.get_nowait()
            except queue.Empty:
                return
            callback()

    def close_camera_window(self):
        """Hide the camera window and release the camera when it is not being viewed."""
        self._camera_init_generation += 1
        self.camera_ready = False
        self.is_paused = True
        self.is_refreshing = False
        capture = self.cap
        self.cap = None
        if capture is not None:
            self.camera_release_complete.clear()
            threading.Thread(
                target=self._release_camera_capture,
                args=(capture, self.camera_release_complete),
                daemon=True,
            ).start()
        if hasattr(self, "cam_win") and self.cam_win.winfo_exists():
            self.cam_win.withdraw()

    def _release_camera_capture(self, capture, completion_event=None):
        try:
            with self.camera_capture_lock:
                capture.release()
        finally:
            if completion_event is not None:
                completion_event.set()

    def init_camera(self, is_boot=True):
        """Start camera initialization without blocking the Tk event loop."""
        if not is_boot and self.is_refreshing:
            return
        self._camera_init_generation += 1
        generation = self._camera_init_generation
        self.camera_ready = False
        self.is_refreshing = True
        self.is_paused = True
        with self.camera_frame_lock:
            self.latest_camera_frame = None
            self.latest_camera_preview_frame = None

        if is_boot:
            if hasattr(self, "cam_win") and self.cam_win.winfo_exists():
                self.cam_win.withdraw()
            self.is_refreshing = False
            return

        self._show_camera_placeholder("相機重啟中")
        if hasattr(self, "cam_win") and self.cam_win.winfo_exists():
            self.cam_win.deiconify()
        camera_index = int(self.settings.get("camera_index", 0))

        if not hasattr(self, 'camera_thread_started'):
            self.camera_thread_started = True
            threading.Thread(target=self.camera_loop, daemon=True).start()

        self.log_status("智慧醫護：相機正在重新啟動。")
        self.tts_speak("相機重啟中")
        threading.Thread(
            target=self._open_camera_in_background,
            args=(generation, camera_index),
            daemon=True,
        ).start()

    def _open_camera_in_background(self, generation, camera_index):
        capture = None
        error = None
        try:
            self.camera_release_complete.wait()
            with self.camera_capture_lock:
                previous_capture = self.cap
                self.cap = None
            if previous_capture is not None:
                self._release_camera_capture(previous_capture)

            capture = cv2.VideoCapture(camera_index)
            if not capture.isOpened():
                raise RuntimeError(f"無法開啟攝影機裝置 {camera_index}")

            camera_width = int(self.settings.get("camera_width", 640))
            camera_height = int(self.settings.get("camera_height", 480))
            if (
                camera_width <= 0
                or camera_height <= 0
                or camera_width * 3 != camera_height * 4
            ):
                camera_width, camera_height = 640, 480
                self.settings["camera_width"] = camera_width
                self.settings["camera_height"] = camera_height

            capture.set(cv2.CAP_PROP_BUFFERSIZE, 1)
            capture.set(cv2.CAP_PROP_FRAME_WIDTH, camera_width)
            capture.set(cv2.CAP_PROP_FRAME_HEIGHT, camera_height)
            capture.set(cv2.CAP_PROP_FPS, 60)
            frame_available = False
            for _ in range(5):
                success, _ = capture.read()
                frame_available = frame_available or success
            if not frame_available:
                raise RuntimeError("無法讀取攝影機畫面")

            with self.camera_capture_lock:
                if generation != self._camera_init_generation or not self.loop_active:
                    capture.release()
                    return
                self.cap = capture
                self.camera_ready = True
                capture = None
        except Exception as exc:
            error = exc
        finally:
            if capture is not None:
                capture.release()
            self._post_camera_ui(
                lambda: self._finish_camera_initialization(generation, error)
            )

    def _finish_camera_initialization(self, generation, error):
        if generation != self._camera_init_generation or not self.loop_active:
            return
        self.is_refreshing = False
        self.is_paused = False
        if error is not None:
            print(f"【相機】初始化失敗：{error}")
            self._show_camera_placeholder("相機啟動失敗")
            self.log_status("智慧醫護：相機連線失敗，請檢查硬體。")
            self.play_system_sound("warn")
            self.tts_speak("相機啟動失敗")
            return
        print("【相機】重新啟動成功。")
        self.tts_speak("相機已就緒")

    def camera_loop(self):
        screen_w, screen_h = pyautogui.size()
        
        while self.loop_active:
            # 智慧安全鎖：重啟期間不執行影像擷取與手勢分析，留給灰階畫面呈現
            if self.is_refreshing or not self.camera_ready:
                time.sleep(0.1)
                continue
                
            with self.camera_capture_lock:
                capture = self.cap
                if capture is None or not capture.isOpened():
                    ret, frame = False, None
                else:
                    ret, frame = capture.read()
            if not ret:
                time.sleep(0.01)
                continue

            frame = self._crop_frame_to_aspect(
                frame,
                int(self.settings.get("camera_width", 640)),
                int(self.settings.get("camera_height", 480)),
            )
            frame = cv2.flip(frame, 1)
            preview_frame = frame.copy()
            fh, fw, _ = frame.shape
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = self.hands.process(rgb_frame)
            fall_detected=False
            fall_score=0
            fall_status = "偵測中..."
            if hasattr(self,"fall_detector") and self.fall_detector is not None and self.settings.get("fall_detection_enabled",True):
                try:
                    frame, fall_status, fall_score, fall_detected = self.fall_detector.process(frame)
                except Exception as e:
                    print(f"【FallDetector】camera_loop 錯誤：{e}")
                    fall_status = "跌倒偵測錯誤"

            if fall_detected or (hasattr(self, "fall_detector") and self.fall_detector is not None and getattr(self.fall_detector, "alert_active", False)):
                status_text = fall_status if fall_status else "🚨 疑似跌倒｜等待長者確認"
                right_hand_detected = False
            else:
                status_text = "偵測中..."
                right_hand_detected = False

            box_width = np.clip(
                float(self.settings.get("camera_box_width", 40)), 10, 100
            ) / 100
            box_height = np.clip(
                float(self.settings.get("camera_box_height", 40)), 10, 100
            ) / 100
            box_center_x = 0.5 + np.clip(
                float(self.settings.get("camera_box_x_offset", 0)), -50, 50
            ) / 100
            box_center_y = 0.5 - np.clip(
                float(self.settings.get("camera_box_y_offset", 0)), -50, 50
            ) / 100
            box_center_x = np.clip(
                box_center_x, box_width / 2, 1 - box_width / 2
            )
            box_center_y = np.clip(
                box_center_y, box_height / 2, 1 - box_height / 2
            )
            box_x_start = box_center_x - box_width / 2
            box_x_end = box_center_x + box_width / 2
            box_y_start = box_center_y - box_height / 2
            box_y_end = box_center_y + box_height / 2
            cv2.rectangle(
                frame,
                (int(box_x_start * fw), int(box_y_start * fh)),
                (int(box_x_end * fw), int(box_y_end * fh)),
                (255, 0, 0),
                2,
            )

            current_trail_color = (255, 200, 100)  # BGR 格式

            # 只有挑戰進行中才獨佔手勢辨識；開始前及結束後恢復一般控制。
            gesture_game_mode = self._process_gesture_game(results)
            gesture_game_blocked = gesture_game_mode
            system_gestures_blocked = (
                gesture_game_blocked or self.modal_gesture_suspended
            )

            # 遊戲模式不使用一般滑鼠控制分支，但仍要在相機畫面顯示雙手節點。
            if gesture_game_mode and results.multi_hand_landmarks:
                for hand_lms in results.multi_hand_landmarks:
                    mp.solutions.drawing_utils.draw_landmarks(
                        frame,
                        hand_lms,
                        self.mp_hands.HAND_CONNECTIONS,
                        mp.solutions.drawing_utils.DrawingSpec(
                            color=(0, 255, 0), thickness=2, circle_radius=3
                        ),
                        mp.solutions.drawing_utils.DrawingSpec(
                            color=(255, 160, 0), thickness=2
                        ),
                    )
                    index_tip = hand_lms.landmark[INDEX]
                    cx = int(index_tip.x * fw)
                    cy = int(index_tip.y * fh)
                    cv2.circle(frame, (cx, cy), 13, (0, 0, 255), -1)
                    cv2.circle(frame, (cx, cy), 6, (255, 255, 255), -1)

            # =========================================================
            # 🆘 雙手高舉 SOS 偵測邏輯
            # =========================================================
            if not system_gestures_blocked and results.multi_hand_landmarks and len(results.multi_hand_landmarks) == 2:
                hands_up_count = 0
                for hand_landmarks in results.multi_hand_landmarks:
                    wrist_y = hand_landmarks.landmark[0].y
                    index_tip_y = hand_landmarks.landmark[INDEX].y
                    middle_tip_y = hand_landmarks.landmark[MIDDLE].y
                    
                    # 判斷食指與中指是否顯著高於手腕 (y 軸越小代表越高)
                    if index_tip_y < (wrist_y - 0.15) and middle_tip_y < (wrist_y - 0.15):
                        hands_up_count += 1

                # 雙手皆滿足高舉條件
                if hands_up_count == 2:
                    self.sos_counter += 1
                    if self.sos_counter >= self.SOS_FRAMES_REQUIRED:
                        now = time.time()
                        if now - self.last_sos_time > self.SOS_COOLDOWN:
                            self.last_sos_time = now
                            self.sos_counter = 0
                            self.trigger_sos_alert()  # 觸發警報與 Email
                else:
                    self.sos_counter = max(0, self.sos_counter - 1)
            else:
                self.sos_counter = max(0, self.sos_counter - 1)
            # =========================================================

            fall_alert_active = fall_detected or (
                hasattr(self, "fall_detector")
                and self.fall_detector is not None
                and getattr(self.fall_detector, "alert_active", False)
            )
            detected_hands = {}
            if results.multi_hand_landmarks and results.multi_handedness:
                for hand_lms, handedness in zip(
                    results.multi_hand_landmarks, results.multi_handedness
                ):
                    side = handedness.classification[0].label
                    detected_hands[side] = hand_lms

            if not system_gestures_blocked and not fall_alert_active:
                right_hand = detected_hands.get("Right")
                left_hand = detected_hands.get("Left")

                if right_hand is not None and not self.is_paused:
                    hand_lms = right_hand
                    right_hand_detected = True
                    status_text = "右手控制中"
                    index_tip = hand_lms.landmark[INDEX]

                    speed_mult = float(self.settings.get("speed_multiplier", 1.5))
                    normalized_x = np.clip(
                        (index_tip.x - box_x_start) / (box_x_end - box_x_start),
                        0,
                        1,
                    )
                    normalized_y = np.clip(
                        (index_tip.y - box_y_start) / (box_y_end - box_y_start),
                        0,
                        1,
                    )
                    raw_x = int(
                        (normalized_x * screen_w - screen_w / 2) * speed_mult
                        + screen_w / 2
                    )
                    raw_y = int(
                        (normalized_y * screen_h - screen_h / 2) * speed_mult
                        + screen_h / 2
                    )
                    tgt_x = max(2, min(screen_w - 2, raw_x))
                    tgt_y = max(2, min(screen_h - 2, raw_y))

                    self.pos_buffer.append((tgt_x, tgt_y))
                    final_x = int(np.mean([p[0] for p in self.pos_buffer]))
                    final_y = int(np.mean([p[1] for p in self.pos_buffer]))
                    pyautogui.moveTo(final_x, final_y, duration=0)

                    finger_states = {
                        tip: hand_lms.landmark[tip].y
                        < hand_lms.landmark[tip - 2].y
                        for tip in (INDEX, MIDDLE, RING, PINKY)
                    }
                    is_right_click = (
                        finger_states[INDEX]
                        and finger_states[MIDDLE]
                    )
                    if is_right_click:
                        if not self.right_click_gesture_active:
                            self.dwell_start_time = time.monotonic()
                            self.right_click_gesture_fired = False
                        self.right_click_gesture_active = True
                        elapsed = time.monotonic() - self.dwell_start_time
                        if (
                            not self.right_click_gesture_fired
                            and elapsed >= float(
                                self.settings.get("dwell_threshold", 1)
                            )
                        ):
                            pyautogui.click(button="right")
                            self.play_system_sound("click")
                            self.right_click_gesture_fired = True
                            self._post_camera_ui(
                                lambda: self.log_status(
                                    "智慧醫護：偵測到右手兩指停留，已執行右鍵。"
                                )
                            )
                        status_text = "兩指停留執行右鍵"
                    else:
                        if self.right_click_gesture_active:
                            self.dwell_start_time = None
                        self.right_click_gesture_active = False
                        self.right_click_gesture_fired = False
                        dist_moved = np.hypot(
                            final_x - self.last_cursor_pos[0],
                            final_y - self.last_cursor_pos[1],
                        )
                        if dist_moved < 18:
                            current_trail_color = (0, 127, 255)
                            if self.dwell_start_time is None:
                                self.dwell_start_time = time.time()
                            elif time.time() - self.dwell_start_time > float(
                                self.settings.get("dwell_threshold", 1)
                            ):
                                pyautogui.click(button="left")
                                self.play_system_sound("click")
                                self._post_camera_ui(
                                    lambda: self.log_status(
                                        "智慧醫護：偵測到長者停留，已自動執行點選！"
                                    )
                                )
                                self.dwell_start_time = time.time()
                        else:
                            current_trail_color = (255, 200, 100)
                            self.dwell_start_time = time.time()
                    self.last_cursor_pos = (final_x, final_y)

                    cx_raw = int(index_tip.x * fw)
                    cy_raw = int(index_tip.y * fh)
                    if hasattr(self, "trail_points") and self.settings.get(
                        "enable_trail", True
                    ):
                        self.trail_points.appendleft((cx_raw, cy_raw))
                else:
                    self.right_click_gesture_active = False
                    self.right_click_gesture_fired = False

                if left_hand is not None:
                    hand_lms = left_hand
                    status_text = "左手快捷觸發"
                    up_fingers = 0
                    if hand_lms.landmark[THUMB].x < hand_lms.landmark[THUMB - 1].x:
                        up_fingers += 1
                    for tip in (INDEX, MIDDLE, RING, PINKY):
                        if hand_lms.landmark[tip].y < hand_lms.landmark[tip - 2].y:
                            up_fingers += 1

                    if up_fingers == 5 and not self.is_paused:
                        self.is_paused = True
                        self.play_system_sound("warn")
                        self._post_camera_ui(
                            lambda: self.log_status(
                                "智慧醫護：偵測到左手全手掌，系統已暫停偵測。"
                            )
                        )
                        self.tts_speak("系統已暫停")
                    elif up_fingers == 3 and self.is_paused:
                        self.is_paused = False
                        self.play_system_sound("execute")
                        self._post_camera_ui(
                            lambda: self.log_status(
                                "智慧醫護：偵測到左手三指，系統已恢復偵測。"
                            )
                        )
                        self.tts_speak("系統已恢復")

                for hand_lms in detected_hands.values():
                    mp.solutions.drawing_utils.draw_landmarks(
                        frame, hand_lms, self.mp_hands.HAND_CONNECTIONS
                    )
                    index_tip_lms = hand_lms.landmark[INDEX]
                    cx, cy = int(index_tip_lms.x * fw), int(index_tip_lms.y * fh)
                    cv2.circle(frame, (cx, cy), 12, (0, 0, 255), -1)
                    cv2.circle(frame, (cx, cy), 6, (255, 255, 255), -1)
            else:
                self.right_click_gesture_active = False

            if not right_hand_detected:
                self.dwell_start_time = None
                if hasattr(self, 'trail_points') and len(self.trail_points) > 0:
                    self.trail_points.pop()

            # ➕ 新增：繪製動態發光淡出路徑
            if hasattr(self, 'trail_points') and self.settings.get("enable_trail", True) and len(self.trail_points) > 1:
                overlay = frame.copy()
                
                current_maxlen = self.trail_points.maxlen if self.trail_points.maxlen else 20
                
                for i in range(len(self.trail_points) - 1):
                    point_alpha = max(0, 1 - (i / current_maxlen))
                    thickness = int(max(1, point_alpha * 8))
                    
                    cv2.line(overlay, self.trail_points[i], self.trail_points[i+1], current_trail_color, thickness)
                
                cv2.addWeighted(overlay, 0.7, frame, 0.3, 0, frame)

            if (
                self.is_paused
                and not gesture_game_blocked
                and not (
                    fall_detected
                    or (
                        hasattr(self, "fall_detector")
                        and self.fall_detector is not None
                        and getattr(self.fall_detector, "alert_active", False)
                    )
                )
            ):
                status_text = self.localize_runtime_text("系統已暫停")

            # 1. 將 BGR 轉 PIL 繪製中文
            pil_frame = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
            draw = ImageDraw.Draw(pil_frame)

            # 2. 若沒有跌倒警報，才顯示一般的手勢狀態文字；讓跌倒警報覆蓋層維持最前層
            if not (fall_detected or (hasattr(self, "fall_detector") and self.fall_detector is not None and getattr(self.fall_detector, "alert_active", False))):
                draw.text((20, 20), status_text, font=self.main_font, fill=(0, 255, 0))
            
            # 3. 轉回 OpenCV BGR 格式
            frame = cv2.cvtColor(np.array(pil_frame), cv2.COLOR_RGB2BGR)
            with self.camera_frame_lock:
                self.latest_camera_frame = frame.copy()
                self.latest_camera_preview_frame = preview_frame

            time.sleep(0.001)

    # =====================
    # 語音控制與 AI 問答 (Gemini)
    # =====================
    def init_speech_control(self):
        self.recognizer = sr.Recognizer()
        self.recognizer.dynamic_energy_threshold = True

        self.voice_commands = {
            "做復健運動": self.show_rehab_video_picker,
            "手指健康操": self.trigger_finger_gesture_game,
            "看懷舊風景": lambda: self.execute_medical_search("台灣懷舊風景 4K 放鬆影片"),
            "衛教健康網": lambda: self.open_link("https://www.hpa.gov.tw/Pages/List.aspx?nodeid=366"),
            "記憶小遊戲": lambda: self.trigger_chimp_game(),
            "照護助理": lambda: self.trigger_ai_chat(),
            "系統參數": lambda: self.open_settings_window(),
            "相機視窗": self.show_camera_window,
            "關閉程式": self.safe_exit,
            "今日天氣": lambda: self.speak_today_weather(),
            "指令說明": lambda: self.show_voice_commands_dialog(), 
        }
    def toggle_microphone(self):
        """切換麥克風狀態（主動點擊觸發，防誤說防背景雜音）"""
        if not getattr(self, "is_listening", False):
            self.is_listening = True
            self.play_system_sound("execute")
            self.log_status("麥克風已開啟，請開始說話...")
            
            # 變更為高對比醒目紅（提示正在錄音）
            self.transition_text(
                self.mic_button,
                "🛑 正在聽...講完自動關",
                configure=lambda: self.mic_button.configure(
                    bg="#DC2626", activebackground="#B91C1C"
                ),
            )
            
            # 🔗 直接對接你原本改好的語音函式
            self.start_voice_listening()
        else:
            # 如果長者沒講完想手動提早關閉，再點一下也可以關
            self.stop_microphone_listening()

    def stop_microphone_listening(self):
        """關閉麥克風，回復到安全防誤說狀態"""
        self.is_listening = False
        self.log_status("麥克風已關閉。")
        if hasattr(self, "mic_button") and self.mic_button and self.mic_button.winfo_exists():
            # 恢復成溫和的安全灰色狀態
            self.transition_text(
                self.mic_button,
                "🎤 點我說話 (關閉中)",
                configure=lambda: self.mic_button.configure(
                    bg="#64748B", activebackground="#475569"
                ),
            )
    def start_voice_listening(self):
        """改造成『點擊觸發式』語音監聽：只在點擊後精準聽取一次，徹底防誤說"""
        def listen_loop():
            pythoncom.CoInitialize()
            
            # 💡 防誤說核心：只有在按鈕被點擊、self.is_listening 為 True 時才真正開啟麥克風
            if not getattr(self, "is_listening", False):
                return

            with sr.Microphone(
                device_index=self.settings.get("microphone_index")
            ) as source:
                # 稍微降低環境噪音影響（無障礙優化：縮短到 0.5 秒讓長者不用等太久）
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                try:
                    # 開始聽取語音
                    audio = self.recognizer.listen(source, timeout=4, phrase_time_limit=6)
                    audio = self.apply_microphone_gain(audio)
                    text = self.recognizer.recognize_google(audio, language="zh-TW")
                    
                    # 🚀 辨識成功：派發指令（例如：簽到、看相簿）
                    self.process_voice_text(text)
                except sr.WaitTimeoutError:
                    self.log_status("語音聽取逾時（長者太久沒說話）")
                except sr.UnknownValueError:
                    self.log_status("沒聽清楚，請再試一次")
                except Exception as e:
                    self.log_status(f"語音辨識出錯: {e}")
                finally:
                    # 🔒 關鍵防誤說：不管成功或失敗，這句話講完就『自動關閉麥克風』並把按鈕恢復原狀
                    if self.root and self.root.winfo_exists():
                        self.root.after(0, self.stop_microphone_listening)

        # 啟動單次聽取執行緒，完全不卡死 UI
        threading.Thread(target=listen_loop, daemon=True).start()

    def apply_microphone_gain(self, audio):
        gain = float(self.settings.get("microphone_gain", 1.0))
        if gain == 1.0:
            return audio
        raw_audio = audio.get_raw_data(convert_width=2)
        amplified_audio = audioop.mul(raw_audio, 2, gain)
        return sr.AudioData(
            amplified_audio, audio.sample_rate, sample_width=2
        )

    def process_voice_text(self, text):
        """處理聽到的話：先比對快捷指令，不符合則送去問 Gemini"""
        self.log_status(f"聽到了：{text}")

        help_keywords = ["指令", "說明", "可以說什麼", "幫忙", "教我"]
        if any(keyword in text for keyword in help_keywords):
            self.root.after(0, self.show_voice_commands_dialog)
            return
        # 1. 快捷關鍵字精準匹配
        for cmd_key, func in self.voice_commands.items():
            if cmd_key in text:
                self.root.after(0, func)
                return

    def trigger_ai_chat(self):
        """手動或自動點擊九宮格「AI照護助理」的動作：彈出專屬溫馨對話視窗並開啟防誤說麥克風"""
        lang = self.settings.get("language", "國語 (華語)")
        lang_dict = I18N.get(lang, I18N["國語 (華語)"])

        # 根據語系預設語音提示與對話開場白
        tts_welcome_map = {
            "國語 (華語)": "為您打開 AI 照護助理，想聊天時請點擊下方的麥克風按鈕。",
            "台文 (臺灣閩南語)": "幫汝開啟 AI 照護助理，想欲開講請點下底的麥克風按鈕。",
            "客語 (臺灣客家語)": "幫您打開 AI 照護助理，想愛開講請點下高嘅麥克風按鈕。",
            "English": "The AI care assistant is open. Click the microphone button below when you want to talk."
        }
        greeting_map = {
            "國語 (華語)": "護理助理：您好！我是您的健康小秘書。今天身體感覺怎麼樣？想和我聊聊什麼呢？\n\n",
            "台文 (臺灣閩南語)": "護理助理：汝好！我是汝的健康小秘書。今仔日身體感覺怎樣？想欲跟我開講啥物咧？\n\n",
            "客語 (臺灣客家語)": "護理助理：您好！我是您的健康小秘書。今晡日身體感覺按仔細？想愛跟我聊講哪細項？\n\n",
            "English": "Care assistant: Hello! I am your health assistant. How are you feeling today? What would you like to talk about?\n\n"
        }

        self.log_status("AI 照護助理已啟動。")
        self.play_system_sound("execute")
        self.tts_speak(tts_welcome_map.get(lang, tts_welcome_map["國語 (華語)"]))

        # 1. 建立獨立的 AI 對話頂層視窗
        ai_win = tk.Toplevel(self.root)
        ai_win.title(self.localize_runtime_text("AI照護助理"))
        ai_win.geometry("800x650")
        ai_win.configure(bg="#F0FDF4") 
        ai_win.grab_set() 

        # 2. 標題列
        lbl_title = tk.Label(
            ai_win, text="🤗 溫馨 AI 照護助理", 
            font=("Microsoft JhengHei", 26, "bold"), bg="#F0FDF4", fg="#166534"
        )
        lbl_title.pack(pady=15)

        # 底部提示文字
        lbl_status = tk.Label(
            ai_win, text="提示：可輸入文字送出，或點擊麥克風開始說話。",
            font=("Microsoft JhengHei", 14, "bold"), bg="#F0FDF4", fg="#15803D"
        )
        lbl_status.pack(side="bottom", pady=(5, 10))

        # 底部動作按鈕容器
        btn_container = tk.Frame(ai_win, bg="#F0FDF4")
        btn_container.pack(side="bottom", pady=15, padx=30, fill="x")

        # 3. 中間對話歷史紀錄顯示區
        chat_history = scrolledtext.ScrolledText(
            ai_win, font=("Microsoft JhengHei", 16), wrap=tk.WORD,
            bg="white", fg="#1E293B", bd=2, relief="solid",
            height=10
        )
        
        chat_history.insert(tk.END, greeting_map.get(lang, greeting_map["國語 (華語)"]))
        chat_history.configure(state="disabled")
        chat_history.see(tk.END)
    
        # 安全更新對話框內部函式
        def update_ui_text(user_text, ai_text=None):
            chat_history.configure(state="normal")
            if user_text:
                speaker_label = "You" if self.settings.get("language") == "English" else "您"
                chat_history.insert(tk.END, f"{speaker_label}: {user_text}\n")
            if ai_text:
                speaker_label = "Care assistant" if self.settings.get("language") == "English" else "護理助理"
                chat_history.insert(tk.END, f"{speaker_label}: {ai_text}\n\n")
            chat_history.configure(state="disabled")
            chat_history.see(tk.END)

        is_ai_responding = False

        def finish_ai_response():
            nonlocal is_ai_responding
            is_ai_responding = False
            if ai_win.winfo_exists():
                text_entry.configure(state="normal")
                send_btn.configure(state="normal")

        def submit_ai_message(query_text):
            nonlocal is_ai_responding
            query_text = query_text.strip()
            if not query_text:
                return
            if is_ai_responding:
                lbl_status.configure(
                    text=self.localize_runtime_text("請先等待助理回覆完成。"),
                    fg="#EF4444",
                )
                return

            is_ai_responding = True
            update_ui_text(user_text=query_text)
            text_entry.delete(0, tk.END)
            text_entry.configure(state="disabled")
            send_btn.configure(state="disabled")
            self.transition_text(
                lbl_status,
                self.localize_runtime_text("🤖 助理正在思考回應..."),
                configure=lambda: lbl_status.configure(fg="#9333EA"),
            )
            self.get_gemini_response_for_ui(
                query_text, ai_win, lbl_status, update_ui_text,
                on_complete=finish_ai_response,
            )

        text_input_frame = tk.Frame(ai_win, bg="#F0FDF4")
        text_input_frame.pack(side="bottom", fill="x", padx=30, pady=(5, 0))

        text_entry = tk.Entry(
            text_input_frame, font=("Microsoft JhengHei", 16),
            bg="white", fg="#1E293B", bd=2, relief="solid",
        )
        text_entry.pack(side="left", expand=True, fill="x", ipady=8, padx=(0, 10))
        text_entry.bind("<Return>", lambda _event: submit_ai_message(text_entry.get()))

        send_btn = tk.Button(
            text_input_frame, text="送出", font=("Microsoft JhengHei", 14, "bold"),
            bg="#16A34A", fg="white", activebackground="#15803D",
            activeforeground="white", bd=2, padx=18, pady=6,
            command=lambda: submit_ai_message(text_entry.get()), cursor="hand2",
        )
        send_btn.pack(side="right")

        chat_history.pack(fill="both", expand=True, padx=30, pady=10)

        # 4. 單次防誤說語音輸入控制
        self.is_ai_chat_listening = False  # 控制當前 AI 視窗麥克風的開關狀態

        def stop_ai_chat_mic():
            """安全回復麥克風按鈕外觀狀態"""
            self.is_ai_chat_listening = False
            if ai_win.winfo_exists():
                btn_mic.configure(text="🎤 點我說話 (關閉中)", bg="#64748B", activebackground="#475569")
                lbl_status.configure(text="提示：請點擊下方麥克風按鈕開始說話...", fg="#15803D")

        def single_voice_listen_loop():
            """執行緒：觸發單次精準聽取，聽完自動關閉"""
            pythoncom.CoInitialize()
            
            recognizer = sr.Recognizer()
            recognizer.energy_threshold = 3000
            recognizer.dynamic_energy_threshold = True

            try:
                with sr.Microphone(
                    device_index=self.settings.get("microphone_index")
                ) as source:
                    if not ai_win.winfo_exists() or not self.is_ai_chat_listening: return
                    
                    recognizer.adjust_for_ambient_noise(source, duration=0.5)
                    
                    if not ai_win.winfo_exists() or not self.is_ai_chat_listening: return
                    ai_win.after(0, lambda: self.transition_text(
                        lbl_status,
                        self.localize_runtime_text("🎙️ 正在傾聽中... 請說話..."),
                        configure=lambda: lbl_status.configure(fg="#EA580C"),
                    ))
                    
                    audio = recognizer.listen(source, timeout=5, phrase_time_limit=6)
                    
                    if not ai_win.winfo_exists() or not self.is_ai_chat_listening: return
                    ai_win.after(0, lambda: self.transition_text(
                        lbl_status,
                        self.localize_runtime_text("⚙️ 正在辨識您的語音..."),
                        configure=lambda: lbl_status.configure(fg="#2563EB"),
                    ))
                    audio = self.apply_microphone_gain(audio)
                    query = recognizer.recognize_google(audio, language="zh-TW")
                    
                    if query.strip() and ai_win.winfo_exists():
                        ai_win.after(0, lambda q=query: submit_ai_message(q))
                        
            except (sr.WaitTimeoutError, sr.UnknownValueError):
                if ai_win.winfo_exists():
                    ai_win.after(0, lambda: self.transition_text(
                        lbl_status,
                        self.localize_runtime_text("💡 沒聽清楚或逾時，請重新點擊再試一次"),
                        configure=lambda: lbl_status.configure(fg="#EF4444"),
                    ))
            except Exception as e:
                print(f"【AI語音助理錯誤】: {e}")
            finally:
                if ai_win.winfo_exists():
                    ai_win.after(0, stop_ai_chat_mic)

        def toggle_ai_chat_mic():
            """點擊麥克風按鈕的切換開關"""
            if not self.is_ai_chat_listening:
                self.is_ai_chat_listening = True
                self.play_system_sound("execute")
                self.transition_text(
                    btn_mic,
                    "🛑 正在聽... 講完自動關",
                    configure=lambda: btn_mic.configure(
                        bg="#DC2626", activebackground="#B91C1C"
                    ),
                )
                threading.Thread(target=single_voice_listen_loop, daemon=True).start()
            else:
                stop_ai_chat_mic()

        # 高無障礙防誤說大按鈕
        btn_mic = tk.Button(
            btn_container, text="🎤 點我說話 (關閉中)", font=("Microsoft JhengHei", 16, "bold"),
            bg="#64748B", fg="white", activebackground="#475569", activeforeground="white",
            bd=3, relief="raised", padx=20, pady=10, cursor="hand2",
            command=toggle_ai_chat_mic
        )
        btn_mic.pack(side="left", expand=True, fill="x", padx=(0, 10))

        # 5. 底部的關閉退出按鈕
        def close_ai_chat():
            self.is_ai_chat_listening = False
            close_tts_map = {
                "國語 (華語)": "好的，期待下次再和您聊天。",
                "台文 (臺灣閩南語)": "好，期待下一次閣跟汝開講。",
                "客語 (臺灣客家語)": "好，期待下一次再跟我開講。"
            }
            self.tts_speak(close_tts_map.get(lang, close_tts_map["國語 (華語)"]))
            ai_win.destroy()

        exit_btn = tk.Button(
            btn_container, text="❌ 聊天結束 ｜ 返回首頁", font=("Microsoft JhengHei", 16, "bold"),
            bg="#EF4444", fg="white", activebackground="#DC2626", activeforeground="white",
            bd=3, relief="raised", padx=20, pady=10, command=close_ai_chat, cursor="hand2"
        )
        exit_btn.pack(side="right", expand=True, fill="x", padx=(10, 0))

        ai_win.protocol("WM_DELETE_WINDOW", close_ai_chat)
        self.localize_window_widgets(ai_win)
        text_entry.focus_set()

    def get_gemini_response_for_ui(self, query_text, window, status_label, ui_callback, on_complete=None):
        """在背景執行緒中呼叫 Gemini API，並將產生的文本即時渲染到對話視窗"""
        def run_api():
            try:
                response = self.gemini_chat.send_message(message=query_text)
                answer = response.text
                
                if window.winfo_exists():
                    window.after(0, lambda: ui_callback(user_text=None, ai_text=answer))
                    window.after(0, lambda: status_label.configure(text=self.localize_runtime_text("🎙️ 回應完畢，請繼續說話..."), fg="#15803D"))
                
                self.play_system_sound("execute")
                self.tts_speak(answer)
                
            except Exception as e:
                print(f"【Gemini 3 API 錯誤】: {e}")
                if window.winfo_exists():
                    window.after(0, lambda: status_label.configure(text=self.localize_runtime_text("❌ 助理線路忙碌中，請再試一次"), fg="#EF4444"))
                self.play_system_sound("warn")
            finally:
                if on_complete and window.winfo_exists():
                    window.after(0, on_complete)

        threading.Thread(target=run_api, daemon=True).start()

    # =====================
    # 多媒體與系統整合功能
    # =====================
    def show_rehab_video_picker(self):
        """Collect exercise preferences, then request a general AI exercise assessment."""
        lang_dict = I18N.get(
            self.settings.get("language", "國語 (華語)"),
            I18N["國語 (華語)"],
        )
        language_name = self.settings.get("language", "國語 (華語)")
        picker = tk.Toplevel(self.root)
        picker.title(lang_dict.get("rehab_picker_title", "復健運動 AI 評估"))
        picker.geometry("920x760")
        picker.resizable(True, True)
        picker.configure(bg="#F8FAFC")
        picker.transient(self.root)
        picker.grab_set()
        self.center_window(picker, 920, 760)

        answers = {}
        content = tk.Frame(picker, bg="#F8FAFC")
        content.pack(fill="both", expand=True, padx=36, pady=30)
        question_steps = [
            (
                "rehab_question_area",
                [
                    ("head_neck", "rehab_head_neck"),
                    ("arms", "rehab_arms"),
                    ("shoulders", "rehab_shoulders"),
                    ("waist_back", "rehab_waist_back"),
                    ("legs", "rehab_legs"),
                ],
            ),
            (
                "rehab_question_position",
                [
                    ("seated", "rehab_position_seated"),
                    ("standing", "rehab_position_standing"),
                    ("supported", "rehab_position_supported"),
                ],
            ),
            (
                "rehab_question_discomfort",
                [
                    ("none", "rehab_discomfort_none"),
                    ("mild", "rehab_discomfort_mild"),
                    ("pain_or_restriction", "rehab_discomfort_pain"),
                ],
            ),
        ]
        area_search_terms = {
            "head_neck": "頸部",
            "arms": "手臂 上肢",
            "shoulders": "肩膀",
            "waist_back": "腰背",
            "legs": "下肢 腿部",
        }
        position_search_terms = {
            "seated": "坐姿",
            "standing": "站姿",
            "supported": "扶椅 坐姿",
        }

        def show_presets():
            for child in content.winfo_children():
                child.destroy()
            tk.Label(
                content,
                text=lang_dict.get("rehab_presets_title", "直接選擇影片類別"),
                font=("Microsoft JhengHei", 23, "bold"),
                bg="#F8FAFC",
                fg="#166534",
            ).pack(pady=(0, 8))
            tk.Label(
                content,
                text=lang_dict.get("rehab_presets_hint", "選擇部位即可查詢預設推薦影片。"),
                font=("Microsoft JhengHei", 14),
                bg="#F8FAFC",
                fg="#334155",
                wraplength=680,
            ).pack(pady=(0, 12))

            presets_frame = tk.Frame(content, bg="#F8FAFC")
            presets_frame.pack(fill="both", expand=True)
            preset_areas = [
                ("head_neck", "rehab_head_neck"),
                ("arms", "rehab_arms"),
                ("shoulders", "rehab_shoulders"),
                ("waist_back", "rehab_waist_back"),
                ("legs", "rehab_legs"),
                (None, "rehab_assess_from_presets"),
            ]
            for index, (area, label_key) in enumerate(preset_areas):
                row, column = divmod(index, 2)
                tk.Button(
                    presets_frame,
                    text=lang_dict.get(label_key, label_key),
                    font=("Microsoft JhengHei", 17, "bold"),
                    bg="#16A34A" if area is None else "#DCFCE7",
                    fg="white" if area is None else "#166534",
                    activebackground="#15803D" if area is None else "#BBF7D0",
                    activeforeground="white" if area is None else "#14532D",
                    bd=2,
                    relief="raised",
                    cursor="hand2",
                    command=lambda selected_area=area: (
                        show_question(0)
                        if selected_area is None
                        else self.execute_medical_search(
                            "長者 坐姿 "
                            + area_search_terms[selected_area]
                            + " 溫和安全運動 影片"
                        )
                    ),
                ).grid(
                    row=row,
                    column=column,
                    sticky="nsew",
                    padx=8,
                    pady=8,
                    ipady=10,
                )
            for column in range(2):
                presets_frame.columnconfigure(column, weight=1, uniform="rehab_preset")
            for row in range(3):
                presets_frame.rowconfigure(row, weight=1, uniform="rehab_preset")
            self.apply_font_size()

        def show_question(step_index=0):
            for child in content.winfo_children():
                child.destroy()
            if step_index >= len(question_steps):
                show_assessment()
                return

            question_key, choices = question_steps[step_index]
            tk.Label(
                content,
                text=lang_dict.get("rehab_picker_title", "復健運動 AI 評估"),
                font=("Microsoft JhengHei", 23, "bold"),
                bg="#F8FAFC",
                fg="#166534",
            ).pack(pady=(0, 6))
            tk.Label(
                content,
                text=lang_dict.get("rehab_step", "問題 {current} / {total}").format(
                    current=step_index + 1,
                    total=len(question_steps),
                ),
                font=("Microsoft JhengHei", 12, "bold"),
                bg="#F8FAFC",
                fg="#64748B",
            ).pack(pady=(0, 14))
            tk.Label(
                content,
                text=lang_dict.get(question_key, ""),
                font=("Microsoft JhengHei", 17, "bold"),
                bg="#F8FAFC",
                fg="#1E293B",
                wraplength=680,
            ).pack(pady=(0, 12))

            selected_choice = [""]
            choice_buttons = []
            options_frame = tk.Frame(content, bg="#F8FAFC")
            options_frame.pack(fill="x", expand=True, pady=(0, 14))
            for value, label_key in choices:
                choice_label = lang_dict.get(label_key, label_key)
                choice_button = tk.Button(
                    options_frame,
                    text=choice_label,
                    font=("Microsoft JhengHei", 16, "bold"),
                    bg="#E2E8F0",
                    activebackground="#CBD5E1",
                    fg="#334155",
                    activeforeground="#1E293B",
                    bd=2,
                    relief="raised",
                    cursor="hand2",
                    wraplength=670,
                    padx=16,
                    pady=12,
                )
                choice_button.pack(fill="x", pady=5, ipady=3)
                choice_buttons.append((value, choice_label, choice_button))

            def select_choice(value):
                selected_choice[0] = value
                for button_value, label, button in choice_buttons:
                    is_selected = button_value == value
                    button.configure(
                        text=f"✓  {label}" if is_selected else label,
                        bg="#86EFAC" if is_selected else "#E2E8F0",
                        activebackground="#4ADE80" if is_selected else "#CBD5E1",
                        fg="#14532D" if is_selected else "#334155",
                        relief="sunken" if is_selected else "raised",
                    )

            for value, _label, button in choice_buttons:
                button.configure(command=lambda selected=value: select_choice(selected))

            nav = tk.Frame(content, bg="#F8FAFC")
            nav.pack(fill="x", side="bottom")
            preset_nav = tk.Frame(content, bg="#F8FAFC")
            preset_nav.pack(fill="x", side="bottom")
            tk.Button(
                preset_nav,
                text=lang_dict.get("rehab_skip_questionnaire", "跳過問卷，直接選擇影片類別"),
                font=("Microsoft JhengHei", 12),
                command=show_presets,
            ).pack(side="left", padx=5, pady=(4, 0))
            if step_index > 0:
                tk.Button(
                    nav,
                    text=lang_dict.get("rehab_back", "Back"),
                    font=("Microsoft JhengHei", 13, "bold"),
                    command=lambda: show_question(step_index - 1),
                ).pack(side="left", padx=5, pady=8)

            def continue_assessment():
                selected = selected_choice[0]
                if not selected:
                    tk.Label(
                        nav,
                        text=lang_dict.get("rehab_select_answer", "請先選擇一個答案。"),
                        font=("Microsoft JhengHei", 11),
                        bg="#F8FAFC",
                        fg="#B91C1C",
                    ).pack(side="left", padx=8)
                    self.apply_font_size()
                    return
                answers[question_key] = selected
                show_question(step_index + 1)

            tk.Button(
                nav,
                text=lang_dict.get(
                    "rehab_next" if step_index < len(question_steps) - 1 else "rehab_assess",
                    "Next",
                ),
                font=("Microsoft JhengHei", 13, "bold"),
                bg="#16A34A",
                fg="white",
                activebackground="#15803D",
                activeforeground="white",
                padx=24,
                pady=8,
                cursor="hand2",
                command=continue_assessment,
            ).pack(side="right", padx=5, pady=8)
            self.apply_font_size()

        def show_assessment():
            for child in content.winfo_children():
                child.destroy()
            tk.Label(
                content,
                text=lang_dict.get("rehab_result_title", "AI 評估建議"),
                font=("Microsoft JhengHei", 23, "bold"),
                bg="#F8FAFC",
                fg="#166534",
            ).pack(pady=(0, 14))
            result_text = tk.Label(
                content,
                text=lang_dict.get("rehab_assessing", "AI 正在整理您的回答，請稍候..."),
                font=("Microsoft JhengHei", 15),
                bg="white",
                fg="#1E293B",
                justify="left",
                anchor="nw",
                wraplength=670,
                padx=18,
                pady=16,
                relief="solid",
                bd=1,
            )
            result_text.pack(fill="both", expand=True, pady=(0, 12))
            tk.Label(
                content,
                text=lang_dict.get("rehab_picker_safety", ""),
                font=("Microsoft JhengHei", 11),
                bg="#F8FAFC",
                fg="#B45309",
                wraplength=680,
            ).pack(pady=(0, 8))
            actions = tk.Frame(content, bg="#F8FAFC")
            actions.pack(fill="x")
            video_button = tk.Button(
                actions,
                text=lang_dict.get("rehab_search_videos", "Search related videos"),
                font=("Microsoft JhengHei", 13, "bold"),
                bg="#16A34A",
                fg="white",
                activebackground="#15803D",
                activeforeground="white",
                padx=18,
                pady=8,
                cursor="hand2",
                state="disabled",
                command=lambda: self.execute_medical_search(
                    "長者 "
                    + position_search_terms[answers["rehab_question_position"]]
                    + " "
                    + area_search_terms[answers["rehab_question_area"]]
                    + " 溫和安全運動 影片"
                ),
            )
            video_button.pack(side="left", padx=5)
            tk.Button(
                actions,
                text=lang_dict.get("rehab_back_to_questions", "Edit answers"),
                font=("Microsoft JhengHei", 12),
                command=lambda: show_question(0),
            ).pack(side="left", padx=5)
            tk.Button(
                actions,
                text=lang_dict.get("rehab_picker_close", "Close"),
                font=("Microsoft JhengHei", 12),
                command=picker.destroy,
            ).pack(side="right", padx=5)

            area_choices = dict(question_steps[0][1])
            position_choices = dict(question_steps[1][1])
            discomfort_choices = dict(question_steps[2][1])
            area = lang_dict.get(
                area_choices[answers["rehab_question_area"]],
                answers["rehab_question_area"],
            )
            position = lang_dict.get(
                position_choices[answers["rehab_question_position"]],
                answers["rehab_question_position"],
            )
            discomfort = lang_dict.get(
                discomfort_choices[answers["rehab_question_discomfort"]],
                answers["rehab_question_discomfort"],
            )
            prompt = (
                "你是長者健康運動資訊助理，不是醫師。請依據以下問答，以"
                f"{language_name}用親切、簡單的語言提供一般性運動影片挑選建議，"
                "限 100 字內。不要診斷、不要保證療效，也不要鼓勵忍痛運動。"
                "若有疼痛、活動受限或醫療限制，優先提醒停止造成不適的動作並詢問醫療專業人員；"
                "否則建議低強度、循序漸進、先從安全支撐的姿勢開始。"
                f"\n活動部位：{area}"
                f"\n運動姿勢偏好：{position}"
                f"\n疼痛或活動限制：{discomfort}"
            )

            def request_assessment():
                try:
                    for attempt in range(3):
                        try:
                            response = self.ai_client.models.generate_content(
                                model=self.model_id,
                                contents=prompt,
                            )
                            break
                        except Exception as exc:
                            is_temporarily_unavailable = (
                                getattr(exc, "code", None) == 503
                                or getattr(exc, "status", None) == "UNAVAILABLE"
                            )
                            if not is_temporarily_unavailable or attempt == 2:
                                raise
                            time.sleep(2 ** attempt)

                    assessment = (response.text or "").strip()
                    if not assessment:
                        raise ValueError("Gemini returned an empty assessment.")

                    def show_result():
                        if picker.winfo_exists():
                            result_text.configure(text=assessment)
                            video_button.configure(state="normal")

                    picker.after(0, show_result)
                    self.tts_speak(assessment)
                except Exception as exc:
                    print(f"【復健運動 AI 評估錯誤】: {exc}")

                    def show_error():
                        if picker.winfo_exists():
                            result_text.configure(
                                text=lang_dict.get(
                                    "rehab_assessment_error",
                                    "AI 評估暫時無法使用，請稍後再試。",
                                )
                            )

                    picker.after(0, show_error)

            threading.Thread(target=request_assessment, daemon=True).start()
            self.apply_font_size()

        show_question()

    def execute_medical_search(self, keyword):
        """自動幫長者開啟 YouTube 進行指定的復健、懷舊多媒體播放"""
        url = f"https://www.youtube.com/results?search_query={quote_plus(keyword)}"
        webbrowser.open(url)
        self.log_status(f"智慧醫護：正在開啟多媒體內容【{keyword}】")
        self.tts_speak("好的，現在帶您一起放鬆做運動")

    def open_link(self, url):
        """開啟外部無障礙遊戲或衛教網站"""
        webbrowser.open(url)
        self.log_status("正在為您開啟精彩內容，請直接使用右手操作。")
        self.tts_speak("好的，為您開啟網頁")

    def test_fall_alert(self):
        """Debug-only helper: force the fall warning flow to open for testing."""
        if not hasattr(self, "fall_detector") or self.fall_detector is None:
            self.fall_detector = FallDetector(self)

        self.play_system_sound("warn")
        self.tts_speak("測試跌倒警報，請確認畫面與語音是否正常。")
        self.log_status("🧪 正在測試跌倒警報：正在模擬疑似跌倒。")

        try:
            self.fall_detector.start_fall_confirmation(force=True)
        except Exception as e:
            print(f"【FallDetector】測試警報啟動失敗：{e}")

    def open_setup_wizard(self, on_complete=None):
        """首次啟動時，引導使用者完成基本身分、通知與無障礙偏好設定。"""
        lang_text = I18N.get(self.settings.get("language", "國語 (華語)"), I18N["國語 (華語)"])
        wizard = tk.Toplevel(self.root)
        wizard.title(lang_text.get("wizard_title", "首次使用設定"))
        wizard.geometry("720x600")
        self.center_window(wizard, 720, 600)
        wizard.resizable(True, True)
        wizard.configure(bg="#F8FAFC")
        wizard.transient(self.root)
        wizard.grab_set()

        selected_language = tk.StringVar(value=self.settings.get("language", "國語 (華語)"))
        user_name = tk.StringVar(value=self.settings.get("user_name", ""))
        receiver_email = tk.StringVar(value=self.settings.get("receiver_email", ""))
        email_enabled = tk.BooleanVar(value=self.settings.get("email_enabled", True))
        tts_enabled = tk.BooleanVar(value=self.settings.get("tts_enabled", True))
        fall_enabled = tk.BooleanVar(value=self.settings.get("fall_detection_enabled", True))
        earthquake_enabled = tk.BooleanVar(value=self.settings.get("earthquake_enabled", True))
        theme_mode = tk.StringVar(value=self.settings.get("theme_mode", "自動 (隨時間)"))
        ui_style_mode = tk.StringVar(value=self.settings.get("ui_style_mode", "精簡模式 (純純九宮格)"))
        target_county = tk.StringVar(value=self.settings.get("target_county", "臺中市"))
        theme_display = tk.StringVar()
        ui_style_display = tk.StringVar()
        county_display = tk.StringVar()
        theme_values = ["自動 (隨時間)", "明亮模式", "黑暗模式"]
        ui_style_values = ["精簡模式 (純純九宮格)", "焦點媒體模式 (專注看影片)"]
        county_values = ["臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市", "基隆市",
                         "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣", "南投縣", "雲林縣",
                         "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣"]

        def display_value(value, lang_key):
            if lang_key == "English":
                return ENGLISH_RUNTIME_TRANSLATIONS.get(value, value)
            return value

        theme_display.set(display_value(theme_mode.get(), self.settings.get("language", "國語 (華語)")))
        ui_style_display.set(display_value(ui_style_mode.get(), self.settings.get("language", "國語 (華語)")))
        county_display.set(display_value(target_county.get(), self.settings.get("language", "國語 (華語)")))
        page_index = [0]
        pages = []

        header = tk.Frame(wizard, bg="#1D4ED8", height=104)
        header.pack(fill="x")
        header.pack_propagate(False)
        title_label = tk.Label(
            header, text=lang_text.get("wizard_welcome", "歡迎使用智慧照護系統"), font=("Microsoft JhengHei", 22, "bold"),
            bg="#1D4ED8", fg="white"
        )
        title_label.pack(anchor="w", padx=28, pady=(18, 2))
        progress_label = tk.Label(
            header, text=lang_text.get("wizard_progress", "第 1 步，共 3 步").format(step=1, total=3), font=("Microsoft JhengHei", 12),
            bg="#1D4ED8", fg="#DBEAFE"
        )
        progress_label.pack(anchor="w", padx=30)

        content = tk.Frame(wizard, bg="#F8FAFC")
        content.pack(fill="both", expand=True, padx=30, pady=18)
        action_bar = tk.Frame(wizard, bg="#F8FAFC")
        action_bar.pack(fill="x", padx=30, pady=(0, 20))

        def add_label(parent, text, size=13, bold=False):
            return tk.Label(
                parent, text=text, font=("Microsoft JhengHei", size, "bold" if bold else "normal"),
                bg="#F8FAFC", fg="#1E293B", anchor="w", justify="left"
            )

        def create_page():
            page = tk.Frame(content, bg="#F8FAFC")
            pages.append(page)
            return page

        page_one = create_page()
        add_label(page_one, lang_text.get("wizard_contact_intro", "先填寫照護聯絡資訊，之後仍可在設定選單修改。"), 14).pack(anchor="w", pady=(0, 22))
        add_label(page_one, lang_text.get("wizard_user_name", "使用者稱呼（選填）"), bold=True).pack(anchor="w")
        tk.Entry(page_one, textvariable=user_name, font=("Microsoft JhengHei", 14), width=34).pack(anchor="w", pady=(5, 18), ipady=5)

        email_verification_code = tk.StringVar(value="")
        email_verification_status = tk.StringVar(value="驗證碼有效時間：尚未寄送")
        email_row = tk.Frame(page_one, bg="#F8FAFC")
        email_row.pack(anchor="w", pady=(0, 6))
        add_label(email_row, lang_text.get("wizard_family_email", "家屬收件 Email"), bold=True).pack(side="left")

        def send_family_email_verification():
            target_email = receiver_email.get().strip()
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", target_email):
                messagebox.showerror("Email 格式錯誤", "請先輸入有效的家屬收件 Email。", parent=wizard)
                return
            verification_code = self.generate_email_verification_code()
            ok, msg = self.send_verification_code_email(target_email, verification_code)
            if not ok:
                messagebox.showerror("驗證碼發送失敗", msg, parent=wizard)
                return
            self.pending_email_verification = {
                "email": target_email,
                "code": verification_code,
                "expires_at": time.time() + 300,
                "verified": False,
            }
            email_verification_code.set("")
            update_email_verification_countdown()
            messagebox.showinfo("驗證碼已寄出", msg, parent=wizard)

        def update_email_verification_countdown():
            pending = getattr(self, "pending_email_verification", {})
            expires_at = float(pending.get("expires_at", 0) or 0)
            remaining = max(0, int(expires_at - time.time()))
            if not expires_at:
                email_verification_status.set("驗證碼有效時間：尚未寄送")
            elif remaining > 0:
                minutes, seconds = divmod(remaining, 60)
                email_verification_status.set(
                    f"驗證碼有效時間：{minutes:02d}:{seconds:02d}"
                )
                wizard.after(1000, update_email_verification_countdown)
            else:
                if pending:
                    pending["code"] = ""
                    pending["verified"] = False
                email_verification_status.set("驗證碼已失效，請重新寄送")

        def verify_family_email_code():
            target_email = receiver_email.get().strip()
            pending = getattr(self, "pending_email_verification", {})
            if not pending or target_email != str(pending.get("email", "")).strip():
                messagebox.showerror("尚未寄送驗證碼", "請先點擊「寄送驗證碼」後再驗證。", parent=wizard)
                return
            if time.time() > float(pending.get("expires_at", 0)):
                pending["code"] = ""
                pending["verified"] = False
                messagebox.showerror("驗證碼已過期", "請重新點擊「寄送驗證碼」取得新的驗證碼。", parent=wizard)
                return
            if email_verification_code.get().strip() != str(pending.get("code", "")).strip():
                messagebox.showerror("驗證碼錯誤", "驗證碼不正確，請重新輸入。", parent=wizard)
                return
            pending["verified"] = True
            messagebox.showinfo("驗證成功", "家屬收件信箱已驗證完成。", parent=wizard)

        entry_email_field = tk.Entry(page_one, textvariable=receiver_email, font=("Microsoft JhengHei", 14), width=34)
        entry_email_field.pack(anchor="w", pady=(5, 8), ipady=5)
        button_row = tk.Frame(page_one, bg="#F8FAFC")
        button_row.pack(anchor="w", pady=(0, 6))
        tk.Button(button_row, text="寄送驗證碼", font=("Microsoft JhengHei", 11, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", bd=0, relief="flat", cursor="hand2", command=send_family_email_verification).pack(side="left")
        tk.Button(button_row, text="驗證", font=("Microsoft JhengHei", 11, "bold"), bg="#10B981", fg="white", activebackground="#059669", activeforeground="white", bd=0, relief="flat", cursor="hand2", command=verify_family_email_code).pack(side="left", padx=(8, 0))
        verification_row = tk.Frame(page_one, bg="#F8FAFC")
        verification_row.pack(anchor="w", pady=(0, 12))
        tk.Label(verification_row, text="驗證碼：", font=("Microsoft JhengHei", 12, "bold"), bg="#F8FAFC", fg="#1E293B").pack(side="left")
        tk.Entry(verification_row, textvariable=email_verification_code, font=("Microsoft JhengHei", 14), width=18).pack(side="left", padx=(8, 0), ipady=4)
        tk.Label(
            page_one,
            textvariable=email_verification_status,
            font=("Microsoft JhengHei", 11),
            bg="#F8FAFC",
            fg="#64748B",
        ).pack(anchor="w", pady=(0, 12))

        add_label(page_one, lang_text.get("wizard_email_hint", "用於每日簽到、SOS 與跌倒警報通知。"), 11).pack(anchor="w", pady=(0, 18))
        tk.Checkbutton(
            page_one, text=lang_text.get("wizard_email_enabled", "啟用家屬 Email 通知"), variable=email_enabled,
            font=("Microsoft JhengHei", 13, "bold"), bg="#F8FAFC", fg="#1D4ED8",
            activebackground="#F8FAFC", selectcolor="#DBEAFE"
        ).pack(anchor="w")

        page_two = create_page()
        add_label(page_two, lang_text.get("wizard_preferences_intro", "選擇您習慣的語言與畫面操作方式。"), 14).pack(anchor="w", pady=(0, 22))
        add_label(page_two, lang_text.get("wizard_language", "系統語言"), bold=True).pack(anchor="w")
        ttk.Combobox(
            page_two, textvariable=selected_language,
            values=["國語 (華語)", "台文 (臺灣閩南語)", "客語 (臺灣客家語)", "English"],
            state="readonly", font=("Microsoft JhengHei", 13), width=28
        ).pack(anchor="w", pady=(5, 18))
        add_label(page_two, lang_text.get("wizard_theme", "畫面主題"), bold=True).pack(anchor="w")
        theme_combo = ttk.Combobox(
            page_two, textvariable=theme_display,
            values=[display_value(value, self.settings.get("language", "國語 (華語)")) for value in theme_values],
            state="readonly", font=("Microsoft JhengHei", 13), width=28
        )
        theme_combo.pack(anchor="w", pady=(5, 18))
        add_label(page_two, lang_text.get("wizard_ui_style", "主畫面樣式"), bold=True).pack(anchor="w")
        ui_style_combo = ttk.Combobox(
            page_two, textvariable=ui_style_display,
            values=[display_value(value, self.settings.get("language", "國語 (華語)")) for value in ui_style_values],
            state="readonly", font=("Microsoft JhengHei", 13), width=28
        )
        ui_style_combo.pack(anchor="w", pady=(5, 18))
        tk.Checkbutton(
            page_two, text=lang_text.get("wizard_tts_enabled", "啟用語音提示"), variable=tts_enabled,
            font=("Microsoft JhengHei", 13, "bold"), bg="#F8FAFC", fg="#1D4ED8",
            activebackground="#F8FAFC", selectcolor="#DBEAFE"
        ).pack(anchor="w")

        page_three = create_page()
        add_label(page_three, lang_text.get("wizard_safety_intro", "最後設定安全功能。建議保留警報功能開啟。"), 14).pack(anchor="w", pady=(0, 22))
        tk.Checkbutton(
            page_three, text=lang_text.get("wizard_fall_enabled", "啟用跌倒警報偵測"), variable=fall_enabled,
            font=("Microsoft JhengHei", 13, "bold"), bg="#F8FAFC", fg="#9A3412",
            activebackground="#F8FAFC", selectcolor="#FFEDD5"
        ).pack(anchor="w", pady=6)
        tk.Checkbutton(
            page_three, text=lang_text.get("wizard_earthquake_enabled", "啟用地震預警"), variable=earthquake_enabled,
            font=("Microsoft JhengHei", 13, "bold"), bg="#F8FAFC", fg="#9A3412",
            activebackground="#F8FAFC", selectcolor="#FFEDD5"
        ).pack(anchor="w", pady=6)
        add_label(page_three, lang_text.get("wizard_county", "居住監測地區"), bold=True).pack(anchor="w", pady=(22, 0))
        county_combo = ttk.Combobox(
            page_three, textvariable=county_display,
            values=[display_value(value, self.settings.get("language", "國語 (華語)")) for value in county_values],
            state="readonly", font=("Microsoft JhengHei", 13), width=28
        )
        county_combo.pack(anchor="w", pady=(5, 18))
        add_label(page_three, lang_text.get("wizard_finish_hint", "完成後會立即套用設定，之後可從「系統參數」再次調整。"), 11).pack(anchor="w", pady=(8, 0))

        def refresh_wizard_language(*args):
            language = selected_language.get()
            theme_combo.configure(values=[display_value(value, language) for value in theme_values])
            ui_style_combo.configure(values=[display_value(value, language) for value in ui_style_values])
            county_combo.configure(values=[display_value(value, language) for value in county_values])
            theme_display.set(display_value(theme_mode.get(), language))
            ui_style_display.set(display_value(ui_style_mode.get(), language))
            county_display.set(display_value(target_county.get(), language))

        selected_language.trace_add("write", refresh_wizard_language)

        def save_setup():
            email = receiver_email.get().strip()
            if email_enabled.get():
                if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                    messagebox.showerror("Email 格式錯誤", "請輸入有效的家屬 Email，或先取消勾選 Email 通知。", parent=wizard)
                    return
                pending = getattr(self, "pending_email_verification", {})
                if (
                    not pending.get("verified", False)
                    or email != str(pending.get("email", "")).strip()
                    or time.time() >= float(pending.get("expires_at", 0) or 0)
                ):
                    messagebox.showerror("Email 尚未驗證", "請先寄送並完成有效期限內的驗證碼驗證。", parent=wizard)
                    return
            selected_theme = dict(zip(
                [display_value(value, selected_language.get()) for value in theme_values], theme_values
            )).get(theme_display.get(), theme_display.get())
            selected_style = dict(zip(
                [display_value(value, selected_language.get()) for value in ui_style_values], ui_style_values
            )).get(ui_style_display.get(), ui_style_display.get())
            selected_county = dict(zip(
                [display_value(value, selected_language.get()) for value in county_values], county_values
            )).get(county_display.get(), county_display.get())
            self.settings.update({
                "user_name": user_name.get().strip(),
                "receiver_email": email,
                "email_enabled": bool(email_enabled.get()),
                "language": selected_language.get(),
                "theme_mode": selected_theme,
                "ui_style_mode": selected_style,
                "tts_enabled": bool(tts_enabled.get()),
                "fall_detection_enabled": bool(fall_enabled.get()),
                "earthquake_enabled": bool(earthquake_enabled.get()),
                "target_county": selected_county,
                "setup_completed": True,
            })
            self.current_style_mode = ui_style_mode.get()
            self.target_county = target_county.get()
            self.save_settings(self.settings)
            self.apply_language()
            self.check_and_apply_theme()
            wizard.grab_release()
            wizard.destroy()
            if on_complete is not None:
                on_complete()

        def close_setup():
            self.settings["setup_completed"] = True
            self.save_settings(self.settings)
            wizard.grab_release()
            wizard.destroy()
            if on_complete is not None:
                on_complete()

        def show_page(index):
            page_index[0] = index
            for page in pages:
                page.pack_forget()
            pages[index].pack(fill="both", expand=True)
            progress_label.configure(text=lang_text.get("wizard_progress", "第 {step} 步，共 {total} 步").format(step=index + 1, total=3))
            back_button.configure(state="normal" if index else "disabled")
            next_button.configure(
                text=lang_text.get("wizard_done", "完成") if index == len(pages) - 1
                else lang_text.get("wizard_next", "下一步")
            )

        back_button = tk.Button(
            action_bar, text=lang_text.get("wizard_back", "上一步"), font=("Microsoft JhengHei", 13, "bold"),
            bg="#CBD5E1", fg="#1E293B", bd=0, padx=18, pady=8,
            command=lambda: show_page(page_index[0] - 1)
        )
        back_button.pack(side="left")
        next_button = tk.Button(
            action_bar, text=lang_text.get("wizard_next", "下一步"), font=("Microsoft JhengHei", 13, "bold"),
            bg="#2563EB", fg="white", activebackground="#1D4ED8", bd=0, padx=24, pady=8,
            command=lambda: save_setup() if page_index[0] == len(pages) - 1 else show_page(page_index[0] + 1)
        )
        next_button.pack(side="right")
        tk.Button(
            action_bar, text=lang_text.get("wizard_later", "稍後設定"), font=("Microsoft JhengHei", 11),
            bg="#F8FAFC", fg="#64748B", bd=0, command=close_setup
        ).pack(side="right", padx=12)
        wizard.protocol("WM_DELETE_WINDOW", close_setup)
        show_page(0)

    def open_settings_window(self):
        """開啟高齡化無障礙設定選單：支援多語言 UI 切換、語音、滑鼠 speed/smooth、Email、軌跡與地震警報"""
        
        self.log_status("正在開啟系統參數設定選單...")
        self.play_system_sound("execute")
        
        curr_lang = self.settings.get("language", "國語 (華語)")
        original_language = curr_lang
        original_font_size = self.settings.get("font_size", "normal")
        font_size_saved = False
        language_saved = False
        self.tts_speak(I18N.get(curr_lang, {}) .get("title", "已為您打開設定選單"))

        settings_win = tk.Toplevel(self.root)
        settings_win.title(I18N.get(curr_lang, I18N["國語 (華語)"]).get("settings_title", "系統參數設定選單"))
        settings_win.geometry("820x720")
        settings_win.minsize(720, 560)
        settings_win.resizable(True, True)
        settings_win.configure(bg="#F1F5F9")
        settings_win.grab_set()

        title_lbl = tk.Label(
            settings_win, font=("Microsoft JhengHei", 20, "bold"),
            bg="#F1F5F9", fg="#1E293B", anchor="w", justify="left"
        )
        title_lbl.pack(fill="x", padx=20, pady=15)

        # ── Ribbon 分類頁籤 ──
        main_container = tk.Frame(settings_win, bg="#F1F5F9")
        main_container.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        tab_style = ttk.Style(settings_win)
        tab_style.configure(
            "Settings.TNotebook",
            background="#F1F5F9",
            borderwidth=0,
        )
        if not hasattr(self, "_settings_tab_font"):
            self._settings_tab_font = tkfont.Font(
                root=self.root,
                family="Microsoft JhengHei",
                size=12,
                weight="bold",
            )
        tab_style.configure(
            "Settings.TNotebook.Tab",
            font=self._settings_tab_font,
            padding=(14, 8),
        )
        notebook = ttk.Notebook(main_container, style="Settings.TNotebook")
        notebook.pack(fill="both", expand=True)
        tab_transition = {"active": False, "current": None}
        tab_background = settings_win.winfo_rgb("#F1F5F9")
        fade_color_options = (
            "bg", "fg", "activebackground", "activeforeground",
            "highlightbackground", "troughcolor", "selectcolor",
        )

        def animate_ribbon_transition(event=None):
            if tab_transition["active"] or not settings_win.winfo_exists():
                return
            target_tab = notebook.select()
            if target_tab == tab_transition["current"]:
                return

            tab_transition["active"] = True
            previous_tab = tab_transition["current"]
            notebook.select(previous_tab)
            notebook.state(["disabled"])

            steps = 15
            interval_ms = 16

            def eased_progress(step):
                progress = step / steps
                return progress * progress * (3 - 2 * progress)

            def collect_palette(tab_id):
                try:
                    tab = settings_win.nametowidget(tab_id)
                except KeyError as exc:
                    raise RuntimeError(f"找不到設定頁籤元件: {tab_id}") from exc
                palette = []

                def visit(widget):
                    try:
                        options = {}
                        for option in fade_color_options:
                            try:
                                value = widget.cget(option)
                                color = widget.winfo_rgb(value)
                            except tk.TclError:
                                continue
                            options[option] = color
                        if options:
                            palette.append((widget, options))
                    except tk.TclError:
                        return
                    for child in widget.winfo_children():
                        visit(child)

                visit(tab)
                return palette

            def apply_palette(palette, progress, fade_out):
                amount = eased_progress(progress)
                if not fade_out:
                    amount = 1 - amount
                for widget, options in palette:
                    try:
                        changes = {}
                        for option, original_color in options.items():
                            color = tuple(
                                round(channel + (neutral - channel) * amount)
                                for channel, neutral in zip(original_color, tab_background)
                            )
                            changes[option] = "#" + "".join(
                                f"{channel // 257:02x}" for channel in color
                            )
                        widget.configure(**changes)
                    except tk.TclError:
                        continue

            def restore_palette(palette):
                for widget, options in palette:
                    try:
                        widget.configure({
                            option: "#" + "".join(
                                f"{channel // 257:02x}" for channel in color
                            )
                            for option, color in options.items()
                        })
                    except tk.TclError:
                        continue

            try:
                outgoing_palette = collect_palette(previous_tab)
                incoming_palette = collect_palette(target_tab)
            except RuntimeError as exc:
                print(f"設定頁籤轉場初始化失敗: {exc}")
                notebook.select(target_tab)
                tab_transition["current"] = target_tab
                notebook.state(["!disabled"])
                tab_transition["active"] = False
                return

            def finish_transition():
                tab_transition["active"] = False
                notebook.state(["!disabled"])

            def fade_in(step):
                if not settings_win.winfo_exists():
                    return
                apply_palette(incoming_palette, step, fade_out=False)
                if step < steps:
                    settings_win.after(interval_ms, lambda: fade_in(step + 1))
                else:
                    finish_transition()

            def fade_out(step):
                if not settings_win.winfo_exists():
                    return
                apply_palette(outgoing_palette, step, fade_out=True)
                if step < steps:
                    settings_win.after(interval_ms, lambda: fade_out(step + 1))
                    return
                notebook.state(["!disabled"])
                notebook.select(target_tab)
                notebook.state(["disabled"])
                tab_transition["current"] = target_tab
                restore_palette(outgoing_palette)
                apply_palette(incoming_palette, steps, fade_out=True)
                settings_win.after(interval_ms, lambda: fade_in(0))

            fade_out(0)

        tab_labels = {
            "國語 (華語)": ("基本操作", "語音與通知", "語音&視訊", "個人化", "安全警報", "進階", "關於"),
            "台文 (臺灣閩南語)": ("基本操作", "語音與通知", "語音&視訊", "個人化", "安全警報", "進階", "關於"),
            "客語 (臺灣客家語)": ("基本操作", "語音與通知", "語音&視訊", "個人化", "安全警報", "進階", "關於"),
            "English": ("General", "Voice & Alerts", "Voice & Video", "Customization", "Safety", "Advanced", "About"),
        }
        initial_tab_labels = tab_labels.get(curr_lang, tab_labels["國語 (華語)"])
        basic_frame = tk.Frame(notebook, bg="#F1F5F9")
        voice_frame = tk.Frame(notebook, bg="#F1F5F9")
        audio_video_frame = tk.Frame(notebook, bg="#F1F5F9")
        visual_frame = tk.Frame(notebook, bg="#F1F5F9")
        safety_frame = tk.Frame(notebook, bg="#F1F5F9")
        tools_frame = tk.Frame(notebook, bg="#F1F5F9")
        about_frame = tk.Frame(notebook, bg="#F1F5F9")
        tab_frames = (
            basic_frame, voice_frame, audio_video_frame,
            visual_frame, safety_frame, tools_frame, about_frame,
        )
        tab_scroll_canvases = {}
        tab_content_frames = {}
        tab_content_window_ids = {}
        for frame, label in zip(tab_frames, initial_tab_labels):
            notebook.add(frame, text=label)
            canvas = tk.Canvas(
                frame, bg="#F1F5F9", highlightthickness=0, borderwidth=0
            )
            scrollbar = ttk.Scrollbar(
                frame, orient="vertical", command=canvas.yview
            )
            content_frame = tk.Frame(canvas, bg="#F1F5F9")
            content_window = canvas.create_window(
                (0, 0), window=content_frame, anchor="nw"
            )
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def update_scroll_region(event, scroll_canvas=canvas):
                scroll_canvas.configure(
                    scrollregion=scroll_canvas.bbox("all")
                )

            def resize_content(event, scroll_canvas=canvas, window_id=content_window):
                scroll_canvas.itemconfigure(window_id, width=event.width)
                scroll_canvas.coords(window_id, 0, 0)

            content_frame.bind("<Configure>", update_scroll_region)
            canvas.bind("<Configure>", resize_content)
            tab_scroll_canvases[str(frame)] = canvas
            tab_content_frames[frame] = content_frame
            tab_content_window_ids[str(frame)] = content_window
            settings_win.after_idle(canvas.yview_moveto, 0)
        about_content = tab_content_frames[about_frame]
        about_icon_photo = self._load_icon_photo(96)
        if about_icon_photo is not None:
            about_icon = tk.Label(
                about_content,
                image=about_icon_photo,
                bg="#F1F5F9",
                borderwidth=0,
            )
            about_icon.image = about_icon_photo
            about_icon.pack(pady=(32, 12))
        about_title = tk.Label(
            about_content,
            text="智慧多媒體互動平台",
            font=("Microsoft JhengHei", 20, "bold"),
            bg="#F1F5F9",
            fg="#1E293B",
            anchor="center",
        )
        about_title.pack(fill="x", pady=(0, 12))
        about_description = tk.Label(
            about_content,
            text="專為高齡者設計的直覺式互動平台",
            font=("Microsoft JhengHei", 12),
            bg="#F1F5F9",
            fg="#475569",
            anchor="center",
            justify="center",
            wraplength=600,
        )
        about_description.pack(fill="x", padx=24)
        about_details = tk.Label(
            about_content,
            text="",
            font=("Microsoft JhengHei", 12),
            bg="#F1F5F9",
            fg="#334155",
            anchor="center",
            justify="center",
            wraplength=650,
        )
        about_details.pack(fill="x", padx=24, pady=(28, 12))
        tab_transition["current"] = notebook.select()
        notebook.bind("<<NotebookTabChanged>>", animate_ribbon_transition, add="+")

        def align_selected_tab_content(event=None):
            selected_tab = notebook.select()

            def reset_selected_tab_scroll():
                if not settings_win.winfo_exists():
                    return
                selected_canvas = tab_scroll_canvases.get(selected_tab)
                if selected_canvas is None:
                    return
                selected_canvas.update_idletasks()
                selected_canvas.coords(
                    tab_content_window_ids[selected_tab], 0, 0
                )
                selected_canvas.configure(
                    scrollregion=selected_canvas.bbox("all")
                )
                selected_canvas.yview_moveto(0)

            settings_win.after_idle(reset_selected_tab_scroll)
            settings_win.after(280, reset_selected_tab_scroll)

        notebook.bind(
            "<<NotebookTabChanged>>", align_selected_tab_content, add="+"
        )

        form_frame = tab_content_frames[basic_frame]

        def on_mouse_wheel(event):
            selected_tab = notebook.select()
            scroll_canvas = tab_scroll_canvases.get(selected_tab)
            if scroll_canvas is None:
                return "break"
            if getattr(event, "num", None) == 4:
                units = -1
            elif getattr(event, "num", None) == 5:
                units = 1
            else:
                units = -int(event.delta / 120)
                if units == 0:
                    units = -1 if event.delta > 0 else 1
            scroll_canvas.yview_scroll(units, "units")
            return "break"

        def set_canvas_scroll_enabled(enabled):
            try:
                if enabled:
                    settings_win.bind_all("<MouseWheel>", on_mouse_wheel)
                    settings_win.bind_all("<Button-4>", on_mouse_wheel)
                    settings_win.bind_all("<Button-5>", on_mouse_wheel)
                else:
                    settings_win.unbind_all("<MouseWheel>")
                    settings_win.unbind_all("<Button-4>")
                    settings_win.unbind_all("<Button-5>")
            except Exception:
                pass

        set_canvas_scroll_enabled(True)

        def disable_scroll_on_combo_focus(event):
            set_canvas_scroll_enabled(False)

        def enable_scroll_after_combo_loses_focus(event):
            set_canvas_scroll_enabled(True)

        def on_window_close():
            stop_microphone_test()
            stop_camera_preview()
            if (
                not language_saved
                and self.settings.get("language") != original_language
            ):
                self.settings["language"] = original_language
                self.apply_language()
            if not font_size_saved:
                self.settings["font_size"] = original_font_size
                self.apply_font_size()
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                try:
                    settings_win.unbind_all(seq)
                except Exception:
                    pass
            settings_win.destroy()

        # 輔助滑動條點擊函式
        def click_to_set_value(event):
            widget = event.widget
            if isinstance(widget, tk.Scale) and widget.winfo_width() > 0:
                fraction = event.x / widget.winfo_width()
                val_range = widget.cget("to") - widget.cget("from")
                clicked_val = widget.cget("from") + (fraction * val_range)
                res = widget.cget("resolution")
                widget.set(max(widget.cget("from"), min(widget.cget("to"), round(clicked_val / res) * res)))

        # ── 介面元件宣吿 ──
        lbl_lang_sec = tk.Label(form_frame, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#475569")
        lbl_lang_sec.configure(anchor="w", justify="left")
        lbl_lang_sec.pack(anchor="w", pady=(10, 5))

        lang_card = tk.Frame(form_frame, bg="#FFFFFF", bd=1, relief="solid", padx=10, pady=5)
        lang_card.pack(fill="x", pady=(0, 10))

        selected_lang = tk.StringVar(value=curr_lang)
        languages = [
            ("國語 (華語)", "國語 (華語)"),
            ("台文 (臺灣閩南語)", "台文 (臺灣閩南語)"),
            ("客語 (臺灣客家語)", "客語 (臺灣客家語)"),
            ("English", "English"),
        ]

        # 1. 鼠標平滑度
        lbl_smooth = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_smooth.pack(anchor="w", pady=(10,0))
        val_smooth = tk.DoubleVar(value=self.settings.get("smooth_factor", 5.0))
        scale_smooth = tk.Scale(form_frame, from_=1.0, to=15.0, resolution=1.0, orient="horizontal", variable=val_smooth, bg="#FFFFFF", bd=0, highlightthickness=0)
        scale_smooth.pack(fill="x", pady=5)
        scale_smooth.bind("<Button-1>", click_to_set_value)

        # 2. 移動速度
        lbl_speed = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_speed.pack(anchor="w", pady=(10,0))
        val_speed = tk.DoubleVar(value=self.settings.get("speed_multiplier", 1.5))
        scale_speed = tk.Scale(form_frame, from_=0.5, to=3.5, resolution=0.1, orient="horizontal", variable=val_speed, bg="#FFFFFF", bd=0, highlightthickness=0)
        scale_speed.pack(fill="x", pady=5)
        scale_speed.bind("<Button-1>", click_to_set_value)

        # 3. 停留點擊
        lbl_dwell = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_dwell.pack(anchor="w", pady=(10,0))
        val_dwell = tk.DoubleVar(value=self.settings.get("dwell_threshold", 1))
        scale_dwell = tk.Scale(form_frame, from_=0.5, to=3.0, resolution=0.1, orient="horizontal", variable=val_dwell, bg="#FFFFFF", bd=0, highlightthickness=0)
        scale_dwell.pack(fill="x", pady=5)
        scale_dwell.bind("<Button-1>", click_to_set_value)

        camera_box_x_offset = tk.IntVar(
            value=self.settings.get("camera_box_x_offset", 0)
        )
        camera_box_y_offset = tk.IntVar(
            value=self.settings.get("camera_box_y_offset", 0)
        )
        camera_box_width = tk.IntVar(
            value=self.settings.get("camera_box_width", 40)
        )
        camera_box_height = tk.IntVar(
            value=self.settings.get("camera_box_height", 40)
        )
        camera_box_controls = (
            ("camera_box_x_offset_lbl", camera_box_x_offset, -50, 50),
            ("camera_box_y_offset_lbl", camera_box_y_offset, -50, 50),
            ("camera_box_width_lbl", camera_box_width, 10, 100),
            ("camera_box_height_lbl", camera_box_height, 10, 100),
        )
        camera_box_labels = {}
        for label_key, variable, minimum, maximum in camera_box_controls:
            label = tk.Label(
                form_frame,
                font=("Microsoft JhengHei", 12),
                bg="#F1F5F9",
                fg="#475569",
            )
            label.pack(anchor="w", pady=(10, 0))
            scale = tk.Scale(
                form_frame,
                from_=minimum,
                to=maximum,
                resolution=1,
                orient="horizontal",
                variable=variable,
                bg="#FFFFFF",
                bd=0,
                highlightthickness=0,
            )
            scale.pack(fill="x", pady=5)
            scale.bind("<Button-1>", click_to_set_value)
            camera_box_labels[label_key] = label

        form_frame = tab_content_frames[voice_frame]

        # 4. 語速
        lbl_vspeed = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_vspeed.pack(anchor="w", pady=(10,0))
        val_vspeed = tk.IntVar(value=self.settings.get("voice_speed", 0))
        scale_vspeed = tk.Scale(form_frame, from_=-5, to=5, resolution=1, orient="horizontal", variable=val_vspeed, bg="#FFFFFF", bd=0, highlightthickness=0)
        scale_vspeed.pack(fill="x", pady=5)
        scale_vspeed.bind("<Button-1>", click_to_set_value)

        # 5. TTS 開關
        lbl_voice_sec = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_voice_sec.pack(anchor="w", pady=(15,0))
        val_tts = tk.BooleanVar(value=self.settings.get("tts_enabled", True))
        chk_tts = tk.Checkbutton(form_frame, variable=val_tts, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#1E293B", bd=0)
        chk_tts.pack(anchor="w", pady=5)

        form_frame = tab_content_frames[audio_video_frame]

        def make_device_row(parent, label_text, values, selected_value):
            row = tk.Frame(parent, bg="#F1F5F9")
            row.pack(fill="x", pady=(8, 3))
            label = tk.Label(
                row, text=label_text, font=("Microsoft JhengHei", 12, "bold"),
                bg="#F1F5F9", fg="#475569",
            )
            label.pack(anchor="w")
            combo = ttk.Combobox(
                row, values=values, font=("Microsoft JhengHei", 11),
                state="readonly", width=60,
            )
            combo.pack(anchor="w", pady=(4, 0))
            combo.bind("<FocusIn>", disable_scroll_on_combo_focus)
            combo.bind("<FocusOut>", enable_scroll_after_combo_loses_focus)
            if selected_value in values:
                combo.set(selected_value)
            elif values:
                combo.current(0)
            return label, combo

        audio_input_devices = self.audio_input_devices
        audio_output_devices = self.audio_output_devices
        audio_device_error = self.audio_device_error

        input_device_labels = [label for label, _ in audio_input_devices]
        input_index_to_label = {
            index: label for label, index in audio_input_devices
        }
        output_device_labels = [label for label, _ in audio_output_devices]
        output_index_to_label = {
            index: label for label, index in audio_output_devices
        }
        stored_input_index = self.settings.get("microphone_index")
        stored_output_index = self.settings.get("speaker_index")
        lbl_input_device, cbo_input_device = make_device_row(
            form_frame,
            I18N.get(curr_lang, {}).get("settings_input_device", "麥克風輸入裝置："),
            input_device_labels,
            input_index_to_label.get(stored_input_index),
        )
        microphone_gain_frame = tk.Frame(form_frame, bg="#F1F5F9")
        microphone_gain_frame.pack(fill="x", pady=(10, 2))
        lbl_microphone_gain = tk.Label(
            microphone_gain_frame,
            text=I18N.get(curr_lang, {}).get(
                "settings_microphone_gain", "麥克風輸入增益："
            ),
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#F1F5F9", fg="#475569",
        )
        lbl_microphone_gain.pack(anchor="w")
        saved_microphone_gain = float(
            self.settings.get("microphone_gain", 1.0)
        )
        saved_microphone_gain = max(0.5, min(6.0, saved_microphone_gain))
        microphone_gain_value = tk.DoubleVar(
            value=saved_microphone_gain
        )
        microphone_gain_scale = tk.Scale(
            microphone_gain_frame,
            from_=0.5, to=6.0, resolution=0.1,
            orient="horizontal", variable=microphone_gain_value,
            bg="#FFFFFF", bd=0, highlightthickness=0,
            command=lambda value: microphone_gain_display.set(
                f"{float(value):.1f}×"
            ),
        )
        microphone_gain_scale.pack(fill="x", pady=(3, 0))
        microphone_gain_scale.bind("<Button-1>", click_to_set_value)
        microphone_gain_display = tk.StringVar(
            value=f"{microphone_gain_value.get():.1f}×"
        )
        microphone_gain_value_label = tk.Label(
            microphone_gain_frame, textvariable=microphone_gain_display,
            font=("Microsoft JhengHei", 10), bg="#F1F5F9", fg="#64748B",
        )
        microphone_gain_value_label.pack(anchor="e")
        lbl_output_device, cbo_output_device = make_device_row(
            form_frame,
            I18N.get(curr_lang, {}).get("settings_output_device", "音訊輸出裝置："),
            output_device_labels,
            output_index_to_label.get(stored_output_index),
        )
        if audio_device_error is not None:
            audio_device_status = tk.Label(
                form_frame,
                text=I18N.get(curr_lang, {}).get(
                    "settings_audio_unavailable",
                        "無法列舉音訊裝置：",
                    ) + str(audio_device_error),
                font=("Microsoft JhengHei", 10),
                bg="#F1F5F9", fg="#B91C1C", wraplength=650, justify="left",
            )
            audio_device_status.pack(anchor="w", pady=5)
        elif not audio_input_devices or not audio_output_devices:
            missing_directions = []
            if not audio_input_devices:
                missing_directions.append("麥克風")
            if not audio_output_devices:
                missing_directions.append("喇叭")
            audio_device_status = tk.Label(
                form_frame,
                text=I18N.get(curr_lang, {}).get(
                    "settings_audio_missing",
                    "找不到可用的音訊裝置：" + "、".join(missing_directions),
                ),
                font=("Microsoft JhengHei", 10),
                bg="#F1F5F9", fg="#B91C1C",
            )
            audio_device_status.pack(anchor="w", pady=5)
        else:
            audio_device_status = None

        waveform_title = tk.Label(
            form_frame,
            text=I18N.get(curr_lang, {}).get(
                "settings_waveform_title", "麥克風語音頻譜測試"
            ),
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#F1F5F9", fg="#475569",
        )
        waveform_title.pack(anchor="w", pady=(16, 4))
        waveform_canvas = tk.Canvas(
            form_frame, height=130, bg="#FFFFFF",
            highlightthickness=1, highlightbackground="#CBD5E1",
        )
        waveform_canvas.pack(fill="x", pady=(0, 8))
        waveform_status = tk.StringVar(
            value=I18N.get(curr_lang, {}).get(
                "settings_waveform_hint",
                "開始測試後，請對著麥克風說話以查看即時頻譜。",
            )
        )
        waveform_status_label = tk.Label(
            form_frame, textvariable=waveform_status,
            font=("Microsoft JhengHei", 10), bg="#F1F5F9", fg="#475569",
        )
        waveform_status_label.pack(anchor="w")
        waveform_state = {
            "active": False,
            "samples": deque([0.0] * 2048, maxlen=2048),
            "sample_rate": 16000,
            "lock": threading.Lock(),
            "stop_event": None,
        }

        def draw_microphone_waveform(event=None):
            width = max(1, waveform_canvas.winfo_width())
            height = max(1, waveform_canvas.winfo_height())
            waveform_canvas.delete("all")
            with waveform_state["lock"]:
                samples = list(waveform_state["samples"])
                sample_rate = waveform_state["sample_rate"]
            waveform_canvas.create_line(
                0, height - 18, width, height - 18, fill="#CBD5E1"
            )
            graph_height = max(1, height - 32)
            for decibels in (-40, -20, 0):
                level = (decibels + 60.0) / 60.0
                y = height - 19 - level * graph_height
                waveform_canvas.create_line(
                    0, y, width, y, fill="#E2E8F0", dash=(2, 3)
                )
                waveform_canvas.create_text(
                    4, y - 2, anchor="sw", fill="#94A3B8",
                    text=f"{decibels} dB",
                )
            spectrum = np.abs(
                np.fft.rfft(
                    np.asarray(samples, dtype=np.float32)
                    * np.hanning(len(samples))
                )
            ) / max(1, len(samples) / 2)
            frequencies = np.fft.rfftfreq(
                len(samples), d=1.0 / sample_rate
            )
            min_frequency = 80.0
            max_frequency = max(
                min_frequency + 1.0, min(8000.0, sample_rate / 2.0)
            )
            bar_count = max(1, min(64, width // 5))
            edges = np.geomspace(
                min_frequency, max_frequency, bar_count + 1
            )
            bar_width = width / bar_count
            gain = microphone_gain_value.get()
            for index in range(bar_count):
                bins = np.where(
                    (frequencies >= edges[index])
                    & (frequencies < edges[index + 1])
                )[0]
                magnitude = (
                    float(np.max(spectrum[bins])) * gain
                    if bins.size else 0.0
                )
                decibels = 20 * np.log10(max(magnitude, 1e-6))
                level = max(0.0, min(1.0, (decibels + 60.0) / 60.0))
                bar_height = level * graph_height
                x0 = index * bar_width + 1
                x1 = (index + 1) * bar_width - 1
                y0 = height - 19 - bar_height
                waveform_canvas.create_rectangle(
                    x0, y0, max(x0 + 1, x1), height - 19,
                    fill="#16A34A", outline="",
                )
            waveform_canvas.create_text(
                6, 5, anchor="nw", fill="#64748B",
                text="80 Hz",
            )
            waveform_canvas.create_text(
                width - 6, 5, anchor="ne", fill="#64748B",
                text=f"{int(max_frequency)} Hz",
            )
            waveform_canvas.create_text(
                width // 2, height - 2, anchor="s", fill="#64748B",
                text=I18N.get(selected_lang.get(), {}).get(
                    "settings_spectrum_frequency_axis", "頻率（對數刻度）"
                ),
            )
            if waveform_state["active"]:
                settings_win.after(40, draw_microphone_waveform)

        def stop_microphone_test():
            stop_event = waveform_state.get("stop_event")
            if stop_event is not None:
                stop_event.set()
            waveform_state["active"] = False
            waveform_state["stop_event"] = None
            if settings_win.winfo_exists():
                microphone_test_btn.configure(
                    text=I18N.get(selected_lang.get(), {}).get(
                        "settings_audio_test_button", "開始麥克風測試"
                    ),
                    bg="#2563EB",
                    activebackground="#1D4ED8",
                )
                waveform_status.set(
                    I18N.get(selected_lang.get(), {}).get(
                        "settings_waveform_hint",
                        "開始測試後，請對著麥克風說話以查看即時頻譜。",
                    )
                )
                draw_microphone_waveform()

        def start_microphone_test():
            selected_device_label = cbo_input_device.get()
            selected_device_index = next(
                (
                    index for label, index in audio_input_devices
                    if label == selected_device_label
                ),
                None,
            )
            if selected_device_index is None:
                messagebox.showerror(
                    I18N.get(selected_lang.get(), {}).get(
                        "settings_audio_test_title", "麥克風測試"
                    ),
                    I18N.get(selected_lang.get(), {}).get(
                        "settings_choose_input", "請先選擇可用的麥克風輸入裝置。"
                    ),
                    parent=settings_win,
                )
                return

            stop_event = threading.Event()
            waveform_state["stop_event"] = stop_event
            waveform_state["active"] = True
            microphone_test_btn.configure(
                text=I18N.get(selected_lang.get(), {}).get(
                    "settings_audio_test_stop", "停止麥克風測試"
                ),
                bg="#DC2626",
                activebackground="#B91C1C",
            )
            waveform_status.set(
                I18N.get(selected_lang.get(), {}).get(
                    "settings_audio_listening", "正在聆聽，請對著麥克風說話..."
                )
            )
            with waveform_state["lock"]:
                waveform_state["samples"].clear()
                waveform_state["samples"].extend([0.0] * 2048)
            draw_microphone_waveform()

            def capture_microphone():
                audio = None
                stream = None
                try:
                    audio = pyaudio.PyAudio()
                    device_info = audio.get_device_info_by_index(
                        selected_device_index
                    )
                    sample_rate = int(device_info.get("defaultSampleRate", 16000))
                    with waveform_state["lock"]:
                        waveform_state["sample_rate"] = sample_rate
                    stream = audio.open(
                        format=pyaudio.paInt16,
                        channels=1,
                        rate=sample_rate,
                        input=True,
                        input_device_index=selected_device_index,
                        frames_per_buffer=1024,
                    )
                    while not stop_event.is_set():
                        raw_samples = stream.read(
                            1024, exception_on_overflow=False
                        )
                        samples = (
                            np.frombuffer(raw_samples, dtype=np.int16)
                            .astype(np.float32) / 32768.0
                        )
                        with waveform_state["lock"]:
                            waveform_state["samples"].extend(samples.tolist())
                except Exception as exc:
                    def report_capture_error(error=exc):
                        if waveform_state.get("stop_event") is stop_event:
                            stop_microphone_test()
                            waveform_status.set(
                                f"{I18N.get(selected_lang.get(), {}).get('settings_audio_test_failed', '麥克風測試失敗')}：{error}"
                            )
                    if settings_win.winfo_exists():
                        settings_win.after(0, report_capture_error)
                finally:
                    try:
                        if stream is not None:
                            stream.close()
                    finally:
                        if audio is not None:
                            audio.terminate()

            threading.Thread(target=capture_microphone, daemon=True).start()

        def toggle_microphone_test():
            if waveform_state["active"]:
                stop_microphone_test()
            else:
                start_microphone_test()

        microphone_test_btn = tk.Button(
            form_frame,
            text=I18N.get(curr_lang, {}).get(
                "settings_audio_test_button", "開始麥克風測試"
            ),
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#2563EB", fg="white", activebackground="#1D4ED8",
            activeforeground="white", bd=0, relief="flat", cursor="hand2",
            command=toggle_microphone_test,
        )
        microphone_test_btn.pack(anchor="w", pady=(4, 0), ipady=5, ipadx=8)
        waveform_canvas.bind("<Configure>", draw_microphone_waveform)
        draw_microphone_waveform()

        form_frame = tab_content_frames[audio_video_frame]
        camera_indices = list(self.camera_indices)
        stored_camera_index = int(self.settings.get("camera_index", 0))
        if stored_camera_index not in camera_indices:
            camera_indices.append(stored_camera_index)
        camera_indices.sort()
        camera_labels = [f"Camera {index}" for index in camera_indices]
        camera_label_to_index = {
            label: index for label, index in zip(camera_labels, camera_indices)
        }
        lbl_camera_device, cbo_camera_device = make_device_row(
            form_frame,
            I18N.get(curr_lang, {}).get("settings_camera_device", "攝影機裝置："),
            camera_labels,
            f"Camera {stored_camera_index}",
        )
        lbl_camera_resolution = tk.Label(
            form_frame,
            text=I18N.get(curr_lang, {}).get(
                "settings_camera_resolution", "攝影機解析度："
            ),
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#F1F5F9", fg="#475569",
        )
        lbl_camera_resolution.pack(anchor="w", pady=(15, 4))
        camera_resolution_values = {
            "320x240": (320, 240),
            "640x480": (640, 480),
            "800x600": (800, 600),
            "1024x768": (1024, 768),
            "1280x960": (1280, 960),
            "1600x1200": (1600, 1200),
        }
        cbo_camera_resolution = ttk.Combobox(
            form_frame, values=list(camera_resolution_values),
            font=("Microsoft JhengHei", 11), state="readonly", width=18,
        )
        current_resolution = (
            int(self.settings.get("camera_width", 640)),
            int(self.settings.get("camera_height", 480)),
        )
        cbo_camera_resolution.set(
            next(
                (
                    label for label, size in camera_resolution_values.items()
                    if size == current_resolution
                ),
                "640x480",
            )
        )
        cbo_camera_resolution.pack(anchor="w")
        cbo_camera_resolution.bind("<FocusIn>", disable_scroll_on_combo_focus)
        cbo_camera_resolution.bind("<FocusOut>", enable_scroll_after_combo_loses_focus)
        camera_status = tk.Label(
            form_frame,
            text=I18N.get(curr_lang, {}).get(
                "settings_camera_hint",
                "儲存後套用；若相機目前已開啟，會立即重新啟動。",
            ),
            font=("Microsoft JhengHei", 10), bg="#F1F5F9", fg="#64748B",
            wraplength=650, justify="left",
        )
        camera_status.pack(anchor="w", pady=(10, 0))

        camera_preview_status = tk.StringVar(
            value=""
        )
        camera_preview_container = tk.Frame(
            form_frame, bg="#0F172A", width=480, height=360
        )
        camera_preview_container.pack(
            anchor="center", expand=False, pady=(14, 6)
        )
        camera_preview_container.pack_propagate(False)
        camera_preview_label = tk.Label(
            camera_preview_container, textvariable=camera_preview_status,
            bg="#0F172A", fg="#E2E8F0",
            font=("Microsoft JhengHei", 12),
            compound="center", anchor="center", justify="center",
        )
        camera_preview_label.pack(fill="both", expand=True)
        camera_preview_state = {
            "active": False,
            "stop_event": None,
            "frame": None,
            "error": None,
            "lock": threading.Lock(),
            "generation": 0,
            "photo": None,
            "thread": None,
        }
        camera_preview_tab_change = {"after_id": None}

        def stop_camera_preview():
            stop_event = camera_preview_state.get("stop_event")
            if stop_event is not None:
                stop_event.set()
                with self._camera_preview_lock:
                    self._camera_preview_stop_events.discard(stop_event)
            camera_preview_state["generation"] += 1
            camera_preview_state["active"] = False
            camera_preview_state["stop_event"] = None
            camera_preview_state["photo"] = None
            if settings_win.winfo_exists():
                camera_preview_label.configure(image="", textvariable=camera_preview_status)
                camera_preview_status.set(
                    I18N.get(selected_lang.get(), {}).get(
                        "settings_camera_preview_waiting", "鏡頭預覽畫面"
                    )
                )

        settings_win.protocol("WM_DELETE_WINDOW", on_window_close)

        def update_camera_preview(generation):
            if (
                not settings_win.winfo_exists()
                or not camera_preview_state["active"]
                or generation != camera_preview_state["generation"]
            ):
                return
            preview_error = camera_preview_state.get("error")
            if preview_error:
                stop_event = camera_preview_state.get("stop_event")
                if stop_event is not None:
                    stop_event.set()
                camera_preview_status.set(
                    f"{I18N.get(selected_lang.get(), {}).get('settings_camera_preview_failed', '攝影機預覽失敗')}：{preview_error}"
                )
                camera_preview_state["active"] = False
                camera_preview_state["stop_event"] = None
                camera_preview_btn.configure(
                    text=I18N.get(selected_lang.get(), {}).get(
                        "settings_camera_preview_start", "開始預覽"
                    )
                )
                return

            if self._is_camera_open():
                with self.camera_frame_lock:
                    frame = (
                        self.latest_camera_preview_frame.copy()
                        if self.latest_camera_preview_frame is not None else None
                    )
            else:
                with camera_preview_state["lock"]:
                    frame = (
                        camera_preview_state["frame"].copy()
                        if camera_preview_state["frame"] is not None else None
                    )
            if frame is not None:
                preview_width = max(1, camera_preview_label.winfo_width())
                preview_height = max(1, camera_preview_label.winfo_height())
                frame_height, frame_width = frame.shape[:2]
                scale = min(
                    preview_width / frame_width,
                    preview_height / frame_height,
                )
                preview_size = (
                    max(1, int(frame_width * scale)),
                    max(1, int(frame_height * scale)),
                )
                preview_frame = cv2.resize(frame, preview_size)
                preview_rgb = cv2.cvtColor(preview_frame, cv2.COLOR_BGR2RGB)
                photo = ImageTk.PhotoImage(image=Image.fromarray(preview_rgb))
                camera_preview_label.configure(
                    image=photo, text="", textvariable="",
                    anchor="center", justify="center",
                )
                camera_preview_state["photo"] = photo
            settings_win.after(
                16, lambda: update_camera_preview(generation)
            )

        def start_camera_preview():
            if camera_preview_state["active"]:
                return
            camera_preview_state["generation"] += 1
            generation = camera_preview_state["generation"]
            stop_event = threading.Event()
            camera_preview_state["stop_event"] = stop_event
            camera_preview_state["active"] = True
            camera_preview_state["error"] = None
            with self._camera_preview_lock:
                if self._shutting_down:
                    stop_event.set()
                    camera_preview_state["active"] = False
                    camera_preview_state["stop_event"] = None
                    return
                self._camera_preview_stop_events.add(stop_event)
            with camera_preview_state["lock"]:
                camera_preview_state["frame"] = None
            camera_preview_btn.configure(
                text=I18N.get(selected_lang.get(), {}).get(
                    "settings_camera_preview_stop", "停止預覽"
                )
            )
            camera_preview_label.configure(image="", text="", textvariable="")
            camera_preview_state["photo"] = None
            camera_preview_status.set("")

            if not self._is_camera_open():
                selected_index = camera_label_to_index.get(
                    cbo_camera_device.get(),
                    int(self.settings.get("camera_index", 0)),
                )
                selected_width, selected_height = camera_resolution_values.get(
                    cbo_camera_resolution.get(), (640, 480)
                )

                def capture_preview():
                    capture = None
                    try:
                        capture = cv2.VideoCapture(
                            selected_index, cv2.CAP_DSHOW
                        )
                        if not capture.isOpened():
                            raise RuntimeError(
                                f"無法開啟攝影機裝置 {selected_index}"
                            )
                        with self._camera_preview_lock:
                            if self._shutting_down or stop_event.is_set():
                                capture.release()
                                capture = None
                                return
                            self._camera_preview_captures[stop_event] = capture
                        capture.set(cv2.CAP_PROP_FRAME_WIDTH, selected_width)
                        capture.set(cv2.CAP_PROP_FRAME_HEIGHT, selected_height)
                        capture.set(cv2.CAP_PROP_FPS, 60)
                        while not stop_event.is_set():
                            success, frame = capture.read()
                            if not success:
                                raise RuntimeError("無法讀取攝影機畫面")
                            if generation != camera_preview_state["generation"]:
                                break
                            frame = self._crop_frame_to_aspect(
                                frame, selected_width, selected_height
                            )
                            with camera_preview_state["lock"]:
                                camera_preview_state["frame"] = frame
                    except Exception as exc:
                        if (
                            not stop_event.is_set()
                            and generation == camera_preview_state["generation"]
                        ):
                            camera_preview_state["error"] = str(exc)
                    finally:
                        if capture is not None:
                            capture.release()
                        with self._camera_preview_lock:
                            self._camera_preview_captures.pop(
                                stop_event, None
                            )
                            self._camera_preview_stop_events.discard(
                                stop_event
                            )

                preview_thread = threading.Thread(
                    target=capture_preview, daemon=True
                )
                camera_preview_state["thread"] = preview_thread
                preview_thread.start()
            update_camera_preview(generation)

        def refresh_camera_preview():
            if self._is_camera_open():
                with self.camera_frame_lock:
                    self.latest_camera_frame = None
                camera_preview_status.set(
                    I18N.get(selected_lang.get(), {}).get(
                        "settings_camera_preview_refreshing", "正在更新攝影機畫面..."
                    )
                )
                return

            preview_thread = camera_preview_state.get("thread")
            stop_camera_preview()
            camera_refresh_btn.configure(state="disabled")

            def reopen_preview():
                if not settings_win.winfo_exists():
                    return
                if preview_thread is not None and preview_thread.is_alive():
                    settings_win.after(50, reopen_preview)
                    return
                camera_preview_state["thread"] = None
                camera_refresh_btn.configure(state="normal")
                if notebook.select() == str(audio_video_frame):
                    start_camera_preview()

            reopen_preview()

        def toggle_camera_preview():
            if camera_preview_state["active"]:
                stop_camera_preview()
            else:
                start_camera_preview()

        camera_preview_actions = tk.Frame(form_frame, bg="#F1F5F9")
        camera_preview_actions.pack(anchor="w", pady=(0, 8))
        camera_preview_btn = tk.Button(
            camera_preview_actions,
            text=I18N.get(curr_lang, {}).get(
                "settings_camera_preview_start", "開始預覽"
            ),
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#2563EB", fg="white", activebackground="#1D4ED8",
            activeforeground="white", bd=0, relief="flat", cursor="hand2",
            command=toggle_camera_preview,
        )
        camera_preview_btn.pack(side="left", ipady=5, ipadx=8)
        camera_refresh_btn = tk.Button(
            camera_preview_actions,
            text=I18N.get(curr_lang, {}).get(
                "settings_camera_preview_refresh", "重新整理畫面"
            ),
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#0F766E", fg="white", activebackground="#115E59",
            activeforeground="white", bd=0, relief="flat", cursor="hand2",
            command=refresh_camera_preview,
        )
        camera_refresh_btn.pack(side="left", padx=(8, 0), ipady=5, ipadx=8)

        def update_camera_preview_for_selected_tab(event=None):
            if camera_preview_tab_change["after_id"] is not None:
                try:
                    settings_win.after_cancel(
                        camera_preview_tab_change["after_id"]
                    )
                except tk.TclError:
                    pass

            def apply_tab_preview_state():
                camera_preview_tab_change["after_id"] = None
                if not settings_win.winfo_exists():
                    return
                if notebook.select() == str(audio_video_frame):
                    start_camera_preview()
                else:
                    stop_camera_preview()

            camera_preview_tab_change["after_id"] = settings_win.after(
                400, apply_tab_preview_state
            )

        notebook.bind(
            "<<NotebookTabChanged>>",
            update_camera_preview_for_selected_tab,
            add="+",
        )

        form_frame = tab_content_frames[voice_frame]

        # 6. Email
        lbl_notify_sec = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_notify_sec.pack(anchor="w", pady=(15,0))
        val_email = tk.BooleanVar(value=self.settings.get("email_enabled", True))
        chk_email = tk.Checkbutton(form_frame, variable=val_email, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#2563EB", bd=0)
        chk_email.pack(anchor="w", pady=(0, 5))

        email_entry_frame = tk.Frame(form_frame, bg="#F1F5F9")
        email_entry_frame.pack(fill="x", pady=2)
        lbl_email_prefix = tk.Label(email_entry_frame, font=("Microsoft JhengHei", 11, "bold"), bg="#F1F5F9", fg="#475569")
        lbl_email_prefix.pack(side="left")

        entry_email = tk.Entry(email_entry_frame, font=("Microsoft JhengHei", 11), bd=1, relief="solid", bg="#FFFFFF", width=25)
        entry_email.insert(0, self.settings.get("receiver_email", ""))
        entry_email.pack(side="left", padx=5)

        pending_email_verification = {
            "code": "",
            "email": self.settings.get("receiver_email", "").strip(),
            "verified": bool(self.settings.get("receiver_email", "").strip()),
            "expires_at": 0,
        }
        email_verification_status = tk.StringVar(value="驗證碼有效時間：尚未寄送")

        def invalidate_email_verification(event=None):
            current_email = entry_email.get().strip()
            if current_email != pending_email_verification.get("email", ""):
                pending_email_verification["code"] = ""
                pending_email_verification["verified"] = False
                pending_email_verification["expires_at"] = 0
                email_verification_status.set("驗證碼有效時間：尚未寄送")

        entry_email.bind("<KeyRelease>", invalidate_email_verification)
        entry_email.bind("<<Paste>>", invalidate_email_verification)
        entry_email.bind("<<Cut>>", invalidate_email_verification)

        def send_email_verification_for_settings():
            target_email = entry_email.get().strip()
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", target_email):
                messagebox.showerror("Email 格式錯誤", "請先輸入有效的家屬 Email，再寄送驗證碼。", parent=self.root)
                return

            verification_code = self.generate_email_verification_code()
            ok, msg = self.send_verification_code_email(target_email, verification_code)
            if not ok:
                messagebox.showerror("驗證碼發送失敗", msg, parent=self.root)
                return

            pending_email_verification["code"] = verification_code
            pending_email_verification["email"] = target_email
            pending_email_verification["verified"] = False
            pending_email_verification["expires_at"] = time.time() + 300
            entry_verification_code.delete(0, tk.END)
            update_settings_email_verification_countdown()
            messagebox.showinfo("驗證碼已寄出", msg, parent=self.root)

        def update_settings_email_verification_countdown():
            expires_at = float(pending_email_verification.get("expires_at", 0) or 0)
            remaining = max(0, int(expires_at - time.time()))
            if not expires_at:
                email_verification_status.set("驗證碼有效時間：尚未寄送")
            elif remaining > 0:
                minutes, seconds = divmod(remaining, 60)
                email_verification_status.set(
                    f"驗證碼有效時間：{minutes:02d}:{seconds:02d}"
                )
                settings_win.after(1000, update_settings_email_verification_countdown)
            else:
                pending_email_verification["code"] = ""
                pending_email_verification["verified"] = False
                email_verification_status.set("驗證碼已失效，請重新寄送")

        def verify_email_code_for_settings():
            target_email = entry_email.get().strip()
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", target_email):
                messagebox.showerror("Email 格式錯誤", "請先輸入有效的家屬 Email。", parent=self.root)
                return

            expected_code = str(pending_email_verification.get("code", "")).strip()
            if not expected_code:
                messagebox.showerror("尚未寄送驗證碼", "請先點擊「寄送驗證碼」後再進行驗證。", parent=self.root)
                return
            if target_email != pending_email_verification.get("email", ""):
                pending_email_verification["code"] = ""
                pending_email_verification["verified"] = False
                pending_email_verification["expires_at"] = 0
                messagebox.showerror("Email 已變更", "Email 已變更，請重新寄送驗證碼。", parent=self.root)
                return
            if time.time() >= float(pending_email_verification.get("expires_at", 0) or 0):
                pending_email_verification["code"] = ""
                pending_email_verification["verified"] = False
                email_verification_status.set("驗證碼已失效，請重新寄送")
                messagebox.showerror("驗證碼已過期", "請重新寄送驗證碼後再進行驗證。", parent=self.root)
                return

            entered_code = entry_verification_code.get().strip()
            if entered_code != expected_code:
                messagebox.showerror("驗證碼錯誤", "驗證碼不正確，請重新輸入。", parent=self.root)
                return

            pending_email_verification["verified"] = True
            messagebox.showinfo("驗證成功", "家屬 Email 已完成驗證，可以儲存。", parent=self.root)

        verification_button_row = tk.Frame(form_frame, bg="#F1F5F9")
        verification_button_row.pack(fill="x", pady=(10, 6))

        send_verification_btn = tk.Button(
            verification_button_row,
            text="寄送驗證碼",
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            bd=0,
            relief="flat",
            cursor="hand2",
            command=send_email_verification_for_settings
        )
        send_verification_btn.pack(side="left")

        verify_code_btn = tk.Button(
            verification_button_row,
            text="驗證",
            font=("Microsoft JhengHei", 11, "bold"),
            bg="#16A34A",
            fg="white",
            activebackground="#15803D",
            activeforeground="white",
            bd=0,
            relief="flat",
            cursor="hand2",
            command=verify_email_code_for_settings
        )
        verify_code_btn.pack(side="left", padx=(10, 0))

        verification_code_row = tk.Frame(form_frame, bg="#F1F5F9")
        verification_code_row.pack(fill="x", pady=2)
        tk.Label(verification_code_row, text="驗證碼：", font=("Microsoft JhengHei", 11, "bold"), bg="#F1F5F9", fg="#475569").pack(side="left")
        entry_verification_code = tk.Entry(verification_code_row, font=("Microsoft JhengHei", 11), bd=1, relief="solid", bg="#FFFFFF", width=20)
        entry_verification_code.pack(side="left", padx=5)
        tk.Label(
            form_frame,
            textvariable=email_verification_status,
            font=("Microsoft JhengHei", 10),
            bg="#F1F5F9",
            fg="#64748B",
        ).pack(anchor="w", pady=(0, 4))

        form_frame = tab_content_frames[visual_frame]

        # 主題
        theme_frame = tk.Frame(form_frame, bg="#F1F5F9")
        theme_frame.pack(fill="x", pady=10)
        lbl_theme = tk.Label(theme_frame, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#475569")
        lbl_theme.pack(side="left")

        def block_combobox_wheel(widget):
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                try:
                    widget.bind(seq, lambda event, seq=seq: "break")
                except Exception:
                    pass
                try:
                    widget.bind_class("TCombobox", seq, lambda event, seq=seq: "break")
                except Exception:
                    pass
                try:
                    widget.bind_class("TCombobox.Popdown", seq, lambda event, seq=seq: "break")
                except Exception:
                    pass

        theme_values = ["自動 (隨時間)", "明亮模式", "黑暗模式"]
        theme_display_values = [self.localize_runtime_text(value) for value in theme_values]
        cbo_theme = ttk.Combobox(theme_frame, values=theme_display_values, font=("Microsoft JhengHei", 12), state="readonly", width=18)
        cbo_theme.set(self.localize_runtime_text(self.settings.get("theme_mode", "自動 (隨時間)")))
        cbo_theme.pack(side="left", padx=10)
        cbo_theme.bind("<FocusIn>", disable_scroll_on_combo_focus)
        cbo_theme.bind("<FocusOut>", enable_scroll_after_combo_loses_focus)

        font_size_options = [
            ("small", "較小"),
            ("normal", "標準"),
            ("large", "較大"),
            ("largest", "特大"),
        ]
        pending_font_size = {"value": original_font_size}
        lbl_font_size = tk.Label(
            form_frame, font=("Microsoft JhengHei", 12, "bold"),
            bg="#F1F5F9", fg="#475569", text="介面字體大小："
        )
        lbl_font_size.pack(anchor="w", pady=(15, 4))
        cbo_font_size = ttk.Combobox(
            form_frame,
            values=[label for _, label in font_size_options],
            font=("Microsoft JhengHei", 12),
            state="readonly",
            width=14,
        )
        cbo_font_size.pack(anchor="w", pady=(0, 8))
        cbo_font_size.bind("<FocusIn>", disable_scroll_on_combo_focus)
        cbo_font_size.bind("<FocusOut>", enable_scroll_after_combo_loses_focus)

        def selected_font_size_key():
            selected_label = cbo_font_size.get()
            if selected_lang.get() == "English":
                return dict(zip(
                    [
                        ENGLISH_RUNTIME_TRANSLATIONS.get(label, label)
                        for _, label in font_size_options
                    ],
                    [key for key, _ in font_size_options]
                )).get(selected_label, pending_font_size["value"])
            return dict(
                    (label, key) for key, label in font_size_options
                ).get(selected_label, pending_font_size["value"])

        def stage_selected_font_size(*_):
            pending_font_size["value"] = selected_font_size_key()

        def apply_selected_font_size():
            stage_selected_font_size()
            self.settings["font_size"] = pending_font_size["value"]
            self.apply_font_size()

        cbo_font_size.bind("<<ComboboxSelected>>", stage_selected_font_size)

        # 7. 軌跡
        lbl_trail_sec = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_trail_sec.pack(anchor="w", pady=(15,0))
        val_trail = tk.BooleanVar(value=self.settings.get("enable_trail", True))
        chk_trail = tk.Checkbutton(form_frame, variable=val_trail, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#059669", bd=0)
        chk_trail.pack(anchor="w", pady=(0, 5))

        lbl_trail_len = tk.Label(form_frame, font=("Microsoft JhengHei", 11, "bold"), bg="#F1F5F9", fg="#475569")
        lbl_trail_len.pack(anchor="w", pady=(2, 0))
        val_trail_len = tk.IntVar(value=int(self.settings.get("trail_length", 20)))
        scale_trail_len = tk.Scale(form_frame, from_=5, to=40, resolution=1, orient="horizontal", variable=val_trail_len, bg="#FFFFFF", bd=0, highlightthickness=0)
        scale_trail_len.pack(fill="x", pady=5)
        scale_trail_len.bind("<Button-1>", click_to_set_value)

        form_frame = tab_content_frames[safety_frame]

        # 8. 地震預警
        lbl_eq_sec = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_eq_sec.pack(anchor="w", pady=(15,0))

        def test_earthquake_alert():
            self.play_system_sound("execute")
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.popup_earthquake_report_ui(
                scale="芮氏規模 6.5 (測試)", intensity="5", origin_time=now_str,
                epicenter="花蓮縣政府東方 25.0 公里", report_content="【測試警報】確認畫面與語音播報是否正常。"
            )

        btn_test_eq = tk.Button(
            form_frame, font=("Microsoft JhengHei", 12, "bold"), bg="#EF4444", fg="white",
            activebackground="#DC2626", activeforeground="white", bd=2, relief="raised", cursor="hand2", padx=10, pady=5,
            command=test_earthquake_alert
        )
        btn_test_eq.pack(anchor="w", pady=(5, 10))

        val_eq = tk.BooleanVar(value=self.settings.get("earthquake_enabled", True))
        chk_eq = tk.Checkbutton(form_frame, variable=val_eq, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#DC2626", bd=0)
        chk_eq.pack(anchor="w", pady=(0, 5))

        # 9. 跌倒警報偵測
        lbl_fall_sec = tk.Label(form_frame, font=("Microsoft JhengHei", 12), bg="#F1F5F9", fg="#475569")
        lbl_fall_sec.pack(anchor="w", pady=(15,0))
        val_fall = tk.BooleanVar(value=self.settings.get("fall_detection_enabled", True))
        chk_fall = tk.Checkbutton(form_frame, variable=val_fall, font=("Microsoft JhengHei", 12, "bold"), bg="#F1F5F9", fg="#7C2D12", bd=0)
        chk_fall.pack(anchor="w", pady=(0, 10))

        btn_test_fall = tk.Button(
            form_frame, text="🚨 測試跌倒警報", font=("Microsoft JhengHei", 12, "bold"), bg="#7C2D12", fg="white",
            activebackground="#9A4D2E", activeforeground="white", bd=2, relief="raised", cursor="hand2", padx=10, pady=5,
            command=self.test_fall_alert
        )
        btn_test_fall.pack(anchor="w", pady=(0, 20))

        eq_county_frame = tk.Frame(form_frame, bg="#F1F5F9")
        eq_county_frame.pack(fill="x", pady=2)
        lbl_county = tk.Label(eq_county_frame, font=("Microsoft JhengHei", 11, "bold"), bg="#F1F5F9", fg="#475569")
        lbl_county.pack(side="left")

        taiwan_counties = ["臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市", "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣", "臺東縣", "澎湖縣", "金門縣", "連江縣"]
        county_display_values = [self.localize_runtime_text(value) for value in taiwan_counties]
        cb_county = ttk.Combobox(eq_county_frame, values=county_display_values, font=("Microsoft JhengHei", 11), state="readonly", width=15)
        cb_county.set(self.localize_runtime_text(self.settings.get("target_county", getattr(self, "target_county", "臺中市"))))
        cb_county.pack(side="left", padx=5)
        cb_county.bind("<FocusIn>", disable_scroll_on_combo_focus)
        cb_county.bind("<FocusOut>", enable_scroll_after_combo_loses_focus)

        eq_intensity_frame = tk.Frame(form_frame, bg="#F1F5F9")
        eq_intensity_frame.pack(fill="x", pady=4)
        lbl_intensity = tk.Label(eq_intensity_frame, font=("Microsoft JhengHei", 11, "bold"), bg="#F1F5F9", fg="#475569")
        lbl_intensity.pack(side="left")

        cb_min_intensity = ttk.Combobox(eq_intensity_frame, values=["0", "1", "2", "3", "4", "5", "6", "7"], font=("Microsoft JhengHei", 11), state="readonly", width=5)
        cb_min_intensity.set(str(self.settings.get("min_alert_intensity", getattr(self, "min_alert_intensity", "3"))))
        cb_min_intensity.pack(side="left", padx=3)
        cb_min_intensity.bind("<FocusIn>", disable_scroll_on_combo_focus)
        cb_min_intensity.bind("<FocusOut>", enable_scroll_after_combo_loses_focus)
        lbl_intensity_suffix = tk.Label(eq_intensity_frame, font=("Microsoft JhengHei", 11), bg="#F1F5F9", fg="#475569")
        lbl_intensity_suffix.pack(side="left")

        form_frame = tab_content_frames[tools_frame]

        btn_voice_help = tk.Button(form_frame, font=("Microsoft JhengHei", 12, "bold"), bg="#0284C7", fg="white", activebackground="#0369A1", activeforeground="white", bd=0, relief="flat", cursor="hand2", command=self.show_voice_commands_dialog)
        btn_voice_help.pack(fill="x", pady=5, ipady=6)

        def open_wizard_from_settings():
            on_window_close()
            self.open_setup_wizard()

        setup_wizard_btn = tk.Button(
            form_frame, font=("Microsoft JhengHei", 12, "bold"),
            bg="#4F46E5", fg="white", activebackground="#4338CA",
            activeforeground="white", bd=0, relief="flat", cursor="hand2",
            command=open_wizard_from_settings
        )
        setup_wizard_btn.pack(fill="x", pady=5, ipady=6)

        clear_settings_btn = tk.Button(
            form_frame,
            text=I18N.get(curr_lang, I18N["國語 (華語)"]).get(
                "btn_clear_settings", "🗑️ 清除設定 JSON 檔"
            ),
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#B91C1C",
            fg="white",
            activebackground="#991B1B",
            activeforeground="white",
            bd=0,
            relief="flat",
            cursor="hand2",
        )
        clear_settings_btn.pack(fill="x", pady=(16, 5), ipady=6)

        bottom_action_frame = tk.Frame(settings_win, bg="#F1F5F9")
        bottom_action_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 12))

        maintenance_action_frame = tk.Frame(bottom_action_frame, bg="#F1F5F9")
        maintenance_action_frame.pack(fill="x", pady=(0, 10))

        reset_btn = tk.Button(
            maintenance_action_frame,
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#F59E0B",
            fg="white",
            activebackground="#D97706",
            activeforeground="white",
            pady=6,
            relief="raised",
            bd=3,
            cursor="hand2",
        )
        reset_btn.pack(side="left", expand=True, fill="x", padx=(0, 5))

        btn_action_frame = tk.Frame(bottom_action_frame, bg="#F1F5F9")
        btn_action_frame.pack(fill="x")

        save_btn = tk.Button(
            btn_action_frame,
            font=("Microsoft JhengHei", 14, "bold"),
            bg="#10B981",
            fg="white",
            pady=8,
            cursor="hand2",
        )
        save_btn.pack(side="left", padx=(0, 5), expand=True, fill="x")

        apply_btn = tk.Button(
            btn_action_frame,
            font=("Microsoft JhengHei", 14, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            pady=8,
            cursor="hand2",
        )
        apply_btn.pack(side="left", padx=5, expand=True, fill="x")

        cancel_btn = tk.Button(
            btn_action_frame,
            font=("Microsoft JhengHei", 14, "bold"),
            bg="#EF4444",
            fg="white",
            pady=8,
            command=on_window_close,
            cursor="hand2",
        )
        cancel_btn.pack(side="right", padx=(10, 0), expand=True, fill="x")

        def align_setting_labels_left(widget):
            for child in widget.winfo_children():
                if (
                    isinstance(child, tk.Label)
                    and child not in (
                        about_title,
                        about_description,
                        about_details,
                    )
                ):
                    child.configure(anchor="w", justify="left")
                align_setting_labels_left(child)

        align_setting_labels_left(settings_win)
        camera_preview_label.configure(
            anchor="center", justify="center", wraplength=600
        )

        # ── 💡 核心：動態重新載入 UI 文字 ──
        def refresh_ui_texts(lang_key):
            txt = I18N.get(lang_key, I18N["國語 (華語)"])

            def set_animated_text(widget, text):
                self.transition_text(widget, text)

            for frame, label in zip(
                tab_frames,
                tab_labels.get(lang_key, tab_labels["國語 (華語)"])
            ):
                notebook.tab(frame, text=label)
            set_animated_text(
                about_title,
                txt.get("about_product_name", "智慧多媒體互動平台"),
            )
            set_animated_text(
                about_description,
                txt.get(
                    "about_description",
                    "專為高齡者設計的直覺式互動平台",
                ),
            )
            set_animated_text(
                about_details,
                txt.get(
                    "about_details",
                    "2026 中科大資訊工程系專題\n\n"
                    "專題成員\n"
                    "1411232070 黃品傑\n"
                    "1411232067 莊道鈞\n\n"
                    "指導老師\n"
                    "陳永隆教授",
                ),
            )
            set_animated_text(title_lbl, txt["title"])
            set_animated_text(lbl_lang_sec, txt["lang_section"])
            set_animated_text(lbl_smooth, txt["smooth_lbl"])
            set_animated_text(lbl_speed, txt["speed_lbl"])
            set_animated_text(lbl_dwell, txt["dwell_lbl"])
            for label_key, label_widget in camera_box_labels.items():
                set_animated_text(label_widget, txt[label_key])
            set_animated_text(lbl_vspeed, txt["vspeed_lbl"])
            set_animated_text(lbl_voice_sec, txt["voice_sec"])
            set_animated_text(chk_tts, txt["chk_tts"])
            set_animated_text(
                lbl_input_device,
                txt.get("settings_input_device", "麥克風輸入裝置："),
            )
            set_animated_text(
                lbl_microphone_gain,
                txt.get("settings_microphone_gain", "麥克風輸入增益："),
            )
            set_animated_text(
                lbl_output_device,
                txt.get("settings_output_device", "音訊輸出裝置："),
            )
            set_animated_text(
                waveform_title,
                txt.get(
                    "settings_waveform_title", "麥克風語音頻譜測試"
                ),
            )
            if not waveform_state["active"]:
                waveform_status.set(
                    txt.get(
                        "settings_waveform_hint",
                        "開始測試後，請對著麥克風說話以查看即時頻譜。",
                    )
                )
            else:
                waveform_status.set(
                    txt.get(
                        "settings_audio_listening",
                        "正在聆聽，請對著麥克風說話...",
                    )
                )
            set_animated_text(
                microphone_test_btn,
                txt.get(
                    "settings_audio_test_stop" if waveform_state["active"]
                    else "settings_audio_test_button",
                    "停止麥克風測試" if waveform_state["active"]
                    else "開始麥克風測試",
                ),
            )
            set_animated_text(
                camera_preview_btn,
                txt.get(
                    "settings_camera_preview_stop"
                    if camera_preview_state["active"]
                    else "settings_camera_preview_start",
                    "停止預覽" if camera_preview_state["active"] else "開始預覽",
                ),
            )
            set_animated_text(
                camera_refresh_btn,
                txt.get("settings_camera_preview_refresh", "重新整理畫面"),
            )
            if not camera_preview_state["active"]:
                camera_preview_status.set("")
            elif not camera_preview_state.get("error"):
                camera_preview_status.set("")
            set_animated_text(
                lbl_camera_device,
                txt.get("settings_camera_device", "攝影機裝置："),
            )
            set_animated_text(
                lbl_camera_resolution,
                txt.get("settings_camera_resolution", "攝影機解析度："),
            )
            set_animated_text(
                camera_status,
                txt.get(
                    "settings_camera_hint",
                    "儲存後套用；若相機目前已開啟，會立即重新啟動。",
                ),
            )
            if audio_device_status is not None:
                if audio_device_error is not None:
                    audio_device_status.configure(
                        text=(
                            txt.get(
                                "settings_audio_unavailable",
                                "無法列舉音訊裝置：",
                            )
                            + str(audio_device_error)
                        )
                    )
                elif not audio_input_devices or not audio_output_devices:
                    audio_device_status.configure(
                        text=txt.get(
                            "settings_audio_missing",
                            "找不到可用的音訊輸入或輸出裝置。",
                        )
                    )
            set_animated_text(lbl_notify_sec, txt["notify_sec"])
            set_animated_text(chk_email, txt["chk_email"])
            set_animated_text(lbl_email_prefix, txt["email_prefix"])
            set_animated_text(lbl_theme, txt["theme_lbl"])
            is_english = lang_key == "English"
            set_animated_text(
                lbl_font_size,
                ENGLISH_RUNTIME_TRANSLATIONS["介面字體大小："]
                if is_english else "介面字體大小：",
            )
            display_font_size_options = [
                ENGLISH_RUNTIME_TRANSLATIONS.get(label, label) if is_english else label
                for _, label in font_size_options
            ]
            cbo_font_size.configure(values=display_font_size_options)
            stored_font_size_label = dict(font_size_options).get(
                pending_font_size["value"], "標準"
            )
            cbo_font_size.set(
                ENGLISH_RUNTIME_TRANSLATIONS.get(stored_font_size_label, stored_font_size_label)
                if is_english else stored_font_size_label
            )
            set_animated_text(lbl_trail_sec, txt["trail_sec"])
            set_animated_text(chk_trail, txt["chk_trail"])
            set_animated_text(lbl_trail_len, txt["trail_len"])
            set_animated_text(lbl_eq_sec, txt["eq_sec"])
            set_animated_text(btn_test_eq, txt["btn_test_eq"])
            set_animated_text(chk_eq, txt["chk_eq"])
            set_animated_text(lbl_fall_sec, txt["fall_sec"])
            set_animated_text(chk_fall, txt["chk_fall"])
            set_animated_text(
                btn_test_fall, txt.get("btn_test_fall", "🚨 測試跌倒警報")
            )
            set_animated_text(lbl_county, txt["county_lbl"])
            set_animated_text(lbl_intensity, txt["intensity_lbl"])
            set_animated_text(lbl_intensity_suffix, txt["intensity_suffix"])
            set_animated_text(btn_voice_help, txt["btn_voice_help"])
            set_animated_text(setup_wizard_btn, txt["btn_setup_wizard"])
            set_animated_text(reset_btn, txt["btn_reset"])
            set_animated_text(
                clear_settings_btn,
                txt.get("btn_clear_settings", "🗑️ 清除設定 JSON 檔"),
            )
            set_animated_text(save_btn, txt["btn_save"])
            set_animated_text(apply_btn, txt.get("btn_apply", "✅ 套用"))
            set_animated_text(cancel_btn, txt["btn_cancel"])

            display_theme_values = [
                ENGLISH_RUNTIME_TRANSLATIONS.get(value, value) if is_english else value
                for value in theme_values
            ]
            cbo_theme.configure(values=display_theme_values)
            stored_theme = self.settings.get("theme_mode", "自動 (隨時間)")
            cbo_theme.set(
                ENGLISH_RUNTIME_TRANSLATIONS.get(stored_theme, stored_theme)
                if is_english else stored_theme
            )

            display_county_values = [
                ENGLISH_RUNTIME_TRANSLATIONS.get(value, value) if is_english else value
                for value in taiwan_counties
            ]
            cb_county.configure(values=display_county_values)
            stored_county = self.settings.get("target_county", getattr(self, "target_county", "臺中市"))
            cb_county.set(
                ENGLISH_RUNTIME_TRANSLATIONS.get(stored_county, stored_county)
                if is_english else stored_county
            )

        # 生成單選 Radiobutton 並綁定即時切換事件
        for text, val in languages:
            rb = tk.Radiobutton(
                lang_card, text=text, value=val, variable=selected_lang,
                font=("Microsoft JhengHei", 11), bg="#FFFFFF", fg="#1E293B",
                activebackground="#FFFFFF", selectcolor="#E2E8F0", cursor="hand2",
                command=lambda: refresh_ui_texts(selected_lang.get())  # 點擊選項時即時更新選單語言
            )
            rb.pack(anchor="w", pady=2)

        def clear_settings_file():
            lang_text = I18N.get(selected_lang.get(), I18N["國語 (華語)"])
            if not messagebox.askyesno(
                lang_text.get("msg_clear_title", "清除設定檔"),
                lang_text.get(
                    "msg_clear_confirm",
                    "確定要刪除 settings.json？重新啟動後將使用預設設定；此工作階段不會再寫入設定檔。",
                ),
                parent=settings_win,
            ):
                return
            try:
                with self._settings_io_lock:
                    if os.path.exists(SETTINGS_FILE):
                        os.remove(SETTINGS_FILE)
                    self._settings_file_cleared = True
            except OSError as exc:
                messagebox.showerror(
                    lang_text.get("msg_clear_title", "清除設定檔"),
                    f"{lang_text.get('msg_clear_failed', '無法清除設定檔')}\n{exc}",
                    parent=settings_win,
                )
                return

            save_btn.configure(state="disabled")
            clear_settings_btn.configure(state="disabled")
            messagebox.showinfo(
                lang_text.get("msg_clear_title", "清除設定檔"),
                lang_text.get(
                    "msg_clear_success",
                    "設定檔已清除。請重新啟動程式以使用預設設定；重新啟動前不會再寫入設定檔。",
                ),
                parent=settings_win,
            )

        # 重置與儲存按鈕綁定
        def reset_to_defaults():
            selected_lang.set(DEFAULT_SETTINGS.get("language", "國語 (華語)"))
            val_smooth.set(DEFAULT_SETTINGS["smooth_factor"])
            val_speed.set(DEFAULT_SETTINGS["speed_multiplier"])
            val_dwell.set(DEFAULT_SETTINGS["dwell_threshold"])
            camera_box_x_offset.set(DEFAULT_SETTINGS["camera_box_x_offset"])
            camera_box_y_offset.set(DEFAULT_SETTINGS["camera_box_y_offset"])
            camera_box_width.set(DEFAULT_SETTINGS["camera_box_width"])
            camera_box_height.set(DEFAULT_SETTINGS["camera_box_height"])
            val_vspeed.set(DEFAULT_SETTINGS["voice_speed"])
            val_tts.set(DEFAULT_SETTINGS["tts_enabled"])
            microphone_gain_value.set(
                float(DEFAULT_SETTINGS.get("microphone_gain", 1.0))
            )
            microphone_gain_display.set(
                f"{microphone_gain_value.get():.1f}×"
            )
            val_email.set(DEFAULT_SETTINGS.get("email_enabled", True))
            entry_email.delete(0, tk.END)
            entry_email.insert(0, DEFAULT_SETTINGS.get("receiver_email", ""))
            entry_verification_code.delete(0, tk.END)
            pending_email_verification["code"] = ""
            pending_email_verification["email"] = ""
            pending_email_verification["verified"] = False
            pending_email_verification["expires_at"] = 0
            email_verification_status.set("驗證碼有效時間：尚未寄送")
            val_trail.set(DEFAULT_SETTINGS.get("enable_trail", True))
            val_trail_len.set(int(DEFAULT_SETTINGS.get("trail_length", 20)))
            cbo_theme.set(DEFAULT_SETTINGS.get("theme_mode", "自動 (隨時間)"))
            val_eq.set(True)
            val_fall.set(DEFAULT_SETTINGS.get("fall_detection_enabled", True))
            cb_county.set("臺中市")
            cb_min_intensity.set("3")
            cbo_input_device.set(
                input_index_to_label.get(
                    DEFAULT_SETTINGS.get("microphone_index"),
                    input_device_labels[0] if input_device_labels else "",
                )
            )
            cbo_output_device.set(
                output_index_to_label.get(
                    DEFAULT_SETTINGS.get("speaker_index"),
                    output_device_labels[0] if output_device_labels else "",
                )
            )
            default_camera_index = int(DEFAULT_SETTINGS.get("camera_index", 0))
            cbo_camera_device.set(
                f"Camera {default_camera_index}"
            )
            default_resolution = (
                int(DEFAULT_SETTINGS.get("camera_width", 640)),
                int(DEFAULT_SETTINGS.get("camera_height", 480)),
            )
            cbo_camera_resolution.set(
                next(
                    (
                        label for label, size in camera_resolution_values.items()
                        if size == default_resolution
                    ),
                    "640x480",
                )
            )
            
            refresh_ui_texts(selected_lang.get())
            default_font_size_label = dict(font_size_options)[
                DEFAULT_SETTINGS.get("font_size", "normal")
            ]
            cbo_font_size.set(
                ENGLISH_RUNTIME_TRANSLATIONS.get(default_font_size_label, default_font_size_label)
                if selected_lang.get() == "English" else default_font_size_label
            )
            pending_font_size["value"] = DEFAULT_SETTINGS.get(
                "font_size", "normal"
            )
            self.play_system_sound("warn")
            self.tts_speak(I18N.get(selected_lang.get(), {}).get("msg_reset", "已重設為預設參數"))

        def apply_settings(save_changes):
            nonlocal font_size_saved, language_saved
            chosen_lang = selected_lang.get()
            receiver_address = entry_email.get().strip()
            if receiver_address:
                if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", receiver_address):
                    messagebox.showerror("Email 格式錯誤", "請輸入有效的家屬 Email。")
                    return
                if (
                    not pending_email_verification.get("verified", False)
                    or receiver_address != pending_email_verification.get("email", "")
                    or (
                        pending_email_verification.get("expires_at", 0)
                        and time.time() >= float(pending_email_verification.get("expires_at", 0))
                    )
                ):
                    messagebox.showerror("Email 尚未驗證", "請先點擊「寄送驗證碼」並完成驗證，再儲存家屬 Email。")
                    return
            elif bool(val_email.get()):
                messagebox.showerror("Email 格式錯誤", "啟用 Email 通知時，請輸入家屬 Email。")
                return
            self.settings["language"] = chosen_lang
            self.settings["smooth_factor"] = int(val_smooth.get())
            self.settings["speed_multiplier"] = float(val_speed.get())
            self.settings["voice_speed"] = int(val_vspeed.get())
            self.settings["camera_box_x_offset"] = int(
                camera_box_x_offset.get()
            )
            self.settings["camera_box_y_offset"] = int(
                camera_box_y_offset.get()
            )
            self.settings["camera_box_width"] = int(camera_box_width.get())
            self.settings["camera_box_height"] = int(camera_box_height.get())
            self.settings["tts_enabled"] = bool(val_tts.get())
            self.settings["email_enabled"] = bool(val_email.get())
            self.settings["receiver_email"] = receiver_address
            selected_theme = cbo_theme.get()
            if chosen_lang == "English":
                selected_theme = dict(zip(
                    [ENGLISH_RUNTIME_TRANSLATIONS.get(value, value) for value in theme_values],
                    theme_values
                )).get(selected_theme, selected_theme)
            self.settings["theme_mode"] = selected_theme
            apply_selected_font_size()
            if save_changes:
                font_size_saved = True
            settings_win.update_idletasks()
            self.settings["enable_trail"] = bool(val_trail.get())
            
            new_trail_len = int(val_trail_len.get())
            self.settings["trail_length"] = new_trail_len

            self.settings["earthquake_enabled"] = bool(val_eq.get())
            self.settings["fall_detection_enabled"] = bool(val_fall.get())
            selected_county = cb_county.get()
            if chosen_lang == "English":
                selected_county = dict(zip(
                    [ENGLISH_RUNTIME_TRANSLATIONS.get(value, value) for value in taiwan_counties],
                    taiwan_counties
                )).get(selected_county, selected_county)
            self.settings["target_county"] = selected_county
            self.settings["min_alert_intensity"] = cb_min_intensity.get()

            selected_input_device = next(
                (
                    index for label, index in audio_input_devices
                    if label == cbo_input_device.get()
                ),
                self.settings.get("microphone_index"),
            )
            selected_output_device = next(
                (
                    index for label, index in audio_output_devices
                    if label == cbo_output_device.get()
                ),
                self.settings.get("speaker_index"),
            )
            selected_camera_index = camera_label_to_index.get(
                cbo_camera_device.get(),
                int(self.settings.get("camera_index", 0)),
            )
            selected_camera_resolution = camera_resolution_values.get(
                cbo_camera_resolution.get(),
                (
                    int(self.settings.get("camera_width", 640)),
                    int(self.settings.get("camera_height", 480)),
                ),
            )
            camera_settings_changed = (
                selected_camera_index != int(self.settings.get("camera_index", 0))
                or selected_camera_resolution
                != (
                    int(self.settings.get("camera_width", 640)),
                    int(self.settings.get("camera_height", 480)),
                )
            )
            camera_was_active = self._is_camera_open()
            self.settings["microphone_index"] = selected_input_device
            self.settings["microphone_gain"] = float(
                microphone_gain_value.get()
            )
            self.settings["speaker_index"] = selected_output_device
            self.settings["camera_index"] = selected_camera_index
            self.settings["camera_width"], self.settings["camera_height"] = (
                selected_camera_resolution
            )
             
            self.target_county = selected_county
            self.min_alert_intensity = cb_min_intensity.get()
            self.trail_points = deque(list(self.trail_points)[:new_trail_len], maxlen=new_trail_len)
            self.dwell_threshold = float(val_dwell.get())
            self.pos_buffer = deque(self.pos_buffer, maxlen=int(val_smooth.get()))

            if save_changes:
                self.save_settings(self.settings)
                language_saved = not getattr(
                    self, "_settings_file_cleared", False
                )
            self.apply_language()
            self.check_and_apply_theme()
            # 如果主畫面（Main Window）也有 UI 切換函式，可於此處同步呼叫
            if hasattr(self, "update_main_ui_language"):
                self.update_main_ui_language()

            message_key = "msg_saved" if save_changes else "msg_applied"
            fallback_message = "設定已儲存" if save_changes else "設定已套用"
            message = I18N.get(chosen_lang, {}).get(
                message_key, fallback_message
            )
            action_text = "儲存" if save_changes else "套用"
            self.log_status(
                f"設定已{action_text}！當前語言：{chosen_lang}，監測地區：{self.target_county}。"
            )
            self.play_system_sound("execute")
            self.tts_speak(message)
            if save_changes:
                on_window_close()
            if camera_settings_changed and camera_was_active:
                self.init_camera(is_boot=False)

        def save_and_close():
            apply_settings(save_changes=True)

        def apply_without_saving():
            apply_settings(save_changes=False)

        reset_btn.config(command=reset_to_defaults)
        clear_settings_btn.config(command=clear_settings_file)
        save_btn.config(command=save_and_close)
        apply_btn.config(command=apply_without_saving)

        # 初次初始化介面文字
        refresh_ui_texts(curr_lang)
        self.apply_font_size()
        align_selected_tab_content()
    # =====================
    # 輔助工具：音效、TTS 與日誌
    # =====================
    def _sanitize_tts_text(self, text):
        """移除 emoji / pictograph symbols，避免 TTS 直接讀出 emoji 名稱。"""
        if text is None:
            return ""

        text = str(text)
        text = text.replace("\uFE0F", "").replace("\u200D", "")
        text = "".join(ch for ch in text if unicodedata.category(ch) != "Cf")
        text = re.sub(r"[\U0001F300-\U0001FAFF\U00002600-\U000026FF\U00002700-\U000027BF]", "", text)
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def _start_tts_worker(self):
        """Start a background TTS worker thread that owns a single SpVoice instance.
        The worker uses SVSFAsync|SVSFPurgeBeforeSpeak so new texts immediately interrupt previous playback.
        """
        # Already started?
        if getattr(self, 'tts_worker', None) is not None:
            return

        self.tts_queue = queue.Queue(maxsize=32)
        self.tts_worker_stop = threading.Event()

        def _worker():
            pythoncom.CoInitialize()
            speaker = None
            voices = None
            try:
                try:
                    speaker = win32com.client.Dispatch("SAPI.SpVoice")
                except Exception as e:
                    print(f"[TTS worker] Dispatch error: {e}")
                    return

                # apply initial settings
                try:
                    speaker.Rate = int(self.settings.get("voice_speed", 0))
                except Exception:
                    pass
                try:
                    speaker.Volume = int(self.settings.get("tts_volume", 100))
                except Exception:
                    pass

                # list voices for debug
                try:
                    voices = speaker.GetVoices()
                    print(f"[TTS worker] voices.Count={getattr(voices, 'Count', 'unknown')}")
                    # optional: print first few
                    try:
                        for i in range(min(3, getattr(voices, 'Count', 0))):
                            try:
                                v = voices.Item(i)
                                desc = None
                                try:
                                    desc = v.GetDescription()
                                except Exception:
                                    try:
                                        desc = v.Name
                                    except Exception:
                                        desc = str(v)
                                print(f"[TTS worker] voice[{i}] = {desc}")
                            except Exception as e:
                                print(f"[TTS worker] voice enumerate error: {e}")
                    except Exception:
                        pass
                except Exception as e:
                    print(f"[TTS worker] GetVoices error: {e}")

                flags = 0x1 | 0x2  # SVSFAsync | SVSFPurgeBeforeSpeak

                while not self.tts_worker_stop.is_set():
                    try:
                        text = self.tts_queue.get(timeout=0.2)
                    except queue.Empty:
                        continue

                    # Drain queue to get latest request (avoid backlog)
                    try:
                        while True:
                            text = self.tts_queue.get_nowait()
                    except queue.Empty:
                        pass

                    if not text:
                        continue

                    try:
                        # refresh runtime settings
                        try:
                            speaker.Rate = int(self.settings.get("voice_speed", 0))
                        except Exception:
                            pass
                        try:
                            speaker.Volume = int(self.settings.get("tts_volume", 100))
                        except Exception:
                            pass

                        try:
                            if voices and getattr(voices, 'Count', 0) > 0:
                                selected_voice = None
                                language = self.settings.get("language", "國語 (華語)")
                                preferred_terms = (
                                    ("English", ("english", "en-us", "en_us", "en-us")),
                                    ("台文 (臺灣閩南語)", ("taiwan", "zh-tw")),
                                    ("客語 (臺灣客家語)", ("taiwan", "zh-tw")),
                                    ("國語 (華語)", ("chinese", "zh-tw", "zh_cn")),
                                )
                                for voice_language, terms in preferred_terms:
                                    if voice_language != language:
                                        continue
                                    for i in range(getattr(voices, 'Count', 0)):
                                        candidate = voices.Item(i)
                                        try:
                                            description = str(candidate.GetDescription()).lower()
                                        except Exception:
                                            description = str(candidate).lower()
                                        if any(term in description for term in terms):
                                            selected_voice = candidate
                                            break
                                    break
                                speaker.Voice = selected_voice or voices.Item(0)
                        except Exception:
                            pass

                        clean_text = self._sanitize_tts_text(text)
                        if not clean_text:
                            continue
                        print(f"[TTS worker] speaking preview: {clean_text[:120]}")
                        # Purge previous and speak async so new Speak with purge can interrupt
                        speaker.Speak(clean_text, flags)
                    except Exception as e:
                        print(f"[TTS worker] speak error: {e}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        self.tts_worker = threading.Thread(target=_worker, daemon=True)
        self.tts_worker.start()

    def tts_speak(self, text):
        # ➕ 新增判斷：如果設定為不啟用，就直接返回不發聲
        if not self.settings.get("tts_enabled", True):
            return

        cleaned = self._sanitize_tts_text(self.localize_runtime_text(text))
        if not cleaned:
            return

        # ensure worker is running
        try:
            if not hasattr(self, 'tts_worker') or self.tts_worker is None:
                self._start_tts_worker()
        except Exception:
            pass

        # enqueue new text; if queue is full, drop oldest to ensure latest is delivered
        try:
            self.tts_queue.put_nowait(cleaned)
        except Exception:
            try:
                # drop one and retry
                _ = self.tts_queue.get_nowait()
                self.tts_queue.put_nowait(cleaned)
            except Exception:
                pass

        # return immediately; worker handles playback (and interruption)

    def play_system_sound(self, sound_type):
        """播放系統回饋音效，幫助長者確認點擊"""
        if threading.current_thread() is not threading.main_thread():
            self._post_camera_ui(
                lambda: self.play_system_sound(sound_type)
            )
            return

        try:
            tone = {
                "click": (1000, 120),
                "execute": (1200, 150),
                "warn": (600, 300),
            }.get(sound_type)
            if tone is None:
                return
            winsound.Beep(*tone)
        except Exception as exc:
            print(f"播放系統回饋音效失敗: {exc}")

    def localize_runtime_text(self, text):
        """Translate legacy runtime messages that predate the language preset system."""
        language = self.settings.get("language", "國語 (華語)")
        runtime_translations = I18N.get(language, {}).get("RUNTIME_TRANSLATIONS", {})
        if isinstance(text, str):
            translated = runtime_translations.get(text, text)
            if translated == text and language == "English":
                for source, target in (
                    ("📊 當前顯示樣式：", "📊 Current display style: "),
                    ("目前難度：", "Current difficulty: "),
                    ("剩餘機會：", "Chances left: "),
                    ("本關經過時間：", "Current round time: "),
                    ("剩餘時間 ", "Time left: "),
                    ("完成：", "Completed: "),
                    ("請做：", "Do: "),
                    ("總共成功闖過：", "Total rounds cleared: "),
                    ("平均每關時間：", "Average time per round: "),
                ):
                    if text.startswith(source):
                        translated = target + text[len(source):]
                        if source == "目前難度：" and translated.endswith(" 個數字"):
                            translated = translated[:-4] + " numbers"
                        elif source == "總共成功闖過：" and translated.endswith(" 關"):
                            translated = translated[:-2] + " rounds"
                        elif source == "平均每關時間：" and translated.endswith(" 秒"):
                            translated = translated[:-2] + " seconds"
                        elif source == "本關經過時間：" and translated.endswith(" 秒"):
                            translated = translated[:-2] + " seconds"
                        break
                if text.startswith("今日已完成簽到！累計連續簽到 "):
                    translated = text.replace("今日已完成簽到！累計連續簽到 ", "Today's check-in is complete! Current streak: ").replace(" 天。", " days.")
                elif text.startswith("若沒有回應，") and text.endswith(" 秒後將通知家屬。"):
                    seconds = text[len("若沒有回應，"): -len(" 秒後將通知家屬。")]
                    translated = f"If there is no response, your family will be notified in {seconds} seconds."
                elif text.startswith("報告長輩，現在"):
                    translated = "Here is the current weather: " + text[len("報告長輩，現在"):]
                elif text.startswith("語音辨識出錯: "):
                    translated = "Voice recognition error: " + text[len("語音辨識出錯: "):]
                elif text.startswith("主畫面已切換為："):
                    translated = "The main screen style is now: " + text.split("：", 1)[1]
                elif text.startswith("已更換，當前功能是："):
                    translated = "Changed. The current feature is: " + text.split("：", 1)[1]
                elif text.startswith("猜錯囉，沒關係，我們再來一次。"):
                    translated = "That was not correct, but that's okay. Let's try again."
                elif text.startswith("挑戰結束。您今天總共成功闖過 "):
                    translated = text.replace("挑戰結束。您今天總共成功闖過 ", "Challenge complete. You successfully cleared ") \
                        .replace(" 關，再接再厲！要再玩一次嗎？", " rounds today. Great effort! Would you like to play again?")
            return translated
        return text

    def localize_window_widgets(self, widget):
        """Apply English labels to widgets created by legacy dialog code."""
        self.apply_font_size()

        def localize_children(parent):
            if self.settings.get("language") == "English":
                if isinstance(parent, (tk.Label, tk.Button, tk.Checkbutton)):
                    current_text = parent.cget("text")
                    translated = self.localize_runtime_text(current_text)
                    if translated != current_text:
                        parent.configure(text=translated)
            for child in parent.winfo_children():
                localize_children(child)

        localize_children(widget)

    def log_status(self, text):
        """即時更新首頁最下方的長者提示列"""
        text = self.localize_runtime_text(text)
        prefix = "Status: " if self.settings.get("language") == "English" else "提示："
        self.transition_text(self.status_bar, f"{prefix}{text}")

    def safe_exit(self):
        """安全關閉系統，釋放相機資源並自動保存"""
        if self._shutting_down:
            return
        if not messagebox.askyesno(
            self.localize_runtime_text("關閉程式"),
            self.localize_runtime_text("確定要關閉程式嗎？"),
            parent=self.root,
            default=messagebox.NO,
        ):
            return

        print("【系統】偵測到退出指令，儲存並釋放資源中...")
        self.loop_active = False
        with self._camera_preview_lock:
            self._shutting_down = True
            preview_stop_events = list(self._camera_preview_stop_events)
            preview_captures = list(self._camera_preview_captures.values())
            self._camera_preview_stop_events.clear()
            self._camera_preview_captures.clear()
            for stop_event in preview_stop_events:
                stop_event.set()

        self.save_settings(self.settings)
        self.camera_ready = False
        capture = self.cap
        self.cap = None
        if capture is not None:
            threading.Thread(
                target=self._release_camera_capture,
                args=(capture,),
                daemon=True,
            ).start()
        for preview_capture in preview_captures:
            threading.Thread(
                target=self._release_camera_capture,
                args=(preview_capture,),
                daemon=True,
            ).start()

        # Signal TTS worker to stop and wake it if necessary
        try:
            if getattr(self, 'tts_worker_stop', None) is not None:
                try:
                    self.tts_worker_stop.set()
                except Exception:
                    pass
                try:
                    # put a sentinel to wake the worker
                    if getattr(self, 'tts_queue', None) is not None:
                        self.tts_queue.put_nowait(None)
                except Exception:
                    pass
        except Exception:
            pass

        faulthandler.disable()
        self._crash_log_file.close()
        self.root.destroy()

if __name__ == "__main__":
    pyautogui.PAUSE = 0
    pyautogui.FAILSAFE = True
    root = tk.Tk()
    app = SmartCareSystem(root)
    root.mainloop()