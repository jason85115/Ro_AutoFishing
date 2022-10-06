from PIL import Image
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from win32 import win32gui
import win32ui, win32con, win32api # win32 gui 必須一起進來
import numpy as np
import pyautogui
from time import time, sleep
import io
import base64


class Ro_AutoFinishing:
    def __init__(self) -> None:
        self.hwnd = None
        self.hwnd_position = { "x": 0, "y":0, "r":0, "b":0 }
        self.global_mouse_position = {"x":0, "y":0}
        self.fish_mouse_position = {"x":0, "y":0}
        self.red_envelope_mouse_position = {"x":0, "y":0, "r":0, "b":0}

        self.fishing = False
        self.get_red_envelopeing = False
        self.chat_get_red_envelopeing = False
        self.tasking = True

        self.Ro_AutoFinishing_GUI = tk.Tk()
        self.init_GUI()
        self.init_listner()
        self.start_main_job()

    def init_GUI(self) -> None:
        self.Ro_AutoFinishing_GUI.title("Ro_AutoFinishing")
        self.Ro_AutoFinishing_GUI.geometry('275x215')

        # first and second row
        self.role = tk.Label(self.Ro_AutoFinishing_GUI, text="asd")
        self.role.grid(row=0, column=0, rowspan=2,columnspan=2)

        # 設定圖片
        blankImage = np.zeros((50,60,3), dtype=np.uint8) # 產生黑色圖片
        blankImage = Image.fromarray(blankImage) # 產生黑色圖片
        imgBytes = io.BytesIO() # 抽取Bytes
        blankImage.save(imgBytes, format='PNG') # 抽取Bytes
        imgBytes = imgBytes.getvalue() # 抽取Bytes
        base64Image = base64.b64encode(imgBytes) # 轉成 base64
        self.role_image = tk.PhotoImage(data=base64Image) # 設定圖片
        self.role.config(image=self.role_image) # 設定圖片

        tk.Label(self.Ro_AutoFinishing_GUI, text="請選擇視窗:").grid(row=0, column=2, sticky="w")

        self.refresh_windows_button = ttk.Button(self.Ro_AutoFinishing_GUI, text="刷新視窗", width=10)
        self.refresh_windows_button.grid(row=0, column=3)

        self.select_windows = ttk.Combobox(self.Ro_AutoFinishing_GUI, values = self._get_all_windows())
        self.select_windows.grid(row=1, column=2, columnspan=2, pady=5, sticky="we")

        # third row
        tk.Label(self.Ro_AutoFinishing_GUI, text="x").grid(row=2, column=0, columnspan=2)
        self.x = tk.Entry(self.Ro_AutoFinishing_GUI)
        self.x.grid(row=2, column=2, padx=0, pady=5, columnspan=2, sticky="we")

        # fourth row
        tk.Label(self.Ro_AutoFinishing_GUI, text="y").grid(row=3, column=0, columnspan=2)
        self.y = tk.Entry(self.Ro_AutoFinishing_GUI)
        self.y.grid(row=3, column=2, padx=0, pady=5, columnspan=2, sticky="we")

        # fifth row
        tk.Label(self.Ro_AutoFinishing_GUI, text="hint: 點選 f 將會抓取滑鼠位置 (水上)", foreground="red").grid(row=4, column=0, columnspan=4, pady=[0,5])

        # sixth row
        self.stop_button = ttk.Button(self.Ro_AutoFinishing_GUI, text="stop")
        self.stop_button.grid(row=5, column=0, columnspan=4)

        self.prevent_save_power_value = tk.IntVar()
        self.prevent_save_power_checkbox = tk.Checkbutton(self.Ro_AutoFinishing_GUI, text="防省電", variable=self.prevent_save_power_value)
        self.prevent_save_power_checkbox.grid(row=5, column= 3)
        
        # seven row
        tk.Label(self.Ro_AutoFinishing_GUI, text="hint: s釣魚開始 r or c搶紅包 p全部結束 ", foreground="red").grid(row=6, column=0, columnspan=4, pady=[0,5])

        # eight row
        tk.Label(self.Ro_AutoFinishing_GUI, text="目前狀態：").grid(row=7, column=0, pady=[0,0])

        self.fishing_label = tk.Label(self.Ro_AutoFinishing_GUI, text="釣魚X", foreground="red")
        self.fishing_label.grid(row=7, column=1, pady=[0,0])

        self.get_red_envelopeing_label = tk.Label(self.Ro_AutoFinishing_GUI, text="紅包X", foreground="red")
        self.get_red_envelopeing_label.grid(row=7, column=2, pady=[0,0])

        self.chat_get_red_envelopeing_label = tk.Label(self.Ro_AutoFinishing_GUI, text="聊紅包X", foreground="red")
        self.chat_get_red_envelopeing_label.grid(row=7, column=3, pady=[0,0])
    
    def init_listner(self) -> None:
        self.Ro_AutoFinishing_GUI.bind('<KeyPress>', self._on_Key_Press)
        self.refresh_windows_button.bind('<Button-1>', self._refresh_all_windows)
        self.select_windows.bind("<<ComboboxSelected>>", self._windows_selected)
        self.Ro_AutoFinishing_GUI.protocol("WM_DELETE_WINDOW", self._terminate)
        self.stop_button.bind('<Button-1>', self.stop_click)

    def start_fishing(self):
        if self.hwnd == None:
            messagebox.showinfo("提示", "請先選取遊戲視窗。")
            return

        if (self.x.get() == "" or self.x.get() == None or self.y.get() == "" or self.y.get() == None):
            messagebox.showinfo("提示", "請先抓取釣魚位置，採用鍵盤f按鈕。")
            return 

        move_x = self.global_mouse_position["x"]
        move_y = self.global_mouse_position["y"]
        pyautogui.mouseDown(move_x, move_y)
        sleep(0.03432 + (np.random.randn() / 100)) 
        pyautogui.mouseUp(move_x, move_y)

        if self.fishing:
            return

        if not self._get_in_game_position():
            return

        self.fishing = True
        self.fishing_label.config(foreground="green")
        self.fishing_label.config(text="釣魚V")
        return

    def start_get_red_envelope(self):
        if self.hwnd == None:
            messagebox.showinfo("提示", "請先選取遊戲視窗。")
            return
        
        if self.chat_get_red_envelopeing:
            messagebox.showinfo("提示", "請先關閉聊天搶紅包。")
            return

        width = self.hwnd_position["r"] - self.hwnd_position["x"]
        height = self.hwnd_position["b"] - self.hwnd_position["y"]

        self.red_envelope_mouse_position["x"] = int(width * 0.563)
        self.red_envelope_mouse_position["y"] = int(height * 0.592)
        self.red_envelope_mouse_position["r"] = int(width * 0.632)
        self.red_envelope_mouse_position["b"] = int(height * 0.757)
        
        self.get_red_envelopeing = True
        self.get_red_envelopeing_label.config(foreground="green")
        self.get_red_envelopeing_label.config(text="紅包V")
        return

    def start_chat_get_red_envelope(self):
        if self.hwnd == None:
            messagebox.showinfo("提示", "請先選取遊戲視窗。")
            return

        if self.get_red_envelopeing:
            messagebox.showinfo("提示", "請先關閉一般搶紅包。")
            return

        self.chat_get_red_envelopeing = True
        self.chat_get_red_envelopeing_label.config(foreground="green")
        self.chat_get_red_envelopeing_label.config(text="聊紅包V")

    def stop_click(self, event):
        self.fishing = False
        self.get_red_envelopeing = False
        self.chat_get_red_envelopeing = False
        self.get_red_envelopeing_label.config(foreground="red")
        self.get_red_envelopeing_label.config(text="紅包X")
        self.fishing_label.config(foreground="red")
        self.fishing_label.config(text="釣魚X")
        self.chat_get_red_envelopeing_label.config(foreground="red")
        self.chat_get_red_envelopeing_label.config(text="聊紅包X")
    
    def stop(self):
        self.fishing = False
        self.get_red_envelopeing = False
        self.chat_get_red_envelopeing = False
        self.get_red_envelopeing_label.config(foreground="red")
        self.get_red_envelopeing_label.config(text="紅包X")
        self.fishing_label.config(foreground="red")
        self.fishing_label.config(text="釣魚X")
        self.chat_get_red_envelopeing_label.config(foreground="red")
        self.chat_get_red_envelopeing_label.config(text="聊紅包X")

    def start_main_job(self):
        def prevent_save_power():
            width = self.hwnd_position["r"] - self.hwnd_position["x"]
            height = self.hwnd_position["b"] - self.hwnd_position["y"]

            refresh_x = int(width * 0.94)
            refresh_y = int(height * 0.77)

            refresh_x = self.hwnd_position["x"] + refresh_x
            refresh_y = self.hwnd_position["y"] + refresh_y
            pyautogui.mouseDown(refresh_x, refresh_y)
            sleep(0.03432 + (np.random.randn() / 100)) 
            pyautogui.mouseUp(refresh_x, refresh_y)

            
        start_time = 0
        while(self.tasking):
            if start_time == 0:
                start_time = time()
            screen_shot = self._get_screen_shot()

            if self.chat_get_red_envelopeing:
                now_time = time()
                if (now_time - start_time) >= 30 and self.prevent_save_power_value.get() == 1:
                    # 防止進入省電模式
                    prevent_save_power()
                    start_time = 0
                
                
                h, w, _ = np.array(screen_shot).shape
                check_padding = h / 13
                chat_red_position = (int)(w * (1.17 / 4))

                screen_shot_chat_room = screen_shot.crop((0, 0, (w * 1/3), h))
                screen_shot_chat_room = np.array(screen_shot_chat_room)
                h_room, w_room, _ = screen_shot_chat_room.shape

                for idx in range(0, (int)(h_room // check_padding)):
                    check_interval_center = h_room - 1 - (int)(idx * check_padding)
                    aver_dot = np.average(screen_shot_chat_room[check_interval_center , chat_red_position-5:chat_red_position+5], axis=0)
                    move_x = chat_red_position + self.hwnd_position["x"]
                    move_y = check_interval_center + self.hwnd_position["y"]
                    if aver_dot[0] >= 240 and (aver_dot[1] >= 180 and aver_dot[1] <= 200) and (aver_dot[2] >= 75 and aver_dot[2] <= 105):
                        pyautogui.mouseDown(move_x, move_y)
                        sleep(0.03432 + (np.random.randn() / 100)) 
                        pyautogui.mouseUp(move_x, move_y)
                        sleep(0.5)
                        prevent_save_power()
                        
                    
            if self.get_red_envelopeing:
                now_time = time()

                if (now_time - start_time) >= 30 and self.prevent_save_power_value.get() == 1:
                    # 防止進入省電模式
                    prevent_save_power()
                    start_time = 0

                get_red = False
                screen_shot_red = screen_shot.crop((self.red_envelope_mouse_position["x"], self.red_envelope_mouse_position["y"], self.red_envelope_mouse_position["r"], self.red_envelope_mouse_position["b"]))
                screen_shot_red = np.array(screen_shot_red)
                h, w, _ = np.array(screen_shot_red).shape
                aver_dot = np.average(screen_shot_red[h//2+15, w//2-5:w//2+5], axis=0)
                
                move_x = w//2 + self.hwnd_position["x"] + self.red_envelope_mouse_position["x"]
                move_y = h//2+15 + self.hwnd_position["y"] + self.red_envelope_mouse_position["y"]

                # 紅 紅包 230 100 93
                if aver_dot[0] >= 230 and (aver_dot[1] >= 100 and aver_dot[1] <= 120) and (aver_dot[2] >= 80 and aver_dot[2] <= 110):
                    get_red = True
                    pyautogui.click(move_x, move_y, clicks=3)

                # 黃 紅包 234 156 94
                if aver_dot[0] >= 230 and (aver_dot[1] >= 140 and aver_dot[1] <= 170) and (aver_dot[2] >= 80 and aver_dot[2] <= 110):
                    get_red = True
                    pyautogui.click(move_x, move_y, clicks=3)

                if get_red:
                    get_red = False
                    sleep(1.0)
                    h, w, _ = np.array(screen_shot).shape
                    dwidth = (int)(w / 10)
                    # 向右移動
                    pyautogui.click(move_x + dwidth, move_y, clicks=3)

            if self.fishing:
                get_fish = False
                # 242 89 90  239 88 88
                h, w, _ = np.array(screen_shot).shape
                dheight = (int)(h / 15)
                screen_shot_fishing = np.array(screen_shot)[self.fish_mouse_position["y"] - dheight, self.fish_mouse_position["x"]]
                if screen_shot_fishing[0] >= 220 and screen_shot_fishing[1] <= 100 and screen_shot_fishing[2] <= 100:
                    dheight = (int)(h / 10)
                    get_fish = True
                    move_x = self.global_mouse_position["x"]
                    move_y = self.global_mouse_position["y"] - dheight
                    pyautogui.mouseDown(move_x, move_y)
                    sleep(0.03432 + (np.random.randn() / 100)) 
                    pyautogui.mouseUp(move_x, move_y)
                    sleep(1.0)
                    dheight = (int)(h / 4)
                    pyautogui.mouseDown(move_x, move_y + dheight)
                    sleep(0.03432 + (np.random.randn() / 100)) 
                    pyautogui.mouseUp(move_x, move_y + dheight)


                if get_fish:
                    get_fish = False
                    sleep(1.0)
                    pyautogui.mouseDown(self.global_mouse_position["x"], self.global_mouse_position["y"])
                    sleep(0.03432 + (np.random.randn() / 100)) 
                    pyautogui.mouseUp(self.global_mouse_position["x"], self.global_mouse_position["y"])

            self.Ro_AutoFinishing_GUI.update_idletasks()
            self.Ro_AutoFinishing_GUI.update()
            
    def _fishing_job(self):
        return

    def _get_red_envelope_job(self):
        return

    def _show_status(self):
        return
    
    def _get_cursor_position(self):
        x,y = win32gui.GetCursorPos()
        self.x.delete(0, tk.END)
        self.x.insert(0, x)

        self.y.delete(0, tk.END)
        self.y.insert(0, y)

        self.global_mouse_position["x"] = x
        self.global_mouse_position["y"] = y
        return

    def _get_in_game_position(self) -> bool:
        game_x = self.global_mouse_position["x"] - self.hwnd_position["x"]
        game_y = self.global_mouse_position["y"] - self.hwnd_position["y"]

        if game_x <= 0 or game_y <=0 or self.global_mouse_position["x"] >= self.hwnd_position["r"] or self.global_mouse_position["y"] >= self.hwnd_position["b"]:
            messagebox.showinfo("提示", "請選擇遊戲框內。")
            return False

        self.fish_mouse_position["x"] = game_x
        self.fish_mouse_position["y"] = game_y

        return True

    def _on_Key_Press(self, event) -> None:
        if event.char == 's':
            self.start_fishing()
        elif event.char == 'p':
            self.stop()
        elif event.char == 'f':
            self._get_cursor_position()
        elif event.char == 'r':
            self.start_get_red_envelope()
        elif event.char == 'c':
            self.start_chat_get_red_envelope()

    def _windows_selected(self, event) -> None:
        
        # 檢查選擇的名稱
        if event.widget.get() == "":
            return
        self.hwnd = event.widget.get().split(',')[0]

        # 設定窗口資訊
        x,y,r,b = win32gui.GetWindowRect(self.hwnd)
        self.hwnd_position["x"] = x
        self.hwnd_position["y"] = y
        self.hwnd_position["r"] = r
        self.hwnd_position["b"] = b

        # 設定圖片
        screen_shot = self._get_screen_shot().crop((15,45,105,130))
        screen_shot = screen_shot.resize((50,50))
        imgBytes = io.BytesIO() # 抽取Bytes
        screen_shot.save(imgBytes, format='PNG') # 抽取Bytes
        imgBytes = imgBytes.getvalue() # 抽取Bytes
        base64Image = base64.b64encode(imgBytes) # 轉成 base64
        self.role_image = tk.PhotoImage(data=base64Image) # 設定圖片
        self.role.config(image=self.role_image) # 設定圖片
        
        return
    
    def _get_all_windows(self) -> list:
        def winEnumHandler(hwnd, resultList):
            if win32gui.IsWindowVisible( hwnd ):
                name = win32gui.GetWindowText(hwnd)
                if name == "":
                    pass
                else:
                    if not name.find("RO"):
                        tmp = str(hwnd) + "," + name
                        resultList.append(tmp)
        results = []
        win32gui.EnumWindows(winEnumHandler, results)
        return results

    def _refresh_all_windows(self, event):
        all_windows = [""]
        all_windows.extend(self._get_all_windows())
        self.select_windows.config(values= all_windows)
        self.select_windows.current(0)

    def _get_screen_shot(self):
        im = pyautogui.screenshot().crop((self.hwnd_position["x"], self.hwnd_position["y"], self.hwnd_position["r"], self.hwnd_position["b"]))
        return im

    def _terminate(self):
        self.tasking = False
        self.Ro_AutoFinishing_GUI.destroy()

        
if __name__ == '__main__':
    gui = Ro_AutoFinishing()
    