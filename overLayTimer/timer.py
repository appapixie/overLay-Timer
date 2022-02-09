# tkinter 
import tkinter as tk
import tkinter.font
import ttkbootstrap as ttk
from tkinter import filedialog
from tkinter.colorchooser import askcolor

# lib
from pygame import mixer
from win32gui import *
import configparser
import subprocess
import keyboard
import requests
import time
import sys
import os
import re


# overLay Class
class overLay:
    def __init__(self):
        
        mixer.init()
        self.timer = tk.Tk()
        self.timer.title('Timer')
        
        # GUI 설정
        self.timer.resizable(False, False)
        self.timer.overrideredirect(True)
        self.timer.attributes('-alpha', 0)
        self.taskManagerCheck()
        self.style = ttk.Style('solar')
        
        # 폼 투명화
        self.transColor = '#002B36'
        self.timer.configure(bg=self.transColor)
        self.timer.wm_attributes('-transparentcolor', self.transColor)
        font = tkinter.font.Font(size=30, weight='bold')
        
        self.nowTimeVal=tk.BooleanVar()
        self.setGui=tk.BooleanVar()
        
        # 라벨
        t = time.strftime('%I:%M:%S %p')
        self.nowLabel=tk.Label(self.timer, text=f'{t}', fg='white', bg=self.transColor, font=tkinter.font.Font(size=12, weight='bold'))
        self.timeLabel=tk.Label(self.timer, text='0 : 0', fg='white', bg=self.transColor, font=font)
        self.nowLabel.pack()
        self.timeLabel.pack()
        
        # 우클릭 메뉴
        self.menu = tk.Menu(self.timer, tearoff=0)
        self.subMenu = tk.Menu(self.timer, tearoff=0)
        
        self.menu.add_cascade(label="타이머 설정", menu=self.subMenu)
        self.subMenu.add_checkbutton(label="현재시간 표시", onvalue=True, offvalue=False, variable=self.nowTimeVal)
        self.subMenu.add_command(label='타이머 폰트 색설정', command=self.selectColor)
        self.subMenu.add_checkbutton(label="타이머 투명도 설정", onvalue=True, offvalue=False, command=self.newGui, variable=self.setGui)
        self.subMenu.add_checkbutton(label="타이머 볼륨 설정", onvalue=True, offvalue=False, command=self.newGui, variable=self.setGui)
        self.menu.add_separator()
        self.menu.add_command(label='팅패확인' , command=self.patch)
        self.menu.add_separator()
        self.menu.add_command(label='설정 불러오기', command=self.load)
        self.menu.add_separator()
        self.menu.add_command(label='종료', command=self.timer.destroy)
        self.timer.bind('<Button-3>', self.menuOpen)
        
        self.loadVar() # 변수들
        self.nowTime() # 현재시간 업뎃
        self.windowPosition() # 오버레이 위치 업뎃
        self.timer.mainloop() 
            
    # 설정 불러오기
    def load(self):
        try:
            self.fileName=filedialog.askopenfilename(initialdir=os.getcwd(), 
                                                    title="타이머 설정파일을 골라주세요.", 
                                                    filetypes = (("timer","*.ini"),("all files","*.*")))
            
            if self.fileName == "":
                return
            
            keyboard.clear_all_hotkeys()
            self.loadVar()
        except:
            pass
    
    # 따이머 리셋
    def timerReset(self):
        self.minute, self.second = self.defaultMinute, self.defaultSecond
        self.timeLabel['text'] = f'{self.minute} : {self.second:.0f}'
        return
        
    # 변수
    def loadVar(self):
        # os.chdir(os.path.dirname(os.path.abspath(__file__)))

        config = configparser.ConfigParser()
        
        try:
            config.read('timer.ini')
            
            self.startKey = config['hotkey']['start']
            self.stopKey = config['hotkey']['stop']
            self.resetKey = config['hotkey']['reset']
            
            # 프리셋
            self.presetKey = []
            self.minutePreset = []
            self.secondPreset = []

            for i in range(5):
                self.presetKey.append(config['hotkey'][f'preset{i+1}'])
                self.minutePreset.append(int(config[f'preset{i+1}']['minute']))
                self.secondPreset.append(int(config[f'preset{i+1}']['second']))
                
            # 오버레이 위치
            self.windowX = int(config['overlayLocation']['x'])
            self.windowY = int(config['overlayLocation']['y'])

            # 기본 타이머값
            self.defaultMinute = int(config['timer']['defaultMinute'])
            self.defaultSecond = int(config['timer']['defaultSecond'])
            
            if self.defaultSecond >= 60:
                self.defaultSecond = 0
                self.defaultMinute += 1
            
            self.minute = self.defaultMinute
            self.second = self.defaultSecond
            
            self.transValue = int(config['etc']['trans'])
            self.volumeValue = int(config['etc']['volume'])
            
            
            if self.transValue < 10:
                self.transValue = 10
            
            self.timeLabel['text'] = f'{self.minute} : {self.second}'
            self.timer.attributes("-alpha", self.transValue / 100)
            mixer.music.set_volume(self.volumeValue / 100)
            
            self.afterID = ''
            self.timerCheck = False
            
            # 핫키 설정
            keyboard.add_hotkey(self.startKey, lambda: self.duplicateCheck())
            keyboard.add_hotkey(self.stopKey, lambda: self.timerStop())
            keyboard.add_hotkey(self.resetKey, lambda: self.timerReset())
                
            keyboard.add_hotkey(self.presetKey[0], lambda: self.timerSet(key=0))
            keyboard.add_hotkey(self.presetKey[1], lambda: self.timerSet(key=1))
            keyboard.add_hotkey(self.presetKey[2], lambda: self.timerSet(key=2))
            keyboard.add_hotkey(self.presetKey[3], lambda: self.timerSet(key=3))
            keyboard.add_hotkey(self.presetKey[4], lambda: self.timerSet(key=4))
            
        except: # 설정파일에 오류가 있으면 다시 설정파일 맨듬
            
            config['timer'] = {
                'defaultMinute': 30,
                'defaultSecond': 0
            }
            
            config['etc'] = {
                'volume': 100,
                'trans': 100
            }
            
            config['overlayLocation'] = {
                'x': 50,
                'y': 180
            }

            config['preset1'] = {
                'minute': '20',
                'second': '0'
            }
            
            config['preset2'] = {
                'minute': '15',
                'second': '0'
            }
            
            config['preset3'] = {
                'minute': '10',
                'second': '0'
            }
            
            config['preset4'] = {
                'minute': '5',
                'second': '0'
            }
            
            config['preset5'] = {
                'minute': '2',
                'second': '0'
            }
            
            config['hotkey'] = {
                'start': 'F10',
                'stop': 'F11',
                'reset': 'F12',
                'preset1': 'ctrl+1',
                'preset2': 'ctrl+2',
                'preset3': 'ctrl+3',
                'preset4': 'ctrl+4',
                'preset5': 'ctrl+5'
            }
            
            with open('./timer.ini', 'w') as f:
                config.write(f)
                tkinter.messagebox.showwarning(title='Error', message='ini 파일이 손상되어 재설정하였습니다.')
                sys.exit()
                   
        return
    
    # 따이머 프리셋
    def timerSet(self, key):
        self.minute =self. minutePreset[key]
        self.second = self.secondPreset[key]
        self.defaultMinute = self.minutePreset[key]
        self.defaultSecond = self.secondPreset[key]
        self.timeLabel['text'] = f'{self.minute} : {self.second:.0f}'
        return
    
    # 투명도 볼륨 설정 구이
    def newGui(self):
        
        # 중복실행 방지
        if self.setGui.get() is False:
            self.option.destroy()
            return
        
        screenWidth = int(self.timer.winfo_screenwidth() / 2 - 260 / 2)
        screenHeight = int(self.timer.winfo_screenheight() / 2 - 130 / 2)
        
        self.option=tk.Toplevel()
        self.option.geometry(f'260x130+{screenWidth}+{screenHeight}')
        self.option.resizable(False, False)
        self.option.wm_attributes('-topmost', 1)
                
        self.trnasVar=tk.IntVar()
        self.volumeVar=tk.IntVar()
        
        
        # 투명도
        self.transLabel=ttk.Label(self.option, text='투명도 : ')
        self.transLabel.place(x=10, y=15)
        
        self.transSpinBox=ttk.Spinbox(self.option, from_=10, to=100, width=4, command=self.transSpin, bootstyle="success")
        self.transSpinBox.place(x=180, y=10)
        self.transSpinBox.set(self.transValue)
        
        self.transScale=ttk.Scale(self.option, length=90, variable=self.trnasVar, command=self.trans, orient="horizontal", from_=10, to=100, bootstyle="success")
        self.transScale.place(x=70, y=18)
        self.transScale.set(self.transValue)
        
        # 볼륨
        self.volumeLabel=ttk.Label(self.option, text='볼  륨 : ')
        self.volumeLabel.place(x=10, y=55)
        
        self.volumeSpinBox=ttk.Spinbox(self.option, from_=0, to=100, width=4, command=self.volSpin, bootstyle="success")
        self.volumeSpinBox.place(x=180, y=50)
        self.volumeSpinBox.set(self.volumeValue)
        
        self.volumeScale=ttk.Scale(self.option, length=90, variable=self.volumeVar, command=self.vol, orient="horizontal", from_=0, to=100, bootstyle="success")
        self.volumeScale.place(x=70, y=58)
        self.volumeScale.set(self.volumeValue)
        
        # 볼륨 테스트 버튼
        self.buttons=ttk.Button(self.option, text='테스트', command=self.volTest, width=6, bootstyle="success")
        self.buttons.place(x=180, y=90)
        
        self.setApply()
        self.option.protocol("WM_DELETE_WINDOW", self.closeWindow)
        return
        
    # 투명도 조절 스핀버튼
    def transSpin(self):
        self.transValue = self.transSpinBox.get()
        self.transScale.set(self.transSpinBox.get())
        return
    
    # 투명도 슬라이더 
    def trans(self, value):
        try:
            self.transValue = float(value)
            self.transSpinBox.set(f'{float(value):.0f}')
        except:
            pass
        return
        
    # 볼륨 스핀버튼
    def volSpin(self):
        self.volumeValue = self.volumeSpinBox.get()
        self.volumeScale.set(self.volumeSpinBox.get())
    
    # 볼륨 슬라이더
    def vol(self, value):
        try:
            self.volumeValue = float(value)
            self.volumeSpinBox.set(f'{float(value):.0f}')
        except:
            pass
                
    # 설정 적용입니땅.
    def setApply(self):
        afterID = self.option.after(100, self.setApply)
        self.timer.attributes("-alpha", float(self.transValue) / 100)
        mixer.music.set_volume(float(self.volumeValue) / 100)
        return
    
    # 볼륨 테스트
    def volTest(self):
        mixer.music.load('alram.mp3')
        mixer.music.play()
        return
    
    # 옵션 닫았을때
    def closeWindow(self):
        self.setGui.set(False)
        self.option.destroy()
        return
    
    # 타이머 중복실행 방지
    # 중복실행 시 시간이 빨리 감
    def duplicateCheck(self):
        if self.timerCheck is True:
            return
        self.timerCheck = True
        self.timerStart()
        return
        
    # 따이머 시작
    def timerStart(self):
        # os.chdir(os.path.dirname(os.path.abspath(__file__)))
        self.afterID = self.timer.after(1000, self.timerStart)
        
        operationTime = self.minute * 60 + self.second - 1
        self.minute = operationTime // 60
        self.second = operationTime % 60
        
        self.timeLabel['text'] = f'{self.minute} : {self.second:.0f}'
        
        # 0 : 0 초 알람
        if self.minute <= 0 and self.second <= 0:
            self.minute = self.defaultMinute
            self.second = self.defaultSecond
            
            try:
                mixer.music.load('alram.mp3')
                mixer.music.play()
            except:
                pass
        return
    
    # 따이머 중지
    def timerStop(self):
        try:
            self.timer.after_cancel(self.afterID)
            self.timerCheck=False
        except:
            pass
    
    # 작업관리자로 실행했는지 체크
    def taskManagerCheck(self):
        
        try:
            output = subprocess.check_output("bcdedit >> nul", shell=True)
            print(output)
        except:
            tkinter.messagebox.showwarning(title='Error', message='작업관리자로 실행해주세요.')
            sys.exit()
            
        return
        
    # overLay 위치
    def windowPosition(self):
        self.timer.after(1, self.windowPosition)
        
        # 활성화된 창
        try:
            windowHandle = GetForegroundWindow()
            windowTitle = GetWindowText(windowHandle)
            
            if windowTitle != 'MapleStory':
                self.timer.wm_attributes('-topmost', 0)
            else:
                self.timer.wm_attributes('-topmost', 1)

            # 창하고 같이 이동
            windowHandle = FindWindow(None, 'MapleStory')
            windowRect = GetWindowRect(windowHandle)
            
            x, y = int(windowRect[0])+self.windowX, int(windowRect[1])+self.windowY
            self.timer.geometry(f'+{x}+{y}')
        except:
            pass
        
    # 현재 시간 업뎃 및 표시
    def nowTime(self):
        if self.nowTimeVal.get() is True:
            self.nowLabel.pack()
            self.timeLabel.pack_forget()
            self.timeLabel.pack()
        else:
            self.nowLabel.pack_forget()
        
        self.timer.after(1000, self.nowTime)
        t = time.strftime('%I:%M:%S %p')
        self.nowLabel['text'] = t
        self.timer.wm_attributes('-topmost', 0)
        return
    
    # 우클릭 메뉴
    def menuOpen(self, event):
        try:
            self.menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            self.menu.grab_release()
        return
    
    # 폰트 색 설정
    def selectColor(self):
        result = askcolor(color='#FFFFFF', title='폰트 색 설정') 
        color = result[1]
        self.timeLabel.configure(fg=color)
        self.nowLabel.configure(fg=color)
        return
        
    # maplestory API
    def patch(self):
        try:
            URL = 'http://api.maplestory.nexon.com/soap/maplestory.asmx'
            res = requests.get(URL)

            if res.status_code == 200:
                GetInspectionInfo = """<?xml version="1.0" encoding="utf-8"?>
                    <soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
                    <soap12:Body>
                        <GetInspectionInfo xmlns="https://api.maplestory.nexon.com/soap/" />
                    </soap12:Body>
                    </soap12:Envelope>
                """
                headers = {'Content-Type': 'application/soap+xml'}
                resp = requests.post(URL, data=GetInspectionInfo, headers=headers)
                scheduledStart = re.findall("<startDateTime>(.*?)</startDateTime>",resp.text)
                scheduledEnd = re.findall("<endDateTime>(.*?)</endDateTime>",resp.text)
                content = re.findall("<strObstacleContents>(.*?)</strObstacleContents>",resp.text)
                
                tkinter.messagebox.showinfo(title='메이플스토리 팅패치 정보', 
                                            message=f'시작예정 : {scheduledStart[0]}\n종료예정 : {scheduledEnd[0]}\n패치내용 : {content[0]}') 
            else:
                tkinter.messagebox.showwarning(title='Not Response 200', message='팅패정보를 불러오지 못하였습니다.')
        except:
            tkinter.messagebox.showerror(title='Error', message='팅패정보를 불러오지 못하였습니다.')
        return
      
timer=overLay()