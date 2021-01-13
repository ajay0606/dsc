import tkinter as tk
from tkinter import ttk
import os
import re
from datetime import datetime
import pyautogui, sys
import subprocess
import time 


pyautogui.PAUSE = 2.5
pyautogui.FAILSAFE = True


class LabelInput(tk.Frame):
   

    def __init__(self, parent, label='', input_class=ttk.Entry,
         input_var=None, input_args=None, label_args=None,
         **kwargs):
        super().__init__(parent, **kwargs)
        input_args = input_args or {}
        label_args = label_args or {}
        self.variable = input_var
        if input_class in (ttk.Checkbutton, ttk.Button, 
        ttk.Radiobutton):
            input_args['text'] = label
            input_args['variable'] = input_var
        else:
            self.label = ttk.Label(self, text=label, **label_args)
            self.label.grid(row=0, column=0, sticky=(tk.W + tk.E))
            input_args['textvariable'] = input_var
        self.input = input_class(self, **input_args)
        self.input.grid(row=1, column=0, sticky=(tk.W + tk.E))
        self.columnconfigure(0, weight=1)

        def grid(self, sticky=(tk.E + tk.W), **kwargs):
            super().grid(sticky=sticky, **kwargs)

    def get(self):
        try:
            if self.variable:
                return self.variable.get()
            elif type(self.input) == tk.Text:
                return self.input.get('1.0', tk.END)
            else:
                return self.input.get()
        except (TypeError, tk.TclError):
            return ''
        
    def set(self, value, *args, **kwargs):
        if type(self.variable) == tk.BooleanVar:
            self.variable.set(bool(value))
        elif self.variable:
            self.variable.set(value, *args, **kwargs)
        elif type(self.input) in (ttk.Checkbutton, ttk.Radiobutton):
            if value:
                self.input.select()
            else:
                self.input.deselect()
        elif type(self.input) == tk.Text:
            self.input.delete('1.0', tk.END)
            self.input.insert('1.0', value)
        else: # input must be an Entry-type widget with no variable
            self.input.delete(0, tk.END)
            self.input.insert(0, value)

class DataRecordForm(tk.Frame):

    '''The input form for our widgets'''
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        #self.reset()
        recordinfo = tk.LabelFrame(self, text='Data')
        self.Path = LabelInput(recordinfo, 'Path', input_var=tk.StringVar(parent, value="D:\pdf"))
        self.Path.grid(row=0, column=0, ipadx=100)
        self.Speed =  LabelInput(recordinfo, 'Speed', input_var=tk.DoubleVar())
        self.Speed.grid(row=3, column=0, padx=10, pady=10)        
        self.Pcount = LabelInput(recordinfo, 'Count', )
        self.Pcount.grid(row=3, column=1, padx=10, pady=10)
        recordinfo.grid(row=0, column=0, sticky=tk.W + tk.E)
        


class Application(tk.Tk):
    ''' main functions of widget app'''
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title('UKL Digital Signature Certificate')
        self.geometry('495x250')
        ttk.Label(self,
        text='Uni Klinger Ltd',
        font=('TkDefaultFont', 16)).grid(row=0)
        self.recordform = DataRecordForm(self)
        self.recordform.grid(row=1, padx=10)
        self.userpin = tk.StringVar()
        self.userlabel = tk.Label(self, text='User Pin',borderwidth = 3, font = ('calibre',10,'bold'))
        self.userlabel.grid(row=2, column=0)
        self.userPin = tk.Entry(self, textvariable=self.userpin, font = ('calibre',10,'normal'), show='*')
        self.userPin.grid(row=3, column=0,)
        self.savebutton = ttk.Button(self, text='Start', command=self.start)
        self.savebutton.grid(sticky=tk.E, row=3, padx=10)
        self.status = tk.StringVar()
        self.statusbar = ttk.Label(self, textvariable=self.status)
        self.statusbar.grid(sticky=(tk.W + tk.E), row=4, padx=10)

    def start(self): 
        global clean
        countdown = 0
        clean =  self.recordform.Path.get()
        clean = re.sub(r'\\', '/', clean)
        #di = os.startfile(clean)
        di = os.chdir(clean)
        #delay = self.recordform.Speed.get()
        delay = 0.25
        userPin = self.userpin.get()
        self.status.set("Process started, Please do not Interrupt !...")
        
        for count, pdf in enumerate(os.listdir(di)):
            pdfe, pext = os.path.splitext(pdf) 
            fname = pdfe
            name = fname[-3:]
            
            
            if name == "-DG" in fname:
                count+=1
            else:
                adob = "C:/Program Files (x86)/Adobe/Acrobat Reader DC/Reader/AcroRd32.exe"
                subprocess.Popen([adob, pdf], shell=True)
                #script
                time.sleep(4)
                ob = pyautogui.moveTo(x=100, y=60, duration=delay)
                pyautogui.click(ob)
                ob = pyautogui.moveTo(x=497, y=436, duration=delay)
                pyautogui.click(ob)
                ob = pyautogui.moveTo(564, 217)
                pyautogui.click(ob)
                pyautogui.scroll(-1000000)
                #Script1
                #signture
                ob = pyautogui.moveTo(x=532, y=147, duration=delay)
                pyautogui.click(ob)
                #signature
                pyautogui.moveTo(1017, 479)
                pyautogui.mouseDown(1017, 479, button='left', duration=delay)
                pyautogui.dragTo(1261, 563, duration=delay, button='left')
                pyautogui.mouseUp(1261, 563, button='left', duration=delay)
                #time.sleep(8)
                # auto object click
                # continue
                img = 'C:/Users/Public/Pictures/req/firstop.PNG'
                while pyautogui.locateCenterOnScreen(img) == None:
                    fir = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                    print('image not found')
                fir = pyautogui.locateCenterOnScreen(img, confidence=0.9)
                print('image found')
                pyautogui.click(fir)
                # continue
                # lock
                imge = 'C:/Users/Public/Pictures/req/lock.PNG'
                while pyautogui.locateCenterOnScreen(imge) == None:
                    lock = pyautogui.locateCenterOnScreen(imge, confidence=0.9)
                    print('image not found')
                lock = pyautogui.locateCenterOnScreen(imge, confidence=0.9)
                print('image found')
                pyautogui.click(lock)
                # lock
                #time.sleep(6)
                # sign
                imgg = 'C:/Users/Public/Pictures/req/secondop.PNG'
                while pyautogui.locateCenterOnScreen(imgg) == None:
                    sec = pyautogui.locateCenterOnScreen(imgg, confidence=0.9)
                    print('image not found')
                sec = pyautogui.locateCenterOnScreen(imgg, confidence=0.9)
                print('image found')
                pyautogui.click(sec)
                # sign
                #time.sleep(20)
                # save As
                imggg = 'C:/Users/Public/Pictures/req/thirdop.PNG'
                while pyautogui.locateCenterOnScreen(imggg) == None:
                    thi = pyautogui.locateCenterOnScreen(imggg, confidence=0.9)
                    print('image not found')
                thi =  pyautogui.locateCenterOnScreen(imggg, confidence=0.9)
                print('image found')
                pyautogui.click(thi)
                # save As
                #time.sleep(6)
                # Yes
                imgi = 'C:/Users/Public/Pictures/req/forthop.PNG'
                while pyautogui.locateCenterOnScreen(imgi) == None:
                    forth = pyautogui.locateCenterOnScreen(imgi, confidence=0.9)
                    print('image not found')
                forth =  pyautogui.locateCenterOnScreen(imgi, confidence=0.9)
                print('image found')
                pyautogui.click(forth)
                # Yes
                #time.sleep(6)
                # User Pin
                imgii = 'C:/Users/Public/Pictures/req/fifth.PNG'
                ptime = 0
                while pyautogui.locateCenterOnScreen(imgii) == None:
                    ptime+=1
                    fifth = pyautogui.locateCenterOnScreen(imgii, confidence=0.9)
                    print('image not found')
                    if ptime == 2:
                        break
            
                fifth = pyautogui.locateCenterOnScreen(imgii, confidence=0.9)
                print('image found')
                #pyautogui.click(fifth)
                # User Pin
                # actual Pin
                print(userPin)
                pyautogui.write(userPin)
                # actual Pin
                # Ok
                ptime = 0
                imgiii = 'C:/Users/Public/Pictures/req/sixth.PNG'
                while pyautogui.locateCenterOnScreen(imgiii) == None:
                    ptime+=1
                    sixth = pyautogui.locateCenterOnScreen(imgiii, confidence=0.9)
                    print('image not found')
                    if ptime == 2:
                        break
                sixth =  pyautogui.locateCenterOnScreen(imgiii, confidence=0.9)
                print('image found')
                pyautogui.click(sixth)
                # Ok

                # manual object click
                #ob = pyautogui.moveTo(961, 599, duration=delay)
                #pyautogui.click(ob)
                # manual object click
                #lock signature
                #ob = pyautogui.moveTo(362, 483, duration=delay)
                #pyautogui.click(ob)
                #lock signature
                #ob = pyautogui.moveTo(969, 608, duration=delay)
                #pyautogui.click(ob)
                # saving file
                #ob = pyautogui.moveTo(514, 487, duration=delay)
                #pyautogui.click(ob)
                #time.sleep(2)
                #ob = pyautogui.moveTo(727, 371, duration=delay)
                #pyautogui.click(ob)
                # pin
                #ob = pyautogui.moveTo(628, 301, duration=delay)
                #pyautogui.click(ob)
                #pyautogui.write('Priyal@3103')
                #ob = pyautogui.moveTo(580, 369, duration=delay)
                #pyautogui.click(ob)
                #pin

                ob = pyautogui.moveTo(1338, 3, duration=delay)
                pyautogui.click(ob)
                pyautogui.moveTo(218, 127, duration=delay)
                pyautogui.move(0, 25, duration=delay)
                rename = "{}-DG{}".format(pdfe, pext)
                os.rename(pdf, rename)
                countdown+=1
                cop = self.recordform.Pcount.set(countdown)
                self.status.set("Process Completed..........")
                

               

if __name__ == "__main__":
    app = Application()
    app.mainloop()
