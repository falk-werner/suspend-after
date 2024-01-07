#!/usr/bin/env python3

import os
from tkinter import *
from tkinter import ttk
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

TIMEOUT = 60

class Main:
    def __init__(self):
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = interface.QueryInterface(IAudioEndpointVolume)
        self.initial_volume = self.volume.GetMasterVolumeLevel() 
        self.min_volume = self.volume.GetVolumeRange()[0]
        self.remaining = 60 * TIMEOUT
        self.root = Tk()
        self.root.title("SuspendAfter")
        self.root.iconbitmap("timer.ico")
        frame = ttk.Frame(self.root, padding=10)
        frame.grid()

        row = 0
        ttk.Label(frame, text="Remaining:").grid(column=0,row=row,padx=5,sticky=E)
        self.remainingText = StringVar()
        self.remainingText.set(f"{int(self.remaining / 60)}m {self.remaining % 60}s")
        ttk.Label(frame, textvariable=self.remainingText).grid(column=1,row=row, sticky=E,padx=5)
        ttk.Button(frame, text="Reset", command=self.reset).grid(column=2,row=row,padx=5)

        row += 1
        ttk.Label(frame, text="Timeout:").grid(column=0,row=row,padx=5,sticky=E)
        self.timeout = StringVar()
        self.timeout.set("60")
        ttk.Entry(frame, textvariable=self.timeout, justify=RIGHT).grid(column=1,row=row,padx=5)
        ttk.Label(frame, text="min").grid(column=2,row=row,sticky=W,padx=5)

        row += 1
        ttk.Label(frame, text="Volume:").grid(column=0,row=row,padx=5,sticky=E)
        self.volumeText = StringVar()
        self.volumeText.set(f"{self.volume.GetMasterVolumeLevel():.02f}")
        ttk.Label(frame, textvariable=self.volumeText).grid(column=1,row=row,sticky=E,padx=5)
        ttk.Label(frame, text="dB").grid(column=2,row=row,sticky=W,padx=5)
    
    def reset(self):
        self.remaining = 60 * int(self.timeout.get())
        self.remainingText.set(f"{int(self.remaining / 60)}m {self.remaining % 60}s")
        self.volume.SetMasterVolumeLevel(self.initial_volume, None)
        self.volumeText.set(f"{self.volume.GetMasterVolumeLevel():.02f}")

    def decrease_volume(self):
        new_volume = max(self.volume.GetMasterVolumeLevel() - 1, self.min_volume)
        self.volume.SetMasterVolumeLevel(new_volume, None)
        self.volumeText.set(f"{self.volume.GetMasterVolumeLevel():.02f}")

    def on_tick(self):
        self.remaining -= 1
        self.remainingText.set(f"{int(self.remaining / 60)}m {self.remaining % 60}s")

        if self.remaining == 0:
            os.system("shutdown /h")
            self.root.destroy()
            return

        if self.remaining < 120 and (self.remaining % 5) == 0:
            self.decrease_volume()
        self.root.after(1000, self.on_tick)

    def Loop(self):
        self.root.after(1000, self.on_tick)
        self.root.mainloop()

if __name__ == "__main__":
    main = Main()
    main.Loop()



