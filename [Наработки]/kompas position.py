import tkinter as tk
import win32gui

kp = []
i = 0


def callback(hwnd, extra):
    global kp, i
    name = win32gui.GetWindowText(hwnd)
    if name.find('КОМПАС-3D') != -1:
        if i == 1:
            kp = win32gui.GetWindowRect(hwnd)
            i = 0
        else:
            i += 1


root = tk.Tk()
root.overrideredirect(1)
root.lift()
root.attributes('-topmost', True)
root.after_idle(root.attributes, '-topmost', True)
root.resizable(False, False)
root.configure(bg='#F2F2F2', highlightthickness=1,
               highlightbackground='#666666', highlightcolor='#666666')


def update_position():
    try:
        while True:
            win32gui.EnumWindows(callback, None)
            w = 28
            h = 100
            ki = 147
            if kp[0] == -8:
                ki = ki + 5
            root.geometry(f'{w}x{h}+{kp[2]-w-2}+{kp[1]+ki}')
            root.update()
    except KeyboardInterrupt:
        root.destroy()


update_position()
root.mainloop()
