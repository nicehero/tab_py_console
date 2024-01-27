import asyncio
from codecs import StreamReader
import tkinter as tk
from tkinter import ttk
import subprocess
from tkinter.scrolledtext import ScrolledText
import os
from async_tkinter_loop import async_handler, async_mainloop,get_event_loop,main_loop
import min2tray
import win32gui, win32api, win32con
import threading
import pystray
cmds = [
"python -u test.py",
]
#cmds = []
title = "标签x控制台xPython"
loop = get_event_loop()
consoles = []
root = tk.Tk()
root.title(title)
max_line_count = 256
tab_control = ttk.Notebook(root)

def get_window_handle():
    return win32gui.FindWindow("TkTopLevel",title)
def to_tray():
    window_handle = get_window_handle()
    min2tray.hide_window(window_handle=window_handle)
icon = None
def iconRun(icon_image_path):
    icon_image = None
    icon_image = min2tray.create_default_image(64, 64, 'black', 'white')
    def show_window():
        window_handle = get_window_handle()
        min2tray.show_window(window_handle=window_handle)
        win32gui.ShowWindow(window_handle, win32con.SW_SHOWNORMAL)
    async def stop_root_async():
        root.destroy()
    async def to_tray_async():
        to_tray()
    def stop_all():
        icon.stop()
        asyncio.run_coroutine_threadsafe(stop_root_async(), loop)
    global icon
    icon = pystray.Icon(
        'test name',
        icon=icon_image,
        title=title,
        menu=pystray.Menu(
            pystray.MenuItem(text="打开界面",action=show_window,default=True),
            pystray.MenuItem(text="隐藏界面",action=lambda:asyncio.run_coroutine_threadsafe(to_tray_async(), loop)),
            pystray.MenuItem(text="退出",action=stop_all)
        )
    )
    icon.run()

thread_icon = threading.Thread(target=iconRun, args=(None,))
thread_icon.start()

tab_control.pack(expand=True, fill=tk.BOTH)

def on_minimize_window():
    to_tray()

def add_console(cmd,index,dirname):
    frame = tk.Frame(tab_control)
    frame.pack(expand=True, fill=tk.BOTH)
    text = ScrolledText(frame)
    text.pack(expand=True, fill=tk.BOTH)
    tab_control.add(frame, text=f"{cmd.split(' ')[-1].split('/')[-1]}")
    c1 = ttk.Frame(frame)
    c1.pack(fill='x')
    start_button = tk.Button(c1, text=f"start", command=lambda index=index: start_console_async(index))
    stop_button = tk.Button(c1, text=f"stop", command=lambda index=index: stop_console(index))
    stop_button.configure(state=tk.DISABLED)
    start_button.pack(fill='x')
    stop_button.pack(fill='x')
    proc = None
    consoles.append([proc,text,start_button,stop_button,frame])

def stop_console(index,isend=False):
    if consoles[index] != None:
        (proc,text,start_button,stop_button,frame) = consoles[index]
        if proc != None:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.check_call(['taskkill', '/PID', str(proc.pid)], startupinfo=si)
            consoles[index][0] = None
            if isend == False:
                stop_button.configure(state=tk.DISABLED)
                start_button.configure(state=tk.NORMAL)

def is_ascii_byte(byte):
    if len(byte) < 1:
        return True
    if 0x00 <= byte[0] <= 0x7F:
        return True
    else:
        return False

@async_handler
async def start_console_async(index):
    (proc,text,start_button,stop_button,frame) = consoles[index]
    start_button.configure(state=tk.DISABLED)
    stop_button.configure(state=tk.NORMAL)
    text.delete('1.0', tk.END)
    cmd = cmds[index]
    dirname = None
    if isinstance(cmd,list):
        if len(cmd) > 1:
            dirname = cmd[1]
        cmd = cmd[0]
    proc = await asyncio.create_subprocess_shell(
        cmd,
        creationflags=subprocess.CREATE_NEW_CONSOLE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
        cwd=dirname
        )
    consoles[index][0] = proc
    print(proc.pid)
    async def readpipe(pipe:asyncio.StreamReader,isStdErr):
        buffer = b''
        lastIsR = False
        while True:
            try:
                out = await pipe.read(1)
            except:
                break
            if out == b'\r':
                lastIsR = True
                continue
            if out == b'\n':
                lastIsR = False
            buffer = buffer + out
            if not is_ascii_byte(out):
                continue
            try:
                buffer = buffer.decode('utf-8')
            except:
                try:
                    buffer = buffer.decode('GBK')
                except:
                    pass
            console_output(index, buffer, 1,lastIsR)
            lastIsR = False
            buffer = b''
            if out == b'':
                break
        if isStdErr:
            console_output(index, '程序结束!', 0,False)
    asyncio.run_coroutine_threadsafe(readpipe(proc.stdout,False), loop)
    asyncio.run_coroutine_threadsafe(readpipe(proc.stderr,True), loop)

def console_output(index, new_text, ret,firstIsR):
    add_text(index,new_text,firstIsR)
    if ret == 0:
        (proc,text,start_button,stop_button,frame) = consoles[index]
        stop_button.configure(state=tk.DISABLED)
        start_button.configure(state=tk.NORMAL)
    
def add_text(index,new_text,firstIsR:bool):
    (proc,text,start_button,stop_button,frame) = consoles[index]
    current_text = text.get('1.0', tk.END)
    last_index = text.index(tk.END)
    line_count = int(last_index.split(".")[0])
    if line_count >= max_line_count and not firstIsR:
        text.delete("1.0",  "2.0")
    if firstIsR == True:
        last_line_index = text.index("end-1c linestart")
        t = text.get(last_line_index, "end")
        if '\n' != t:
            text.delete(last_line_index, "end")
            if last_line_index != '1.0':
                text.insert(tk.END, '\n')
    text.insert(tk.END, new_text)
    text.see(tk.END)

for index in range(len(cmds)):
    add_console(cmds[index],index,None)

def wndProc(hwnd, msg, wParam, lParam):
    if msg == win32con.WM_DROPFILES:
        hdrop = wParam
        num_files = win32api.DragQueryFile(hdrop, -1)
        for i in range(num_files):
            file_path = win32api.DragQueryFile(hdrop, i)
            # 获取文件名
            filename = os.path.basename(file_path)
            dirname = os.path.dirname(file_path)
            extension = os.path.splitext(file_path)[1]
            file_path = file_path.replace('\\','/')
            dirname = dirname.replace('\\','/')
            if extension == '.py':
                cmd = f"python -u {file_path}"
                async def add_console_async(cmd,index,dirname):
                    cmds.append([cmd,dirname])
                    add_console(cmd,index,dirname)
                asyncio.run_coroutine_threadsafe(add_console_async(cmd,len(cmds),dirname),loop)
            else:
                print(f'Dropped file: {file_path}')
        win32api.DragFinish(hdrop)
    elif msg == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        root.destroy()
    elif msg == win32con.WM_SIZE:
        if wParam == win32con.SIZE_MINIMIZED:
            to_tray()
    else:
        return win32gui.DefWindowProc(hwnd, msg, wParam, lParam)
    return 0

def start_drag_event():
    print(root.winfo_id())
    window_handle = win32gui.GetParent(root.winfo_id())
    #window_handle = get_window_handle()
    win32gui.SetWindowLong(window_handle, win32con.GWL_WNDPROC, wndProc)
    win32gui.DragAcceptFiles(window_handle, True)

if __name__ == "__main__":
    root.after(100,start_drag_event)
    #go main
    loop.run_until_complete(main_loop(root))
    #clear subprocesses
    if icon:
        icon.stop()
    for i in range(len(consoles)):
        stop_console(i,True)

