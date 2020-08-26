try:
    import Tkinter as tk
    from Tkinter import *
except ImportError:
    import tkinter as tk
    from tkinter import *

class DateEntry(tk.Frame):
    def __init__(self, master, frame_look={}, **look):
        try:
            args = dict(relief=tk.SUNKEN, border=1)
            args.update(frame_look)
            tk.Frame.__init__(self, master, **args)

            args = {'relief': tk.FLAT}
            args.update(look)

            self.entry_1 = tk.Entry(self, width=2, **args)
            self.label_1 = tk.Label(self, text='/', **args)
            self.entry_2 = tk.Entry(self, width=2, **args)
            self.label_2 = tk.Label(self, text='/', **args)
            self.entry_3 = tk.Entry(self, width=4, **args)

            self.entry_1.pack(side=tk.LEFT)
            self.label_1.pack(side=tk.LEFT)
            self.entry_2.pack(side=tk.LEFT)
            self.label_2.pack(side=tk.LEFT)
            self.entry_3.pack(side=tk.LEFT)

            self.entry_1.bind('<KeyRelease>', self._e1_check)
            self.entry_2.bind('<KeyRelease>', self._e2_check)
            self.entry_3.bind('<KeyRelease>', self._e3_check)
        except:
            pass

    def _backspace(self, entry):
        try:
            cont = entry.get()
            entry.delete(0, tk.END)
            entry.insert(0, cont[:-1])
        except:
            pass

    def _e1_check(self, e):
        try:
            cont = self.entry_1.get()
            if len(cont) >= 2:
                self.entry_2.focus()
            if len(cont) > 2 or not cont[-1].isdigit():
                self._backspace(self.entry_1)
                self.entry_1.focus()
        except:
            pass

    def _e2_check(self, e):
        try:
            cont = self.entry_2.get()
            if len(cont) >= 2:
                self.entry_3.focus()
            if len(cont) > 2 or not cont[-1].isdigit():
                self._backspace(self.entry_2)
                self.entry_2.focus()
        except:
            pass

    def _e3_check(self, e):
        try:
            cont = self.entry_3.get()
            if len(cont) > 4 or not cont[-1].isdigit():
                self._backspace(self.entry_3)
        except:
            pass

    def get(self):
        return self.entry_2.get() + self.entry_1.get() + self.entry_3.get() 

