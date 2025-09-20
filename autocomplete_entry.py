import tkinter as tk

class AutocompleteEntry(tk.Entry):
    def __init__(self, master, options, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.options = options
        self.var = tk.StringVar()
        self.config(textvariable=self.var)
        self.var.trace("w", self.changed)
        self.bind("<Down>", self.move_down)
        self.bind("<Up>", self.move_up)
        self.bind("<Return>", self.select_current)
        self.listbox_up = False

    def changed(self, *args):
        typed = self.var.get()
        if typed == "":
            self.hide_listbox()
            return
        matches = [o for o in self.options if typed.lower() in o.lower()]
        if matches:
            if not self.listbox_up:
                self.show_listbox(self.row)
            self.listbox.delete(0, tk.END)
            for option in matches:
                self.listbox.insert(tk.END, option)
        else:
            self.hide_listbox()

    def show_listbox(self, row):
        if not self.listbox_up:
            max_show = 4
            self.listbox = tk.Listbox(self.master, height=max_show)
            self.listbox.bind("<<ListboxSelect>>", self.on_click)
            self.focus_set()
            x = self.winfo_x()
            entry_height = self.winfo_height()
            if row >= 4:
                y = self.winfo_y() - entry_height * (max_show+1)
            else:
                y = self.winfo_y() + entry_height
            self.listbox.place(x=x, y=y, width=self.winfo_width())
            self.listbox_up = True

    def hide_listbox(self):
        if self.listbox_up:
            self.listbox.destroy()
            self.listbox_up = False

    def on_click(self, event):
        self.select_current()

    def select_current(self, event=None):
        if not self.winfo_exists():
            return
        if self.listbox_up and self.listbox.curselection():
            index = self.listbox.curselection()[0]
            value = self.listbox.get(index)
            self.var.set(value)
            self.icursor(tk.END)
        else:
            if self.var.get() not in self.options:
                self.var.set('')
                
        self.icursor(tk.END)
        self.hide_listbox()
        
        return "break"

    def move_down(self, event):
        if self.listbox_up:
            current = self.listbox.curselection()
            if not current:
                self.listbox.selection_set(0)
            else:
                index = current[0]
                if index < self.listbox.size() - 1:
                    self.listbox.selection_clear(index)
                    self.listbox.selection_set(index + 1)
            return "break"

    def move_up(self, event):
        if self.listbox_up:
            current = self.listbox.curselection()
            if not current:
                self.listbox.selection_set(self.listbox.size() - 1)
            else:
                index = current[0]
                if index > 0:
                    self.listbox.selection_clear(index)
                    self.listbox.selection_set(index - 1)
            return "break"
