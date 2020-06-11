import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import shutil
import csv
import xlrd
import itertools


class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        """Use our completion list as our drop down selection menu, arrows move through menu."""
        self._completion_list = sorted(completion_list, key=str.lower)  # Work with a sorted list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self._completion_list  # Setup our popup menu

    def autocomplete(self, delta=0):
        """autocomplete the Combobox, delta may be 0/1/-1 to cycle through possible hits"""
        if delta:  # need to delete selection otherwise we would fix the current position
            self.delete(self.position, tk.END)
        else:  # set position to end so selection starts where textentry ended
            self.position = len(self.get())
        # collect hits
        _hits = []
        for element in self._completion_list:
            if element.lower().startswith(self.get().lower()):  # Match case insensitively
                _hits.append(element)
        # if we have a new hit list, keep this in mind
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        # only allow cycling if we are in a known hit list
        if _hits == self._hits and self._hits:
            self._hit_index = (self._hit_index + delta) % len(self._hits)
        # now finally perform the auto completion
        if self._hits:
            self.delete(0, tk.END)
            self.insert(0, self._hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        """event handler for the keyrelease event on this widget"""
        if event.keysym == "BackSpace":
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        if event.keysym == "Left":
            if self.position < self.index(tk.END):  # delete the selection
                self.delete(self.position, tk.END)
            else:
                self.position = self.position - 1  # delete one character
                self.delete(self.position, tk.END)
        if event.keysym == "Right":
            self.position = self.index(tk.END)  # go to end (no selection)
        if len(event.keysym) == 1:
            self.autocomplete()
        # No need for up/down, we'll jump to the popup
        # list at the position of the autocompletion


class MainApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")
        self.state("zoomed")
        self.title("Price List")

        container = tk.Frame(self, borderwidth=0, highlightthickness=0)
        container.pack(side="top", fill="both")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for page in [MainPage, PriceList, Analyzer]:
            page_name = page.__name__
            frame = page(master=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky=tk.NSEW)

        self.show_frame("MainPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()


class MainPage(tk.Frame):
    def __init__(self, master, controller):
        tk.Frame.__init__(self, master)
        self.controller = controller
        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(self, text="Database", font="Arial 28 bold")
        price_list_button = tk.Button(self, text="Price List", font="Arial 14", width=30,
                                      command=lambda: self.controller.show_frame("PriceList"))
        analyze_button = tk.Button(self, text="Analyze", font="Arial 14", width=30,
                                   command=lambda: self.controller.show_frame("Analyzer"))
        title_label.pack(pady=(80, 10))
        price_list_button.pack(pady=10)
        analyze_button.pack(pady=10)


class PriceList(tk.Frame):
    def __init__(self, master, controller):
        tk.Frame.__init__(self, master)
        self.controller = controller
        self.width, self.height = self.winfo_screenwidth(), self.winfo_screenheight()
        self.configure(width=self.width, height=self.height)
        self.propagate()
        self.tables = set()
        self.table_frames_dict = dict()
        self.table_size = 14
        self.scale = 1

        self.back_button = tk.Button(self, text="<", font="Times 8 bold",
                                     command=lambda: self.controller.show_frame("MainPage"))
        self.files = sorted(self.price_list_options())
        self.price_list_var = tk.StringVar()
        self.price_list_label = tk.Label(self, text="Price List of: ")
        self.price_list_file_option = AutocompleteCombobox(self, textvariable=self.price_list_var, values=self.files,
                                                           width=50)
        self.price_list_file_option.set_completion_list(self.files)
        self.price_list_file_option.bind("<Return>", lambda event: self.create_tables())
        self.create_func =  self.price_list_file_option.bind("<<ComboboxSelected>>", lambda event: self.create_tables())
        self.price_list_search_label = tk.Label(self, text="Search:")
        self.search_only_var = tk.IntVar()
        self.search_only_box = tk.Checkbutton(self, text="Search Only", variable = self.search_only_var, command= self.search_only)
        self.price_list_search_bar = tk.Entry(self, width=30)
        self.price_list_search_bar.bind("<Return>", func=lambda event: self.search())
        self.big_frame = tk.Frame(self, bg="blue")
        self.price_list_canvas = tk.Canvas(self.big_frame, bg="white")
        self.price_list_canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        self.y_scrollbar = tk.Scrollbar(self.big_frame, orient="vertical", command=self.price_list_canvas.yview)
        self.x_scrollbar = tk.Scrollbar(self.big_frame, orient="horizontal", command=self.price_list_canvas.xview)
        self.price_list_frame = tk.Frame(self.price_list_canvas)
        self.price_list_frame.bind("<Configure>", self.set_scrollregion)
        self.zoom_in_button = tk.Button(self, text="Zoom In", command=self.zoom_in)
        self.zoom_out_button = tk.Button(self, text="Zoom Out", command=self.zoom_out)
        self.add_button = tk.Button(self, text="Add a new price list file", command=self.add_file)

        self.back_button.place(x=0, y=0)
        self.price_list_label.place(x=30, y=30)
        self.price_list_file_option.place(x=120, y=30)
        self.search_only_box.place(x = self.width - 370, y = 30)
        self.price_list_search_label.place(x=self.width - 250, y=30)
        self.price_list_search_bar.place(x=self.width - 205, y=30)
        self.big_frame.place(width=self.width - 50, height=self.height - 140, x=30, y=60)
        self.x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.price_list_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=tk.TRUE)
        self.price_list_canvas.create_window(0, 0, anchor=tk.NW, window=self.price_list_frame)
        self.price_list_canvas.configure(xscrollcommand=self.x_scrollbar.set)
        self.price_list_canvas.configure(yscrollcommand=self.y_scrollbar.set)
        self.zoom_in_button.place(x=30, y=700)
        self.zoom_out_button.place(x=90, y=700)
        self.add_button.place(x=self.width - 150, y=700)

    def on_mousewheel(self, event):
        self.price_list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def set_scrollregion(self, event):
        self.price_list_canvas.configure(scrollregion=self.price_list_canvas.bbox("all"))

    def price_list_options(self):
        with open("price_list_file.txt", "r") as file:
            files = file.read().split("\n")
        return files

    def create_tables(self):
        for widget in self.price_list_frame.winfo_children():
            widget.destroy()

        file = self.price_list_var.get()

        if file not in self.table_frames_dict:
            self.tables = set()
            with open(f"{file}.csv", "r") as old_file:
                sniffer = csv.Sniffer().sniff(old_file.read(1024))
                dialect = sniffer.delimiter
                old_file.seek(0)
                csv_reader = csv.reader(old_file, delimiter=dialect)
                data = list(itertools.zip_longest(*csv_reader))
                for i, col in enumerate(data):
                    a = [word for word in col if word]
                    average = sum(len(word) for word in a) / len(a)
                    for j, field in enumerate(col):
                        if field:
                            label = tk.Label(self.price_list_frame, text=field, bg="white", relief="groove",
                                             font=f"Arial {self.table_size}", anchor=tk.W)
                            if len(field) >= average + 8:
                                label.grid(row=j, column=i, sticky=tk.NSEW, columnspan=4)
                            else:
                                label.grid(row=j, column=i, sticky=tk.NSEW)
                            self.tables.add(label)

    def search_only(self):
        state = self.search_only_var.get()
        if state == 0:
            self.create_func = self.price_list_file_option.bind("<<ComboboxSelected>>",
                                                                lambda event: self.create_tables())
        else:
            self.price_list_file_option.unbind("<<ComboboxSelected>>", self.create_func)

    def search(self):
        self.price_list_frame.destroy()
        self.price_list_frame = tk.Frame(self.price_list_canvas)
        self.price_list_canvas.create_window(0, 0, anchor=tk.NW, window=self.price_list_frame)
        self.tables = set()
        file = self.price_list_var.get()
        searched = self.price_list_search_bar.get().lower()
        if not searched:
            self.create_tables()
            return
        with open(f"{file}.csv", "r") as old_file:
            sniffer = csv.Sniffer().sniff(old_file.read(1024))
            dialect = sniffer.delimiter
            old_file.seek(0)
            csv_reader = csv.reader(old_file, delimiter=dialect)
            for i, row in enumerate(csv_reader):
                line = " ".join(row).lower()
                if searched in line:
                    for j, field in enumerate(row):
                        label = tk.Label(self.price_list_frame, text=field, bg="white", relief="groove",
                                         font=f"Arial {self.table_size}", anchor=tk.W)
                        label.grid(row=i, column=j, sticky=tk.NSEW)
                        self.tables.add(label)
                else:
                    continue

    def zoom_in(self):
        try:
            if self.table_size <= 26:
                self.table_size += 2
                self.zoom_out_button.configure(state="normal")
                for label in self.tables:
                    label.configure(font=f"Arial {self.table_size}")
            else:
                self.zoom_in_button.configure(state="disabled")
        except:
            return

    def zoom_out(self):
        try:
            if self.table_size >= 8:
                self.table_size -= 2
                self.zoom_in_button.configure(state="normal")
                for label in self.tables:
                    label.configure(font=f"Arial {self.table_size}")
            else:
                self.zoom_out_button.configure(state="disabled")
        except:
            return

    def add_file(self):
        file = filedialog.askopenfilename(title="Select a file",
                                          filetypes=(
                                          ("CSV Files", "*.csv"), ("Excel Files", ".xlsx .xls"), ("All Files", "*.*")))
        if not file:
            return
        base = os.path.basename(file)
        file_name, file_extension = os.path.splitext(base)

        try:
            with open("price_list_file.txt", "r+") as f:
                current_list = f.read().split("\n")

                if file_extension == ".xlsx" or file_extension == ".xls":
                    wb = xlrd.open_workbook(file)
                    shs = wb.sheet_names()
                    for sh_name in shs:
                        new_file_name = f"{sh_name}_{file_name}"
                        if new_file_name in current_list:
                            response = messagebox.askyesno(title=f"{new_file_name} Already Exists",
                                                           message="Overwrite the old file?")
                            if response == 0:
                                continue
                            else:
                                with open(f"{new_file_name}.csv", "w") as new_csv_file:
                                    sh = wb.sheet_by_name(sh_name)
                                    wr = csv.writer(new_csv_file, quoting=csv.QUOTE_ALL)
                                    for row in range(sh.nrows):
                                        wr.writerow(sh.row_values(row))
                        else:
                            with open(f"{new_file_name}.csv", "w") as new_csv_file:
                                sh = wb.sheet_by_name(sh_name)
                                wr = csv.writer(new_csv_file, quoting=csv.QUOTE_ALL)
                                for row in range(sh.nrows):
                                    wr.writerow(sh.row_values(row))
                                f.write(f"\n{new_file_name}")

                elif file_extension == ".csv":
                    if file_name in current_list:
                        response = messagebox.askyesno(title=f"{file_name} Already Exists",
                                                       message="Overwrite the old file?")
                        if response == 0:
                            raise AssertionError
                        else:
                            dir_path = os.path.dirname(os.path.realpath(__file__))
                            shutil.copy(file, dir_path)
                    else:
                        dir_path = os.path.dirname(os.path.realpath(__file__))
                        shutil.copy(file, dir_path)
                        f.write(f"\n{file_name}")
        except AssertionError:
            pass
        self.files = self.price_list_options()
        self.price_list_file_option.configure(values=self.files)


class Analyzer(tk.Frame):
    def __init__(self, master, controller):
        tk.Frame.__init__(self, master)
        self.controller = controller
        self.configure(bg="blue")


def main():
    root = MainApp()
    root.mainloop()


if __name__ == '__main__':
    main()
