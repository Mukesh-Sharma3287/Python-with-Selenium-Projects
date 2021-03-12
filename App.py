import tkinter as tk
from tkinter import filedialog
from BPTravel import BPTravel
from logger import App_Logger
import os
from tkinter import messagebox


class GUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Web Automation Tool Edge")
        # driverpath = "./chromedriver"
        self.driverpath = "./msedgedriver"
        self.url = "http://bptravel.blueprism.com"
        self.file_object = open("logs.txt", 'a+')
        self.log_writer = App_Logger()
        self.success = False

        # this removes the maximize button
        self.window.resizable(0, 0)
        window_height = 300
        window_width = 800

        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))

        self.window.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        #self.window.geometry('880x600')
        self.window.configure(background='#ffffff')

        #self.window.attributes('-fullscreen', True)
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)

        header = tk.Label(self.window, text="R&D Microsoft Edge Automation with Python", width=60, height=1, fg="white", bg="#363e75",
                          font=('times', 18, 'bold'))
        header.place(x=0, y=0)
        self.button()

        self.filetxt = tk.Entry(self.window, width=50, textvariable="Select the file path", bg="white", fg="black",
                                font=('times', 15))
        self.start_button()
        self.filetxt.place(x=205, y=80)
        self.window.mainloop()

    def button(self):
        self.button = tk.Button(self.window, text="Browse File", command = self.file_dialog, width=15, height=1, fg="white", bg="#363e75", font=('times', 10))
        self.button.place(x=80, y=80)

    def start_button(self):
        self.Start = tk.Button(self.window, text="Start",command = self.create_quotes, fg="white", bg="#363e75", width=10, height=1,
                               activebackground="#118ce1", font=('times', 12, 'bold'))
        self.Start.place(x=320, y=150)

    def file_dialog(self):
        self.FilePath = filedialog.askopenfilenames(filetypes=[('Excel File', 'xlsx .xls .csv')], initialdir='/',title='Please select Input file')
        if len(self.FilePath) > 0:
            self.filetxt.insert(0, self.FilePath[0])
    def create_quotes(self):
        try:
            if not os.path.exists(self.filetxt.get()):
                messagebox.showinfo("File", "File does not exist")
            else:
                self.bt = BPTravel(driver_path=self.driverpath, url=self.url)
                self.bt.login_bp_travel()
                self.bt.create_quotes(exl_path=self.FilePath[0])
                self.success = True
            if self.success:
                messagebox.showinfo("BP Travel", "Done")
        except Exception as ex:
            self.log_writer.log(self.file_object, "Error occured while web processing BP Travel quotes!! Error:: %s" %ex)
            self.file_object.close()
            raise ex

if __name__ == "__main__":
    gui = GUI()






