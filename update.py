import string
import psutil
import gspread
import os
import datetime
import traceback
from tkinter import *


class Memory:
    gc = gspread.service_account(filename=os.path.abspath("credentials.json"))
    sh = gc.open_by_key("1eHOQcO2Wr6GDohcv9BuogsJequ0M6v1M9eaUI0ys59c")
    worksheet = sh.sheet1

    def time_now(self):
        current_datetime = datetime.datetime.now()
        return str(current_datetime)

    def getmemory(self):
        disk_list = []
        disks = []

        for c in string.ascii_uppercase:
            disk = c + ':'
            if os.path.isdir(disk):
                disk_list.append(disk)

        for i in disk_list:
            free = psutil.disk_usage(i).free / (1024 * 1024 * 1024)
            disks.append(str(round(free)))
        return disks

    def update_data(self, value, memory):
        Memory.worksheet.update(value, memory)

    def update_time(self, value, time):
        Memory.worksheet.update(value, time)

    def get_and_write_index(self, path):
        f = open(path, "r")
        return f.read()

    def window(self):
        window = Tk()
        window.title("Предупреждение!")
        window.geometry('1730x300')

        lbl = Label(window,
                    text="Заканчивается свободное дисковое пространство на диске. Возможна некорректная работа программы и базы данных.",
                    font=("Arial Bold", 20), fg="red")
        lbl2 = Label(window,
                     text=" Обратитесь к Вашим IT-специалистам или в отдел технической поддержки компании Меридиан. Тел. 89223276585, WhatsApp, Viber!",
                     font=("Arial Bold", 20), fg="red")
        lbl.grid()
        lbl2.grid()
        window.mainloop()


try:
    cr = Memory()
    try:
        kek = ''
        for i in cr.getmemory():
            kek += i + '\n'
        cr.update_data("B"+cr.get_and_write_index("C:\index.txt"), kek)
        cr.update_time("C"+cr.get_and_write_index("C:\index.txt"), cr.time_now())
        for i in cr.getmemory():
            if int(i) < 25:
                cr.window()
                break


    except FileNotFoundError:
        cr.update_data("B" + cr.get_and_write_index("index.txt"), cr.getmemory())
        cr.update_time("C" + cr.get_and_write_index("index.txt"), cr.time_now())
        for i in cr.getmemory():
            print(i)
            if int(i) < 25:
                cr.window()
                break

except Exception:
    f = open("log.log", "w")
    f.write(traceback.format_exc())





