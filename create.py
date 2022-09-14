import string
import psutil
import gspread
import os
import datetime
import traceback


class Memory:
    gc = gspread.service_account(filename=os.path.abspath("credentials.json"))
    sh = gc.open_by_key("1eHOQcO2Wr6GDohcv9BuogsJequ0M6v1M9eaUI0ys59c")
    worksheet = sh.sheet1

    def time_now(self):
        current_datetime = datetime.datetime.now()
        return str(current_datetime)

    def getmemory(self):
        disk_list = []
        disks = ''
        for c in string.ascii_uppercase:
            disk = c + ':'
            if os.path.isdir(disk):
                disk_list.append(disk)

        for i in disk_list:
            free = psutil.disk_usage(i).free / (1024 * 1024 * 1024)
            a = f"{free:.4}"
            disks += a + "\n"
        return disks

    def create_client(self, user, memory, time):
        client = [user, memory, time]
        Memory.worksheet.append_row(client)

    def get_and_write_index(self, path, user):
        f = open(path, "w")
        values_list = Memory.worksheet.col_values(1)
        f.write(str(len(values_list)))
        f.close()

try:
    cr = Memory()
    while True:
        cl = input("Введите название клиента: ")
        if Memory.worksheet.find(cl) == None:
            cr.create_client(cl, cr.getmemory(), cr.time_now())
            cr.get_and_write_index("C:\index.txt", cl)
            cr.get_and_write_index("index.txt", cl)
            print(f"Клиент {cl} успешно добавлен!")

            break
        else:
            print("Клиент уже существует. Попробуйте еще раз!")
except:
    f = open("log.log", "w")
    f.write(traceback.format_exc())

