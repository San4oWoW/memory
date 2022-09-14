import smtplib                                              # Импортируем библиотеку по работе с SMTP
import os                                                   # Функции для работы с операционной системой, не зависящие от используемой операционной системы

# Добавляем необходимые подклассы - MIME-типы
import mimetypes                                            # Импорт класса для обработки неизвестных MIME-типов, базирующихся на расширении файла
from email import encoders                                  # Импортируем энкодер
from email.mime.base import MIMEBase                        # Общий тип
from email.mime.text import MIMEText                        # Текст/HTML
from email.mime.image import MIMEImage                      # Изображения
from email.mime.audio import MIMEAudio                      # Аудио
from email.mime.multipart import MIMEMultipart              # Многокомпонентный объект
import gspread

gc = gspread.service_account(filename=os.path.abspath("credentials.json"))
sh = gc.open_by_key("1eHOQcO2Wr6GDohcv9BuogsJequ0M6v1M9eaUI0ys59c")
worksheet = sh.sheet1

def send_email_with_file(addr_to, msg_subj, msg_text, files):
    addr_from = "helpme.meridian@gmail.com"                         # Отправитель
    password  = "xubhdbdubpixbgzr"                                  # Пароль

    msg = MIMEMultipart()                                   # Создаем сообщение
    msg['From']    = addr_from                              # Адресат
    msg['To']      = addr_to                                # Получатель
    msg['Subject'] = msg_subj                               # Тема сообщения

    body = msg_text                                         # Текст сообщения
    msg.attach(MIMEText(body, 'plain'))                     # Добавляем в сообщение текст

    process_attachement(msg, files)

    #======== Этот блок настраивается для каждого почтового провайдера отдельно ===============================================
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    #server.set_debuglevel(True)                            # Включаем режим отладки, если не нужен - можно закомментировать
    server.login(addr_from, password)                       # Получаем доступ
    server.send_message(msg)                                # Отправляем сообщение
    server.quit()                                           # Выходим
    #==========================================================================================================================

def send_email(addr_to, msg_subj, msg_text):
    addr_from = "helpme.meridian@gmail.com"                         # Отправитель
    password  = "xubhdbdubpixbgzr"                                  # Пароль

    msg = MIMEMultipart()                                   # Создаем сообщение
    msg['From']    = addr_from                              # Адресат
    msg['To']      = addr_to                                # Получатель
    msg['Subject'] = msg_subj                               # Тема сообщения

    body = msg_text                                         # Текст сообщения
    msg.attach(MIMEText(body, 'plain'))                     # Добавляем в сообщение текст


    #======== Этот блок настраивается для каждого почтового провайдера отдельно ===============================================
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    #server.set_debuglevel(True)                            # Включаем режим отладки, если не нужен - можно закомментировать
    server.login(addr_from, password)                       # Получаем доступ
    server.send_message(msg)                                # Отправляем сообщение
    server.quit()                                           # Выходим

def process_attachement(msg, files):                        # Функция по обработке списка, добавляемых к сообщению файлов
    for f in files:
        if os.path.isfile(f):                               # Если файл существует
            attach_file(msg,f)                              # Добавляем файл к сообщению
        elif os.path.exists(f):                             # Если путь не файл и существует, значит - папка
            dir = os.listdir(f)                             # Получаем список файлов в папке
            for file in dir:                                # Перебираем все файлы и...
                attach_file(msg,f+"/"+file)                 # ...добавляем каждый файл к сообщению

def attach_file(msg, filepath):                             # Функция по добавлению конкретного файла к сообщению
    filename = os.path.basename(filepath)                   # Получаем только имя файла
    ctype, encoding = mimetypes.guess_type(filepath)        # Определяем тип файла на основе его расширения
    if ctype is None or encoding is not None:               # Если тип файла не определяется
        ctype = 'application/octet-stream'                  # Будем использовать общий тип
    maintype, subtype = ctype.split('/', 1)                 # Получаем тип и подтип
    if maintype == 'text':                                  # Если текстовый файл
        with open(filepath) as fp:                          # Открываем файл для чтения
            file = MIMEText(fp.read(), _subtype=subtype)    # Используем тип MIMEText
            fp.close()                                      # После использования файл обязательно нужно закрыть
    elif maintype == 'image':                               # Если изображение
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'audio':                               # Если аудио
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    else:                                                   # Неизвестный тип файла
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)              # Используем общий MIME-тип
            file.set_payload(fp.read())                     # Добавляем содержимое общего типа (полезную нагрузку)
            fp.close()
            encoders.encode_base64(file)                    # Содержимое должно кодироваться как Base64
    file.add_header('Content-Disposition', 'attachment', filename=filename) # Добавляем заголовки
    msg.attach(file)                                        # Присоединяем файл к сообщению

def sort(d):
    mas1 = list(d)
    mas2 = []
    mas_i = []
    s = ''
    try:
        for i in range(len(mas1)):
            s += mas1[i]
            if mas1[i] == '\n':
                s = s[:-1]
                mas2.append(s)
                mas_i.append(i)
                s = ''

        mas2.append(d[mas_i[len(mas_i) - 1] + 1:])

        mylist = [x for x in mas2 if x]
        return mylist
    except Exception:
        mas2.append(d)
        return mas2

def client_names_sort():
    values_list = worksheet.col_values(1)
    values_list2 = worksheet.col_values(2)
    values_list.pop(0)
    values_list2.pop(0)
    add = []
    add_errors = []
    out = ''
    out_err = ''
    for i in range(len(values_list2)):
        try:
            mas = sort(values_list2[i])
            for j in mas:
                if float(j) < 25:
                    add.append(values_list[i])
                    break
        except Exception:
            add_errors.append(values_list[i])

    for i in add_errors:
        out_err += i + "\n"

    for i in add:
        out += i + "\n"

    for i in add:
        if len(add) == 1:
            if len(out_err) != 0:
                if len(add_errors) > 1:
                    return f"У клиента: {out[:-2]} \n\nКритический уровень памяти.\nСвяжитесь для решения!\n\nОшибка данных клиенов {out_err[:-2]}."
                else:
                    return f"У клиента: {out[:-2]} \n\nКритический уровень памяти.\nСвяжитесь для решения!\n\nОшибка данных клиента {out_err[:-2]}."
            else:
                return f"У клиента: {out[:-2]} \n\nКритический уровень памяти.\nСвяжитесь для решения!"
        else:
            if len(out_err) != 0:
                if len(add_errors) > 1:
                    return f"У клиентов:\n {out[:-2]} \n\nКритический уровень памяти.\nСвяжитесь для решения!\n\nОшибка данных клиентов {out_err[:-2]}."
                else:
                    return f"У клиентов:\n {out[:-2]} \n\nКритический уровень памяти.\nСвяжитесь для решения!\n\nОшибка данных клиента {out_err[:-2]}."
            else:
                return f"У клиентов:\n {out[:-2]} \n\nКритический уровень памяти.\nСвяжитесь для решения!"


f = open("list.txt", 'w')
f.write(client_names_sort())
f.close()

# Использование функции send_email()
addr_to   = "support@meridiant.ru"                                # Получатель
files = ["list.txt"]                                 # Список файлов, если вложений нет, то files=[]                        # Если нужно отправить все файлы из заданной папки, нужно указать её
send_email(addr_to, "Критический уровень памяти на модуле управления клиента!", client_names_sort())

addr_to2   = "helpme.meridian@gmail.com"                                # Получатель
send_email(addr_to2, "Критический уровень памяти на модуле управления клиента!", client_names_sort())