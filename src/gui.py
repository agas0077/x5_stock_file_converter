from tkinter import *
from pathlib import Path
from Dialog import *
from EntryField import *
from Lables import *
from script import processFile

# Создаем окно с необходимыми параметрами
window = Tk()
window.title('Преобразователь недельных стоков X5')
window.geometry("500x300")
window.resizable(width=False, height=False)
window.configure(background='#2e8b57')

# Создаем экемпляр класса диалогового окна
dialog = Dialog()
# Создаем экземпляр класса поля ввода
sheetField = EntryField(window, "Sheet1")
# Создаем экземпляр лейбла
lb = Lables()


# Функция, определяющая адреса файлов, пароль и запускающая процесс блокировки файлов
def runConverting():
    tup = dialog.getPaths()
    sheet = sheetField.getFieldValue()

    # Удаляем сообщения если есть
    lb.clearMessages()
    lb.destroyLables()

    # Проверяем выбран ли файлы
    if len(tup) == 0:
        return messagebox.showerror("Ошибка", "Необходимо выбрать файлы")

    i = 0
    errors = 0
    for path in tup:
        
        lb.addMessageToList(f"{Path(path)}", 6)
        lb.destroyLables()
        lb.createLables(background='#2e8b57')

        # Обновляем окно
        window.update()
        
        # Запускаем процесс
        res = processFile(path, sheet)

        # Вставляем сообщение
        message = f"Файл обработан. Данные в новом файле {res}"

        # Добавляем сообщение об успешном выполнении
        lb.addMessageToList(f"{message}", 6)
        lb.destroyLables()
        lb.createLables(background='#2e8b57')

        # Обновляем окно
        window.update()

        i += 1

    # Удаляем сообщения об успешном выполнении и выводим сообщение об успешном завершении процесса
    lb.destroyLables()
    lb.addMessageToList(f"ВСЕ ФАЙЛЫ БЫЛИ ОБРАБОТАНЫ! Ошибок: {errors}", 6)
    lb.createLables(background='#2e8b57')
        

# Инициализация кнопки выбора файлов
chooseFileBtn = Button(window, text="Выбрать файлы", command=dialog.callDialog, width=25, height=1)
chooseFileBtn.pack(side=TOP, padx=5, pady=5)

# Инициализация поля ввода пароля
sheetField.configure(width=30, justify=CENTER)
sheetField.pack(side=TOP, padx=5, pady=5)

# Инициализация кнопки запуска
showFileBtn = Button(window, text="Запустить обработку файлов", command=runConverting, width=25, height=1)
showFileBtn.pack(padx=5, pady=5)

window.mainloop()
