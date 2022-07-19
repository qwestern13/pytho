import sys
from this import d
from winreg import SetValue  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets
from click import progressbar
import Ui_design  # Это наш конвертированный файл дизайна
import os
from docx2pdf import convert
import math

class ExampleApp(QtWidgets.QMainWindow, Ui_design.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.btnBrowse.clicked.connect(self.browse_folder)

    def browse_folder(self):
        #self.listWidget.clear()  # На случай, если в списке уже есть элементы
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите папку")
        # открыть диалог выбора директории и установить значение переменной
        # равной пути к выбранной директории

        if directory:  # не продолжать выполнение, если пользователь не выбрал директорию
            # for file_name in os.listdir(directory):  # для каждого файла в директории
            #    self.listWidget.addItem(file_name)   # добавить файл в listWidget 
            numberOfFiles = 0
            for file_name in os.listdir(directory):
                if ".docx" in file_name:
                    numberOfFiles += 1     
            
            
            progressValue = math.ceil(100 / numberOfFiles) 

            for file_name in os.listdir(directory):
                if ".docx" in file_name:
                    convert(directory + '/' + file_name)
                    
                    if progressValue >= 100:
                        progressValue = 100
                    else:
                        progressValue = progressValue + progressValue    

                    self.progressBar.setValue(progressValue) 

                    print(str(progressValue))
                    
def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()