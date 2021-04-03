"""
Модуль позволяет вести журнал автотестов.
Методы:
----------
create_xlsx_logger:
    Порождает экземпляр объекта журнала логирования в Excel документе.

"""

import os
import subprocess
from abc import abstractmethod
from enum import Enum
from time import time_ns, gmtime, strftime

import xlsxwriter
import pyautogui


def create_screenshot(directory: str) -> str:
    """
    Метод делает скриншот экрана и сохраняет результат в
    указанную директорию.
    Параметры:
    ----------
    directory: str
        Директория, в которую будет сохранен файл скриншота
    """
    image = pyautogui.screenshot()
    filename: str = "%s/%s.png" % (directory.rstrip('/'), time_ns())
    image.save(filename)

    return os.path.abspath(filename)


class LoggerMessageType(Enum):
    """
    Перечисление типов сообщений для журнала
    """
    INFO = "INFO"
    SUCCESS = "SUCCESS"
    WARNING = "WARNING"
    ERROR = "ERROR"


class ITestCaseLogger:
    """
    Интерфейс используется для журналирования результатов
    автоматического тестирования с возможностью сохранения
    скриншотов экрана.

    """

    @abstractmethod
    def save(self, open_file: bool = False) -> None:
        """
        Сохранение лога.
        Параметры:
        ----------
        open_file: bool
            Нужно ли открывать файл после сохранения
        """
        raise NotImplementedError

    @abstractmethod
    def info(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        """
        Вставка информационного сообщения
        Параметры:
        ----------
        case_name: str
            Наименование тестового кейса
        message: str
            Текст сообщения
        make_screenshot: bool
            Нужно ли вставлять скриншот
        """
        raise NotImplementedError

    @abstractmethod
    def success(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        """
        Вставка сообщения об успехе.
        Параметры:
        ----------
        case_name: str
            Наименование тестового кейса
        message: str
            Текст сообщения
        make_screenshot: bool
            Нужно ли вставлять скриншот
        """
        raise NotImplementedError

    @abstractmethod
    def warning(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        """
        Вставка предупреждающего сообщения
        Параметры:
        ----------
        case_name: str
            Наименование тестового кейса
        message: str
            Текст сообщения
        make_screenshot: bool
            Нужно ли вставлять скриншот
        """
        raise NotImplementedError

    @abstractmethod
    def error(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        """
        Вставка сообщения об ошибке.
        Параметры:
        ----------
        case_name: str
            Наименование тестового кейса
        message: str
            Текст сообщения
        make_screenshot: bool
            Нужно ли вставлять скриншот
        """
        raise NotImplementedError

    @abstractmethod
    def delete(self) -> None:
        """Удаление лога"""
        raise NotImplementedError


def create_xlsx_logger(directory: str) -> ITestCaseLogger:
    """
    Метод порождает экземпляр объекта журнала логирования в Excel документе.
    Параметры:
    ----------
    directory: str
        Директория, в которой будет сохранен файл журнала XLSX и необходимые скриншоты
    """
    return XLSXLogger(directory=directory)


class XLSXLogger(ITestCaseLogger):
    """
    Класс предоставляет возможность вести журнал логирования в Excel документе.
    Унаследован от общего интерфейса ITestCaseLogger.
    Параметры:
    ----------
    directory: str
        Директория, в которую будет сохранен журнал и скриншоты

    """

    def __init__(self, directory: str = "") -> None:
        self.work_directory = '%s/%s' % (directory.rstrip('/'), time_ns())
        self.screenshots_directory = '%s/screenshots/' % (self.work_directory)
        self.filename = '%s/%s.xlsx' % (self.work_directory, 'result')
        os.makedirs(self.screenshots_directory, exist_ok=True)
        self.workbook = xlsxwriter.Workbook(self.filename)
        self._create_worksheet()

    def _create_worksheet(self) -> None:
        """Создание и настройка листа"""
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.set_column(0, 0, 21)
        self.worksheet.set_column(1, 1, 8)
        self.worksheet.set_column(2, 2, 50)
        self.worksheet.set_column(3, 3, 50)
        self.worksheet.write('A1', 'Время')
        self.worksheet.write('B1', 'Тип')
        self.worksheet.write('C1', 'Наименование кейса')
        self.worksheet.write('D1', 'Сообщение')
        self.worksheet.write('E1', 'Скриншот')
        self._formats = {}
        self._current_row = 2

    def _get_format(self, message_type: LoggerMessageType) -> xlsxwriter.format.Format:
        """
        Получение формата ячеек для форматирования в соответствии с типом сообщения.
        Параметры:
        ----------
        message_type: LoggerMessageType
            Тип сообщения

        """
        document_format: xlsxwriter.format.Format = self._formats.get(
            message_type.value)

        if document_format is None:
            document_format = self.workbook.add_format()
            document_format.set_border(style=1)

            if message_type.value == LoggerMessageType.SUCCESS.value:
                document_format.set_bg_color('#d4ffd4')
            elif message_type.value == LoggerMessageType.WARNING.value:
                document_format.set_bg_color('#fff8d4')
            elif message_type.value == LoggerMessageType.ERROR.value:
                document_format.set_bg_color('#ffd4d4')

            self._formats[message_type.value] = document_format

        return document_format

    def _write_log(self, message_type: LoggerMessageType,
                   case_name: str, message: str = "",
                   make_screenshot: bool = False) -> None:
        """
        Общий метод создания записи сообщения в журнале
        Параметры:
        ----------
        message_type: LoggerMessageType
            Тип сообщения
        case_name: str
            Наименование тестового кейса
        message: str
            Текст сообщения
        make_screenshot: bool
            Нужно ли вставлять скриншот

        """
        document_format = self._get_format(message_type)
        self.worksheet.write('A%s' % self._current_row,
                             strftime("%a, %d %b %Y %H:%M:%S", gmtime()), document_format)
        self.worksheet.write('B%s' % self._current_row,
                             message_type.value, document_format)
        self.worksheet.write('C%s' % self._current_row,
                             case_name, document_format)
        self.worksheet.write('D%s' % self._current_row,
                             message, document_format)
        if make_screenshot:
            screenshot_file = create_screenshot(self.screenshots_directory)
            self.worksheet.write_url(
                'E%s' % self._current_row, screenshot_file,
                cell_format=document_format, string='Скриншот')
        else:
            self.worksheet.write('E%s' % self._current_row,
                                 '', document_format)

        self._current_row += 1

    def info(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        self._write_log(message_type=LoggerMessageType.INFO, case_name=case_name,
                        message=message, make_screenshot=make_screenshot)

    def success(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        self._write_log(message_type=LoggerMessageType.SUCCESS, case_name=case_name,
                        message=message, make_screenshot=make_screenshot)

    def warning(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        self._write_log(message_type=LoggerMessageType.WARNING, case_name=case_name,
                        message=message, make_screenshot=make_screenshot)

    def error(self, case_name: str, message: str = "", make_screenshot: bool = False) -> None:
        self._write_log(message_type=LoggerMessageType.ERROR, case_name=case_name,
                        message=message, make_screenshot=make_screenshot)

    def save(self, open_file: bool = False) -> None:
        self._close()
        if open_file:
            self._open()

    def _close(self) -> None:
        """Сохранение и закрытие документа"""
        self.workbook.close()

    def _open(self) -> None:
        """Открытие документа на просмотр"""
        subprocess.call(["open", (self.filename)])

    def delete(self) -> None:
        os.remove(self.filename)
