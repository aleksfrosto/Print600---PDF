#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Print600.ru - ввод данных   (версия 1.7)

Заполняет одинаковые PDF-макеты данными из списка (ФИО) и сохраняет
результат одним многостраничным PDF: одна запись = одна страница макета.

Зависимости:  pip install pymupdf fonttools openpyxl tkinterdnd2
Запуск:       python print600_vvod.py
"""

import base64
import copy
import csv
import io
import json
import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, simpledialog
import tkinter.font as tkfont

try:
    import pymupdf as fitz
except ImportError:  # старые версии PyMuPDF
    import fitz

from fontTools.ttLib import TTFont

try:  # перетаскивание файлов мышью (если библиотеки нет — просто без него)
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    TkinterDnD = DND_FILES = None

APP_TITLE = "Print600.ru - ввод данных"
APP_VERSION = "1.7"
MM = 72 / 25.4                    # пунктов в миллиметре
PDF_FONT_TAG = "P600F"            # имя шрифта в ресурсах страницы
SAMPLE_TEXT = "Иванов Иван Иванович"
FONT_EXTS = (".ttf", ".otf", ".ttc", ".otc")
LIST_EXTS = (".xlsx", ".xlsm", ".csv", ".txt")
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")
TEMPLATE_EXTS = (".pdf",) + IMAGE_EXTS
PROMO_URL = "https://print600.ru/flyer/"
DONATE_URL = "https://tips.tips/000485170"
QR_PNG = "iVBORw0KGgoAAAANSUhEUgAAAPAAAADwAQAAAAAWLtQ/AAACtklEQVR42u1ZQa6bMBB9g5GMVKlm1y4qwQ16BOjJcNSL9CYlPQm5gdmZlvh1gZNP8gmL/t/qg+oNCU8ZezQzb54nQqytBP/hpUU/PRsGNV5eNtPDcG9+2xlsSNJC1KBJsqtIgCT91uPdi9QAAs4ieV8CBwuKZBF2O8/zRI3a9wUP9e4KWPUZABztThzL12FDtgDCOYV2ONUAIKTfLSN7IAV6mUiYenjXfzh9bpsvhACA2bDfKXRs0QcA0A7HsgK4/w485MBRxFKNqs96qVrgICIi2R60Q2URFOmnjkoGiJ9ph3rZ7qbjbbqCFAs1ApoOwEGgOG7JsSkDH8G0z4KXAH3V4UjLEKZX3ZEEJnpO1nedw1FoujeTyHQo2KKRixw2pCWEJGcimH/9aPUqXHx7bb/jgngUXfQbYVn8p6vG1c9V+NPq0bp11fMHlFviJCIy+XLJcqFXdEiuqSdAmIc1TW+Mc2mzu6+nF4SEpYTXZuSiqy4BPQUNQxISlm9zH+2K8RxlvQK/R10+ht3srrV48pf47THeZQzJSSpSPBRd05oWQJBR3xTwtLND+zivaeX7D/M47Q//mJnIDmjAJ2aygJD0s4DaN9Fq8gdwuv7r+Lm8LWCRihSLJBZvfSQJpoNxtwZj9R43opl+kewaMpGgR0O6irSJBGSaTJAC6FogIIyjj0QVpqvBprXilMhoEKC8dijaKmat5h7mTPYqpcp7uCCJAJyzIQeKFg0DoP0uBockoDzMFF25p9xbZuKCGFxoC29c/MfBIcez8egr0oJguhnxL+vwM5U7TvE2HVDZBoDy2pnYgdW4OigOmylgBKghwzWgULeTYNn2WCJbhbUzDpGedzCOKaIaHNJB8jg4THAW4/Y4Z0puGbmBcIR2pqvapzvelgeH1xb7FUCW5XOhCfn/x9/S+g3U0XBzd5C0swAAAABJRU5ErkJggg=="
CACHE_VERSION = 2
FAKE_BOLD = 0.035                 # обводка для имитации жирного, доля кегля
FAKE_ITALIC = 0.21                # наклон для имитации курсива (≈12°)
# варианты выбора: (код для настроек, подпись — переводится функцией tr)
ALIGN_OPTS = [("left", "По левому краю"), ("center", "По центру"), ("right", "По правому краю")]
ZOOM_OPTS = [("fit", "По размеру окна")] + [(z, z) for z in ("50%", "75%", "100%", "150%", "200%", "300%", "400%")]
ORDER_OPTS = [("rows", "Последний номер меняется чаще: 1-1, 1-2, 1-3 …"),
              ("cols", "Первый номер меняется чаще: 1-1, 2-1, 3-1 …")]
# подписи старых версий (для переноса настроек)
OLD_ALIGN = {"По левому краю": "left", "По центру": "center", "По правому краю": "right",
             "Left": "left", "Center": "center", "Right": "right"}
OLD_ORDER = {"Ряд за рядом: 1-1, 1-2, 1-3 …": "rows", "Место за местом: 1-1, 2-1, 3-1 …": "cols"}
MAX_NUMBERS = 3
ALL = -1                          # «все номера сразу»

# --------------------------------------------------------------------------
#  Язык интерфейса
# --------------------------------------------------------------------------

LANG = "ru"
# словарь перевода на английский (ключ — исходная русская строка)
EN = {
    "Print600.ru - ввод данных": "Print600.ru - Data Input",
    "Иванов Иван Иванович": "John Michael Smith",
    "шт": "pcs", "мм": "mm", "ФИО": "names", "нумерация": "numbering", "номер: ": "number: ",
    "все": "all", "выключен": "off", "пробел": "space", "если не влезает": "if it does not fit",
    "по центру": "centered", "слева": "left", "справа": "right", "не выбран": "not selected",
    "жирный": "bold", "курсив": "italic", "Ж": "B", "К": "I", "Ч": "U",
    # ход работы
    "Размещение данных…": "Placing data…", "Оптимизация шрифтов…": "Optimizing fonts…",
    "Сохранение файла…": "Saving file…", "{t}  {d} из {n}": "{t}  {d} of {n}",
    # разделы и кнопки
    "Макет PDF": "PDF layout", "Место на макете": "Position on layout", "Данные": "Data",
    "Шрифт": "Font", "Цвет": "Color", "Перенос на новую строку": "Line breaks",
    "Сформировать PDF…": "Create PDF…", "Отменить": "Undo", "Вернуть": "Redo",
    "Масштаб:": "Zoom:", "Открыть…": "Open…", "Страница макета:": "Layout page:",
    "макет не выбран": "no layout selected", "из {n}": "of {n}",
    "X, мм": "X, mm", "Y, мм": "Y, mm", "Выравнивание": "Alignment", "Макс. ширина, мм": "Max width, mm",
    "= ширина макета": "= layout width", "Центр по горизонтали": "Center horizontally",
    "Центр страницы": "Page center",
    "Щёлкните по макету — текст встанет в эту точку; потяните за текст — он сдвинется. X — точка "
    "выравнивания, Y — базовая линия первой строки (от левого верхнего угла). Если текст шире макс. "
    "ширины, кегль уменьшается автоматически (0 — без ограничения).":
        "Click on the layout to place the text there; drag the text to move it. X is the alignment point, "
        "Y is the baseline of the first line (from the top-left corner). If the text is wider than the max "
        "width, the font size is reduced automatically (0 = no limit).",
    "По левому краю": "Left", "По центру": "Center", "По правому краю": "Right",
    "По размеру окна": "Fit to window",
    # данные: ФИО
    "Список ФИО": "Name list", "Нумерация": "Numbering",
    "По одной записи в строке. Ctrl+V, правая кнопка мыши или перетаскивание файла — можно копировать "
    "столбец прямо из Excel.":
        "One entry per line. Use Ctrl+V, the right mouse button or drag a file in — you can copy a column "
        "straight from Excel.",
    "Вставить": "Paste", "Копировать": "Copy", "Вырезать": "Cut", "Выделить всё": "Select all",
    "Очистить список": "Clear list", "Загрузить из файла…": "Load from file…", "Очистить": "Clear",
    "Показать запись №": "Show record #", "Самая длинная": "Longest",
    # данные: нумерация
    "1 номер": "1 number", "2 номера": "2 numbers", "3 номера": "3 numbers",
    "Например: номер билета · ряд и место · сектор, ряд и место.":
        "For example: ticket number · row and seat · section, row and seat.",
    "Номер {n}": "Number {n}", "с": "from", "по": "to", "шаг": "step", "цифр": "digits",
    "текст до": "text before", "после": "after", "Порядок": "Order", "Размещение": "Placement",
    "в одной строке": "on one line", "отдельно (своё место и шрифт)": "separately (own position and font)",
    "Последний номер меняется чаще: 1-1, 1-2, 1-3 …": "Last number changes fastest: 1-1, 1-2, 1-3 …",
    "Первый номер меняется чаще: 1-1, 2-1, 3-1 …": "First number changes fastest: 1-1, 2-1, 3-1 …",
    "например, номер билета": "e.g. ticket number", "например, ряд": "e.g. row",
    "например, место": "e.g. seat", "например, сектор": "e.g. section",
    "Итого: {t} {w}": "Total: {t} {w}", "Первый: «{a}»": "First: “{a}”", "последний: «{b}»": "last: “{b}”",
    "Пробелы в «до» и «после» учитываются: «от » + 1 + « шт» = «от 1 шт». Цифр — минимальное число "
    "знаков: 3 → 001.":
        "Spaces in “before” and “after” count: “No. ” + 1 + “ pcs” = “No. 1 pcs”. Digits is the minimum "
        "number of digits: 3 → 001.",
    # поля
    "Настраиваются:": "Editing:", "Все номера": "All numbers",
    "Номер 1": "Number 1", "Номер 2": "Number 2", "Номер 3": "Number 3",
    "«Все номера» — изменения применяются ко всем номерам сразу, а перемещение сдвигает их вместе. "
    "Чтобы настроить один номер, выберите его здесь или дважды щёлкните по нему на макете.":
        "“All numbers” applies changes to every number at once, and moving shifts them together. To edit a "
        "single number, select it here or double-click it on the layout.",
    "Номер — своим шрифтом и цветом": "Number in its own font and color",
    "Слова (текст до и после)": "Words (text before/after)", "Номер": "Number",
    # шрифт
    "Поиск": "Search", "Только с кириллицей": "Cyrillic only", "Шрифт из файла…": "Font from file…",
    "Размер, pt": "Size, pt", "Начертание: {f}": "Style: {f}",
    "(имитация — в шрифте нет такого начертания)": "(simulated — the font has no such style)",
    "загрузка списка шрифтов…": "loading font list…", "Загрузка шрифтов…": "Loading fonts…",
    "Загрузка списка шрифтов…": "Loading font list…", "шрифты не найдены": "no fonts found",
    "Выберите шрифт": "Choose a font",
    # цвет и пипетка
    "CMYK (печать)": "CMYK (print)", "Палитра…": "Palette…", "Цвет текста (RGB)": "Text color (RGB)",
    "Пипетка": "Eyedropper", "Пипетка вкл.": "Eyedropper on", "Готово": "Done",
    "Щелчок по макету — применить цвет\nEsc, «Готово» или крестик — выйти":
        "Click on the layout to apply the color\nEsc, “Done” or the cross to exit",
    "Наведите курсор на макет": "Hover over the layout",
    "Цвет текста взят с макета: {c}": "Text color picked from the layout: {c}",
    "Применено: {c}\nЩелчок — взять другой · Esc — выйти": "Applied: {c}\nClick to pick another · Esc to exit",
    # перенос
    "Заменить на перенос пробел №": "Replace with a line break: space #",
    "только если текст не помещается в макс. ширину": "only if the text does not fit the max width",
    "Межстрочный интервал, %": "Line spacing, %",
    "Например, «1» — фамилия на первой строке, имя и отчество на второй; «1» и «2» — три строки. Для "
    "нумерации: «Ряд 5» → «Ряд» над «5».":
        "For example, “1” puts the first word on the first line and the rest on the second; “1” and “2” "
        "give three lines. For numbering: “Row 5” → “Row” above “5”.",
    # предпросмотр
    "Запись {i} из {n}": "Record {i} of {n}", "Образец текста": "Sample text", "ширина {w} мм": "width {w} mm",
    "строк: {n}": "lines: {n}", "кегль уменьшен до {p}%, чтобы влезть": "font size reduced to {p}% to fit",
    "(!) текст выходит за край страницы": "(!) text goes beyond the page edge",
    "(!) в шрифте нет символов: {s}": "(!) characters missing from the font: {s}",
    "Ошибка предпросмотра: {e}": "Preview error: {e}", "Ошибка шрифта: {e}": "Font error: {e}",
    # файлы и сообщения
    "Шрифты": "Fonts", "Файл со списком ФИО": "Name list file", "Файл шрифта": "Font file",
    "Сохранить результат": "Save result",
    "Это не PDF-файл.": "This is not a PDF file.", "Файл защищён паролем.": "The file is password-protected.",
    "В файле нет страниц.": "The file has no pages.",
    "Не удалось открыть макет:\n{path}\n\n{e}": "Could not open the layout:\n{path}\n\n{e}",
    "Сначала откройте PDF-макет.": "Please open a PDF layout first.",
    "В файле нет данных.": "The file contains no data.",
    "В выбранных столбцах нет данных.": "The selected columns contain no data.",
    "Не удалось прочитать файл:\n{e}": "Could not read the file:\n{e}",
    "В таблице несколько столбцов. Первая строка:\n\n{lines}\n\nКакие столбцы взять? Номера через запятую "
    "(например: 1  или  1,2,3).\nНесколько столбцов объединятся через пробел.":
        "The table has several columns. First row:\n\n{lines}\n\nWhich columns should be used? Enter numbers "
        "separated by commas (e.g. 1  or  1,2,3).\nSeveral columns will be joined with a space.",
    "Загружено записей: {n}.\n\nЗаменить текущий список?\n(«Нет» — добавить в конец)":
        "Entries loaded: {n}.\n\nReplace the current list?\n(“No” appends to the end)",
    "Для чтения Excel установите пакет openpyxl:\npip install openpyxl":
        "To read Excel files, install the openpyxl package:\npip install openpyxl",
    "Старый формат .xls не поддерживается.\nПересохраните файл в Excel как .xlsx или .csv.":
        "The old .xls format is not supported.\nPlease re-save the file in Excel as .xlsx or .csv.",
    "Ошибка при чтении шрифтов:\n{e}": "Error reading fonts:\n{e}",
    "Не удалось загрузить шрифт «{f}»:\n{e}": "Could not load the font “{f}”:\n{e}",
    "Не удалось загрузить шрифт:\n{e}": "Could not load the font:\n{e}",
    "Этот шрифт не поддерживается.\nНужен шрифт TrueType или OpenType (.ttf, .otf, .ttc).":
        "This font is not supported.\nA TrueType or OpenType font (.ttf, .otf, .ttc) is required.",
    "Выберите шрифт.": "Please choose a font.", "Список ФИО пуст.": "The name list is empty.",
    "Нет номеров.": "There are no numbers.",
    "Получится {n} страниц — это займёт время и большой файл. Продолжить?":
        "This will produce {n} pages — it will take time and create a large file. Continue?",
    "В шрифтах нет некоторых символов — они не будут видны:":
        "Some characters are missing from the fonts and will not be visible:",
    "Продолжить?": "Continue?",
    "Нельзя сохранять поверх исходного макета.\nВыберите другое имя.":
        "You cannot overwrite the source layout.\nPlease choose another name.",
    "Не удалось сформировать PDF:\n\n{e}": "Could not create the PDF:\n\n{e}", "Ошибка.": "Error.",
    "Готово!\n\nСтраниц: {n}\n{path}": "Done!\n\nPages: {n}\n{path}",
    "Страниц с уменьшенным кеглем (текст не влезал в макс. ширину): {n}":
        "Pages with a reduced font size (text did not fit the max width): {n}",
    "Готово: {f}": "Done: {f}", "Открыть файл?": "Open the file?",
    "Выберите макет": "Choose a layout",
    "Макеты (PDF, JPG, PNG, TIFF)": "Layouts (PDF, JPG, PNG, TIFF)",
    "Изображения": "Images",
    "Разрешение, dpi": "Resolution, dpi",
    "Макет можно просто перетащить мышью в окно предпросмотра. Поддерживаются PDF, JPG, PNG и TIFF; у картинки размер макета считается по её разрешению.": "You can simply drag a layout onto the preview area. PDF, JPG, PNG and TIFF are supported; for an image the layout size is calculated from its resolution.",
    "Нажмите, чтобы выбрать макет": "Click to choose a layout",
    "PDF, JPG, PNG, TIFF — или перетащите файл сюда": "PDF, JPG, PNG, TIFF — or drag a file here",
    "PDF, JPG, PNG, TIFF": "PDF, JPG, PNG, TIFF",
    "Можно перетащить макет (PDF, JPG, PNG, TIFF), список (.xlsx, .csv, .txt) или файл шрифта (.ttf, .otf).": "You can drag a layout (PDF, JPG, PNG, TIFF), a list (.xlsx, .csv, .txt) or a font file (.ttf, .otf).",
    "Напечатать быстро и дешево можно тут": "Print it fast and cheap here",
    "Готово! Напечатать быстро и дешево можно тут": "Done! Print it fast and cheap here",
    "Все файлы": "All files", "Списки": "Lists", "Текст": "Text",
    "Поддержать разработку": "Support us",
    "Если наша программа помогла решить вашу проблему, будем рады вашей поддержке": "If this program solved your problem, we would be glad of your support",
    "Открыть ссылку": "Open the link",
    "Копировать ссылку": "Copy the link",
    "Ссылка скопирована": "Link copied",
    "Закрыть": "Close",
    "Сбросить настройки": "Reset",
    "Сбросить все настройки?\n\nШрифт, цвет, место, нумерация и перенос вернутся к исходным. Макет и список останутся.\nДействие можно отменить по Ctrl+Z.": "Reset all settings?\n\nFont, color, position, numbering and line breaks will return to their defaults. The layout and the list will stay.\nYou can undo this with Ctrl+Z.",
    "Настройки сброшены (Ctrl+Z — вернуть).": "Settings have been reset (Ctrl+Z to undo).",
    "Непредвиденная ошибка:\n\n{e}\n\nПодробности записаны в файл:\n{f}":
        "Unexpected error:\n\n{e}\n\nDetails were written to:\n{f}",
}


def set_lang(lang):
    global LANG
    LANG = lang if lang in ("ru", "en") else "ru"


def tr(text):
    """Перевод строки интерфейса на текущий язык (русский — исходный)."""
    return EN.get(text, text) if LANG == "en" else text


def count_word(n):
    """«макет / макета / макетов» или «layout(s)»."""
    if LANG == "en":
        return "layout" if n == 1 else "layouts"
    return plural(n, "макет", "макета", "макетов")

HEADER_WORDS = {"фио", "ф.и.о.", "ф.и.о", "ф. и. о.", "фамилия имя отчество",
                "фамилия", "имя", "name", "full name", "fio"}
PAD = 20                          # отступ листа в окне предпросмотра, px

# фирменные цвета
C_ACCENT = "#0093DA"
C_ACCENT_DARK = "#007CB9"
C_ACCENT_LIGHT = "#D9ECF7"
C_BG = "#F2F6FA"
C_CARD = "#FFFFFF"
C_HOVER = "#F6FAFD"
C_BORDER = "#D3E2EE"
C_TEXT = "#1E2B37"
C_MUTED = "#6A7D8E"
C_CANVAS = "#DCE5ED"
C_MARK = "#E0245E"
C_DONATE = "#FF6B00"          # яркая кнопка поддержки
C_DONATE_DARK = "#E25D00"
C_HEAD = "#D9ECF7"            # фон заголовков разделов
C_HEAD_HOVER = "#C8E3F4"
C_HEAD_BORDER = "#B7D5EA"
C_HEAD_TEXT = "#3D5A72"
ICON_PNG = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAFSElEQVR4nO1be1BUVRz+dmGldmXBEkRDa0pAXsobZkVSExIRTcSiQYomchhFnREdJrFinNTMsrHMSUNqRBxqDMosHV+8gwmRBJNHUE2WjbDyWNmFcDf6axfOPXvv4nCdM7n3++ue73fOd7/z3XvP3nt3rgwCmOkTOiJU/7/gRluDjK9mtfCgTJwLa0EQxIM6cS7GBiE3b9jL5AFyrnKhjvYAOWBfR98M85zt/gyQ2ePRHwu7PwOkAFgbYA0pANYGWEMKgLUB1rD7ABzvpXOA3xysSoyHJjIMHh7uUDtPRk9vH7q6tWj+uQXnLlagurYew8PDvBrhoUFIfm45oiJC4O42FXK5HN3dt3G58SpKT51BRXXtuLyIpTOuO8HJKiX27MzFimVxkMl43y0AAHa9ewCHCwopXqFQYN/bbyBp5TLB8RfKKrExewf0BoPVulg6Zti8BFxd1DhZlI+VCc/anLwQPjmw16ZpAFiyKAaF+R/B0cHhvuqYYTOA93a/Bb853pa20WRC/ucnEJ+UCp+gaPiGxGDJ8ueRlb0dpd+eweDQEKWRkrwSsYtjCO6zwmIEa2IRELEIHxw8QtTCQuYhM+Pl+6YzFoKXwMIFGhz79EOCe3X9Fpy/VCkoSuxAJkPVua8xa+ZjFu7a9VYkrE7DyMjork8UHEK0JsLS7uvXIWzBUst6IpYOF4JnwCtpKUT7/KXKe5o8AIQGzyVMA8DpsxcI02ZuLFxd1Hjm6WjRdbjgDUChUEATGUZwNXX1WL8uHadPHkPLlUp0Nteiruw7HHx/N6LCQ6zqhAbNpbjW9g6Ka2lrp7jgoADRdbjg/Rn09/WGk9MkgsvdtgkKhYLgZkyfhhUJcViREIeiL0qwY+c7MJn+tdR9vJ+itG91dVNcV5eW4ny8ZouuwwXvGeDuNpXiuJPnIvWFJORsySI4F7Wa6mfQD9LcIM25uoyOFUuHC94A1M7OVvnW9g4sT34JfqExSMvYiG7tbaKekZ6Kx2d5WtrKhx+iNIwmI8XdNdKcSqUUXYcL3gCG71pfNbfl7kTTtesY0BtQUV2Lvfs/JuqODg5YFrfY0jYM0j+LCkf6yrPG6fWjNzFi6XDBG4BON0BxA3oDrjZfJ7iaunqqn9fsJy3b/TodVVcq6SNijRs7ViwdLngD6Pztd4rr7e0bF+c0aXTxbGvvpOrT3N0ozsMK1/bL6FixdLjgDeDGnzehvd1DcFOmuFL9rHFj14WGn5qouq+PF8XN8aFX6iuNzaLrcCF4I/TN6bNEe7JKiXmBfgQ3PyqcGvfj5UbLdkNjE/648RdRT1i6hHquSIyPJdp9/TpcrKgWXYcLwQAKCoupe/t9u95EoL8vVEolYuZHIWfLBqJ+8+9buFA+usORkREcPFxA9PH39UZe7lY8MsUVarUzsjdlQsMJ8kjBceL2VSwdLmw+Dq9NWY3dea8LdbHAaDIhfd1mVNbUUbWjh/ZTDzJ8aGhswpq1r8FoMt03HTNsPg0eL/4K2/P2YGjoH8F+ff06ZGzItjp5AMjcnIOSU9/b2h0ullchLSOL17RYOmaM+6+xGdOnIe3FZCyM1sDTczpUKhV0ujto7/gVZRXVKPqyFDrdHZs6EWHBWLMqEZHhwXB3mwqZTAattsfyJqe86ofx2BFNR/pvkLUB1pACYG2ANaQAWBtgDSkA1gZYQwqAtQHWkAJgbYA1pABYG2ANKQDWBlhDCoC1Adaw+wBkKOmd0CsxbdLRCRl4VLd1QuNnBz4xofF2fwZIAbA2wBpSAKwNsIbdB2D3n85KH08LDXhQghD6fP4/r1w6cBoVaZIAAAAASUVORK5CYII="


# --------------------------------------------------------------------------
#  Вспомогательное
# --------------------------------------------------------------------------

def config_dir():
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    from_store = "\\windowsapps\\" in exe_dir.lower() + "\\"   # установлено из Microsoft Store (MSIX)
    if getattr(sys, "frozen", False) and not from_store:
        # портативная версия (exe): настройки рядом с программой,
        # если туда можно писать; иначе — в профиле пользователя
        path = os.path.join(exe_dir, "Print600.ru_data")
        try:
            os.makedirs(path, exist_ok=True)
            probe = os.path.join(path, ".probe")
            with open(probe, "w") as f:
                f.write("ok")
            os.remove(probe)
            return path
        except OSError:
            pass
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    path = os.path.join(base, "Print600")
    os.makedirs(path, exist_ok=True)
    return path


def num(value, default=0.0):
    """Число из строки; понимает и запятую, и точку."""
    try:
        return float(str(value).strip().replace(",", "."))
    except (ValueError, TypeError):
        return default


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


def fmt(v):
    """12.0 -> '12', 12.35 -> '12.35'."""
    return f"{v:.2f}".rstrip("0").rstrip(".")


def open_pdf(path):
    """Открывает PDF через Python — так работают пути с русскими буквами на Windows."""
    with open(path, "rb") as f:
        data = f.read()
    return fitz.open(stream=data, filetype="pdf")


def save_pdf(doc, path):
    """Сохраняет PDF через Python (кириллица в пути) и атомарно заменяет файл."""
    data = doc.tobytes(garbage=3, deflate=True)
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)


def image_dpi(data):
    """Разрешение из метаданных картинки (0, если его нет или оно явно экранное)."""
    try:
        pix = fitz.Pixmap(data)
        dpi = max(pix.xres or 0, pix.yres or 0)
    except Exception:
        return 0
    return dpi if dpi >= 150 else 0


def pdf_from_image(path, dpi=None):
    """Одностраничный PDF из JPG/PNG/TIFF: размер страницы = пиксели ÷ dpi.
    Возвращает (байты PDF, ширина px, высота px, использованное dpi)."""
    with open(path, "rb") as f:
        data = f.read()
    pix = fitz.Pixmap(data)
    used = int(dpi or image_dpi(data) or 300)
    doc = fitz.open()
    page = doc.new_page(width=pix.width * 72.0 / used, height=pix.height * 72.0 / used)
    page.insert_image(page.rect, stream=data)
    out = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return out, pix.width, pix.height, used


def open_pdf_src(src):
    """Макет из файла (путь) или из готовых байтов PDF (для картинок)."""
    if isinstance(src, (bytes, bytearray)):
        return fitz.open(stream=src, filetype="pdf")
    return open_pdf(src)


def open_with_system(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# --------------------------------------------------------------------------
#  Шрифты
# --------------------------------------------------------------------------

def system_font_dirs():
    home = os.path.expanduser("~")
    if sys.platform.startswith("win"):
        dirs = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")]
        if os.environ.get("LOCALAPPDATA"):  # шрифты, установленные «для текущего пользователя»
            dirs.append(os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Windows", "Fonts"))
    elif sys.platform == "darwin":
        dirs = ["/System/Library/Fonts", "/Library/Fonts", os.path.join(home, "Library", "Fonts")]
    else:
        dirs = ["/usr/share/fonts", "/usr/local/share/fonts",
                os.path.join(home, ".fonts"), os.path.join(home, ".local", "share", "fonts")]
    return [d for d in dirs if os.path.isdir(d)]


def _ttc_count(path):
    with open(path, "rb") as f:
        head = f.read(12)
    if head[:4] != b"ttcf":
        return 0
    return int.from_bytes(head[8:12], "big")


def _face_info(tt):
    """Семейство, начертание, кириллица и параметры подчёркивания одного шрифта."""
    if "glyf" not in tt and "CFF " not in tt:
        return None  # растровые, CFF2 и прочая экзотика
    names = tt["name"]
    family = (names.getDebugName(1) or names.getDebugName(16) or "").strip()
    style = (names.getDebugName(2) or names.getDebugName(17) or "Regular").strip()
    if not family:
        return None
    full = (names.getDebugName(4) or f"{family} {style}").strip()
    if "OS/2" in tt:
        sel = tt["OS/2"].fsSelection
        bold, italic = bool(sel & 0x20), bool(sel & 0x01)
    else:
        mac = tt["head"].macStyle
        bold, italic = bool(mac & 1), bool(mac & 2)
    cmap = tt.getBestCmap() or {}
    upm = float(tt["head"].unitsPerEm or 1000)
    ul_pos, ul_th = -0.1, 0.05
    try:
        post = tt["post"]
        if post.underlineThickness > 0:
            ul_pos, ul_th = post.underlinePosition / upm, post.underlineThickness / upm
    except Exception:
        pass
    if not -0.5 < ul_pos < -0.02:
        ul_pos = -0.1
    if not 0.01 < ul_th < 0.2:
        ul_th = 0.05
    return {"family": " ".join(family.split()), "style": " ".join(style.split()),
            "full": " ".join(full.split()), "bold": bold, "italic": italic,
            "cyr": 0x0416 in cmap and 0x044F in cmap,          # «Ж» и «я»
            "ul_pos": round(ul_pos, 4), "ul_th": round(ul_th, 4)}


def scan_font_file(path):
    """Начертания в файле: [{'family', 'style', 'bold', 'italic', 'index', ...}]."""
    faces = []
    try:
        count = _ttc_count(path) if path.lower().endswith((".ttc", ".otc")) else 0
        for idx in (range(count) if count else [0]):
            tt = TTFont(path, fontNumber=idx if count else -1, lazy=True)
            try:
                info = _face_info(tt)
            finally:
                tt.close()
            if info:
                info["index"] = idx
                faces.append(info)
    except Exception:
        pass
    return faces


def scan_system_fonts(cache_file):
    """Все начертания шрифтов системы. Результат кэшируется."""
    try:
        with open(cache_file, encoding="utf-8") as f:
            cache = json.load(f)
        files = cache["files"] if cache.get("v") == CACHE_VERSION else {}
    except Exception:
        files = {}
    fresh = {}
    for folder in system_font_dirs():
        for root, _dirs, names in os.walk(folder):
            for fn in names:
                if not fn.lower().endswith(FONT_EXTS):
                    continue
                path = os.path.join(root, fn)
                try:
                    st = os.stat(path)
                except OSError:
                    continue
                sig = [st.st_size, int(st.st_mtime)]
                old = files.get(path)
                faces = old["faces"] if old and old.get("sig") == sig else scan_font_file(path)
                fresh[path] = {"sig": sig, "faces": faces}
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump({"v": CACHE_VERSION, "files": fresh}, f, ensure_ascii=False)
    except Exception:
        pass
    result = []
    for path, rec in fresh.items():
        for face in rec["faces"]:
            result.append(dict(face, path=path))
    return result


def group_families(faces):
    """Группирует начертания в семейства (как в Word: Обычный / Ж / К / ЖК)."""
    fams = {}
    for face in faces:
        key = face["family"].casefold() + ("|file" if face.get("from_file") else "")
        fam = fams.get(key)
        if fam is None:
            fam = fams[key] = {"name": face["family"], "faces": {}, "from_file": bool(face.get("from_file"))}
        fam["faces"].setdefault((face["bold"], face["italic"]), face)
    system_names = {f["name"].casefold() for f in fams.values() if not f["from_file"]}
    result = []
    for fam in fams.values():
        if fam["from_file"] and fam["name"].casefold() in system_names:
            fam["name"] += " (файл)"   # не переводится: имя шрифта должно быть постоянным
        regular = fam["faces"].get((False, False)) or next(iter(fam["faces"].values()))
        fam["cyr"] = regular["cyr"]
        result.append(fam)
    result.sort(key=lambda f: f["name"].casefold())
    return result


def resolve_face(family, bold, italic):
    """Подбирает начертание. Если нужного нет — возвращает ближайшее и что имитировать."""
    faces = family["faces"]
    if (bold, italic) in faces:
        return faces[(bold, italic)], False, False
    if italic and (bold, False) in faces:
        return faces[(bold, False)], False, True
    if bold and (False, italic) in faces:
        return faces[(False, italic)], True, False
    base = faces.get((False, False)) or next(iter(faces.values()))
    return base, bold and not base["bold"], italic and not base["italic"]


_FONT_CACHE = {}


def load_font(face):
    """(байты шрифта, fitz.Font). Начертание из .ttc извлекается в отдельный TTF."""
    key = (face["path"], face["index"])
    if key not in _FONT_CACHE:
        if face["path"].lower().endswith((".ttc", ".otc")):
            tt = TTFont(face["path"], fontNumber=face["index"])
            buf = io.BytesIO()
            tt.save(buf)
            tt.close()
            data = buf.getvalue()
        else:
            with open(face["path"], "rb") as f:
                data = f.read()
        _FONT_CACHE[key] = (data, fitz.Font(fontbuffer=data))
    return _FONT_CACHE[key]


# --------------------------------------------------------------------------
#  Размещение текста и сборка PDF
# --------------------------------------------------------------------------

STYLE_KEYS = ("size", "bold", "italic", "underline", "mode", "c", "m", "y", "k", "r", "g", "b")
FIELD_KEYS = ("x", "ypos", "align", "maxw", "br1", "br2", "br3", "wrap", "leading", "num_sep")


def default_style():
    return {"family": "", "size": "14", "bold": False, "italic": False, "underline": False,
            "mode": "CMYK", "c": "0", "m": "0", "y": "0", "k": "100", "r": "0", "g": "0", "b": "0"}


def default_field():
    return {"x": "0", "ypos": "0", "align": "center", "maxw": "0",
            "br1": False, "br2": False, "br3": False, "wrap": True, "leading": "120",
            "num_sep": False, "init": False, "styles": {"text": default_style(), "num": None}}


def clean_field(saved):
    """Поле из сохранённых настроек (с подстановкой значений по умолчанию)."""
    f = default_field()
    if isinstance(saved, dict):
        for k in FIELD_KEYS + ("init",):
            if k in saved:
                f[k] = saved[k]
        f["align"] = OLD_ALIGN.get(f["align"], f["align"])
        if f["align"] not in ("left", "center", "right"):
            f["align"] = "center"
        styles = saved.get("styles") or {}
        for part in ("text", "num"):
            st = styles.get(part)
            if isinstance(st, dict):
                f["styles"][part] = dict(default_style(), **{k: v for k, v in st.items()
                                                             if k in STYLE_KEYS + ("family",)})
    return f


def font_tag(face):
    import hashlib
    return "P6" + hashlib.md5(f"{face['path']}|{face['index']}".encode("utf-8")).hexdigest()[:8]


def num_range(frm, to, step):
    step = max(1, abs(int(step)))
    return range(frm, to + 1, step) if frm <= to else range(frm, to - 1, -step)


def plural(n, one, few, many):
    n = abs(n) % 100
    if 11 <= n <= 14:
        return many
    return {1: one, 2: few, 3: few, 4: few}.get(n % 10, many)


def fmt_num(n, digits):
    return f"{n:0{digits}d}" if digits > 0 else str(n)


def split_runs(runs, breaks):
    """Заменяет выбранные по счёту пробелы (1, 2, 3…) на перенос строки.
    runs — [(текст, часть)], результат — строки из таких же кусков."""
    lines = [[]]
    count = 0
    for text, key in runs:
        buf = ""
        for ch in text:
            if ch == " ":
                count += 1
                if count in breaks:
                    if buf:
                        lines[-1].append((buf, key))
                    buf = ""
                    lines.append([])
                    continue
            buf += ch
        if buf:
            lines[-1].append((buf, key))
    return [line for line in lines if line]


def seg_width(text, st, k=1.0):
    size = st["size"] * k
    w = st["font"].text_length(text, fontsize=size)
    if st["fake_bold"]:
        w += size * FAKE_BOLD
    return w


def line_size(line, styles, k=1.0):
    return max((styles[key]["size"] for _, key in line), default=styles["text"]["size"]) * k


def layout_runs(runs, spec):
    """Строки, коэффициент уменьшения кегля и ширины строк."""
    runs = [(t, key) for t, key in runs if t]
    styles = spec["styles"]

    def line_w(line, k=1.0):
        return sum(seg_width(t, styles[key], k) for t, key in line)

    lines = [runs]
    max_w = spec["max_w"] * MM
    if spec["breaks"] and runs:
        too_wide = max_w > 0 and line_w(runs) > max_w
        if too_wide or not spec["wrap_if_needed"]:
            lines = split_runs(runs, spec["breaks"]) or [runs]
    widths = [line_w(line) for line in lines]
    k = 1.0
    for _ in range(2):   # второй проход — поправка на имитацию жирного
        if max_w > 0 and max(widths) > max_w + 0.01:
            k *= max_w / max(widths)
            widths = [line_w(line, k) for line in lines]
    return lines, k, widths


def place_runs(page, runs, spec):
    """
    Ставит на страницу текст из кусков разного оформления (слова / номер).
    Координаты — мм от левого верхнего угла видимой страницы, y — базовая линия
    первой строки. Возвращает сведения о размещении (включая рамку текста).
    """
    lines, k, widths = layout_runs(runs, spec)
    x0, y0 = spec["x"] * MM, spec["y"] * MM
    if not lines or not lines[0]:
        return {"k": 1.0, "lines": [], "widths": [0], "lefts": [x0], "baselines": [y0], "bbox": None}
    styles = spec["styles"]
    D = page.derotation_matrix
    R = fitz.Matrix(D.a, D.b, D.c, D.d, 0, 0)
    shear = ~R * fitz.Matrix(1, 0, FAKE_ITALIC, 1, 0, 0) * R
    for key in {key for line in lines for _, key in line}:
        page.insert_font(fontname=styles[key]["tag"], fontbuffer=styles[key]["data"])
    sizes = [line_size(line, styles, k) for line in lines]
    y = y0
    lefts, baselines = [], []
    for i, (line, w) in enumerate(zip(lines, widths)):
        if i:
            y += spec["leading"] * max(sizes[i - 1], sizes[i])
        x = x0
        if spec["align"] == "center":
            x -= w / 2
        elif spec["align"] == "right":
            x -= w
        lefts.append(x)
        baselines.append(y)
        for text, key in line:
            st = styles[key]
            size = st["size"] * k
            sw = seg_width(text, st, k)
            point = fitz.Point(x, y) * D
            extra = {"render_mode": 2, "border_width": FAKE_BOLD} if st["fake_bold"] else {}
            page.insert_text(point, text, fontsize=size, fontname=st["tag"], color=st["color"],
                             fill=st["color"], rotate=page.rotation,
                             morph=(point, shear) if st["fake_italic"] else None, **extra)
            if st["underline"]:
                top = y - size * st["ul_pos"]
                thick = max(0.25, size * st["ul_th"])
                page.draw_rect(fitz.Rect(x, top, x + sw, top + thick) * D, color=None,
                               fill=st["color"], width=0)
            x += sw
    right = max(left + w for left, w in zip(lefts, widths))
    bbox = fitz.Rect(min(lefts), baselines[0] - sizes[0] * 1.05, right, baselines[-1] + sizes[-1] * 0.35)
    return {"k": k, "lines": lines, "widths": widths, "lefts": lefts, "baselines": baselines,
            "bbox": bbox, "size": max(sizes)}


def materialize_inherited(doc):
    """Прописывает унаследованные от дерева страниц атрибуты прямо в страницы,
    чтобы копии страниц были полностью самостоятельными."""
    for pno in range(doc.page_count):
        xref = doc.page_xref(pno)
        for key in ("Resources", "MediaBox", "CropBox", "Rotate"):
            if doc.xref_get_key(xref, key)[0] != "null":
                continue
            parent, seen = doc.xref_get_key(xref, "Parent"), set()
            while parent[0] == "xref":
                pxref = int(parent[1].split()[0])
                if pxref in seen:
                    break
                seen.add(pxref)
                kind, value = doc.xref_get_key(pxref, key)
                if kind != "null":
                    doc.xref_set_key(xref, key, value)
                    break
                parent = doc.xref_get_key(pxref, "Parent")


def build_pdf(template_path, out_path, count, items, page_index, progress=None):
    """Размножает страницу макета: count страниц, на каждой — поля items(i) = [(куски, настройки)].
    Возвращает число страниц, где кегль уменьшался по ширине."""
    doc = open_pdf_src(template_path)
    materialize_inherited(doc)
    n_tpl = doc.page_count
    shrunk = 0
    for i in range(count):
        doc.fullcopy_page(page_index)
        page = doc[doc.page_count - 1]
        small = False
        for runs, spec in items(i):
            if place_runs(page, runs, spec)["k"] < 1 - 1e-6:
                small = True
        shrunk += small
        if progress and (i % 20 == 0 or i == count - 1):
            progress(i + 1, count, tr("Размещение данных…"))
    doc.delete_pages(0, n_tpl - 1)  # убираем исходные страницы макета
    if progress:
        progress(count, count, tr("Оптимизация шрифтов…"))
    try:
        doc.subset_fonts()  # оставляем в файле только нужные символы шрифта
    except Exception:
        pass
    meta = doc.metadata or {}
    meta["creator"] = APP_TITLE
    try:
        doc.set_metadata(meta)
    except Exception:
        pass
    if progress:
        progress(count, count, tr("Сохранение файла…"))
    save_pdf(doc, out_path)
    doc.close()
    return shrunk


# --------------------------------------------------------------------------
#  Список ФИО
# --------------------------------------------------------------------------

def clean_line(line):
    return re.sub(r"\s+", " ", line.replace("\t", " ")).strip()


def parse_names(text):
    return [s for s in (clean_line(l) for l in text.splitlines()) if s]


def read_text_file(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "cp1251", "utf-16"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def read_table(path):
    """Строки таблицы (список списков строк) из .xlsx / .csv / .txt."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            raise RuntimeError(tr("Для чтения Excel установите пакет openpyxl:\npip install openpyxl"))
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = [["" if c is None else str(c).strip() for c in row]
                for row in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    if ext == ".xls":
        raise RuntimeError(tr("Старый формат .xls не поддерживается.\nПересохраните файл в Excel как .xlsx или .csv."))
    text = read_text_file(path)
    if ext == ".csv" or ("\t" in text or ";" in text):
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=";,\t")
            return [[c.strip() for c in row] for row in csv.reader(io.StringIO(text), dialect)]
        except csv.Error:
            pass
    return [[line] for line in text.splitlines()]


def make_icon(kind, size, color, fill=None):
    """Значок, нарисованный программно (не зависит от шрифтов системы)."""
    doc = fitz.open()
    pg = doc.new_page(width=size, height=size)

    def rgb(c):
        return tuple(int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))

    col = rgb(color)
    fil = rgb(fill) if fill else None
    sh = pg.new_shape()

    def P(x, y):
        return fitz.Point(x * size, y * size)

    box = fitz.Rect(P(.1, .1), P(.9, .9))
    if kind in ("down", "right"):
        pts = [P(.2, .34), P(.8, .34), P(.5, .7)] if kind == "down" else [P(.34, .2), P(.7, .5), P(.34, .8)]
        sh.draw_polyline(pts + [pts[0]])
        sh.finish(color=None, fill=col, closePath=True)
    elif kind == "close":
        sh.draw_line(P(.24, .24), P(.76, .76))
        sh.draw_line(P(.76, .24), P(.24, .76))
        sh.finish(color=col, width=size * .12, lineCap=1)
    elif kind in ("undo", "redo"):
        m = (lambda x: x) if kind == "undo" else (lambda x: 1 - x)
        sh.draw_bezier(P(m(.8), .86), P(m(.92), .44), P(m(.66), .3), P(m(.36), .32))
        sh.finish(color=col, width=size * .12, lineCap=1, closePath=False)
        sh.draw_polyline([P(m(.1), .32), P(m(.42), .1), P(m(.42), .54), P(m(.1), .32)])
        sh.finish(color=None, fill=col, closePath=True)
    elif kind == "heart":
        sh.draw_circle(P(.3, .36), size * .2)
        sh.draw_circle(P(.7, .36), size * .2)
        sh.finish(color=None, fill=col)
        sh.draw_polyline([P(.1, .42), P(.9, .42), P(.5, .88), P(.1, .42)])
        sh.finish(color=None, fill=col, closePath=True)
    elif kind == "box":          # квадрат галочки (выкл.)
        sh.draw_rect(box, radius=.22)
        sh.finish(color=col, fill=fil, width=size * .08)
    elif kind == "box_on":       # квадрат галочки (вкл.)
        sh.draw_rect(box, radius=.22)
        sh.finish(color=col, fill=col, width=size * .08)
        sh.draw_polyline([P(.27, .52), P(.44, .68), P(.74, .34)])
        sh.finish(color=(1, 1, 1), width=size * .12, lineCap=1, lineJoin=1, closePath=False)
    elif kind == "ring":         # переключатель (выкл.)
        sh.draw_circle(P(.5, .5), size * .4)
        sh.finish(color=col, fill=fil, width=size * .08)
    elif kind == "ring_on":      # переключатель (вкл.)
        sh.draw_circle(P(.5, .5), size * .4)
        sh.finish(color=col, fill=(1, 1, 1), width=size * .08)
        sh.draw_circle(P(.5, .5), size * .2)
        sh.finish(color=None, fill=col)
    sh.commit()
    pix = pg.get_pixmap(alpha=True)
    doc.close()
    return tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")))


# --------------------------------------------------------------------------
#  Сворачиваемый раздел
# --------------------------------------------------------------------------

class Section(tk.Frame):
    """Карточка с цветным заголовком: щелчок по заголовку сворачивает/разворачивает."""

    def __init__(self, app, parent, num_, title, key):
        super().__init__(parent, bg=C_CARD, highlightthickness=1,
                         highlightbackground=C_HEAD_BORDER, highlightcolor=C_HEAD_BORDER)
        self.app, self.key = app, key
        self.expanded = app.settings.get("sections", {}).get(key, True)
        px = app.px
        self.head = tk.Frame(self, bg=C_HEAD, cursor="hand2")
        self.head.pack(fill="x")
        d = px(22)
        self.badge = tk.Canvas(self.head, width=d, height=d, bg=C_HEAD, highlightthickness=0, cursor="hand2")
        self.badge.create_oval(1, 1, d - 1, d - 1, fill=C_ACCENT, outline="")
        self.badge.create_text(d / 2, d / 2, text=str(num_), fill="white", font=app.f_small_bold)
        self.badge.pack(side="left", padx=(px(12), px(9)), pady=px(9))
        self.title = tk.Label(self.head, text=title, bg=C_HEAD, fg=C_TEXT, font=app.f_title, cursor="hand2")
        self.title.pack(side="left")
        self.chev = tk.Label(self.head, bg=C_HEAD, cursor="hand2")
        self.chev.pack(side="right", padx=(px(6), px(12)))
        self.summary = tk.Label(self.head, bg=C_HEAD, fg=C_HEAD_TEXT, font=app.f_small, cursor="hand2")
        self.summary.pack(side="right")
        self.swatch = None
        self.body = ttk.Frame(self, padding=(px(14), px(10), px(14), px(12)))
        self._head_widgets = [self.head, self.badge, self.title, self.chev, self.summary]
        for w in self._head_widgets:
            w.bind("<Button-1>", self.toggle)
            w.bind("<Enter>", lambda e: self._hover(True))
            w.bind("<Leave>", lambda e: self._hover(False))
        self._apply()

    def add_swatch(self):
        s = self.app.px(14)
        self.swatch = tk.Canvas(self.head, width=s, height=s, bg="#000000", highlightthickness=1,
                                highlightbackground="white", cursor="hand2")
        self.swatch.pack(side="right", padx=(self.app.px(6), 0), before=self.summary)
        self.swatch.bind("<Button-1>", self.toggle)
        return self.swatch

    def _hover(self, on):
        bg = C_HEAD_HOVER if on else C_HEAD
        for w in self._head_widgets:
            w.configure(bg=bg)

    def _apply(self):
        self.chev.configure(image=self.app.icons["down" if self.expanded else "right"])
        if self.expanded:
            self.body.pack(fill="x")
        else:
            self.body.pack_forget()

    def toggle(self, _e=None):
        self.expanded = not self.expanded
        self._apply()
        self.app.settings.setdefault("sections", {})[self.key] = self.expanded

    def set_summary(self, text, limit=38):
        if len(text) > limit:
            text = text[:limit - 1] + "…"
        self.summary.configure(text=text)


# --------------------------------------------------------------------------
#  Окно программы
# --------------------------------------------------------------------------

NUM_HINTS = {1: ["например, номер билета"],
             2: ["например, ряд", "например, место"],
             3: ["например, сектор", "например, ряд", "например, место"]}


class App(tk.Tk):
    UNDO_LIMIT = 150

    def __init__(self):
        super().__init__()
        self.cfg_dir = config_dir()
        self.settings = self._load_settings()
        set_lang(self.settings.get("lang", "ru"))
        self.title(tr(APP_TITLE))
        try:
            self._icon = tk.PhotoImage(data=ICON_PNG)
            self.iconphoto(True, self._icon)
        except tk.TclError:
            pass
        self.dnd = False
        if TkinterDnD is not None:
            try:
                TkinterDnD._require(self)
                self.dnd = True
            except Exception:
                self.dnd = False
        self.scale = max(1.0, self.winfo_fpixels("1i") / 96.0)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        default_geo = f"{min(self.px(1380), sw - 40)}x{min(self.px(900), sh - 80)}"
        self.geometry(self.settings.get("geometry", default_geo))
        self.minsize(min(self.px(1120), max(900, sw - 60)), min(self.px(620), max(560, sh - 80)))

        self.tpl_doc = None
        self.tpl_path = None
        self.tpl_src = None          # байты PDF, если макет — картинка
        self._img_dpi = 0
        self._img_size = (0, 0)
        self.page_w = self.page_h = 0.0      # размер видимой страницы, пт
        self.faces = []
        self.families = []
        self._fam_index = {}
        self.family_view = []
        self.family = None
        self.names = []
        self.fields = [default_field() for _ in range(MAX_NUMBERS)]
        self._editing = (0, "text")          # какое поле (или ALL) / какая часть сейчас в редакторе
        self._all_base = None
        self._loading = False
        self.busy = False
        self.picking = False
        self._fonts_loading = True
        self._zoom = 1.0
        self._jobs = {}
        self._base_img = None
        self._ovl_imgs = []
        self._base_pix = None
        self._cmyk_pix = None
        self._anchor_px = {}
        self._fbox = {}
        self._drag = None
        self._hover = None
        self._badge_cache = {}
        self._hex_lock = False
        self._history, self._future = [], []
        self._last_snap = None
        self._packinfo = {}
        self._combos = {}

        self._make_fonts()
        s16, s17 = self.px(16), self.px(17)
        self.icons = {
            "down": make_icon("down", s16, C_ACCENT), "right": make_icon("right", s16, C_ACCENT),
            "close": make_icon("close", s16, C_MUTED), "undo": make_icon("undo", s16, C_ACCENT_DARK),
            "redo": make_icon("redo", s16, C_ACCENT_DARK), "undo_off": make_icon("undo", s16, "#A9B8C5"),
            "redo_off": make_icon("redo", s16, "#A9B8C5"),
            "cb_off": make_icon("box", s17, "#9DB7CB", "#FFFFFF"),
            "cb_hover": make_icon("box", s17, C_ACCENT, "#EEF7FD"),
            "cb_on": make_icon("box_on", s17, C_ACCENT),
            "rb_off": make_icon("ring", s17, "#9DB7CB", "#FFFFFF"),
            "rb_hover": make_icon("ring", s17, C_ACCENT, "#EEF7FD"),
            "rb_on": make_icon("ring_on", s17, C_ACCENT),
            "heart": make_icon("heart", s16, "#FFFFFF"),
        }
        self._make_vars()
        self._setup_theme()
        self._build_ui()
        self._bind_global()
        self._apply_settings()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._start_font_scan()

    def px(self, n):
        return int(round(n * self.scale))

    # ---------------- шрифты интерфейса и тема ----------------

    def _make_fonts(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(size=10)
            except tk.TclError:
                pass
        base = tkfont.nametofont("TkDefaultFont")
        names = ("f_bold", "f_title", "f_small", "f_small_bold", "f_go", "f_empty", "f_sub",
                 "f_badge", "f_tb_b", "f_tb_i", "f_tb_u", "f_big", "f_total", "f_link")
        for n in names:
            setattr(self, n, base.copy())
        self.f_bold.configure(weight="bold")
        self.f_title.configure(size=11, weight="bold")
        self.f_small.configure(size=9)
        self.f_small_bold.configure(size=9, weight="bold")
        self.f_go.configure(size=11, weight="bold")
        self.f_empty.configure(size=14, weight="bold")
        self.f_sub.configure(size=11)
        self.f_badge.configure(size=10, weight="bold")
        self.f_tb_b.configure(size=11, weight="bold")
        self.f_tb_i.configure(size=11, slant="italic")
        self.f_tb_u.configure(size=11, underline=1)
        self.f_big.configure(size=12, weight="bold")
        self.f_total.configure(size=11, weight="bold")
        self.f_link.configure(size=9, weight="bold", underline=1)

    def _setup_theme(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        self.configure(bg=C_BG)
        px = self.px
        st.configure(".", background=C_CARD, foreground=C_TEXT, bordercolor=C_BORDER,
                     lightcolor=C_CARD, darkcolor=C_BORDER, troughcolor=C_BG, focuscolor=C_ACCENT,
                     selectbackground=C_ACCENT, selectforeground="white", fieldbackground="white",
                     insertcolor=C_TEXT, font=tkfont.nametofont("TkDefaultFont"))
        st.configure("TFrame", background=C_CARD)
        st.configure("Panel.TFrame", background=C_BG)
        st.configure("TLabel", background=C_CARD, foreground=C_TEXT)
        st.configure("Panel.TLabel", background=C_BG, foreground=C_MUTED)
        st.configure("Hint.TLabel", foreground=C_MUTED)
        st.configure("Bold.TLabel", font=self.f_bold)
        st.configure("Total.TLabel", font=self.f_total, foreground=C_ACCENT_DARK)
        st.configure("TButton", background=C_ACCENT_LIGHT, foreground=C_ACCENT_DARK,
                     bordercolor=C_ACCENT_LIGHT, lightcolor=C_ACCENT_LIGHT, darkcolor=C_ACCENT_LIGHT,
                     padding=(px(10), px(4)), relief="flat", focusthickness=0)
        st.map("TButton",
               background=[("disabled", "#EEF2F5"), ("pressed", "#B9DBF0"), ("active", "#C7E3F4")],
               bordercolor=[("pressed", "#B9DBF0"), ("active", "#C7E3F4")],
               lightcolor=[("pressed", "#B9DBF0"), ("active", "#C7E3F4")],
               darkcolor=[("pressed", "#B9DBF0"), ("active", "#C7E3F4")],
               foreground=[("disabled", "#9AAAB8")])
        acc_states = [("disabled", "#9CCDE9"), ("pressed", "#006FA6"), ("active", C_ACCENT_DARK)]
        st.configure("Accent.TButton", background=C_ACCENT, foreground="white", bordercolor=C_ACCENT,
                     lightcolor=C_ACCENT, darkcolor=C_ACCENT, font=self.f_go, padding=(px(12), px(9)))
        st.map("Accent.TButton", background=acc_states, bordercolor=acc_states, lightcolor=acc_states,
               darkcolor=acc_states, foreground=[("disabled", "white")])
        st.configure("Pick.TButton", background=C_ACCENT, foreground="white", bordercolor=C_ACCENT,
                     lightcolor=C_ACCENT, darkcolor=C_ACCENT, padding=(px(10), px(4)))
        st.map("Pick.TButton", background=[("active", C_ACCENT_DARK)])
        don = [("pressed", "#CC5200"), ("active", C_DONATE_DARK)]
        st.configure("Donate.TButton", background=C_DONATE, foreground="white", bordercolor=C_DONATE,
                     lightcolor=C_DONATE, darkcolor=C_DONATE, font=self.f_bold, padding=(px(12), px(4)))
        st.map("Donate.TButton", background=don, bordercolor=don, lightcolor=don, darkcolor=don)
        st.configure("Tool.TButton", background=C_BG, bordercolor=C_BG, lightcolor=C_BG, darkcolor=C_BG,
                     foreground=C_ACCENT_DARK, padding=(px(8), px(3)))
        st.map("Tool.TButton", background=[("disabled", C_BG), ("active", C_ACCENT_LIGHT)],
               foreground=[("disabled", "#A9B8C5")])
        st.configure("Toolbutton", background=C_CARD, bordercolor=C_BORDER, lightcolor=C_CARD,
                     darkcolor=C_CARD, padding=(px(8), px(2)), relief="solid", anchor="center")
        st.map("Toolbutton",
               background=[("selected", C_ACCENT), ("active", C_ACCENT_LIGHT)],
               foreground=[("selected", "white")],
               bordercolor=[("selected", C_ACCENT)],
               lightcolor=[("selected", C_ACCENT)], darkcolor=[("selected", C_ACCENT)])
        st.configure("B.Toolbutton", font=self.f_tb_b)
        st.configure("I.Toolbutton", font=self.f_tb_i)
        st.configure("U.Toolbutton", font=self.f_tb_u)
        st.configure("Seg.Toolbutton", background="#EEF5FB", lightcolor="#EEF5FB", darkcolor="#EEF5FB",
                     padding=(px(10), px(5)), font=self.f_bold, foreground=C_ACCENT_DARK)
        st.configure("Lang.Toolbutton", background=C_BG, lightcolor=C_BG, darkcolor=C_BG,
                     padding=(px(8), px(2)), font=self.f_small_bold, foreground=C_ACCENT_DARK)
        st.configure("Bar.Toolbutton", background=C_ACCENT_DARK, foreground="white", bordercolor=C_ACCENT_DARK,
                     lightcolor=C_ACCENT_DARK, darkcolor=C_ACCENT_DARK, padding=(px(10), px(4)), font=self.f_bold)
        st.map("Bar.Toolbutton", background=[("selected", "white"), ("active", "#1AA0DD")],
               foreground=[("selected", C_ACCENT_DARK)], bordercolor=[("selected", "white")],
               lightcolor=[("selected", "white")], darkcolor=[("selected", "white")])
        for w in ("TEntry", "TSpinbox", "TCombobox"):
            st.configure(w, fieldbackground="white", bordercolor=C_BORDER, lightcolor="white",
                         darkcolor="white", padding=(px(4), px(2)), arrowcolor=C_ACCENT, background=C_CARD)
            st.map(w, bordercolor=[("focus", C_ACCENT)], lightcolor=[("focus", C_ACCENT)],
                   fieldbackground=[("readonly", "#F7FAFC")])
        st.configure("TSpinbox", arrowsize=px(11))
        st.configure("TCombobox", arrowsize=px(13))
        self.option_add("*TCombobox*Listbox.selectBackground", C_ACCENT)
        self.option_add("*TCombobox*Listbox.font", tkfont.nametofont("TkDefaultFont"))
        ic = self.icons
        gap = self.px(17) + self.px(7)
        try:   # галочки и переключатели — рисованные значки (масштабируются вместе с экраном)
            st.element_create("P6.check", "image", ic["cb_off"], ("selected", ic["cb_on"]),
                              ("active", ic["cb_hover"]), width=gap, sticky="w")
            st.element_create("P6.radio", "image", ic["rb_off"], ("selected", ic["rb_on"]),
                              ("active", ic["rb_hover"]), width=gap, sticky="w")
            for cls, el in (("Checkbutton", "P6.check"), ("Radiobutton", "P6.radio")):
                st.layout("T" + cls, [(f"{cls}.padding", {"sticky": "nswe", "children": [
                    (el, {"side": "left", "sticky": ""}),
                    (f"{cls}.focus", {"side": "left", "sticky": "w",
                                      "children": [(f"{cls}.label", {"sticky": "nswe"})]})]})])
        except tk.TclError:
            for w in ("TCheckbutton", "TRadiobutton"):
                st.configure(w, indicatorsize=px(15))
        for w in ("TCheckbutton", "TRadiobutton"):
            st.configure(w, background=C_CARD, padding=(0, px(2)))
            st.map(w, background=[("active", C_CARD)])
        for w in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            st.configure(w, background="#C9D8E4", troughcolor=C_BG, bordercolor=C_BG, arrowsize=px(14),
                         lightcolor="#C9D8E4", darkcolor="#C9D8E4", arrowcolor=C_MUTED, gripcount=0)
            st.map(w, background=[("active", "#AFC6D8")])
        st.configure("Accent.Horizontal.TProgressbar", troughcolor=C_ACCENT_LIGHT, background=C_ACCENT,
                     bordercolor=C_ACCENT_LIGHT, lightcolor=C_ACCENT, darkcolor=C_ACCENT, thickness=px(6))
        st.configure("TPanedwindow", background=C_BG)
        st.configure("Sash", sashthickness=px(6), gripcount=0, background=C_BG, bordercolor=C_BG,
                     lightcolor=C_BG, darkcolor=C_BG)

    def _style_tk(self, w):
        """Оформление классических tk-виджетов (список, текст) в общем стиле."""
        w.configure(bg="white", fg=C_TEXT, relief="flat", borderwidth=0, highlightthickness=1,
                    highlightbackground=C_BORDER, highlightcolor=C_ACCENT, selectbackground=C_ACCENT,
                    selectforeground="white")
        if isinstance(w, tk.Text):
            w.configure(insertbackground=C_TEXT, padx=self.px(6), pady=self.px(4))

    # ---------------- переменные ----------------

    def _make_vars(self):
        S, B = tk.StringVar, tk.BooleanVar
        self.v_tpl_name = S(value=tr("макет не выбран"))
        self.v_dpi = S(value="300")
        self.v_page = S(value="1")
        self.v_page_info = S(value="")
        self.v_prev_idx = S(value="1")
        self.v_font_filter = S()
        self.v_cyr_only = B(value=True)
        self.v_font_name = S(value="—")
        self.v_size = S(value="14")
        self.v_bold, self.v_italic, self.v_underline = B(value=False), B(value=False), B(value=False)
        self.v_mode = S(value="CMYK")
        self.v_c, self.v_m, self.v_y, self.v_k = S(value="0"), S(value="0"), S(value="0"), S(value="100")
        self.v_r, self.v_g, self.v_b = S(value="0"), S(value="0"), S(value="0")
        self.v_hex = S(value="#000000")
        self.v_x, self.v_ypos = S(value="0"), S(value="0")
        self.v_align = S(value="center")
        self.v_maxw = S(value="0")
        self.v_br = [B(value=False), B(value=False), B(value=False)]
        self.v_wrap_needed = B(value=True)
        self.v_leading = S(value="120")
        self.v_num_sep = B(value=False)
        self.v_zoom = S(value="fit")
        self.v_info = S(value="")
        self.v_status = S(value="")
        self.v_lang = S(value=LANG)
        self.v_data_mode = S(value="fio")
        self.v_field = S(value="all")
        self.v_part = S(value="text")
        self.v_n = {"n_count": S(value="1"), "n_order": S(value="rows"), "n_layout": S(value="same")}
        defaults = {1: ("1", "100"), 2: ("1", "10"), 3: ("1", "10")}
        for k in range(1, MAX_NUMBERS + 1):
            frm, to = defaults[k]
            self.v_n.update({f"n{k}_from": S(value=frm), f"n{k}_to": S(value=to), f"n{k}_step": S(value="1"),
                             f"n{k}_digits": S(value="0"), f"n{k}_before": S(value=""), f"n{k}_after": S(value="")})

        # выпадающие списки: в переменной — код, на экране — подпись на текущем языке
        self._make_combo("align", self.v_align, ALIGN_OPTS)
        self._make_combo("order", self.v_n["n_order"], ORDER_OPTS)
        self._make_combo("zoom", self.v_zoom, ZOOM_OPTS)

        self._num_defaults = {k: v.get() for k, v in self.v_n.items()}
        self.v_dpi.trace_add("write", lambda *_: self.schedule("dpi", 400))

        # редактор: значения текущего поля / текущей части
        self._evars = {
            "size": self.v_size, "bold": self.v_bold, "italic": self.v_italic, "underline": self.v_underline,
            "mode": self.v_mode, "c": self.v_c, "m": self.v_m, "y": self.v_y, "k": self.v_k,
            "r": self.v_r, "g": self.v_g, "b": self.v_b,
            "x": self.v_x, "ypos": self.v_ypos, "align": self.v_align, "maxw": self.v_maxw,
            "br1": self.v_br[0], "br2": self.v_br[1], "br3": self.v_br[2],
            "wrap": self.v_wrap_needed, "leading": self.v_leading, "num_sep": self.v_num_sep,
        }
        # общие настройки, которые тоже отменяются по Ctrl+Z
        self._gvars = dict(self.v_n, page=self.v_page, data_mode=self.v_data_mode)

        redraw = [self.v_prev_idx, self.v_size, self.v_mode, self.v_c, self.v_m, self.v_y,
                  self.v_k, self.v_r, self.v_g, self.v_b, self.v_x, self.v_ypos,
                  self.v_align, self.v_maxw, self.v_wrap_needed, self.v_leading,
                  self.v_underline] + self.v_br + list(self.v_n.values())
        for v in redraw:
            v.trace_add("write", lambda *_: self.schedule("overlay", 40))
        for v in (self.v_bold, self.v_italic):
            v.trace_add("write", lambda *_: self._on_style_toggle())
        for v in (self.v_r, self.v_g, self.v_b):
            v.trace_add("write", lambda *_: self._rgb_to_hex())
        for v in (self.v_c, self.v_m, self.v_y, self.v_k, self.v_r, self.v_g, self.v_b, self.v_mode):
            v.trace_add("write", lambda *_: self._update_swatch())
        self.v_mode.trace_add("write", lambda *_: self._on_mode_change())
        self.v_page.trace_add("write", lambda *_: self.schedule("page", 200))
        self.v_zoom.trace_add("write", lambda *_: self.schedule("base"))
        self.v_font_filter.trace_add("write", lambda *_: self._refill_fonts())
        self.v_cyr_only.trace_add("write", lambda *_: self._refill_fonts())
        for v in list(self._evars.values()) + list(self._gvars.values()):
            v.trace_add("write", lambda *_: self._note_change())
        for v in (self.v_data_mode, self.v_n["n_count"], self.v_n["n_layout"], self.v_field, self.v_part):
            v.trace_add("write", lambda *_: self._retarget())
        self.v_num_sep.trace_add("write", lambda *_: self._on_num_sep())
        self.v_lang.trace_add("write", lambda *_: self.after_idle(lambda: self.set_language(self.v_lang.get())))

    def _make_combo(self, name, code_var, opts):
        disp = tk.StringVar()
        self._combos[name] = (code_var, disp, opts)
        code_var.trace_add("write", lambda *_: self._sync_combo(name))
        disp.trace_add("write", lambda *_: self._combo_picked(name))
        self._sync_combo(name)

    def _sync_combo(self, name):
        code, disp, opts = self._combos[name]
        labels = {c: tr(label) for c, label in opts}
        label = labels.get(code.get(), next(iter(labels.values())))
        if disp.get() != label:
            disp.set(label)

    def _combo_picked(self, name):
        code, disp, opts = self._combos[name]
        for c, label in opts:
            if tr(label) == disp.get():
                if code.get() != c:
                    code.set(c)
                return

    def _combobox(self, parent, name, width):
        _code, disp, opts = self._combos[name]
        return ttk.Combobox(parent, textvariable=disp, values=[tr(label) for _, label in opts],
                            state="readonly", width=width)

    # ---------------- интерфейс ----------------

    def _show(self, w, visible, **kw):
        """Показать/скрыть виджет, запоминая, как он был размещён."""
        if kw:
            self._packinfo[w] = kw
        if visible and not w.winfo_manager():
            w.pack(**self._packinfo.get(w, {}))
        elif not visible and w.winfo_manager():
            w.pack_forget()

    def _bind_global(self):
        """Привязки клавиш и колеса — один раз на всё время работы."""
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._on_wheel, add="+")
            self.bind_class("TCombobox", seq, lambda e: (self._on_wheel(e), "break")[1])
        self.bind_all("<Control-o>", lambda e: self.open_template())
        self.bind_all("<Control-KeyPress>", self._on_ctrl_key, add="+")
        self.bind("<Escape>", lambda e: self._stop_pick())

    def _build_ui(self):
        px = self.px
        self._part_frames = []
        paned = self.paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # --- левая панель ---
        outer = tk.Frame(paned, bg=C_BG)
        paned.add(outer, weight=0)

        bottom = tk.Frame(outer, bg=C_BG)
        bottom.pack(side="bottom", fill="x", padx=(px(12), px(10)), pady=(px(6), px(10)))
        self.btn_go = ttk.Button(bottom, text=tr("Сформировать PDF…"), style="Accent.TButton",
                                 command=self.generate)
        self.btn_go.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate", style="Accent.Horizontal.TProgressbar")
        self.lbl_status = ttk.Label(bottom, textvariable=self.v_status, style="Panel.TLabel", wraplength=px(420))
        self.lbl_status.pack(anchor="w", pady=(px(6), 0))


        # полоса выбора номера (видна, когда номера размещены отдельно)
        self.field_bar = tk.Frame(outer, bg=C_ACCENT)
        tk.Label(self.field_bar, text=tr("Настраиваются:"), bg=C_ACCENT, fg="white",
                 font=self.f_bold).pack(anchor="w", padx=px(12), pady=(px(8), px(4)))
        row = tk.Frame(self.field_bar, bg=C_ACCENT)
        row.pack(fill="x", padx=px(12))
        self.fb_btns = []
        for text, val in ((tr("Все номера"), "all"), (tr("Номер 1"), "0"), (tr("Номер 2"), "1"),
                          (tr("Номер 3"), "2")):
            b = ttk.Radiobutton(row, text=text, value=val, variable=self.v_field, style="Bar.Toolbutton")
            b.pack(side="left", padx=(0, px(4)))
            self.fb_btns.append(b)
        self._packinfo[self.fb_btns[3]] = {"side": "left", "padx": (0, px(4)), "after": self.fb_btns[2]}
        tk.Label(self.field_bar, text=tr("«Все номера» — изменения применяются ко всем номерам сразу, а "
                                         "перемещение сдвигает их вместе. Чтобы настроить один номер, выберите "
                                         "его здесь или дважды щёлкните по нему на макете."),
                 bg=C_ACCENT, fg=C_ACCENT_LIGHT, font=self.f_small, wraplength=px(440),
                 justify="left").pack(anchor="w", padx=px(12), pady=(px(6), px(8)))

        self.scroll_area = tk.Frame(outer, bg=C_BG)
        self.scroll_area.pack(fill="both", expand=True)
        self._packinfo[self.field_bar] = {"side": "top", "fill": "x", "padx": (px(12), px(10)),
                                          "pady": (px(10), 0), "before": self.scroll_area}
        self.lcanvas = tk.Canvas(self.scroll_area, highlightthickness=0, width=px(500), bg=C_BG)
        lsb = ttk.Scrollbar(self.scroll_area, orient="vertical", command=self.lcanvas.yview)
        self.lcanvas.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y", pady=px(10))
        self.lcanvas.pack(side="left", fill="both", expand=True)
        panel = tk.Frame(self.lcanvas, bg=C_BG, padx=px(12), pady=px(10))
        win = self.lcanvas.create_window(0, 0, window=panel, anchor="nw")
        panel.bind("<Configure>", lambda e: self.lcanvas.configure(scrollregion=self.lcanvas.bbox("all")))
        self.lcanvas.bind("<Configure>", lambda e: self.lcanvas.itemconfigure(win, width=e.width))

        self.sections = {}
        order = [("tpl", tr("Макет PDF"), self._build_template_box),
                 ("pos", tr("Место на макете"), self._build_position_box),
                 ("data", tr("Данные"), self._build_data_box),
                 ("font", tr("Шрифт"), self._build_font_box),
                 ("color", tr("Цвет"), self._build_color_box),
                 ("wrap", tr("Перенос на новую строку"), self._build_wrap_box)]
        for num_, (key, title, builder) in enumerate(order, start=1):
            sec = Section(self, panel, num_, title, key)
            sec.pack(fill="x", pady=(0, px(10)))
            self.sections[key] = sec
            builder(sec.body)
        self.sec_swatch = self.sections["color"].add_swatch()
        self._update_swatch()

        self.update_idletasks()   # ширина панели — по содержимому, чтобы всё помещалось
        width = max(px(510), panel.winfo_reqwidth())
        self.lcanvas.configure(width=width)
        self._pane_w = getattr(self, "_pane_w", 0) or width + px(26)
        self.after_idle(self._apply_sash)

        # --- правая панель: предпросмотр ---
        right = tk.Frame(paned, bg=C_BG)
        paned.add(right, weight=1)
        bar = tk.Frame(right, bg=C_BG)
        bar.pack(fill="x", padx=(px(6), px(12)), pady=(px(10), px(6)))
        ic = self.icons
        self.btn_undo = ttk.Button(bar, style="Tool.TButton", command=self.undo,
                                   image=(ic["undo"], "disabled", ic["undo_off"]))
        self.btn_undo.pack(side="left")
        self._tooltip(self.btn_undo, tr("Отменить") + "  (Ctrl+Z)")
        self.btn_redo = ttk.Button(bar, style="Tool.TButton", command=self.redo,
                                   image=(ic["redo"], "disabled", ic["redo_off"]))
        self.btn_redo.pack(side="left", padx=(px(2), 0))
        self._tooltip(self.btn_redo, tr("Вернуть") + "  (Ctrl+Y)")
        ttk.Button(bar, text=tr("Сбросить настройки"), style="Tool.TButton",
                   command=self.reset_settings).pack(side="left", padx=(px(14), 0))
        ttk.Button(bar, text=tr("Поддержать разработку"), style="Donate.TButton", command=self.show_donate,
                   image=ic["heart"], compound="left").pack(side="left", padx=(px(10), 0))
        row2 = tk.Frame(right, bg=C_BG)
        row2.pack(fill="x", padx=(px(6), px(12)), pady=(0, px(8)))
        for text, val in (("EN", "en"), ("RU", "ru")):
            ttk.Radiobutton(bar, text=text, value=val, variable=self.v_lang,
                            style="Lang.Toolbutton").pack(side="right")
        ttk.Frame(bar, width=px(12), style="Panel.TFrame").pack(side="right")
        self._combobox(row2, "zoom", 15).pack(side="right")
        ttk.Label(row2, text=tr("Масштаб:"), style="Panel.TLabel").pack(side="right", padx=(px(12), px(6)))

        promo = tk.Frame(row2, bg=C_ACCENT_LIGHT, cursor="hand2", highlightthickness=1,
                         highlightbackground=C_HEAD_BORDER)
        promo.pack(side="left", fill="x", expand=True)
        self.lbl_promo = tk.Label(promo, text=tr("Напечатать быстро и дешево можно тут"), bg=C_ACCENT_LIGHT,
                                  fg=C_TEXT, font=self.f_bold, cursor="hand2")
        self.lbl_promo.pack(side="left", padx=(px(12), px(6)), pady=px(7))
        link = tk.Label(promo, text="— print600.ru", bg=C_ACCENT_LIGHT, fg=C_ACCENT_DARK,
                        font=self.f_link, cursor="hand2")
        link.pack(side="left", pady=px(7))
        self._promo = [promo, self.lbl_promo, link]
        for w in self._promo:
            w.bind("<Button-1>", lambda e: webbrowser.open(PROMO_URL))

        info = tk.Frame(right, bg=C_BG)
        info.pack(side="bottom", fill="x", padx=(px(6), px(12)), pady=(px(4), px(8)))
        ttk.Label(info, textvariable=self.v_info, style="Panel.TLabel").pack(side="left", fill="x")

        wrap = tk.Frame(right, bg=C_BG)
        wrap.pack(fill="both", expand=True, padx=(px(6), px(12)))
        self.pcanvas = tk.Canvas(wrap, bg=C_CANVAS, highlightthickness=1, highlightbackground=C_BORDER,
                                 highlightcolor=C_BORDER, cursor="hand2")
        xs = ttk.Scrollbar(wrap, orient="horizontal", command=self.pcanvas.xview)
        ys = ttk.Scrollbar(wrap, orient="vertical", command=self.pcanvas.yview)

        def xscroll(*a):
            xs.set(*a)
            self.after_idle(self._place_badge)

        def yscroll(*a):
            ys.set(*a)
            self.after_idle(self._place_badge)

        self.pcanvas.configure(xscrollcommand=xscroll, yscrollcommand=yscroll)
        self.pcanvas.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.pcanvas.bind("<Configure>", self._on_canvas_resize)
        self.pcanvas.bind("<Button-1>", self._on_canvas_press)
        self.pcanvas.bind("<Double-Button-1>", self._on_canvas_double)
        self.pcanvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.pcanvas.bind("<ButtonRelease-1>", lambda e: setattr(self, "_drag", None))
        self.pcanvas.bind("<Motion>", self._on_canvas_motion)
        self.pcanvas.bind("<Leave>", lambda e: self._set_hover(None))
        self._build_pick_panel()

        if self.dnd:
            for w in (self.pcanvas, self.txt, self.lb_fonts):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<DropEnter>>", self._on_drag_enter)
                    w.dnd_bind("<<DropPosition>>", lambda e: "copy")
                    w.dnd_bind("<<DropLeave>>", self._on_drag_leave)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass
        self._update_undo_buttons()
        self._update_mode_ui()
        self._update_summaries()
        self.after(50, self._draw_empty)

    def _apply_sash(self):
        """Ширина левой панели (в том числе после перестроения окна при смене языка)."""
        try:
            self.paned.sashpos(0, self._pane_w)
        except tk.TclError:
            pass

    def _tooltip(self, widget, text):
        """Подсказка при наведении — значки без подписей должны оставаться понятными."""
        tip = {"win": None}

        def show(_e=None):
            if tip["win"] or not widget.winfo_exists():
                return
            win = tip["win"] = tk.Toplevel(self)
            win.overrideredirect(True)
            win.configure(bg=C_TEXT)
            tk.Label(win, text=text, bg=C_TEXT, fg="white", font=self.f_small,
                     padx=self.px(8), pady=self.px(4)).pack()
            win.geometry(f"+{widget.winfo_rootx()}+{widget.winfo_rooty() + widget.winfo_height() + self.px(4)}")

        def hide(_e=None):
            if tip["win"]:
                tip["win"].destroy()
                tip["win"] = None

        widget.bind("<Enter>", show, add="+")
        widget.bind("<Leave>", hide, add="+")
        widget.bind("<Button-1>", hide, add="+")

    def _spin(self, parent, var, lo, hi, step, width=7, fmt_="%.1f"):
        return ttk.Spinbox(parent, textvariable=var, from_=lo, to=hi, increment=step,
                           width=width, format=fmt_)

    def _row(self, parent, top=8):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(self.px(top), 0))
        return row

    def _hint(self, parent, text, top=8):
        lbl = ttk.Label(parent, text=text, style="Hint.TLabel", wraplength=self.px(430), justify="left")
        lbl.pack(anchor="w", pady=(self.px(top), 0))
        return lbl

    def _part_selector(self, box):
        """«Слова / Номер» — для нумерации, где номер оформлен отдельно."""
        px = self.px
        frame = ttk.Frame(box)
        ttk.Checkbutton(frame, text=tr("Номер — своим шрифтом и цветом"),
                        variable=self.v_num_sep).pack(anchor="w")
        seg = ttk.Frame(frame)
        for text, val in ((tr("Слова (текст до и после)"), "text"), (tr("Номер"), "num")):
            ttk.Radiobutton(seg, text=text, value=val, variable=self.v_part,
                            style="Seg.Toolbutton").pack(side="left", fill="x", expand=True)
        self._packinfo[seg] = {"fill": "x", "pady": (px(6), 0)}
        tk.Frame(frame, bg=C_BORDER, height=1).pack(side="bottom", fill="x", pady=(px(10), 0))
        self._part_frames.append((frame, seg))
        return frame

    def _build_template_box(self, box):
        row = self._row(box, 0)
        ttk.Entry(row, textvariable=self.v_tpl_name, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(row, text=tr("Открыть…"), command=self.open_template).pack(side="left", padx=(self.px(6), 0))
        self.row_page = ttk.Frame(box)   # показывается, только если в макете несколько страниц
        ttk.Label(self.row_page, text=tr("Страница макета:")).pack(side="left")
        self.sp_page = ttk.Spinbox(self.row_page, textvariable=self.v_page, from_=1, to=1, increment=1,
                                   width=5, format="%.0f")
        self.sp_page.pack(side="left", padx=self.px(6))
        ttk.Label(self.row_page, textvariable=self.v_page_info, style="Hint.TLabel").pack(side="left")
        self.row_dpi = ttk.Frame(box)    # показывается, если макет — картинка
        ttk.Label(self.row_dpi, text=tr("Разрешение, dpi")).pack(side="left")
        self._spin(self.row_dpi, self.v_dpi, 10, 2400, 10, width=6, fmt_="%.0f").pack(side="left", padx=self.px(6))
        self.lbl_img = ttk.Label(self.row_dpi, text="", style="Hint.TLabel")
        self.lbl_img.pack(side="left")
        self.tpl_hint = self._hint(box, tr("Макет можно просто перетащить мышью в окно предпросмотра. "
                                           "Поддерживаются PDF, JPG, PNG и TIFF; у картинки размер макета "
                                           "считается по её разрешению."))

    def _build_position_box(self, box):
        px = self.px
        row = self._row(box, 0)
        ttk.Label(row, text=tr("X, мм")).pack(side="left")
        self._spin(row, self.v_x, -2000, 5000, 0.5).pack(side="left", padx=(px(6), px(16)))
        ttk.Label(row, text=tr("Y, мм")).pack(side="left")
        self._spin(row, self.v_ypos, -2000, 5000, 0.5).pack(side="left", padx=px(6))
        row = self._row(box)
        ttk.Label(row, text=tr("Выравнивание")).pack(side="left")
        self._combobox(row, "align", 16).pack(side="left", padx=px(6))
        row = self._row(box)
        ttk.Label(row, text=tr("Макс. ширина, мм")).pack(side="left")
        self._spin(row, self.v_maxw, 0, 5000, 1).pack(side="left", padx=px(6))
        ttk.Button(row, text=tr("= ширина макета"), command=self.maxw_to_page).pack(side="left")
        row = self._row(box)
        ttk.Button(row, text=tr("Центр по горизонтали"), command=self.center_h).pack(side="left")
        ttk.Button(row, text=tr("Центр страницы"), command=self.center_page).pack(side="left", padx=px(6))
        self._hint(box, tr("Щёлкните по макету — текст встанет в эту точку; потяните за текст — "
                           "он сдвинется. X — точка выравнивания, Y — базовая линия первой строки "
                           "(от левого верхнего угла). Если текст шире макс. ширины, кегль "
                           "уменьшается автоматически (0 — без ограничения)."))

    def _build_data_box(self, box):
        px = self.px
        seg = ttk.Frame(box)
        seg.pack(fill="x")
        for text, val in ((tr("Список ФИО"), "fio"), (tr("Нумерация"), "num")):
            ttk.Radiobutton(seg, text=text, value=val, variable=self.v_data_mode,
                            style="Seg.Toolbutton").pack(side="left", fill="x", expand=True)

        # --- список ФИО ---
        self.f_fio = ttk.Frame(box)
        self._hint(self.f_fio, tr("По одной записи в строке. Ctrl+V, правая кнопка мыши или перетаскивание "
                                  "файла — можно копировать столбец прямо из Excel."), top=10)
        tf = ttk.Frame(self.f_fio)
        tf.pack(fill="both", pady=(px(6), px(6)))
        self.txt = tk.Text(tf, height=8, width=20, wrap="none", undo=False, font="TkDefaultFont")
        self._style_tk(self.txt)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.txt.bind("<<Modified>>", self._on_text_modified)
        self.menu_txt = tk.Menu(self, tearoff=0, bg="white", fg=C_TEXT, activebackground=C_ACCENT_LIGHT,
                                activeforeground=C_TEXT, relief="flat", bd=1)
        for item in ((tr("Вставить"), "Ctrl+V", lambda: self._txt_event("<<Paste>>")),
                                  (tr("Копировать"), "Ctrl+C", lambda: self._txt_event("<<Copy>>")),
                                  (tr("Вырезать"), "Ctrl+X", lambda: self._txt_event("<<Cut>>")),
                                  None,
                                  (tr("Выделить всё"), "Ctrl+A", lambda: self._select_all(self.txt)),
                                  (tr("Отменить"), "Ctrl+Z", self.undo),
                                  (tr("Очистить список"), "", lambda: self.txt.delete("1.0", "end"))):
            if item is None:
                self.menu_txt.add_separator()
            else:
                self.menu_txt.add_command(label=item[0], accelerator=item[1], command=item[2])
        self.txt.bind("<Button-3>", lambda e: (self.txt.focus_set(),
                                               self.menu_txt.tk_popup(e.x_root, e.y_root)))
        row = ttk.Frame(self.f_fio)
        row.pack(fill="x")
        ttk.Button(row, text=tr("Загрузить из файла…"), command=self.load_names_file).pack(side="left")
        ttk.Button(row, text=tr("Вставить"), command=lambda: self._txt_event("<<Paste>>")).pack(side="left", padx=px(6))
        ttk.Button(row, text=tr("Очистить"), command=lambda: self.txt.delete("1.0", "end")).pack(side="left")

        # --- нумерация ---
        self.f_num = ttk.Frame(box)
        seg = self._row(self.f_num, 12)
        for text, val in ((tr("1 номер"), "1"), (tr("2 номера"), "2"), (tr("3 номера"), "3")):
            ttk.Radiobutton(seg, text=text, value=val, variable=self.v_n["n_count"],
                            style="Seg.Toolbutton").pack(side="left", fill="x", expand=True)
        self._hint(self.f_num, tr("Например: номер билета · ряд и место · сектор, ряд и место."), top=4)
        self.f_counters, self.lbl_chint = [], []
        for k in range(1, MAX_NUMBERS + 1):
            fr = ttk.Frame(self.f_num)
            self._build_counter(fr, k)
            self.f_counters.append(fr)
        self.f_multi = ttk.Frame(self.f_num)
        row = self._row(self.f_multi, 12)
        ttk.Label(row, text=tr("Порядок")).pack(side="left")
        self._combobox(row, "order", 36).pack(side="left", padx=px(6))
        row = self._row(self.f_multi, 8)
        ttk.Label(row, text=tr("Размещение")).pack(side="left", padx=(0, px(8)))
        ttk.Radiobutton(row, text=tr("в одной строке"), value="same",
                        variable=self.v_n["n_layout"]).pack(side="left", padx=(0, px(10)))
        ttk.Radiobutton(row, text=tr("отдельно (своё место и шрифт)"), value="separate",
                        variable=self.v_n["n_layout"]).pack(side="left")
        self.lbl_total = ttk.Label(self.f_num, style="Total.TLabel")
        self.lbl_total.pack(anchor="w", pady=(px(12), 0))
        self.lbl_example = ttk.Label(self.f_num, style="Hint.TLabel", wraplength=px(430), justify="left")
        self.lbl_example.pack(anchor="w", pady=(px(2), 0))
        self._hint(self.f_num, tr("Пробелы в «до» и «после» учитываются: «от » + 1 + « шт» = «от 1 шт». "
                                  "Цифр — минимальное число знаков: 3 → 001."))

        # --- общий ряд: какая запись в предпросмотре ---
        self.row_prev = self._row(box, 10)
        ttk.Label(self.row_prev, text=tr("Показать запись №")).pack(side="left")
        self.sp_prev = ttk.Spinbox(self.row_prev, textvariable=self.v_prev_idx, from_=1, to=1, increment=1,
                                   width=7, format="%.0f")
        self.sp_prev.pack(side="left", padx=px(6))
        ttk.Button(self.row_prev, text=tr("Самая длинная"), command=self.show_longest).pack(side="left")
        self._packinfo[self.f_fio] = {"fill": "x", "before": self.row_prev}
        self._packinfo[self.f_num] = {"fill": "x", "before": self.row_prev}

    def _build_counter(self, parent, n):
        px = self.px
        v = lambda k: self.v_n[f"n{n}_{k}"]   # noqa: E731
        head = self._row(parent, 12)
        ttk.Label(head, text=tr("Номер {n}").format(n=n), style="Bold.TLabel").pack(side="left")
        hint = ttk.Label(head, text="", style="Hint.TLabel")
        hint.pack(side="left", padx=px(8))
        self.lbl_chint.append(hint)
        g = self._row(parent, 6)
        for col, (label, key, w, lo, hi) in enumerate(((tr("с"), "from", 7, -999999, 9999999),
                                                       (tr("по"), "to", 7, -999999, 9999999),
                                                       (tr("шаг"), "step", 4, 1, 100000),
                                                       (tr("цифр"), "digits", 3, 0, 12))):
            ttk.Label(g, text=label).grid(row=0, column=col * 2, sticky="w", padx=(0 if col == 0 else px(10), px(4)))
            self._spin(g, v(key), lo, hi, 1, width=w, fmt_="%.0f").grid(row=0, column=col * 2 + 1)
        t = self._row(parent, 6)
        ttk.Label(t, text=tr("текст до")).pack(side="left")
        ttk.Entry(t, textvariable=v("before"), width=11).pack(side="left", fill="x", expand=True, padx=(px(4), px(10)))
        ttk.Label(t, text=tr("после")).pack(side="left")
        ttk.Entry(t, textvariable=v("after"), width=11).pack(side="left", fill="x", expand=True, padx=(px(4), 0))

    def _build_font_box(self, box):
        px = self.px
        first = self._part_selector(box)
        row = self._row(box, 0)
        self._packinfo[first] = {"fill": "x", "pady": (0, px(10)), "before": row}
        ttk.Label(row, text=tr("Поиск")).pack(side="left")
        ttk.Entry(row, textvariable=self.v_font_filter).pack(side="left", fill="x", expand=True, padx=(px(6), 0))
        lf = ttk.Frame(box)
        lf.pack(fill="x", pady=(px(6), px(4)))
        self.lb_fonts = tk.Listbox(lf, height=8, width=20, exportselection=False, activestyle="none")
        self._style_tk(self.lb_fonts)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.lb_fonts.yview)
        self.lb_fonts.configure(yscrollcommand=sb.set)
        self.lb_fonts.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")
        self.lb_fonts.bind("<<ListboxSelect>>", self._on_font_select)
        row = ttk.Frame(box)
        row.pack(fill="x")
        ttk.Checkbutton(row, text=tr("Только с кириллицей"), variable=self.v_cyr_only).pack(side="left")
        ttk.Button(row, text=tr("Шрифт из файла…"), command=self.add_font_file).pack(side="right")
        row = self._row(box, 10)
        ttk.Label(row, text=tr("Размер, pt")).pack(side="left")
        self._spin(row, self.v_size, 2, 999, 0.5).pack(side="left", padx=(px(6), px(16)))
        for text, var, st in ((tr("Ж"), self.v_bold, "B.Toolbutton"), (tr("К"), self.v_italic, "I.Toolbutton"),
                              (tr("Ч"), self.v_underline, "U.Toolbutton")):
            ttk.Checkbutton(row, text=text, variable=var, style=st, width=3).pack(side="left", padx=(0, px(4)))
        row = self._row(box)
        ttk.Label(row, textvariable=self.v_font_name, style="Hint.TLabel",
                  wraplength=px(430), justify="left").pack(side="left")

    def _build_color_box(self, box):
        px = self.px
        first = self._part_selector(box)
        row = self._row(box, 0)
        self._packinfo[first] = {"fill": "x", "pady": (0, px(10)), "before": row}
        ttk.Radiobutton(row, text=tr("CMYK (печать)"), value="CMYK", variable=self.v_mode).pack(side="left")
        ttk.Radiobutton(row, text="RGB", value="RGB", variable=self.v_mode).pack(side="left", padx=px(12))
        self.swatch = tk.Canvas(row, width=px(46), height=px(24), highlightthickness=1,
                                highlightbackground=C_BORDER, bg="#000000")
        self.swatch.pack(side="right")
        self.btn_pick = ttk.Button(row, text=tr("Пипетка"), command=self.toggle_pick)
        self.btn_pick.pack(side="right", padx=px(8))

        self.f_cmyk = ttk.Frame(box)
        for lbl, var in (("C", self.v_c), ("M", self.v_m), ("Y", self.v_y), ("K", self.v_k)):
            ttk.Label(self.f_cmyk, text=lbl, style="Bold.TLabel").pack(side="left")
            self._spin(self.f_cmyk, var, 0, 100, 1, width=4, fmt_="%.0f").pack(side="left", padx=(px(3), px(8)))
        ttk.Label(self.f_cmyk, text="%", style="Hint.TLabel").pack(side="left")
        ttk.Button(self.f_cmyk, text="K100", command=lambda: self._set_cmyk(0, 0, 0, 100)).pack(side="right")

        self.f_rgb = ttk.Frame(box)
        for lbl, var in (("R", self.v_r), ("G", self.v_g), ("B", self.v_b)):
            ttk.Label(self.f_rgb, text=lbl, style="Bold.TLabel").pack(side="left")
            self._spin(self.f_rgb, var, 0, 255, 1, width=4, fmt_="%.0f").pack(side="left", padx=(px(3), px(8)))
        e = ttk.Entry(self.f_rgb, textvariable=self.v_hex, width=8)
        e.pack(side="left")
        e.bind("<Return>", self._hex_to_rgb)
        e.bind("<FocusOut>", self._hex_to_rgb)
        ttk.Button(self.f_rgb, text=tr("Палитра…"), command=self.pick_color).pack(side="right")
        self._show_color_frame()
        self._update_swatch()

    def _build_wrap_box(self, box):
        px = self.px
        row = self._row(box, 0)
        ttk.Label(row, text=tr("Заменить на перенос пробел №")).pack(side="left")
        for i, var in enumerate(self.v_br, start=1):
            ttk.Checkbutton(row, text=str(i), variable=var).pack(side="left", padx=(px(8), 0))
        ttk.Checkbutton(box, text=tr("только если текст не помещается в макс. ширину"),
                        variable=self.v_wrap_needed).pack(anchor="w", pady=(px(8), 0))
        row = self._row(box)
        ttk.Label(row, text=tr("Межстрочный интервал, %")).pack(side="left")
        self._spin(row, self.v_leading, 60, 400, 5, width=5, fmt_="%.0f").pack(side="left", padx=px(6))
        self._hint(box, tr("Например, «1» — фамилия на первой строке, имя и отчество на второй; "
                           "«1» и «2» — три строки. Для нумерации: «Ряд 5» → «Ряд» над «5»."))

    def _build_pick_panel(self):
        """Плавающая панель пипетки поверх предпросмотра."""
        px = self.px
        p = self.pick_panel = tk.Frame(self.pcanvas, bg=C_CARD, highlightthickness=2,
                                       highlightbackground=C_ACCENT, highlightcolor=C_ACCENT)
        inner = ttk.Frame(p, padding=(px(12), px(8), px(12), px(10)))
        inner.pack(fill="both")
        head = ttk.Frame(inner)
        head.pack(fill="x")
        dot = tk.Canvas(head, width=px(10), height=px(10), bg=C_CARD, highlightthickness=0)
        dot.create_oval(1, 1, px(10) - 1, px(10) - 1, fill=C_ACCENT, outline="")
        dot.pack(side="left", padx=(0, px(6)))
        ttk.Label(head, text=tr("Пипетка"), font=self.f_big).pack(side="left")
        close = tk.Label(head, image=self.icons["close"], bg=C_CARD, cursor="hand2")
        close.pack(side="right")
        close.bind("<Button-1>", lambda e: self._stop_pick())
        body = ttk.Frame(inner)
        body.pack(fill="x", pady=(px(8), 0))
        self.pk_swatch = tk.Canvas(body, width=px(58), height=px(58), highlightthickness=1,
                                   highlightbackground=C_BORDER, bg=C_CANVAS)
        self.pk_swatch.pack(side="left", anchor="n")
        grid = ttk.Frame(body)
        grid.pack(side="left", padx=(px(12), 0))
        self.pk_rows = []
        for i in range(4):
            name = ttk.Label(grid, text="C", style="Bold.TLabel", width=2)
            name.grid(row=i, column=0, sticky="w")
            bar = tk.Canvas(grid, width=px(120), height=px(9), bg="#EEF3F7", highlightthickness=0)
            bar.grid(row=i, column=1, padx=px(6), pady=px(2))
            val = ttk.Label(grid, text="—", width=5, anchor="e")
            val.grid(row=i, column=2, sticky="e")
            self.pk_rows.append((name, bar, val))
        self.pk_hex = ttk.Label(inner, text="", style="Hint.TLabel")
        self.pk_hex.pack(anchor="w", pady=(px(6), 0))
        self.pk_msg = ttk.Label(inner, text=tr("Щелчок по макету — применить цвет\nEsc, «Готово» или крестик — выйти"),
                                style="Hint.TLabel", justify="left")
        self.pk_msg.pack(anchor="w", pady=(px(2), px(8)))
        ttk.Button(inner, text=tr("Готово"), style="Pick.TButton", command=self._stop_pick).pack(fill="x")

    # ---------------- поддержка и сброс ----------------

    def show_donate(self):
        """Окно с QR-кодом и ссылкой для поддержки."""
        px = self.px
        win = tk.Toplevel(self)
        win.title(tr("Поддержать разработку"))
        win.configure(bg=C_CARD)
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text=tr("Если наша программа помогла решить вашу проблему, будем рады вашей поддержке"),
                 bg=C_CARD, fg=C_TEXT, font=self.f_bold, wraplength=px(300),
                 justify="left").pack(padx=px(18), pady=(px(16), px(12)), anchor="w")
        self._qr_img = tk.PhotoImage(data=QR_PNG)
        tk.Label(win, image=self._qr_img, bg=C_CARD, highlightthickness=1,
                 highlightbackground=C_BORDER).pack()
        link = tk.Label(win, text=DONATE_URL, bg=C_CARD, fg=C_ACCENT_DARK, font=self.f_link, cursor="hand2")
        link.pack(pady=(px(10), 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(DONATE_URL))
        row = ttk.Frame(win, padding=(px(18), px(12)))
        row.pack(fill="x")
        ttk.Button(row, text=tr("Открыть ссылку"), style="Pick.TButton",
                   command=lambda: webbrowser.open(DONATE_URL)).pack(side="left")

        def copy_link():
            self.clipboard_clear()
            self.clipboard_append(DONATE_URL)
            msg.configure(text=tr("Ссылка скопирована"))

        ttk.Button(row, text=tr("Копировать ссылку"), command=copy_link).pack(side="left", padx=px(8))
        ttk.Button(row, text=tr("Закрыть"), command=win.destroy).pack(side="right")
        msg = tk.Label(win, text="", bg=C_CARD, fg=C_MUTED, font=self.f_small)
        msg.pack(pady=(0, px(10)))
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _promo_flash(self):
        """После формирования PDF ненадолго подсветить блок о печати."""
        done, normal = tr("Готово! Напечатать быстро и дешево можно тут"), tr("Напечатать быстро и дешево можно тут")
        self.lbl_promo.configure(text=done)
        for w in self._promo:
            w.configure(bg="#BEE1F6")

        def back():
            self._jobs.pop("promo", None)
            try:
                self.lbl_promo.configure(text=normal)
                for w in self._promo:
                    w.configure(bg=C_ACCENT_LIGHT)
            except tk.TclError:
                pass

        if "promo" in self._jobs:
            self.after_cancel(self._jobs.pop("promo"))
        self._jobs["promo"] = self.after(6000, back)

    def reset_settings(self):
        """Вернуть настройки к исходным (макет и список остаются, действие отменяется по Ctrl+Z)."""
        if not messagebox.askyesno(tr(APP_TITLE), tr("Сбросить все настройки?\n\nШрифт, цвет, место, "
                                                     "нумерация и перенос вернутся к исходным. Макет и "
                                                     "список останутся.\nДействие можно отменить по Ctrl+Z.")):
            return
        self._stop_pick()
        self._flush_history()
        fam = self._default_family()
        self._loading = True
        try:
            self.fields = [default_field() for _ in range(MAX_NUMBERS)]
            if fam:
                for f in self.fields:
                    for st in f["styles"].values():
                        if st is not None:
                            st["family"] = fam["name"]
            for k, v in self.v_n.items():
                v.set(self._num_defaults[k])
            self.v_data_mode.set("fio")
            self.v_field.set("all")
            self.v_part.set("text")
            self.v_num_sep.set(False)
            self.v_prev_idx.set("1")
            self.v_zoom.set("fit")
            self.v_font_filter.set("")
        finally:
            self._loading = False
        for key, sec in self.sections.items():
            self.settings.setdefault("sections", {})[key] = True
            if not sec.expanded:
                sec.toggle()
        self._editing = (0, "text")
        if fam:
            self._refill_fonts()
            self._set_family(fam, note=False)
        self._load_editor(self._editing)
        self._update_mode_ui()
        if self.tpl_doc:
            self._on_page_changed(force_center=True, new_template=True)
        self._note_change()
        self.v_status.set(tr("Настройки сброшены (Ctrl+Z — вернуть)."))

    # ---------------- язык ----------------

    def set_language(self, lang):
        """Переключение языка: окно перестраивается, все данные и настройки сохраняются."""
        if lang == LANG:
            return
        if self.busy:
            self.v_lang.set(LANG)
            return
        self._stop_pick()
        self._flush_history()
        self._store_editor(self._editing)
        names = self.txt.get("1.0", "end-1c")
        for k in list(self._jobs):
            self.after_cancel(self._jobs.pop(k))
        try:
            pos = self.paned.sashpos(0)
            if pos > self.px(200):
                self._pane_w = pos          # сохраняем ширину панели, если её меняли мышью
        except tk.TclError:
            pass
        set_lang(lang)
        self.settings["lang"] = lang
        self.title(tr(APP_TITLE))
        for w in self.winfo_children():
            w.destroy()
        self._packinfo, self._fbox, self._anchor_px = {}, {}, {}
        self._build_ui()
        for name in self._combos:
            self._sync_combo(name)
        self._loading = True
        try:
            self.txt.insert("1.0", names)
            self.txt.edit_modified(False)
        finally:
            self._loading = False
        self.names = parse_names(names)
        self._refill_fonts()
        self._load_editor(self._editing)
        self._update_mode_ui()
        self.v_status.set("")
        if self.tpl_doc:
            self.v_page_info.set(tr("из {n}").format(n=self.tpl_doc.page_count))
            self.sp_page.configure(to=self.tpl_doc.page_count)
            if self.tpl_doc.page_count > 1:
                self.row_page.pack(fill="x", pady=(self.px(8), 0), before=self.tpl_hint)
            self.pcanvas.configure(cursor="crosshair")
            self.after(80, self._render_base)
        else:
            self.v_tpl_name.set(tr("макет не выбран"))
        self._update_undo_buttons()

    # ---------------- поля, части и редактор ----------------

    def _numbering(self):
        return self.v_data_mode.get() == "num"

    def _count(self):
        return int(clamp(num(self.v_n["n_count"].get(), 1), 1, MAX_NUMBERS))

    def _separate(self):
        return self._numbering() and self._count() > 1 and self.v_n["n_layout"].get() == "separate"

    def _used_fields(self):
        return list(range(self._count())) if self._separate() else [0]

    def _active_field(self):
        if not self._separate():
            return 0
        v = self.v_field.get()
        if v == "all" or not v.isdigit() or int(v) >= self._count():
            return ALL
        return int(v)

    def _active_part(self):
        return "num" if self._numbering() and self.v_num_sep.get() and self.v_part.get() == "num" else "text"

    def _active_set(self):
        """Поля, которые сейчас редактируются (одно или все)."""
        return self._used_fields() if self._editing[0] == ALL else [self._editing[0]]

    @staticmethod
    def _rep(fi):
        return 0 if fi == ALL else fi

    def _family_by_name(self, name):
        return self._fam_index.get(name.casefold()) if name else None

    @staticmethod
    def _style_for(f, part):
        if f["styles"].get(part) is None:
            f["styles"][part] = copy.deepcopy(f["styles"]["text"])
        return f["styles"][part]

    def _editor_values(self):
        vals = {k: self._evars[k].get() for k in FIELD_KEYS + STYLE_KEYS}
        vals["family"] = self.family["name"] if self.family else None
        return vals

    def _store_editor(self, target=None):
        """Значения из полей ввода — в модель. В режиме «Все номера» переносятся только изменённые
        настройки: положение — сдвигом, остальное — одинаковым значением для всех номеров."""
        fi, part = target or self._editing
        if fi == ALL:
            cur = self._editor_values()
            base = self._all_base or cur
            changed = {k: v for k, v in cur.items() if base.get(k) != v and v is not None}
            for t in self._used_fields():
                f = self.fields[t]
                for k, v in changed.items():
                    if k in ("x", "ypos"):
                        f[k] = fmt(round(num(f[k]) + num(v) - num(base.get(k)), 2))
                    elif k in FIELD_KEYS:
                        f[k] = v
                    else:
                        self._style_for(f, part)[k] = v
            self._all_base = cur
            return
        f = self.fields[fi]
        for k in FIELD_KEYS:
            f[k] = self._evars[k].get()
        st = self._style_for(f, part)
        for k in STYLE_KEYS:
            st[k] = self._evars[k].get()
        if self.family:
            st["family"] = self.family["name"]

    def _load_editor(self, target):
        """Модель — в поля ввода (для «Все номера» — значения номера 1)."""
        fi, part = target
        f = self.fields[self._rep(fi)]
        st = f["styles"].get(part) or f["styles"]["text"]
        self._loading = True
        try:
            for k in FIELD_KEYS:
                if self._evars[k].get() != f[k]:
                    self._evars[k].set(f[k])
            for k in STYLE_KEYS:
                if self._evars[k].get() != st[k]:
                    self._evars[k].set(st[k])
            fam = self._family_by_name(st.get("family"))
            if fam and fam is not self.family:
                self._set_family(fam, note=False)
        finally:
            self._loading = False
        self._all_base = self._editor_values() if fi == ALL else None
        self._on_style_toggle()
        self._update_swatch()
        self._show_color_frame()
        self.schedule("overlay", 20)

    def _retarget(self):
        """Переключение поля/части: сохранить редактор в старую цель и загрузить новую."""
        if self._loading:
            return
        if self._separate() and any(not self.fields[k].get("init") for k in range(1, self._count())):
            self._store_editor(self._editing)
            self._init_fields()          # новые номера сразу встают под предыдущими
            if self._editing[0] == ALL:
                self._all_base = self._editor_values()
        new = (self._active_field(), self._active_part())
        if new != self._editing:
            self._store_editor(self._editing)
            self._editing = new
            self._load_editor(new)
        self._update_mode_ui()

    def _init_fields(self):
        for k in range(1, self._count()):
            if self.fields[k].get("init"):
                continue
            prev = self.fields[k - 1]
            f = copy.deepcopy(prev)
            sizes = [num(st.get("size", "14"), 14) for st in prev["styles"].values() if st]
            f["ypos"] = fmt(round(num(prev["ypos"]) + max(sizes) * 1.3 * 25.4 / 72, 1))
            f["init"] = True
            self.fields[k] = f

    def _on_num_sep(self):
        if self._loading:
            return
        on = self.v_num_sep.get()
        self._store_editor(self._editing)        # в режиме «Все номера» галочка переходит ко всем
        if on:
            for t in self._active_set():
                self._style_for(self.fields[t], "num")
        self.v_part.set("num" if on else "text")
        self._retarget()

    def _update_mode_ui(self):
        if not hasattr(self, "f_num"):
            return
        num_mode = self._numbering()
        n = self._count()
        self._show(self.f_fio, not num_mode)
        self._show(self.f_num, num_mode)
        for fr in self.f_counters + [self.f_multi]:   # заново, чтобы сохранить порядок блоков
            if fr.winfo_manager():
                fr.pack_forget()
        for k, fr in enumerate(self.f_counters):
            if k < n:
                fr.pack(fill="x", before=self.lbl_total)
            self.lbl_chint[k].configure(text=tr(NUM_HINTS[n][k]) if k < n else "")
        if n > 1:
            self.f_multi.pack(fill="x", before=self.lbl_total)
        self._show(self.field_bar, self._separate())
        self._show(self.fb_btns[3], n > 2)
        for frame, seg in self._part_frames:
            self._show(frame, num_mode)
            self._show(seg, num_mode and self.v_num_sep.get())
        self.schedule("overlay", 20)

    # ---------------- записи: ФИО или нумерация ----------------

    def _cfg(self):
        g = {k: v.get() for k, v in self.v_n.items()}
        n = self._count()

        def rng(k):
            return num_range(int(num(g[f"n{k}_from"], 1)), int(num(g[f"n{k}_to"], 1)),
                             int(num(g[f"n{k}_step"], 1)) or 1)

        ks = range(1, n + 1)
        return {"fio": not self._numbering(), "n": n, "ranges": [rng(k) for k in ks],
                "digits": [int(clamp(num(g[f"n{k}_digits"]), 0, 12)) for k in ks],
                "before": [g[f"n{k}_before"] for k in ks], "after": [g[f"n{k}_after"] for k in ks],
                "order": g["n_order"] if g["n_order"] in ("rows", "cols") else "rows",
                "separate": self._separate()}

    def _record_count(self, cfg):
        if cfg["fio"]:
            return len(self.names)
        total = 1
        for r in cfg["ranges"]:
            total *= len(r)
        return total

    def _numbers(self, i, cfg):
        """Значения номеров для страницы i (смешанная система счисления)."""
        rs = cfg["ranges"]
        vals = [None] * len(rs)
        order = range(len(rs) - 1, -1, -1) if cfg["order"] == "rows" else range(len(rs))
        for k in order:
            size = len(rs[k])
            vals[k] = rs[k][i % size]
            i //= size
        return vals

    def _record(self, i, cfg):
        """Поля страницы: [(номер поля, [(текст, 'text'|'num'), …])]."""
        if cfg["fio"]:
            return [(0, [(self.names[i], "text")])]
        vals = self._numbers(i, cfg)
        parts = [[(cfg["before"][k], "text"), (fmt_num(vals[k], cfg["digits"][k]), "num"),
                  (cfg["after"][k], "text")] for k in range(cfg["n"])]
        if cfg["separate"]:
            return [(k, parts[k]) for k in range(cfg["n"])]
        return [(0, [run for part in parts for run in part])]

    def _record_text(self, i, cfg):
        return " | ".join("".join(t for t, _ in runs) for _, runs in self._record(i, cfg))

    def _render_style(self, S):
        fam = self._family_by_name(S.get("family")) or self.family
        if not fam:
            return None
        face, fake_b, fake_i = resolve_face(fam, bool(S["bold"]), bool(S["italic"]))
        data, font = load_font(face)
        if S["mode"] == "CMYK":
            color = tuple(clamp(num(S[c]), 0, 100) / 100 for c in "cmyk")
        else:
            color = tuple(clamp(round(num(S[c])), 0, 255) / 255 for c in "rgb")
        return {"data": data, "font": font, "tag": font_tag(face), "face": face,
                "size": clamp(num(S["size"], 14), 1, 2000), "color": color,
                "fake_bold": fake_b, "fake_italic": fake_i, "underline": bool(S["underline"]),
                "ul_pos": face["ul_pos"], "ul_th": face["ul_th"]}

    def _field_spec(self, fi):
        f = self.fields[fi]
        text_st = self._render_style(f["styles"]["text"])
        if text_st is None:
            return None
        num_st = text_st
        if self._numbering() and f["num_sep"] and f["styles"].get("num"):
            num_st = self._render_style(f["styles"]["num"]) or text_st
        align = f["align"] if f["align"] in ("left", "center", "right") else "center"
        return {"x": num(f["x"]), "y": num(f["ypos"]), "align": align,
                "max_w": max(0.0, num(f["maxw"])),
                "breaks": {i for i in (1, 2, 3) if f[f"br{i}"]}, "wrap_if_needed": bool(f["wrap"]),
                "leading": clamp(num(f["leading"], 120), 30, 1000) / 100,
                "styles": {"text": text_st, "num": num_st}}

    def _specs(self):
        """Настройки всех используемых полей или None, если шрифты ещё не готовы."""
        self._store_editor(self._editing)
        specs = {fi: self._field_spec(fi) for fi in self._used_fields()}
        return specs if all(specs.values()) else None

    # ---------------- клавиатура и мышь ----------------

    def _on_wheel(self, e):
        w = self.winfo_containing(e.x_root, e.y_root)
        if w is None or isinstance(w, (tk.Listbox, tk.Text)):
            return
        step = -1 if (getattr(e, "num", 0) == 4 or getattr(e, "delta", 0) > 0) else 1
        path = str(w)
        if path.startswith(str(self.lcanvas)):
            self.lcanvas.yview_scroll(step * 2, "units")
        elif path.startswith(str(self.pcanvas)):
            if e.state & 0x0001:  # Shift — по горизонтали
                self.pcanvas.xview_scroll(step * 2, "units")
            else:
                self.pcanvas.yview_scroll(step * 2, "units")

    def _on_ctrl_key(self, e):
        """Горячие клавиши по физической клавише — работают и в русской раскладке."""
        if not (e.state & 0x4):
            return None
        if sys.platform.startswith("win"):
            codes = {86: "<<Paste>>", 67: "<<Copy>>", 88: "<<Cut>>", 90: "undo", 89: "redo", 65: "all"}
        elif sys.platform == "darwin":
            return None
        else:
            codes = {55: "<<Paste>>", 54: "<<Copy>>", 53: "<<Cut>>", 52: "undo", 29: "redo", 38: "all"}
        action = codes.get(e.keycode)
        if action == "undo" and e.state & 0x1:
            action = "redo"                     # Ctrl+Shift+Z
        if action == "undo":
            self.undo()
            return "break"
        if action == "redo":
            self.redo()
            return "break"
        w = e.widget
        if action is None or not isinstance(w, (tk.Text, tk.Entry)):
            return None
        if action == "all":
            self._select_all(w)
            return "break"
        if e.keysym.lower() in ("v", "c", "x"):
            return None  # латинская раскладка — работает штатно
        w.event_generate(action)
        return "break"

    def _select_all(self, w):
        if isinstance(w, tk.Text):
            w.tag_add("sel", "1.0", "end-1c")
            w.mark_set("insert", "end")
        elif isinstance(w, tk.Entry):
            w.select_range(0, "end")
            w.icursor("end")

    def _txt_event(self, ev):
        self.txt.focus_set()
        self.txt.event_generate(ev)

    # ---------------- отмена действий (Ctrl+Z / Ctrl+Y) ----------------

    def _snapshot(self):
        self._store_editor(self._editing)
        return {"fields": copy.deepcopy(self.fields),
                "g": {k: v.get() for k, v in self._gvars.items()},
                "names": self.txt.get("1.0", "end-1c")}

    def _note_change(self):
        if self._loading:
            return
        if "hist" in self._jobs:
            self.after_cancel(self._jobs.pop("hist"))
        self._jobs["hist"] = self.after(350, self._commit_history)

    def _commit_history(self):
        self._jobs.pop("hist", None)
        if self._loading:
            return
        snap = self._snapshot()
        if self._last_snap is None:
            self._last_snap = snap
        elif snap != self._last_snap:
            self._history.append(self._last_snap)
            del self._history[:-self.UNDO_LIMIT]
            self._future.clear()
            self._last_snap = snap
        self._update_undo_buttons()

    def _reset_history(self):
        if "hist" in self._jobs:
            self.after_cancel(self._jobs.pop("hist"))
        self._history.clear()
        self._future.clear()
        self._last_snap = self._snapshot()
        self._update_undo_buttons()

    def _flush_history(self):
        if "hist" in self._jobs:
            self.after_cancel(self._jobs.pop("hist"))
            self._commit_history()

    def undo(self):
        if self.busy:
            return
        self._flush_history()
        if not self._history:
            self.bell()
            return
        self._future.append(self._last_snap)
        self._restore(self._history.pop())

    def redo(self):
        if self.busy:
            return
        self._flush_history()
        if not self._future:
            self.bell()
            return
        self._history.append(self._last_snap)
        self._restore(self._future.pop())

    def _restore(self, snap):
        self._loading = True
        try:
            self.fields = copy.deepcopy(snap["fields"])
            for k, v in self._gvars.items():
                if v.get() != snap["g"][k]:
                    v.set(snap["g"][k])
            if self.txt.get("1.0", "end-1c") != snap["names"]:
                self.txt.delete("1.0", "end")
                self.txt.insert("1.0", snap["names"])
                self._refresh_names()
        finally:
            self._loading = False
        self._editing = (self._active_field(), self._active_part())
        self._load_editor(self._editing)
        self._update_mode_ui()
        if "hist" in self._jobs:
            self.after_cancel(self._jobs.pop("hist"))
        self._last_snap = snap
        self._update_undo_buttons()

    def _update_undo_buttons(self):
        if hasattr(self, "btn_undo") and self.btn_undo.winfo_exists():
            self.btn_undo.state(["!disabled"] if self._history or "hist" in self._jobs else ["disabled"])
            self.btn_redo.state(["!disabled"] if self._future else ["disabled"])

    # ---------------- перетаскивание файлов ----------------

    def _on_drag_enter(self, e):
        if e.widget is self.pcanvas:
            if self.tpl_doc:
                c = self.pcanvas
                c.delete("dndhl")
                c.create_rectangle(c.canvasx(3), c.canvasy(3), c.canvasx(c.winfo_width() - 4),
                                   c.canvasy(c.winfo_height() - 4), outline=C_ACCENT, width=3,
                                   tags="dndhl")
            else:
                self._draw_empty(highlight=True)
        return "copy"

    def _on_drag_leave(self, e):
        self.pcanvas.delete("dndhl")
        if not self.tpl_doc:
            self._draw_empty()
        return "copy"

    def _on_drop(self, e):
        self._on_drag_leave(e)
        try:
            files = [f for f in self.tk.splitlist(e.data) if os.path.isfile(f)]
        except Exception:
            files = []
        if not files:
            return "copy"
        path = files[0]
        ext = os.path.splitext(path)[1].lower()
        if ext in TEMPLATE_EXTS:
            self.after(10, lambda: self.open_template(path))
        elif ext in LIST_EXTS:
            self.after(10, lambda: self.load_names_file(path))
        elif ext in FONT_EXTS:
            self.after(10, lambda: self.add_font_file(path))
        else:
            messagebox.showwarning(tr(APP_TITLE), tr("Можно перетащить макет (PDF, JPG, PNG, TIFF), "
                                                     "список (.xlsx, .csv, .txt) или файл шрифта (.ttf, .otf)."))
        return "copy"

    # ---------------- отложенные перерисовки ----------------

    def schedule(self, what, delay=80):
        """what: 'page' (смена страницы) > 'base' (фон) > 'overlay' (только текст)."""
        if self.busy:
            return
        order = {"overlay": 0, "base": 1, "page": 2, "dpi": 3}
        for k in [k for k in self._jobs if k in order]:
            if order[k] <= order[what]:
                self.after_cancel(self._jobs.pop(k))
            else:
                return  # уже запланирована более полная перерисовка
        self._jobs[what] = self.after(delay, lambda: self._run_job(what))

    def _run_job(self, what):
        self._jobs.pop(what, None)
        try:
            if what == "dpi":
                self._reload_image()
            elif what == "page":
                self._on_page_changed()
            elif what == "base":
                self._render_base()
            else:
                self._render_overlay()
        except Exception as ex:  # предпросмотр не должен ронять программу
            self.v_info.set(tr("Ошибка предпросмотра: {e}").format(e=ex))

    # ---------------- макет ----------------

    def open_template(self, path=None, dpi=None):
        if self.busy:
            return
        if not path:
            imgs = "*.jpg *.jpeg *.png *.tif *.tiff *.bmp"
            path = filedialog.askopenfilename(
                title=tr("Выберите макет"),
                initialdir=os.path.dirname(self.tpl_path) if self.tpl_path else None,
                filetypes=[(tr("Макеты (PDF, JPG, PNG, TIFF)"), "*.pdf " + imgs), ("PDF", "*.pdf"),
                           (tr("Изображения"), imgs), (tr("Все файлы"), "*.*")])
            if not path:
                return
        path = os.path.abspath(path)
        is_image = os.path.splitext(path)[1].lower() in IMAGE_EXTS
        try:
            if is_image:
                src, iw, ih, used = pdf_from_image(path, dpi)
                doc = open_pdf_src(src)
                self._img_size, self._img_dpi = (iw, ih), used
            else:
                src = None
                doc = open_pdf(path)
                if not doc.is_pdf:
                    raise RuntimeError(tr("Это не PDF-файл."))
                if doc.needs_pass:
                    raise RuntimeError(tr("Файл защищён паролем."))
                if doc.page_count == 0:
                    raise RuntimeError(tr("В файле нет страниц."))
        except Exception as ex:
            messagebox.showerror(tr(APP_TITLE), tr("Не удалось открыть макет:\n{path}\n\n{e}").format(path=path, e=ex))
            return
        self._stop_pick()
        if self.tpl_doc:
            self.tpl_doc.close()
        first_time = self.tpl_path is None and not self.settings.get("fields") and not self.settings.get("x")
        new_file = path != self.tpl_path
        self.tpl_doc, self.tpl_path, self.tpl_src = doc, path, src
        if is_image:
            self._loading = True
            try:
                self.v_dpi.set(str(self._img_dpi))
            finally:
                self._loading = False
            self.lbl_img.configure(text="  {} × {} px".format(*self._img_size))
            self.row_dpi.pack(fill="x", pady=(self.px(8), 0), before=self.tpl_hint)
        else:
            self.row_dpi.pack_forget()
        self.v_tpl_name.set(os.path.basename(path))
        self.sp_page.configure(to=doc.page_count)
        if doc.page_count > 1:
            self.row_page.pack(fill="x", pady=(self.px(8), 0), before=self.tpl_hint)
        else:
            self.row_page.pack_forget()
        if not 1 <= int(num(self.v_page.get(), 1)) <= doc.page_count:
            self.v_page.set("1")
        self.pcanvas.configure(cursor="crosshair")
        self._on_page_changed(force_center=first_time, new_template=new_file)

    def _reload_image(self):
        """Пересобрать макет-картинку с другим разрешением."""
        if not (self.tpl_path and self.tpl_src is not None):
            return
        d = int(clamp(num(self.v_dpi.get(), 300), 10, 2400))
        if d != self._img_dpi:
            self.open_template(self.tpl_path, dpi=d)

    def _page_index(self):
        if not self.tpl_doc:
            return 0
        return int(clamp(num(self.v_page.get(), 1), 1, self.tpl_doc.page_count)) - 1

    def _on_page_changed(self, force_center=False, new_template=False):
        if not self.tpl_doc:
            return
        page = self.tpl_doc[self._page_index()]
        self.page_w, self.page_h = page.rect.width, page.rect.height
        self.v_page_info.set(tr("из {n}").format(n=self.tpl_doc.page_count))
        pw, ph = self.page_w / MM, self.page_h / MM
        self._store_editor(self._editing)
        for i, f in enumerate(self.fields):   # поля, оказавшиеся за краем, — в центр
            x, y = num(f["x"], -1), num(f["ypos"], -1)
            if force_center or not (0 <= x <= pw and 0 <= y <= ph):
                f["x"], f["ypos"], f["align"] = fmt(round(pw / 2, 2)), fmt(round(ph / 2 + i * 8, 2)), "center"
            mw = num(f["maxw"])
            if (new_template and mw <= 0) or mw > pw + 0.01:
                f["maxw"] = fmt(round(pw, 1))   # по умолчанию — ширина макета
        self._load_editor(self._editing)
        self._render_base()

    def center_h(self):
        if self.tpl_doc:
            self.v_x.set(fmt(round(self.page_w / MM / 2, 2)))
            self.v_align.set("center")

    def center_page(self):
        if self.tpl_doc:
            self.center_h()
            self.v_ypos.set(fmt(round(self.page_h / MM / 2, 2)))

    def maxw_to_page(self):
        if self.tpl_doc:
            self.v_maxw.set(fmt(round(self.page_w / MM, 1)))

    # ---------------- предпросмотр ----------------

    def _on_canvas_resize(self, _e=None):
        if self.tpl_doc:
            if self.v_zoom.get() == "fit":
                self.schedule("base", 150)
            self._place_badge()
        else:
            self._draw_empty()

    def _draw_empty(self, highlight=False):
        if self.tpl_doc or not self.pcanvas.winfo_exists():
            return
        c = self.pcanvas
        c.delete("all")
        px = self.px
        w, h = max(c.winfo_width(), 300), max(c.winfo_height(), 240)
        c.configure(scrollregion=(0, 0, w, h))
        c.xview_moveto(0)
        c.yview_moveto(0)
        m = px(24)
        c.create_rectangle(m, m, w - m, h - m, fill=C_ACCENT_LIGHT if highlight else C_BG,
                           outline=C_ACCENT, width=2, dash=() if highlight else (10, 6))
        cx, cy = w / 2, h / 2 - px(20)
        dw, dh, fold = px(34), px(44), px(14)
        c.create_polygon(cx - dw, cy - dh, cx + dw - fold, cy - dh, cx + dw, cy - dh + fold,
                         cx + dw, cy + dh, cx - dw, cy + dh, fill="white", outline=C_BORDER, width=1)
        c.create_polygon(cx + dw - fold, cy - dh, cx + dw, cy - dh + fold, cx + dw - fold, cy - dh + fold,
                         fill=C_ACCENT_LIGHT, outline=C_BORDER)
        c.create_rectangle(cx - dw - px(8), cy + px(2), cx + dw - px(10), cy + px(24), fill=C_ACCENT, outline="")
        c.create_text(cx - px(9), cy + px(13), text="PDF", fill="white", font=self.f_bold)
        c.create_text(cx, cy + dh + px(34), text=tr("Нажмите, чтобы выбрать макет"),
                      fill=C_TEXT, font=self.f_empty)
        sub = tr("PDF, JPG, PNG, TIFF — или перетащите файл сюда") if self.dnd else tr("PDF, JPG, PNG, TIFF")
        c.create_text(cx, cy + dh + px(62), text=sub, fill=C_MUTED, font=self.f_sub)

    def _zoom_factor(self):
        z = self.v_zoom.get()
        if z == "fit":
            cw = max(80, self.pcanvas.winfo_width() - 2 * PAD)
            ch = max(80, self.pcanvas.winfo_height() - 2 * PAD)
            return max(0.05, min(cw / self.page_w, ch / self.page_h))
        return num(z.rstrip("%"), 100) / 100 * self.winfo_fpixels("1i") / 72

    def _render_base(self):
        if not self.tpl_doc:
            return
        self._zoom = z = self._zoom_factor()
        page = self.tpl_doc[self._page_index()]
        pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
        self._base_pix, self._cmyk_pix = pix, None
        if self.picking:
            self._make_cmyk_pix()
        self._base_img = tk.PhotoImage(data=pix.tobytes("ppm"))
        c = self.pcanvas
        c.delete("all")
        for off, col in ((5, "#C3CFDA"), (3, "#B3C1CE")):   # мягкая тень листа
            c.create_rectangle(PAD + off, PAD + off, PAD + pix.width + off, PAD + pix.height + off,
                               fill=col, outline="")
        c.create_image(PAD, PAD, image=self._base_img, anchor="nw")
        c.configure(scrollregion=(0, 0, pix.width + 2 * PAD, pix.height + 2 * PAD))
        self._anchor_px = {}
        self._render_overlay()

    def _current_face(self):
        """(начертание, имитировать жирный, имитировать курсив) или None."""
        if not self.family:
            return None
        return resolve_face(self.family, self.v_bold.get(), self.v_italic.get())

    def _render_overlay(self):
        c = self.pcanvas
        c.delete("ovl")
        self._update_summaries()
        if not self.tpl_doc:
            return
        z = self._zoom
        active = self._active_set()
        rep = self._rep(self._editing[0])
        info = []
        self._fbox, self._anchor_px, self._ovl_imgs = {}, {}, []
        try:
            specs = self._specs()
        except Exception as ex:
            specs = None
            info.append(tr("Ошибка шрифта: {e}").format(e=ex))
        if specs:
            cfg = self._cfg()
            count = self._record_count(cfg)
            if count:
                idx = int(clamp(num(self.v_prev_idx.get(), 1), 1, count))
                items = self._record(idx - 1, cfg)
            else:
                idx, items = 0, [(0, [(tr(SAMPLE_TEXT), "text")])]
            for fi, runs in items:
                spec = specs[fi]
                tmp = fitz.open()
                pg = tmp.new_page(width=self.page_w, height=self.page_h)
                res = place_runs(pg, runs, spec)
                tag = f"f{fi}"
                if res["bbox"]:
                    pad = res["size"] * 0.45
                    clip = (res["bbox"] + (-pad, -pad, pad, pad)) & pg.rect
                    if not clip.is_empty:   # рисуем только область текста — это быстро
                        pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), clip=clip, alpha=True)
                        img = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")))
                        self._ovl_imgs.append(img)
                        c.create_image(PAD + pix.x, PAD + pix.y, image=img, anchor="nw", tags=("ovl", tag))
                    b = res["bbox"]
                    self._fbox[fi] = [PAD + b.x0 * z, PAD + b.y0 * z, PAD + b.x1 * z, PAD + b.y1 * z]
                tmp.close()
                if fi == rep:
                    if idx:
                        info.append(tr("Запись {i} из {n}").format(i=idx, n=count))
                        if not cfg["fio"]:
                            info[-1] += f" ({self._record_text(idx - 1, cfg)})"
                    else:
                        info.append(tr("Образец текста"))
                    info.append(tr("ширина {w} мм").format(w=f"{max(res['widths']) / MM:.1f}"))
                    if len(res["lines"]) > 1:
                        info.append(tr("строк: {n}").format(n=len(res["lines"])))
                    if res["k"] < 1 - 1e-6:
                        info.append(tr("кегль уменьшен до {p}%, чтобы влезть").format(p=f"{res['k'] * 100:.0f}"))
                    if res["bbox"] and (res["bbox"].x0 < -0.5 or res["bbox"].x1 > self.page_w + 0.5
                                        or res["baselines"][-1] > self.page_h):
                        info.append(tr("(!) текст выходит за край страницы"))
                missing = sorted({ch for t, key in runs for ch in t if not ch.isspace()
                                  and not spec["styles"][key]["font"].has_glyph(ord(ch))})
                if missing:
                    info.append(tr("(!) в шрифте нет символов: {s}").format(s=" ".join(missing)))
            many = len(self._used_fields()) > 1
            for fi in self._used_fields():   # точки привязки полей
                f = self.fields[fi]
                ax, ay = PAD + num(f["x"]) * MM * z, PAD + num(f["ypos"]) * MM * z
                self._anchor_px[fi] = (ax, ay)
                on = fi in active
                col = C_MARK if on else "#8FA3B6"
                tags = ("ovl", f"f{fi}")
                if on and num(f["maxw"]) > 0:   # зона максимальной ширины
                    mw = num(f["maxw"]) * MM * z
                    al = f["align"] if f["align"] in ("left", "center", "right") else "center"
                    x0 = {"left": ax, "center": ax - mw / 2, "right": ax - mw}[al]
                    for dx in (x0, x0 + mw):
                        c.create_line(dx, ay - 6, dx, ay + 6, fill=C_ACCENT, width=2, tags=tags)
                    c.create_line(x0, ay + 4, x0 + mw, ay + 4, fill=C_ACCENT, dash=(4, 3), tags=tags)
                c.create_line(ax - 14, ay, ax + 14, ay, fill=col, width=1, tags=tags)
                c.create_line(ax, ay - 14, ax, ay + 14, fill=col, width=1, tags=tags)
                c.create_oval(ax - 3, ay - 3, ax + 3, ay + 3, outline=col, width=2, tags=tags)
                if many:
                    c.create_text(ax + 8, ay - 12, text=str(fi + 1), fill=col, font=self.f_small_bold,
                                  anchor="w", tags=tags)
        elif self._fonts_loading:
            info.append(tr("Загрузка списка шрифтов…"))
        elif not info:
            info.append(tr("Выберите шрифт"))
        self.v_info.set("  ·  ".join(info))
        self._place_badge()

    def _update_summaries(self):
        if not hasattr(self, "sections"):
            return
        sec = self.sections
        fi = self._editing[0]
        prefix = ""
        if self._separate():
            prefix = (tr("все") if fi == ALL else f"№{fi + 1}") + " · "
        part = tr("номер: ") if self._editing[1] == "num" else ""
        if self.tpl_doc:
            sec["tpl"].set_summary(f"{fmt(round(self.page_w / MM, 1))} × {fmt(round(self.page_h / MM, 1))} "
                                   + tr("мм"))
            align = {"left": tr("слева"), "center": tr("по центру"), "right": tr("справа")}.get(
                self.v_align.get(), tr("по центру"))
            sec["pos"].set_summary(prefix + f"X {fmt(num(self.v_x.get()))} · Y {fmt(num(self.v_ypos.get()))} "
                                   + tr("мм") + f" · {align}")
        else:
            sec["tpl"].set_summary(tr("не выбран"))
            sec["pos"].set_summary("")
        cfg = self._cfg()
        count = self._record_count(cfg)
        pcs = tr("шт")
        if cfg["fio"]:
            sec["data"].set_summary(tr("ФИО") + f" · {count} {pcs}")
        else:
            rs = cfg["ranges"]
            if len(rs) > 1:
                total = " × ".join(str(len(r)) for r in rs) + f" = {count}"
            else:
                total = f"{fmt_num(rs[0][0], cfg['digits'][0])}–{fmt_num(rs[0][-1], cfg['digits'][0])} · {count}"
            sec["data"].set_summary(tr("нумерация") + f" · {total} {pcs}")
            shown = total if len(rs) > 1 else str(count)
            self.lbl_total.configure(text=tr("Итого: {t} {w}").format(t=shown, w=count_word(count)))
            if count:
                ex = tr("Первый: «{a}»").format(a=self._record_text(0, cfg))
                if count > 1:
                    ex += "   ·   " + tr("последний: «{b}»").format(b=self._record_text(count - 1, cfg))
                self.lbl_example.configure(text=ex)
        self.sp_prev.configure(to=max(1, count))
        if count and num(self.v_prev_idx.get(), 1) > count:
            self.v_prev_idx.set("1")
        if self.family:
            marks = "".join(tr(m) for m, v in (("Ж", self.v_bold), ("К", self.v_italic), ("Ч", self.v_underline))
                            if v.get())
            sec["font"].set_summary(prefix + part + f"{fmt(num(self.v_size.get(), 14))} pt {marks}".strip()
                                    + " · " + self.family["name"])
        if self.v_mode.get() == "CMYK":
            sec["color"].set_summary(prefix + part + "C{} M{} Y{} K{}".format(*(fmt(v) for v in self._cmyk100())))
        else:
            sec["color"].set_summary(prefix + part + self.v_hex.get())
        br = [str(i + 1) for i, v in enumerate(self.v_br) if v.get()]
        if br:
            sec["wrap"].set_summary(prefix + tr("пробел") + " " + ", ".join(br)
                                    + (" · " + tr("если не влезает") if self.v_wrap_needed.get() else ""))
        else:
            sec["wrap"].set_summary(prefix + tr("выключен"))

    # --- табличка с размером ---

    def _badge_image(self, w, h):
        key = (w, h)
        if key not in self._badge_cache:
            d = fitz.open()
            pg = d.new_page(width=w, height=h)
            pg.draw_rect(pg.rect, color=None, fill=(0.12, 0.17, 0.22), fill_opacity=0.62,
                         radius=(min(0.5, 8 / w), min(0.5, 8 / h)))
            pix = pg.get_pixmap(alpha=True)
            d.close()
            self._badge_cache[key] = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")))
        return self._badge_cache[key]

    def _place_badge(self):
        if not hasattr(self, "pcanvas") or not self.pcanvas.winfo_exists():
            return
        c = self.pcanvas
        c.delete("badge")
        if not self.tpl_doc:
            return
        mm = tr("мм")
        lines = [f"{fmt(round(self.page_w / MM, 1))} × {fmt(round(self.page_h / MM, 1))} {mm}"]
        if self._hover:
            lines.append("X {:.1f}   Y {:.1f} ".format(*self._hover) + mm)
        tw = max(self.f_badge.measure(lines[0]), self.f_badge.measure("X 0000.0   Y 0000.0 " + mm))
        w = int(tw + self.px(26))
        h = int(self.f_badge.metrics("linespace") * len(lines) + self.px(14))
        right = c.canvasx(c.winfo_width()) - self.px(12)
        top = c.canvasy(0) + self.px(12)
        c.create_image(right - w, top, image=self._badge_image(w, h), anchor="nw", tags="badge")
        c.create_text(right - w / 2, top + h / 2, text="\n".join(lines), fill="white",
                      font=self.f_badge, justify="center", tags="badge")

    def _set_hover(self, hover):
        if hover != self._hover:
            self._hover = hover
            self._place_badge()

    def _canvas_mm(self, e):
        x = (self.pcanvas.canvasx(e.x) - PAD) / self._zoom / MM
        y = (self.pcanvas.canvasy(e.y) - PAD) / self._zoom / MM
        return x, y

    def _on_canvas_motion(self, e):
        if not self.tpl_doc:
            return
        x, y = self._canvas_mm(e)
        inside = 0 <= x <= self.page_w / MM and 0 <= y <= self.page_h / MM
        self._set_hover((round(x, 1), round(y, 1)) if inside else None)
        if self.picking:
            self._pick_hover(e)

    def _hit_field(self, e):
        """Номер поля, по тексту которого щёлкнули (сначала — редактируемые)."""
        cx, cy = self.pcanvas.canvasx(e.x), self.pcanvas.canvasy(e.y)
        active = self._active_set()
        for fi in sorted(self._fbox, key=lambda f: f not in active):
            x0, y0, x1, y1 = self._fbox[fi]
            if x0 - 4 <= cx <= x1 + 4 and y0 - 4 <= cy <= y1 + 4:
                return fi
        return None

    def _on_canvas_press(self, e):
        if self.busy:
            return
        self.pcanvas.focus_set()        # чтобы Ctrl+Z отменял действия, а не набор текста
        if not self.tpl_doc:
            self.open_template()
            return
        if self.picking:
            self._pick_apply(e)
            return
        mm = self._canvas_mm(e)
        hit = self._hit_field(e)
        if hit is not None:             # взяли за текст — тянем, не прыгая
            if self._editing[0] != ALL and hit != self._editing[0]:
                self.v_field.set(str(hit))
            self._drag = (mm, (num(self.v_x.get()), num(self.v_ypos.get())))
        else:                           # щелчок мимо текста — текст встаёт в эту точку
            self._set_anchor(mm)
            self._drag = (mm, (num(self.v_x.get()), num(self.v_ypos.get())))

    def _on_canvas_double(self, e):
        """Двойной щелчок по номеру — настроить только его."""
        if self.tpl_doc and not self.picking and self._separate():
            hit = self._hit_field(e)
            if hit is not None:
                self.v_field.set(str(hit))

    def _on_canvas_drag(self, e):
        if not self._drag or not self.tpl_doc or self.busy or self.picking:
            return
        mm = self._canvas_mm(e)
        (px0, py0), (ax, ay) = self._drag
        self._set_anchor((ax + mm[0] - px0, ay + mm[1] - py0))
        self._on_canvas_motion(e)

    def _set_anchor(self, mm):
        """Мгновенно сдвигает текст на экране (одно поле или все вместе); точная перерисовка — следом."""
        x = round(clamp(mm[0], 0, self.page_w / MM), 1)
        y = round(clamp(mm[1], 0, self.page_h / MM), 1)
        rep = self._rep(self._editing[0])
        old = self._anchor_px.get(rep)
        if old:
            nx, ny = PAD + x * MM * self._zoom, PAD + y * MM * self._zoom
            dx, dy = nx - old[0], ny - old[1]
            for fi in self._active_set():
                self.pcanvas.move(f"f{fi}", dx, dy)
                if fi in self._anchor_px:
                    ax, ay = self._anchor_px[fi]
                    self._anchor_px[fi] = (ax + dx, ay + dy)
                if fi in self._fbox:
                    b = self._fbox[fi]
                    self._fbox[fi] = [b[0] + dx, b[1] + dy, b[2] + dx, b[3] + dy]
        self.v_x.set(fmt(x))
        self.v_ypos.set(fmt(y))

    # --- пипетка ---

    def toggle_pick(self):
        if self.picking:
            self._stop_pick()
        else:
            self.start_pick()

    def start_pick(self):
        if not self.tpl_doc:
            messagebox.showinfo(tr(APP_TITLE), tr("Сначала откройте PDF-макет."))
            return
        self.picking = True
        self._make_cmyk_pix()
        self.pcanvas.configure(cursor="tcross")
        self.btn_pick.configure(style="Pick.TButton", text=tr("Пипетка вкл."))
        self._pick_show(None)
        self.pick_panel.place(x=self.px(14), y=self.px(14))
        self.pick_panel.lift()

    def _stop_pick(self):
        if self.picking:
            self.picking = False
            try:
                self.pick_panel.place_forget()
                self.pcanvas.configure(cursor="crosshair" if self.tpl_doc else "hand2")
                self.btn_pick.configure(style="TButton", text=tr("Пипетка"))
            except tk.TclError:
                pass

    def _make_cmyk_pix(self):
        if self._cmyk_pix is None and self.tpl_doc and self._base_pix:
            page = self.tpl_doc[self._page_index()]
            self._cmyk_pix = page.get_pixmap(matrix=fitz.Matrix(self._zoom, self._zoom),
                                             colorspace=fitz.csCMYK, alpha=False)

    def _pick_sample(self, e):
        """Цвет макета под курсором: (cmyk в %, rgb) или None."""
        x = int(self.pcanvas.canvasx(e.x) - PAD)
        y = int(self.pcanvas.canvasy(e.y) - PAD)
        pix = self._base_pix
        if not pix or not (0 <= x < pix.width and 0 <= y < pix.height):
            return None
        rgb = tuple(pix.pixel(x, y)[:3])
        self._make_cmyk_pix()
        cp = self._cmyk_pix
        cmyk = tuple(int(round(v / 2.55)) for v in cp.pixel(min(x, cp.width - 1), min(y, cp.height - 1))[:4])
        return cmyk, rgb

    def _pick_show(self, sample):
        cmyk_mode = self.v_mode.get() == "CMYK"
        chans = [("C", "#00AEEF"), ("M", "#EC008C"), ("Y", "#F2CF00"), ("K", "#2B2B2B")] if cmyk_mode else \
                [("R", "#E53935"), ("G", "#43A047"), ("B", "#1E88E5"), ("", "")]
        values = None
        if sample:
            cmyk, rgb = sample
            values = [v / 100 for v in cmyk] if cmyk_mode else [v / 255 for v in rgb] + [None]
            self.pk_swatch.configure(bg="#%02x%02x%02x" % rgb)
            self.pk_hex.configure(text="C{} M{} Y{} K{}   ·   #{:02X}{:02X}{:02X}".format(*cmyk, *rgb))
        else:
            self.pk_swatch.configure(bg=C_CANVAS)
            self.pk_hex.configure(text=tr("Наведите курсор на макет"))
        bw, bh = self.px(120), self.px(9)
        for i, (name, bar, val) in enumerate(self.pk_rows):
            label, color = chans[i]
            bar.delete("all")
            if not label:
                name.configure(text="")
                val.configure(text="")
                bar.configure(bg=C_CARD)
                continue
            bar.configure(bg="#EEF3F7")
            name.configure(text=label)
            v = values[i] if values else None
            if v is None:
                val.configure(text="—")
            else:
                bar.create_rectangle(0, 0, max(1, bw * v), bh, fill=color, outline="")
                val.configure(text=f"{round(v * 100)} %" if cmyk_mode else str(round(v * 255)))

    def _pick_hover(self, e):
        self._pick_show(self._pick_sample(e))

    def _pick_apply(self, e):
        sample = self._pick_sample(e)
        if not sample:
            return
        cmyk, rgb = sample
        if self.v_mode.get() == "CMYK":
            self._set_cmyk(*cmyk)
            msg = "C{} M{} Y{} K{}".format(*cmyk)
        else:
            for var, val in zip((self.v_r, self.v_g, self.v_b), rgb):
                var.set(str(val))
            msg = "#{:02X}{:02X}{:02X}".format(*rgb)
        self.v_status.set(tr("Цвет текста взят с макета: {c}").format(c=msg))
        self.pk_msg.configure(text=tr("Применено: {c}\nЩелчок — взять другой · Esc — выйти").format(c=msg))

    # ---------------- список ФИО ----------------

    def _on_text_modified(self, _e=None):
        if not self.txt.edit_modified():
            return
        self.txt.edit_modified(False)
        if "names" in self._jobs:
            self.after_cancel(self._jobs.pop("names"))
        self._jobs["names"] = self.after(250, self._refresh_names)
        self._note_change()

    def _refresh_names(self):
        self._jobs.pop("names", None)
        self.names = parse_names(self.txt.get("1.0", "end"))
        self.schedule("overlay", 40)

    def load_names_file(self, path=None):
        if not path:
            path = filedialog.askopenfilename(
                title=tr("Файл со списком ФИО"),
                filetypes=[(tr("Списки"), "*.xlsx *.xlsm *.csv *.txt"), ("Excel", "*.xlsx *.xlsm"),
                           ("CSV", "*.csv"), (tr("Текст"), "*.txt"), (tr("Все файлы"), "*.*")])
            if not path:
                return
        try:
            rows = read_table(path)
        except Exception as ex:
            messagebox.showerror(tr(APP_TITLE), tr("Не удалось прочитать файл:\n{e}").format(e=ex))
            return
        rows = [r for r in rows if any(c.strip() for c in r)]
        if not rows:
            messagebox.showwarning(tr(APP_TITLE), tr("В файле нет данных."))
            return
        ncols = max(len(r) for r in rows)
        used = [i for i in range(ncols) if any(i < len(r) and r[i].strip() for r in rows)]
        cols = used[:1]
        if len(used) > 1:
            first = rows[0]
            lines = "\n".join(f"{i + 1}: {first[i] if i < len(first) else ''}" for i in used[:12])
            ans = simpledialog.askstring(
                tr(APP_TITLE),
                tr("В таблице несколько столбцов. Первая строка:\n\n{lines}\n\nКакие столбцы взять? "
                   "Номера через запятую (например: 1  или  1,2,3).\n"
                   "Несколько столбцов объединятся через пробел.").format(lines=lines),
                initialvalue=str(used[0] + 1), parent=self)
            if ans is None:
                return
            cols = [int(n) - 1 for n in re.findall(r"\d+", ans) if 0 < int(n) <= ncols]
            if not cols:
                return
        names = [clean_line(" ".join(r[i] for i in cols if i < len(r))) for r in rows]
        names = [n for n in names if n]
        if names and names[0].casefold() in HEADER_WORDS:
            names = names[1:]  # строка-заголовок
        if not names:
            messagebox.showwarning(tr(APP_TITLE), tr("В выбранных столбцах нет данных."))
            return
        replace = True
        if self.txt.get("1.0", "end").strip():
            ans = messagebox.askyesnocancel(
                tr(APP_TITLE), tr("Загружено записей: {n}.\n\nЗаменить текущий список?\n"
                                  "(«Нет» — добавить в конец)").format(n=len(names)))
            if ans is None:
                return
            replace = ans
        if replace:
            self.txt.delete("1.0", "end")
        elif not self.txt.get("end-2c", "end-1c") == "\n":
            self.txt.insert("end", "\n")
        self.txt.insert("end", "\n".join(names))
        self.v_data_mode.set("fio")
        self._refresh_names()
        self._flush_history()

    def show_longest(self):
        cfg = self._cfg()
        count = self._record_count(cfg)
        if not count:
            return
        if cfg["fio"]:
            current = self._current_face()
            if not current:
                return
            _data, font = load_font(current[0])
            lengths = [font.text_length(n, fontsize=10) for n in self.names]
            best = lengths.index(max(lengths))
        elif count <= 50000:
            best = max(range(count), key=lambda i: len(self._record_text(i, cfg)))
        else:
            best = count - 1
        self.v_prev_idx.set(str(best + 1))

    # ---------------- шрифты ----------------

    def _start_font_scan(self):
        self.v_font_name.set(tr("загрузка списка шрифтов…"))
        self.lb_fonts.insert("end", "  " + tr("Загрузка шрифтов…"))
        q = queue.Queue()
        cache = os.path.join(self.cfg_dir, "fonts_cache.json")

        def work():
            try:
                q.put(("ok", scan_system_fonts(cache)))
            except Exception as ex:
                q.put(("err", str(ex)))

        def poll():
            try:
                kind, res = q.get_nowait()
            except queue.Empty:
                self.after(150, poll)
                return
            if kind == "err":
                messagebox.showerror(tr(APP_TITLE), tr("Ошибка при чтении шрифтов:\n{e}").format(e=res))
                res = []
            self.faces = res
            self._fonts_loading = False
            for p in self.settings.get("extra_fonts", []):
                if os.path.isfile(p):
                    self._add_font_faces(p)
            self._set_families(group_families(self.faces))
            self._restore_font()

        threading.Thread(target=work, daemon=True).start()
        self.after(150, poll)

    def _set_families(self, fams):
        self.families = fams
        self._fam_index = {f["name"].casefold(): f for f in fams}

    def _default_family(self, prefer=""):
        """Шрифт по умолчанию: сохранённый, затем привычные, затем первый с кириллицей."""
        wanted = [prefer, "Arial", "Times New Roman", "Calibri", "DejaVu Sans", "Liberation Sans", "Helvetica"]
        found = next((self._family_by_name(w) for w in wanted if self._family_by_name(w)), None)
        if found:
            return found
        return next((f for f in self.families if f["cyr"]), self.families[0] if self.families else None)

    def _restore_font(self):
        fi, part = self._editing
        f0 = self.fields[self._rep(fi)]
        cur = (f0["styles"].get(part) or f0["styles"]["text"]).get("family", "")
        default = self._default_family(cur)
        if not default:
            self.v_font_name.set(tr("шрифты не найдены"))
            return
        for f in self.fields:   # стилям без шрифта (или с удалённым шрифтом) — шрифт по умолчанию
            for st in f["styles"].values():
                if st is not None and not self._family_by_name(st.get("family")):
                    st["family"] = default["name"]
        if not default["cyr"]:
            self.v_cyr_only.set(False)
        self._refill_fonts()
        self._load_editor(self._editing)
        self._reset_history()   # история отмены начинается с загруженного состояния

    def _refill_fonts(self):
        flt = self.v_font_filter.get().strip().casefold()
        cyr = self.v_cyr_only.get()
        self.family_view = [f for f in self.families if (not cyr or f["cyr"]) and flt in f["name"].casefold()]
        self.lb_fonts.delete(0, "end")
        for f in self.family_view:
            self.lb_fonts.insert("end", f["name"])
        if self.family in self.family_view:
            i = self.family_view.index(self.family)
            self.lb_fonts.selection_set(i)
            self.lb_fonts.see(i)

    def _on_font_select(self, _e=None):
        sel = self.lb_fonts.curselection()
        if sel and sel[0] < len(self.family_view):
            self._set_family(self.family_view[sel[0]])

    def _set_family(self, family, note=True):
        changed = family is not self.family
        self.family = family
        self.lb_fonts.selection_clear(0, "end")
        if family in self.family_view:
            i = self.family_view.index(family)
            self.lb_fonts.selection_set(i)
            self.lb_fonts.see(i)
        self._on_style_toggle()
        if changed and note:
            self._note_change()

    def _on_style_toggle(self):
        """Выбор начертания по кнопкам Ж / К и загрузка шрифта."""
        current = self._current_face()
        if not current:
            return
        face, fake_b, fake_i = current
        try:
            self.config(cursor="watch")
            self.update_idletasks()
            load_font(face)
        except Exception as ex:
            messagebox.showerror(tr(APP_TITLE), tr("Не удалось загрузить шрифт «{f}»:\n{e}").format(f=face["full"], e=ex))
            return
        finally:
            self.config(cursor="")
        label = tr("Начертание: {f}").format(f=face["full"])
        fakes = [tr(n) for n, f in (("жирный", fake_b), ("курсив", fake_i)) if f]
        if fakes:
            label += " + " + ", ".join(fakes) + " " + tr("(имитация — в шрифте нет такого начертания)")
        self.v_font_name.set(label)
        self.schedule("overlay", 40)

    def _add_font_faces(self, path):
        added = []
        for face in scan_font_file(path):
            if any(f["path"] == path and f["index"] == face["index"] for f in self.faces):
                continue
            face = dict(face, path=path, from_file=True)
            self.faces.append(face)
            added.append(face)
        return added

    def add_font_file(self, path=None):
        if not path:
            path = filedialog.askopenfilename(title=tr("Файл шрифта"),
                                              filetypes=[(tr("Шрифты"), "*.ttf *.otf *.ttc *.otc"),
                                                         (tr("Все файлы"), "*.*")])
            if not path:
                return
        path = os.path.abspath(path)
        self._add_font_faces(path)
        faces = [f for f in self.faces if f["path"] == path]
        if not faces:
            messagebox.showerror(tr(APP_TITLE), tr("Этот шрифт не поддерживается.\n"
                                                   "Нужен шрифт TrueType или OpenType (.ttf, .otf, .ttc)."))
            return
        extra = self.settings.setdefault("extra_fonts", [])
        if path not in extra:
            extra.append(path)
        self._set_families(group_families(self.faces))
        fam = next(f for f in self.families if faces[0] in f["faces"].values())
        if not fam["cyr"]:
            self.v_cyr_only.set(False)
        self.v_font_filter.set("")
        self._refill_fonts()
        self._set_family(fam)

    # ---------------- цвет ----------------

    def _on_mode_change(self):
        self._show_color_frame()
        if self.picking:
            self._pick_show(None)

    def _show_color_frame(self):
        if not hasattr(self, "f_rgb") or not self.f_rgb.winfo_exists():
            return
        pad = (self.px(10), 0)
        if self.v_mode.get() == "CMYK":
            self.f_rgb.pack_forget()
            self.f_cmyk.pack(fill="x", pady=pad)
        else:
            self.f_cmyk.pack_forget()
            self.f_rgb.pack(fill="x", pady=pad)

    def _rgb255(self):
        return tuple(int(clamp(round(num(v.get())), 0, 255)) for v in (self.v_r, self.v_g, self.v_b))

    def _cmyk100(self):
        return tuple(clamp(num(v.get()), 0, 100) for v in (self.v_c, self.v_m, self.v_y, self.v_k))

    def _screen_rgb(self):
        if self.v_mode.get() == "CMYK":  # приблизительно, только для экрана
            c, m, y, k = (v / 100 for v in self._cmyk100())
            return tuple(int(255 * (1 - v) * (1 - k)) for v in (c, m, y))
        return self._rgb255()

    def _update_swatch(self):
        if not hasattr(self, "swatch") or not self.swatch.winfo_exists():
            return
        col = "#%02x%02x%02x" % self._screen_rgb()
        self.swatch.configure(bg=col)
        if getattr(self, "sec_swatch", None) and self.sec_swatch.winfo_exists():
            self.sec_swatch.configure(bg=col)

    def _rgb_to_hex(self):
        if self._hex_lock:
            return
        self._hex_lock = True
        self.v_hex.set("#%02X%02X%02X" % self._rgb255())
        self._hex_lock = False

    def _hex_to_rgb(self, _e=None):
        h = self.v_hex.get().strip().lstrip("#")
        if len(h) == 3:
            h = "".join(ch * 2 for ch in h)
        if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
            self._rgb_to_hex()
            return
        self._hex_lock = True
        for var, i in ((self.v_r, 0), (self.v_g, 2), (self.v_b, 4)):
            var.set(str(int(h[i:i + 2], 16)))
        self._hex_lock = False
        self._rgb_to_hex()

    def _set_cmyk(self, c, m, y, k):
        for var, val in zip((self.v_c, self.v_m, self.v_y, self.v_k), (c, m, y, k)):
            var.set(str(val))

    def pick_color(self):
        res = colorchooser.askcolor(color="#%02x%02x%02x" % self._rgb255(), title=tr("Цвет текста (RGB)"))
        if res and res[0]:
            r, g, b = (int(round(v)) for v in res[0])
            self.v_r.set(str(r))
            self.v_g.set(str(g))
            self.v_b.set(str(b))

    # ---------------- формирование PDF ----------------

    def _set_progress(self, done, total, text):
        self.progress.configure(maximum=max(1, total), value=done)
        self.v_status.set(tr("{t}  {d} из {n}").format(t=text, d=done, n=total))
        self.update()

    def _missing_glyphs(self, specs, cfg):
        """Символы, которых нет в выбранных шрифтах: [(имя шрифта, 'символы')]."""
        need = {}   # (поле, часть) -> набор символов

        def add(fi, key, text):
            need.setdefault((fi, key), set()).update(ch for ch in text if not ch.isspace())

        if cfg["fio"]:
            for n in self.names:
                add(0, "text", n)
        else:
            neg = any(min(r) < 0 for r in cfg["ranges"])
            digits = "0123456789" + ("-" if neg else "")
            for k in range(cfg["n"]):
                fi = k if cfg["separate"] else 0
                add(fi, "text", cfg["before"][k] + cfg["after"][k])
                add(fi, "num", digits)
        result = []
        for (fi, key), chars in need.items():
            st = specs[fi]["styles"][key]
            miss = sorted(ch for ch in chars if not st["font"].has_glyph(ord(ch)))
            if miss:
                result.append((st["face"]["full"], " ".join(miss[:40])))
        return result

    def generate(self):
        if self.busy:
            return
        self._stop_pick()
        if "names" in self._jobs:
            self.after_cancel(self._jobs.pop("names"))
            self._refresh_names()
        if not self.tpl_doc:
            messagebox.showwarning(tr(APP_TITLE), tr("Сначала откройте PDF-макет."))
            return
        cfg = self._cfg()
        count = self._record_count(cfg)
        if not count:
            messagebox.showwarning(tr(APP_TITLE), tr("Список ФИО пуст.") if cfg["fio"] else tr("Нет номеров."))
            return
        try:
            specs = self._specs()
        except Exception as ex:
            messagebox.showerror(tr(APP_TITLE), tr("Не удалось загрузить шрифт:\n{e}").format(e=ex))
            return
        if not specs:
            messagebox.showwarning(tr(APP_TITLE), tr("Выберите шрифт."))
            return
        if count > 5000 and not messagebox.askyesno(
                tr(APP_TITLE), tr("Получится {n} страниц — это займёт время и большой файл. Продолжить?").format(n=count)):
            return
        missing = self._missing_glyphs(specs, cfg)
        if missing and not messagebox.askyesno(
                tr(APP_TITLE), tr("В шрифтах нет некоторых символов — они не будут видны:") + "\n\n"
                + "\n".join(f"«{name}»: {chars}" for name, chars in missing) + "\n\n" + tr("Продолжить?")):
            return
        stem = os.path.splitext(os.path.basename(self.tpl_path))[0]
        kind = tr("ФИО") if cfg["fio"] else tr("нумерация")
        out_path = filedialog.asksaveasfilename(
            title=tr("Сохранить результат"), defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialdir=self.settings.get("out_dir") or os.path.dirname(self.tpl_path),
            initialfile=f"{stem}_{kind}_{count}{tr('шт')}.pdf")
        if not out_path:
            return
        if os.path.abspath(out_path) == os.path.abspath(self.tpl_path):
            messagebox.showerror(tr(APP_TITLE), tr("Нельзя сохранять поверх исходного макета.\nВыберите другое имя."))
            return
        self.settings["out_dir"] = os.path.dirname(out_path)
        self._save_settings()

        def items(i):
            return [(runs, specs[fi]) for fi, runs in self._record(i, cfg)]

        self.busy = True
        self.btn_go.state(["disabled"])
        self.config(cursor="watch")
        self.progress.pack(fill="x", pady=(self.px(8), 0), before=self.lbl_status)
        try:
            shrunk = build_pdf(self.tpl_src or self.tpl_path, out_path, count, items,
                               self._page_index(), self._set_progress)
        except Exception as ex:
            self.v_status.set(tr("Ошибка."))
            self.progress.configure(value=0)
            messagebox.showerror(tr(APP_TITLE), tr("Не удалось сформировать PDF:\n\n{e}").format(e=ex))
            return
        finally:
            self.busy = False
            self.btn_go.state(["!disabled"])
            self.config(cursor="")
            self.progress.pack_forget()
        msg = tr("Готово!\n\nСтраниц: {n}\n{path}").format(n=count, path=out_path)
        if shrunk:
            msg += "\n\n" + tr("Страниц с уменьшенным кеглем (текст не влезал в макс. ширину): {n}").format(n=shrunk)
        self.v_status.set(tr("Готово: {f}").format(f=os.path.basename(out_path)))
        self._promo_flash()
        if messagebox.askyesno(tr(APP_TITLE), msg + "\n\n" + tr("Открыть файл?")):
            open_with_system(out_path)

    # ---------------- настройки ----------------

    def _settings_file(self):
        return os.path.join(self.cfg_dir, "settings.json")

    def _load_settings(self):
        try:
            with open(self._settings_file(), encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _migrate_old(self, s):
        """Настройки версий 1.2–1.3 (одно поле) — в новую модель."""
        f = default_field()
        st = f["styles"]["text"]
        for key in ("size", "mode"):
            if key in s:
                st[key] = str(s[key])
        for key in ("bold", "italic", "underline"):
            st[key] = bool(s.get(key, False))
        if isinstance(s.get("cmyk"), list) and len(s["cmyk"]) == 4:
            st.update(zip("cmyk", (str(v) for v in s["cmyk"])))
        if isinstance(s.get("rgb"), list) and len(s["rgb"]) == 3:
            st.update(zip("rgb", (str(v) for v in s["rgb"])))
        st["family"] = s.get("font_family", "")
        f["x"], f["ypos"] = str(s.get("x", "0")), str(s.get("y", "0"))
        f["align"] = s.get("align") if s.get("align") in ("left", "center", "right") else "center"
        f["maxw"] = str(s.get("max_w", "0"))
        for i in (1, 2, 3):
            f[f"br{i}"] = i in s.get("breaks", [])
        f["wrap"] = bool(s.get("wrap_if_needed", True))
        f["leading"] = str(s.get("leading", "120"))
        return f

    def _apply_settings(self):
        s = self.settings
        zoom = s.get("zoom", "fit")
        self.v_zoom.set(zoom if zoom in dict(ZOOM_OPTS) else "fit")
        self.v_cyr_only.set(s.get("cyr_only", True))
        self._loading = True
        try:
            saved = s.get("fields")
            if isinstance(saved, list) and saved:
                self.fields = [clean_field(saved[k] if k < len(saved) else None) for k in range(MAX_NUMBERS)]
            elif "x" in s or "size" in s:
                self.fields = [self._migrate_old(s)] + [default_field() for _ in range(MAX_NUMBERS - 1)]
            self.v_page.set(str(s.get("page", 1)))
            self.v_data_mode.set(s.get("data_mode", "fio"))
            saved_num = dict(s.get("num") or {})
            if "n_count" not in saved_num and "n_double" in saved_num:   # настройки версии 1.4
                saved_num["n_count"] = "2" if saved_num.get("n_double") else "1"
            if saved_num.get("n_order") in OLD_ORDER:
                saved_num["n_order"] = OLD_ORDER[saved_num["n_order"]]
            for k, v in saved_num.items():
                if k in self.v_n:
                    self.v_n[k].set(v)
            field = s.get("field", "all")
            self.v_field.set({"0": "all"}.get(field, field) if "n_double" in (s.get("num") or {}) else field)
            self.v_part.set(s.get("part", "text"))
        finally:
            self._loading = False
        self._editing = (self._active_field(), self._active_part())
        self._load_editor(self._editing)
        self._update_mode_ui()
        tpl = s.get("template")
        if tpl and os.path.isfile(tpl):
            dpi = int(clamp(num(s.get("dpi", 300), 300), 10, 2400))
            self.after(200, lambda: self.open_template(tpl, dpi=dpi))

    def _save_settings(self):
        self._store_editor(self._editing)
        s = self.settings
        s.update({
            "fields": self.fields, "data_mode": self.v_data_mode.get(),
            "num": {k: v.get() for k, v in self.v_n.items()},
            "field": self.v_field.get(), "part": self.v_part.get(), "lang": LANG,
            "zoom": self.v_zoom.get(), "cyr_only": self.v_cyr_only.get(),
            "page": self._page_index() + 1, "template": self.tpl_path, "geometry": self.geometry(),
            "dpi": self._img_dpi or 300,
        })
        try:
            with open(self._settings_file(), "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def report_callback_exception(self, exc, val, tb):
        """Непредвиденная ошибка: показываем окно и пишем в error.log (у exe нет консоли)."""
        import traceback
        text = "".join(traceback.format_exception(exc, val, tb))
        try:
            with open(os.path.join(self.cfg_dir, "error.log"), "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except Exception:
            pass
        self.busy = False
        self._loading = False
        try:
            self.btn_go.state(["!disabled"])
            self.config(cursor="")
        except Exception:
            pass
        messagebox.showerror(tr(APP_TITLE), tr("Непредвиденная ошибка:\n\n{e}\n\nПодробности записаны в файл:\n{f}")
                             .format(e=val, f=os.path.join(self.cfg_dir, "error.log")))

    def _on_close(self):
        if self.busy:
            return
        self._save_settings()
        for job in list(self._jobs.values()):   # снимаем отложенные перерисовки
            try:
                self.after_cancel(job)
            except tk.TclError:
                pass
        self._jobs.clear()
        self.destroy()


def main():
    if sys.platform.startswith("win"):
        try:  # чёткий интерфейс на экранах с масштабированием
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(TEMPLATE_EXTS):
        app.after(300, lambda: app.open_template(sys.argv[1]))
    app.mainloop()


if __name__ == "__main__":
    main()
