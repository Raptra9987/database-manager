# -*- coding: utf-8 -*-
"""
Database Manager  -  a Windows 11 styled GUI
==========================================================================
Merges many files into a single Excel database. It is not limited to
contacts: the columns come from a "schema", so products, students or any
list of your own work just as well.

How it is used:
  1) "New database"   -> you pick the type, that is, the set of columns
     "Open database"  -> the columns are read from the file and numbering
                         continues from the last row
  2) "Add contacts"   -> you are asked who the file came from and that goes
                         into the "Source" column
  3) Repeated rows are marked red. Which column identifies a duplicate is
     part of the schema, and the marking is kept in Excel too.

Speaks three languages: Uzbek, Russian and English.

Run with:  python database_manager.py
"""

import os
import re
import csv
import sys
import json
import copy
import ctypes
import shutil
import datetime
import traceback
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import font as tkfont

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import sv_ttk
    HAS_SV = True
except ImportError:
    HAS_SV = False

try:                                   # lets files be dropped on the window
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    HAS_DND = False


# ================================================================= translations
#
# Three strings per key:  (Uzbek, Russian, English).
# When you add a string, fill in all three.

LANG_ORDER = ("uz", "ru", "en")
LANG_NAMES = {"uz": "O'zbekcha", "ru": "Русский", "en": "English"}
LANG_SHORT = {"uz": "UZ", "ru": "RU", "en": "EN"}

LANG = "uz"                            # current language (read from settings)

TR = {
    # --- dastur nomi va logotip
    "app_name": ("Baza Menejeri", "Менеджер баз", "Database Manager"),
    "brand1": ("Baza Menejeri", "Менеджер баз", "Database Manager"),

    # --- baza turlari (shablonlar)
    "tpl_choose": ("Baza turini tanlang", "Выберите тип базы",
                   "Choose the database type"),
    "tpl_contacts": ("Kontaktlar", "Контакты", "Contacts"),
    "tpl_products": ("Mahsulotlar", "Товары", "Products"),
    "tpl_students": ("Talabalar", "Студенты", "Students"),
    "tpl_custom": ("O'z ustunlarim…", "Свои столбцы…", "My own columns…"),
    "tpl_opened": ("Fayldan o'qilgan", "Прочитано из файла", "Read from file"),

    # --- shablon ustunlari
    "h_pname": ("Nomi", "Название", "Name"),
    "h_code": ("Kod / Artikul", "Код / Артикул", "Code / SKU"),
    "h_price": ("Narxi", "Цена", "Price"),
    "h_qty": ("Soni", "Количество", "Quantity"),
    "h_group": ("Guruh", "Группа", "Group"),

    # --- ustunlarni o'zim tuzish
    "sch_title": ("O'z ustunlarim", "Свои столбцы", "My own columns"),
    "sch_head": ("Baza ustunlarini yozing", "Укажите столбцы базы",
                 "Define the database columns"),
    "sch_body": ("Har bir ustunni alohida qatorga yozing.\n"
                 "Tartib raqami va manba ustuni pastdagi tugmachalar bilan "
                 "qo'shiladi.",
                 "Каждый столбец — с новой строки.\n"
                 "Номер по порядку и столбец источника добавляются "
                 "флажками ниже.",
                 "One column per line.\n"
                 "The row number and source columns are added with the "
                 "checkboxes below."),
    "sch_key": ("Dublikat qaysi ustun bo'yicha topilsin?",
                "По какому столбцу искать дубликаты?",
                "Which column identifies duplicates?"),
    "sch_no": ("Tartib raqami ustuni bo'lsin (№)",
               "Добавить столбец с номером по порядку (№)",
               "Add a row-number column (#)"),
    "sch_src": ("Manba ustuni bo'lsin (fayl kimdan olingani)",
                "Добавить столбец источника (от кого получен файл)",
                "Add a source column (who the file came from)"),
    "sch_none": ("— dublikat tekshirilmasin —", "— не искать дубликаты —",
                 "— do not check duplicates —"),
    "sch_err": ("Ustun yozilmadi", "Столбцы не указаны", "No columns given"),
    "sch_err_b": ("Kamida bitta ustun nomini yozing.",
                  "Укажите хотя бы один столбец.",
                  "Enter at least one column name."),
    "sch_mode": ("Solishtirish usuli", "Способ сравнения", "Comparison method"),
    "km_phone": ("telefon raqami sifatida", "как телефонный номер",
                 "as a phone number"),
    "km_text": ("matn sifatida", "как текст", "as text"),
    "km_number": ("raqam sifatida", "как число", "as a number"),

    # --- yon panel bo'limlari
    "sec_base": ("BAZA", "БАЗА", "DATABASE"),
    "sec_contacts": ("KONTAKTLAR", "КОНТАКТЫ", "CONTACTS"),
    "sec_output": ("CHIQARISH", "ЭКСПОРТ", "OUTPUT"),

    # --- yon panel tugmalari
    "nav_new": ("Yangi baza", "Новая база", "New database"),
    "nav_open": ("Eski bazani ochish", "Открыть базу", "Open database"),
    "nav_recent": ("Oxirgi fayllar", "Недавние файлы", "Recent files"),
    "nav_add": ("Kontakt qo'shish", "Добавить контакты", "Add contacts"),
    "nav_manual": ("Qo'lda qo'shish", "Добавить вручную", "Add manually"),
    "nav_save": ("Saqlash", "Сохранить", "Save"),
    "nav_saveas": ("Boshqa nom bilan", "Сохранить как", "Save as"),
    "nav_excel": ("Excelda ochish", "Открыть в Excel", "Open in Excel"),
    "nav_lang": ("Til", "Язык", "Language"),
    "nav_dark": ("Tungi rejim", "Тёмная тема", "Dark mode"),
    "nav_light": ("Kunduzgi rejim", "Светлая тема", "Light mode"),

    # --- jadval sarlavhalari (Excelga ham shular yoziladi)
    "h_no": ("№", "№", "#"),
    "h_name": ("Ism Familiya", "Имя Фамилия", "Full Name"),
    "h_phone": ("Telefon", "Телефон", "Phone"),
    "h_id": ("ID", "ID", "ID"),
    "h_source": ("Manba (kimdan olindi)", "Источник (от кого получено)",
                 "Source (from whom)"),

    # --- ko'rsatkich kartalari
    "card_total": ("Jami qator", "Всего строк", "Total rows"),
    "card_dup": ("Dublikat", "Дубликаты", "Duplicates"),
    "card_src": ("Manba raqam", "Источников", "Sources"),
    "card_view": ("Ko'rinmoqda", "Показано", "Shown"),

    # --- qidiruv va filtrlar
    "search_ph": ("Ism, raqam yoki manba…", "Имя, номер или источник…",
                  "Name, number or source…"),
    "only_dups": ("Faqat dublikatlar", "Только дубликаты", "Duplicates only"),
    "last9": ("Oxirgi 9 raqam bo'yicha", "По последним 9 цифрам",
              "Match last 9 digits"),
    "btn_del_bar": ("  O'chirish  ", "  Удалить  ", "  Delete  "),
    "btn_src_bar": ("  Manbani o'zgartirish  ", "  Изменить источник  ",
                    "  Change source  "),

    # --- bo'sh holat
    "empty_title": ("Baza hozircha bo'sh", "База пока пуста",
                    "The database is empty"),
    "empty_text": (
        "Chapdagi «Yangi baza» tugmasi bilan yangi ro'yxat boshlang\n"
        "yoki «Eski bazani ochish» orqali mavjud Excel faylni yuklang.",
        "Начните новый список кнопкой «Новая база» слева\n"
        "или загрузите готовый файл через «Открыть базу».",
        "Start a new list with “New database” on the left,\n"
        "or load an existing file with “Open database”."),

    # --- yuqori sarlavha va pastki qator
    "new_base_label": ("Yangi baza", "Новая база", "New database"),
    "not_saved": ("hali saqlanmagan", "ещё не сохранено", "not saved yet"),
    "keys_hint": (
        "Ctrl+S saqlash   ·   Ctrl+Z ortga   ·   Ctrl+C nusxa   ·   "
        "ikki marta bosib tahrirlash",
        "Ctrl+S сохранить   ·   Ctrl+Z отменить   ·   Ctrl+C копировать   ·   "
        "двойной клик — правка",
        "Ctrl+S save   ·   Ctrl+Z undo   ·   Ctrl+C copy   ·   "
        "double-click to edit"),
    "hint_start": ("Boshlash uchun baza tanlang", "Выберите базу, чтобы начать",
                   "Choose a database to start"),
    "hint_new": ("Yangi baza — raqamlash 1 dan boshlanadi",
                 "Новая база — нумерация начнётся с 1",
                 "New database — numbering starts at 1"),
    "hint_next": ("Keyingi qator raqami: {n}", "Номер следующей строки: {n}",
                  "Next row number: {n}"),
    "hint_added": ("Yangi qatorlar yashil rangda  ·  "
                   "Ctrl+Z bilan ortga qaytarish mumkin",
                   "Новые строки выделены зелёным  ·  Ctrl+Z отменит",
                   "New rows are green  ·  Ctrl+Z undoes it"),
    "hint_saved": ("Saqlandi: {path}", "Сохранено: {path}", "Saved: {path}"),

    # --- umumiy tugmalar
    "btn_ok": ("OK", "ОК", "OK"),
    "btn_close": ("Yopish", "Закрыть", "Close"),
    "btn_cancel": ("Bekor", "Отмена", "Cancel"),
    "btn_save": ("Saqlash", "Сохранить", "Save"),
    "btn_add": ("Qo'shish", "Добавить", "Add"),
    "btn_delete": ("O'chirish", "Удалить", "Delete"),
    "dlg_msg": ("Xabar", "Сообщение", "Message"),
    "dlg_err": ("Xato", "Ошибка", "Error"),

    # --- saqlanmagan o'zgarishlar
    "unsaved_title": ("Saqlanmagan o'zgarishlar", "Несохранённые изменения",
                      "Unsaved changes"),
    "unsaved_head": ("Bazada saqlanmagan o'zgarishlar bor",
                     "В базе есть несохранённые изменения",
                     "The database has unsaved changes"),
    "unsaved_body": ("Davom etishdan oldin saqlaymizmi?",
                     "Сохранить перед продолжением?",
                     "Save before continuing?"),
    "btn_dont_save": ("Saqlamasdan", "Не сохранять", "Don't save"),

    # --- fayl tanlash oynalari
    "fd_open": ("Eski Excel bazani tanlang", "Выберите файл базы",
                "Choose an existing database"),
    "fd_add": ("Kontaktlar faylini tanlang (bir nechta bo'lishi mumkin)",
               "Выберите файлы контактов (можно несколько)",
               "Choose contact files (multiple allowed)"),
    "fd_save": ("Bazani saqlash", "Сохранить базу", "Save database"),
    "ft_all": ("Barcha fayllar", "Все файлы", "All files"),
    "ft_contacts": ("Kontakt fayllari", "Файлы контактов", "Contact files"),
    "ft_text": ("Matnli fayl", "Текстовый файл", "Text file"),
    "ft_vcard": ("Telefon eksporti (vCard)", "Экспорт с телефона (vCard)",
                 "Phone export (vCard)"),
    "ft_xlsx": ("Excel fayl", "Файл Excel", "Excel file"),

    # --- kutish oynasi
    "busy_load": ("Yuklanmoqda:  {name}", "Загрузка:  {name}",
                  "Loading:  {name}"),
    "busy_read": ("O'qilmoqda ({n}/{total}):  {name}",
                  "Чтение ({n}/{total}):  {name}",
                  "Reading ({n}/{total}):  {name}"),

    # --- manba raqamini so'rash
    "src_title": ("Manba raqami", "Номер источника", "Source number"),
    "src_head": ("Bu kontaktlar kimning raqamidan olindi?",
                 "С чьего номера получены эти контакты?",
                 "Whose number did these contacts come from?"),
    "src_body": (
        "Fayl:  {file}\nBu fayldagi kontaktlar:  {n} ta\n\n"
        "Yozgan raqamingiz shu fayldan keladigan {n} ta qatorning "
        "oxirgi «Manba» ustuniga qo'yiladi.",
        "Файл:  {file}\nКонтактов в файле:  {n}\n\n"
        "Введённый номер будет записан в столбец «Источник» "
        "для всех {n} строк из этого файла.",
        "File:  {file}\nContacts in this file:  {n}\n\n"
        "The number you enter goes into the “Source” column "
        "for all {n} rows from this file."),
    "srcempty_title": ("Manba bo'sh", "Источник не указан", "Source is empty"),
    "srcempty_head": ("Manba raqami yozilmadi", "Номер источника не введён",
                      "No source number entered"),
    "srcempty_body": (
        "«Manba» ustuni bo'sh qoladi — keyin kimning kontakti ekanini "
        "bilib bo'lmaydi.\nBaribir davom etamizmi?",
        "Столбец «Источник» останется пустым — потом будет не понять, "
        "чьи это контакты.\nВсё равно продолжить?",
        "The “Source” column will stay empty — you won't know whose "
        "contacts these are later.\nContinue anyway?"),
    "btn_back": ("Orqaga qaytish", "Вернуться", "Go back"),
    "btn_leave_empty": ("Bo'sh qoldirish", "Оставить пустым", "Leave empty"),

    # --- dublikatlar
    "dup_title": ("Dublikat raqamlar", "Дублирующиеся номера",
                  "Duplicate numbers"),
    "dup_head": ("{n} ta takrorlangan raqam topildi", "Найдено дубликатов: {n}",
                 "{n} duplicate numbers found"),
    "dup_body": (
        "«{file}» faylidagi ba'zi raqamlar bazada allaqachon bor.\n\n"
        "Qo'shish — hammasi qo'shiladi, dublikatlar qizil belgilanadi\n"
        "Tashlab ketish — faqat yangi raqamlar qo'shiladi",
        "Некоторые номера из файла «{file}» уже есть в базе.\n\n"
        "Добавить — добавятся все, дубликаты выделятся красным\n"
        "Пропустить — добавятся только новые номера",
        "Some numbers in “{file}” are already in the database.\n\n"
        "Add — everything is added, duplicates marked red\n"
        "Skip — only new numbers are added"),
    "btn_skip": ("Tashlab ketish", "Пропустить", "Skip"),

    # --- xatolar
    "err_read": ("Faylni o'qib bo'lmadi", "Не удалось прочитать файл",
                 "Could not read the file"),
    "err_file_read": ("«{file}» o'qilmadi", "«{file}» не прочитан",
                      "“{file}” could not be read"),
    "err_file_empty": ("«{file}» bo'sh", "«{file}» пуст", "“{file}” is empty"),
    "err_file_empty_b": ("Faylda kontakt qatorlari topilmadi.",
                         "В файле не найдено строк с контактами.",
                         "No contact rows found in the file."),
    "err_norow": ("Qator tanlanmagan", "Строка не выбрана", "No row selected"),
    "err_norow_multi": ("Avval jadvaldan bir yoki bir nechta qatorni tanlang.",
                        "Сначала выберите одну или несколько строк в таблице.",
                        "Select one or more rows in the table first."),
    "err_norow_one": ("Avval jadvaldan qatorni tanlang.",
                      "Сначала выберите строку в таблице.",
                      "Select a row in the table first."),
    "err_norow_any": ("Avval jadvaldan qator(lar)ni tanlang.",
                      "Сначала выберите строку или строки.",
                      "Select row(s) in the table first."),
    "err_empty_base": ("Baza bo'sh", "База пуста", "The database is empty"),
    "err_empty_base_b": ("Saqlash uchun hech qanday qator yo'q.",
                         "Нечего сохранять — нет ни одной строки.",
                         "There are no rows to save."),
    "err_busy": ("Fayl band", "Файл занят", "File is in use"),
    "err_busy_b": ("Fayl Excelda ochiq bo'lsa kerak.\n"
                   "Excelni yoping va qaytadan saqlang.\n\n"
                   "Eski fayl o'zgarmadi.",
                   "Похоже, файл открыт в Excel.\n"
                   "Закройте Excel и сохраните ещё раз.\n\n"
                   "Старый файл не изменён.",
                   "The file is probably open in Excel.\n"
                   "Close Excel and save again.\n\n"
                   "The old file is unchanged."),
    "err_save": ("Saqlab bo'lmadi", "Не удалось сохранить", "Could not save"),
    "err_save_b": ("{err}\n\nEski fayl o'zgarmadi.",
                   "{err}\n\nСтарый файл не изменён.",
                   "{err}\n\nThe old file is unchanged."),
    "err_notfound": ("Fayl topilmadi", "Файл не найден", "File not found"),
    "err_notfound_b": ("Avval bazani saqlang.", "Сначала сохраните базу.",
                       "Save the database first."),
    "err_open": ("Ochib bo'lmadi", "Не удалось открыть", "Could not open"),
    "err_unexpected": ("Kutilmagan xato", "Непредвиденная ошибка",
                       "Unexpected error"),
    "err_unexpected_b": ("{name}: {msg}\n\nBatafsil ma'lumot shu faylda:\n{log}",
                         "{name}: {msg}\n\nПодробности в файле:\n{log}",
                         "{name}: {msg}\n\nDetails are in this file:\n{log}"),

    # --- tahrirlash
    "edit_src_title": ("Manbani o'zgartirish", "Изменить источник",
                       "Change source"),
    "edit_src_head": ("Tanlangan {n} ta qator uchun manba",
                      "Источник для выбранных строк ({n})",
                      "Source for {n} selected row(s)"),
    "edit_src_body": ("Yangi qiymatni kiriting:", "Введите новое значение:",
                      "Enter the new value:"),
    "edit_row_title": ("Qatorni tahrirlash", "Редактировать строку", "Edit row"),
    "edit_row_head": ("{no}-qator", "Строка {no}", "Row {no}"),
    "edit_row_body": ("Qiymatlarni o'zgartiring:", "Измените значения:",
                      "Change the values:"),
    "add_title": ("Yangi kontakt", "Новый контакт", "New contact"),
    "add_head": ("Yangi kontakt qo'shish", "Добавить новый контакт",
                 "Add a new contact"),
    "add_body": ("Kamida ism yoki telefon to'ldirilishi kerak:",
                 "Нужно заполнить хотя бы имя или телефон:",
                 "At least a name or a phone is required:"),
    "add_empty": ("Bo'sh kontakt", "Пустой контакт", "Empty contact"),
    "add_empty_b": ("Hech bo'lmasa ism yoki telefon yozilishi kerak.",
                    "Нужно указать хотя бы имя или телефон.",
                    "You must enter at least a name or a phone."),

    # --- o'chirish
    "del_title": ("O'chirish", "Удаление", "Delete"),
    "del_head": ("{n} ta qator o'chirilsinmi?", "Удалить строк: {n}?",
                 "Delete {n} row(s)?"),
    "del_body": ("Xato bo'lsa Ctrl+Z bilan qaytarib olsangiz bo'ladi.",
                 "Если ошиблись — Ctrl+Z вернёт.",
                 "If it's a mistake, Ctrl+Z brings them back."),

    # --- qisqa xabarlar
    "t_loaded": ("{n} ta qator yuklandi", "Загружено строк: {n}",
                 "{n} rows loaded"),
    "t_added": ("{n} ta qator qo'shildi", "Добавлено строк: {n}",
                "{n} rows added"),
    "t_skipped": (", {n} ta dublikat tashlandi", ", пропущено дубликатов: {n}",
                  ", {n} duplicates skipped"),
    "t_src_changed": ("{n} ta qatorning manbasi o'zgartirildi",
                      "Источник изменён у строк: {n}",
                      "Source changed for {n} row(s)"),
    "t_row_updated": ("Qator yangilandi", "Строка обновлена", "Row updated"),
    "t_contact_added": ("Kontakt qo'shildi", "Контакт добавлен",
                        "Contact added"),
    "t_deleted": ("{n} ta qator o'chirildi", "Удалено строк: {n}",
                  "{n} rows deleted"),
    "t_saved": ("{n} ta qator saqlandi", "Сохранено строк: {n}",
                "{n} rows saved"),
    "t_copied": ("{n} ta qator nusxa olindi", "Скопировано строк: {n}",
                 "{n} rows copied"),
    "t_select_first": ("Avval qator tanlang", "Сначала выберите строку",
                       "Select a row first"),
    "t_recent_empty": ("Oxirgi fayllar ro'yxati hozircha bo'sh",
                       "Список недавних файлов пуст",
                       "The recent files list is empty"),
    "t_recent_cleared": ("Ro'yxat tozalandi", "Список очищен", "List cleared"),
    "menu_clear": ("Ro'yxatni tozalash", "Очистить список", "Clear the list"),

    # --- vositalar menyusi
    "nav_tools": ("Vositalar", "Инструменты", "Tools"),
    "tool_paste": ("Buferdan qo'yish   (Ctrl+V)",
                   "Вставить из буфера   (Ctrl+V)",
                   "Paste from clipboard   (Ctrl+V)"),
    "tool_normalize": ("Telefon formatini birxillashtirish",
                       "Привести телефоны к одному виду",
                       "Normalize phone format"),
    "tool_problems": ("Faqat shubhali qatorlar",
                      "Только подозрительные строки",
                      "Suspicious rows only"),
    "tool_csv": ("CSV ga eksport", "Экспорт в CSV", "Export to CSV"),
    "tool_vcard": ("vCard ga eksport  (telefon uchun)",
                   "Экспорт в vCard  (для телефона)",
                   "Export to vCard  (for phone)"),

    # --- buferdan qo'yish
    "clip_label": ("buferdan", "из буфера", "clipboard"),
    "err_clip": ("Buferda kontakt yo'q", "В буфере нет контактов",
                 "No contacts in the clipboard"),
    "err_clip_b": ("Exceldan yoki matndan qatorlarni nusxa oling "
                   "va qaytadan urinib ko'ring.",
                   "Скопируйте строки из Excel или из текста и повторите.",
                   "Copy rows from Excel or from text and try again."),

    # --- manbani qayta ishlatish
    "apply_all": ("Qolgan fayllarga ham shu raqam",
                  "Этот же номер для остальных файлов",
                  "Use this number for the remaining files"),

    # --- telefon formati va shubhali qatorlar
    "t_normalized": ("{n} ta raqam formati o'zgartirildi",
                     "Формат изменён у номеров: {n}",
                     "{n} phone numbers reformatted"),
    "t_normalized_none": ("Hamma raqam allaqachon bir xil ko'rinishda",
                          "Все номера уже в едином виде",
                          "All numbers are already in the same format"),
    "u_normalize": ("telefon formati birxillashtirildi",
                    "формат телефонов приведён к одному виду",
                    "phone format normalized"),
    "t_problems": ("{n} ta shubhali qator", "Подозрительных строк: {n}",
                   "{n} suspicious rows"),
    "t_no_problems": ("Shubhali qator topilmadi", "Подозрительных строк нет",
                      "No suspicious rows found"),

    # --- eksport
    "fd_csv": ("CSV ga saqlash", "Сохранить в CSV", "Save as CSV"),
    "fd_vcard": ("vCard ga saqlash", "Сохранить в vCard", "Save as vCard"),
    "ft_csv": ("CSV fayl", "Файл CSV", "CSV file"),
    "ft_vcf": ("vCard fayl", "Файл vCard", "vCard file"),
    "t_exported": ("{n} ta kontakt eksport qilindi",
                   "Экспортировано контактов: {n}", "{n} contacts exported"),
    "busy_save": ("Saqlanmoqda:  {name}", "Сохранение:  {name}",
                  "Saving:  {name}"),

    # --- bo'sh ekrandagi oxirgi fayllar
    "empty_recent": ("Oxirgi ochilganlar:", "Недавно открытые:",
                     "Recently opened:"),

    # --- ortga qaytarish
    "undo_word": ("Ortga qaytarildi", "Отменено", "Undone"),
    "redo_word": ("Qaytadan bajarildi", "Возвращено", "Redone"),
    "undo_empty": ("Ortga qaytariladigan amal yo'q", "Нечего отменять",
                   "Nothing to undo"),
    "redo_empty": ("Qaytadan bajariladigan amal yo'q", "Нечего возвращать",
                   "Nothing to redo"),
    "u_files": ("{n} ta fayl qo'shildi", "добавлено файлов: {n}",
                "{n} files added"),
    "u_src": ("{n} ta qatorning manbasi o'zgartirildi",
              "изменён источник у {n} строк", "source changed for {n} rows"),
    "u_row": ("{no}-qator tahrirlandi", "строка {no} изменена",
              "row {no} edited"),
    "u_manual": ("qo'lda kontakt qo'shildi", "контакт добавлен вручную",
                 "contact added manually"),
    "u_del": ("{n} ta qator o'chirildi", "удалено строк: {n}",
              "{n} rows deleted"),
}


def t(key, **kw):
    """Return the text in the current language, or the key if it is unknown."""
    row = TR.get(key)
    if row is None:
        return key
    try:
        text = row[LANG_ORDER.index(LANG)]
    except (ValueError, IndexError):
        text = row[0]
    return text.format(**kw) if kw else text


# ===================================================================== schema
#
# A schema describes the shape of a database. The application is not tied to
# contacts: the columns come from here, so any kind of list works.
#
#   cols     - the columns. Each one holds:
#                key   - internal name (stable, also written to the file)
#                tkey  - translation key (for the built-in templates)
#                title - literal text (for user-defined columns)
#                width - width in the table
#                role  — "no" | "name" | "phone" | "id" | "src" | ""
#   key_col  - column duplicates are matched on (None disables the check)
#   key_mode — "phone" | "text" | "number"

SCHEMA_SHEET = "_schema"          # hidden sheet the schema is stored in


def col(key, width, role="", tkey=None, title=None):
    return {"key": key, "width": width, "role": role,
            "tkey": tkey, "title": title}


TEMPLATES = {
    "kontakt": {
        "tkey": "tpl_contacts",
        "cols": [col("no", 55, "no", "h_no"),
                 col("name", 270, "name", "h_name"),
                 col("phone", 190, "phone", "h_phone"),
                 col("id", 130, "id", "h_id"),
                 col("source", 220, "src", "h_source")],
        "key_col": "phone", "key_mode": "phone",
    },
    "mahsulot": {
        "tkey": "tpl_products",
        "cols": [col("no", 55, "no", "h_no"),
                 col("name", 260, "name", "h_pname"),
                 col("code", 160, "id", "h_code"),
                 col("price", 110, "", "h_price"),
                 col("qty", 90, "", "h_qty"),
                 col("source", 200, "src", "h_source")],
        "key_col": "code", "key_mode": "text",
    },
    "talaba": {
        "tkey": "tpl_students",
        "cols": [col("no", 55, "no", "h_no"),
                 col("name", 260, "name", "h_name"),
                 col("group", 120, "", "h_group"),
                 col("phone", 180, "phone", "h_phone"),
                 col("id", 130, "id", "h_id"),
                 col("source", 200, "src", "h_source")],
        "key_col": "phone", "key_mode": "phone",
    },
}


def default_schema():
    """The default schema: contacts."""
    return copy.deepcopy(TEMPLATES["kontakt"]) | {"name": "kontakt"}


def col_title(c):
    """Column header. Built-in templates are translated, custom ones are not."""
    if c.get("tkey"):
        return t(c["tkey"])
    return c.get("title") or c["key"]


def schema_title(schema):
    """Display name of the schema, shown in the sidebar."""
    if schema.get("tkey"):
        return t(schema["tkey"])
    return t("tpl_opened")


def schema_fields(schema):
    return [c["key"] for c in schema["cols"]]


def headers(schema):
    """Column headers for the table and for Excel, in the current language."""
    return [col_title(c) for c in schema["cols"]]


def role_col(schema, role):
    """Key of the first column with the given role, or None."""
    for c in schema["cols"]:
        if c.get("role") == role:
            return c["key"]
    return None


def data_cols(schema, with_src=False):
    """
    The columns read from a file. The row-number column is never read, it
    is regenerated. The source column depends on the situation:
      - ADDING a contact file: not read, the user types the number
      - OPENING our own database: read, otherwise the value would be lost
    """
    skip = ("no",) if with_src else ("no", "src")
    return [c for c in schema["cols"] if c.get("role") not in skip]


def blank_row(schema):
    return {c["key"]: "" for c in schema["cols"]}


def row_key(row, schema, last9=True):
    """Comparison key for duplicate detection; empty when the schema has none."""
    kc = schema.get("key_col")
    if not kc:
        return ""
    value = row.get(kc)
    mode = schema.get("key_mode", "text")
    if mode == "phone":
        return norm_phone(value, last9)
    if mode == "number":
        return re.sub(r"\D", "", cell_text(value))
    return cell_text(value).casefold()


def title_variants(c):
    """The header text in all three languages, used to recognise a file."""
    if c.get("tkey") and c["tkey"] in TR:
        return {s.strip().lower() for s in TR[c["tkey"]] if s.strip()}
    return {cell_text(c.get("title") or c["key"]).lower()}


APP_NAME = "Database Manager"

# settings and the error log live here
_APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
CONFIG_DIR = os.path.join(_APPDATA, "DatabaseManager")
OLD_CONFIG_DIR = os.path.join(_APPDATA, "BazaMenejeri")   # a previous name
CONFIG_PATH = os.path.join(CONFIG_DIR, "settings.json")
LOG_PATH = os.path.join(CONFIG_DIR, "errors.log")

RECENT_LIMIT = 8            # length of the recent-files list
SOURCE_LIMIT = 20           # eslab qolinadigan manba raqamlari soni
UNDO_LIMIT = 30             # changed amalni ortga qaytarish mumkin
UNDO_ROW_BUDGET = 300_000   # undo tarixida total changed row saqlanadi (xotira)

FIELDS = ["no", "name", "phone", "id", "source"]

# Keywords used to recognise column headers automatically.
# All three languages must be recognised, so a database saved in one
# language still lines up when opened in another.
KEYS_SOURCE = ("manba", "kimdan", "source", "olindi", "egasi",
               "источник", "от кого", "получено")
KEYS_PHONE = ("telefon", "tel", "phone", "nomer", "raqam", "number",
              "mobil", "телефон", "номер")
KEYS_NAME = ("ism", "familiya", "fio", "ф.и.о", "фио", "имя", "name",
             "фамилия")
KEYS_ID = ("id", "идент")
KEYS_NO = ("№", "no", "n", "t/r", "tr", "order", "#")

# pulls a phone number out of a line of free text
PHONE_RE = re.compile(r"(?<!\d)(\+?\d[\d\s\-()]{6,}\d)(?!\d)")
LETTER_RE = re.compile(r"[^\W\d_]", re.UNICODE)

TEXT_EXT = (".txt", ".text", ".log", ".tsv", ".tab")
VCARD_EXT = (".vcf", ".vcard")
TABLE_EXT = (".xlsx", ".xlsm", ".csv")

# --- Segoe icon glyphs. Present in both the Windows 11 font (SegoeIcons)
IC_LOGO = "\uE779"      # contact list
IC_NEW = "\uE7C3"      # updated hujjat
IC_OPEN = "\uE838"      # folder ochish
IC_ADD = "\uE710"      # plyus
IC_SAVE = "\uE74E"      # saqlash
IC_SAVEAS = "\uE792"      # boshqa name bilan saqlash
IC_EXTERN = "\uE8A7"      # tashqi dasturda ochish
IC_EDIT = "\uE70F"      # qalam
IC_DELETE = "\uE74D"      # savat
IC_SEARCH = "\uE721"      # lupa
IC_SUN = "\uE706"      # kunduzgi tema
IC_MOON = "\uE708"      # tungi tema
IC_PHONE = "\uE717"      # telefon
IC_WARN = "\uE7BA"      # ogohlantirish
IC_INFO = "\uE946"      # ma'lumot
IC_ERROR = "\uE783"      # err
IC_CHECK = "\uE73E"      # bajarildi
IC_EMPTY = "\uE8A1"      # empty list
IC_ROWS = "\uE8FD"      # rows soni
IC_PEOPLE = "\uE716"      # sources
IC_FILTER = "\uE71C"      # shown rows
IC_FOLDER = "\uE8B7"      # folder
IC_UNDO = "\uE7A7"      # ortga qaytarish
IC_HISTORY = "\uE81C"      # recent files
IC_PERSONADD = "\uE8FA"      # add a row by hand
IC_COPY = "\uE8C8"      # nusxa olish
IC_GLOBE = "\uE774"      # til tanlash
IC_TOOLS = "\uEC7A"      # tools

# colours written into Excel (identical in both themes)
XL_DUP_BG = "FFC7CE"
XL_DUP_FG = "9C0006"
XL_HEAD_BG = "2F5597"

PALETTES = {
    "light": {
        "bg":          "#f3f3f3",   # asosiy maydon foni
        "sidebar":     "#fbfbfb",   # yon panel
        "surface":     "#ffffff",   # cards, jadval
        "text":        "#1a1a1a",
        "muted":       "#616161",
        "faint":       "#8a8a8a",
        "border":      "#e5e5e5",
        "accent":      "#0f6cbd",
        "accent_soft": "#eaf3fb",
        "accent_dim":  "#0c5a9e",
        "on_accent":   "#ffffff",
        "nav_hover":   "#ececec",
        "stripe":      "#f7f8fa",
        "dup_bg":      "#ffe2e3",
        "dup_fg":      "#b42318",
        "dup_soft":    "#fdeeee",
        "new_bg":      "#e4f6e9",
        "new_fg":      "#166534",
        "new_soft":    "#eefaf1",
        "bad_bg":      "#fdf0d8",   # shubhali row (qisqa raqam, harf...)
        "bad_fg":      "#8a5300",
    },
    "dark": {
        "bg":          "#191919",
        "sidebar":     "#202020",
        "surface":     "#242424",
        "text":        "#f5f5f5",
        "muted":       "#a3a3a3",
        "faint":       "#7d7d7d",
        "border":      "#2f2f2f",
        "accent":      "#4cc2ff",
        "accent_soft": "#12313f",
        "accent_dim":  "#3ba9e0",
        "on_accent":   "#08202b",
        "nav_hover":   "#2d2d2d",
        "stripe":      "#282828",
        "dup_bg":      "#4a2124",
        "dup_fg":      "#ffb4ab",
        "dup_soft":    "#331a1c",
        "new_bg":      "#1e3a26",
        "new_fg":      "#a6e5b8",
        "new_soft":    "#172a1c",
        "bad_bg":      "#3d3016",
        "bad_fg":      "#f0c476",
    },
}


# ==================================================== settings and error log

def load_config():
    """Read the saved settings. Returns {} when the file is missing or broken."""
    path = CONFIG_PATH
    if not os.path.exists(path):
        # the application had other names before; keep the old settings
        old = os.path.join(OLD_CONFIG_DIR, os.path.basename(CONFIG_PATH))
        if os.path.exists(old):
            path = old

    try:
        # utf-8-sig so a file written with a BOM is still read
        with open(path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(data):
    """Write the settings; an interrupted write cannot corrupt the old file."""
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        tmp = CONFIG_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CONFIG_PATH)
    except Exception:
        pass                              # the app runs even if settings do not save


def log_error(exc_type, exc, tb):
    """
    Write the error to a file. The app starts through `pythonw`, so there
    is no console: without this log an error would be entirely invisible.
    """
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 62 + "\n")
            f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
            traceback.print_exception(exc_type, exc, tb, file=f)
    except Exception:
        pass


def system_lang():
    """Pick the initial language from the Windows UI language (first run)."""
    try:
        lid = ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF
    except Exception:
        return "uz"
    return {0x43: "uz", 0x19: "ru", 0x09: "en"}.get(lid, "uz")


def find_icon_file(name="icon.ico"):
    """Locate the icon file, both inside the .exe bundle and next to the script."""
    here = os.path.dirname(os.path.abspath(__file__))
    for folder in (getattr(sys, "_MEIPASS", None), here):
        if folder:
            p = os.path.join(folder, name)
            if os.path.exists(p):
                return p
    return None


def message_box(text, title=APP_NAME, icon=0x10):
    """Show a message through the Windows API, for when Tk is not up yet."""
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, icon)
    except Exception:
        pass


# ============================================================ helper functions

def cell_text(v):
    """Turn an Excel cell into clean text (998901234567.0 -> 998901234567)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v).strip()


def norm_phone(raw, last9=True):
    """Normalise a phone number for comparison."""
    digits = re.sub(r"\D", "", cell_text(raw))
    if not digits:
        return ""
    if last9 and len(digits) >= 9:
        return digits[-9:]
    return digits


def pretty_phone(raw):
    """
    Bring a number into one shape:
        901112233, 998901112233, +998 90-111-22-33  ->  +998 90 111 22 33
    Anything that does not look Uzbek only gets its punctuation stripped.
    """
    text = cell_text(raw)
    digits = re.sub(r"\D", "", text)
    if not digits:
        return text

    if len(digits) == 9:                     # no country code, assume local
        digits = "998" + digits
    if digits.startswith("998") and len(digits) == 12:
        return (f"+{digits[:3]} {digits[3:5]} {digits[5:8]} "
                f"{digits[8:10]} {digits[10:]}")
    if len(digits) >= 10:                    # chet el raqami
        return "+" + digits
    return text                              # juda qisqa — tegmaymiz


def row_problem(row, schema):
    """
    Return why a row deserves attention, or None when it is fine.
    Which checks apply depends on the schema.
    """
    pc = role_col(schema, "phone")
    if pc:
        phone = cell_text(row.get(pc))
        digits = re.sub(r"\D", "", phone)
        if not digits:
            return "phone_yoq"
        if LETTER_RE.search(phone):  # harf uzunlikdan before — sabab aniqroq
            return "harf"
        if len(digits) < 9:
            return "qisqa"

    nc = role_col(schema, "name")
    if nc and not cell_text(row.get(nc)):
        return "ism_yoq"

    kc = schema.get("key_col")       # a row whose duplicate key is empty
    if kc and not cell_text(row.get(kc)):
        return "kalit_yoq"
    return None


def read_table(path):
    """Read an xlsx / xlsm / csv file as a raw list of rows."""
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows

    if ext == ".csv":
        for enc in ("utf-8-sig", "cp1251", "latin-1"):
            try:
                with open(path, "r", newline="", encoding=enc) as f:
                    sample = f.read(8192)
                    f.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                    except csv.Error:
                        dialect = csv.excel
                    return [r for r in csv.reader(f, dialect)]
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not determine the encoding of the CSV file.")

    if ext == ".xls":
        raise ValueError("The old .xls format is not supported.\n"
                         "Open it in Excel and save it as .xlsx.")

    raise ValueError("Only .xlsx, .xlsm and .csv files are supported.")


def match_header(row, schema, with_src=False):
    """
    Sarlavha qatorini sxema ustunlariga solishtiradi.
    Qaytaradi: {ustun_kaliti: ustun_raqami} yoki topilmasa None.
    """
    cols = data_cols(schema, with_src)
    variants = {c["key"]: title_variants(c) for c in cols}

    mapping = {}
    for idx, val in enumerate(row or []):
        text = cell_text(val).lower().rstrip(":").strip()
        if not text:
            continue
        for c in cols:
            if c["key"] in mapping:
                continue
            if text in variants[c["key"]]:
                mapping[c["key"]] = idx
                break

    needed = 1 if len(cols) == 1 else 2
    return mapping if len(mapping) >= needed else None


def semantic_header(row, schema, with_src=False):
    """
    When the header names do not match the schema, fall back to keywords
    ("telefon", "имя", "name", "манба" ...). Only useful when the schema
    actually has a column with that role.
    """
    roles = {}
    for idx, val in enumerate(row or []):
        text = cell_text(val).lower().rstrip(":").strip()
        if not text:
            continue
        if any(k in text for k in KEYS_SOURCE):
            roles.setdefault("src", idx)
        elif any(k in text for k in KEYS_PHONE):
            roles.setdefault("phone", idx)
        elif any(k in text for k in KEYS_NAME):
            roles.setdefault("name", idx)
        elif text in KEYS_ID or text.startswith("id"):
            roles.setdefault("id", idx)
        elif text in KEYS_NO:
            roles.setdefault("no", idx)

    if "phone" not in roles and "name" not in roles:
        return None

    mapping = {}
    for c in data_cols(schema, with_src):
        idx = roles.get(c.get("role"))
        if idx is not None:
            mapping[c["key"]] = idx
    return mapping or None


def infer_columns(table, schema, with_src=False):
    """
    With no header at all, work the columns out from their CONTENT: which
    column holds phone-like values, which holds names made of letters, which
    is a plain row number. Only meaningful when the schema has a phone or a
    name column.
    """
    ncols = max((len(r) for r in table), default=0)
    if not ncols:
        return {}

    stats = {}
    for c in range(ncols):
        vals = [cell_text(r[c]) for r in table if c < len(r) and cell_text(r[c])]
        if not vals:
            continue
        phone_ratio = sum(1 for v in vals
                          if len(re.sub(r"\D", "", v)) >= 7) / len(vals)
        letters = sum(len(LETTER_RE.findall(v)) for v in vals) / len(vals)
        seq_ratio = sum(1 for v in vals
                        if v.isdigit() and len(v) <= 4) / len(vals)
        stats[c] = (phone_ratio, letters, seq_ratio)

    roles = {}

    # phone: a column where at least half the values have 7+ digits
    cand = [(s[0], -c, c) for c, s in stats.items() if s[0] >= 0.5]
    if cand:
        roles["phone"] = max(cand)[2]

    # name: most letters on average (an ID like "u1001" does not qualify)
    cand = [(s[1], -c, c) for c, s in stats.items()
            if c != roles.get("phone") and s[1] >= 2]
    if cand:
        roles["name"] = max(cand)[2]

    # row number: a column of short whole numbers
    used = set(roles.values())
    cand = [c for c, s in stats.items() if c not in used and s[2] >= 0.8]
    if cand:
        roles["no"] = min(cand)

    # whatever is left: ID first, then source
    used = set(roles.values())
    rest = [c for c in sorted(stats) if c not in used]
    if rest:
        roles["id"] = rest[0]
    if len(rest) > 1:
        roles["src"] = rest[1]

    mapping = {}
    for c in data_cols(schema, with_src):
        idx = roles.get(c.get("role"))
        if idx is not None:
            mapping[c["key"]] = idx
    return mapping


def rows_from_table(table, schema, with_src=False):
    """
    Xom jadvalni sxema qatorlariga aylantiradi. Uch bosqich:
      1) header names that match the schema columns
      2) keywords or content (for phone, name ... columns)
      3) failing everything else, plain left-to-right order
    """
    if not table:
        return []

    cols = data_cols(schema, with_src)
    start, mapping = 0, None

    for i in range(min(5, len(table))):          # header upper qatorlarda
        m = (match_header(table[i], schema, with_src)
             or semantic_header(table[i], schema, with_src))
        if m:
            mapping, start = m, i + 1
            break

    if mapping is None:
        mapping = infer_columns(table, schema, with_src)
    if not mapping:                              # by position
        mapping = {c["key"]: i for i, c in enumerate(cols)}

    out = []
    for raw in table[start:]:
        if not raw:
            continue

        row = blank_row(schema)
        has_value = False
        for c in cols:
            idx = mapping.get(c["key"])
            if idx is None or idx >= len(raw):
                continue
            value = cell_text(raw[idx])
            row[c["key"]] = value
            if value:
                has_value = True

        if has_value:                            # skip entirely empty rows
            out.append(row)
    return out


# ----------------------------------------------------- text (.txt / .vcf) files

def read_text_lines(path):
    """Read any text file, detecting its encoding."""
    with open(path, "rb") as f:
        raw = f.read()

    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        encodings = ("utf-16",)
    elif raw[:3] == b"\xef\xbb\xbf":
        encodings = ("utf-8-sig",)
    else:
        encodings = ("utf-8", "cp1251", "cp1252", "latin-1")

    for enc in encodings:
        try:
            text = raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
        if "\x00" in text:            # wrong encoding picked
            continue
        return text.splitlines()

    raise ValueError("Could not determine the text encoding of the file.")


def clean_name(s):
    """Strip separators and a leading list number from a name."""
    s = re.sub(r"\s{2,}", " ", s).strip(" \t-–—:;,|.")
    m = re.match(r"^\d{1,4}\s*[.)\-:]?\s+(.*)$", s)
    if m and LETTER_RE.search(m.group(1)):
        s = m.group(1)
    return s.strip(" \t-–—:;,|.")


def parse_text(lines):
    """
    Turn a plain text list into a RAW table; mapping onto columns happens
    later, against the schema. Two shapes are understood:

      1) ajratgichli jadval   ->  Alisher Karimov | +998901112233 | u1001
      2) erkin text           ->  1. Alisher Karimov — +998 90 111 22 33
    """
    body = [ln for ln in lines if ln.strip()]
    if not body:
        return []

    # --- 1) a delimited table
    for d in ("\t", "|", ";", ","):
        if all(ln.count(d) >= 1 for ln in body):
            table = [[c.strip() for c in ln.split(d)] for ln in body]
            if max(len(r) for r in table) > 1:
                return table

    # --- 2) free text: pull the number out of each line
    table = []
    for ln in body:
        m = PHONE_RE.search(ln)
        if not m:
            continue
        phone = re.sub(r"[\s\-()]", "", m.group(1))
        name = clean_name(ln[:m.start()] + " " + ln[m.end():])
        table.append([name, phone])
    return table


def parse_vcard(lines):
    """Read a .vcf (vCard) file exported from a phone."""
    unfolded = []                      # long lines may have been folded
    for ln in lines:
        if ln[:1] in (" ", "\t") and unfolded:
            unfolded[-1] += ln[1:]
        else:
            unfolded.append(ln)

    out, cur = [], None                # har bir karta -> [ism, telefon, ID]
    for ln in unfolded:
        s = ln.strip()
        up = s.upper()
        if up.startswith("BEGIN:VCARD"):
            cur = ["", "", ""]
        elif up.startswith("END:VCARD"):
            if cur and (cur[0] or cur[1]):
                out.append(cur)
            cur = None
        elif cur is not None and ":" in s:
            key, val = s.split(":", 1)
            k = key.split(";")[0].upper()
            val = val.strip()
            if k == "FN" and val:
                cur[0] = val
            elif k == "N" and val and not cur[0]:
                parts = [p.strip() for p in val.split(";")[:2] if p.strip()]
                cur[0] = " ".join(reversed(parts))
            elif k == "TEL" and val and not cur[1]:
                cur[1] = val
            elif k in ("UID", "X-ID") and val and not cur[2]:
                cur[2] = val
    return out


def read_any(path):
    """Detect the file type and return a RAW table, not yet mapped to columns."""
    ext = os.path.splitext(path)[1].lower()

    if ext in VCARD_EXT:
        return parse_vcard(read_text_lines(path))
    if ext in TEXT_EXT:
        return parse_text(read_text_lines(path))
    if ext in TABLE_EXT:
        return read_table(path)
    if ext == ".xls":
        raise ValueError("The old .xls format is not supported.\n"
                         "Open it in Excel and save it as .xlsx.")

    # unknown extension: try to read it as text
    return parse_text(read_text_lines(path))


def load_rows(path, schema, with_src=False):
    """Read a file and map its rows onto the columns of the given schema."""
    return rows_from_table(read_any(path), schema, with_src)


# ------------------------------------------ carrying the schema with the file

def read_schema(path):
    """
    Work out the schema of a file:
      1) saved by us: restored exactly from the hidden sheet
      2) a foreign Excel file: columns are built from the header row
      3) no header either: None, and the caller keeps the current schema
    """
    if os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm"):
        return None

    try:
        wb = load_workbook(path, data_only=True)
        try:
            if SCHEMA_SHEET in wb.sheetnames:
                raw = wb[SCHEMA_SHEET]["A1"].value
                schema = json.loads(raw) if raw else None
                if isinstance(schema, dict) and schema.get("cols"):
                    return schema
            head = [cell_text(v) for v in
                    next(wb.active.iter_rows(values_only=True), ()) or ()]
        finally:
            wb.close()
    except Exception:
        return None

    return schema_from_header(head)


def schema_from_header(head):
    """Build a schema from the header row of a foreign Excel file."""
    names = [h for h in head if cell_text(h)]
    if len(names) < 2:
        return None

    cols, used = [], set()
    for i, name in enumerate(head):
        name = cell_text(name)
        if not name:
            continue
        low = name.lower().rstrip(":").strip()

        if low in KEYS_NO and not any(c["role"] == "no" for c in cols):
            role = "no"
        elif any(k in low for k in KEYS_SOURCE):
            role = "src"
        elif any(k in low for k in KEYS_PHONE):
            role = "phone"
        elif any(k in low for k in KEYS_NAME):
            role = "name"
        elif low in KEYS_ID or low.startswith("id"):
            role = "id"
        else:
            role = ""
        if role and role != "" and any(c["role"] == role for c in cols):
            role = ""                     # no two columns share one role

        key = re.sub(r"\W+", "_", low).strip("_") or f"c{i}"
        while key in used:
            key += "_"
        used.add(key)
        cols.append(col(key, 180, role, title=name))

    if not cols:
        return None

    # duplicate key: prefer a phone column, then an ID column, otherwise none
    key = next((c for c in cols if c["role"] == "phone"), None)
    mode = "phone"
    if key is None:
        key = next((c for c in cols if c["role"] == "id"), None)
        mode = "text"

    return {"name": "file", "cols": cols,
            "key_col": key["key"] if key else None, "key_mode": mode}


# ================================================================== appearance

def _colorref(hex_color):
    """#RRGGBB -> Windows COLORREF (0x00BBGGRR)."""
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return ctypes.c_int((b << 16) | (g << 8) | r)


def style_titlebar(win, pal, dark):
    """Tint the Windows 11 title bar to match the application."""
    if not sys.platform.startswith("win"):
        return
    try:
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        dwm = ctypes.windll.dwmapi

        def set_attr(attr, value):
            """DwmSetWindowAttribute; True when it succeeded."""
            try:
                return dwm.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), 4) == 0
            except Exception:
                return False

        # Dark title bar: attribute 20 on Windows 10 2004+ and Windows 11.
        # On older Windows 10 (1809 ... 1909) the same thing lived at 19.
        mode = ctypes.c_int(1 if dark else 0)
        if not set_attr(20, mode):
            set_attr(19, mode)

        # 34/35/36 tint the title bar. Windows 11 only; on Windows 10 they are
        # ignored and the ordinary system title bar is shown.
        for attr, color in ((35, pal["sidebar"]),      # header foni
                            (36, pal["text"]),         # header matni
                            (34, pal["border"])):      # ramka
            set_attr(attr, _colorref(color))
    except Exception:
        pass                                   # simply skipped on older Windows


class Art:
    """
    Renders icons and rounded rectangles with PIL. Without Pillow the
    application still runs, it just looks plainer.
    """

    FONT_PATHS = (r"C:\Windows\Fonts\SegoeIcons.ttf",   # Windows 11
                  r"C:\Windows\Fonts\segmdl2.ttf")      # Windows 10

    def __init__(self):
        self.ok = False
        self._cache = {}
        self._icon_font = None
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk
            self._Image, self._Draw = Image, ImageDraw
            self._Font, self._Tk = ImageFont, ImageTk
            self._resample = getattr(getattr(Image, "Resampling", Image),
                                     "LANCZOS")
            for p in self.FONT_PATHS:
                if os.path.exists(p):
                    self._icon_font = p
                    break
            self.ok = True
        except Exception:
            self.ok = False

    def icon(self, glyph, size=16, color="#000000"):
        """Render a Segoe icon glyph into a coloured image."""
        if not self.ok or not self._icon_font:
            return None
        key = ("i", glyph, size, color)
        if key in self._cache:
            return self._cache[key]
        try:
            box = size + 6
            img = self._Image.new("RGBA", (box, box), (0, 0, 0, 0))
            drw = self._Draw.Draw(img)
            fnt = self._Font.truetype(self._icon_font, size)
            drw.text((box / 2, box / 2), glyph, font=fnt, fill=color, anchor="mm")
            photo = self._Tk.PhotoImage(img)
        except Exception:
            return None
        self._cache[key] = photo
        return photo

    def round_rect(self, w, h, radius, fill, outline=None, width=1):
        """A rounded rectangle with smooth edges, used behind buttons and cards."""
        if not self.ok or w < 1 or h < 1:
            return None
        key = ("r", w, h, radius, fill, outline, width)
        if key in self._cache:
            return self._cache[key]
        try:
            ss = 3                                   # supersample for smooth edges
            img = self._Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
            drw = self._Draw.Draw(img)
            drw.rounded_rectangle([0, 0, w * ss - 1, h * ss - 1],
                                  radius=radius * ss, fill=fill,
                                  outline=outline,
                                  width=width * ss if outline else 0)
            photo = self._Tk.PhotoImage(img.resize((w, h), self._resample))
        except Exception:
            return None
        self._cache[key] = photo
        return photo

    def round_rect_flat(self, w, h, radius, fill, outline, key_color):
        """
        A rounded rectangle for the toast. Because the window uses colour-key
        transparency the edge pixels must not blend, so alpha is cut hard;
        otherwise a halo of the key colour shows around the edge.
        """
        if not self.ok or w < 1 or h < 1:
            return None
        ck = ("f", w, h, radius, fill, outline, key_color)
        if ck in self._cache:
            return self._cache[ck]
        try:
            ss = 3
            img = self._Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
            drw = self._Draw.Draw(img)
            drw.rounded_rectangle([0, 0, w * ss - 1, h * ss - 1],
                                  radius=radius * ss, fill=fill,
                                  outline=outline, width=ss)
            img = img.resize((w, h), self._resample)
            mask = img.getchannel("A").point(lambda v: 255 if v >= 128 else 0)
            flat = self._Image.new("RGB", (w, h), key_color)
            flat.paste(img.convert("RGB"), mask=mask)
            photo = self._Tk.PhotoImage(flat)
        except Exception:
            return None
        self._cache[ck] = photo
        return photo


class NavButton(tk.Canvas):
    """A rounded sidebar button that lights up under the pointer."""

    H, R = 38, 8

    def __init__(self, parent, app, glyph, text, command, width=220,
                 accent=False, height=None):
        self.H = height or NavButton.H
        super().__init__(parent, width=width, height=self.H,
                         highlightthickness=0, bd=0, takefocus=0)
        self.app, self.glyph, self.command = app, glyph, command
        self.accent, self.W = accent, width
        self._hover = self._down = False
        self._imgs = {}

        self._bg = self.create_image(0, 0, anchor="nw")
        self._ic = self.create_image(15, self.H // 2, anchor="w")
        self._tx = self.create_text(46, self.H // 2, anchor="w", text=text)

        for seq, fn in (("<Enter>", self._enter), ("<Leave>", self._leave),
                        ("<Button-1>", self._press),
                        ("<ButtonRelease-1>", self._release)):
            self.bind(seq, fn)
        self.configure(cursor="hand2")

    def _enter(self, _=None):
        self._hover = True
        self.restyle()

    def _leave(self, _=None):
        self._hover = self._down = False
        self.restyle()

    def _press(self, _=None):
        self._down = True
        self.restyle()

    def _release(self, _=None):
        was = self._down
        self._down = False
        self.restyle()
        if was and self._hover:
            self.command()

    def restyle(self):
        pal, art = self.app.pal, self.app.art
        self.configure(bg=pal["sidebar"])

        if self.accent:
            fill = pal["accent_dim"] if (self._hover or self._down) else pal["accent"]
            fg = pal["on_accent"]
        else:
            fill = pal["nav_hover"] if (self._hover or self._down) else pal["sidebar"]
            fg = pal["text"]

        self._imgs["bg"] = art.round_rect(self.W, self.H, self.R, fill)
        if self._imgs["bg"]:
            self.itemconfigure(self._bg, image=self._imgs["bg"])
        self._imgs["ic"] = art.icon(self.glyph, 17, fg)
        if self._imgs["ic"]:
            self.itemconfigure(self._ic, image=self._imgs["ic"])
        self.itemconfigure(self._tx, fill=fg,
                           font=(self.app.f_ui, 10, "bold" if self.accent else ""))

    def set_glyph(self, glyph, text=None):
        self.glyph = glyph
        if text is not None:
            self.itemconfigure(self._tx, text=text)
        self.restyle()


class StatCard(tk.Canvas):
    """A statistic card: icon, caption and a large number."""

    W, H, R = 176, 84, 12

    def __init__(self, parent, app, glyph, caption, width=None, height=None):
        self.W = width or StatCard.W
        self.H = height or StatCard.H
        super().__init__(parent, width=self.W, height=self.H,
                         highlightthickness=0, bd=0, takefocus=0)
        self.app, self.glyph, self.tone = app, glyph, "normal"
        self._imgs = {}
        # caption sits a third of the way down, the number below it
        y_cap, y_val = round(self.H * 0.33), round(self.H * 0.69)
        self._bg = self.create_image(0, 0, anchor="nw")
        self._ic = self.create_image(18, y_cap, anchor="w")
        self._cap = self.create_text(42, y_cap, anchor="w", text=caption)
        self._val = self.create_text(18, y_val, anchor="w", text="0")

    def set(self, value, tone="normal"):
        self.tone = tone
        self.itemconfigure(self._val, text=f"{value:,}".replace(",", " "))
        self.restyle()

    def restyle(self):
        pal, art = self.app.pal, self.app.art
        self.configure(bg=pal["bg"])

        if self.tone == "danger":
            fill, line, accent = pal["dup_soft"], pal["dup_bg"], pal["dup_fg"]
        elif self.tone == "good":
            fill, line, accent = pal["new_soft"], pal["new_bg"], pal["new_fg"]
        else:
            fill, line, accent = pal["surface"], pal["border"], pal["accent"]

        self._imgs["bg"] = art.round_rect(self.W, self.H, self.R, fill,
                                          outline=line, width=1)
        if self._imgs["bg"]:
            self.itemconfigure(self._bg, image=self._imgs["bg"])
        self._imgs["ic"] = art.icon(self.glyph, 15, accent)
        if self._imgs["ic"]:
            self.itemconfigure(self._ic, image=self._imgs["ic"])

        self.itemconfigure(self._cap, fill=pal["muted"], font=(self.app.f_ui, 9))
        self.itemconfigure(self._val, fill=accent if self.tone != "normal"
                           else pal["text"], font=(self.app.f_disp, 22, "bold"))


class SearchBox(tk.Canvas):
    """A rounded search box with an icon inside it."""

    H, R = 38, 8

    def __init__(self, parent, app, on_change, width=300, height=None):
        self.H = height or SearchBox.H
        super().__init__(parent, width=width, height=self.H,
                         highlightthickness=0, bd=0, takefocus=0)
        self.app, self.W = app, width
        self._imgs = {}
        self._ph_on = True
        self.PLACEHOLDER = t("search_ph")

        self._bg = self.create_image(0, 0, anchor="nw")
        self._ic = self.create_image(14, self.H // 2, anchor="w")

        self.entry = tk.Entry(self, relief="flat", bd=0, highlightthickness=0)
        self.create_window(38, self.H // 2, anchor="w", window=self.entry,
                           width=width - 52, height=self.H - 12)
        self.entry.insert(0, self.PLACEHOLDER)
        self.entry.bind("<FocusIn>", self._focus_in)
        self.entry.bind("<FocusOut>", self._focus_out)
        self.entry.bind("<KeyRelease>", on_change)
        self.entry.bind("<Escape>", lambda e: (self.entry.delete(0, "end"),
                                               on_change(e)))

    def query(self):
        return "" if self._ph_on else self.entry.get().strip().lower()

    def _focus_in(self, _=None):
        if self._ph_on:
            self._ph_on = False
            self.entry.delete(0, "end")
            self.entry.configure(fg=self.app.pal["text"])

    def _focus_out(self, _=None):
        if not self.entry.get():
            self._ph_on = True
            self.entry.insert(0, self.PLACEHOLDER)
            self.entry.configure(fg=self.app.pal["faint"])

    def restyle(self):
        pal, art = self.app.pal, self.app.art
        self.configure(bg=pal["bg"])
        self._imgs["bg"] = art.round_rect(self.W, self.H, self.R,
                                          pal["surface"], outline=pal["border"])
        if self._imgs["bg"]:
            self.itemconfigure(self._bg, image=self._imgs["bg"])
        self._imgs["ic"] = art.icon(IC_SEARCH, 15, pal["muted"])
        if self._imgs["ic"]:
            self.itemconfigure(self._ic, image=self._imgs["ic"])
        self.entry.configure(bg=pal["surface"], font=(self.app.f_ui, 10),
                             insertbackground=pal["text"],
                             fg=pal["faint"] if self._ph_on else pal["text"])


class Toast(tk.Toplevel):
    """A small self-dismissing message in the bottom-right corner."""

    KEY = "#ff00fe"                     # the colour made transparent

    def __init__(self, parent, art, pal, f_ui, glyph, text, ms=2800):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        try:
            self.attributes("-transparentcolor", self.KEY)
        except tk.TclError:
            pass
        self.configure(bg=self.KEY)

        font = tkfont.Font(family=f_ui, size=10)
        w = min(460, max(240, font.measure(text) + 76))
        h = 54

        cv = tk.Canvas(self, width=w, height=h, bg=self.KEY,
                       highlightthickness=0, bd=0)
        cv.pack()
        self._bg = art.round_rect_flat(w, h, 10, pal["surface"],
                                       pal["border"], self.KEY)
        if self._bg:
            cv.create_image(0, 0, anchor="nw", image=self._bg)
        self._ic = art.icon(glyph, 18, pal["accent"])
        if self._ic:
            cv.create_image(20, h // 2, anchor="w", image=self._ic)
        cv.create_text(48, h // 2, anchor="w", text=text, fill=pal["text"],
                       font=(f_ui, 10))
        cv.bind("<Button-1>", lambda e: self._close())

        try:
            x = parent.winfo_rootx() + parent.winfo_width() - w - 30
            y = parent.winfo_rooty() + parent.winfo_height() - h - 62
            self.geometry(f"{w}x{h}+{x}+{y}")
        except tk.TclError:
            pass
        self.after(ms, self._close)

    def _close(self):
        try:
            self.destroy()
        except tk.TclError:
            pass


class Dialog(tk.Toplevel):
    """A modal dialog for questions and messages."""

    def __init__(self, parent, pal, art, f_ui, f_disp, title, glyph,
                 glyph_color, heading, body, buttons, entry=None,
                 values=None, check=None):
        super().__init__(parent)
        self.result = None
        self.checked = False
        self._entry_widget = None

        self.withdraw()
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.configure(bg=pal["bg"])

        wrap = tk.Frame(self, bg=pal["bg"], padx=26, pady=24)
        wrap.pack(fill="both", expand=True)

        head = tk.Frame(wrap, bg=pal["bg"])
        head.pack(fill="x")

        self._img = art.icon(glyph, 30, glyph_color)
        if self._img is not None:
            tk.Label(head, image=self._img, bg=pal["bg"]).pack(
                side="left", padx=(0, 15), anchor="n")

        tk.Label(head, text=heading, font=(f_disp, 14, "bold"),
                 bg=pal["bg"], fg=pal["text"], justify="left",
                 anchor="w", wraplength=390).pack(side="left", fill="x")

        if body:
            tk.Label(wrap, text=body, font=(f_ui, 10), bg=pal["bg"],
                     fg=pal["muted"], justify="left", anchor="w",
                     wraplength=450).pack(fill="x", pady=(14, 0))

        # entry: a string -> one field;  [(label, value), ...] -> several
        self.vars = []
        self._multi = entry is not None and not isinstance(entry, str)
        if entry is not None:
            fields = [(None, entry)] if isinstance(entry, str) else list(entry)
            for i, (label, value) in enumerate(fields):
                if label:
                    tk.Label(wrap, text=label, font=(f_ui, 9), bg=pal["bg"],
                             fg=pal["muted"], anchor="w").pack(
                        fill="x", pady=(12 if i else 18, 3))
                var = tk.StringVar(value=value)
                if i == 0 and values:      # offer previously used values
                    e = ttk.Combobox(wrap, textvariable=var, font=(f_ui, 12),
                                     values=list(values))
                else:
                    e = ttk.Entry(wrap, textvariable=var, font=(f_ui, 12))
                e.pack(fill="x", pady=(0 if label else 18, 0), ipady=5)
                e.bind("<Return>", lambda _e: self._pick(buttons[0][1]))
                self.vars.append(var)
                if i == 0:
                    e.focus_set()
                    if hasattr(e, "select_range"):
                        e.select_range(0, "end")
                    self._entry_widget = e
        self.var = self.vars[0] if self.vars else None

        # check: (label, initial state) -> one checkbox underneath
        self._check_var = None
        if check is not None:
            self._check_var = tk.BooleanVar(value=bool(check[1]))
            ttk.Checkbutton(wrap, text=check[0], variable=self._check_var
                            ).pack(anchor="w", pady=(16, 0))

        row = tk.Frame(wrap, bg=pal["bg"])
        row.pack(fill="x", pady=(24, 0))
        for text, value, accent in reversed(buttons):
            ttk.Button(row, text=text,
                       style="Accent.TButton" if accent else "TButton",
                       command=lambda v=value: self._pick(v)
                       ).pack(side="right", padx=(8, 0), ipadx=10, ipady=2)

        self.bind("<Escape>", lambda _e: self._pick(None))
        self.protocol("WM_DELETE_WINDOW", lambda: self._pick(None))

        self.update_idletasks()
        self._center(parent)
        style_titlebar(self, pal, pal is PALETTES["dark"])
        self.deiconify()
        if self._entry_widget is None:
            self.focus_set()
        self.grab_set()
        self.wait_window(self)

    def _center(self, parent):
        w, h = self.winfo_width(), self.winfo_height()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
        except tk.TclError:
            return
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 3}")

    def _pick(self, value):
        if self._check_var is not None:
            self.checked = bool(self._check_var.get())
        if value is not None and self._entry_widget is not None:
            if self._multi:
                self.result = (value, [v.get().strip() for v in self.vars])
            else:
                self.result = (value, self.var.get().strip())
        else:
            self.result = value
        self.destroy()


class SchemaDialog(tk.Toplevel):
    """
    Where the user writes their own columns: one per line, with the
    duplicate key and the extra columns chosen underneath.
    """

    def __init__(self, app, boshlangich=None):
        super().__init__(app.root)
        pal, f_ui, f_disp = app.pal, app.f_ui, app.f_disp
        self.app = app
        self.result = None

        self.withdraw()
        self.title(t("sch_title"))
        self.resizable(False, False)
        self.transient(app.root)
        self.configure(bg=pal["bg"])

        wrap = tk.Frame(self, bg=pal["bg"], padx=26, pady=24)
        wrap.pack(fill="both", expand=True)

        head = tk.Frame(wrap, bg=pal["bg"])
        head.pack(fill="x")
        self._img = app.art.icon(IC_TOOLS, 30, pal["accent"])
        if self._img is not None:
            tk.Label(head, image=self._img, bg=pal["bg"]).pack(
                side="left", padx=(0, 15), anchor="n")
        tk.Label(head, text=t("sch_head"), font=(f_disp, 14, "bold"),
                 bg=pal["bg"], fg=pal["text"], anchor="w").pack(side="left")

        tk.Label(wrap, text=t("sch_body"), font=(f_ui, 10), bg=pal["bg"],
                 fg=pal["muted"], justify="left", anchor="w",
                 wraplength=430).pack(fill="x", pady=(14, 0))

        self.text = tk.Text(wrap, width=44, height=8, font=(f_ui, 11),
                            bg=pal["surface"], fg=pal["text"], relief="flat",
                            insertbackground=pal["text"], padx=10, pady=8,
                            highlightthickness=1, highlightbackground=pal["border"],
                            highlightcolor=pal["accent"])
        self.text.pack(fill="x", pady=(16, 0))
        self.text.insert("1.0", "\n".join(boshlangich or []))
        self.text.bind("<KeyRelease>", lambda _e: self._refresh_keys())

        self.no_var = tk.BooleanVar(value=True)
        self.src_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(wrap, text=t("sch_no"), variable=self.no_var
                        ).pack(anchor="w", pady=(16, 0))
        ttk.Checkbutton(wrap, text=t("sch_src"), variable=self.src_var
                        ).pack(anchor="w", pady=(4, 0))

        tk.Label(wrap, text=t("sch_key"), font=(f_ui, 9), bg=pal["bg"],
                 fg=pal["muted"], anchor="w").pack(fill="x", pady=(18, 3))
        self.key_var = tk.StringVar(value=t("sch_none"))
        self.key_box = ttk.Combobox(wrap, textvariable=self.key_var,
                                    font=(f_ui, 11), state="readonly")
        self.key_box.pack(fill="x", ipady=3)
        self.key_box.bind("<<ComboboxSelected>>", lambda _e: None)

        tk.Label(wrap, text=t("sch_mode"), font=(f_ui, 9), bg=pal["bg"],
                 fg=pal["muted"], anchor="w").pack(fill="x", pady=(12, 3))
        self.mode_var = tk.StringVar(value=t("km_text"))
        ttk.Combobox(wrap, textvariable=self.mode_var, font=(f_ui, 11),
                     state="readonly",
                     values=[t("km_text"), t("km_phone"), t("km_number")]
                     ).pack(fill="x", ipady=3)

        row = tk.Frame(wrap, bg=pal["bg"])
        row.pack(fill="x", pady=(24, 0))
        ttk.Button(row, text=t("btn_cancel"), command=self._cancel
                   ).pack(side="right", padx=(8, 0), ipadx=10, ipady=2)
        ttk.Button(row, text=t("btn_save"), style="Accent.TButton",
                   command=self._ok).pack(side="right", ipadx=10, ipady=2)

        self._refresh_keys()
        self.bind("<Escape>", lambda _e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self.update_idletasks()
        try:
            px, py = app.root.winfo_rootx(), app.root.winfo_rooty()
            pw, ph = app.root.winfo_width(), app.root.winfo_height()
            self.geometry(f"+{px + (pw - self.winfo_width()) // 2}"
                          f"+{max(0, py + (ph - self.winfo_height()) // 4)}")
        except tk.TclError:
            pass
        style_titlebar(self, pal, pal is PALETTES["dark"])
        self.deiconify()
        self.text.focus_set()
        self.grab_set()
        self.wait_window(self)

    def _column_names(self):
        return [s.strip() for s in self.text.get("1.0", "end").splitlines()
                if s.strip()]

    def _refresh_keys(self):
        """Keep the duplicate-key dropdown in step with the typed column names."""
        names = self._column_names()
        self.key_box.configure(values=[t("sch_none")] + names)
        if self.key_var.get() not in names:
            self.key_var.set(t("sch_none"))

    def _cancel(self):
        self.result = None
        self.destroy()

    def _ok(self):
        names = self._column_names()
        if not names:
            self.app.error(t("sch_err"), t("sch_err_b"), IC_INFO)
            return

        cols, used = [], set()
        if self.no_var.get():
            cols.append(col("no", 55, "no", "h_no"))
            used.add("no")

        key_name = None
        for i, name in enumerate(names):
            key = re.sub(r"\W+", "_", name.lower()).strip("_") or f"c{i}"
            while key in used:
                key += "_"
            used.add(key)
            cols.append(col(key, 200, "", title=name))
            if name == self.key_var.get():
                key_name = key

        if self.src_var.get():
            cols.append(col("source", 200, "src", "h_source"))

        mode = {t("km_phone"): "phone", t("km_number"): "number"}.get(
            self.mode_var.get(), "text")

        self.result = {"name": "custom", "cols": cols,
                       "key_col": key_name,
                       "key_mode": mode if key_name else "text"}
        self.destroy()


class Busy(tk.Toplevel):
    """
    Shown during a long operation such as reading a large file, so the
    application does not look frozen.
    """

    def __init__(self, app, text):
        super().__init__(app.root)
        pal = app.pal
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=pal["border"])

        inner = tk.Frame(self, bg=pal["surface"], padx=30, pady=22)
        inner.pack(padx=1, pady=1)
        self._lbl = tk.Label(inner, text=text, bg=pal["surface"], fg=pal["text"],
                             font=(app.f_ui, 10), anchor="w", width=34)
        self._lbl.pack(fill="x")
        self._bar = ttk.Progressbar(inner, mode="indeterminate", length=280)
        self._bar.pack(pady=(14, 0))
        self._bar.start(12)

        self.update_idletasks()
        try:
            p = app.root
            self.geometry(
                f"+{p.winfo_rootx() + (p.winfo_width() - self.winfo_width()) // 2}"
                f"+{p.winfo_rooty() + (p.winfo_height() - self.winfo_height()) // 2}")
        except tk.TclError:
            pass
        self.update()

    def step(self, text):
        try:
            self._lbl.configure(text=text)
            self.update()
        except tk.TclError:
            pass

    def close(self):
        try:
            self._bar.stop()
            self.destroy()
        except tk.TclError:
            pass


# ================================================================= application

# The sidebar width is not a fixed constant: DatabaseApp._ui_metrics()
# measures the text, so display scaling and longer languages both fit.


def pretty_path(path, limit=82):
    """Shorten a long path so it fits in the window."""
    if not path:
        return t("not_saved")
    if len(path) <= limit:
        return path
    parts = path.replace("/", "\\").split("\\")
    short = f"{parts[0]}\\…\\" + "\\".join(parts[-2:])
    return short if len(short) <= limit else "…\\" + parts[-1]


class DatabaseApp:

    def __init__(self, root):
        self.root = root
        self.rows = []          # asosiy baza
        self.path = None        # current file path
        self.dirty = False      # are there unsaved changes
        self.dup_keys = set()   # normalised keys that occur more than once
        self.new_from = None    # index of the first row added last time
        self.view_index = []
        self.schema = default_schema()   # the shape of the database

        self.cfg = load_config()
        self.recent = [p for p in self.cfg.get("recent", [])
                       if isinstance(p, str)][:RECENT_LIMIT]
        self.last_dir = self.cfg.get("last_dir") or ""
        self.sources = [s for s in self.cfg.get("sources", [])
                        if isinstance(s, str)][:SOURCE_LIMIT]
        self._undo, self._redo = [], []          # ortga/qaytadan qaytarish steki
        self.sort_col, self.sort_desc = None, False
        self._sort_numeric = False       # saralanayotgan column sonlimi
        self.hint = ""          # pastki maslahat matni (til almashsa saqlanadi)
        self.last_check = False                  # last dialogdagi tugmacha

        global LANG                              # tilni sozlamadan tiklaymiz
        LANG = self.cfg.get("lang") if self.cfg.get("lang") in LANG_ORDER \
            else system_lang()

        self.art = Art()
        self.f_ui = self._pick_font("Segoe UI Variable Text", "Segoe UI")
        self.f_disp = self._pick_font("Segoe UI Variable Display", "Segoe UI")

        start_theme = self.cfg.get("theme")
        if start_theme not in PALETTES:
            start_theme = "light"

        self.theme = tk.StringVar(value=start_theme)
        self.last9 = tk.BooleanVar(value=bool(self.cfg.get("last9", True)))
        self.only_dups = tk.BooleanVar(value=False)
        self.only_bad = tk.BooleanVar(value=False)   # faqat shubhali rows

        self.pal = PALETTES["light"]
        self._navs = []         # yon panel tugmalari
        self._cards = []        # statistic cards

        root.report_callback_exception = self._on_tk_error

        self._build_ui()
        self._setup_window()
        self._apply_theme(start_theme)
        self.refresh()

    # ------------------------------------------------------------------ language

    def set_hint(self, text):
        """Set the hint line at the bottom and remember it."""
        self.hint = text
        self.hint_var.set(text)

    def show_langs(self):
        """The language menu."""
        menu = self._menu()
        for code in LANG_ORDER:
            mark = "  ● " if code == LANG else "     "
            menu.add_command(label=mark + LANG_NAMES[code],
                             command=lambda c=code: self.set_lang(c))
        self._popup(menu, self.btn_lang)

    # --------------------------------------------------------------------- tools

    def _menu(self):
        """An empty menu coloured to match the theme."""
        pal = self.pal
        return tk.Menu(self.root, tearoff=0, bd=0, activeborderwidth=0,
                       bg=pal["surface"], fg=pal["text"],
                       activebackground=pal["accent"],
                       activeforeground=pal["on_accent"],
                       font=(self.f_ui, 10))

    def _popup(self, menu, button):
        try:
            menu.tk_popup(button.winfo_rootx() + button.winfo_width() + 6,
                          button.winfo_rooty())
        finally:
            menu.grab_release()

    def show_tools(self):
        """The tools menu; only entries that apply to this schema are shown."""
        has_phone = role_col(self.schema, "phone") is not None
        menu = self._menu()
        menu.add_command(label="  " + t("tool_paste"), command=self.paste_rows)
        menu.add_separator()
        if has_phone:
            menu.add_command(label="  " + t("tool_normalize"),
                             command=self.normalize_phones)
        menu.add_checkbutton(label="  " + t("tool_problems"),
                             variable=self.only_bad, command=self.toggle_bad)
        menu.add_separator()
        menu.add_command(label="  " + t("tool_csv"), command=self.export_csv)
        if has_phone or role_col(self.schema, "name"):
            menu.add_command(label="  " + t("tool_vcard"),
                             command=self.export_vcard)
        self._popup(menu, self.btn_tools)

    def set_lang(self, code):
        """Switch the language and rebuild the interface."""
        global LANG
        if code == LANG or code not in LANG_ORDER:
            return
        LANG = code
        self._save_config()
        self._rebuild_ui()
        self.toast(LANG_NAMES[code], IC_GLOBE)

    def _rebuild_ui(self):
        """
        Rebuild the whole interface. Needed when the language or the
        columns change, because both live inside the widgets.
        """
        for w in self.root.winfo_children():
            if isinstance(w, tk.Toplevel):
                continue                      # ochiq dialoglarga tegmaymiz
            w.destroy()
        self._navs, self._cards = [], []

        self._build_ui()
        self._apply_theme(self.theme.get())
        self.refresh()

    # ------------------------------------------------------------ error handling

    def _on_tk_error(self, exc_type, exc, tb):
        """An unexpected error inside Tkinter: log it and show it."""
        log_error(exc_type, exc, tb)
        try:
            self.error(t("err_unexpected"),
                       t("err_unexpected_b", name=exc_type.__name__,
                         msg=exc, log=LOG_PATH))
        except Exception:
            message_box(f"{exc_type.__name__}: {exc}\n\nLog: {LOG_PATH}")

    @staticmethod
    def _pick_font(*names):
        fams = set(tkfont.families())
        for n in names:
            if n in fams:
                return n
        return "Segoe UI"

    # ------------------------------------------------------------------ interface

    def _setup_window(self):
        """One-time window setup; not repeated when the language changes."""
        r = self.root
        # The minimum size follows the text size: at 150% scaling both the
        # sidebar and the table need more room.
        # It must not end up larger than the screen either.
        min_w = min(self.m["sidebar_w"] + int(830 * self.m["scale"]),
                    r.winfo_screenwidth() - 40)
        min_h = min(max(690, self.m["min_h"]), r.winfo_screenheight() - 80)
        r.minsize(min_w, min_h)
        self._restore_window(int(1300 * self.m["scale"]),
                             int(820 * self.m["scale"]))
        r.protocol("WM_DELETE_WINDOW", self.on_close)

        ico = find_icon_file()
        if ico:
            try:
                r.iconbitmap(ico)
            except tk.TclError:
                pass

        if HAS_DND:                       # fayllarni oynaga sudrab tashlash
            try:
                r.drop_target_register(DND_FILES)
                r.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        for seq, fn in (("<Control-s>", lambda e: self.save()),
                        ("<Control-o>", lambda e: self.open_base()),
                        ("<Control-n>", lambda e: self.new_base()),
                        ("<Control-f>",
                         lambda e: self.searchbox.entry.focus_set()),
                        ("<Control-z>", self.undo),
                        ("<Control-y>", self.redo),
                        ("<Control-Shift-Z>", self.redo),
                        ("<Control-c>", self.copy_selected),
                        ("<Control-v>", self.paste_rows),
                        ("<Control-a>", self.select_all)):
            r.bind(seq, fn)

    def _ui_metrics(self):
        """
        Work out the sidebar and card sizes FROM THE TEXT WIDTH.

        Fixed pixel values cannot be trusted: at 125% or 150% display
        scaling (ordinary on laptops) the font grows, and on top of that
        Russian labels are longer than English ones. Together they push the
        text out of the button, so we measure instead of guessing.
        """
        f = tkfont.Font(family=self.f_ui, size=10)
        fb = tkfont.Font(family=self.f_ui, size=10, weight="bold")
        fc = tkfont.Font(family=self.f_ui, size=9)
        fd = tkfont.Font(family=self.f_disp, size=22, weight="bold")

        navs = [t(k) for k in ("nav_new", "nav_open", "nav_recent", "nav_add",
                                 "nav_manual", "nav_save", "nav_saveas",
                                 "nav_excel", "nav_tools", "nav_dark",
                                 "nav_light")]
        navs.append(f"{t('nav_lang')}  ·  {LANG_SHORT.get(LANG, 'UZ')}")
        wide = max(max(f.measure(s), fb.measure(s)) for s in navs)

        line_h = f.metrics("linespace")
        m = {
            "nav_w": min(460, max(244, wide + 62)),   # icon + text + margin
            "nav_h": max(38, line_h + 16),
            "sec_h": max(19, fc.metrics("linespace") + 4),
            "brand_h": max(46, line_h * 2 + 10),
            "search_h": max(38, line_h + 16),
            "card_h": max(84, line_h + fd.metrics("linespace") + 26),
        }
        m["sidebar_w"] = m["nav_w"] + 28
        captions = [t(k) for k in ("card_total", "card_dup", "card_src",
                                   "card_view")]
        m["card_w"] = max(176, max(fc.measure(s) for s in captions) + 60)
        m["scale"] = max(1.0, line_h / 20)          # used for column widths

        # height needed for every sidebar button to fit
        m["min_h"] = (34 + m["brand_h"] + 3 * (m["sec_h"] + 14)
                      + 9 * (m["nav_h"] + 2)          # the buttons above
                      + 2 * (m["nav_h"] + 8) + 40)    # language + theme + slack
        return m

    def _build_ui(self):
        r = self.root
        r.title(t("app_name"))
        self.m = self._ui_metrics()

        # ================================================================== sidebar
        self.sidebar = tk.Frame(r, width=self.m["sidebar_w"])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.sb_divider = tk.Frame(r, width=1)
        self.sb_divider.pack(side="left", fill="y")

        # --- logo
        brand = tk.Frame(self.sidebar)
        brand.pack(fill="x", padx=14, pady=(18, 14))
        self.brand_frame = brand
        self.logo = tk.Label(brand)
        self.logo.pack(side="left", padx=(2, 12))
        btexts = tk.Frame(brand)
        btexts.pack(side="left")
        self.brand_texts = btexts
        self.lbl_brand = tk.Label(btexts, text=t("brand1"), anchor="w")
        self.lbl_brand.pack(anchor="w")
        # second line: the current database type (Contacts / Products / ...)
        self.lbl_brand2 = tk.Label(btexts, text=schema_title(self.schema),
                                   anchor="w")
        self.lbl_brand2.pack(anchor="w")

        # --- sections
        self.sec_labels = []

        def section(title):
            lb = tk.Label(self.sidebar, text=title, anchor="w")
            lb.pack(fill="x", padx=22, pady=(10, 4))
            self.sec_labels.append(lb)

        def nav(glyph, text, cmd, accent=False):
            b = NavButton(self.sidebar, self, glyph, text, cmd,
                          width=self.m["nav_w"], accent=accent,
                          height=self.m["nav_h"])
            b.pack(padx=14, pady=1)
            self._navs.append(b)
            return b

        section(t("sec_base"))
        self.btn_new = nav(IC_NEW, t("nav_new"), self.new_base)
        nav(IC_OPEN, t("nav_open"), self.open_base)
        self.btn_recent = nav(IC_HISTORY, t("nav_recent"), self.show_recent)

        section(t("sec_contacts"))
        nav(IC_ADD, t("nav_add"), self.add_contacts, accent=True)
        nav(IC_PERSONADD, t("nav_manual"), self.add_manual)

        section(t("sec_output"))
        nav(IC_SAVE, t("nav_save"), self.save)
        nav(IC_SAVEAS, t("nav_saveas"), self.save_as)
        nav(IC_EXTERN, t("nav_excel"), self.open_in_excel)
        self.btn_tools = nav(IC_TOOLS, t("nav_tools"), self.show_tools)

        # --- at the bottom: theme and language (packed bottom-up)
        self.btn_theme = NavButton(self.sidebar, self, IC_MOON, t("nav_dark"),
                                   self.toggle_theme, width=self.m["nav_w"],
                                   height=self.m["nav_h"])
        self.btn_theme.pack(side="bottom", padx=14, pady=(4, 16))
        self._navs.append(self.btn_theme)

        self.btn_lang = NavButton(
            self.sidebar, self, IC_GLOBE,
            f"{t('nav_lang')}  ·  {LANG_SHORT.get(LANG, 'UZ')}",
            self.show_langs, width=self.m["nav_w"], height=self.m["nav_h"])
        self.btn_lang.pack(side="bottom", padx=14, pady=(4, 0))
        self._navs.append(self.btn_lang)

        # ============================================================== main area
        self.content = tk.Frame(r)
        self.content.pack(side="left", fill="both", expand=True)

        # --- title area
        top = tk.Frame(self.content)
        top.pack(fill="x", padx=28, pady=(24, 4))
        self.top_frame = top
        self.lbl_file = tk.Label(top, text=t("new_base_label"), anchor="w")
        self.lbl_file.pack(anchor="w")
        self.lbl_path = tk.Label(top, text=t("not_saved"), anchor="w")
        self.lbl_path.pack(anchor="w", pady=(2, 0))

        # --- statistic cards
        cards = tk.Frame(self.content)
        cards.pack(fill="x", padx=28, pady=(18, 4))
        self.cards_frame = cards

        def card(glyph, caption):
            c = StatCard(cards, self, glyph, caption,
                         width=self.m["card_w"], height=self.m["card_h"])
            c.pack(side="left", padx=(0, 12))
            self._cards.append(c)
            return c

        self.card_total = card(IC_ROWS, t("card_total"))
        self.card_dup = card(IC_WARN, t("card_dup"))
        self.card_src = card(IC_PEOPLE, t("card_src"))
        self.card_view = card(IC_FILTER, t("card_view"))

        # --- search and filters
        bar = tk.Frame(self.content)
        bar.pack(fill="x", padx=28, pady=(18, 10))
        self.bar_frame = bar

        self.searchbox = SearchBox(bar, self, lambda e: self.refresh_table(),
                                   height=self.m["search_h"])
        self.searchbox.pack(side="left")

        sw = "Switch.TCheckbutton" if HAS_SV else "TCheckbutton"
        ttk.Checkbutton(bar, text=t("only_dups"), style=sw,
                        variable=self.only_dups,
                        command=self.refresh_table).pack(side="left", padx=(22, 0))
        # "last 9 digits" only means anything when matching on a phone number
        if self.schema.get("key_mode") == "phone":
            ttk.Checkbutton(bar, text=t("last9"), style=sw,
                            variable=self.last9,
                            command=self._toggle_last9
                            ).pack(side="left", padx=(22, 0))

        ttk.Button(bar, text=t("btn_del_bar"),
                   command=self.delete_selected).pack(side="right", ipady=2)
        ttk.Button(bar, text=t("btn_src_bar"),
                   command=self.edit_source).pack(side="right", padx=(0, 8),
                                                  ipady=2)

        # --- table
        self.card = ttk.Frame(self.content,
                              style="Card.TFrame" if HAS_SV else "TFrame",
                              padding=6)
        self.card.pack(fill="both", expand=True, padx=28)

        flds = schema_fields(self.schema)
        self.tree = ttk.Treeview(self.card, columns=flds, show="headings",
                                 selectmode="extended")
        # widths come from the schema (scaled for the display), then any
        # width the user set last time is applied on top
        k = self.m["scale"]
        widths = {c["key"]: int(c["width"] * k) for c in self.schema["cols"]}
        saved = (self.cfg.get("colw") or {}).get(self._schema_id(), {})
        for f, w in saved.items():
            if f in widths and isinstance(w, int) and 30 <= w <= 1400:
                widths[f] = w

        for c, h in zip(self.schema["cols"], headers(self.schema)):
            f, role = c["key"], c.get("role", "")
            anchor = "center" if role == "no" else "w"
            self.tree.heading(f, text=h, anchor=anchor,
                              command=lambda ff=f: self.sort_by(ff))
            self.tree.column(f, width=widths[f], anchor=anchor,
                             stretch=(role in ("name", "src")))

        vsb = ttk.Scrollbar(self.card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # --- empty-state message
        self.empty = tk.Frame(self.card)
        self.empty_icon = tk.Label(self.empty)
        self.empty_icon.pack()
        self.empty_title = tk.Label(self.empty, text=t("empty_title"))
        self.empty_title.pack(pady=(14, 4))
        self.empty_text = tk.Label(self.empty, justify="center",
                                   text=t("empty_text"))
        self.empty_text.pack()
        self.empty_recent = tk.Frame(self.empty)   # recent files go in here
        self.empty_recent.pack(pady=(18, 0))

        # --- bottom row
        foot = tk.Frame(self.content)
        foot.pack(fill="x", padx=28, pady=(10, 16))
        self.foot_frame = foot
        self.hint_var = tk.StringVar(value=self.hint or t("hint_start"))
        self.lbl_hint = tk.Label(foot, textvariable=self.hint_var, anchor="w")
        self.lbl_hint.pack(side="left")
        self.lbl_keys = tk.Label(foot, anchor="e", text=t("keys_hint"))
        self.lbl_keys.pack(side="right")

        # --- table klaviaturasi
        self.tree.bind("<Double-1>", lambda e: self.edit_row())
        self.tree.bind("<Return>", lambda e: self.edit_row())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())

        self._update_headings()      # keep the sort marker across a language switch

    def _restore_window(self, w, h):
        """Restore the previous window size, or centre a default one."""
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()

        placed = False
        geo = self.cfg.get("geometry")
        m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$", geo or "")
        if m:
            gw, gh, gx, gy = (int(v) for v in m.groups())
            # a saved size is useless if the screen or the scaling changed
            min_w = self.m["sidebar_w"] + int(830 * self.m["scale"])
            if (gw >= min_w and gh >= max(690, self.m["min_h"])
                    and gw <= sw and gh <= sh
                    and -20 <= gx <= sw - 200 and -20 <= gy <= sh - 200):
                self.root.geometry(geo)
                placed = True

        if not placed:
            w, h = min(w, sw - 80), min(h, sh - 120)
            self.root.geometry(
                f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 2 - 30)}")

        if self.cfg.get("zoomed"):     # last safar yoyilgan holda yopilgan
            try:
                self.root.state("zoomed")
            except tk.TclError:
                pass

    # ------------------------------------------------------------------ settings

    def _save_config(self, geometry=False):
        if geometry:
            try:
                zoomed = (self.root.state() == "zoomed")
                self.cfg["zoomed"] = zoomed
                if not zoomed:
                    self.cfg["geometry"] = self.root.geometry()
                # column widths are stored per database type
                barcha = self.cfg.get("colw")
                if not isinstance(barcha, dict):
                    barcha = {}
                barcha[self._schema_id()] = {
                    f: int(self.tree.column(f, "width"))
                    for f in schema_fields(self.schema)}
                self.cfg["colw"] = barcha
            except (tk.TclError, ValueError):
                pass
        self.cfg.update({"theme": self.theme.get(),
                         "lang": LANG,
                         "last9": bool(self.last9.get()),
                         "last_dir": self.last_dir,
                         "recent": self.recent,
                         "sources": self.sources})
        save_config(self.cfg)

    def _toggle_last9(self):
        self.refresh()
        self._save_config()

    def _dialog_dir(self):
        return self.last_dir if os.path.isdir(self.last_dir or "") else None

    # --------------------------------------------------------------------- theme

    def toggle_theme(self):
        self._apply_theme("dark" if self.theme.get() == "light" else "light")
        self.refresh_table()
        self._save_config()

    def _apply_theme(self, name):
        self.theme.set(name)
        pal = self.pal = PALETTES[name]
        dark = (name == "dark")

        if HAS_SV:
            sv_ttk.set_theme(name)

        style = ttk.Style()
        style.configure("Treeview", rowheight=34, font=(self.f_ui, 10),
                        background=pal["surface"], fieldbackground=pal["surface"],
                        foreground=pal["text"], borderwidth=0)
        style.configure("Treeview.Heading", font=(self.f_ui, 10, "bold"),
                        padding=(12, 10))
        style.map("Treeview", background=[("selected", pal["accent"])],
                  foreground=[("selected", pal["on_accent"])])

        self.root.configure(bg=pal["bg"])
        for w in (self.sidebar, self.brand_frame, self.brand_texts, self.logo,
                  self.lbl_brand, self.lbl_brand2):
            w.configure(bg=pal["sidebar"])
        for w in (self.content, self.top_frame, self.cards_frame,
                  self.bar_frame, self.foot_frame, self.lbl_file,
                  self.lbl_path, self.lbl_hint, self.lbl_keys):
            w.configure(bg=pal["bg"])
        for w in (self.empty, self.empty_icon, self.empty_title,
                  self.empty_text, self.empty_recent):
            w.configure(bg=pal["surface"])

        self.sb_divider.configure(bg=pal["border"])

        self.lbl_brand.configure(fg=pal["text"], font=(self.f_disp, 13, "bold"))
        self.lbl_brand2.configure(fg=pal["muted"], font=(self.f_ui, 9))
        self.lbl_file.configure(fg=pal["text"], font=(self.f_disp, 20, "bold"))
        self.lbl_path.configure(fg=pal["muted"], font=(self.f_ui, 9))
        self.lbl_hint.configure(fg=pal["muted"], font=(self.f_ui, 9))
        self.lbl_keys.configure(fg=pal["faint"], font=(self.f_ui, 9))
        self.empty_title.configure(fg=pal["text"], font=(self.f_disp, 14, "bold"))
        self.empty_text.configure(fg=pal["muted"], font=(self.f_ui, 10))

        for lb in self.sec_labels:
            lb.configure(bg=pal["sidebar"], fg=pal["faint"],
                         font=(self.f_ui, 8, "bold"))

        self._logo_img = self.art.icon(IC_LOGO, 28, pal["accent"])
        if self._logo_img:
            self.logo.configure(image=self._logo_img)
        self._empty_img = self.art.icon(IC_EMPTY, 44, pal["faint"])
        if self._empty_img:
            self.empty_icon.configure(image=self._empty_img)

        self.btn_theme.set_glyph(IC_SUN if dark else IC_MOON,
                                 t("nav_light") if dark else t("nav_dark"))
        for b in self._navs:
            b.restyle()
        for c in self._cards:
            c.restyle()
        self.searchbox.restyle()

        self.tree.tag_configure("dup", background=pal["dup_bg"],
                                foreground=pal["dup_fg"])
        self.tree.tag_configure("new", background=pal["new_bg"],
                                foreground=pal["new_fg"])
        self.tree.tag_configure("bad", background=pal["bad_bg"],
                                foreground=pal["bad_fg"])
        self.tree.tag_configure("odd", background=pal["stripe"],
                                foreground=pal["text"])
        self.tree.tag_configure("even", background=pal["surface"],
                                foreground=pal["text"])

        style_titlebar(self.root, pal, dark)

    # --------------------------------------------------------------- computation

    def key_of(self, row):
        """Duplicate key of a row, taken from the column the schema names."""
        return row_key(row, self.schema, self.last9.get())

    def compute_dups(self):
        counts = {}
        for row in self.rows:
            k = self.key_of(row)
            if k:
                counts[k] = counts.get(k, 0) + 1
        self.dup_keys = {k for k, c in counts.items() if c > 1}

    def renumber(self):
        nc = role_col(self.schema, "no")
        if not nc:                       # this schema has no row-number column
            return
        for i, row in enumerate(self.rows, start=1):
            row[nc] = i

    def refresh(self):
        self.compute_dups()
        self.refresh_table()

    @staticmethod
    def _as_number(v):
        try:
            return float(v.replace(" ", "").replace(",", "."))
        except ValueError:
            return None

    def _column_numeric(self, field):
        """
        Are all the values in this column numbers? Then price/quantity
        columns sort numerically instead of as text, so 100 does not come
        before 20.
        """
        seen = 0
        for r in self.rows:
            v = cell_text(r.get(field, ""))
            if not v:
                continue
            if self._as_number(v) is None:
                return False
            seen += 1
            if seen >= 200:          # this many is enough to decide
                break
        return seen > 0

    def _sort_key(self, field, row):
        """Sort key; empty values always sort last."""
        v = cell_text(row.get(field, ""))
        role = next((c.get("role", "") for c in self.schema["cols"]
                     if c["key"] == field), "")

        if role == "phone":
            d = re.sub(r"\D", "", v)      # by digits, which groups by dialling code
            return (0 if d else 1, 0.0, d)

        if self._sort_numeric:
            number = self._as_number(v)
            return (0 if v else 1, number if number is not None else 0.0, "")
        return (0 if v else 1, 0.0, v.casefold())

    def sort_by(self, field):
        """Clicking a header cycles: ascending -> descending -> original order."""
        if self.sort_col != field:
            self.sort_col, self.sort_desc = field, False
        elif not self.sort_desc:
            self.sort_desc = True
        else:
            self.sort_col, self.sort_desc = None, False
        self._update_headings()
        self.refresh_table()

    def _fill_empty_recent(self):
        """Draw the recent files as a clickable list on the empty screen."""
        for w in self.empty_recent.winfo_children():
            w.destroy()

        items = [p for p in self.recent if os.path.exists(p)][:5]
        pal = self.pal
        self.empty_recent.configure(bg=pal["surface"])
        if not items:
            return

        tk.Label(self.empty_recent, text=t("empty_recent"), bg=pal["surface"],
                 fg=pal["faint"], font=(self.f_ui, 9)).pack(pady=(0, 6))
        for p in items:
            lb = tk.Label(self.empty_recent, text=pretty_path(p, 56),
                          bg=pal["surface"], fg=pal["accent"],
                          font=(self.f_ui, 10), cursor="hand2")
            lb.pack(pady=1)
            lb.bind("<Button-1>", lambda _e, pp=p: self.open_recent(pp))

    def _schema_id(self):
        """A short identifier for the schema, used as a settings key."""
        return self.schema.get("name") or "-".join(schema_fields(self.schema))

    def _update_headings(self):
        for f, h in zip(schema_fields(self.schema), headers(self.schema)):
            mark = ""
            if self.sort_col == f:
                mark = "  ▼" if self.sort_desc else "  ▲"
            self.tree.heading(f, text=h + mark)

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        self.view_index = []

        q = self.searchbox.query()
        only_d = self.only_dups.get()
        only_b = self.only_bad.get()
        flds = schema_fields(self.schema)
        dup_rows = shown = 0

        order = list(range(len(self.rows)))
        if self.sort_col:
            self._sort_numeric = self._column_numeric(self.sort_col)
            order.sort(key=lambda i: self._sort_key(self.sort_col, self.rows[i]),
                       reverse=self.sort_desc)

        for i in order:
            row = self.rows[i]
            key = self.key_of(row)
            is_dup = bool(key) and key in self.dup_keys
            if is_dup:
                dup_rows += 1
            if only_d and not is_dup:
                continue

            is_bad = row_problem(row, self.schema) is not None
            if only_b and not is_bad:
                continue
            if q and q not in " ".join(str(row.get(f, "")).lower()
                                       for f in flds):
                continue

            if is_dup:
                tag = "dup"
            elif is_bad:
                tag = "bad"
            elif self.new_from is not None and i >= self.new_from:
                tag = "new"
            else:
                tag = "odd" if shown % 2 else "even"

            self.tree.insert("", "end", iid=str(i),
                             values=[row.get(f, "") for f in flds], tags=(tag,))
            self.view_index.append(i)
            shown += 1

        sc = role_col(self.schema, "src")
        sources = ({cell_text(r.get(sc)) for r in self.rows if cell_text(r.get(sc))}
                   if sc else set())
        self.card_total.set(len(self.rows))
        self.card_dup.set(dup_rows, "danger" if dup_rows else "normal")
        self.card_src.set(len(sources))
        self.card_view.set(shown)

        if self.rows:
            self.empty.place_forget()
        else:
            self._fill_empty_recent()
            self.empty.place(relx=0.5, rely=0.44, anchor="center")

        name = os.path.basename(self.path) if self.path else t("new_base_label")
        self.lbl_file.configure(text=name + (" ●" if self.dirty else ""))
        self.lbl_path.configure(text=pretty_path(self.path))
        self.root.title(f"{t('app_name')} — {name}"
                        f"{' •' if self.dirty else ''}")

    # -------------------------------------------------------------------- dialogs

    def ask(self, title, glyph, heading, body, buttons, entry=None, color=None,
            values=None, check=None):
        """
        Show the dialog. When `check` is given, the checkbox state is left
        in `self.last_check` so the shape of the result does not change.
        """
        dlg = Dialog(self.root, self.pal, self.art, self.f_ui, self.f_disp,
                     title, glyph, color or self.pal["accent"], heading, body,
                     buttons, entry, values, check)
        self.last_check = dlg.checked
        return dlg.result

    def info(self, heading, body="", glyph=IC_CHECK):
        self.ask(t("dlg_msg"), glyph, heading, body,
                 [(t("btn_ok"), True, True)])

    def error(self, heading, body="", glyph=IC_WARN):
        self.ask(t("dlg_err"), glyph, heading, body,
                 [(t("btn_close"), True, True)], color=self.pal["dup_fg"])

    def toast(self, text, glyph=IC_CHECK):
        """A short message that does not interrupt the work."""
        try:
            Toast(self.root, self.art, self.pal, self.f_ui, glyph, text)
        except Exception:
            pass

    # ---------------------------------------------------------------- undo / redo

    def _push_undo(self, label):
        """Call BEFORE a change: stores a copy of the database."""
        self._undo.append((label, [dict(r) for r in self.rows], self.new_from))

        # 30 copies of a very large database eats memory, so trim the history
        total = sum(len(step[1]) for step in self._undo)
        while len(self._undo) > 1 and (len(self._undo) > UNDO_LIMIT
                                       or total > UNDO_ROW_BUDGET):
            total -= len(self._undo[0][1])
            self._undo.pop(0)

        self._redo.clear()

    def _restore(self, stack, other, word, empty_msg):
        if not stack:
            self.toast(empty_msg, IC_INFO)
            return
        label, rows, new_from = stack.pop()
        other.append((label, [dict(r) for r in self.rows], self.new_from))
        self.rows, self.new_from = rows, new_from
        self.dirty = True
        self.refresh()
        self.toast(f"{word}: {label}", IC_UNDO)

    def undo(self, *_):
        self._restore(self._undo, self._redo, t("undo_word"), t("undo_empty"))

    def redo(self, *_):
        self._restore(self._redo, self._undo, t("redo_word"), t("redo_empty"))

    # ------------------------------------------------------- copying / selecting

    def _in_entry(self):
        """Is focus in a text field? Then Ctrl+C / Ctrl+A must not hit the table."""
        try:
            return isinstance(self.root.focus_get(), (tk.Entry, ttk.Entry))
        except (KeyError, tk.TclError):     # focus may be in another window
            return False

    def copy_selected(self, *_):
        if self._in_entry():
            return
        sel = {int(i) for i in self.tree.selection()}
        if not sel:
            self.toast(t("t_select_first"), IC_INFO)
            return
        flds = schema_fields(self.schema)
        lines = ["\t".join(headers(self.schema))]
        for i in self.view_index:                 # in the order shown on screen
            if i in sel:
                lines.append("\t".join(str(self.rows[i].get(f, ""))
                                       for f in flds))
        self.root.clipboard_clear()
        self.root.clipboard_append("\r\n".join(lines))
        self.toast(t("t_copied", n=len(sel)), IC_COPY)

    def select_all(self, *_):
        if self._in_entry():
            return
        items = self.tree.get_children()
        if items:
            self.tree.selection_set(items)

    # -------------------------------------------------------------------- actions

    def confirm_discard(self):
        if not self.dirty:
            return True
        ans = self.ask(t("unsaved_title"), IC_SAVE, t("unsaved_head"),
                       t("unsaved_body"),
                       [(t("btn_save"), "save", True),
                        (t("btn_dont_save"), "skip", False),
                        (t("btn_cancel"), None, False)])
        if ans is None:
            return False
        return self.save() if ans == "save" else True

    def new_base(self, *_):
        """New database: first choose its type, that is, its set of columns."""
        menu = self._menu()
        for kod, tpl in TEMPLATES.items():
            mark = "  ● " if self.schema.get("name") == kod else "     "
            menu.add_command(label=mark + t(tpl["tkey"]),
                             command=lambda k=kod: self._new_from_template(k))
        menu.add_separator()
        menu.add_command(label="     " + t("tpl_custom"),
                         command=self._new_custom)
        self._popup(menu, self.btn_new)

    def _new_from_template(self, kod):
        tpl = copy.deepcopy(TEMPLATES[kod])
        tpl["name"] = kod
        self._start_base(tpl)

    def _new_custom(self):
        """The dialog for defining your own columns."""
        if self.dirty and not self.confirm_discard():
            return
        res = SchemaDialog(self, [col_title(c) for c in data_cols(self.schema)]).result
        if res:
            self._start_base(res, confirmed=True)

    def _start_base(self, schema, confirmed=False):
        """Open an empty database and switch it to the given schema."""
        if not confirmed and not self.confirm_discard():
            return
        self.schema = schema
        self.rows, self.path, self.dirty, self.new_from = [], None, False, None
        self._undo.clear()
        self._redo.clear()
        self.sort_col, self.sort_desc = None, False
        self.hint = t("hint_new")
        self._rebuild_ui()               # the columns changed, so rebuild
        self.toast(schema_title(schema), IC_NEW)

    def open_base(self, *_):
        if not self.confirm_discard():
            return
        path = filedialog.askopenfilename(
            title=t("fd_open"),
            initialdir=self._dialog_dir(),
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"),
                       (t("ft_all"), "*.*")])
        if path:
            self._load_base(path)

    def _load_base(self, path):
        """
        Load a file as the database. The columns come from the file first:
        from the hidden sheet if we saved it, otherwise from the header row.
        If neither works the current schema is kept.
        """
        busy = Busy(self, t("busy_load", name=os.path.basename(path)))
        try:
            file_schema = read_schema(path)
            schema = file_schema or self.schema
            # with_src=True: this is our own database, so read the source column too
            rows = load_rows(path, schema, with_src=True)
        except Exception as e:
            busy.close()                  # err oynasi ustma-ust tushmasin
            self.error(t("err_read"), str(e))
            return
        busy.close()

        columns_changed = (schema is not self.schema)
        self.schema = schema
        self.rows = rows
        self.renumber()
        self.path, self.dirty, self.new_from = path, False, None
        self._undo.clear()
        self._redo.clear()
        self.sort_col, self.sort_desc = None, False
        self._remember(path)
        self.hint = t("hint_next", n=len(rows) + 1)

        if columns_changed:
            self._rebuild_ui()            # columns boshqacha — jadval qaytadan
        else:
            self.refresh()
            self.set_hint(self.hint)
        self.toast(t("t_loaded", n=len(rows)), IC_OPEN)

    # -------------------------------------------------------------- recent files

    def _remember(self, path):
        """Move a file to the front of the recent-files list."""
        p = os.path.abspath(path)
        self.recent = [p] + [x for x in self.recent
                             if os.path.normcase(x) != os.path.normcase(p)]
        del self.recent[RECENT_LIMIT:]
        self.last_dir = os.path.dirname(p)
        self._save_config()

    def show_recent(self):
        items = [p for p in self.recent if os.path.exists(p)]
        if not items:
            self.toast(t("t_recent_empty"), IC_INFO)
            return

        menu = self._menu()
        for p in items:
            menu.add_command(label="  " + pretty_path(p, 58),
                             command=lambda pp=p: self.open_recent(pp))
        menu.add_separator()
        menu.add_command(label="  " + t("menu_clear"), command=self._clear_recent)
        self._popup(menu, self.btn_recent)

    def open_recent(self, path):
        if not os.path.exists(path):
            self.error(t("err_notfound"), path, IC_WARN)
            self.recent = [p for p in self.recent if p != path]
            self._save_config()
            return
        if self.confirm_discard():
            self._load_base(path)

    def _clear_recent(self):
        self.recent = []
        self._save_config()
        self.toast(t("t_recent_cleared"), IC_HISTORY)

    def add_contacts(self):
        paths = filedialog.askopenfilenames(
            title=t("fd_add"),
            filetypes=[(t("ft_contacts"),
                        "*.xlsx *.xlsm *.csv *.txt *.tsv *.vcf *.vcard"),
                       ("Excel", "*.xlsx *.xlsm"),
                       (t("ft_text"), "*.txt *.tsv *.csv"),
                       (t("ft_vcard"), "*.vcf *.vcard"),
                       (t("ft_all"), "*.*")],
            initialdir=self._dialog_dir())
        if paths:
            self._import_paths(paths)

    def _on_drop(self, event):
        """Accept files dropped onto the window."""
        try:
            paths = [p for p in self.root.tk.splitlist(event.data)
                     if os.path.isfile(p)]
        except Exception:
            return getattr(event, "action", "copy")

        if paths:
            # Return from the event at once. Opening a dialog here would freeze
            # the application the file was dragged from (Explorer) until we answer.
            self.root.after(60, lambda: self._import_paths(paths))
        return getattr(event, "action", "copy")

    def _remember_source(self, src):
        """Remember a source value so it can be picked from the list next time."""
        if not src:
            return
        self.sources = [src] + [s for s in self.sources if s != src]
        del self.sources[SOURCE_LIMIT:]
        self._save_config()

    def _ask_source(self, label, count, offer_all=False):
        """
        Ask for the source value, offering previously used ones.
        Returns (value, apply_to_rest), or (None, False) when cancelled.
        """
        while True:
            res = self.ask(
                t("src_title"), IC_PHONE, t("src_head"),
                t("src_body", file=label, n=count),
                [(t("btn_add"), "ok", True), (t("btn_cancel"), None, False)],
                entry="", values=self.sources,
                check=(t("apply_all"), False) if offer_all else None)
            if res is None:
                return None, False

            apply_all = bool(self.last_check)
            if res[1]:
                self._remember_source(res[1])
                return res[1], apply_all

            # left empty: ask for confirmation
            if self.ask(t("srcempty_title"), IC_WARN,
                        t("srcempty_head"), t("srcempty_body"),
                        [(t("btn_back"), None, True),
                         (t("btn_leave_empty"), "empty", False)]) == "empty":
                return "", apply_all

    def _import_paths(self, paths):
        added_total = skipped_total = 0
        start_index = len(self.rows)
        self._push_undo(t("u_files", n=len(paths)))
        shared_src = None                   # "use this for the remaining files"

        for n, path in enumerate(paths, start=1):
            fname = os.path.basename(path)
            busy = Busy(self, t("busy_read", n=n, total=len(paths), name=fname))
            try:
                incoming = load_rows(path, self.schema)
            except Exception as e:
                busy.close()
                self.error(t("err_file_read", file=fname), str(e))
                continue
            busy.close()

            if not incoming:
                self.error(t("err_file_empty", file=fname),
                           t("err_file_empty_b"), IC_EMPTY)
                continue

            # --- ask for the source BY HAND; nothing is filled in automatically
            if shared_src is not None:
                src = shared_src
            else:
                src, apply_all = self._ask_source(
                    fname, len(incoming), offer_all=(n < len(paths)))
                if src is None:
                    continue                # this file was cancelled
                if apply_all:
                    shared_src = src

            # --- count the duplicates up front
            existing = {self.key_of(r) for r in self.rows}
            existing.discard("")

            seen, dup_count = set(), 0
            for row in incoming:
                k = self.key_of(row)
                if k and (k in existing or k in seen):
                    dup_count += 1
                if k:
                    seen.add(k)

            skip_dups = False
            if dup_count:
                ans = self.ask(
                    t("dup_title"), IC_WARN, t("dup_head", n=dup_count),
                    t("dup_body", file=fname),
                    [(t("btn_add"), "add", True),
                     (t("btn_skip"), "skip", False),
                     (t("btn_cancel"), None, False)],
                    color=self.pal["dup_fg"])
                if ans is None:
                    continue
                skip_dups = (ans == "skip")

            # --- append
            src_col = role_col(self.schema, "src")
            added = set()
            for row in incoming:
                k = self.key_of(row)
                if skip_dups and k and (k in existing or k in added):
                    skipped_total += 1
                    continue
                if k:
                    added.add(k)
                if src_col:
                    row[src_col] = src    # aynan siz yozgan raqam
                self.rows.append(row)
                added_total += 1

        if added_total:
            self.renumber()
            self.new_from = start_index
            self.dirty = True
            self.last_dir = os.path.dirname(os.path.abspath(paths[-1]))
            self._save_config()
            self.refresh()
            msg = t("t_added", n=added_total)
            if skipped_total:
                msg += t("t_skipped", n=skipped_total)
            self.toast(msg, IC_ADD)
            self.set_hint(t("hint_added"))
        else:
            self._undo.pop()              # nothing changed, drop the undo step
            self.refresh()

    # ------------------------------------------------------ paste from clipboard

    def paste_rows(self, *_):
        """Append rows copied straight out of Excel or a text list."""
        if self._in_entry():
            return
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""

        table = parse_text(text.splitlines()) if text.strip() else []
        incoming = rows_from_table(table, self.schema) if table else []
        if not incoming:
            self.error(t("err_clip"), t("err_clip_b"), IC_INFO)
            return

        src, _ = self._ask_source(t("clip_label"), len(incoming))
        if src is None:
            return

        self._push_undo(t("u_files", n=1))
        start_index = len(self.rows)
        src_col = role_col(self.schema, "src")
        for row in incoming:
            if src_col:
                row[src_col] = src
            self.rows.append(row)
        self.renumber()
        self.new_from = start_index
        self.dirty = True
        self.refresh()
        self.toast(t("t_added", n=len(incoming)), IC_ADD)
        self.set_hint(t("hint_added"))

    # --------------------------------------------------------------- data quality

    def normalize_phones(self):
        """Bring every phone number into the same shape."""
        pc = role_col(self.schema, "phone")
        if not self.rows or not pc:
            self.error(t("err_empty_base"), t("err_empty_base_b"), IC_EMPTY)
            return

        updated = [pretty_phone(r.get(pc)) for r in self.rows]
        changed = sum(1 for r, y in zip(self.rows, updated) if r.get(pc) != y)
        if not changed:
            self.toast(t("t_normalized_none"), IC_CHECK)
            return

        self._push_undo(t("u_normalize"))
        for r, y in zip(self.rows, updated):
            r[pc] = y
        self.dirty = True
        self.refresh()
        self.toast(t("t_normalized", n=changed), IC_PHONE)

    def toggle_bad(self):
        """The "suspicious rows only" filter."""
        if self.only_bad.get():
            changed = sum(1 for r in self.rows if row_problem(r, self.schema))
            if not changed:
                self.only_bad.set(False)
                self.toast(t("t_no_problems"), IC_CHECK)
                return
            self.toast(t("t_problems", n=changed), IC_WARN)
        self.refresh_table()

    # -------------------------------------------------------------------- export

    def export_csv(self):
        if not self.rows:
            self.error(t("err_empty_base"), t("err_empty_base_b"), IC_EMPTY)
            return
        flds = schema_fields(self.schema)
        path = filedialog.asksaveasfilename(
            title=t("fd_csv"), defaultextension=".csv",
            initialdir=self._dialog_dir(), initialfile="baza.csv",
            filetypes=[(t("ft_csv"), "*.csv")])
        if not path:
            return
        try:
            # utf-8-sig so Excel opens the CSV with the right encoding
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(headers(self.schema))
                for r in self.rows:
                    w.writerow([r.get(f_, "") for f_ in flds])
        except PermissionError:
            self.error(t("err_busy"), t("err_busy_b"))
            return
        except Exception as e:
            log_error(type(e), e, e.__traceback__)
            self.error(t("err_save"), str(e))
            return
        self._remember(path)
        self.toast(t("t_exported", n=len(self.rows)), IC_SAVEAS)

    def export_vcard(self):
        """Write a .vcf to load back onto a phone; needs a name or phone column."""
        nc = role_col(self.schema, "name")
        pc = role_col(self.schema, "phone")
        if not self.rows or not (nc or pc):
            self.error(t("err_empty_base"), t("err_empty_base_b"), IC_EMPTY)
            return
        ic = role_col(self.schema, "id")
        sc = role_col(self.schema, "src")
        path = filedialog.asksaveasfilename(
            title=t("fd_vcard"), defaultextension=".vcf",
            initialdir=self._dialog_dir(), initialfile="contacts.vcf",
            filetypes=[(t("ft_vcf"), "*.vcf")])
        if not path:
            return

        def esc(s):
            return (cell_text(s).replace("\\", "\\\\").replace(";", "\\;")
                    .replace(",", "\\,").replace("\n", "\\n"))

        try:
            with open(path, "w", encoding="utf-8", newline="\r\n") as f:
                for r in self.rows:
                    tel = cell_text(r.get(pc)) if pc else ""
                    ism = (cell_text(r.get(nc)) if nc else "") or tel
                    part = ism.split()
                    familiya = part[-1] if len(part) > 1 else ""
                    otismi = " ".join(part[:-1]) if len(part) > 1 else ism
                    f.write("BEGIN:VCARD\nVERSION:3.0\n")
                    f.write(f"FN:{esc(ism)}\n")
                    f.write(f"N:{esc(familiya)};{esc(otismi)};;;\n")
                    if tel:
                        f.write(f"TEL;TYPE=CELL:{esc(tel)}\n")
                    if ic and cell_text(r.get(ic)):
                        f.write(f"UID:{esc(r.get(ic))}\n")
                    if sc and cell_text(r.get(sc)):
                        f.write(f"NOTE:{esc(t('h_source'))}: "
                                f"{esc(r.get(sc))}\n")
                    f.write("END:VCARD\n")
        except PermissionError:
            self.error(t("err_busy"), t("err_busy_b"))
            return
        except Exception as e:
            log_error(type(e), e, e.__traceback__)
            self.error(t("err_save"), str(e))
            return
        self._remember(path)
        self.toast(t("t_exported", n=len(self.rows)), IC_SAVEAS)

    def edit_source(self):
        sc = role_col(self.schema, "src")
        if not sc:                        # this schema has no source column
            self.toast(t("t_no_problems"), IC_INFO)
            return
        sel = self.tree.selection()
        if not sel:
            self.error(t("err_norow"), t("err_norow_multi"), IC_INFO)
            return
        idxs = [int(i) for i in sel]
        res = self.ask(t("edit_src_title"), IC_EDIT,
                       t("edit_src_head", n=len(idxs)), t("edit_src_body"),
                       [(t("btn_save"), "ok", True),
                        (t("btn_cancel"), None, False)],
                       entry=cell_text(self.rows[idxs[0]].get(sc)),
                       values=self.sources)
        if res is None:
            return
        self._push_undo(t("u_src", n=len(idxs)))
        for i in idxs:
            self.rows[i][sc] = res[1]
        self._remember_source(res[1])
        self.dirty = True
        self.refresh_table()
        self.toast(t("t_src_changed", n=len(idxs)), IC_EDIT)

    def edit_row(self):
        """Edit every column of one row, opened by double-clicking it."""
        sel = self.tree.selection()
        if not sel:
            self.error(t("err_norow"), t("err_norow_one"), IC_INFO)
            return
        if len(sel) > 1:                  # bir changed tanlangan — faqat manba
            self.edit_source()
            return

        i = int(sel[0])
        row = self.rows[i]
        # every column except the row-number one is editable
        cols = [c for c in self.schema["cols"] if c.get("role") != "no"]
        nc = role_col(self.schema, "no")

        res = self.ask(t("edit_row_title"), IC_EDIT,
                       t("edit_row_head", no=row.get(nc, i + 1) if nc else i + 1),
                       t("edit_row_body"),
                       [(t("btn_save"), "ok", True),
                        (t("btn_cancel"), None, False)],
                       entry=[(col_title(c), cell_text(row.get(c["key"])))
                              for c in cols])
        if res is None:
            return

        updated = dict(zip((c["key"] for c in cols), res[1]))
        if all(cell_text(row.get(k)) == v for k, v in updated.items()):
            return                        # nothing changed

        self._push_undo(t("u_row", no=row.get(nc, i + 1) if nc else i + 1))
        row.update(updated)
        self.dirty = True
        self.refresh()                    # the key may have changed, so recount
        self.toast(t("t_row_updated"), IC_EDIT)

    def add_manual(self):
        """Add a single row by hand rather than from a file."""
        cols = [c for c in self.schema["cols"] if c.get("role") != "no"]
        res = self.ask(t("add_title"), IC_PERSONADD, t("add_head"),
                       t("add_body"),
                       [(t("btn_add"), "ok", True),
                        (t("btn_cancel"), None, False)],
                       entry=[(col_title(c), "") for c in cols])
        if res is None:
            return

        row = blank_row(self.schema)
        row.update(dict(zip((c["key"] for c in cols), res[1])))
        if not any(cell_text(v) for k, v in row.items()
                   if k != role_col(self.schema, "no")):
            self.error(t("add_empty"), t("add_empty_b"), IC_INFO)
            return

        self._push_undo(t("u_manual"))
        self.rows.append(row)
        self.renumber()
        self.new_from = len(self.rows) - 1
        self.dirty = True
        self.refresh()
        self.toast(t("t_contact_added"), IC_PERSONADD)

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            self.error(t("err_norow"), t("err_norow_any"), IC_INFO)
            return
        ans = self.ask(t("del_title"), IC_DELETE, t("del_head", n=len(sel)),
                       t("del_body"),
                       [(t("btn_delete"), True, True),
                        (t("btn_cancel"), None, False)],
                       color=self.pal["dup_fg"])
        if not ans:
            return
        n = len(sel)
        self._push_undo(t("u_del", n=n))
        for i in sorted((int(x) for x in sel), reverse=True):
            del self.rows[i]
        self.renumber()
        self.new_from = None
        self.dirty = True
        self.refresh()
        self.toast(t("t_deleted", n=n), IC_DELETE)

    # -------------------------------------------------------------------- saving

    def save(self, *_):
        if not self.rows:
            self.error(t("err_empty_base"), t("err_empty_base_b"), IC_EMPTY)
            return False
        return self._write(self.path) if self.path else self.save_as()

    def save_as(self):
        path = filedialog.asksaveasfilename(
            title=t("fd_save"), defaultextension=".xlsx",
            initialdir=self._dialog_dir(),
            initialfile=os.path.basename(self.path) if self.path else "baza.xlsx",
            filetypes=[(t("ft_xlsx"), "*.xlsx")])
        return self._write(path) if path else False

    def _write(self, path):
        """
        Write to a temporary file first, take a `.bak` copy of the old one,
        and only then swap it in. If the write fails the old database is
        still there.
        """
        self.compute_dups()
        head = headers(self.schema)          # column nomlari — joriy tilda
        tmp = path + ".writing.tmp"

        # a large database takes seconds to write; do not look frozen
        busy = (Busy(self, t("busy_save", name=os.path.basename(path)))
                if len(self.rows) > 2000 else None)
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Database"

            head_fill = PatternFill("solid", start_color=XL_HEAD_BG)
            head_font = Font(bold=True, color="FFFFFF", size=11)
            dup_fill = PatternFill("solid", start_color=XL_DUP_BG)
            dup_font = Font(bold=True, color=XL_DUP_FG)
            thin = Side(style="thin", color="D9D9D9")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            ws.append(head)
            for c in range(1, len(head) + 1):
                cell = ws.cell(row=1, column=c)
                cell.fill, cell.font, cell.border = head_fill, head_font, border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 24

            cols = self.schema["cols"]
            markaz = Alignment(horizontal="center")

            for r_i, row in enumerate(self.rows, start=2):
                is_dup = False
                key = self.key_of(row)
                if key and key in self.dup_keys:
                    is_dup = True

                for c_i, c in enumerate(cols, start=1):
                    cell = ws.cell(row=r_i, column=c_i,
                                   value=row.get(c["key"], ""))
                    cell.border = border
                    if c.get("role") == "no":
                        cell.alignment = markaz
                    else:
                        # keep numbers as text so +998... does not lose its leading characters
                        cell.number_format = "@"
                    if is_dup:              # dasturdagidek — butun row qizil
                        cell.fill, cell.font = dup_fill, dup_font

            for i, c in enumerate(cols, start=1):
                # roughly convert the on-screen pixel width into Excel's own unit
                ws.column_dimensions[get_column_letter(i)].width = \
                    max(7, min(60, round(c["width"] / 7.5)))
            ws.freeze_panes = "A2"
            last = get_column_letter(len(head))
            ws.auto_filter.ref = f"A1:{last}{len(self.rows) + 1}"

            # write the schema to a hidden sheet so reopening the file restores
            # exactly these columns
            sh = wb.create_sheet(SCHEMA_SHEET)
            sh["A1"] = json.dumps(self.schema, ensure_ascii=False)
            sh.sheet_state = "hidden"

            wb.save(tmp)

            if os.path.exists(path):        # do not lose the previous version
                try:
                    shutil.copy2(path, path + ".bak")
                except Exception:
                    pass
            os.replace(tmp, path)           # bitta qadamda almashtiriladi
        except PermissionError:
            if busy:
                busy.close()
            self.error(t("err_busy"), t("err_busy_b"))
            return False
        except Exception as e:
            if busy:
                busy.close()
            log_error(type(e), e, e.__traceback__)
            self.error(t("err_save"), t("err_save_b", err=e))
            return False
        finally:
            if os.path.exists(tmp):         # yarim yozilgan file qolmasin
                try:
                    os.remove(tmp)
                except OSError:
                    pass

        if busy:
            busy.close()

        self.path, self.dirty, self.new_from = path, False, None
        self._remember(path)
        self.refresh()
        self.set_hint(t("hint_saved", path=path))
        self.toast(t("t_saved", n=len(self.rows)), IC_SAVE)
        return True

    def open_in_excel(self):
        if not self.path or not os.path.exists(self.path):
            self.error(t("err_notfound"), t("err_notfound_b"), IC_INFO)
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.path)                       # noqa
            else:
                subprocess.Popen(["open", self.path])
        except Exception as e:
            self.error(t("err_open"), str(e))

    def on_close(self):
        if self.confirm_discard():
            self._save_config(geometry=True)
            self.root.destroy()


# ===================================================================== startup

def make_icon(path="icon.ico"):
    """Generate the icon for the .exe; build.bat calls this."""
    from PIL import Image, ImageDraw, ImageFont

    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    drw = ImageDraw.Draw(img)
    drw.rounded_rectangle([0, 0, size - 1, size - 1], radius=56,
                          fill=PALETTES["light"]["accent"])

    font_file = next((p for p in Art.FONT_PATHS if os.path.exists(p)), None)
    if font_file:
        drw.text((size / 2, size / 2), IC_LOGO,
                 font=ImageFont.truetype(font_file, 150), fill="#ffffff",
                 anchor="mm")

    img.save(path, sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])
    return path


def make_root():
    """A window that accepts dropped files, falling back to a plain one."""
    if HAS_DND:
        try:
            return TkinterDnD.Tk()
        except Exception:
            pass
    return tk.Tk()


def main():
    if "--make-icon" in sys.argv:           # only used by the build script
        made = make_icon()
        try:
            print(made)
        except Exception:                   # pythonw has no console
            pass
        return

    sys.excepthook = log_error              # no console, so log the error

    try:                                   # crisp text on Windows
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    try:                                   # own icon in the taskbar
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "DatabaseManager.App")
    except Exception:
        pass

    root = make_root()
    DatabaseApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:               # report even if the window never opens
        log_error(type(exc), exc, exc.__traceback__)
        message_box(f"Dastur ishga tushmadi.\n\n{type(exc).__name__}: {exc}\n\n"
                    f"Batafsil ma'lumot:\n{LOG_PATH}")
        raise
