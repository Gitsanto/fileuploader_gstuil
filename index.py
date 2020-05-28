#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Author: Santosh Rai
# Version: 1.00 (20200520)
# upload files to GCP with gsutil

import platform
import os
import io
import sys
import codecs
import csv
import subprocess
import datetime as dt

sys.stdout = codecs.getwriter("utf-8")(sys.stdout)

if sys.version_info[0] == 3:
    # for Python3
    from tkinter import Frame, Label, Message, StringVar, Canvas, Button, Menu, DoubleVar
    import tkinter.filedialog as filedialogprompt
    from tkinter.ttk import Scrollbar
    from tkinter.constants import *
    import time
    from tkinter import ttk
    from tkinter import messagebox as message_box


else:
    # for Python2
    from Tkinter import Frame, Label, Message, StringVar, Canvas, Button, Menu, DoubleVar
    import tkFileDialog as filedialogprompt
    from ttk import Scrollbar
    from Tkconstants import *
    import time
    import tkMessageBox as message_box


OS = platform.system()

ASCII_ZENKAKU_CHARS = (
    u'ａ', u'ｂ', u'ｃ', u'ｄ', u'ｅ', u'ｆ', u'ｇ', u'ｈ', u'ｉ', u'ｊ', u'ｋ',
    u'ｌ', u'ｍ', u'ｎ', u'ｏ', u'ｐ', u'ｑ', u'ｒ', u'ｓ', u'ｔ', u'ｕ', u'ｖ',
    u'ｗ', u'ｘ', u'ｙ', u'ｚ',
    u'Ａ', u'Ｂ', u'Ｃ', u'Ｄ', u'Ｅ', u'Ｆ', u'Ｇ', u'Ｈ', u'Ｉ', u'Ｊ', u'Ｋ',
    u'Ｌ', u'Ｍ', u'Ｎ', u'Ｏ', u'Ｐ', u'Ｑ', u'Ｒ', u'Ｓ', u'Ｔ', u'Ｕ', u'Ｖ',
    u'Ｗ', u'Ｘ', u'Ｙ', u'Ｚ',
    u'！', u'”', u'＃', u'＄', u'％', u'＆', u'’', u'（', u'）', u'＊', u'＋',
    u'，', u'－', u'．', u'／', u'：', u'；', u'＜', u'＝', u'＞', u'？', u'＠',
    u'［', u'￥', u'］', u'＾', u'＿', u'‘', u'｛', u'｜', u'｝', u'～', u'　'
)

ASCII_HANKAKU_CHARS = (
    u'a', u'b', u'c', u'd', u'e', u'f', u'g', u'h', u'i', u'j', u'k',
    u'l', u'm', u'n', u'o', u'p', u'q', u'r', u's', u't', u'u', u'v',
    u'w', u'x', u'y', u'z',
    u'A', u'B', u'C', u'D', u'E', u'F', u'G', u'H', u'I', u'J', u'K',
    u'L', u'M', u'N', u'O', u'P', u'Q', u'R', u'S', u'T', u'U', u'V',
    u'W', u'X', u'Y', u'Z',
    u'!', u'"', u'#', u'$', u'%', u'&', u'\'', u'(', u')', u'*', u'+',
    u',', u'-', u'.', u'/', u':', u';', u'<', u'=', u'>', u'?', u'@',
    u'[', u"¥", u']', u'^', u'_', u'`', u'{', u'|', u'}', u'~', u' '
)

KANA_ZENKAKU_CHARS = (
    u'ア', u'イ', u'ウ', u'エ', u'オ', u'カ', u'キ', u'ク', u'ケ', u'コ',
    u'サ', u'シ', u'ス', u'セ', u'ソ', u'タ', u'チ', u'ツ', u'テ', u'ト',
    u'ナ', u'ニ', u'ヌ', u'ネ', u'ノ', u'ハ', u'ヒ', u'フ', u'ヘ', u'ホ',
    u'マ', u'ミ', u'ム', u'メ', u'モ', u'ヤ', u'ユ', u'ヨ',
    u'ラ', u'リ', u'ル', u'レ', u'ロ', u'ワ', u'ヲ', u'ン',
    u'ァ', u'ィ', u'ゥ', u'ェ', u'ォ', u'ッ', u'ャ', u'ュ', u'ョ',
    u'。', u'、', u'・', u'゛', u'゜', u'「', u'」', u'ー'
)

KANA_HANKAKU_CHARS = (
    u'ｱ', u'ｲ', u'ｳ', u'ｴ', u'ｵ', u'ｶ', u'ｷ', u'ｸ', u'ｹ', u'ｺ',
    u'ｻ', u'ｼ', u'ｽ', u'ｾ', u'ｿ', u'ﾀ', u'ﾁ', u'ﾂ', u'ﾃ', u'ﾄ',
    u'ﾅ', u'ﾆ', u'ﾇ', u'ﾈ', u'ﾉ', u'ﾊ', u'ﾋ', u'ﾌ', u'ﾍ', u'ﾎ',
    u'ﾏ', u'ﾐ', u'ﾑ', u'ﾒ', u'ﾓ', u'ﾔ', u'ﾕ', u'ﾖ',
    u'ﾗ', u'ﾘ', u'ﾙ', u'ﾚ', u'ﾛ', u'ﾜ', u'ｦ', u'ﾝ',
    u'ｧ', u'ｨ', u'ｩ', u'ｪ', u'ｫ', u'ｯ', u'ｬ', u'ｭ', u'ｮ',
    u'｡', u'､', u'･', u'ﾞ', u'ﾟ', u'｢', u'｣', u'ｰ'
)

DIGIT_ZENKAKU_CHARS = (
    u'０', u'１', u'２', u'３', u'４', u'５', u'６', u'７', u'８', u'９'
)

DIGIT_HANKAKU_CHARS = (
    u'0', u'1', u'2', u'3', u'4', u'5', u'6', u'7', u'8', u'9'
)

KANA_TEN_MAP = (
    (u'ガ', u'ｶ'), (u'ギ', u'ｷ'), (u'グ', u'ｸ'), (u'ゲ', u'ｹ'), (u'ゴ', u'ｺ'),
    (u'ザ', u'ｻ'), (u'ジ', u'ｼ'), (u'ズ', u'ｽ'), (u'ゼ', u'ｾ'), (u'ゾ', u'ｿ'),
    (u'ダ', u'ﾀ'), (u'ヂ', u'ﾁ'), (u'ヅ', u'ﾂ'), (u'デ', u'ﾃ'), (u'ド', u'ﾄ'),
    (u'バ', u'ﾊ'), (u'ビ', u'ﾋ'), (u'ブ', u'ﾌ'), (u'ベ', u'ﾍ'), (u'ボ', u'ﾎ'),
    (u'ヴ', u'ｳ')
)

KANA_MARU_MAP = (
    (u'パ', u'ﾊ'), (u'ピ', u'ﾋ'), (u'プ', u'ﾌ'), (u'ペ', u'ﾍ'), (u'ポ', u'ﾎ')
)


ascii_zh_table = {}
ascii_hz_table = {}
kana_zh_table = {}
kana_hz_table = {}
digit_zh_table = {}
digit_hz_table = {}

for (az, ah) in zip(ASCII_ZENKAKU_CHARS, ASCII_HANKAKU_CHARS):
    ascii_zh_table[az] = ah
    ascii_hz_table[ah] = az

for (kz, kh) in zip(KANA_ZENKAKU_CHARS, KANA_HANKAKU_CHARS):
    kana_zh_table[kz] = kh
    kana_hz_table[kh] = kz

for (dz, dh) in zip(DIGIT_ZENKAKU_CHARS, DIGIT_HANKAKU_CHARS):
    digit_zh_table[dz] = dh
    digit_hz_table[dh] = dz

kana_ten_zh_table = {}
kana_ten_hz_table = {}
kana_maru_zh_table = {}
kana_maru_hz_table = {}

for (ktz, kth) in KANA_TEN_MAP:
    kana_ten_zh_table[ktz] = kth
    kana_ten_hz_table[kth] = ktz

for (kmz, kmh) in KANA_MARU_MAP:
    kana_maru_zh_table[kmz] = kmh
    kana_maru_hz_table[kmh] = kmz

del ASCII_ZENKAKU_CHARS, ASCII_HANKAKU_CHARS, \
    KANA_ZENKAKU_CHARS, KANA_HANKAKU_CHARS,\
    DIGIT_ZENKAKU_CHARS, DIGIT_HANKAKU_CHARS,\
    KANA_TEN_MAP, KANA_MARU_MAP

kakko_zh_table = {
    u'｟': u'⦅', u'｠': u'⦆',
    u'『': u'｢', u'』': u'｣',
    u'〚': u'⟦', u'〛': u'⟧',
    u'〔': u'❲', u'〕': u'❳',
    u'〘': u'⟬', u'〙': u'⟭',
    u'《': u'⟪', u'》': u'⟫',
    u'【': u'(', u'】': u')',
    u'〖': u'(', u'〗': u')'
}

kakko_hz_table = {}
for k, v in kakko_zh_table.items():
    kakko_hz_table[v] = k




class MOJI:
    @staticmethod
    def zen2han(text="", ascii_=True, digit=True, kana=True, kakko=True, ignore=()):
        
        result = []

        for c in text:
            if c in ignore:
                result.append(c)
            elif ascii_ and (c in ascii_zh_table):
                result.append(ascii_zh_table[c])
            elif digit and (c in digit_zh_table):
                result.append(digit_zh_table[c])
            # elif kana and (c in kana_zh_table):
            #     result.append(kana_zh_table[c])
            elif kana and (c in kana_ten_zh_table):
                result.append(kana_ten_zh_table[c] + u'ﾞ')
            elif kana and (c in kana_maru_zh_table):
                result.append(kana_maru_zh_table[c] + u'ﾟ')
            elif kakko and (c in kakko_zh_table):
                result.append(kakko_zh_table[c])
            else:
                result.append(c)

        return "".join(result)

    @staticmethod
    def han2zen(text, ascii_=True, digit=True, kana=True, kakko=True, ignore=()):
        result = []

        for i, c in enumerate(text):
            if c == u'ﾞ' or c == u'ﾟ':
                continue
            elif c in ignore:
                result.append(c)
            elif ascii_ and (c in ascii_hz_table):
                result.append(ascii_hz_table[c])
            elif digit and (c in digit_hz_table):
                result.append(digit_hz_table[c])
            elif kana and (c in kana_ten_hz_table) and (text[i+1] == u'ﾞ'):
                result.append(kana_ten_hz_table[c])
            elif kana and (c in kana_maru_hz_table) and (text[i+1] == u'ﾟ'):
                result.append(kana_maru_hz_table[c])
            elif kana and (c in kana_hz_table):
                result.append(kana_hz_table[c])
            elif kakko and (c in kakko_hz_table):
                result.append(kakko_hz_table[c])
            else:
                result.append(c)

        return "".join(result)

class Mousewheel_Support(object):

    # implemetation of singleton pattern
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = object.__new__(cls)
        return cls._instance

    def __init__(self, root, horizontal_factor=2, vertical_factor=2):

        self._active_area = None

        if isinstance(horizontal_factor, int):
            self.horizontal_factor = horizontal_factor
        else:
            raise Exception("Vertical factor must be an integer.")

        if isinstance(vertical_factor, int):
            self.vertical_factor = vertical_factor
        else:
            raise Exception("Horizontal factor must be an integer.")

        if OS == "Linux":
            root.bind_all('<4>', self._on_mousewheel,  add='+')
            root.bind_all('<5>', self._on_mousewheel,  add='+')
        else:
            # Windows and MacOS
            root.bind_all("<MouseWheel>", self._on_mousewheel,  add='+')

    def _on_mousewheel(self, event):
        if self._active_area:
            self._active_area.onMouseWheel(event)

    def _mousewheel_bind(self, widget):
        self._active_area = widget

    def _mousewheel_unbind(self):
        self._active_area = None

    def add_support_to(self, widget=None, xscrollbar=None, yscrollbar=None, what="units", horizontal_factor=None, vertical_factor=None):
        if xscrollbar is None and yscrollbar is None:
            return

        if xscrollbar is not None:
            horizontal_factor = horizontal_factor or self.horizontal_factor

            xscrollbar.onMouseWheel = self._make_mouse_wheel_handler(
                widget, 'x', self.horizontal_factor, what)
            xscrollbar.bind('<Enter>', lambda event,
                            scrollbar=xscrollbar: self._mousewheel_bind(scrollbar))
            xscrollbar.bind('<Leave>', lambda event: self._mousewheel_unbind())

        if yscrollbar is not None:
            vertical_factor = vertical_factor or self.vertical_factor

            yscrollbar.onMouseWheel = self._make_mouse_wheel_handler(
                widget, 'y', self.vertical_factor, what)
            yscrollbar.bind('<Enter>', lambda event,
                            scrollbar=yscrollbar: self._mousewheel_bind(scrollbar))
            yscrollbar.bind('<Leave>', lambda event: self._mousewheel_unbind())

        main_scrollbar = yscrollbar if yscrollbar is not None else xscrollbar

        if widget is not None:
            if isinstance(widget, list) or isinstance(widget, tuple):
                list_of_widgets = widget
                for widget in list_of_widgets:
                    widget.bind(
                        '<Enter>', lambda event: self._mousewheel_bind(widget))
                    widget.bind(
                        '<Leave>', lambda event: self._mousewheel_unbind())

                    widget.onMouseWheel = main_scrollbar.onMouseWheel
            else:
                widget.bind(
                    '<Enter>', lambda event: self._mousewheel_bind(widget))
                widget.bind('<Leave>', lambda event: self._mousewheel_unbind())

                widget.onMouseWheel = main_scrollbar.onMouseWheel

    @staticmethod
    def _make_mouse_wheel_handler(widget, orient, factor=1, what="units"):
        view_command = getattr(widget, orient+'view')

        if OS == 'Linux':
            def onMouseWheel(event):
                if event.num == 4:
                    view_command("scroll", (-1)*factor, what)
                elif event.num == 5:
                    view_command("scroll", factor, what)

        elif OS == 'Windows':
            def onMouseWheel(event):
                view_command("scroll", (-1) *
                             int((event.delta/120)*factor), what)

        elif OS == 'Darwin':
            def onMouseWheel(event):
                view_command("scroll", event.delta, what)

        return onMouseWheel


class Scrolling_Area(Frame, object):

    def __init__(self, master, width=None, anchor=N, height=None, mousewheel_speed=2, scroll_horizontally=True, xscrollbar=None, scroll_vertically=True, yscrollbar=None, outer_background=None, inner_frame=Frame, **kw):
        Frame.__init__(self, master, class_=self.__class__)

        if outer_background:
            self.configure(background=outer_background)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._width = width
        self._height = height

        self.canvas = Canvas(self, background=outer_background,
                             highlightthickness=0, width=width, height=height)
        self.canvas.grid(row=0, column=0, sticky=N+E+W+S)

        if scroll_vertically:
            if yscrollbar is not None:
                self.yscrollbar = yscrollbar
            else:
                self.yscrollbar = Scrollbar(self, orient=VERTICAL)
                self.yscrollbar.grid(row=0, column=1, sticky=N+S)

            self.canvas.configure(yscrollcommand=self.yscrollbar.set)
            self.yscrollbar['command'] = self.canvas.yview
        else:
            self.yscrollbar = None

        if scroll_horizontally:
            if xscrollbar is not None:
                self.xscrollbar = xscrollbar
            else:
                self.xscrollbar = Scrollbar(self, orient=HORIZONTAL)
                self.xscrollbar.grid(row=1, column=0, sticky=E+W)

            self.canvas.configure(xscrollcommand=self.xscrollbar.set)
            self.xscrollbar['command'] = self.canvas.xview
        else:
            self.xscrollbar = None

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.innerframe = inner_frame(self.canvas, **kw)
        self.innerframe.pack(anchor=anchor)

        self.canvas.create_window(
            0, 0, window=self.innerframe, anchor='nw', tags="inner_frame")

        self.canvas.bind('<Configure>', self._on_canvas_configure)

        Mousewheel_Support(self).add_support_to(
            self.canvas, xscrollbar=self.xscrollbar, yscrollbar=self.yscrollbar)

    @property
    def width(self):
        return self.canvas.winfo_width()

    @width.setter
    def width(self, width):
        self.canvas.configure(width=width)

    @property
    def height(self):
        return self.canvas.winfo_height()

    @height.setter
    def height(self, height):
        self.canvas.configure(height=height)

    def set_size(self, width, height):
        self.canvas.configure(width=width, height=height)

    def _on_canvas_configure(self, event):
        width = max(self.innerframe.winfo_reqwidth(), event.width)
        height = max(self.innerframe.winfo_reqheight(), event.height)

        self.canvas.configure(scrollregion="0 0 %s %s" % (width, height))
        self.canvas.itemconfigure("inner_frame", width=width, height=height)

    def update_viewport(self):
        self.update()

        window_width = self.innerframe.winfo_reqwidth()
        window_height = self.innerframe.winfo_reqheight()

        if self._width is None:
            canvas_width = window_width
        else:
            canvas_width = min(self._width, window_width)

        if self._height is None:
            canvas_height = window_height
        else:
            canvas_height = min(self._height, window_height)

        self.canvas.configure(scrollregion="0 0 %s %s" % (
            window_width, window_height), width=canvas_width, height=canvas_height)
        self.canvas.itemconfigure(
            "inner_frame", width=window_width, height=window_height)


class Cell(Frame):
    """Base class for cells"""


class Data_Cell(Cell):
    def __init__(self, master, variable, anchor=W, bordercolor=None, borderwidth=1, padx=0, pady=0, background=None, foreground=None, font=None):
        Cell.__init__(self, master, background=background, highlightbackground=bordercolor,
                      highlightcolor=bordercolor, highlightthickness=borderwidth, bd=0)

        self._message_widget = Message(
            self, textvariable=variable, font=font, background=background, foreground=foreground)
        self._message_widget.pack(
            expand=True, padx=padx, pady=pady, anchor=anchor)


class Header_Cell(Cell):
    def __init__(self, master, text, bordercolor=None, borderwidth=1, padx=0, pady=0, background=None, foreground=None, font=None, anchor=CENTER, separator=True):
        Cell.__init__(self, master, background=background, highlightbackground=bordercolor,
                      highlightcolor=bordercolor, highlightthickness=borderwidth, bd=0)
        self.pack_propagate(False)

        self._header_label = Label(
            self, text=text, background=background, foreground=foreground, font=font)
        self._header_label.pack(padx=padx, pady=pady, expand=True)

        if separator and bordercolor is not None:
            separator = Frame(self, height=2, background=bordercolor,
                              bd=0, highlightthickness=0, class_="Separator")
            separator.pack(fill=X, anchor=anchor)

        self.update()
        height = self._header_label.winfo_reqheight() + 2*padx
        width = self._header_label.winfo_reqwidth() + 40*pady

        self.configure(height=height, width=width)


class Table(Frame):
    def __init__(self, master, columns, column_weights=None, column_minwidths=None, height=500, minwidth=20, minheight=20, padx=5, pady=5, cell_font=None, cell_foreground="black", cell_background="white", cell_anchor=W, header_font=None, header_background="white", header_foreground="black", header_anchor=CENTER, bordercolor="#999999", innerborder=True, outerborder=True, stripped_rows=("#EEEEEE", "white"), on_change_data=None, mousewheel_speed=2, scroll_horizontally=False, scroll_vertically=True):
        outerborder_width = 1 if outerborder else 0

        Frame.__init__(self, master, bd=0)

        self._cell_background = cell_background
        self._cell_foreground = cell_foreground
        self._cell_font = cell_font
        self._cell_anchor = cell_anchor

        self._stripped_rows = stripped_rows

        self._padx = padx
        self._pady = pady

        self._bordercolor = bordercolor
        self._innerborder_width = 1 if innerborder else 0

        self._data_vars = []

        self._columns = columns

        self._number_of_rows = 0
        self._number_of_columns = len(columns)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._head = Frame(self, highlightbackground=bordercolor,
                           highlightcolor=bordercolor, highlightthickness=outerborder_width, bd=0)
        self._head.grid(row=0, column=0, sticky=E+W)

        header_separator = False if outerborder else True

        for j in range(len(columns)):
            column_name = columns[j]

            header_cell = Header_Cell(self._head, text=column_name, borderwidth=self._innerborder_width, font=header_font, background=header_background,
                                      foreground=header_foreground, padx=padx, pady=pady, bordercolor=bordercolor, anchor=header_anchor, separator=header_separator)
            header_cell.grid(row=0, column=j, sticky=N+E+W+S)

        add_scrollbars = scroll_horizontally or scroll_vertically
        if add_scrollbars:
            if scroll_horizontally:
                xscrollbar = Scrollbar(self, orient=HORIZONTAL)
                xscrollbar.grid(row=2, column=0, sticky=E+W)
            else:
                xscrollbar = None

            if scroll_vertically:
                yscrollbar = Scrollbar(self, orient=VERTICAL)
                yscrollbar.grid(row=1, column=1, sticky=N+S)
            else:
                yscrollbar = None

            scrolling_area = Scrolling_Area(self, width=self._head.winfo_reqwidth(
            ), height=height, scroll_horizontally=scroll_horizontally, xscrollbar=xscrollbar, scroll_vertically=scroll_vertically, yscrollbar=yscrollbar)
            scrolling_area.grid(row=1, column=0, sticky=E+W)

            self._body = Frame(scrolling_area.innerframe, highlightbackground=bordercolor,
                               highlightcolor=bordercolor, highlightthickness=outerborder_width, bd=0)
            self._body.pack()

            def on_change_data():
                scrolling_area.update_viewport()

        else:
            self._body = Frame(self, height=height, highlightbackground=bordercolor,
                               highlightcolor=bordercolor, highlightthickness=outerborder_width, bd=0)
            self._body.grid(row=1, column=0, sticky=N+E+W+S)

        if column_weights is None:
            for j in range(len(columns)):
                self._body.grid_columnconfigure(j, weight=1)
        else:
            for j, weight in enumerate(column_weights):
                self._body.grid_columnconfigure(j, weight=weight)

        if column_minwidths is not None:
            for j, minwidth in enumerate(column_minwidths):
                if minwidth is None:
                    header_cell = self._head.grid_slaves(row=0, column=j)[0]
                    minwidth = header_cell.winfo_reqwidth()

                self._body.grid_columnconfigure(j, minsize=minwidth)
        else:
            for j in range(len(columns)):
                header_cell = self._head.grid_slaves(row=0, column=j)[0]
                minwidth = header_cell.winfo_reqwidth()

                self._body.grid_columnconfigure(j, minsize=minwidth)

        self._on_change_data = on_change_data

    def _append_n_rows(self, n):
        number_of_rows = self._number_of_rows
        number_of_columns = self._number_of_columns

        for i in range(number_of_rows, number_of_rows+n):
            list_of_vars = []
            for j in range(number_of_columns):
                var = StringVar()
                list_of_vars.append(var)

                if self._stripped_rows:
                    cell = Data_Cell(self._body, borderwidth=self._innerborder_width, variable=var, bordercolor=self._bordercolor, padx=self._padx,
                                     pady=self._pady, background=self._stripped_rows[i % 2], foreground=self._cell_foreground, font=self._cell_font, anchor=self._cell_anchor)
                else:
                    cell = Data_Cell(self._body, borderwidth=self._innerborder_width, variable=var, bordercolor=self._bordercolor, padx=self._padx,
                                     pady=self._pady, background=self._cell_background, foreground=self._cell_foreground, font=self._cell_font, anchor=self._cell_anchor)

                cell.grid(row=i, column=j, sticky=N+E+W+S)

            self._data_vars.append(list_of_vars)

        if number_of_rows == 0:
            for j in range(self.number_of_columns):
                header_cell = self._head.grid_slaves(row=0, column=j)[0]
                data_cell = self._body.grid_slaves(row=0, column=j)[0]
                data_cell.bind("<Configure>", lambda event, header_cell=header_cell: header_cell.configure(
                    width=event.width), add="+")

        self._number_of_rows += n

    def _pop_n_rows(self, n):
        number_of_rows = self._number_of_rows
        number_of_columns = self._number_of_columns

        for i in range(number_of_rows-n, number_of_rows):
            for j in range(number_of_columns):
                self._body.grid_slaves(row=i, column=j)[0].destroy()

            self._data_vars.pop()

        self._number_of_rows -= n

    def set_data(self, data):
        n = len(data)
        m = len(data[0])

        number_of_rows = self._number_of_rows

        if number_of_rows > n:
            self._pop_n_rows(number_of_rows-n)
        elif number_of_rows < n:
            self._append_n_rows(n-number_of_rows)

        for i in range(n):
            for j in range(m):
                self._data_vars[i][j].set(data[i][j])

        if self._on_change_data is not None: self._on_change_data()

    def get_data(self):
        number_of_rows = self._number_of_rows
        number_of_columns = self.number_of_columns

        data = []
        for i in range(number_of_rows):
            row = []
            row_of_vars = self._data_vars[i]
            for j in range(number_of_columns):
                cell_data = row_of_vars[j].get()
                row.append(cell_data)

            data.append(row)
        return data

    @property
    def number_of_rows(self):
        return self._number_of_rows

    @property
    def number_of_columns(self):
        return self._number_of_columns

    def row(self, index, data=None):
        if data is None:
            row = []
            row_of_vars = self._data_vars[index]

            for j in range(self.number_of_columns):
                row.append(row_of_vars[j].get())

            return row
        else:
            number_of_columns = self.number_of_columns

            if len(data) != number_of_columns:
                raise ValueError("data has no %d elements: %s" %
                                 (number_of_columns, data))

            row_of_vars = self._data_vars[index]
            for j in range(number_of_columns):
                row_of_vars[index][j].set(data[j])

            if self._on_change_data is not None: self._on_change_data()

    def column(self, index, data=None):
        number_of_rows = self._number_of_rows

        if data is None:
            column = []

            for i in range(number_of_rows):
                column.append(self._data_vars[i][index].get())

            return column
        else:
            if len(data) != number_of_rows:
                raise ValueError("data has no %d elements: %s" %
                                 (number_of_rows, data))

            for i in range(number_of_columns):
                self._data_vars[i][index].set(data[i])

            if self._on_change_data is not None: self._on_change_data()

    def clear(self):
        number_of_rows = self._number_of_rows
        number_of_columns = self._number_of_columns

        for i in range(number_of_rows):
            for j in range(number_of_columns):
                self._data_vars[i][j].set("")

        if self._on_change_data is not None: self._on_change_data()

    def delete_row(self, index):
        i = index
        while i < self._number_of_rows:
            row_of_vars_1 = self._data_vars[i]
            row_of_vars_2 = self._data_vars[i+1]

            j = 0
            while j < self.number_of_columns:
                row_of_vars_1[j].set(row_of_vars_2[j])

            i += 1

        self._pop_n_rows(1)

        if self._on_change_data is not None: self._on_change_data()

    def insert_row(self, data, index=END):
        self._append_n_rows(1)

        if index == END:
            index = self._number_of_rows - 1

        i = self._number_of_rows-1
        while i > index:
            row_of_vars_1 = self._data_vars[i-1]
            row_of_vars_2 = self._data_vars[i]

            j = 0
            while j < self.number_of_columns:
                row_of_vars_2[j].set(row_of_vars_1[j])
                j += 1
            i -= 1

        list_of_cell_vars = self._data_vars[index]
        for cell_var, cell_data in zip(list_of_cell_vars, data):
            cell_var.set(cell_data)

        if self._on_change_data is not None: self._on_change_data()

    def cell(self, row, column, data=None):
        """Get the value of a table cell"""
        if data is None:
            return self._data_vars[row][column].get()
        else:
            self._data_vars[row][column].set(data)
            if self._on_change_data is not None: self._on_change_data()

    def __getitem__(self, index):
        if isinstance(index, tuple):
            row, column = index
            return self.cell(row, column)
        else:
            raise Exception("Row and column indices are required")

    def __setitem__(self, index, value):
        if isinstance(index, tuple):
            row, column = index
            self.cell(row, column, value)
        else:
            raise Exception("Row and column indices are required")

    def on_change_data(self, callback):
        self._on_change_data = callback


# class Window(Frame):
#     def __init__(self, master=None):
#         Frame.__init__(self, master)
#         self.master = master

#         menu = Menu(self.master)
#         self.master.config(menu=menu)

#         fileMenu = Menu(menu)
#         fileMenu.add_command(label="Item")
#         fileMenu.add_command(label="Exit", command=self.exitProgram)
#         menu.add_cascade(label="File", menu=fileMenu)

#         editMenu = Menu(menu)
#         editMenu.add_command(label="Undo")
#         editMenu.add_command(label="Redo")
#         menu.add_cascade(label="Edit", menu=editMenu)

#     def exitProgram(self):
#         exit()


def getListOfFiles(dirName):
    try:
        # create a list of file and sub directories
        # names in the given directory
        listOfFile = os.listdir(dirName)
        allFiles = list()
        # Iterate over all the entries
        for entry in listOfFile:
            # Create full path
            fullPath = os.path.join(dirName, entry).replace('\\', '/')
            # If entry is a directory then get the list of files in this directory
            if os.path.isdir(fullPath):
                allFiles = allFiles + getListOfFiles(fullPath)
            else:
                allFiles.append(fullPath)

        return allFiles
    except (IOError, ValueError):
        message_box.showwarning("Warning","指定されたパスが見つかりません")
    except:
        message_box.showwarning("Warning","An unexpected error occurred")
        raise


def runUploadCommand(importPath):
    try:
      
        print(importPath)
        # output = subprocess.check_output('move "%s" "%s"' % (importPath.encode('utf-8'),destPath), stderr=subprocess.STDOUT,shell=True)
        output = subprocess.check_output('gsutil cp -r "%s" "%s"' % (importPath.encode('utf-8'), destPath), stderr=subprocess.STDOUT,
                shell=True)
        return output
    except AssertionError as error:
        # message_box.showwarning("Warning",'"%s"Upload出来ませんでした"'  % (importPath.encode('utf-8')))
        # print(error)
        return "NG"

def writeCSV(data):
    filename = 'GCP_Upload_' + str(dt.datetime.now().strftime('%Y%m%d%H%M')) +'.csv'
    with open(filename, 'w+') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)
    csvfile.close()

def getFolderPath():
    progressbarframe.pack_forget()
    dirName = filedialogprompt.askdirectory()
    # if folder not seleted
    if not dirName:
        return
    # showing selected folder
    folderPathlable.config(text=dirName)

def upload():
    try:
    # os.system('cmd /k "color a & date"')
      dirName = folderPathlable.cget("text")
     
      proceedNext = message_box.askquestion ('Confirmation','Uploadの作業を開始しますか？',icon = 'warning')
      
      if not proceedNext == 'yes':
          return  
        # root.destroy()
    
    # if folder not seleted
      if not dirName:
          message_box.showwarning("Warning","フォルダを選択してください。")
          return
      
    # showing progress bar
      progressbarframe.pack()

    # showing selected folder
      folderPathlable.config(text= dirName)
      
    # Get the list of all files in directory tree at given path
      listOfFiles = getListOfFiles(dirName)
      
    # Create shared variable and set initial value.
      MAX = len(listOfFiles)
      k = 1
      data = []
      fileName = ""
      result =""
      
      progressbar["maximum"] = MAX
       # Print the files
      for elem in listOfFiles:
            # progress_var.set(k)

            # fileName = moji.zen2han(os.path.basename(elem))
            # fileName = os.path.basename(elem)
            # folderNamePath = os.path.dirname(elem)

            # fullPath = folderNamePath + '/' + fileName
            # print(fullPath)
            # run gsutil command
            result = runUploadCommand(elem)
            
            # add to userform table
            # table.insert_row([k+1,elem,result])
            data.append([k,elem.encode('utf-8'),result])
            progressbar["value"] = k
            progressbar.update()
            k +=1
            root.update_idletasks()
            time.sleep(0.02)

      print ("****************")
      
      # write result on csv
      writeCSV(data)

      table.insert_row([dt.datetime.now(),"%s ファイルのアプロード作業を完了" %(str(MAX))])

      message_box.showinfo("報告","%s ファイルのアプロード作業を完了です。 確認してください。" %(str(MAX)))
      
    except:
        message_box.showwarning("Warning","Upload出来ませんでした")
        progressbarframe.pack_forget()
        raise

def center(win):
    # """
    # centers a tkinter window
    # :param win: the root or Toplevel window to center
    # """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()      


# def loginCMD():
#       output = subprocess.check_output('gcloud auth login', stderr=subprocess.STDOUT,shell=True)
#       print( output)

# def login():
#       print(loginCMD)  
#       loginframe.pack_forget()

if __name__ == "__main__":
    try:
        from Tkinter import Tk
        import ttk
    except ImportError:
        from tkinter import Tk

    root = Tk()

    # for window menu
    # app = Window(root)
    
    # Create shared variable and set initial value.
    MAX = 0
    
    frame = Frame(root)
    frame.pack()

    # loginframe = Frame(root,bg="skyblue",width=500,height=800)
    # loginframe.pack(side=TOP, fill="both", expand=True)

    topframe = Frame(root,bg="grey",width=500)
    topframe.pack( side = TOP )

    folderpathframe = Frame(topframe,bg="skyblue",width=500)
    folderpathframe.pack( side = BOTTOM )

    uploadFrame = Frame(folderpathframe,bg="skyblue",width=500,height=200)
    uploadFrame.pack( side = BOTTOM )

    # popUpFrame = Tk.Toplevel()

    
    resultFrame = Frame(root,bg="skyblue",width=500)
    resultFrame.pack( side = TOP )

    pframe = Frame(uploadFrame,bg="skyblue",width=500)
    pframe.pack( side = BOTTOM )

    progressbarframe = Frame(pframe,width=500)
    progressbarframe.pack( side = TOP )

    
    # loginTop = Label(loginframe,width=200,height=100)
    # loginTop.pack()

    # loginButton = Button(loginframe,text="  Login  ", command=login)
    # loginButton.pack()

    # loginLabel = Label(loginframe,width=500,height=50)
    # loginLabel.pack(side="left")

    headerTitle = Label(topframe,  text="Upload to GCP              ",fg="blue",font=("Meiryo UI", 48,'bold'),width=400)
    headerTitle.pack(pady=4,side="top")

    headerText = Label(topframe, text="説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　\n ※STEP1: GCPにアプロードするフォルダ先を選択ください            　                                                                         \n※STEP2:確認上アプロードボタンを押してください。  　                                                                                         \n",font=("Meiryo UI", 11))
    headerText.pack(pady=4,side="top")

    selectFolderButton = Button(folderpathframe,text='      フォルダ選択    ', command=getFolderPath,
                    bg='blue', fg='white', font=('Meiryo UI', 12, 'bold'))
    selectFolderButton.pack(padx=5, pady=5,side="left")

    folderPathlable = Label(folderpathframe,text="Select the folder to upload", fg='red', font=("Meiryo UI", 11))
    folderPathlable.pack(padx=5, pady=5,side="left")
    

    # uploadText = Label(uploadFrame,text="STEP2 :アプロード開始                                                                                                                                  ",bg="skyblue", font=("Meiryo UI", 11))
    # uploadText.pack(padx=5, pady=5,side="top")
    uploadText = Label(uploadFrame,text="",bg="skyblue", width = 400, font=("Meiryo UI", 11))
    uploadText.pack(padx=5, pady=5,side="top")
    # uploadText2 = Label(uploadFrame,text="アプロード開始",bg="skyblue",width=500, font=("Meiryo UI", 11))
    # uploadText2.pack(side="top")

    uploadButton = Button(uploadFrame,text='       アプロード      ', command=upload,
                    bg='blue', fg='white', font=('Meiryo UI', 12, 'bold'))
    uploadButton.pack(padx=5, pady=5,side="left")

    uploadPath = Label(uploadFrame,text=destPath,
                    bg='blue', fg='white', font=('Meiryo UI', 11, 'bold'))
    uploadPath.pack(padx=5, pady=5,side="left")

    progressBarLabel = Label(progressbarframe, text="Progress Bar",bg="skyblue")
    progressBarLabel.pack(padx=5, pady=5,side="left")

    # progressbar 
    progressbar = ttk.Progressbar(progressbarframe, orient = 'horizontal', length = 400, mode = 'determinate')
    progressbar.pack(padx=5, pady=5,fill=X)

    progressbarframe.pack_forget()

    # resultLabel = Label(resultFrame, text="Result",bg="skyblue")
    # resultLabel.pack(padx=5, pady=5,side="top")

    # table = Table(root, ["SN", "File Path", "File Name", "Result"], column_minwidths=[10, 400, None, None])
    table = Table(resultFrame, ["Date", "Result"], column_minwidths=[10, 400])
    table.pack(expand=False, fill=X,padx=10,pady=10)
    
    root.update()

    root.title("Upload Files to Google Cloud")
    # Set the starting size of the window
    # root.geometry("%sx%s"%(root.winfo_reqwidth(),500))
    root.geometry("800x500")
    # root.geometry("500x500") #You want the size of the app to be 500x500
    root.resizable(False, False) #Don't allow resizing in the x or y direction

     # centers a tkinter window
    center(root)

    root.config(bg="skyblue")

    moji = MOJI()
 
    root.mainloop()
