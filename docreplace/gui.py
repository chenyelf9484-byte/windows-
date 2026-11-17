import importlib.util
import copy
import datetime
import io
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import zipfile
from dataclasses import asdict
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
from . import __version__
from .engine import (Control, enumerate_files, scan_file, save_file, SUPPORTED, IMAGES)
from .rules import Rule, Matcher


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('文档替换工具 · '+__version__)
        self.geometry('1180x820')
        self.minsize(960, 680)
        self.configure(bg='#eef2f6')
        self.results, self.rows, self.reports = [], {}, []
        self.events = queue.Queue()
        self.busy = False
        self.control = Control()
        self.ocr_available = importlib.util.find_spec('rapidocr_onnxruntime') is not None
        self.ocr = tk.BooleanVar(value=False)
        self.embedded = tk.BooleanVar(value=True)
        self.standalone = tk.BooleanVar(value=False)
        self.recursive = tk.BooleanVar(value=True)
        self.masters = tk.BooleanVar(value=False)
        self.ignore_case = tk.BooleanVar(value=False)
        self.mode = tk.StringVar(value='copy')
        self.output = tk.StringVar()
        self.encoding = tk.StringVar(value='auto')
        self.extensions = tk.StringVar(value='.docx .xlsx .pptx .txt .md .csv .tsv .log .doc .xls .ppt .wps .et .dps')
        self.max_mb = tk.StringVar(value='200')
        self.min_mb = tk.StringVar(value='0')
        self.date_after = tk.StringVar()
        self.font = tk.StringVar(value=self.default_font())
        self.status = tk.StringVar(value='选择文件夹或文件，添加替换规则，然后扫描。')
        self.worker_buttons = []
        self.style_ui()
        self.build()
        self.after(100, self.poll)
        self.protocol('WM_DELETE_WINDOW', self.close)

    def default_font(self):
        for p in ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf',
                  '/System/Library/Fonts/PingFang.ttc', '/System/Library/Fonts/STHeiti Light.ttc']:
            if Path(p).exists():
                return p
        return ''

    def style_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        font = 'Microsoft YaHei UI' if os.name == 'nt' else 'PingFang SC'
        self.option_add('*Font', (font, 10))
        style.configure('.', font=(font, 10), background='#eef2f6')
        style.configure('TButton', padding=(12, 7))
        style.configure('Accent.TButton', background='#225a91', foreground='white', padding=(18, 9))
        style.map('Accent.TButton', background=[('active', '#194773'), ('disabled', '#a2b2c4')])
        style.configure('Treeview', background='white', fieldbackground='white', rowheight=30)
        style.configure('Treeview.Heading', padding=7, font=(font, 10, 'bold'))
        style.configure('TNotebook.Tab', padding=(22, 10))
        style.configure('Title.TLabel', font=(font, 22, 'bold'), foreground='#183b5d')

    def button(self, parent, text, command, **kwargs):
        b = ttk.Button(parent, text=text, command=command, **kwargs)
        b.pack(side='left', padx=(0, 7), pady=4)
        self.worker_buttons.append(b)
        return b

    def build(self):
        head = ttk.Frame(self, padding=(24, 18, 24, 12))
        head.pack(fill='x')
        ttk.Label(head, text='文档替换工具', style='Title.TLabel').pack(side='left')
        ttk.Label(head, text='本机处理  /  扫描后预览  /  按选择保存', foreground='#536a80').pack(side='right')
        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill='both', expand=True, padx=20)
        self.paths_tab = ttk.Frame(self.tabs, padding=15)
        self.rules_tab = ttk.Frame(self.tabs, padding=15)
        self.results_tab = ttk.Frame(self.tabs, padding=15)
        for tab, label in [(self.paths_tab, '扫描范围'), (self.rules_tab, '替换规则'), (self.results_tab, '结果与保存')]:
            self.tabs.add(tab, text=label)
        self.build_paths()
        self.build_rules()
        self.build_results()
        footer = ttk.Frame(self, padding=(20, 10))
        footer.pack(fill='x')
        self.scan_button = self.button(footer, '开始扫描', self.start_scan, style='Accent.TButton')
        self.pause_button = ttk.Button(footer, text='暂停', command=self.pause, state='disabled')
        self.pause_button.pack(side='left', padx=5)
        self.cancel_button = ttk.Button(footer, text='取消', command=self.cancel, state='disabled')
        self.cancel_button.pack(side='left', padx=5)
        self.progress = ttk.Progressbar(footer, mode='indeterminate', length=120)
        self.progress.pack(side='right')
        ttk.Label(self, textvariable=self.status, wraplength=1100, padding=(20, 0, 20, 12)).pack(fill='x')

    def build_paths(self):
        p = self.paths_tab
        ttk.Label(p, text='扫描路径（可直接填写 C:\\ 扫描磁盘，也可添加多个文件夹或文件）').pack(anchor='w')
        self.path_list = tk.Listbox(p, height=6, selectmode='extended', bg='white', relief='solid', bd=1)
        self.path_list.pack(fill='x', pady=8)
        bar = ttk.Frame(p)
        bar.pack(fill='x')
        self.button(bar, '添加文件夹', lambda: self.add_path(filedialog.askdirectory()))
        self.button(bar, '添加文件', self.add_files)
        self.button(bar, '输入路径', self.input_path)
        self.button(bar, '移除选中', self.remove_path)
        ttk.Checkbutton(bar, text='包含子文件夹', variable=self.recursive).pack(side='right')
        form = ttk.Frame(p)
        form.pack(fill='x', pady=10)
        for row, (label, var) in enumerate([('扫描扩展名（空格分隔）', self.extensions),
                                           ('单文件最大 MB', self.max_mb), ('单文件最小 MB', self.min_mb),
                                           ('修改时间不早于（YYYY-MM-DD，可空）', self.date_after)]):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky='w', padx=(0, 15), pady=4)
            ttk.Entry(form, textvariable=var).grid(row=row, column=1, sticky='ew')
        form.columnconfigure(1, weight=1)
        ttk.Label(p, text='排除目录（每行一个完整路径；系统目录和工具备份目录自动排除）').pack(anchor='w')
        self.exclusions = tk.Text(p, height=3, bg='white', relief='solid', bd=1)
        self.exclusions.pack(fill='x', pady=6)
        line = ttk.Frame(p)
        line.pack(fill='x', pady=5)
        ttk.Checkbutton(line, text='启用图片识别（默认关闭）', variable=self.ocr, command=self.ocr_toggle, state='normal' if self.ocr_available else 'disabled').pack(side='left')
        self.embedded_check = ttk.Checkbutton(line, text='文档内图片', variable=self.embedded, state='disabled')
        self.embedded_check.pack(side='left', padx=15)
        self.standalone_check = ttk.Checkbutton(line, text='独立图片', variable=self.standalone, state='disabled')
        self.standalone_check.pack(side='left')
        line2 = ttk.Frame(p)
        line2.pack(fill='x', pady=5)
        ttk.Label(line2, text='文本编码').pack(side='left')
        ttk.Combobox(line2, textvariable=self.encoding, values=['auto', 'utf-8', 'gb18030', 'utf-16-le', 'utf-16-be', 'big5'],
                     width=15, state='readonly').pack(side='left', padx=10)
        ttk.Checkbutton(line2, text='扫描 PPT 母版与版式', variable=self.masters).pack(side='left', padx=15)

    def build_rules(self):
        p = self.rules_tab
        ttk.Label(p, text='按原文一次匹配，较长关键词优先；替换内容可以留空。').pack(anchor='w', pady=(0, 10))
        line = ttk.Frame(p)
        line.pack(fill='x')
        self.source, self.target = tk.StringVar(), tk.StringVar()
        ttk.Label(line, text='查找').pack(side='left')
        ttk.Entry(line, textvariable=self.source, width=30).pack(side='left', padx=8)
        ttk.Label(line, text='替换为').pack(side='left')
        ttk.Entry(line, textvariable=self.target, width=30).pack(side='left', padx=8)
        self.button(line, '添加规则', self.add_rule)
        self.rules_tree = ttk.Treeview(p, columns=('source', 'target'), show='headings', selectmode='extended')
        for col, label in [('source', '原关键词'), ('target', '替换文字')]:
            self.rules_tree.heading(col, text=label)
            self.rules_tree.column(col, width=350)
        self.rules_tree.pack(fill='both', expand=True, pady=12)
        self.rules_tree.bind('<Double-1>', self.edit_rule)
        bar = ttk.Frame(p)
        bar.pack(fill='x')
        self.button(bar, '删除选中', self.delete_rules)
        self.button(bar, '读取规则', self.load_rules)
        self.button(bar, '保存规则', self.save_rules)
        ttk.Checkbutton(bar, text='忽略英文大小写', variable=self.ignore_case).pack(side='right')
        ttk.Label(p, text='双击规则可编辑。规则在开始扫描时固定；修改规则后请重新扫描。', foreground='#536a80').pack(anchor='w', pady=10)

    def build_results(self):
        p = self.results_tab
        bar = ttk.Frame(p)
        bar.pack(fill='x')
        self.button(bar, '全选文字', lambda: self.select_units(True))
        self.button(bar, '全部取消', lambda: self.select_units(False))
        self.button(bar, '切换勾选', self.toggle_selected)
        self.button(bar, '打开原文件', self.open_source)
        self.button(bar, '图片预览与编辑', self.image_editor)
        table = ttk.Frame(p)
        table.pack(fill='both', expand=True, pady=8)
        self.tree = ttk.Treeview(table, columns=('check', 'file', 'location', 'before', 'after'), show='headings', selectmode='extended')
        for col, label, width in [('check', '修改', 55), ('file', '文件 / 状态', 210), ('location', '位置', 190),
                                  ('before', '原内容', 220), ('after', '替换为 / 说明', 240)]:
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, minwidth=45, stretch=col != 'check')
        scroll = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        horizontal = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        scroll.grid(row=0, column=1, sticky='ns')
        horizontal.grid(row=1, column=0, sticky='ew')
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree.bind('<Double-1>', self.toggle_selected)
        self.tree.bind('<<TreeviewSelect>>', self.show_detail)
        self.detail = tk.Text(p, height=5, bg='#f7f9fc', relief='flat', wrap='word', state='disabled')
        self.detail.pack(fill='x', pady=5)
        modes = ttk.Frame(p)
        modes.pack(fill='x', pady=8)
        for text, value in [('另存副本', 'copy'), ('备份后修改原文件', 'backup'), ('直接修改原文件并保存（不备份）', 'direct')]:
            ttk.Radiobutton(modes, text=text, value=value, variable=self.mode).pack(side='left', padx=(0, 18))
        out = ttk.Frame(p)
        out.pack(fill='x')
        ttk.Label(out, text='副本输出目录').pack(side='left')
        ttk.Entry(out, textvariable=self.output).pack(side='left', fill='x', expand=True, padx=8)
        self.button(out, '选择目录', self.pick_output)
        action = ttk.Frame(p)
        action.pack(fill='x', pady=8)
        self.save_button = self.button(action, '执行勾选项目', self.start_save, style='Accent.TButton')
        self.button(action, '导出任务报告', self.export_report)
        ttk.Label(action, text='图片默认仅标记，需预览确认后勾选。', foreground='#536a80').pack(side='right')

    def add_path(self, path):
        if path and path not in self.path_list.get(0, 'end'):
            self.path_list.insert('end', path)

    def add_files(self):
        for path in filedialog.askopenfilenames():
            self.add_path(path)

    def input_path(self):
        self.add_path(simpledialog.askstring('输入路径', '磁盘、文件夹或文件完整路径：', parent=self))

    def remove_path(self):
        for i in reversed(self.path_list.curselection()):
            self.path_list.delete(i)

    def pick_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output.set(path)

    def ocr_toggle(self):
        for widget in (self.embedded_check, self.standalone_check):
            widget.configure(state='normal' if self.ocr.get() else 'disabled')

    def add_rule(self):
        if not self.source.get():
            messagebox.showerror('规则', '关键词不能为空')
            return
        self.rules_tree.insert('', 'end', values=(self.source.get(), self.target.get()))
        self.source.set('')
        self.target.set('')

    def edit_rule(self, event=None):
        if self.busy or not self.rules_tree.selection():
            return
        item = self.rules_tree.selection()[0]
        source, target = self.rules_tree.item(item, 'values')
        new_source = simpledialog.askstring('编辑关键词', '关键词', initialvalue=source, parent=self)
        if not new_source:
            return
        new_target = simpledialog.askstring('编辑替换文字', '替换为（可以留空）', initialvalue=target, parent=self)
        if new_target is not None:
            self.rules_tree.item(item, values=(new_source, new_target))

    def delete_rules(self):
        for i in self.rules_tree.selection():
            self.rules_tree.delete(i)

    def matcher(self):
        return Matcher([Rule(*self.rules_tree.item(i, 'values')) for i in self.rules_tree.get_children()], self.ignore_case.get())

    def load_rules(self):
        path = filedialog.askopenfilename(filetypes=[('规则 JSON', '*.json')])
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text('utf-8-sig'))
            matcher = Matcher([Rule(r['source'], r['target']) for r in data['rules']], data.get('ignore_case', False))
            self.rules_tree.delete(*self.rules_tree.get_children())
            for r in matcher.rules:
                self.rules_tree.insert('', 'end', values=(r.source, r.target))
            self.ignore_case.set(matcher.ignore_case)
        except Exception as exc:
            messagebox.showerror('读取失败', str(exc))

    def save_rules(self):
        try:
            data = self.matcher().to_json()
            path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('规则 JSON', '*.json')])
            if path:
                Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), 'utf-8')
        except Exception as exc:
            messagebox.showerror('保存失败', str(exc))

    def set_busy(self, busy):
        self.busy = busy
        for b in self.worker_buttons:
            b.configure(state='disabled' if busy else 'normal')
        self.pause_button.configure(state='normal' if busy else 'disabled', text='暂停')
        self.cancel_button.configure(state='normal' if busy else 'disabled')
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def start_scan(self):
        try:
            matcher = self.matcher()
            roots = list(self.path_list.get(0, 'end'))
            if not roots:
                raise ValueError('请先选择扫描路径')
            options = {'embedded': self.embedded.get(), 'standalone': self.standalone.get(),
                       'masters': self.masters.get(), 'encoding': self.encoding.get(), 'font': self.font.get()}
            extensions = {('.'+e.lstrip('.')).lower() for e in self.extensions.get().replace(',', ' ').split()}
            if self.ocr.get() and self.standalone.get():
                extensions |= IMAGES
            if not extensions or extensions - SUPPORTED:
                raise ValueError('扩展名列表为空或包含不支持的格式：'+str(extensions-SUPPORTED))
            if not self.ocr.get():
                extensions -= IMAGES
            if not extensions:
                raise ValueError('关闭图片识别后没有可扫描格式，请选择文档格式或启用 OCR')
            scan_args = {'recursive': self.recursive.get(), 'extensions': extensions,
                         'max_mb': float(self.max_mb.get()), 'min_mb': float(self.min_mb.get()),
                         'modified_after': datetime.datetime.strptime(self.date_after.get(), '%Y-%m-%d').timestamp() if self.date_after.get() else 0,
                         'excluded': self.exclusions.get('1.0', 'end').strip().splitlines()}
            if not 0 <= scan_args['min_mb'] <= scan_args['max_mb']:
                raise ValueError('文件大小范围无效')
            if self.output.get():
                scan_args['excluded'].append(self.output.get())
            use_ocr = self.ocr.get()
        except Exception as exc:
            messagebox.showerror('扫描设置', str(exc))
            return
        self.results, self.rows, self.reports = [], {}, []
        self.scan_roots = roots
        self.tree.delete(*self.tree.get_children())
        self.control = Control()
        self.set_busy(True)
        self.tabs.select(self.results_tab)
        def worker():
            try:
                ocr = None
                if use_ocr:
                    self.events.put(('status', '正在加载离线图片识别组件…'))
                    from .ocr import OCR
                    ocr = OCR()
                count = 0
                scan_args['on_error'] = lambda p, e: self.events.put(('path_error', (p, e)))
                for path in enumerate_files(roots, control=self.control, **scan_args):
                    count += 1
                    self.events.put(('status', '已扫描 %d 个文件 · %s' % (count, path)))
                    result = scan_file(path, matcher, options, ocr)
                    self.events.put(('result', result))
                self.events.put(('status', '扫描%s：%d 个文件。请预览并选择保存方式。' % ('已取消' if self.control.cancel.is_set() else '完成', count)))
            except Exception as exc:
                self.events.put(('error', str(exc)))
            finally:
                self.events.put(('done', None))
        threading.Thread(target=worker, daemon=True).start()

    def start_save(self):
        selected = [r for r in self.results if any(u.selected for u in r.units)]
        if not selected:
            messagebox.showinfo('保存', '请先扫描并勾选修改项目')
            return
        mode, output = self.mode.get(), self.output.get()
        if mode == 'copy' and not output:
            messagebox.showerror('保存', '另存副本需要选择输出目录')
            return
        count = sum(sum(u.selected for u in r.units) for r in selected)
        msg = '将修改 %d 个文件中的 %d 项。\n保存方式：%s' % (len(selected), count,
                {'copy': '另存副本', 'backup': '备份后修改原文件', 'direct': '直接覆盖原文件，不生成备份'}[mode])
        if not messagebox.askyesno('执行替换', msg+'\n是否执行？'):
            return
        # 工作线程只读取快照，用户切换设置不改变本批任务。
        selected = copy.deepcopy(selected)
        roots = list(getattr(self, 'scan_roots', []))
        self.control = Control()
        self.set_busy(True)
        def worker():
            try:
                for i, result in enumerate(selected):
                    if not self.control.checkpoint():
                        break
                    bases = [Path(r).resolve() for r in roots if Path(r).is_dir() and (Path(r).resolve() in Path(result.path).resolve().parents)]
                    root = max(bases, key=lambda p: len(p.parts)) if bases else Path(result.path).parent
                    self.events.put(('status', '正在保存 %d/%d · %s' % (i+1, len(selected), result.path)))
                    self.events.put(('saved', save_file(result, mode, output, root)))
                self.events.put(('status', '保存%s。成功和失败详情见列表；修改原文件后再次执行需要重新扫描。' % ('已取消' if self.control.cancel.is_set() else '完成')))
            except Exception as exc:
                self.events.put(('error', str(exc)))
            finally:
                self.events.put(('done', None))
        threading.Thread(target=worker, daemon=True).start()

    def pause(self):
        if self.control.running.is_set():
            self.control.running.clear()
            self.pause_button.configure(text='继续')
            self.status.set('将在当前文件处理完后暂停。')
        else:
            self.control.running.set()
            self.pause_button.configure(text='暂停')

    def cancel(self):
        self.control.cancel.set()
        self.control.running.set()
        self.status.set('正在取消，当前文件将处理完成后停止。')

    def poll(self):
        for _ in range(150):
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == 'status':
                self.status.set(value)
            elif kind == 'result':
                self.results.append(value)
                for unit in value.units:
                    row = self.tree.insert('', 'end', values=('☑' if unit.selected else '☐', Path(value.path).name, unit.location, unit.before[:150], unit.after[:150]))
                    self.rows[row] = (value, unit)
                for warning in value.warnings:
                    self.tree.insert('', 'end', values=('', value.path, '注意', '', warning))
                if value.error or not value.units:
                    self.tree.insert('', 'end', values=('', value.path, '失败' if value.error else '未命中', '', value.error))
            elif kind == 'saved':
                self.reports.append(asdict(value))
                self.tree.insert('', 'end', values=('', value.source, value.status, value.backup, value.detail))
            elif kind == 'path_error':
                self.reports.append({'source': value[0], 'status': '扫描失败', 'detail': value[1]})
                self.tree.insert('', 'end', values=('', value[0], '扫描失败', '', value[1]))
            elif kind == 'error':
                self.status.set('任务失败：'+value)
                messagebox.showerror('任务失败', value)
            elif kind == 'done':
                self.set_busy(False)
        self.after(100, self.poll)

    def refresh_row(self, row):
        result, unit = self.rows[row]
        self.tree.item(row, values=('☑' if unit.selected else '☐', Path(result.path).name, unit.location, unit.before[:150], unit.after[:150]))

    def toggle_selected(self, event=None):
        if self.busy:
            return
        for row in self.tree.selection():
            if row in self.rows:
                unit = self.rows[row][1]
                if unit.kind == 'image' and not unit.selected and not unit.payload.get('confirmed'):
                    messagebox.showinfo('图片修改', '请先打开“图片预览与编辑”，确认替换区域和效果。')
                    continue
                unit.selected = not unit.selected
                self.refresh_row(row)

    def select_units(self, selected):
        for row, (_, unit) in self.rows.items():
            if not selected or unit.kind == 'text':
                unit.selected = selected
                self.refresh_row(row)

    def show_detail(self, event=None):
        rows = self.tree.selection()
        text = ''
        if rows and rows[0] in self.rows:
            result, unit = self.rows[rows[0]]
            text = '%s\n%s\n原文：%s\n替换：%s' % (result.path, unit.location, unit.before, unit.after)
        elif rows:
            text = '\n'.join(str(v) for v in self.tree.item(rows[0], 'values'))
        self.detail.configure(state='normal')
        self.detail.delete('1.0', 'end')
        self.detail.insert('1.0', text)
        self.detail.configure(state='disabled')

    def open_source(self):
        rows = self.tree.selection()
        if rows and rows[0] in self.rows:
            path = self.rows[rows[0]][0].path
            if os.name == 'nt':
                os.startfile(path)
            else:
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', path])

    def image_editor(self):
        rows = self.tree.selection()
        if not rows or rows[0] not in self.rows or self.rows[rows[0]][1].kind != 'image':
            messagebox.showinfo('图片', '请先选择一条图片命中记录')
            return
        row = rows[0]
        result, unit = self.rows[row]
        try:
            from PIL import Image, ImageTk, ImageDraw
            from .ocr import render, load_image
            part = unit.payload['part']
            if Path(result.path).suffix.lower() in IMAGES:
                data = Path(result.path).read_bytes()
            elif '_modern_data' in result.options:
                with zipfile.ZipFile(io.BytesIO(result.options['_modern_data'])) as z:
                    data = z.read(part)
            else:
                with zipfile.ZipFile(result.path) as z:
                    data = z.read(part)
            work = copy.deepcopy(unit)
            font = result.options.get('font') or self.default_font()
            win = tk.Toplevel(self)
            win.title('图片修改 · 整行重绘，请检查背景和邻近文字')
            win.geometry('1000x690')
            canvas = tk.Canvas(win, bg='#dce4ed', height=440)
            canvas.pack(fill='both', expand=True)
            target = tk.StringVar(value=unit.after)
            coords = tk.StringVar(value=','.join(str(int(v)) for v in (min(p[0] for p in work.payload['box']), min(p[1] for p in work.payload['box']), max(p[0] for p in work.payload['box']), max(p[1] for p in work.payload['box']))))
            font_var = tk.StringVar(value=font)
            fg, bg = tk.StringVar(value=result.options.get('foreground', '#000000')), tk.StringVar(value=result.options.get('background', '#ffffff'))
            form = ttk.Frame(win, padding=10)
            form.pack(fill='x')
            for n, (label, variable) in enumerate([('替换后的整行文字', target), ('选区 x0,y0,x1,y1（像素）', coords), ('字体文件', font_var)]):
                ttk.Label(form, text=label).grid(row=n, column=0, sticky='w', pady=3)
                ttk.Entry(form, textvariable=variable, width=80).grid(row=n, column=1, sticky='ew')
            form.columnconfigure(1, weight=1)
            actions = ttk.Frame(win, padding=10)
            actions.pack(fill='x')
            state = {'preview': False}
            def read_work():
                x0, y0, x1, y1 = [int(v.strip()) for v in coords.get().split(',')]
                if not (x1 > x0 >= 0 and y1 > y0 >= 0):
                    raise ValueError('选区坐标无效')
                work.payload['box'] = [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]
                work.after = target.get()
            def preview():
                try:
                    read_work()
                    combined = [work if u is unit else u for u in result.units
                                if u is unit or (u.kind == 'image' and u.selected and u.payload['part'] == part
                                                 and not u.payload.get('replacement') and not work.payload.get('replacement'))]
                    after = render(data, combined, font_var.get(), fg.get(), bg.get())
                    original = load_image(data)
                    ImageDraw.Draw(original).rectangle((work.payload['box'][0][0], work.payload['box'][0][1], work.payload['box'][2][0], work.payload['box'][2][1]), outline='red', width=3)
                    images = [original, Image.open(io.BytesIO(after)).convert('RGB')]
                    canvas.delete('all')
                    photos = []
                    for i, im in enumerate(images):
                        im.thumbnail((470, 405))
                        photo = ImageTk.PhotoImage(im)
                        photos.append(photo)
                        canvas.create_image(10+i*490, 28, anchor='nw', image=photo)
                        canvas.create_text(15+i*490, 12, text='原图 / 红框为选区' if i == 0 else '修改后', anchor='w')
                    canvas.photos = photos
                    state['preview'] = True
                except Exception as exc:
                    messagebox.showerror('图片预览失败', str(exc), parent=win)
            def changed(*_):
                state['preview'] = False
            for var in (target, coords, font_var, fg, bg):
                var.trace_add('write', changed)
            def choose_color(variable):
                color = colorchooser.askcolor(variable.get(), parent=win)[1]
                if color:
                    variable.set(color)
            def choose_font():
                p = filedialog.askopenfilename(parent=win, filetypes=[('字体', '*.ttf *.ttc *.otf')])
                if p:
                    font_var.set(p)
            def choose_image():
                p = filedialog.askopenfilename(parent=win, filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp')])
                if p:
                    work.payload['replacement'] = Path(p).read_bytes()
                    state['preview'] = False
                    preview()
            def confirm():
                if not state['preview']:
                    messagebox.showinfo('确认', '请先点击预览，检查当前设置的效果。', parent=win)
                    return
                unit.after = work.after
                unit.payload = copy.deepcopy(work.payload)
                unit.payload['confirmed'] = True
                unit.selected = True
                for rid, (res, other) in self.rows.items():
                    if (res is result and other is not unit and other.kind == 'image'
                            and other.payload['part'] == part
                            and (unit.payload.get('replacement') or other.payload.get('replacement'))):
                        other.selected = False
                        other.payload.pop('confirmed', None)
                        self.refresh_row(rid)
                # 同文档字体和颜色统一，已确认其他图片需重新预览。
                new_style = (font_var.get(), fg.get(), bg.get())
                old_style = (result.options.get('font'), result.options.get('foreground', '#000000'), result.options.get('background', '#ffffff'))
                if new_style != old_style:
                    for rid, (res, other) in self.rows.items():
                        if res is result and other.kind == 'image' and other is not unit:
                            other.selected = False
                            other.payload.pop('confirmed', None)
                            self.refresh_row(rid)
                result.options.update(font=font_var.get(), foreground=fg.get(), background=bg.get())
                self.refresh_row(row)
                win.destroy()
            for label, command in [('选字体', choose_font), ('字色', lambda: choose_color(fg)), ('背景色', lambda: choose_color(bg)),
                                   ('替换整张图片', choose_image), ('预览', preview), ('确认并勾选', confirm)]:
                ttk.Button(actions, text=label, command=command).pack(side='left', padx=5)
            win.transient(self)
            win.grab_set()
            preview()
        except Exception as exc:
            messagebox.showerror('图片编辑', str(exc))

    def export_report(self):
        path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('任务报告 JSON', '*.json')])
        if not path:
            return
        try:
            scan = [{'path': r.path, 'error': r.error, 'warnings': r.warnings,
                     'hits': [{'location': u.location, 'before': u.before, 'after': u.after,
                               'count': u.count, 'selected': u.selected, 'kind': u.kind} for u in r.units]} for r in self.results]
            Path(path).write_text(json.dumps({'version': __version__, 'scan': scan, 'save': self.reports}, ensure_ascii=False, indent=2), 'utf-8')
            self.status.set('报告已保存：'+path)
        except Exception as exc:
            messagebox.showerror('报告保存失败', str(exc))

    def close(self):
        if self.busy:
            self.cancel()
            messagebox.showinfo('正在停止', '等待当前文件处理结束后再关闭，避免保存中断。')
            return
        self.destroy()


def main():
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()
