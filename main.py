# -*- coding: UTF-8 -*-
import urllib.request
import time,json
import datetime
import os
import re
import tkinter.messagebox
import time
from tkinter import ttk
import tkinter.font as tkFont
import calendar
import tkinter as tk
from tkinter import *
import threading
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter.filedialog import askdirectory

def gettoken(client_id, client_secret):
    url = 'http://webapi.cninfo.com.cn/api-cloud-platform/oauth2/token'  # api.before.com需要根据具体访问域名修改
    #post_data = "grant_type=client_credentials&client_id=%s&client_secret=%s" % (client_id, client_secret)
    post_data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret
    }
    post_data = urllib.parse.urlencode(post_data).encode(encoding='UTF-8')
    req = urllib.request.urlopen(url, data=post_data)
    responsecontent = req.read()
    responsedict = json.loads(responsecontent)
    token = responsedict["access_token"]
    return token

def apipost(scode, sdate, edate, tokent):
    url = 'http://webapi.cninfo.com.cn/api/info/p_info3015' # apitest2.com需要根据具体访问域名修改
    post_data = {
        "scode": scode,
        "sdate": sdate,
        "edate": edate,
        "access_token": tokent,
    }
    params = urllib.parse.urlencode(post_data).encode(encoding='UTF8')
    req = urllib.request.urlopen(url, params)
    content = req.read()
    responsedict = json.loads(content)
    resultcode = responsedict["resultcode"]
    responsedict["resultmsg"], responsedict["resultcode"]
    if (responsedict["resultmsg"] == "success" and len(responsedict["records"]) >= 1):
        responsedict["records"]  # 接收到的具体数据内容
    else:
        print
        'no data'
    return(resultcode,responsedict)

def createfold(SECNAME,file_name,input_dir):
    fold_SEC = SECNAME.strip('*')
    sdir = input_dir + '\\' + fold_SEC
    isExists = os.path.exists(sdir)
    if not isExists:
        os.makedirs(sdir)
    downPath = os.path.join(sdir, file_name)
    return downPath

def splitscode(scodeall):
    scodelist = re.split(r'[.,;\s]\s*',scodeall)
    return scodelist

def download(i,records,input_dir):
    text = records[i]
    SECNAME = text["SECNAME"]
    date_long = text["F001D"]
    d = datetime.datetime.strptime(date_long, '%Y-%m-%d %H:%M:%S')
    date_short = datetime.datetime.strftime(d, '%Y-%m-%d')
    Market = text["F009V"]
    title1 = text["F002V"]
    title2 = title1.encode("utf-8")
    title = title2.replace(b'\xef\xbc\x9a', b'-')
    title = title.decode("utf-8")
    title = title.replace('*','')
    down_url = text["F003V"]
    type2 = text["F006V"]
    file_name = date_short + title + ".pdf"
    save_dir = createfold(SECNAME, file_name, input_dir)
    FileIsExists = os.path.isfile(save_dir)
    try:
        urllib.request.urlretrieve(down_url, save_dir)
        info.insert(tk.END,"已下载:"+file_name+'\n')
        info.see(tk.END)
        root.update()
    except FileIsExists as Ture:
        size = os.path.getsize(path)
        if formatSize(size) == 0:
            urllib.request.urlretrieve(down_url, save_dir)
            info.insert(tk.END, "已下载:" + file_name +'\n')
            info.see(tk.END)
            root.update()

def step1(scodeall,sdate,edate):
    general = {}
    client_id = '7aedaba3c8f246ab8810973590ca34aa'
    client_secret = '8199d312efd44df4b050bf7320c7c4d1'  # client_id,client_secret通过我的凭证获取
    token = gettoken(client_id, client_secret)
    scodelist = splitscode(scodeall)
    for each in scodelist:
        scode = each
        resultcode, responsedict = apipost(scode, sdate, edate, token)  # 以http post方法获取数据
        if resultcode == 405:
            token = gettoken(client_id, client_secret)
            resultcode, responsedict = apipost(scode, sdate, edate, token)  # post请求
        elif resultcode == 200:
            general[each] = responsedict
        else:
            info_error(resultcode)
    print(general)
    return general

def step2(general,input_dir):
    time_start = time.time()
    for u in general:
        responsedict = general[u]
        try:
            count = responsedict["total"]
            records = responsedict["records"]
            for i in range(0, count):
                th = threading.Thread(target=download(i, records, input_dir), args=(i,))
                th.setDaemon(True)  # 守护线程
                th.start()
                percent = (i + 1) / count * 100
                time_end = time.time()
                time_delta = time_end - time_start
                update_progress_bar(percent, time_delta)
            time.sleep(1)
            info_finish(u)
        except:
            info_noann(u)

def update_progress_bar(percent,time_delta):
        hour = int(time_delta / 3600)
        minute = int(time_delta / 60) - hour * 60
        second = time_delta % 60
        green_length = int(sum_length * percent / 100)
        canvas_progress_bar.coords(canvas_shape, (0, 0, green_length, 25))
        canvas_progress_bar.itemconfig(canvas_text, text='%02d:%02d:%02d' % (hour, minute, second))
        var_progress_bar_percent.set('%0.2f  %%' % percent)

datetime2 = calendar.datetime.datetime
timedelta = calendar.datetime.timedelta

class Calendar:

    def __init__(s, point=None, position=None):
        # point    提供一个基点，来确定窗口位置
        # position 窗口在点的位置 'ur'-右上, 'ul'-左上, 'll'-左下, 'lr'-右下
        # s.master = tk.Tk()
        s.master = tk.Toplevel()
        s.master.withdraw()
        fwday = calendar.SUNDAY

        year = datetime2.now().year
        month = datetime2.now().month
        locale = None
        sel_bg = '#ecffc4'
        sel_fg = '#05640e'

        s._date = datetime2(year, month, 1)
        s._selection = None  # 设置为未选中日期

        s.G_Frame = ttk.Frame(s.master)

        s._cal = s.__get_calendar(locale, fwday)

        s.__setup_styles()  # 创建自定义样式
        s.__place_widgets()  # pack/grid 小部件
        s.__config_calendar()  # 调整日历列和安装标记
        # 配置画布和正确的绑定，以选择日期。
        s.__setup_selection(sel_bg, sel_fg)

        # 存储项ID，用于稍后插入。
        s._items = [s._calendar.insert('', 'end', values='') for _ in range(6)]

        # 在当前空日历中插入日期
        s._update()

        s.G_Frame.pack(expand=1, fill='both')
        s.master.overrideredirect(1)
        s.master.update_idletasks()
        width, height = s.master.winfo_reqwidth(), s.master.winfo_reqheight()
        if point and position:
            if position == 'ur':
                x, y = point[0], point[1] - height
            elif position == 'lr':
                x, y = point[0], point[1]
            elif position == 'ul':
                x, y = point[0] - width, point[1] - height
            elif position == 'll':
                x, y = point[0] - width, point[1]
        else:
            x, y = (s.master.winfo_screenwidth() - width) / 2, (s.master.winfo_screenheight() - height) / 2
        s.master.geometry('%dx%d+%d+%d' % (width, height, x, y))  # 窗口位置居中
        s.master.after(300, s._main_judge)
        s.master.deiconify()
        s.master.focus_set()
        s.master.wait_window()  # 这里应该使用wait_window挂起窗口，如果使用mainloop,可能会导致主程序很多错误

    def __get_calendar(s, locale, fwday):
        # 实例化适当的日历类
        if locale is None:
            return calendar.TextCalendar(fwday)
        else:
            return calendar.LocaleTextCalendar(fwday, locale)

    def __setitem__(s, item, value):
        if item in ('year', 'month'):
            raise AttributeError("attribute '%s' is not writeable" % item)
        elif item == 'selectbackground':
            s._canvas['background'] = value
        elif item == 'selectforeground':
            s._canvas.itemconfigure(s._canvas.text, item=value)
        else:
            s.G_Frame.__setitem__(s, item, value)

    def __getitem__(s, item):
        if item in ('year', 'month'):
            return getattr(s._date, item)
        elif item == 'selectbackground':
            return s._canvas['background']
        elif item == 'selectforeground':
            return s._canvas.itemcget(s._canvas.text, 'fill')
        else:
            r = ttk.tclobjs_to_py({item: ttk.Frame.__getitem__(s, item)})
            return r[item]

    def __setup_styles(s):
        # 自定义TTK风格
        style = ttk.Style(s.master)
        arrow_layout = lambda dir: (
            [('Button.focus', {'children': [('Button.%sarrow' % dir, None)]})]
        )
        style.layout('L.TButton', arrow_layout('left'))
        style.layout('R.TButton', arrow_layout('right'))

    def __place_widgets(s):
        # 标头框架及其小部件
        Input_judgment_num = s.master.register(s.Input_judgment)  # 需要将函数包装一下，必要的
        hframe = ttk.Frame(s.G_Frame)
        gframe = ttk.Frame(s.G_Frame)
        bframe = ttk.Frame(s.G_Frame)
        hframe.pack(in_=s.G_Frame, side='top', pady=5, anchor='center')
        gframe.pack(in_=s.G_Frame, fill=tk.X, pady=5)
        bframe.pack(in_=s.G_Frame, side='bottom', pady=5)

        lbtn = ttk.Button(hframe, style='L.TButton', command=s._prev_month)
        lbtn.grid(in_=hframe, column=0, row=0, padx=12)
        rbtn = ttk.Button(hframe, style='R.TButton', command=s._next_month)
        rbtn.grid(in_=hframe, column=5, row=0, padx=12)

        s.CB_year = ttk.Combobox(hframe, width=5, values=[str(year) for year in
                                                          range(datetime2.now().year, datetime2.now().year - 11, -1)],
                                 validate='key', validatecommand=(Input_judgment_num, '%P'))
        s.CB_year.current(0)
        s.CB_year.grid(in_=hframe, column=1, row=0)
        s.CB_year.bind('<KeyPress>', lambda event: s._update(event, True))
        s.CB_year.bind("<<ComboboxSelected>>", s._update)
        tk.Label(hframe, text='年', justify='left').grid(in_=hframe, column=2, row=0, padx=(0, 5))

        s.CB_month = ttk.Combobox(hframe, width=3, values=['%02d' % month for month in range(1, 13)], state='readonly')
        s.CB_month.current(datetime2.now().month - 1)
        s.CB_month.grid(in_=hframe, column=3, row=0)
        s.CB_month.bind("<<ComboboxSelected>>", s._update)
        tk.Label(hframe, text='月', justify='left').grid(in_=hframe, column=4, row=0)

        # 日历部件
        s._calendar = ttk.Treeview(gframe, show='', selectmode='none', height=7)
        s._calendar.pack(expand=1, fill='both', side='bottom', padx=5)

        ttk.Button(bframe, text="确 定", width=6, command=lambda: s._exit(True)).grid(row=0, column=0, sticky='ns',
                                                                                    padx=20)
        ttk.Button(bframe, text="取 消", width=6, command=s._exit).grid(row=0, column=1, sticky='ne', padx=20)

        tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=0, relwidth=1, relheigh=2 / 200)
        tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=198 / 200, relwidth=1, relheigh=2 / 200)
        tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=0, relwidth=2 / 200, relheigh=1)
        tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=198 / 200, rely=0, relwidth=2 / 200, relheigh=1)

    def __config_calendar(s):
        # cols = s._cal.formatweekheader(3).split()
        cols = ['日', '一', '二', '三', '四', '五', '六']
        s._calendar['columns'] = cols
        s._calendar.tag_configure('header', background='grey90')
        s._calendar.insert('', 'end', values=cols, tag='header')
        # 调整其列宽
        font = tkFont.Font()
        maxwidth = max(font.measure(col) for col in cols)
        for col in cols:
            s._calendar.column(col, width=maxwidth, minwidth=maxwidth,
                               anchor='center')

    def __setup_selection(s, sel_bg, sel_fg):
        def __canvas_forget(evt):
            canvas.place_forget()
            s._selection = None

        s._font = tkFont.Font()
        s._canvas = canvas = tk.Canvas(s._calendar, background=sel_bg, borderwidth=0, highlightthickness=0)
        canvas.text = canvas.create_text(0, 0, fill=sel_fg, anchor='w')

        canvas.bind('<Button-1>', __canvas_forget)
        s._calendar.bind('<Configure>', __canvas_forget)
        s._calendar.bind('<Button-1>', s._pressed)

    def _build_calendar(s):
        year, month = s._date.year, s._date.month

        # update header text (Month, YEAR)
        header = s._cal.formatmonthname(year, month, 0)

        # 更新日历显示的日期
        cal = s._cal.monthdayscalendar(year, month)
        for indx, item in enumerate(s._items):
            week = cal[indx] if indx < len(cal) else []
            fmt_week = [('%02d' % day) if day else '' for day in week]
            s._calendar.item(item, values=fmt_week)

    def _show_select(s, text, bbox):
        """为新的选择配置画布。"""
        x, y, width, height = bbox

        textw = s._font.measure(text)

        canvas = s._canvas
        canvas.configure(width=width, height=height)
        canvas.coords(canvas.text, (width - textw) / 2, height / 2 - 1)
        canvas.itemconfigure(canvas.text, text=text)
        canvas.place(in_=s._calendar, x=x, y=y)

    def _pressed(s, evt=None, item=None, column=None, widget=None):
        """在日历的某个地方点击。"""
        if not item:
            x, y, widget = evt.x, evt.y, evt.widget
            item = widget.identify_row(y)
            column = widget.identify_column(x)

        if not column or not item in s._items:
            # 在工作日行中单击或仅在列外单击。
            return

        item_values = widget.item(item)['values']
        if not len(item_values):  # 这个月的行是空的。
            return

        text = item_values[int(column[1]) - 1]
        if not text:  # 日期为空
            return

        bbox = widget.bbox(item, column)
        if not bbox:  # 日历尚不可见
            s.master.after(20, lambda: s._pressed(item=item, column=column, widget=widget))
            return

        # 更新，然后显示选择
        text = '%02d' % text
        s._selection = (text, item, column)
        s._show_select(text, bbox)

    def _prev_month(s):
        """更新日历以显示前一个月。"""
        s._canvas.place_forget()
        s._selection = None

        s._date = s._date - timedelta(days=1)
        s._date = datetime2(s._date.year, s._date.month, 1)
        s.CB_year.set(s._date.year)
        s.CB_month.set(s._date.month)
        s._update()

    def _next_month(s):
        """更新日历以显示下一个月。"""
        s._canvas.place_forget()
        s._selection = None

        year, month = s._date.year, s._date.month
        s._date = s._date + timedelta(
            days=calendar.monthrange(year, month)[1] + 1)
        s._date = datetime2(s._date.year, s._date.month, 1)
        s.CB_year.set(s._date.year)
        s.CB_month.set(s._date.month)
        s._update()

    def _update(s, event=None, key=None):
        """刷新界面"""
        if key and event.keysym != 'Return': return
        year = int(s.CB_year.get())
        month = int(s.CB_month.get())
        if year == 0 or year > 9999: return
        s._canvas.place_forget()
        s._date = datetime2(year, month, 1)
        s._build_calendar()  # 重建日历

        if year == datetime2.now().year and month == datetime2.now().month:
            day = datetime2.now().day
            for _item, day_list in enumerate(s._cal.monthdayscalendar(year, month)):
                if day in day_list:
                    item = 'I00' + str(_item + 2)
                    column = '#' + str(day_list.index(day) + 1)
                    s.master.after(100, lambda: s._pressed(item=item, column=column, widget=s._calendar))

    def _exit(s, confirm=False):
        """退出窗口"""
        if not confirm: s._selection = None
        s.master.destroy()

    def _main_judge(s):
        """判断窗口是否在最顶层"""
        try:
            # s.master 为 TK 窗口
            # if not s.master.focus_displayof(): s._exit()
            # else: s.master.after(10, s._main_judge)

            # s.master 为 toplevel 窗口
            if s.master.focus_displayof() == None or 'toplevel' not in str(s.master.focus_displayof()):
                s._exit()
            else:
                s.master.after(10, s._main_judge)
        except:
            s.master.after(10, s._main_judge)

        # s.master.tk_focusFollowsMouse() # 焦点跟随鼠标

    def selection(s):
        """返回表示当前选定日期的日期时间。"""
        if not s._selection: return None

        year, month = s._date.year, s._date.month
        return str(datetime2(year, month, int(s._selection[0])))[:10]

    def Input_judgment(s, content):
        """输入判断"""
        # 如果不加上==""的话，就会发现删不完。总会剩下一个数字
        if content.isdigit() or content == "":
            return True
        else:
            return False

def info_finish(scode):
    # 弹出对话框
    result = tk.messagebox.showinfo(title = '信息提示！',message='公司'+scode+'下载完成！')
    # 返回值为：ok
    print(result)

def info_noann(scode):
    # 弹出对话框
    result = tk.messagebox.showinfo(title='信息提示！', message='公司' + scode + '此期间无公告！')
    # 返回值为：ok
    print(result)

def info_confirm(conbg):
    # 弹出对话框
    totalanns = 0
    for m in conbg:
        print(conbg[m]["count"])
        numofeach = conbg[m]["count"]
        totalanns = numofeach+totalanns
    totalanns = str(totalanns)
    result = tk.messagebox.askokcancel(title='信息提示！', message='共有'+totalanns+'个文件'+'\n'+
                                      "确定开始下载吗" )
    print(result)
    return(result)

def info_error(resultcode):
    if resultcode == -1:
        result = tk.messagebox.showinfo(title='信息提示！', message='系统繁忙')
    elif resultcode == 401:
        result = tk.messagebox.showinfo(title='信息提示！', message='未经授权的访问')
    elif resultcode == 402:
        result = tk.messagebox.showinfo(title='信息提示！', message='不合法的参数')
    elif resultcode == 407:
        result = tk.messagebox.showinfo(title='信息提示！', message='免费试用次数已用完')
    elif resultcode == 408:
        result = tk.messagebox.showinfo(title='信息提示！', message='	用户没有余额')
    else:
        pass
    print(result)

root = tk.Tk()
root.resizable(False, False)
root.title('公告下载器 v3.0 by Jonysun from PCCPA 9-5')
width, height = root.winfo_reqwidth() + 350, 250  # 窗口大小
x, y = (root.winfo_screenwidth() - width) / 2, (root.winfo_screenheight() - height) / 2
root.geometry('%dx%d+%d+%d' % (width, height+80, x, y))  # 窗口位置居中
root.update()

label1 = ttk.Label(root, text="公司代码:").place(x=2, y=35)
label2 = ttk.Label(root, text="开始日期:").place(x=2, y=60)
label3 = ttk.Label(root, text="结束日期:").place(x=2, y=90)
label4 = ttk.Label(root, text="保存路径:").place(x=2, y=5)
label5 = tk.Label(root, text="结束日期为必填项", fg="red").place(x=260, y=90)

def select_Path():
    """选取本地路径"""
    path_ = tk.filedialog.askdirectory()  # 返回目录名
    path.set(path_)

def clickb1():
    entry2scode = entry_scode.get()
    sdate = entry_sdate.get()
    edate = entry_edate.get()
    input_dir = input.get()
    conbg = step1(entry2scode,sdate,edate)
    confirming = info_confirm(conbg)
    if confirming == True:
        step2(conbg,input_dir)

def loadtxt():
    txt_path = filedialog.askopenfilename(title='打开txt文件', filetypes=[('记事本', '*.txt'), ('All Files', '*')])
    if txt_path != []:
        lines = []
        with open(txt_path, 'r') as file_to_read:
            while True:
                line = file_to_read.readline()
                if not line:
                    break
                line = line.strip('\n')
                lines.append(line)
        scodelist_ = ";".join(lines)
    scodelist.set(scodelist_)

dt = datetime.datetime.now()
scodelist=tk.StringVar()
today = dt.strftime("%Y-%m-%d")
default_path =os.path.abspath('.')
path = StringVar(value=default_path)
input = ttk.Entry(root, textvariable=path, width=58)
input.place(x=60,y=5)
entry_scode = ttk.Entry(root, textvariable=scodelist, width=58)
entry_scode.place(x=60,y=35)
sdate_str = tk.StringVar()
entry_sdate = ttk.Entry(root, textvariable=sdate_str)
entry_sdate.place(x=60,y=60)
edate_str = tk.StringVar(value=today)
entry_edate = ttk.Entry(root, textvariable=edate_str)
entry_edate.place(x=60,y=90)
readme = "支持多个股票代码输入，两个股票代码之间以英文分号、空格等连接"+"\n"+"导入的TXT文件中股票代码每行输入一个"+"\n"+"如有问题或者建议请发送邮件至sunqize@pccpa.cn"
info = tk.StringVar()
info = scrolledtext.ScrolledText(root, height=13, width=72)
info.place(x=20, y=120)
info.insert(tk.END, readme)

sdate_str_gain = lambda: [
    sdate_str.set(date)
    for date in [Calendar((x, y), 'ur').selection()]
    if date]

edate_str_gain = lambda: [
    edate_str.set(date)
    for date in [Calendar((x, y), 'ur').selection()]
    if date]

button0 = ttk.Button(root, text="保存位置", command=select_Path, width=7).place(x=480, y=5)
button1 = ttk.Button(root, text="执行", command=clickb1, width=7).place(x=480, y=60)
button3 = ttk.Button(root, text="退出", command=root.quit, width=7).place(x=480, y=90)
button4 = ttk.Button(root, text="导入txt", width=7, command=loadtxt).place(x=480, y=35)
button_s = ttk.Button(root, text='选择',width=5, command=sdate_str_gain).place(x=210, y=60)
button_e = ttk.Button(root, text='选择',width=5, command=edate_str_gain).place(x=210, y=90)



# 进度条
sum_length = 420
canvas_progress_bar = Canvas(root, width=sum_length, height=20)
canvas_shape = canvas_progress_bar.create_rectangle(0, 0, 0, 25, fill='green')
canvas_text = canvas_progress_bar.create_text(185, 4, anchor=NW)
canvas_progress_bar.itemconfig(canvas_text, text='00:00:00')
var_progress_bar_percent = StringVar()
var_progress_bar_percent.set('00.00  %')
label_progress_bar_percent = Label(root, textvariable=var_progress_bar_percent, fg='#F5F5F5', bg='#535353')
canvas_progress_bar.place(x=250, y=310, anchor=CENTER)
label_progress_bar_percent.place(x=500, y=310, anchor=CENTER)

root.mainloop()