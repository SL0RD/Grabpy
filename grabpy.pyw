import wx
import win32con
import desktopmagic.screengrab_win32 as grab
import win32gui, win32clipboard
import pyhk
import pysftp
import os, time, datetime, ConfigParser, logging

config = ConfigParser.RawConfigParser()
config_path = '%s\\grabpy\\' % os.environ['APPDATA']
config_file = '%s\\grabpy.ini' % config_path
log_file = '{0}logs\\grabpy.log'.format(config_path)
grabs_dir = '{0}grabs\\'.format(config_path)

if not os.path.exists(config_path):
    os.makedirs(config_path)

if not os.path.exists(config_path+'logs'):
    os.makedirs(config_path+'logs')

if not os.path.exists(config_path+'grabs'):
    os.makedirs(config_path+'grabs')

default_main_conf = {'copy_to_clipboard': True,
                     'remove_local': False,
                     'prt_scn_func': None,
                     'logging_level': 'INFO',
                     'grab_name': '%H%M%S-%d%m%Y',
                     'append_to_name': 'http://example.com/'}

default_sftp_conf = {'use_sftp': False,
                     'sftp_host': 'example.com',
                     'sftp_user': 'user',
                     'sftp_pass': 'secretpassword',
                     'sftp_port': 22,
                     'sftp_path': 'remote/path/on/server/',
                     'sftp_use_key': False,
                     'sftp_key': '/path/to/key'}

logging.basicConfig(filename=log_file,
                    format='%(asctime)s - %(levelname)s:%(message)s',
                    level='DEBUG',
                    datefmt='[%m/%d/%Y %I:%M:%S %p]')

if not os.path.isfile(config_file):
    open(config_file, 'a').close()

    config.add_section('Main App')
    for o in default_main_conf:
        v = default_main_conf[o]
        config.set('Main App', o, v)

    config.add_section('SFTP Settings')
    for o in default_sftp_conf:
        v = default_sftp_conf[o]
        config.set('SFTP Settings', o, v)

    logging.info("No config file found, creating default config file.")
    with open(config_file, 'wb') as default_config_file:
        config.write(default_config_file)


def str_to_bool(s):
    if s == 'True':
        return True
    else:
        return False


def check_config(conf_file):
    config.read(conf_file)
    missing_options = {}
    logging.info('Checking config file')
    for o in default_main_conf:
        if not config.has_option('Main App', o):
            logging.warning('Config option missing, adding defaults.')
            missing_options[o] = default_main_conf[o]

    print missing_options


def getconfig(conf_file):
    global_config = {}

    config.read(conf_file)

    global_config['main'] = {}
    global_config['sftp'] = {}

    for o in config.options('Main App'):
        global_config['main'][o] = config.get('Main App', o)
    for o in config.options('SFTP Settings'):
        global_config['sftp'][o] = config.get('SFTP Settings', o)

    return global_config

check_config(config_file)


def gen_filename():
    return get_timestamp(app.conf['main']['grab_name'])+'.png'


def get_timestamp(date_format=None):
    if date_format is None:
        date_format = "%H%M%S%d%m%Y"
    ts = time.time()
    ts = datetime.datetime.fromtimestamp(ts).strftime(date_format)
    return str(ts)


def set_clipboard(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text)
    win32clipboard.CloseClipboard()

TRAY_TOOLTIP = 'Grabpy'
TRAY_ICON = 'icon.ico'

hot = pyhk.pyhk()


def create_menu_item(menu, label, func):
    item = wx.MenuItem(menu, -1, label)
    menu.Bind(wx.EVT_MENU, func, id=item.GetId())
    menu.AppendItem(item)
    return item


class Settings(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, title="Settings", style=wx.DEFAULT_FRAME_STYLE & ~wx.MAXIMIZE_BOX ^wx.RESIZE_BORDER)
        ico = wx.Icon(TRAY_ICON, wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        self.config = getconfig(config_file)
        self.panel = wx.Panel(self)

        log_levels = ['INFO', 'DEBUG', 'WARNING', 'ERROR']
        prntscrn_func = ['None', 'Whole Screen', 'Active Window', 'Selection']

        font = wx.SystemSettings_GetFont(wx.SYS_SYSTEM_FONT)
        font.SetPointSize(9)

        self.mainTitle = wx.StaticText(self.panel, -1, 'Main App Settings')
        self.mainTitle.SetFont(font)
        self.copy_toClipboard = wx.CheckBox(self.panel, -1, 'Copy URL to clipboard')
        self.copy_toClipboard.SetValue(str_to_bool(app.conf['main']['copy_to_clipboard']))
        self.append_clip_label = wx.StaticText(self.panel, -1, 'Append to clipboard:')
        self.append_clip = wx.TextCtrl(self.panel)
        self.append_clip.SetValue(app.conf['main']['append_to_name'])
        self.prn_scn_label = wx.StaticText(self.panel, -1, 'PrintScreen Function:')
        self.prt_scrn_func = wx.ComboBox(self.panel, -1, choices=prntscrn_func, style=wx.CB_READONLY)
        self.prt_scrn_func.SetValue(app.conf['main']['prt_scn_func'])

        self.logLevel_label = wx.StaticText(self.panel, -1, 'Logging Level:')
        self.logLevel_cb = wx.ComboBox(self.panel, -1, choices=log_levels, style=wx.CB_READONLY)

        self.logLevel_cb.SetValue(app.conf['main']['logging_level'])

        self.sftpTitle = wx.StaticText(self.panel, -1, 'SFTP Settings')
        self.sftpTitle.SetFont(font)
        self.sftp_upload = wx.CheckBox(self.panel, -1, 'Enable SFTP Upload')
        self.sftp_upload.SetValue(str_to_bool(app.conf['sftp']['use_sftp']))
        self.host_label = wx.StaticText(self.panel, -1, 'SFTP Hostname:')
        self.port_label = wx.StaticText(self.panel, -1, 'SFTP Port:')
        self.username_label = wx.StaticText(self.panel, -1, 'SFTP Username:')
        self.passwd_label = wx.StaticText(self.panel, -1, 'SFTP Password:')
        self.path_label = wx.StaticText(self.panel, -1, 'SFTP Path:')
        self.host_value = wx.TextCtrl(self.panel)
        self.port_value = wx.TextCtrl(self.panel)
        self.username_value = wx.TextCtrl(self.panel)
        self.passwd_value = wx.TextCtrl(self.panel, style=wx.TE_PASSWORD)
        self.path_value = wx.TextCtrl(self.panel)
        self.use_key = wx.CheckBox(self.panel, -1, 'Use SSH key for auth')
        self.kpath_label = wx.StaticText(self.panel, -1, 'SFTP Key Path:')
        self.kpath_value = wx.TextCtrl(self.panel)
        self.host_value.SetValue(app.conf['sftp']['sftp_host'])
        self.username_value.SetValue(app.conf['sftp']['sftp_user'])
        self.port_value.SetValue(str(app.conf['sftp']['sftp_port']))
        self.path_value.SetValue(app.conf['sftp']['sftp_path'])
        self.use_key.SetValue(str_to_bool(app.conf['sftp']['sftp_use_key']))
        self.kpath_value.SetValue(app.conf['sftp']['sftp_key'])

        self.applyButton = wx.Button(self.panel, label='Apply')
        self.okayButton = wx.Button(self.panel, label='OK')

        self.windowSizer = wx.BoxSizer()
        self.windowSizer.Add(self.panel, 1, wx.ALL | wx.EXPAND, 10)

        self.mainSizer = wx.GridSizer(5, 2)
        self.mainSizer.Add(self.mainTitle)
        self.mainSizer.Add(wx.StaticText(self))
        self.mainSizer.Add(self.prn_scn_label)
        self.mainSizer.Add(self.prt_scrn_func)
        self.mainSizer.Add(wx.StaticText(self))
        self.mainSizer.Add(self.copy_toClipboard)
        self.mainSizer.Add(self.append_clip_label)
        self.mainSizer.Add(self.append_clip)
        self.mainSizer.Add(self.logLevel_label)
        self.mainSizer.Add(self.logLevel_cb)

        self.sizer = wx.GridSizer(10, 2)
        self.sizer.Add(self.sftpTitle)
        self.sizer.Add(wx.StaticText(self))
        self.sizer.Add(wx.StaticText(self))
        self.sizer.Add(self.sftp_upload)
        self.sizer.Add(self.host_label,  flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.host_value)
        self.sizer.Add(self.port_label, flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.port_value)
        self.sizer.Add(self.username_label,  flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.username_value)
        self.sizer.Add(self.passwd_label,  flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.passwd_value)
        self.sizer.Add(self.path_label,  flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.path_value)
        self.sizer.Add(wx.StaticText(self))
        self.sizer.Add(self.use_key)
        self.sizer.Add(self.kpath_label, flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.kpath_value)

        self.sizer.Add(self.applyButton, flag=wx.ALIGN_RIGHT)
        self.sizer.Add(self.okayButton)

        self.border = wx.BoxSizer(wx.VERTICAL)
        self.border.Add(self.mainSizer, 1, wx.ALL, 5)
        self.border.Add(wx.StaticLine(self.panel, size=(300, 1)))
        self.border.Add(self.sizer, 1, wx.ALL, 5)
        self.panel.SetSizerAndFit(self.border)
        self.SetSizerAndFit(self.windowSizer)

        self.applyButton.Bind(wx.EVT_BUTTON, self.OnApply)
        self.okayButton.Bind(wx.EVT_BUTTON, self.OnOkay)

    def OnOkay(self, event):
        self.Close(True)

    def get_new_values(self):

        self.new_clip = self.copy_toClipboard.GetValue()
        self.new_prtscrf = str(self.prt_scrn_func.GetValue())
        self.new_append = str(self.append_clip.GetValue())
        self.new_logging = str(self.logLevel_cb.GetValue())

        self.newStatus = self.sftp_upload.GetValue()
        self.newHostname = str(self.host_value.GetValue())
        self.newPort = int(self.port_value.GetValue())
        self.newUsername = str(self.username_value.GetValue())
        self.new_pass = str(self.passwd_value.GetValue())
        self.newPath = str(self.path_value.GetValue())
        self.new_key_status = str(self.use_key.GetValue())
        self.new_key_path = str(self.kpath_value.GetValue())

        return {'main': {'logging_level': self.new_logging,
                         'prt_scn_func': self.new_prtscrf,
                         'append_to_name': self.new_append,
                         'grab_name': app.conf['main']['grab_name'],
                         'copy_to_clipboard': self.new_clip},
                'sftp': {'use_sftp': self.newStatus,
                        'sftp_host': self.newHostname,
                        'sftp_user': self.newUsername,
                        'sftp_pass': self.new_pass,
                        'sftp_port': self.newPort,
                        'sftp_path': self.newPath,
                        'sftp_use_key': self.new_key_status,
                        'sftp_key': self.new_key_path}}

    def OnApply(self, event):
        new_conf_values = self.get_new_values()
        print new_conf_values
        result = app.save_config(config_file, new_conf_values)
        if result == 0:
            print 'error saving'
            return 0
        else:
            print 'success'
        app.conf = new_conf_values
        app.set_prtscn_hk()


class TaskBarIcon(wx.TaskBarIcon):

    def __init__(self):
        self.tgul = ''
        self.dupload = None
        super(TaskBarIcon, self).__init__()
        self.set_icon(TRAY_ICON)
        self.Bind(wx.EVT_TASKBAR_LEFT_DOWN, self.on_left_down)
        hot.addHotkey(['Alt', 'Shift', 'F10'], self.hk_selectable_area)
        hot.addHotkey(['Alt', 'Shift', 'F11'], self.get_active_window, isThread=True)
        hot.addHotkey(['Alt', 'Shift', 'F12'], self.get_whole_screen, isThread=True)

    def hk_selectable_area(self):
        frame = SelectableFrame()
        frame.Show()

    def toggle_upload(self, event):
        if app.conf['sftp']['use_sftp'] == 'True':
            app.conf['sftp']['use_sftp'] = 'False'
            logging.debug('Disabling automatic file upload')
            app.save_config(config_file, app.conf)
        else:
            app.conf['sftp']['use_sftp'] = 'True'
            logging.debug('Enabling automatic file upload')
            app.save_config(config_file, app.conf)

    def selectable_area(self, event):
        frame = SelectableFrame()
        frame.Show()

    def CreatePopupMenu(self):

        if not hasattr(self, 'Disable Upload'):
            self.dupload = wx.NewId()
            self.Bind(wx.EVT_MENU, self.toggle_upload, id=self.dupload)

        menu = wx.Menu()
        self.tgul = menu.Append(self.dupload, 'Disable Upload', 'Disable Upload', kind=wx.ITEM_CHECK)
        create_menu_item(menu, 'Settings', self.on_settings)
        menu.AppendSeparator()
        create_menu_item(menu, 'Capture selection', self.selectable_area)
        menu.AppendSeparator()
        create_menu_item(menu, 'Exit', self.on_exit)
        stat = str(app.conf['sftp']['use_sftp'])
        print '"'+stat+'"'
        if stat == 'False':
            print self.tgul.IsChecked()
            if self.tgul.IsChecked():
                return
            else:
                self.tgul.Check()
        return menu

    def on_settings(self, event):
        logging.info('Launching settings window')
        frame = Settings()
        frame.Show()

    def set_icon(self, path):
        icon = wx.IconFromBitmap(wx.Bitmap(path))
        self.SetIcon(icon, TRAY_TOOLTIP)

    def on_left_down(self, event):
        print 'Tray icon was left-clicked.'

    def on_exit(self, event):
        wx.CallAfter(self.Destroy)

    def get_active_window(self):

        hwnd = win32gui.GetForegroundWindow()
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        logging.info("Grabbing window with coordinates (%d,%d,%d,%d)" % (l, t, r, b))

        filename = gen_filename()
        path = grabs_dir + filename
        activewindow = grab.getRectAsImage((l, t, r, b))
        activewindow.save(path, format='png')
        app.put_sftp(path)
        clip = app.conf['main']['append_to_name'] + filename
        set_clipboard(clip)

    def get_whole_screen(self):
        logging.info("Grabbing entire screen")
        wholescreen = grab.getScreenAsImage()
        filename = gen_filename()
        path = grabs_dir + filename
        wholescreen.save(path, format='png')
        app.put_sftp(path)
        clip = app.conf['main']['append_to_name'] + filename
        set_clipboard(clip)


class SelectableFrame(wx.Frame):

    c1 = None
    c2 = None
    rc1 = None
    rc2 = None
    sel_coords = ()

    def __init__(self):

        wx.Frame.__init__(self, None, title="Selectable Area", size=wx.DisplaySize(), style=wx.STAY_ON_TOP)
        self.ShowFullScreen(True)
        self.SetFocus()

        self.panel = wx.Panel(self, size=self.GetSize())
        self.panel.Bind(wx.EVT_MOTION, self.OnMouseMove)
        self.panel.Bind(wx.EVT_LEFT_DOWN, self.OnMouseDown)
        self.panel.Bind(wx.EVT_LEFT_UP, self.OnMouseUp)
        self.panel.Bind(wx.EVT_PAINT, self.OnPaint)

        self.SetCursor(wx.StockCursor(wx.CURSOR_CROSS))
        self.SetBackgroundColour(wx.BLUE)
        self.SetTransparent(50)
        self.Refresh()

    def OnMouseMove(self, event):
        if event.Dragging() and event.LeftIsDown():
            self.c2 = event.GetPosition()
            self.Refresh()

    def OnMouseDown(self, event):
        self.c1 = event.GetPosition()
        self.rc1 = win32gui.GetCursorPos()

    def OnMouseUp(self, event):
        self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        self.rc2 = win32gui.GetCursorPos()
        self.SetTransparent(0)

        l, t = self.rc1
        r, b = self.rc2

        if t > b:
            ob = b
            b = t
            t = ob
        if l > r:
            ore = r
            r = l
            l = ore
        filename = gen_filename()
        path = grabs_dir + filename
        selection = grab.getRectAsImage((l,t,r,b))
        selection.save(path, format='png')
        app.put_sftp(path)
        clip = app.conf['main']['append_to_name'] + filename
        set_clipboard(clip)
        self.Close()

    def OnPaint(self, event):
        if self.c1 is None or self.c2 is None: return

        dc = wx.PaintDC(self.panel)
        dc.SetPen(wx.Pen(wx.YELLOW, 2))
        dc.SetBrush(wx.Brush(wx.Colour(0, 0, 0), wx.TRANSPARENT))

        dc.DrawRectangle(self.c1.x, self.c1.y, self.c2.x - self.c1.x, self.c2.y - self.c1.y)

        self.sel_coords = (self.c1.x, self.c1.y, self.c2.x, self.c2.y)


class MyApp(wx.App, TaskBarIcon, SelectableFrame):
    def __init__(self):
        wx.App.__init__(self)
        self.conf = getconfig(config_file)

        if self.conf['main']['prt_scn_func'] == 'Selection':
            self.sn_sl = hot.addHotkey(['Snapshot'], self.hk_selectable_area, up=True)
            self.sn_id = self.sn_sl
        elif self.conf['main']['prt_scn_func'] == 'Whole Screen':
            self.sn_ws = hot.addHotkey(['Snapshot'], self.get_whole_screen, up=True, isThread=True)
            self.sn_id = self.sn_ws
        elif self.conf['main']['prt_scn_func'] == 'Active Window':
            self.sn_aw = hot.addHotkey(['Snapshot'], self.get_active_window, up=True, isThread=True)
            self.sn_id = self.sn_aw
        else:
            return

        if str_to_bool(self.conf['sftp']['sftp_use_key']):
            self.cinfo = {'host': self.conf['sftp']['sftp_host'],
                          'username': self.conf['sftp']['sftp_user'],
                          'private_key': self.conf['sftp']['sftp_key'],
                          'port': int(self.conf['sftp']['sftp_port'])}
        else:
            self.cinfo = {'host': self.conf['sftp']['sftp_host'],
                          'username': self.conf['sftp']['sftp_user'],
                          'password': self.conf['sftp']['sftp_pass'],
                          'port': int(self.conf['sftp']['sftp_port'])}

        logging.basicConfig(filename=log_file,
                    format='%(asctime)s - %(levelname)s:%(message)s',
                    level=self.conf['main']['logging_level'],
                    datefmt='[%m/%d/%Y %I:%M:%S %p]')

    def OnInit(self):
        self.tray = TaskBarIcon()
        return True

    def set_prtscn_hk(self):

        if self.conf['main']['prt_scn_func'] == 'Selection':
            if self.sn_id:
                hot.removeHotkey(id=self.sn_id)
            self.sn_sl = hot.addHotkey(['Snapshot'], self.hk_selectable_area, up=True)
            self.sn_id = self.sn_sl
        elif self.conf['main']['prt_scn_func'] == 'Whole Screen':
            if self.sn_id:
                hot.removeHotkey(id=self.sn_id)
            self.sn_ws = hot.addHotkey(['Snapshot'], self.get_whole_screen, up=True, isThread=True)
            self.sn_id = self.sn_sl
        elif self.conf['main']['prt_scn_func'] == 'Active Window':
            if self.sn_id:
                hot.removeHotkey(id=self.sn_id)
            self.sn_aw = hot.addHotkey(['Snapshot'], self.get_active_window, up=True, isThread=True)
            self.sn_id = self.sn_aw
        else:
            if self.sn_id:
                hot.removeHotkey(id=self.sn_id)
            return

    def put_sftp(self, filename):
        if self.conf['sftp']['use_sftp'] == "True":
            print self.cinfo
            with pysftp.Connection(**self.cinfo) as sftp:
                logging.debug("Connected to ftp server")
                with sftp.cd(self.conf['sftp']['sftp_path']):
                    logging.debug("Uploading file")
                    print filename
                    sftp.put(filename, preserve_mtime=True)
            logging.debug('disconnecting')
            self.tray.ShowBalloon('', 'Upload complete', 3000)

    def save_config(self, cfile, conf_dic):
        main = conf_dic['main']
        sftp = conf_dic['sftp']
        for o in main:
            v = main[o]
            res = self.setconfigoption('Main App', o, v)
            if res == 0:
                print 'error: '+o+' : '+v
                return 0

        for o in sftp:
            v = sftp[o]
            res = self.setconfigoption('SFTP Settings', o, v)
            if res == 0:
                print 'error: '+o+' : '+v
                return 0

        with open(cfile, 'wb') as conf_file:
            config.write(conf_file)

    def setconfigoption(self, section, option, value):
        if not config.has_section(section):
            logging.error('Section not found in config')
            return 0
        if not config.has_option(section, option):
            logging.error('Option not found in section')
            return 0

        try:
            config.set(section, option, value)
            return 1
        except:
            print "Config Error"
            return 0


if __name__ == '__main__':
    app = MyApp()
    app.MainLoop()