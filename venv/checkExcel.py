import wx;
from os import remove;
from base64 import b64decode;
from pandas import read_excel;

from wxBg_jpg import img as imgBg;
from wxFile_png import img as imgFile;
from wxQuery_png import img as imgQuery;
from wxTime_png import img as imgTime;
from wxUser_png import img as imgUser;
from wxIcon_ico import img as imgIcon;

imgList = [imgBg, imgFile, imgQuery, imgTime, imgUser, imgIcon]
imgNames = ["imgBg.jpg", "imgFile.png", "imgQuery.png", "imgTime.png", "imgUser.png", "imgIcon.ico"]


def openImg():
    for i in range(0, len(imgList)):
        TEMP = open(imgNames[i], 'wb')
        TEMP.write(b64decode(imgList[i]))
        TEMP.close()


def removeImg():
    for i in range(0, len(imgList)):
        remove(imgNames[i])


def sortKey(list):
    return list[1]


def FilterInfo(fileSrc, wName, wTime):
    _fileSrc = fileSrc.strip()
    if _fileSrc == "":
        return "Tips：请先选择文件"
    _wName = wName.strip()
    if _wName == "":
        return "Tips：姓名不能为空"
    try:
        excel_d = read_excel(_fileSrc, sheet_name=None)
    except IOError:
        return "Error：没有找到文件或读取文件失败"
    peoples = []
    strs = ""
    count = 0
    for sheet in excel_d:
        for people in excel_d[sheet].iloc:
            nTime = "00:00" if str(people[3]) == "nan" else str(people[3])
            nPeople = "" if str(people[1]) == "nan" else str(people[1])
            if nTime > wTime and nPeople.find(_wName) >= 0:
                count += 1
                peoples.append(people)
    peoples.sort(key=sortKey)
    curPeople = peoples[0][1]
    for _people in peoples:
        if _people[1] != curPeople:
            curPeople = _people[1]
            strs += "\n"
        for info in _people:
            infoStr = str(info)
            if infoStr == "nan":
                infoStr = ""
            strs += infoStr + "　"
        strs += "\n"
    strs = "Total： " + str(count) + " 条记录\n\n" + strs
    return strs


class TransparentText(wx.StaticText):
    def __init__(self, parent, id=wx.ID_ANY, label='', pos=wx.DefaultPosition, size=wx.DefaultSize,
                 style=wx.TRANSPARENT_WINDOW, name='transparenttext'):
        wx.StaticText.__init__(self, parent, id, label, pos, size, style, name)
        self.Bind(wx.EVT_PAINT, self.on_paint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda event: None)
        self.Bind(wx.EVT_SIZE, self.on_size)

    def on_paint(self, event):
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False)
        bdc = wx.PaintDC(self)
        dc = wx.GCDC(bdc)
        font_color = "#FFF"
        dc.SetFont(font)
        dc.SetTextForeground(font_color)
        dc.DrawText(self.GetLabel(), 0, 0)

    def on_size(self, event):
        self.Refresh()
        event.Skip()


class createWin(wx.Frame):
    # auth @Lee
    def __init__(self):
        # 窗口宽高
        width = 800
        height = 750

        # 创建临时图片
        openImg()

        # 主体框架
        PFrame = wx.Frame.__init__(self, None, title='Late Come Query Tool', size=(width, height),
                                   style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER ^ wx.MAXIMIZE_BOX)
        wx.Frame.SetMinSize(self, (width, height))
        wx.Frame.SetMaxSize(self, (width, height))
        self.Center()


        Icon = wx.Icon("imgIcon.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(Icon)

        # 主体容器
        cPanel = wx.Panel(parent=self, size=(width, height))

        # 主体背景容器
        bgImg = wx.Image("imgBg.jpg", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        cPanelBg = wx.StaticBitmap(cPanel, -1, bgImg, (0, 0))

        # 字体
        sFont = wx.Font(13, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, "微软雅黑")
        cFont = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, "微软雅黑")

        textBgColor = "#252525"
        textColor = "#FFF"

        # 文件
        FilePanel = wx.Panel(parent=cPanelBg, pos=(40, 40), size=(580, 40))
        FilePanel.SetBackgroundColour(textBgColor)
        FileImage = wx.Image('imgFile.png', wx.BITMAP_TYPE_PNG).ConvertToBitmap()
        FileIcon = wx.StaticBitmap(FilePanel, -1, FileImage, (8, 8))
        self.FileValue = wx.TextCtrl(parent=FilePanel, pos=(40, 8), size=(530, 24), style=wx.TE_READONLY | wx.NO_BORDER)
        self.FileValue.SetBackgroundColour(textBgColor), self.FileValue.SetForegroundColour(textColor)
        self.FileValue.Bind(wx.EVT_LEFT_UP, self.OnOpenFile)
        self.FileValue.SetFont(sFont)

        # 姓名
        NamePanel = wx.Panel(parent=cPanelBg, pos=(40, 95), size=(285, 40))
        NamePanel.SetBackgroundColour(textBgColor)
        NameImage = wx.Image('imgUser.png', wx.BITMAP_TYPE_PNG).ConvertToBitmap()
        NameIcon = wx.StaticBitmap(NamePanel, -1, NameImage, (8, 8))
        self.NameValue = wx.TextCtrl(parent=NamePanel, pos=(40, 8), size=(235, 24),
                                     style=wx.TE_PROCESS_ENTER | wx.NO_BORDER)
        self.NameValue.SetBackgroundColour(textBgColor), self.NameValue.SetForegroundColour(textColor)
        self.NameValue.Bind(wx.EVT_TEXT_ENTER, self.ReadFile)
        self.NameValue.SetFont(sFont)

        # 时间
        TimePanel = wx.Panel(parent=cPanelBg, pos=(340, 95), size=(280, 40))
        TimePanel.SetBackgroundColour(textBgColor)
        TimeImage = wx.Image('imgTime.png', wx.BITMAP_TYPE_PNG).ConvertToBitmap()
        TimeIcon = wx.StaticBitmap(TimePanel, -1, TimeImage, (8, 8))
        self.TimeValue = wx.TextCtrl(parent=TimePanel, pos=(40, 8), size=(235, 24),
                                     style=wx.TE_PROCESS_ENTER | wx.NO_BORDER)
        self.TimeValue.SetBackgroundColour(textBgColor), self.TimeValue.SetForegroundColour(textColor)
        self.TimeValue.Bind(wx.EVT_TEXT_ENTER, self.ReadFile)
        self.TimeValue.SetFont(sFont)

        # 查询按钮
        BMP = wx.Image("imgQuery.png", wx.BITMAP_TYPE_PNG).ConvertToBitmap()
        self.QueryBtn = wx.BitmapButton(cPanelBg, -1, BMP, pos=(640, 41), size=(95, 95), style=wx.NO_BORDER)
        self.QueryBtn.SetBackgroundColour(textBgColor)
        self.QueryBtn.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.QueryBtn.Bind(wx.EVT_BUTTON, self.ReadFile)

        # 文本域
        DataPanel = wx.Panel(parent=cPanelBg, pos=(38, 175), size=(708, 496))
        DataPanel.SetBackgroundColour(textBgColor)
        self.QueryData = wx.TextCtrl(parent=DataPanel, pos=(10, 10), size=(719, 476),
                                     style=wx.TE_MULTILINE | wx.NO_BORDER)
        self.QueryData.SetBackgroundColour(textBgColor), self.QueryData.SetForegroundColour(textColor)
        self.QueryData.SetFont(cFont)

        # 删除临时图片
        removeImg()

    def OnOpenFile(self, event):
        wildcard = 'Excel files(*.xls;*.xlsx)|*.xls;*.xlsx'
        dialog = wx.FileDialog(
            self, message="请选择Excel文件",
            defaultFile="",
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_CHANGE_DIR
        )
        if dialog.ShowModal() == wx.ID_OK:
            self.FileValue.SetValue(dialog.GetPath())
            dialog.Destroy

    def ReadFile(self, event):
        gInfo = FilterInfo(self.FileValue.GetLineText(0), self.NameValue.GetLineText(0), self.TimeValue.GetLineText(0))
        self.QueryData.SetValue(gInfo)


if __name__ == '__main__':
    app = wx.App()
    MainFrame = createWin()
    MainFrame.Show()
    app.MainLoop()
