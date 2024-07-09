import wx
import pcbnew

from .thread import *
from .result_event import *
from .process import *

class KiCadToPCBWayForm(wx.Frame):
    def __init__(self):
        wx.Dialog.__init__(
            self,
            None,
            id=wx.ID_ANY,
            title=u"PCBWay is processing...",
            pos=wx.DefaultPosition,
            size=wx.DefaultSize,
            style=wx.DEFAULT_DIALOG_STYLE)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.m_gaugeStatus = wx.Gauge(
            self, wx.ID_ANY, 100, wx.DefaultPosition, wx.Size(
                300, 20), wx.GA_HORIZONTAL)
        self.m_gaugeStatus.SetValue(0)
        bSizer1.Add(self.m_gaugeStatus, 0, wx.ALL, 5)

        self.SetSizer(bSizer1)
        self.Layout()
        bSizer1.Fit(self)

        self.Centre(wx.BOTH)

        EVT_RESULT(self, self.updateDisplay)
        PCBWayThread(self)

    def updateDisplay(self, status):
        if status.data == -1:
            pcbnew.Refresh()
            self.Destroy()
        else:
            self.m_gaugeStatus.SetValue(int(status.data))


class PCBWayPlugin(pcbnew.ActionPlugin):
    def __init__(self):
        self.name = "PCBWay Plug-in for KiCad"  # 插件名称
        self.category = "Manufacturing"  # 描述性类别名称
        self.description = "Start prototype and assembly by sending files to PCBWay with just one click."  # 对插件及其功能的描述
        self.pcbnew_icon_support = hasattr(self, "show_toolbar_button")
        self.show_toolbar_button = True  # 可选，默认为 False
        self.icon_file_name = os.path.join(
            os.path.dirname(__file__), 'icon.png')  # 可选，默认为 ""
        self.dark_icon_file_name = os.path.join(
            os.path.dirname(__file__), 'icon.png')

    def Run(self):
        # 在用户操作时执行的插件的入口函数
        KiCadToPCBWayForm().Show()
