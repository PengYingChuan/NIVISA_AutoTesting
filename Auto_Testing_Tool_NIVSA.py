# =========== PyVisa  ===========
import visa
from pyvisa import attributes, constants
from ctypes import c_int, byref, c_byte, c_uint8
#from win32file import * # The base COM port and file IO functions.
#import win32con # constants.
#from collections import defaultdict
import time

READ_DATA = c_byte
status = constants.StatusCode
data_out_buffer = c_uint8 * 12
COM_PORTS=[]

# =========== wxPython  ===========
import wx
import os
import ctypes
from comtypes.GUID import GUID
from ctypes import wintypes
import numpy as np
import matplotlib
matplotlib.use('WXAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import \
    FigureCanvasWxAgg as FigCanvas
from collections import defaultdict
import matplotlib.pyplot as plt
import pylab
import random
import win32com.client

import pdb

wildcard = "CSV (*.csv)|*.csv|" \
            "All files (*.*)|*.*"

##########################################################################
## Notification Definition
###########################################################################
GWL_WNDPROC = -4
WM_DESTROY  = 2
DBT_DEVTYP_DEVICEINTERFACE = 0x00000005  # device interface class
DBT_DEVICEREMOVECOMPLETE = 0x8004  # device is gone
DBT_DEVICEARRIVAL = 0x8000  # system detected a new device
WM_DEVICECHANGE = 0x0219
## It's probably not neccesary to make this distinction, but it never hurts to be safe
if 'unicode' in wx.PlatformInfo:
    SetWindowLong = ctypes.windll.user32.SetWindowLongW
    CallWindowProc = ctypes.windll.user32.CallWindowProcW
else:
    SetWindowLong = ctypes.windll.user32.SetWindowLongA
    CallWindowProc = ctypes.windll.user32.CallWindowProcA

## Create a type that will be used to cast a python callable to a c callback function
## first arg is return type, the rest are the arguments
#WndProcType = ctypes.WINFUNCTYPE(c_int, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
WndProcType = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, ctypes.c_uint, ctypes.c_int, ctypes.c_int)

if 'unicode' in wx.PlatformInfo:
    RegisterDeviceNotification = ctypes.windll.user32.RegisterDeviceNotificationW
else:
    RegisterDeviceNotification = ctypes.windll.user32.RegisterDeviceNotificationA
RegisterDeviceNotification.restype = wintypes.HANDLE
RegisterDeviceNotification.argtypes = [wintypes.HANDLE, wintypes.c_void_p, wintypes.DWORD]

UnregisterDeviceNotification = ctypes.windll.user32.UnregisterDeviceNotification
UnregisterDeviceNotification.restype = wintypes.BOOL
UnregisterDeviceNotification.argtypes = [wintypes.HANDLE]

class DEV_BROADCAST_DEVICEINTERFACE(ctypes.Structure):
    _fields_ = [("dbcc_size", ctypes.c_ulong),
                  ("dbcc_devicetype", ctypes.c_ulong),
                  ("dbcc_reserved", ctypes.c_ulong),
                  ("dbcc_classguid", GUID),
				  #("dbcc_classguid", GUID.GUID),
                  ("dbcc_name", ctypes.c_wchar * 256)]

class DEV_BROADCAST_HDR(ctypes.Structure):
    _fields_ = [("dbch_size", wintypes.DWORD),
                ("dbch_devicetype", wintypes.DWORD),
                ("dbch_reserved", wintypes.DWORD)]

class WndProcHookMixin:
    """
    This class can be mixed in with any wxWindows window class in order to hook it's WndProc function.
    You supply a set of message handler functions with the function addMsgHandler. When the window receives that
    message, the specified handler function is invoked. If the handler explicitly returns False then the standard
    WindowProc will not be invoked with the message. You can really screw things up this way, so be careful.
    This is not the correct way to deal with standard windows messages in wxPython (i.e. button click, paint, etc)
    use the standard wxWindows method of binding events for that. This is really for capturing custom windows messages
    or windows messages that are outside of the wxWindows world.
    """
    def __init__(self):
        self.__msgDict = {}
        ## We need to maintain a reference to the WndProcType wrapper
        ## because ctypes doesn't
        self.__localWndProcWrapped = None
        self.rtnHandles = []

    def hookWndProc(self):
        self.__localWndProcWrapped = WndProcType(self.localWndProc)
        self.__oldWndProc = SetWindowLong(self.GetHandle(),
                                        GWL_WNDPROC,
                                        self.__localWndProcWrapped)
    def unhookWndProc(self):
        SetWindowLong(self.GetHandle(),
                        GWL_WNDPROC,
                        self.__oldWndProc)

        ## Allow the ctypes wrapper to be garbage collected
        self.__localWndProcWrapped = None

    def addMsgHandler(self,messageNumber,handler):
        self.__msgDict[messageNumber] = handler

    def localWndProc(self, hWnd, msg, wParam, lParam):
        # call the handler if one exists
        # performance note: "in" is the fastest way to check for a key
        # when the key is unlikely to be found
        # (which is the case here, since most messages will not have handlers).
        # This is called via a ctypes shim for every single windows message
        # so dispatch speed is important
        if msg in self.__msgDict:
            # if the handler returns false, we terminate the message here
            # Note that we don't pass the hwnd or the message along
            # Handlers should be really, really careful about returning false here
            if self.__msgDict[msg](wParam,lParam) == False:
                return

        # Restore the old WndProc on Destroy.
        if msg == WM_DESTROY: self.unhookWndProc()

        return CallWindowProc(self.__oldWndProc,
                                hWnd, msg, wParam, lParam)

    def registerDeviceNotification(self, guid, devicetype=DBT_DEVTYP_DEVICEINTERFACE):
        devIF = DEV_BROADCAST_DEVICEINTERFACE()
        devIF.dbcc_size = ctypes.sizeof(DEV_BROADCAST_DEVICEINTERFACE)
        devIF.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE

        if guid:
            #devIF.dbcc_classguid = GUID.GUID(guid)
            devIF.dbcc_classguid = GUID(guid)
        return RegisterDeviceNotification(self.GetHandle(), ctypes.byref(devIF), 0)

    def unregisterDeviceNotification(self, handle):
        if UnregisterDeviceNotification(handle) == 0:
            raise Exception("Unable to unregister device notification messages")
###########################################################################
## Class Top_Option
###########################################################################
class Top_Option ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Anpec Auto Testing Tool", pos = wx.DefaultPosition, size = wx.Size( 462,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
        #favicon = wx.Icon(r'./icon/anpec.ico', wx.BITMAP_TYPE_ICO, 16, 16)
       # wx.Frame.SetIcon(self, favicon)

        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
        self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_SCROLLBAR ) )

        bSizer7 = wx.BoxSizer( wx.VERTICAL )

        self.Bt1_None = wx.RadioButton( self, wx.ID_ANY, u"0. None", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt1_None, 0, wx.ALL, 5 )

        self.Bt2_USB_I2C = wx.RadioButton( self, wx.ID_ANY, u"1. Control :USB I2C Control", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt2_USB_I2C, 0, wx.ALL, 5 )

        self.Bt3_VICI = wx.RadioButton( self, wx.ID_ANY, u"2. Sweep  : X-> To Change Vi voltage       , Y-> To measure current of Vi", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt3_VICI, 0, wx.ALL, 5 )

        self.Bt4_VIVO = wx.RadioButton( self, wx.ID_ANY, u"3. Sweep  : X-> To Change Vi voltage       , Y-> To measure voltage of Vo", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt4_VIVO, 0, wx.ALL, 5 )

        self.Bt5_VIFO = wx.RadioButton( self, wx.ID_ANY, u"4. Sweep  : X-> To Change Vi voltage       , Y-> To measure frequency    ", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt5_VIFO, 0, wx.ALL, 5 )

        self.Bt6_COVO = wx.RadioButton( self, wx.ID_ANY, u"5. Sweep  : X-> To Change Loading          , Y-> To measure voltage of Vo", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt6_COVO, 0, wx.ALL, 5 )

        self.Bt7_REGVO = wx.RadioButton( self, wx.ID_ANY, u"6. Sweep  : X-> To Change Regs (I2C)       , Y-> To measure voltage of Vo", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt7_REGVO, 0, wx.ALL, 5 )

        #self.Bt8_TOFO = wx.RadioButton( self, wx.ID_ANY, u"7. Sweep  : X-> To measure Temperature, Y-> To measure frequency         ", wx.DefaultPosition, wx.DefaultSize, 0 )
        #bSizer7.Add( self.Bt8_TOFO, 0, wx.ALL, 5 )

        #self.Bt9_TOVO = wx.RadioButton( self, wx.ID_ANY, u"8. Sweep  : X-> To measure Temperature, Y-> To measure Voltage           ", wx.DefaultPosition, wx.DefaultSize, 0 )
        #bSizer7.Add( self.Bt9_TOVO, 0, wx.ALL, 5 )

        self.Bt10_EFFI = wx.RadioButton( self, wx.ID_ANY, u"9. Efficiency                                                            ", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer7.Add( self.Bt10_EFFI, 0, wx.ALL, 5 )


        self.SetSizer( bSizer7 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.Bt1_None.Bind( wx.EVT_RADIOBUTTON, self.Ev0_None )
        self.Bt2_USB_I2C.Bind( wx.EVT_RADIOBUTTON, self.Ev1_USB_I2C )
        self.Bt3_VICI.Bind( wx.EVT_RADIOBUTTON, self.Ev2_VICI )
        self.Bt4_VIVO.Bind( wx.EVT_RADIOBUTTON, self.Ev3_VIVO )
        self.Bt5_VIFO.Bind( wx.EVT_RADIOBUTTON, self.Ev4_VIFO )
        self.Bt6_COVO.Bind( wx.EVT_RADIOBUTTON, self.Ev5_COVO )
        self.Bt7_REGVO.Bind( wx.EVT_RADIOBUTTON, self.Ev6_REGVO )
        #self.Bt8_TOFO.Bind( wx.EVT_RADIOBUTTON, self.Ev7_TOFO )
        #self.Bt9_TOVO.Bind( wx.EVT_RADIOBUTTON, self.Ev8_TOVO )
        self.Bt10_EFFI.Bind( wx.EVT_RADIOBUTTON, self.Ev9_EFFI )

    def __del__( self ):
        pass

    # Virtual event handlers, overide them in your derived class
    def Ev0_None( self, event ):
        pass

    def Ev1_USB_I2C( self, event ):
        Op_1 = USB_I2C(None)
        Op_1.Show()
        self.Destroy()
        print "Open_1_USBI2C_FRAME"

    def Ev2_VICI( self, event ):
        Op_2 = Sweep_2_VICI(None, 'Input Voltage & Input Current', 'Input Voltage(V)', 'Input Current(A)', wx.Size( 920,450 ) )
        Op_2.Show()
        self.Destroy()
        print "Open_2_ViCo_FRAME"

    def Ev3_VIVO( self, event ):
        Op_3 = Sweep_3_VIVO(None, 'Input Voltage & Output Voltage', 'Input Voltage(V)', 'Output Voltage(V)', wx.Size( 920,450 ) )
        Op_3.Show()
        self.Destroy()
        print "Open_3_ViVo_FRAME"

    def Ev4_VIFO( self, event ):
        #wx.MessageBox('Sorry, Not Implement!', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        Op_4 = Sweep_4_VIFO(None, 'Input Voltage & Output Frequency', 'Sweep Input Voltage(V):', 'Record Output Frequency(Hz)', wx.Size( 1000,450 ) )
        Op_4.Show()
        self.Destroy()
        print "Open_4_ViFo_FRAME"

    def Ev5_COVO( self, event ):
        Op_5 = Sweep_5_COVO(None, 'Output Current & Output Voltage', 'Output Loading(A)', 'Output Voltage(V)', wx.Size( 940,450 ) )
        Op_5.Show()
        self.Destroy()
        print "Open_5_CoVo_FRAME"

    def Ev6_REGVO( self, event ):
        Op_6 = Sweep_6_REGVO(None, 'Register & Output Voltage', 'Register', 'Output Voltage(V)', wx.Size( 940,450 ) )
        Op_6.Show()
        self.Destroy()
        print "Open_6_RegVo_FRAME"

    def Ev7_TOFO( self, event ):
        """
                Op_7 = Sweep(None, 'Temperature & Output Frequency', 'Time Duration:', 'Record Temp(degree) and Output Frequency(Hz):')
                Op_7.Show()
                self.Destroy()
                print "Open_7_ToFo_FRAME"
                """
        wx.MessageBox('Sorry, Not Implement!', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

    def Ev8_TOVO( self, event ):
        wx.MessageBox('Sorry, Not Implement!', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

    def Ev9_EFFI( self, event ):
        Op_9 = Sweep_9_EFFI(None, 'Loading VS. Efficiency', 'Loading(A)', 'Efficiency(%)', wx.Size( 600,720 ))
        Op_9.Show()
        self.Destroy()
        print "Open_9_EFFI_FRAME"

class BoundControlBox(wx.Panel):
    """ A static box with a couple of radio buttons and a text
        box. Allows to switch between an automatic mode and a
        manual mode with an associated value.
    """
    def __init__(self, parent, ID, label, initval):
        wx.Panel.__init__(self, parent, ID)

        self.value = initval

        box = wx.StaticBox(self, -1, label)
        sizer = wx.StaticBoxSizer(box, wx.VERTICAL)

        self.radio_auto = wx.RadioButton(self, -1,
            label="Auto", style=wx.RB_GROUP)
        self.radio_manual = wx.RadioButton(self, -1,
            label="Manual")
        self.manual_text = wx.TextCtrl(self, -1,
            size=(35,-1),
            value=str(initval),
            style=wx.TE_PROCESS_ENTER)

        self.Bind(wx.EVT_UPDATE_UI, self.on_update_manual_text, self.manual_text)
        self.Bind(wx.EVT_TEXT_ENTER, self.on_text_enter, self.manual_text)

        manual_box = wx.BoxSizer(wx.HORIZONTAL)
        manual_box.Add(self.radio_manual, flag=wx.ALIGN_CENTER_VERTICAL)
        manual_box.Add(self.manual_text, flag=wx.ALIGN_CENTER_VERTICAL)

        sizer.Add(self.radio_auto, 0, wx.ALL, 10)
        sizer.Add(manual_box, 0, wx.ALL, 10)

        self.SetSizer(sizer)
        sizer.Fit(self)

    def on_update_manual_text(self, event):
        self.manual_text.Enable(self.radio_manual.GetValue())

    def on_text_enter(self, event):
        self.value = self.manual_text.GetValue()

    def is_auto(self):
        return self.radio_auto.GetValue()

    def manual_value(self):
        return self.value

    def __del__( self ):
        pass

###########################################################################
## Class Sweep
###########################################################################
class Sweep ( wx.Frame ):

    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        wx.Frame.__init__( self, parent, id = wx.ID_ANY, title = new_title, pos = wx.DefaultPosition, size = new_size, style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
        #favicon = wx.Icon(r'./icon/anpec.ico', wx.BITMAP_TYPE_ICO, 16, 16)
        #wx.Frame.SetIcon(self, favicon)
        self.Bind(wx.EVT_CLOSE, self.Ev_OnClose)
        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
        self.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_GRAYTEXT ) )
        self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INFOBK ) )

        self.Horizotal_Frame = wx.BoxSizer( wx.HORIZONTAL )

        self.Vertical_Frame_Left = wx.BoxSizer( wx.VERTICAL )
        self.SWEEP_FRAME( sweep_item )
        self.MESSAGE_FRAME(changed_item)
        self.Horizotal_Frame.Add(self.Vertical_Frame_Left, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_TOP, 5 )

        self.Vertical_Frame_Right = wx.BoxSizer( wx.VERTICAL )
        self.FIGURE_FRAME(new_title, sweep_item, changed_item)
        self.Horizotal_Frame.Add(self.Vertical_Frame_Right, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_TOP, 5 )

        self.SetSizer( self.Horizotal_Frame )
        self.Layout()
        self.Centre( wx.BOTH )


        self.currentDirectory = os.getcwd()

    def SWEEP_FRAME(self, sweep_item):
        Subframe_SWEEP = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Sweep" ), wx.VERTICAL )

        four_SWEEP = wx.GridSizer( 0, 2, 0, 0 )

        InputType = wx.BoxSizer( wx.VERTICAL )

        self.Label_Input = wx.StaticText( self, wx.ID_ANY, label=sweep_item, pos=(140, 40), style=0)
        self.Label_Input.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Input.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Label_Input.Wrap( -1 )

        InputType.AddSpacer(38)
        InputType.Add( self.Label_Input, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_CENTER_HORIZONTAL, 5 )


        four_SWEEP.Add( InputType, 1, wx.EXPAND, 5 )

        Range = wx.BoxSizer( wx.VERTICAL )

        bSizer5 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Start = wx.StaticText( self, wx.ID_ANY, u"Start", wx.DefaultPosition, wx.Size( 32,-1 ), 0 )
        self.Label_Start.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Start.Wrap( -1 )
        bSizer5.Add( self.Label_Start, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Stop = wx.StaticText( self, wx.ID_ANY, u"Stop", wx.DefaultPosition, wx.Size( 32,-1 ), 0 )
        self.Label_Stop.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Stop.Wrap( -1 )
        bSizer5.Add( self.Label_Stop, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Step = wx.StaticText( self, wx.ID_ANY, u"Step", wx.DefaultPosition, wx.Size( 32,-1 ), 0 )
        self.Label_Step.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Step.Wrap( -1 )

        bSizer5.Add( self.Label_Step, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        Range.Add( bSizer5, 1, wx.EXPAND, 5 )

        bSizer6 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_Start = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.Add( self.Text_Start, 0, wx.BOTTOM|wx.LEFT, 5 )

        self.Text_Stop = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.Add( self.Text_Stop, 0, wx.BOTTOM|wx.LEFT, 5 )

        self.Text_Step = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.Add( self.Text_Step, 0, wx.BOTTOM|wx.LEFT, 5 )


        Range.Add( bSizer6, 1, wx.EXPAND, 5 )


        four_SWEEP.Add( Range, 1, wx.EXPAND, 5 )

        bSizer13 = wx.BoxSizer( wx.VERTICAL )

        self.Bt_RETURN = wx.Button( self, wx.ID_ANY, u"Retrun", wx.Point( -1,-1 ), wx.DefaultSize, 0 )
        self.Bt_RETURN.SetBackgroundColour( wx.Colour( 0, 128, 192 ) )

        bSizer13.Add( self.Bt_RETURN, 0, wx.ALIGN_BOTTOM|wx.ALL, 5 )

        self.Bt_Check_Instrument = wx.Button( self, wx.ID_ANY, u"Connect", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_Check_Instrument.SetBackgroundColour( wx.Colour( 0, 128, 0 ) )
        bSizer13.Add( self.Bt_Check_Instrument, 0, wx.ALL, 5 )

        four_SWEEP.Add( bSizer13, 1, wx.EXPAND, 5 )

        self.Bt_RUN = wx.Button( self, wx.ID_ANY, u"Run", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_RUN.SetBackgroundColour( wx.Colour( 255, 128, 64 ) )

        four_SWEEP.Add( self.Bt_RUN, 0, wx.ALL|wx.ALIGN_BOTTOM|wx.ALIGN_RIGHT, 5 )


        Subframe_SWEEP.Add( four_SWEEP, 1, wx.EXPAND, 5 )


        self.Vertical_Frame_Left.Add( Subframe_SWEEP, 1, wx.EXPAND, 5 )

        # Connect Events
        self.Bt_RETURN.Bind( wx.EVT_BUTTON, self.Ev_RETURN )
        self.Bt_RUN.Bind( wx.EVT_BUTTON, self.Ev_RUN )
        self.Bt_Check_Instrument.Bind( wx.EVT_BUTTON, self.Ev_Check_Instrument )

    def MESSAGE_FRAME(self, changed_item):

        Subframe_MESSAGE = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Message" ), wx.VERTICAL )

        bSizer2 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Output = wx.StaticText( self, wx.ID_ANY, changed_item, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Label_Output.Wrap( -1 )
        self.Label_Output.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Label_Output.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INFOTEXT ) )
        bSizer2.Add( self.Label_Output, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Bt_SAVE_DATA = wx.Button( self, wx.ID_ANY, u"S-DATA", wx.Point( -1,-1 ), wx.DefaultSize, 0 )
        self.Bt_SAVE_DATA.SetBackgroundColour( wx.Colour( 255, 0, 128 ) )
        #bSizer2.AddSpacer(45)
        bSizer2.Add( self.Bt_SAVE_DATA, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Bt_SAVE_FIG = wx.Button( self, wx.ID_ANY, u"S-FIG", wx.Point( -1,-1 ), wx.DefaultSize, 0 )
        self.Bt_SAVE_FIG.SetBackgroundColour( wx.Colour( 128, 128, 192 ) )
        #bSizer2.AddSpacer(45)
        bSizer2.Add( self.Bt_SAVE_FIG, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        Subframe_MESSAGE.Add( bSizer2, 0, wx.ALL, 5 )

        self.Text_Output = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE )
        Subframe_MESSAGE.Add( self.Text_Output, 1, wx.ALL|wx.EXPAND, 5 )

        self.Vertical_Frame_Left.Add( Subframe_MESSAGE, 1, wx.EXPAND, 5 )

        # Connect Events
        self.Bt_SAVE_DATA.Bind( wx.EVT_BUTTON, self.Ev_SAVE_DATA )
        self.Bt_SAVE_FIG.Bind( wx.EVT_BUTTON, self.Ev_SAVE_FIG )

    def FIGURE_FRAME(self, new_title, sweep_item, changed_item):
        Subframe_FIGURE = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Figure" ), wx.VERTICAL )

        self.panel = wx.Panel(self)

        self.init_plot(new_title, sweep_item, changed_item)
        self.canvas = FigCanvas(self.panel, -1, self.fig)


        Subframe_FIGURE.Add( self.panel, 0, wx.ALL, 5 )
        self.Vertical_Frame_Right.Add( Subframe_FIGURE, 1, wx.EXPAND, 5 )

    def init_plot(self, new_title, sweep_item, changed_item):
        self.dpi = 100
        self.fig = Figure((5.5, 3.5), dpi=self.dpi)

        self.axes = self.fig.add_subplot(111)
        self.axes.set_axis_bgcolor('white')
        self.axes.set_title(new_title, size=10)
        self.axes.set_xlabel(sweep_item, size=8)
        self.axes.set_ylabel(changed_item, size=8)
        self.axes.grid(True, color='gray')

        pylab.setp(self.axes.get_xticklabels(), fontsize=6)
        pylab.setp(self.axes.get_yticklabels(), fontsize=6)

        # plot the data as a line series, and save the reference
        # to the plotted line series
        #

        self.plot_data = self.axes.plot(
            #self.data,
            [],
            linewidth=1,
            color=(1,1,0),
            )[0]

    # Virtual event handlers, overide them in your derived class
    def plot_draw(self, x_data, y_data, color, label):

        #xmin = float(self.Text_Start.Value)
        #xmax = float(self.Text_Stop.Value)
        xmin = float(min(x_data))
        xmax = float(max(x_data))
        ymin=float(min(y_data))
        ymax=float(max(y_data))

        self.axes.set_xbound(lower=0., upper=1.05*xmax)
        self.axes.set_ybound(lower=0., upper=1.05*ymax)
        self.axes.set_autoscale_on(False)
        #self.axes.set_ylim([0,110])

        pylab.setp(self.axes.get_xticklabels())
        if label==[]:
            self.plot_data = self.axes.plot(x_data, y_data, linewidth=1, color=color)
        else:
            self.plot_data = self.axes.plot(x_data, y_data, linewidth=1,color=color, label=label)
            legend = self.axes.legend(loc='lower right', shadow=True, fontsize=8)
            legend.get_frame().set_facecolor('0.75')

        self.canvas.draw()

    def Ev_RETURN( self, event ):
        TOP=Top_Option(None)
        TOP.Show()
        self.Destroy()
        print "Push Return"

    def Ev_RUN( self, event ):
        print "Push Run"

    def Ev_Check_Instrument( self, event ):
        self.Text_Output.AppendText(u"Checking Instruments... \n")
        self.RESOURCE = self.Check_Resource()
        print self.RESOURCE
        if self.RESOURCE <> {}:
            for inst in self.RESOURCE:
                if inst == "PRODIGI":
                    print "E-Load @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'E-Load @ COM'+self.RESOURCE[inst][4:5]+u'\n')
                elif inst == "KEI":
                    print "Keithely @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'Keithely @ COM'+self.RESOURCE[inst][4:5]+u'\n')
                elif inst == "HEWLE":
                    print "Agilent @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'Agilent @ COM'+self.RESOURCE[inst][4:5]+u'\n')
                elif inst == "USB":
                    print "USB I2C is Connecting @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'USB I2C @ '+self.RESOURCE[inst][0:3]+u'\n')
        else:
            self.Text_Output.AppendText(u" No Instruments !! \n")

    def Ev_OnClose( self, event ):
        print "push close!"
        self.Destroy()

    def Ev_SAVE_DATA(self, event ):
        dlg = wx.FileDialog(
            self, message="Save file as ...",
            defaultDir=self.currentDirectory,
            defaultFile="", wildcard=wildcard, style=wx.SAVE
            )
        if dlg.ShowModal() == wx.ID_OK:
            self.save_path = dlg.GetPath()
            print "You chose the following filename: %s" % self.save_path
        dlg.Destroy()
        self.Do_SAVE()

    def Do_SAVE(self):
        pass

    def Ev_SAVE_FIG(self, event):
        file_choices = "PNG (*.png)|*.png"

        dlg = wx.FileDialog(
            self,
            message="Save plot as...",
            defaultDir=os.getcwd(),
            defaultFile="plot.png",
            wildcard=file_choices,
            style=wx.SAVE)

        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.canvas.print_figure(path, dpi=self.dpi)
            #self.flash_status_message("Saved to %s" % path)
            self.Text_Output.AppendText("Saved to %s" % path)

    def Check_Resource(self):
        TOTAL_RESOURCE=defaultdict(dict)
        rm = visa.ResourceManager()
        NI_VISA_Resource1 = rm.list_resources("ASRL?*::INSTR")
        for Resource in NI_VISA_Resource1:
            #print Resource
            try:
                RS232 = rm.open_resource(Resource)
                RS232.Baud_rate = 9600
                RS232.write('*IDN?')
                name = RS232.read(termination="T")
                TOTAL_RESOURCE[name]=Resource
                RS232.clear() # very important, or Agilent E3632 will missing
                RS232.close()
            except:
                print "Initial instrument fail: ", Resource


        NI_VISA_Resource2 = rm.list_resources('USB?*0x0451::0xBB01?*')
        #print NI_VISA_Resource2

        try:
            USB = rm.open_resource(str(NI_VISA_Resource2[0]))
            USB.timeout = 2500
            USB.chunk_size = 256
            lib = rm.visalib
            status = lib.enable_event(USB.session, constants.VI_EVENT_USB_INTR, 1, 0)
            status = lib.usb_control_in(USB.session, 128, 6, 1023, 0, 12)
            status = lib.get_attribute(USB.session, constants.VI_ATTR_MODEL_NAME)
            print(status[0][0:3])
            print NI_VISA_Resource2
            TOTAL_RESOURCE[status[0][0:3]]=NI_VISA_Resource2[0]
            USB.close()
        except:
            pass

        try:
            scope = rm.open_resource("TCPIP0::"+self.Text_IP_Address.Value)
            name = scope.query('*IDN?')
            print "IP is " + self.Text_IP_Address.Value
            scope.close()
            TOTAL_RESOURCE[name[0:6]] = name
        except:
            pass
            #wx.MessageBox('C!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        rm.close()
        #print TOTAL_RESOURCE
        return TOTAL_RESOURCE

    def LeCroy_Init(self):
        self.LeCroy = win32com.client.Dispatch("LeCroy.ActiveDSOCtrl.1")
        self.LeCroy.MakeConnection("IP:" + self.Text_IP_Address.Label)
        self.LeCroy.WriteString("C1:TRA ON", 1)
        self.LeCroy.WriteString("ASET", 1)
        self.LeCroy.WriteString("VBS app.Measure.ShowMeasure = true", 1)
        self.LeCroy.WriteString("""VBS 'app.Measure.P1.View = true '  """, 1)
        self.LeCroy.WriteString("""VBS 'app.Measure.P1.Source1 ="C1" '  """, 1)
        self.LeCroy.WriteString("""VBS 'app.Measure.P1.ParamEngine="FREQ" '  """, 1)

    def Agilent_E3632A_Init(self, Resource):
        print Resource
        self.AGILENT_E3632 = self.rm.open_resource(Resource)
        self.AGILENT_E3632.Baud_rate = 9600
        self.AGILENT_E3632.Parity = constants.VI_ASRL_PAR_NONE
        self.AGILENT_E3632.Start_bits = 1
        self.AGILENT_E3632.Data_bits = 8
        self.AGILENT_E3632.Stop_bits = constants.VI_ASRL_STOP_TWO
        self.AGILENT_E3632.Bytes_in_buffer = 16
        self.AGILENT_E3632.write(u'*RST\r')
        #self.AGILENT_E3632.write(u'SYST:LOC\r')
        self.AGILENT_E3632.write(u'OUTP ON\r')
        time.sleep(0.5) # Very import delay 0.5-sec, can't not be removed
        self.AGILENT_E3632.write(u'SYST:REM\r')
        self.AGILENT_E3632.write(u'VOLTage:PROTection:STATe OFF\r')
        #self.AGILENT_E3632.write(u'CURR:PROT:STAT OFF')
        #self.AGILENT_E3632.write(u'SYSTem:BEEPer:STATe OFF')

    def Agilent_E3632A_Curr(self, volt):

        self.AGILENT_E3632.write(u'VOLT '+str(volt)+'\r')
        #self.AGILENT_E3632.clear() # Disable to avoid Beeper noise from agilent
        time.sleep(0.5)# Very import delay 0.5-sec, can't not be removed
        try:
            self.AGILENT_E3632.write("MEAS:CURR?\r")
            time.sleep(0.1)
            current = self.AGILENT_E3632.read(termination="\r")
            #self.AGILENT_E3632.clear()  # Disable to avoid Beeper noise from agilent
            print "Agilent set Voltage "+ str(volt)+ ", Measure Current: " + current
        except:
            print "Measure Current Error!"

        return current

    def Agilent_E3632A_Volt(self, volt):

        self.AGILENT_E3632.write(u'VOLT '+str(volt)+'\r')
        #self.AGILENT_E3632.clear()  # Disable to avoid Beeper noise from agilent
        time.sleep(0.5)# Very import delay 0.5-sec, can't not be removed

    def Agilent_E3632A_close(self):
        self.AGILENT_E3632.write(u'OUTP OFF\r')
        self.AGILENT_E3632.close()

    def Keithley_2700_Init(self, Resource):
        print Resource
        # COM port setting
        self.KEITHLEY_2700 = self.rm.open_resource(Resource)

        self.KEITHLEY_2700.Baud_rate = 9600
        self.KEITHLEY_2700.Parity = constants.VI_ASRL_PAR_NONE
        self.KEITHLEY_2700.Start_bits = 1
        self.KEITHLEY_2700.Data_bits = 8
        self.KEITHLEY_2700.Stop_bits = constants.VI_ASRL_STOP_ONE
        self.KEITHLEY_2700.Bytes_in_buffer = 16

        time.sleep(0.5)
        # Initial setting
        self.KEITHLEY_2700.write(u'*RST')
        #self.KEITHLEY_2700.write(u'*rst; status:preset; *cls')
        self.KEITHLEY_2700.write(u'SYSTem:BEEPer 0')
        self.KEITHLEY_2700.write(u'FORMat:ELEMents READing')
        self.KEITHLEY_2700.write(u'INIT:CONT ON\r')
        self.KEITHLEY_2700.write(u'TRIG:COUN 1\r')
        self.KEITHLEY_2700.write("SENS:FUNC 'VOLT:DC'\r")
        self.KEITHLEY_2700.write(u'VOLT:DC:RANG 10\r')

        #self.KEITHLEY_2700.write("OUTPUT OFF;:INIT")

    def Keithley_2700_close(self):
        self.KEITHLEY_2700.write(u'OUTP OFF\r')
        self.KEITHLEY_2700.close()

    def Prodigit_3311C_Init(self, Resource):
        print Resource
        # COM port setting
        self.PRODIGIT_3311C = self.rm.open_resource(Resource)
        self.PRODIGIT_3311C.Baud_rate = 9600
        self.PRODIGIT_3311C.Parity = constants.VI_ASRL_PAR_NONE
        self.PRODIGIT_3311C.Start_bits = 1
        self.PRODIGIT_3311C.Data_bits = 8
        self.PRODIGIT_3311C.Stop_bits = constants.VI_ASRL_STOP_ONE
        self.PRODIGIT_3311C.Bytes_in_buffer = 16
        # Initial setting
        self.PRODIGIT_3311C.write(u'STATE:REMO\r')
        self.PRODIGIT_3311C.write(u'STATE:LOCAL\r')
        self.PRODIGIT_3311C.write(u'STATE:MODE CC;\r')
        self.PRODIGIT_3311C.write(u'STATE:PRES OFF;\r')
        self.PRODIGIT_3311C.write(u'STATE:SENSE OFF;\r')
        self.PRODIGIT_3311C.write(u'STATE:LEVEL HIGH;\r')
        self.PRODIGIT_3311C.write(u'STATE:DYN OFF;\r')
        self.PRODIGIT_3311C.write(u'STATE:LOAD ON;\r')

    def Prodigit_3311C_close(self):
        self.PRODIGIT_3311C.write(u'STATE:LOAD OFF;\r')
        self.PRODIGIT_3311C.close()

    def USB_I2C_Init(self, Resource):
        # USB port setting
        #USB_DEVICE = rm.list_resources('USB?*0x0451::0xBB01?*')
        print Resource
        self.USB_I2C = self.rm.open_resource(Resource)
        self.USB_I2C.timeout = 2500
        self.USB_I2C.chunk_size = 256
        self.lib = self.rm.visalib

    def USB_I2C_close(self):
        self.USB_I2C.close()

class Sweep_2_VICI( Sweep ):
    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        super(Sweep_2_VICI, self).__init__( parent, new_title, sweep_item, changed_item, new_size )
        self.current=[]

    def Ev_RUN( self, event ):
        print "Push Run"
        self.current=[]
        try:
            Resource_AGILENT = self.RESOURCE["HEWLE"]
            self.rm = visa.ResourceManager()
            self.Agilent_E3632A_Init(Resource_AGILENT)
        except:
            wx.MessageBox('Agilent E3236A does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            if (float(self.Text_Start.Value) == float(self.Text_Stop.Value)) or (self.Text_Step.Value == ''):
                self.volt = [float(self.Text_Start.Value)]
            else:
                self.volt = np.arange(float(self.Text_Start.Value),
                                  float(self.Text_Stop.Value),
                                  float(self.Text_Step.Value))

        except:
            wx.MessageBox('Please give All Data for Sweep!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            for volt in self.volt:
                curr = self.Agilent_E3632A_Curr(volt)
                msg = u'Set Voltage:' + str(volt) + u', Get Current:' + str(curr) + u'\n'
                self.Text_Output.AppendText(msg)
                self.current.append(curr)
        except:
            wx.MessageBox('Data measurement Error!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        self.Agilent_E3632A_close()
        self.rm.close()

        print self.current
        color = (random.random(), random.random(), random.random())
        self.plot_draw(self.volt, self.current, color, [])


    def Do_SAVE(self):
        # Save File
        data = np.array([self.volt, self.current], dtype=np.float)
        np.savetxt(self.save_path, np.transpose(data), delimiter=", ")

class Sweep_3_VIVO( Sweep ):
    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        super(Sweep_3_VIVO, self).__init__( parent, new_title, sweep_item, changed_item, new_size )

    def Ev_RUN( self, event ):
        print "Push Run"
        self.volt_OUT=[]
        # Check Instrument
        try:
            Resource_AGILENT = self.RESOURCE["HEWLE"]
            self.rm = visa.ResourceManager()
            self.Agilent_E3632A_Init(Resource_AGILENT)
        except:
            wx.MessageBox('Agilent E3236A does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        # Check Instrument
        try:
            Resource_KE2700 = self.RESOURCE["KEI"]
            #self.rm = visa.ResourceManager()
            self.Keithley_2700_Init(Resource_KE2700)
        except:
            wx.MessageBox('Keithley does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            if (float(self.Text_Start.Value) == float(self.Text_Stop.Value)) or self.Text_Step.Value =='':
                self.volt_IN=[float(self.Text_Start.Value)]
            else:
                self.volt_IN = np.arange(float(self.Text_Start.Value),
                                         float(self.Text_Stop.Value),
                                         float(self.Text_Step.Value))
        except:
            wx.MessageBox('Please give All Data for Sweep!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)


        try:
            xmin = float(self.Text_Start.Value)
            xmax = float(self.Text_Stop.Value)
            for volt in self.volt_IN:
                self.Agilent_E3632A_Volt(volt)
                try:
                    self.KEITHLEY_2700.write(":READ?")
                    time.sleep(0.1)
                    temp_Volt = self.KEITHLEY_2700.read(termination="\r")
                    #self.KEITHLEY_2700.clear()
                    self.volt_OUT.append(temp_Volt)
                    msg = u'Set Voltage:' + str(volt) + u', Get Voltage:' + str(temp_Volt) + u'\n'
                    self.Text_Output.AppendText(msg)
                    #self.plot_draw(self.volt_OUT)
                except:
                    self.Text_Output.AppendText("Measure Current Error!")
        except:
            wx.MessageBox('Data measurement Error!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)


        self.Agilent_E3632A_close()
        self.Keithley_2700_close()
        self.rm.close()

        #print self.volt_OUT
        color = (random.random(), random.random(), random.random())
        self.plot_draw(self.volt_IN, self.volt_OUT, color, [])

    def Do_SAVE(self):
        # Save File
        data = np.array([self.volt_IN, self.volt_OUT], dtype=np.float)
        np.savetxt(self.save_path, np.transpose(data), delimiter=", ")

class Sweep_4_VIFO( Sweep ):
    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        super(Sweep_4_VIFO, self).__init__( parent, new_title, sweep_item, changed_item, new_size )

    def SWEEP_FRAME(self, sweep_item):
        Subframe_SWEEP = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Sweep" ), wx.VERTICAL )

        four_SWEEP = wx.GridSizer( 0, 2, 0, 0 )

        InputType = wx.BoxSizer( wx.VERTICAL )

        self.Label_Input = wx.StaticText( self, wx.ID_ANY, label=sweep_item, pos=(140, 40), style=0)
        self.Label_Input.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Input.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Label_Input.Wrap( -1 )

        #InputType.AddSpacer(10)
        InputType.Add( self.Label_Input, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Label_Scope_IP = wx.StaticText( self, wx.ID_ANY, label='LeCrory Scope IP Address:', pos=(140, 40), style=0)
        self.Label_Scope_IP.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Scope_IP.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Label_Scope_IP.Wrap( -1 )

        InputType.AddSpacer(38)
        InputType.Add( self.Label_Scope_IP, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        four_SWEEP.Add( InputType, 1, wx.EXPAND, 5 )

        Range = wx.BoxSizer( wx.VERTICAL )

        bSizer5 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Start = wx.StaticText( self, wx.ID_ANY, u"Start", wx.DefaultPosition, wx.Size( 32,-1 ), 0 )
        self.Label_Start.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Start.Wrap( -1 )
        bSizer5.Add( self.Label_Start, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Stop = wx.StaticText( self, wx.ID_ANY, u"Stop", wx.DefaultPosition, wx.Size( 32,-1 ), 0 )
        self.Label_Stop.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Stop.Wrap( -1 )
        bSizer5.Add( self.Label_Stop, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Step = wx.StaticText( self, wx.ID_ANY, u"Step", wx.DefaultPosition, wx.Size( 32,-1 ), 0 )
        self.Label_Step.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Step.Wrap( -1 )

        bSizer5.Add( self.Label_Step, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        Range.Add( bSizer5, 1, wx.EXPAND, 5 )

        bSizer6 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_Start = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.Add( self.Text_Start, 0, wx.BOTTOM|wx.LEFT, 5 )

        self.Text_Stop = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.Add( self.Text_Stop, 0, wx.BOTTOM|wx.LEFT, 5 )

        self.Text_Step = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.Add( self.Text_Step, 0, wx.BOTTOM|wx.LEFT, 5 )

        Range.Add( bSizer6, 1, wx.EXPAND, 5 )

        bSizer6_1 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_IP_Address = wx.TextCtrl( self, wx.ID_ANY, '169.254.203.86', wx.DefaultPosition, wx.Size( 120,24 ), 0 )
        bSizer6_1.Add( self.Text_IP_Address, 0, wx.BOTTOM|wx.LEFT, 5 )
        Range.Add( bSizer6_1, 1, wx.EXPAND, 5 )

        four_SWEEP.Add( Range, 1, wx.EXPAND, 5 )

        bSizer13 = wx.BoxSizer( wx.VERTICAL )

        self.Bt_RETURN = wx.Button( self, wx.ID_ANY, u"Retrun", wx.Point( -1,-1 ), wx.DefaultSize, 0 )
        self.Bt_RETURN.SetBackgroundColour( wx.Colour( 0, 128, 192 ) )

        bSizer13.Add( self.Bt_RETURN, 0, wx.ALIGN_BOTTOM|wx.ALL, 5 )

        self.Bt_Check_Instrument = wx.Button( self, wx.ID_ANY, u"Connect", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_Check_Instrument.SetBackgroundColour( wx.Colour( 0, 128, 0 ) )
        bSizer13.Add( self.Bt_Check_Instrument, 0, wx.ALL, 5 )

        four_SWEEP.Add( bSizer13, 1, wx.EXPAND, 5 )

        self.Bt_RUN = wx.Button( self, wx.ID_ANY, u"Run", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_RUN.SetBackgroundColour( wx.Colour( 255, 128, 64 ) )

        four_SWEEP.Add( self.Bt_RUN, 0, wx.ALL|wx.ALIGN_BOTTOM|wx.ALIGN_RIGHT, 5 )


        Subframe_SWEEP.Add( four_SWEEP, 1, wx.EXPAND, 5 )


        self.Vertical_Frame_Left.Add( Subframe_SWEEP, 1, wx.EXPAND, 5 )

        # Connect Events
        self.Bt_RETURN.Bind( wx.EVT_BUTTON, self.Ev_RETURN )
        self.Bt_RUN.Bind( wx.EVT_BUTTON, self.Ev_RUN )
        self.Bt_Check_Instrument.Bind( wx.EVT_BUTTON, self.Ev_Check_Instrument )

    def Ev_RUN( self, event ):
        print "Push Run"
        self.FREQ=[]
        # Check Instrument
        try:
            Resource_AGILENT = self.RESOURCE["HEWLE"]
            self.rm = visa.ResourceManager()
            self.Agilent_E3632A_Init(Resource_AGILENT)
        except:
            wx.MessageBox('Agilent E3236A does not connect well.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        # Check Instrument
        try:
            self.RESOURCE["LECROY"]
            self.LeCroy_Init()
            print "do Lecroy Initial ..."
        except:
            wx.MessageBox('LeCroy does not connect well.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)


        try:
            if (float(self.Text_Start.Value) == float(self.Text_Stop.Value)) or self.Text_Start.Value =='':
                self.volt_IN = [float(self.Text_Start.Value)]
            else:
                self.volt_IN = np.arange(float(self.Text_Start.Value),
                                         float(self.Text_Stop.Value),
                                         float(self.Text_Step.Value))
        except:
            wx.MessageBox('Please give All Data for Sweep!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)


        try:
            xmin = float(self.Text_Start.Value)
            xmax = float(self.Text_Stop.Value)
            for volt in self.volt_IN:
                self.Agilent_E3632A_Volt(volt)
                try:
                    self.LeCroy.WriteString("""VBS? 'return=app.Measure.P1.Out.Result.Value'  """, 1)
                    #time.sleep(0.1)
                    temp_Freq = self.LeCroy.ReadString(80)
                    #print "The frequency is: " + str(temp_Freq)
                    self.FREQ.append(temp_Freq)
                    msg = u'Set Voltage:' + str(volt) + u', Get Freq:' + str(temp_Freq) + u'\n'
                    self.Text_Output.AppendText(msg)
                    #self.plot_draw(self.volt_OUT)
                except:
                    self.Text_Output.AppendText("Measure Frequenncy Error!")
        except:
            wx.MessageBox('Data measurement Error!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)


        self.Agilent_E3632A_close()
        self.LeCroy.Disconnect()
        self.rm.close()


        color = (random.random(), random.random(), random.random())
        self.plot_draw(self.volt_IN, self.FREQ, color, [])

    def Ev_Check_Instrument( self, event ):

        self.Text_Output.AppendText(u"Checking Instruments... \n")
        self.RESOURCE = self.Check_Resource()
        print self.RESOURCE
        if self.RESOURCE <> {}:
            for inst in self.RESOURCE:
                if inst == "PRODIGI":
                    print "E-Load @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'E-Load @ COM'+self.RESOURCE[inst][4:5]+u'\n')
                elif inst == "KEI":
                    print "Keithely @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'Keithely @ COM'+self.RESOURCE[inst][4:5]+u'\n')
                elif inst == "HEWLE":
                    print "Agilent @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'Agilent @ COM'+self.RESOURCE[inst][4:5]+u'\n')
                elif inst == "USB":
                    print "USB I2C is Connecting @", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'USB I2C @ '+self.RESOURCE[inst][0:3]+u'\n')
                elif inst == "LECROY":
                    print "LeCroy is Connecting", self.RESOURCE[inst]
                    self.Text_Output.AppendText(u'Connecting:'+self.RESOURCE[inst]+u'\n')
        else:
            self.Text_Output.AppendText(u" No Instruments !! \n")

        #self.Ev_Check_Scope()

    """
    def Ev_Check_Scope(self):
        LeCroy = win32com.client.Dispatch("LeCroy.ActiveDSOCtrl.1")
        print "IP:" + self.Text_IP_Address.Label
        print LeCroy.MakeConnection("IP:" + self.Text_IP_Address.Label)
        print LeCroy.WriteString("C1:VDIV 0.5", 1)
        LeCroy.Disconnect()
    """

    def Do_SAVE(self):
        # Save File
        data = np.array([self.volt_IN, self.volt_OUT], dtype=np.float)
        np.savetxt(self.save_path, np.transpose(data), delimiter=", ")

class Sweep_5_COVO( Sweep ):
    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        super(Sweep_5_COVO, self).__init__( parent, new_title, sweep_item, changed_item, new_size )

    def Ev_RUN( self, event ):
        print "Push Run"
        self.volt_OUT=[]
        try:
            Resource_PRODIGIT = self.RESOURCE["PRODIGI"]
            self.rm = visa.ResourceManager()
            self.Prodigit_3311C_Init(Resource_PRODIGIT)
        except:
            wx.MessageBox('Prodigit 3311C does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        try:
            Resource_KE2700 = self.RESOURCE["KEI"]
            self.Keithley_2700_Init(Resource_KE2700)
        except:
           wx.MessageBox('Keithley does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            self.curr_OUT = np.arange(float(self.Text_Start.Value),
                                     float(self.Text_Stop.Value),
                                     float(self.Text_Step.Value))
        except:
            wx.MessageBox('Please give All Data for Sweep!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            for curr in self.curr_OUT:
                self.PRODIGIT_3311C.write(u'CURR:HIGH '+ str(curr) + u'\r')
                self.PRODIGIT_3311C.write(u'CURR:LOW '+ str(curr) + u'\r')
                try:
                    self.KEITHLEY_2700.write("READ?")
                    time.sleep(0.3)
                    temp_volt = self.KEITHLEY_2700.read(termination="\r")
                    #self.KEITHLEY_2700.clear()
                    self.volt_OUT.append(temp_volt)
                    msg = u'Set Load:' + str(curr) + u', Get Voltage:' + str(temp_volt) + u'\n'
                    self.Text_Output.AppendText(msg)
                except:
                    self.Text_Output.AppendText("Measure Current Error!")
        except:
            wx.MessageBox('Data measurement Error!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        self.Prodigit_3311C_close()
        self.Keithley_2700_close()
        self.rm.close()

        #print self.volt_OUT
        color = (random.random(), random.random(), random.random())
        self.plot_draw(self.curr_OUT, self.volt_OUT, color, [])

    def Do_SAVE(self):
        # Save File
        data = np.array([self.curr_OUT, self.volt_OUT], dtype=np.float)
        np.savetxt(self.save_path, np.transpose(data), delimiter=", ")

class Sweep_6_REGVO( Sweep, WndProcHookMixin ):
    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        super(Sweep_6_REGVO, self).__init__( parent, new_title, sweep_item, changed_item, new_size )
        self.windows_notification()

    def SWEEP_FRAME(self, sweep_item):
        Subframe_SWEEP = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Sweep" ), wx.VERTICAL )


        TWO_SIDE = wx.BoxSizer( wx.HORIZONTAL )

        InputType = wx.BoxSizer( wx.VERTICAL )

        self.Device_Check = wx.StaticText( self, wx.ID_ANY, label="Push Reset Button!", pos=(5, 40), style=0)
        self.Device_Check.SetForegroundColour( wx.Colour( 0, 0, 255 ) )
        self.Device_Check.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Device_Check.Wrap( -1 )
        InputType.Add( self.Device_Check, 0, wx.ALIGN_BOTTOM|wx.ALL, 5 )

        self.Bt_RETURN = wx.Button( self, wx.ID_ANY, label=u"Retrun", pos=(5, 120), style=0 )
        self.Bt_RETURN.SetBackgroundColour( wx.Colour( 0, 128, 192 ) )
        InputType.AddSpacer(20)
        InputType.Add( self.Bt_RETURN, 0, wx.ALIGN_BOTTOM|wx.ALL, 5 )

        self.Bt_Check_Instrument = wx.Button( self, wx.ID_ANY, u"Connect", pos=(5, 160), style=0 )
        self.Bt_Check_Instrument.SetBackgroundColour( wx.Colour( 0, 128, 0 ) )
        #InputType.AddSpacer()
        InputType.Add( self.Bt_Check_Instrument, 0, wx.ALL, 5 )

        self.Bt_RUN = wx.Button( self, wx.ID_ANY, u"Run", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_RUN.SetBackgroundColour( wx.Colour( 255, 128, 64 ) )
        InputType.Add( self.Bt_RUN, 0, wx.ALL, 5 )

        #------------------------------------------------------------------------------------

        TWO_SIDE.Add( InputType, 1, wx.EXPAND, 5 )

        Range = wx.BoxSizer( wx.VERTICAL )
        #------------------------------------------------------------------------------------
        bSizer5 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Device_ID = wx.StaticText( self, wx.ID_ANY, u"Device ID", (5, 160), wx.Size( 50,-1 ), 0 )
        self.Label_Device_ID.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Device_ID.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Device_ID.Wrap( -1 )
        bSizer5.Add( self.Label_Device_ID, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Reg_Address = wx.StaticText( self, wx.ID_ANY, u"Reg. Addr", (5, 160), wx.Size( 50,-1 ), 0 )
        self.Label_Reg_Address.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Reg_Address.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Reg_Address.Wrap( -1 )
        bSizer5.Add( self.Label_Reg_Address, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )
        Range.Add( bSizer5, 1, wx.TOP, 5 )

        #------------------------------------------------------------------------------------
        bSizer7 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_Device_ID = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, size=( 24,24 ))
        bSizer7.AddSpacer(15)
        bSizer7.Add( self.Text_Device_ID, 0, wx.BOTTOM|wx.LEFT, 5 )

        self.Text_Reg_Address = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer7.AddSpacer(25)
        bSizer7.Add( self.Text_Reg_Address, 0, wx.BOTTOM|wx.LEFT, 5 )

        Range.Add( bSizer7, 1, wx.TOP, 5 )

        #------------------------------------------------------------------------------------
        bSizer9 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Reg_Start = wx.StaticText( self, wx.ID_ANY, u"Reg. Start", wx.DefaultPosition, wx.Size( 50,-1 ), 0 )
        self.Label_Reg_Start.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Reg_Start.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Reg_Start.Wrap( -1 )
        bSizer9.Add( self.Label_Reg_Start, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Reg_Stop = wx.StaticText( self, wx.ID_ANY, u"Reg. Stop", wx.DefaultPosition, wx.Size( 50,-1 ), 0 )
        self.Label_Reg_Stop.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Reg_Stop.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Reg_Stop.Wrap( -1 )
        bSizer9.Add( self.Label_Reg_Stop, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        Range.Add( bSizer9, 1, wx.EXPAND, 5 )
        #------------------------------------------------------------------------------------
        bSizer6 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_Reg_Start = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer6.AddSpacer(15)
        bSizer6.Add( self.Text_Reg_Start, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Text_Reg_Stop = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer6.AddSpacer(25)
        bSizer6.Add( self.Text_Reg_Stop, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        Range.Add( bSizer6, 1, wx.EXPAND, 5 )

        #------------------------------------------------------------------------------------

        TWO_SIDE.Add( Range, 1, wx.EXPAND, 5 )

        Subframe_SWEEP.Add(TWO_SIDE, 1, wx.EXPAND, 5 )

        self.Vertical_Frame_Left.Add( Subframe_SWEEP, 1, wx.EXPAND, 5 )

        # Connect Events
        self.Bt_RETURN.Bind( wx.EVT_BUTTON, self.Ev_RETURN )
        self.Bt_RUN.Bind( wx.EVT_BUTTON, self.Ev_RUN )
        self.Bt_Check_Instrument.Bind( wx.EVT_BUTTON, self.Ev_Check_Instrument )

    def Ev_RUN( self, event ):
        print "Push Run"
        self.volt_OUT=[]
        data_out = c_uint8 * 12

        try:
            Resource_KE2700 = self.RESOURCE["KEI"]
            self.rm = visa.ResourceManager()
            self.Keithley_2700_Init(Resource_KE2700)
        except:
           wx.MessageBox('Keithley does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            Resource_USB = self.RESOURCE["USB"]
            self.USB_I2C_Init(Resource_USB)
        except:
           wx.MessageBox('USB does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        try:
            self.REG = np.arange(int(self.Text_Reg_Start.Value, 16),
                                 int(self.Text_Reg_Stop.Value, 16))
        except:
            wx.MessageBox('Please give All Data for Sweep!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            for reg in self.REG:
                ii = data_out(0x12, int(self.Text_Device_ID.Value,16), 0x01, int(self.Text_Reg_Address.Value, 16), reg) # """  Write """
                self.lib.enable_event(self.USB_I2C.session, constants.VI_EVENT_USB_INTR, 1, 0)
                self.lib.usb_control_out(self.USB_I2C.session, 33, 9, 0, 3, ii)
                try:
                    self.KEITHLEY_2700.write(u':READ?')
                    time.sleep(0.1)
                    temp_volt  = self.KEITHLEY_2700.read(termination="\r")
                    self.volt_OUT.append(temp_volt)
                    msg = u'Set Reg:' + hex(reg) + u', Get Voltage:' + str(temp_volt) + u'\n'
                    self.Text_Output.AppendText(msg)
                except:
                    self.Text_Output.AppendText("Measure Current Error!")
        except:
            wx.MessageBox('Data measurement Error!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        self.Keithley_2700_close()
        self.USB_I2C_close()
        self.rm.close()

        #print self.volt_OUT
        color = (random.random(), random.random(), random.random())
        self.plot_draw(self.REG, self.volt_OUT, color, [])

    def Do_SAVE(self):
        # Save File
        data = np.array([self.REG, self.volt_OUT], dtype=np.float)
        np.savetxt(self.save_path, np.transpose(data), delimiter=", ")

    def windows_notification(self):
        """ Windows Notification """
        WndProcHookMixin.__init__(self)
        self.Bind(wx.EVT_CLOSE, self.onClose)
        """Change the following guid to the GUID of the device you want notifications for"""
        self.devNotifyHandle = self.registerDeviceNotification(guid="{A5DCBF10-6530-11D2-901F-00C04FB951ED}")
        self.addMsgHandler(WM_DEVICECHANGE, self.onDeviceChange)
        self.hookWndProc()
        ################################################################################################################

    def onDeviceChange(self,wParam,lParam):
        if lParam:
            dbh = DEV_BROADCAST_HDR.from_address(lParam)
            if dbh.dbch_devicetype == DBT_DEVTYP_DEVICEINTERFACE:
                dbd = DEV_BROADCAST_DEVICEINTERFACE.from_address(lParam)
                #Verify that the USB VID and PID match our assigned VID and PID
                if 'VID_0451&PID_BB01' in dbd.dbcc_name:
                    if wParam == DBT_DEVICEARRIVAL:
                        self.Device_Check.Label='Device Exit!'
                        self.Device_Check.SetForegroundColour( wx.Colour( 0, 255, 0 ) )
                    elif wParam == DBT_DEVICEREMOVECOMPLETE:
                        self.Device_Check.Label='Device Not Exit!'
                        self.Device_Check.SetForegroundColour( wx.Colour( 255, 0, 0 ) )
        return True

    def onClose(self, event):
        self.unregisterDeviceNotification(self.devNotifyHandle)
        event.Skip()

class Sweep_9_EFFI( Sweep ):
    def __init__( self, parent, new_title, sweep_item, changed_item, new_size ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = new_title, pos = wx.DefaultPosition, size = new_size, style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
        #favicon = wx.Icon(r'./icon/anpec.ico', wx.BITMAP_TYPE_ICO, 16, 16)
        #wx.Frame.SetIcon(self, favicon)
        self.Bind(wx.EVT_CLOSE, self.Ev_OnClose)
        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
        self.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_GRAYTEXT ) )
        self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INFOBK ) )

        #self.Horizotal_Frame = wx.BoxSizer( wx.HORIZONTAL )

        self.Vertical_Frame_Left = wx.BoxSizer( wx.VERTICAL )
        self.SWEEP_FRAME(sweep_item)
        self.MESSAGE_FRAME(new_title, sweep_item, changed_item)
        #self.Horizotal_Frame.Add(self.Vertical_Frame_Left, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        #self.Vertical_Frame_Right = wx.BoxSizer( wx.VERTICAL )
        #self.create_main_panel()
        #self.Horizotal_Frame.Add(self.Vertical_Frame_Right, 0, wx.EXPAND, 5 )

        self.SetSizer( self.Vertical_Frame_Left )
        self.Layout()
        self.Centre( wx.BOTH )

        self.currentDirectory = os.getcwd()

    def SWEEP_FRAME(self, sweep_item):
        Subframe_SWEEP = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Sweep" ), wx.VERTICAL )

        TWO_SIDE = wx.BoxSizer( wx.HORIZONTAL )

        InputType = wx.BoxSizer( wx.VERTICAL )

        self.Label_Input = wx.StaticText( self, wx.ID_ANY, label=sweep_item, pos=(5, 80), style=0)
        self.Label_Input.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Input.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Label_Input.Wrap( -1 )

        InputType.AddSpacer(5)
        InputType.Add( self.Label_Input, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Bt_RETURN = wx.Button( self, wx.ID_ANY, label=u"Retrun", pos=(5, 120), style=0 )
        self.Bt_RETURN.SetBackgroundColour( wx.Colour( 0, 128, 192 ) )
        InputType.Add( self.Bt_RETURN, 0, wx.ALIGN_BOTTOM|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Bt_Check_Instrument = wx.Button( self, wx.ID_ANY, u"Connect", pos=(5, 160), style=0 )
        self.Bt_Check_Instrument.SetBackgroundColour( wx.Colour( 0, 128, 0 ) )
        InputType.Add( self.Bt_Check_Instrument, 0, wx.ALIGN_BOTTOM|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Bt_RUN = wx.Button( self, wx.ID_ANY, u"Run",pos=(5, 200), style=0 )
        self.Bt_RUN.SetBackgroundColour( wx.Colour( 255, 128, 64 ) )
        InputType.Add( self.Bt_RUN, 0, wx.ALIGN_BOTTOM|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
        #------------------------------------------------------------------------------------

        TWO_SIDE.Add( InputType, 1, wx.EXPAND, 5 )

        SAVE = wx.BoxSizer( wx.VERTICAL )
        SAVE.AddSpacer(65)
        self.Bt_SAVE_DATA = wx.Button( self, wx.ID_ANY, u"S-DATA",pos=(30, 160), style=0 )
        self.Bt_SAVE_DATA.SetBackgroundColour( wx.Colour( 255, 0, 128 ) )
        SAVE.Add( self.Bt_SAVE_DATA, 0, wx.ALIGN_BOTTOM|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Bt_SAVE_FIG = wx.Button( self, wx.ID_ANY, u"S-FIG", pos=(30, 200), style=0 )
        self.Bt_SAVE_FIG.SetBackgroundColour( wx.Colour( 128, 128, 192 ) )
        #bSizer2.AddSpacer(45)
        SAVE.Add( self.Bt_SAVE_FIG, 0, wx.ALIGN_BOTTOM|wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        TWO_SIDE.Add( SAVE, 1, wx.EXPAND, 5 )

        Range = wx.BoxSizer( wx.VERTICAL )
        #------------------------------------------------------------------------------------
        bSizer5 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Volt_Start = wx.StaticText( self, wx.ID_ANY, u"Volt Start", (5, 160), wx.Size( 50,-1 ), 0 )
        self.Label_Volt_Start.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Volt_Start.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Volt_Start.Wrap( -1 )
        bSizer5.Add( self.Label_Volt_Start, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Volt_Stop = wx.StaticText( self, wx.ID_ANY, u"Volt Stop", (5, 160), wx.Size( 50,-1 ), 0 )
        self.Label_Volt_Stop.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Volt_Stop.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Volt_Stop.Wrap( -1 )
        bSizer5.Add( self.Label_Volt_Stop, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Volt_Step = wx.StaticText( self, wx.ID_ANY, u"Volt Step", (5, 160), wx.Size( 50,-1 ), 0 )
        self.Label_Volt_Step.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Volt_Step.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Volt_Step.Wrap( -1 )
        bSizer5.Add( self.Label_Volt_Step, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )
        Range.Add( bSizer5, 1, wx.EXPAND, 5 )
        #------------------------------------------------------------------------------------
        bSizer7 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_Volt_Start = wx.TextCtrl( self, wx.ID_ANY, '3.3', wx.DefaultPosition, size=( 24,24 ))
        bSizer7.AddSpacer(15)
        bSizer7.Add( self.Text_Volt_Start, 0, wx.BOTTOM|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Text_Volt_Stop = wx.TextCtrl( self, wx.ID_ANY, '3.4', wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer7.AddSpacer(25)
        bSizer7.Add( self.Text_Volt_Stop, 0, wx.BOTTOM|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Text_Volt_Step = wx.TextCtrl( self, wx.ID_ANY, '0.1', wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer7.AddSpacer(25)
        bSizer7.Add( self.Text_Volt_Step, 0, wx.BOTTOM|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        Range.Add( bSizer7, 1, wx.EXPAND, 5 )

        #------------------------------------------------------------------------------------
        bSizer9 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Curr_Start = wx.StaticText( self, wx.ID_ANY, u"Curr Start", wx.DefaultPosition, wx.Size( 50,-1 ), 0 )
        self.Label_Curr_Start.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Curr_Start.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Curr_Start.Wrap( -1 )
        bSizer9.Add( self.Label_Curr_Start, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Curr_Stop = wx.StaticText( self, wx.ID_ANY, u"Curr Stop", wx.DefaultPosition, wx.Size( 50,-1 ), 0 )
        self.Label_Curr_Stop.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Curr_Stop.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Curr_Stop.Wrap( -1 )
        bSizer9.Add( self.Label_Curr_Stop, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Curr_Step_uA = wx.StaticText( self, wx.ID_ANY, u"0.01~0.1", wx.DefaultPosition, wx.Size( 45,-1 ), 0 )
        self.Label_Curr_Step_uA.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Curr_Step_uA.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Curr_Step_uA.Wrap( -1 )
        bSizer9.Add( self.Label_Curr_Step_uA, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Curr_Step_mA = wx.StaticText( self, wx.ID_ANY, u"0.1~1", wx.DefaultPosition, wx.Size( 45,-1 ), 0 )
        self.Label_Curr_Step_mA.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Curr_Step_mA.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Curr_Step_mA.Wrap( -1 )
        bSizer9.Add( self.Label_Curr_Step_mA, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Curr_Step_A = wx.StaticText( self, wx.ID_ANY, u"1~10", wx.DefaultPosition, wx.Size( 40,-1 ), 0 )
        self.Label_Curr_Step_A.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Curr_Step_A.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Curr_Step_A.Wrap( -1 )
        bSizer9.Add( self.Label_Curr_Step_A, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        self.Label_Curr_Step_10A = wx.StaticText( self, wx.ID_ANY, u"10~100", wx.DefaultPosition, wx.Size( 40,-1 ), 0 )
        self.Label_Curr_Step_10A.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )
        self.Label_Curr_Step_10A.SetFont( wx.Font( 9, 72, 90, 90, False, wx.EmptyString ) )
        self.Label_Curr_Step_10A.Wrap( -1 )
        bSizer9.Add( self.Label_Curr_Step_10A, 0, wx.TOP|wx.RIGHT|wx.LEFT|wx.ALIGN_BOTTOM, 5 )

        Range.Add( bSizer9, 1, wx.EXPAND, 5 )
        #------------------------------------------------------------------------------------
        bSizer6 = wx.BoxSizer( wx.HORIZONTAL )

        self.Text_Curr_Start = wx.TextCtrl( self, wx.ID_ANY, '0.01', wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.AddSpacer(5)
        bSizer6.Add( self.Text_Curr_Start, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.Text_Curr_Stop = wx.TextCtrl( self, wx.ID_ANY, '2', wx.DefaultPosition, wx.Size( 36,24 ), 0 )
        bSizer6.AddSpacer(15)
        bSizer6.Add( self.Text_Curr_Stop, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Text_Curr_Step_uA = wx.TextCtrl( self, wx.ID_ANY, '5', wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer6.AddSpacer(10)
        bSizer6.Add( self.Text_Curr_Step_uA, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Text_Curr_Step_mA = wx.TextCtrl( self, wx.ID_ANY, '5', wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer6.AddSpacer(20)
        bSizer6.Add( self.Text_Curr_Step_mA, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Text_Curr_Step_A = wx.TextCtrl( self, wx.ID_ANY, '5', wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer6.AddSpacer(20)
        bSizer6.Add( self.Text_Curr_Step_A, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Text_Curr_Step_10A = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 24,24 ), 0 )
        bSizer6.AddSpacer(20)
        bSizer6.Add( self.Text_Curr_Step_10A, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        Range.Add( bSizer6, 1, wx.EXPAND, 5 )

        #------------------------------------------------------------------------------------

        TWO_SIDE.Add( Range, 1, wx.EXPAND, 5 )

        Subframe_SWEEP.Add(TWO_SIDE, 1, wx.EXPAND, 5 )

        self.Vertical_Frame_Left.Add( Subframe_SWEEP, 1, wx.EXPAND, 5 )

        # Connect Events
        self.Bt_RETURN.Bind( wx.EVT_BUTTON, self.Ev_RETURN )
        self.Bt_RUN.Bind( wx.EVT_BUTTON, self.Ev_RUN )
        self.Bt_Check_Instrument.Bind( wx.EVT_BUTTON, self.Ev_Check_Instrument )

    def MESSAGE_FRAME(self, new_title, sweep_item, changed_item):

        Subframe_MESSAGE = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Message" ), wx.VERTICAL )

        """
        bSizer2 = wx.BoxSizer( wx.HORIZONTAL )

        self.Label_Output = wx.StaticText( self, wx.ID_ANY, changed_item, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Label_Output.Wrap( -1 )
        self.Label_Output.SetFont( wx.Font( 9, 74, 90, 92, False, wx.EmptyString ) )
        self.Label_Output.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INFOTEXT ) )
        bSizer2.Add( self.Label_Output, 0,wx.ALL|wx.ALIGN_RIGHT, 5 )

        Subframe_MESSAGE.Add( bSizer2, 0, wx.ALL, 5 )
        """

        self.Text_Output = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, (400,80), wx.TE_MULTILINE )
        Subframe_MESSAGE.Add( self.Text_Output, 0, wx.ALL|wx.ALIGN_LEFT|wx.EXPAND, 5 )

        bSizer3 = wx.BoxSizer( wx.HORIZONTAL )

        Subframe_FIGURE = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Figure" ), wx.VERTICAL )
        self.panel = wx.Panel(self)
        self.init_plot(new_title, sweep_item, changed_item)
        self.canvas = FigCanvas(self.panel, -1, self.fig)
        #self.xmin_control = BoundControlBox(self.panel, -1, "X min", 0)
        #self.xmax_control = BoundControlBox(self.panel, -1, "X max", 50)
        #self.ymin_control = BoundControlBox(self.panel, -1, "Y min", 0)
        #self.ymax_control = BoundControlBox(self.panel, -1, "Y max", 100)
        Subframe_FIGURE.Add( self.panel, 0, wx.ALL, 5 )
        bSizer3.Add( Subframe_FIGURE, 1, wx.EXPAND, 5 )


        Subframe_MESSAGE.Add( bSizer3, 1, wx.ALL|wx.EXPAND, 5 )

        self.Vertical_Frame_Left.Add( Subframe_MESSAGE, 1, wx.EXPAND, 5 )

        # Connect Events
        self.Bt_SAVE_DATA.Bind( wx.EVT_BUTTON, self.Ev_SAVE_DATA )
        self.Bt_SAVE_FIG.Bind( wx.EVT_BUTTON, self.Ev_SAVE_FIG )

    def create_main_panel(self):
        Subframe_FIGURE = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Figure" ), wx.VERTICAL )

        self.panel = wx.Panel(self)

        self.init_plot()
        self.canvas = FigCanvas(self.panel, -1, self.fig)

        #self.xmin_control = BoundControlBox(self.panel, -1, "X min", 0)
        #self.xmax_control = BoundControlBox(self.panel, -1, "X max", 50)
        #self.ymin_control = BoundControlBox(self.panel, -1, "Y min", 0)
        #self.ymax_control = BoundControlBox(self.panel, -1, "Y max", 100)

        Subframe_FIGURE.Add( self.panel, 0, wx.ALL, 5 )
        self.Vertical_Frame_Right.Add( Subframe_FIGURE, 1, wx.EXPAND, 5 )

    def Ev_RUN( self, event ):

        self.Efficiency = defaultdict(dict)

        print "Push Run"

        try:
            Resource_AGILENT = self.RESOURCE["HEWLE"]
            self.rm = visa.ResourceManager()
            self.Agilent_E3632A_Init(Resource_AGILENT)
        except:
            wx.MessageBox('Agilent E3236A does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            Resource_PRODIGIT = self.RESOURCE["PRODIGI"]
            self.rm = visa.ResourceManager()
            self.Prodigit_3311C_Init(Resource_PRODIGIT)
        except:
            wx.MessageBox('Prodigit 3311C does not well connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            if float(self.Text_Curr_Stop.Value) < 0.1:
                self.PRODIGIT_CURR_OUT = np.linspace(0.01, float(self.Text_Curr_Stop.Value), int(self.Text_Curr_Step_uA.Value, 16))
            elif (0.1 <= float(self.Text_Curr_Stop.Value)) and (float(self.Text_Curr_Stop.Value) < 1):
                self.PRODIGIT_CURR_OUT = np.append(np.linspace(0.01, 0.1, int(self.Text_Curr_Step_uA.Value,16)),
                                           np.linspace(0.1, float(self.Text_Curr_Stop.Value), int(self.Text_Curr_Step_mA.Value,16)))
            elif (1 <= float(self.Text_Curr_Stop.Value)) and (float(self.Text_Curr_Stop.Value) < 10):
                self.PRODIGIT_CURR_OUT = np.append(np.linspace(0.01, 0.1, int(self.Text_Curr_Step_uA.Value,16)),
                                           np.linspace(0.1, 1.0, int(self.Text_Curr_Step_mA.Value,16)))
                self.PRODIGIT_CURR_OUT = np.append(self.PRODIGIT_CURR_OUT,
                                           np.linspace(1.0, float(self.Text_Curr_Stop.Value), int(self.Text_Curr_Step_A.Value,16)))
            elif (10 <= float(self.Text_Curr_Stop.Value)) and (float(self.Text_Curr_Stop.Value) < 100):
                self.PRODIGIT_CURR_OUT = np.append(np.linspace(0.01, 0.1, int(self.Text_Curr_Step_uA.Value,16)),
                                           np.linspace(0.1, 1.0, int(self.Text_Curr_Step_mA.Value,16)))
                self.PRODIGIT_CURR_OUT = np.append(self.PRODIGIT_CURR_OUT,
                                           np.linspace(1.0, 10.0, int(self.Text_Curr_Step_A.Value,16)))
                self.PRODIGIT_CURR_OUT = np.append(self.PRODIGIT_CURR_OUT,
                                           np.linspace(10.0, float(self.Text_Curr_Stop.Value), int(self.Text_Curr_Step_10A.Value,16)))
        except:
            wx.MessageBox('Sweep Range setting error', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        if (float(self.Text_Volt_Start.Value) == float(self.Text_Volt_Stop.Value)) or (self.Text_Volt_Step.Value == ''):
            self.AGILENT_VOLT_IN =[float(self.Text_Volt_Start.Value)]
        else:
            self.AGILENT_VOLT_IN = np.arange(float(self.Text_Volt_Start.Value), float(self.Text_Volt_Stop.Value), float(self.Text_Volt_Step.Value))

        for temp_IN_V in self.AGILENT_VOLT_IN:
            temp_total_effi=[]
            temp_total_IN_I=[]
            temp_total_IN_V=[]
            temp_total_OUT_I=[]
            temp_total_OUT_V=[]
            color_index = random
            self.Agilent_E3632A_Volt(temp_IN_V)
            #pdb.set_trace()
            try:
                for temp_OUT_I in self.PRODIGIT_CURR_OUT:
                    self.PRODIGIT_3311C.write(u'CURR:HIGH '+ str(temp_OUT_I) + u'\r')
                    self.PRODIGIT_3311C.write(u'CURR:LOW '+ str(temp_OUT_I) + u'\r')
                    #pdb.set_trace()
                    #self.AGILENT_E3632.clear()  # Disable to avoid Beeper noise from agilent
                    time.sleep(0.1)# Very import delay 0.5-sec, can't not be removed
                    try:
                        self.AGILENT_E3632.write("MEAS:CURR?\r")
                        self.PRODIGIT_3311C.write("MEAS:VOL?\r")
                        time.sleep(0.5)
                        temp_IN_I = float(self.AGILENT_E3632.read(termination='\r'))
                        temp_OUT_V = float(self.PRODIGIT_3311C.read())
                        temp_E = (temp_OUT_V * temp_OUT_I) / (temp_IN_V * temp_IN_I) * 100
                        if temp_E > 100:
                            temp_E = 100
                        temp_total_effi.append(temp_E)
                        temp_total_IN_I.append(temp_IN_I)
                        temp_total_IN_V.append(temp_IN_V)
                        temp_total_OUT_I.append(temp_OUT_I)
                        temp_total_OUT_V.append(temp_OUT_V)
                        #print(temp_total_effi)
                        #pdb.set_trace()
                        msg = u'Set Output Load:' + str(temp_OUT_I) + u', Output Voltage: '+ str(temp_OUT_V) + u'\n'
                        self.Text_Output.AppendText(msg)
                        msg = u'Get Input Curr:' + str(temp_IN_I) + u', Get Input Voltage:' + str(temp_IN_V) + u'\n'
                        self.Text_Output.AppendText(msg)
                    except:
                        self.Text_Output.AppendText("Efficiency Measure Current Error!")

            except:
                wx.MessageBox('Sweep error', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

            self.Efficiency[temp_IN_V] = temp_total_effi, temp_total_IN_I, temp_total_IN_V, temp_total_OUT_I, temp_total_OUT_V
            #print self.Efficiency[temp_IN_V]
            color = (random.random(), random.random(), random.random())
            label = 'Volt=' + str(temp_IN_V) + u'(V)'
            #pdb.set_trace()
            self.plot_draw(self.PRODIGIT_CURR_OUT, self.Efficiency[temp_IN_V][0], color, label)
            #self.axes.set_ylim([0,110])

        #print self.Efficiency
        self.Agilent_E3632A_close()
        self.Prodigit_3311C_close()
        self.rm.close()

    def Do_SAVE(self):
        # Save File
        #data = np.zeros(len(self.PRODIGIT_CURR_OUT))
        for temp_IN_V in self.AGILENT_VOLT_IN:
            #pdb.set_trace()
            data = np.c_[self.Efficiency[temp_IN_V][1]] # Input Current
            #data = np.c_[data, np.c_[self.Efficiency[temp_IN_V][1]]] # Input Current
            data = np.c_[data, np.c_[self.Efficiency[temp_IN_V][2]]] # Input Voltage
            data = np.c_[data, np.c_[self.Efficiency[temp_IN_V][3]]] # Output Current
            data = np.c_[data, np.c_[self.Efficiency[temp_IN_V][4]]] # Output Voltage
            data = np.c_[data, np.c_[self.Efficiency[temp_IN_V][0]]] # Efficiency
            #pdb.set_trace()
            index = self.save_path.find('.csv')
            np.savetxt(self.save_path[:index]+"_"+str(temp_IN_V)+"V"+self.save_path[index:], data, delimiter=", ")

        #data = np.vstack((np.insert(self.AGILENT_VOLT_IN,0 ,0) , data))

        #try:

        #except:
        #    self.Text_Output.AppendText('Please Close .CSV file')

class USB_I2C ( wx.Frame, WndProcHookMixin ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"USB_I2C", pos = wx.DefaultPosition, size = wx.Size( 312,345 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
        #favicon = wx.Icon(r'./icon/anpec.ico', wx.BITMAP_TYPE_ICO, 16, 16)
        #wx.Frame.SetIcon(self, favicon)

        self.windows_notification()

        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
        self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INFOBK ) )

        bSizer8 = wx.BoxSizer( wx.VERTICAL )

        USB_I2C = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"USB_I2C" ), wx.VERTICAL )

        USB_I2C_HORIZOTAL = wx.BoxSizer( wx.HORIZONTAL )

        bSizer10 = wx.BoxSizer( wx.VERTICAL )

        self.Device_ID = wx.StaticText( self, wx.ID_ANY, u"Device_ID(hex):  0x", wx.Point( -1,-1 ), wx.Size( -1,24 ), 0 )
        self.Device_ID.Wrap( -1 )
        bSizer10.Add( self.Device_ID, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Reg = wx.StaticText( self, wx.ID_ANY, u"Reg(hex): 0x", wx.DefaultPosition, wx.Size( -1,24 ), 0 )
        self.Reg.Wrap( -1 )
        bSizer10.Add( self.Reg, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )

        self.Data = wx.StaticText( self, wx.ID_ANY, u"Data(hex): 0x", wx.DefaultPosition, wx.Size( -1,24 ), 0 )
        self.Data.Wrap( -1 )
        bSizer10.Add( self.Data, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )


        USB_I2C_HORIZOTAL.Add( bSizer10, 1, wx.EXPAND, 5 )

        bSizer11 = wx.BoxSizer( wx.VERTICAL )

        self.Text_Device_ID = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,-1 ), 0 )
        bSizer11.Add( self.Text_Device_ID, 0, wx.ALL, 5 )

        self.Text_Reg = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,-1 ), 0 )
        bSizer11.Add( self.Text_Reg, 0, wx.ALL, 5 )

        self.Text_Data = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 36,-1 ), 0 )
        bSizer11.Add( self.Text_Data, 0, wx.ALL, 5 )


        USB_I2C_HORIZOTAL.Add( bSizer11, 1, wx.EXPAND, 5 )

        bSizer13 = wx.BoxSizer( wx.VERTICAL )

        self.Device_Check = wx.StaticText( self, wx.ID_ANY, u"Push Reset Button!", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Device_Check.SetForegroundColour( wx.Colour( 0, 0, 255 ) )
        self.Device_Check.Wrap( -1 )

        bSizer13.Add( self.Device_Check, 0, wx.ALL, 5 )

        self.Bt_READ = wx.Button( self, wx.ID_ANY, u"Read", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_READ.SetBackgroundColour( wx.Colour( 255,128,64 ) )
        bSizer13.Add( self.Bt_READ, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )

        self.Bt_WRITE = wx.Button( self, wx.ID_ANY, u"Write", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_WRITE.SetBackgroundColour( wx.Colour( 0, 128, 0 ) )
        bSizer13.Add( self.Bt_WRITE, 0, wx.ALIGN_RIGHT|wx.TOP|wx.RIGHT|wx.LEFT, 5 )


        USB_I2C_HORIZOTAL.Add( bSizer13, 1, wx.EXPAND, 5 )


        USB_I2C.Add( USB_I2C_HORIZOTAL, 1, wx.EXPAND, 5 )

        bSizer12 = wx.BoxSizer( wx.VERTICAL )


        USB_I2C.Add( bSizer12, 1, wx.EXPAND, 5 )

        self.Bt_RETURN = wx.Button( self, wx.ID_ANY, u"Return", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Bt_RETURN.SetBackgroundColour( wx.Colour( 0, 128, 192 ) )
        USB_I2C.Add( self.Bt_RETURN, 0, wx.ALL, 5 )


        bSizer8.Add( USB_I2C, 1, wx.EXPAND, 5 )

        MESSAGE = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Message" ), wx.VERTICAL )

        self.Label_Output_Message = wx.StaticText( self, wx.ID_ANY, u"Output", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.Label_Output_Message.Wrap( -1 )
        MESSAGE.Add( self.Label_Output_Message, 0, wx.ALL, 5 )

        self.Text_Ouput_message = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 300,100 ), wx.TE_MULTILINE )
        MESSAGE.Add( self.Text_Ouput_message, 0, wx.ALL, 5 )


        bSizer8.Add( MESSAGE, 1, wx.EXPAND, 5 )


        self.SetSizer( bSizer8 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.Bt_READ.Bind( wx.EVT_BUTTON, self.Ev_READ )
        self.Bt_WRITE.Bind( wx.EVT_BUTTON, self.Ev_WRITE )
        self.Bt_RETURN.Bind( wx.EVT_BUTTON, self.Ev_RETURN )

        #self.onDeviceChange()

    def windows_notification(self):
        """ Windows Notification """
        WndProcHookMixin.__init__(self)
        self.Bind(wx.EVT_CLOSE, self.onClose)
        """Change the following guid to the GUID of the device you want notifications for"""
        self.devNotifyHandle = self.registerDeviceNotification(guid="{A5DCBF10-6530-11D2-901F-00C04FB951ED}")
        self.addMsgHandler(WM_DEVICECHANGE, self.onDeviceChange)
        self.hookWndProc()
        ################################################################################################################

    def __del__( self ):
        pass
    # Virtual event handlers, overide them in your derived class

    def USB_I2C_Device_Check():
        rm = visa.ResourceManager()
        USB = rm.list_resources('USB?*0x0451::0xBB01?*')
        inst = rm.open_resource(str(USB[0]))
        lib = rm.visalib
        attribute_state = c_int
        status = lib.enable_event(inst.session, constants.VI_EVENT_USB_INTR, 1, 0)
        status = lib.usb_control_in(inst.session, 128, 6, 1023, 0, 12)
        status = lib.get_attribute(inst.session, constants.VI_ATTR_MODEL_NAME)
        print "USB device plugin"
        print(status[0])
        inst.close()
        rm.close()

    def USB_I2C_Read(self, Device_ID, Reg_Address):
        rm = visa.ResourceManager()
        try:
            USB = rm.list_resources('USB?*0x0451::0xBB01?*')
            inst = rm.open_resource(str(USB[0]))
            inst.timeout = 2500
            inst.chunk_size = 256
            lib = rm.visalib
            status = lib.enable_event(inst.session, constants.VI_EVENT_USB_INTR, 1, 0)
            out_value = data_out_buffer(0x02, Device_ID, 0x01, Reg_Address, 0x00) # Read
            status = lib.usb_control_out(inst.session, 33, 9, 0, 3, out_value)
            #out_value = data_out_buffer(0x02, Device_ID, 0x01, Reg_Address, 0x55) # Read
            status = lib.usb_control_out(inst.session, 33, 9, 0, 3, out_value)
            status = lib.wait_on_event(inst.session, constants.VI_EVENT_USB_INTR, 500)
            inst_Rd = status[1]
            status = lib.get_attribute(inst_Rd, constants.VI_ATTR_USB_RECV_INTR_DATA)
            #print(status[0])
            READ_DATA = status[0]
        except:
            wx.MessageBox('USB_I2C does not connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        try:
            #print hex(READ_DATA[4])[2:]
            #print hex(READ_DATA[4])[2:]+hex(READ_DATA[5])[2:]
            status = lib.close(inst_Rd)
            status = lib.disable_event(inst.session, constants.VI_EVENT_USB_INTR, 1)
            if len(hex(READ_DATA[4])[2:]) == 1:
                data_h = str(0) + hex(READ_DATA[4])[2:]
            else:
                data_h = hex(READ_DATA[4])[2:]
            """
            if len(hex(READ_DATA[5])[2:]) == 1:
                data_l = str(0) + hex(READ_DATA[5])[2:]
            else:
                data_l = hex(READ_DATA[5])[2:]
            """
            return data_h
        except:
            wx.MessageBox('USB_I2C Read Data Error ~', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

        rm.close()

    def USB_I2C_Write(self, Device_ID, Reg_Address, Reg_Data):
        rm = visa.ResourceManager()
        try:
            USB = rm.list_resources('USB?*0x0451::0xBB01?*')
            inst = rm.open_resource(str(USB[0]))
            inst.timeout = 2500
            inst.chunk_size = 256
            lib = rm.visalib
            data_out = c_uint8 * 12
            #print Reg_Data
            Reg_Data_H = Reg_Data/256
            Reg_Data_L = Reg_Data%256
            #print Reg_Data_H
            #print Reg_Data_L
            ii = data_out(0x12, Device_ID, 0x01, Reg_Address, Reg_Data_H, Reg_Data_L) # """  Write """
            status = lib.enable_event(inst.session, constants.VI_EVENT_USB_INTR, 1, 0)
            status = lib.usb_control_out(inst.session, 33, 9, 0, 3, ii)
            inst.close()
        except:
            wx.MessageBox('USB_I2C does not connect!.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        rm.close()

    def Ev_READ( self, event ):
        self.Text_Ouput_message.AppendText('USB-I2C Read!\n')
        try:
            self.Text_Data.Value = self.USB_I2C_Read(int(self.Text_Device_ID.Value, 16), int(self.Text_Reg.Value, 16))
        except:
            wx.MessageBox('Please Check Device_ID and Register_Address.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)

    def Ev_WRITE( self, event ):
        try:
            self.USB_I2C_Write(int(self.Text_Device_ID.Value, 16), int(self.Text_Reg.Value, 16), int(self.Text_Data.Value, 16))
        except:
            wx.MessageBox('Please Check Device_ID and Register_Address.', 'Error',wx.ICON_EXCLAMATION | wx.STAY_ON_TOP)
        self.Text_Ouput_message.AppendText('USB-I2C Write!\n')

    def Ev_RETURN( self, event ):
        TOP=Top_Option(None)
        TOP.Show()
        self.Destroy()
        print "Push Return"

    def onDeviceChange(self,wParam,lParam):
        if lParam:
            dbh = DEV_BROADCAST_HDR.from_address(lParam)
            if dbh.dbch_devicetype == DBT_DEVTYP_DEVICEINTERFACE:
                dbd = DEV_BROADCAST_DEVICEINTERFACE.from_address(lParam)
                #Verify that the USB VID and PID match our assigned VID and PID
                if 'VID_0451&PID_BB01' in dbd.dbcc_name:
                    if wParam == DBT_DEVICEARRIVAL:
                      self.Device_Check.Label='Device Exit!'
                      self.Device_Check.SetForegroundColour( wx.Colour( 0, 255, 0 ) )
                    elif wParam == DBT_DEVICEREMOVECOMPLETE:
                      self.Device_Check.Label='Device Not Exit!'
                      self.Device_Check.SetForegroundColour( wx.Colour( 255, 0, 0 ) )
        return True

    def onClose(self, event):
        self.unregisterDeviceNotification(self.devNotifyHandle)
        event.Skip()

if __name__ == '__main__':
    app=wx.App()
    TOP=Top_Option(None)
    TOP.Show()
    app.MainLoop()