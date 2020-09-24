VERSION 5.00
Begin VB.UserControl MouseWheel 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fest Einfach
   CanGetFocus     =   0   'False
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   1860
   ToolboxBitmap   =   "MouseWheel.ctx":0000
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   15
      Top             =   255
   End
   Begin VB.Image img 
      Height          =   225
      Left            =   0
      Picture         =   "MouseWheel.ctx":00FA
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "MouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Attribute GetSystemMetrics.VB_Description = "Win API"
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As tMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Attribute GetMessage.VB_Description = "Win API"
Private Declare Function TranslateMessage Lib "user32" (lpMsg As tMSG) As Long
Attribute TranslateMessage.VB_Description = "Win API"
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As tMSG) As Long
Attribute DispatchMessage.VB_Description = "Win API"
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long
Attribute GetCursorPos.VB_Description = "Win API"
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Attribute WindowFromPoint.VB_Description = "Win API"

Private Const SM_MOUSEWHEELPRESENT  As Long = 75
Attribute SM_MOUSEWHEELPRESENT.VB_VarDescription = "Win API constant"
Private Const WM_MOUSEWHEEL         As Long = &H20A
Attribute WM_MOUSEWHEEL.VB_VarDescription = "Win API constant"

Private Type tPOINT
    X As Long
    Y As Long
End Type
Private MouseCoords As tPOINT
Attribute MouseCoords.VB_VarDescription = "Mouse Cursor Coordinates"

Private Type tMSG
    hWnd        As Long
    nMsg        As Long
    wParam      As Long
    lParam      As Long
    time        As Long
    pt          As tPOINT
End Type
Private MSG As tMSG
Attribute MSG.VB_VarDescription = "Windows Message Structure"

Public Enum KeyDown
    KeyNone = 0
    KeyShift = 4
    KeyCntl = 8
    KeyBoth = 12
End Enum

Private myEnabled           As Boolean
Attribute myEnabled.VB_VarDescription = "Private property"
Private myAutoDisconnect    As Boolean
Attribute myAutoDisconnect.VB_VarDescription = "Private property"
Private WheelIsPresent      As Boolean
Attribute WheelIsPresent.VB_VarDescription = "True if a wheel mouse is connected"
Private hWndConnected       As Long
Attribute hWndConnected.VB_VarDescription = "The hnd of the control to which the wheel is connected"

Public Event WheelScroll(ByVal ConnectedTo As Long, ByVal Direction As Long, ByVal Shift As KeyDown)
Attribute WheelScroll.VB_Description = "Fired when user scrolls the wheel"

Public Property Let AutoDisconnect(ByVal nuAutoDisconnect As Boolean)
Attribute AutoDisconnect.VB_Description = "Determines whether or not the wheel is automatically disconnected from a control, when the mouse cursor leaves the control."
Attribute AutoDisconnect.VB_ProcData.VB_Invoke_PropertyPut = ";Verhalten"

    myAutoDisconnect = (nuAutoDisconnect <> False)
    PropertyChanged "AutoDisconnect"

End Property

Public Property Get AutoDisconnect() As Boolean

    AutoDisconnect = myAutoDisconnect

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets or returns whether the control will react to user input."

    Enabled = myEnabled

End Property

Public Property Let Enabled(ByVal nuEnabled As Boolean)

    myEnabled = (nuEnabled <> False) And WheelIsPresent
    PropertyChanged "Enabled"

End Property

Public Property Get IsConnected() As Boolean
Attribute IsConnected.VB_Description = "Returns whether the wheel is connected or idle."

    IsConnected = tmr.Enabled

End Property

Private Sub tmr_Timer()
Attribute tmr_Timer.VB_Description = "Timer event"

    GetMessage MSG, Parent.hWnd, 0, 0
    TranslateMessage MSG
    DispatchMessage MSG
    GetCursorPos MouseCoords 'get current mouse location
    If WindowFromPoint(MouseCoords.X, MouseCoords.Y) = hWndConnected Then
        With MSG
            If .nMsg = WM_MOUSEWHEEL Then
                RaiseEvent WheelScroll(hWndConnected, Sgn(.wParam), .wParam And &HFFFF&) 'last param has key
            End If
        End With 'MSG
      Else 'NOT WINDOWFROMPOINT(MOUSECOORDS.X,...
        If myAutoDisconnect Then
            WheelDisconnect
        End If
    End If

End Sub

Private Sub UserControl_Initialize()
Attribute UserControl_Initialize.VB_Description = "Called when the control is initialized"

    WheelIsPresent = GetSystemMetrics(SM_MOUSEWHEELPRESENT)

End Sub

Private Sub UserControl_InitProperties()
Attribute UserControl_InitProperties.VB_Description = "Initializes the control properties to their default values"

    WheelIsPresent = GetSystemMetrics(SM_MOUSEWHEELPRESENT)
    myEnabled = WheelIsPresent

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Attribute UserControl_ReadProperties.VB_Description = "Get properties from property bag"

    With PropBag
        Enabled = .ReadProperty("Enabled", True) And WheelIsPresent
        AutoDisconnect = .ReadProperty("AutoDisconnect", False)
    End With 'PROPBAG

End Sub

Private Sub UserControl_Resize()
Attribute UserControl_Resize.VB_Description = "Called when the user resizes the control"

    Size img.Width + 60, img.Height + 60

End Sub

Private Sub UserControl_Terminate()
Attribute UserControl_Terminate.VB_Description = "Called when control is destroyed"

    tmr.Enabled = False

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Attribute UserControl_WriteProperties.VB_Description = "Write properties into property bag"

    With PropBag
        .WriteProperty "Enabled", myEnabled, True
        .WriteProperty "AutoDisconnect", myAutoDisconnect, False
    End With 'PROPBAG

End Sub

Public Sub WheelConnect(ByVal hWnd As Long)
Attribute WheelConnect.VB_Description = "Used to connect the wheel to a control."
Attribute WheelConnect.VB_UserMemId = 0

    hWndConnected = hWnd
    tmr.Enabled = myEnabled

End Sub

Public Function WheelDisconnect() As Boolean
Attribute WheelDisconnect.VB_Description = "Used to disconnect the wheel from a control."

    tmr.Enabled = False

End Function

Public Property Get WheelPresent() As Boolean
Attribute WheelPresent.VB_Description = "Returns whether a wheel mouse is connectd to this computer."

    WheelPresent = WheelIsPresent

End Property

':) Ulli's VB Code Formatter V2.13.3 (02.08.2002 00:10:43) 41 + 116 = 157 Lines
