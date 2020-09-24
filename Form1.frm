VERSION 5.00
Object = "{90D1C92B-C265-4C72-8A34-9938EAEA05C8}#1.0#0"; "Mousewheel.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Wheel Test"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3870
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1680
      Left            =   1958
      ScaleHeight     =   1620
      ScaleWidth      =   1620
      TabIndex        =   2
      Top             =   187
      Width           =   1680
   End
   Begin WheelCtl.MouseWheel MouseWheel1 
      Left            =   1373
      Top             =   742
      _ExtentX        =   529
      _ExtentY        =   503
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   233
      TabIndex        =   1
      Text            =   "0"
      Top             =   1282
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   263
      TabIndex        =   0
      Top             =   187
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Arc As Single

Private Sub Form_Load()

    Show
    Label1.ToolTipText = IIf(MouseWheel1.WheelPresent And MouseWheel1.Enabled, "Scroll Mouse Wheel", IIf(MouseWheel1.Enabled, "You have no Mouse Wheel", "Mouse Wheel is disabled"))
    Text1.ToolTipText = Label1.ToolTipText
    Arc = -0.01
    Picture1.Circle (Picture1.Width / 2.1, Picture1.Height / 2), Picture1.Height * 0.45, vbRed, -0.001, Arc

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'MouseWheel1.WheelDisconnect

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'connect wheel

    MouseWheel1.WheelConnect hWnd 'Label1 has no hwnd so we take the form's hwnd instead _
                             this has the effect that unless the wheel is disconnected by _
                             Form_MouseMove above it will remain active and connected to _
                             the label

End Sub

Private Sub MouseWheel1_WheelScroll(ByVal ConnectedTo As Long, ByVal Direction As Long, ByVal Shift As WheelCtl.KeyDown)

    Select Case Shift 'shift could also be used to control horizontal or vertical scrolling
      Case WheelCtl.KeyCntl
        Direction = Direction * 2
      Case WheelCtl.KeyShift
        Direction = Direction * 4
      Case WheelCtl.KeyBoth
        Direction = Direction * 8
    End Select
    Select Case ConnectedTo
      Case Text1.hWnd
        Text1 = Val(Text1) + Direction
      Case Picture1.hWnd
        Arc = Arc - Direction / 12
        If Arc < -0.001 And Arc > -6.25 Then
            Picture1.Cls
            Picture1.Circle (Picture1.Width / 2 - 30, Picture1.Height / 2 - 30), Picture1.Height * 0.45, vbRed, -0.001, Arc
          Else 'NOT ARC...
            Arc = Arc + Direction / 10
        End If
      Case Else  'NOT CONNECTEDTO...
        Label1 = Val(Label1) + Direction
    End Select

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'connect wheel

    MouseWheel1.WheelConnect Picture1.hWnd

End Sub

Private Sub Text1_Change()

    Text1.BackColor = QBColor(Val(Text1) And 15)
    Text1.ForeColor = QBColor(15 - (Val(Text1) And 15))

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'connect wheel

    MouseWheel1.WheelConnect Text1.hWnd
    Text1.SetFocus

End Sub

':) Ulli's VB Code Formatter V2.13.3 (02.08.2002 01:50:15) 2 + 78 = 80 Lines
