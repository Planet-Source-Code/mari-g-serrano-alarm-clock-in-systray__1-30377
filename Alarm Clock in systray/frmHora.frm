VERSION 5.00
Begin VB.Form frmHora 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   135
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   10
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   0
      Top             =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu PopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuApagar 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuCambiar 
         Caption         =   "Change time of Alarm "
      End
      Begin VB.Menu mnufecha 
         Caption         =   "fecha..."
      End
   End
End
Attribute VB_Name = "frmHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()
    Me.ScaleWidth = 1
    Me.ScaleHeight = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu PopUp, , , , mnuApagar
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static CtrMov As Boolean

With Me
    If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
        ReleaseCapture
        CtrMov = False
        Label1.ForeColor = &HFF8080
        Label1.FontUnderline = False
        ''debug.Print "LostMouseFocus"
    Else
        SetCapture .hWnd
        If CtrMov = False Then
            CtrMov = True
          Label1.ForeColor = vbBlue '
          Label1.FontUnderline = True
          ''  debug.Print "GetMouseFocus"
        End If
    End If
End With
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    form_MouseMove Button, Shift, X, Y
End Sub

Private Sub mnuApagar_Click()
    Unload Form1
    Unload Me
End Sub

Private Sub mnuCambiar_Click()
    On Error GoTo err
   
    Shell App.Path & "\" & App.EXEName & ".EXE", vbNormalFocus
    Unload Form1
    Unload Me
    Exit Sub
err:
    MsgBox App.Path & "\" & App.EXEName & ".EXE not found"
    
End Sub

Private Sub Timer1_Timer()
Static b As Integer
If b < 3 Then b = b + 1
If b = 2 Then iniT


    Label1 = Time
    mnufecha.Caption = "Today is " & WeekdayName(Weekday(Date, vbMonday)) _
                        & ", " & Day(Date) & " of " & _
                        MonthName(Month(Date)) & " of " & Year(Date)
    
    mnuCambiar.Caption = "Chamge alarm time (" & sAlarma & ")"
End Sub

Sub iniT()
'push the new clock
 Dim thWnd As Long, rctemp As RECT
    Me.Visible = True
    thWnd = FindWindow("Shell_TrayWnd", vbNullString)
    thWnd = FindWindowEx(thWnd, 0, "TrayNotifyWnd", vbNullString)
    thWnd = FindWindowEx(thWnd, 0, "TrayClockWClass", vbNullString) 'uncomment
    
    GetWindowRect thWnd, rctemp
        With frmHora
            .Top = 0
            .Left = 0
            .Height = frmHora.Height * (rctemp.Bottom - rctemp.Top) / frmHora.ScaleHeight
            .Width = frmHora.Width * (rctemp.Right - rctemp.Left) / frmHora.ScaleWidth
        End With
    SetParent frmHora.hWnd, thWnd

End Sub
