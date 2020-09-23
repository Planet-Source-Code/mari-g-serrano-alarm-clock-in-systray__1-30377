VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "ALARMA"
   ClientHeight    =   735
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   1845
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlarma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   -30
      Width           =   75
   End
   Begin VB.Menu popUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuApagar 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuCambiar 
         Caption         =   "Change time of alarm "
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   On Error Resume Next
   If App.PrevInstance Then End
   If Trim(Command$()) = "" Then
      
      sAlarma = InputBox("Time of the Alarm: (The Time is: " & Time & ")", "TaskBar Alarm", "8:27")
      sMensaje = Space$(25) & InputBox("Message: ", _
        "TaskBar Alarm", "This is a Message to show when the alarm must appear!!!") & " "
   Else
      sAlarma = Mid(Command$(), InStr(1, Command$(), "/") + 1, _
                InStr(1, Command$(), ";") - 1 - InStr(1, Command$(), "/"))
      sMensaje = Space$(25) & Mid(Command$(), InStr(1, Command$(), ";") + 1) & " "
   End If
   If Hour(sAlarma) = 0 And Minute(sAlarma) = 0 And Second(sAlarma) = 0 Then _
        MsgBox sAlarma & " no es una hora válida": End

   Timer1.Enabled = False
   Timer.Enabled = True
   frmHora.Show
End Sub
Private Sub Form_DblClick()
    Unload frmHora
    Unload Me
End Sub


Sub PonerenTray()
'show the message in the trayNotifyWnd..)
    Dim hWnd As Long, rctemp As RECT
    Me.Caption = ""
    Me.Visible = True
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    hWnd = FindWindowEx(hWnd, 0, "TrayNotifyWnd", vbNullString)
    GetWindowRect hWnd, rctemp
    With Me
        .Top = 0
        .Left = 0
        .Height = Me.Height * (rctemp.Bottom - rctemp.Top) / Me.ScaleHeight
        .Width = Me.Width * (rctemp.Right - rctemp.Left) / Me.ScaleWidth
    End With
    Label1.Move 0, 0, Me.Width, Me.Height
    SetParent Me.hWnd, hWnd
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu PopUp, , , , mnuApagar
End Sub

Private Sub Label1_DblClick()
    Unload frmHora
    Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub mnuApagar_Click()
    Unload frmHora
    Unload Me
End Sub

Private Sub mnuCambiar_Click()
    On Error GoTo err
    
    Shell App.Path & "\" & App.EXEName & ".EXE", vbNormalFocus
    Unload frmHora
    Unload Me
    Exit Sub
err:
    MsgBox App.Path & "\" & App.EXEName & ".EXE not found"
End Sub

Private Sub Timer_Timer()
    If sAlarma = "" Then sAlarma = CStr(Time)
    
    If Time > CDate(sAlarma) Then
        PonerenTray
        Timer.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    
    Static count As Integer
    count = count + 1
    If count = Len(sMensaje) + 1 Then count = 1
    Static b As Boolean
    b = Not b
    Me.BackColor = IIf(b, vbBlue, &H8000000F)
    Label1 = Mid(sMensaje, count) 'Time
    
End Sub


