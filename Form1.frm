VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Telephonic connection"
   ClientHeight    =   4605
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   4350
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7752
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2040
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2055
      Top             =   3690
   End
   Begin VB.Frame Frame3 
      Height          =   150
      Left            =   0
      TabIndex        =   25
      Top             =   -120
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Number to dial:"
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   735
         Left            =   1440
         Picture         =   "Form1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Dial"
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Height          =   570
         Left            =   -15
         TabIndex        =   28
         Top             =   225
         Width           =   2175
      End
      Begin VB.Label NumberToDialLabel 
         AutoSize        =   -1  'True
         Caption         =   "Number to dial:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   2400
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   7
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   6
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   5
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   4
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   3
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   2
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   1
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label FastDialingLabel 
         Caption         =   " Fast Dialing"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   35
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   34
         Top             =   3600
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   33
         Top             =   3120
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   32
         Top             =   2640
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   31
         Top             =   1680
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   30
         Top             =   1200
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   29
         Top             =   720
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         MousePointer    =   7  'Size N S
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   225
         Width           =   2025
      End
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu menuexit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu menuminimize 
         Caption         =   "&Minimize"
      End
   End
   Begin VB.Menu menumodify 
      Caption         =   "Modify"
   End
   Begin VB.Menu menutools 
      Caption         =   "Tools"
   End
   Begin VB.Menu menuhelp 
      Caption         =   "?"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Dial(num As String)
  Static busy As Boolean

num = Trim(num)
If busy Then
 Exit Sub
End If

busy = True
    
   ' Open the com port.
   On Error Resume Next
   MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: port not available"
       Exit Sub
    End If
   On Error GoTo 0
   
   'Empty the input buffer
   MSComm1.InBufferCount = 0
   
   'Send the attention command (AT) to the modem
   'and wait for the OK response
      'NOTE:look into modem documentation for the
      'complete list of Hayes compatible commands.
   'vbCr=chr(13) and vbNewLine=vbCrLf=chr(13)+chr(10)
   ret = Pause(1)
   Do
     If ret = vbOK Then MSComm1.Output = "AT" & vbCr
     If (ret = vbAbort) Or ExitPause Then GoTo AbortExit
     ret = Pause(500)
     If ret = vbOK Then Buffer$ = Buffer$ & MSComm1.Input
   Loop Until InStr(Buffer$, "OK" & vbNewLine)
    
   'Dial the number.
   'Output the attention command (AT) using dial tone (DT)
   'The semicolon means to the modem: after stay listening
   'for more commands (don't vorget it, don't remove it)
      'NOTE:look into modem documentation for the
      'complete list of Hayes compatible commands.
   MSComm1.Output = "AT" & "DT" & " " & num & ";" & vbCr
        
   'wait the number is composed and sent by the modem
   ret = Pause(15000) '15 seconds, interruptable pause
   If ret = vbAbort Then GoTo AbortExit

AbortExit:
   
   'Close the com port
   If MSComm1.PortOpen And Not ClosingModem Then MSComm1.PortOpen = False
   
         
busy = False

End Sub

Private Sub Combo1_DropDown()
 DropDownCombo = True
End Sub

Private Sub Combo1_GotFocus()
  If DropDownCombo Then DropDownCombo = False: Exit Sub
  Combo1.Text = ""
  ExitPause = True
  BlankCombo = True
End Sub

Private Sub Command1_Click(Index As Integer) 'keyboard
 If BlankCombo Then BlankCombo = False: ExitPause = True: Combo1.Text = ""
 Select Case Index
  Case 0 To 9
    Combo1.Text = Combo1.Text & Index
  Case 10
    Combo1.Text = Combo1.Text & "*"
  Case 11
    Combo1.Text = Combo1.Text & "#"
 End Select
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Command1", Index
End Sub

Private Sub Command2_Click(Index As Integer) 'fast dial
If Command2(Index).Caption = "" Then
 Enabled = False
 IndexSave = Index
 Dialog1.Show modal, Me
 Exit Sub
End If
 Combo1.Text = Command2(Index).Tag
 ExitPause = True
 BlankCombo = True
End Sub

Private Sub Command2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 2 Then Exit Sub
If Command2(Index).Caption = "" Then Exit Sub
 'indexsave is used by Dialog1
 'indexsave is used to drag and drop buttons also (reorder)
 Enabled = False
 IndexSave = Index
 Dialog1.Show modal, Me
 Dialog1.Text1 = Command2(Index).Caption
 Dialog1.Text2 = Command2(Index).Tag
 
End Sub

Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Command2", Command2(Index).Caption, Command2(Index).Tag
End Sub

Private Sub Command3_Click() 'dial
  For a = 0 To Combo1.ListCount - 1
    If Combo1.List(a) = Combo1.Text Then found = True
  Next
  If Not found Then Combo1.AddItem Combo1.Text

  Dial Combo1.Text
  
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Command3"
End Sub

Private Sub Command4_Click() 'close
  
  ClosingModem = True
  
  'interrupt Pause subroutine eventually running
  ExitPause = True 'stop current pause
  
  
  Do
     'Open the com port
     If Not MSComm1.PortOpen Then MSComm1.PortOpen = True

     If MSComm1.PortOpen Then
        'Empty the input buffer
        MSComm1.InBufferCount = 0
        Do
           'Deconnect modem
           'Commands: attention (AT) and halt (H)
              'NOTE:look into modem documentation for the
              'complete list of Hayes compatible commands.
           'don't insert semicolon here
           'vbCr=chr(13)
           MSComm1.Output = "AT" & "H" & vbCr
           Pause 500
           Buffer$ = Buffer$ & MSComm1.Input
        Loop Until InStr(Buffer$, "OK" & vbNewLine)
   
  
       Do
          'Close the com port
          If MSComm1.PortOpen Then MSComm1.PortOpen = False
          If Not MSComm1.PortOpen Then Exit Do
       Loop
  
       Exit Do
     End If
  Loop
  
  
  Combo1.Text = ""
  
  ClosingModem = False

End Sub



Private Sub file_Click()

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Command4"
End Sub

Private Sub FastDialingLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "FastDialingLabel"
End Sub

Private Sub Form_Load()

 FormLoad = True
 
 SaveWidth = Width
 SaveHeight = Height
 
 'clean up register,debug only
 'DeleteSetting App.ProductName
 
 For a = 0 To 7
   Command2(a).Caption = GetSetting(App.ProductName, "FastDialName", a, "")
   Command2(a).Tag = GetSetting(App.ProductName, "FastDialNumber", a, "")
 Next
    
 'port settings and number (1 is mouse)
 MSComm1.Settings = "9600,n,8,1"
 MSComm1.CommPort = 2
 'Set InputLen to 0 and the MSComm will read the whole
 'buffer content when Input property will be used
 MSComm1.InputLen = 0

 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Form1"
End Sub

Private Sub Form_Resize()
 If FormLoad Then FormLoad = False: Exit Sub
 On Error Resume Next
 Width = SaveWidth
 Height = SaveHeight
 On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Frame1"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Frame2"
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "Frame3"
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  'label that, as a result, drags a fast dial button
  'to another of the 8 places, swapping
  
  'Index is the drop index
  'IndexSource is the drag index
  'let's swap them
  If Source.Name <> "Label1" Then Exit Sub
  IndexSource = Source.Index
  
  If Index = IndexSource Then Exit Sub 'is a simple click
  
  SaveSetting App.ProductName, "FastDialName", IndexSource, Command2(Index).Caption
  SaveSetting App.ProductName, "FastDialNumber", IndexSource, Command2(Index).Tag
  SaveSetting App.ProductName, "FastDialName", Index, Command2(IndexSource).Caption
  SaveSetting App.ProductName, "FastDialNumber", Index, Command2(IndexSource).Tag
  Command2(IndexSource).Caption = GetSetting(App.ProductName, "FastDialName", IndexSource, "error")
  Command2(IndexSource).Tag = GetSetting(App.ProductName, "FastDialNumber", IndexSource, "error")
  Command2(Index).Caption = GetSetting(App.ProductName, "FastDialName", Index, "error")
  Command2(Index).Tag = GetSetting(App.ProductName, "FastDialNumber", Index, "error")

End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "Label1"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Combo1"
End Sub

Private Sub menuexit_Click()
 Unload Me
End Sub

Private Sub menuhelp_Click()
 frmAbout.Show 0, Form1  '0=not modal 1=vbmodal, must be not modal
End Sub

Private Sub menuminimize_Click()
 WindowState = vbMinimized
End Sub


Private Sub menumodify_Click()
 'to do, connection options
End Sub

Private Sub NumberToDialLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "NumberToDialLabel"
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "StatusBar1"
End Sub

Private Sub Timer1_Timer() 'always running
 'highlightes the fast dial button whenever
 'the content of it (its Tag, i.e. the telephon number)
 'equates the Combo1.Text (i.e. the number shown in the
 'Combo1 window
 For a = 0 To 7
  If Command2(a).Tag = Combo1.Text Then
    If Command2(a).Caption <> "" Then
      If Command2(a).BackColor <> vbWindowBackground Then Command2(a).BackColor = vbWindowBackground
    End If
  Else
    If Command2(a).BackColor <> vbButtonFace Then Command2(a).BackColor = vbButtonFace
  End If
 Next
End Sub
