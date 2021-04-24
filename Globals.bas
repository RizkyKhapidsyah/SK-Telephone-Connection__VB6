Attribute VB_Name = "Globals"
Public FormLoad As Boolean
Public IndexSave As Integer
Public SaveWidth As Integer
Public SaveHeight As Integer
Public ExitPause As Boolean
Public BlankCombo As Boolean
Public DropDownCombo As Boolean
Public ClosingModem As Boolean

'++++++++++++++++++++++ PAUSE MILLISECONDS ++++++++++++++++++++++
Public Function Pause(ByVal milliseconds As Double) As Integer
'pauses milliseconds
'pause may be interrupted making the global ExitPause true
'returns: vbOK=1 vbAbort=3
Dim Start, Finish, TotalTime
    'interruption from now on only
    ExitPause = False
    Start = Timer * 1000 ' Set start time.
    Do While Timer * 1000 < Start + milliseconds
        DoEvents    ' Yield to other processes.
        If Timer * 1000 < Start Then 'trepassing midnight
          Start = Start - 24 * 60 * 60 * 1000
        End If
        'interruption
        If ExitPause Then ExitPause = False: Pause = vbAbort: Exit Function
    Loop
    Pause = vbOK
End Function

Public Sub Help(ByVal objectname As String, Optional ByRef more = "", Optional ByRef moremore = "")
  'addressed by many controls of many forms
    Select Case objectname
     Case "NumberToDialLabel"
        msg = "Enter number: directly, from keyboard or fast dialing"
     Case "FastDialingLabel"
        msg = "Buttons configurable with Names and TF numbers"
     Case "Command1" 'telephon keyboard
        Select Case more 'more is index
           Case 0 To 9 ' numbers
              msg = "Telephon Keyboard [" & more & "]"
           Case 10 ' asterisk
              msg = "Telephon Keyboard [*]"
           Case 11 ' diesis
              msg = "Telephon Keyboard [#]"
        End Select
     Case "Command3" 'dial
        msg = "Click to dial (turn on modem)"
     Case "Command4" 'stop modem
        msg = "Click to deconnect modem"
     Case "Label1" 'reorder fast dial list
        msg = "Click and drag to reorder the list"
     Case "Command2" 'fast dial
        Select Case more ' more is button caption
         Case ""
           msg = "Empty button. Click left to configure"
         Case Else
           ' moremore is object Tag
           msg = "Phone: " & moremore & ". Click right to reconfigure"
        End Select
     Case "Combo1" 'combo
        msg = "Memorizes dialed numbers. See pull down"
     Case "StatusBar1" 'help
        msg = "This is the help window..."
     Case "Frame3" 'under menu frame
        msg = "Menu"
     Case "Text1" 'fast dial configure (Dialog1)
        msg = "Insert Name"
     Case "Text2" 'fast dial configure (Dialog1)
        msg = "Insert phone Number"
     Case "CancelButton" 'fast dial configure (Dialog1)
        msg = "Cancel"
     Case "OKButton" 'fast dial configure (Dialog1)
        msg = "OK"
     Case "PrintCommand" 'print flowers (frmAbout)
        msg = "Print current size image (resize to enlarge) on default printer"
     Case "SaveCommand" 'print flowers (frmAbout)
        msg = "Save full size image as c:\Telephonic Connection About.bmp"
     Case "frmAbout" 'form About
        msg = "Thanks for using 'Tel. Connection'. Please rate it on Planet"
     Case "Pietro_Cecchi" 'email address (frmAbout)
        msg = "Click here to send a message to the Author"
     Case Else 'any other object can clear help line
        msg = ""
    End Select

  
  Form1.StatusBar1.Panels(1).Text = msg
End Sub

