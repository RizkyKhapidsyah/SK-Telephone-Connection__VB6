VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fast Dialing"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3915
   Icon            =   "Dialog1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1515
      Top             =   705
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Fill (or blank) both fields.   Otherwise the OK button will stay disabled."
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Number to dial:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Input name and number to be associated to this button."
      Height          =   795
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
 Unload Me
End Sub

Private Sub CancelButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "CancelButton"
End Sub

Private Sub Form_Load()
 With Form1
  Top = .Top + .Height - .ScaleHeight
  Left = .Left + .Width - .ScaleWidth
 End With
 Label20.Caption = "Fill (or blank) both fields." & vbNewLine & "Otherwise the OK button will stay disabled."
 Label20.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "Dialog1"
End Sub

Private Sub Form_Resize()
 Caption = "Fast dial, button " & IndexSave + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Form1.Enabled = True 'reenables controls on Form1
                      'that were disabled when this
                      'form was called up
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "Label10"
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "Text1"
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "Text2"
End Sub

Private Sub OKButton_Click()
 SaveSetting App.ProductName, "FastDialName", IndexSave, Text1.Text
 SaveSetting App.ProductName, "FastDialNumber", IndexSave, Text2.Text
 Form1.Command2(IndexSave).Caption = Text1.Text 'name
 Form1.Command2(IndexSave).Tag = Text2.Text     'TF nr.
 Unload Me
End Sub

Private Sub OKButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Help "OKButton"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Text1"
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Help "Text2"
End Sub

Private Sub Timer1_Timer() 'always running
 'monitors the inputs
 Dim condition
 condition = ((Text1.Text = "") And (Text2.Text <> "")) Or ((Text2.Text = "") And (Text1.Text <> ""))
 If OKButton.Enabled <> Not condition Then OKButton.Enabled = Not condition
 If Label20.Visible <> condition Then Label20.Visible = condition
End Sub
