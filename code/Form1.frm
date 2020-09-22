VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Creating Controls at Run-Time"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Dummy"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dummy"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Make sure if you wanna duplicate something
' like this that you need to set the control's
' index to 0.
'
' http://www.thegtproject.com

Private Sub Command1_Click(Index As Integer)
Call MsgBox("You click button #: " & Index & "", vbOKOnly, "Creating Controls")
End Sub

Private Sub Form_Load()
' This is just testing the subs
' Note: YOU CAN LOAD AS MANY AS YOU WANT!
Call CreateButton("Hello", 30, 30, "500", "1000")
Call CreateLabel("Hello", "30", "700", "250", "1000")
Call CreateButton("Hello", 30, "1000", "500", "1000")
Call CreateLabel("Hello", "30", "1600", "250", "1000")
End Sub

Sub CreateButton(Cptn As String, X As Integer, Y As Integer, HT As Integer, WT As Integer)
' Im using a sub just for easier use.
'You dont need a sub to do this so
'you can create controls without a sub!

Dim CurIndex As Integer
' This is to set curindex as the highest
'loaded index of command1 + 1.
CurIndex = Command1.UBound + 1
' Now we are gonna load the new control
'with the curindex value
Load Command1(CurIndex)
' Now to set up all the preferences:
Command1(CurIndex).Caption = Cptn
Command1(CurIndex).Top = Y
Command1(CurIndex).Left = X
Command1(CurIndex).Width = WT
Command1(CurIndex).Height = HT
' And finally make the new control visible!
Command1(CurIndex).Visible = True
End Sub

Sub CreateLabel(Cptn As String, X As Integer, Y As Integer, HT As Integer, WT As Integer)
' Im using a sub just for easier use.
'You dont need a sub to do this so
'you can create controls without a sub!

Dim CurIndex As Integer

' This is to set curindex as the highest
'loaded index of label1 + 1.
CurIndex = Label1.UBound + 1
' Now we are gonna load the new control
'with the curindex value
Load Label1(CurIndex)
' Now to set up all the preferences:
Label1(CurIndex).Caption = Cptn
Label1(CurIndex).Left = X
Label1(CurIndex).Top = Y
Label1(CurIndex).Height = HT
Label1(CurIndex).Width = WT
' And finally make the new control visible!
Label1(CurIndex).Visible = True
End Sub
