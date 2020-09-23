VERSION 5.00
Begin VB.Form frmUDL 
   Caption         =   "UDL Thinger"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "c:\MyUDL.udl"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "UDL File Name:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Basic UDL Usage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1515
   End
End
Attribute VB_Name = "frmUDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' some notes

''Using an UDL file.
''Create UDL Manualy like this or use the app.
''Create a new textfile on your desktop and rename it as 'MyFirstUDL.udl'.
''Having done that you'll see that the icon has changed.
''Double-click You can setup a data connection to any database and test it
''
''You can reference a UDL file from your code like this:
''
''Dim cnn As ADODB.Connection
''Set cnn = New ADODB.Connection
''cnn.Open "File Name=C:\MyFirstUDL.udl"
''
''What 's the use?
''You can add this file to your setup-project.
''So if your program has to change dynamically to another database
''you just have to call the UDL file,
''make the proper adjustments via the UDL-interface
''and reconnect to the database via the UDL.
Option Explicit


Private Sub cmdCreate_Click()

MsgBox ResultMessage(CreateUDL(txtFileName.Text))

End Sub

Private Sub cmdEdit_Click()
    
MsgBox ResultMessage(EditUDL(txtFileName.Text))

End Sub

Private Sub cmdTest_Click()

MsgBox ResultMessage(TestUDL(txtFileName.Text))

End Sub

Private Sub cmdView_Click()

MsgBox ResultMessage(ViewUDL(txtFileName.Text))

End Sub
Function ResultMessage(iResult As Integer) As String

If (iResult And udlClickCancel) = udlClickCancel Then: ResultMessage = " Cancel Clicked "
If (iResult And udlBadFile) = udlBadFile Then: ResultMessage = ResultMessage & " File Problem "
If (iResult And udlConnectFail) = udlConnectFail Then: ResultMessage = ResultMessage & " Connection Faied "
If (iResult And udlNotSaved) = udlNotSaved Then: ResultMessage = ResultMessage & " File Not Saved "

If (iResult And udlClickOk) = udlClickOk Then: ResultMessage = ResultMessage & " Ok Clicked "
If (iResult And udlGoodFile) = udlGoodFile Then: ResultMessage = ResultMessage & " File Ok "
If (iResult And udlConnectPass) = udlConnectPass Then: ResultMessage = ResultMessage & " Connection Passed "
If (iResult And udlSaved) = udlSaved Then: ResultMessage = ResultMessage & " UDL Saved "

End Function

