Attribute VB_Name = "modUDL"
' Written / modified by Reg Besseling regbes@hotmail.com
' thanks to
' Andreas Hofmann
' http://www.planet-source-code.com/vb/authors/ShowBio.asp?lngAuthorId=195700078&lngWId=1

' www.shoutsoft.com

' Feel free to use this wehever/whenever you like just credit everybody listed above

' require refrences to the following
' 1.Microsoft Scripting Runtime
' 2.OLE automation
' 3.Microsoft OLE DB servicses component
' 4.Microsoft Active X data Objects

' you will have to add your own errortrapping to this
' if you have lots of energy you can fix the comments

Option Explicit
Option Base 0
Option Compare Text

Public Const udlClickCancel As Integer = 1
Public Const udlBadFile As Integer = 2
Public Const udlConnectFail As Integer = 4
Public Const udlNotSaved As Integer = 8

Public Const udlClickOk As Integer = 16
Public Const udlGoodFile As Integer = 32
Public Const udlConnectPass As Integer = 64
Public Const udlSaved As Integer = 128
'--------------------
'Private Members
'--------------------

Private Const CON_UDL_LINE1 = "[oledb]"
Private Const CON_UDL_LINE2 = "; Everything after this line is an OLE DB initstring"


Function SaveUDL(sUDLFIleName As String, sUDLConnectionString As String) As Boolean
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       03 Jul 2004
' Purpose     :       Saves a connection string into a UDL file
' Assumptions :       Valid connection string is passed
' Effects     :       Overwrites the file or creates it
' Arguments   :       ---Inputs/Outputs into the procedure arguments byval
' Results     :       True or false for success or failure
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' Errortrapping of your preference required
  '--------------
  'Define variables and objects
  '--------------
  Dim FSO As Scripting.FileSystemObject
  Dim TXT As Scripting.TextStream
  '--------------
  'Create Objects
  '--------------
  Set FSO = New Scripting.FileSystemObject
  '--------
  'Defaults
  '--------
  SaveUDL = False
  '---------------------
  'Do the Work
  '---------------------
  ' Create a File in Unicode-Mode
  Set TXT = FSO.CreateTextFile(sUDLFIleName, True, True)
  TXT.WriteLine CON_UDL_LINE1
  TXT.WriteLine CON_UDL_LINE2
  TXT.WriteLine sUDLConnectionString
  TXT.Close
  SaveUDL = True
  
ProcedureClose:
  '--------------------------
  'Set Object to nothing here
  '--------------------------
  Set TXT = Nothing
  Set FSO = Nothing

End Function

Function UDLFileExsists(sUDLFIleName As String, NewOrExsists As String, _
                          Extension As String) As Boolean
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       03 Jul 2004
' Purpose     :       ---What the procedure does (not how).
' Assumptions :       ---Expected before calling this procedure
' Effects     :       ---List of effected controls, not passed as arguments
' Arguments   :       ---Inputs/Outputs into the procedure arguments byval
' Results     :       ---Explanation of results passed back if this is a function
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' Errortrapping of your preference required
  '--------------
  'Define variables and objects
  '--------------
  '--------------
  'Create Objects
  '--------------
  '--------
  'Defaults
  '--------
  UDLFileExsists = False
  '---------------------
  'Do the Work
  '---------------------
  ' first check if the UDL file exsists and has a udl extension
  Select Case NewOrExsists
    Case "Exsists"
    
    If Dir(sUDLFIleName) = vbNullString Or UCase(Right(sUDLFIleName, 3)) <> UCase(Extension) Then
      UDLFileExsists = False
    Else
      UDLFileExsists = True
    End If
    
    Case "New"
  
    If UCase(Right(sUDLFIleName, 3)) <> UCase(Extension) _
    Or Dir(Left$(sUDLFIleName, Len(sUDLFIleName) - InStr(1, StrReverse(sUDLFIleName), "\") + 1) _
    , vbDirectory) = vbNullString Then
      UDLFileExsists = False
    Else
      UDLFileExsists = True
    End If
    
  End Select
  
ProcedureClose:
  '--------------------------
  'Set Object to nothing here
  '--------------------------

End Function

Function TestUDL(sUDLFIleName As String) As Integer
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       29 Jun 2004
' Purpose     :       Checks if a UDL File EXSISTS and if it does
' Assumptions :       Requires a Valid Filename as input
' Effects     :       ---List of effected controls, not passed as arguments
' Arguments   :       Requires a Valid Filename as input
' Results     :       passes back a result code as defined in the global constants
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' Errortrapping of your preference required
  '--------------
  'Define variables and objects
  '--------------
  Dim connUDL As ADODB.Connection
  
  On Error GoTo ErrHandler
  '--------------
  'Create Objects
  '--------------
  Set connUDL = New ADODB.Connection
  '--------
  'Defaults
  '--------
  TestUDL = 0
  '---------------------
  'Do the Work
  '---------------------
  ' first check if the UDL file exsists and has a udl extension
  If UDLFileExsists(sUDLFIleName, "Exsists", "UDL") Then
    TestUDL = TestUDL + udlGoodFile
  Else
    TestUDL = TestUDL + udlBadFile
    GoTo ProcedureClose
  End If
  
  connUDL.Open "File Name=" & sUDLFIleName
  TestUDL = TestUDL + udlConnectPass
  
  GoTo ProcedureClose
ErrHandler:
  ' ADODB errors are expected and must not be trapped
  If Err.Source = "Microsoft OLE DB Provider for ODBC Drivers" Then
    Err.Clear
    TestUDL = TestUDL + udlConnectFail
  End If
  
ProcedureClose:
  '--------------------------
  'Set Object to nothing here
  '--------------------------
  Set connUDL = Nothing
End Function

Function CreateUDL(sUDLFIleName As String) As Integer
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       29 Jun 2004
' Purpose     :       Create a Correctly populated UDL File
' Assumptions :        Requires a Valid UDL Filename as input and DIr that Exsists
' Effects     :       Creates a UDL file with the name passed or overwrites exsisting file
' Arguments   :       A valid Filename
' Results     :       passes back a result code as defined in the global constants
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' Errortrapping of your preference required
  '--------------
  'Define variables and objects
  '--------------
  Dim DataLink As DataLinks
  Dim connUDL As Connection
  '--------------
  'Create Objects
  '--------------
  Set DataLink = New DataLinks
  Set connUDL = New ADODB.Connection
  '--------
  'Defaults
  '--------
  CreateUDL = 0
  '---------------------
  'Do the Work
  '---------------------
  'Check that the extension is "UDL" and the directory exsits
  
  If UDLFileExsists(sUDLFIleName, "New", "udl") Then
    CreateUDL = CreateUDL + udlGoodFile
  Else
    CreateUDL = CreateUDL + udlBadFile
    GoTo ProcedureClose
  End If
    
    ' should use prompt new but sets the conn to nothing if cancle clicked
    'problem if cancle clicked nothing is returned how do we test for nothing
    ' so we use datalinks.promptedit not the best but ok
    
  If DataLink.PromptEdit(connUDL) Then
    ' clicked ok
    CreateUDL = CreateUDL + udlClickOk
    
    If SaveUDL(sUDLFIleName, connUDL.ConnectionString) Then
      CreateUDL = CreateUDL + udlSaved
    Else
      CreateUDL = CreateUDL + udlNotSaved
    End If
    
  Else
    'clicked cancel
    CreateUDL = CreateUDL + udlClickCancel
  End If
  
ProcedureClose:
  '--------------------------
  'Set Object to nothing here
  '--------------------------
  Set connUDL = Nothing
  Set DataLink = Nothing
End Function

Function EditUDL(sUDLFIleName As String) As Integer
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       29 Jun 2004
' Purpose     :       Shows the Data Link Properties dialog populated
'                     with details from a UDL file and allows you to save changes
' Assumptions :       That the file is a correctly populated UDL file
' Effects     :       ---List of effected controls, not passed as arguments
' Arguments   :       A Valid UDL File Name
' Results     :       See constant declarations above for return codes
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' Wa can combine View and change into one proc but im a bit tired now
' Errortrapping of your preference required

  '--------------
  'Define variables and objects
  '--------------
  Dim connUDL As ADODB.Connection
  Dim DataLink As DataLinks
  '--------------
  'Create Objects
  '--------------
  Set connUDL = New ADODB.Connection
  Set DataLink = New DataLinks
  '--------
  'Defaults
  '--------
  EditUDL = 0
  '---------------------
  'Do the Work
  '---------------------
  ' first check if the UDL file exsists and it has a UDL extension
  If UDLFileExsists(sUDLFIleName, "Exsists", "UDL") Then
    EditUDL = EditUDL + udlGoodFile
  Else
    EditUDL = EditUDL + udlBadFile
    GoTo ProcedureClose
  End If
  
  ' get the connection string fropm the file
  connUDL.ConnectionString = ConnectionStringFromFile(sUDLFIleName)
  
  ' show the dialog box and get the new string back
  ' show the udl edit form with datalinks.promptedit
  If DataLink.PromptEdit(connUDL) Then
    ' ok clicked
    EditUDL = EditUDL + udlClickOk
    
    If SaveUDL(sUDLFIleName, connUDL.ConnectionString) Then
      EditUDL = EditUDL + udlSaved
    Else
      EditUDL = EditUDL + udlNotSaved
    End If
    
  Else
    'cancel pressed
    EditUDL = EditUDL + udlClickCancel
  End If
 

ProcedureClose:
  '--------------------------
  'Set Object to nothing here
  '--------------------------
  Set connUDL = Nothing
  Set DataLink = Nothing

End Function

Function ViewUDL(sUDLFIleName As String) As Integer
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       29 Jun 2004
' Purpose     :       Shows the Data Link Properties dialog populated
'                     with details from a UDL file
' Assumptions :       That the file is a correctly populated UDL file
' Effects     :       ---List of effected controls, not passed as arguments
' Arguments   :       A Valid UDL File Name
' Results     :       See constant declarations above for return codes
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' We can combine View and change into one proc but im a bit lazy now
' Errortrapping of your preference required
    '--------------
    'Define variables and objects
    '--------------
Dim connUDL As ADODB.Connection
Dim DataLink As DataLinks
    '--------------
    'Create Objects
    '--------------
Set connUDL = New ADODB.Connection
Set DataLink = New DataLinks
    '--------
    'Defaults
    '--------
ViewUDL = 0
    '---------------------
    'Do the Work
    '---------------------
' first check if the UDL file exsists and it has a UDL extension
If UDLFileExsists(sUDLFIleName, "Exsists", "UDL") Then
  ViewUDL = ViewUDL + udlGoodFile
Else
  ViewUDL = ViewUDL + udlBadFile
  GoTo ProcedureClose
End If

' get the connection string fropm the file
connUDL.ConnectionString = ConnectionStringFromFile(sUDLFIleName)
' show the udl edit form with datalinks.promptedit
If DataLink.PromptEdit(connUDL) Then
  ViewUDL = ViewUDL + udlClickOk
Else
  ViewUDL = ViewUDL + udlClickCancel
End If

ProcedureClose:
    '--------------------------
    'Set Object to nothing here
    '--------------------------

    Set connUDL = Nothing
    Set DataLink = Nothing


End Function

Public Function ConnectionStringFromFile(ByVal sUDLFIleName As String) As String
'--------------------------------------------------------------------------------
' Author      :       Reg Besseling regbes@hotmail.com
' Date        :       03 Jul 2004
' Purpose     :       extract a connection string from a UDL file and convert it from unicode
' Assumptions :       That all UDL Files connection strings srart with "Provider="
' Effects     :       ---List of effected controls, not passed as arguments
' Arguments   :       A valid udl file name
' Results     :       An ADODB connection string converted from unicode
'--------------------------------------------------------------------------------
'                 Version Information May Appear Below Here
' Errortrapping of your preference required

  '--------------
  'Define variables and objects
  '--------------
  Dim iFileNum As Integer
  Dim sUDLFileContents As String
  Dim iLocation As Integer
  '--------------
  'Create Objects
  '--------------
  '--------
  'Defaults
  '--------
  ConnectionStringFromFile = ""
  '---------------------
  'Do the Work
  '---------------------
  iFileNum = FreeFile
  Open sUDLFIleName For Binary As #iFileNum
  sUDLFileContents = Input(LOF(iFileNum), #iFileNum)
  Close #iFileNum
  sUDLFileContents = StrConv(sUDLFileContents, vbFromUnicode)
  iLocation = InStr(sUDLFileContents, "Provider=")
  If iLocation > 0 Then
    ConnectionStringFromFile = Mid(sUDLFileContents, iLocation)
  End If
  
ProcedureClose:
  '--------------------------
  'Set Object to nothing here
  '--------------------------
End Function
 
