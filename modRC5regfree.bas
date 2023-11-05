Attribute VB_Name = "Module3"
'---------------------------------------------------------------------------------------
' Module    : Module3
' Author    : Olaf Schmidt
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

'The following two Declares are used, to ensure regfree mode when running compiled...
'For that to work, the App-Zip will need a \Bin\-Folder besides the compiled Executable, which should contain:
'an up-to-date copy of the 3 Base-Dlls: DirectCOM.dll, vbRichClient5.dll and vb_cairo_sqlite.dll ...
'usually copied from "C:\RC5\", where the registered version of the RC5 should reside on your Dev-Machine
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetInstanceEx Lib "DirectCOM" (spFName As Long, spClassName As Long, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object

Const SubDir As String = "\BIN\" ' <- the usual "\Bin\" subdirectory (placed beside the compiled Executable)

'---------------------------------------------------------------------------------------
' Property  : New_c
' Author    : Olaf Schmidt
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get New_c() As cConstructor
  Static oConstr As cConstructor
   On Error GoTo New_c_Error

  If Not oConstr Is Nothing Then Set New_c = oConstr: Exit Property
  
  If App.LogMode Then 'we run compiled - and ensure that we load the "Entry-Object" of the RC5 regfree
    LoadLibraryW StrPtr(App.Path & SubDir & "\DirectCOM.dll")
    Set oConstr = GetInstanceEx(StrPtr(App.Path & SubDir & "vbRichClient5.dll"), StrPtr("cConstructor"))
  Else 'we run in the IDE, and instantiate from a registered version of the RC5 (usually placed on: "C:\RC5\.." on your Dev-Machine)
    Set oConstr = CreateObject("vbRichClient5.cConstructor") 'load the Constructor-instance from the registered version
  End If
  Set New_c = oConstr

   On Error GoTo 0
   Exit Property

New_c_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure New_c of Module Module3"
End Property

'---------------------------------------------------------------------------------------
' Property  : Cairo
' Author    : Olaf Schmidt
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Cairo() As cCairo
  Static oCairo As cCairo
   On Error GoTo Cairo_Error

  If oCairo Is Nothing Then Set oCairo = New_c.Cairo 'ensure the static on the first run
  Set Cairo = oCairo

   On Error GoTo 0
   Exit Property

Cairo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Cairo of Module Module3"
End Property
