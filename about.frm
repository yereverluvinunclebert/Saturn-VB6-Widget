VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   0  'None
   Caption         =   "About SteamyDock"
   ClientHeight    =   9615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11385
   ControlBox      =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAbout 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      Picture         =   "about.frx":058A
      ScaleHeight     =   641
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   761
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    
        ' display the version number on the general panel
    'Call displayVersionNumber
    
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Hide
End Sub


'---------------------------------------------------------------------------------------
' Procedure : displayVersionNumber
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub displayVersionNumber()
   On Error GoTo displayVersionNumber_Error
   ''If debugflg = 1 Then msgBox "%displayVersionNumber"

'     about.lblMajorVersion.Caption = App.Major
'     about.lblMinorVersion.Caption = App.Minor
'     about.lblRevisionNum.Caption = App.Revision

   On Error GoTo 0
   Exit Sub

displayVersionNumber_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayVersionNumber of Form dock"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : lblPunklabsLink_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblPunklabsLink_Click()
   On Error GoTo lblPunklabsLink_Click_Error
   ''If debugflg = 1 Then msgBox "%lblPunklabsLink_Click"

        'Call ShellExecute(Me.hWnd, "Open", "http://www.punklabs.com", vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

lblPunklabsLink_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblPunklabsLink_Click of Form about"

End Sub

Private Sub picAbout_Click()
    Me.Hide
End Sub
