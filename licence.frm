VERSION 5.00
Begin VB.Form frmLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Agreement Accept or Decline..."
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "licence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox licencePicture 
      BackColor       =   &H00414D38&
      Height          =   4245
      Left            =   -60
      Picture         =   "licence.frx":1084A
      ScaleHeight     =   4185
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   -15
      Width           =   5115
      Begin VB.TextBox txtLicenceTextBox 
         BackColor       =   &H00414D38&
         ForeColor       =   &H00FFFFFF&
         Height          =   2685
         Left            =   315
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   405
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   4200
         TabIndex        =   4
         ToolTipText     =   "If you accept the program will run"
         Top             =   3795
         Width           =   600
      End
      Begin VB.Label declineLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Decline"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   3375
         TabIndex        =   3
         ToolTipText     =   "If you decline the program will close"
         Top             =   3795
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Left            =   360
         TabIndex        =   2
         Top             =   3180
         Width           =   4380
      End
      Begin VB.Label licenceAgreement 
         BackStyle       =   0  'Transparent
         Caption         =   "Licence Agreement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   315
         TabIndex        =   1
         Top             =   120
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : declineLabel_Click
' Author    : beededea
' Date      : 27/10/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub declineLabel_Click()
    'Dim ofrm As Form
    Dim slicence As String: slicence = "0"
    
    On Error GoTo declineLabel_Click_Error

    MsgBox "Please uninstall and remove Steamydock" & vbCr & "from your computer."

    Call saturnForm_Unload
    
    sPutINISetting "Software\Saturn", "Licence", slicence, StSettingsFile
    End

   On Error GoTo 0
   Exit Sub

declineLabel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure declineLabel_Click of Form licence"
End Sub

Private Sub Form_Load()

    If PrPrefsFont <> vbNullString Then
        Call changeFormFont(frmLicence, PrPrefsFont, Val(PrPrefsFontSize), 0, False, CBool(PrPrefsFontItalics), vbWhite)
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Label2_Click
' Author    : beededea
' Date      : 27/10/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Label2_Click()
    Dim slicence As String: slicence = "0"

    On Error GoTo Label2_Click_Error

    frmLicence.Hide
    slicence = "1"
    
    sPutINISetting "Software\Saturn", "Licence", slicence, StSettingsFile

   On Error GoTo 0
   Exit Sub

Label2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Label2_Click of Form licence"

End Sub

