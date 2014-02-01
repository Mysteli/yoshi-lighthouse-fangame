VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4200
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   150
      Picture         =   "frmSendGetData.frx":3332
      ScaleHeight     =   975
      ScaleWidth      =   3900
      TabIndex        =   1
      Top             =   180
      Width           =   3900
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   1380
      Width           =   4200
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Me.Caption = Options.Game_Name & " (esc to cancel)"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    If KeyAscii = vbKeyEscape Then
        Call DestroyTCP
        frmLoad.Hide
        frmMenu.Show
        frmMenu.picMain.visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' When the form close button is pressed
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Call DestroyTCP
        frmLoad.Hide
        frmMenu.Show
        frmMenu.picMain.visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
