VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00B0C8A8&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   2880
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   6570
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00B0C8A8&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   4800
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   1680
         Width           =   480
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   330
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H00B0C8A8&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   1935
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H00B0C8A8&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Left            =   3360
         MaskColor       =   &H00C8D8C0&
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtCName 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   225
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblSprite 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ Change Sprite ]"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   210
         Left            =   2280
         TabIndex        =   25
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Character:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblCAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.PictureBox picRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00B0C8A8&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   2880
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   6570
      Begin VB.TextBox txtRPass2 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "�"
         TabIndex        =   13
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtRPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "�"
         TabIndex        =   10
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtRUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   225
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label txtRAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.PictureBox picLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00B0C8A8&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   2880
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   6570
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00B0C8A8&
         Caption         =   "Save Password?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtLPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   20
         PasswordChar    =   "�"
         TabIndex        =   3
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtLUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00A0B898&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00403028&
         Height          =   225
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   1
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblLAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00B0C8A8&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   2880
      ScaleHeight     =   3195
      ScaleWidth      =   6570
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   6570
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00B0C8A8&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   2880
      ScaleHeight     =   3195
      ScaleWidth      =   6570
      TabIndex        =   27
      Top             =   2760
      Width           =   6570
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is an example of the news. Not very exciting, I know, but it's better than nothing, amirite? "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00909058&
         Height          =   1575
         Left            =   1800
         TabIndex        =   28
         Top             =   1080
         Width           =   3135
      End
   End
   Begin VB.Image imgButton 
      Height          =   495
      Index           =   4
      Left            =   7755
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   495
      Index           =   3
      Left            =   6255
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   495
      Index           =   2
      Left            =   4755
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   495
      Index           =   1
      Left            =   3255
      Top             =   6480
      Width           =   1335
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClass_Click()
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
End Sub

Private Sub Form_Load()
    Dim tmpTxt As String, tmpArray() As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' general menu stuff
    Me.Caption = Options.Game_Name
    
    ' load news
    Open App.Path & "\data files\news.txt" For Input As #1
        Line Input #1, tmpTxt
    Close #1
    ' split breaks
    tmpArray() = Split(tmpTxt, "<br />")
    lblNews.Caption = vbNullString
    For i = 0 To UBound(tmpArray)
        lblNews.Caption = lblNews.Caption & tmpArray(i) & vbNewLine
    Next

    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    If Options.savePass = 1 Then
        txtLPass.text = Trim$(Options.Password)
        chkPass.value = Options.savePass
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not EnteringGame Then DestroyGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            If Not picLogin.visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.visible = False
                picLogin.visible = True
                picRegister.visible = False
                picCharacter.visible = False
                picMain.visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        Case 2
            If Not picRegister.visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.visible = False
                picLogin.visible = False
                picRegister.visible = True
                picCharacter.visible = False
                picMain.visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        Case 3
            If Not picCredits.visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.visible = True
                picLogin.visible = False
                picRegister.visible = False
                picCharacter.visible = False
                picMain.visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        Case 4
            Call DestroyGame
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    changeButtonState_Menu Index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    If Not MenuButton(Index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Menu Index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Menu = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Menu = Index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Menu -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub lblSprite_Click()
Dim spritecount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If optMale.value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblSprite_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCredits_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Register
Private Sub txtRAccept_Click()
    Dim name As String
    Dim Password As String
    Dim PasswordAgain As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' New Char
Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MenuState(MENU_STATE_ADDCHAR)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
