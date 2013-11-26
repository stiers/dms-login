VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1965
   ClientLeft      =   9030
   ClientTop       =   5385
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   1650
      TabIndex        =   1
      Top             =   315
      Width           =   2505
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   270
      TabIndex        =   4
      Top             =   1425
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3180
      TabIndex        =   5
      Top             =   1425
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1650
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   795
      Width           =   2505
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   270
      Top             =   180
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   1095
      Left            =   285
      Top             =   195
      Width           =   4020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&User Name:"
      Height          =   195
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   375
      Width           =   840
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Index           =   1
      Left            =   465
      TabIndex        =   2
      Top             =   855
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUsername As String
Private mPassword As String
Private mCancel As Boolean

Private Sub cmdCancel_Click()
    mCancel = True
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mCancel = False
    mUsername = Me.txtUserName.Text
    mPassword = Me.txtPassword.Text
    
    Unload Me
End Sub

Public Function GetLogIn(ByRef UserName As String, ByRef Password As String, Owner As Object) As Boolean
    Me.txtUserName.Text = UserName
    
    Me.Show vbModal, Owner
    
    UserName = mUsername
    Password = mPassword
    
    GetLogIn = Not mCancel
End Function

Private Sub Form_Activate()
    If Len(Me.txtUserName.Text) > 0 Then Me.txtPassword.SetFocus
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
