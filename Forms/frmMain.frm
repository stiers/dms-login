VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Login Example"
   ClientHeight    =   4410
   ClientLeft      =   5070
   ClientTop       =   4080
   ClientWidth     =   10020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   10020
   Begin VB.CommandButton cmdDataEntry 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6660
      TabIndex        =   3
      Top             =   675
      Width           =   2130
   End
   Begin VB.CommandButton cmdCSR 
      Caption         =   "Customer Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3945
      TabIndex        =   2
      Top             =   675
      Width           =   2130
   End
   Begin VB.CommandButton cmdManageUsers 
      Caption         =   "Manage Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1230
      TabIndex        =   1
      Top             =   675
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3330
      TabIndex        =   0
      Top             =   2655
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdCSR_Click()
    frmCSR.Show vbModal, Me
End Sub

Private Sub cmdDataEntry_Click()
    frmDataEntry.Show vbModal, Me
End Sub

Private Sub cmdManageUsers_Click()
    frmManageUsers.Show vbModal, Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    GetLocation Me
    ResizeForm Me

    'Load the database
    Set DBConn = LoadDatabase(App.Path & "\dbGBC.mdb")
    
    'Do a count of records to see if there any users in the table
    Set rsData = DBConn.Execute("SELECT Count(*) FROM tblUsers")
    
    'If there are users in the table, then prompt for login
    If rsData.Fields(0).Value > 0 Then
        If Not DoLogin Then
            Unload Me 'Login unsuccessful, so unload
            Exit Sub
        End If
    End If
    
    Me.cmdCSR.Enabled = FormAllowed("frmCSR")
    Me.cmdDataEntry.Enabled = FormAllowed("frmDataEntry")
End Sub

Private Function DoLogin() As Boolean
    Dim UserName As String, Password As String, ret As Boolean
    Dim LoginSuccessful As Boolean
    Dim MD5 As New clsMD5
    
    Randomize
    
    'Get the user that last logged in from the registry
    UserName = GetSetting(App.EXEName, "Settings", "LastLogIn", "")
    
    'Prompt user to enter username and password
    ret = frmLogin.GetLogIn(UserName, Password, Me)
    
    Do While ret
        Set rsData = DBConn.Execute("SELECT [ID], [UserName], [Password], [UserType], IIF(IsNull([UserGroupID]), -1, [UserGroupID]) AS UserGroup FROM tblUsers WHERE UserName = '" & Replace(UserName, "'", "''") & "'")
        
        If rsData.RecordCount = 0 Then GoTo Bye
        
        'If a record was found, it means the user exists
        If rsData("UserGroup").Value <> -1 Or rsData("UserType").Value = 2 Then
            'Check if the password is correct
            If UCase(MD5.DigestStrToHexStr(Password)) = UCase(rsData("Password").Value) Then
                
                'Password is correct, so save the user that just logged in
                LogInUserID = rsData("ID").Value
                LogInUserName = rsData("UserName").Value
                LogInUserType = rsData("UserType").Value
                
                'Save the username in the registry
                SaveSetting App.EXEName, "Settings", "LastLogIn", rsData("UserName").Value
                
                LoginSuccessful = True
                Exit Do
            End If
        End If
        
Bye:    If Not LoginSuccessful Then
            ret = False
            
            If MsgBox("Invalid login, do you want to try again ?", vbQuestion + vbYesNo, "Invalid Login") = vbYes Then
                'To prevent brute force password cracking from the application
                Sleep 200 + 300 * Rnd
                
                'If login was not successfull, prompt again until Cancel is clicked
                ret = frmLogin.GetLogIn(UserName, Password, Me)
            End If
        End If
    Loop
    
    DoLogin = LoginSuccessful
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not (DBConn Is Nothing) Then
        DBConn.Close
        Set DBConn = Nothing
    End If
End Sub

Private Sub Form_Resize()
    ResizeControls Me
End Sub
