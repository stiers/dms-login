VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Users"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdManageGroups 
      Caption         =   "&Manage Groups"
      Height          =   420
      Left            =   2145
      TabIndex        =   16
      Top             =   5085
      Width           =   1455
   End
   Begin VB.OptionButton optUserType 
      Caption         =   "USER"
      Height          =   195
      Index           =   1
      Left            =   1395
      TabIndex        =   5
      Top             =   4140
      Width           =   1050
   End
   Begin VB.OptionButton optUserType 
      Caption         =   "ADMIN"
      Height          =   195
      Index           =   2
      Left            =   2745
      TabIndex        =   6
      Top             =   4140
      Width           =   1050
   End
   Begin VB.ComboBox cmbUserGroup 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4500
      Width           =   4110
   End
   Begin VB.TextBox txtOldPassword 
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2685
      Width           =   4110
   End
   Begin VB.TextBox txtPassword2 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3645
      Width           =   4110
   End
   Begin VB.TextBox txtPassword1 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3165
      Width           =   4110
   End
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Top             =   2205
      Width           =   4110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   180
      TabIndex        =   9
      Top             =   5085
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4050
      TabIndex        =   8
      Top             =   5085
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   1860
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   3281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   4140
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group:"
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   4560
      Width           =   480
   End
   Begin VB.Label lblOldPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   2760
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmation Password:"
      Height          =   420
      Left            =   180
      TabIndex        =   12
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   2280
      Width           =   840
   End
End
Attribute VB_Name = "frmManageUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'List all the users in the ListView
Private Sub ListUsers()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DBConn.Execute("SELECT ID, UserName, UserType, GroupName FROM tblUsers AS u LEFT JOIN tblUserGroup AS ug ON u.UserGroupID = ug.GroupID")
    
    lstUsers.ListItems.Clear
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        Do Until rsData.EOF Or rsData.BOF
            With lstUsers.ListItems.Add(, , rsData("ID").Value & "")
                .SubItems(1) = rsData("UserName").Value & ""
                .SubItems(2) = IIf(Val(rsData("UserType").Value & "") = 1, "USER", "ADMIN")
                .SubItems(3) = rsData("GroupName").Value & ""
            End With
            
            rsData.MoveNext
        Loop
    End If
End Sub

Private Sub ListGroups()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DBConn.Execute("SELECT GroupID, GroupName FROM tblUserGroup ORDER BY GroupName")
    
    cmbUserGroup.Clear
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        Do Until rsData.EOF Or rsData.BOF
            cmbUserGroup.AddItem rsData("GroupName").Value & ""
            cmbUserGroup.ItemData(cmbUserGroup.NewIndex) = rsData("GroupID").Value
            
            rsData.MoveNext
        Loop
    End If
    
    cmbUserGroup.Enabled = cmbUserGroup.ListCount > 0 And LogInUserType = 2 And optUserType(1).Value
End Sub

Private Sub cmdAdd_Click()
    Dim Request As String, rsData As Recordset, GroupID As Long
    Dim MD5 As New clsMD5, NewPassword As String, OldPassword As String
    
    If Len(Trim(Me.txtUserName.Text)) = 0 Then
        MsgBox "You must enter a User Name.", vbInformation
        Exit Sub
    End If
    
    If Len(txtPassword1.Text) > 0 Or Len(txtPassword2.Text) > 0 Then
        If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "Confirm password must the same as Password field", vbExclamation
            Exit Sub
        End If
    End If
    
    If Not optUserType(1).Value And Not optUserType(2).Value Then
        optUserType(2).Value = (LogInUserID = 0 And Len(LogInUserName) = 0 And LogInUserType = 0) And _
            DBConn.Execute("SELECT Count(*) FROM tblUsers").Fields(0).Value = 0
    End If
    
    If (Not optUserType(2).Value) And ((LogInUserID = 0 And Len(LogInUserName) = 0 And LogInUserType = 0) Or _
            DBConn.Execute("SELECT Count(*) FROM tblUsers").Fields(0).Value = 0) Then
        MsgBox "First user added must be of type ADMIN.", vbInformation
        Exit Sub
    End If
    
    If LogInUserType = 1 And optUserType(2).Value Then
        MsgBox "Only an ADMIN can add another ADMIN.", vbInformation
        Exit Sub
    End If
    
    If optUserType(1).Value Then
        If cmbUserGroup.ListIndex = -1 Then
            MsgBox "You must select a group for this user.", vbInformation
            Exit Sub
        Else
            GroupID = cmbUserGroup.ItemData(cmbUserGroup.ListIndex)
        End If
    End If
    
    'Get the hash of the passwords
    NewPassword = UCase(MD5.DigestStrToHexStr(Me.txtPassword1.Text))
    OldPassword = UCase(MD5.DigestStrToHexStr(Me.txtOldPassword.Text))
    
    If cmdAdd.Caption = "&Add User" Then
        If DBConn.Execute("SELECT Count(*) FROM tblUsers WHERE [UserName] = '" & Trim(Me.txtUserName.Text) & "'").Fields(0).Value > 0 Then
            MsgBox "A user with this User Name already exists, please choose another User Name.", vbInformation
            Exit Sub
        End If
        
        'Prepare the INSERT statement
        Request = "INSERT INTO tblUsers ([UserName], [Password], [UserType], [UserGroupID]) VALUES(" & _
            "'" & Replace(Trim(Me.txtUserName.Text), "'", "''") & "'," & _
            "'" & NewPassword & "'," & _
            IIf(optUserType(2).Value, 2, 1) & "," & _
            IIf(optUserType(1).Value, GroupID, "NULL") & ")"
    Else
        If DBConn.Execute("SELECT Count(*) FROM tblUsers WHERE [ID] <> " & Me.lstUsers.SelectedItem.Text & " AND [UserName] = '" & Trim(Me.txtUserName.Text) & "'").Fields(0).Value > 0 Then
            MsgBox "A user with this User Name already exists, please choose another User Name.", vbInformation
            Exit Sub
        End If
        
        If LogInUserType = 2 And optUserType(1).Value Then
            If DBConn.Execute("SELECT Count(*) FROM tblUsers WHERE UserType = 2 AND [ID] <> " & Me.lstUsers.SelectedItem.Text).Fields(0).Value = 0 Then
                MsgBox "At least one ADMIN must exist.", vbInformation
                Exit Sub
            End If
        End If
        
        'If logged in as a diferent user than the one we are changing now
        If LogInUserID <> Val(lstUsers.SelectedItem.Text) Then
            'Validate user password if logged in as a diferent user
            Set rsData = DBConn.Execute("SELECT Password FROM tblUsers WHERE ID = " & Me.lstUsers.SelectedItem.Text)
            If OldPassword <> rsData("Password").Value Then
                MsgBox "Invalid old password." & vbNewLine & "You must enter the valid password for user selected.", vbInformation
                Exit Sub
            End If
        End If
        
        'Prepare the UPDATE statement
        Request = "UPDATE tblUsers SET UserName = '" & Replace(Me.txtUserName.Text, "'", "''") & "'" & _
            ", [UserType] = " & IIf(Me.optUserType(1).Value, 1, 2) & _
            ", [UserGroupID] = " & IIf(optUserType(2).Value, "NULL", GroupID)
        
        If Len(Me.txtPassword1.Text) > 0 Then Request = Request & ", [Password] = '" & NewPassword & "'"
        Request = Request & " WHERE ID = " & Me.lstUsers.SelectedItem.Text
    End If
    
    'Execute the request
    Debug.Print Request
    DBConn.Execute Request
    
    'Reset controls
    ListUsers
    ClearUpdate
End Sub

Private Sub ClearUpdate()
    txtUserName.Text = ""
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtOldPassword.Text = ""
    txtOldPassword.Enabled = False
    lblOldPassword.Enabled = False
    lstUsers.Enabled = True
    Me.cmbUserGroup.ListIndex = -1
    
    optUserType(1).Value = False
    optUserType(2).Value = False
    
    Me.cmdExit.Caption = "&Exit"
    cmdAdd.Caption = "&Add User"
    
    optUserType(1).Enabled = LogInUserType = 2
    optUserType(2).Enabled = LogInUserType = 2
    cmbUserGroup.Enabled = cmbUserGroup.ListCount > 0 And LogInUserType = 2 And optUserType(1).Value
End Sub

Private Sub cmdExit_Click()
    If Me.cmdExit.Caption = "&Exit" Then
        Unload Me
    Else
        ClearUpdate
    End If
End Sub

Private Sub cmdManageGroups_Click()
    frmManageGroups.Show vbModal, Me
    ListGroups
End Sub

Private Sub Form_Load()
    Dim IsFirstTime As Boolean
    
    ListUsers
    ListGroups
    
    IsFirstTime = (LogInUserID = 0 And Len(LogInUserName) = 0 And LogInUserType = 0)
    
    Me.cmdManageGroups.Enabled = (LogInUserType = 2) Or IsFirstTime
    Me.optUserType(1).Enabled = Not IsFirstTime And (LogInUserType = 2)
    Me.optUserType(2).Enabled = Not IsFirstTime And (LogInUserType = 2)
End Sub

Private Sub lstUsers_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lstUsers_DblClick()
    Dim GroupID As Long, K As Long
    
    If Not (lstUsers.SelectedItem Is Nothing) Then
        Me.txtUserName.Text = lstUsers.SelectedItem.SubItems(1)
        optUserType(1).Value = lstUsers.SelectedItem.SubItems(2) = "USER"
        optUserType(2).Value = lstUsers.SelectedItem.SubItems(2) = "ADMIN"
        
        If optUserType(1).Value Then
            GroupID = Val(DBConn.Execute("SELECT UserGroupID FROM tblUsers WHERE [ID] = " & Me.lstUsers.SelectedItem.Text).Fields(0).Value & "")
            
            For K = 0 To Me.cmbUserGroup.ListCount - 1
                If Me.cmbUserGroup.ItemData(K) = GroupID Then
                    Me.cmbUserGroup.ListIndex = K
                    Exit For
                End If
            Next K
        Else
            Me.cmbUserGroup.ListIndex = -1
        End If
        
        optUserType(1).Enabled = LogInUserType = 2
        optUserType(2).Enabled = LogInUserType = 2
        cmbUserGroup.Enabled = cmbUserGroup.ListCount > 0 And LogInUserType = 2 And optUserType(1).Value
        
        Me.lstUsers.Enabled = False
        Me.cmdAdd.Caption = "&Update"
        Me.cmdExit.Caption = "&Cancel"
        
        'Enable the txtOldPassword field when logged in as a diferent user
        If LogInUserID <> Val(lstUsers.SelectedItem.Text) Then
            txtOldPassword.Enabled = True
            lblOldPassword.Enabled = True
        End If
    End If
End Sub

Private Sub optUserType_Click(Index As Integer)
    cmbUserGroup.Enabled = Index = 1 And cmbUserGroup.ListCount > 0 And LogInUserType = 2
End Sub
