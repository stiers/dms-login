VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Groups"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddForm 
      Caption         =   "Add Form"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4680
      TabIndex        =   6
      Top             =   4095
      Width           =   1500
   End
   Begin VB.ComboBox cmbForms 
      Enabled         =   0   'False
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4110
      Width           =   3840
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "Add New Group"
      Height          =   330
      Left            =   4635
      TabIndex        =   3
      Top             =   1980
      Width           =   1500
   End
   Begin MSComctlLib.ListView lstForms 
      Height          =   1590
      Left            =   180
      TabIndex        =   2
      Top             =   2430
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Form Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   2445
      TabIndex        =   1
      Top             =   5040
      Width           =   1500
   End
   Begin MSComctlLib.ListView lstGroups 
      Height          =   1590
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   2805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Group Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowed Forms:"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Groups:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   135
      Width           =   555
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   12
      X2              =   408
      Y1              =   316
      Y2              =   316
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   12
      X2              =   408
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Label lblForm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   4170
      Width           =   390
   End
End
Attribute VB_Name = "frmManageGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListGroups()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DBConn.Execute("SELECT GroupID, GroupName FROM tblUserGroup ORDER BY GroupName")
    
    lstGroups.ListItems.Clear
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        Do Until rsData.EOF Or rsData.BOF
            With lstGroups.ListItems.Add(, , rsData("GroupID").value)
                .SubItems(1) = rsData("GroupName").value & ""
            End With
            
            rsData.MoveNext
        Loop
    End If
End Sub

Private Sub ListGroupForms_LST(ByVal GroupID As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DBConn.Execute("SELECT PrivilegeID, FormName FROM tblUserGroupPrivileges AS ugp " & _
        "INNER JOIN tblForms AS f ON ugp.FormID = f.FormID " & _
        "WHERE ugp.GroupID = " & GroupID & " ORDER BY FormName")
    
    lstForms.ListItems.Clear
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        Do Until rsData.EOF Or rsData.BOF
            With lstForms.ListItems.Add(, , rsData("PrivilegeID").value)
                .SubItems(1) = rsData("FormName").value & ""
            End With
            
            rsData.MoveNext
        Loop
    End If
End Sub

Private Sub ListGroupForms_CMB(ByVal GroupID As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DBConn.Execute("SELECT f.FormID, f.FormName FROM tblForms AS f " & _
        "LEFT JOIN tblUserGroupPrivileges AS ugp ON ((ugp.FormID = f.FormID) AND (ugp.GroupID = " & GroupID & "))" & _
        "WHERE ugp.FormID IS NULL ORDER BY FormName")
    
    cmbForms.Clear
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        Do Until rsData.EOF Or rsData.BOF
            cmbForms.AddItem rsData("FormName").value & ""
            cmbForms.ItemData(cmbForms.NewIndex) = rsData("FormID").value
            
            rsData.MoveNext
        Loop
    End If
    
    cmbForms.Enabled = cmbForms.ListCount > 0
    lblForm.Enabled = cmbForms.Enabled
    cmdAddForm.Enabled = cmbForms.Enabled
End Sub

Private Sub cmdAddForm_Click()
    Dim RecordsAffected As Long
    
    If cmbForms.ListIndex = -1 Then
        MsgBox "Select an item first", vbInformation
    Else
        DBConn.Execute "INSERT INTO tblUserGroupPrivileges (GroupID, FormID) " & _
            "VALUES(" & lstGroups.SelectedItem.Text & "," & cmbForms.ItemData(cmbForms.ListIndex) & ")"
        
        ListGroupForms_LST Val(Me.lstGroups.SelectedItem.Text)
        ListGroupForms_CMB Val(Me.lstGroups.SelectedItem.Text)
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewGroup_Click()
    Dim GroupName As String, RecordsAffected As Long
    
    GroupName = InputBox("Type the new group name:", "Enter group name", "")
    
    If Len(Trim(GroupName)) > 0 Then
        DBConn.Execute "INSERT INTO tblUserGroup (GroupName) " & _
            "VALUES('" & Replace(Trim(GroupName), "'", "''") & "')", RecordsAffected
        
        If RecordsAffected = 0 Then
            MsgBox "Error, no records affected.", vbExclamation
        Else
            ListGroups
            lstGroups_ItemClick Me.lstGroups.ListItems(1)
        End If
    End If
End Sub

Private Sub Form_Load()
    ListGroups
End Sub

Private Sub lstForms_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim RecordsAffected As Long
    
    If KeyCode = 46 Then
        If Not (Me.lstForms.SelectedItem Is Nothing) Then
            If MsgBox("Are you sure you want to delete this entry ?", vbYesNo + vbQuestion, "Delete entry ?") = vbYes Then
                DBConn.Execute "DELETE FROM tblUserGroupPrivileges WHERE PrivilegeID = " & Val(Me.lstForms.SelectedItem.Text), RecordsAffected
                
                If RecordsAffected = 0 Then
                    MsgBox "Error, No records affected.", vbExclamation
                Else
                    ListGroupForms_LST Val(Me.lstGroups.SelectedItem.Text)
                    ListGroupForms_CMB Val(Me.lstGroups.SelectedItem.Text)
                End If
            End If
        End If
    End If
End Sub

Private Sub lstGroups_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lstForms.Enabled = True
    
    ListGroupForms_LST Val(Item.Text)
    ListGroupForms_CMB Val(Item.Text)
End Sub
