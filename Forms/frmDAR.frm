VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDAR 
   BackColor       =   &H00004000&
   Caption         =   "Daily Activity Report — Guill-Bern Corporation"
   ClientHeight    =   8670
   ClientLeft      =   915
   ClientTop       =   1905
   ClientWidth     =   15420
   Icon            =   "frmDAR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   15420
   Begin VB.TextBox dar_txtSearch 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid grdActivity 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12515
      _Version        =   393216
      Cols            =   22
   End
   Begin GuillBernApp.jcbutton cmdManageDAR 
      Height          =   555
      Left            =   13440
      TabIndex        =   1
      Top             =   8040
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   979
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Manage"
   End
   Begin GuillBernApp.jcbutton cmdHistory 
      Height          =   555
      Left            =   11640
      TabIndex        =   2
      Top             =   8040
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   979
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&History"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   10680
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DisplayActivity(UserID As Long, UserType As Integer)

'    If UserType = 1 Then
'        Request = "SELECT * FROM tbldar WHERE UserID = '" & UserID & "'"
'    Else
'        Request = "SELECT * FROM tbldar"
'    End If
    
    Request = "SELECT * FROM tbldar"
    
    Set rsData = DBConn.Execute(Request)
    
    With grdActivity
        .Rows = rsData.RecordCount + 1
        .Cols = rsData.Fields.Count
        
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Service Date"
        .TextMatrix(0, 3) = "Time In"
        .TextMatrix(0, 4) = "Time Out"
        .TextMatrix(0, 5) = "Job Order"
        .TextMatrix(0, 6) = "Account"
        .TextMatrix(0, 7) = "Brand"
        .TextMatrix(0, 8) = "Product"
        .TextMatrix(0, 9) = "Model"
        .TextMatrix(0, 10) = "Serial Number"
        .TextMatrix(0, 11) = "Status"
        .TextMatrix(0, 12) = "Details"
        .TextMatrix(0, 13) = "Contact Person"
        .TextMatrix(0, 14) = "Position"
        .TextMatrix(0, 15) = "Contact Number"
        .TextMatrix(0, 16) = "Plan Details"
        .TextMatrix(0, 17) = "Plan Date"
        .TextMatrix(0, 18) = "Transport"
        .TextMatrix(0, 19) = "Meals"
        .TextMatrix(0, 20) = "Materials"
        .TextMatrix(0, 21) = "Accommodation"
        
        For RowData = 1 To rsData.RecordCount
            .TextMatrix(RowData, 0) = rsData("ReportID").value
            
            For ColData = 1 To rsData.Fields.Count - 1
                .TextMatrix(RowData, ColData) = rsData.Fields(ColData)
            Next ColData
            
            rsData.MoveNext
        Next RowData
        
        .ColWidth(0) = 350
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 3000
        .ColWidth(4) = 3000
        .ColWidth(5) = 3000
        .ColWidth(6) = 3000
        .ColWidth(7) = 3000
        .ColWidth(8) = 3000
        .ColWidth(9) = 3000
        .ColWidth(10) = 3000
        .ColWidth(11) = 3000
        .ColWidth(12) = 3000
        .ColWidth(13) = 3000
        .ColWidth(14) = 3000
        .ColWidth(15) = 3000
        .ColWidth(16) = 3000
        .ColWidth(17) = 2000
        .ColWidth(18) = 1500
        .ColWidth(19) = 1500
        .ColWidth(20) = 1500
        .ColWidth(21) = 1500
    End With
End Sub

Private Sub cmdHistory_Click()
    frmHistory.Show vbModal, Me
End Sub

Private Sub cmdManageDAR_Click()
    frmManageDAR.Show vbModal, Me
End Sub

Private Sub dar_txtSearch_Change()
    Dim SearchKey As String
    
    SearchKey = dar_txtSearch.Text
    
    If Len(SearchKey) <= 0 Then
        DisplayActivity LogInUserID, LogInUserType
        Exit Sub
    End If
    
    If UserType = 1 Then
        Request = "SELECT * FROM tbldar WHERE UserID = '" & UserID & "' AND MATCH(JobOrder, JobAccount, Brand, Product, Model, SerialNumber, JobStatus, JobDetails) AGAINST('%" & SearchKey & "%' IN BOOLEAN MODE)"
    Else
        Request = "SELECT * FROM tbldar WHERE MATCH(JobOrder, JobAccount, Brand, Product, Model, SerialNumber, JobStatus, JobDetails) AGAINST('%" & SearchKey & "%' IN BOOLEAN MODE)"
    End If
    
    Set rsData = DBConn.Execute(Request)
    
    With grdActivity
        .Rows = rsData.RecordCount + 1
        .Cols = rsData.Fields.Count
        
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Service Date"
        .TextMatrix(0, 3) = "Time In"
        .TextMatrix(0, 4) = "Time Out"
        .TextMatrix(0, 5) = "Job Order"
        .TextMatrix(0, 6) = "Account"
        .TextMatrix(0, 7) = "Brand"
        .TextMatrix(0, 8) = "Product"
        .TextMatrix(0, 9) = "Model"
        .TextMatrix(0, 10) = "Serial Number"
        .TextMatrix(0, 11) = "Status"
        .TextMatrix(0, 12) = "Details"
        .TextMatrix(0, 13) = "Contact Person"
        .TextMatrix(0, 14) = "Position"
        .TextMatrix(0, 15) = "Contact Number"
        .TextMatrix(0, 16) = "Plan Details"
        .TextMatrix(0, 17) = "Plan Date"
        .TextMatrix(0, 18) = "Transport"
        .TextMatrix(0, 19) = "Meals"
        .TextMatrix(0, 20) = "Materials"
        .TextMatrix(0, 21) = "Accommodation"
        
        For RowData = 1 To rsData.RecordCount
            .TextMatrix(RowData, 0) = rsData("ReportID").value
            
            For ColData = 1 To rsData.Fields.Count - 1
                .TextMatrix(RowData, ColData) = rsData.Fields(ColData)
            Next ColData
            
            rsData.MoveNext
        Next RowData
        
        .ColWidth(0) = 350
    End With
    
End Sub

Private Sub Form_Activate()
    DisplayActivity LogInUserID, LogInUserType
End Sub

Private Sub Form_Load()
    DisplayActivity LogInUserID, LogInUserType
    
    Me.cmdManageDAR.Enabled = FormAllowed("frmManageDAR")
    Me.cmdHistory.Visible = FormAllowed("frmHistory")
End Sub

Private Sub grdActivity_DblClick()
    With frmManageDAR
        .dar_dteTimeIn.value = grdActivity.TextMatrix(grdActivity.Row, 3)
        .dar_dteTimeOut.value = grdActivity.TextMatrix(grdActivity.Row, 4)
        .dar_cboJobType.Text = grdActivity.TextMatrix(grdActivity.Row, 5)
        .dar_cboJobAccount.Text = grdActivity.TextMatrix(grdActivity.Row, 6)
        .dar_cboBrand.Text = grdActivity.TextMatrix(grdActivity.Row, 7)
        .dar_cboProduct.Text = grdActivity.TextMatrix(grdActivity.Row, 8)
        .dar_cboModel.Text = grdActivity.TextMatrix(grdActivity.Row, 9)
        .dar_txtSerial.Text = grdActivity.TextMatrix(grdActivity.Row, 10)
        .dar_cboStatus.Text = grdActivity.TextMatrix(grdActivity.Row, 11)
        .dar_txtJobDetails.Text = grdActivity.TextMatrix(grdActivity.Row, 12)
        .dar_txtContactPerson.Text = grdActivity.TextMatrix(grdActivity.Row, 13)
        .dar_txtContactPosition.Text = grdActivity.TextMatrix(grdActivity.Row, 14)
        .dar_txtContactNumber.Text = grdActivity.TextMatrix(grdActivity.Row, 15)
        .dar_txtPlanDetails.Text = grdActivity.TextMatrix(grdActivity.Row, 16)
        .dar_dtePlanDate.value = grdActivity.TextMatrix(grdActivity.Row, 17)
        .dar_txtTransportation.Text = grdActivity.TextMatrix(grdActivity.Row, 18)
        .dar_txtMeals.Text = grdActivity.TextMatrix(grdActivity.Row, 19)
        .dar_txtMaterials.Text = grdActivity.TextMatrix(grdActivity.Row, 20)
        .dar_txtAccommodation.Text = grdActivity.TextMatrix(grdActivity.Row, 21)
    End With
    
    frmManageDAR.dar_lblReportID.Visible = True
    frmManageDAR.Show vbModal, Me
End Sub
