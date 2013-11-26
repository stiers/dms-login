VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSalesReport 
   BackColor       =   &H00004000&
   Caption         =   "Sales Report — Guill-Bern Corporation"
   ClientHeight    =   8430
   ClientLeft      =   240
   ClientTop       =   2550
   ClientWidth     =   14490
   Icon            =   "frmSalesReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14490
   Begin VB.Frame Frame14 
      BackColor       =   &H00004000&
      Caption         =   "Date"
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
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   3495
      Begin MSComCtl2.DTPicker sls_dteRefDate 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   100139009
         CurrentDate     =   41520
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "Reference Number"
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
      Height          =   855
      Left            =   3720
      TabIndex        =   18
      Top             =   5400
      Width           =   3495
      Begin VB.TextBox sls_txtRefNumber 
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00004000&
      Caption         =   "Company"
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
      Height          =   855
      Left            =   7320
      TabIndex        =   17
      Top             =   5400
      Width           =   3495
      Begin VB.TextBox sls_txtCompany 
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00004000&
      Caption         =   "Contact Person"
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
      Height          =   855
      Left            =   10920
      TabIndex        =   16
      Top             =   5400
      Width           =   3495
      Begin VB.TextBox sls_txtContact 
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00004000&
      Caption         =   "Equipment"
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
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   3495
      Begin VB.TextBox sls_txtEquipment 
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00004000&
      Caption         =   "Brand"
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
      Height          =   855
      Left            =   3720
      TabIndex        =   14
      Top             =   6360
      Width           =   3495
      Begin VB.TextBox sls_txtBrand 
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00004000&
      Caption         =   "Model"
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
      Height          =   855
      Left            =   7320
      TabIndex        =   13
      Top             =   6360
      Width           =   3495
      Begin VB.TextBox sls_txtModel 
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00004000&
      Caption         =   "Price"
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
      Height          =   855
      Left            =   10920
      TabIndex        =   12
      Top             =   6360
      Width           =   3495
      Begin VB.TextBox sls_txtPrice 
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.TextBox sls_txtSearch 
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
      Left            =   11040
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid grdSales 
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GuillBernApp.jcbutton cmdDeleteSalesReport 
      Height          =   555
      Left            =   12480
      TabIndex        =   11
      Top             =   7560
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete"
   End
   Begin GuillBernApp.jcbutton cmdUpdateSalesReport 
      Height          =   555
      Left            =   10680
      TabIndex        =   10
      Top             =   7560
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Update"
   End
   Begin GuillBernApp.jcbutton cmdAddSalesReport 
      Height          =   555
      Left            =   8880
      TabIndex        =   9
      Top             =   7560
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
      Caption         =   "Add"
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
      Left            =   9720
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DisplaySales()
    Set rsData = DBConn.Execute("SELECT * FROM tblsales")
    
    With grdSales
        .Rows = rsData.RecordCount + 1
        .Cols = rsData.Fields.Count
        
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Reference Number"
        .TextMatrix(0, 3) = "Company"
        .TextMatrix(0, 4) = "Contact Person"
        .TextMatrix(0, 5) = "Equipment"
        .TextMatrix(0, 6) = "Brand"
        .TextMatrix(0, 7) = "Model"
        .TextMatrix(0, 8) = "Price"
        
        For RowData = 1 To rsData.RecordCount
            .TextMatrix(RowData, 0) = rsData("SalesID").Value
            
            For ColData = 0 To rsData.Fields.Count - 1
                .TextMatrix(RowData, ColData) = rsData.Fields(ColData)
            Next ColData
            
            rsData.MoveNext
        Next RowData
        
        .ColWidth(0) = 350
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2500
        .ColWidth(5) = 2000
        .ColWidth(6) = 1100
        .ColWidth(7) = 1700
        .ColWidth(8) = 1000
    End With
End Sub

Private Sub cmdAddSalesReport_Click()
    Request = "INSERT INTO tblsales(ReferenceDate, ReferenceNumber, CompanyName, ContactPerson, EqName, EqBrand, EqModel, EqPrice) " & _
        "VALUES ('" & sls_dteRefDate.Value & "', '" & sls_txtRefNumber.Text & "', '" & sls_txtCompany.Text & "', '" & sls_txtContact.Text & "', '" & sls_txtEquipment.Text & "', '" & sls_txtBrand.Text & "', '" & sls_txtModel.Text & "', '" & sls_txtPrice.Text & "')"
    
    Debug.Print Request
    DBConn.Execute Request
    MsgBox "Successfully Added!"
    
    DisplaySales
End Sub

Private Sub cmdDeleteSalesReport_Click()
    Dim SalesID As Integer
    
    SalesID = grdSales.TextMatrix(grdSales.Row, 0)

    Request = "DELETE FROM tblsales WHERE SalesID = '" & SalesID & "'"
    
    Debug.Print Request
    DBConn.Execute Request
    MsgBox "Successfully Deleted!"
    
    DisplaySales
End Sub

Private Sub cmdUpdateSalesReport_Click()
    Dim SalesID As Integer
    
    SalesID = grdSales.TextMatrix(grdSales.Row, 0)

    Request = "UPDATE tblsales SET ReferenceDate = '" & sls_dteRefDate.Value & "', ReferenceNumber = '" & sls_txtRefNumber.Text & "', CompanyName = '" & sls_txtCompany.Text & "', ContactPerson = '" & sls_txtContact.Text & "', EqName = '" & sls_txtEquipment.Text & "', EqBrand = '" & sls_txtBrand.Text & "', EqModel = '" & sls_txtModel.Text & "', EqPrice = '" & sls_txtPrice.Text & "' " & _
        "WHERE SalesID = '" & SalesID & "'"
    
    Debug.Print Request
    DBConn.Execute Request
    MsgBox "Successfully Updated!"
    
    DisplaySales
End Sub

Private Sub Form_Load()
    DisplaySales
    
    sls_dteRefDate.Value = Format(Now, "mm/dd/yyyy")
    
    Me.cmdAddSalesReport.Enabled = True
    Me.cmdUpdateSalesReport.Enabled = False
    Me.cmdDeleteSalesReport.Enabled = False
End Sub

Private Sub grdSales_DblClick()
    Dim SalesID As Integer
    
    Me.cmdAddSalesReport.Enabled = False
    Me.cmdUpdateSalesReport.Enabled = True
    Me.cmdDeleteSalesReport.Enabled = True
    
    With grdSales
        SalesID = .TextMatrix(.Row, 0)
        sls_dteRefDate = .TextMatrix(.Row, 1)
        sls_txtRefNumber = .TextMatrix(.Row, 2)
        sls_txtCompany = .TextMatrix(.Row, 3)
        sls_txtContact = .TextMatrix(.Row, 4)
        sls_txtEquipment = .TextMatrix(.Row, 5)
        sls_txtBrand = .TextMatrix(.Row, 6)
        sls_txtModel = .TextMatrix(.Row, 7)
        sls_txtPrice = .TextMatrix(.Row, 8)
    End With
End Sub

Private Sub sls_txtSearch_Change()
    Dim SearchKey As String
    
    SearchKey = sls_txtSearch.Text
    
    If Len(SearchKey) <= 0 Then
        DisplaySales
        Exit Sub
    End If
        
    
    Set rsData = DBConn.Execute("SELECT * FROM tblsales WHERE MATCH(ReferenceDate, ReferenceNumber, CompanyName, ContactPerson, EqName, EqBrand, EqModel) AGAINST('%" & SearchKey & "%' IN BOOLEAN MODE)")
    
    With grdSales
        .Rows = rsData.RecordCount + 1
        .Cols = rsData.Fields.Count
        
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Reference Number"
        .TextMatrix(0, 3) = "Company"
        .TextMatrix(0, 4) = "Contact Person"
        .TextMatrix(0, 5) = "Equipment"
        .TextMatrix(0, 6) = "Brand"
        .TextMatrix(0, 7) = "Model"
        .TextMatrix(0, 8) = "Price"
        
        For RowData = 1 To rsData.RecordCount
            .TextMatrix(RowData, 0) = rsData("SalesID").Value
            
            For ColData = 0 To rsData.Fields.Count - 1
                .TextMatrix(RowData, ColData) = rsData.Fields(ColData)
            Next ColData
            
            rsData.MoveNext
        Next RowData
        
        .ColWidth(0) = 350
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 2500
        .ColWidth(5) = 2000
        .ColWidth(6) = 1000
        .ColWidth(7) = 2000
        .ColWidth(8) = 900
    End With
End Sub
