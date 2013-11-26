VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmManageDAR 
   BackColor       =   &H00004000&
   Caption         =   "Daily Activity Report — Guill-Bern Corporation"
   ClientHeight    =   8445
   ClientLeft      =   2745
   ClientTop       =   2340
   ClientWidth     =   15645
   Icon            =   "frmManageDAR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   15645
   Begin VB.TextBox txtTSFR 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   50
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00004000&
      Caption         =   "Instrument or Equipment"
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
      Height          =   4095
      Left            =   4560
      TabIndex        =   36
      Top             =   120
      Width           =   6855
      Begin VB.Frame Frame8 
         BackColor       =   &H00004000&
         Caption         =   "Product"
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
         Left            =   240
         TabIndex        =   45
         Top             =   1320
         Width           =   3135
         Begin VB.ComboBox dar_cboProduct 
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
            Height          =   360
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame9 
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
         Left            =   240
         TabIndex        =   44
         Top             =   2160
         Width           =   3135
         Begin VB.ComboBox dar_cboModel 
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
            Height          =   360
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00004000&
         Caption         =   "Serial Number"
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
         Left            =   240
         TabIndex        =   43
         Top             =   3000
         Width           =   3135
         Begin VB.TextBox dar_txtSerial 
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
            TabIndex        =   9
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00004000&
         Caption         =   "Details"
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
         Height          =   2535
         Left            =   3480
         TabIndex        =   42
         Top             =   1320
         Width           =   3135
         Begin VB.TextBox dar_txtJobDetails 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame7 
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
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   3135
         Begin VB.ComboBox dar_cboBrand 
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
            Height          =   360
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00004000&
         Caption         =   "Status"
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
         Left            =   3480
         TabIndex        =   37
         Top             =   480
         Width           =   3135
         Begin VB.ComboBox dar_cboStatus 
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
            Height          =   360
            ItemData        =   "frmManageDAR.frx":0442
            Left            =   120
            List            =   "frmManageDAR.frx":045B
            TabIndex        =   10
            Top             =   360
            Width           =   2895
         End
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00004000&
      Caption         =   "Contact Person who signed the TSFR"
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
      Height          =   4095
      Left            =   11520
      TabIndex        =   34
      Top             =   120
      Width           =   3975
      Begin VB.Frame Frame15 
         BackColor       =   &H00004000&
         Caption         =   "Position"
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
         Left            =   240
         TabIndex        =   47
         Top             =   1560
         Width           =   3495
         Begin VB.TextBox dar_txtContactPosition 
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
            TabIndex        =   13
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00004000&
         Caption         =   "Contact Number"
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
         Left            =   240
         TabIndex        =   46
         Top             =   2640
         Width           =   3495
         Begin VB.TextBox dar_txtContactNumber 
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
            TabIndex        =   14
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00004000&
         Caption         =   "Name"
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
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   3495
         Begin VB.TextBox dar_txtContactPerson 
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
            TabIndex        =   12
            Top             =   360
            Width           =   3255
         End
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00004000&
      Caption         =   "Next Action Plan"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   31
      Top             =   4320
      Width           =   8775
      Begin VB.Frame Frame18 
         BackColor       =   &H00004000&
         Caption         =   "Details"
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
         Height          =   2175
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   4815
         Begin VB.TextBox dar_txtPlanDetails 
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
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame Frame19 
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
         Height          =   735
         Left            =   5160
         TabIndex        =   32
         Top             =   360
         Width           =   3375
         Begin MSComCtl2.DTPicker dar_dtePlanDate 
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
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
            Format          =   100270081
            CurrentDate     =   41519
         End
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H00004000&
      Caption         =   "Expenses Per Amount"
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
      Height          =   2775
      Left            =   9000
      TabIndex        =   29
      Top             =   4320
      Width           =   6495
      Begin VB.Frame Frame22 
         BackColor       =   &H00004000&
         Caption         =   "Materials"
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
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   2895
         Begin VB.TextBox dar_txtMaterials 
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
            TabIndex        =   18
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame23 
         BackColor       =   &H00004000&
         Caption         =   "Meals"
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
         Left            =   3360
         TabIndex        =   40
         Top             =   360
         Width           =   2895
         Begin VB.TextBox dar_txtMeals 
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
            TabIndex        =   19
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame24 
         BackColor       =   &H00004000&
         Caption         =   "Accommodation"
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
         Left            =   3360
         TabIndex        =   39
         Top             =   1440
         Width           =   2895
         Begin VB.TextBox dar_txtAccommodation 
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
            TabIndex        =   20
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame21 
         BackColor       =   &H00004000&
         Caption         =   "Transportation"
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
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   2895
         Begin VB.TextBox dar_txtTransportation 
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
            TabIndex        =   17
            Top             =   360
            Width           =   2655
         End
      End
   End
   Begin VB.Frame Frame25 
      BackColor       =   &H00004000&
      Caption         =   "Service Profile"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame2 
         BackColor       =   &H00004000&
         Caption         =   "Time In"
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
         Height          =   735
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1815
         Begin MSComCtl2.DTPicker dar_dteTimeIn 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   100270082
            CurrentDate     =   41519
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00004000&
         Caption         =   "Time Out"
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
         Height          =   735
         Left            =   2160
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
         Begin MSComCtl2.DTPicker dar_dteTimeOut 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   100270082
            CurrentDate     =   41519
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00004000&
         Caption         =   "Type of Job Order"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   3735
         Begin VB.ComboBox dar_cboJobType 
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
            Height          =   360
            ItemData        =   "frmManageDAR.frx":050B
            Left            =   120
            List            =   "frmManageDAR.frx":051B
            TabIndex        =   4
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00004000&
         Caption         =   "Account"
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
         Left            =   240
         TabIndex        =   1
         Top             =   3000
         Width           =   3735
         Begin VB.ComboBox dar_cboJobAccount 
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
            Height          =   360
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Label dar_lblReportID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "lblReportID"
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
         Height          =   240
         Left            =   3120
         TabIndex        =   48
         Top             =   360
         Width           =   945
      End
      Begin VB.Label dar_lblUserName 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "lblUserName"
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
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label dar_lblServiceDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "lblServiceDate"
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
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
   End
   Begin GuillBernApp.jcbutton cmdDeleteDAR 
      Height          =   555
      Left            =   13560
      TabIndex        =   23
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
      Caption         =   "Delete"
   End
   Begin GuillBernApp.jcbutton cmdUpdateDAR 
      Height          =   555
      Left            =   11760
      TabIndex        =   22
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
      Caption         =   "Update"
   End
   Begin GuillBernApp.jcbutton cmdAddDAR 
      Height          =   555
      Left            =   9960
      TabIndex        =   21
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TSFR #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   49
      Top             =   7320
      Width           =   660
   End
End
Attribute VB_Name = "frmManageDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportID As Integer

Private Sub cmdAddDAR_Click()
    Request = "INSERT INTO tbldar(UserID, ServiceDate, TimeIn, TimeOut, JobOrder, JobAccount, Brand, Product, Model, SerialNumber, JobStatus, JobDetails, ContactPerson, ContactPosition, ContactNumber, PlanDetails, PlanDate, TransportExpense, MealExpense, MaterialExpense, AccommodationExpense ) " & _
        "VALUES ('" & LogInUserID & "', '" & dar_lblServiceDate.Caption & "', '" & dar_dteTimeIn.Value & "', '" & dar_dteTimeOut.Value & "', '" & dar_cboJobType.Text & "', '" & dar_cboJobAccount.Text & "', '" & dar_cboBrand.Text & "', '" & dar_cboProduct.Text & "', '" & dar_cboModel.Text & "', '" & dar_txtSerial.Text & "', '" & dar_cboStatus.Text & "', '" & dar_txtJobDetails.Text & "', '" & dar_txtContactPerson.Text & "', '" & dar_txtContactPosition.Text & "', '" & dar_txtContactNumber.Text & "', '" & dar_txtPlanDetails.Text & "', '" & dar_dtePlanDate.Value & "', '" & dar_txtTransportation.Text & "', '" & dar_txtMeals.Text & "', '" & dar_txtMaterials.Text & "', '" & dar_txtAccommodation.Text & "')"
    
    Debug.Print Request
    DBConn.Execute Request
    MsgBox "Successfully Added!"
    Unload Me
End Sub

Private Sub cmdDeleteDAR_Click()
    ReportID = frmDAR.grdActivity.TextMatrix(frmDAR.grdActivity.Row, 0)
    
    Request = "DELETE FROM tbldar WHERE ReportID = '" & ReportID & "'"
    
    Debug.Print Request
    DBConn.Execute Request
    MsgBox "Successfully Deleted!"
    Unload Me
End Sub

Private Sub cmdUpdateDAR_Click()
    ReportID = frmDAR.grdActivity.TextMatrix(frmDAR.grdActivity.Row, 0)

    Request = "UPDATE tbldar SET TimeIn = '" & dar_dteTimeIn.Value & "', TimeOut = '" & dar_dteTimeOut.Value & "', JobOrder = '" & dar_cboJobType.Text & "', JobAccount = '" & dar_cboJobAccount.Text & "', Brand = '" & dar_cboBrand.Text & "', Product = '" & dar_cboProduct.Text & "', Model = '" & dar_cboModel.Text & "', SerialNumber = '" & dar_txtSerial.Text & "', JobStatus = '" & dar_cboStatus.Text & "', JobDetails = '" & dar_txtJobDetails.Text & "', ContactPerson = '" & dar_txtContactPerson.Text & "', ContactPosition = '" & dar_txtContactPosition.Text & "', ContactNumber = '" & dar_txtContactNumber.Text & "', PlanDetails = '" & dar_txtPlanDetails.Text & "', PlanDate = '" & dar_dtePlanDate.Value & "', TransportExpense = '" & dar_txtTransportation.Text & "', MealExpense = '" & dar_txtMeals.Text & "', MaterialExpense = '" & dar_txtMaterials.Text & "', AccommodationExpense = '" & dar_txtAccommodation.Text & "' " & _
        "WHERE ReportID = '" & ReportID & "'"
        
    Debug.Print Request
    DBConn.Execute Request
    MsgBox "Successfully Updated!"
    Unload Me
End Sub

Private Sub dar_cboBrand_GotFocus()
    Set rsData = DBConn.Execute("SELECT DISTINCT Brand FROM tbldar")
    
    While Not rsData.EOF
        dar_cboBrand.AddItem rsData!Brand
        rsData.MoveNext
    Wend
End Sub

Private Sub dar_cboJobAccount_GotFocus()
    Set rsData = DBConn.Execute("SELECT DISTINCT JobAccount FROM tbldar")
    
    While Not rsData.EOF
        dar_cboJobAccount.AddItem rsData!JobAccount
        rsData.MoveNext
    Wend
End Sub

Private Sub dar_cboJobType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub dar_cboModel_GotFocus()
    Set rsData = DBConn.Execute("SELECT DISTINCT Model FROM tbldar")
    
    While Not rsData.EOF
        dar_cboModel.AddItem rsData!Model
        rsData.MoveNext
    Wend
End Sub

Private Sub dar_cboProduct_GotFocus()
    Set rsData = DBConn.Execute("SELECT DISTINCT Product FROM tbldar")
    
    While Not rsData.EOF
        dar_cboProduct.AddItem rsData!Product
        rsData.MoveNext
    Wend
End Sub

Private Sub dar_cboStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    If Not frmDAR.grdActivity.TextMatrix(frmDAR.grdActivity.Row, 0) = "" Then
        ReportID = frmDAR.grdActivity.TextMatrix(frmDAR.grdActivity.Row, 0)
    End If

    dar_lblReportID.Visible = False
    dar_lblReportID.Caption = "DAR: " & ReportID
    dar_lblUserName.Caption = LogInUserName
    dar_lblServiceDate.Caption = Format(Now, "mm/dd/yyyy")
    dar_dteTimeIn.Value = Format(Now, "hh:mm AMPM")
    dar_dteTimeOut.Value = Format(Now, "hh:mm AMPM")
    dar_dtePlanDate.Value = Format(Now, "mm/dd/yyyy")
End Sub
