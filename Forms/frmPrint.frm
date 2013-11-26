VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Print"
   ClientHeight    =   5355
   ClientLeft      =   4965
   ClientTop       =   3930
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9825
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "&Print Preview"
      Height          =   555
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Print Form"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   562
      TabIndex        =   0
      Top             =   2160
      Width           =   8700
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.cmdPrintPreview.Enabled = FormAllowed("frmPrintPreview")
End Sub
