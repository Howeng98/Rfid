VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "please input VendorID and  ProductID"
   ClientHeight    =   2280
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   3840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   420
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Text            =   "&HC02E"
      Top             =   1080
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2280
      TabIndex        =   2
      Text            =   "&H1325"
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "ProductID"
      Height          =   372
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "VendorID"
      Height          =   372
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   852
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyVendorIDm As String
Public MyProductID As String

Private Sub Command1_Click()
    MyVendorIDm = Me.Text1.Text
    MyProductID = Me.Text2.Text
'    frmmain.MyVendorID = Me.Text1.Text
'    frmmain.MyProductID = Me.Text2.Text
Me.Hide

End Sub
