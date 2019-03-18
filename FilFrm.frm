VERSION 5.00
Begin VB.Form FilFrm 
   Caption         =   "Save to file"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   675
      Left            =   225
      TabIndex        =   3
      Top             =   990
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   675
      Left            =   3405
      TabIndex        =   2
      Top             =   1005
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   285
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   270
      Left            =   30
      TabIndex        =   1
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "FilFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    G_Filename = Trim(Text1.Text)
    Unload Me
End Sub
