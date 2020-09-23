VERSION 5.00
Begin VB.Form cfgfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Setting Up System"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close application"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cfg_form 
      Caption         =   "Save changes"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox n4 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox n3 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox n2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Default system name"
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox n1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "The Admin"
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "System Root:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Working dir:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "System Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Admin Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "cfgfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cfg_form_Click()
Dim fs As Integer
fs = FreeFile
Open "system.cfg" For Output As fs
Print #fs, n1.Text + ";" + n2.Text + ";" + n3.Text + ";" + n4.Text + ";"
Close fs
cfg.adminname = n1.Text
cfg.sysname = n2.Text
cfg.workdir = n3.Text
cfg.sysroot = n4.Text
Unload Me
End Sub

Private Sub Command2_Click()
FreeConsole
End
End Sub

Private Sub Form_Load()
n3.Text = App.path
n4.Text = "c:\"

End Sub
