VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form actsfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network Activities"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "actsfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label Label5 
      Caption         =   "Bytes recv/sec:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Bytes sent/sec:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Number of users connected:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Bandwitch per user:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Total Bandwitch:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "actsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lastselected = 0
End Sub

