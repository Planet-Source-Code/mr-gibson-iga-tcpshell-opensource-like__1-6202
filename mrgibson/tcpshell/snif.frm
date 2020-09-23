VERSION 5.00
Begin VB.Form snif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sniffing node 0"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "snif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = "Sniffing node" + Str(sniffnode)

End Sub

Private Sub Form_Unload(Cancel As Integer)
sniffnode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 10 Then
sock_out sniffnode, Chr(KeyAscii)
Text1.SelStart = Len(Text1.Text)
End If
KeyAscii = 0
End Sub
