VERSION 5.00
Begin VB.Form frmMissing 
   Caption         =   "Missing Animals"
   ClientHeight    =   3675
   ClientLeft      =   3780
   ClientTop       =   3690
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4680
   Begin VB.PictureBox picMissing 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   2280
      Picture         =   "frmMissing.frx":0000
      ScaleHeight     =   2415
      ScaleWidth      =   2175
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdListMissing 
      Caption         =   "List Missing Animals"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewMissing 
      Caption         =   "New Missing Animal"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmMissing.Hide
frmPCHS_Main.Show
End Sub

Private Sub cmdListMissing_Click()
frmListMissing.Show
frmMissing.Hide
End Sub

Private Sub cmdNewMissing_Click()
frmNewMissing.Show
frmMissing.Hide
End Sub

