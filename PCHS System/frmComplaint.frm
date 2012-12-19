VERSION 5.00
Begin VB.Form frmComplaint 
   Caption         =   "Animal Complaint"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBiteComplaint 
      Caption         =   "Bite Complaint"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearchComplaints 
      Caption         =   "Search Complaints"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdListComplaints 
      Caption         =   "List Complaints"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewComplaint 
      Caption         =   "New Complaint"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmComplaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
frmComplaint.Hide
frmPCHS_Main.Show
End Sub

Private Sub cmdBiteComplaint_Click()
frmComplaint.Hide
frmNewBiteComplaint.Show
End Sub

Private Sub cmdListComplaints_Click()
frmComplaint.Hide
frmListComplaints.Show
End Sub

Private Sub cmdNewComplaint_Click()
frmComplaint.Hide
frmNewComplaint.Show
End Sub

