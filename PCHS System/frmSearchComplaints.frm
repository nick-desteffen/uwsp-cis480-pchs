VERSION 5.00
Begin VB.Form frmSearchComplaints 
   Caption         =   "Search Complaints"
   ClientHeight    =   1560
   ClientLeft      =   6060
   ClientTop       =   2865
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4200
   Begin VB.TextBox txtComplaint 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSearchComplaints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
Dim intComplaint As Integer
Dim intMsgBox As Integer

Dim rstSearch As ADODB.Recordset
Set rstSearch = New ADODB.Recordset

If IsNumeric(txtComplaint.Text) = True And txtComplaint.Text <> "" Then
    intComplaint = txtComplaint.Text
Else
    intMsgBox = MsgBox("Please enter a valid complaint number!", vbOKOnly, "Invalid Number")
    Exit Sub
End If

Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM COMPLAINT WHERE COMPLAINT_NUMBER = " & intComplaint)

If rstSearch.EOF = False Then
    With rstSearch
        rstSearch.MoveFirst
        Do While Not rstSearch.EOF
            If Not IsNull(![COMPLAINT_NUMBER]) Then
                frmShowComplaint.intComplaint = intComplaint
                frmShowComplaint.Show
            End If
            rstSearch.MoveNext
        Loop
    End With
Else
    intMsgBox = MsgBox("No matches found.", vbOKOnly, "None found")
    Unload Me
    frmPCHS_Main.Show
End If

Unload Me
End Sub

