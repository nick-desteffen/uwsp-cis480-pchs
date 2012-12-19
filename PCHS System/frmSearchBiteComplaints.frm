VERSION 5.00
Begin VB.Form frmSearchBiteComplaints 
   Caption         =   "Search Bite Complaints"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBiteComplaint 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSearchBiteComplaints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
Dim intBiteComplaint As Integer
Dim intMsgBox As Integer

Dim rstSearch As ADODB.Recordset
Set rstSearch = New ADODB.Recordset

If IsNumeric(txtBiteComplaint.Text) = True And txtBiteComplaint.Text <> "" Then
    intBiteComplaint = txtBiteComplaint.Text
Else
    intMsgBox = MsgBox("Please enter a valid bite complaint number!", vbOKOnly, "Invalid Number")
    Exit Sub
End If

Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM BITE WHERE BITE_NUMBER = " & intBiteComplaint)

If rstSearch.EOF = False Then
    With rstSearch
        rstSearch.MoveFirst
        Do While Not rstSearch.EOF
            If Not IsNull(![BITE_NUMBER]) Then
                frmShowBiteComplaint.intBiteComplaint = intBiteComplaint
                frmShowBiteComplaint.Show
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

