VERSION 5.00
Begin VB.Form frmCompleteSearch 
   Caption         =   "Complete Search"
   ClientHeight    =   2790
   ClientLeft      =   3000
   ClientTop       =   2385
   ClientWidth     =   7395
   Icon            =   "frmCompleteSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   7395
   Begin VB.TextBox txtLname 
      Height          =   285
      Left            =   1200
      TabIndex        =   18
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtFname 
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewSurrender 
      Caption         =   "View Surrender"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewReceipt 
      Caption         =   "View Receipt"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox cboSurrenderNum 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox cboReceiptNum 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdViewBiteComplaint 
      Caption         =   "View Bite Report"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewComplaint 
      Caption         =   "View Complaint"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewAdoption 
      Caption         =   "View Adoption"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cboBiteComplaintNum 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cboComplaintNum 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cboAdoptionNum 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdSearchName 
      Caption         =   "Search"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblLname 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblFname 
      Caption         =   "First Name"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblSurrenderNum 
      Caption         =   "Surrendered Animals:"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblReceiptNum 
      Caption         =   "Receipts:"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblBiteComplaintNum 
      Caption         =   "Bite Reports:"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblComplaintNum 
      Caption         =   "Complaints:"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Adoptions:"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuBack 
         Caption         =   "Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmCompleteSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************
'* This form takes a user inputted name and searches for all the adoptions, surrenders,
'* reciepts, bite reports, and animal complaints this person has.  The user selects the
'* report number from the appropriate type and clicks the view button to view it.
'*
'**************************************************************************************

Public intPersonNum As Integer             'Number of the person if found

Public Sub cmdSearchName_Click()
'********************************************************************************************
'* Runs after the search button is pressed.  It searches through the adoption table, surrender
'* table, reciept table, bite table, and complaint table and returns all the reports that
'* pertain to the searched person.
'*
'* Written by: Nick DeSteffen
'* Written on: 12-08-2002
'********************************************************************************************
Dim rstSearch As New ADODB.Recordset        'Used for interfacing with the database
Dim intMsgBox As Integer                    'Used for messageboxes
Dim strSQL As String                        'SQL Statement

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset

strSQL = "SELECT * FROM PERSON WHERE PERSON_LNAME = '" & txtLname.Text & "'"
Set rstSearch = Open_Recordsets.objConnection.Execute(strSQL)

If rstSearch.EOF <> True Then
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 12
    
    frmPeople.Show (vbModal)

    If intPersonNum <> 0 Then
        
        'Finds all the adoptions
        Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT ADOPTION_NUMBER FROM ADOPTION WHERE ADOPTION_ADOPTORNUM = " & intPersonNum & " ORDER BY ADOPTION_NUMBER")
        If rstSearch.EOF <> True Then
            With rstSearch
                rstSearch.MoveFirst
                Do While Not rstSearch.EOF
                    cboAdoptionNum.AddItem (![ADOPTION_NUMBER])
                    rstSearch.MoveNext
                Loop
            End With
        End If
        Set rstSearch = Nothing
        
        'Finds all the complaints
        Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT COMPLAINT_NUMBER FROM COMPLAINT WHERE COMPLAINT_PERSON = " & intPersonNum & " ORDER BY COMPLAINT_NUMBER")
        If rstSearch.EOF <> True Then
            With rstSearch
                rstSearch.MoveFirst
                Do While Not rstSearch.EOF
                    cboComplaintNum.AddItem (![COMPLAINT_NUMBER])
                    rstSearch.MoveNext
                Loop
            End With
        End If
        Set rstSearch = Nothing
        
        'Finds all the bite reports
        Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT BITE_NUMBER FROM BITE WHERE BITE_PERSON = " & intPersonNum & " ORDER BY BITE_NUMBER")
        If rstSearch.EOF <> True Then
            With rstSearch
                rstSearch.MoveFirst
                Do While Not rstSearch.EOF
                    cboBiteComplaintNum.AddItem (![BITE_NUMBER])
                    rstSearch.MoveNext
                Loop
            End With
        End If
        Set rstSearch = Nothing
        
        'Finds all the receipts
        Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT RECEIPT_NUMBER FROM RECEIPT WHERE RECEIPT_PERSON = " & intPersonNum & " ORDER BY RECEIPT_NUMBER")
        If rstSearch.EOF <> True Then
            With rstSearch
                rstSearch.MoveFirst
                Do While Not rstSearch.EOF
                    cboReceiptNum.AddItem (![RECEIPT_NUMBER])
                    rstSearch.MoveNext
                Loop
            End With
        End If
        Set rstSearch = Nothing
        
        'Finds all the animals surrendered by this person
        Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT SURRENDER_ANIMAL_NUMBER FROM SURRENDER WHERE SURRENDER_OWNER = " & intPersonNum & " ORDER BY SURRENDER_ANIMAL_NUMBER")
        If rstSearch.EOF <> True Then
            With rstSearch
                rstSearch.MoveFirst
                Do While Not rstSearch.EOF
                    cboSurrenderNum.AddItem (![SURRENDER_ANIMAL_NUMBER])
                    rstSearch.MoveNext
                Loop
            End With
        End If
        Set rstSearch = Nothing
            
    End If
Else
    intMsgBox = MsgBox("No matches found.", vbOKOnly, "Not Found")
End If
Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub cmdViewAdoption_Click()
If cboAdoptionNum.Text <> "" Then
    frmShowContract.intAdoptionNum = cboAdoptionNum.Text
    frmShowExamine.intAdoptionNum = cboAdoptionNum.Text
    frmVerifyAdoption.intAdoptionNum = cboAdoptionNum.Text
    frmShowNeuter.intAdoptionNum = cboAdoptionNum.Text
    frmShowContract.Show
    frmShowExamine.Show
    frmVerifyAdoption.Show
    frmShowNeuter.Show
End If

End Sub

Private Sub cmdViewBiteComplaint_Click()
If cboBiteComplaintNum.Text <> "" Then
    frmShowBite.intBiteNum = cboBiteComplaintNum.Text
    frmShowBite.Show
End If
End Sub

Private Sub cmdViewComplaint_Click()
If cboComplaintNum.Text <> "" Then
    frmShowComplaint.intComplaintNum = cboComplaintNum.Text
    frmShowComplaint.Show
End If
End Sub

Private Sub cmdViewReceipt_Click()
If cboReceiptNum.Text <> "" Then
    frmShowReciept.intReceiptNum = cboReceiptNum.Text
    frmShowReciept.Show
End If
End Sub

Private Sub cmdViewSurrender_Click()
If cboSurrenderNum.Text <> "" Then
    frmShowSurrender.intAnimalNum = cboSurrenderNum.Text
    frmShowSurrender.Show
End If
End Sub

Private Sub mnuAbout_Click()
Call About.About
End Sub

Private Sub mnuBack_Click()
Unload Me
frmPCHS_Main.Show
End Sub

Private Sub mnuExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmPCHS_Main.Show
End Sub

