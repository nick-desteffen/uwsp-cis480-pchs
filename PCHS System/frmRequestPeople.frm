VERSION 5.00
Begin VB.Form frmPeople 
   Caption         =   "Please select the correct person"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "frmRequestPeople.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFname 
      Enabled         =   0   'False
      Height          =   1425
      Left            =   960
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ListBox lstDOB 
      Enabled         =   0   'False
      Height          =   1425
      Left            =   4680
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox lstLicense 
      Enabled         =   0   'False
      Height          =   1425
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Use None"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdPeople 
      Caption         =   "Use Selected"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox lstPeopleNum 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblFname 
      Caption         =   "First Name"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblMany 
      Caption         =   "The following people were found in the system with the name:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblDOB 
      Caption         =   "Date of Birth"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblLicense 
      Caption         =   "Drivers License"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblNum 
      Caption         =   "Number"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************
'* This form is passed the first and last name of a person to be searched for
'* from another form.  It searches through the database and if a match is found
'* it displays the match with the person's number, driver's license, and DOB.
'* The user then selects which person to use and clicks use selected, or use
'* none.  The form then grabs all the person's information and passes it back
'* to the original form.
'*
'* Written by: Nick DeSteffen
'* Written on: 11-15-2002
'*******************************************************************************

Public bolMatchFound As Boolean            'True = person already in database
Dim bolBad As Boolean                   'Whethere or not the person is bad
Public strFname As String               'Person's first name
Public strLname As String               'Person's last name
Public intType As Integer               'Form number thats calling this form

Private Sub cmdNone_Click()
Select Case intType
        Case 1: frmNewRequest.intPersonNum = 0
                frmNewRequest.bolMatchFound = False
        Case 2: frmNewMissing.intPersonNum = 0
                frmNewMissing.bolMatchFound = False
        Case 3: frmSurrender.intPersonNum = 0
                frmSurrender.bolMatchFound = False
        Case 4: frmNewAdoption.intPersonNum = 0
                frmNewAdoption.bolMatchFound = False
        Case 5: frmNewComplaint.intPersonNum = 0
                frmNewComplaint.bolMatchFound = False
        Case 6: frmNewDonation.intPersonNum = 0
                frmNewDonation.bolMatchFound = False
        Case 7: frmNewBiteComplaint.intPersonNum = 0
                frmNewBiteComplaint.bolMatchFound = False
        Case 8: frmLicense.intPersonNum = 0
                frmLicense.bolMatchFound = False
        Case 9: frmReclaim.intPersonNum = 0
                frmReclaim.bolMatchFound = False
        Case 10: frmMiscReceipt.intPersonNum = 0
                 frmMiscReceipt.bolMatchFound = False
        Case 11: frmBadPerson.intPersonNum = 0
                 frmBadPerson.bolMatchFound = False
        Case 12: frmCompleteSearch.intPersonNum = 0
    End Select
    Unload Me
End Sub
Private Sub cmdPeople_Click()

Dim intMsgBox As Integer        'Used for messageboxes
Dim rstBad As ADODB.Recordset   'Recordset used for interfacing with the database

On Error GoTo ErrorHandler

Set rstBad = New ADODB.Recordset

If lstPeopleNum.Text <> "" Then
    Set rstBad = Open_Recordsets.objConnection.Execute("SELECT PERSON_BAD FROM PERSON WHERE PERSON_NUMBER = " & lstPeopleNum.Text)
    With rstBad
        If ![PERSON_BAD] = True Then
            intMsgBox = MsgBox("This person is marked as a bad person!" & Chr(13) & "                 Continue process?", vbYesNo, "WARNING")
            If intMsgBox = 7 Then
                Exit Sub
            End If
        End If
    End With
    Select Case intType
        Case 1: frmNewRequest.intPersonNum = lstPeopleNum.Text
        Case 2: frmNewMissing.intPersonNum = lstPeopleNum.Text
        Case 3: frmSurrender.intPersonNum = lstPeopleNum.Text
        Case 4: frmNewAdoption.intPersonNum = lstPeopleNum.Text
        Case 5: frmNewComplaint.intPersonNum = lstPeopleNum.Text
        Case 6: frmNewDonation.intPersonNum = lstPeopleNum.Text
        Case 7: frmNewBiteComplaint.intPersonNum = lstPeopleNum.Text
        Case 8: frmLicense.intPersonNum = lstPeopleNum.Text
        Case 9: frmReclaim.intPersonNum = lstPeopleNum.Text
        Case 10: frmMiscReceipt.intPersonNum = lstPeopleNum.Text
        Case 11: frmBadPerson.intPersonNum = lstPeopleNum.Text
        Case 12: frmCompleteSearch.intPersonNum = lstPeopleNum.Text
    End Select
    Unload Me
Else
    intMsgBox = MsgBox("Please select a person!", vbOKOnly, "Select please")
    Exit Sub
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub
Private Sub Form_Load()

Dim rstSearch As New ADODB.Recordset        'Recordset used for interfacing with the database
Dim intMsgBox As Integer                    'Used for messageboxes

On Error GoTo ErrorHandler

lblName.Caption = strFname & " " & strLname
Set rstSearch = New ADODB.Recordset

Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM PERSON WHERE PERSON_LNAME = '" & strLname & "'")
If rstSearch.EOF = False Then
    With rstSearch
        rstSearch.MoveFirst
        Do While Not rstSearch.EOF
            lstPeopleNum.AddItem ![PERSON_NUMBER]
            If Not IsNull(![person_license]) Then
                lstLicense.AddItem ![person_license]
            Else
                lstLicense.AddItem "No License"
            End If
            If Not IsNull(![person_dob]) Then
                lstDOB.AddItem Format(![person_dob], "MM/DD/YYYY")
            Else
                lstDOB.AddItem "No Birth Date"
            End If
            If Not IsNull(![PERSON_FNAME]) Then
                lstFname.AddItem ![PERSON_FNAME]
            Else
                lstFname.AddItem "None"
            End If
            rstSearch.MoveNext
        Loop
    End With
End If

Set rstSearch = Nothing

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Sub

Private Sub lstPeopleNum_Click()
    lstLicense.ListIndex = lstPeopleNum.ListIndex
    lstDOB.ListIndex = lstPeopleNum.ListIndex
    lstFname.ListIndex = lstPeopleNum.ListIndex
End Sub
