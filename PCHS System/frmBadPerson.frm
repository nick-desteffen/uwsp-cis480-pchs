VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBadPerson 
   Caption         =   "New Bad Person"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6135
   Icon            =   "frmBadPerson.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearchPeople 
      Caption         =   "Search People"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame frmPersonal 
      Caption         =   "Person Information"
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtLname 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtFname 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   2760
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37579
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblCIty 
         Caption         =   "City"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
      End
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
Attribute VB_Name = "frmBadPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************************
'* This form takes down the information of a person and saves it to the database.
'* The person is flagged as "Bad" when it is saved.
'*
'***********************************************************************************

Public intPersonNum As Integer          'Person's number
Public bolMatchFound As Boolean            'True = person already in database

'********************************************************************************************
'* Runs after the user clicks save.  It either adds a new person to the database and marks
'* them as "bad" or updates somobody already in the database and marks them as "bad".
'*
'* Written by: Nick DeSteffen
'* Written on: 12-07-2002
'********************************************************************************************
Private Sub cmdSave_Click()
Dim strLname As String              'First name of the person
Dim strFname As String              'Last name of the person
Dim strAddress As String            'Address of the person
Dim strCity As String               'City of the person
Dim strState As String              'State of the person
Dim strZip As String                'Zip code of the person
Dim strPhone As String              'Telephone number of the person
Dim strEmail As String              'Email address of the person
Dim dteDOB As Date                  'Person's date of birth
Dim strLicense As String            'Person's drivers license

Dim intMsgBox As Integer            'Used to retrieve message box information
Dim strSQL As String                'SQL string
Dim rstInsert As ADODB.Recordset    'Recordset with the person's information

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'Checks to see if all required fields are populated with valid data
If txtLname.Text = "" Or txtFname.Text = "" Or txtAddress.Text = "" _
                   Or txtCity.Text = "" Or txtState.Text = "" Or txtZip.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the personal information fields!", vbOKOnly, "Error")
    Exit Sub
End If

If Verify_Data.Check_Zip(txtZip.Text) = True Then
    strZip = txtZip.Text
Else
    intMsgBox = MsgBox("Please enter a valid zip code!" & Chr(13) & "Valid formats are ##### or #####-####.", vbOKOnly, "Invalid Zip Code")
    Exit Sub
End If

If Verify_Data.Check_Phone(txtPhone.Text) = True Then
    strPhone = txtPhone.Text
Else
    intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

strLname = Replace(txtLname.Text, "'", "''")
strFname = Replace(txtFname.Text, "'", "''")
strAddress = Replace(txtAddress.Text, "'", "''")
strCity = Replace(txtCity.Text, "'", "''")
strState = Replace(txtState.Text, "'", "''")
strEmail = Replace(txtEmail.Text, "'", "''")
strLicense = Replace(txtLicense.Text, "'", "''")
dteDOB = dtpDOB.Value

'Updates person if there is a match found or inserts if one wasn't found

If bolMatchFound = True Then
intMsgBox = MsgBox("Update database with new values?", vbYesNo, "Update?")
    If intMsgBox = 6 Then
                        
        strSQL = "UPDATE PERSON SET PERSON_LNAME = '" & strLname & "', "
        strSQL = strSQL & "PERSON_FNAME = '" & strFname & "', "
        strSQL = strSQL & "PERSON_ADDRESS = '" & strAddress & "', "
        strSQL = strSQL & "PERSON_CITY = '" & strCity & "', "
        strSQL = strSQL & "PERSON_STATE = '" & strState & "', "
        strSQL = strSQL & "PERSON_ZIP = '" & strZip & "', "
        strSQL = strSQL & "PERSON_TELEPHONE = '" & strPhone & "', "
        strSQL = strSQL & "PERSON_EMAIL = '" & strEmail & "', "
        strSQL = strSQL & "PERSON_LICENSE = '" & strLicense & "', "
        strSQL = strSQL & "PERSON_DOB = '" & dteDOB & "' "
        strSQL = strSQL & "WHERE PERSON_NUMBER = " & intPersonNum
        Open_Recordsets.objConnection.Execute (strSQL)
    End If

Else

intMsgBox = MsgBox("Add new person to database?", vbYesNo, "Add?")
    If intMsgBox = 6 Then

        strSQL = "INSERT INTO PERSON (PERSON_LNAME, "
        strSQL = strSQL & "PERSON_FNAME, "
        strSQL = strSQL & "PERSON_ADDRESS, "
        strSQL = strSQL & "PERSON_CITY, "
        strSQL = strSQL & "PERSON_STATE, "
        strSQL = strSQL & "PERSON_ZIP, "
        strSQL = strSQL & "PERSON_TELEPHONE, "
        strSQL = strSQL & "PERSON_EMAIL, "
        strSQL = strSQL & "PERSON_LICENSE, "
        strSQL = strSQL & "PERSON_DOB)"
        strSQL = strSQL & "VALUES ('"
        strSQL = strSQL & strLname & "', '" & strFname & "', '" & strAddress & "', '" & strCity & _
             "', '" & strState & "', '" & strZip & "', '" & strPhone & "', '" & strEmail & "', '" & strLicense & "', '" & dteDOB & "')"
        Open_Recordsets.objConnection.Execute (strSQL)
        
        'Selects the person's number if they were a new person in the system
                
        strSQL = "SELECT PERSON_NUMBER FROM PERSON WHERE PERSON_LNAME = '" & strLname & "' AND PERSON_FNAME = '" & strFname & "'"

        Set rstInsert = Open_Recordsets.objConnection.Execute(strSQL)
        
        If rstInsert.EOF = False Then
            With rstInsert
                rstInsert.MoveFirst
                Do While Not rstInsert.EOF
                    If Not IsNull(![PERSON_NUMBER]) Then
                        intPersonNum = ![PERSON_NUMBER]
                    Else
                        intPersonNum = 0
                    End If
                    rstInsert.MoveNext
                Loop
            End With
        End If
    End If
End If

'Updates the person table and marks them as "Bad"

Open_Recordsets.objConnection.Execute ("UPDATE PERSON SET PERSON_BAD = -1 WHERE PERSON_NUMBER = " & intPersonNum)
frmPCHS_Main.Show
Unload Me

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Public Sub cmdSearchPeople_Click()
'********************************************************************************************
'* Runs after the first and last names have been entered, searches through the person table
'* and populates the other boxes if a match is found.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-21-2002
'********************************************************************************************
Dim strFname As String              'First name of the person
Dim strLname As String              'Last name of the person
Dim strAddress As String            'Address of the person
Dim strCity As String               'City of the person
Dim strState As String              'State of the person
Dim strZip As String                'Zip code of the person
Dim strPhone As String              'Telephone number of the person
Dim strEmail As String              'Email address of the person
Dim strLicense As String            'Drivers license of the person
Dim dteDOB As Date                  'Person's date of birth

Dim intMsgBox As Integer            'Used to retrieve message box information
Dim strSQL As String                'SQL string
Dim rstSearch As ADODB.Recordset    'Recordset with the person's information

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset

strSQL = "SELECT * FROM PERSON WHERE PERSON_LNAME = '" & txtLname.Text & "'"

Set rstSearch = Open_Recordsets.objConnection.Execute(strSQL)

If rstSearch.EOF <> True Then
    
    'Displays the people form with matches, user selects a match
    
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 11
    
    frmPeople.Show (vbModal)

    'If a match is chosen the Search_Person module populates the text boxes with appropriate data
        
    If intPersonNum <> 0 Then
        bolMatchFound = True
        strFname = Replace(txtFname.Text, "'", "''")
        strLname = Replace(txtLname.Text, "'", "''")

        Call Search_Person.Search_Person(strFname, strLname, strAddress, strCity, strState, strZip, strPhone, strEmail, strLicense, dteDOB, intPersonNum)
        
        txtFname.Text = strFname
        txtLname.Text = strLname
        txtAddress.Text = strAddress
        txtCity.Text = strCity
        txtState.Text = strState
        txtZip.Text = strZip
        txtLicense.Text = strLicense
        dtpDOB.Value = dteDOB
        txtPhone.Text = strPhone
        txtEmail.Text = strEmail
    End If
Else
    intMsgBox = MsgBox("No matches found.", vbOKOnly, "Not Found")
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbExclamation, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub Form_Load()
dtpDOB.Value = Date
End Sub

Private Sub mnuAbout_Click()
Call About.About
End Sub

Private Sub mnuBack_Click()
frmListAnimals.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()
frmPCHS_Main.Show
Unload Me
End Sub

Private Sub mnuExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub
