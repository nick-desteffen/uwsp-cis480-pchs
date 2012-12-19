VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNewMissing 
   Caption         =   "New Missing Animal"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6165
   Icon            =   "frmNewMissing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearchName 
      Caption         =   "Search People"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame frmRequest 
      Caption         =   "Requested Animal Information"
      Height          =   2415
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   5895
      Begin VB.ComboBox cboColor 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Frame frmSex 
         Caption         =   "Sex of Animal"
         Height          =   1215
         Left            =   4080
         TabIndex        =   30
         Top             =   840
         Width           =   1335
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.ComboBox cboAnimalType 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cboBreed 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboAge 
         Height          =   315
         ItemData        =   "frmNewMissing.frx":08CA
         Left            =   1440
         List            =   "frmNewMissing.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblColor 
         Caption         =   "Color of Animal"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblBreed 
         Caption         =   "Breed of Animal"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblAnimalType 
         Caption         =   "Type of Animal"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblAge 
         Caption         =   "Age of Animal"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame frmPersonal 
      Caption         =   "Person Information"
      Height          =   3255
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   5895
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3600
         TabIndex        =   36
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47251457
         CurrentDate     =   37579
      End
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
         TabIndex        =   8
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblCIty 
         Caption         =   "City"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.Label lblMissingDate 
      Caption         =   "Date"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   240
      Width           =   2055
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
Attribute VB_Name = "frmNewMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************************
'* This form is used to enter in information about a new animal reported missing missing.
'* It saves the information to the missing table.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-05-2002
'*******************************************************************************************

'********************************************************************************************
'* Array that contains all the types of animals
'********************************************************************************************
Private Type combo_info
    name As String
    Number As Integer
End Type

Public bolMatchFound As Boolean            'True = person already in database
Public intPersonNum As Integer             'Number of the person if found
'********************************************************************************************
'* Called when the value in the animal type combo box is changed
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cboAnimalType_Click()

cboBreed.Clear

Dim breeds() As combo_info          'Array containing all the breeds
Dim looper As Integer               'Loop control variable
Dim strSQL As String                'SQL Statement

Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstType = New ADODB.Recordset

'Populates dog breed recordset
If cboAnimalType.Text = "Dog" Then
    
    cboBreed.Enabled = True

    looper = 0
    Set rstType = Nothing

    strSQL = "SELECT BREED_NUMBER, BREED_NAME FROM DOG_BREEDS"

    Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
    With rstType
        rstType.MoveFirst
        Do While Not rstType.EOF
            ReDim Preserve breeds(looper)
            If Not IsNull(![BREED_NUMBER]) Then
                breeds(looper).Number = (![BREED_NUMBER])
            End If
            If Not IsNull(![BREED_NAME]) Then
                breeds(looper).name = (![BREED_NAME])
            End If
            rstType.MoveNext
            looper = looper + 1
        Loop
    End With
End If

'Populates the combo box
    For looper = 0 To UBound(breeds)
        cboBreed.AddItem (breeds(looper).name)
    Next looper
'Closes the connection
    rstType.Close
    Set rstType = Nothing
    
    'Populates cat breed recordset

ElseIf cboAnimalType.Text = "Cat" Then

    cboBreed.Enabled = True
       
    looper = 0
    Set rstType = Nothing

    strSQL = "SELECT BREED_NUMBER, BREED_NAME FROM CAT_BREEDS"

    Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
    With rstType
        rstType.MoveFirst
        Do While Not rstType.EOF
            ReDim Preserve breeds(looper)
            If Not IsNull(![BREED_NUMBER]) Then
                breeds(looper).Number = (![BREED_NUMBER])
            End If
            If Not IsNull(![BREED_NAME]) Then
                breeds(looper).name = (![BREED_NAME])
            End If
            rstType.MoveNext
            looper = looper + 1
        Loop
    End With
End If

'Populates the combobox
    For looper = 0 To UBound(breeds)
        cboBreed.AddItem (breeds(looper).name)
    Next looper
'Closes the connection
    rstType.Close
    Set rstType = Nothing
Else
    cboBreed.Enabled = False
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Sub

'********************************************************************************************
'* Runs when the user clicks save, saves the new missing animal to the database
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cmdSave_Click()

Dim strLname As String              'Last name of the person
Dim strFname As String              'First name of the person
Dim strAddress As String            'Address of the person
Dim strCity As String               'City of the person
Dim strState As String              'State of the person
Dim strZip As String                'Zip code of the person
Dim strPhone As String              'Phone number of the person
Dim strEmail As String              'Email address of the person
Dim dteDOB As Date                  'Person's date of birth
Dim strLicense As String            'Person's drivers license number

Dim intType As Integer              'Type of animal
Dim intBreed As Integer             'Breed of animal
Dim strSex As String                'Sex of the animal
Dim intColor As Integer             'Color of the animal
Dim strAge As String                'Age of the animal

Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL Statement
Dim rstInsert As ADODB.Recordset    'Used for interfacing with the database

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'Checks to see if all required fields are populated
If txtLname.Text = "" Or txtFname.Text = "" Or txtAddress.Text = "" _
                   Or txtCity.Text = "" Or txtState.Text = "" Or txtZip.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the required fields!", vbOKOnly, "Error")
    Exit Sub
End If

If cboColor.Text = "" Then
    intMsgBox = MsgBox("Please select the color!", vbOKOnly, "Error")
    Exit Sub
End If
If cboAnimalType.Text = "" Then
    intMsgBox = MsgBox("Please select the animal type!", vbOKOnly, "Error")
    Exit Sub
End If

If cboAge.Text = "" Then
    intMsgBox = MsgBox("Please select the age!", vbOKOnly, "Error")
    Exit Sub
End If

If ((cboAnimalType.Text = "Dog") Or (cboAnimalType.Text = "Cat")) And (cboBreed.Text = "") Then
    intMsgBox = MsgBox("Please choose the breed of the animal!", vbOKOnly, "Error")
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
strAge = Replace(cboAge.Text, "'", "''")
strEmail = Replace(txtEmail.Text, "'", "''")
strLicense = Replace(txtLicense.Text, "'", "''")
dteDOB = dtpDOB.Value

'Chooses the sex of the animal
If optMale.Value = True Then
    strSex = "M"
ElseIf optFemale.Value = True Then
    strSex = "F"
Else
    intMsgBox = MsgBox("Please select the sex of the animal!", vbOKOnly, "Error")
    Exit Sub
End If

intType = Get_Types.Get_Types(cboAnimalType.Text)
intColor = Get_Colors.Get_Colors(cboColor.Text)
If intType = 1 Or intType = 2 Then
    intBreed = Get_Breeds.Get_Breeds(intType, cboBreed.Text)
Else
    intBreed = 0
End If


If bolMatchFound = True Then
intMsgBox = MsgBox("Update database with new values?", vbYesNo, "Update?")
    If intMsgBox = 6 Then
                
            'Updates person if there is a match found and values have changed
        
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
    
        'Inserts new person into database

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

If intPersonNum <> 0 And intMsgBox = 6 Then

    'Inserts new animal request into database
    strSQL = "INSERT INTO MISSING (MISSING_PERSON, "
    strSQL = strSQL & "MISSING_TYPE, "
    strSQL = strSQL & "MISSING_BREED, "
    strSQL = strSQL & "MISSING_SEX, "
    strSQL = strSQL & "MISSING_AGE, "
    strSQL = strSQL & "MISSING_COLOR) "
    strSQL = strSQL & "VALUES ("
    strSQL = strSQL & intPersonNum & ", " & intType & ", " & intBreed & ", '" & strSex & _
             "', '" & strAge & "', " & intColor & ")"
    
    Open_Recordsets.objConnection.Execute (strSQL)
ElseIf intMsgBox <> 6 Then
    Exit Sub
Else
    intMsgBox = MsgBox("There was an error, please restart the missing animal form!", vbCritical, "Error")
    Exit Sub
End If

Unload Me
frmPCHS_Main.Show

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Public Sub cmdSearchName_Click()
'********************************************************************************************
'* Runs after the first and last names have been entered, searches through the person table
'* and populates the other boxes if a match is found.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-21-2002
'********************************************************************************************
Dim strFname As String          'First name of the person
Dim strLname As String          'Last name of the person
Dim strAddress As String        'Address of the person
Dim strCity As String           'City of the person
Dim strState As String          'State of the person
Dim strZip As String            'Zip code of the person
Dim strPhone As String          'Telephone number of the person
Dim strEmail As String          'Person's email address
Dim strLicense As String        'Person's drivers license
Dim dteDOB As Date              'Person's date of birth

Dim intMsgBox As Integer                'Used for messageboxes
Dim rstSearch As New ADODB.Recordset    'Used for interfacing with the database

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset
Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM PERSON WHERE PERSON_LNAME = '" & txtLname.Text & "'")

If rstSearch.EOF <> True Then
    
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 2
    
    frmPeople.Show (vbModal)

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
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

'********************************************************************************************
'* Runs when the form is loaded.  Populates the animal type combo box.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-09-2002
'********************************************************************************************
Public Sub Form_Load()

Dim types() As combo_info           'Array containing animal types
Dim colors() As combo_info          'Array containing animal colors
Dim looper As Integer               'Loop control variable

Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL Statement

On Error GoTo ErrorHandler

lblMissingDate.Caption = "Date: " & Date
dtpDOB.Value = Date

Set rstType = New ADODB.Recordset
looper = 0

'Populates the recordset
strSQL = "SELECT TYPE_NUMBER, TYPE_NAME FROM ANIMAL_TYPES"

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)
looper = 0
With rstType
    rstType.MoveFirst
    Do While Not rstType.EOF
        ReDim Preserve types(looper)
        If Not IsNull(![TYPE_NUMBER]) Then
            types(looper).Number = (![TYPE_NUMBER])
        End If
        If Not IsNull(![TYPE_NAME]) Then
            types(looper).name = (![TYPE_NAME])
        End If
    rstType.MoveNext
    looper = looper + 1
    Loop
End With
'Populates the combo box
For looper = 0 To UBound(types)
    cboAnimalType.AddItem (types(looper).name)
Next looper

'Populates the color recordset
Set rstType = Nothing
looper = 0
strSQL = "SELECT COLOR_NUMBER, COLOR_NAME FROM COLOR"

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
With rstType
    rstType.MoveFirst
    Do While Not rstType.EOF
        ReDim Preserve colors(looper)
        If Not IsNull(![COLOR_NUMBER]) Then
            colors(looper).Number = (![COLOR_NUMBER])
        End If
        If Not IsNull(![COLOR_NAME]) Then
            colors(looper).name = (![COLOR_NAME])
            cboColor.AddItem (![COLOR_NAME])
        End If
    rstType.MoveNext
    looper = looper + 1
    Loop
End With
End If

'Closes the connection
rstType.Close
Set rstType = Nothing

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
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
