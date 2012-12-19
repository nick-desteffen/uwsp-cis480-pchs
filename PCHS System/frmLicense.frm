VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLicense 
   Caption         =   "New License"
   ClientHeight    =   7995
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6165
   Icon            =   "frmLicense.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVet 
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox txtLicenseNum 
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox txtMunicipality 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearchName 
      Caption         =   "Search People"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   37
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4800
      TabIndex        =   20
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame frmRequest 
      Caption         =   "Licensed Animal Information"
      Height          =   2655
      Left            =   120
      TabIndex        =   32
      Top             =   3480
      Width           =   5895
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   43
         Top             =   360
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpRabies 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   41
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   47972353
         CurrentDate     =   37593
      End
      Begin VB.CheckBox chkNeuter 
         Caption         =   "Animal is spayed or neutered"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox cboBreed 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cboAnimalType 
         Height          =   315
         ItemData        =   "frmLicense.frx":08CA
         Left            =   1440
         List            =   "frmLicense.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Frame frmSex 
         Caption         =   "Sex of Animal"
         Height          =   1215
         Left            =   4080
         TabIndex        =   33
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Name of Animal"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblRabies 
         Caption         =   "Date rabies vacc was administered:"
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblAnimalType 
         Caption         =   "Type of Animal"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblBreed 
         Caption         =   "Breed of Animal"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblColor 
         Caption         =   "Color of Animal"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame frmPersonal 
      Caption         =   "Person Information"
      Height          =   3255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox txtFname 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtLname 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
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
         Format          =   47972353
         CurrentDate     =   37579
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCIty 
         Caption         =   "City"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1095
      End
   End
   Begin VB.Label lblVet 
      Caption         =   "Veternarian:"
      Height          =   255
      Left            =   960
      TabIndex        =   40
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label lblLicenseNum 
      Caption         =   "License Number:"
      Height          =   255
      Left            =   600
      TabIndex        =   39
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblMunicipality 
      Caption         =   "Municipality of license:"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   6240
      Width           =   1695
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
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************
'* This form generates a license for a dog or cat and stores the information about the animal
'* and owner in the database.
'*
'* Written by: Nick DeSteffen
'* Written on: 12-04-2002
'********************************************************************************************

'********************************************************************************************
'* Array that contains all the types of animals
'********************************************************************************************
Private Type combo_info
    name As String
    Number As Integer
End Type

Public bolMatchFound As Boolean            'True = person already in database
Public intPersonNum As Integer             'Number of the person if found
Public intLicenseNum As Integer         'Animal's license number
'********************************************************************************************
'* Called when the value in the animal type combo box is changed
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cboAnimalType_Click()

cboBreed.Clear

Dim breeds() As combo_info          'Array containing breeds

Dim rstType As ADODB.Recordset      'Used for interfacing with the database
Dim looper As Integer               'Loop control variable
Dim strSQL As String                'SQL Statement
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
'* Runs when the user clicks save, saves the new animal license to the database
'*
'* Written by: Nick DeSteffen
'* Written on: 12-03-2002
'********************************************************************************************
Private Sub cmdSave_Click()

Dim strLname As String              'Last name of the person
Dim strFname As String              'First name of the person
Dim strAddress As String            'Address of the person
Dim strCity As String               'City of the person
Dim strState As String              'State of the person
Dim strZip As String                'Zip code of the person
Dim strPhone As String              'Telephone number of the person
Dim intType As Integer              'Type of animal
Dim intBreed As Integer             'Breed of the animal
Dim strSex As String                'Sex of the animal
Dim intColor As Integer             'Color of the animal
Dim strEmail As String              'Email address of the person
Dim dteDOB As Date                  'Date of birth of the person
Dim strLicense As String            'Drivers license of the person
Dim strMunicipality As String       'Municipality of the license
Dim strVet As String                'Vet of the animal
Dim dteRabies As Date               'Date the animal recieved a rabies vaccination
Dim intNeuter As Integer            'Whether or not the animal is neutered
Dim strName As String               'Name of the animal

Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL statement
Dim rstInsert As ADODB.Recordset    'Recordset used for updating the database

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'Checks to see if all required fields are populated
If txtLname.Text = "" Or txtFname.Text = "" Or txtAddress.Text = "" _
                   Or txtCity.Text = "" Or txtState.Text = "" Or txtZip.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the required fields!", vbOKOnly, "Error")
    Exit Sub
End If
If cboColor.Text = "" Or cboAnimalType.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the required fields!", vbOKOnly, "Error")
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
strMunicipality = Replace(txtMunicipality.Text, "'", "''")
strEmail = Replace(txtEmail.Text, "'", "''")
strLicense = Replace(txtLicense.Text, "'", "''")
dteDOB = dtpDOB.Value

strVet = Replace(txtVet.Text, "'", "''")
strName = Replace(txtName.Text, "'", "''")
If chkNeuter.Value = 1 Then: intNeuter = -1

If IsNumeric(txtLicenseNum.Text) = True And txtLicenseNum.Text <> "" Then
    intLicenseNum = txtLicenseNum.Text
Else
    intMsgBox = MsgBox("Please enter a valid number for the license number.", vbOKOnly, "Invalid number")
    Exit Sub
End If

'Chooses the sex of the animal
If optMale.Value = True Then
    strSex = "M"
ElseIf optFemale.Value = True Then
    strSex = "F"
Else
    intMsgBox = MsgBox("Please select the sex of the animal!", vbOKOnly, "Error")
    Exit Sub
End If

'Calls functions to return the type, color, and breed numbers of the animal
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

'Updates the license table with the new animal license

If intPersonNum <> 0 And intMsgBox = 6 Then

    If Not IsNull(dtpRabies.Value) Then
        dteRabies = dtpRabies.Value
        
        strSQL = "INSERT INTO LICENSE (LICENSE_OWNER, "
        strSQL = strSQL & "LICENSE_TYPE, "
        strSQL = strSQL & "LICENSE_BREED, "
        strSQL = strSQL & "LICENSE_SEX, "
        strSQL = strSQL & "LICENSE_COLOR, "
        strSQL = strSQL & "LICENSE_NUMBER, "
        strSQL = strSQL & "LICENSE_SPAYED, "
        strSQL = strSQL & "LICENSE_MUNICIPALITY, "
        strSQL = strSQL & "LICENSE_RABIES, "
        strSQL = strSQL & "LICENSE_NAME, "
        strSQL = strSQL & "LICENSE_VET) "
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & intPersonNum & ", " & intType & ", " & intBreed & ", '" & strSex
        strSQL = strSQL & "', " & intColor & ", " & intLicenseNum & ", " & intNeuter & ", '"
        strSQL = strSQL & strMunicipality & "', '" & dteRabies & "', '" & strName & "', '" & strVet & "')"
    Else
        strSQL = "INSERT INTO LICENSE (LICENSE_OWNER, "
        strSQL = strSQL & "LICENSE_TYPE, "
        strSQL = strSQL & "LICENSE_BREED, "
        strSQL = strSQL & "LICENSE_SEX, "
        strSQL = strSQL & "LICENSE_COLOR, "
        strSQL = strSQL & "LICENSE_NUMBER, "
        strSQL = strSQL & "LICENSE_SPAYED, "
        strSQL = strSQL & "LICENSE_MUNICIPALITY, "
        strSQL = strSQL & "LICENSE_RABIES, "
        strSQL = strSQL & "LICENSE_NAME, "
        strSQL = strSQL & "LICENSE_VET) "
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & intPersonNum & ", " & intType & ", " & intBreed & ", '" & strSex
        strSQL = strSQL & "', " & intColor & ", " & intLicenseNum & ", " & intNeuter & ", '"
        strSQL = strSQL & strMunicipality & "', Null, '" & strName & "', '" & strVet & "')"
    End If
    
    Open_Recordsets.objConnection.Execute (strSQL)
    
    'Displays the new license and calls the new receipt function
    
    frmShowLicense.intLicense = intLicenseNum
    frmShowLicense.Show
    frmNewReciept.intPersonNum = intPersonNum
    frmNewReciept.cboReason.ListIndex = 3
    frmNewReciept.intType = 3
    frmNewReciept.intNumber = intLicenseNum
    frmPCHS_Main.Show
    frmNewReciept.Show
    
ElseIf intMsgBox <> 6 Then
    Exit Sub
Else
    intMsgBox = MsgBox("There was an error, please restart the animal license form!", vbCritical, "Error")
    Exit Sub
End If
Unload Me

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
    
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 8
    
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

Dim types() As combo_info           'Array of animal types
Dim colors() As combo_info          'Array of animal colors
Dim looper As Integer               'Loop control variable

Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL Statement

On Error GoTo ErrorHandler

dtpDOB.Value = Date
Set rstType = New ADODB.Recordset
looper = 0
dtpRabies.Value = Null


'Populates the color recordset
Set rstType = Nothing
looper = 0
strSQL = "SELECT COLOR_NUMBER, COLOR_NAME FROM COLOR"

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

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

