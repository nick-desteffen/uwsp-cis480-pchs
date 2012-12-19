VERSION 5.00
Begin VB.Form frmNewAnimal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Animal"
   ClientHeight    =   7020
   ClientLeft      =   330
   ClientTop       =   615
   ClientWidth     =   4575
   ClipControls    =   0   'False
   Icon            =   "frmNewAnimal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   4575
   Begin VB.TextBox txtFound 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   4800
      Width           =   2295
   End
   Begin VB.ComboBox cboAge 
      Height          =   315
      ItemData        =   "frmNewAnimal.frx":08CA
      Left            =   1440
      List            =   "frmNewAnimal.frx":08D7
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CheckBox chkMix 
      Caption         =   "Mixed breed"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox chkDeclawed 
      Caption         =   "Declawed"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtTemperment 
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtComments 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CheckBox chkNeutor 
      Caption         =   "Spayed/Neutered"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   3000
      TabIndex        =   16
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Frame frmAcquired 
      Caption         =   "How Acquired"
      Height          =   1815
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton optStray 
         Caption         =   "Stray"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optSurrender 
         Caption         =   "Surrendered"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optImpounded 
         Caption         =   "Impounded"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optAbandon 
         Caption         =   "Abandoned"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame frmSex 
      Caption         =   "Sex of Animal"
      Height          =   1215
      Left            =   2640
      TabIndex        =   21
      Top             =   360
      Width           =   1335
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ComboBox cboBreed 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmNewAnimal.frx":08F6
      Left            =   1440
      List            =   "frmNewAnimal.frx":08F8
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3360
      Width           =   2295
   End
   Begin VB.ComboBox cboAnimalType 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblFound 
      Caption         =   "Where Found"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblTemper 
      Caption         =   "Temperment"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblComments 
      Caption         =   "Staff Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblAge 
      Caption         =   "Approximate Age"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblColor 
      Caption         =   "Color"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label lblBreed 
      Caption         =   "Breed"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuBack 
         Caption         =   "Back"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmNewAnimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************
'* This form is used when a new animal is entered into the system.  The user chooses its
'* characteristics and how it was acquired and saves it to the system.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************

'********************************************************************************************
'* Array that contains all the types of animals
'********************************************************************************************
Private Type combo_info
    name As String
    Number As Integer
End Type
Public intNum As Integer               'Number of the animal

'********************************************************************************************
'* Called when the value in the animal type combo box is changed
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cboAnimalType_Click()

Dim breeds() As combo_info          'Array of animal breeds
Dim rstType As ADODB.Recordset      'Used for interfacing with the database
Dim looper As Integer               'loop control variable
Dim strSQL As String                'SQL Statement
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

cboBreed.Clear
chkMix.Enabled = False
chkDeclawed.Enabled = False
chkMix.Value = 0
chkDeclawed.Value = 0
cboBreed.Enabled = False

Set rstType = New ADODB.Recordset

'Populates dog breed recordset
If cboAnimalType.Text = "Dog" Or cboAnimalType.Text = "Cat" Then
    
    cboBreed.Enabled = True
    chkMix.Enabled = True

    looper = 0
    Set rstType = Nothing

    If cboAnimalType.Text = "Dog" Then
        strSQL = "SELECT BREED_NUMBER, BREED_NAME FROM DOG_BREEDS"
    ElseIf cboAnimalType.Text = "Cat" Then
        strSQL = "SELECT BREED_NUMBER, BREED_NAME FROM CAT_BREEDS"
        chkDeclawed.Enabled = True
    End If

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
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub
'********************************************************************************************
'* Runs when the user clicks save, saves the new animal to the database
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cmdSave_Click()

Dim strAcquired As String           'How the animal was acquired
Dim strName As String               'Name of the animal
Dim intType As Integer              'Type of the animal
Dim intBreed As Integer             'Breed of the animal
Dim intMixed As Integer             'Whether or not the animal is mixed
Dim strSex As String                'Sex of the animal
Dim intNeutor As Integer            'Whether or not the animal is neutered
Dim intDeclawed As Integer          'Whether or not the animal is declawed
Dim strAge As String                'Approximate age of the animal
Dim intColor As Integer             'Color of the animal
Dim strFound As String              'Where the animal was found if stray or abandoned
Dim strTemper As String             'Temperment of the animal
Dim strComments As String           'Staff comments about the animal

Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL statement

On Error GoTo ErrorHandler

'Checks to see if all required fields are populated
If cboColor.Text = "" Or cboAge.Text = "" Or cboAnimalType.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the required fields!", vbOKOnly, "Error")
    Exit Sub
End If
If ((cboAnimalType.Text = "Dog") Or (cboAnimalType.Text = "Cat")) And (cboBreed.Text = "") Then
        intMsgBox = MsgBox("Please choose the breed of the animal!", vbOKOnly, "Error")
        Exit Sub
End If

'Chooses how the animal was acquired
If optStray.Value = True Then
    strAcquired = "T"
ElseIf optSurrender.Value = True Then
    strAcquired = "S"
ElseIf optImpounded.Value = True Then
    strAcquired = "I"
ElseIf optAbandon.Value = True Then
    strAcquired = "A"
Else
    intMsgBox = MsgBox("Please select how the animal was acquired!", vbOKOnly, "Error")
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

strName = Replace(txtName.Text, "'", "''")
strTemper = Replace(txtTemperment.Text, "'", "''")
strAge = cboAge.Text
strComments = Replace(txtComments.Text, "'", "''")
strFound = Replace(txtFound.Text, "'", "''")

If chkNeutor.Value = 1 Then: intNeutor = -1
If chkDeclawed.Value = 1 Then: intDeclawed = -1
If chkMix.Value = 1 Then: intMixed = -1

'Returns the type, breed, and color number of the animal

intType = Get_Types.Get_Types(cboAnimalType.Text)
intColor = Get_Colors.Get_Colors(cboColor.Text)
If intType = 1 Or intType = 2 Then
    intBreed = Get_Breeds.Get_Breeds(intType, cboBreed.Text)
Else
    intBreed = 0
End If

'Inserts new animal into database
strSQL = "INSERT INTO ANIMALS (ANIMAL_NAME, "
strSQL = strSQL & "ANIMAL_TYPE, "
strSQL = strSQL & "ANIMAL_BREED, "
strSQL = strSQL & "ANIMAL_SEX, "
strSQL = strSQL & "ANIMAL_COLOR, "
strSQL = strSQL & "ANIMAL_AGE, "
strSQL = strSQL & "ANIMAL_TEMPERMENT, "
strSQL = strSQL & "ANIMAL_SPAY_NEUTER, "
strSQL = strSQL & "ANIMAL_STATUS, "
strSQL = strSQL & "ANIMAL_COMMENTS, "
strSQL = strSQL & "ANIMAL_ACQUIRED, "
strSQL = strSQL & "ANIMAL_MIX, "
strSQL = strSQL & "ANIMAL_DECLAWED, "
strSQL = strSQL & "ANIMAL_FOUND) "
strSQL = strSQL & "VALUES ('"
strSQL = strSQL & strName & "', " & intType & ", " & intBreed & ", '" & strSex & "', " & intColor & ", '" & strAge & "', '"
strSQL = strSQL & strTemper & "', " & intNeutor & ", 'R', '" & strComments & "', '" & strAcquired & "', " & intMixed & ", " & intDeclawed & ", '" & strFound & "')"

intMsgBox = MsgBox("Add new animal to database?", vbYesNo, "Add new animal?")
If intMsgBox = 6 Then
     Open_Recordsets.objConnection.Execute (strSQL)
    Unload Me
    
    'If the animal was surrendered then additional information is filled out
    
    If strAcquired = "S" Then
        intMsgBox = MsgBox("Please fill out additional information about the surrendered animal.", vbOKOnly, "Surrendered Animal")
        intNum = Get_IDs.Get_Animal_Num(strName, intType, intBreed, strSex, intColor, strAge, strTemper, intNeutor, strAcquired, intMixed)
        frmSurrender.Show
    Else
        intMsgBox = MsgBox("One new animal record added to database!", vbOKOnly, "New record added")
        frmPCHS_Main.Show
    End If
    
    Call Search_Missing.Search_Missing(intType, intBreed, intColor, strAge, strSex)
    Call Search_Request.Search_Requests(intType, intBreed, intColor, strAge, strSex)
ElseIf intMsgBox <> 6 Then
    Exit Sub
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
Dim looper As Integer               'looper used for arrays

Dim rstType As ADODB.Recordset      'Recordset used for inserting into the database
Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL statement

On Error GoTo ErrorHandler

Set rstType = New ADODB.Recordset
looper = 0

'Populates the animal type recordset
strSQL = "SELECT TYPE_NUMBER, TYPE_NAME FROM ANIMAL_TYPES"

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)
ReDim Preserve types(looper)
If rstType.EOF = False Then
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
End If

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
frmPCHS_Main.Show
Unload Me
End Sub

Private Sub mnuExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub

Private Sub optAbandon_Click()
If optAbandon.Value = True Then
    txtFound.Enabled = True
End If
End Sub

Private Sub optImpounded_Click()
If optImpounded.Value = True Then
    txtFound.Enabled = False
End If
End Sub

Private Sub optStray_Click()
If optStray.Value = True Then
    txtFound.Enabled = True
End If
End Sub

Private Sub optSurrender_Click()
If optSurrender.Value = True Then
    txtFound.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmPCHS_Main.Show
End Sub
