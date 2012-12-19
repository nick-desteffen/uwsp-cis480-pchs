VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditAnimal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Animal"
   ClientHeight    =   7350
   ClientLeft      =   630
   ClientTop       =   1200
   ClientWidth     =   5265
   ControlBox      =   0   'False
   Icon            =   "frmEditAnimal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5265
   Begin VB.CheckBox chkDeclaw 
      Caption         =   "Declawed"
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CheckBox chkNeuter 
      Caption         =   "Spayed/Neutered"
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Frame frmVacc 
      Caption         =   "Vaccinations"
      Height          =   1815
      Left            =   1080
      TabIndex        =   14
      Top             =   3240
      Width           =   3135
      Begin MSComCtl2.DTPicker dtpBordetella 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   48365569
         CurrentDate     =   37579
      End
      Begin MSComCtl2.DTPicker dtpDeworm 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   48365569
         CurrentDate     =   37579
      End
      Begin MSComCtl2.DTPicker dtpRabies 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   48365569
         CurrentDate     =   37579
      End
      Begin MSComCtl2.DTPicker dtpVacc 
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   48365569
         CurrentDate     =   37579
      End
      Begin VB.Label lblBordetella 
         Caption         =   " Bordetella:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblDeoworm 
         Caption         =   "Dewormed:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDuramune 
         Caption         =   "Duramune:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRabies 
         Caption         =   "Rabies:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmEditAnimal.frx":08CA
      Left            =   1560
      List            =   "frmEditAnimal.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtTemperment 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtComments 
      Height          =   495
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5880
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblMixed 
      Caption         =   "Mixed: "
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblAnimalBreed 
      Caption         =   "Animal Breed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblAnimalType 
      Caption         =   "Animal Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblAnimalNum 
      Caption         =   "Animal Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblTemper 
      Caption         =   "Temperment "
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblComments 
      Caption         =   "Staff Comments"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
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
Attribute VB_Name = "frmEditAnimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************************
'* This form is used to change information about an active animal.  Vaccinations, name,
'* temperment, comments, spay/neuter, declawed can all be changed via this form.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'*******************************************************************************************

Public intAnimalNum As Integer          'Number of the animal being edited

'********************************************************************************************
'* Runs when the user clicks save, saves the new information about the selected animal to the
'* database.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cmdSave_Click()

Dim strName As String                   'Name of the animal
Dim strTemper As String                 'Temperment of the animal
Dim strComments As String               'Staff comments about the animal
Dim strStatus As String                 'Status of the animal
Dim intNeuter As Integer                'Whether the animal is neutered or not
Dim intDeclaw As Integer                'Whether the animal is declawed or not
Dim dteVacc As Date                     'Date the animal recieved felovac or duramane
Dim dteRabies As Date                   'Date the animal recieved a rabies vaccination
Dim dteDeworm As Date                   'Date the animal was dewormed
Dim dteBordetella As Date               'Date the animal recieved bordetella

Dim intMsgBox As Integer                'Used for message boxes
Dim strSQL As String                    'SQL statement
Dim rstInsert As ADODB.Recordset        'Recordset used for updating the database

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'Checks to see if all required fields are populated
If cboStatus.Text = "" Then
    intMsgBox = MsgBox("Please adjust the animal's status!", vbOKOnly, "Error")
    Exit Sub
End If

strName = Replace(txtName.Text, "'", "''")
strTemper = Replace(txtTemperment.Text, "'", "''")
strComments = Replace(txtComments.Text, "'", "''")
strStatus = Left(cboStatus.Text, 1)

If chkNeuter.Value = 1 Then: intNeuter = -1
If chkDeclaw.Value = 1 Then: intDeclaw = -1

'Updates all the information except the vaccinations

strSQL = "UPDATE ANIMALS SET ANIMAL_NAME = '" & strName
strSQL = strSQL & "', ANIMAL_TEMPERMENT = '" & strTemper
strSQL = strSQL & "', ANIMAL_COMMENTS = '" & strComments
strSQL = strSQL & "', ANIMAL_STATUS = '" & strStatus
strSQL = strSQL & "', ANIMAL_DATE_STATUS = '" & Date
strSQL = strSQL & "', ANIMAL_SPAY_NEUTER = " & intNeuter
strSQL = strSQL & ", ANIMAL_DECLAWED = " & intDeclaw
strSQL = strSQL & " WHERE ANIMAL_NUMBER = " & intAnimalNum

intMsgBox = MsgBox("Apply changed information to animal database?", vbYesNo, "Save changes?")
If intMsgBox = 6 Then
    
    'Updates all the vaccinations separately because a date can be set to null
    
    If Not IsNull(dtpVacc.Value) Then
        dteVacc = dtpVacc.Value
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_VACC = '" & dteVacc & "' WHERE ANIMAL_NUMBER = " & intAnimalNum)
    Else
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_VACC = NULL WHERE ANIMAL_NUMBER = " & intAnimalNum)
    End If
    
    If Not IsNull(dtpRabies.Value) Then
        dteRabies = dtpRabies.Value
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_RABIES = '" & dteRabies & "' WHERE ANIMAL_NUMBER = " & intAnimalNum)
    Else
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_RABIES = NULL WHERE ANIMAL_NUMBER = " & intAnimalNum)
    End If
    
    If Not IsNull(dtpDeworm.Value) Then
        dteDeworm = dtpDeworm.Value
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_DEWORM = '" & dteDeworm & "' WHERE ANIMAL_NUMBER = " & intAnimalNum)
    Else
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_DEWORM = NULL WHERE ANIMAL_NUMBER = " & intAnimalNum)
    End If
    
    If Not IsNull(dtpBordetella.Value) Then
        dteBordetella = dtpBordetella.Value
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_BORDETELLA = '" & dteBordetella & "' WHERE ANIMAL_NUMBER = " & intAnimalNum)
    Else
        Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_BORDETELLA = NULL WHERE ANIMAL_NUMBER = " & intAnimalNum)
    End If
    
    Open_Recordsets.objConnection.Execute (strSQL)
    intMsgBox = MsgBox("Animal record has been updated in database!", vbOKOnly, "New record added")
    Unload Me
    frmListAnimals.Show
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

Dim intType As Integer              'Type number of the animal
Dim strStatus As String             'Status of the animal
Dim intBreed As Integer             'Breed of the animal

Dim strSQL As String                'SQL statement
Dim rstType As ADODB.Recordset      'Recordset used for updating the database
Dim intMsgBox As Integer            'Used for message boxes

On Error GoTo ErrorHandler

Set rstType = New ADODB.Recordset

intAnimalNum = frmListAnimals.dgdCurrentAnimals.Columns("Number").CellValue(frmListAnimals.dgdCurrentAnimals.Bookmark)

cboStatus.AddItem "D - Deceased", 0
cboStatus.AddItem "E - Euthanized", 1
cboStatus.AddItem "R - Residing", 2
cboStatus.AddItem "C - Reclaimed", 3
cboStatus.AddItem "P - Pending Adoption", 4

'Populates all the fields with information from the selected animal

strSQL = "SELECT ANIMAL_NAME, "
strSQL = strSQL & "ANIMAL_BREED, "
strSQL = strSQL & "ANIMAL_TYPE, "
strSQL = strSQL & "ANIMAL_TEMPERMENT, "
strSQL = strSQL & "ANIMAL_COMMENTS, "
strSQL = strSQL & "ANIMAL_MIX, "
strSQL = strSQL & "ANIMAL_SPAY_NEUTER, "
strSQL = strSQL & "ANIMAL_DECLAWED, "
strSQL = strSQL & "ANIMAL_VACC, "
strSQL = strSQL & "ANIMAL_RABIES, "
strSQL = strSQL & "ANIMAL_DEWORM, "
strSQL = strSQL & "ANIMAL_BORDETELLA, "
strSQL = strSQL & "SWITCH([ANIMAL_STATUS]='A', 'A - Adopted', [ANIMAL_STATUS]='D', 'D - Deceased', [ANIMAL_STATUS]='E', 'E - Euthanized', [ANIMAL_STATUS]='R', 'R - Residing', [ANIMAL_STATUS]='C', 'C - Reclaimed',[ANIMAL_STATUS]='P', 'P - Pending Adoption', True, 'Status is invalid') AS STATUS,"
strSQL = strSQL & "ANIMAL_NUMBER "
strSQL = strSQL & "FROM ANIMALS "
strSQL = strSQL & "WHERE ANIMAL_NUMBER = " & intAnimalNum

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
With rstType
    rstType.MoveFirst
    Do While Not rstType.EOF
        
        If Not IsNull(![ANIMAL_NAME]) Then
            txtName.Text = ![ANIMAL_NAME]
        Else
            txtName.Text = "None"
        End If
    
        If Not IsNull(![ANIMAL_BREED]) Then
            intBreed = ![ANIMAL_BREED]
        Else
            lblAnimalBreed.Caption = ""
        End If
        
        If Not IsNull(![animal_type]) Then
            intType = ![animal_type]
        Else
            intType = 0
        End If
            
        If Not IsNull(![ANIMAL_TEMPERMENT]) Then
            txtTemperment.Text = ![ANIMAL_TEMPERMENT]
        Else
            txtTemperment.Text = ""
        End If
        
        If Not IsNull(![ANIMAL_COMMENTS]) Then
            txtComments.Text = ![ANIMAL_COMMENTS]
        Else
            txtComments.Text = ""
        End If
        
        If Not IsNull(![animal_number]) Then
            lblAnimalNum.Caption = "Animal Number: " & ![animal_number]
        Else
            lblAnimalNum.Caption = "Animal Number: "
        End If
        
        If Not IsNull(![Status]) Then
            cboStatus.Text = ![Status]
        Else
            strStatus = ""
        End If
        
        If ![ANIMAL_MIX] = -1 Then
            lblMixed.Caption = "Mixed: Yes"
        Else
            lblMixed.Caption = "Mixed: No"
        End If
        
        If Not IsNull(![ANIMAL_VACC]) Then
            dtpVacc.Value = ![ANIMAL_VACC]
        Else
            dtpVacc.Value = Null
        End If
        
        If Not IsNull(![ANIMAL_RABIES]) Then
            dtpRabies.Value = ![ANIMAL_RABIES]
        Else
            dtpRabies.Value = Null
        End If
            
        If Not IsNull(![ANIMAL_DEWORM]) Then
            dtpDeworm.Value = ![ANIMAL_DEWORM]
        Else
            dtpDeworm.Value = Null
        End If
        
        If Not IsNull(![ANIMAL_BORDETELLA]) Then
            dtpBordetella.Value = ![ANIMAL_BORDETELLA]
        Else
            dtpBordetella.Value = Null
        End If
        
        If ![ANIMAL_SPAY_NEUTER] = -1 Then
            chkNeuter.Value = 1
        End If
        
        If ![ANIMAL_DECLAWED] = -1 Then
            chkDeclaw.Value = 1
        End If
            
    rstType.MoveNext
    Loop
End With
End If

Set rstType = Nothing

'If the type is not a cat then declawed is disabled

If intType <> 2 Then
    chkDeclaw.Enabled = False
End If

'Selects the type of animal and displays it in a label

strSQL = "SELECT TYPE_NAME FROM ANIMAL_TYPES WHERE TYPE_NUMBER = " & intType

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
With rstType
    rstType.MoveFirst
    If Not IsNull(![TYPE_NAME]) Then
        lblAnimalType.Caption = "Type: " & ![TYPE_NAME]
    Else
        lblAnimalType.Caption = "Type: "
    End If
End With
End If

'If the animal is a cat or dog selects the breed of the animal and displays it in a label

If intType = 1 Or intType = 2 Then
    If intType = 1 Then
        strSQL = "SELECT BREED_NAME FROM DOG_BREEDS WHERE BREED_NUMBER = " & intBreed
        lblDuramune.Caption = "Duramune:"
    ElseIf intType = 2 Then
        strSQL = "SELECT BREED_NAME FROM CAT_BREEDS WHERE BREED_NUMBER = " & intBreed
        lblDuramune.Caption = "Felovac:"
        lblBordetella.Enabled = False
        dtpBordetella.Enabled = False
    End If

    Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

    If rstType.EOF = False Then
    With rstType
        rstType.MoveFirst
        If Not IsNull(![BREED_NAME]) Then
            lblAnimalBreed.Caption = "Animal Breed: " & ![BREED_NAME]
        Else
            lblAnimalBreed.Caption = ""
        End If
    End With
    End If
Else
    lblAnimalBreed.Caption = ""
End If

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

Private Sub cmdCancel_Click()
Unload Me
frmListAnimals.Show
End Sub
