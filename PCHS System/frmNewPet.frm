VERSION 5.00
Begin VB.Form frmNewPet 
   Caption         =   "New Pet"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   Icon            =   "frmNewPet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtReason 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CheckBox chkHave 
      Caption         =   "Still have?"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CheckBox chkNeuter 
      Caption         =   "Spayed / Neutered"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox chkRabies 
      Caption         =   "Has rabies shot"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox cboAge 
      Height          =   315
      ItemData        =   "frmNewPet.frx":08CA
      Left            =   1080
      List            =   "frmNewPet.frx":08D7
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Frame frmSex 
      Caption         =   "Sex"
      Height          =   975
      Left            =   720
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frmKept 
      Caption         =   "Where Kept"
      Height          =   975
      Left            =   2280
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
      Begin VB.OptionButton optOut 
         Caption         =   "Outside"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optIn 
         Caption         =   "Inside"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Reason for not having pet anymore."
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label lblType 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblAge 
      Caption         =   "Age:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmNewPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************************
'* This form is displayed after an adoption form is filled out.  It asks which animals the
'* adoptor has had within the past 5 years and will loop until the user chooses no more pets
'* or hits cancel.
'*
'* Written by: Nick DeSteffen
'* Written on: 11-25-2002
'*******************************************************************************************

'********************************************************************************************
'* Saves the information on the form about the pet to the pet table.
'*
'* Written by: Nick DeSteffen
'* Written on: 11-26-2002
'********************************************************************************************
Private Sub cmdSave_Click()
Dim strName As String               'Name of the pet
Dim strSex As String                'Sex of the pet
Dim strAge As String                'Age of the pet
Dim intRabies As Integer            'Animal has rabies shot
Dim intNeuter As Integer            'Animal is spayed / neutered
Dim strKept As String               'Where the animal is kept
Dim intHave As Integer              'Still have the animal
Dim strReason As String             'Reason for not having the animal
Dim intType As Integer              'Type of animal

Dim intMsgBox As Integer            'Used for message box function
Dim strSQL As String                'SQL statement

On Error GoTo ErrorHandler

strName = Replace(txtName.Text, "'", "''")
strAge = cboAge.Text
If optMale.Value = True Then: strSex = "M"
If optFemale.Value = True Then: strSex = "F"

If optIn.Value = True Then: strKept = "Inside"
If optOut.Value = True Then: strKept = "Outside"

If chkRabies.Value = 1 Then: intRabies = -1
If chkNeuter.Value = 1 Then: intNeuter = -1

If chkHave.Value = 1 Then: intHave = -1

If (chkHave.Value = -1 And txtReason.Text = "") Then
    intMsgBox = MsgBox("Please explain the reason for not having this pet anymore.", vbOKOnly, "Reason")
    Exit Sub
Else
    strReason = Replace(txtReason.Text, "'", "''")
End If

intType = Get_Types.Get_Types(cboType.Text)

strSQL = "INSERT INTO PET (PET_ADOPTION,"
strSQL = strSQL & " PET_NAME, PET_SEX, PET_SPAY_NEUTER, PET_LOCATION,"
strSQL = strSQL & " PET_AGE, PET_STILL_HAVE, PET_HAVE_REASON, PET_RABIES_VAC, PET_TYPE) VALUES ("
strSQL = strSQL & frmNewAdoption.intAdoptionNum & ", '" & strName & "', '" & strSex
strSQL = strSQL & "', " & intNeuter & ", '" & strKept & "', '" & strAge & "', " & intHave
strSQL = strSQL & ", '" & strReason & "', " & intRabies & ", " & intType & ")"

Open_Recordsets.objConnection.Execute (strSQL)


intMsgBox = MsgBox("Add another animal?", vbYesNo, "Add Another?")
If intMsgBox = 6 Then
    txtName.Text = ""
    cboType.Text = ""
    chkRabies.Value = 0
    chkHave.Value = 0
    chkNeuter.Value = 0
    txtReason.Text = ""
Else
    Unload Me
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
'* Written on: 11-26-2002
'********************************************************************************************
Private Sub Form_Load()

Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL Statement

On Error GoTo ErrorHandler

Set rstType = New ADODB.Recordset
strSQL = "SELECT TYPE_NAME FROM ANIMAL_TYPES"

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
    With rstType
        rstType.MoveFirst
            Do While Not rstType.EOF
                If Not IsNull(![TYPE_NAME]) Then
                    cboType.AddItem ![TYPE_NAME]
                End If
                rstType.MoveNext
            Loop
    End With
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub chkHave_Click()
If chkHave.Value = 1 Then
    txtReason.Enabled = False
End If
If chkHave.Value = 0 Then
    txtReason.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

