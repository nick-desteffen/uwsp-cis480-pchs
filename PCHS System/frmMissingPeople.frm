VERSION 5.00
Begin VB.Form frmMissingPeople 
   Caption         =   "Please select the correct person"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDOB 
      Height          =   1425
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox lstLicense 
      Height          =   1425
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Use None"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdPeople 
      Caption         =   "Use Selected"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ListBox lstPeopleNum 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblName 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmMissingPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNone_Click()
frmNewMissing.intPersonNum = 0
Unload Me
End Sub
Private Sub cmdPeople_Click()
If lstPeopleNum.Text <> "" Then
    frmNewMissing.intPersonNum = lstPeopleNum.Text
    Unload Me
Else
    intMsgBox = MsgBox("Please select a person!", vbOKOnly, "Select please")
    Exit Sub
End If
End Sub
Private Sub Form_Load()

lblName.Caption = frmNewMissing.txtFname.Text & " " & frmNewMissing.txtLname.Text

Dim rstSearch As New ADODB.Recordset
Dim objConnection As New ADODB.Connection

Set rstSearch = New ADODB.Recordset
Set objConnection = New ADODB.Connection

objConnection.ConnectionString = frmPCHS_Main.strConnectionString
objConnection.Open

Set rstSearch = objConnection.Execute("SELECT * FROM PERSON WHERE PERSON_FNAME = '" & frmNewMissing.txtFname.Text & "' AND PERSON_LNAME = '" & frmNewMissing.txtLname.Text & "'")

If rstSearch.EOF = False Then
    With rstSearch
        rstSearch.MoveFirst
        Do While Not rstSearch.EOF
            lstPeopleNum.AddItem ![PERSON_NUMBER]
            If Not IsNull(![PERSON_LICENSE]) Then
                lstLicense.AddItem ![PERSON_LICENSE]
            Else
                lstLicense.AddItem "No License"
            End If
            If Not IsNull(![PERSON_DOB]) Then
                lstDOB.AddItem Format(![PERSON_DOB], "MM/DD/YYYY")
            Else
                lstDOB.AddItem "No Birth Date"
            End If
            rstSearch.MoveNext
        Loop
    End With
End If

Set objConnection = Nothing
Set rstSearch = Nothing

End Sub
