VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFollowUp 
   Caption         =   "Follow Up Information"
   ClientHeight    =   7665
   ClientLeft      =   1890
   ClientTop       =   630
   ClientWidth     =   9180
   Icon            =   "frmFollowUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   9180
   Begin VB.CommandButton cmdReceipt 
      Caption         =   "Generate Receipt"
      Height          =   495
      Left            =   3840
      TabIndex        =   44
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify Information"
      Height          =   495
      Left            =   6480
      TabIndex        =   43
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame frmAdoption 
      Caption         =   "Adoption Information"
      Height          =   2175
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   8895
      Begin MSComCtl2.DTPicker dtpTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   42
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47972354
         CurrentDate     =   37579
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   840
         TabIndex        =   41
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47972353
         CurrentDate     =   37579
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmFollowUp.frx":08CA
         Left            =   6720
         List            =   "frmFollowUp.frx":08DD
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtFollowUp 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtAgent 
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtNeuterSponsor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status of adoption"
         Height          =   255
         Left            =   6840
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblFollowUp 
         Caption         =   "Follow up instructions"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblTime 
         Caption         =   "Time"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblDate 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Agent"
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblPickUpDate 
         Caption         =   "Pick up:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblNeuterSponsor 
         Caption         =   "Neutor donation amount:"
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame frmAnimal 
      Caption         =   "Animal Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   5760
      Width           =   8895
      Begin VB.Label lblAnimalType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblAnimalNum 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdContract 
      Caption         =   "Generate Contract"
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Update"
      Height          =   495
      Left            =   7800
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame frmPersonal 
      Caption         =   "Adopter Information"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   8895
      Begin VB.TextBox txtLicense 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtFname 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtLname 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtCity 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtZip 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPhone 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtDOB 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   1
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCIty 
         Caption         =   "City"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.Label lblCurDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddd, MMMM dd, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblAdoptionNum 
      Caption         =   "Adoption Number:"
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
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   4815
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
End
Attribute VB_Name = "frmFollowUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************
'* This form is used after sombody fills out an adoption form.  From it the adoption
'* status can be adjusted and several forms can be displayed.  The adoption verification
'* form, adoption contract, spay/neuter verification form, reciept, surrendered animal
'* information, and spay/neuter donation voucher can all be displayed from this form.
'*
'* Written by: Nick DeSteffen
'* Written on: 11-15-2002
'**************************************************************************************

Public intPersonNum As Integer          'Number of the person
Public intAdoptionNum As Integer        'Number of the adoption - Required
Public intAnimalNum As Integer          'Number of the animal
Public strAnimalAcquire As String       'How the animal was acquired
Public intDonationAmount As Integer     'Amount sombody donated for animal

'Shows the contract, examination form, surrendered information form, and spay/neuter voucher form
Private Sub cmdContract_Click()
frmShowContract.intAdoptionNum = intAdoptionNum
frmShowExamine.intAdoptionNum = intAdoptionNum

frmShowContract.Show
frmShowExamine.Show

If intDonationAmount > 0 Then
    frmShowNeuter.intAdoptionNum = intAdoptionNum
    frmShowNeuter.Show
End If

If strAnimalAcquire = "S" Then
    frmShowSurrender.intAnimalNum = intAnimalNum
    frmShowSurrender.Show
End If

End Sub

'Calls the receipt function to create a receipt for the transaction
Private Sub cmdReceipt_Click()
frmNewReciept.intPersonNum = intPersonNum
    frmNewReciept.cboReason.ListIndex = 0
    frmNewReciept.intType = 1
    frmNewReciept.intNumber = intAdoptionNum
    frmNewReciept.Show
End Sub

'********************************************************************************************
'* Runs when the user clicks save, saves the follow up adoption information to the database
'*
'********************************************************************************************
Private Sub cmdSave_Click()

Dim strStatus As String             'Status of the adoption
Dim strFollowUp As String           'Follow up directions for the adoption
Dim dteDate As Date                 'Date of the adoption
Dim strAgent As String              'Name of agent
Dim intNeuterSponsor As Double      'Amount of money donated by sponsor
Dim dteTime As Date                 'Time of adoption

Dim intMsgBox As Integer            'Used for message boxes
Dim strSQL As String                'SQL statement

On Error GoTo ErrorHandler

'Checks to see if all required fields are populated
If cboStatus.Text = "" Then
    intMsgBox = MsgBox("Please adjust the adoption status!", vbOKOnly, "Error")
    Exit Sub
End If

strStatus = Left(cboStatus.Text, 1)
strFollowUp = Replace(txtFollowUp.Text, "'", "''")
strAgent = Replace(txtAgent.Text, "'", "''")

If strStatus = "A" Then
    dteDate = dtpDate.Value
    dteTime = dtpTime.Value
End If
    
'Updates the database with the selected values

strSQL = "UPDATE ADOPTION SET ADOPTION_STATUS = '" & strStatus
strSQL = strSQL & "', ADOPTION_FOLLOW_UP = '" & strFollowUp
strSQL = strSQL & "', ADOPTION_AGENT = '" & strAgent
strSQL = strSQL & "', ADOPTION_DATE = '" & dteDate
strSQL = strSQL & "', ADOPTION_TIME = '" & dteTime
strSQL = strSQL & "' WHERE ADOPTION_NUMBER = " & intAdoptionNum

Open_Recordsets.objConnection.Execute (strSQL)

If strStatus = "H" Or strStatus = "C" Then
    strSQL = "UPDATE ANIMALS SET ANIMAL_STATUS = 'A', ANIMAL_DATE_STATUS = '" & Date & "' WHERE ANIMAL_NUMBER = " & intAnimalNum
    Open_Recordsets.objConnection.Execute (strSQL)
End If

If strStatus = "D" Then
    strSQL = "UPDATE ANIMALS SET ANIMAL_STATUS = 'R', ANIMAL_DATE_STATUS = Null WHERE ANIMAL_NUMBER = " & intAnimalNum
    Open_Recordsets.objConnection.Execute (strSQL)
End If

frmActiveAdoptions.Show
Unload Me

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

'********************************************************************************************
'* Called when the form loads, populates all the information fields with appropriate data.
'********************************************************************************************
Private Sub Form_Load()

Dim strStatus As String                     'Status of the adoption

Dim rstType As ADODB.Recordset              'Recordset used to get information from the database
Dim strSQL As String                        'SQL statement
Dim intMsgBox As Integer                    'Used for messageboxes

On Error GoTo ErrorHandler

Set rstType = New ADODB.Recordset
dtpDate.Value = Date

intAdoptionNum = frmActiveAdoptions.dgdActiveAdoptions.Columns("Adoption Num").CellValue(frmActiveAdoptions.dgdActiveAdoptions.Bookmark)
lblAdoptionNum.Caption = "Adoption Number: " & intAdoptionNum

'Populates all the fields with information from the selected adoption

strSQL = "SELECT Person.person_lname, "
strSQL = strSQL & "Person.person_fname, "
strSQL = strSQL & "Person.person_address, "
strSQL = strSQL & "Person.person_number, "
strSQL = strSQL & "Person.person_city, "
strSQL = strSQL & "Person.person_state, "
strSQL = strSQL & "Person.person_zip, "
strSQL = strSQL & "Person.person_telephone, "
strSQL = strSQL & "Person.person_dob, "
strSQL = strSQL & "Person.person_email, "
strSQL = strSQL & "Person.person_license, "
strSQL = strSQL & "Animals.animal_number, "
strSQL = strSQL & "Animal_Types.type_name, "
strSQL = strSQL & "Adoption.adoption_date, "
strSQL = strSQL & "Adoption.adoption_time, "
strSQL = strSQL & "Adoption.adoption_agent, "
strSQL = strSQL & "Adoption.adoption_status, "
strSQL = strSQL & "Adoption.adoption_follow_up, "
strSQL = strSQL & "Adoption.adoption_number "
strSQL = strSQL & "From Animal_Types "
strSQL = strSQL & "INNER JOIN (Person INNER JOIN "
strSQL = strSQL & "(Animals INNER JOIN Adoption "
strSQL = strSQL & "ON Animals.animal_number = Adoption.adoption_animal) "
strSQL = strSQL & "ON Person.person_number = Adoption.adoption_adoptorNum) "
strSQL = strSQL & "ON Animal_Types.type_number = Animals.animal_type "
strSQL = strSQL & "WHERE Adoption.adoption_number = " & intAdoptionNum

Set rstType = Open_Recordsets.objConnection.Execute(strSQL)

If rstType.EOF = False Then
With rstType
    rstType.MoveFirst
    Do While Not rstType.EOF

        If Not IsNull(![adoption_status]) Then
            If ![adoption_status] = "A" Then
                cboStatus.ListIndex = 0
            ElseIf ![adoption_status] = "P" Then
                cboStatus.ListIndex = 1
            ElseIf ![adoption_status] = "D" Then
                cboStatus.ListIndex = 2
            ElseIf ![adoption_status] = "C" Then
                cboStatus.ListIndex = 4
            ElseIf ![adoption_status] = "H" Then
                cboStatus.ListIndex = 3
            Else
                cboStatus.ListIndex = 1
            End If
        End If
            
        If Not IsNull(![PERSON_FNAME]) Then
            txtFname.Text = ![PERSON_FNAME]
        Else
            txtFname.Text = "None"
        End If
        
        If Not IsNull(![PERSON_LNAME]) Then
            txtLname.Text = ![PERSON_LNAME]
        Else
            txtLname.Text = "None"
        End If
        
        If Not IsNull(![person_address]) Then
            txtAddress.Text = ![person_address]
        Else
            txtAddress.Text = "None"
        End If
        
        If Not IsNull(![person_city]) Then
            txtCity.Text = ![person_city]
        Else
            txtCity.Text = "None"
        End If
        
         If Not IsNull(![person_state]) Then
            txtState.Text = ![person_state]
        Else
            txtState.Text = "None"
        End If
        
         If Not IsNull(![person_zip]) Then
            txtZip.Text = ![person_zip]
        Else
            txtZip.Text = "None"
        End If
        
         If Not IsNull(![person_telephone]) Then
            txtPhone.Text = ![person_telephone]
        Else
            txtPhone.Text = "None"
        End If
        
         If Not IsNull(![person_dob]) Then
            txtDOB.Text = ![person_dob]
        Else
            txtDOB.Text = "Invalid"
        End If
        
         If Not IsNull(![person_email]) Then
            txtEmail.Text = ![person_email]
        Else
            txtEmail.Text = "None"
        End If
        
         If Not IsNull(![person_license]) Then
            txtLicense.Text = ![person_license]
        Else
            txtLicense.Text = "None"
        End If
        
        If Not IsNull(![adoption_follow_up]) Then
            txtFollowUp.Text = ![adoption_follow_up]
        Else
            txtFollowUp.Text = ""
        End If
        
        If Not IsNull(![adoption_date]) Then
            dtpDate.Value = ![adoption_date]
        End If
        
        If Not IsNull(![adoption_time]) Then
            dtpTime.Value = ![adoption_time]
        End If
        
        If Not IsNull(![adoption_agent]) Then
            txtAgent.Text = ![adoption_agent]
        Else
            txtAgent.Text = ""
        End If
        
        If Not IsNull(![animal_number]) Then
            lblAnimalNum.Caption = "Animal Number: " & ![animal_number]
            intAnimalNum = ![animal_number]
        Else
            lblAnimalNum.Caption = "Animal Number: "
        End If
        
        If Not IsNull(![TYPE_NAME]) Then
            lblAnimalType.Caption = "Animal Type: " & ![TYPE_NAME]
        Else
            lblAnimalType.Caption = ""
        End If
        
        If Not IsNull(![PERSON_NUMBER]) Then
            intPersonNum = ![PERSON_NUMBER]
        Else
            intPersonNum = 0
        End If
            
    rstType.MoveNext
    Loop
End With
End If

'Checks to see if the animal was surrendered to display the information at adoption time

Set rstType = Open_Recordsets.objConnection.Execute("SELECT ANIMAL_ACQUIRED FROM ANIMALS WHERE ANIMAL_NUMBER = " & intAnimalNum)

If rstType.EOF <> True Then
    With rstType
        rstType.MoveFirst
        Do While rstType.EOF = False
            strAnimalAcquire = ![animal_acquired]
        rstType.MoveNext
        Loop
    End With
End If
    
'Checks to see if sombody made a donation to have this animal spayed or neutered and displys the amount

Set rstType = Open_Recordsets.objConnection.Execute("SELECT DONATION.DONATION_AMOUNT FROM DONATION, ANIMALS WHERE ANIMALS.ANIMAL_NEUTER_SPONSOR = DONATION.DONATION_NUMBER AND ANIMALS.ANIMAL_NUMBER = " & intAnimalNum)
If rstType.EOF <> True Then
    With rstType
        rstType.MoveFirst
        Do While rstType.EOF = False
            txtNeuterSponsor.Text = "$" & ![DONATION_AMOUNT]
            intDonationAmount = ![DONATION_AMOUNT]
        rstType.MoveNext
        Loop
    End With
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

Private Sub cmdCancel_Click()
frmActiveAdoptions.Show
Unload Me
End Sub

Private Sub cmdVerify_Click()
frmVerifyAdoption.intAdoptionNum = intAdoptionNum
frmVerifyAdoption.Show
End Sub

Private Sub mnuBack_Click()
frmActiveAdoptions.Show
Unload Me
End Sub

Private Sub mnuExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub

