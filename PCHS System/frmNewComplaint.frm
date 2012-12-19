VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNewComplaint 
   Caption         =   "New Complaint"
   ClientHeight    =   6210
   ClientLeft      =   1305
   ClientTop       =   1725
   ClientWidth     =   9600
   Icon            =   "frmNewComplaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9600
   Begin TabDlg.SSTab sstNewComplaint 
      Height          =   4215
      Left            =   240
      TabIndex        =   35
      Top             =   1200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Complainant"
      TabPicture(0)   =   "frmNewComplaint.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDOB"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEmail"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOwnerPhone"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblOwnerState"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblOwnerZip"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblOwnerCity"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblOwnerLname"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblOwnerAddress"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblOwnerFname"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLicense(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpDOB"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSearchName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPhone"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtState"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtZip"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCity"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLname"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAddress"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFname"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtEmail"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtLicense"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Owner "
      TabPicture(1)   =   "frmNewComplaint.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblComplainantMunicipality"
      Tab(1).Control(1)=   "lblPhone"
      Tab(1).Control(2)=   "lblState"
      Tab(1).Control(3)=   "lblZip"
      Tab(1).Control(4)=   "lblCIty"
      Tab(1).Control(5)=   "lblAddress"
      Tab(1).Control(6)=   "lblFname"
      Tab(1).Control(7)=   "lblLname"
      Tab(1).Control(8)=   "txtOwnerMunicipality"
      Tab(1).Control(9)=   "txtOwnerPhone"
      Tab(1).Control(10)=   "txtOwnerState"
      Tab(1).Control(11)=   "txtOwnerZip"
      Tab(1).Control(12)=   "txtOwnerCity"
      Tab(1).Control(13)=   "txtOwnerAddress"
      Tab(1).Control(14)=   "txtOwnerFname"
      Tab(1).Control(15)=   "txtOwnerLname"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Violation "
      TabPicture(2)   =   "frmNewComplaint.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmSex"
      Tab(2).Control(1)=   "dtpViolationDate"
      Tab(2).Control(2)=   "txtQuantity"
      Tab(2).Control(3)=   "chkHowling"
      Tab(2).Control(4)=   "chkRunAtLarge"
      Tab(2).Control(5)=   "chkDefecate"
      Tab(2).Control(6)=   "txtViolation"
      Tab(2).Control(7)=   "txtComments"
      Tab(2).Control(8)=   "cboBreed"
      Tab(2).Control(9)=   "cboAnimalType"
      Tab(2).Control(10)=   "cboColor"
      Tab(2).Control(11)=   "dtpViolationTime"
      Tab(2).Control(12)=   "lblQuantity"
      Tab(2).Control(13)=   "lblViolationDate"
      Tab(2).Control(14)=   "lblViolationTimes"
      Tab(2).Control(15)=   "lblViolation"
      Tab(2).Control(16)=   "lblComments"
      Tab(2).Control(17)=   "lblAnimalType"
      Tab(2).Control(18)=   "lblBreed"
      Tab(2).Control(19)=   "lblColor"
      Tab(2).ControlCount=   20
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2400
         Width           =   3255
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
         Left            =   4560
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3720
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
      Begin VB.TextBox txtOwnerLname 
         Height          =   285
         Left            =   -73560
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearchName 
         Caption         =   "Search People"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtOwnerFname 
         Height          =   285
         Left            =   -73560
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtOwnerAddress 
         Height          =   285
         Left            =   -73560
         TabIndex        =   13
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtOwnerCity 
         Height          =   285
         Left            =   -73560
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtOwnerZip 
         Height          =   285
         Left            =   -70560
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtOwnerState 
         Height          =   285
         Left            =   -71400
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtOwnerPhone 
         Height          =   285
         Left            =   -73560
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtOwnerMunicipality 
         Height          =   285
         Left            =   -70800
         TabIndex        =   18
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Frame frmSex 
         Caption         =   "Sex"
         Height          =   615
         Left            =   -72120
         TabIndex        =   36
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpViolationDate 
         Height          =   375
         Left            =   -69360
         TabIndex        =   27
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47382529
         CurrentDate     =   37581
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   -70200
         TabIndex        =   24
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkHowling 
         Caption         =   "Howling"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -69120
         TabIndex        =   31
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chkRunAtLarge 
         Caption         =   "Runs at large"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69120
         TabIndex        =   30
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkDefecate 
         Caption         =   "Defecate"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69120
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtViolation 
         Height          =   405
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2520
         Width           =   4935
      End
      Begin VB.TextBox txtComments 
         Height          =   615
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   3240
         Width           =   4935
      End
      Begin VB.ComboBox cboBreed 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox cboAnimalType 
         Height          =   315
         Left            =   -74760
         TabIndex        =   19
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         Left            =   -74760
         TabIndex        =   21
         Top             =   1920
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpViolationTime 
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
         Left            =   -67680
         TabIndex        =   28
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47382530
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47382529
         CurrentDate     =   37579
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblOwnerFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   600
         TabIndex        =   65
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblOwnerAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   64
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblOwnerLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   600
         TabIndex        =   63
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblOwnerCity 
         Caption         =   "City"
         Height          =   255
         Left            =   960
         TabIndex        =   62
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblOwnerZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4200
         TabIndex        =   61
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblOwnerState 
         Caption         =   "State"
         Height          =   255
         Left            =   3120
         TabIndex        =   60
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblOwnerPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3120
         TabIndex        =   57
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   -74400
         TabIndex        =   56
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   -74400
         TabIndex        =   55
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   -74280
         TabIndex        =   54
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblCIty 
         Caption         =   "City"
         Height          =   255
         Left            =   -74040
         TabIndex        =   53
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   -70920
         TabIndex        =   52
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         Height          =   255
         Left            =   -71880
         TabIndex        =   51
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   -74760
         TabIndex        =   50
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblComplainantMunicipality 
         Caption         =   "Municipality"
         Height          =   255
         Left            =   -71760
         TabIndex        =   49
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Number of Animals Involed:"
         Height          =   255
         Left            =   -72240
         TabIndex        =   48
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblViolationDate 
         Caption         =   "Date of violation"
         Height          =   255
         Left            =   -69360
         TabIndex        =   47
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblViolationTimes 
         Caption         =   "Time"
         Height          =   255
         Left            =   -67680
         TabIndex        =   46
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblViolation 
         Caption         =   "Violation"
         Height          =   255
         Left            =   -74760
         TabIndex        =   45
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblComments 
         Caption         =   "Comments from Complainant"
         Height          =   255
         Left            =   -74760
         TabIndex        =   44
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lblAnimalType 
         Caption         =   "Type of Animal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblBreed 
         Caption         =   "Breed of Animal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblColor 
         Caption         =   "Color of Animal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.TextBox txtReferredTo 
      Height          =   285
      Left            =   2040
      TabIndex        =   33
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtTakenBy 
      Height          =   285
      Left            =   3240
      TabIndex        =   32
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   8040
      TabIndex        =   34
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   37
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lblReferredTo 
      Caption         =   "Complaint Referred To"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblTakenby 
      Caption         =   "Taken By"
      Height          =   255
      Left            =   2400
      TabIndex        =   39
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblComplaintDate 
      Caption         =   "Date "
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   240
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu mnuBack 
         Caption         =   "Back"
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Index           =   3
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
Attribute VB_Name = "frmNewComplaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************
'* This form takes down all the information required when sombody calls and complains about
'* an animal.  It is recorded to the complaint table.
'*
'* Written by: Kalonji Kadima
'* Written on: 11-25-2002
'********************************************************************************************

'********************************************************************************************
'* Array that contains all the types of animals
'********************************************************************************************
Private Type combo_info
    name As String
    Number As Integer
End Type

Public bolMatchFound As Boolean         'True = person already in database
Dim intType As Integer                  'Type of animal
Public intPersonNum As Integer          'Number of the person if found

'********************************************************************************************
'* Called when the value in the animal type combo box is changed
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cboAnimalType_Click()

Dim breeds() As combo_info          'Array containing all the breeds
Dim looper As Integer               'Loop control variable
        
Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'used for messageboxes
Dim strSQL As String                'SQL statement

On Error GoTo ErrorHandler

Set rstType = New ADODB.Recordset

chkHowling.Enabled = False
chkRunAtLarge.Enabled = False
chkDefecate.Enabled = False

'Populates dog breed recordset
If cboAnimalType.Text = "Dog" Then
    
    cboBreed.Enabled = True
    chkHowling.Enabled = True
    chkRunAtLarge.Enabled = True
    chkDefecate.Enabled = True
    
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
    chkRunAtLarge.Enabled = True
    chkDefecate.Enabled = True

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

Private Sub cmdSave_Click()

'Complainant variables

Dim strFname As String              'First name of the complainant
Dim strLname As String              'Last name of the complainant
Dim strAddress As String            'Address of the complainant
Dim strCity As String               'City of the complainant
Dim strState As String              'State of the complainant
Dim strZip As String                'Zipcode of the complainant
Dim strPhone As String              'Telephone number of the complainant
Dim strEmail As String              'Email address of complainant
Dim strLicense As String            'Drivers license of the complainant
Dim dteDOB As Date                  'Date of birth of complainant

'Original owner variables

Dim strOwnerFname As String         'First name of the owner
Dim strOwnerLname As String         'Last name of the owner
Dim strOwnerAddress As String       'Address of the owner
Dim strOwnerCity As String          'City of owner
Dim strOwnerState As String         'State of owner
Dim strOwnerZip As String           'Zipcode of owner
Dim strOwnerPhone As String         'Phone number of owner
Dim strOwnerMunicipality As String  'Location

'General variables

Dim strSex As String                'Sex of the owner
Dim strTakenBy As String            'Person who took down the complaint
Dim strReferredTo As String         'Person the complaint was referred to
Dim intType As Integer              'Type of animal
Dim intBreed As Integer             'Breed of animal
Dim intColor As Integer             'Color of animal
Dim intQuantity As Integer          'Number of animals involved
Dim intHowling As Integer           'Whether the animal has a howling problem
Dim intRunAtLarge As Integer        'Whether the animal runs at large
Dim intDefecate As Integer          'Whether the animal defecates
Dim dteViolationDate As Date        'Date of the violation
Dim dteViolationTime As String      'Time of the violation
Dim strViolation As String          'Description of the violation
Dim strComments As String           'Comments from the complainant

Dim rstInsert As ADODB.Recordset    'Used to interface with the database
Dim strSQL As String                'SQL Statement
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'First tab

If txtFname.Text = "" Or txtLname.Text = "" Or txtAddress.Text = "" Or txtCity.Text = "" Or txtState.Text = "" Or txtZip.Text = "" Or txtPhone.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the personal information fields!", vbOKOnly, "Error!")
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

'Second tab

If txtTakenBy.Text = "" Then
    intMsgBox = MsgBox("Please enter your name in the complaint taken by field!", vbOKOnly, "Error!")
    Exit Sub
End If

If txtOwnerFname.Text = "" Or txtOwnerLname.Text = "" Or txtOwnerAddress.Text = "" Or txtOwnerCity.Text = "" Or txtOwnerState.Text = "" Or txtOwnerZip.Text = "" Or txtOwnerPhone.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the animal owner fields!", vbOKOnly, "Error!")
    Exit Sub
End If
  
strTakenBy = Replace(txtTakenBy.Text, "'", "''")
strReferredTo = Replace(txtReferredTo.Text, "'", "''")
strOwnerFname = Replace(txtOwnerFname.Text, "'", "''")
strOwnerLname = Replace(txtOwnerLname.Text, "'", "''")
strOwnerAddress = Replace(txtOwnerAddress.Text, "'", "''")
strOwnerCity = Replace(txtOwnerCity.Text, "'", "''")
strOwnerState = Replace(txtOwnerState.Text, "'", "''")
strOwnerMunicipality = Replace(txtOwnerMunicipality.Text, "'", "''")


If Verify_Data.Check_Zip(txtOwnerZip.Text) = True Then
    strOwnerZip = txtOwnerZip.Text
Else
    intMsgBox = MsgBox("Please enter a valid zip code!" & Chr(13) & "Valid formats are ##### or #####-####.", vbOKOnly, "Invalid Zip Code")
    Exit Sub
End If

If Verify_Data.Check_Phone(txtOwnerPhone.Text) = True Then
    strOwnerPhone = txtOwnerPhone.Text
Else
    intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

'Third tab

'Chooses the sex of the animal
If optMale.Value = True Then
    strSex = "M"
ElseIf optFemale.Value = True Then
    strSex = "F"
Else
    intMsgBox = MsgBox("Please select the sex of the animal!", vbOKOnly, "Error")
    Exit Sub
End If

'Checks to see if all required fields are populated
If cboColor.Text = "" Then
    intMsgBox = MsgBox("Please choose the color of the animal!", vbOKOnly, "Error")
    Exit Sub
End If

If cboAnimalType.Text = "" Then
    intMsgBox = MsgBox("Please choose the type of animal!", vbOKOnly, "Error")
    Exit Sub
End If

If ((cboAnimalType.Text = "Dog") Or (cboAnimalType.Text = "Cat")) And (cboBreed.Text = "") Then
        intMsgBox = MsgBox("Please choose the breed of the animal!", vbOKOnly, "Error")
        Exit Sub
End If

'Returns the type, breed, and color number of the animal

intType = Get_Types.Get_Types(cboAnimalType.Text)
intColor = Get_Colors.Get_Colors(cboColor.Text)
If intType = 1 Or intType = 2 Then
    intBreed = Get_Breeds.Get_Breeds(intType, cboBreed.Text)
Else
    intBreed = 0
End If

If IsNumeric(txtQuantity.Text) = True Then
    intQuantity = txtQuantity.Text
Else
    intMsgBox = MsgBox("Please enter a numeric value for the number of animals involved.", vbOKOnly, "Error!")
    Exit Sub
End If

If chkHowling.Value = 1 Then: intHowling = -1
If chkRunAtLarge.Value = 1 Then: intRunAtLarge = -1
If chkDefecate.Value = 1 Then: intDefecate = -1

dteViolationDate = dtpViolationDate.Value
dteViolationTime = dtpViolationTime.Value
strViolation = Replace(txtViolation.Text, "'", "''")
strComments = Replace(txtComments.Text, "'", "''")

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

intMsgBox = MsgBox("Add new complaint information to database?", vbYesNo, "Add?")

'Inserts new complaint into database

strSQL = "INSERT INTO COMPLAINT (COMPLAINT_PERSON, "
strSQL = strSQL & "COMPLAINT_OWNER_FNAME, "
strSQL = strSQL & "COMPLAINT_OWNER_LNAME, "
strSQL = strSQL & "COMPLAINT_ADDRESS, "
strSQL = strSQL & "COMPLAINT_CITY, "
strSQL = strSQL & "COMPLAINT_STATE, "
strSQL = strSQL & "COMPLAINT_ZIP, "
strSQL = strSQL & "COMPLAINT_PHONE, "
strSQL = strSQL & "COMPLAINT_MUNICIPALITY, "
strSQL = strSQL & "COMPLAINT_SEX, "
strSQL = strSQL & "COMPLAINT_TYPE, "
strSQL = strSQL & "COMPLAINT_BREED, "
strSQL = strSQL & "COMPLAINT_COLOR, "
strSQL = strSQL & "COMPLAINT_QUANTITY, "
strSQL = strSQL & "COMPLAINT_HOWLING, "
strSQL = strSQL & "COMPLAINT_AT_LARGE, "
strSQL = strSQL & "COMPLAINT_DEFECATE, "
strSQL = strSQL & "COMPLAINT_VIOLATION_DATE, "
strSQL = strSQL & "COMPLAINT_VIOLATION_TIME, "
strSQL = strSQL & "COMPLAINT_TAKEN_BY, "
strSQL = strSQL & "COMPLAINT_PERSON_REFERRED, "
strSQL = strSQL & "COMPLAINT_VIOLATION_INFO, "
strSQL = strSQL & "COMPLAINT_COMMENTS) "
strSQL = strSQL & "VALUES ("
strSQL = strSQL & intPersonNum & ", '" & strOwnerFname & "', '" & strOwnerLname & "', '" & strOwnerAddress
strSQL = strSQL & "','" & strOwnerCity & "', '" & strOwnerState & "', '" & strOwnerZip
strSQL = strSQL & "', '" & strOwnerPhone & "', '" & strOwnerMunicipality & "', '" & strSex
strSQL = strSQL & "', " & intType & ", " & intBreed & ", " & intColor & ", " & intQuantity
strSQL = strSQL & ", " & intHowling & ", " & intRunAtLarge & ", " & intDefecate
strSQL = strSQL & ", '" & dteViolationDate & "', '" & dteViolationTime
strSQL = strSQL & "', '" & strTakenBy & "', '" & strReferredTo
strSQL = strSQL & "', '" & strViolation & "', '" & strComments & "')"

Open_Recordsets.objConnection.Execute (strSQL)

frmPCHS_Main.Show
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

Dim strFname As String              'First name of the complainant
Dim strLname As String              'Last name of the complainant
Dim strAddress As String            'Address of the complainant
Dim strCity As String               'City of the complainant
Dim strState As String              'State of the complainant
Dim strZip As String                'Zipcode of the complainant
Dim strPhone As String              'Telephone number of the complainant
Dim strEmail As String              'Email address of complainant
Dim strLicense As String            'Drivers license of the complainant
Dim dteDOB As Date                  'Date of birth of complainant

Dim rstSearch As New ADODB.Recordset 'Recordset used for interfacing with the database
Dim intMsgBox As Integer             'Used for messageboxes
Dim strSQL As String                 'SQL statement

Set rstSearch = New ADODB.Recordset

On Error GoTo ErrorHandler

strSQL = "SELECT * FROM PERSON WHERE PERSON_LNAME = '" & txtLname.Text & "'"

Set rstSearch = Open_Recordsets.objConnection.Execute(strSQL)

If rstSearch.EOF <> True Then
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 5
    
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
Private Sub Form_Load()

Dim types() As combo_info           'Array containing types of animals
Dim colors() As combo_info          'Array containing colors of animals

Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes
Dim strSQL As String                'SQL Statement
Dim looper As Integer               'Loop control variable

On Error GoTo ErrorHandler

lblComplaintDate.Caption = "Date Received: " & Date
dtpDOB.Value = Date
dtpViolationDate.Value = Date

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

Private Sub mnuBack_Click(Index As Integer)
Unload Me
frmPCHS_Main.Show
End Sub

Private Sub mnuExit_Click(Index As Integer)
Open_Recordsets.objConnection.Close
End
End Sub

Private Sub mnuAbout_Click()
Call About.About
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmPCHS_Main.Show
End Sub
