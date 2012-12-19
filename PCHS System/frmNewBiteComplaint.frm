VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNewBiteComplaint 
   Caption         =   "New Bite Report"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8280
   Icon            =   "frmNewBiteComplaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   57
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6840
      TabIndex        =   58
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtTakenBy 
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin TabDlg.SSTab sstNewBiteComplaint 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Victim "
      TabPicture(0)   =   "frmNewBiteComplaint.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHowBiteHappen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblParentsName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLocationOfBite"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTime"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDateofBite"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblOwnerHomePhone"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblOwnerState"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblOwnerZip"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblOwnerCity"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblOwnerLname"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblOwnerAddress"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblOwnerFname"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblEmail"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDOB"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblLicense"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dtpDOB"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dtpBiteTime"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dtpBiteDate"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtHow"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtParents"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtBiteLocation"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdSearchName"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtPhone"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtState"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtZip"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtCity"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtLname"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtAddress"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtFname"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtEmail"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtLicense"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Owner "
      TabPicture(1)   =   "frmNewBiteComplaint.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblVictimWorkPhone"
      Tab(1).Control(1)=   "lblVictimMunicipality"
      Tab(1).Control(2)=   "lblFname"
      Tab(1).Control(3)=   "lblVictimAddress"
      Tab(1).Control(4)=   "lblVictimCity"
      Tab(1).Control(5)=   "lblVictimZip"
      Tab(1).Control(6)=   "lblVictimState"
      Tab(1).Control(7)=   "lblVictimHomePhone"
      Tab(1).Control(8)=   "lblLname"
      Tab(1).Control(9)=   "lblOwnedStray"
      Tab(1).Control(10)=   "txtOwnerWorkPhone"
      Tab(1).Control(11)=   "txtMunicipality"
      Tab(1).Control(12)=   "txtOwnerPhone"
      Tab(1).Control(13)=   "txtOwnerAddress"
      Tab(1).Control(14)=   "txtOwnerCity"
      Tab(1).Control(15)=   "txtOwnerZip"
      Tab(1).Control(16)=   "txtOwnerState"
      Tab(1).Control(17)=   "txtOwnerFname"
      Tab(1).Control(18)=   "txtOwnerLname"
      Tab(1).Control(19)=   "Frame1"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Animal"
      TabPicture(2)   =   "frmNewBiteComplaint.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtVetPhone"
      Tab(2).Control(1)=   "txtVetClinic"
      Tab(2).Control(2)=   "cboAge"
      Tab(2).Control(3)=   "txtMarkings"
      Tab(2).Control(4)=   "cboAnimalType"
      Tab(2).Control(5)=   "cboBreed"
      Tab(2).Control(6)=   "cboColor"
      Tab(2).Control(7)=   "txtAnimalName"
      Tab(2).Control(8)=   "frmSex"
      Tab(2).Control(9)=   "chkNeuter"
      Tab(2).Control(10)=   "dtpVaccinateDate"
      Tab(2).Control(11)=   "lblVaccinateDate"
      Tab(2).Control(12)=   "lblVetPhone"
      Tab(2).Control(13)=   "lblVetClinic"
      Tab(2).Control(14)=   "lblAge"
      Tab(2).Control(15)=   "lblMarkings"
      Tab(2).Control(16)=   "lblType"
      Tab(2).Control(17)=   "lblBreed"
      Tab(2).Control(18)=   "lblColor"
      Tab(2).Control(19)=   "lblName"
      Tab(2).ControlCount=   20
      TabCaption(3)   =   "Medical"
      TabPicture(3)   =   "frmNewBiteComplaint.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtPhysicianState"
      Tab(3).Control(1)=   "txtPhysicianZip"
      Tab(3).Control(2)=   "txtPhysicianCity"
      Tab(3).Control(3)=   "txtPhysicianAddress"
      Tab(3).Control(4)=   "txtPhysicianPhone"
      Tab(3).Control(5)=   "txtPhysician"
      Tab(3).Control(6)=   "txtClinicDrHospital"
      Tab(3).Control(7)=   "chkMedical"
      Tab(3).Control(8)=   "Label3"
      Tab(3).Control(9)=   "Label2"
      Tab(3).Control(10)=   "Label1"
      Tab(3).Control(11)=   "PhysicianAddress"
      Tab(3).Control(12)=   "lblPhysicianPhone"
      Tab(3).Control(13)=   "lblFamilyPhysician"
      Tab(3).Control(14)=   "lblClinicDrHospital"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Quarantine"
      TabPicture(4)   =   "frmNewBiteComplaint.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cboWhereQuarantine"
      Tab(4).Control(1)=   "txtHumaneOfficer"
      Tab(4).Control(2)=   "txtExamVetPhone"
      Tab(4).Control(3)=   "txtExamVet"
      Tab(4).Control(4)=   "dtpLastQuarantine"
      Tab(4).Control(5)=   "dtpFirstQuarantine"
      Tab(4).Control(6)=   "lblInformed"
      Tab(4).Control(7)=   "lblHumaneOfficer"
      Tab(4).Control(8)=   "lblExamVetPhone"
      Tab(4).Control(9)=   "lblExamVet"
      Tab(4).Control(10)=   "lblWhereQuarantined"
      Tab(4).Control(11)=   "lblLastQuarantine"
      Tab(4).Control(12)=   "lblFirstQuarantine"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "Other"
      TabPicture(5)   =   "frmNewBiteComplaint.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtEuthanizeBy"
      Tab(5).Control(1)=   "txtPreparedBy"
      Tab(5).Control(2)=   "txtComments"
      Tab(5).Control(3)=   "dtpLabDate"
      Tab(5).Control(4)=   "dtpEuthanizeDate"
      Tab(5).Control(5)=   "lblDateEuthanized"
      Tab(5).Control(6)=   "lblEuthanizedBy"
      Tab(5).Control(7)=   "lblSpecimen"
      Tab(5).Control(8)=   "lblLabDate"
      Tab(5).Control(9)=   "lblComments"
      Tab(5).ControlCount=   10
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   -73680
         TabIndex        =   115
         Top             =   600
         Width           =   1935
         Begin VB.OptionButton optOwned 
            Caption         =   "Owned"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optStray 
            Caption         =   "Stray"
            Height          =   255
            Left            =   1080
            TabIndex        =   116
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtOwnerLname 
         Height          =   285
         Left            =   -73800
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cboWhereQuarantine 
         Height          =   315
         ItemData        =   "frmNewBiteComplaint.frx":0972
         Left            =   -74760
         List            =   "frmNewBiteComplaint.frx":097F
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtFname 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtLname 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3720
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearchName 
         Caption         =   "Search People"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtOwnerFname 
         Height          =   285
         Left            =   -73800
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtOwnerState 
         Height          =   285
         Left            =   -71520
         TabIndex        =   22
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtOwnerZip 
         Height          =   285
         Left            =   -70560
         TabIndex        =   23
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtOwnerCity 
         Height          =   285
         Left            =   -73800
         TabIndex        =   21
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtOwnerAddress 
         Height          =   285
         Left            =   -73800
         TabIndex        =   20
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtOwnerPhone 
         Height          =   285
         Left            =   -73800
         TabIndex        =   24
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtMunicipality 
         Height          =   285
         Left            =   -73800
         TabIndex        =   26
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox txtOwnerWorkPhone 
         Height          =   285
         Left            =   -73800
         TabIndex        =   25
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtHumaneOfficer 
         Height          =   285
         Left            =   -74760
         TabIndex        =   46
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtExamVetPhone 
         Height          =   285
         Left            =   -71760
         TabIndex        =   51
         Top             =   3660
         Width           =   1935
      End
      Begin VB.TextBox txtExamVet 
         Height          =   285
         Left            =   -74760
         TabIndex        =   50
         Top             =   3660
         Width           =   1935
      End
      Begin VB.TextBox txtPhysicianState 
         Height          =   285
         Left            =   -70920
         TabIndex        =   43
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtPhysicianZip 
         Height          =   285
         Left            =   -70080
         TabIndex        =   44
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtPhysicianCity 
         Height          =   285
         Left            =   -73080
         TabIndex        =   42
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtPhysicianAddress 
         Height          =   285
         Left            =   -73080
         TabIndex        =   41
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtPhysicianPhone 
         Height          =   315
         Left            =   -73080
         TabIndex        =   45
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtPhysician 
         Height          =   285
         Left            =   -73080
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtClinicDrHospital 
         Height          =   285
         Left            =   -72120
         TabIndex        =   39
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chkMedical 
         Caption         =   "Recieved medical treatment"
         Height          =   195
         Left            =   -74760
         TabIndex        =   38
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtEuthanizeBy 
         Height          =   285
         Left            =   -72480
         TabIndex        =   53
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtPreparedBy 
         Height          =   285
         Left            =   -72480
         TabIndex        =   54
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtComments 
         Height          =   735
         Left            =   -73680
         TabIndex        =   56
         Top             =   2400
         Width           =   5415
      End
      Begin VB.TextBox txtVetPhone 
         Height          =   285
         Left            =   -71280
         TabIndex        =   37
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtVetClinic 
         Height          =   285
         Left            =   -74400
         TabIndex        =   36
         Top             =   3480
         Width           =   2895
      End
      Begin VB.ComboBox cboAge 
         Height          =   315
         ItemData        =   "frmNewBiteComplaint.frx":09AE
         Left            =   -74280
         List            =   "frmNewBiteComplaint.frx":09BB
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtMarkings 
         Height          =   285
         Left            =   -70440
         TabIndex        =   34
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cboAnimalType 
         Height          =   315
         Left            =   -74280
         TabIndex        =   28
         Top             =   900
         Width           =   1815
      End
      Begin VB.ComboBox cboBreed 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74280
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   1260
         Width           =   1815
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         Left            =   -74280
         TabIndex        =   30
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtAnimalName 
         Height          =   315
         Left            =   -74280
         TabIndex        =   27
         Top             =   540
         Width           =   1575
      End
      Begin VB.Frame frmSex 
         Caption         =   "Sex of Animal"
         Height          =   1215
         Left            =   -72240
         TabIndex        =   32
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.CheckBox chkNeuter 
         Caption         =   "Spayed/Neutered"
         Height          =   255
         Left            =   -70440
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtBiteLocation 
         Height          =   285
         Left            =   4680
         TabIndex        =   16
         Top             =   3540
         Width           =   2775
      End
      Begin VB.TextBox txtParents 
         Height          =   285
         Left            =   360
         TabIndex        =   15
         Top             =   3540
         Width           =   4095
      End
      Begin VB.TextBox txtHow 
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   4140
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker dtpVaccinateDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   35
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   19660801
         CurrentDate     =   37593
      End
      Begin MSComCtl2.DTPicker dtpBiteDate 
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   37583
      End
      Begin MSComCtl2.DTPicker dtpBiteTime 
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
         Left            =   5640
         TabIndex        =   14
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660802
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker dtpLabDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   55
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   19660801
         CurrentDate     =   37582
      End
      Begin MSComCtl2.DTPicker dtpEuthanizeDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   52
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   19660801
         CurrentDate     =   37582
      End
      Begin MSComCtl2.DTPicker dtpLastQuarantine 
         Height          =   375
         Left            =   -72960
         TabIndex        =   48
         Top             =   2220
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   37582
      End
      Begin MSComCtl2.DTPicker dtpFirstQuarantine 
         Height          =   375
         Left            =   -74760
         TabIndex        =   47
         Top             =   2220
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   37582
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   37582
      End
      Begin VB.Label lblOwnedStray 
         Caption         =   "This animal is"
         Height          =   255
         Left            =   -74760
         TabIndex        =   118
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   -74640
         TabIndex        =   114
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Left            =   240
         TabIndex        =   113
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3240
         TabIndex        =   112
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   960
         TabIndex        =   111
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblOwnerFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   600
         TabIndex        =   110
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblOwnerAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   109
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblOwnerLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   600
         TabIndex        =   108
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblOwnerCity 
         Caption         =   "City"
         Height          =   255
         Left            =   1080
         TabIndex        =   107
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblOwnerZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4320
         TabIndex        =   106
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblOwnerState 
         Caption         =   "State"
         Height          =   255
         Left            =   3240
         TabIndex        =   105
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblOwnerHomePhone 
         Caption         =   "Home Phone"
         Height          =   255
         Left            =   360
         TabIndex        =   104
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblVictimHomePhone 
         Caption         =   "Home Phone"
         Height          =   255
         Left            =   -74880
         TabIndex        =   103
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblVictimState 
         Caption         =   "State"
         Height          =   255
         Left            =   -72000
         TabIndex        =   102
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblVictimZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   -70920
         TabIndex        =   101
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblVictimCity 
         Caption         =   "City"
         Height          =   255
         Left            =   -74160
         TabIndex        =   100
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblVictimAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   -74520
         TabIndex        =   99
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   -74640
         TabIndex        =   98
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblVictimMunicipality 
         Caption         =   "Municipality"
         Height          =   255
         Left            =   -74760
         TabIndex        =   97
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lblVictimWorkPhone 
         Caption         =   "Work Phone"
         Height          =   255
         Left            =   -74880
         TabIndex        =   96
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblInformed 
         Caption         =   $"frmNewBiteComplaint.frx":09DA
         Height          =   495
         Left            =   -74880
         TabIndex        =   95
         Top             =   600
         Width           =   7335
      End
      Begin VB.Label lblHumaneOfficer 
         Caption         =   "Person issuing quarantine orders:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   94
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblExamVetPhone 
         Caption         =   "Phone"
         Height          =   255
         Left            =   -71760
         TabIndex        =   93
         Top             =   3420
         Width           =   495
      End
      Begin VB.Label lblExamVet 
         Caption         =   "Name of examing Veterinarian or Clinic"
         Height          =   255
         Left            =   -74760
         TabIndex        =   92
         Top             =   3420
         Width           =   2895
      End
      Begin VB.Label lblWhereQuarantined 
         Caption         =   "Where will animal be quarantined"
         Height          =   255
         Left            =   -74760
         TabIndex        =   91
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblLastQuarantine 
         Caption         =   "Last day of quarantine"
         Height          =   255
         Left            =   -72960
         TabIndex        =   90
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label lblFirstQuarantine 
         Caption         =   "First day of quarantine"
         Height          =   255
         Left            =   -74760
         TabIndex        =   89
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "State"
         Height          =   255
         Left            =   -71400
         TabIndex        =   88
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Zip"
         Height          =   255
         Left            =   -70440
         TabIndex        =   87
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "City"
         Height          =   255
         Left            =   -73440
         TabIndex        =   86
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label PhysicianAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   -73800
         TabIndex        =   85
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblPhysicianPhone 
         Caption         =   "Phone"
         Height          =   255
         Left            =   -73680
         TabIndex        =   84
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblFamilyPhysician 
         Caption         =   "Regular family physician"
         Height          =   255
         Left            =   -74880
         TabIndex        =   83
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblClinicDrHospital 
         Caption         =   "If yes, name of Clinic, Dr. or hospital:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   82
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblDateEuthanized 
         Caption         =   "Date died or euthanized:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   81
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblEuthanizedBy 
         Caption         =   "Euthanized by:"
         Height          =   255
         Left            =   -73560
         TabIndex        =   80
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblSpecimen 
         Caption         =   "Specimen prepared by:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   79
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblLabDate 
         Caption         =   "Date sent to State Lab:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   78
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblComments 
         Caption         =   "Comments:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   77
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblDateofBite 
         Caption         =   "Date bite occurred"
         Height          =   255
         Left            =   5640
         TabIndex        =   76
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblTime 
         Caption         =   "Time:"
         Height          =   255
         Left            =   5640
         TabIndex        =   75
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblVaccinateDate 
         Caption         =   "Rabies vaccination expiration date"
         Height          =   255
         Left            =   -74400
         TabIndex        =   74
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblVetPhone 
         Caption         =   "Phone"
         Height          =   255
         Left            =   -71280
         TabIndex        =   73
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblVetClinic 
         Caption         =   "Veterinarian or Clinic"
         Height          =   255
         Left            =   -74400
         TabIndex        =   72
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblAge 
         Caption         =   "Age"
         Height          =   255
         Left            =   -74640
         TabIndex        =   71
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblMarkings 
         Caption         =   "Special Markings"
         Height          =   255
         Left            =   -70440
         TabIndex        =   70
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblType 
         Caption         =   "Type"
         Height          =   255
         Left            =   -74760
         TabIndex        =   69
         Top             =   900
         Width           =   495
      End
      Begin VB.Label lblBreed 
         Caption         =   "Breed"
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label lblColor 
         Caption         =   "Color"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblLocationOfBite 
         Caption         =   "Physical location of bite"
         Height          =   255
         Left            =   4680
         TabIndex        =   63
         Top             =   3300
         Width           =   1815
      End
      Begin VB.Label lblParentsName 
         Caption         =   "Name of parents if victim is under 18 years of age"
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   3300
         Width           =   3615
      End
      Begin VB.Label lblHowBiteHappen 
         Caption         =   "How did bite happen"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   3900
         Width           =   1575
      End
   End
   Begin VB.Label lblDateReported 
      Caption         =   "Date"
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblTakenby 
      Caption         =   "Report Taken By:"
      Height          =   255
      Left            =   3480
      TabIndex        =   59
      Top             =   240
      Width           =   1455
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
Attribute VB_Name = "frmNewBiteComplaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************
'* This form is used to take down information when sombody calls and reports an animal bite.
'* It saves it to the bite table.
'*
'* Written by: Kalonji Kadima
'* Written on: 12-01-2002
'********************************************************************************************

'********************************************************************************************
'* Array that contains all the types of animals
'********************************************************************************************
Private Type combo_info
    name As String
    Number As Integer
End Type

Public bolMatchFound As Boolean            'True = person already in database
Public intPersonNum As Integer          'Number of the person if found

'********************************************************************************************
'* Called when the value in the animal type combo box is changed
'*
'* Written by: Nick DeSteffen
'* Written on: 10-10-2002
'********************************************************************************************
Private Sub cboAnimalType_Click()

Dim breeds() As combo_info          'Array containing all the breeds
Dim rstType As ADODB.Recordset      'Recordset used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes
Dim looper As Integer               'Loop control variable
Dim strSQL As String                'SQL statement

On Error GoTo ErrorHandler

cboBreed.Clear
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

Dim strTakenBy As String           'Name of person who took down report

'1st Tab variables

Dim strFname As String             'First name of the victim
Dim strLname As String             'Last name of the victim
Dim strAddress As String           'Address of the victim
Dim strCity As String              'City of the victim
Dim strState As String             'State of the victim
Dim strZip As String               'Zipcode of the victim
Dim strPhone As String             'Telephone number of the victim
Dim strEmail As String             'Email address of victim
Dim strLicense As String           'Drivers license of the victim
Dim dteDOB As Date                 'Date of birth of victim
Dim dteBiteDate As String          'Date the bite occurred
Dim dteBiteTime As String          'Time the bite occurred
Dim strParents As String           'Name of victim's parents if under 18
Dim strBiteLocation As String      'Physical location of the bite
Dim strHow As String               'How the bite happened

'2nd Tab variables

Dim strOwnedStray As String        'Whether the animal is owned or stray
Dim strOwnerFname As String        'First name of the owner
Dim strOwnerLname As String        'Last name of the owner
Dim strOwnerAddress As String      'Address of the owner
Dim strOwnerCity As String         'City of owner
Dim strOwnerState As String        'State of owner
Dim strOwnerZip As String          'Zipcode of owner
Dim strOwnerPhone As String        'Phone number of owner
Dim strOwnerWorkPhone As String    'Work phone number of owner
Dim strMunicipality As String      'Location

'3rd Tab variables

Dim strSex As String               'Sex of the animal
Dim strAnimalName As String        'Name of animal
Dim intType As Integer             'Type of animal
Dim intBreed As Integer            'Breed of animal
Dim intColor As Integer            'Color of animal
Dim strAge As String               'Approximate age of animal
Dim strVetClinic As String         'Name of vet or clinic
Dim strVetPhone As String          'Phone number of the vet or clinic
Dim dteVaccinateDate As Date       'Date of vaccination
Dim intNeuter As Integer           'Whether the animal has been spayed/neutered
Dim strMarkings As String          'Special markings on animal

'4th Tab variables

Dim intMedical As Integer          'Whether the victim had medical treatment
Dim strClinicDrHospital As String  'Name of clinic, Dr, or hospital
Dim strPhysician As String         'Name of family physician
Dim strPhysicianAddress As String  'Address of family physician
Dim strPhysicianCity As String     'City of family physician
Dim strPhysicianState As String    'State of family physician
Dim strPhysicianZip As String      'Zipcode of family physician
Dim strPhysicianPhone As String    'Phone number of family physician

'5th Tab variables

Dim strHumaneOfficer As String      'Name of Humane Officer
Dim dteFirstQuarantine As Date     'Date of first quarantine
Dim dteLastQuarantine As Date      'Date of last quarantine
Dim strWhereQuarantine As String   'Place where animal will be quarantined
Dim strExamVet As String           'Name of vet examining animal
Dim strExamVetPhone As String      'Phone number of vet examining animal

'6th Tab variables

Dim dteEuthanizeDate As Date       'Date animal was euthanized
Dim strEuthanizeBy As String       'Name of vet who euthanized animal
Dim strPreparedBy As String        'Name of vet who prepared the specimen
Dim strComments As String          'Comments about specimen
Dim dteLabDate As Date             'Date specimen sent to lab

Dim rstInsert As ADODB.Recordset   'Recordset used for interfacing with the database
Dim strSQL As String               'SQL statement
Dim intMsgBox As Integer           'Used for messageboxes
Dim intBiteNum As Integer          'Bite number in the system

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'*************** Checks to see if all the required fields are filled in **************

If txtTakenBy.Text = "" Then
    intMsgBox = MsgBox("Enter your name in the appropriate field", vbOKOnly, "Error!")
    Exit Sub
End If

strTakenBy = Replace(txtTakenBy.Text, "'", "''")

'First tab

If txtFname.Text = "" Or txtLname.Text = "" Or txtAddress.Text = "" Or txtCity.Text = "" Or txtState.Text = "" Or txtZip.Text = "" Or txtPhone.Text = "" Or dtpDOB.Value = Now Then
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
    strZip = Replace(txtZip.Text, "'", "''")
Else
    intMsgBox = MsgBox("Please enter a valid zip code!" & Chr(13) & "Valid formats are ##### or #####-####.", vbOKOnly, "Invalid Zip Code")
    Exit Sub
End If

If Verify_Data.Check_Phone(txtPhone.Text) = True Then
    strPhone = Replace(txtPhone.Text, "'", "''")
Else
    intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

dteBiteDate = dtpBiteDate.Value
dteBiteTime = dtpBiteTime.Value

If (Date - 6574) <= dtpDOB.Value And txtParents.Text = "" Then
    intMsgBox = MsgBox("Please enter the names of the parents in the field.", vbOKOnly, "Error")
    Exit Sub
End If
strParents = Replace(txtParents.Text, "'", "''")
 
If txtBiteLocation.Text = "" Then
    intMsgBox = MsgBox("Please enter in the location of the bite.", vbOKOnly, "Error")
    Exit Sub
End If

If txtHow.Text = "" Then
    intMsgBox = MsgBox("Please explain how the bite occured.", vbOKOnly, "Error")
    Exit Sub
End If

strBiteLocation = Replace(txtBiteLocation.Text, "'", "''")
strHow = Replace(txtHow.Text, "'", "''")

'Second tab

If optOwned.Value = True Then
    strOwnedStray = "Owned"
ElseIf optStray.Value = True Then
    strOwnedStray = "Stray"
Else
    intMsgBox = MsgBox("Please choose whether the animal was owned or a stray.", vbOKOnly, "Error")
    Exit Sub
End If

If optOwned.Value = True Then
    If txtOwnerFname.Text = "" Or txtOwnerLname.Text = "" Or txtOwnerAddress.Text = "" Or txtOwnerCity.Text = "" Or txtOwnerState.Text = "" Or txtOwnerZip.Text = "" Or txtOwnerPhone.Text = "" Then
        intMsgBox = MsgBox("Please fill in all the owner information fields!", vbOKOnly, "Error!")
        Exit Sub
    End If

    strOwnerFname = Replace(txtOwnerFname.Text, "'", "''")
    strOwnerLname = Replace(txtOwnerLname.Text, "'", "''")
    strOwnerAddress = Replace(txtOwnerAddress.Text, "'", "''")
    strOwnerCity = Replace(txtOwnerCity.Text, "'", "''")
    strOwnerState = Replace(txtOwnerState.Text, "'", "''")
    strMunicipality = Replace(txtMunicipality.Text, "'", "''")

    If Verify_Data.Check_Zip(txtOwnerZip.Text) = True Then
        strOwnerZip = Replace(txtOwnerZip.Text, "'", "''")
    Else
        intMsgBox = MsgBox("Please enter a valid zip code!" & Chr(13) & "Valid formats are ##### or #####-####.", vbOKOnly, "Invalid Zip Code")
        Exit Sub
    End If

    If Verify_Data.Check_Phone(txtOwnerPhone.Text) = True Then
        strOwnerPhone = Replace(txtOwnerPhone.Text, "'", "''")
    Else
        intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
        Exit Sub
    End If

    If Verify_Data.Check_Phone(txtOwnerWorkPhone.Text) = True Or txtOwnerWorkPhone.Text = "" Then
        strOwnerWorkPhone = txtOwnerWorkPhone.Text
    Else
        intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
        Exit Sub
    End If
End If

'Third tab

If cboColor.Text = "" Then
    intMsgBox = MsgBox("Please choose the color of the animal.", vbOKOnly, "Error")
    Exit Sub
End If

If cboAnimalType.Text = "" Then
    intMsgBox = MsgBox("Please choose the type of animal.", vbOKOnly, "Error")
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

'Chooses the sex of the animal
If optMale.Value = True Then
    strSex = "M"
ElseIf optFemale.Value = True Then
    strSex = "F"
Else
    intMsgBox = MsgBox("Please select the sex of the animal.", vbOKOnly, "Error")
    Exit Sub
End If
   
If chkNeuter.Value = 1 Then: intNeuter = -1
strAnimalName = Replace(txtAnimalName.Text, "'", "''")
strAge = cboAge.Text
strMarkings = Replace(txtMarkings.Text, "'", "''")

If Not IsNull(dtpVaccinateDate) Then
    dteVaccinateDate = dtpVaccinateDate.Value
End If
    
strVetClinic = Replace(txtVetClinic.Text, "'", "''")
    
If Verify_Data.Check_Phone(txtVetPhone.Text) = True Or txtVetPhone.Text = "" Then
    strVetPhone = Replace(txtVetPhone.Text, "'", "''")
Else
    intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

'Fourth tab

If chkMedical.Value = 1 Then: intMedical = -1

strClinicDrHospital = Replace(txtClinicDrHospital.Text, "'", "''")

strPhysician = Replace(txtPhysician.Text, "'", "''")
strPhysicianAddress = Replace(txtPhysicianAddress.Text, "'", "''")
strPhysicianCity = Replace(txtPhysicianCity.Text, "'", "''")
strPhysicianState = Replace(txtPhysicianState.Text, "'", "''")

If Verify_Data.Check_Zip(txtPhysicianZip.Text) = True Or txtPhysicianZip.Text = "" Then
    strPhysicianZip = txtPhysicianZip.Text
Else
    intMsgBox = MsgBox("Please enter a valid zip code!" & Chr(13) & "Valid formats are ##### or #####-####.", vbOKOnly, "Invalid Zip Code")
    Exit Sub
End If

If Verify_Data.Check_Phone(txtPhysicianPhone.Text) = True Or txtPhysicianPhone.Text = "" Then
    strPhysicianPhone = txtPhysicianPhone.Text
Else
    intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

'Fifth tab

If cboWhereQuarantine.Text = "" Then
    intMsgBox = MsgBox("Please choose where the animal is to be quarantined.", vbOKOnly, "Error")
    Exit Sub
End If

dteFirstQuarantine = dtpFirstQuarantine.Value
dteLastQuarantine = dtpLastQuarantine.Value
strWhereQuarantine = cboWhereQuarantine.Text
strExamVet = Replace(txtExamVet.Text, "'", "''")

If Verify_Data.Check_Phone(txtExamVetPhone.Text) = True Or txtExamVetPhone.Text = "" Then
    strExamVetPhone = txtExamVetPhone.Text
Else
    intMsgBox = MsgBox("Please enter a valid telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

'Sixth tab
If Not IsNull(dtpEuthanizeDate.Value) = True Then
    dteEuthanizeDate = dtpEuthanizeDate.Value
End If
If Not IsNull(dtpLabDate.Value) = True Then
    dteLabDate = dtpLabDate.Value
End If
strEuthanizeBy = Replace(txtEuthanizeBy.Text, "'", "''")
strPreparedBy = Replace(txtPreparedBy.Text, "'", "''")
strComments = Replace(txtComments.Text, "'", "''")
strHumaneOfficer = Replace(txtHumaneOfficer.Text, "'", "''")


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

intMsgBox = MsgBox("Add new bite complaint information to database?", vbYesNo, "Add?")

'Inserts new complaint into database

If intPersonNum <> 0 And intMsgBox = 6 Then
        
    strSQL = "INSERT INTO BITE (BITE_PERSON, "
    strSQL = strSQL & "BITE_REPORT_TAKEN_BY, "
    strSQL = strSQL & "BITE_DATE_OCCURED, "
    strSQL = strSQL & "BITE_TIME_OCCURED, "
    strSQL = strSQL & "BITE_OWNER_FNAME, "
    strSQL = strSQL & "BITE_OWNER_LNAME, "
    strSQL = strSQL & "BITE_ADDRESS, "
    strSQL = strSQL & "BITE_CITY, "
    strSQL = strSQL & "BITE_STATE, "
    strSQL = strSQL & "BITE_ZIP, "
    strSQL = strSQL & "BITE_HOME_PHONE, "
    strSQL = strSQL & "BITE_WORK_PHONE, "
    strSQL = strSQL & "BITE_MUNICIPALITY, "
    strSQL = strSQL & "BITE_VICTIM_PARENTS, "
    strSQL = strSQL & "BITE_BITE_LOCATION, "
    strSQL = strSQL & "BITE_HOW_HAPPENED, "
    strSQL = strSQL & "BITE_OWNED_STRAY, "
    strSQL = strSQL & "BITE_ANIMAL_NAME, "
    strSQL = strSQL & "BITE_TYPE, "
    strSQL = strSQL & "BITE_BREED, "
    strSQL = strSQL & "BITE_COLOR, "
    strSQL = strSQL & "BITE_AGE, "
    strSQL = strSQL & "BITE_SEX, "
    strSQL = strSQL & "BITE_NEUTERED, "
    strSQL = strSQL & "BITE_MARKINGS, "
    strSQL = strSQL & "BITE_ANIMAL_VET_CLINIC, "
    strSQL = strSQL & "BITE_ANIMAL_VET_PHONE, "
    strSQL = strSQL & "BITE_VACCINATED_DATE, "
    strSQL = strSQL & "BITE_MEDICAL, "
    strSQL = strSQL & "BITE_PHYSICIAN_NAME, "
    strSQL = strSQL & "BITE_PHYSICIAN_ADDRESS, "
    strSQL = strSQL & "BITE_PHYSICIAN_CITY, "
    strSQL = strSQL & "BITE_PHYSICIAN_STATE, "
    strSQL = strSQL & "BITE_PHYSICIAN_ZIP, "
    strSQL = strSQL & "BITE_PHYSICIAN_PHONE, "
    strSQL = strSQL & "BITE_INFORMED, "
    strSQL = strSQL & "BITE_FIRST_QUARANTINE, "
    strSQL = strSQL & "BITE_LAST_QUARANTINE, "
    strSQL = strSQL & "BITE_QUARANTINE_PLACE, "
    strSQL = strSQL & "BITE_QUARANTINE_VET, "
    strSQL = strSQL & "BITE_QUARANTINE_VET_PHONE, "
    strSQL = strSQL & "BITE_EUTHANIZED_DATE, "
    strSQL = strSQL & "BITE_EUTHANIZER, "
    strSQL = strSQL & "BITE_PREPARED_BY, "
    strSQL = strSQL & "BITE_LAB_DATE, "
    strSQL = strSQL & "BITE_COMMENTS,"
    strSQL = strSQL & "BITE_TREATED) "
    strSQL = strSQL & "VALUES ("
    strSQL = strSQL & intPersonNum & ", '" & strTakenBy
    strSQL = strSQL & "', '" & dteBiteDate & "', '" & dteBiteTime & "', '" & strOwnerFname & "', '" & strOwnerLname
    strSQL = strSQL & "', '" & strOwnerAddress & "', '" & strOwnerCity & "', '" & strOwnerState
    strSQL = strSQL & "', '" & strOwnerZip & "', '" & strOwnerPhone & "', '" & strOwnerWorkPhone
    strSQL = strSQL & "', '" & strMunicipality & "', '" & strParents
    strSQL = strSQL & "', '" & strBiteLocation & "', '" & strHow & "', '" & strOwnedStray
    strSQL = strSQL & "', '" & strAnimalName & "', " & intType & ", " & intBreed & ", " & intColor
    strSQL = strSQL & ", '" & strAge & "', '" & strSex & "', " & intNeuter & ", '" & strMarkings
    strSQL = strSQL & "', '" & strVetClinic & "', '" & strVetPhone
    strSQL = strSQL & "', " & Check_Date(dtpVaccinateDate.Value) & ", " & intMedical
    strSQL = strSQL & ", '" & strPhysician & "', '" & strPhysicianAddress & "', '" & strPhysicianCity
    strSQL = strSQL & "', '" & strPhysicianState & "', '" & strPhysicianZip & "', '" & strPhysicianPhone
    strSQL = strSQL & "', '" & strHumaneOfficer & "', '" & dteFirstQuarantine & "', '" & dteLastQuarantine
    strSQL = strSQL & "', '" & strWhereQuarantine & "', '" & strExamVet & "', '" & strExamVetPhone
    strSQL = strSQL & "', " & Check_Date(dtpEuthanizeDate.Value) & ", '" & strEuthanizeBy & "', '" & strPreparedBy
    strSQL = strSQL & "', " & Check_Date(dtpLabDate.Value) & ", '" & strComments & "', '" & strClinicDrHospital & "')"

    Open_Recordsets.objConnection.Execute (strSQL)
ElseIf intMsgBox <> 6 Then
    Exit Sub
Else
    intMsgBox = MsgBox("There was an error, please restart the bite complaint form!", vbCritical, "Error")
    Exit Sub
End If

Set rstInsert = Open_Recordsets.objConnection.Execute("SELECT MAX(BITE_NUMBER) AS NUM FROM BITE")

'Gets the newest bite number

If rstInsert.EOF = False Then
With rstInsert
    rstInsert.MoveFirst
    Do While Not rstInsert.EOF
        intBiteNum = ![NUM]
        rstInsert.MoveNext
    Loop
End With
End If
    
'Displays the bite form
    
frmShowBite.intBiteNum = intBiteNum
frmShowBite.Show
    
frmPCHS_Main.Show
Unload Me

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & Err.Number & Chr(13) & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
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

Dim strFname As String              'First name of the original owner
Dim strLname As String              'Last name of the original owner
Dim strAddress As String            'Address of the original owner
Dim strCity As String               'City of the original owner
Dim strState As String              'State of the original owner
Dim strZip As String                'Zipcode of the original owner
Dim strPhone As String              'Telephone number of the original owner
Dim strEmail As String              'Email address of original owner
Dim strLicense As String            'Drivers license of the original owner
Dim dteDOB As Date                  'Date of birth of original owner

Dim intType As Integer              'Which form is being searched from
Dim rstSearch As New ADODB.Recordset 'Recordset used for interfacing with the database
Dim intMsgBox As Integer             'Used for messageboxes

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset
Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM PERSON WHERE PERSON_LNAME = '" & txtLname.Text & "'")

If rstSearch.EOF <> True Then
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 7
    
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

Dim types() As combo_info               'Array of animal types
Dim colors() As combo_info              'Array of colors
Dim looper As Integer                   'Loop control variable

Dim rstType As ADODB.Recordset          'Recordset for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes
Dim strSQL As String                    'SQL statement

On Error GoTo ErrorHandler

lblDateReported.Caption = "Date Reported: " & Date
dtpDOB.Value = Date
dtpBiteDate.Value = Date
dtpBiteTime.Value = "12:00:00 AM"
dtpVaccinateDate.Value = Null
dtpFirstQuarantine.Value = Date
dtpLastQuarantine.Value = Date
dtpEuthanizeDate.Value = Null
dtpLabDate.Value = Null

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

Private Sub cmdCancel_Click()
Unload Me
frmPCHS_Main.Show
End Sub

'Function which checks to see if a date picker is null or not, if its null it returns
'Null in a string for entering into the SQL statement, otherwise it returns the date with
' aposterphies around it

Function Check_Date(dteCheck) As String

If IsNull(dteCheck) = True Then
    Check_Date = "Null"
Else
    Check_Date = "'" & dteCheck & "'"
End If

End Function
