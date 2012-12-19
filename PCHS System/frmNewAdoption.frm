VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNewAdoption 
   Caption         =   "New Adoption"
   ClientHeight    =   7980
   ClientLeft      =   465
   ClientTop       =   1035
   ClientWidth     =   9630
   Icon            =   "frmNewAdoption.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleMode       =   0  'User
   ScaleWidth      =   9647.341
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   8160
      TabIndex        =   66
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   67
      Top             =   7320
      Width           =   1215
   End
   Begin TabDlg.SSTab sstAdoption 
      Height          =   6495
      Left            =   120
      TabIndex        =   71
      Top             =   720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Adoptor"
      TabPicture(0)   =   "frmNewAdoption.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAgent"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmPersonal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmEmployment"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAgent"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Housing/Vet"
      TabPicture(1)   =   "frmNewAdoption.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmVet"
      Tab(1).Control(1)=   "frmLandlord"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "chkAllergies"
      Tab(1).Control(4)=   "txtSpouse"
      Tab(1).Control(5)=   "txtAge"
      Tab(1).Control(6)=   "txtChildren"
      Tab(1).Control(7)=   "txtAdults"
      Tab(1).Control(8)=   "txtHomeLength"
      Tab(1).Control(9)=   "lblSpouse"
      Tab(1).Control(10)=   "lblAge"
      Tab(1).Control(11)=   "lblChildren"
      Tab(1).Control(12)=   "lblAdults"
      Tab(1).Control(13)=   "lblNumAdults"
      Tab(1).Control(14)=   "lblHomeLength"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Additional"
      TabPicture(2)   =   "frmNewAdoption.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAdjustConcerns"
      Tab(2).Control(1)=   "cboHearAbout"
      Tab(2).Control(2)=   "chkResponsibleLife"
      Tab(2).Control(3)=   "chkAdjust"
      Tab(2).Control(4)=   "txtWho"
      Tab(2).Control(5)=   "txtResponsible"
      Tab(2).Control(6)=   "txtStillHaveReason"
      Tab(2).Control(7)=   "txtSurrenderReason"
      Tab(2).Control(8)=   "txtNeuterReason"
      Tab(2).Control(9)=   "txtNeuterCost"
      Tab(2).Control(10)=   "chkOtherAdults"
      Tab(2).Control(11)=   "chkPrevious"
      Tab(2).Control(12)=   "chkStillHave"
      Tab(2).Control(13)=   "chkSurrender"
      Tab(2).Control(14)=   "chkNeuter"
      Tab(2).Control(15)=   "chkAfford"
      Tab(2).Control(16)=   "lblAdjustConcerns(0)"
      Tab(2).Control(17)=   "lblHearAbout(0)"
      Tab(2).Control(18)=   "lblWho"
      Tab(2).Control(19)=   "lblResposible"
      Tab(2).Control(20)=   "lblStillHaveReason"
      Tab(2).Control(21)=   "lblSurrenderReason"
      Tab(2).Control(22)=   "lblNeuterReason"
      Tab(2).Control(23)=   "lblNeuterCost"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "Dog Information"
      TabPicture(3)   =   "frmNewAdoption.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblHoursAway"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "Label10"
      Tab(3).Control(5)=   "Label11"
      Tab(3).Control(6)=   "lblAnimalLocation(0)"
      Tab(3).Control(7)=   "lblAdoptReason(0)"
      Tab(3).Control(8)=   "lblOutdoors(0)"
      Tab(3).Control(9)=   "lblWhenNotHome"
      Tab(3).Control(10)=   "lblHousebreak"
      Tab(3).Control(11)=   "txtHoursAway"
      Tab(3).Control(12)=   "cboLocation"
      Tab(3).Control(13)=   "cboOutside"
      Tab(3).Control(14)=   "txtDogProb1"
      Tab(3).Control(15)=   "txtDogProb2"
      Tab(3).Control(16)=   "txtDogProb3"
      Tab(3).Control(17)=   "txtDogProb4"
      Tab(3).Control(18)=   "chkFamiliar"
      Tab(3).Control(19)=   "cboWhenNotHome"
      Tab(3).Control(20)=   "txtHousebreak"
      Tab(3).Control(21)=   "chkHouse"
      Tab(3).Control(22)=   "cboReason"
      Tab(3).ControlCount=   23
      TabCaption(4)   =   "Cat Information"
      TabPicture(4)   =   "frmNewAdoption.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblProb4"
      Tab(4).Control(1)=   "lblProb3"
      Tab(4).Control(2)=   "lblProb2"
      Tab(4).Control(3)=   "lblProbs"
      Tab(4).Control(4)=   "lblProb1"
      Tab(4).Control(5)=   "lblAnimalLocation(1)"
      Tab(4).Control(6)=   "lblOutdoors(1)"
      Tab(4).Control(7)=   "lblAdoptReason(1)"
      Tab(4).Control(8)=   "cboCatKept"
      Tab(4).Control(9)=   "cboCatOutside"
      Tab(4).Control(10)=   "cboCatReason"
      Tab(4).Control(11)=   "txtProb4"
      Tab(4).Control(12)=   "txtProb3"
      Tab(4).Control(13)=   "txtProb2"
      Tab(4).Control(14)=   "txtProb1"
      Tab(4).Control(15)=   "chkEverOwned"
      Tab(4).ControlCount=   16
      Begin VB.CheckBox chkEverOwned 
         Caption         =   "Owned cat before "
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74760
         TabIndex        =   58
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtProb1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71400
         TabIndex        =   62
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtProb2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71400
         TabIndex        =   63
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtProb3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71400
         TabIndex        =   64
         Top             =   2520
         Width           =   4575
      End
      Begin VB.TextBox txtProb4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71400
         TabIndex        =   65
         Top             =   3120
         Width           =   4575
      End
      Begin VB.ComboBox cboCatReason 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":0956
         Left            =   -74760
         List            =   "frmNewAdoption.frx":096C
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cboCatOutside 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":09AF
         Left            =   -74760
         List            =   "frmNewAdoption.frx":09BF
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cboCatKept 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":09F8
         Left            =   -74760
         List            =   "frmNewAdoption.frx":0A05
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cboReason 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":0A20
         Left            =   -74760
         List            =   "frmNewAdoption.frx":0A39
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkHouse 
         Caption         =   "Dog house and fence already installed"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74760
         TabIndex        =   50
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtHousebreak 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74760
         TabIndex        =   52
         Top             =   4590
         Width           =   3495
      End
      Begin VB.ComboBox cboWhenNotHome 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":0A96
         Left            =   -74760
         List            =   "frmNewAdoption.frx":0AA9
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   3870
         Width           =   2055
      End
      Begin VB.CheckBox chkFamiliar 
         Caption         =   "Familiar with breed chosen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtDogProb4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70440
         TabIndex        =   57
         Top             =   3240
         Width           =   4335
      End
      Begin VB.TextBox txtDogProb3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70440
         TabIndex        =   56
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox txtDogProb2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70440
         TabIndex        =   55
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox txtDogProb1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70440
         TabIndex        =   54
         Top             =   1440
         Width           =   4335
      End
      Begin VB.ComboBox cboOutside 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":0AFA
         Left            =   -74760
         List            =   "frmNewAdoption.frx":0B0D
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cboLocation 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":0B6E
         Left            =   -74760
         List            =   "frmNewAdoption.frx":0B7B
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtHoursAway 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74760
         TabIndex        =   53
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox txtAgent 
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtAdjustConcerns 
         Height          =   285
         Left            =   -70080
         TabIndex        =   44
         Top             =   2160
         Width           =   3975
      End
      Begin VB.ComboBox cboHearAbout 
         Height          =   315
         ItemData        =   "frmNewAdoption.frx":0BB5
         Left            =   -70080
         List            =   "frmNewAdoption.frx":0BC5
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Frame frmVet 
         Caption         =   "Veterinarian"
         Height          =   2295
         Left            =   -70800
         TabIndex        =   92
         Top             =   840
         Width           =   4935
         Begin VB.CheckBox chkVet 
            Caption         =   "Has a veterinarian"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtVetPhone 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   28
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtVetName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtVetConsider 
            Height          =   285
            Left            =   240
            TabIndex        =   29
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblVetName 
            Caption         =   "Name of veterinarian"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   95
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblVetPhone 
            Caption         =   "Phone Number"
            Height          =   255
            Left            =   2400
            TabIndex        =   94
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblVetName 
            Caption         =   "If you don't have a current vet, which vet are you considering?"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   93
            Top             =   1320
            Width           =   4455
         End
      End
      Begin VB.CheckBox chkResponsibleLife 
         Caption         =   "Prepared to be responsible for animal for entire life"
         Height          =   255
         Left            =   -70080
         TabIndex        =   42
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Frame frmLandlord 
         Caption         =   "Landlord information"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   89
         Top             =   1920
         Width           =   2175
         Begin VB.TextBox txtLandLordTelephone 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtLandlord 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblTelephone 
            Caption         =   "Telephone Number:"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblLandlord 
            Caption         =   "Name:"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Housing Type"
         Height          =   975
         Left            =   -74760
         TabIndex        =   87
         Top             =   840
         Width           =   3375
         Begin VB.ComboBox cboHomeType 
            Height          =   315
            ItemData        =   "frmNewAdoption.frx":0BE9
            Left            =   1440
            List            =   "frmNewAdoption.frx":0BFC
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optRents 
            Caption         =   "Rents"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optOwns 
            Caption         =   "Owns"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblType 
            Caption         =   "Type of home:"
            Height          =   255
            Left            =   1440
            TabIndex        =   88
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame frmEmployment 
         Caption         =   "Employment"
         Height          =   1095
         Left            =   240
         TabIndex        =   84
         Top             =   4440
         Width           =   8895
         Begin VB.TextBox txtWorkPhone 
            Height          =   285
            Left            =   3360
            TabIndex        =   14
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtEmployer 
            Height          =   285
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblWorkPhone 
            Caption         =   "Work Phone Number"
            Height          =   255
            Left            =   3360
            TabIndex        =   86
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblEmployer 
            Caption         =   "Employer/Source of Income"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkAdjust 
         Caption         =   "Prepared to allow animal time to adjust to new home"
         Height          =   255
         Left            =   -70080
         TabIndex        =   43
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CheckBox chkAllergies 
         Caption         =   "Someone in household has allergies"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   5880
         Width           =   2895
      End
      Begin VB.TextBox txtSpouse 
         Height          =   285
         Left            =   -74760
         TabIndex        =   24
         Top             =   5520
         Width           =   2895
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   -73320
         TabIndex        =   22
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtChildren 
         Height          =   285
         Left            =   -74040
         TabIndex        =   21
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtAdults 
         Height          =   285
         Left            =   -74760
         TabIndex        =   20
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox txtHomeLength 
         Height          =   285
         Left            =   -74760
         TabIndex        =   23
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtWho 
         Height          =   285
         Left            =   -74760
         TabIndex        =   31
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtResponsible 
         Height          =   285
         Left            =   -74760
         TabIndex        =   32
         Top             =   2070
         Width           =   2295
      End
      Begin VB.TextBox txtStillHaveReason 
         Height          =   285
         Left            =   -74760
         TabIndex        =   35
         Top             =   3390
         Width           =   4455
      End
      Begin VB.TextBox txtSurrenderReason 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74760
         TabIndex        =   37
         Top             =   4230
         Width           =   4455
      End
      Begin VB.TextBox txtNeuterReason 
         Height          =   285
         Left            =   -74760
         TabIndex        =   40
         Top             =   5790
         Width           =   4455
      End
      Begin VB.TextBox txtNeuterCost 
         Height          =   285
         Left            =   -74760
         TabIndex        =   38
         Top             =   4920
         Width           =   855
      End
      Begin VB.CheckBox chkOtherAdults 
         Caption         =   "Other adults in  household aware you are looking for a pet"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   4455
      End
      Begin VB.CheckBox chkPrevious 
         Caption         =   "Adopted a pet from Portage County Humane Society before"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   2520
         Width           =   4575
      End
      Begin VB.CheckBox chkStillHave 
         Caption         =   "Still have this pet"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CheckBox chkSurrender 
         Caption         =   "Surrendered a pet from this Humane Society before"
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   3720
         Width           =   3975
      End
      Begin VB.CheckBox chkNeuter 
         Caption         =   "Plan on having pet spayed/neutered"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   5280
         Width           =   4575
      End
      Begin VB.CheckBox chkAfford 
         Caption         =   "Can afford cost of pet ($300 - $500 anually)"
         Height          =   255
         Left            =   -70080
         TabIndex        =   41
         Top             =   840
         Width           =   4095
      End
      Begin VB.Frame frmPersonal 
         Caption         =   "Adopter Information"
         Height          =   3015
         Left            =   240
         TabIndex        =   72
         Top             =   1320
         Width           =   8895
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   375
            Left            =   3720
            TabIndex        =   111
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   47382529
            CurrentDate     =   37579
         End
         Begin VB.CheckBox chkStudent 
            Caption         =   "Student"
            Height          =   255
            Left            =   6000
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtStudentLocation 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   12
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtPhone 
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtState 
            Height          =   285
            Left            =   3720
            TabIndex        =   6
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtZip 
            Height          =   285
            Left            =   4560
            TabIndex        =   7
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtLname 
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtAddress 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtFname 
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Top             =   2280
            Width           =   3255
         End
         Begin VB.TextBox txtLicense 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Top             =   2640
            Width           =   3255
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search People"
            Height          =   375
            Left            =   7320
            TabIndex        =   3
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblStudent 
            Caption         =   "If student, where?"
            Height          =   255
            Left            =   6000
            TabIndex        =   83
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblDOB 
            Caption         =   "DOB"
            Height          =   255
            Left            =   3120
            TabIndex        =   82
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label lblEmail 
            Caption         =   "Email"
            Height          =   255
            Left            =   840
            TabIndex        =   81
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label lblPhone 
            Caption         =   "Phone Number"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblState 
            Caption         =   "State"
            Height          =   255
            Left            =   3120
            TabIndex        =   79
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblZip 
            Caption         =   "Zip"
            Height          =   255
            Left            =   4200
            TabIndex        =   78
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblCIty 
            Caption         =   "City"
            Height          =   255
            Left            =   960
            TabIndex        =   77
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblLname 
            Caption         =   "Last Name"
            Height          =   255
            Left            =   600
            TabIndex        =   76
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblAddress 
            Caption         =   "Address"
            Height          =   255
            Left            =   720
            TabIndex        =   75
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblFname 
            Caption         =   "First Name"
            Height          =   255
            Left            =   600
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblLicense 
            Caption         =   "Drivers License"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   73
            Top             =   2640
            Width           =   1215
         End
      End
      Begin VB.Label lblAdoptReason 
         Caption         =   "I want to adopt this cat becuase:"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   130
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lblOutdoors 
         Caption         =   "If this pet is left outdoors it will:"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   129
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblAnimalLocation 
         Caption         =   "Where will the animal be kept?"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   128
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblProb1 
         Caption         =   "Urinating/Defecating outside litter box:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   127
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label lblProbs 
         Caption         =   "How do you plan to prevent behavior problems, such as:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   126
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblProb2 
         Caption         =   "Scratching furniture:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   125
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label lblProb3 
         Caption         =   "Running away:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   124
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label lblProb4 
         Caption         =   "Jumping on counters:"
         Height          =   255
         Left            =   -71400
         TabIndex        =   123
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label lblHousebreak 
         Caption         =   "How do you plan to housebreak your dog/puppy?"
         Height          =   255
         Left            =   -74760
         TabIndex        =   122
         Top             =   4350
         Width           =   3615
      End
      Begin VB.Label lblWhenNotHome 
         Caption         =   "When you are not at home, this pet will be:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   121
         Top             =   3630
         Width           =   3015
      End
      Begin VB.Label lblOutdoors 
         Caption         =   "If this pet is left outdoors it will:"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   120
         Top             =   2550
         Width           =   2175
      End
      Begin VB.Label lblAdoptReason 
         Caption         =   "I want to adopt this dog/puppy becuase:"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   119
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblAnimalLocation 
         Caption         =   "Where will the animal be kept?"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   118
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Running Away:"
         Height          =   255
         Left            =   -70440
         TabIndex        =   117
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Separation Anxiety:"
         Height          =   255
         Left            =   -70440
         TabIndex        =   116
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Chewing:"
         Height          =   255
         Left            =   -70440
         TabIndex        =   115
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Barking:"
         Height          =   255
         Left            =   -70440
         TabIndex        =   114
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "How do you plan on preventing behavior problems, such as:"
         Height          =   255
         Left            =   -70440
         TabIndex        =   113
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblHoursAway 
         Caption         =   "Average number of hours animal will be left alone:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   112
         Top             =   5040
         Width           =   3735
      End
      Begin VB.Label lblAgent 
         Caption         =   "Humane Society Agent:"
         Height          =   255
         Left            =   360
         TabIndex        =   110
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblAdjustConcerns 
         Caption         =   "Adjusting Concerns"
         Height          =   255
         Index           =   0
         Left            =   -70080
         TabIndex        =   109
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblHearAbout 
         Caption         =   "Learned about Humane Society from:"
         Height          =   255
         Index           =   0
         Left            =   -70080
         TabIndex        =   108
         Top             =   2760
         Width           =   3495
      End
      Begin VB.Label lblSpouse 
         Caption         =   "Name (s) of Spouse or other adults in household"
         Height          =   255
         Left            =   -74760
         TabIndex        =   107
         Top             =   5280
         Width           =   3495
      End
      Begin VB.Label lblAge 
         Caption         =   "Ages of childern:"
         Height          =   255
         Left            =   -73320
         TabIndex        =   106
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblChildren 
         Caption         =   "Children"
         Height          =   255
         Left            =   -74160
         TabIndex        =   105
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lblAdults 
         Caption         =   "Adults"
         Height          =   255
         Left            =   -74760
         TabIndex        =   104
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblNumAdults 
         Caption         =   "Number of people living in household:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   103
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label lblHomeLength 
         Caption         =   "How long have you lived at this address? (Years)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   102
         Top             =   4680
         Width           =   3495
      End
      Begin VB.Label lblWho 
         Caption         =   "Person pet is being adopted for:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   101
         Top             =   1230
         Width           =   2535
      End
      Begin VB.Label lblResposible 
         Caption         =   "Person who pet will be primary responsibility of:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   100
         Top             =   1830
         Width           =   3495
      End
      Begin VB.Label lblStillHaveReason 
         Caption         =   "If not, explain"
         Height          =   255
         Left            =   -74760
         TabIndex        =   99
         Top             =   3150
         Width           =   975
      End
      Begin VB.Label lblSurrenderReason 
         Caption         =   "If yes, explain"
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   3990
         Width           =   975
      End
      Begin VB.Label lblNeuterReason 
         Caption         =   "If no, explain"
         Height          =   255
         Left            =   -74760
         TabIndex        =   97
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label lblNeuterCost 
         Caption         =   "Estimated cost to spay or neuter animal:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   96
         Top             =   4680
         Width           =   4335
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
      Left            =   6240
      TabIndex        =   70
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblAnimalNum 
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
      Left            =   3000
      TabIndex        =   69
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblAnimal 
      Caption         =   "Animal Number: "
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
      TabIndex        =   68
      Top             =   120
      Width           =   2655
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
Attribute VB_Name = "frmNewAdoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************************
'* This form is used to record all the information needed to be checked or verified in an
'* adoption.  It is all saved in the adoption table.
'*******************************************************************************************

Public bolMatchFound As Boolean            'True = person already in database
Dim intType As Integer                  'Type of animal
Dim intReceiptNum As Integer            'Number of the receipt if found
Public intPersonNum As Integer          'Number of the person if found
Public intAnimalNum As Integer          'Number of the animal - Required
Public intAdoptionNum As Integer        'Number of the adoption

Private Sub cmdSave_Click()

'Adopter variables

Dim strFname As String              'First name of the Adopter
Dim strLname As String              'Last name of the Adopter
Dim strAddress As String            'Address of the Adopter
Dim strCity As String               'City of the Adopter
Dim strState As String              'State of the Adopter
Dim strZip As String                'Zipcode of the Adopter
Dim strPhone As String              'Telephone number of the Adopter
Dim strEmail As String              'Email address of Adopter
Dim strLicense As String            'Drivers license of the Adopter
Dim dteDOB As Date                  'Date of birth of Adopter

'General variables

Dim strAgent As String              'Agent handling the adoption

Dim intOwnRent As Integer           'Whether adopter owns house, apartment, etc.
Dim strLandlord As String           'Name of landlord
Dim strLandlordTelephone As String  'Telephone number of landlord
Dim strStatus As String             'Status of the adoption
Dim strHomeType As String           'Type of home adopter lives in
Dim intAdults As Integer            'Number of adults in the household
Dim intChildren As Integer          'Number of children in the household
Dim strAge As String                'Ages of children in the household
Dim strSpouse As String             'Name of spouse that lives in household
Dim intStudent As Integer           'Whether the adopter is a student
Dim strStudentLocation As String    'Where the student goes to school
Dim strEmployer As String           'Name of employer
Dim strWorkPhone As String          'Work phone number of adopter
Dim intHomeLength As Integer        'How long adopter has lived at present home
Dim intOtherAdults As Integer       'Other adults in the household aware of adoption
Dim strWho As String                'Who the pet is for
Dim strResponsible As String        'Who is responsible for the pet
Dim intResponsibleLife As Integer   'Whether adopter will be responsible for the pet's entire life
Dim intPrevious As Integer          'Whether adopted an pet from PCHS before
Dim intStillHave As Integer         'Whether still have the pet
Dim strStillHaveReason As String    'If don't have pet reason
Dim intSurrender As Integer         'Whether adopter surrendered a pet before
Dim strSurrenderReason As String    'If surrendered pet reason
Dim intNeuter As Integer            'Whether adopter plans to neuter the pet
Dim strNeuterReason As String       'If not going to neuter reason
Dim intNeuterCost As Double         'Cost of neuter
Dim intAfford As Double             'Whether adopter can afford the pet
Dim intAllergies As Integer         'If anyone in the house has allergies
Dim intAdjust As Integer            'If willing to allow animal time to adjust to new home
Dim strAdjustConcerns As String     'Any concerns of pet adjusting
Dim strHearAbout As String          'Where adopter heard of PCHS

'Dog and cat specific variables

Dim strLocation As String           'Where animal will be kept
Dim strOutside As String            'Where pet will be left if outside
Dim intFamiliar As Integer          'Whether adopter is familiar with the breed of dog adopting
Dim intHouse As Integer             'Whether dog house and fence have been built
Dim strReason As String             'Reason for adopting this pet
Dim strWhenNotHome As String        'Where pet will be when adopter is not home
Dim strHousebreak As String         'How does adopter plan to housebreak the pet
Dim intHoursAway As Integer         'How many hours will adopter leave pet alone
Dim intEverOwned As Integer         'If everowned a cat before
Dim strCatKept As String            'Where cat is kept
Dim strCatOutside As String         'Where cat is kept if outside
Dim strCatReason As String          'Reason for adopting cat
Dim strProb1 As String              'Urinating/defecating problem
Dim strProb2 As String              'Scratching furniture problem
Dim strProb3 As String              'Running away problem
Dim strProb4 As String              'Jumping on counters problem
Dim strDogProb1 As String           'Barking problem
Dim strDogProb2 As String           'Chewing problem
Dim strDogProb3 As String           'Separation anxiety problem
Dim strDogProb4 As String           'Running away problem

'Vet variables

Dim intVet As Integer               'Whether adopter has a vet
Dim strVetName As String            'Name of vet
Dim strVetPhone As String           'Phone number of vet
Dim strVetConsider As String        'Name of vet adopter is considering

Dim intMsgBox As Integer            'Used for messagboxes
Dim rstInsert As ADODB.Recordset    'Used for inserting data into the database
Dim strSQL As String                'SQL string

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'Gets information from the first tab

strAgent = Replace(txtAgent.Text, "'", "''")

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

strStatus = "P"

strEmployer = Replace(txtEmployer.Text, "'", "''")

If Verify_Data.Check_Phone(txtWorkPhone.Text) = True Or txtWorkPhone.Text = "" Then
    strWorkPhone = txtWorkPhone.Text
Else
    intMsgBox = MsgBox("Please enter a valid work telephone number!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
    Exit Sub
End If

If chkStudent.Value = 1 Then: intStudent = -1
strStudentLocation = Replace(txtStudentLocation.Text, "'", "''")

'Gets information from the second tab

strHomeType = cboHomeType.Text

If cboHomeType.Text = "" Then
    intMsgBox = MsgBox("Please choose your type of home!", vbOKOnly, "Error!")
    Exit Sub
End If
If optRents.Value = True And txtLandlord.Text = "" Then
    intMsgBox = MsgBox("Please enter the landlords name!", vbOKOnly, "Error!")
    Exit Sub
End If

If optOwns.Value = True Then: intOwnRent = -1
If optRents.Value = True Then: intOwnRent = 1

If intOwnRent = 1 Then
    strLandlord = Replace(txtLandlord.Text, "'", "''")
    If Verify_Data.Check_Phone(txtLandLordTelephone.Text) = True Then
        strLandlordTelephone = Replace(txtLandLordTelephone.Text, "'", "''")
    Else
        intMsgBox = MsgBox("Please enter a valid telephone number for your landlord!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
        Exit Sub
    End If
End If

If IsNumeric(txtAdults.Text) = True Then
    intAdults = txtAdults.Text
Else
    intMsgBox = MsgBox("Please enter a numeric value for the number of household adults.", vbOKOnly, "Error!")
    Exit Sub
End If

If IsNumeric(txtChildren.Text) = True Then
    intChildren = txtChildren.Text
Else
    intMsgBox = MsgBox("Please enter a numeric value for the number of household childern.", vbOKOnly, "Error!")
    Exit Sub
End If

strAge = Replace(txtAge.Text, "'", "''")
strSpouse = Replace(txtSpouse.Text, "'", "''")
If IsNumeric(txtHomeLength.Text) = True Then
    intHomeLength = txtHomeLength.Text
Else
    intMsgBox = MsgBox("Please enter a numeric value for the length of time living at the home.", vbOKOnly, "Error!")
    Exit Sub
End If
If chkAllergies.Value = 1 Then: intAllergies = -1

If chkVet.Value = 1 Then: intVet = -1

If intVet = -1 Then
    strVetName = Replace(txtVetName.Text, "'", "''")
    If Verify_Data.Check_Phone(txtVetPhone.Text) = True Then
        strVetPhone = Replace(txtVetPhone.Text, "'", "''")
    Else
        intMsgBox = MsgBox("Please enter a valid telephone number for your vet!" & Chr(13) & "Valid formats are ####### or ###-###-####.", vbOKOnly, "Invalid Telephone Number")
        Exit Sub
    End If
Else
    strVetConsider = Replace(txtVetConsider.Text, "'", "''")
End If

'Gets information from the third tab

If chkOtherAdults.Value = 1 Then: intOtherAdults = -1
strWho = Replace(txtWho.Text, "'", "''")
strResponsible = Replace(txtResponsible.Text, "'", "''")
If chkStillHave.Value = 1 Then: intStillHave = -1
strStillHaveReason = Replace(txtStillHaveReason.Text, "'", "''")

If chkPrevious.Value = 1 Then
    intPrevious = -1
    If intStillHave <> -1 And strStillHaveReason = "" Then
        intMsgBox = MsgBox("Please explain why you don't have this animal anymore", vbOKOnly, "Explain")
        Exit Sub
    End If
End If

If chkSurrender.Value = 1 Then: intSurrender = -1
strSurrenderReason = Replace(txtSurrenderReason.Text, "'", "''")

If intSurrender = -1 And strSurrenderReason = "" Then
    intMsgBox = MsgBox("Please explain why you surrendered this animal.", vbOKOnly, "Explain")
    Exit Sub
End If

If chkNeuter.Value = 1 Then: intNeuter = -1
strNeuterReason = Replace(txtNeuterReason.Text, "'", "''")

If intNeuter <> -1 And strNeuterReason = "" Then
    intMsgBox = MsgBox("Please explain why you aren't having this animal neutered.", vbOKOnly, "Explain")
    Exit Sub
End If

If IsNumeric(txtNeuterCost.Text) = True Then
    intNeuterCost = txtNeuterCost.Text
Else
    intMsgBox = MsgBox("Please enter a valid estimated spay/neuter cost.", vbOKOnly, "Error!")
    Exit Sub
End If

If chkAfford.Value = 1 Then: intAfford = -1
If chkResponsibleLife.Value = 1 Then intResponsibleLife = -1
If chkAdjust.Value = 1 Then: intAdjust = -1
strAdjustConcerns = Replace(txtAdjustConcerns.Text, "'", "''")
strHearAbout = cboHearAbout.Text

'Information from the dog Tab
If intType = 1 Then
    strReason = cboReason.Text
    strLocation = cboLocation.Text
    strOutside = cboOutside.Text
    If chkFamiliar.Value = 1 Then: intFamiliar = -1
    If chkHouse.Value = 1 Then: intHouse = -1
    strWhenNotHome = cboWhenNotHome.Text
    strHousebreak = Replace(txtHousebreak.Text, "'", "''")

    If IsNumeric(txtHoursAway.Text) = True Then
        intHoursAway = txtHoursAway.Text
    Else
        intMsgBox = MsgBox("Please enter a valid number for the number of hours the dog will be left alone.", vbOKOnly, "Error!")
        Exit Sub
    End If

strProb1 = Replace(txtDogProb1.Text, "'", "''")
strProb2 = Replace(txtDogProb2.Text, "'", "''")
strProb3 = Replace(txtDogProb3.Text, "'", "''")
strProb4 = Replace(txtDogProb4.Text, "'", "''")

ElseIf intType = 2 Then

'Information from the cat tab

    If chkEverOwned.Value = 1 Then: intEverOwned = -1
    strCatKept = cboCatKept.Text
    strCatOutside = cboCatOutside.Text
    strCatReason = cboCatReason.Text
    
    strProb1 = Replace(txtProb1.Text, "'", "''")
    strProb2 = Replace(txtProb2.Text, "'", "''")
    strProb3 = Replace(txtProb3.Text, "'", "''")
    strProb4 = Replace(txtProb4.Text, "'", "''")

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

'Insert new adoption information into database

strSQL = "INSERT INTO ADOPTION (ADOPTION_ADOPTORNUM, "
strSQL = strSQL & "ADOPTION_ANIMAL, "
strSQL = strSQL & "ADOPTION_STATUS, "
strSQL = strSQL & "ADOPTION_AGENT, "
strSQL = strSQL & "ADOPTION_EMPLOYER, "
strSQL = strSQL & "ADOPTION_WORK_PHONE, "
strSQL = strSQL & "ADOPTION_STUDENT, "
strSQL = strSQL & "ADOPTION_STUDENT_LOCATION, "
strSQL = strSQL & "ADOPTION_OWN_RENT, "
strSQL = strSQL & "ADOPTION_HOME_TYPE, "
strSQL = strSQL & "ADOPTION_LANDLORD, "
strSQL = strSQL & "ADOPTION_LANDLORD_TELEPHONE, "
strSQL = strSQL & "ADOPTION_ADULTS, "
strSQL = strSQL & "ADOPTION_CHILDREN, "
strSQL = strSQL & "ADOPTION_AGE, "
strSQL = strSQL & "ADOPTION_HOME_LENGTH, "
strSQL = strSQL & "ADOPTION_SPOUSE, "
strSQL = strSQL & "ADOPTION_ALLERGIES, "
strSQL = strSQL & "ADOPTION_VET, "
strSQL = strSQL & "ADOPTION_VET_NAME, "
strSQL = strSQL & "ADOPTION_VET_PHONE, "
strSQL = strSQL & "ADOPTION_VET_CONSIDER, "
strSQL = strSQL & "ADOPTION_OTHER_ADULTS, "
strSQL = strSQL & "ADOPTION_WHO, "
strSQL = strSQL & "ADOPTION_RESPONSIBLE, "
strSQL = strSQL & "ADOPTION_PREVIOUS, "
strSQL = strSQL & "ADOPTION_STILL_HAVE, "
strSQL = strSQL & "ADOPTION_STILL_HAVE_REASON, "
strSQL = strSQL & "ADOPTION_SURRENDER, "
strSQL = strSQL & "ADOPTION_SURRENDER_REASON, "
strSQL = strSQL & "ADOPTION_NEUTER_COST, "
strSQL = strSQL & "ADOPTION_NEUTER, "
strSQL = strSQL & "ADOPTION_NEUTER_REASON, "
strSQL = strSQL & "ADOPTION_AFFORD, "
strSQL = strSQL & "ADOPTION_RESPONSIBLE_LIFE, "
strSQL = strSQL & "ADOPTION_ADJUST, "
strSQL = strSQL & "ADOPTION_ADJUST_CONCERNS, "
strSQL = strSQL & "ADOPTION_HEAR_ABOUT, "
strSQL = strSQL & "ADOPTION_REASON, "
strSQL = strSQL & "ADOPTION_FAMILIAR, "
strSQL = strSQL & "ADOPTION_LOCATION, "
strSQL = strSQL & "ADOPTION_OUTSIDE, "
strSQL = strSQL & "ADOPTION_DOG_HOUSE, "
strSQL = strSQL & "ADOPTION_WHEN_NOT_HOME, "
strSQL = strSQL & "ADOPTION_HOUSEBREAK, "
strSQL = strSQL & "ADOPTION_HOURS_AWAY, "
strSQL = strSQL & "ADOPTION_EVER_OWNED, "
strSQL = strSQL & "ADOPTION_PROBLEM1, "
strSQL = strSQL & "ADOPTION_PROBLEM2, "
strSQL = strSQL & "ADOPTION_PROBLEM3, "
strSQL = strSQL & "ADOPTION_PROBLEM4,"
strSQL = strSQL & "ADOPTION_CAT_KEPT, "
strSQL = strSQL & "ADOPTION_CAT_OUTSIDE, "
strSQL = strSQL & "ADOPTION_CAT_REASON)"
strSQL = strSQL & "VALUES ("
strSQL = strSQL & intPersonNum & ", " & intAnimalNum & ", '" & strStatus
strSQL = strSQL & "', '" & strAgent '& "', " & intNeuterSponsor & ", '" & dteDate & "', '" & dteTime & "', " & intReceiptNum
strSQL = strSQL & "', '" & strEmployer & "', '" & strWorkPhone & "', " & intStudent & ", '" & strStudentLocation
strSQL = strSQL & "', " & intOwnRent & ", '" & strHomeType & "', '" & strLandlord & "', '" & strLandlordTelephone
strSQL = strSQL & "', " & intAdults & ", " & intChildren & ", '" & strAge & "', " & intHomeLength & ", '" & strSpouse
strSQL = strSQL & "', " & intAllergies & ", " & intVet & ", '" & strVetName & "', '" & strVetPhone & "', '" & strVetConsider
strSQL = strSQL & "', " & intOtherAdults & ", '" & strWho & "', '" & strResponsible & "', " & intPrevious & ", " & intStillHave
strSQL = strSQL & ", '" & strStillHaveReason & "', " & intSurrender & ", '" & strSurrenderReason & "', " & intNeuterCost
strSQL = strSQL & ", " & intNeuter & ", '" & strNeuterReason & "', " & intAfford & ", " & intResponsibleLife
strSQL = strSQL & ", " & intAdjust & ", '" & strAdjustConcerns & "', '" & strHearAbout & "', '" & strReason
strSQL = strSQL & "', " & intFamiliar & ", '" & strLocation & "', '" & strOutside & "', " & intHouse
strSQL = strSQL & ", '" & strWhenNotHome & "', '" & strHousebreak & "', " & intHoursAway & ", " & intEverOwned
strSQL = strSQL & ", '" & strProb1 & "', '" & strProb2 & "', '" & strProb3 & "', '" & strProb4 & "', '" & strCatKept & "', '" & strCatReason & "', '" & strCatOutside & "')"

Open_Recordsets.objConnection.Execute (strSQL)

'Updates the animal table to show that the animal is in holding

Open_Recordsets.objConnection.Execute ("UPDATE ANIMALS SET ANIMAL_STATUS = 'P' WHERE ANIMAL_NUMBER = " & intAnimalNum)

    Set rstInsert = Open_Recordsets.objConnection.Execute("SELECT ADOPTION_NUMBER FROM ADOPTION WHERE ADOPTION_ANIMAL = " & intAnimalNum & " AND ADOPTION_STATUS = 'P'")
    
    If rstInsert.EOF = False Then
    With rstInsert
        rstInsert.MoveFirst
        Do While Not rstInsert.EOF
            If Not IsNull(![ADOPTION_NUMBER]) Then
                intAdoptionNum = ![ADOPTION_NUMBER]
            End If
            rstInsert.MoveNext
        Loop
    End With
    End If

'Displays the pet form, used to get all the pets the adoptor has had in the last 5 years.

intMsgBox = MsgBox("Please list all pets the adoptor has had in the past five years.", vbYesNo, "Add pets?")
If intMsgBox = 6 Then
    frmNewPet.Show (vbModal)
End If

frmPCHS_Main.Show
Unload Me

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmListAnimals.Show
End Sub

Public Sub cmdSearch_Click()
'********************************************************************************************
'* Runs after the first and last names have been entered, searches through the person table
'* and populates the other boxes if a match is found.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-21-2002
'********************************************************************************************
Dim strFname As String          'Person's first name
Dim strLname As String          'Person's last name
Dim strAddress As String        'Person's address
Dim strCity As String           'Person's city
Dim strState As String          'Person's state
Dim strZip As String            'Person's zip code
Dim strPhone As String          'Person's phone number
Dim strEmail As String          'Person's email
Dim strLicense As String        'Person's driver's license
Dim dteDOB As Date              'Person's DOB

Dim rstSearch As New ADODB.Recordset    'Used for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset
Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM PERSON WHERE PERSON_LNAME = '" & txtLname.Text & "'")

If rstSearch.EOF <> True Then
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 4
    
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

'*****************************************************************************************************
'* Called when the form loads, enables and disables objects on the forms based on the type of animal
'* being adopted.
'*****************************************************************************************************
Private Sub Form_Load()

Dim rstInsert As ADODB.Recordset    'Used for interfacing with the database
Dim intMsgBox As Integer            'Used for messageboxes

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset
dtpDOB.Value = Date

intAnimalNum = frmListAnimals.dgdCurrentAnimals.Columns("Number").CellValue(frmListAnimals.dgdCurrentAnimals.Bookmark)
Set rstInsert = Open_Recordsets.objConnection.Execute("SELECT ANIMAL_TYPE FROM ANIMALS WHERE ANIMAL_NUMBER = " & intAnimalNum)

If rstInsert.EOF = False Then
    With rstInsert
        rstInsert.MoveFirst
        Do While Not rstInsert.EOF
            intType = ![animal_type]
            rstInsert.MoveNext
        Loop
    End With
End If

lblAnimalNum.Caption = intAnimalNum
lblCurDate = Format(Now, "dddd, MMMM dd, yyyy")

If intType = 1 Then
    cboReason.Enabled = True
    chkFamiliar.Enabled = True
    cboLocation.Enabled = True
    cboOutside.Enabled = True
    chkHouse.Enabled = True
    cboWhenNotHome.Enabled = True
    txtHousebreak.Enabled = True
    txtHoursAway.Enabled = True
    txtDogProb1.Enabled = True
    txtDogProb2.Enabled = True
    txtDogProb3.Enabled = True
    txtDogProb4.Enabled = True
ElseIf intType = 2 Then
    chkEverOwned.Enabled = True
    cboCatKept.Enabled = True
    cboCatOutside.Enabled = True
    cboCatReason.Enabled = True
    txtProb1.Enabled = True
    txtProb2.Enabled = True
    txtProb3.Enabled = True
    txtProb4.Enabled = True
End If

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

Private Sub mnuBack_Click(Index As Integer)
Unload Me
frmListAnimals.Show
End Sub

Private Sub mnuExit_Click(Index As Integer)
Open_Recordsets.objConnection.Close
End
End Sub

Private Sub optOwns_Click()
    txtLandlord.Enabled = False
    txtLandLordTelephone.Enabled = False
End Sub

Private Sub optRents_Click()
If optRents.Value = True Then
    txtLandlord.Enabled = True
    txtLandLordTelephone.Enabled = True
Else
    txtLandlord.Enabled = False
    txtLandLordTelephone.Enabled = False
End If
End Sub

Private Sub chkNeuter_Click()
If chkNeuter.Value = 1 Then
    txtNeuterReason.Enabled = False
Else
    txtNeuterReason.Enabled = True
End If
End Sub

Private Sub chkStillHave_Click()
If chkStillHave.Value = 1 Then
    txtStillHaveReason.Enabled = False
Else
    txtStillHaveReason.Enabled = True
End If
End Sub

Private Sub chkStudent_Click()
If chkStudent.Value = 1 Then
    txtStudentLocation.Enabled = True
Else
    txtStudentLocation.Enabled = False
End If
End Sub

Private Sub chkSurrender_Click()
If chkSurrender.Value = 1 Then
    txtSurrenderReason.Enabled = True
Else
    txtSurrenderReason.Enabled = False
End If
End Sub

Private Sub chkVet_Click()
If chkVet.Value = 1 Then
    txtVetName.Enabled = True
    txtVetPhone.Enabled = True
    txtVetConsider.Enabled = False
Else
    txtVetName.Enabled = False
    txtVetPhone.Enabled = False
    txtVetConsider.Enabled = True
End If
End Sub
