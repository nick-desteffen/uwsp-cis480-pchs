VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSurrender 
   Caption         =   "Surrendered animal information"
   ClientHeight    =   7440
   ClientLeft      =   270
   ClientTop       =   555
   ClientWidth     =   9585
   ControlBox      =   0   'False
   Icon            =   "frmSurrender.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleMode       =   0  'User
   ScaleWidth      =   10139.92
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7920
      TabIndex        =   49
      Top             =   6840
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   50
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Owner information"
      TabPicture(0)   =   "frmSurrender.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLicense"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDOB"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEmail"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPhone"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblZip"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblState"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCity"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblAddress"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLname"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblFname"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtLicense"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtEmail"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSearch"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPhone"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtZip"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtState"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCity"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAddress"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtLname"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtFname"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dtpDOB"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Medical Information"
      TabPicture(1)   =   "frmSurrender.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMedical"
      Tab(1).Control(1)=   "lblMedHist"
      Tab(1).Control(2)=   "lblVet"
      Tab(1).Control(3)=   "txtMedical"
      Tab(1).Control(4)=   "txtMedHist"
      Tab(1).Control(5)=   "txtVet"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Additional"
      TabPicture(2)   =   "frmSurrender.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblGames"
      Tab(2).Control(1)=   "lblFood"
      Tab(2).Control(2)=   "lblSpecialDiet"
      Tab(2).Control(3)=   "lblBehavior"
      Tab(2).Control(4)=   "txtAfriad"
      Tab(2).Control(5)=   "lblStrangers"
      Tab(2).Control(6)=   "lblOtherAnimals"
      Tab(2).Control(7)=   "lblOther"
      Tab(2).Control(8)=   "txtGames"
      Tab(2).Control(9)=   "txtFood"
      Tab(2).Control(10)=   "txtSpecialDiet"
      Tab(2).Control(11)=   "txtBehavior"
      Tab(2).Control(12)=   "txtAfraid"
      Tab(2).Control(13)=   "txtStrangers"
      Tab(2).Control(14)=   "txtOtherAnimals"
      Tab(2).Control(15)=   "txtOther"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Animal History"
      TabPicture(3)   =   "frmSurrender.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblKept"
      Tab(3).Control(1)=   "lblBiteDetails"
      Tab(3).Control(2)=   "lblReason"
      Tab(3).Control(3)=   "lblWhere"
      Tab(3).Control(4)=   "lblTime"
      Tab(3).Control(5)=   "cboWhereKept"
      Tab(3).Control(6)=   "chkLoose"
      Tab(3).Control(7)=   "chkBite"
      Tab(3).Control(8)=   "txtBiteDetails"
      Tab(3).Control(9)=   "txtReason"
      Tab(3).Control(10)=   "txtWhereGot"
      Tab(3).Control(11)=   "txtTime"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Childern"
      TabPicture(4)   =   "frmSurrender.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblAges"
      Tab(4).Control(1)=   "lblNoChildern"
      Tab(4).Control(2)=   "chkChildern"
      Tab(4).Control(3)=   "txtAges"
      Tab(4).Control(4)=   "txtNoChildern"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Cat / Dog"
      TabPicture(5)   =   "frmSurrender.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).Control(1)=   "Frame2"
      Tab(5).ControlCount=   2
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3480
         TabIndex        =   88
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48234497
         CurrentDate     =   37579
      End
      Begin VB.Frame Frame2 
         Caption         =   "Additional Cat Information"
         Height          =   4815
         Left            =   -70320
         TabIndex        =   82
         Top             =   840
         Width           =   4455
         Begin VB.CheckBox chkWindow 
            Caption         =   "Cat likes looking out windows"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   2040
            Width           =   2655
         End
         Begin VB.CheckBox chkHigh 
            Caption         =   "Cat likes high places"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkPost 
            Caption         =   "Cat uses a scratching post"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   3840
            Width           =   3975
         End
         Begin VB.CheckBox chkToilet 
            Caption         =   "Cat likes toilet water"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   3480
            Width           =   3735
         End
         Begin VB.CheckBox chkNip 
            Caption         =   "Cat likes cat nip"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   3120
            Width           =   3735
         End
         Begin VB.CheckBox chkLaps 
            Caption         =   "Cat likes sitting on laps"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   2760
            Width           =   3615
         End
         Begin VB.CheckBox chkSunny 
            Caption         =   "Cat likes sunny spots"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   2400
            Width           =   3615
         End
         Begin VB.CheckBox chkCounter 
            Caption         =   "Cat likes counter tops"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1680
            Width           =   3615
         End
         Begin VB.CheckBox chkLitter 
            Caption         =   "Cat is litter trained"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Additional Dog Information"
         Height          =   4815
         Left            =   -74880
         TabIndex        =   81
         Top             =   840
         Width           =   4455
         Begin VB.CheckBox chkPeopleFood 
            Caption         =   "Dog was fed people food"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtSensitive 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   39
            Top             =   3960
            Width           =   3855
         End
         Begin VB.TextBox txtCommands 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   38
            Top             =   3360
            Width           =   3855
         End
         Begin VB.TextBox txtGoOut 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   37
            Top             =   2760
            Width           =   3855
         End
         Begin VB.CheckBox chkBrokenTrain 
            Caption         =   "Dog has broken training before"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   1920
            Width           =   3375
         End
         Begin VB.CheckBox chkHousetrained 
            Caption         =   "Dog is housetrained"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   1680
            Width           =   3615
         End
         Begin VB.CheckBox chkNoDamage 
            Caption         =   "Dog can be trusted without causing damage"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   1440
            Width           =   3495
         End
         Begin VB.ComboBox cboDogKept 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmSurrender.frx":0972
            Left            =   2760
            List            =   "frmSurrender.frx":0982
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtHours 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   31
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblSensitive 
            Caption         =   "Sensitive areas the dog has"
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   3720
            Width           =   2895
         End
         Begin VB.Label lblCommands 
            Caption         =   "Commands the dog knows"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label lblGoOut 
            Caption         =   "How the dog notifys it wants to go out"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label lblDogAlone 
            Caption         =   "Where the dog was kept when left alone"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   84
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblAlone 
            Caption         =   "Average number of hours left alone"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   83
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.TextBox txtOther 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   5300
         Width           =   5295
      End
      Begin VB.TextBox txtNoChildern 
         Height          =   285
         Left            =   -74280
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   2180
         Width           =   5295
      End
      Begin VB.TextBox txtAges 
         Height          =   285
         Left            =   -74280
         TabIndex        =   29
         Top             =   1580
         Width           =   2655
      End
      Begin VB.CheckBox chkChildern 
         Caption         =   "Trusted with childern"
         Height          =   255
         Left            =   -74280
         TabIndex        =   28
         Top             =   980
         Width           =   2055
      End
      Begin VB.TextBox txtTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -74280
         TabIndex        =   22
         Top             =   1840
         Width           =   375
      End
      Begin VB.TextBox txtWhereGot 
         Height          =   285
         Left            =   -74280
         TabIndex        =   23
         Top             =   2440
         Width           =   3375
      End
      Begin VB.TextBox txtReason 
         Height          =   285
         Left            =   -74280
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1240
         Width           =   5295
      End
      Begin VB.TextBox txtBiteDetails 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74280
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3400
         Width           =   3375
      End
      Begin VB.CheckBox chkBite 
         Caption         =   "Ever bitten anybody?"
         Height          =   375
         Left            =   -74280
         TabIndex        =   24
         Top             =   2800
         Width           =   1935
      End
      Begin VB.CheckBox chkLoose 
         Caption         =   "Let run loose"
         Height          =   255
         Left            =   -72240
         TabIndex        =   27
         Top             =   4120
         Width           =   1335
      End
      Begin VB.ComboBox cboWhereKept 
         Height          =   315
         ItemData        =   "frmSurrender.frx":09AC
         Left            =   -74280
         List            =   "frmSurrender.frx":09B9
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4120
         Width           =   1935
      End
      Begin VB.TextBox txtOtherAnimals 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1020
         Width           =   5295
      End
      Begin VB.TextBox txtStrangers 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1620
         Width           =   5295
      End
      Begin VB.TextBox txtAfraid 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2220
         Width           =   5295
      End
      Begin VB.TextBox txtBehavior 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2820
         Width           =   5295
      End
      Begin VB.TextBox txtSpecialDiet 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4620
         Width           =   5295
      End
      Begin VB.TextBox txtFood 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   4020
         Width           =   5295
      End
      Begin VB.TextBox txtGames 
         Height          =   285
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3420
         Width           =   5295
      End
      Begin VB.TextBox txtVet 
         Height          =   285
         Left            =   -74640
         TabIndex        =   10
         Top             =   1025
         Width           =   2895
      End
      Begin VB.TextBox txtMedHist 
         Height          =   285
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2225
         Width           =   5295
      End
      Begin VB.TextBox txtMedical 
         Height          =   285
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1625
         Width           =   3855
      End
      Begin VB.TextBox txtFname 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   710
         Width           =   1575
      End
      Begin VB.TextBox txtLname 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1070
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1430
         Width           =   2895
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1790
         Width           =   1575
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   1790
         Width           =   375
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   1790
         Width           =   1215
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2150
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search People"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtLicense 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lblOther 
         Caption         =   "Any other information about the animal"
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         Top             =   5060
         Width           =   3015
      End
      Begin VB.Label lblNoChildern 
         Caption         =   "If not, why?"
         Height          =   255
         Left            =   -74280
         TabIndex        =   78
         Top             =   1940
         Width           =   975
      End
      Begin VB.Label lblAges 
         Caption         =   "If so, what ages?"
         Height          =   255
         Left            =   -74280
         TabIndex        =   77
         Top             =   1340
         Width           =   1695
      End
      Begin VB.Label lblTime 
         Caption         =   "Length of time had animal (Years)"
         Height          =   255
         Left            =   -74280
         TabIndex        =   76
         Top             =   1605
         Width           =   2415
      End
      Begin VB.Label lblWhere 
         Caption         =   "Where recieved animal from"
         Height          =   255
         Left            =   -74280
         TabIndex        =   75
         Top             =   2200
         Width           =   2175
      End
      Begin VB.Label lblReason 
         Caption         =   "Reason for giving the animal up:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   74
         Top             =   1000
         Width           =   2415
      End
      Begin VB.Label lblBiteDetails 
         Caption         =   "If so, details"
         Height          =   255
         Left            =   -74280
         TabIndex        =   73
         Top             =   3160
         Width           =   975
      End
      Begin VB.Label lblKept 
         Caption         =   "Where the animal was kept"
         Height          =   255
         Left            =   -74280
         TabIndex        =   72
         Top             =   3760
         Width           =   2175
      End
      Begin VB.Label lblOtherAnimals 
         Caption         =   "Reaction around other animals"
         Height          =   255
         Left            =   -74760
         TabIndex        =   71
         Top             =   780
         Width           =   2655
      End
      Begin VB.Label lblStrangers 
         Caption         =   "Reaction to strangers"
         Height          =   255
         Left            =   -74760
         TabIndex        =   70
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label txtAfriad 
         Caption         =   "What the animal is afraid of"
         Height          =   255
         Left            =   -74760
         TabIndex        =   69
         Top             =   1980
         Width           =   2175
      End
      Begin VB.Label lblBehavior 
         Caption         =   "Behavioral problems"
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   2580
         Width           =   3135
      End
      Begin VB.Label lblSpecialDiet 
         Caption         =   "Any Special Diets"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   4380
         Width           =   2895
      End
      Begin VB.Label lblFood 
         Caption         =   "What kind of food the animal was fed"
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   3780
         Width           =   3375
      End
      Begin VB.Label lblGames 
         Caption         =   "What games and toys the animal likes"
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   3180
         Width           =   2895
      End
      Begin VB.Label lblVet 
         Caption         =   "Name of current Veternarian"
         Height          =   255
         Left            =   -74640
         TabIndex        =   64
         Top             =   785
         Width           =   2295
      End
      Begin VB.Label lblMedHist 
         Caption         =   "Medical history of animal"
         Height          =   255
         Left            =   -74640
         TabIndex        =   63
         Top             =   1985
         Width           =   2175
      End
      Begin VB.Label lblMedical 
         Caption         =   "Any present medical problems:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   62
         Top             =   1385
         Width           =   2295
      End
      Begin VB.Label lblFname 
         Caption         =   "First Name"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   710
         Width           =   855
      End
      Begin VB.Label lblLname 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   1070
         Width           =   855
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   1430
         Width           =   735
      End
      Begin VB.Label lblCity 
         Caption         =   "City"
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   1790
         Width           =   375
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         Height          =   255
         Left            =   3000
         TabIndex        =   57
         Top             =   1790
         Width           =   495
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         Height          =   255
         Left            =   3960
         TabIndex        =   56
         Top             =   1790
         Width           =   495
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2150
         Width           =   1095
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   720
         TabIndex        =   54
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblDOB 
         Caption         =   "DOB"
         Height          =   255
         Left            =   2880
         TabIndex        =   53
         Top             =   2150
         Width           =   375
      End
      Begin VB.Label lblLicense 
         Caption         =   "Drivers License"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3000
         Width           =   1095
      End
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
      Left            =   120
      TabIndex        =   80
      Top             =   120
      Width           =   2655
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
      Left            =   2880
      TabIndex        =   51
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSurrender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************************
'* This form is used to enter information about a new surrendered animal into the
'* system.  It is called when a new animal is entered in and has been surrendered.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-21-2002
'***********************************************************************************

Public bolMatchFound As Boolean         'True = person already in database
Dim intType As Integer                  'Type of animal
Public intPersonNum As Integer          'Number of the person if found
Public intAnimal As Integer             'Number of the animal - Required

Private Sub cmdSave_Click()
'********************************************************************************************
'* Runs after the user clicks save, it saves previous owner info to the person table and
'* additional animal information to the surrender table.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-21-2002
'********************************************************************************************

'Original owner variables

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

'General variables

Dim strVet As String                'Name of the vet
Dim strPresentMed As String         'Present Medical condition
Dim strMedHist As String            'Medical history
Dim strAnimalReaction As String     'Reaction towards other animals
Dim strStrangerReaction As String   'Reaction towards strangers
Dim strAfraid As String             'What the animal is afraid of
Dim strBehaviorProb As String       'Behavioral problems the animal has
Dim strGames As String              'What games and toys the animal likes
Dim strFood As String               'Food the animal eats
Dim strDiet As String               'Any special diets the animal has
Dim strOther As String              'Any other info
Dim strReason As String             'Reason for giving animal up - Required
Dim dblTime As Double               'Length of tiem having the animal - Required
Dim strWhere As String              'Where got the animal from
Dim intBitten As Integer            'If the animal has bitten before - Required
Dim strDetails As String            'Details of the bite
Dim strKept As String               'Where the animal was kept - Required
Dim intLoose As Integer             'If the animal was let run loose - Required
Dim intTrusted As Integer           'If the animal was trusted with childern - Required
Dim strAges As String               'Ages of childern trusted with
Dim strNotTrusted As String         'Reason for not trusted witch childern

'Cat specific variables

Dim intLitter As Integer            'Cat is litter trained
Dim intHigh As Integer              'Cat likes high places
Dim intCounter As Integer           'Cat likes high places
Dim intWindow As Integer            'Cat likes looking out windows
Dim intSunny As Integer             'Cat likes sunny places
Dim intLaps As Integer              'Cat likes sitting on laps
Dim intNip As Integer               'Cat likes nip
Dim intToilet As Integer            'Cat likes toilet water
Dim intPost As Integer              'Cat uses a scratching post

'Dog specific variables

Dim intHours As Integer             'Average number of hours the dog is left alone
Dim strDogKept As String            'Where the dog is kept
Dim intDamage As Integer            'Dog can be trusted without causing damage
Dim intTrained As Integer           'Dog is housetrained
Dim intBroken As Integer            'Dog has broken housetraining
Dim intFood As Integer              'Dog eats people food
Dim strGoOut As String              'How the dog signals go out
Dim strCommands As String           'Commands the dog knows
Dim strSensitive As String          'Sensitive areas on the dog

Dim intMsgBox As Integer            'Used for messageboxes
Dim rstInsert As ADODB.Recordset    'Used for interfacing with the database
Dim strSQL As String                'SQL Statement

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset

'Checks to see if all required fields are populated
If txtFname.Text = "" Or txtLname.Text = "" Or txtAddress.Text = "" Or txtCity.Text = "" Or txtState.Text = "" Or txtZip.Text = "" Or txtPhone.Text = "" Then
    intMsgBox = MsgBox("Please fill out all the required fields!", vbOKOnly, "Error!")
    Exit Sub
End If
If txtTime.Text = "" Or cboWhereKept.Text = "" Or (chkBite.Value = 1 And txtBiteDetails.Text = "") Or (chkChildern.Value = 0 And txtNoChildern.Text = "") Then
    intMsgBox = MsgBox("Please fill out all the required fields!", vbOKOnly, "Error!")
    Exit Sub
End If

strLname = Replace(txtLname.Text, "'", "''")
strFname = Replace(txtFname.Text, "'", "''")
strAddress = Replace(txtAddress.Text, "'", "''")
strCity = Replace(txtCity.Text, "'", "''")
strState = Replace(txtState.Text, "'", "''")

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

strEmail = Replace(txtEmail.Text, "'", "''")
strLicense = Replace(txtLicense.Text, "'", "''")
dteDOB = dtpDOB.Value
strVet = Replace(txtVet.Text, "'", "''")
strPresentMed = Replace(txtMedical.Text, "'", "''")
strMedHist = Replace(txtMedHist.Text, "'", "''")
strAnimalReaction = Replace(txtOtherAnimals.Text, "'", "''")
strStrangerReaction = Replace(txtStrangers.Text, "'", "''")
strAfraid = Replace(txtAfraid.Text, "'", "''")
strBehaviorProb = Replace(txtBehavior.Text, "'", "''")
strGames = Replace(txtGames.Text, "'", "''")
strFood = Replace(txtFood.Text, "'", "''")
strDiet = Replace(txtSpecialDiet.Text, "'", "''")
strOther = Replace(txtOther.Text, "'", "''")
strReason = Replace(txtReason.Text, "'", "''")
If IsNumeric(txtTime.Text) Then
    dblTime = txtTime.Text
Else
    intMsgBox = MsgBox("Please choose a vaild number for the amount of time the animal was left alone.", vbOKOnly, "Invalid Number")
    Exit Sub
End If

strWhere = Replace(txtWhereGot.Text, "'", "''")
intBitten = chkBite.Value
strDetails = Replace(txtBiteDetails.Text, "'", "''")
strKept = cboWhereKept.Text
intLoose = chkLoose.Value
intTrusted = chkChildern.Value
strAges = Replace(txtAges.Text, "'", "''")
strNotTrusted = Replace(txtNoChildern.Text, "'", "''")

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
    End If
End If

strSQL = "SELECT PERSON_NUMBER FROM PERSON WHERE PERSON_LNAME = '" & strLname & "'"

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

strSQL = "INSERT INTO SURRENDER (SURRENDER_OWNER, "
strSQL = strSQL & "SURRENDER_ANIMAL_NUMBER, "
strSQL = strSQL & "SURRENDER_VET, "
strSQL = strSQL & "SURRENDER_REASON, "
strSQL = strSQL & "SURRENDER_MEDICAL, "
strSQL = strSQL & "SURRENDER_MEDICAL_HIST, "
strSQL = strSQL & "SURRENDER_BITTEN, "
strSQL = strSQL & "SURRENDER_BITTEN_DETAILS, "
strSQL = strSQL & "SURRENDER_WHERE_GOT, "
strSQL = strSQL & "SURRENDER_HOW_LONG, "
strSQL = strSQL & "SURRENDER_WHERE_KEPT, "
strSQL = strSQL & "SURRENDER_RUN_LOOSE, "
strSQL = strSQL & "SURRENDER_CHILDERN, "
strSQL = strSQL & "SURRENDER_CHILDERN_AGE, "
strSQL = strSQL & "SURRENDER_CHILDERN_NOT_TRUSTED, "
strSQL = strSQL & "SURRENDER_OTHER_ANIMALS, "
strSQL = strSQL & "SURRENDER_GAMES, "
strSQL = strSQL & "SURRENDER_BEHAVIOR, "
strSQL = strSQL & "SURRENDER_AFRAID, "
strSQL = strSQL & "SURRENDER_STRANGERS, "
strSQL = strSQL & "SURRENDER_FOOD, "
strSQL = strSQL & "SURRENDER_DIET, "
strSQL = strSQL & "SURRENDER_OTHER) VALUES ("
strSQL = strSQL & intPersonNum & ", " & intAnimal & ", '" & strVet & "', '" & strReason & "', '" & strPresentMed
strSQL = strSQL & "', '" & strMedHist & "', " & intBitten & ", '" & strDetails & "', '" & strWhere & "', "
strSQL = strSQL & dblTime & ", '" & strKept & "', " & intLoose & ", " & intTrusted & ", '" & strAges & "', '" & strNotTrusted & "', '"
strSQL = strSQL & strAnimalReaction & "', '" & strGames & "', '" & strBehaviorProb & "', '" & strAfraid & "', '"
strSQL = strSQL & strStrangerReaction & "', '" & strFood & "', '" & strDiet & "', '" & strOther & "')"

Open_Recordsets.objConnection.Execute (strSQL)

If intType = 1 Then
    intHours = CInt(txtHours.Text)
    strDogKept = cboDogKept.Text
    intDamage = chkNoDamage.Value
    intTrained = chkHousetrained.Value
    intBroken = chkBrokenTrain.Value
    intFood = chkPeopleFood.Value
    strGoOut = Replace(txtGoOut.Text, "'", "''")
    strCommands = Replace(txtCommands.Text, "'", "''")
    strSensitive = Replace(txtSensitive.Text, "'", "''")
    
    strSQL = "INSERT INTO SURRENDER_DOG (SDOG_NUMBER, "
    strSQL = strSQL & "SDOG_HOURS_ALONE, "
    strSQL = strSQL & "SDOG_WHERE_KEPT, "
    strSQL = strSQL & "SDOG_NO_DAMAGE, "
    strSQL = strSQL & "SDOG_HOUSETRAINED, "
    strSQL = strSQL & "SDOG_BROKEN_TRAIN, "
    strSQL = strSQL & "SDOG_GO_OUT, "
    strSQL = strSQL & "SDOG_COMMANDS, "
    strSQL = strSQL & "SDOG_SENSITIVE, "
    strSQL = strSQL & "SDOG_PEOPLE_FOOD) VALUES ("
    strSQL = strSQL & intAnimal & ", " & intHours & ", '" & strDogKept & "', " & intDamage & ", " & intTrained & ", " & intBroken & ", '" & strGoOut & "', '"
    strSQL = strSQL & strCommands & "', '" & strSensitive & "', " & intFood & ")"

    Open_Recordsets.objConnection.Execute (strSQL)

ElseIf intType = 2 Then
    
    intLitter = chkLitter.Value
    intHigh = chkHigh.Value
    intCounter = chkCounter.Value
    intWindow = chkWindow.Value
    intSunny = chkSunny.Value
    intLaps = chkLaps.Value
    intNip = chkNip.Value
    intToilet = chkToilet.Value
    intPost = chkPost.Value
    
    strSQL = "INSERT INTO SURRENDER_CAT (SCAT_NUMBER, "
    strSQL = strSQL & "SCAT_LITTER, "
    strSQL = strSQL & "SCAT_HIGH, "
    strSQL = strSQL & "SCAT_COUNTER, "
    strSQL = strSQL & "SCAT_WINDOWS, "
    strSQL = strSQL & "SCAT_SUNNY, "
    strSQL = strSQL & "SCAT_LAPS, "
    strSQL = strSQL & "SCAT_NIP, "
    strSQL = strSQL & "SCAT_TOILET, "
    strSQL = strSQL & "SCAT_POST) VALUES ("
    strSQL = strSQL & intAnimal & ", " & intLitter & ", " & intHigh & ", " & intCounter & ", " & intWindow & ", " & intSunny & ", " & intLaps & ", " & intNip & ", " & intToilet & ", " & intPost & ")"
    
    Open_Recordsets.objConnection.Execute (strSQL)
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

Public Sub cmdSearch_Click()
'********************************************************************************************
'* Runs after the first and last names have been entered, searches through the person table
'* and populates the other boxes if a match is found.
'*
'* Written by: Nick DeSteffen
'* Written on: 10-21-2002
'********************************************************************************************
Dim strFname As String          'First name of person
Dim strLname As String          'Last name of person
Dim strAddress As String        'Address of person
Dim strCity As String           'City of person
Dim strState As String          'State of person
Dim strZip As String            'Zip code of person
Dim strPhone As String          'Telephone number of person
Dim strEmail As String          'Email address of person
Dim strLicense As String        'Driver's license of person
Dim dteDOB As Date              'DOB of person

Dim rstSearch As New ADODB.Recordset    'Used for interfacing with the databae
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstSearch = New ADODB.Recordset
Set rstSearch = Open_Recordsets.objConnection.Execute("SELECT * FROM PERSON WHERE PERSON_FNAME = '" & txtFname.Text & "' AND PERSON_LNAME = '" & txtLname.Text & "'")

If rstSearch.EOF <> True Then
    
    frmPeople.strFname = txtFname.Text
    frmPeople.strLname = txtLname.Text
    frmPeople.intType = 3
    
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

Private Sub Form_Load()

Dim rstInsert As ADODB.Recordset        'Used for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstInsert = New ADODB.Recordset
intAnimal = frmNewAnimal.intNum
lblAnimalNum.Caption = intAnimal
dtpDOB.Value = Date

Set rstInsert = Open_Recordsets.objConnection.Execute("SELECT ANIMAL_TYPE FROM ANIMALS WHERE ANIMAL_NUMBER = " & intAnimal)

With rstInsert
    .MoveFirst
    intType = ![animal_type]
End With

If intType = 1 Then
    txtHours.Enabled = True
    cboDogKept.Enabled = True
    chkNoDamage.Enabled = True
    chkHousetrained.Enabled = True
    chkBrokenTrain.Enabled = True
    chkPeopleFood.Enabled = True
    txtGoOut.Enabled = True
    txtCommands.Enabled = True
    txtSensitive.Enabled = True
ElseIf intType = 2 Then
    chkLitter.Enabled = True
    chkHigh.Enabled = True
    chkCounter.Enabled = True
    chkWindow.Enabled = True
    chkSunny.Enabled = True
    chkLaps.Enabled = True
    chkNip.Enabled = True
    chkToilet.Enabled = True
    chkPost.Enabled = True
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub chkBite_Click()
If chkBite.Value = 1 Then
    txtBiteDetails.Enabled = True
ElseIf chkBite.Value = 0 Then
    txtBiteDetails.Enabled = False
End If
End Sub

