VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListAnimals 
   Caption         =   "Animals Currently Residing at PCHS"
   ClientHeight    =   7350
   ClientLeft      =   225
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmListAnimals.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10995
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print List"
      Height          =   615
      Left            =   6000
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dgdCurrentAnimals 
      Height          =   4095
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "animal_number"
         Caption         =   "Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "animal_name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "type_name"
         Caption         =   "Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "BREED_NAME"
         Caption         =   "Breed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "animal_sex"
         Caption         =   "Sex"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "color_name"
         Caption         =   "Color"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "animal_age"
         Caption         =   "Age"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "dte"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "donation_amount"
         Caption         =   "Donation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   959.811
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdReclaim 
      Caption         =   "Reclaim Animal"
      Height          =   615
      Left            =   7560
      TabIndex        =   4
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdopt 
      Caption         =   "Adopt Animal"
      Height          =   615
      Left            =   9120
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Animal"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Caption         =   "Animals currently residing at the Portage County Humane Society"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   10575
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
      Index           =   4
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmListAnimals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************************
'* This form displays all the animals currently at the humane society.  From this form users
'* can edit an animal, reclaim an animal, adopt an animal, or generate a detailed printout of
'* all the animals
'*
'*******************************************************************************************

Private Sub cmdAdopt_Click()
frmNewAdoption.Show
frmListAnimals.Hide
End Sub

Private Sub cmdBack_Click()
Unload Me
frmPCHS_Main.Show
End Sub

Private Sub cmdEdit_Click()
frmEditAnimal.Show
Unload Me
End Sub

Private Sub cmdPrint_Click()
frmShowAnimals.Show
End Sub

Private Sub cmdReclaim_Click()
frmReclaim.Show
Unload Me
End Sub

Private Sub Form_Load()
frmPCHS_Main.Hide
Call Open_Recordsets.Open_Animals
End Sub

Private Sub mnuAbout_Click()
Call About.About
End Sub

Private Sub mnuBack_Click(Index As Integer)
Unload Me
frmPCHS_Main.Show
End Sub

Private Sub mnuExit_Click(Index As Integer)
Open_Recordsets.objConnection.Close
End
End Sub
