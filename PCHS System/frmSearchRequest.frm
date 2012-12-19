VERSION 5.00
Begin VB.Form frmSearchRequest 
   Caption         =   "Search Animal"
   ClientHeight    =   3615
   ClientLeft      =   3585
   ClientTop       =   3975
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   7410
   Begin VB.ComboBox cboAnimalType 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.ComboBox cboBreed 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblType 
      Caption         =   "Type of Animal"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblBreed 
      Caption         =   "Breed of Animal"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblColor 
      Caption         =   "Color of Animal"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu mnuBack 
         Caption         =   "Back"
         Index           =   2
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmSearchRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************
'* Array that contains all the types of animals
'*
'* Written by: Nick DeSteffen and Kalonji Kadima
'* Written on: 10-09-2002
'********************************************************************************************
Private Type combo_info
    name As String
    Number As Integer
End Type

Private Sub cboAnimalType_Click()

End Sub

Private Sub cboBreed_Change()

End Sub

Private Sub cmdCancel_Click()
Unload Me
frmRequest.Show
End Sub

Private Sub cmdSearch_Click()

Dim rstRequest As ADODB.Recordset
Dim objConnection As ADODB.Connection

Dim intRequest As Integer
Dim strAnimal As String

Set objConnection = New ADODB.Connection
Set rstRequest = New ADODB.Recordset

objConnection.ConnectionString = frmPCHS_Main.strConnectionString
objConnection.Open

strAnimal = txtAnimal.Text


intRequest = Search_Requests.Search_Requests(intType, intBreed, intColor)

rstRequest.Close

End Sub

Private Sub mnuBack_Click(Index As Integer)
frmSearchRequest.Hide
frmRequest.Show
End Sub

Private Sub Form_Load()

Dim types() As combo_info
Dim colors() As combo_info
Dim breeds() As combo_info

End Sub


