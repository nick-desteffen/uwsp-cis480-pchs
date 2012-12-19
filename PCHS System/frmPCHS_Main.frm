VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPCHS_Main 
   Caption         =   "Portage County Humane Society System"
   ClientHeight    =   7095
   ClientLeft      =   345
   ClientTop       =   630
   ClientWidth     =   8145
   Icon            =   "frmPCHS_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   8145
   Begin VB.CommandButton cmdNewDonation 
      Caption         =   "New Donation"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdNewRequest 
      Caption         =   "New Request"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdNewAnimal 
      Caption         =   "New Animal"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdNewMissing 
      Caption         =   "New Missing"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdListAnimals 
      Caption         =   "List Current Animals"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdListAdoptions 
      Caption         =   "List Active Adoptions"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   4320
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog dlgLocation 
      Left            =   2520
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   5880
      Width           =   1815
   End
   Begin VB.PictureBox picDog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   3240
      Picture         =   "frmPCHS_Main.frx":08CA
      ScaleHeight     =   6975
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAnimal 
      Caption         =   "Animals"
      Begin VB.Menu mnuNewAnimal 
         Caption         =   "New Animal"
      End
      Begin VB.Menu mnuListAnimals 
         Caption         =   "List Currently Residing Animals"
      End
   End
   Begin VB.Menu mnuRequest 
      Caption         =   "Requests"
      Begin VB.Menu mnuNewRequest 
         Caption         =   "New Animal Request"
      End
      Begin VB.Menu mnuListRequest 
         Caption         =   "List Requested Animals"
      End
   End
   Begin VB.Menu mnuMissing 
      Caption         =   "Missing"
      Begin VB.Menu mnuNewMissing 
         Caption         =   "New Missing Animal"
      End
      Begin VB.Menu mnuListMissing 
         Caption         =   "List Missing Animals"
      End
   End
   Begin VB.Menu mnuComplaints 
      Caption         =   "Complaints"
      Begin VB.Menu mnuNewComplaints 
         Caption         =   "New Complaint"
      End
      Begin VB.Menu mnuBiteComplaint 
         Caption         =   "New Bite Report"
      End
      Begin VB.Menu mnuUpdateBite 
         Caption         =   "Update Bite Report"
      End
   End
   Begin VB.Menu mnuDonations 
      Caption         =   "Donations"
      Begin VB.Menu mnuNewDonation 
         Caption         =   "New Donation"
      End
   End
   Begin VB.Menu mnuLicense 
      Caption         =   "License"
      Begin VB.Menu mnuNewLicense 
         Caption         =   "New License"
      End
      Begin VB.Menu mnuSearchLicense 
         Caption         =   "Search Licenses"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuDetailAcquisition 
         Caption         =   "Detailed Acquisition Report"
      End
      Begin VB.Menu mnuDetailDisposition 
         Caption         =   "Detailed Disposition Report"
      End
      Begin VB.Menu mnuSumAcquisition 
         Caption         =   "Summarized Acquisition Report"
         Begin VB.Menu mnuSumDog 
            Caption         =   "Dogs"
         End
         Begin VB.Menu mnuSumCats 
            Caption         =   "Cats"
         End
         Begin VB.Menu mnuSumOther 
            Caption         =   "Other Animals"
         End
      End
      Begin VB.Menu mnuSumDisposition 
         Caption         =   "Summarized Disposition Report"
         Begin VB.Menu mnuDogsOut 
            Caption         =   "Dogs"
         End
         Begin VB.Menu mnuCatsOut 
            Caption         =   "Cats"
         End
         Begin VB.Menu mnuOtherOut 
            Caption         =   "Other Animals"
         End
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Other"
      Begin VB.Menu mnuLocation 
         Caption         =   "Database Location"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search Database"
      End
      Begin VB.Menu mnuBadPerson 
         Caption         =   "New Bad Person"
      End
      Begin VB.Menu mnuPurge 
         Caption         =   "Purge Records"
      End
      Begin VB.Menu mnuReciept 
         Caption         =   "Generate Reciept"
      End
      Begin VB.Menu mnuLabel 
         Caption         =   "Print Mailing Labels"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmPCHS_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
'* This is the main form.  When it runs it looks at the file PCHS.INI which is
'* located in the same location as PCHS.EXE for the location of the database file
'* PCHS.MDB.  It then sets the ADO connection string.  The following forms can
'* be called from this form.
'*
'* New Animal
'* List Animals
'* New Missing
'* List Missing
'* New Request
'* List Reuests
'* New Animal Complaint
'* New Bite report
'* Update Bite Report
'* New Donation
'* New License
'* Search Licenses
'* All the reports
'* New Bad Person
'* Search Database
'* New general receipt
'* Purge Records
'* Print Mailing labels
'* Database Location
'*******************************************************************************

    
    Dim RtnStr As String    'value read from INI file
    Dim sRetVal As Long     'length of returned string
    Dim strPath As String   'Path of the database

    Dim SourceFolder As String
    Dim objFSO As Object
    Dim ObjFile As Variant
    Dim ObjFolder As Variant
    
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public strConnectionString As String

Private Sub cmdExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub

Private Sub cmdListAnimals_Click()
frmPCHS_Main.Hide
frmListAnimals.Show
End Sub

Private Sub cmdNewAnimal_Click()
frmNewAnimal.Show
frmPCHS_Main.Hide
End Sub

Private Sub cmdNewDonation_Click()
frmNewDonation.Show
frmPCHS_Main.Hide
End Sub

Private Sub cmdNewMissing_Click()
frmNewMissing.Show
frmPCHS_Main.Hide
End Sub

Private Sub cmdNewRequest_Click()
frmNewRequest.Show
frmPCHS_Main.Hide
End Sub

Private Sub cmdListAdoptions_Click()
frmActiveAdoptions.Show
End Sub

Private Sub Form_Load()
Call ReadFiles
frmPCHS_Main.strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strPath
Call Open_Recordsets.Open_Conn
End Sub

Private Sub mnuAbout_Click()
Call About.About
End Sub

Private Sub mnuBadPerson_Click()
frmBadPerson.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuBiteComplaint_Click()
frmNewBiteComplaint.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuCatsOut_Click()
frmCatOut.Show
End Sub

Private Sub mnuDetailAcquisition_Click()
frmDetailedAcquisition.Show
End Sub

Private Sub mnuDetailDisposition_Click()
frmDetailedDisposition.Show
End Sub

Private Sub mnuDogsOut_Click()
frmDogOut.Show
End Sub

Private Sub mnuExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub

Private Sub mnuLabel_Click()
frmLabels.Show
End Sub

Private Sub mnuListAnimals_Click()
frmListAnimals.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuListMissing_Click()
frmListMissing.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuListRequest_Click()
frmListRequest.Show
frmPCHS_Main.Hide
End Sub

Public Sub mnuLocation_Click()
dlgLocation.ShowOpen
If dlgLocation.FileName <> "" Then
    strPath = dlgLocation.FileName
    sRetVal = WritePrivateProfileString("Database", "Path", strPath, App.Path & "\PCHS.ini")
    Call ReadFiles
End If
frmPCHS_Main.strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strPath
End Sub

Sub ReadFromINI()
   On Error Resume Next
    ObjFile.name = App.Path & "\PCHS.ini"
    sRetVal = GetPrivateProfileString("Database", "Path", "0", RtnStr, 255, SourceFolder & "\" & ObjFile.name)
    strPath = Left$(RtnStr, sRetVal)  ' extract the returned string from the buffer
  End Sub

Sub ReadFiles()
 On Error GoTo ErrorHandler   ' Enable error-handling routine.
   RtnStr = Space(255)
   Set objFSO = CreateObject("Scripting.FileSystemObject")
       SourceFolder = App.Path
       If objFSO.FolderExists(SourceFolder) Then
           Set ObjFolder = objFSO.GetFolder(SourceFolder)
           For Each ObjFile In ObjFolder.Files
             If InStr(ObjFile, ".ini") > 0 Then
                Call ReadFromINI '// Ok
             End If
           Next
        End If
        Set ObjFolder = Nothing
        Set ObjFile = Nothing
        Set objFSO = Nothing
 Exit Sub           ' Exit to avoid handler.
ErrorHandler:       ' Error-handling routine.
    Set ObjFolder = Nothing
    Set ObjFile = Nothing
    Set objFSO = Nothing
    End
End Sub

Private Sub mnuNewAnimal_Click()
frmNewAnimal.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuNewComplaints_Click()
frmNewComplaint.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuNewDonation_Click()
frmNewDonation.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuNewLicense_Click()
frmLicense.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuNewMissing_Click()
frmNewMissing.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuNewRequest_Click()
frmNewRequest.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuOtherOut_Click()
frmOtherOut.Show
End Sub

Private Sub mnuPurge_Click()
frmDelRecords.Show
End Sub

Private Sub mnuReciept_Click()
frmMiscReceipt.Show
End Sub

Private Sub mnuSearch_Click()
frmCompleteSearch.Show
frmPCHS_Main.Hide
End Sub

Private Sub mnuSearchLicense_Click()
frmSearchLicense.Show
End Sub

Private Sub mnuSumCats_Click()
frmCatsIn.Show
End Sub

Private Sub mnuSumDog_Click()
frmDogsIn.Show
End Sub

Private Sub mnuSumOther_Click()
frmOtherIn.Show
End Sub

Private Sub mnuUpdateBite_Click()
frmUpdateBite.Show
End Sub
