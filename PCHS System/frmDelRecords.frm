VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDelRecords 
   Caption         =   "Delete Records"
   ClientHeight    =   2925
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4785
   Icon            =   "frmDelRecords.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelDonations 
      Caption         =   "Delete Donations"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelReciepts 
      Caption         =   "Delete Reciepts"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelAdoptions 
      Caption         =   "Delete Adoptions"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelAnimals 
      Caption         =   "Delete Animals"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19791873
      CurrentDate     =   37585
   End
   Begin VB.Label Label1 
      Caption         =   "Delete Records older than:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   3375
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
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmDelRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************
'* This form is used to purge old records.  The user selects the date to purge
'* records older than.  They then click the button to purge the records they
'* want to purge.
'******************************************************************************

'************* Deletes the adoptions ******************************************
Private Sub cmdDelAdoptions_Click()
Dim intMsgBox As Integer        'Used for messageboxes
Dim strSQL As String            'SQL Statement

On Error GoTo ErrorHandler

intMsgBox = MsgBox("Are you sure you want to delete the adoption records?" & Chr(13) & "                THIS CANNOT BE UNDONE!", vbYesNo, "WARNING")
If intMsgBox = 6 Then
    strSQL = "DELETE FROM ADOPTION WHERE ADOPTION_DATE_START <= #" & dtpDate.Value & "# AND (ADOPTION_STATUS <> 'H' OR ADOPTION_STATUS <> 'A')"
    Open_Recordsets.objConnection.Execute (strSQL)
Else
    Exit Sub
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

'************* Deletes the animals ******************************************
Private Sub cmdDelAnimals_Click()
Dim intMsgBox As Integer        'Used for messageboxes
Dim strSQL As String            'SQL Statement

On Error GoTo ErrorHandler

intMsgBox = MsgBox("Are you sure you want to delete the animal records?" & Chr(13) & "               THIS CANNOT BE UNDONE!", vbYesNo, "WARNING")
If intMsgBox = 6 Then
    strSQL = "DELETE FROM ANIMALS WHERE ANIMAL_DATE <= #" & dtpDate.Value & "# AND (ANIMAL_STATUS <> 'R' OR ANIMAL_STATUS <> 'P')"
    Open_Recordsets.objConnection.Execute (strSQL)
Else
    Exit Sub
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

'************* Deletes the donations ******************************************
Private Sub cmdDelDonations_Click()
Dim intMsgBox As Integer        'Used for messageboxes
Dim strSQL As String            'SQL Statement

On Error GoTo ErrorHandler

intMsgBox = MsgBox("Are you sure you want to delete the donation records?" & Chr(13) & "                THIS CANNOT BE UNDONE!", vbYesNo, "WARNING")
If intMsgBox = 6 Then
    strSQL = "DELETE FROM DONATION WHERE DONATION_DATE <= #" & dtpDate.Value & "#"
    Open_Recordsets.objConnection.Execute (strSQL)
Else
    Exit Sub
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

'************* Deletes the receipts ******************************************
Private Sub cmdDelReciepts_Click()
Dim intMsgBox As Integer        'Used for messageboxes
Dim strSQL As String            'SQL Statement

On Error GoTo ErrorHandler
intMsgBox = MsgBox("Are you sure you want to delete the reciept records?" & Chr(13) & "                THIS CANNOT BE UNDONE!", vbYesNo, "WARNING")
If intMsgBox = 6 Then
    strSQL = "DELETE FROM RECIEPT WHERE RECIEPT_DATE <= #" & dtpDate.Value & "#"
    Open_Recordsets.objConnection.Execute (strSQL)
Else
    Exit Sub
End If

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End
End Sub

Private Sub Form_Load()
dtpDate.Value = Date
End Sub

Private Sub mnuAbout_Click()
Call About.About
End Sub

Private Sub mnuBack_Click()
frmPCHS_Main.Show
Unload Me
End Sub

Private Sub mnuExit_Click()
Open_Recordsets.objConnection.Close
End
End Sub
