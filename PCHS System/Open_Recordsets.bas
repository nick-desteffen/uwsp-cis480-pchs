Attribute VB_Name = "Open_Recordsets"
Option Explicit
'*************************************************************************************
'* This module is used to open all the recoredsets used to display information in data
'* grids.  When a form with a datagrid is opened the appropriate sub is called and the
'* datagrid is populated.  This module also contains the connection string to the
'* database.
'*
'* Written by: Nick DeSteffen
'* Written on: 12-04-2002
'*************************************************************************************

Public objConnection As ADODB.Connection        'Connection string used throughout program
'*****************************************************************************
'* This sub opens the connection to the database.
'*****************************************************************************
Public Sub Open_Conn()

On Error GoTo ErrorHandler

Dim intMsgBox As Integer        'Used for messageboxes

Set objConnection = New ADODB.Connection
objConnection.ConnectionString = frmPCHS_Main.strConnectionString
objConnection.Open
objConnection.CursorLocation = adUseClient
Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
        intMsgBox = MsgBox("The path to the database is invalid." & Chr(13) & "Please change it.", vbOKOnly, "Invalid path")
        Call frmPCHS_Main.mnuLocation_Click
        objConnection.ConnectionString = frmPCHS_Main.strConnectionString
        Resume
    Else
        intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
        End
    End If
    
End Sub

'*****************************************************************************************
'* Gets all the requests and displays them.
'*****************************************************************************************
Public Sub Open_Requests()

Dim strRequestSQL As String             'SQL Statement
Dim rstRequest As ADODB.Recordset       'Recordset used for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstRequest = New ADODB.Recordset

strRequestSQL = "SELECT Requests.request_number,"
strRequestSQL = strRequestSQL & " Person.person_fname,"
strRequestSQL = strRequestSQL & " Person.person_lname,"
strRequestSQL = strRequestSQL & " Animal_Types.type_name,"
strRequestSQL = strRequestSQL & " Requests.request_sex,"
strRequestSQL = strRequestSQL & " Requests.request_age,"
strRequestSQL = strRequestSQL & " Color.color_name,"
strRequestSQL = strRequestSQL & " Breeds.BREED_NAME,"
strRequestSQL = strRequestSQL & " Format(Requests.request_date, 'mm/dd/yyyy') AS request_date"
strRequestSQL = strRequestSQL & " FROM Person INNER JOIN"
strRequestSQL = strRequestSQL & " ((Color RIGHT JOIN (Animal_Types"
strRequestSQL = strRequestSQL & " INNER JOIN Requests "
strRequestSQL = strRequestSQL & " ON Animal_Types.type_number = Requests.request_type)"
strRequestSQL = strRequestSQL & " ON Color.color_number = Requests.request_color)"
strRequestSQL = strRequestSQL & " LEFT JOIN Breeds ON (Requests.request_breed = Breeds.BREED_NUMBER)"
strRequestSQL = strRequestSQL & " AND (Requests.request_type = Breeds.TYPE))"
strRequestSQL = strRequestSQL & " ON Person.person_number = Requests.request_person;"

rstRequest.Open strRequestSQL, objConnection, adOpenKeyset, adLockOptimistic, adCmdText
Set frmListRequest.dbgCurrentRequests.DataSource = rstRequest

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Sub

'*****************************************************************************************
'* Gets all the missing animals and displays them.
'*****************************************************************************************
Public Sub Open_Missing()

Dim strMissingSQL As String             'SQL Statement
Dim rstMissing As ADODB.Recordset       'Recordset used for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstMissing = New ADODB.Recordset

strMissingSQL = "SELECT MISSING.MISSING_NUMBER,"
strMissingSQL = strMissingSQL & " FORMAT(MISSING.MISSING_DATE, 'MM/DD/YYYY') AS DTE,"
strMissingSQL = strMissingSQL & " PERSON.PERSON_FNAME,"
strMissingSQL = strMissingSQL & " PERSON.PERSON_LNAME AS LNME,"
strMissingSQL = strMissingSQL & " ANIMAL_TYPES.TYPE_NAME,"
strMissingSQL = strMissingSQL & " MISSING.MISSING_SEX,"
strMissingSQL = strMissingSQL & " MISSING.MISSING_AGE,"
strMissingSQL = strMissingSQL & " COLOR.COLOR_NAME,"
strMissingSQL = strMissingSQL & " DOG_BREEDS.BREED_NAME"
strMissingSQL = strMissingSQL & " FROM PERSON,"
strMissingSQL = strMissingSQL & " MISSING,"
strMissingSQL = strMissingSQL & " ANIMAL_TYPES,"
strMissingSQL = strMissingSQL & " COLOR,"
strMissingSQL = strMissingSQL & " DOG_BREEDS"
strMissingSQL = strMissingSQL & " Where PERSON.PERSON_NUMBER = MISSING.MISSING_PERSON"
strMissingSQL = strMissingSQL & " AND ANIMAL_TYPES.TYPE_NUMBER = MISSING.MISSING_TYPE"
strMissingSQL = strMissingSQL & " AND COLOR.COLOR_NUMBER = MISSING.MISSING_COLOR"
strMissingSQL = strMissingSQL & " AND MISSING.MISSING_TYPE = 1"
strMissingSQL = strMissingSQL & " AND MISSING.MISSING_BREED = DOG_BREEDS.BREED_NUMBER"
strMissingSQL = strMissingSQL & " Union"
strMissingSQL = strMissingSQL & " SELECT MISSING.MISSING_NUMBER,"
strMissingSQL = strMissingSQL & " FORMAT(MISSING.MISSING_DATE, 'MM/DD/YYYY') AS DTE,"
strMissingSQL = strMissingSQL & " PERSON.PERSON_FNAME,"
strMissingSQL = strMissingSQL & " PERSON.PERSON_LNAME AS LNME,"
strMissingSQL = strMissingSQL & " ANIMAL_TYPES.TYPE_NAME,"
strMissingSQL = strMissingSQL & " MISSING.MISSING_SEX,"
strMissingSQL = strMissingSQL & " MISSING.MISSING_AGE,"
strMissingSQL = strMissingSQL & " COLOR.COLOR_NAME,"
strMissingSQL = strMissingSQL & " CAT_BREEDS.BREED_NAME"
strMissingSQL = strMissingSQL & " FROM PERSON,"
strMissingSQL = strMissingSQL & " MISSING,"
strMissingSQL = strMissingSQL & " ANIMAL_TYPES,"
strMissingSQL = strMissingSQL & " COLOR,"
strMissingSQL = strMissingSQL & " CAT_BREEDS"
strMissingSQL = strMissingSQL & " Where PERSON.PERSON_NUMBER = MISSING.MISSING_PERSON"
strMissingSQL = strMissingSQL & " AND ANIMAL_TYPES.TYPE_NUMBER = MISSING.MISSING_TYPE"
strMissingSQL = strMissingSQL & " AND COLOR.COLOR_NUMBER = MISSING.MISSING_COLOR"
strMissingSQL = strMissingSQL & " AND MISSING.MISSING_TYPE = 2"
strMissingSQL = strMissingSQL & " AND MISSING.MISSING_BREED = CAT_BREEDS.BREED_NUMBER"
strMissingSQL = strMissingSQL & " UNION SELECT  MISSING.MISSING_NUMBER,"
strMissingSQL = strMissingSQL & " FORMAT(MISSING.MISSING_DATE, 'MM/DD/YYYY') AS DTE,"
strMissingSQL = strMissingSQL & " PERSON.PERSON_FNAME,"
strMissingSQL = strMissingSQL & " PERSON.PERSON_LNAME AS LNME,"
strMissingSQL = strMissingSQL & " ANIMAL_TYPES.TYPE_NAME,"
strMissingSQL = strMissingSQL & " MISSING.MISSING_SEX,"
strMissingSQL = strMissingSQL & " MISSING.MISSING_AGE,"
strMissingSQL = strMissingSQL & " COLOR.COLOR_NAME,"
strMissingSQL = strMissingSQL & " IIF(MISSING.MISSING_BREED, 0, 'None')"
strMissingSQL = strMissingSQL & " FROM PERSON,"
strMissingSQL = strMissingSQL & " MISSING,"
strMissingSQL = strMissingSQL & " ANIMAL_TYPES,"
strMissingSQL = strMissingSQL & " Color"
strMissingSQL = strMissingSQL & " Where PERSON.PERSON_NUMBER = MISSING.MISSING_PERSON"
strMissingSQL = strMissingSQL & " AND ANIMAL_TYPES.TYPE_NUMBER = MISSING.MISSING_TYPE"
strMissingSQL = strMissingSQL & " AND COLOR.COLOR_NUMBER = MISSING.MISSING_COLOR"
strMissingSQL = strMissingSQL & " AND MISSING.MISSING_TYPE <> 1"
strMissingSQL = strMissingSQL & " AND MISSING.MISSING_TYPE <> 2"
strMissingSQL = strMissingSQL & " ORDER BY DTE, MISSING_NUMBER, LNME;"

rstMissing.Open strMissingSQL, objConnection, adOpenKeyset, adLockOptimistic, adCmdText
Set frmListMissing.dbgCurrentMissing.DataSource = rstMissing

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Sub

'*****************************************************************************************
'* Gets all the active animals and displays them.
'*****************************************************************************************
Public Sub Open_Animals()

Dim strAnimalSQL As String              'SQL Statement
Dim rstAnimals As ADODB.Recordset       'Recordset used for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstAnimals = New ADODB.Recordset

strAnimalSQL = "SELECT Animals.animal_number,"
strAnimalSQL = strAnimalSQL & " Animals.animal_name,"
strAnimalSQL = strAnimalSQL & " Animal_Types.type_name,"
strAnimalSQL = strAnimalSQL & " Breeds.BREED_NAME,"
strAnimalSQL = strAnimalSQL & " Animals.animal_sex,"
strAnimalSQL = strAnimalSQL & " Color.color_name,"
strAnimalSQL = strAnimalSQL & " Animals.animal_age,"
strAnimalSQL = strAnimalSQL & " Format(Animals.animal_date,'mm/dd/yyyy') AS dte,"
strAnimalSQL = strAnimalSQL & " Donation.donation_amount"
strAnimalSQL = strAnimalSQL & " FROM Donation RIGHT JOIN"
strAnimalSQL = strAnimalSQL & " (Breeds RIGHT JOIN"
strAnimalSQL = strAnimalSQL & " (Color INNER JOIN"
strAnimalSQL = strAnimalSQL & " (Animal_Types INNER JOIN Animals"
strAnimalSQL = strAnimalSQL & " ON Animal_Types.type_number = Animals.animal_type)"
strAnimalSQL = strAnimalSQL & " ON Color.color_number = Animals.animal_color)"
strAnimalSQL = strAnimalSQL & " ON (Breeds.BREED_NUMBER = Animals.animal_breed)"
strAnimalSQL = strAnimalSQL & " AND (Breeds.TYPE = Animals.animal_type))"
strAnimalSQL = strAnimalSQL & " ON Donation.donation_number = Animals.animal_neuter_sponsor"
strAnimalSQL = strAnimalSQL & " WHERE (((Animals.animal_status)='R'))"
strAnimalSQL = strAnimalSQL & " ORDER BY Animals.animal_number;"

rstAnimals.Open strAnimalSQL, objConnection, adOpenKeyset, adLockOptimistic, adCmdText
Set frmListAnimals.dgdCurrentAnimals.DataSource = rstAnimals

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End


End Sub

'*****************************************************************************************
'* Gets all the active animals and displays them.
'*****************************************************************************************
Public Sub Open_Adoptions()

Dim strAdoptionSQL As String            'SQL Statement
Dim rstAdoptions As ADODB.Recordset     'Recordset used for interfacing with the database
Dim intMsgBox As Integer                'Used for messageboxes

On Error GoTo ErrorHandler

Set rstAdoptions = New ADODB.Recordset

strAdoptionSQL = "SELECT Adoption.adoption_number,"
strAdoptionSQL = strAdoptionSQL & " Person.person_fname,"
strAdoptionSQL = strAdoptionSQL & " Person.person_telephone,"
strAdoptionSQL = strAdoptionSQL & " Format([Adoption].[adoption_date_start],'mm/dd/yyyy') AS Expr1,"
strAdoptionSQL = strAdoptionSQL & " Switch([Adoption].[adoption_status]='P','Pending Verification',[Adoption].[adoption_status]='A','Aproved',[Adoption].[adoption_status]='D','Declined',[Adoption].[adoption_status]='C','Completed',[Adoption].[adoption_status]='H','Checkup') AS Expr2,"
strAdoptionSQL = strAdoptionSQL & " Animal_Types.type_name,"
strAdoptionSQL = strAdoptionSQL & " Person.person_lname,"
strAdoptionSQL = strAdoptionSQL & " Animals.animal_number"
strAdoptionSQL = strAdoptionSQL & " FROM Person INNER JOIN"
strAdoptionSQL = strAdoptionSQL & " ((Animal_Types INNER JOIN Animals"
strAdoptionSQL = strAdoptionSQL & " ON Animal_Types.type_number = Animals.animal_type)"
strAdoptionSQL = strAdoptionSQL & " INNER JOIN Adoption ON Animals.animal_number = Adoption.adoption_animal)"
strAdoptionSQL = strAdoptionSQL & " ON Person.person_number = Adoption.adoption_adoptorNum"
strAdoptionSQL = strAdoptionSQL & " WHERE (((Switch([Adoption].[adoption_status]='P','Pending Verification',[Adoption].[adoption_status]='A','Aproved',[Adoption].[adoption_status]='D','Declined',[Adoption].[adoption_status]='C','Completed',[Adoption].[adoption_status]='H','Checkup'))='Checkup')) OR (((Switch([Adoption].[adoption_status]='P','Pending Verification',[Adoption].[adoption_status]='A','Aproved',[Adoption].[adoption_status]='D','Declined',[Adoption].[adoption_status]='C','Completed',[Adoption].[adoption_status]='H','Checkup'))='Aproved')) OR (((Switch([Adoption].[adoption_status]='P','Pending Verification',[Adoption].[adoption_status]='A','Aproved',[Adoption].[adoption_status]='D','Declined',[Adoption].[adoption_status]='C','Completed',[Adoption].[adoption_status]='H','Checkup'))='Pending Verification'))"
strAdoptionSQL = strAdoptionSQL & " ORDER BY Adoption.adoption_number;"

rstAdoptions.Open strAdoptionSQL, objConnection, adOpenKeyset, adLockOptimistic, adCmdText
Set frmActiveAdoptions.dgdActiveAdoptions.DataSource = rstAdoptions

Exit Sub
ErrorHandler:
    intMsgBox = MsgBox("A fatal error has occurred! " & Err.Description & Chr(13) & "The program must exit now.", vbCritical, "Fatal Error!")
    Open_Recordsets.objConnection.Close
    Set Open_Recordsets.objConnection = Nothing
    End

End Sub
