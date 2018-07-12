Attribute VB_Name = "Parameter"
' courtesy of Laurent Franceschetti
' (copyright Settlenext)

Option Compare Database
 
Const LOCAL_TABLE = "LocalParameter"
Const CENTRAL_TABLE = "Parameter"
 
Enum ParamType
'Type of parameter
    MDB = 1
    Central = 2
    Default = 3
End Enum
 
 


 
Function myParam(ParamId As String, Optional intParamType As ParamType = Default) As Variant
' Returns the value of a parameter
'   - NULL if not found
'   - "" if null value of the parameter or empty string
 
    'On Error Resume Next
    myParam = Null
    'We trim the string because the record may be filled with blanks (e.g. with an UPDATE)
    'Debug.Print intParamType
 
    Select Case intParamType
    Case Is = MDB
        myParam = Trim(DLookup("[ParamValue]", LOCAL_TABLE, "ParamId = """ & ParamId & """"))
 
        If IsNull(myParam) And DCount("[ParamValue]", LOCAL_TABLE, "ParamId = """ & ParamId & """") > 0 Then
            myParam = ""
        End If
    Case Is = Central
        myParam = Trim(DLookup("[ParamValue]", CENTRAL_TABLE, "ParamId = """ & ParamId & """"))
        ' If null check whether it is not found or empty:
        If IsNull(myParam) And DCount("[ParamValue]", CENTRAL_TABLE, "ParamId = """ & ParamId & """") > 0 Then
            myParam = ""
        End If
    Case Is = Default
        'First local, then central
        If DoesTblExist(LOCAL_TABLE) Then
            ' First local
            myParam = Trim(DLookup("[ParamValue]", LOCAL_TABLE, "ParamId = """ & ParamId & """"))
            ' If null check whether it is not found or empty:
            If IsNull(myParam) And DCount("[ParamValue]", LOCAL_TABLE, "ParamId = """ & ParamId & """") > 0 Then
                myParam = ""
            End If
            ' Then remote
            If IsNull(myParam) Then
                myParam = Trim(DLookup("[ParamValue]", CENTRAL_TABLE, "ParamId = """ & ParamId & """"))
                ' If null check whether it is not found or empty:
                If IsNull(myParam) And DCount("[ParamValue]", CENTRAL_TABLE, "ParamId = """ & ParamId & """") > 0 Then
                    myParam = ""
                End If
            End If
        ElseIf DoesTblExist(CENTRAL_TABLE) Then
        ' Case where CENTRAL_TABLE exists and not LOCAL_TABLE (initial situation)
                myParam = Trim(DLookup("[ParamValue]", CENTRAL_TABLE, "ParamId = """ & ParamId & """"))
                ' If null check whether it is not found or empty:
                If IsNull(myParam) And DCount("[ParamValue]", CENTRAL_TABLE, "ParamId = """ & ParamId & """") > 0 Then
                    myParam = ""
                End If
        End If
    End Select
 
 
 
    'Translate CR:
    If Not IsNull(myParam) Then
        myParam = Replace(myParam, "%vbCrLF%", vbCrLf)
    End If
 
    If IsNull(myParam) Then
        Debug.Print "Cannot find " & ParamId
    End If
    Exit Function
 
 
End Function
 
Function myParamInt(ParamId As String, Optional intParamType As ParamType = Default) As Integer
'Returns a Long when expected
    On Error Resume Next
    myParamInt = 0
    myParamInt = CInt(myParam(ParamId, intParamType))
End Function
 
Sub UpdateParam(ParamId As String, ParamValue, Optional intParamType As ParamType = Default)
'Updates a parameter with a value
'Creates the table if it does not exist
    'On Error GoTo Err_UpdateParam
    Dim strQuery As String
 
    'If IsNull(ParamValue) Then
    '    ParamValue = ""
    '    Debug.Print "Value set to null string"
    'End If
 
    If Not IsNull(myParam(ParamId, intParamType)) Then
        strQuery = "UPDATE #table# SET ParamValue= """ & ParamValue & """ WHERE ParamId = """ & ParamId & """;"
    Else
        strQuery = "INSERT INTO #table# (ParamId,ParamValue) VALUES (""" & ParamId & """ , """ & ParamValue & """);"
    End If
 
    Select Case intParamType
    Case Is = Default
        '2008-12 LF Check if Parameter table exists, then the LocalParameter, if nothing, create the central table
        If DoesTblExist(CENTRAL_TABLE) Then
            strQuery = Replace(strQuery, "#table#", CENTRAL_TABLE)
        ElseIf DoesTblExist(LOCAL_TABLE) Then
            strQuery = Replace(strQuery, "#table#", LOCAL_TABLE)
        Else
            CheckParamTable (Central)
            strQuery = Replace(strQuery, "#table#", CENTRAL_TABLE)
        End If
    Case Is = MDB
        CheckParamTable (MDB)
        strQuery = Replace(strQuery, "#table#", LOCAL_TABLE)
    Case Is = Central
        CheckParamTable (Central)
        strQuery = Replace(strQuery, "#table#", CENTRAL_TABLE)
    End Select
    Debug.Print "Query is: " & strQuery
    Application.CurrentDb.Execute strQuery
    Exit Sub
 
Err_UpdateParam:
    Err.Raise 15001, "UpdateParam", "Error could not update parameter " & ParamId & " with " & ParamValue & "." & vbCrLf & _
            Err.number & ": " & Err.Description
End Sub
 
Private Sub CheckParamTable(intParamType As ParamType)
'Check if a parameter table does exist
    Dim strQuery As String
 
    strQuery = "CREATE TABLE #table# (ParamId varchar(20) PRIMARY KEY, ParamValue varchar(50));"
 
    Select Case intParamType
    Case Is = MDB
        If Not DoesTblExist(LOCAL_TABLE) Then
            strQuery = Replace(strQuery, "#table#", LOCAL_TABLE)
            'MsgBox strQuery
            Application.CurrentDb.Execute strQuery
        End If
    Case Is = Central
        If Not DoesTblExist(CENTRAL_TABLE) Then
            strQuery = Replace(strQuery, "#table#", CENTRAL_TABLE)
            MsgBox "Table Parameter created; if you want to share it, you will have to push it centrally.", vbExclamation
            'MsgBox strQuery
            Application.CurrentDb.Execute strQuery
        End If
    End Select
End Sub
 
 
      Private Function DoesTblExist(strTblName As String) As Boolean
         On Error Resume Next
 
         Set tbl = CurrentDb.TableDefs(strTblName)
         If Err.number = 3265 Then   ' Item not found.
            DoesTblExist = False
            Exit Function
         End If
         DoesTblExist = True
      End Function





