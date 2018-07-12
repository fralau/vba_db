Attribute VB_Name = "Database"
'--------------------------------------------------------------------------------
' PURPOSE
'       Database generic functions
'
'
' AUTHOR
'       Laurent Franceschetti, 2010-2011
'       Update 2015
'
' USAGE

'
' REQUIRES
'       RegExp2 (Module)
'       Microsoft Scripting Runtime (Library)
'
' --------------------------------------------------------------------------------

 
Option Compare Database
Option Explicit
 
 
Enum SQLSyntaxType
    SQL_JET = 1
    SQL_SQLServer = 2
    SQL_Oracle = 3
    SQL_MYSQL = 4
    SQL_UNKNOWN = 5
End Enum
 
 
' A field may contain a word plus quotes (Oracle: "", MYSQL ``, Access [], SQL Server: '')
Const NOT_FIELD_PATTERN = "[^\w`'""\[\]]"

' Constant for Memo type (emulates VB naming conventions)
Const vbMemo = -1


#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If


'------------------------------------------------
' Quoting
'------------------------------------------------

Function SQLQuote(AnyVariable, Optional SQLSyntax As SQLSyntaxType = SQL_JET) As String
' Quotes smartly any variable
    Select Case VarType(AnyVariable)
        Case vbString:
            If AnyVariable = "" Then
                SQLQuote = "NULL"
            Else
                AnyVariable = Replace(AnyVariable, "'", """") ' replace simple quotes by double to avoid trouble
                SQLQuote = Quote(AnyVariable)
            End If
        Case vbDate:
            Select Case SQLSyntax:
                Case SQL_JET: SQLQuote = ToJetDate(AnyVariable)
                Case SQL_Oracle: SQLQuote = ToOracleDate(AnyVariable)
                Case SQL_MYSQL: SQLQuote = ToMySQLDate(AnyVariable)
                'TODO: insert SQL Server
                Case Else: SQLQuote = Quote(AnyVariable)
            End Select
        Case vbBoolean:
            ' Needed because of foreign versions of Access!
            If AnyVariable = True Then
                SQLQuote = "TRUE"
            Else
                SQLQuote = "FALSE"
            End If
        
        Case vbSingle, vbDouble:
            ' Necessary in case a Windows machine has a locale with "," (2017/04)
            SQLQuote = Replace(CStr(AnyVariable), ",", ".")
            
            
        Case vbNull, vbEmpty:
            SQLQuote = "NULL"
        Case Else:
            SQLQuote = AnyVariable
    End Select
End Function
 
Function ReplaceSQL(ByRef strMain As String, strToReplace As String, Replacement, Optional SQLSyntax As SQLSyntaxType = SQL_JET)
' Properly converts the variables into SQL equivalents
' this function also modifies the first parameter -- can be used as a sub
    ' USES: RegExp2
    strMain = RegReplace(strMain, strToReplace & "\b", SQLQuote(Replacement, SQLSyntax))
 
    ' Use the back tick as a replacement for "
    strMain = Replace(strMain, "`", """")
    ReplaceSQL = strMain
End Function
 
 
Function ReplaceEsc(str) As String
' Replace escape sequences
    ReplaceEsc = str
    ReplaceEsc = Replace(ReplaceEsc, "\n", vbCrLf)
    ReplaceEsc = Replace(ReplaceEsc, "\t", vbTab)
End Function
 
 
 
Function Interp(strSQL As String, ParamArray Args() As Variant) As String
' Interpolates a string, converting to SQL equivalent
' USAGE:
'   Interp(Query, 1stparameter, 2ndparameter,...., ConnectSequence)
'   In the query, the variables to replace are @1, @2...
'   If a variable is marked as @@1, then it is taken litterally (string value)
    
    Interp = strSQL
 
    'Debug.Print "Ubound: " & UBound(Args)
    If UBound(Args) >= 0 Then
 
        ' Determine syntax:
        ' If last argument starts with "ODBC" or "DATABASE" then consider it as a connection string
        ' This will be used to find the SQL Syntax
        Dim LastArg As String
        LastArg = CStr(Args(UBound(Args)))
        Dim SQLSyntax As SQLSyntaxType
        If Left(LastArg, 4) = "ODBC" Or Left(LastArg, 8) = "DATABASE" Or InStr(1, LastArg, "DSN=") Then
            SQLSyntax = ConnectType(LastArg)
            'Debug.Print "Syntax: " & SQLSyntax
        End If
 
 
        ' Interpolate
        Dim i As Long
        For i = LBound(Args) To UBound(Args)
            'Debug.Print i & " " & Args(i)
            ' If Not IsNull(Args(0)) Then Interp = Replace(Interp, "@1", Args(0)) -- the array starts at 0
            'If Not IsNull(Args(i)) Then
                ' If double @@, raw interpolation
                Interp = Replace(Interp, "@@" & i + 1, Nz(Args(i), "NULL"))
 
                ' Else interpolate according to syntax
                Interp = ReplaceSQL(Interp, "@" & i + 1, Args(i), SQLSyntax)
            'End If
        Next i
    End If
 
    Interp = ReplaceEsc(Interp)
 
 
End Function

Function Interp2(Statement As String, Dict As Dictionary, Optional Connect As String = "") As String
' Interpret the fields in a statement with a dictionary
'
' USAGE:
'   In the statement:
'   - the fields to replace and interpret as SQL are prefixed with a @ (e.g. @MY_FIELD)
'   - the fields to replace litterally are prefixed with a @@ (e.g. @MY_FIELD)
'   If the Connect string is provided, this allows to adapt the SQL to the target DB (useful for dates)

' This function builds on Interp.
    
    ' Check last parameter if it is a Connection string
    'Dim Connect As String
    'Connect = ""
    'If UBound(Args) > 0 Then
    '    Dim LastArg As String
    '    LastArg = CStr(Args(UBound(Args)))
    '    If ConnectType(LastArg) <> SQL_UNKNOWN Then Connect = LastArg
    '    Debug.Print "Connect:", Connect
    'End If
    
    ' Continue
    'On Error GoTo Err_Interp2
    Dim r As String
    Dim FieldName, FieldValue
    Dim FieldKey As String
 
 
 
 
 
 
 
    ' the regexp for identifying a field name in a table (@MY_FIELD):
    ' a field name may have a point in it

    Const FIELD_NAME_MATCH = "@@?((\w|\.|%|-)+)"
 
    r = Statement
    For Each FieldName In RegMatches(Statement, FIELD_NAME_MATCH)
 
        'Debug.Print FieldName & ":";
        If Left(FieldName, 2) = "@@" Then
            'Double @: Litteral interpretation
            FieldKey = Split(FieldName, "@")(2)
            'Debug.Print FieldName, Dict(FieldKey)
            
            If Not RegFind(FieldKey, "^\d*") Then
                If Not Dict.Exists(FieldKey) Then Err.Raise "5011", , Interp("Requested field @1 does not exist in Dictionary", FieldKey)
                ' Replace with the value in the dictionary with the litteral value
                r = RegReplace(r, CStr(FieldName) & "\b", Dict(FieldKey))
            End If
 
 
        Else
            ' Single @: Translate
            ' Find the actual key (after the @)
            FieldKey = Split(FieldName, "@")(1)
            'Debug.Print FieldName, Dict(FieldKey)
            
            If Not RegFind(FieldKey, "^\d*") Then
                If Not Dict.Exists(FieldKey) Then Err.Raise "5011", , Interp("Requested field @1 does not exist in Dictionary", FieldKey)
                ' Replace with the value in the dictionary with the interpolated value (for the specific database)
                r = RegReplace(r, CStr(FieldName) & "\b", Interp("@1", Dict(FieldKey), Connect))
            End If
        End If
        'Debug.Print Dict(FieldKey)
    Next
 
    ' Interpret the rest
    'If UBound(Args) > 0 Then R = Interp(R, Args)
    Interp2 = r
    Exit Function
 
Err_Interp2:
    Err.Raise Err.number, , "Cannot interpret: " & Err.Description
End Function


Function InterpRaw(strSQL As String, ParamArray Args() As Variant) As String
' Interpolates a string, raw value
    InterpRaw = strSQL
    Dim i As Long
    For i = LBound(Args) To UBound(Args)
        ' Debug.Print i & " " & Args(i)
        ' If Not IsNull(Args(0)) Then Interp = Replace(Interp, "@1", Args(0)) -- the array starts at 0
        If Not IsNull(Args(i)) Then InterpRaw = Replace(InterpRaw, "@" & i + 1, Args(i))
    Next i
    InterpRaw = ReplaceEsc(InterpRaw)
End Function
 
 
Function Quote(ByVal str As String) As String
    ' Quote ONLY if not quoted yet
    If Left(str, 1) <> "'" And Right(str, 1) <> "'" Then
        Quote = "'" & str & "'"
    Else
        Quote = str
    End If
End Function
 
Function QuoteList(strList As String) As String
' Quote the elements of a list
' e.g. QuoteList("foo,bar,baz") = "'foo','bar','baz'"
    QuoteList = strList
    QuoteList = RegReplace(QuoteList, "(.+?)(,|$)", "'$1'$2")
End Function
 
 
Function OutRow(A As Dictionary, ParamArray Args())
' Output a dictionary
    Dim Item
    Dim r As String
    r = ""
    For Each Item In Args
        r = Concat(", ", r, Concat("=", Item, SQLQuote(A(Item))))
    Next
    OutRow = "(" & r & ")"
End Function
 
'------------------------------------------------
' Creation of filters
'------------------------------------------------

Function ANDFilter(ByRef strFilter As String, ByVal strClause, Optional Direction = True) As String
' Adds an AND clause to a filter
    strFilter = Trim(strFilter)
 
    ' Quit the function if no clause
    If IsNull(strClause) Then Exit Function
    If strClause = "" Then Exit Function
 
 
 
 
    If strFilter <> "" Then strFilter = strFilter & " AND "
 
    If Direction = False Then strFilter = strFilter & "NOT "
 
    strFilter = strFilter & Trim("(" & strClause & ")")
    ANDFilter = strFilter
End Function
 
Function ORFilter(ByRef strFilter As String, ByVal strClause As String, Optional Direction = True) As String
' Adds an OR clause to a filter
    strFilter = Trim(strFilter)
    If strFilter <> "" Then strFilter = strFilter & " OR "
 
    If Direction = False Then strFilter = strFilter & "NOT "
 
    ' Enclose the filter into parentheses, to allow further and
    strFilter = "(" & strFilter & Trim("(" & strClause & ")") & ")"
    ORFilter = strFilter
End Function
 
 
 
 
 
'------------------------------------------------
' Power Database Functions
'------------------------------------------------

Sub DBExecute(strSQL As String, Connect As String, Optional ByRef TargetConnection As ADODB.Connection = Nothing, _
    Optional FailOnError As Boolean = True)
' Execute a query on a remote connection
' USAGE:
'   - If a TargetConnection is provided, it opens it (if needed)
'     This is necessary for
'   - Can execute multiple subqueries, separated by a ; (but escape ; with \)
    On Error GoTo Err_DBExecute
    
    'Const Separator = "<#separator#>"
    ' Replace escaped ;
    'strSQL = Replace(strSQL, "\;", Separator)
 
    ' Remove newlines before ;
    'strSQL = Replace(strSQL, vbCrLf & ";", ";")
    'strSQL = RegReplace(strSQL, "\n+?;", ";")
    
    ' Remove ODBC
    Connect = RegReplace(Connect, "ODBC\s*;", "")
 
    ' Initialize TargetConnection if not defined
    Dim NewTargetConnection As Boolean, DefaultTargetConnection As Boolean
    DefaultTargetConnection = IsMissing(TargetConnection)
    If TargetConnection Is Nothing Then
        NewTargetConnection = True
    ElseIf TargetConnection.STATE = adStateClosed Then
        NewTargetConnection = True
    Else
        NewTargetConnection = False
    End If
 
    If NewTargetConnection Then
        'Debug.Print "Created new connection"
        NewTargetConnection = True
        Set TargetConnection = New ADODB.Connection
        TargetConnection.Open Connect
    Else
        'Debug.Print "Not created new connection"
    End If
 
 
    ' Work by subqueries
    Dim subQuery
    For Each subQuery In SplitSQL(strSQL)
 
        subQuery = Trim(CStr(subQuery))
 
        ' Return escaped separator
 
        If subQuery <> "" Then
            'Debug.Print "------------------"
            'Debug.Print subQuery
            TargetConnection.Execute subQuery

        End If
 
    Next
 

 
    If DefaultTargetConnection Then
    ' Close
        Debug.Print "Closing..."
        TargetConnection.Close
        Set TargetConnection = Nothing
    End If
 
    Exit Sub
Err_DBExecute:
    If FailOnError Then
        Debug.Print strSQL
        Err.Raise Err.number, , "Cannot execute query: " & Err.Description & ". " & "Offending query: " & strSQL
    End If
End Sub






 
Function DBLookup(Domain As String, Optional Filter As String = "", Optional OrderBy As String = "", Optional Connect As String = "") As Collection
' Makes a lookup in a table that returns a collection of rows
' Returns a dictionary
'
' EXAMPLE :
'       Dim Row
'       For Each Row in DBLookup("User ","Department = 'Accounting'")
'           Debug.Print Row("UserName") & ": " & Row("UserFirstName") & " " & Row("UserLastName)
'       Next
'
' NOTE: As Domain, you can use either a table, an Access Query or an SQL statement

    On Error GoTo Err_DBLookup
 
 
    ' Prepare the query
    Dim strSQL As String
    If InStr(1, Domain, " ", vbTextCompare) = 0 Then
    ' If a single name (table or query)
        strSQL = "SELECT * FROM " & Domain
    Else
        strSQL = Domain
    End If
    'WAS: strSQL = "SELECT * FROM (" & Domain & ") as tmp0"
    
    ' Filter:
    strSQL = strSQL & IIf(Filter <> "", " WHERE " & Filter, "")
 
    ' Order By:
    strSQL = strSQL & IIf(OrderBy <> "", " ORDER BY " & OrderBy, "")
 
 
 
 
    ' Create the resulting object
    Set DBLookup = New Collection
 
    ' Each result row
    Dim ResultRow As Dictionary
 
 
    ' Fields
    Dim Field As Variant
 
    ' Open the dataset
    Dim rst As DAO.Recordset
 
 
    Dim DB As DAO.Database
    If Connect = "" Then
        Set DB = CurrentDb
    Else
        Set DB = DAO.OpenDatabase("", False, False, Connect)
    End If
    'Debug.Print Domain
        Set rst = DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
 
 
    ' Run through the dataset
    While Not rst.EOF
 
        'Debug.Print "Row"
        
        Set ResultRow = New Dictionary
        ResultRow.CompareMode = TextCompare ' Make the dictionary case-insensitive
 
        ' Add the fields to the dictionary
        For Each Field In rst.Fields
            'Debug.Print field.Name & " " & field.Value
            ResultRow.Add Field.Name, Field.Value
        Next
 
        ' Add the row to the collection
        DBLookup.Add ResultRow
 
        ' Re-assign ResultRow
        Set ResultRow = Nothing
 
        ' Next
        rst.MoveNext
    Wend
 
 
 
    rst.Close
    Set rst = Nothing
    Set DB = Nothing
 
 
    Exit Function
 
Err_DBLookup:
    Debug.Print strSQL
    Set rst = Nothing
    Set DB = Nothing
    If Err.number = 3170 Then
        Err.Raise Err.number, "DBLookup", "Cannot execute query as requested: ODBC string incorrect."
    Else
        Err.Raise Err.number, "DBLookup", "Cannot execute query as requested: " & Err.Description & " Offending query: " & strSQL
    End If
End Function
 





Public Function ALookup(Domain As String, Optional Filter As String = "", Optional OrderBy As String = "", Optional Connect As String = "") As Dictionary
' Allows to make a lookup that returns a whole row
' Returns a dictionary
    On Error GoTo Err_Alookup
 
    'Set ALookup = New Dictionary
    
    Dim rst
    Set rst = DBLookup(Domain, Filter, OrderBy, Connect)
 
 
    ' Take first item
    If rst.Count > 0 Then
        Set ALookup = rst.Item(1)
    Else
        'Set ALookup = Nothing
    End If
 
    Set rst = Nothing
    Exit Function
Err_Alookup:
    Dim Output As String
    Output = InterpRaw("Alookup(""@1"", ""@2"", ""@3"", ""@4"")", Domain, Filter, OrderBy, Connect)
    'Debug.Print output
    If Err.number = 3170 Then
        Err.Raise Err.number, "DBLookup", "Cannot execute query as requested: ODBC string incorrect."
    Else
        Err.Raise Err.number, "ALookup", Err.Description
    End If
End Function
 
Public Function RDLookup(Field As String, Domain As String, Optional Filter As String = "", Optional OrderBy As String = "", Optional Connect As String = "")
    ' Provides the result of a field
    ' Like DLookup
    On Error GoTo Err_RDlookup
    Dim Row As Dictionary
    ' Get the full row

    Set Row = ALookup(Domain, Filter, OrderBy, Connect)
    ' Get the field
    'Debug.Print "Row found"
    RDLookup = Row(Field)
    Exit Function
 
Err_RDlookup:
    'Debug.Print Err.Description
    If Err.number = 3170 Then
        Err.Raise Err.number, "DBLookup", "Cannot execute query as requested: ODBC string incorrect."
    Else
        Err.Raise Err.number, "RDLookup", Interp("(field @1): ", Field) & Err.Description
    End If
End Function
 
 
Function MDLookup(Field As String, Domain As String, Optional Filter As String = "", Optional OrderBy As String = "", Optional Connect As String = "") As Collection
' Returns a collection containing a list of UNIQUE items from a table
' You may use a Jet or VBA function
' e.g. :
'   - MDLookup("TradeDate","Transaction")
'   - MDLookup("Year(TradeDate)", "Transaction")
'   - MDLookup("Year(TradeDate) AS TradeYear", "Transaction")
' If you use a function, it must be a function implemented on the target database (Jet/VBA by default), otherwise the target database)
    
    On Error GoTo Err_MDLookup
 
 
    Dim strSQL As String
 
 
    ' If DOMAIN is something else than a pure name, surround with a parenthesis and rename
    If RegFind(Domain, NOT_FIELD_PATTERN) Then
        Domain = "(" & Domain & ") AS TMP_0102259"
    End If
 
 
    ' Create
    strSQL = "SELECT " & FieldSQL(Field) & " FROM " & Domain
    If Filter <> "" Then strSQL = strSQL & vbCrLf & "WHERE " & Filter
 
    strSQL = strSQL & vbCrLf & "GROUP BY " & FieldFormula(Field) ' better than DISTINCT
    If OrderBy <> "" Then strSQL = strSQL & ", " & OrderBy ' you have to add it to the GROUP BY
    
 
    If OrderBy <> "" Then strSQL = strSQL & vbCrLf & "ORDER BY " & FieldFormula(OrderBy)
 
 
 
 
 
    'Debug.Print strSQL
 
    Dim Row As Dictionary
 
    Set MDLookup = New Collection
    For Each Row In DBLookup(strSQL, , , Connect)
        'Debug.Print "Field: " & FieldName(Field) & "; Value Is : " & Row(FieldName(Field))
        MDLookup.Add (Row(FieldName(Field)))
    Next
    Exit Function
 
Err_MDLookup:
    Err.Raise Err.number, "MDLookup", "Cannot execute MDLookup query as requested: " & Err.Description
End Function
 
 
 
 
 
 
Sub DBUpdate(AssignStatement As String, Domain As String, Optional Filter As String = "", _
            Optional Connect As String = "", Optional ByRef TargetConnection As ADODB.Connection = Nothing)
' Update a table.
' Usage.
'   AssignStatement: the statement for updating, e.g.: "Field1 = '" & FieldContent & "', Field2 = TRUE"
    
    On Error GoTo Err_DBUpdate
    Dim strSQL As String
    strSQL = "UPDATE " & Domain & vbCrLf & "SET " & AssignStatement
 
    ' Add filter:
    If Filter <> "" Then
        strSQL = strSQL & vbCrLf & "WHERE " & Filter
    Else
        'Debug.Print "Filter is empty"
    End If
 
    'Debug.Print strSQL
    ' Execute query
    If Connect = "" Then
        CurrentDb.Execute strSQL
    Else
        ' Execute with the optional target connection (in case of multiple queries)
        DBExecute strSQL, Connect, TargetConnection
    End If
 
    Exit Sub
 
Err_DBUpdate:
    Debug.Print strSQL ' new
    Err.Raise Err.number, "DBUpdate", "Cannot execute DBUpdate as requested: " & _
                            vbCrLf & strSQL & " " & vbCrLf & _
                            Err.Description
End Sub
 
Public Function ExtDCount(Column As String, Domain As String, Optional Filter As String = "", Optional Connect As String = "")
' Works as DCount, but with a remote database + will  also work with a query
    Dim strSQL As String
    If InStr(1, Domain, " ", vbTextCompare) = 0 Then
    ' If a single name (table or query)
        strSQL = "SELECT Count(@Column) as NoFound FROM @Domain"
    Else
        strSQL = "SELECT Count(@Column) as NoFound FROM (@Domain) as tmp"
    End If
 
    strSQL = Replace(strSQL, "@Column", Column)
    strSQL = Replace(strSQL, "@Domain", Domain)
 
    'Debug.Print strSQL
    ExtDCount = ALookup(strSQL, Filter, , Connect)("NoFound")
End Function
 
 
Function ColumnExists(TableName As String, FieldName As String) As Boolean
' Checks whether a table exist
    On Error Resume Next
    ColumnExists = False
    ColumnExists = (CurrentDb.TableDefs(TableName).Fields(FieldName).Name <> "")
End Function
 
 
'----------------------------------------------------------
' Experimental: make CREATE Table Statement
'----------------------------------------------------------

Function MakeCreateStatement(TableName As String, Connect As String, SourceDomain As String, _
                Optional SourceFilter As String = "", Optional SourceOrderBy As String = "", Optional SourceConnect As String = "") As String
' Make a statement for a certain connection
    Dim r As String
    
    ' Get the field list
    Dim FieldList As Dictionary
    Set FieldList = GetDefinition(SourceDomain, SourceFilter, SourceOrderBy, SourceConnect)
    
    ' Create the rows
    Dim Field
    For Each Field In FieldList.Keys
        r = Concat("," & vbCrLf, r, Field & " " & GetSQLType(FieldList(Field), ConnectType(Connect)))
    Next
    
    
    r = "CREATE TABLE " & TableName & "(" & vbCrLf & r & ")"
    MakeCreateStatement = r
End Function



Function GetDefinition(Domain As String, _
                Optional Filter As String = "", Optional OrderBy As String = "", Optional Connect As String = "") As Dictionary
' Create a dictionary containing the VB types of tables
    
    Const MAX_NO_ROWS = 2000
    
    Dim Row As Dictionary
    Dim Result As New Dictionary
    
    Dim Column
    Dim ColumnFilled As Long
    
    Dim NoRows As Long
    NoRows = 0
    
    
    Dim i As Long
    ' Run through all rows
    For Each Row In DBLookup(Domain, Filter, OrderBy, Connect)
        ' Get through the columns of the row
        'Debug.Print "Next"
        NoRows = NoRows + 1
        
        
        i = 0

        For Each Column In Row.Keys
            If Result.Exists(Column) Then
                ' Assign the vartype
                If Result(Column) <> 0 Then
                    ' OK it is assigned
                    i = i + 1
                    

                Else
                    ' Not assigned yet
                    If Not IsNull(Row(Column)) Then
                        Result(Column) = VarType(Row(Column))
                        i = i + 1
                    End If
                End If
            Else
                ' not found yet
                If Not IsNull(Row(Column)) Then
                    ' New one
                    Result(Column) = VarType(Row(Column))
                    
                    i = i + 1
                Else
                    ' Create the column with a "zero" type
                    Result(Column) = 0
                End If
            End If
            
            'Check string, just in case
            If Result(Column) = vbString And Not IsNull(Row(Column)) Then
                If Len(Row(Column)) > 255 Then
                    Result(Column) = vbMemo
                End If
            End If
            


            
            ' If the result table (containing the types) is full, then stop
            If i = Row.Count Then Exit For
            
            ' Stop after a no of rows
            If NoRows > MAX_NO_ROWS Then Exit For
        Next
        'Debug.Print Interp("@1 found from @2", i, Row.Count)
    Next
    

    
    ' Assign the result
    Set GetDefinition = Result
End Function



Function GetSQLType(VBAType, SQLSyntax As SQLSyntaxType) As String
' Make a conversion of the field type

    ' vbMemo is an application specific constant
    
    Dim r As String
    'Debug.Print SQLSyntax
    
    If SQLSyntax = SQL_Oracle Then
        Select Case VBAType
            Case vbInteger, vbLong:
                r = "INT"
            Case vbDate:
                r = "DATE"
            Case vbSingle, vbDouble:
                r = "DOUBLE PRECISION"
            Case vbMemo:
                r = "CLOB"
            Case Else:
                ' Anything else is varchar
                r = "VARCHAR2(255)"
        End Select
    ElseIf SQLSyntax = SQL_MYSQL Then
            Select Case VBAType
            Case vbInteger, vbLong:
                r = "INT"
            Case vbDate:
                r = "DATE"
            Case vbSingle, vbDouble:
                r = "FLOAT"
            Case vbMemo:
                r = "TEXT"
            Case Else:
                ' Anything else is varchar
                r = "VARCHAR(255)"
        End Select
    ElseIf SQLSyntax = SQL_UNKNOWN Then
        Err.Raise 5023, , "Unknown or unidentified Connect chain"
    
    Else
        ' Assume Access
        Select Case VBAType
            Case vbInteger, vbLong:
                r = "INTEGER"
            Case vbDate:
                r = "DATE"
            Case vbSingle, vbDouble:
                r = "DOUBLE"
            Case vbMemo:
                r = "MEMO"
            Case Else:
                ' Anything else is varchar
                r = "VARCHAR(255)"
        End Select
    End If
        
    
    GetSQLType = r
End Function

'----------------------------------------------------------
' Save Row
'----------------------------------------------------------
 
Public Sub SaveRow(Row As Dictionary, _
                    OutTable As String, _
                    Optional Connect As String = "", _
                    Optional ByVal CurrentConnection As ADODB.Connection = Nothing)
' Save the Row to the OutTable Table
' Dynamically saves every field that matches between the dictionary and the OutTable
' If an adodb connection is provided, uses it (this is used to prevent continuous open/close)

 
    On Error GoTo Err_SaveRow
    Dim myConnection As ADODB.Connection
 
    ' If no current connection, open it
    
 
    If CurrentConnection Is Nothing Then
        If Connect <> "" Then
            ' Open a new connection
            Set myConnection = New ADODB.Connection
            myConnection.Open Connect
        Else
            ' It is the current database
            Set myConnection = CurrentProject.Connection
        End If
    Else
        ' Use the connection
        Set myConnection = CurrentConnection
    End If
 
 
 
 
 
 
 
    Dim strSQL As String ' the query
    Dim ColumnList As String, ValueList As String ' to build the query
    
 
    Dim FieldName, strFieldName As String
    For Each FieldName In Row.Keys
        strFieldName = FieldName
        'Debug.Print "Found: " & FieldName
        
        ' Save only if the column exists in OutTable
        If ColumnExists(OutTable, strFieldName) Then
 
            'Debug.Print FieldName & " " & Row(strFieldName)
            
            ' Add to field list
            ColumnList = ColumnList & ", " & strFieldName
 
            ' Add to value list
            ValueList = ValueList & ", " & SQLQuote(Row(strFieldName), ConnectType(Connect))
        Else
            ' DO NOTHING
            'Debug.Print FieldName & " not found"
        End If
    Next
 
    ' Remove comma
    ColumnList = RemoveFirstChar(ColumnList)
    ValueList = RemoveFirstChar(ValueList)
 
    ' Remove row and re-create
    'strSQL = InterpRaw("DELETE FROM @1 WHERE TransactionId = '@2'", OutTable, Me.TransactionId)

 
    'myConnection.Execute strSQL
    'ExecuteRemote strSQL, Connect
    
    strSQL = InterpRaw("INSERT INTO @1 \n   (@2) \n   VALUES (@3)", OutTable, ColumnList, ValueList)
    'Debug.Print "Query: [ " & vbCrLf & strSQL & vbCrLf & "]" & vbCrLf
    
    'Debug.Print strSQL
    myConnection.Execute strSQL
    
    
    
    
 
    If CurrentConnection Is Nothing Then
    ' If we had to create a new connection, close it
        myConnection.Close
        Set myConnection = Nothing
    End If
    Exit Sub
 
Err_SaveRow:
    'Debug.Print "Column List: " & ColumnList
    'Debug.Print "Value List: " & ValueList
    Debug.Print "Offending query: " & vbCrLf & strSQL
    Err.Raise Err.number, "SaveRow", "Cannot save record. Offending query: " & vbCrLf & vbCrLf & _
                "**************** START OF QUERY ************" & vbCrLf & _
                strSQL & vbCrLf & "**************** END OF QUERY ************" & vbCrLf & Err.Description
End Sub
 
 
Private Function RemoveFirstChar(myString As String) As String
' Remove the first character from a string -- also trim
    RemoveFirstChar = Trim(Right(myString, Len(myString) - 1))
End Function
 
Function FieldSQL(FieldExpression As String) As String
     ' Gets the SQL Wording of a field expression
     ' Normalizes a Field wording so that it always has a AS if needed
     ' FieldSQL("Year(TradeDate)") => Year(TradeDate) As YearTradeDate
    Dim myFieldName As String
    myFieldName = FieldName(FieldExpression)
    If myFieldName = FieldExpression Then
        ' No difference, a simple field name (e.g. TradeDate)
        FieldSQL = FieldExpression
    Else
        ' IF there should be a difference, i.e. something after the AS
        FieldSQL = FieldExpression
        'If no as, then ad the field name:
        If Not RegFind(FieldExpression, " AS ") Then
            FieldSQL = FieldExpression & " AS " & FieldName(FieldExpression)
        End If
    End If
 End Function
 
Function FieldName(ByRef FieldExpression As String) As String
    ' Gets the name of a field expression
    ' can be e.g. "TradeDate" or "Year(TradeDate)"
    ' e.g. FieldName("Year(TradeDate) AS YearTradeDate") => YearTradeDate
    ' If the Field Expression does not contain AS, it uses a default

 
    FieldName = Trim(FieldExpression)
    If RegFind(FieldExpression, NOT_FIELD_PATTERN) Then
        ' An actual expression
        If Not RegFind(FieldExpression, "\sAS\s") Then
            ' No as found, turn the expression into a variable name
            FieldName = RegReplace(FieldExpression, NOT_FIELD_PATTERN, "")
        Else
            FieldName = Trim(Replace(RegMatch(FieldExpression, "\sAS\s\w+"), " AS ", ""))
        End If
    End If
 End Function
 
Function FieldFormula(FieldExpression As String) As String
 ' Gets the formula in a field
 ' e.g. FieldFormula("Year(TradeDate) AS YearTradeDate") => Year(TradeDate)
    FieldFormula = Split(FieldExpression, " AS ")(0)
 End Function
 
 
 
 
 
'------------------------------------------------
' Date Functions
'------------------------------------------------

 
Function ToSQLDate(ByVal myDate As Date, Optional AddMonth As Integer = 0, Optional ByVal SQLSyntax As SQLSyntaxType = SQL_JET) As String
' Converts a VBA date to a database date
    Select Case SQLSyntax
        Case SQL_JET
            ToSQLDate = ToJetDate(myDate, AddMonth)
        Case SQL_Oracle
            ToSQLDate = ToOracleDate(myDate, AddMonth)
        Case SQL_MYSQL
            ToSQLDate = ToMySQLDate(myDate, AddMonth)
        Case Else
            Err.Raise 5012, "Cannot produce date for that syntax"
        End Select
End Function
 
Function ToOracleDate(ByVal myDate As Date, Optional AddMonth As Integer = 0) As String
'Converts a VBA date to Oracle format (text), for a query
    Dim OracleFormula As String
 
    OracleFormula = Format(myDate, "YYYYMMDD hh:mm:ss")
 
    ToOracleDate = "TO_DATE('" & OracleFormula & "', 'YYYYMMDD HH24:MI:SS')"
 
 
    If AddMonth <> 0 Then
        ToOracleDate = "ADD_MONTHS(" & ToOracleDate & "," & AddMonth & ")"
    End If
End Function
 
 
Function ToJetDate(ByVal myDate As Date, Optional AddMonth As Integer = 0) As String
'Converts a VBA date to Jet Format (text), for a query
    'ToJetDate = Format(myDate, "#YYYY-MM-DD#")
        ToJetDate = "#" & Format(DateAdd("m", AddMonth, myDate), "YYYY-MM-DD h:m:s") & "#"
End Function
 
Function ToMySQLDate(ByVal myDate As Date, Optional AddMonth As Integer = 0) As String
'Converts a VBA date to MySQL format (text), for a query
    ToMySQLDate = "'" & Format(DateAdd("m", AddMonth, myDate), "YYYY-MM-DD hh:mm:ss") & "'"
End Function
 
 
 
 
 
'------------------------------------------------
' ODBC - DSN Related functions
'------------------------------------------------

Public Function ConnectType(Connect As String) As SQLSyntaxType
' Get the syntax type from an ODBC chain, whether DSNless or DSN
    
    On Error Resume Next
    ' In case of error
    ConnectType = SQL_UNKNOWN
    
    Dim DSN As String
    If Connect = "" Then
        ' Default table
        ConnectType = SQL_JET
    ElseIf InStr(1, Connect, "DATABASE=") Then
        ' Access Database
        ConnectType = SQL_JET
    ElseIf InStr(1, Connect, "DSN=") Then
        ' DSN
        DSN = Replace(RegMatch(Connect, "DSN=\w*"), "DSN=", "")
        ConnectType = DSNType(DSN)
    ElseIf InStr(1, Connect, "ODBC=") Then
        ' ODBC chain (not DSN)
        ConnectType = GetDriverType(Connect)
    Else
        ' We don't know
        ConnectType = SQL_UNKNOWN
    End If
End Function
 
 
Public Function DSNType(DSN As String) As SQLSyntaxType
' Get the syntax type from the DSN name
    Dim SystemDriverName As String, UserDriverName As String
    UserDriverName = RegKeyRead(InterpRaw("HKCU\Software\ODBC\ODBC.INI\@1\Driver", DSN))
    SystemDriverName = RegKeyRead(InterpRaw("HLM\Software\ODBC\ODBC.INI\@1\Driver", DSN))
    'Debug.Print "U: " & UserDriverName
    'Debug.Print "S: " & SystemDriverName
    If UserDriverName <> "" Then
        DSNType = GetDriverType(UserDriverName)
    ElseIf SystemDriverName <> "" Then
        DSNType = GetDriverType(SystemDriverName)
    Else
        Err.Raise 5010, "DSNType", "DSN '" & DSN & "' not found"
    End If
End Function
 
 
Public Function GetDriverType(DriverName As String) As SQLSyntaxType
' Get the syntax type, from the driver name
    If InStr(1, DriverName, "MYSQL") > 0 Then
        GetDriverType = SQL_MYSQL
    ElseIf InStr(1, DriverName, "sqora") > 0 Then
        GetDriverType = SQL_Oracle
    ElseIf InStr(1, DriverName, "SQLSRV") > 0 Then
        GetDriverType = SQL_Oracle
    ElseIf InStr(1, DriverName, "ACEODBC") > 0 Then
        GetDriverType = SQL_JET
    Else
        GetDriverType = SQL_UNKNOWN
    End If
End Function
 
 
Function RegKeyRead(i_RegKey As String) As String
' Reads the value for the registry key i_RegKey
' If the key cannot be found, the return value is ""
' Credits: See http://vba-corner.livejournal.com/3054.html
    On Error Resume Next
 
    Dim myWS As Object
    'Access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    RegKeyRead = myWS.RegRead(i_RegKey)
End Function


Function Concat(SEPARATOR, ParamArray Args() As Variant) As String
' Concats n items with a separator
    Dim Item
    Concat = ""
    
    For Each Item In Args
        If IsNull(Item) Then Item = ""
        Concat = Concat & IIf(Concat <> "" And Item <> "", SEPARATOR, "") & Item
    Next
End Function

'-----------------------------------------------------------------
' Spit a SQL statement
'-----------------------------------------------------------------
Function SplitSQL(Script As String) As Collection
' Split a SQL string into statements difficulty is that the separator is a ";"
' and it can be found elsewhere
' WARNING: it may miss cases in embedded comments

    Dim CorrScript As String
    Dim r As New Collection
    
    'Debug.Print "-------------"
    'Debug.Print Script
    'Debug.Print "-------------"
    
    
    Const SEPARATOR = "#SEP0214346d945#"
    Const TRUE_SEPARATOR = "#SEPas019k3dj082kwsj92831klj#"
    Const STRING_LITT = "'.*?'"
    
    Dim CorrItem As String
    
    Dim Item
    
    ' Correct the simple comments:
    CorrScript = Script
    CorrScript = EscapeStr(CorrScript, "'(.|\n)*?'", ";", SEPARATOR) ' within strings
    CorrScript = EscapeStr(CorrScript, "^--.*?$", ";", SEPARATOR) ' within -- comments
    CorrScript = EscapeStr(CorrScript, "/\*(.|\n)*?\*/", ";", SEPARATOR) ' within /* */ comments
    
    
    
    
    'Debug.Print "------ PROTECTED -----------------"
    'Debug.Print CorrScript
    'Debug.Print "----------------------------------"
    
    ' Find finally the correct separator
    CorrScript = Replace(CorrScript, ";", TRUE_SEPARATOR)
    
    ' Return the previous ones
    CorrScript = Replace(CorrScript, SEPARATOR, ";")
    
    ' Split now
    
    
    Dim Statement
    For Each Statement In Split(CorrScript, TRUE_SEPARATOR)
        'Debug.Print "NEW STATEMENT: " & Statement
        r.Add CStr(Statement)
    Next
    
    Set SplitSQL = r
End Function

Private Function EscapeStr(Statement As String, Pattern As String, Sequence As String, EscapedSequence As String) As String
' On a statement, find a regexp pattern and within that pattern, replaces Sequence by EscapedSequence
' It also handles multiline statements
    Dim CorrItem As String
    Dim Item
    'Debug.Print "Escaping " & Statement
    EscapeStr = Statement
    For Each Item In RegMatches(EscapeStr, Pattern, , True)
        'Debug.Print "Found: [" & item & "] with  pattern " & Pattern
        CorrItem = Replace(CStr(Item), Sequence, EscapedSequence)
        EscapeStr = Replace(EscapeStr, CStr(Item), CorrItem)
    Next
    'Debug.Print "Result: " & EscapeStr
End Function

'----------------------------------------------------------
' New functions
'----------------------------------------------------------

Function NewLine() As String
' Sometimes needed in queries
' Addition 2015
    NewLine = vbCrLf
End Function

Function TableInterp(Clause As String) As String
' Interpolate a clause with the columns names
    ' the regexp for identifying a field name in a table (#MY_FIELD):
    ' a field name may have a point in it

    Const FIELD_NAME_MATCH = "#((\w|\.|%|-)+)"
    Dim FieldName
    Dim r As String, f As String
    
    
    ' double quotes:
    r = Clause
    r = Replace(r, "'", "''")
    
    ' quote the clause
    r = Quote(r)
    
    For Each FieldName In RegMatches(Clause, FIELD_NAME_MATCH)
        'Debug.Print FieldName
        f = Mid(FieldName, 2)
        r = Replace(r, FieldName, Interp("' & @@1 & '", f))
    Next
    TableInterp = r
End Function
