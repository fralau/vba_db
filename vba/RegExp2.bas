Attribute VB_Name = "RegExp2"
' --------------------------------------------------------------------------------
' PURPOSE
'       Manage Regular expressions in a simple way
'
'
' AUTHOR
'       Laurent Franceschetti, June/July 2009
'
' USAGE
'       See TestRegMatches subroutine
'
' REQUIRES
'       Microsoft VB Script Regular Expressions (Library)
'
' --------------------------------------------------------------------------------
 
Option Compare Database
Option Explicit
 
Function RegMatch(ByVal sourceString As String, Pattern As String, _
                        Optional IgnoreCase = True, Optional Multiline = False) As String
' Gets first match of a string
    On Error GoTo Err_RegMatch
    RegMatch = ""
    Dim Reg As RegExp
    Set Reg = New RegExp
    Reg.IgnoreCase = IgnoreCase ' If false, it will make a difference between normal an capitals.
    Reg.Multiline = Multiline ' If true, ^ and $ will be interpreted as start of line
    Reg.Global = False ' Only the first one
    
    Dim myMatches As MatchCollection
 
 
    Reg.Pattern = Pattern
    Set myMatches = Reg.Execute(sourceString)
 
    If myMatches.Count > 0 Then
        RegMatch = myMatches(0)
    End If
 
    Set Reg = Nothing
    Set myMatches = Nothing
    Exit Function
 
Err_RegMatch:
    Err.Raise "5012", , "Error in Regular expression. " & vbCrLf & Err.number & " " & Err.Description
End Function
 
Function RegFind(ByVal sourceString As String, Pattern As String, _
                        Optional IgnoreCase = True, Optional Multiline = False) As Boolean
' Determines whether a pattern is found. For use in a SQL query
    RegFind = (Len(RegMatch(sourceString, Pattern, IgnoreCase, Multiline)) > 0)
End Function
 
 
 
Function RegReplace(sourceString As String, Pattern As String, ReplaceVar As String, _
                        Optional IgnoreCase = True, Optional GlobalReplace = True, Optional Multiline = False) As String
' Replace a pattern in a string; by default replaces all occurrences
' To replace only the first occurence, set GlobalReplace to False.
    RegReplace = ""
    Dim Reg As RegExp
    Set Reg = New RegExp
    Reg.IgnoreCase = IgnoreCase
    Reg.Multiline = Multiline ' If true, ^ and $ will be interpreted as start of line
    Reg.Global = GlobalReplace ' This makes the replace "global" or not
    Reg.Pattern = Pattern
    RegReplace = Reg.Replace(sourceString, ReplaceVar)
 
    Set Reg = Nothing
End Function
 
 
 
 
Function RegMatches(sourceString As String, Pattern As String, _
                        Optional IgnoreCase = True, Optional Multiline = False) As MatchCollection
' Gets all matches of a string
 
    Dim Reg As RegExp
    Set Reg = New RegExp
    Reg.IgnoreCase = IgnoreCase
    Reg.Multiline = Multiline ' If true, ^ and $ will be interpreted as start of line
    Reg.Global = True ' This makes the search "global"
 
    Dim myMatches As MatchCollection
    Dim myMatch As Match
 
    Reg.Pattern = Pattern
    Set myMatches = Reg.Execute(sourceString)
 
    Set RegMatches = myMatches
 
    Set Reg = Nothing
    Set myMatches = Nothing
End Function
 
Sub TestRegMatches(sourceString As String, Pattern As String)
' Test Regmatches; you can use it as example, or a debug tool for your regular expressions.
    Dim myMatch As Match
    For Each myMatch In RegMatches(sourceString, Pattern)
        Debug.Print "Found " & myMatch.Value
    Next
End Sub


