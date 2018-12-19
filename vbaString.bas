Attribute VB_Name = "Module2"
'++==
'
'   vbaString.bas
'   VBA String Library
'
'--==
Option Explicit

'++
' strReplace()
' Substitutes a phrase for another phrase in a given string.
'
' Args:
'   haystack - the string to process
'   needle   - the value to substitute out
'   subst    - the value to take needle's place
'
' Ret:
'   a string with the substituted values.
'
'--
Public Function strReplace(ByRef haystack As String, ByVal needle As String, ByVal subst As String) As String
        Dim off As Long:   off = 1
        Dim ret As String: ret = ""
        Dim start As Long: start = 1
        Dim skip As Long:  skip = Len(needle)
        
        If Len(haystack) < 1 Then
            strReplace = haystack
            Exit Function
        End If
        
        off = InStr(off, haystack, needle)
        
        ' obtain the first phrase
        If off >= 1 Then ret = ret & Mid(haystack, 1, off - 1) & subst
            
        ' all other phrase
        Do While (off > 0)
            start = off + skip
            off = InStr(start, haystack, needle)
            If off > 0 Then ret = ret & subst & Mid(haystack, start, off - start)
        Loop
        
        ' last phrase
        off = start
        If off <= Len(haystack) Then ret = ret & Mid(haystack, off, Len(haystack) - off + 1)

        strReplace = ret
End Function

' ++
' strSplit()
'   Split a string into an array along a given delimiter
'
'   Args:
'         src - the string to split
'       delim - the delimiter to split on
'
'   Ret:
'       An array of strings containing the substrings split along a delimiter
'
' Comment: Some modern versions of VBA have this function - Strings.Split()
' --
Public Function strSplit(ByRef src As String, ByVal delim As String) As String()
        Dim off As Long:   off = 1
        Dim start As Long: start = 1
        Dim skip As Long:  skip = Len(delim)
        Dim N As Long:     N = 1
        Dim i As Long:     i = 0
        Dim phrases() As String
        
        If Len(src) < 1 Then
            ReDim phrases(0)
            phrases(0) = src
            strSplit = phrases
            Exit Function
        End If
        
        'count the number of occurances of the delimiter
        off = InStr(off, src, delim)
        
        Do While (off > 0)
            start = off + skip
            off = InStr(start, src, delim)
            N = N + 1
        Loop
              
        ReDim phrases(N)
        off = InStr(1, src, delim)
        
        ' obtain the first phrase
        If off >= 1 Then
            phrases(i) = Mid(src, 1, off - 1)
            i = i + 1
        End If
        
        ' all other phrases
        Do While (off > 0)
            start = off + skip
            off = InStr(start, src, delim)
            If off > 0 Then
                phrases(i) = Mid(src, start, off - start)
                i = i + 1
            End If
        Loop
        
        ' last phrase
        off = start
        If off <= Len(src) Then phrases(i) = Mid(src, off, Len(src) - off + 1)
        
        strSplit = phrases
        
End Function
' ++
' strJoin()
'   Collapse a string array into a single string with array elements seperated by
'   a given delimiter. The complement of strSplit()
'
'   Args:
'         ar  - the string array to join / collapse
'       delim - the delimiter to seperate elements
'
'   Ret:
'       The joined string
'
'   Comment: Some versions of VBA have a built in version - Strings.Join
' --
Public Function strJoin(ByRef ar() As String, ByVal delim As String) As String
    Dim i As Long
    Dim retstr As String
    Dim N As Long:            N = UBound(ar)
    
    For i = 0 To N - 1
        retstr = retstr & ar(i) & delim
    Next i
    
    retstr
End Function

' ++
'   strLen()
'       Explicit alias of Len() which is commonly defined in other languages.
'   Args:
'      str - string to calculate the length
'
'   Ret:
'       The length of the string (Long)
' --
Public Function strLen(ByRef str As String)
    strLen = Len(str)
End Function

' ++
'   sizeOf()
'       Explicit alias of Len() which is defined in ANSI C and other languages.
'       In this usage it will calcualte the number of bytes a data type uses in memory.
'
'   Args:
'      var - variable to calculate length
'
'   Ret:
'       The data type size (Long)
' --
Public Function sizeOf(ByRef var As Variant)
    sizeOf = Len(var)
End Function
