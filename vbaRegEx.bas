Attribute VB_Name = "vbaRegEx"
Option Explicit
'++++
'   Regular Expressions
'
'  NOTICE: References the MS VBScript Regular Expression v. 5.5 library
'  Built into Windows 7 and later, but also available for download
'
'  TO-DO: Ideally the RegEx parser can be either compiled and linked ( C implementation )
'  or ported into native VBA so that the proprietary dependance is removed.
'  tiny-regex-c looks like an attractive open source project to follow
'
'----

'+
' grep()
'   A unix grep-like regular expression function. Returns matching elements of a given regex
'   for a given string array.
'
'   Args:
'       ar - array of strings to search with regular expressions (String())
'    regex - regular expression to match
'    flags - a list of mode flags which determine the match and output parameters (string)
'          'i' - case insensitive
'          'o' - output only the match
'          'm' - multiline
'          'g' - global
'
'   Ret:
'           An array of elements that match the given regex.
'
'   Note:
'       grep -io 'test'  behaves similarly to grep( ..., 'test', 'io' )
'
'   TO-DO: Implement 'v' option ( exclude matching elements )
'          Impelement 'o' option that includes all matches
'
'-
Function grep(ByRef ar() As String, ByVal regex As String, Optional ByVal flags) As String()
    'Declaration and Initalization
    Dim re As Object:          Set re = CreateObject("VBScript.RegExp")
    Dim N As Long:             N = UBound(ar)
    Dim indxMatch() As Long:   ReDim indxMatch(N)
    Dim M As Long:             M = 0
    Dim i As Long, j As Long:  j = 0
    Dim flgO As Boolean:       flgO = False
    
    'setup regex object
    re.Pattern = regex
    
    If IsMissing(flags) = True Then flags = ""
    
    If InStr(flags, "g") > 0 Then re.Global = True Else re.Global = False
    If InStr(flags, "i") > 0 Then re.IgnoreCase = True Else re.IgnoreCase = False
    If InStr(flags, "m") > 0 Then re.MultiLine = True Else re.MultiLine = False
    If InStr(flags, "o") > 0 Then flgO = True Else flgO = False
    
    'see if any of the lines match
    For i = 0 To N - 1
        If re.test(ar(i)) = True Then
            indxMatch(j) = i
            j = j + 1
        End If
    Next i
    
    M = j
    
    'assemble return array from matches lines
    Dim ret() As String: ReDim ret(M)
    
    For j = 0 To M - 1
        ret(j) = ar(indxMatch(j))
    Next j
    
     If flgO = False Then
        grep = ret
        Exit Function
     Else
        'O option
        grep = ret
        
        For j = 0 To M - 1
            reMatches = re.Execute(ret(j))
        Next j
     End If
     
     
End Function
'+
' sed()
'   A unix sed-like (s command) regular expression function. Substitutes matching regex of
'   for a given string array.
'
'   Args:
'          src - string to process
'    regex_find - regular expression to match
'        subst - value to replace the matching regex
'       flags - 'i' - case insensitive
'               'g' - global substitution (otherwise only first match )
'               'm' - multiline
'   Ret:
'       The substituted string
'
'   Note: This follows Javascript RegEx dialect
'         group delimiters ( ) and multiple denotes { } are NOT escaped.
'         '(', ')', '{', '}' are escaped.
'         Backsubstitution notation is using $1, $2, $3... notation
'         Entire Match is $&
'-
Function sed(ByRef src As String, ByRef regex_find, ByRef subst, Optional ByVal flags As String = "")
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")

    flags = LCase(flags)
    re.Pattern = regex_find

    If InStr(flags, "g") > 0 Then re.Global = True Else re.Global = False
    If InStr(flags, "i") > 0 Then re.IgnoreCase = True Else re.IgnoreCase = False
    If InStr(flags, "m") > 0 Then re.MultiLine = True Else re.MultiLine = False
       
    sed = re.Replace(src, subst)
End Function
