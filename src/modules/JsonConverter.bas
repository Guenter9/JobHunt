Attribute VB_Name = "JsonConverter"
'==============================================================================
' Module     : JsonConverter
' Description: Lightweight JSON parser and serializer for VBA.
'              Requires: Microsoft Scripting Runtime (Scripting.Dictionary)
'
' Public API:
'   ParseJson(sJson As String) As Object
'       Returns a Scripting.Dictionary for JSON objects,
'               a VBA Collection for JSON arrays,
'               a String/Double/Long/Boolean for scalar values.
'
'   ConvertToJson(vValue As Variant, Optional iIndent As Long = 0) As String
'       Serializes a Dictionary, Collection, or primitive to a JSON string.
'       iIndent = 0 (no whitespace), >0 = spaces per level.
'==============================================================================
Option Explicit

' ---- Public entry points -----------------------------------------------------

Public Function ParseJson(ByVal sJson As String) As Object
    Dim lPos As Long
    lPos = 1
    sJson = Trim(sJson)
    Dim vResult As Variant
    vResult = ParseValueIntoVariant(sJson, lPos)
    If IsObject(vResult) Then
        Set ParseJson = vResult
    Else
        ' Scalar root – wrap in a Dictionary for consistency
        Dim d As New Scripting.Dictionary
        d("value") = vResult
        Set ParseJson = d
    End If
End Function

Public Function ConvertToJson(ByVal vValue As Variant, _
                              Optional ByVal iIndent As Long = 0, _
                              Optional ByVal iLevel As Long = 0) As String
    Dim sOut  As String
    Dim sPad  As String
    Dim sPad1 As String
    Dim vKey  As Variant
    Dim bFirst As Boolean

    If iIndent > 0 Then
        sPad  = String(iLevel * iIndent, " ")
        sPad1 = String((iLevel + 1) * iIndent, " ")
    End If

    If IsNull(vValue) Or IsEmpty(vValue) Then
        ConvertToJson = "null"
        Exit Function
    End If

    Select Case TypeName(vValue)
        Case "Dictionary"
            bFirst = True
            sOut = "{"
            For Each vKey In vValue.Keys
                If Not bFirst Then sOut = sOut & ","
                If iIndent > 0 Then sOut = sOut & vbNewLine & sPad1
                sOut = sOut & """" & EscapeJson(CStr(vKey)) & """:" & _
                       IIf(iIndent > 0, " ", "") & _
                       ConvertToJson(vValue(vKey), iIndent, iLevel + 1)
                bFirst = False
            Next vKey
            If iIndent > 0 And Not bFirst Then sOut = sOut & vbNewLine & sPad
            sOut = sOut & "}"
        Case "Collection"
            bFirst = True
            sOut = "["
            Dim vItem As Variant
            For Each vItem In vValue
                If Not bFirst Then sOut = sOut & ","
                If iIndent > 0 Then sOut = sOut & vbNewLine & sPad1
                sOut = sOut & ConvertToJson(vItem, iIndent, iLevel + 1)
                bFirst = False
            Next vItem
            If iIndent > 0 And Not bFirst Then sOut = sOut & vbNewLine & sPad
            sOut = sOut & "]"
        Case "Boolean"
            sOut = IIf(vValue, "true", "false")
        Case "Date"
            sOut = """" & Format(vValue, "YYYY-MM-DD HH:MM:SS") & """"
        Case "String"
            sOut = """" & EscapeJson(CStr(vValue)) & """"
        Case Else
            If IsNumeric(vValue) Then
                sOut = CStr(vValue)
                ' Ensure decimal point is "." regardless of locale
                sOut = Replace(sOut, ",", ".")
            Else
                sOut = """" & EscapeJson(CStr(vValue)) & """"
            End If
    End Select

    ConvertToJson = sOut
End Function

' ---- Internal parser ---------------------------------------------------------


Private Function ParseObject(ByRef s As String, ByRef pos As Long) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.CompareMode = vbTextCompare    ' case-insensitive keys
    pos = pos + 1   ' skip {
    SkipWs s, pos

    If Mid(s, pos, 1) = "}" Then
        pos = pos + 1
        Set ParseObject = dict
        Exit Function
    End If

    Do
        SkipWs s, pos
        Dim sKey As String
        sKey = ParseString(s, pos)
        SkipWs s, pos
        pos = pos + 1   ' skip :
        SkipWs s, pos
        Dim vv As Variant
        vv = ParseValueIntoVariant(s, pos)
        If IsObject(vv) Then
            Set dict(sKey) = vv
        Else
            dict(sKey) = vv
        End If
        SkipWs s, pos
        If Mid(s, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        End If
        pos = pos + 1   ' skip ,
    Loop

    Set ParseObject = dict
End Function

Private Function ParseArray(ByRef s As String, ByRef pos As Long) As Collection
    Dim col As New Collection
    pos = pos + 1   ' skip [
    SkipWs s, pos

    If Mid(s, pos, 1) = "]" Then
        pos = pos + 1
        Set ParseArray = col
        Exit Function
    End If

    Do
        SkipWs s, pos
        Dim vv As Variant
        vv = ParseValueIntoVariant(s, pos)
        If IsObject(vv) Then
            col.Add vv
        Else
            col.Add vv
        End If
        SkipWs s, pos
        If Mid(s, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        End If
        pos = pos + 1   ' skip ,
    Loop

    Set ParseArray = col
End Function

' Wrapper that returns either a primitive or an object without losing the Set requirement
Private Function ParseValueIntoVariant(ByRef s As String, ByRef pos As Long) As Variant
    SkipWs s, pos
    If pos > Len(s) Then Exit Function

    Dim ch As String
    ch = Mid(s, pos, 1)

    Select Case ch
        Case "{"
            Set ParseValueIntoVariant = ParseObject(s, pos)
        Case "["
            Set ParseValueIntoVariant = ParseArray(s, pos)
        Case """"
            ParseValueIntoVariant = ParseString(s, pos)
        Case "t"
            pos = pos + 4
            ParseValueIntoVariant = True
        Case "f"
            pos = pos + 5
            ParseValueIntoVariant = False
        Case "n"
            pos = pos + 4
            ParseValueIntoVariant = Null
        Case Else
            ParseValueIntoVariant = ParseNumber(s, pos)
    End Select
End Function

Private Function ParseString(ByRef s As String, ByRef pos As Long) As String
    pos = pos + 1   ' skip opening "
    Dim sOut As String
    sOut = ""

    Do While pos <= Len(s)
        Dim ch As String
        ch = Mid(s, pos, 1)
        If ch = """" Then
            pos = pos + 1
            Exit Do
        ElseIf ch = "\" Then
            pos = pos + 1
            Dim esc As String
            esc = Mid(s, pos, 1)
            Select Case esc
                Case """": sOut = sOut & """"
                Case "\": sOut = sOut & "\"
                Case "/": sOut = sOut & "/"
                Case "b": sOut = sOut & Chr(8)
                Case "f": sOut = sOut & Chr(12)
                Case "n": sOut = sOut & vbLf
                Case "r": sOut = sOut & vbCr
                Case "t": sOut = sOut & vbTab
                Case "u"
                    Dim hexCode As String
                    hexCode = Mid(s, pos + 1, 4)
                    sOut = sOut & ChrW(CLng("&H" & hexCode))
                    pos = pos + 4
                Case Else: sOut = sOut & esc
            End Select
        Else
            sOut = sOut & ch
        End If
        pos = pos + 1
    Loop

    ParseString = sOut
End Function

Private Function ParseNumber(ByRef s As String, ByRef pos As Long) As Variant
    Dim start As Long
    start = pos
    Do While pos <= Len(s)
        Dim ch As String
        ch = Mid(s, pos, 1)
        If InStr("0123456789+-eE.", ch) = 0 Then Exit Do
        pos = pos + 1
    Loop
    Dim sNum As String
    sNum = Mid(s, start, pos - start)
    ' Normalize to local decimal sep for Val() / CDbl()
    Dim sDecSep As String
    sDecSep = Mid(CStr(0.5), 2, 1)   ' locale-safe: "0.5" -> "."
    If sDecSep <> "." Then sNum = Replace(sNum, ".", sDecSep)
    If InStr(sNum, sDecSep) > 0 Or InStr(LCase(sNum), "e") > 0 Then
        ParseNumber = CDbl(sNum)
    Else
        ParseNumber = CLng(sNum)
    End If
End Function

Private Sub SkipWs(ByRef s As String, ByRef pos As Long)
    Do While pos <= Len(s) And InStr(" " & vbTab & vbCr & vbLf, Mid(s, pos, 1)) > 0
        pos = pos + 1
    Loop
End Sub

' ---- Escape helpers ----------------------------------------------------------

Private Function EscapeJson(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function
