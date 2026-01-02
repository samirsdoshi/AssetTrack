Public Function GetCell(strRetVal, cellNumber)
    i = InStr(UCase(strRetVal), cellNumber & "|")
    If i = 0 Then
        GetCell = ""
        Exit Function
    End If
    r = InStrRev(UCase(strRetVal), ">", i)
    
    StartCellText = r + 1

   ' Now... to find the end of this cell's text, we look for either <TABLE
   ' or <TD - whichever comes first (but we have to check if they exist or not)
   ' We don't include nested tables in the cell data because those tables have
   ' cells of their own.
   If (InStr(r, UCase(strRetVal), "<TABLE") > 0) And _
         (InStr(r, UCase(strRetVal), "<TABLE") < _
              InStr(r, UCase(strRetVal), "</TD>")) Then
      thiscelltext = Mid(strRetVal, StartCellText, _
               InStr(r, UCase(strRetVal), "<TABLE") - StartCellText)
   Else
      thiscelltext = Mid(strRetVal, StartCellText + Len(cellNumber & "|"), _
               InStr(r, UCase(strRetVal), "</TD>") - StartCellText - Len(cellNumber & "|"))
   End If
   GetCell = thiscelltext
   
End Function
Public Function numbercell(strResult, ByRef maxcell)
        q = 1
        i = 1
        Do While InStr(i, UCase(strResult), "<TD") > 0
               ' find next <TD
               i = InStr(i, UCase(strResult), "<TD")
               
               ' fomd the end of the <TD
               r = InStr(i, strResult, ">")
               strResult = Left(strResult, r) & q & "|" & _
                     Right(strResult, Len(strResult) - r)
        
               'Number the cells: the string equals all the html we've check,
               'our cell number, and then the html we've yet to check
               ' Let the next loop start looking after this <TD tag we found
               i = r + 1
         
               ' increase the count of which cell we're at
               q = q + 1
        Loop
        maxcell = q
        numbercell = strResult
End Function
Public Function cleanUp(strval)
    strval = Replace(strval, "<b>", "")
    strval = Replace(strval, "</b>", "")
    strval = Replace(strval, "&nbsp;", "")
    strval = Replace(strval, "$", "")
    strval = Replace(strval, "</strong>", "")
    strval = Replace(strval, "<strong>", "")
    strval = Replace(strval, "<font color=" & Chr(34) & "red" & Chr(34) & ">", "")
    strval = Replace(strval, "</font>", "")
    strval = Replace(strval, "&mdash;", "")
   cleanUp = strval 'RemoveWhiteSpace(strval)
End Function
Public Function checkNum(strval)
    If IsNumeric(strval) Then
        checkNum = strval
    Else
        checkNum = "0"
    End If
End Function
Function RemoveWhiteSpace(strText)
 Dim RegEx
 Set RegEx = New RegExp
 RegEx.Pattern = "\s+"
 RegEx.MultiLine = True
 RegEx.Global = True
 strText = RegEx.Replace(strText, " ")
 RemoveWhiteSpace = strText
End Function
Public Function min(n1, n2)
    min = n1
    If (n2 < n1) Then
        min = n2
    End If
End Function

Public Function max(n1, n2)
    max = n1
    If (n2 > n1) Then
        max = n2
    End If
End Function

Public Function URLEncode(sRawURL As String) As String
    On Error GoTo Catch
    Dim iLoop As Integer
    Dim sRtn As String
    Dim sTmp As String
    Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/.?=_-$(){}~&"


    If Len(sRawURL) > 0 Then
        ' Loop through each char


        For iLoop = 1 To Len(sRawURL)
            sTmp = Mid(sRawURL, iLoop, 1)


            If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                ' If not ValidChar, convert to HEX and p
                '     refix with %
                sTmp = Hex(Asc(sTmp))


                If sTmp = "20" Then
                    sTmp = "+"
                ElseIf Len(sTmp) = 1 Then
                    sTmp = "%0" & sTmp
                Else
                    sTmp = "%" & sTmp
                End If
            End If
            sRtn = sRtn & sTmp
        Next iLoop
        URLEncode = sRtn
    End If
Finally:
    Exit Function
Catch:
    URLEncode = ""
    Resume Finally
End Function
Sub createfile(fpath)
    Set g_fs = CreateObject("Scripting.FileSystemObject")
    Set g_fh = g_fs.CreateTextFile(fpath, True)
End Sub

Sub writetofile(msg)
    g_fh.WriteLine (msg & vbCrLf)
End Sub

Sub closefile()
   g_fh.Flush
   g_fh.Close
End Sub

Public Function nullif(v, defvalue)
    If IsNull(v) Then
        nullif = defvalue
    Else
        nullif = v
    End If
End Function
Public Function empty2def(v, defvalue)
    empty2def = defvalue
    If Not IsEmpty(v) Then
        If Len(v) > 0 Then
            empty2def = v
        End If
    End If
End Function
Public Function Lpad(MyValue, MyPadChar, MyPaddedLength)
    If (Len(MyValue) < MyPaddedLength) Then
        Lpad = String(MyPaddedLength - Len(MyValue), MyPadChar) & MyValue
    Else
        Lpad = MyValue
    End If
End Function
Public Function Rpad(MyValue, MyPadChar, MyPaddedLength)
    If (Len(MyValue) < MyPaddedLength) Then
        Rpad = MyValue & String(MyPaddedLength - Len(MyValue), MyPadChar)
    Else
        Rpad = MyValue
    End If
End Function
