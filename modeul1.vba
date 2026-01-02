Public aconn As ADODB.Connection

Public Sub reallocate(assetid, asofdate, amount)
    Call opendb
    On Error Resume Next
    strsql = "select * from assetinv where assetid=" & assetid & " and asofdate='" & mysqldate(asofdate, vbShortDate) & "' and amount=" & amount & " order by assetinvid desc"
    Set rs = aconn.Execute(strsql, , 1)
    assetinvid = "0"
    While Not rs.EOF
        assetinvid = assetinvid & "," & rs.fields("assetinvid").Value
        rs.movenext
    Wend
    rs.Close
    
    aconn.begintrans
    aconn.Execute "delete from assetinvalloc where assetinvid in (" & assetinvid & ")", , 1
    aconn.Execute "delete from assetinvsecind where assetinvid in (" & assetinvid & ")", , 1
    aconn.Execute "delete from assetinvinter where assetinvid in (" & assetinvid & ")", , 1
    aconn.Execute "delete from assetinv where assetinvid in (" & assetinvid & ")", , 1
    If Err.Number = 0 Then
        aconn.committrans
    Else
        aconn.rollbacktrans
    End If
    
    Call allocate(assetid, asofdate, amount)
End Sub
Public Sub allocate(assetid, asofdate, amount)
    
    Call opendb
    On Error Resume Next
    aconn.begintrans
    strsql = "insert into assetinv(assetid, asofdate, amount) values (" & assetid & ",'" & mysqldate(asofdate, vbShortDate) & "'," & amount & ")"
    aconn.Execute strsql, , 1
    
    strsql = "select top 1 * from assetinv order by assetinvid desc"
    Set rs = aconn.Execute(strsql, , 1)
    assetinvid = rs.fields("assetinvid").Value
    rs.Close
    
    strsql = "select tcode, tval1, tval2, prct from templatedetails td inner join asset a on td.templateid=a.templateid where a.assetid=" & assetid
    Set rs = aconn.Execute(strsql, , 1)
    While Not rs.EOF
        
        tcode = rs.fields("tcode").Value
        tval1 = nullif(rs.fields("tval1").Value, "0")
        tval2 = nullif(rs.fields("tval2").Value, "0")
        prct = rs.fields("prct").Value
        Select Case LCase(tcode)
            Case "alloc"
                strsql = "insert into assetinvalloc(assetinvid,alloccode,amount) values(" & assetinvid & "," & CInt(tval1) & "," & Round(amount * (prct / 100), 2) & ")"
                aconn.Execute strsql, , 1
            Case "secind"
                strsql = "insert into assetinvsecind(assetinvid,sec_id,ind_id,amount) values(" & assetinvid & "," & CInt(tval1) & "," & CInt(tval2) & "," & Round(amount * (prct / 100), 2) & ")"
                aconn.Execute strsql, , 1
                                
            Case "inter"
                strsql = "insert into assetinvinter(assetinvid,intercode,amount) values(" & assetinvid & "," & CInt(tval1) & "," & Round(amount * (prct / 100), 2) & ")"
                aconn.Execute strsql, , 1
        End Select
        rs.movenext
    Wend
    If Err.Number = 0 Then
        aconn.committrans
    Else
        Range("l2").Value = Err.Description
        aconn.rollbacktrans
    End If
    Call closedb
End Sub

Public Sub opendb()
    If aconn Is Nothing Then
        Set aconn = CreateObject("adodb.connection")
        aconn.Open "mysql57", "root", "sa123"
        aconn.Execute "use asset"
    End If
End Sub
Public Sub closedb()
    On Error Resume Next
    aconn.Close
    Set aconn = Nothing
End Sub



Public Sub addTemplateDetAlloc(templateid, arr)
    For i = 1 To UBound(arr)
        arr1 = Split(arr(i - 1), ",")
        strAllocName = arr1(0)
        allocPrct = CDbl(arr1(1))
        Set rs = aconn.Execute("select alloccode from alloctype where allocdesc='" & strAllocName & "'", , 1)
        alloccode = rs.fields("alloccode").Value
        rs.Close
        
        strsql = "insert into templatedetails(templateid, tcode, tval1, tval2, prct) values("
        strsql = strsql & templateid & ",'alloc'," & alloccode & ",''," & allocPrct & ")"
        aconn.Execute strsql, , 1
        
    Next
End Sub
Public Sub addTemplateDetInter(templateid, arr)
    For i = 1 To UBound(arr)
        arr1 = Split(arr(i - 1), ",")
        strInterName = arr1(0)
        allocPrct = CDbl(arr1(1))
        
        Set rs = aconn.Execute("select intercode from inter where inter_name='" & strInterName & "'", , 1)
        intercode = rs.fields("intercode").Value
        rs.Close
        
        strsql = "insert into templatedetails(templateid, tcode, tval1, tval2, prct) values("
        strsql = strsql & templateid & ",'inter'," & intercode & ",''," & allocPrct & ")"
        aconn.Execute strsql, , 1
    
    Next

End Sub
Public Sub addTemplateDetSecInd(templateid, arr)
    For i = 1 To UBound(arr)
        arr1 = Split(arr(i - 1), ",")
        strSector = arr1(0)
        strInd = arr1(1)
        allocPrct = CDbl(arr1(2))
        
        Set rs = aconn.Execute("select sec_id from sector where sec_name='" & strSector & "'", , 1)
        sec_id = rs.fields("sec_id").Value
        rs.Close
        
        If (Len(strInd) > 0 And strInd <> "0") Then
            Set rs = aconn.Execute("select ind_id from industry where ind_name='" & strInd & "'", , 1)
            ind_id = rs.fields("ind_id").Value
            rs.Close
        Else
            ind_id = 0
        End If
        
        strsql = "insert into templatedetails(templateid, tcode, tval1, tval2, prct) values("
        strsql = strsql & templateid & ",'secind'," & sec_id & "," & ind_id & "," & allocPrct & ")"
        aconn.Execute strsql, , 1
    
    Next

End Sub
Public Sub deleteTemplateDetails(templateid)
    aconn.Execute ("delete from templatedetails where templateid=" & templateid)
End Sub
Public Sub normalizefullview()
    Sheets("fullview").Select
    Call opendb
    j = 1
    Range("k1:n2000").Clear
    curraccount = ""
    oldaccount = ""
    Set objDictionary = CreateObject("Scripting.Dictionary")
    For i = 1 To 2000
        If (Range("a" & i).Value = "") Then
            Exit For
        End If
        account = Range("a" & i).Value
        fund = LCase(Range("a" & i + 1).Value)
        curraccount = ""
        If (InStr(account, "CollegeAdv") > 0) Then
            curraccount = "CollegeAdv"
        End If
        If (InStr(account, "- Individual") > 0 And InStr(fund, "rowe") > 0 And InStr(fund, "price") > 0) Then
            curraccount = "TRPInv"
        End If
        If (InStr(account, "Rollover IRA") > 0 And InStr(fund, "rowe") > 0 And InStr(fund, "price") > 0) Then
            curraccount = "TRPRollover"
        End If
        If (InStr(account, "Roth IRA") > 0 And InStr(fund, "rowe") > 0 And InStr(fund, "price") > 0) Then
            curraccount = "TRPRoth"
        End If
        If (InStr(account, "Traditional IRA") > 0 And InStr(fund, "rowe") > 0 And InStr(fund, "price") > 0) Then
            curraccount = "TRPRollover"
            Rem Sangeetas account
        End If
        If (InStr(account, "GAPSHARE 401(K) PLAN") > 0) Then
            curraccount = "TRPRps"
        End If
        If (InStr(account, "Individual - TOD") > 0) Then
            curraccount = "FidelityInv"
        End If
        If (Len(curraccount) = 0 And InStr(account, "Rollover IRA") > 0) Then
            curraccount = "FidelityIRA"
        End If
        If (InStr(account, "Brokerage Account - 10498558") > 0) Then
            curraccount = "Vanguard"
        End If
        If (InStr(account, "Brokerage Account - 10498558") > 0) Then
            curraccount = "Vanguard"
        End If
        If (InStr(account, "Brokerage Account - 10498558") > 0) Then
            curraccount = "Vanguard"
        End If
        If (InStr(account, "Samir S Doshi - Rollover IRA") > 0) Then
            curraccount = "Vanguard IRA"
        End If
        If (curraccount = "" And oldaccount <> "") Then
            fund = Range("a" & i).Value
            ticker = Range("b" & i).Value
            amount = Range("f" & i).Value
            Debug.Print oldaccount & ":" & ticker & ":"; amount
            Range("k" & j).Value = oldaccount & "_" & ticker
            Range("l" & j).Value = amount
            If (Not objDictionary.exists(curraccount)) Then
                objDictionary.Item(curraccount) = 0
            End If
            objDictionary.Item(curraccount) = objDictionary.Item(curraccount) + amount
            j = j + 1
        Else
            oldaccount = curraccount
        End If
    Next
    k = 8
    For Each objKey In objDictionary
        Range("N" & k) = objKey
        Range("O" & k) = objDictionary.Item(objKey)
        k = k + 1
    Next
End Sub

Sub ReallocateAssetRef()
    Sheets("assetref").Select
    Call opendb
    For i = 1 To 2000
        ticker = Range("a" & i).Value
        heldat = Range("J" & i).Value
        Debug.Print ticker
        If ticker = "ENDOFPORTFOLIO" Then
            Exit For
        End If
        ticker = filterTicker(ticker)
        If Len(ticker) > 0 Then
            Rem heldAt = getHeldAt(ticker)
            If Len(heldat) > 0 Then
                strsql = "select assetid from asset where (ticker='" & ticker & "'  or assetname = '" & ticker & "') "
                Set rs = aconn.Execute(strsql, , 1)
                assetid = 0
                If Not rs.EOF Then
                    assetid = rs.fields("assetid").Value
                End If
                rs.Close
                asofdate = Range("N4").Value
                amount = Range("e" & i).Value
                amount = Replace(amount, "$", "")
                amount = Replace(amount, ",", "")
                If assetid <> 0 Then
                    Range("a" & i).Font.Color = vbRed
                    Call allocateassetref(assetid, asofdate, amount, heldat)
                    Range("a" & i).Font.Color = vbBlack
                Else
                    Debug.Print "***** ASSET NOT FOUND ****** " & ticker
                    Range("a" & i).Font.Color = vbRed
                End If
            Else
                    Debug.Print "***** ASSET NOT FOUND ****** " & ticker
                    Range("a" & i).Font.Color = vbRed
            End If
        End If
    Next
    Call closedb
    MsgBox "done"

End Sub

Private Function getHeldAt(ticker)
        getHeldAt = ""
        For i = 1 To Len(ticker)
            If Asc(Mid(ticker, i, 1)) = 160 Then
                ticker = Mid(ticker, 1, i - 1) & " " & Mid(ticker, i + 1, Len(ticker) - i + 1)
            End If
        Next

        If ticker = "CollegeAdv" Or ticker = "FidelityInv" _
        Or ticker = "FidelityIRA" Or ticker = "FidelityRoth" Or ticker = "TRPInv" _
        Or ticker = "TRPRoth" Or ticker = "Vanguard" Or ticker = "Fidelity401k" _
        Or ticker = "WellsFargo401k" Then
            getHeldAt = ticker
        End If
        If ticker = "CollegeAdvantage 529 Savings Plan - xxx2686" Then
            getHeldAt = "CollegeAdv"
        End If
        If ticker = "Fidelity Investments (owned) - INDIVIDUAL - TOD" Then
            getHeldAt = "FidelityInv"
        End If
        If ticker = "Fidelity Investments (owned) - ROLLOVER IRA" Then
            getHeldAt = "FidelityIRA"
        End If
        If ticker = "Fidelity Investments (owned) - ROTH IRA" Then
            getHeldAt = "FidelityRoth"
        End If
        If ticker = "T. Rowe Price - Investments - Individual" Then
            getHeldAt = "TRPInv"
        End If
        If ticker = "T. Rowe Price - Investments - Roth IRA" Then
            getHeldAt = "TRPRoth"
        End If
        If ticker = "Vanguard Investments - Samir S Doshi" Then
            getHeldAt = "Vanguard"
        End If
        If ticker = "Vanguard Investments - Brokerage Account" Then
            getHeldAt = "Vanguard"
        End If
        If ticker = "Fidelity NetBenefits (owned) - Investments - HP 401(K) PLAN" Then
            getHeldAt = "Fidelity401k"
        End If
        If ticker = "Wells Fargo Retirement Services - Limited Brands 401k Savings & Retirement Plan" Then
            getHeldAt = "WellsFargo401k"
        End If
        If ticker = "Etrade" Or ticker = "Robinhood" Or ticker = "Ameritrade" Or ticker = "TradeStation" Then
            getHeldAt = ticker
        End If
End Function

Public Sub allocateassetref(assetid, asofdate, amount, heldat)
    Sheets("assetref").Select
    
    dupasset = False
    If amount = 0 Then
        Exit Sub
    End If
    aconn.begintrans
    strsql = "select assetinvid, amount from assetinv where assetid=" & assetid & " and asofdate='" & mysqldate(asofdate, vbShortDate) & "' and heldat='" & heldat & "'"
    Set rs = aconn.Execute(strsql, , 1)
    If Not rs.EOF Then
        assetinvid = rs.fields("assetinvid").Value
        orgamount = rs.fields("amount").Value
        rs.Close
        dupasset = True
        If (orgamount <> amount) Then
            aconn.Execute "update assetinv set amount=amount + " & amount & " where assetinvid=" & assetinvid
        End If
    Else
        strsql = "select max(assetinvid) from assetinv"
        Set rs = aconn.Execute(strsql, , 1)
        assetinvid = rs.fields(0).Value + 1
        rs.Close
        strsql = "insert into assetinv(assetinvid, assetid, asofdate, amount, heldat) values (" & assetinvid & "," & assetid & ",'" & mysqldate(asofdate, vbShortDate) & "'," & amount & ",'" & heldat & "')"
        aconn.Execute strsql, , 1
    End If
    
   
    strsql = "select tcode, tval1, tval2, prct from templatedetails td inner join asset a on td.templateid=a.templateid where a.assetid=" & assetid
    Set rs = aconn.Execute(strsql, , 1)
    If rs.EOF Then
        Err.Raise 100
    End If
    While Not rs.EOF
        
        tcode = rs.fields("tcode").Value
        tval1 = nullif(rs.fields("tval1").Value, "0")
        tval2 = nullif(rs.fields("tval2").Value, "0")
        prct = rs.fields("prct").Value
        Select Case LCase(tcode)
            Case "alloc"
                If dupasset Then
                    strsql = "update assetinvalloc set amount=amount+ " & Round(amount * (prct / 100), 2) & " where assetinvid=" & assetinvid
                    strsql = strsql & " and alloccode=" & tval1
                    aconn.Execute strsql, , 1
                Else
                    strsql = "insert into assetinvalloc(assetinvid,alloccode,amount) values(" & assetinvid & "," & CInt(tval1) & "," & Round(amount * (prct / 100), 2) & ")"
                    aconn.Execute strsql, , 1
                End If
            Case "secind"
                If dupasset Then
                    strsql = "update assetinvsecind set amount=amount+ " & Round(amount * (prct / 100), 2) & " where assetinvid=" & assetinvid
                    strsql = strsql & " and sec_id=" & CInt(tval1) & " and ind_id=" & CInt(tval2)
                    aconn.Execute strsql, , 1
                Else
                    strsql = "insert into assetinvsecind(assetinvid,sec_id,ind_id,amount) values(" & assetinvid & "," & CInt(tval1) & "," & CInt(tval2) & "," & Round(amount * (prct / 100), 2) & ")"
                    aconn.Execute strsql, , 1
                End If
            Case "inter"
                If dupasset Then
                    strsql = "update assetinvinter set amount=amount+ " & Round(amount * (prct / 100), 2) & " where assetinvid=" & assetinvid
                    strsql = strsql & " and intercode=" & CInt(tval1)
                    aconn.Execute strsql, , 1
                Else
                    strsql = "insert into assetinvinter(assetinvid,intercode,amount) values(" & assetinvid & "," & CInt(tval1) & "," & Round(amount * (prct / 100), 2) & ")"
                    aconn.Execute strsql, , 1
                End If
        End Select
        rs.movenext
    Wend
    If Err.Number = 0 Then
        aconn.committrans
    Else
        Range("N5").Value = Err.Description
        aconn.rollbacktrans
    End If

End Sub
Function mysqldate(dt, opt)
    mysqldate = "" & Year(dt) & "-" & Month(dt) & "-" & Day(dt)
End Function
Sub deleteAssetInfo()
    Sheets("assetref").Select
    Call opendb
    asofdate = Range("N4").Value
    
    strsql = "delete from assetgain where assetdate='" & mysqldate(asofdate, vbShortDate) & "'"
    aconn.Execute strsql, , 1
    strsql = "delete from assetgain where assetdate<'" & mysqldate(DateAdd("m", -24, asofdate), vbShortDate) & "'"
    aconn.Execute strsql, , 1
    strsql = "select * from assetinv where asofdate='" & mysqldate(asofdate, vbShortDate) & "'"
    Set rs = aconn.Execute(strsql, , 1)
    While Not rs.EOF
        assetinvid = rs.fields("assetinvid").Value
        aconn.Execute "delete from assetinvalloc where assetinvid=" & assetinvid, , 1
        aconn.Execute "delete from assetinvsecind where assetinvid=" & assetinvid, , 1
        aconn.Execute "delete from assetinvinter where assetinvid=" & assetinvid, , 1
        aconn.Execute "delete from assetinv where assetinvid=" & assetinvid, , 1
        rs.movenext
    Wend
    rs.Close
    MsgBox "done"
End Sub


Function filterTicker(ticker)
    filterTicker = Trim(ticker)
    If (InStr(ticker, "Go to Site |") > 0) Then
        filterTicker = ""
    End If
    If (Mid(ticker, 1, 6) = "Symbol") Then
        filterTicker = ""
    End If
    If (Mid(ticker, 1, 5) = "Total") Then
        filterTicker = ""
    End If
    If (Mid(ticker, 1, 5) = "samir") Then
        filterTicker = ""
    End If
End Function

Sub calcGain()
    Dim cmd As ADODB.Command
    Dim p1 As ADODB.Parameter
    Dim rsGain As ADODB.Recordset
    
    Sheets("assetref").Select
    Err.Clear
    Call closedb
    Call opendb
    
    strsql = "delete from assetgain where assetdate='" & Range("N4").Value & "'"
        aconn.Execute strsql, , 1
    
    Rem get gains for funds/stocks
    dtToday = CDate(Range("N4").Value)
    If Weekday(dtToday) = 1 Then
        dtToday = DateAdd("d", -2, dtToday)
    Else
        If Weekday(dtToday) = 7 Then
            dtToday = DateAdd("d", -1, dtToday)
        End If
    End If
    Dim arr(6)
    arr(0) = dtToday
    arr(1) = tgDateAdd("ww", -1, dtToday)
    arr(2) = tgDateAdd("ww", -2, dtToday)
    arr(3) = tgDateAdd("ww", -4, dtToday)
    arr(4) = tgDateAdd("ww", -12, dtToday)
    arr(5) = tgDateAdd("ww", -24, dtToday)
    arr(6) = tgDateAdd("ww", -52, dtToday)
    Rem strsql = "select distinct ticker from asset a inner join assetinv b on a.assetid=b.assetid where asofdate='" & Range("N4").Value & "' and benchmark<>''"
    strsql = "select distinct ticker from asset a where benchmark<>''"
    Set rs = aconn.Execute(strsql, , 1)
    While Not rs.EOF
        If Not calcgainfrommorningstar(rs.fields("ticker").Value, arr) Then
            Call calcgainfromyahoo(rs.fields("ticker").Value, arr)
        End If
        rs.movenext
    Wend
    
    Call calcBenchmarkGain_Click
    
    Call closedb
    MsgBox "done"

End Sub
Public Sub testGain()
   Dim arr(6)
       Call openDailyDB
    Call opendb
    dtToday = CDate("09/12/2014")
    arr(0) = dtToday
    arr(1) = tgDateAdd("ww", -1, dtToday)
    arr(2) = tgDateAdd("ww", -2, dtToday)
    arr(3) = tgDateAdd("ww", -4, dtToday)
    arr(4) = tgDateAdd("ww", -12, dtToday)
    arr(5) = tgDateAdd("ww", -24, dtToday)
    arr(6) = tgDateAdd("ww", -52, dtToday)
    Call calcgainfromyahoo("FNARX", arr)
     
End Sub
Private Sub objHttpsend(objHttp)
    On Error Resume Next
    Err.Clear
    For i = 1 To 5
        objHttp.send
        If Err.Number = 0 Then
            Exit Sub
        End If
    Next
    On Error GoTo 0
End Sub
Public Function calcgainfrommorningstar(ticker, arr)
  Dim gain(6)
  calcgainfrommorningstar = False
    Debug.Print ticker
    If (ticker = "VDMIX") Then
        kk = 1
    End If
    Set objHttp = CreateObject("Microsoft.XMLHTTP")
    strurl = "http://performance.morningstar.com/Performance/fund/trailing-total-returns.action?t=" & ticker & "&ops=clear"
    Debug.Print strurl
    objHttp.Open "GET", strurl, False
    Call objHttpsend(objHttp)
    strResult = objHttp.responsetext
    strResult = numbercell(strResult, maxcell)
    For k = 1 To maxcell
         cellval = GetCell(strResult, k)
         If InStr(cellval, "1-Day") > 0 Then
            oneweek = GetCell(strResult, k + 12)
            onemonth = GetCell(strResult, k + 13)
            threemonth = GetCell(strResult, k + 14)
            oneyear = GetCell(strResult, k + 16)
            Exit For
         End If
         
    Next
    If empty2def(cleanUp(oneweek), "0") = "0" And empty2def(cleanUp(onemonth), "0") = "0" And _
        empty2def(cleanUp(threemonth), "0") = "0" And empty2def(cleanUp(oneyear), "0") = "0" Then
        calcgainfrommorningstar = False
    Else
        
        strsql = "insert into AssetGain values('"
        strsql = strsql & ticker & "','"
        strsql = strsql & Range("N4").Value & "',"
        strsql = strsql & empty2def(cleanUp(oneweek), 0) & ","
        strsql = strsql & "0,"
        strsql = strsql & empty2def(cleanUp(onemonth), 0) & ","
        strsql = strsql & empty2def(cleanUp(threemonth), 0) & ","
        strsql = strsql & "0,"
        strsql = strsql & empty2def(cleanUp(oneyear), 0) & ")"
        Debug.Print strsql
         aconn.Execute strsql, , 1
        calcgainfrommorningstar = True
    End If
Set objHttp = Nothing

End Function
Public Sub calcgainfromyahoo(ticker, arr)
    Dim gain(6)
    On Error GoTo han_err:
    Debug.Print ticker
    If ticker = "FCASH" Then
        Exit Sub
    End If
    Set objHttp = CreateObject("Microsoft.XMLHTTP")
    Rem strurl = "http://finance.yahoo.com/q/hp?s=" & rs.fields("ticker").Value & "&a=" & Month(arr(i + 1)) - 1 & "&b=" & Day(arr(i + 1)) & "&c=" & Year(arr(i + 1)) & "&d=" & Month(arr(i)) - 1 & "&e=" & Day(arr(i)) & "&f=" & Year(arr(i)) & "&g=d"
    strurl = "http://ichart.finance.yahoo.com/table.csv?s=" & ticker & "&a=" & Month(arr(UBound(arr))) - 1 & "&b=" & Day(arr(UBound(arr))) & "&c=" & Year(arr(UBound(arr))) & "&d=" & Month(arr(0)) - 1 & "&e=" & Day(arr(0)) & "&f=" & Year(arr(0)) & "&g=d&ignore=.csv"
    Debug.Print strurl
    objHttp.Open "GET", strurl, False
    objHttp.send
    strResult = objHttp.responsetext
    If (Mid(strResult, 1, Len("<!doctype")) = "<!doctype") Then
        Exit Sub
    End If
    resarr = Split(strResult, Chr(10))
    For i = 0 To UBound(arr)
        For j = 1 To UBound(resarr)
            If (Len(resarr(j) > 0)) Then
                tarr = Split(resarr(j), ",")
                If UBound(tarr) > 0 Then
                    If CDate(tarr(0)) = CDate(arr(i)) Then
                        If (i = 0) Then
                            currprice = tarr(4)
                        Else
                           dprice2 = tarr(4)
                           gain(i) = Round(((currprice - dprice2) / dprice2) * 100, 2)
                        End If
                        Exit For
                    End If
                End If
            End If
        Next
        'index1 = InStr(strResult, Year(arr(i)) & "-" & Lpad(Month(arr(i)), "0", 2) & "-" & Lpad(Day(arr(i)), "0", 2))
        'index2 = InStr(strResult, Year(arr(i + 1)) & "-" & Lpad(Month(arr(i + 1)), "0", 2) & "-" & Lpad(Day(arr(i + 1)), "0", 2))
        'If (index1 > 0 And index2 > 0) Then
        '    index11 = index1
        '    index22 = index2
        '    For k = 1 To 4
        '        index11 = InStr(index11, strResult, ",")
        '        index22 = InStr(index22, strResult, ",")
        '    Next
        '    dprice1 = CDbl(Mid(strResult, index11 + 1, InStr(index11 + 1, strResult, ",") - index11 - 1))
        '    If (i = 0) Then
        '        currprice = dprice1
        '    End If
        '    dprice2 = CDbl(Mid(strResult, index22 + 1, InStr(index22 + 1, strResult, ",") - index22 - 1))
        '    gain(i) = Round(((currprice - dprice2) / dprice2) * 100, 2)
        'End If
    Next
    strsql = "insert into AssetGain values('"
    strsql = strsql & ticker & "','"
    strsql = strsql & Range("N4").Value & "',"
    strsql = strsql & empty2def(gain(1), 0) & ","
    strsql = strsql & empty2def(gain(2), 0) & ","
    strsql = strsql & empty2def(gain(3), 0) & ","
    strsql = strsql & empty2def(gain(4), 0) & ","
    strsql = strsql & empty2def(gain(5), 0) & ","
    strsql = strsql & empty2def(gain(6), 0) & ")"
    Debug.Print strsql
     aconn.Execute strsql, , 1
han_err:
     Set objHttp = Nothing
     Exit Sub
End Sub
Rem obsolete. dailystockhist in iportdb does not contain 6 months of data
Public Sub old_calcBenchmarkGain_Click()
   Rem get gains for benchmarks
    Call closedb
    Call openDailyDB
    Call opendb
    Sheets("assetref").Select
    Set rs = aconn.Execute("select distinct benchmark from asset", , 1)
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = dailydbconn
    cmd.CommandType = 1
    cmd.CommandText = "select * from [_StkGain]"
    While Not rs.EOF
        If Not IsNull(rs.fields("benchmark").Value) Then
            Set p1 = New ADODB.Parameter
            p1.Type = 200
            p1.Direction = &H1
            p1.Name = "@ticker"
            p1.Size = 10
            p1.Value = rs.fields("benchmark").Value
            
            Set p2 = New ADODB.Parameter
            p2.Type = 7
            p2.Direction = &H1
            p2.Name = "@parmdate"
            p2.Value = CDate(Range("N4").Value)
            p2.Size = 8
            
            cmd.Parameters.append p2
            cmd.Parameters.append p1
            
            aconn.Execute "delete from assetgain where ticker='" & rs.fields("benchmark").Value & "' and assetdate='" & Range("N4").Value & "'"
            Set rsGain = cmd.Execute
            If Not rsGain.EOF Then
                strsql = "insert into AssetGain values('"
                strsql = strsql & rs.fields("benchmark").Value & "','"
                strsql = strsql & Range("N4").Value & "',"
                strsql = strsql & empty2def(rsGain.fields(1).Value, "0") & ","
                strsql = strsql & empty2def(rsGain.fields(2).Value, "0") & ","
                strsql = strsql & empty2def(rsGain.fields(3).Value, "0") & ","
                strsql = strsql & empty2def(rsGain.fields(4).Value, "0") & ","
                strsql = strsql & empty2def(rsGain.fields(5).Value, "0") & ")"
                aconn.Execute strsql, , 1
            End If
            rsGain.Close
            cmd.Parameters.Delete 1
            cmd.Parameters.Delete 0
        End If
        rs.movenext
    Wend
    rs.Close
    MsgBox "done"
End Sub
Public Sub calcBenchmarkGain_Click()
   Rem get gains for benchmarks
    Call closedb
    Call opendb
    Sheets("assetref").Select
    dtToday = CDate(Range("N4").Value)
    If Weekday(dtToday) = 1 Then
        dtToday = DateAdd("d", -2, dtToday)
    Else
        If Weekday(dtToday) = 7 Then
            dtToday = DateAdd("d", -1, dtToday)
        End If
    End If
    Dim arr(6)
    arr(0) = dtToday
    arr(1) = tgDateAdd("ww", -1, dtToday)
    arr(2) = tgDateAdd("ww", -2, dtToday)
    arr(3) = tgDateAdd("ww", -4, dtToday)
    arr(4) = tgDateAdd("ww", -12, dtToday)
    arr(5) = tgDateAdd("ww", -24, dtToday)
    arr(6) = tgDateAdd("ww", -52, dtToday)
    
    Set rs = aconn.Execute("select distinct benchmark from asset where benchmark<>'' union select distinct ticker from asset  where benchmark is null", , 1)
    While Not rs.EOF
        If Not IsNull(rs.fields("benchmark").Value) Then
            Debug.Print rs.fields("benchmark").Value
            aconn.Execute "delete from assetgain where ticker='" & rs.fields("benchmark").Value & "' and assetdate='" & Range("N4").Value & "'"
           
            If Not calcgainfrommorningstar(rs.fields("benchmark").Value, arr) Then
                Call calcgainfromyahoo(rs.fields("benchmark").Value, arr)
            End If

        End If
        rs.movenext
    Wend
    rs.Close
    MsgBox "done"
End Sub

Sub test1()
    Call calcGainByTicker("PRDGX", CDate(Range("N4").Value))
    MsgBox "done"
End Sub
Sub test2()
    Call calcgainfrommorningstar("XLU", Null)
    MsgBox "done"
End Sub

Sub calcGainByTicker(ticker, orgDtToday)
    Dim cmd As ADODB.Command
    Dim p1 As ADODB.Parameter
    Dim rsGain As ADODB.Recordset
    
    Err.Clear
    Call closedb
    Call openDailyDB
    Call opendb
    
    strsql = "delete from assetgain where ticker='" & ticker & "' and assetdate='" & orgDtToday & "'"
    aconn.Execute strsql, , 1
    
    Rem get gains for funds/stocks
    Set objHttp = CreateObject("Microsoft.XMLHTTP")
    dtToday = orgDtToday
    If Weekday(dtToday) = 1 Then
        dtToday = DateAdd("d", -2, dtToday)
    Else
        If Weekday(dtToday) = 7 Then
            dtToday = DateAdd("d", -1, dtToday)
        End If
    End If
    Dim arr(5)
    arr(0) = dtToday
    arr(1) = tgDateAdd("ww", -1, dtToday)
    arr(2) = tgDateAdd("ww", -2, dtToday)
    arr(3) = tgDateAdd("ww", -4, dtToday)
    arr(4) = tgDateAdd("ww", -12, dtToday)
    arr(5) = tgDateAdd("ww", -24, dtToday)
    Dim gain(5)
    Rem strurl = "http://finance.yahoo.com/q/hp?s=" & rs.fields("ticker").Value & "&a=" & Month(arr(i + 1)) - 1 & "&b=" & Day(arr(i + 1)) & "&c=" & Year(arr(i + 1)) & "&d=" & Month(arr(i)) - 1 & "&e=" & Day(arr(i)) & "&f=" & Year(arr(i)) & "&g=d"
    strurl = "http://ichart.finance.yahoo.com/table.csv?s=" & ticker & "&a=" & Month(arr(UBound(arr))) - 1 & "&b=" & Day(arr(UBound(arr))) & "&c=" & Year(arr(UBound(arr))) & "&d=" & Month(arr(0)) - 1 & "&e=" & Day(arr(0)) & "&f=" & Year(arr(0)) & "&g=d&ignore=.csv"
    objHttp.Open "GET", strurl, False
    objHttp.send Query
    strResult = objHttp.responsetext
    For i = 0 To UBound(arr) - 1
        index1 = InStr(strResult, Year(arr(i)) & "-" & Lpad(Month(arr(i)), "0", 2) & "-" & Lpad(Day(arr(i)), "0", 2))
        index2 = InStr(strResult, Year(arr(i + 1)) & "-" & Lpad(Month(arr(i + 1)), "0", 2) & "-" & Lpad(Day(arr(i + 1)), "0", 2))
        If (index1 > 0 And index2 > 0) Then
            index11 = index1
            index22 = index2
            For k = 1 To 4
                index11 = InStr(index11, strResult, ",")
                index22 = InStr(index22, strResult, ",")
            Next
            dprice1 = CDbl(Mid(strResult, index11 + 1, InStr(index11 + 1, strResult, ",") - index11 - 1))
            If (i = 0) Then
                currprice = dprice1
            End If
            dprice2 = CDbl(Mid(strResult, index22 + 1, InStr(index22 + 1, strResult, ",") - index22 - 1))
            gain(i) = Round(((currprice - dprice2) / dprice2) * 100, 2)
        End If
    Next
    strsql = "insert into AssetGain values('"
    strsql = strsql & rs.fields("ticker").Value & "','"
    strsql = strsql & orgDtToday & "',"
    strsql = strsql & empty2def(gain(0), 0) & ","
    strsql = strsql & empty2def(gain(1), 0) & ","
    strsql = strsql & empty2def(gain(2), 0) & ","
    strsql = strsql & empty2def(gain(3), 0) & ","
    strsql = strsql & empty2def(gain(4), 0) & ","
    strsql = strsql & empty2def(gain(5), 0) & ")"
    aconn.Execute strsql, , 1

End Sub

Function isMktOpen(dtDate)
    isMktOpen = True
    wday = Weekday(dtDate)
    If (wday = 1 Or wday = 7) Then
        isMktOpen = False
        Exit Function
    End If
    strsql = "select * from holiday where year=" & Year(dtDate) & " and holiday_date='" & mysqldate(dtDate, vbShortDate) & "'"
    Set rs = aconn.Execute(strsql, , 1)
    If Not rs.EOF Then
        isMktOpen = False
    End If
    rs.Close
End Function

Function tgDateAdd(dtype, numdays, dtDate)
    tdate = DateAdd(dtype, numdays, dtDate)
    If (isMktOpen(tdate)) Then
    Else
        For i = 1 To 3
            tdate = DateAdd("d", -1, tdate)
            If (isMktOpen(tdate)) Then
                Exit For
            End If
        Next
    End If
    tgDateAdd = tdate
End Function
