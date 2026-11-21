' ============================================================
'  BO TRO GIUP - BOT TONG QUAN MAU (ban don gian de hoc)
'  Khoa hoc Thuc chien AI cho can bo thue - Ngo Hoang Thuy Huong - Zalo 0328 878 297
'  Chay bang CHAY.bat. KHONG bam dup file nay.
'  Toan bo file chi dung ky tu ASCII (khong dau) de wscript khong bao loi.
'
'  Cach dung lam KHUNG cho bot khac:
'    - Chi sua 2 cho co danh dau  >>> SUA O DAY <<<
'      1) Hang so MAN_HINH (ma man hinh TMS)
'      2) Sub CacBuocTrenTMS (cac buoc lay tu kich ban ghi)
'    - Phan con lai (doc MST, ghi Excel, dung bot, ghi log) giu nguyen.
' ============================================================
Option Explicit

Const TEN_TOOL = "BOT TONG QUAN MAU - BO TRO GIUP"

' >>> SUA O DAY (1) <<<  Ma man hinh TMS can tra
Const MAN_HINH = "FMCACOV"

' ID cac o tren man FMCACOV (lay tu kich ban ghi). Doi man hinh thi doi theo kich ban moi.
Const ID_MST_F4 = "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]"
Const ID_TEN = "wnd[0]/usr/subHEAD_DATA_SUBSCREEN:SAPLZFG_PIT01_10:0201/txtZST_PIT_FMCACOV-BPNAM"

Const CHO_TOI_DA_GIAY = 30

Dim fso, folder, session
Dim mstFile, stopPath, logPath, progPath, outFile
Dim dsMst(), nMst
Dim xl, wb, wsKq, wsTd
Dim dongKq, dictCot, soCot

Main

' ---------- Chuong trinh chinh ----------
Sub Main
        Dim i, mst, kq, ten
        Set fso = CreateObject("Scripting.FileSystemObject")
        folder = fso.GetParentFolderName(WScript.ScriptFullName)
        mstFile = folder & "\MST_CHAY.txt"
        stopPath = folder & "\STOP_TONG_QUAN.txt"
        logPath = folder & "\LOG_TONG_QUAN.txt"
        progPath = folder & "\PROGRESS_TONG_QUAN.txt"
        outFile = folder & "\KET_QUA_TONG_QUAN.xlsx"

        On Error Resume Next
        If fso.FileExists(logPath) Then fso.DeleteFile logPath, True
        On Error GoTo 0
        Ghi "=== BAT DAU ==="

        NapMst
        If nMst < 1 Then
                MsgBox "Chua co MST. Mo app bang CHAY.bat, dan MST roi bam LAY TONG QUAN.", 48, TEN_TOOL
                WScript.Quit 1
        End If
        Ghi "So MST: " & nMst

        If Not KetNoiSap() Then
                MsgBox "Chua mo SAP GUI, chua dang nhap TMS, hoac chua bat Scripting.", 16, TEN_TOOL
                WScript.Quit 1
        End If

        If Not MoExcel() Then WScript.Quit 1

        VaoManHinh

        For i = 0 To nMst - 1
                If fso.FileExists(stopPath) Then
                        Ghi "Gap file STOP - dung bot"
                        GhiTienDo i + 1, "DA DUNG"
                        Exit For
                End If
                If Not KetNoiSap() Then
                        GhiTd i, dsMst(i), "", "MAT KET NOI SAP - DA DUNG"
                        Exit For
                End If

                mst = dsMst(i)
                Ghi "MST " & (i + 1) & "/" & nMst & " = " & mst
                GhiTienDo i + 1, mst & "|dang chay"

                ten = ""
                kq = TraMotMst(mst, ten)
                GhiTd i, mst, ten, kq
                Ghi "  ket qua: " & kq
                LuuFile
        Next

        LuuFile
        GhiTienDo nMst, "XONG"
        Ghi "=== KET THUC ==="
        DongExcel
        MsgBox "Xong. Ket qua: " & outFile, 64, TEN_TOOL
End Sub

' ---------- Tra 1 MST: nhap, doc luoi, ghi Excel. Tra ve trang thai ----------
Function TraMotMst(mst, ByRef ten)
        Dim sbar, grid, n, loi
        TraMotMst = ""

        On Error Resume Next
        Err.Clear
        CacBuocTrenTMS mst
        loi = Err.Number
        If loi <> 0 Then
                TraMotMst = "LOI NHAP MST: " & Err.Description
                GhiLoiChiTiet Err.Number, Err.Description
                Err.Clear
                On Error GoTo 0
                DongPopup
                VaoManHinh
                Exit Function
        End If
        On Error GoTo 0

        Ghi "  b6: doc thanh trang thai"
        sbar = DocO("wnd[0]/sbar")
        If sbar <> "" And Not CoO(ID_TEN) Then
                TraMotMst = "KHONG CO DU LIEU: " & sbar
                DongPopup
                Exit Function
        End If

        ten = DocO(ID_TEN)
        Ghi "  b7: chon tab danh sach giao dich"
        BamTabDsGiaoDich
        Ghi "  b8: tim luoi ket qua"
        Set grid = TimLuoi()
        If grid Is Nothing Then
                TraMotMst = "KHONG THAY LUOI KET QUA"
                Exit Function
        End If

        Ghi "  b9: doc luoi va ghi Excel"
        n = NapDuSoDong(grid)
        GhiLuoi grid, n, mst, ten
        TraMotMst = "OK (" & n & " dong)"
End Function

' >>> SUA O DAY (2) <<<
' Cac buoc thao tac cho 1 MST, chep tu kich ban ghi.
' Dung bien mst thay cho MST mau. Sau moi lan bam phim goi ChoSap.
' TRUOC moi thao tac goi Ghi "  bN: <viec dang lam>" de khi loi, dong log cuoi
' cho biet bot chet o buoc nao.
' Khong dung On Error trong Sub nay: loi se duoc TraMotMst bat va ghi lai.
Sub CacBuocTrenTMS(mst)
        Ghi "  b1: bam F4 mo o tim MST"
        session.findById("wnd[0]").sendVKey 4
        ChoSap
        Ghi "  b2: chon tab tim kiem trong popup"
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001").Select
        ChoSap
        Ghi "  b3: nhap MST " & mst & " va Enter"
        session.findById(ID_MST_F4).Text = mst
        session.findById("wnd[1]").sendVKey 0
        ChoSap
        Ghi "  b4: Enter xac nhan doi tuong"
        session.findById("wnd[0]").sendVKey 0
        ChoSap
        Ghi "  b5: Enter hien thi tong quan"
        session.findById("wnd[0]").sendVKey 0
        ChoSap
End Sub

' ---------- Vao man hinh can tra ----------
Sub VaoManHinh
        On Error Resume Next
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & MAN_HINH
        session.findById("wnd[0]").sendVKey 0
        ChoSap
        On Error GoTo 0
End Sub

' ---------- Ket noi SAP GUI dang mo ----------
Function KetNoiSap()
        Dim sapGui, app, conn
        KetNoiSap = False
        Set session = Nothing
        On Error Resume Next
        Set sapGui = GetObject("SAPGUI")
        Set app = sapGui.GetScriptingEngine
        Set conn = app.Children(0)
        Set session = conn.Children(0)
        On Error GoTo 0
        If session Is Nothing Then Exit Function
        KetNoiSap = True
End Function

' ---------- Cho SAP xu ly xong (toi da CHO_TOI_DA_GIAY) ----------
Sub ChoSap
        Dim t0
        t0 = Timer
        On Error Resume Next
        WScript.Sleep 100
        Do While session.Busy
                WScript.Sleep 100
                If Abs(Timer - t0) > CHO_TOI_DA_GIAY Then Exit Do
        Loop
        On Error GoTo 0
End Sub

' ---------- Kiem tra 1 o co tren man hinh khong ----------
Function CoO(id)
        On Error Resume Next
        Err.Clear
        session.findById id
        CoO = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
End Function

' ---------- Doc chu trong 1 o, khong co thi tra ve rong ----------
Function DocO(id)
        DocO = ""
        On Error Resume Next
        DocO = session.findById(id).Text & ""
        On Error GoTo 0
End Function

' ---------- Dong cac popup dang mo ----------
Sub DongPopup
        On Error Resume Next
        If CoO("wnd[3]") Then session.findById("wnd[3]").sendVKey 12
        If CoO("wnd[2]") Then session.findById("wnd[2]").sendVKey 12
        If CoO("wnd[1]") Then session.findById("wnd[1]").sendVKey 12
        ChoSap
        On Error GoTo 0
End Sub

' ---------- Chon tab danh sach giao dich (neu co) ----------
Sub BamTabDsGiaoDich
        Dim tabs, i, kid, t, chon
        If Not CoO("wnd[0]/usr/tabsDATA_DISP") Then Exit Sub
        On Error Resume Next
        Set tabs = session.findById("wnd[0]/usr/tabsDATA_DISP")
        chon = ""
        For i = 0 To tabs.Children.Count - 1
                Set kid = Nothing
                Set kid = tabs.Children(i)
                If Not kid Is Nothing Then
                        t = BoDau(kid.Text & " " & kid.Name)
                        If InStr(t, "ds giao") > 0 Then chon = kid.Name
                        If chon = "" And InStr(t, "giao dich") > 0 And InStr(t, "bu tru") = 0 And InStr(t, "tong quan") = 0 Then chon = kid.Name
                End If
        Next
        If chon <> "" Then
                tabs.selectTab chon
                ChoSap
        End If
        On Error GoTo 0
End Sub

' ---------- Tim luoi ket qua (GuiGridView) lon nhat tren man hinh ----------
Function TimLuoi()
        Dim best, bestRc
        Set best = Nothing
        bestRc = -1
        On Error Resume Next
        TimLuoiDeQuy session.findById("wnd[0]"), best, bestRc
        On Error GoTo 0
        Set TimLuoi = best
End Function

Sub TimLuoiDeQuy(goc, ByRef best, ByRef bestRc)
        Dim i, kid, cols, rc
        If goc Is Nothing Then Exit Sub
        On Error Resume Next
        Set cols = Nothing
        Set cols = goc.ColumnOrder
        rc = -1
        rc = goc.RowCount
        If Not cols Is Nothing And rc > bestRc Then
                If InStr(goc.Id & "", "IMAGE_CONTAINER") = 0 Then
                        Set best = goc
                        bestRc = rc
                End If
        End If
        For i = 0 To goc.Children.Count - 1
                Set kid = Nothing
                Set kid = goc.Children(i)
                If Not kid Is Nothing Then TimLuoiDeQuy kid, best, bestRc
        Next
        On Error GoTo 0
End Sub

' ---------- Cuon luoi de SAP nap het dong ----------
Function NapDuSoDong(grid)
        Dim rc, rc2, i
        NapDuSoDong = 0
        On Error Resume Next
        rc = CLng(0 & grid.RowCount)
        For i = 1 To 40
                If rc < 1 Then Exit For
                grid.firstVisibleRow = rc - 1
                WScript.Sleep 80
                rc2 = CLng(0 & grid.RowCount)
                If rc2 <= rc Then Exit For
                rc = rc2
        Next
        grid.firstVisibleRow = 0
        On Error GoTo 0
        NapDuSoDong = rc
End Function

' ---------- Ghi toan bo luoi vao sheet Ket qua ----------
Sub GhiLuoi(grid, n, mst, ten)
        Dim r, c, nm, cols, vals(), j, tieuDe
        On Error Resume Next
        Set cols = grid.ColumnOrder
        For c = 0 To cols.Count - 1
                nm = cols.Item(c) & ""
                If Not dictCot.Exists(nm) Then
                        tieuDe = ""
                        tieuDe = grid.GetDisplayedColumnTitle(nm) & ""
                        If tieuDe = "" Then tieuDe = nm
                        soCot = soCot + 1
                        dictCot.Add nm, soCot
                        wsKq.Cells(1, soCot).Value = tieuDe
                End If
        Next
        For r = 0 To n - 1
                ReDim vals(soCot - 1)
                vals(0) = mst
                vals(1) = ten
                For c = 0 To cols.Count - 1
                        nm = cols.Item(c) & ""
                        j = dictCot(nm) - 1
                        vals(j) = grid.GetCellValue(r, nm) & ""
                Next
                dongKq = dongKq + 1
                wsKq.Range(wsKq.Cells(dongKq, 1), wsKq.Cells(dongKq, soCot)).Value = vals
        Next
        On Error GoTo 0
End Sub

' ---------- Mo Excel, tao file ket qua ----------
Function MoExcel()
        MoExcel = False
        On Error Resume Next
        If fso.FileExists(outFile) Then fso.DeleteFile outFile, True
        If fso.FileExists(outFile) Then
                MsgBox "Dang mo file KET_QUA_TONG_QUAN.xlsx. Dong file do roi chay lai.", 16, TEN_TOOL
                Exit Function
        End If
        Err.Clear
        Set xl = CreateObject("Excel.Application")
        If Err.Number <> 0 Then
                MsgBox "Khong mo duoc Excel tren may nay.", 16, TEN_TOOL
                Exit Function
        End If
        xl.Visible = False
        xl.DisplayAlerts = False
        Set wb = xl.Workbooks.Add
        Do While wb.Worksheets.Count < 2
                wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count)
        Loop
        Set wsKq = wb.Worksheets(1)
        Set wsTd = wb.Worksheets(2)
        wsKq.Name = "Ket qua"
        wsTd.Name = "Tien do"
        wsKq.Cells.NumberFormat = "@"
        wsTd.Cells.NumberFormat = "@"
        wsKq.Cells(1, 1).Value = "MST"
        wsKq.Cells(1, 2).Value = "Ten NNT"
        wsTd.Cells(1, 1).Value = "STT"
        wsTd.Cells(1, 2).Value = "MST"
        wsTd.Cells(1, 3).Value = "Ten NNT"
        wsTd.Cells(1, 4).Value = "Trang thai"
        wsKq.Rows(1).Font.Bold = True
        wsTd.Rows(1).Font.Bold = True
        Set dictCot = CreateObject("Scripting.Dictionary")
        soCot = 2
        dongKq = 1
        Err.Clear
        wb.SaveAs outFile, 51
        If Err.Number <> 0 Then
                MsgBox "Khong luu duoc file ket qua vao thu muc: " & folder, 16, TEN_TOOL
                xl.Quit
                Exit Function
        End If
        On Error GoTo 0
        MoExcel = True
End Function

' ---------- Ghi 1 dong sheet Tien do ----------
Sub GhiTd(i, mst, ten, trangThai)
        On Error Resume Next
        wsTd.Cells(i + 2, 1).Value = CStr(i + 1)
        wsTd.Cells(i + 2, 2).Value = mst
        wsTd.Cells(i + 2, 3).Value = ten
        wsTd.Cells(i + 2, 4).Value = trangThai
        On Error GoTo 0
End Sub

Sub LuuFile
        On Error Resume Next
        wb.Save
        On Error GoTo 0
End Sub

Sub DongExcel
        On Error Resume Next
        wsKq.Columns.AutoFit
        wsTd.Columns.AutoFit
        wb.Save
        wb.Close False
        xl.Quit
        On Error GoTo 0
End Sub

' ---------- Doc danh sach MST, bo trung ----------
Sub NapMst
        Dim ts, d, dict
        nMst = 0
        ReDim dsMst(0)
        If Not fso.FileExists(mstFile) Then Exit Sub
        Set dict = CreateObject("Scripting.Dictionary")
        Set ts = fso.OpenTextFile(mstFile, 1, False, -2)
        Do Until ts.AtEndOfStream
                d = ChuanMst(ts.ReadLine)
                If d <> "" Then
                        If Not dict.Exists(d) Then
                                dict.Add d, 1
                                ReDim Preserve dsMst(nMst)
                                dsMst(nMst) = d
                                nMst = nMst + 1
                        End If
                End If
        Loop
        ts.Close
End Sub

' MST hop le: 10 so, 10 so-3 so, 12 so (CCCD) hoac 13 so
Function ChuanMst(s)
        Dim t, re
        ChuanMst = ""
        t = Replace(Replace(Replace(Trim(s & ""), " ", ""), vbTab, ""), "'", "")
        Set re = New RegExp
        re.Pattern = "^(\d{10}|\d{10}-\d{3}|\d{12}|\d{13})$"
        If re.Test(t) Then ChuanMst = t
End Function

' ---------- Bo dau tieng Viet (de so sanh ten tab) ----------
Function BoDau(s)
        Dim i, c, r, n
        r = ""
        s = LCase(s & "")
        For i = 1 To Len(s)
                c = Mid(s, i, 1)
                n = AscW(c)
                Select Case n
                        Case 224, 225, 226, 227, 259, 7841, 7843, 7845, 7847, 7849, 7851, 7853, 7855, 7857, 7859, 7861, 7863
                                r = r & "a"
                        Case 232, 233, 234, 7865, 7867, 7869, 7871, 7873, 7875, 7877, 7879
                                r = r & "e"
                        Case 236, 237, 297, 7881, 7883
                                r = r & "i"
                        Case 242, 243, 244, 245, 417, 7885, 7887, 7889, 7891, 7893, 7895, 7897, 7899, 7901, 7903, 7905, 7907
                                r = r & "o"
                        Case 249, 250, 361, 432, 7909, 7911, 7913, 7915, 7917, 7919, 7921
                                r = r & "u"
                        Case 253, 7923, 7925, 7927, 7929
                                r = r & "y"
                        Case 272, 273
                                r = r & "d"
                        Case Else
                                r = r & c
                End Select
        Next
        BoDau = r
End Function

' ---------- Ghi chi tiet loi vao log (de dan cho Gemini sua) ----------
Sub GhiLoiChiTiet(soLoi, moTa)
        Dim tc, cuaSo, w
        On Error Resume Next
        tc = "" : cuaSo = ""
        tc = session.Info.Transaction & ""
        For w = 3 To 0 Step -1
                If CoO("wnd[" & w & "]") Then
                        cuaSo = "wnd[" & w & "] = " & session.findById("wnd[" & w & "]").Text
                        Exit For
                End If
        Next
        Ghi "  !!! LOI so " & soLoi & ": " & moTa
        Ghi "  !!! Man hinh: " & tc & " | Cua so tren cung: " & cuaSo
        Ghi "  !!! Thanh trang thai: " & DocO("wnd[0]/sbar")
        On Error GoTo 0
End Sub

' ---------- Nhat ky va tien do ----------
Sub Ghi(s)
        Dim ts
        On Error Resume Next
        Set ts = fso.OpenTextFile(logPath, 8, True)
        ts.WriteLine Now & "  " & s
        ts.Close
        On Error GoTo 0
End Sub

Sub GhiTienDo(dong, noiDung)
        Dim ts
        On Error Resume Next
        Set ts = fso.OpenTextFile(progPath, 2, True)
        ts.Write dong & "/" & nMst & "|" & noiDung
        ts.Close
        On Error GoTo 0
End Sub
