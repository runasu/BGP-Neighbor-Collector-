'#$Language="VBScript"
'#$Interface="1.0"
Option Explicit

' =====================================================
'  SecureCRT VBScript: RR-INET BGP Collector + Excel
'  Versi: 2026-01-30 (untuk parser v2.7)
'
'  Fitur:
'   - Logging gabungan + file per-host (dengan marker ">>>")
'   - Auto-detect Python (prioritas: py -3)
'   - Jalankan parser Python v2.7 (Excel cantik:
'     header warna, freeze, autofilter, auto-fit, angka ribuan,
'     conditional State + data bars, Delta dengan 3-color scale + icon set,
'     table style, tab color, Summary + Dashboard charts)
'   - Preflight package: cek/instal pandas/xlsxwriter/openpyxl (opsional)
'   - Opsi buka Excel setelah selesai
' =====================================================

' =========================
' KONFIGURASI UTAMA
' =========================
Const ForReading = 1

' === Kredensial (kosongkan untuk diprompt) ===
Const CFG_UserSSH = "nec"   ' boleh "" agar diprompt
Const CFG_PassSSH = ""            ' boleh "" agar diprompt (masked)
Const CFG_FallbackPass = ""       ' opsional: dipakai jika CFG_PassSSH kosong

' === Folder & File ===
Dim TargetFolder : TargetFolder = "C:\Users\0684115\Desktop\RR-INET\Lab"     ' <== Sesuaikan
Dim HostsFile    : HostsFile    = TargetFolder & "\lab.txt"          ' Daftar host
Dim LogFile      : LogFile      = TargetFolder & "\bgp_peers_all.txt"    ' Log gabungan (default)
Dim OutXlsx      : OutXlsx      = TargetFolder & "\sample.xlsx" ' Output Excel

' === Parser Python ===
Dim PythonExe    : PythonExe    = "C:\Python39\python.exe"                ' <== Boleh salah; akan AUTODETECT
Dim ParserPath   : ParserPath   = TargetFolder & "\parse_bgp_to_excel_v2_8.py"  ' <== pakai v2.8 agar styling maksimal

' === Opsi Excel / Eksekusi ===
Dim ExcelMode           : ExcelMode           = "per-host"   ' "per-host" atau "single"
Dim PreferEngine        : PreferEngine        = "xlsxwriter" ' paksa xlsxwriter untuk visual terbaik
Dim OpenExcelWhenDone   : OpenExcelWhenDone   = True         ' True: buka Excel setelah selesai
Dim AutoInstallPackages : AutoInstallPackages = True         ' True: cek/instal paket Python

' === Prompt Junos ===
Const PROMPT_OP  = "> "
Const PROMPT_CFG = "#"

' === Variabel global ===
Dim fso, file, g_user, g_pass, LogsRoot, TodayFolder

Sub Main
    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ===== PERSIAPAN FOLDER =====
    If Not fso.FolderExists(TargetFolder) Then
        fso.CreateFolder TargetFolder
        If Err.Number <> 0 Then
            crt.Dialog.MessageBox "Gagal membuat folder: " & TargetFolder & vbCrLf & _
                                   "Error: " & Err.Description
            Exit Sub
        End If
        Err.Clear
    End If

    ' Folder logs\YYYYMMDD
    LogsRoot    = TargetFolder & "\logs"
    If Not fso.FolderExists(LogsRoot) Then fso.CreateFolder LogsRoot
    TodayFolder = LogsRoot & "\" & DateToYYYYMMDD(Date)
    If Not fso.FolderExists(TodayFolder) Then fso.CreateFolder TodayFolder

    ' ===== BUKA HOST LIST =====
    If Not fso.FileExists(HostsFile) Then
        crt.Dialog.MessageBox "File host list tidak ditemukan: " & HostsFile
        Exit Sub
    End If

    Set file = fso.OpenTextFile(HostsFile, ForReading, False)
    If Err.Number <> 0 Then
        crt.Dialog.MessageBox "Gagal membuka file host list: " & HostsFile & vbCrLf & Err.Description
        Exit Sub
    End If

    ' ===== LOGGING: SATU FILE (gabungan) =====
    If SupportsLogUsingSessionOptions() Then
        crt.Session.LogUsingSessionOptions = False  ' aman bila properti tersedia
    End If
    crt.Session.LogFileName = LogFile
    If Not crt.Session.Logging Then
        crt.Session.Log True, True   ' start logging (append)
    End If

    ' ===== STABILITAS INTERAKSI =====
    crt.Screen.Synchronous = True
    crt.Screen.IgnoreCase  = True

    ' ===== AMBIL KREDENSIAL =====
    g_user = CFG_UserSSH
    If Trim(g_user) = "" Then
        g_user = crt.Dialog.Prompt("Masukkan username SSH:", "Login SSH", "", False)
        If g_user = "" Then
            DoCleanup
            Exit Sub
        End If
    End If

    g_pass = CFG_PassSSH
    If Trim(g_pass) = "" Then
        If Trim(CFG_FallbackPass) <> "" Then
            g_pass = CFG_FallbackPass
        Else
            g_pass = crt.Dialog.Prompt("Masukkan password SSH:", "Login SSH", "", True)
            If g_pass = "" Then
                DoCleanup
                Exit Sub
            End If
        End If
    End If

    ' ===== LOOP: PROSES SETIAP HOST =====
    Do Until file.AtEndOfStream
        Dim line, hostName, hostIp, nameOrIp
        line = Trim(file.ReadLine)
        If Len(line) > 0 Then
            hostName = ""
            hostIp   = ""

            If InStr(line, ";") > 0 Then
                hostName = Trim(Split(line, ";")(0))
                hostIp   = Trim(Split(line, ";")(1))
            Else
                hostName = line
                hostIp   = line  ' bisa hostname atau IP
            End If

            nameOrIp = hostIp

            ' === Pemisah log gabungan per host ===
            LogSeparator "HOST: " & hostName & "  (" & nameOrIp & ")"

            ' === File log per-host ===
            Dim perHostFile
            perHostFile = TodayFolder & "\" & SanitizeFileName(hostName) & "_" & TimeToHHMMSS(Now) & ".txt"

            ' === SSH ===
            Dim sshCmd
            If Trim(g_user) <> "" Then
                sshCmd = "ssh -l " & g_user & " " & nameOrIp
            Else
                sshCmd = "ssh " & nameOrIp
            End If
            crt.Screen.Send sshCmd & vbCr

            ' === Handshake / login ===
            If HandleSshLogin() Then
                ' === Perintah Junos ===
                JalankanPerintahJunos_SaveLog perHostFile

                ' === Keluar device ===
                crt.Screen.Send "exit" & vbCr
                WaitAnyPrompt 10
            Else
                AbortToLocalPrompt
                ' Catat gagal login di file per-host
                WriteTextFile perHostFile, _
                    "LOGIN GAGAL ke " & hostName & " (" & nameOrIp & ") pada " & Now & vbCrLf, False
            End If

            crt.Sleep 300 ' jeda antar-host
        End If
    Loop

    ' ===== TUTUP LOG =====
    DoCleanup

    ' ===== JALANKAN PARSER (auto-detect Python + optional auto-install packages) =====
    RunParserIfAvailable
End Sub

' ====================================================
' Cleanup routine: tutup file dan stop logging
' ====================================================
Sub DoCleanup()
    On Error Resume Next
    If Not file Is Nothing Then
        file.Close
        Set file = Nothing
    End If
    If crt.Session.Logging Then
        crt.Session.Log False
    End If
End Sub

' ====================================================
' Jalankan parser Python jika file tersedia
' ====================================================
Sub RunParserIfAvailable()
    On Error Resume Next

    If Not fso.FileExists(ParserPath) Then
        crt.Dialog.MessageBox "Parser Python tidak ditemukan: " & ParserPath & vbCrLf & _
                              "Letakkan 'parse_bgp_to_excel_v2_7.py' di TargetFolder lalu jalankan ulang."
        Exit Sub
    End If

    Dim logCandidate : logCandidate = LogFile
    If Not fso.FileExists(logCandidate) Then
        Dim altLog : altLog = TargetFolder & "\bgp_peers_all"
        If fso.FileExists(altLog) Then logCandidate = altLog
    End If
    If Not fso.FileExists(logCandidate) Then
        crt.Dialog.MessageBox "File log gabungan tidak ditemukan:" & vbCrLf & _
                              "  " & LogFile & vbCrLf & "atau" & vbCrLf & _
                              "  " & TargetFolder & "\bgp_peers_all"
        Exit Sub
    End If

    Dim pyPath : pyPath = DetectPythonExe()
    If pyPath = "" Then
        If fso.FileExists(PythonExe) Then
            pyPath = PythonExe
        Else
            crt.Dialog.MessageBox "Python tidak ditemukan. Opsi:" & vbCrLf & _
                "1) Install Python 3.x dan centang 'Add to PATH'." & vbCrLf & _
                "2) Pastikan 'py -3' atau 'python' bisa dipanggil dari Command Prompt." & vbCrLf & _
                "3) Edit variabel PythonExe di script ini."
            Exit Sub
        End If
    End If

    ' Pre-create parser_run.log
    Dim parserLog : parserLog = TargetFolder & "\parser_run.log"
    Dim tf
    Set tf = fso.OpenTextFile(parserLog, 2, True)
    tf.WriteLine String(70, "=")
    tf.WriteLine "WAKTU   : " & Now
    tf.WriteLine "PYTHON  : " & pyPath
    tf.WriteLine "PARSER  : " & ParserPath
    tf.WriteLine "LOG     : " & logCandidate
    tf.WriteLine "OUT XLS : " & OutXlsx
    tf.WriteLine "MODE    : " & ExcelMode
    tf.WriteLine "ENGINE  : " & PreferEngine
    tf.WriteLine String(70, "=")
    tf.Close
    Set tf = Nothing

    If AutoInstallPackages Then
        EnsurePythonPackages pyPath, parserLog
    End If

    ' Build command (pakai --engine xlsxwriter agar styling maksimal)
    Dim cmd, sh, rc
    If LCase(Left(pyPath, 4)) = "py -" Then
        cmd = "cmd /d /c " & pyPath & _
              " """ & ParserPath & """ --log """ & logCandidate & _
              """ --hosts """ & HostsFile & """ --out """ & OutXlsx & _
              """ --mode " & ExcelMode & " --engine " & PreferEngine & _
              " >> """ & parserLog & """ 2>&1"
    Else
        cmd = "cmd /d /c """ & pyPath & """ """ & ParserPath & _
              """ --log """ & logCandidate & """ --hosts """ & HostsFile & _
              """ --out """ & OutXlsx & """ --mode " & ExcelMode & _
              " --engine " & PreferEngine & " >> """ & parserLog & """ 2>&1"
    End If

    AppendLog parserLog, "CMDLINE : " & cmd
    AppendLog parserLog, String(70, "-")

    Set sh = CreateObject("WScript.Shell")
    rc = sh.Run(cmd, 1, True)

    If rc = 0 Then
        crt.Dialog.MessageBox "Selesai. Excel dibuat: " & OutXlsx
        If OpenExcelWhenDone And fso.FileExists(OutXlsx) Then
            sh.Run "cmd /c start """" """ & OutXlsx & """", 0, False
        End If
    Else
        Dim tailText : tailText = TailFile(parserLog, 20)
        crt.Dialog.MessageBox "Parser gagal (exit code=" & rc & "). Lihat " & parserLog & vbCrLf & vbCrLf & tailText
    End If

    ' Rapikan sesi lokal
    crt.Screen.Synchronous = True
    crt.Screen.Send "exit" & vbCr
    crt.Session.Disconnect
End Sub

' Append log helper
Sub AppendLog(ByVal path, ByVal content)
    On Error Resume Next
    Dim tf
    Set tf = fso.OpenTextFile(path, 8, True)
    tf.WriteLine content
    tf.Close
End Sub

' Pastikan paket tersedia
Sub EnsurePythonPackages(ByVal pyPath, ByVal parserLog)
    On Error Resume Next
    Dim missing : missing = ""
    If Not PipShow(pyPath, "pandas") Then missing = missing & " pandas"
    If Not PipShow(pyPath, "xlsxwriter") Then missing = missing & " xlsxwriter"
    If Not PipShow(pyPath, "openpyxl") Then missing = missing & " openpyxl"

    If Trim(missing) = "" Then
        AppendLog parserLog, "PKG OK  : pandas/xlsxwriter/openpyxl sudah terpasang."
        Exit Sub
    End If

    AppendLog parserLog, "PKG MISS:" & missing
    AppendLog parserLog, "PKG INST: mencoba install --user" & missing

    Dim cmd, sh, rc
    Set sh = CreateObject("WScript.Shell")
    If LCase(Left(pyPath, 4)) = "py -" Then
        cmd = "cmd /d /c " & pyPath & " -m pip install --user" & missing & " >> """ & parserLog & """ 2>&1"
    Else
        cmd = "cmd /d /c """ & pyPath & """ -m pip install --user" & missing & " >> """ & parserLog & """ 2>&1"
    End If
    rc = sh.Run(cmd, 1, True)
    If rc = 0 Then
        AppendLog parserLog, "PKG INST: sukses."
    Else
        AppendLog parserLog, "PKG INST: gagal (exit code=" & rc & "). Lanjut eksekusi; parser akan coba jalan dengan yang ada."
    End If
End Sub

Function PipShow(ByVal pyPath, ByVal pkg)
    On Error Resume Next
    Dim cmd, sh, rc
    Set sh = CreateObject("WScript.Shell")
    If LCase(Left(pyPath, 4)) = "py -" Then
        cmd = "cmd /c " & pyPath & " -m pip show " & pkg & " >nul 2>&1"
    Else
        cmd = "cmd /c """ & pyPath & """ -m pip show " & pkg & " >nul 2>&1"
    End If
    rc = sh.Run(cmd, 0, True)
    PipShow = (rc = 0)
End Function

' Tail n baris terakhir
Function TailFile(ByVal path, ByVal nLines)
    On Error Resume Next
    TailFile = ""
    If Not fso.FileExists(path) Then Exit Function
    Dim txt, arr, startIdx, i
    txt = fso.OpenTextFile(path, 1, False).ReadAll
    arr = Split(txt, vbCrLf)
    startIdx = UBound(arr) - nLines + 1
    If startIdx < 0 Then startIdx = 0
    For i = startIdx To UBound(arr)
        TailFile = TailFile & arr(i) & vbCrLf
    Next
End Function

' Deteksi python.exe
Function DetectPythonExe()
    On Error Resume Next
    Dim sh, execObj, line, candidate
    Set sh = CreateObject("WScript.Shell")
    DetectPythonExe = ""

    Set execObj = sh.Exec("cmd /c py -3 -c ""import sys;print(sys.executable)""")
    If Not execObj Is Nothing Then
        line = Trim(execObj.StdOut.ReadAll)
        If line <> "" And fso.FileExists(line) Then
            If IsRealPython(line) Then DetectPythonExe = line : Exit Function
        End If
    End If

    Set execObj = sh.Exec("cmd /c where python")
    If Not execObj Is Nothing Then
        line = execObj.StdOut.ReadAll
        If Trim(line) <> "" Then
            Dim arr, i
            arr = Split(line, vbCrLf)
            For i = LBound(arr) To UBound(arr)
                candidate = Trim(arr(i))
                If candidate <> "" And fso.FileExists(candidate) Then
                    If IsRealPython(candidate) Then DetectPythonExe = candidate : Exit Function
                End If
            Next
        End If
    End If

    Dim bases, ver, p
    bases = Array( _
        "C:\Users\" & GetEnv("USERNAME") & "\AppData\Local\Programs\Python", _
        "C:\Program Files\Python", _
        "C:\Program Files (x86)\Python" _
    )
    For Each p In bases
        For ver = 8 To 13
            candidate = p & "\Python3" & ver & "\python.exe"
            If fso.FileExists(candidate) Then
                If IsRealPython(candidate) Then DetectPythonExe = candidate : Exit Function
            End If
        Next
    Next

    Set execObj = sh.Exec("cmd /c py -3 -V")
    If Not execObj Is Nothing Then
        line = Trim(execObj.StdOut.ReadAll & execObj.StdErr.ReadAll)
        If InStr(1, line, "Python", vbTextCompare) > 0 Then
            DetectPythonExe = "py -3"
        End If
    End If
End Function

Function IsRealPython(ByVal exePath)
    On Error Resume Next
    IsRealPython = False
    If exePath = "" Then Exit Function
    If InStr(1, LCase(exePath), "\windowsapps\python.exe", vbTextCompare) > 0 Then Exit Function

    Dim testCmd, sh, rc
    Set sh = CreateObject("WScript.Shell")
    If LCase(Left(exePath, 4)) = "py -" Then
        testCmd = "cmd /c " & exePath & " -c ""import sys;print('ok')"""
    Else
        testCmd = "cmd /c """ & exePath & """ -c ""import sys;print('ok')"""
    End If

    rc = sh.Run(testCmd, 0, True)
    If rc = 0 Then IsRealPython = True
End Function

Function GetEnv(ByVal name)
    On Error Resume Next
    GetEnv = CreateObject("WScript.Shell").Environment("PROCESS")(name)
    If GetEnv = "" Then GetEnv = CreateObject("WScript.Shell").Environment("USER")(name)
    If GetEnv = "" Then GetEnv = CreateObject("WScript.Shell").Environment("SYSTEM")(name)
End Function

Function SupportsLogUsingSessionOptions()
    On Error Resume Next
    Dim tmp
    tmp = crt.Session.LogUsingSessionOptions
    If Err.Number <> 0 Then
        Err.Clear
        SupportsLogUsingSessionOptions = False
    Else
        SupportsLogUsingSessionOptions = True
    End If
    On Error GoTo 0
End Function

Sub LogSeparator(ByVal text)
    On Error Resume Next
    crt.Session.LogWriteLine String(70, "-")
    crt.Session.LogWriteLine text & "  |  " & Now
    crt.Session.LogWriteLine String(70, "-")
End Sub

Sub LogMarker(ByVal cmdText)
    On Error Resume Next
    crt.Session.LogWriteLine ">>> " & cmdText
End Sub

Function HandleSshLogin()
    HandleSshLogin = False
    Dim cek_ssh
    cek_ssh = Array("password: ", "Password:", "(yes/no/[fingerprint])?", "(yes/no)", _
                    "continue connecting", "Permission denied", PROMPT_OP, PROMPT_CFG)

    Dim idx, tries
    tries = 0
    Do
        idx = crt.Screen.WaitForStrings(cek_ssh, 30)
        Select Case idx
            Case 0: Exit Function
            Case 1, 2: crt.Screen.Send g_pass & vbCr
            Case 3, 4, 5: crt.Screen.Send "yes" & vbCr
            Case 6: Exit Function
            Case 7, 8: HandleSshLogin = True : Exit Function
        End Select
        tries = tries + 1
        If tries > 6 Then Exit Do
    Loop
End Function

Function WaitAnyPrompt(ByVal timeoutSec)
    Dim idx
    idx = crt.Screen.WaitForStrings(Array(PROMPT_OP, PROMPT_CFG), timeoutSec)
    WaitAnyPrompt = (idx > 0)
End Function

Sub AbortToLocalPrompt()
    On Error Resume Next
    crt.Screen.Send Chr(3)
    WaitAnyPrompt 3
    crt.Sleep 300
    crt.Screen.Send Chr(3)
    WaitAnyPrompt 3
End Sub

Sub JalankanPerintahJunos_SaveLog(ByVal perHostFile)
    Dim cap

    ' Bangunkan prompt
    crt.Screen.Send vbCr
    WaitAnyPrompt 5
    crt.Sleep 200

    ' 1) Matikan pager
    LogMarker "set cli screen-length 0"
    crt.Screen.Send "set cli screen-length 0" & vbCr
    cap = crt.Screen.ReadString(Array(PROMPT_OP, PROMPT_CFG), 10)
    AppendTextFile perHostFile, ">>> set cli screen-length 0" & vbCrLf & cap & vbCrLf

    ' 2) Timestamp
    LogMarker "set cli timestamp"
    crt.Screen.Send "set cli timestamp" & vbCr
    cap = crt.Screen.ReadString(Array(PROMPT_OP, PROMPT_CFG), 10)
    AppendTextFile perHostFile, ">>> set cli timestamp" & vbCrLf & cap & vbCrLf

    ' 3) Perintah utama
    LogMarker "show bgp neighbor | no-more"
    crt.Screen.Send "show bgp neighbor | no-more" & vbCr
    cap = crt.Screen.ReadString(Array(PROMPT_OP, PROMPT_CFG), 180)
    AppendTextFile perHostFile, ">>> show bgp neighbor | no-more" & vbCrLf & cap & vbCrLf
End Sub

Sub WriteTextFile(ByVal path, ByVal content, ByVal appendMode)
    On Error Resume Next
    Dim tf
    If Not appendMode Then
        If fso.FileExists(path) Then fso.DeleteFile path, True
    End If
    Set tf = fso.OpenTextFile(path, 8, True)
    tf.Write content
    tf.Close
End Sub

Sub AppendTextFile(ByVal path, ByVal content)
    On Error Resume Next
    Dim tf
    Set tf = fso.OpenTextFile(path, 8, True)
    tf.Write content
    tf.Close
End Sub

Function DateToYYYYMMDD(d)
    Dim y, m, dd
    y  = Year(d)
    m  = Right("0" & Month(d), 2)
    dd = Right("0" & Day(d),   2)
    DateToYYYYMMDD = y & m & dd
End Function

Function TimeToHHMMSS(t)
    Dim hh, mm, ss
    hh = Right("0" & Hour(t),   2)
    mm = Right("0" & Minute(t), 2)
    ss = Right("0" & Second(t), 2)
    TimeToHHMMSS = hh & mm & ss
End Function

Function SanitizeFileName(ByVal s)
    Dim badChars, i
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = 0 To UBound(badChars)
        s = Replace(s, badChars(i), "_")
    Next
    SanitizeFileName = s
End Function