Attribute VB_Name = "modProc"
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public infoProcInfo As PROCESSENTRY32
Public tempClear As Integer
Public noClear As Integer

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Sub enumProc()
Dim procType As String
procType = ""
    servProc = 0
    uknProc = 0
    sysProc = 0
    If monitor <> 1 Or firewallStatus <> 1 Then
        If noClear = 0 Then
            frmProc.lstvwProc.ListItems.Clear
        End If
    Else
        If tempClear = 1 Then
            frmProc.lstvwProc.ListItems.Clear
            tempClear = 0
        End If
    End If
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    r = Process32Next(hSnapShot, uProcess)
    Do While r
        ProcessName = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        If UCase(ProcessName) = UCase("services.exe") Then
            servPID = uProcess.th32ProcessID
        ElseIf UCase(ProcessName) = UCase("explorer.exe") Then
            expPID = uProcess.th32ProcessID
        ElseIf UCase(ProcessName) = UCase("system") Then
            sysPID(1) = uProcess.th32ProcessID
        ElseIf UCase(ProcessName) = UCase("smss.exe") Then
            sysPID(2) = uProcess.th32ProcessID
        ElseIf UCase(ProcessName) = UCase("winlogon.exe") Then
            sysPID(3) = uProcess.th32ProcessID
        ElseIf UCase(ProcessName) = UCase("csrss.exe") Then
            sysPID(4) = uProcess.th32ProcessID
        ElseIf UCase(ProcessName) = UCase("lsass.exe") Then
            sysPID(5) = uProcess.th32ProcessID
        End If
        If popView = 1 Then
            If uProcess.th32ParentProcessID = servPID Then
                servProc = servProc + 1
                procType = "Service"
            ElseIf uProcess.th32ParentProcessID = expPID Then
                uknProc = uknProc + 1
                procType = "Unknown"
            ElseIf uProcess.th32ParentProcessID = sysPID(1) Or uProcess.th32ParentProcessID = sysPID(2) Or uProcess.th32ParentProcessID = sysPID(3) Or uProcess.th32ParentProcessID = sysPID(4) Or uProcess.th32ParentProcessID = sysPID(5) Then
                sysProc = sysProc + 1
                procType = "System"
            End If
            Call popLstvw(ProcessName, uProcess, getPriority(uProcess.th32ProcessID), getProcMem(uProcess.th32ProcessID), getCTime(uProcess.th32ProcessID), procType)
        End If
        
        If refreshPort = 1 Then
            If uProcess.th32ProcessID = tempPID Then
                  tempName = ProcessName
                  foundName = 1
            End If
        End If
        
        If checkParent = 1 Then
            If uProcess.th32ProcessID = infoProcInfo.th32ParentProcessID Then
                parentProc = ProcessName
            End If
        End If
        If infoCheck = 1 Then
            If checkPID <> 1 Then
                If UCase(ProcessName) = UCase(infoProc) Then
                    infoProcInfo = uProcess
                    infoProcMem = getProcMem(uProcess.th32ProcessID)
                    infoProcPri = getPriority(uProcess.th32ProcessID)
                    found = 1
                End If
            Else
                If uProcess.th32ProcessID = external Then
                    infoProcInfo = uProcess
                    infoProcMem = getProcMem(uProcess.th32ProcessID)
                    infoProcPri = getPriority(uProcess.th32ProcessID)
                    found = 1
                End If
            End If
        End If
        
        r = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
    If dontlbl <> 1 Then
        Call updateProcStats
    End If
    If popView = 0 Then
        popView = 1
    End If
End Sub

Public Sub popLstvw(procName, procArray As PROCESSENTRY32, priority As Long, procMem As PROCESS_MEMORY_COUNTERS, creationDate As String, procType As String)
Dim procArr(1 To 17) As Variant
Dim miss As Integer
Dim tmpPri As String
miss = 0

    If procType = "" Then
        procType = "Other"
    End If
    
    Select Case priority
    
    Case 32
        tmpPri = "Normal"
    Case 64
        tmpPri = "Idle"
    Case 128
        tmpPri = "High"
    Case 256
        tmpPri = "RealTime"
    End Select
    
    procArr(1) = procArray.th32ProcessID
    procArr(2) = procArray.cntThreads
    procArr(3) = procArray.th32ParentProcessID
    procArr(4) = procArray.pcPriClassBase
    procArr(5) = procArray.szExeFile
    procArr(6) = tmpPri
    procArr(7) = procMem.PageFaultCount
    procArr(8) = procMem.PeakWorkingSetSize
    procArr(9) = procMem.WorkingSetSize
    procArr(10) = procMem.QuotaPeakPagedPoolUsage
    procArr(11) = procMem.QuotaPagedPoolUsage
    procArr(12) = procMem.QuotaPeakNonPagedPoolUsage
    procArr(13) = procMem.QuotaNonPagedPoolUsage
    procArr(14) = procMem.PagefileUsage
    procArr(15) = procMem.PeakPagefileUsage
    procArr(16) = creationDate
    procArr(17) = procType
    
    Set lstItem = frmProc.lstvwProc.ListItems.Add(, , procName)
    
    For i = 2 To 18
        If cols(i) = 1 Then
            lstItem.SubItems((i - 1) - miss) = procArr(i - 1)
        Else
            miss = miss + 1
        End If
    Next i
    
End Sub

