Attribute VB_Name = "modOther"
Public colHead As ColumnHeader
Public lstItem As ListItem
Public cSysTime As SYSTEMTIME
Public eSysTime As SYSTEMTIME
Public kSysTime As SYSTEMTIME
Public uSysTime As SYSTEMTIME
Public infoCheck  As Integer
Public infoProc As String
Public infoProcMem As PROCESS_MEMORY_COUNTERS
Public infoProcPri As Long
Public popView As Integer
Public servProc As Integer
Public uknProc As Integer
Public expPID As Integer
Public servPID As Integer
Public found As Integer
Public sysProc As Integer
Public external As Integer
Public checkPID As Integer
Public checkParent As Integer
Public parentProc As String
Public sysPID(1 To 5) As Integer
Public cols(1 To 18) As Integer
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const PROCESS_SET_INFORMATION As Long = (&H200)
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long

Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


Public Function getPriority(pid As Long)
    hWnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    pri = GetPriorityClass(hWnd)
    CloseHandle hWnd
    getPriority = pri
End Function

Public Function getProcMem(pid As Long) As PROCESS_MEMORY_COUNTERS
Dim procMem As PROCESS_MEMORY_COUNTERS
    
    hWnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    procMem.cb = LenB(procMem)
    GetProcessMemoryInfo hWnd, procMem, procMem.cb
    CloseHandle hWnd
    getProcMem = procMem
End Function

Public Sub updateColArray()
    For i = 0 To 17
        If frmCol.chk(i).Value = 1 Then
            cols(i + 1) = 1
        Else
            cols(i + 1) = 0
        End If
    Next i

    cols(1) = 1
End Sub

Public Function getCTime(pid As Long) As String
Dim createTime As FILETIME, exitTime As FILETIME, kernelTime As FILETIME, userTime As FILETIME
Dim test As Integer
    hWnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    GetProcessTimes hWnd, createTime, exitTime, kernelTime, userTime
    Call convertTimes(createTime, exitTime, kernelTime, userTime)
    CloseHandle hWnd
    
    getCTime = parseSysTime(cSysTime)
End Function

Public Sub convertTimes(cTime As FILETIME, eTime As FILETIME, kTime As FILETIME, uTime As FILETIME)
    FileTimeToSystemTime cTime, cSysTime
    FileTimeToSystemTime eTime, eSysTime
    FileTimeToSystemTime kTime, kSysTime
    FileTimeToSystemTime uTime, uSysTime
End Sub

Public Function parseSysTime(cTime As SYSTEMTIME) As String
Dim dayof As Integer, monthof As String, dayS As String
    
    Select Case cTime.wDayOfWeek
    
        Case 0
            dayS = "Sun"
        Case 1
            dayS = "Mon"
        Case 2
            dayS = "Tues"
        Case 3
            dayS = "Wed"
        Case 4
            dayS = "Thur"
        Case 5
            dayS = "Fri"
        Case 6
            dayS = "Sat"
            
    End Select
    
    Select Case cTime.wMonth
    
        Case 1
            monthof = "Jan"
        Case 2
            monthof = "Feb"
        Case 3
            monthof = "Mar"
        Case 4
            monthof = "Apr"
        Case 5
            monthof = "May"
        Case 6
            monthof = "Jun"
        Case 7
            monthof = "Jul"
        Case 8
            monthof = "Aug"
        Case 9
            monthof = "Sept"
        Case 10
            monthof = "Oct"
        Case 11
            monthof = "Nov"
        Case 12
            monthof = "Dec"
            
    End Select
    dayof = cTime.wDay
    
    If cTime.wHour > 5 And cTime.wHour < 18 Then
        If cTime.wHour - 5 = 12 Then
            timeof = "12:" & cTime.wMinute & "pm"
        Else
            timeof = cTime.wHour - 5 & ":" & cTime.wMinute & "am"
        End If
    ElseIf cTime.wHour > 17 And cTime.wHour < 25 Then
        timeof = cTime.wHour - 17 & ":" & cTime.wMinute & "pm"
    ElseIf cTime.wHour > 0 And cTime.wHour < 6 Then
        If cTime.wHour + 7 = 12 Then
            timeof = "12:" & cTime.wMinute & "am"
        Else
            timeof = cTime.wHour + 7 & ":" & cTime.wMinute & "pm"
        End If
    End If
    parseSysTime = dayS & " " & monthof & " " & dayof & ", " & timeof
End Function

Public Sub updateProcStats()
    If servProc > 0 Then
        servProc = servProc + 1
    End If
    If uknProc > 0 Then
        uknProc = uknProc + 1
    End If
    frmProc.lblProc.Caption = frmProc.lstvwProc.ListItems.Count
    frmProc.lblServProc.Caption = servProc
    frmProc.lblUnkProc.Caption = uknProc
    frmProc.lblSysProc.Caption = sysProc
End Sub
