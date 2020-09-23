Attribute VB_Name = "modProcPort"
'Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
'Const MIB_TCP_STATE_CLOSED As Long = 1
'Const MIB_TCP_STATE_CLOSING As Long = 9
'Const MIB_TCP_STATE_DELETE_TCB As Long = 12
'Const MIB_TCP_STATE_ESTAB As Long = 5
'Const MIB_TCP_STATE_FIN_WAIT1 As Long = 6
'Const MIB_TCP_STATE_FIN_WAIT2 As Long = 7
'Const MIB_TCP_STATE_LAST_ACK As Long = 10
'Const MIB_TCP_STATE_LISTEN As Long = 2
'Const MIB_TCP_STATE_SYN_RCVD As Long = 4
'Const MIB_TCP_STATE_SYN_SENT As Long = 3
'Const MIB_TCP_STATE_TIME_WAIT As Long = 11
Public refreshPort As Integer
Public tempPID As Long
Public tempName As String
Public foundName As Integer
Public checkforID As Integer
Public tempProcName As Long
Public monitor As Integer
Public monitorType As String
Public monitorFor As Variant
Public dontlbl As Integer

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private Declare Function GetProcessHeap Lib "kernel32" () As Long


Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long


Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (pTcpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long


Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private pTablePtr As Long
    Private pDataRef As Long
    Private nRows As Long
    Private nCurrentRow As Long
    Private udtRow As MIB_TCPROW
    Private nState As Long
    Private nLocalAddr As Long
    Private nLocalPort As Long
    Private nRemoteAddr As Long
    Private nRemotePort As Long
    Private nProcId As Long


Public Function GetIPAddress(dwAddr As Long) As String
    Dim arrIpParts(3) As Byte
    CopyMemory arrIpParts(0), dwAddr, 4
    GetIPAddress = CStr(arrIpParts(0)) & "." & _
    CStr(arrIpParts(1)) & "." & _
    CStr(arrIpParts(2)) & "." & _
    CStr(arrIpParts(3))
End Function


Public Function GetPort(ByVal dwPort As Long) As Long
    GetPort = htons(dwPort)
End Function


Public Function RefreshStack() As Boolean
    Dim nRet As Long
    pDataRef = 0
    nRet = AllocateAndGetTcpExTableFromStack(pTablePtr, 0, GetProcessHeap, 0, 2)

    If nRet = 0 Then
        CopyMemory nRows, ByVal pTablePtr, 4
        RefreshStack = True
    Else
        RefreshStack = False
    End If
End Function


Public Function GetEntryCount() As Long
    GetEntryCount = nRows - 2 '// The last entry is always an EOF of sorts
End Function


Public Function EnumEntries() As Boolean
Dim i As Long
    If checkforID = 0 Then
        frmProc.lstvwNetProc.ListItems.Clear
    End If
    portPID = 0
    foundName = 0
    refreshPort = 1
    EnumEntries = True
    
    If nRows = 0 Or pTablePtr = 0 Then
        EnumEntries = False
        Exit Function
    End If


    For i = 0 To nRows '// read 24 bytes at a time
        CopyMemory nState, ByVal pTablePtr + (pDataRef + 4), 4
        CopyMemory nLocalAddr, ByVal pTablePtr + (pDataRef + 8), 4
        CopyMemory nLocalPort, ByVal pTablePtr + (pDataRef + 12), 4
        CopyMemory nRemoteAddr, ByVal pTablePtr + (pDataRef + 16), 4
        CopyMemory nRemotePort, ByVal pTablePtr + (pDataRef + 20), 4
        CopyMemory nProcId, ByVal pTablePtr + (pDataRef + 24), 4
    
        DoEvents
        If checkforID = 0 Then
            If nRemoteAddr <> 0 Or nRemotePort <> 0 Or nLocalPort <> 0 Then
                popView = 0
                tempPID = nProcId
                noClear = 1
                Call enumProc
                noClear = 0
                popView = 1
                If foundName = 0 Then
                    tempName = "Unknown"
                End If
                If nProcId < 70000 And nProcId > 0 And nState > 0 And nState < 13 Then
                    Set lstItem = frmProc.lstvwNetProc.ListItems.Add(, , tempName)
                    lstItem.SubItems(1) = nProcId
                    lstItem.SubItems(2) = GetIPAddress(nLocalAddr)
                    lstItem.SubItems(3) = GetPort(nLocalPort)
                    lstItem.SubItems(4) = GetIPAddress(nRemoteAddr)
                    lstItem.SubItems(5) = GetPort(nRemotePort)
                    lstItem.SubItems(6) = getState(nState) & " (" & nState & ")"
                    If monitor = 1 Then
                        Select Case monitorType
                        
                            Case "Process Name"
                                If UCase(tempName) = UCase(monitorFor) Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                End If
                            Case "Process ID"
                                If nProcId = monitorFor Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                End If
                            Case "Local Port"
                                If GetPort(nLocalPort) = monitorFor Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                End If
                            Case "Remote Port"
                                If GetPort(nRemotePort) = monitorFor Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                End If
                            Case "Remote ID"
                                If GetIPAddress(nRemoteAddr) = monitorFor Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                End If
                        End Select
                        
                    End If
                    
                    If firewallStatus = 1 Then
                        If block = 1 Then
                            If frmProc.lstProcessName.ListCount <> 0 Then
                                For z = 0 To frmProc.lstProcessName.ListCount - 1
                                    If UCase(tempName) = UCase(frmProc.lstProcessName.List(z)) Then
                                        TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    End If
                                Next z
                            End If
                            If frmProc.lstRemoteIP.ListCount <> 0 Then
                                For n = 0 To frmProc.lstRemoteIP.ListCount - 1
                                    If nRemoteAddr = frmProc.lstRemoteIP.List(n) Then
                                        TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        
                                    End If
                                Next n
                            End If
                            If frmProc.lstRemotePort.ListCount <> 0 Then
                                For m = 0 To frmProc.lstRemotePort.ListCount - 1
                                    If nRemotePort = frmProc.lstRemotePort.List(im) Then
                                        TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        
                                    End If
                                Next m
                            End If
                            If frmProc.lstLocalPort.ListCount <> 0 Then
                                For p = 0 To frmProc.lstLocalPort.ListCount - 1
                                    If nLocalPort = frmProc.lstLocalPort.List(p) Then
                                        TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                        
                                    End If
                                Next p
                            End If
                        ElseIf block = 0 Then
                            For z = 0 To frmProc.lstProcessName.ListCount - 1
                                If UCase(tempName) <> UCase(frmProc.lstProcessName.List(z)) Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    
                                End If
                            Next z
                            For n = 0 To frmProc.lstRemoteIP.ListCount - 1
                                If nRemoteAddr <> frmProc.lstRemoteIP.List(n) Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    
                                End If
                            Next n
                            For m = 0 To frmProc.lstRemotePort.ListCount - 1
                                If nRemotePort <> frmProc.lstRemotePort.List(m) Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                    
                                End If
                            Next m
                            For p = 0 To frmProc.lstLocalPort.ListCount - 1
                                If nLocalPort <> frmProc.lstLocalPort.List(p) Then
                                    TerminateThisConnection nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort
                                   
                                End If
                            Next p
                        End If
                    End If
                    
                End If
                foundName = 0
            End If
        Else
            If nProcId = tempProcName Then
                checkforID = 0
                Call TerminateThisConnection(nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort)
            End If
        End If
        
        pDataRef = pDataRef + 24
        DoEvents
    Next i
    refreshPort = 0
End Function


Public Sub TerminateThisConnection(xLocalAddr As Long, xLocalPort As Long, xRemoteAddr As Long, xRemotePort As Long)
    blocked = blocked + 1
    udtRow.dwLocalAddr = xLocalAddr
    udtRow.dwLocalPort = xLocalPort
    udtRow.dwRemoteAddr = xRemoteAddr
    udtRow.dwRemotePort = xRemotePort
    udtRow.dwState = 12
    SetTcpEntry udtRow
    Call RefreshStack
    Call EnumEntries
End Sub

Public Function getState(stateOf As Long) As String

    Select Case stateOf
    
        Case 1
            getState = "Closed"
        Case 2
            getState = "Listening"
        Case 3
            getState = "SYN Sent"
        Case 4
            getState = "SYN Recieved"
        Case 5
            getState = "Established"
        Case 6
            getState = "FIN Wait 1"
        Case 7
            getState = "FIN Wait 2"
        Case 8
            getState = "Close Wait"
        Case 9
            getState = "Closing"
        Case 10
            getState = "Last ACK"
        Case 11
            getState = "Time Wait"
        Case 12
            getState = "Delete TCB"
    End Select

End Function
