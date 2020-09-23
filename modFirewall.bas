Attribute VB_Name = "modFirewall"
Public ruleType As Integer
Public firewallStatus As Integer
Public block As Integer
Public blocked As Integer
Private Function Registry_Read(Key_Path, Key_Name) As Variant
    
    On Error Resume Next
    
    Dim Registry As Object
    
    Set Registry = CreateObject("WScript.Shell")
    
    Registry_Read = Registry.regread(Key_Path & Key_Name)
    
End Function

Public Function isWinXp() As Boolean
    
    Dim Operating_System As String

    Operating_System = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\", "PRODUCTNAME")

    If Operating_System = "" Then

         Operating_System = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\", "PRODUCTNAME")

    End If
    
    If UCase(Operating_System) = UCase("microsoft windows xp") Then
        isWinXp = True
    Else
        isWinXp = False
    End If

End Function


Public Sub disableAll()
    
    frmProc.cmdAddRule.Enabled = False
    frmProc.cmdClear.Enabled = False
    frmProc.cmdCloseConn.Enabled = False
    frmProc.cmdDelete.Enabled = False
    frmProc.cmdEnumPortProc.Enabled = False
    frmProc.cmdFirewall.Enabled = False
    frmProc.cmdSave.Enabled = False
    frmProc.cmdLoad.Enabled = False
    frmProc.cmbFirewall.Enabled = False
    frmProc.cmdMonitor.Enabled = False
    frmProc.cmbMonitor.Enabled = False
    frmProc.txtMonitor.Enabled = False
    frmProc.lblMonitor.Caption = "Windows XP Only"
    frmProc.Frame4.Enabled = False
    frmProc.Frame6.Enabled = False
    frmProc.lblFirewallStat.Caption = "Windows XP Only"
    frmProc.cmdFirewallTab.Enabled = False
End Sub
