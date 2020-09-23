VERSION 5.00
Begin VB.Form frmPriority 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Process Priority"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPriPID 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cmbPriority 
      Height          =   315
      ItemData        =   "frmPriority.frx":0000
      Left            =   960
      List            =   "frmPriority.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdPriNormal 
      Caption         =   "Set Default"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdApplyPri 
      Caption         =   "Change Priority"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Priority:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Process ID:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApplyPri_Click()
Dim pid As Long
Dim priHwnd As Long
Dim temp As MSComctlLib.ListItem
    If txtPriPID.Text <> "" Or txtPriPID.Text <> 0 Then
        pid = txtPriPID.Text
        priHwnd = OpenProcess(PROCESS_SET_INFORMATION, False, pid)
        
        Select Case cmbPriority.ListIndex
        
        Case 0
            SetPriorityClass priHwnd, REALTIME_PRIORITY_CLASS
        Case 1
            SetPriorityClass priHwnd, HIGH_PRIORITY_CLASS
        Case 2
            SetPriorityClass priHwnd, NORMAL_PRIORITY_CLASS
        Case 3
            SetPriorityClass priHwnd, IDLE_PRIORITY_CLASS
        End Select
        CloseHandle priHwnd
        external = 1
        checkPID = pid
        Call frmProc.lstvwProc_ItemClick(temp)
        external = 0
        checkPID = 0
        Unload frmPriority
    Else
        MsgBox "You must enter a valid Process ID", vbOKOnly, "Invalid Process ID"
        Exit Sub
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload frmPriority
End Sub

Private Sub cmdPriNormal_Click()
Dim pid As Long
Dim priHwnd As Long
Dim temp As MSComctlLib.ListItem
    If txtPriPID.Text <> "" Or txtPriPID.Text <> 0 Then
        pid = txtPriPID.Text
        priHwnd = OpenProcess(PROCESS_SET_INFORMATION, False, pid)
        SetPriorityClass priHwnd, NORMAL_PRIORITY_CLASS
        CloseHandle hWnd
        external = 1
        checkPID = pid
        Call frmProc.lstvwProc_ItemClick(temp)
        external = 0
        checkPID = 0
        Unload frmPriority
    Else
        MsgBox "You must enter a valid Process ID", vbOKOnly, "Invalid Process ID"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    txtPriPID.Text = infoProcInfo.th32ProcessID
End Sub
