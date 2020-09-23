VERSION 5.00
Begin VB.Form frmCol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Columns"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Restore Defaults"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chk 
      Caption         =   "Process Type"
      Height          =   255
      Index           =   17
      Left            =   4080
      TabIndex        =   17
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CheckBox chk 
      Caption         =   "Creation Time"
      Height          =   255
      Index           =   16
      Left            =   4080
      TabIndex        =   16
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox chk 
      Caption         =   "Parent Process"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Base Priority"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Exe Name"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Process Priority"
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Page Faults"
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Peak Working Set Size"
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.CheckBox chk 
      Caption         =   "Working Set Size"
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Peak Paged Pool Usage"
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox chk 
      Caption         =   "Paged Pool Usage"
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Peak Non Paged Pool Usage"
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox chk 
      Caption         =   "Non Paged Pool Usage"
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.CheckBox chk 
      Caption         =   "Page File Usage"
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Peak Page FIle Usage"
      Height          =   255
      Index           =   15
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox chk 
      Caption         =   "Threads"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Process ID"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chk 
      Caption         =   "Process Name"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   2  'Grayed
      Width           =   1695
   End
End
Attribute VB_Name = "frmCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
   Call updateColArray
   Call frmProc.updateCol
   Call enumProc
   frmCol.Hide
End Sub

Private Sub cmdDefault_Click()
    For i = 0 To 17
        chk(i).Value = 0
    Next i
    chk(0).Value = 1
    chk(1).Value = 1
    chk(2).Value = 1
    chk(3).Value = 1
    chk(5).Value = 1
    chk(6).Value = 1
End Sub

