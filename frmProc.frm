VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Monitor"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9340
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmProc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdFirewallTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdMemory"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdtabProcinfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdProcList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Running Processes"
      TabPicture(1)   =   "frmProc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "lstvwProc"
      Tab(1).Control(4)=   "cmdStopTimer"
      Tab(1).Control(5)=   "txtTimer"
      Tab(1).Control(6)=   "cmdStartTimer"
      Tab(1).Control(7)=   "tmrEnum"
      Tab(1).Control(8)=   "Command1"
      Tab(1).Control(9)=   "Command5"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Process Information"
      TabPicture(2)   =   "frmProc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label17"
      Tab(2).Control(1)=   "lblInfoParent"
      Tab(2).Control(2)=   "Label22"
      Tab(2).Control(3)=   "lblInfoPID"
      Tab(2).Control(4)=   "lblInfoType"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "lblInfoProcName"
      Tab(2).Control(7)=   "Frame2"
      Tab(2).Control(8)=   "Frame3"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Internet Processes"
      TabPicture(3)   =   "frmProc.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblMonitor"
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(2)=   "lblArc"
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(4)=   "lstvwNetProc"
      Tab(3).Control(5)=   "cmdEnumPortProc"
      Tab(3).Control(6)=   "cmdCloseConn"
      Tab(3).Control(7)=   "cmdMonitor"
      Tab(3).Control(8)=   "cmbMonitor"
      Tab(3).Control(9)=   "tmrMonitor"
      Tab(3).Control(10)=   "txtMonitor"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Firewall"
      TabPicture(4)   =   "frmProc.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblAttempts"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lblFirewallStat"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cd1"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdLoad"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "tmrFirewall"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdSave"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Frame6"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Frame5"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Frame4"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "cmdFirewall"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "cmbFirewall"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).ControlCount=   12
      Begin VB.CommandButton cmdProcList 
         Caption         =   "Running Processes"
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   120
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Caption         =   "Summary"
         Height          =   2175
         Left            =   2520
         TabIndex        =   65
         Top             =   120
         Width           =   2775
         Begin VB.Label Label1 
            Caption         =   "Processes Running:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "System Processes:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Unknown Processes:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblProc 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2160
            TabIndex        =   70
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblSysProc 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2160
            TabIndex        =   69
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblUnkProc 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2160
            TabIndex        =   68
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Service Processes:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblServProc 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2160
            TabIndex        =   66
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdtabProcinfo 
         Caption         =   "Process Information"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdMemory 
         Caption         =   "Internet Processes"
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdFirewallTab 
         Caption         =   "Firewall"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Enumerate Processes"
         Height          =   375
         Left            =   -74880
         TabIndex        =   61
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Select Columns..."
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   480
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Controls"
         Height          =   3735
         Left            =   -70800
         TabIndex        =   54
         Top             =   720
         Width           =   1215
         Begin VB.CommandButton cmdGetProc 
            Caption         =   "Get Process.."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdTabRunProc 
            Caption         =   "Running Proc"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdKillProc 
            Caption         =   "Kill Process.."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton cmdSetPri 
            Caption         =   "Set Priority.."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Process Information"
         Height          =   3735
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   4095
         Begin VB.Label lblInfoSet 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1440
            TabIndex        =   53
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label lblInfoPkPgPool 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1680
            TabIndex        =   52
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblInfoPgPool 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1440
            TabIndex        =   51
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label lblInfoPkPool 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2040
            TabIndex        =   50
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lblInfoPool 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1800
            TabIndex        =   49
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label lblInfoPkPgFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1800
            TabIndex        =   48
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Label lblInfoPgFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1560
            TabIndex        =   47
            Top             =   3360
            Width           =   2055
         End
         Begin VB.Label lblInfoPkSet 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblinfoPgFaults 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3360
            TabIndex        =   45
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblInfoThreads 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3360
            TabIndex        =   44
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblInfoProcPri 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1560
            TabIndex        =   43
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblInfoBase 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1320
            TabIndex        =   42
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblInfoExe 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1080
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label label92 
            Caption         =   "Pk Non-Pg Pool Usage:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label label91 
            Caption         =   "Non-Pg Pool Usage:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label label90 
            Caption         =   "Pk Page File Usage:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label label89 
            Caption         =   "Page File Usage:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label label99 
            Caption         =   "EXE Name:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   855
         End
         Begin VB.Label label97 
            Caption         =   "Process Priority:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblInfoFault 
            Caption         =   "Page Faults:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   34
            Top             =   600
            Width           =   975
         End
         Begin VB.Label label93 
            Caption         =   "Pg Pool Usage:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label label98 
            Caption         =   "Base Priority:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lbael95 
            Caption         =   "Working Set Sz:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label label94 
            Caption         =   "Pk Pg Pool Usage:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label label96 
            Caption         =   "Pk Working Set Sz:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Threads:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Timer tmrEnum 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -73080
         Top             =   120
      End
      Begin VB.CommandButton cmdStartTimer 
         Caption         =   "Start Timer"
         Height          =   255
         Left            =   -71520
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtTimer 
         Height          =   285
         Left            =   -71160
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdStopTimer 
         Caption         =   "Stop Timer"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -70560
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdEnumPortProc 
         Caption         =   "Enumerate Processes"
         Height          =   555
         Left            =   -74880
         TabIndex        =   22
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdCloseConn 
         Caption         =   "Close Process Connection"
         Height          =   375
         Left            =   -74880
         TabIndex        =   21
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdMonitor 
         Caption         =   "Enable Monitor"
         Height          =   375
         Left            =   -71400
         TabIndex        =   20
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox cmbMonitor 
         Height          =   315
         ItemData        =   "frmProc.frx":008C
         Left            =   -71400
         List            =   "frmProc.frx":009F
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin VB.Timer tmrMonitor 
         Enabled         =   0   'False
         Interval        =   800
         Left            =   -72720
         Top             =   1080
      End
      Begin VB.TextBox txtMonitor 
         Height          =   285
         Left            =   -71400
         TabIndex        =   18
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cmbFirewall 
         Height          =   315
         ItemData        =   "frmProc.frx":00E1
         Left            =   -74880
         List            =   "frmProc.frx":00F1
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdFirewall 
         Caption         =   "Enable Firewall"
         Height          =   375
         Left            =   -71760
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Options"
         Height          =   855
         Left            =   -72480
         TabIndex        =   14
         Top             =   120
         Width           =   3015
         Begin VB.CheckBox chkBlock 
            Caption         =   "Block By Default"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Rules"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   2295
         Begin VB.ListBox lstRemoteIP 
            Height          =   3375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ListBox lstRemotePort 
            Height          =   3375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ListBox lstLocalPort 
            Height          =   3375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ListBox lstProcessName 
            Height          =   3375
            ItemData        =   "frmProc.frx":0127
            Left            =   120
            List            =   "frmProc.frx":0129
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear List"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Rule"
            Height          =   375
            Left            =   1080
            TabIndex        =   8
            Top             =   3600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "New Rule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -72480
         TabIndex        =   3
         Top             =   3120
         Width           =   3015
         Begin VB.TextBox txtRule 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton cmdAddRule 
            Caption         =   "Add Rule"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblBlock 
            Caption         =   "Block if (rule) equals:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Rules"
         Height          =   255
         Left            =   -70680
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Timer tmrFirewall 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   -70080
         Top             =   1080
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Rules"
         Height          =   255
         Left            =   -70680
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   -72360
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstvwNetProc 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   23
         Top             =   1920
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstvwProc 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   60
         Top             =   1320
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblInfoProcName 
         Caption         =   "No Process Selected"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   91
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   90
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblInfoType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74280
         TabIndex        =   89
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblInfoPID 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70320
         TabIndex        =   88
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Process ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71160
         TabIndex        =   87
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblInfoParent 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -71160
         TabIndex        =   86
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Parent Process:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72360
         TabIndex        =   85
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Click on a process to get more info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   84
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Auto-Enumerate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71040
         TabIndex        =   83
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Timer Interval (s):"
         Height          =   255
         Left            =   -72480
         TabIndex        =   82
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblMonitor 
         Caption         =   "Status: Not Monitoring"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -71880
         TabIndex        =   81
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "Monitor For:"
         Height          =   255
         Left            =   -72360
         TabIndex        =   80
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblArc 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   -72720
         TabIndex        =   79
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Click on a process to get more info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   78
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label lblFirewallStat 
         Alignment       =   2  'Center
         Caption         =   "Status: Disabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -72600
         TabIndex        =   77
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label lblAttempts 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   -70080
         TabIndex        =   76
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Blocked Processes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72480
         TabIndex        =   75
         Top             =   2280
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFirewall_Click()
    Frame6.Enabled = True
    Select Case cmbFirewall.List(cmbFirewall.ListIndex)
    
        Case "Process Name"
            ruleType = 1
            lblBlock.Caption = "Block if " & cmbFirewall.List(cmbFirewall.ListIndex) & " equals:"
            lstProcessName.Visible = True
            lstRemoteIP.Visible = False
            lstRemotePort.Visible = False
            lstLocalPort.Visible = False
        Case "Remote IP"
            ruleType = 2
            lblBlock.Caption = "Block if " & cmbFirewall.List(cmbFirewall.ListIndex) & " equals:"
            lstProcessName.Visible = False
            lstRemoteIP.Visible = True
            lstRemotePort.Visible = False
            lstLocalPort.Visible = False
        Case "Remote Port"
            ruleType = 3
            lblBlock.Caption = "Block if " & cmbFirewall.List(cmbFirewall.ListIndex) & " equals:"
            lstProcessName.Visible = False
            lstRemoteIP.Visible = False
            lstRemotePort.Visible = True
            lstLocalPort.Visible = False
        Case "Local Port"
            ruleType = 4
            lblBlock.Caption = "Block if " & cmbFirewall.List(cmbFirewall.ListIndex) & " equals:"
            lstProcessName.Visible = False
            lstRemoteIP.Visible = False
            lstRemotePort.Visible = False
            lstLocalPort.Visible = True
            
    End Select
End Sub

Private Sub cmbMonitor_Click()
    lblArc.Caption = cmbMonitor.List(cmbMonitor.ListIndex) & ":"
End Sub

Private Sub cmdAddRule_Click()
Select Case cmbFirewall.List(cmbFirewall.ListIndex)
    
        Case "Process Name"
            If txtRule.Text <> "" Then
                lstProcessName.AddItem txtRule.Text
                txtRule.Text = ""
            Else
                MsgBox "No Rule Defined", vbOKOnly, "No Rule"
            End If
        Case "Remote IP"
            If txtRule.Text <> "" Then
                lstRemoteIP.AddItem txtRule.Text
                txtRule.Text = ""
            Else
                MsgBox "No Rule Defined", vbOKOnly, "No Rule"
            End If
        Case "Remote Port"
            If txtRule.Text <> "" Then
                lstRemotePort.AddItem txtRule.Text
                txtRule.Text = ""
            Else
                MsgBox "No Rule Defined", vbOKOnly, "No Rule"
            End If
        Case "Local Port"
            If txtRule.Text <> "" Then
                lstLocalPort.AddItem txtRule.Text
                txtRule.Text = ""
            Else
                MsgBox "No Rule Defined", vbOKOnly, "No Rule"
            End If
    End Select
End Sub

Private Sub cmdClear_Click()
    Select Case cmbFirewall.List(cmbFirewall.ListIndex)
    
        Case "Process Name"
            lstProcessName.Clear
        Case "Remote IP"
            lstRemoteIP.Clear
        Case "Remote Port"
           lstRemotePort.Clear
        Case "Local Port"
           lstLocalPort.Clear
    End Select
End Sub

Private Sub cmdCloseConn_Click()
Dim tempid As Long
On Error Resume Next
    tempid = lstvwNetProc.ListItems(lstvwNetProc.SelectedItem.Index).SubItems(1)
    tempProcName = InputBox("Enter Process ID", "Terminate Connection", tempid)
    checkforID = 1
    Call RefreshStack
    Call EnumEntries
End Sub

Private Sub cmdDelete_Click()
    Select Case cmbFirewall.List(cmbFirewall.ListIndex)
        
        Case "Process Name"
            If lstProcessName.ListIndex > -1 Then
                lstProcessName.RemoveItem lstProcessName.ListIndex
            End If
        Case "Remote IP"
            If lstRemoteIP.ListIndex > -1 Then
                lstRemoteIP.RemoveItem lstRemoteIP.ListIndex
            End If
        Case "Remote Port"
           If lstRemotePort.ListIndex > -1 Then
                lstRemotePort.RemoveItem lstRemotePort.ListIndex
            End If
        Case "Local Port"
           If lstLocalPort.ListIndex > -1 Then
                lstLocalPort.RemoveItem lstLocalPort.ListIndex
            End If
    End Select
End Sub

Private Sub cmdEnumPortProc_Click()
    dontlbl = 1
    Call RefreshStack
    Call EnumEntries
    dontlbl = 0
End Sub

Private Sub cmdFirewall_Click()
    If firewallStatus <> 1 Then
        If chkBlock.Value = 1 Then
            block = 0
        Else
            block = 1
        End If
        If chkBlock.Value = 1 Then
            If lstProcessName.ListCount = 0 And lstRemoteIP.ListCount = 0 And lstRemotePort.ListCount = 0 And lstLocalPort.ListCount = 0 Then
                MsgBox "Cannot Enable firewall. No rules Specified", vbCritical, "Error"
                Exit Sub
            End If
        End If
        firewallStatus = 1
        lblFirewallStat.Caption = "Status: Enabled"
        lblFirewallStat.ForeColor = &HC000&
        cmdFirewall.Caption = "Disable Firewall"
        tmrFirewall.Enabled = True
    Else
        firewallStatus = 0
        lblFirewallStat.Caption = "Status: Disabled"
        lblFirewallStat.ForeColor = &HFF&
        cmdFirewall.Caption = "Enable Firewall"
        tmrFirewall.Enabled = False
    End If
End Sub

Private Sub cmdFirewallTab_Click()
    Tab1.Tab = 4
End Sub

Private Sub cmdGetProc_Click()
Dim pName As String
Dim temp As MSComctlLib.ListItem
    pName = InputBox("Enter Process Name", "Get Process Info")
    If pName <> "" Then
        external = 1
        infoProc = pName
        Call lstvwProc_ItemClick(temp)
        external = 0
    End If
End Sub

Private Sub cmdKillProc_Click()
On Error Resume Next
Dim checkFail As Long
Dim pid As Long
    pid = InputBox("Enter Process ID to terminate", "Terminate Process", infoProcInfo.th32ProcessID)
    If pid > 0 Then
        hWndof = OpenProcess(PROCESS_TERMINATE, False, pid)
        checkFail = TerminateProcess(hWndof, 0)
        If checkFail = 0 Then
            MsgBox "Unable to terminate process." & vbNewLine & "Verify process is not of type System/Service", vbCritical, "Error"
        End If
        CloseHandle hWndof
    End If

End Sub

Private Sub cmdLoad_Click()
Dim fcgName As String
Dim linef As String
Dim numLst As Integer
    cd1.DefaultExt = "fcg"
    cd1.Filter = "Firewall Rules List (*.FCG)|*.FCG|All Files (*.*)|*.*"
    cd1.ShowOpen
    fcgName = cd1.FileName
    If fcgName = "" Then
        Exit Sub
    End If
    Open fcgName For Input As #1
        Do While Not EOF(1)
            Line Input #1, linef
            numLst = Left(linef, 1)
            Select Case numLst
            
                Case 1
                    lstProcessName.AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case 2
                    lstRemoteIP.AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case 3
                    lstRemotePort.AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
                Case 4
                    lstLocalPort.AddItem Right(linef, Len(linef) - InStr(1, linef, ":"))
            End Select
        Loop
    Close #1
    MsgBox "Rule List Loaded Sucsessfully!", vbOKOnly, "Firewall"
End Sub

Private Sub cmdMemory_Click()
    Tab1.Tab = 3
End Sub

Private Sub cmdMonitor_Click()
    If monitor = 1 Then
        tmrMonitor.Enabled = False
        cmdMonitor.Caption = "Enable Monitor"
        lblMonitor.Caption = "Status: Not Monitoring"
        monitor = 0
        monitorType = ""
    ElseIf monitor = 0 Then
        If cmbMonitor.List(cmbMonitor.ListIndex) <> "" Then
            tmrMonitor.Enabled = True
            monitor = 1
            lblMonitor.Caption = "Status: Monitoring"
            cmdMonitor.Caption = "Stop Monitoring"
            monitorType = cmbMonitor.List(cmbMonitor.ListIndex)
            monitorFor = txtMonitor.Text
        End If
    End If
End Sub

Private Sub cmdProcList_Click()
    Tab1.Tab = 1
End Sub

Private Sub cmdSave_Click()
Dim fcgName As String
    cd1.DialogTitle = "Save Rule List As..."
    cd1.Filter = "Firewall Rules List (*.FCG)|*.FCG|All Files (*.*)|*.*"
    cd1.DefaultExt = "fcg"
    cd1.ShowSave
    fcgName = cd1.FileName
    If fcgName <> "" Then
        Open fcgName For Append As #1
        Close #1
        Open fcgName For Output As #1
        
        For i = 0 To lstProcessName.ListCount - 1
            Print #1, "1:" & lstProcessName.List(i)
        Next i
        
        For i = 0 To lstRemoteIP.ListCount - 1
            Print #1, "2:" & lstRemoteIP.List(i)
        Next i
        
        For i = 0 To lstRemotePort.ListCount - 1
            Print #1, "3:" & lstRemotePort.List(i)
        Next i
        
        For i = 0 To lstLocalPort.ListCount - 1
            Print #1, "4:" & lstLocalPort.List(i)
        Next i
        
        Close #1
        MsgBox "Rule List Saved", vbOKOnly, "Firewall"
    End If
End Sub

Private Sub cmdSetPri_Click()
    Load frmPriority
    frmPriority.Show
End Sub

Private Sub cmdStartTimer_Click()
    If txtTimer.Text <> "" Then
        If txtTimer.Text < 1 Or txtTimer.Text > 99 Then
            MsgBox "Invalid value. Must be between 1 and 99", vbCritical, "Invalid Value"
            txtTimer.Text = ""
            Exit Sub
        End If
    End If
    cmdStartTimer.Enabled = False
    cmdStopTimer.Enabled = True
    Command5.Enabled = False
    Command1.Enabled = False
    If txtTimer.Text <> "" Then
        tmrEnum.Interval = (txtTimer.Text * 1000)
    End If
    txtTimer.Enabled = False
    tmrEnum.Enabled = True
End Sub

Private Sub cmdStopTimer_Click()
    tmrEnum.Enabled = False
    cmdStartTimer.Enabled = True
    cmdStopTimer.Enabled = False
    txtTimer.Enabled = True
    Command5.Enabled = True
    Command1.Enabled = True
    tmrEnum.Interval = 1000
End Sub

Private Sub cmdtabProcinfo_Click()
    Tab1.Tab = 2
End Sub

Private Sub cmdTabRunProc_Click()
    Tab1.Tab = 1
End Sub

Private Sub Command1_Click()
    frmCol.Show
End Sub

Private Sub Command5_Click()
    tempClear = 1
    Call enumProc
End Sub

Private Sub Form_Load()
    blocked = 0
    monitor = 0
    dontlbl = 0
    noClear = 0
    tempClear = 0
    monitorType = ""
    popView = 0
    refreshPort = 0
    checkPID = 0
    Call enumProc 'To get Needed PID's. Does not populate listview
    Load frmCol
    Call updateColArray
    Call updateCol
    frmCol.Hide
    Call addNetCols
    If isWinXp = False Then
        Call disableAll
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmCol
End Sub

Private Sub lblInfoParent_Click()
Dim temp As MSComctlLib.ListItem
    If lblInfoParent.Caption <> "" Then
        infoProc = Left(lblInfoParent.Caption, InStr(1, lblInfoParent.Caption, " ") - 1)
        If InStr(1, infoProc, "exe") > 1 Then
            external = 1
            Call lstvwProc_ItemClick(temp)
            external = 0
        End If
    End If
End Sub

Private Sub lblInfoProcPri_Click()
    If lblInfoProcPri.Caption <> "" Then
        Call cmdSetPri_Click
    End If
End Sub

Private Sub lstvwNetProc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstvwNetProc.Sorted = True
    If lstvwNetProc.SortOrder = lvwAscending Then
        lstvwNetProc.SortOrder = lvwDescending
    Else
        lstvwNetProc.SortOrder = lvwAscending
    End If
    lstvwNetProc.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lstvwNetProc_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim temp As MSComctlLib.ListItem
    external = 1
    infoProc = Item.Text
    Call lstvwProc_ItemClick(temp)
    external = 0
End Sub

Private Sub lstvwProc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstvwProc.Sorted = True
    If lstvwProc.SortOrder = lvwAscending Then
        lstvwProc.SortOrder = lvwDescending
    Else
        lstvwProc.SortOrder = lvwAscending
    End If
    lstvwProc.SortKey = ColumnHeader.Index - 1
End Sub

Public Sub updateCol()
    lstvwProc.ColumnHeaders.Clear
    If cols(1) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "name", "Process Name", TextWidth("Process Name") * 1.5)
    End If
    
    If cols(2) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pid", "PID", 650)
    End If
    
    If cols(3) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "thread", "Threads", TextWidth("Threads") * 1.3)
    End If
    
    If cols(4) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "parent", "Parent", TextWidth("Parent") * 2.2)
    End If
    
    If cols(5) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pri", "Base Priority", TextWidth("Base Priority") * 1.5)
    End If
    
    If cols(6) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "exe", "EXE Name", TextWidth("EXE NAME") * 1.8)
    End If
    
    If cols(7) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pri2", "Process Priority", TextWidth("Process Priority") * 1.5)
    End If
    
    If cols(8) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "fault", "Page Faults", TextWidth("Page Faults") * 1.3)
    End If
    
    If cols(9) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pkwrkset", "PeakWorkingSetSz", TextWidth("PeakWorkingSetSz") * 1.2)
    End If
    
    If cols(10) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "wrkset", "WorkingSetSz", TextWidth("WorkingSetSz") * 1.5)
    End If
    
    If cols(11) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "peakpool", "PagedPeakPoolUsage", TextWidth("PagedPeakPoolUsage") * 1.2)
    End If
    
    If cols(12) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pool", "PagedPoolUsage", TextWidth("PagedPoolUsage") * 1.2)
    End If
    
    If cols(13) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "nonpeakpool", "NonPagedPeakPoolUsage", TextWidth("NonPeakPagedPoolUsage") * 1.1)
    End If
    
    If cols(14) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "nonpool", "NonPagedPoolUsage", TextWidth("NonPagedPoolUsage") * 1.2)
    End If
    
    If cols(15) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pgfile", "PagefileUsage", TextWidth("PagefileUsage") * 1.3)
    End If
    
    If cols(16) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "pkpgfile", "PeakPagefileUsage", TextWidth("PeakPagefileUsage") * 1.3)
    End If
    
    If cols(17) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "cTime", "Creation Date", TextWidth("Creation Date") * 1.9)
    End If
    
    If cols(18) = 1 Then
        Set colHead = lstvwProc.ColumnHeaders.Add(lstvwProc.ColumnHeaders.Count + 1, "Typ", "Type", TextWidth("Type") * 2.5)
    End If
End Sub

Public Sub lstvwProc_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim infoType As String
Dim priPrc As String
    found = 0
    If external <> 1 Then
        infoProc = Item.Text
    End If
    If checkPID <> 0 Then
        external = checkPID
        checkPID = 1
    End If
    parentProc = ""
    infoCheck = 1
    Call enumProc
    infoCheck = 0
    If found = 0 Then
        MsgBox "Process " & infoProc & " not found", vbOKOnly, "Invalid Process"
        Exit Sub
    End If
    checkParent = 1
    Call enumProc
    checkParent = 0
    
    If infoProcInfo.th32ParentProcessID = servPID Then
        infoType = "Service"
    ElseIf infoProcInfo.th32ParentProcessID = expPID Then
        infoType = "Unknown"
    ElseIf infoProcInfo.th32ParentProcessID = sysPID(1) Or infoProcInfo.th32ParentProcessID = sysPID(2) Or infoProcInfo.th32ParentProcessID = sysPID(3) Or infoProcInfo.th32ParentProcessID = sysPID(4) Or infoProcInfo.th32ParentProcessID = sysPID(5) Then
        infoType = "System"
    Else
        infoType = "Unknown"
    End If
    
    Select Case infoProcPri
    
    Case 32
        priPrc = "Normal - 32"
    Case 64
        priPrc = "Idle - 64"
    Case 128
        priPrc = "High - 128"
    Case 256
        priPrc = "RealTime - 256"
    End Select
    
    
    lblInfoProcName.Caption = infoProc
    lblInfoPID.Caption = infoProcInfo.th32ProcessID
    lblInfoType.Caption = infoType
    lblInfoParent.Caption = parentProc & " - " & infoProcInfo.th32ParentProcessID
    lblInfoExe.Caption = infoProcInfo.szExeFile
    lblInfoBase.Caption = infoProcInfo.pcPriClassBase
    lblInfoProcPri.Caption = priPrc
    lblInfoThreads.Caption = infoProcInfo.cntThreads
    lblinfoPgFaults.Caption = infoProcMem.PageFaultCount
    lblInfoPkSet.Caption = infoProcMem.PeakWorkingSetSize
    lblInfoSet.Caption = infoProcMem.WorkingSetSize
    lblInfoPkPgPool.Caption = infoProcMem.QuotaPeakPagedPoolUsage
    lblInfoPgPool.Caption = infoProcMem.QuotaPagedPoolUsage
    lblInfoPkPool.Caption = infoProcMem.QuotaPeakNonPagedPoolUsage
    lblInfoPool.Caption = infoProcMem.QuotaNonPagedPoolUsage
    lblInfoPkPgFile.Caption = infoProcMem.PeakPagefileUsage
    lblInfoPgFile.Caption = infoProcMem.PagefileUsage
    
    Tab1.Tab = 2
End Sub

Private Sub tmrEnum_Timer()
    tempClear = 1
    Call enumProc
End Sub

Public Sub addNetCols()
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "Process Name", "Process Name", TextWidth("Process Name") * 1.5)
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "Process ID", "Process ID", TextWidth("Process ID") * 1.3)
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "Local Address", "Local Address", TextWidth("Local Address") * 1.5)
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "Local Port", "Local Port", TextWidth("Local Port") * 1.5)
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "Remote Address", "Remote Address", TextWidth("Remote Address") * 1.5)
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "Remote Port", "Remote Port", TextWidth("Remote Port") * 1.3)
    Set colHead = lstvwNetProc.ColumnHeaders.Add(lstvwNetProc.ColumnHeaders.Count + 1, "State", "State", TextWidth("State") * 4)
End Sub

Private Sub tmrFirewall_Timer()
    dontlbl = 1
    Call RefreshStack
    Call EnumEntries
    dontlbl = 0
    lblAttempts.Caption = blocked
End Sub

Private Sub tmrMonitor_Timer()
    dontlbl = 1
    Call RefreshStack
    Call EnumEntries
    dontlbl = 0
End Sub
