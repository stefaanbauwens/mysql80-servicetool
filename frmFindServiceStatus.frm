VERSION 5.00
Begin VB.Form frmFindServiceStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MySQL80 Service Tool"
   ClientHeight    =   2280
   ClientLeft      =   3390
   ClientTop       =   2865
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "frmFindServiceStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7500
   Begin VB.Timer tTimer 
      Interval        =   1000
      Left            =   8000
      Top             =   0
   End
   Begin VB.Frame fFrame 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   7455
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   600
         Left            =   2925
         TabIndex        =   3
         Top             =   1575
         Width           =   1635
      End
      Begin VB.CommandButton cmdServiceStart 
         Caption         =   "Service Start"
         Height          =   525
         Left            =   3750
         TabIndex        =   2
         Top             =   735
         Width           =   3525
      End
      Begin VB.CommandButton cmdServiceStop 
         Caption         =   "Service Stop"
         Height          =   525
         Left            =   105
         TabIndex        =   1
         Top             =   735
         Width           =   3525
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Start, Stop, Pause and get status of MySQL80 Service"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   7155
      End
      Begin VB.Line lLine 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         X1              =   135
         X2              =   7260
         Y1              =   1470
         Y2              =   1485
      End
   End
End
Attribute VB_Name = "frmFindServiceStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RetServiceName As String
Dim sState As String

Public Function ServiceStatus(ComputerName As String, ServiceName As String) As String
    Dim ServiceStat As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim hServiceStatus As Long

    ServiceStatus = ""
    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            hServiceStatus = QueryServiceStatus(hService, ServiceStat)
            If hServiceStatus <> 0 Then
                Select Case ServiceStat.dwCurrentState
                Case SERVICE_STOPPED
                    ServiceStatus = "Stopped"
                Case SERVICE_START_PENDING
                    ServiceStatus = "Start Pending"
                Case SERVICE_STOP_PENDING
                    ServiceStatus = "Stop Pending"
                Case SERVICE_RUNNING
                    ServiceStatus = "Running"
                Case SERVICE_CONTINUE_PENDING
                    ServiceStatus = "Coninue Pending"
                Case SERVICE_PAUSE_PENDING
                    ServiceStatus = "Pause Pending"
                Case SERVICE_PAUSED
                    ServiceStatus = "Paused"
                End Select
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Function

Public Sub ServicePause(ComputerName As String, ServiceName As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_PAUSE, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub

Public Sub ServiceStart(ComputerName As String, ServiceName As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = StartService(hService, 0, 0)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub

Public Sub ServiceStop(ComputerName As String, ServiceName As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_STOP, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub

Private Sub cmdServiceStart_Click()
    If RetServiceName = "" Then Call AskParaMeter
    ServiceStart "", RetServiceName
End Sub

Private Sub cmdServiceStop_Click()
    If RetServiceName = "" Then Call AskParaMeter
    ServiceStop "", RetServiceName
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "Minimize" Then
        frmFindServiceStatus.WindowState = 1
    Else
        Unload Me
        End
    End If
End Sub

Public Sub AskParaMeter()
    RetServiceName = "MySQL80"
End Sub

Private Sub Form_Load()
    tTimer_Timer
End Sub

Private Sub tTimer_Timer()
    If RetServiceName = "" Then Call AskParaMeter
    sState = ServiceStatus("", Trim(RetServiceName))
    lblCaption.Caption = "[" & Time$ & "] " & RetServiceName & " service status is : " & sState
    frmFindServiceStatus.Caption = RetServiceName & ": " & sState
    If sState = "Stopped" Then
        cmdClose.Caption = "Close"
    Else
        cmdClose.Caption = "Minimize"
    End If
    If sState = "" Then
        MsgBox "Run this tool as adminstrator !!!", vbCritical, App.Title
        End
    End If
End Sub
