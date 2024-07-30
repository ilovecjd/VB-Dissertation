VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Simulator"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   LinkTopic       =   "MainForm"
   ScaleHeight     =   7785
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.OptionButton Option_Create 
      Caption         =   "Create"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton Option_Load 
      Caption         =   "Load"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton btnGenBoardNProject 
      Caption         =   "������Ʈ����"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox Hr_Init_M 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Hr_Init_H 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtWeeklyProb 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtSimulationWeeks 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ɼ�"
      Height          =   855
      Left            =   3120
      TabIndex        =   14
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "�����η�(�߱�)"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "�����η�(���)"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1515
      Width           =   1935
   End
   Begin VB.Label ������Ʈ�߻��� 
      Caption         =   "������Ʈ�߻���"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label SimulTearm 
      Caption         =   "�ð�(��)"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1


Const STR_DATA_FILE = "data.xlsm"
Const STR_RUN_LOG_FILE = "run_log.txt"
Const STR_START_EXCEL = "����"
Const STR_END_EXCEL = "����"

Enum LoadOrCreate
        Load
        Create
End Enum


Private Sub btnGenBoardNProject_Click()
    
    ' �Է°����� ������Ʈ �Ѵ�.
    GlobalEnv.WeeklyProb = txtWeeklyProb.Text
    'GlobalEnv.Cash_Init
    'GlobalEnv.Hr_Init_H
    'GlobalEnv.Hr_Init_L
    'GlobalEnv.Hr_Init_M
    'GlobalEnv.Hr_LeadTime
    'GlobalEnv.Problem
    GlobalEnv.SimulationWeeks = txtSimulationWeeks.Text
    
    ' ��ú��� ���� �Ǵ� �ε� -> ������Ʈ ���� �Ǵ� �ε�
    
        ' ���̺���� ���� �����ϰų� �������� �ε��ϰų�.
    ' ���� ó���� �������� ����ڰ� �����ؼ� ����ϵ��� ����.
    gTableInitialized = (TableInit = 1)
    
    If gTableInitialized = False Then
        Call BuildTables ' ���̺��� ���� �����ϰ� ������ ����Ѵ�.
        
        Call PrintProjectHeader ' Project ��Ʈ�� ����� ����Ѵ�.
        Call CreateProjects     ' ������Ʈ�� �����Ѵ�.
        
    Else
        Call LoadTablesFromExcel ' ������ ��ϵ� ����� ���̺��� ä���.
    End If


    Call PrintDashboard ' ��ú��带 ��Ʈ�� ����Ѵ�
    ' Call PrintProjectHeader ' ������Ʈ�� ��Ʈ�� ����Ѵ�
    ' Call PrintProjectAll ' ������Ʈ ��ü�� ����Ѵ�
End Sub

' ���α׷����� ����� �⺻���� ������������ �����Ѵ�.
' ���� ���� ����
' ���� ���� ��ȿ�� �˻�
' �⺻ ȯ�� ������ ��𿡼� �������°� ���� (���� ���� �Ǵ� ����Ʈ ����)
' ��ưŬ�� �� ==> �⺻ ȯ�溯���� ��ϵ� ��� ���� �Ұ��ΰ�?? �������� �ε��� ���ΰ�?
'
Private Sub Form_Load()
        
    GCurrentPath = App.Path ' ���α׷��� ��� ����
    
    ' data.xlsm ������ ������ ���α׷� ����, run_log.txt ������ ������ ���� �� ��� ����
    Call CheckFiles
    
    gProjectLoadOrCreate = LoadOrCreate.Load
    Option_Load.Value = True
        
        
    ' �ùķ��̼��� �⺻ ȯ�� ������
    GlobalEnv.WeeklyProb = 1.25
    GlobalEnv.Cash_Init = 1000
    GlobalEnv.Hr_Init_H = 13
    GlobalEnv.Hr_Init_L = 6
    GlobalEnv.Hr_Init_M = 21
    GlobalEnv.Hr_LeadTime = 3
    GlobalEnv.Problem = 100
    GlobalEnv.SimulationWeeks = 156 ' 3��(52�� * 3��)
    

    
    ' ȭ�鿡 ���̴� �ʱ� �� ����
    txtSimulationWeeks.Text = GlobalEnv.SimulationWeeks '"156"
    txtWeeklyProb.Text = GlobalEnv.WeeklyProb '"1.25"
    
    Call ModifyExcel ' ����� ������ ��Ʈ ���� ����
    
    Prologue

    'Call LoadExcelEnv ' ����� ������ ��Ʈ ���� ����
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Call CleanUpExcel ' �����ִ� ������ ����. ����� ������ �������� �ʴ´�.
    Call WriteLog(STR_END_EXCEL) ' �α����Ͽ� ���Ḧ ǥ���Ѵ�.
    
End Sub

' �α������� ������ ����, data ������ ������ ��� �� ���α׷� ����
Public Sub CheckFiles()
    
    Dim filePath As String
    filePath = GCurrentPath & "\" & STR_RUN_LOG_FILE
        
    Dim fileNum As Integer
    fileNum = FreeFile
        
    If Dir(filePath) = "" Then  ' �α� ������ �����ϴ��� Ȯ��
        
        Open filePath For Output As #fileNum ' ������ �������� ������ �� ���� ����
        Close #fileNum ' �� ���Ϸ� ����� ���� �ƹ� ���뵵 ���� ����
        
    End If
        
    filePath = GCurrentPath & "\" & STR_DATA_FILE ' Data.xlsm ���� ��� ����
    
    ' ������ ����(��������)�� �����ϴ��� Ȯ��
    If Dir(filePath) = "" Then
        MsgBox "Data.xlsm ������ ������ �ٽ� ������ �ּ���", vbCritical
        End
    End If
        
End Sub



' data.xlsm ������ �̹� ���� �ִ��� Ȯ���ϰ� ���� ������ �ٷ� �ݿ��ǰ� �����Ѵ�.
Public Sub ModifyExcel()
    
    Dim filePath As String

    filePath = GCurrentPath & "\" & STR_DATA_FILE

    ' �̹� ���� ���� Excel ���ø����̼� ��ü ��������
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    If xlApp Is Nothing Then
        ' ���� ���ø����̼� ��ü �ʱ�ȭ
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.ScreenUpdating = True
        'MsgBox "������ ���� ���� �ƴմϴ�. ������ ������ �� �ٽ� �õ��Ͻʽÿ�."
        'End
        'Exit Sub
    End If
    
    ' ��ũ�� ���� �Ǵ� �̹� ���� �ִ� ��ũ�� ����
    On Error Resume Next
    
    Set xlWb = xlApp.Workbooks.Open(filePath)
    
    If Err.Number <> 0 Then ' ��ũ���� �̹� ���� ������
        Err.Clear
        Set xlWb = xlApp.Workbooks(filePath)
        
    End If
    
    On Error GoTo 0
    
    Set gWsDashboard = xlWb.Sheets(DBOARD_SHEET_NAME)
    Set gWsProject = xlWb.Sheets(PROJECT_SHEET_NAME)
    Set gWsActivity_Struct = xlWb.Sheets(ACTIVITY_SHEET_NAME)
    
    ' ��Ʈ�� Clear �ϰ� ������.
    ' Call ClearSheet(gWsProject)
    'xlWb.Save
    
    
    ' ���� ������ �ǽð����� �� �� �ְ� �ϱ� ���� ȭ�� ������Ʈ
    'xlApp.Visible = True
    'xlApp.ScreenUpdating = True

    ' ���� ���� ���� (���� ����)
    ' xlWb.Save

End Sub


' ���� ��Ʈ���� �ʱ�ȭ(clear) �ϰ� �����Ѵ�.
Sub LoadExcelEnv()
    
    Call CheckPreviousRun
    
    ' ���� ���ø����̼� ��ü �ʱ�ȭ
    Set xlApp = CreateObject("Excel.Application")
    
    ' ���� ��ũ�� ����
    Set xlWb = xlApp.Workbooks.Open(GCurrentPath & "\" & STR_DATA_FILE)
        
    Set gWsDashboard = xlWb.Sheets(DBOARD_SHEET_NAME)
    Set gWsProject = xlWb.Sheets(PROJECT_SHEET_NAME)
    Set gWsActivity_Struct = xlWb.Sheets(ACTIVITY_SHEET_NAME)
    
    ' ��Ʈ�� Clear �ϰ� ������.
    Call ClearSheet(gWsProject)
    xlWb.Save

End Sub


Private Sub WriteLog(status As String)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open (GCurrentPath & "\" & STR_RUN_LOG_FILE) For Output As #fileNum
    Print #fileNum, status
    Close #fileNum
    
End Sub
    


' ���α׷� ���� ������ data.xlsm ������ Open /Close ���¸� Ȯ���Ѵ�.
' ���� ������ ���� �־����� �ݰ� �α����Ͽ� "����"�̶�� �ٽ� ����Ѵ�.
Private Sub CheckPreviousRun()

    If ReadLog() <> STR_END_EXCEL Then
        TerminateExcelInstances
    End If
    
    Call WriteLog(STR_START_EXCEL)
    
End Sub

' ���� ������ Open / Close ���¸� �α����Ͽ� �����.
Private Function ReadLog() As String

    On Error Resume Next
    
    Dim status As String
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open (GCurrentPath & "\" & STR_RUN_LOG_FILE) For Input As #fileNum
    Input #fileNum, status
    Close #fileNum
    ReadLog = status
    
End Function

' �̹� ���� �ִ� ��� ������ �ν���Ʈ���� close �Ѵ�.
Private Sub TerminateExcelInstances()

    On Error Resume Next
    Dim objWMI As Object
    Dim objProcess As Object
    Dim objProcesses As Object
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set objProcesses = objWMI.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
    
    For Each objProcess In objProcesses
        objProcess.Terminate
    Next objProcess
    
    Set objProcess = Nothing
    Set objProcesses = Nothing
    Set objWMI = Nothing
    
End Sub


' �����ִ� ������ �����Ѵ�.
' ����� ������ �������� �ʴ´�.
Private Sub CleanUpExcel()

    On Error Resume Next
    
    If Not xlWb Is Nothing Then
        xlWb.Close SaveChanges:=False
        Set xlWb = Nothing
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
        
End Sub

Private Sub Option_Create_Click()
    gProjectLoadOrCreate = LoadOrCreate.Create
    

End Sub

Private Sub Option_Load_Click()
    gProjectLoadOrCreate = LoadOrCreate.Load
End Sub
