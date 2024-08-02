VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Simulator"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   LinkTopic       =   "MainForm"
   ScaleHeight     =   8400
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows �⺻��
   Begin simulator.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   7080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      Value           =   0
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "�����Ȳ"
   End
   Begin VB.TextBox txtProblemCount 
      Height          =   300
      Left            =   2600
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   5110
      Width           =   1000
   End
   Begin VB.OptionButton Option_Create 
      Caption         =   "Create"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   6120
      Width           =   1215
   End
   Begin VB.OptionButton Option_Load 
      Caption         =   "Load"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtCash 
      Height          =   300
      Left            =   2600
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   4428
      Width           =   1000
   End
   Begin VB.TextBox txtLeadTime 
      Height          =   300
      Left            =   2600
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3750
      Width           =   1000
   End
   Begin VB.TextBox txtHr_L 
      Height          =   300
      Left            =   2600
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3072
      Width           =   1000
   End
   Begin VB.CommandButton btnGenBoardNProject 
      Caption         =   "������Ʈ����"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtHr_M 
      Height          =   300
      Left            =   2600
      TabIndex        =   3
      Top             =   2394
      Width           =   1000
   End
   Begin VB.TextBox txtHr_H 
      Height          =   300
      Left            =   2600
      TabIndex        =   2
      Top             =   1716
      Width           =   1000
   End
   Begin VB.TextBox txtWeeklyProb 
      Height          =   300
      Left            =   2600
      TabIndex        =   1
      Top             =   1038
      Width           =   1000
   End
   Begin VB.TextBox txtSimulationWeeks 
      Height          =   300
      Left            =   2600
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����ɼ�"
      Height          =   855
      Left            =   3240
      TabIndex        =   14
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��� ����
      Caption         =   "��������"
      Height          =   250
      Left            =   400
      TabIndex        =   18
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      Caption         =   "�ڱ�"
      Height          =   250
      Left            =   400
      TabIndex        =   17
      Top             =   4470
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      Caption         =   "LeadTime"
      Height          =   250
      Left            =   400
      TabIndex        =   16
      Top             =   3785
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      Caption         =   "�����η�(�ʱ�)"
      Height          =   250
      Left            =   400
      TabIndex        =   15
      Top             =   3100
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      Caption         =   "�����η�(�߱�)"
      Height          =   250
      Left            =   400
      TabIndex        =   7
      Top             =   2415
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      Caption         =   "�����η�(���)"
      Height          =   250
      Left            =   400
      TabIndex        =   6
      Top             =   1730
      Width           =   1500
   End
   Begin VB.Label ������Ʈ�߻��� 
      Alignment       =   2  '��� ����
      Caption         =   "������Ʈ�߻���"
      Height          =   250
      Left            =   400
      TabIndex        =   5
      Top             =   1045
      Width           =   1500
   End
   Begin VB.Label SimulTearm 
      Alignment       =   2  '��� ����
      Caption         =   "�ð�(��)"
      Height          =   250
      Left            =   400
      TabIndex        =   4
      Top             =   360
      Width           =   1500
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
    
    Dim Res As Integer
    
    ' �Է°����� ������Ʈ �Ѵ�.
    GlobalEnv.WeeklyProb = txtWeeklyProb.Text
    'GlobalEnv.Cash_Init
    'GlobalEnv.Hr_Init_H
    'GlobalEnv.Hr_Init_L
    'GlobalEnv.Hr_Init_M
    'GlobalEnv.Hr_LeadTime
    'GlobalEnv.Problem
    GlobalEnv.SimulationWeeks = txtSimulationWeeks.Text
            
    If gProjectLoadOrCreate = LoadOrCreate.Load Then
        Res = MsgBox("������ Data.xlsm ������ ������Ʈ���� �״�� ��� �մϴ�." & vbNewLine & "��� ���� �Ұ���?", vbYesNo, "�⺻ ȯ�� ����")
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click �Լ� ����
        Else
            ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
            'gTotalProjectNum = GetLastColumnValue(FindRowWithKeyword("��"))
            Call LoadTablesFromExcel ' ������ ��ϵ� ����� ���̺��� ä���.
        End If
    Else
        Res = MsgBox("Data.xlsm������ ������ ����� �ű� ������Ʈ���� ���� �մϴ�" & vbNewLine & "��� ���� �ұ��?", vbYesNo, "�⺻ ȯ�� ����")
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click �Լ� ����
            
        Else
            ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
            Dim I As Integer
            For I = 1 To GlobalEnv.SimulationWeeks
                gPrintDurationTable(I) = I
            Next I
        
            Call CreateOrderTable ' Order ���̺��� ����
            Call CreateProjects     ' ������Ʈ�� �����Ѵ�.
            
        End If
        
    End If
        
    Call PrintDashboard ' ��ú��带 ��Ʈ�� ����Ѵ�
    Call PrintProjectHeader ' Project ��Ʈ�� ����� ����Ѵ�.
    Call PrintProjectAll ' ������Ʈ ��ü�� ����Ѵ�

End Sub






' ���α׷����� ����� �⺻���� ������������ �����Ѵ�.
' ���� ���� ����  / ���� ���� ��ȿ�� �˻�
' �⺻ ȯ�� ������ ��𿡼� �������°� ���� (���� ���� �Ǵ� ����Ʈ ����)
' ��ưŬ�� �� ==> �⺻ ȯ�溯���� ��ϵ� ��� ���� �Ұ��ΰ�?? �������� �ε��� ���ΰ�?
'
Private Sub Form_Load()
        
    GCurrentPath = App.Path ' ���α׷��� ��� ����
    
    ' data.xlsm ������ ������ ���α׷� ����, run_log.txt ������ ������ ���� �� ��� ����
    Call CheckFiles
    
    Call ModifyExcel ' ����� ������ ��Ʈ�� Object���� �����Ѵ�.
    
    Call LoadExcelEnv
    
    gProjectLoadOrCreate = LoadOrCreate.Load ' �⺻ ������ ���� ���Ͽ� ��ϵ� ������ �ε��ؼ� ���
    Option_Load.value = True
        
    ' �ùķ��̼��� �⺻ ȯ�� ������
'    GlobalEnv.WeeklyProb = 1.25
'    GlobalEnv.Cash_Init = 1000
'    GlobalEnv.Hr_Init_H = 13
'    GlobalEnv.Hr_Init_L = 6
'    GlobalEnv.Hr_Init_M = 21
'    GlobalEnv.Hr_LeadTime = 3
'    GlobalEnv.Problem = 100
'    GlobalEnv.SimulationWeeks = 156 ' 3��(52�� * 3��)
    
    'GlobalEnv.SimulationWeeks = GetLastColumnValue(FindRowWithKeyword("��"))
    'gTotalProjectNum = GetLastColumnValue(FindRowWithKeyword("����"))
'    Public gOrderTable() As Variant
'    Public gProjectTable() As clsProject
'    Public gPrintDurationTable() As Variant
    
    
    ' ȭ�鿡 ���̴� �ʱ� �� ����
    txtSimulationWeeks.Text = GlobalEnv.SimulationWeeks '"156"
    txtWeeklyProb.Text = GlobalEnv.WeeklyProb '"1.25"
    txtCash = GlobalEnv.Cash_Init
    txtHr_H = GlobalEnv.Hr_Init_H
    txtHr_M = GlobalEnv.Hr_Init_M
    txtHr_L = GlobalEnv.Hr_Init_L
    txtLeadTime = GlobalEnv.Hr_LeadTime
    txtProblemCount = GlobalEnv.Problem
    
    
    
        
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
        MsgBox "Data.xlsm ������ ������ ���α׷��� �ٽ� ������ �ּ���", vbCritical
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
    
    Set gWsParameters = xlWb.Sheets(PARAMETERS_SHEET_NAME)
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


' ���� ��Ʈ���� �ʱ�ȭ�� �ʿ��� ������ �����´�.
Sub LoadExcelEnv()
    
    Dim posY As Long, posX As Long
    
    With gWsParameters
    ' �ùķ��̼��� �⺻ ȯ�� ������
    posX = 2: posY = 2: GlobalEnv.SimulationWeeks = .Cells(posY, posX) '156 ' 3��(52�� * 3��)
    posY = posY + 1: GlobalEnv.WeeklyProb = .Cells(posY, posX)
    posY = posY + 1: GlobalEnv.Hr_Init_H = .Cells(posY, posX)
    posY = posY + 1: GlobalEnv.Hr_Init_L = .Cells(posY, posX)
    posY = posY + 1: GlobalEnv.Hr_Init_M = .Cells(posY, posX)
    posY = posY + 1: GlobalEnv.Hr_LeadTime = .Cells(posY, posX)
    posY = posY + 1: GlobalEnv.Cash_Init = .Cells(posY, posX)
    posY = posY + 1: GlobalEnv.Problem = .Cells(posY, posX)
    End With
    
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
