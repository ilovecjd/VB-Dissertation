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
   Begin VB.CommandButton ScreenUpdating 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6600
      TabIndex        =   22
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Run 
      Caption         =   "�ùķ��̼ǽ���"
      Height          =   615
      Left            =   4680
      TabIndex        =   21
      Top             =   6120
      Width           =   1455
   End
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
      Top             =   5110
      Width           =   1000
   End
   Begin VB.OptionButton Option_Create 
      Caption         =   "Create"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   6120
      Width           =   975
   End
   Begin VB.OptionButton Option_Load 
      Caption         =   "Load"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtCash 
      Height          =   300
      Left            =   2600
      TabIndex        =   11
      Top             =   4428
      Width           =   1000
   End
   Begin VB.TextBox txtLeadTime 
      Height          =   300
      Left            =   2600
      TabIndex        =   10
      Top             =   3750
      Width           =   1000
   End
   Begin VB.TextBox txtHr_L 
      Height          =   300
      Left            =   2600
      TabIndex        =   9
      Top             =   3072
      Width           =   1000
   End
   Begin VB.CommandButton btnGenBoardNProject 
      Caption         =   "������Ʈ����"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
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
      Left            =   2040
      TabIndex        =   14
      Top             =   5880
      Width           =   2295
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


' ������Ʈ���� ���� �Ѵ�.
' 1. ���� ������Ʈ���� �״�� ���
'   1.1 ���� data.xlsm ���Ͽ��� �ε�
'
' 2. ������Ʈ�� ���Ӱ� ����
'   2.1 ȯ�溯�� ������Ʈ
'   2.2 ���ο� ������Ʈ�� ����
'   2.2 data.xlsm ������ ��Ʈ�� ������Ʈ
Private Sub btnGenBoardNProject_Click()
    
    Dim Res As Integer
    Dim i   As Integer ' song �빮�ڷ� �ڵ� ���� �Ǵµ�...������ �𸣰���. ���� ���� ���̶�...

    ' �Է°����� ������Ʈ �Ѵ�.
    GlobalEnv.SimulationWeeks = txtSimulationWeeks.Text
    GlobalEnv.WeeklyProb = txtWeeklyProb.Text
    GlobalEnv.Hr_Init_H = txtHr_H.Text
    GlobalEnv.Hr_Init_L = txtHr_M.Text
    GlobalEnv.Hr_Init_M = txtHr_L.Text
    GlobalEnv.Hr_LeadTime = txtLeadTime.Text
    GlobalEnv.Cash_Init = txtCash.Text
    GlobalEnv.ProblemCnt = txtProblemCount.Text

    '1. ���� ������Ʈ���� �״�� ���
    If gProjectLoadOrCreate = LoadOrCreate.Load Then
        Res = MsgBox("������ Data.xlsm ������ ������Ʈ���� �״�� ��� �մϴ�." & vbNewLine & "��� ���� �Ұ���?", vbYesNo, "�⺻ ȯ�� ����")
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click �Լ� ����
        Else
            ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
            'gTotalProjectNum = GetLastColumnValue(FindRowWithKeyword("��"))
            
            ' data.xlsm ���Ͽ��� order ���̺�� project ���̺��� �о���δ�.
            ' song ���Ͼ��� ������ ��ȿ�� ������ �� ����.
            Call LoadTablesFromExcel ' ������ ��ϵ� ����� ���̺��� ä���.
        End If
        
    '2. ������Ʈ�� ���Ӱ� ����
    Else
        Res = MsgBox("Data.xlsm������ ������ ����� �ű� ������Ʈ���� ���� �մϴ�" & vbNewLine & "��� ���� �ұ��?", vbYesNo, "�⺻ ȯ�� ����")
        
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click �Լ� ����
            
        Else
            
            ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
            
            For i = 1 To GlobalEnv.SimulationWeeks
                gPrintDurationTable(i) = i
            Next i
        
            Call CreateOrderTable   ' Order ���̺��� �����ϰ� '��'�� �Է��Ѵ�.
            Call CreateProjects     ' ������Ʈ�� �����Ѵ�.
            Call PrintDashboard     ' Order ���̺�� �η������� ��ú��� ��Ʈ�� ����Ѵ�.
            Call PrintProjectHeader ' Project ��Ʈ�� ����� ����Ѵ�.
            Call PrintProjectAll    ' ������Ʈ ��ü�� ����Ѵ�
            
        End If
        
    End If
        
    

End Sub






' ���α׷����� ����� �⺻���� ������������ �����Ѵ�.
' ���� ���� ����  / ���� ���� ��ȿ�� �˻�
' �⺻ ȯ�� ������ ��𿡼� �������°� ���� (���� ���� �Ǵ� ����Ʈ ����)
' ��ưŬ�� �� ==> �⺻ ȯ�溯���� ��ϵ� ��� ���� �Ұ��ΰ�?? �������� �ε��� ���ΰ�?
'
Private Sub Form_Load()
        
    GCurrentPath = App.Path ' ���α׷� �� data.xlsm ������ ���
    
    'run_log.txt ������ ������ ���� �� ��� ����, data.xlsm ������ ������ ��� �� ���α׷� ����,
    Call CheckFiles
    
    ' data.xlsm ������ ��Ʈ ��ü���� ����
    Call ModifyExcel
    
    ' data.xlsm ������ parameters�� Dashboard ��Ʈ�� ����鿡 ���� ��ȿ�� üũ
    Call CheckDataFile
    
    ' data.xlsm ���Ͽ��� �ùķ��̼��� �⺻���������� �����´�.
    Call LoadEnvFromExcel
        
    ' ������Ʈ���� data.xlsm ���Ͽ� ��ϵ� ������ ����ϴ°��� ����Ʈ ����
    gProjectLoadOrCreate = LoadOrCreate.Load
    Option_Load.value = True

    ' ȭ�鿡 ���̴� �ʱ� �� ����
    txtSimulationWeeks.Text = GlobalEnv.SimulationWeeks '156 = 3��(52�� * 3��)
    txtWeeklyProb.Text = GlobalEnv.WeeklyProb           '1.25
    txtCash = GlobalEnv.Cash_Init                       '1000
    txtHr_H = GlobalEnv.Hr_Init_H                       '13
    txtHr_M = GlobalEnv.Hr_Init_M                       '21
    txtHr_L = GlobalEnv.Hr_Init_L                       '6
    txtLeadTime = GlobalEnv.Hr_LeadTime                 '3
    txtProblemCount = GlobalEnv.ProblemCnt              '100
        
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



' data.xlsm ������ ��Ʈ ��ü���� ����
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
    End If
    
    ' ��ũ�� ���� �Ǵ� �̹� ���� �ִ� ��ũ�� ����
    On Error Resume Next
    
    Set xlWb = xlApp.Workbooks.Open(filePath)
    
    If Err.Number <> 0 Then ' ��ũ���� �̹� ���� ������
        Err.Clear
        Set xlWb = xlApp.Workbooks(filePath)
        ' song ������ �ν���Ʈ�� ���� �ִ� ��쿡 ���� ���� ó�� �ʿ�.
    End If
    
    On Error GoTo 0 '���� ��ü�� ����� ���� �ʱⰪ���� ����
    
    ' song ������ �ν���Ʈ�� ���� �ִ� ��쿡 ���� ���� ó�� �ʿ�.
    Set gWsParameters = xlWb.Sheets(PARAMETERS_SHEET_NAME)
    Set gWsDashboard = xlWb.Sheets(DBOARD_SHEET_NAME)
    Set gWsProject = xlWb.Sheets(PROJECT_SHEET_NAME)
    Set gWsActivity_Struct = xlWb.Sheets(ACTIVITY_SHEET_NAME)
    
End Sub


' ���� ��Ʈ���� �ʱ�ȭ�� �ʿ��� ������ �����´�.
Sub LoadEnvFromExcel()
    
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
    posY = posY + 1: GlobalEnv.ProblemCnt = .Cells(posY, posX)
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

' �ùķ��̼��� �����Ѵ�.
Private Sub Run_Click()

    Dim i As Integer
    'song �ùķ��̼��� �غ�Ǿ����� üũ�ؾ���.
    
    Dim Company As clsCompany
    
    Set Company = New clsCompany
    Call Company.Init(1)    ' �ʱ�ȭ.ȸ�� ID(���� ���ǿ��� ���� ȸ�縦 �), ������Ʈ ����

    Debug.Print VBA.String(200, vbNewLine)
    
    For i = 1 To GlobalEnv.SimulationWeeks
        Call Company.Decision(i)    ' i��° �Ⱓ�� �����ؾ� �� �ϵ�
        'Call ClearTableArea(gWsDashboard, DONG_TABLE_INDEX)
        'Call PrintSimulationResults(Company)
    Next
    
    Call ClearTableArea(gWsDashboard, DONG_TABLE_INDEX)
    Call PrintSimulationResults(Company)
    
End Sub

Function ClearTableArea(ws As Worksheet, startRow As Long)
    
    With ws
        Dim endRow As Long ' ��������
        Dim endCol As Long ' ��������
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' ���� ������ ������ �����Ѵ�.
        .Range(.Cells(startRow, 1), .Cells(endRow, endCol)).UnMerge
        .Range(.Cells(startRow, 1), .Cells(endRow, endCol)).Clear
        .Range(.Cells(startRow, 1), .Cells(endRow, endCol)).ClearContents
    End With

End Function


Private Function PrintSimulationResults(Company As clsCompany)
    
    'Call ClearSheet(gWsDashboard)          '��Ʈ�� ��� ������ ����� �� ���� ����

    Dim startRow    As Long
    Dim arrHeader   As Variant
    Dim tempArray() As Integer
        
    arrHeader = Array("��", "����", "prjNum")
    arrHeader = PivotArray(arrHeader)
    
    startRow = DONG_TABLE_INDEX
    tempArray = Company.PropertyDoingTable
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' �����׸��� ����
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '�Ⱓ�� ����
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, tempArray)      ' ������ ���´�.

    startRow = startRow + Company.comDoingTableSize + 2
    tempArray = Company.PropertyDoneTable
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' �����׸��� ����
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '�Ⱓ�� ����
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, tempArray)       ' ������ ���´�.

    startRow = startRow + Company.comDoneTableSize + 2
    tempArray = Company.PropertyDefferTable
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' �����׸��� ����
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '�Ⱓ�� ����
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, tempArray)     ' ������ ���´�.


    Exit Function

    
End Function

' data.xlsm ������ parameters, dashboard ��Ʈ�� ��ȿ�� üũ
Private Function CheckDataFile() As Boolean
        
    Dim arrHeader As Variant
    Dim posY As Long, posX As Long, i As Integer
    Dim strErr As String
    
    CheckDataFile = True
    strErr = "������ Ȯ���ϼ���."
        
    With gWsParameters
    
        strErr = strErr & vbNewLine & PARAMETERS_SHEET_NAME & ": "
        
        arrHeader = Array("SimulTerm", "avgProjects", "Hr_Init_H", "Hr_Init_M", "Hr_Init_L", "Hr_LeadTime", "Cash_Init", "ProblemCnt")
        arrHeader = PivotArray(arrHeader)
                
        posX = 1: posY = 2
        
        For i = LBound(arrHeader) To UBound(arrHeader)
            If arrHeader(i, 1) = .Cells(posY, posX) Then
            
            Else
                strErr = strErr & arrHeader(i, 1) & ", "
                CheckDataFile = False
            End If
            
            posY = posY + 1
            
        Next i
        
    End With
        
        
    With gWsDashboard
    
        strErr = strErr & vbNewLine & DBOARD_SHEET_NAME & ": "
        
        arrHeader = Array("��", "����", "����")
        arrHeader = PivotArray(arrHeader)
                
        posX = 1: posY = 2
        
        For i = LBound(arrHeader) To UBound(arrHeader)
            If arrHeader(i, 1) = .Cells(posY, posX) Then
            
            Else
                strErr = strErr & arrHeader(i, 1) & ", "
                CheckDataFile = False
            End If
            
            posY = posY + 1
            
        Next i
        
    End With
    
    ' song PROJECT_SHEET_NAME �� üũ�� �ʿ� ���ٰ� ������
    ' song ACTIVITY_SHEET_NAME �� üũ�� ���� ����
        
    If CheckDataFile = False Then
        Call MsgBox(strErr, vbCritical, "�߿�")
    End If
    
End Function

Private Sub ScreenUpdating_Click()
xlApp.ScreenUpdating = True
End Sub
