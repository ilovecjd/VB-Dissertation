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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox Chk_ProjectDebug 
      Caption         =   "Proj Debug"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CheckBox Chk_SimiDebug 
      Caption         =   "Simul Debug"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton ScreenUpdating 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8160
      TabIndex        =   22
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Run 
      Caption         =   "시뮬레이션시작"
      Height          =   615
      Left            =   4680
      TabIndex        =   21
      Top             =   6120
      Width           =   1455
   End
   Begin simulator.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   7560
      Width           =   9615
      _ExtentX        =   11668
      _ExtentY        =   661
      Value           =   0
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "진행상황"
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
      Caption         =   "프로젝트생성"
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
      Caption         =   "생성옵션"
      Height          =   855
      Left            =   2040
      TabIndex        =   14
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "문제갯수"
      Height          =   250
      Left            =   400
      TabIndex        =   18
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "자금"
      Height          =   250
      Left            =   400
      TabIndex        =   17
      Top             =   4470
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "LeadTime"
      Height          =   250
      Left            =   400
      TabIndex        =   16
      Top             =   3785
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "보유인력(초급)"
      Height          =   250
      Left            =   400
      TabIndex        =   15
      Top             =   3100
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "보유인력(중급)"
      Height          =   250
      Left            =   400
      TabIndex        =   7
      Top             =   2415
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "보유인력(고급)"
      Height          =   250
      Left            =   400
      TabIndex        =   6
      Top             =   1730
      Width           =   1500
   End
   Begin VB.Label 프로젝트발생빈 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "프로젝트발생빈도"
      Height          =   250
      Left            =   400
      TabIndex        =   5
      Top             =   1045
      Width           =   1500
   End
   Begin VB.Label SimulTearm 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "시간(주)"
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
Const STR_START_EXCEL = "시작"
Const STR_END_EXCEL = "종료"
Const NUN_OF_COMPANY As Integer = 1 ' 시뮬레이션 할 회사 수. 멀티 쓰레드를 고려해서 배열로

' data.xlsm 파일에서 Project를 로드 할지, 새로 생성할지 표시하는 플래그값
Enum LoadOrCreate
        Load
        Create
End Enum

Private m_Companies(NUN_OF_COMPANY) As clsCompany




Private Sub CheckDebug_Click()

End Sub

Private Sub Chk_ProjectDebug_Click()

    If Chk_ProjectDebug.value = 1 Then
        g_ProjDebug = True
    Else
        g_ProjDebug = False
    End If
End Sub

Private Sub Chk_SimiDebug_Click()

    If Chk_SimiDebug.value = 1 Then
        g_SimulDebug = True
    Else
        g_SimulDebug = False
    End If
    
End Sub

' 프로그램의 기본적인 환경변수들을 설정한다.
' data.xlsm 파일 존재 여부와 유효성 검사
' 기본 환경 변수를 어디에서 가져오는가 결정 (엑셀 파일 또는 디폴트 값들)
' 프로젝트는 기본 환경변수에 기록된 대로 새로 생성 또는 data.xlsm파일 에서 로드
Private Sub Form_Load()
        
    GCurrentPath = App.Path ' 프로그램 및 data.xlsm 파일의 경로
        
    g_SimulDebug = True ' 시뮬레이션은 디버깅 모드 on 이 디폴트
    g_ProjDebug = False ' 프로젝트 생성은 디버깅 off 가 디폴트
    
    gProjectLoadOrCreate = LoadOrCreate.Load
    
    
    'run_log.txt 파일이 없으면 생성 후 계속 진행, data.xlsm 파일이 없으면 경고 후 프로그램 종료,
    Call CheckFiles
    
    ' data.xlsm 파일을 열고 WorkSheet Object 들을 설정
    Call ModifyExcel
    
    ' data.xlsm 파일의 parameters와 Dashboard 시트의 내용들에 대한 유효성 체크
    If Not CheckDataFile Then
        Exit Sub
    End If
    
    ' data.xlsm 파일에서 시뮬레이션의 기본설정값들을 가져온다.
    Call LoadEnvFromExcel
    Call SetUserParameters
    
    ' company 들을 생성한다.
    Dim index As Integer
    For index = 1 To NUN_OF_COMPANY
        Set m_Companies(index) = New clsCompany
        m_Companies(index).Init (index)
    Next
    
    ' 컨드롤들의 상태 설정
    ProgressBar1.Max = GlobalEnv.SimulationWeeks
    ProgressBar1.Min = 0
    ProgressBar1.Text = "프로젝트를 생성하세요"
    
    
    If g_SimulDebug Then
        Chk_SimiDebug.value = 1
    Else
        Chk_SimiDebug.value = 0
    End If
    
    
    If g_ProjDebug Then
        Chk_ProjectDebug.value = 1
    Else
        Chk_ProjectDebug.value = 0
    End If
    
    
    If gProjectLoadOrCreate = LoadOrCreate.Load Then
        Option_Load.value = True
    Else
        Option_Create.value = True
    End If
        
        
    Run.Enabled = False
       
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Dim retCode As Integer
    
    retCode = MsgBox("data.xlsm 을 저장하고 종료하려면 Yes," + vbNewLine _
            + "저장없이 종료하려면 No" + vbNewLine _
            + "계속 실행하려면 Cancel 을 선택 하시오", vbYesNoCancel, "프로그램 종료")
            
    Select Case retCode
    
    Case vbYes, vbNo
    
    Case vbCancel
        Cancel = 1
        Exit Sub ' 프로그램을 종료하지 않는다.
    
    Case Else
        Call MsgBox("알수 없는 상태로 종료합니다." + vbNewLine + _
                    " data.xlsm 파일은 저장되지 않습니다." + vbNewLine + _
                    "문제를 확인하세요", vbCritical, "알수 없는 상태로 종료")
    End Select
        
    Call CleanUpExcel(retCode) ' 열려있는 엑셀을 종료. 변경된 내용은 저장하지 않는다.
    Call WriteLog(STR_END_EXCEL) ' 로그파일에 종료를 표시한다.
    
End Sub

' 입력값들을 업데이트 한다.
' maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
Private Function GetUsetParameters()
    GlobalEnv.SimulationWeeks = txtSimulationWeeks.Text
    GlobalEnv.Hr_TableSize = txtSimulationWeeks.Text + 80
    GlobalEnv.WeeklyProb = txtWeeklyProb.Text
    GlobalEnv.Hr_Init_H = txtHr_H.Text
    GlobalEnv.Hr_Init_L = txtHr_M.Text
    GlobalEnv.Hr_Init_M = txtHr_L.Text
    GlobalEnv.Hr_LeadTime = txtLeadTime.Text
    GlobalEnv.Cash_Init = txtCash.Text
    GlobalEnv.ProblemCnt = txtProblemCount.Text
End Function

' 화면에 보여줄 초기 값 설정
Private Function SetUserParameters()
    txtSimulationWeeks.Text = GlobalEnv.SimulationWeeks '156 = 3년(52주 * 3년)
    txtWeeklyProb.Text = GlobalEnv.WeeklyProb           '1.25
    txtCash = GlobalEnv.Cash_Init                       '1000
    txtHr_H = GlobalEnv.Hr_Init_H                       '13
    txtHr_M = GlobalEnv.Hr_Init_M                       '21
    txtHr_L = GlobalEnv.Hr_Init_L                       '6
    txtLeadTime = GlobalEnv.Hr_LeadTime                 '3
    txtProblemCount = GlobalEnv.ProblemCnt              '100
End Function

' 프로젝트들을 생성 한다.
' 1. 기존 프로젝트들을 그대로 사용
'   1.1 기존 data.xlsm 파일에서 로드
'
' 2. 프로젝트를 새롭게 생성
'   2.1 환경변수 업데이트
'   2.2 새로운 프로젝트들 생성
'   2.2 data.xlsm 파일의 시트들 업데이트
Private Sub btnGenBoardNProject_Click()
    
    Dim Res As Integer
    Dim index   As Integer
    
    Call GetUsetParameters
    
    ReDim gWeekNumberTable(1 To GlobalEnv.SimulationWeeks)
            
    For index = 1 To GlobalEnv.SimulationWeeks
        gWeekNumberTable(index) = index
    Next index

    '1. data.xlsm 파일에 있는 기존 프로젝트들을 그대로 사용 ' song 추가적인 데이터 유효성 검증필요?
    If gProjectLoadOrCreate = LoadOrCreate.Load Then
    
        Res = MsgBox("기존의 Data.xlsm 파일의 프로젝트들을 그대로 사용 합니다." & vbNewLine & "계속 진행 할가요?", vbYesNo, "기본 환경 설정")
        If (vbNo = Res) Then
            Exit Sub
        Else
            For index = 1 To NUN_OF_COMPANY
                Call m_Companies(index).LoadProjectsFromExcel   ' data.xlsm 파일에서 order 테이블과 project 테이블을 읽어들인다.
            Next
        End If 'If (vbNo = Res) Then
        
    '2. 화면에 입력된 값으로 프로젝트를 새롭게 생성
    Else
        Res = MsgBox("Data.xlsm파일의 내용을 지우고 신규 프로젝트들을 생성 합니다" & vbNewLine & "계속 진행 할까요?", vbYesNo, "기본 환경 설정")
        
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click 함수 종료
            
        Else
            For index = 1 To NUN_OF_COMPANY
                Call m_Companies(index).CreateProjects      ' 프로젝트 생성
                Call m_Companies(index).PrintCreateProjectResults(m_Companies(index).m_totalProjectNum)      ' 프로젝트 전체를 출력한다
                'Call m_Companies(index).PrintProjectHeader  ' Project 시트의 헤더를 기록한다.
                'Call m_Companies(index).PrintProjectAll     ' 프로젝트 전체를 출력한다
                'Call m_Companies(index).PrintProjectAll     ' Order 테이블과 인력정보를 대시보드 시트에 출력한다.
            Next

        End If  ' If (vbNo = Res) Then
        
    End If  ' If gProjectLoadOrCreate = LoadOrCreate.Load Then
    
    ' song Call
    
    Run.Enabled = True
    

End Sub


' 로그파일이 없으면 생성, data 파일이 없으면 경고 후 프로그램 종료
Public Sub CheckFiles()
    
    Dim filePath As String
    filePath = GCurrentPath & "\" & STR_RUN_LOG_FILE
        
    Dim fileNum As Integer
    fileNum = FreeFile
        
    If Dir(filePath) = "" Then  ' 로그 파일이 존재하는지 확인
        
        Open filePath For Output As #fileNum ' 파일이 존재하지 않으면 새 파일 생성
        Close #fileNum ' 빈 파일로 만들기 위해 아무 내용도 쓰지 않음
        
    End If
        
    filePath = GCurrentPath & "\" & STR_DATA_FILE ' Data.xlsm 파일 경로 설정
    
    ' 데이터 파일(엑셀파일)이 존재하는지 확인
    If Dir(filePath) = "" Then
        MsgBox "Data.xlsm 파일을 복사후 프로그램을 다시 시작해 주세요", vbCritical
        End
    End If
        
End Sub



' data.xlsm 파일의 시트 객체들을 설정
' song 다른 엑셀이 열려 있었으면 xlApp 는 닫지 않는 것으로 코드 수정 필요
Public Sub ModifyExcel()
    
    Dim filePath As String

    filePath = GCurrentPath & "\" & STR_DATA_FILE

    ' 이미 실행 중인 Excel 애플리케이션 객체 가져오기
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    If xlApp Is Nothing Then
        ' 엑셀 애플리케이션 객체 초기화
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    xlApp.Visible = True
    xlApp.ScreenUpdating = True
    
    ' 워크북 열기 또는 이미 열려 있는 워크북 참조
    On Error Resume Next
    
    Set xlWb = xlApp.Workbooks.Open(filePath)
    
    If Err.Number <> 0 Then ' 워크북이 이미 열려 있으면
        Err.Clear
        Set xlWb = xlApp.Workbooks(filePath)
        ' song 엑셀의 인스턴트만 남아 있는 경우에 대한 예외 처리 필요.
    End If
    
    On Error GoTo 0 '오류 객체에 저장된 값을 초기값으로 변경
    
    ' song 엑셀의 인스턴트만 남아 있는 경우에 대한 예외 처리 필요.
    Set g_WsParameters = xlWb.Sheets(PARAMETERS_SHEET_NAME)
    Set g_WsDashboard = xlWb.Sheets(DBOARD_SHEET_NAME)
    Set g_WsProject = xlWb.Sheets(PROJECT_SHEET_NAME)
    Set g_WsActivity_Struct = xlWb.Sheets(ACTIVITY_SHEET_NAME)
    Set g_WsDebugInfo = xlWb.Sheets(DEBUGINFO_SHEET_NAME)
    
End Sub


' 엑셀 시트에서 초기화에 필요한 값들을 가져온다.
Sub LoadEnvFromExcel()
    
    Dim posY As Long, posX As Long
    
    With g_WsParameters
    
    ' 시뮬레이션의 기본 환경 변수들
    posX = 2: posY = 2: GlobalEnv.SimulationWeeks = .Cells(posY, posX) '156 ' 3년(52주 * 3년)
    
    ' maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
    GlobalEnv.Hr_TableSize = GlobalEnv.SimulationWeeks + 80
    
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


' 프로그램 실행 이전의 data.xlsm 파일의 Open /Close 상태를 확인한다.
' 엑셀 파일이 열려 있었으면 닫고 로그파일에 "시작"이라고 다시 기록한다.
Private Sub CheckPreviousRun()

    If ReadLog() <> STR_END_EXCEL Then
        TerminateExcelInstances
    End If
    
    Call WriteLog(STR_START_EXCEL)
    
End Sub

' 엑셀 파일이 Open / Close 상태를 로그파일에 기록함.
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

' 이미 열려 있는 모든 엑셀의 인스턴트들을 close 한다.
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


' 열려있는 엑섹을 종료한다.
' 변경된 내용은 저장하지 않는다.
Private Sub CleanUpExcel(retCode As Integer)

    On Error Resume Next
    
    If Not xlWb Is Nothing Then
    
        If retCode = vbYes Then
            xlWb.Close SaveChanges:=True
            Set xlWb = Nothing
            
        Else
            xlWb.Close SaveChanges:=False
            Set xlWb = Nothing
        End If
        
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


' 시뮬레이션을 시작한다.
Private Sub Run_Click()

    Dim index As Integer
    'song 시뮬레이션이 준비되었는지 체크해야함.
    
    Dim Company As clsCompany
    
    Set Company = m_Companies(1)
    'song 생성시 초기화 하자.Call Company.Init(1)    ' 초기화.회사 ID(같은 조건에서 여러 회사를 운영), 프로젝트 갯수
    
    For index = 1 To GlobalEnv.SimulationWeeks
        Call Company.Decision(index)    ' i번째 기간에 결정해야 할 일들
        Call Company.DebugDashboard(index)
    Next
       
    Call Company.PrintSimulationResults(GlobalEnv.SimulationWeeks)
    
End Sub

Function ClearTableArea(ws As Worksheet, startRow As Long)
    
    With ws
        Dim endRow As Long ' 마지막행
        Dim endCol As Long ' 마지막열
        endRow = .UsedRange.Rows.Count + .UsedRange.row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' 엑셀 파일의 셀들을 정리한다.
        .Range(.Cells(startRow, 1), .Cells(endRow, endCol)).UnMerge
        .Range(.Cells(startRow, 1), .Cells(endRow, endCol)).Clear
        .Range(.Cells(startRow, 1), .Cells(endRow, endCol)).ClearContents
    End With

End Function


' data.xlsm 파일의 parameters, dashboard 시트의 유효성 체크
Private Function CheckDataFile() As Boolean
        
    Dim arrHeader As Variant
    Dim posY As Long, posX As Long, index As Integer
    Dim strErr As String
    
    CheckDataFile = True
    strErr = "다음을 확인하세요."
        
    With g_WsParameters
    
        strErr = strErr & vbNewLine & PARAMETERS_SHEET_NAME & ": "
        
        arrHeader = Array("SimulTerm", "avgProjects", "Hr_Init_H", "Hr_Init_M", "Hr_Init_L", "Hr_LeadTime", "Cash_Init", "ProblemCnt")
        arrHeader = PivotArray(arrHeader)
                
        posX = 1: posY = 2
        
        For index = LBound(arrHeader) To UBound(arrHeader)
            If arrHeader(index, 1) = .Cells(posY, posX) Then
            
            Else
                strErr = strErr & arrHeader(index, 1) & ", "
                CheckDataFile = False
            End If
            
            posY = posY + 1
            
        Next index
        
    End With
        
        
    With g_WsDashboard
    
        strErr = strErr & vbNewLine & DBOARD_SHEET_NAME & ": "
        
        arrHeader = Array("주", "누계", "발주")
        arrHeader = PivotArray(arrHeader)
                
        posX = 1: posY = 2
        
        For index = LBound(arrHeader) To UBound(arrHeader)
            If arrHeader(index, 1) = .Cells(posY, posX) Then
            
            Else
                strErr = strErr & arrHeader(index, 1) & ", "
                CheckDataFile = False
            End If
            
            posY = posY + 1
            
        Next index
        
    End With
    
    ' song PROJECT_SHEET_NAME 은 체크가 필요 없다고 생각됨
    ' song ACTIVITY_SHEET_NAME 의 체크는 추후 진행
        
    If CheckDataFile = False Then
        Call MsgBox(strErr, vbCritical, "중요")
    End If
    
End Function

Private Sub ScreenUpdating_Click()
xlApp.ScreenUpdating = True
End Sub

