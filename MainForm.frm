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
      Caption         =   "프로젝트생성"
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
      Caption         =   "생성옵션"
      Height          =   855
      Left            =   3240
      TabIndex        =   14
      Top             =   5880
      Width           =   3495
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

Enum LoadOrCreate
        Load
        Create
End Enum


Private Sub btnGenBoardNProject_Click()
    
    Dim Res As Integer
    
    ' 입력값들을 업데이트 한다.
    GlobalEnv.WeeklyProb = txtWeeklyProb.Text
    'GlobalEnv.Cash_Init
    'GlobalEnv.Hr_Init_H
    'GlobalEnv.Hr_Init_L
    'GlobalEnv.Hr_Init_M
    'GlobalEnv.Hr_LeadTime
    'GlobalEnv.Problem
    GlobalEnv.SimulationWeeks = txtSimulationWeeks.Text
            
    If gProjectLoadOrCreate = LoadOrCreate.Load Then
        Res = MsgBox("기존의 Data.xlsm 파일의 프로젝트들을 그대로 사용 합니다." & vbNewLine & "계속 진행 할가요?", vbYesNo, "기본 환경 설정")
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click 함수 종료
        Else
            ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
            'gTotalProjectNum = GetLastColumnValue(FindRowWithKeyword("월"))
            Call LoadTablesFromExcel ' 엑셀에 기록된 값들로 테이블을 채운다.
        End If
    Else
        Res = MsgBox("Data.xlsm파일의 내용을 지우고 신규 프로젝트들을 생성 합니다" & vbNewLine & "계속 진행 할까요?", vbYesNo, "기본 환경 설정")
        If (vbNo = Res) Then
            Exit Sub ' btnGenBoardNProject_Click 함수 종료
            
        Else
            ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
            Dim I As Integer
            For I = 1 To GlobalEnv.SimulationWeeks
                gPrintDurationTable(I) = I
            Next I
        
            Call CreateOrderTable ' Order 테이블을 생성
            Call CreateProjects     ' 프로젝트를 생성한다.
            
        End If
        
    End If
        
    Call PrintDashboard ' 대시보드를 시트에 출력한다
    Call PrintProjectHeader ' Project 시트의 헤더를 기록한다.
    Call PrintProjectAll ' 프로젝트 전체를 출력한다

End Sub






' 프로그램에서 사용할 기본적인 전역변수들을 설정한다.
' 파일 존재 여부  / 엑셀 파일 유효성 검사
' 기본 환경 변수를 어디에서 가져오는가 결정 (엑셀 파일 또는 디폴트 값들)
' 버튼클릭 시 ==> 기본 환경변수에 기록된 대로 생성 할것인가?? 엑셀에서 로드할 것인가?
'
Private Sub Form_Load()
        
    GCurrentPath = App.Path ' 프로그램의 경로 저장
    
    ' data.xlsm 파일이 없으면 프로그램 종료, run_log.txt 파일이 없으면 생성 후 계속 진행
    Call CheckFiles
    
    Call ModifyExcel ' 사용할 엑셀과 시트의 Object들을 설정한다.
    
    Call LoadExcelEnv
    
    gProjectLoadOrCreate = LoadOrCreate.Load ' 기본 설정은 엑셀 파일에 기록된 값들을 로드해서 사용
    Option_Load.value = True
        
    ' 시뮬레이션의 기본 환경 변수들
'    GlobalEnv.WeeklyProb = 1.25
'    GlobalEnv.Cash_Init = 1000
'    GlobalEnv.Hr_Init_H = 13
'    GlobalEnv.Hr_Init_L = 6
'    GlobalEnv.Hr_Init_M = 21
'    GlobalEnv.Hr_LeadTime = 3
'    GlobalEnv.Problem = 100
'    GlobalEnv.SimulationWeeks = 156 ' 3년(52주 * 3년)
    
    'GlobalEnv.SimulationWeeks = GetLastColumnValue(FindRowWithKeyword("월"))
    'gTotalProjectNum = GetLastColumnValue(FindRowWithKeyword("누계"))
'    Public gOrderTable() As Variant
'    Public gProjectTable() As clsProject
'    Public gPrintDurationTable() As Variant
    
    
    ' 화면에 보이는 초기 값 설정
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
    
    Call CleanUpExcel ' 열려있는 엑섹을 종료. 변경된 내용은 저장하지 않는다.
    Call WriteLog(STR_END_EXCEL) ' 로그파일에 종료를 표시한다.
    
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



' data.xlsm 파일이 이미 열려 있는지 확인하고 편집 내용이 바로 반영되게 오픈한다.
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
        xlApp.Visible = True
        xlApp.ScreenUpdating = True
        'MsgBox "엑셀이 실행 중이 아닙니다. 엑셀을 실행한 후 다시 시도하십시오."
        'End
        'Exit Sub
    End If
    
    ' 워크북 열기 또는 이미 열려 있는 워크북 참조
    On Error Resume Next
    
    Set xlWb = xlApp.Workbooks.Open(filePath)
    
    If Err.Number <> 0 Then ' 워크북이 이미 열려 있으면
        Err.Clear
        Set xlWb = xlApp.Workbooks(filePath)
        
    End If
    
    On Error GoTo 0
    
    Set gWsParameters = xlWb.Sheets(PARAMETERS_SHEET_NAME)
    Set gWsDashboard = xlWb.Sheets(DBOARD_SHEET_NAME)
    Set gWsProject = xlWb.Sheets(PROJECT_SHEET_NAME)
    Set gWsActivity_Struct = xlWb.Sheets(ACTIVITY_SHEET_NAME)
    
    ' 시트를 Clear 하고 저장함.
    ' Call ClearSheet(gWsProject)
    'xlWb.Save
    
    
    ' 변경 내용을 실시간으로 볼 수 있게 하기 위해 화면 업데이트
    'xlApp.Visible = True
    'xlApp.ScreenUpdating = True

    ' 변경 사항 저장 (선택 사항)
    ' xlWb.Save

End Sub


' 엑셀 시트에서 초기화에 필요한 값들을 가져온다.
Sub LoadExcelEnv()
    
    Dim posY As Long, posX As Long
    
    With gWsParameters
    ' 시뮬레이션의 기본 환경 변수들
    posX = 2: posY = 2: GlobalEnv.SimulationWeeks = .Cells(posY, posX) '156 ' 3년(52주 * 3년)
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
