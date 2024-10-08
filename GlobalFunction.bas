Attribute VB_Name = "GlobalFunction"
Option Explicit
Option Base 1

' Define Global Variables
Public xlApp                As Object   ' 엑셀 애플리케이션 객체
Public xlWb                 As Object   ' 엑셀 워크북 객체

Public g_WsParameters        As Object   ' Parameters 시트 객체
Public g_WsDashboard         As Object   ' Dashboard 시트 객체
Public g_WsProject           As Object   ' Project 시트 객체
Public g_WsActivity_Struct   As Object   ' Activity_Struct 시트 객체
Public g_WsDebugInfo         As Object   ' dbuginfo 시트 객체

Public Const PARAMETERS_SHEET_NAME = "parameters"
Public Const DBOARD_SHEET_NAME = "dashboard"
Public Const PROJECT_SHEET_NAME = "project"
Public Const ACTIVITY_SHEET_NAME = "activity_struct"
Public Const DEBUGINFO_SHEET_NAME = "debuginfo"

Public GCurrentPath As String
Public gProjectLoadOrCreate As Integer ' 프로그램 시작시 프로젝트를 생성할지 기존 데이터를 로드할지 결정하는 변수
Public g_SimulDebug As Boolean  ' 시뮬레이션 결과를 디버깅 정보로 출력할지 말지 결정하는 변수
Public g_ProjDebug As Boolean   ' 프로젝트 생성 결과를 디버깅 정보로 출력할지 말지 결정하는 변수

Public Const P_TYPE_EXTERNAL = 0
Public Const P_TYPE_INTERNAL = 1

Public Const MAX_ACT As Integer = 4
Public Const MAX_N_CF As Integer = 3
Public Const PRJ_SHEET_HEADER_W As Integer = 16
Public Const PRJ_SHEET_HEADER_H As Integer = 7
Public Const RND_HR_H = 20
Public Const RND_HR_M = 70
Public Const MAX_PRJ_TYPE As Integer = 5
Public Const RND_PRJ_TYPE1 As Integer = 20
Public Const RND_PRJ_TYPE2 As Integer = 70
Public Const RND_PRJ_TYPE3 As Integer = 20
Public Const RND_PRJ_TYPE4 As Integer = 70
Public Const RND_PRJ_TYPE5 As Integer = 20
Public Const ORDER_TABLE_INDEX As Long = 1
Public Const DONG_TABLE_INDEX = 21
Public Const PROJECT_TABLE_INDEX As Long = 3


Private gExcelInitialized As Boolean ' 전역변수일 필요 없음
Private gTableInitialized As Boolean ' 전역변수일 필요 없음

Public gTotalProjectNum As Integer
Public GlobalEnv As Environment_
Public gOrderTable() As Variant 'GlobalEnv 에 포함 여부 체크 필요
Public gProjectTable() As clsProject ' 필요한가??
Public gWeekNumberTable() As Integer

Public Type Environment_
    SimulationWeeks As Integer
    Hr_TableSize    As Integer ' maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
    WeeklyProb      As Double
    Hr_Init_H       As Integer
    Hr_Init_M       As Integer
    Hr_Init_L       As Integer
    Hr_LeadTime     As Integer
    Cash_Init       As Integer
    ProblemCnt      As Integer
    status          As Integer  ' 프로그램의 동작 상태. 0:프로젝트 미생성, 1:프로젝트 생성,
End Type

Public Type Activity_
    activityType    As Integer
    duration        As Integer
    startDate       As Integer
    endDate         As Integer
    highSkill       As Integer
    midSkill        As Integer
    lowSkill        As Integer
End Type

Public Property Get GetExcelEnv() As Environment_
    GetExcelEnv = GlobalEnv
End Property

Public Property Get GetExcelInitialized() As Boolean
    GetExcelInitialized = gExcelInitialized
End Property

Public Property Let LetExcelInitialized(value As Boolean)
    gExcelInitialized = value
End Property

Public Property Get GetTableInitialized() As Boolean
    GetTableInitialized = gTableInitialized
End Property

Public Property Let LetTableInitialized(value As Boolean)
    gTableInitialized = value
End Property

Public Property Get GetTotalProjectNum() As Integer
    GetTotalProjectNum = gTotalProjectNum
End Property

Public Property Get GetOrderTable() As Variant
    GetOrderTable = gOrderTable
End Property

Public Property Get GetProjectTable() As Variant
    GetProjectTable = gProjectTable
End Property




' 발생한 프로젝트의 갯수를 테이블에 기록한다.
Public Function CreateOrderTable()
    Dim week As Integer
    Dim projectCount As Integer
    Dim sum As Integer
    
    ReDim gOrderTable(2, GlobalEnv.SimulationWeeks)

    For week = 1 To GlobalEnv.SimulationWeeks
        projectCount = PoissonRandom(GlobalEnv.WeeklyProb)        ' 이번주 발생하는 프로젝트 갯수
        gOrderTable(1, week) = sum
        gOrderTable(2, week) = projectCount
        
        ' 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다. ==> 이전주까지 발생한 프로젝트 갯수후위연산. vba에서 do while 문법 모름... ㅎㅎ
        sum = sum + projectCount
    Next week

    gTotalProjectNum = sum
    
End Function




' 프로젝트를 생성하고
Public Function CreateProjects()

    Dim week As Integer
    Dim Id As Integer
    Dim startPrjNum As Integer
    Dim endPrjNum As Integer
    Dim preTotal As Integer
    Dim tempPrj As clsProject

    If gTotalProjectNum <= 0 Then
        MsgBox "gTotalProjectNum is 0", vbExclamation
        Exit Function
    End If
    
    ' 프로젝트 갯수를 관리하는 테이블에 발생한 프로젝트의 갯수를 기록한다.
    ReDim gProjectTable(gTotalProjectNum)

    MainForm.ProgressBar1.Max = GlobalEnv.SimulationWeeks
    MainForm.ProgressBar1.Min = 0
    MainForm.ProgressBar1.Text = "프로젝트 생성중"
    
    For week = 1 To GlobalEnv.SimulationWeeks
        preTotal = gOrderTable(1, week)
        startPrjNum = preTotal + 1
        endPrjNum = gOrderTable(2, week) + preTotal

        If startPrjNum = 0 Then GoTo Continue
        If startPrjNum > endPrjNum Then GoTo Continue

        ' 이번 주에 발생한 프로젝트들을 생성하고 초기화 한다.
        For Id = startPrjNum To endPrjNum
            Set tempPrj = New clsProject
            Call tempPrj.Init(P_TYPE_EXTERNAL, Id, week)
            Set gProjectTable(Id) = tempPrj
        Next Id

Continue:
        MainForm.ProgressBar1.value = week
    Next week
End Function






Public Function GetVariableValue(rng As Object, variableName As String) As Variant
    Dim dataArray As Variant
    Dim matchIndex As Variant

    dataArray = rng.value
    matchIndex = Application.Match(variableName, Application.index(dataArray, 0, 1), 0)
    
    If Not IsError(matchIndex) Then
        GetVariableValue = dataArray(matchIndex, 2)
    Else
        GetVariableValue = "Variable not found"
    End If
End Function

Sub PrintArrayWithLine(ws As Object, startRow As Long, startCol As Long, dataArray As Variant)
        
    Dim numRows As Long
    Dim numCols As Long
        
    Call GetArraySize(dataArray, numRows, numCols)
    With ws
        .Range(.Cells(startRow, startCol), .Cells(startRow + numRows - 1, startCol + numCols - 1)).value = dataArray
        .Range(.Cells(startRow, startCol), .Cells(startRow + numRows - 1, startCol + numCols - 1)).Borders.LineStyle = xlContinuous
        .Range(.Cells(startRow, startCol), .Cells(startRow + numRows - 1, startCol + numCols - 1)).Borders.Weight = xlThin
        .Range(.Cells(startRow, startCol), .Cells(startRow + numRows - 1, startCol + numCols - 1)).Borders.ColorIndex = xlAutomatic
    End With
    
End Sub



Function GetArraySize(arr As Variant, ByRef rowCount As Long, ByRef colCount As Long)
    On Error GoTo ErrorHandler

    If IsArray(arr) Then
        ' 초기화
        rowCount = 0
        colCount = 0

        ' 1차원 배열인 경우
        On Error Resume Next
        rowCount = UBound(arr, 2)
        If Err.Number <> 0 Then
            ' 1차원 배열
            rowCount = 1
            colCount = UBound(arr, 1) - LBound(arr, 1) + 1
            Err.Clear
        Else
            ' 2차원 배열
            rowCount = UBound(arr, 1) - LBound(arr, 1) + 1
            colCount = UBound(arr, 2) - LBound(arr, 2) + 1
        End If
        On Error GoTo 0
    Else
        rowCount = 0
        colCount = 0
    End If

    Exit Function

ErrorHandler:
    rowCount = 0
    colCount = 0
    
End Function

'Function PrintProjectHeader()
'    Call ClearSheet(gWsProject)
'
'    Dim arrHeader As Variant
'    Dim strHeader As String
'
'    strHeader = "타입,순번,발주일,시작가능,기간,시작,수익,경험,성공%,지급횟수,CF1%,CF2%,CF3%,선금,중도금,잔금"
'    arrHeader = Split(strHeader, ",")
'    arrHeader = ConvertToBase1(arrHeader)
'    arrHeader = ConvertTo1xN(arrHeader)
'    Call PrintArrayWithLine(gWsProject, 2, 1, arrHeader)
'
'    strHeader = ",Dur,start,end,HR_H,HR_M,HR_L,,,,mon_cf1,mon_cf2,mon_cf3"
'    arrHeader = Split(strHeader, ",")
'    arrHeader = ConvertToBase1(arrHeader)
'    arrHeader = ConvertTo1xN(arrHeader)
'    Call PrintArrayWithLine(gWsProject, 3, 1, arrHeader)
'End Function


Function ConvertToBase1(arr As Variant) As Variant
    Dim index As Integer
    Dim newArr() As Variant
    ReDim newArr(1 To UBound(arr) - LBound(arr) + 1)
    For index = LBound(arr) To UBound(arr)
        newArr(index - LBound(arr) + 1) = arr(index)
    Next index
    ConvertToBase1 = newArr
End Function

Function PivotArray(arr As Variant) As Variant
    Dim index As Integer
    Dim rowCount As Integer
    Dim result() As Variant
    
    ' 1차원 배열의 크기 구하기
    rowCount = UBound(arr) - LBound(arr) + 1
    
    ' 2차원 배열 크기 설정 (rowCount x 1)
    ReDim result(1 To rowCount, 1 To 1)
    
    ' 1차원 배열을 2차원 배열로 변환
    For index = LBound(arr) To UBound(arr)
        result(index, 1) = arr(index)
    Next index
    
    PivotArray = result
End Function


Function PrintDashboard()
    Call ClearSheet(g_WsDashboard)

    Dim arrHeader As Variant
    arrHeader = Array("주", "누계", "발주")
    arrHeader = PivotArray(arrHeader)

    Call PrintArrayWithLine(g_WsDashboard, 2, 1, arrHeader)
    Call PrintArrayWithLine(g_WsDashboard, 2, 2, gWeekNumberTable)
    Call PrintArrayWithLine(g_WsDashboard, 3, 2, gOrderTable)
    
    arrHeader = Array("투입", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(g_WsDashboard, 6, 1, arrHeader)
    
    arrHeader = Array("여유", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(g_WsDashboard, 11, 1, arrHeader)
    
    arrHeader = Array("총원", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(g_WsDashboard, 16, 1, arrHeader)
    
End Function

Function ClearSheet(ws As Object)
    With ws
        Dim endRow As Long
        Dim endCol As Long
        endRow = .UsedRange.Rows.Count + .UsedRange.row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).UnMerge
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).Clear
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).ClearContents
    End With
End Function

Public Function PoissonRandom(lambda As Double) As Integer
    Dim L As Double
    Dim p As Double
    Dim k As Integer
    L = Exp(-lambda)
    p = 1
    k = 0
    Do
        k = k + 1
        p = p * Rnd()
    Loop While p > L
    PoissonRandom = k - 1
End Function


Function FindRowWithKeyword(ws As Object, keyword As String) As Long
    Dim lastRow As Long
    Dim index As Long
    
    ' 엑셀 워크시트의 마지막 행 구하기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).row ' xlUp = -4162

    ' 1번 열을 순회하며 키워드 찾기
    For index = 1 To lastRow
        If InStr(1, ws.Cells(index, 1).value, keyword, vbTextCompare) > 0 Then
            FindRowWithKeyword = index
            Exit Function
        End If
    Next index

    ' 키워드를 찾지 못한 경우
    FindRowWithKeyword = 0
End Function

Function GetLastColumnValue(ws As Object, rowNumber As Long) As Variant
    Dim lastCol As Long
    
    ' 특정 행의 마지막 열 번호 구하기
    lastCol = ws.Cells(rowNumber, ws.Columns.Count).End(-4159).Column ' xlToLeft = -4159

    ' 마지막 열의 값 반환
    GetLastColumnValue = ws.Cells(rowNumber, lastCol).value
End Function





