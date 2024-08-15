VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_orderTable()      As Variant  ' 주별 발주 프로젝트의 누계와 주차별 갯수, week 갯수 가변
Private m_projectTable()    As Variant  ' 발주된 모든 프로젝트 object 를 보관, 가변

Public m_totalProjectNum   As Integer  ' 발주된 프로젝트의 총 갯수



Private doingTable()     As Integer  ' 기간별로 진행중인 프로젝트들의 ID를 관리하는 테이블
Private doneTable()      As Integer  ' 기간별로 완료된 프로젝트들의 ID를 관리하는 테이블
Private defferTable()    As Integer  ' 기간별로 보류된 프로젝트들의 ID를 관리하는 테이블

Private candidateTable() As Integer   ' 후보　프로젝트들의 ID를 관리하는 테이블

Public doingTableSize    As Integer      ' DoneTable의 현재크기 song ==> 동적으로 수정하자.
Public doneTableSize     As Integer      ' DoneTable의 현재크기 song ==> 동적으로 수정하자.
Public defferTableSize   As Integer      ' DoneTable의 현재크기 song ==> 동적으로 수정하자.

Private doingHR_H_Table()  As Integer  ' 기간별로 진행중인 프로젝트에 투입된 상급자
Private doingHR_M_Table()  As Integer  ' 기간별로 진행중인 프로젝트에 투입된 중급자
Private doingHR_L_Table()  As Integer  ' 기간별로 진행중인 프로젝트에 투입된 초급자

Private totalHR_H_Table()  As Integer  ' 기간별로 회사가 보유한 상급자
Private totalHR_M_Table()  As Integer  ' 기간별로 회사가 보유한 중급자
Private totalHR_L_Table()  As Integer  ' 기간별로 회사가 보유한 초급자
 
Public companyID As Integer ' 회사 이름

Private Const MAX_CANDIDATES As Integer = 50 ' SONG 임의로 지정. 동적으로 변하게 할 필요까지는 없어 보임.


Public debugArrayWidth As Long
Public debugArrayHeight As Long

Private m_AllDashboardInfo() As Variant
Private m_AllProjectInfo() As Variant


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Property 들
Public Property Get PropertyDoingTable() As Integer()
    PropertyDoingTable = doingTable
End Property

Public Property Let PropertyDoingTable(value() As Integer)
    doingTable = value
End Property

Public Property Get PropertyDoneTable() As Integer()
    PropertyDoneTable = doneTable
End Property

Public Property Let PropertyDoneTable(value() As Integer)
    doneTable = value
End Property

Public Property Get PropertyDefferTable() As Integer()
    PropertyDefferTable = defferTable
End Property

Public Property Let PropertyDefferTable(value() As Integer)
    defferTable = value
End Property

' Property 들
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 생성자와 소멸자
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    ' 종료할때 버퍼의 크기를 확인해서 가까운 값으로 초기값을 결정하기 위한 디버깅 메세지
    Debug.Print "Class is being terminated"; doingTableSize, doneTableSize, defferTableSize
End Sub



' Company의 기본 값들을 설정한다.
' Company ID와 발생할 프로젝트 테이블을 저장한다.
' 프로젝트, 인력, 자금을 관리할 테이블 생성
Public Sub Init(Id As Integer)
    
    companyID = Id
    
    doingTableSize = 10
    doneTableSize = 10
    defferTableSize = 10
    
    ' song debugInfo 배열의 크기는 고민해 보자.
    debugArrayHeight = 70
    debugArrayWidth = 200
    

    ' 지속적으로 사용하는 테이블들을 정해진 크기로 할당한다.
    ' 시뮬레이션 이후로도 회사는 계속 운영되므로 테이블크기는 시뮬레이션 기간보다 커야한다.
    ' 시뮬레이션 기간 + 프로젝트가 가질수 있는 최대 기간 만큼의 크기 이어야 한다.
    ReDim doingTable(1 To doingTableSize, 1 To GlobalEnv.SimulationWeeks)     ' 진행중인 프로젝트들의 ID를 관리하는 테이블
    ReDim doneTable(1 To doneTableSize, 1 To GlobalEnv.SimulationWeeks)       ' 완료된 프로젝트들의 ID를 관리하는 테이블
    ReDim defferTable(1 To defferTableSize, 1 To GlobalEnv.SimulationWeeks)   ' 보류된 프로젝트들의 ID를 관리하는 테이블
    
    ReDim candidateTable(1 To MAX_CANDIDATES) ' 후보 프로젝트들의 ID를 관리하는 테이블
    
    ReDim doingHR_H_Table(1 To GlobalEnv.Hr_TableSize)  ' 프로젝트에 투입된 상급자
    ReDim doingHR_M_Table(1 To GlobalEnv.Hr_TableSize)  ' 프로젝트에 투입된 중급자
    ReDim doingHR_L_Table(1 To GlobalEnv.Hr_TableSize)  ' 프로젝트에 투입된 초급자
    
    ReDim totalHR_H_Table(1 To GlobalEnv.Hr_TableSize)  ' 회사가 보유한 모든 상급자
    ReDim totalHR_M_Table(1 To GlobalEnv.Hr_TableSize)  ' 회사가 보유한 모든 중급자
    ReDim totalHR_L_Table(1 To GlobalEnv.Hr_TableSize)  ' 회사가 보유한 모든 초급자

    Dim index As Integer
    For index = 1 To GlobalEnv.Hr_TableSize
        totalHR_H_Table(index) = GlobalEnv.Hr_Init_H    ' 회사가 보유한 상급자 초기화
        totalHR_M_Table(index) = GlobalEnv.Hr_Init_M    ' 회사가 보유한 중급자 초기화
        totalHR_L_Table(index) = GlobalEnv.Hr_Init_L    ' 회사가 보유한 초급자 초기화
    Next


    ReDim m_AllDashboardInfo(1 To debugArrayHeight, 1 To debugArrayWidth)
    
End Sub


' 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
Public Function Decision(thisWeek As Integer)

    ' 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
    Call CheckLastWeek(thisWeek)
    
    ' 2. 진행 가능한 후보프로젝트들을  candidateTable에 모은다
    Call SelectionOfCandidates(thisWeek)

    ' 3. 신규 프로젝트 선택 및 진행프로젝트 업데이트
    ' 이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
    Call SelectNewProject(thisWeek)

    ' Call comPrintDashboard()
    '
    
End Function

Private Function comPrintDashboard()
    
    'Call ClearSheet(gWsDashboard)          '시트의 모든 내용을 지우고 셀 병합 해제

    Dim startRow    As Long
    Dim arrHeader   As Variant
    arrHeader = Array("월", "누계", "prjNum")

    startRow = 6
    Call PrintArrayWithLine(g_WsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(g_WsDashboard, startRow + 1, 2, gWeekNumberTable) '기간을 적고
    Call PrintArrayWithLine(g_WsDashboard, startRow + 2, 2, doingTable)       ' 내용을 적는다.

    startRow = startRow + doingTableSize + 2
    Call PrintArrayWithLine(g_WsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(g_WsDashboard, startRow + 1, 2, gWeekNumberTable) '기간을 적고
    Call PrintArrayWithLine(g_WsDashboard, startRow + 2, 2, doneTable)    ' 내용을 적는다.

    startRow = startRow + doneTableSize + 2
    Call PrintArrayWithLine(g_WsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(g_WsDashboard, startRow + 1, 2, gWeekNumberTable) '기간을 적고
    Call PrintArrayWithLine(g_WsDashboard, startRow + 2, 2, defferTable)      ' 내용을 적는다.
    
End Function

' 후보 프로젝트들(candidateTable)중에서 가능한것들을 찾는다.
' 가능한 프로젝트들 중에서 정책에 따라서 진행할 프로젝트를 선택한다.
Private Function SelectNewProject(thisWeek As Integer)

    Dim cProject        As clsProject
    Dim cnt             As Integer
    Dim Id              As Integer

    Dim startProjectNum As Integer
    Dim endProjectNum   As Integer


    If thisWeek = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
        Exit Function
    End If
    
    cnt = m_orderTable(2, thisWeek)
    If cnt = 0 Then ' 이번주에 발주된 프로젝트가 없다.
        Exit Function
    End If

    startProjectNum = m_orderTable(1, thisWeek) + 1  ' 이번기간에 처음 프로젝트
    endProjectNum = m_orderTable(1, thisWeek) + cnt  ' 이번기간의 마지막 프로젝트

    ' song ==> 일단 모든 프로젝트를 다 하는것으로 하자. 추후 조건을 추가한다. 인원, 정책
    For Id = startProjectNum To endProjectNum

        Set cProject = m_projectTable(Id)
        
        If cProject.projectType = P_TYPE_INTERNAL Then   ' 내부프로젝트이면

        ' song ==> 프로젝트를 선택할지 말지 결정하는 루틴이 필요함. 현재는 모두 선택하는 것으로
        Else
            cProject.startDate = cProject.possiblestartDate
            
            ' song 시뮬레이션 기간 이후에 시작되는 프로젝트는 무시한다.
            If cProject.startDate > GlobalEnv.SimulationWeeks Then
                Exit Function
            End If
            
            ' song 프로젝트 추가시 Call AddWorkProject(cProject) '프로젝트 관리 테이블들, 프로젝트와 관련된 테이블들 모두 업데이트 하는 코드로 변경예정
            Call AddProjectEntry(doingTable, doingTableSize, cProject.startDate, cProject.projectNum) ' Doing table에 추가
            Call AddProjectHR(thisWeek, cProject) ' 이번주부터 해당 프로젝트에 투입되는 인력정보 갱신
            
            ' 프로젝트 완료시 Call RemoveHrTable(cProject) ' HR Table에서 인력 정보를 제거한다.
            
        End If
    Next

End Function

'' 신규프로젝트를 할 수 있는 인원이 있는가?
Private Function comCheckHR(cProject As clsProject)

'    needHR_H As Integer
'    needHR_M As Integer
'    needHR_L As Integer

    

End Function

' 완료프로젝트 검사 및 진행프로젝트 업데이트
' 1. 지난 기간의 정보를 이번기간에 복사하고
' 2. 지난 기간에 진행중인 프로젝트중 완료된 것이 있는가?
' 3. 완료된 프로젝트들만 이번기간에서 삭제
Private Function CheckLastWeek(thisWeek As Integer)

    Dim cProject As clsProject
    Dim cnt     As Integer
    Dim index       As Integer
    Dim Id      As Integer


    If thisWeek = 1 Then ' song ==> 현재까지 코드에서는 처음 시작에는 지난주가 없다
        Exit Function
    End If
    
    cnt = doingTable(1, thisWeek - 1) ' 지난주에 진행 중이던 프로젝트의 갯수
    If cnt = 0 Then ' song ==> 지난주에 진행중이던 프로젝트가 없다.
        Exit Function
    End If

    ' 지난 기간의 정보를 이번기간에 복사하고
    'Call CopyLastToCurrent(doingTable, thisWeek, cnt)


    Dim writePos As Integer
    writePos = 2 ' 첫줄은 "누계"를 기록하므로 2행부터
    
    '복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
    For index = 1 To cnt

        Id = doingTable(index + 1, thisWeek - 1) ' 지난주 정보

        If Id = 0 Then
            Exit For
        End If

        Set cProject = m_projectTable(Id) ' ID에 해당하는 projct 객체를 가져온다.
        
        If cProject.projectType = P_TYPE_EXTERNAL Then   ' 외내부프로젝트이면
        
            ' 프로젝트 Object 검사
            If thisWeek < (cProject.startDate + cProject.projectDuration - 1) Then ' 아직 안끝났으면
                'Call RemoveFromTable(doingTable, doingTableSize, thisWeek, cProject.projectNum) ' doing table에서 삭제
                'index = index - 1 ' 이번 행을 삭제하고 아래것을 가져왔으므로 이번 행부터 이어서 검사해야 한다.
                If IsEnoughHR(thisWeek, cProject) Then
                    doingTable(writePos, thisWeek) = cProject.projectNum
                    writePos = writePos + 1
                End If
            End If
        Else
            Debug.Print "내부 프로젝트"
            
        End If
    Next

End Function

' 지난 기간의 정보를 이번기간에 복사하고 복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
' 복사할 테이블, 현재기간, 이전기간의 프로젝트 갯수
Private Function CopyLastToCurrent(prjTable As Variant, ByVal thisWeek As Integer, ByVal lastProjectCnt As Integer)

    ' song ==> 예외 처리 필요
    Dim index As Integer
    
    For index = 1 To (lastProjectCnt + 1)
        prjTable(index, thisWeek) = prjTable(index, thisWeek - 1)
    Next
    
End Function



' 2차원 배열의 크기를 변경하고 기존 값을 복사한다.
Public Function Resize2DArray(ByRef arr As Variant, newRows As Integer, newCols As Integer)

    Dim oldRows As Integer
    Dim oldCols As Integer
    Dim index As Integer
    Dim j As Integer
    
    oldRows = UBound(arr, 1)
    oldCols = UBound(arr, 2)
    
    Dim tempArray() As Variant
    ReDim tempArray(1 To newRows, 1 To newCols)
    
    For index = 1 To oldRows
        For j = 1 To oldCols
            tempArray(index, j) = arr(index, j)
        Next j
    Next index
    
    ' 자료형의 문제로 두번 복사해야만 한다.
    ReDim arr(1 To newRows, 1 To newCols)

    For index = 1 To oldRows
        For j = 1 To oldCols
            arr(index, j) = tempArray(index, j)
        Next j
    Next index
    
End Function

' 테이블의 week에 projectID 를 추가한다.
' 테이블의 크기를 벗어나면 테이블의 크기를 변경한다.
' tableSize 를 변경한다. (리턴함)
Public Function AddProjectEntry(Table As Variant, TableRow As Integer, addWeek As Integer, projectID As Integer) As Integer

    Dim addRow As Integer ' ProjectId를 기록할 행번호
    
    ' 1번행은 해더(누계)   addRow = 현재 주(week)의 누계 + 헤더(누계) 1행 + 다음 1행
    addRow = Table(1, addWeek) + 2
    
    ' 추가할 행이 테이블의 끝을 벗어나면 배열을 다시 할당
    If addRow >= TableRow Then
        TableRow = addRow '  리턴해주고 있지만 안전하게 원본인 멤버변수도 변경한다. ByRef
        Call Resize2DArray(Table, TableRow, GlobalEnv.SimulationWeeks)
    End If
    
    ' 프로젝트 ID를 Table의 next 위치에 추가
    Table(addRow, addWeek) = projectID
    
    ' 첫 번째 행의 값을 1 증가시켜 누계를 업데이트
    Table(1, addWeek) = Table(1, addWeek) + 1
    
    ' 최종 변경된 행의 크기를 반환
    AddProjectEntry = TableRow
    
End Function

' '' 프로젝트 관리 테이블에 프로젝트를 넣는 범용적인 함수 (Doing, Done, Deffer )
' Sub AddProjectEntry(prjTable As Variant, week As Integer, projectID As Integer)

'     Dim sumRow As Integer
'     Dim lastEmptyRow As Integer

'     ' 배열의 크기 확인
'     If UBound(prjTable, 2) < week Or week < 1 Or week > m_SimulationDuration Then
'         MsgBox "Week is out of range", vbExclamation
'         Exit Sub
'     End If

'     ' 합계 행 설정
'     sumRow = 1

'     ' 비어있는 마지막 행 찾기 (합계를 통해 결정)
'     lastEmptyRow = prjTable(sumRow, week) + 2
'     If lastEmptyRow > (sumRow + m_MaxIDsPerWeek) Then
'         MsgBox "No empty slot available in week " & week, vbExclamation
'         Exit Sub
'     End If

'     ' 프로젝트 ID 추가
'     prjTable(lastEmptyRow, week) = projectID

'     ' 합계 및 프로젝트 수 증가
'     prjTable(sumRow, week) = prjTable(sumRow, week) + 1
 
' End Sub

Sub RemoveFromTable(prjTable As Variant, tableSize As Integer, ByVal week As Integer, ByVal projectID As Integer)

    Dim index       As Integer
    Dim sumRow  As Integer ' sum 이 들어있는 위치
    Dim idCount As Integer  ' id들의 갯수
    Dim found   As Boolean
        
    sumRow = 1
    found = False
    idCount = prjTable(sumRow, week)

    ' 배열의 크기 확인
    If UBound(prjTable, 2) < week Or week < 1 Or week > GlobalEnv.SimulationWeeks Then
        MsgBox "Week is out of range", vbExclamation
        Exit Sub
    End If
    ' song ==> tableSize의 크기도 검증하자.
    
    ' 해당 주에서 projectID 찾기
    For index = (sumRow + 1) To (sumRow + idCount)
        If prjTable(index, week) = projectID Then
            found = True
            Exit For
        End If
    Next index

    ' ID를 발견하면 삭제 및 아래 행들 이동
    If found Then

        For index = index To tableSize

            If index = tableSize Then 'ID가 마지막 행에 들어 있으면
                prjTable(index, week) = Empty
            Else
                prjTable(index, week) = prjTable(index + 1, week)
            End If

        Next index

        ' 합계 및 프로젝트 수 감소
        prjTable(sumRow, week) = prjTable(sumRow, week) - 1
        
    Else
        MsgBox "Project ID " & projectID & " not found in week " & week, vbExclamation
    End If

End Sub


' 진행 가능한 모든 후보프로젝트들을 모은다.
Function SelectionOfCandidates(thisWeek As Integer)

    Dim cProject        As clsProject
    Dim cnt             As Integer
    Dim Id              As Integer
    Dim index           As Integer
    

    Dim startProjectNum As Integer
    Dim endProjectNum   As Integer


    If thisWeek = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
        Exit Function
    End If
    
    '　candidateTable　을　＇０＇으로　초기화　한다．
    For cnt = 1 To MAX_CANDIDATES
        candidateTable(cnt) = 0
    Next
    
    cnt = m_orderTable(2, thisWeek)
    If cnt = 0 Then ' 이번주에 발주된 프로젝트가 없다.
        Exit Function
    End If

    startProjectNum = m_orderTable(1, thisWeek) + 1  ' 이번기간의 처음 프로젝트
    endProjectNum = m_orderTable(1, thisWeek) + cnt  ' 이번기간의 마지막 프로젝트

    ' 이번기간에 발주된 프로젝트을 모두 모은다.
    index = 1
    For Id = startProjectNum To endProjectNum

        Set cProject = m_projectTable(Id)
        

        If cProject.projectType = P_TYPE_INTERNAL Then   ' 내부프로젝트이면
        ' song ==> 내부 프로젝트는 별도 테이블로 관리하는것이 맞을것 같다. 추후 결정
        
        Else '
            If IsEnoughHR(thisWeek, cProject) = True Then
                candidateTable(index) = cProject.projectNum
            End If
        End If
        
        index = index + 1
    Next
    
End Function

 



Function IsEnoughHR(thisWeek As Integer, cProject As clsProject) As Boolean
        
    Dim activityCnt As Integer
    Dim activity    As Activity_
    Dim actStart    As Integer, actEnd As Integer ' activity 의 시작과 끝
        
    Dim index As Integer, j As Integer
    
    IsEnoughHR = True
    
    activityCnt = cProject.numActivities ' activity 갯수
    
    ' 원본 인력 테이블을 복사해서 프로젝트 인력을 추가 할 수 있는지 확인한다.
    Dim tempH_Table() As Integer
    Dim tempM_Table() As Integer
    Dim tempL_Table() As Integer
    
    ReDim tempH_Table(GlobalEnv.Hr_TableSize)
    ReDim tempM_Table(GlobalEnv.Hr_TableSize)
    ReDim tempL_Table(GlobalEnv.Hr_TableSize)
        
    For index = 1 To GlobalEnv.Hr_TableSize
        tempH_Table(index) = doingHR_H_Table(index)
        tempM_Table(index) = doingHR_M_Table(index)
        tempL_Table(index) = doingHR_L_Table(index)
    Next
    
    
    '2중 루프 activity->기간-> 등급업데이트 순서로
    ' activity들을 순서대로 가져온다.
    For index = 1 To activityCnt
    
        ' Activity 테이블에서 정보를 가져온다.
        activity = GetProjectActivity(cProject, index)
        
        'activity 시작시간을 시뮬레이션의 절대 시간으로 변경
        'song 리드타임을 고려하자
        actStart = activity.startDate
        actEnd = activity.endDate
        
        ' activity의 수행 기간동안
        For j = actStart To actEnd
            
            ' activity의 필요한 인력을 임시 테이블에 추가한다.
            Dim tempCnt As Integer
            
            tempCnt = tempH_Table(j) + activity.highSkill
            tempH_Table(j) = tempCnt
            
            tempCnt = tempM_Table(j) + activity.midSkill
            tempM_Table(j) = tempCnt
            
            tempCnt = tempL_Table(j) + activity.lowSkill
            tempL_Table(j) = tempCnt
            
        Next j
        
    Next index
        
    
    ' 인력 풀과 비교해서 투입에 필요한 인원이 총 인원보다 크면 False 리턴
    ' 코딩 편의상 모든 기간을 체크한다.
    For index = 1 To GlobalEnv.Hr_TableSize
    
        If tempH_Table(index) > totalHR_H_Table(index) Then
            IsEnoughHR = False
            Exit Function
        End If
        
        If tempM_Table(index) > totalHR_M_Table(index) Then
            IsEnoughHR = False
            Exit Function
        End If
        
        If tempL_Table(index) > totalHR_L_Table(index) Then
            IsEnoughHR = False
            Exit Function
        End If
        
    Next
    
End Function


' song 에러이면 null 리턴하는 코드와 사용시 검사하는 코드 추가 필요
'srcTable 크기와 동일한 Table를 생성하고 srcTable 의 내용을 복사해서 리턴해 준다.
Function CopyTableToRedimTable(srcTable() As Variant) As Variant()

    ' srcTable의 크기를 구한다.
'    desTable() As Variant
'
'    CopyTableToRedimTable = desTable

End Function

' project ID 에 필요한 인력정보를 인력정보 테이블에 업데이트 한다.
Function AddProjectHR(thisWeek As Integer, cProject As clsProject)
    
    Dim activityCnt As Integer
    Dim activity    As Activity_
    Dim actStart    As Integer, actEnd As Integer ' activity 의 시작과 끝
            
    Dim index As Integer, j As Integer
    
    activityCnt = cProject.numActivities ' activity 갯수
        
    '2중 루프 activity->기간-> 등급업데이트 순서로
    ' activity들을 순서대로 가져온다.
    For index = 1 To activityCnt
    
        ' Activity 테이블에서 정보를 가져온다.
        activity = GetProjectActivity(cProject, index)
        
        'activity 시작시간을 시뮬레이션의 절대 시간으로 변경
        actStart = activity.startDate
        actEnd = activity.endDate
        
        ' 프로젝트 시작 시점부터 activity의 수행 기간동안
        For j = actStart To actEnd
            
            ' activity의 필요한 인력을 테이블에 추가한다.
            Dim tempCnt As Integer
            
            tempCnt = doingHR_H_Table(j) + activity.highSkill
            doingHR_H_Table(j) = tempCnt
            
            tempCnt = doingHR_M_Table(j) + activity.midSkill
            doingHR_M_Table(j) = tempCnt
            
            tempCnt = doingHR_L_Table(j) + activity.lowSkill
            doingHR_L_Table(j) = tempCnt
            
        Next
        
    Next
    
        
    ' HR_H의 테이블을 갱신한다.
    ' 정보 테이블에 정보를 기록한다.
End Function

' 프로젝트의 Activities 배열에서 index 에 해당하는 activity를 가져온다.
' 사용자 정의 type 구조체는 직접 접근이 안되어서 개별로 데이터를 가져와야 한다.
Private Function GetProjectActivity(cProject As clsProject, actNum As Integer) As Activity_

    Dim activity As Activity_
    
    ' 개별 필드를 가져와서 Activity 구조체에 할당
    activity.activityType = cProject.GetActivityActivityType(actNum)
    activity.duration = cProject.GetActivityDuration(actNum)
    activity.startDate = cProject.GetActivitystartDate(actNum)
    activity.endDate = cProject.GetActivityEndDate(actNum)
    activity.highSkill = cProject.GetActivityHighSkill(actNum)
    activity.midSkill = cProject.GetActivityMidSkill(actNum)
    activity.lowSkill = cProject.GetActivityLowSkill(actNum)
    
    ' 완성된 Activity 구조체를 반환
    GetProjectActivity = activity
End Function




Private Function SetProjectActivity(cProject As clsProject, index As Integer, activity As Activity_)

    Call cProject.SetActivityDuration(index, activity.activityType)
    Call cProject.SetActivityDuration(index, activity.duration)
    Call cProject.SetActivitystartDate(index, activity.startDate)
    Call cProject.SetActivityEndDate(index, activity.endDate)
    Call cProject.SetActivityHighSkill(index, activity.highSkill)
    Call cProject.SetActivityMidSkill(index, activity.midSkill)
    Call cProject.SetActivityLowSkill(index, activity.lowSkill)
    
End Function

Public Function CreateProjects()

    ' 발생한 프로젝트의 갯수를 테이블에 기록한다.
    Dim week As Integer
    Dim cnt As Integer
    Dim sum As Integer
    
    ReDim m_orderTable(2, GlobalEnv.SimulationWeeks)

    For week = 1 To GlobalEnv.SimulationWeeks
    
        cnt = PoissonRandom(GlobalEnv.WeeklyProb)        ' 이번주 발생하는 프로젝트 갯수
        m_orderTable(1, week) = sum             ' 누계
        m_orderTable(2, week) = cnt    ' 발생 프로젝트갯수
        
        ' 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다. ==> 이전주까지 발생한 프로젝트 갯수후위연산. vba에서 do while 문법 모름... ㅎㅎ
        sum = sum + cnt
        
    Next week

    m_totalProjectNum = sum
       
    
    ''''''''''''''''''''''
    ' 디버깅 모드 이면 비교를 위해 발주 현황은 출력해 놓는다.
    If g_ProjDebug Then
        Dim arrHeader As Variant
        
        arrHeader = Array("주", "누계", "발주")
        arrHeader = PivotArray(arrHeader)
    
        ReDim tempArray(1 To GlobalEnv.SimulationWeeks) As Integer
        For cnt = 1 To GlobalEnv.SimulationWeeks
            tempArray(cnt) = cnt
        Next
        Call PrintArrayWithLine(g_WsDashboard, 2, 1, arrHeader)
        Call PrintArrayWithLine(g_WsDashboard, 2, 2, tempArray)
        Call PrintArrayWithLine(g_WsDashboard, 3, 2, m_orderTable)
        
    End If
    ' 디버깅 모드 이면 비교를 위해 발주 현황은 출력해 놓는다.
    ''''''''''''''''''''''
    
    Dim Id As Integer
    Dim startNum As Integer
    Dim endNum As Integer
    Dim preTotal As Integer
    Dim tempPrj As clsProject
    
    ' 프로젝트 관리하는 테이블에 생성한 프로젝트 object 를 넣는다.
    ReDim m_projectTable(m_totalProjectNum)

    MainForm.ProgressBar1.Max = GlobalEnv.SimulationWeeks
    MainForm.ProgressBar1.Min = 0
    MainForm.ProgressBar1.Text = "프로젝트 생성중"
    
    For week = 1 To GlobalEnv.SimulationWeeks
        preTotal = m_orderTable(1, week) ' 지난주까지의 발주 프로젝트 누계
        startNum = preTotal + 1          ' 신규 프로젝트의 시작번호 = 누계 +1
        endNum = preTotal + m_orderTable(2, week)  ' 마지막 프로젝트의 시작번호  =  지난주 누계 + 이번주 발생

        If startNum = 0 Then GoTo Continue
        If startNum > endNum Then GoTo Continue

        ' 이번 주에 발생한 프로젝트들을 생성하고 초기화 한다.
        For Id = startNum To endNum
            Set tempPrj = New clsProject
            Call tempPrj.Init(P_TYPE_EXTERNAL, Id, week)
            Set m_projectTable(Id) = tempPrj
        Next Id
        
Continue:
        MainForm.ProgressBar1.value = week
        Call DebugProjectSheet(endNum)
        
    Next week
    
End Function
'
'Call CreateOrderTable   ' Order 테이블을 생성하고 '주'을 입력한다.
'            Call CreateProjects     ' Order 테이블의 내용에 따라서 프로젝트를 생성한다.
'            Call PrintDashboard     ' Order 테이블과 인력정보를 대시보드 시트에 출력한다.
'            Call PrintProjectHeader ' Project 시트의 헤더를 기록한다.
'            Call PrintProjectAll    ' 프로젝트 전체를 출력한다

' data.xlsm 파일에서 order 테이블과 project 테이블을 읽어들인다.
Public Function LoadProjectsFromExcel() As Boolean

    ' Load Order Table
    Dim startIndex As Long
    
    startIndex = ORDER_TABLE_INDEX + 2
    
    ReDim m_orderTable(1 To 2, 1 To GlobalEnv.SimulationWeeks)
    
    With g_WsDashboard
        m_orderTable = .Range(.Cells(startIndex, 2), .Cells(startIndex + 1, GlobalEnv.SimulationWeeks + 1)).value
    End With

    m_totalProjectNum = m_orderTable(1, GlobalEnv.SimulationWeeks) + m_orderTable(2, GlobalEnv.SimulationWeeks)
    

    ' data.xlsm 파일에서 프로젝트를 읽어들인다.
    Dim projectID As Integer
    Dim startRow As Long, endRow As Long
    
    Dim tempProject As clsProject
    Dim projectInfo As Variant
    
    ReDim m_projectTable(m_totalProjectNum)

    For projectID = 1 To m_totalProjectNum
    
        ' 프로젝트 생성 및 설정
        Set tempProject = New clsProject 'CF 배열과 Activity 배열을 할당한다.
        
        startRow = PROJECT_TABLE_INDEX + (projectID - 1) * PRJ_SHEET_HEADER_H + 1
        endRow = startRow + PRJ_SHEET_HEADER_H - 1

        With g_WsProject
            projectInfo = .Range(.Cells(startRow, 1), .Cells(endRow, PRJ_SHEET_HEADER_W)).value
        End With

        tempProject.projectType = projectInfo(1, 1)
        tempProject.projectNum = projectInfo(1, 2)
        tempProject.projectDuration = projectInfo(1, 3)
        tempProject.possiblestartDate = projectInfo(1, 4)
        tempProject.endDate = projectInfo(1, 5)
        tempProject.orderDate = projectInfo(1, 6)
        tempProject.profit = projectInfo(1, 7)
        tempProject.experience = projectInfo(1, 8)
        tempProject.successProbability = projectInfo(1, 9)
        
        ' Cash Folwer 설정
        Dim tempCF(1 To MAX_N_CF) As Integer
        Dim index As Integer
        
        For index = 1 To MAX_N_CF
            tempCF(index) = projectInfo(1, 10 + index)
        Next index
        
'        Call tempProject.CopyCashFlows(tempCF)

        ' payment 설정
        tempProject.firstPayment = projectInfo(1, 14)
        tempProject.middlePayment = projectInfo(1, 15)
        tempProject.finalPayment = projectInfo(1, 16)

        tempProject.numActivities = projectInfo(2, 1)
        tempProject.firstPaymentMonth = projectInfo(2, 11)
        tempProject.middlePaymentMonth = projectInfo(2, 12)
        tempProject.finalPaymentMonth = projectInfo(2, 13)
        
        
        ' Activity 설정.
        Dim tempAct As Activity_
        
        For index = 1 To tempProject.numActivities
            tempAct.duration = projectInfo(1 + index, 3)
            tempAct.startDate = projectInfo(1 + index, 4)
            tempAct.endDate = projectInfo(1 + index, 5)
            tempAct.highSkill = projectInfo(1 + index, 7)
            tempAct.midSkill = projectInfo(1 + index, 8)
            tempAct.lowSkill = projectInfo(1 + index, 9)
            Call SetProjectActivity(tempProject, index, tempAct)
        Next index
        
        Set m_projectTable(projectID) = tempProject
        
    Next projectID
    
End Function ' LoadProjects() As Boolean















' PROJECT_SHEET_NAME 시트에 헤더를 생성한다.
' VB 6.0에서 Option Base 1을 사용했더라도, Split 함수는 기본적으로 0 기반 배열을 반환
' 따라서 1 기반 배열로 변환하는 코드를 추가 함
Public Function PrintProjectHeader()

    Dim MyArray() As Variant, tempArray() As String, strHeader As String
        
    strHeader = "pType,PRJ_ID,기간,시작가능,끝,발주일,총수익,경험,성공%,CF갯수,CF1%,CF2%,CF3%,선금,중도,잔금"

    tempArray = Split(strHeader, ",")
    MyArray = ConvertToBase1(tempArray) ' TempArray를 1 기반 배열로 변환
    Call PrintArrayWithLine(g_WsProject, 1, 1, MyArray)
    
    'strHeader = ",Dur,start,end,HR_H,HR_M,HR_L,,,mon_cf1,mon_cf2,mon_cf3,,,,"
    strHeader = "act갯수,,Dur,start,end,,HR_H,HR_M,HR_L,,mon_cf1,mon_cf2,mon_cf3,,,,"
    tempArray = Split(strHeader, ",")
    MyArray = ConvertToBase1(tempArray)
    Call PrintArrayWithLine(g_WsProject, 2, 1, MyArray)
    
End Function


Function PrintProjectAll()

    Dim temp As clsProject
    Dim index As Integer

    For index = 1 To m_totalProjectNum
        Set temp = m_projectTable(index)
        Call temp.PrintProjectInfomation
    Next index
End Function



''''''''''''''''''''''''''''''''''''''''''''''
' print functions

' 이번주 까지의 프로젝트 정보를 모은다.
Private Function GetAllDashboardInfo(thisWeek As Integer)
    
    Dim index As Integer, j As Integer
    Dim strHeader As String
    Dim tempArray As Variant
    
    ' 1번열 완성
    strHeader = "주,누계,발주,,투입,HR_H,HR_M,HR_L,,여유,HR_H,HR_M,HR_L,,총원,HR_H,HR_M,HR_L,,,주,누계,prjNum,,,,,,,,,,주,누계,prjNum,,,,,,,,,,주,누계,prjNum,,,,,,,,,,"
    tempArray = Split(strHeader, ",") ' 0 번 인덱스를 사용함에유의
    
    For index = 1 To 48
        m_AllDashboardInfo(index + 1, 1) = tempArray(index - 1)
    Next
    
    ' 기간 1 ~ 156
    For index = 1 To thisWeek ' GlobalEnv.SimulationWeeks
        'orderTable
        m_AllDashboardInfo(2, index + 1) = index
        m_AllDashboardInfo(3, index + 1) = m_orderTable(1, index)
        m_AllDashboardInfo(4, index + 1) = m_orderTable(1, index)
                
        ' 투입 HR
        m_AllDashboardInfo(7, index + 1) = doingHR_H_Table(index)
        m_AllDashboardInfo(8, index + 1) = doingHR_M_Table(index)
        m_AllDashboardInfo(9, index + 1) = doingHR_L_Table(index)
        
        ' 여유 HR
        m_AllDashboardInfo(12, index + 1) = totalHR_H_Table(index) - doingHR_H_Table(index)
        m_AllDashboardInfo(13, index + 1) = totalHR_M_Table(index) - doingHR_M_Table(index)
        m_AllDashboardInfo(14, index + 1) = totalHR_L_Table(index) - doingHR_L_Table(index)
        
        ' 총 HR
        m_AllDashboardInfo(17, index + 1) = totalHR_H_Table(index)
        m_AllDashboardInfo(18, index + 1) = totalHR_M_Table(index)
        m_AllDashboardInfo(19, index + 1) = totalHR_L_Table(index)
        
        ' 진행, 완료, 보류 중인 프로젝트
        m_AllDashboardInfo(22, index + 1) = index
        m_AllDashboardInfo(34, index + 1) = index
        m_AllDashboardInfo(46, index + 1) = index
        
        For j = 1 To 10
            m_AllDashboardInfo(22 + j, index + 1) = doingTable(j, index)
            m_AllDashboardInfo(34 + j, index + 1) = doneTable(j, index)
            m_AllDashboardInfo(46 + j, index + 1) = defferTable(j, index)
        Next
        
    Next
    
End Function

' 테두리를 그린다.
Public Function PrintLine(ws As Worksheet, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, isAround)

    With ws
    
        Dim rng As Object
        Set rng = .Range(.Cells(Top, Left), .Cells(Bottom, Right))
        
        If isAround Then
            rng.BorderAround LineStyle:=1, Weight:=xlThin, ColorIndex:=1
        Else
            rng.Borders.LineStyle = xlContinuous
            rng.Borders.Weight = xlThin
            rng.Borders.ColorIndex = 1
        End If
        
    End With
    
End Function


' 기간이 변경되면 배열크기 수정해야 하는데 신경쓰기 싫어서 조금 크게
Public Function PrintSimulationResults(thisWeek As Integer)
    
    Dim index As Integer, j As Integer
            
    ' 정보를 모을 배열을 초기화 한다.
    For index = 1 To debugArrayHeight
        For j = 1 To debugArrayWidth
            m_AllDashboardInfo(index, j) = ""
        Next
    Next
        
    Call GetAllDashboardInfo(thisWeek) ' 정보를 모은다
    
    Call ClearSheet(g_WsDashboard) ' 시트를 지운다.
        
    With g_WsDashboard
        
        ' 모은 정보를 시트에 출력한다.
        .Range(.Cells(1, 1), .Cells(debugArrayHeight, GlobalEnv.SimulationWeeks)).value = m_AllDashboardInfo
        
        ' 테두리를 그린다.
        Call PrintLine(g_WsDashboard, 1, 2, GlobalEnv.SimulationWeeks, 4, False)
        Call PrintLine(g_WsDashboard, 1, 7, GlobalEnv.SimulationWeeks, 9, False)
        Call PrintLine(g_WsDashboard, 1, 12, GlobalEnv.SimulationWeeks, 14, False)
        Call PrintLine(g_WsDashboard, 1, 17, GlobalEnv.SimulationWeeks, 19, False)
        Call PrintLine(g_WsDashboard, 1, 22, GlobalEnv.SimulationWeeks, 32, False)
        Call PrintLine(g_WsDashboard, 1, 34, GlobalEnv.SimulationWeeks, 44, False)
        Call PrintLine(g_WsDashboard, 1, 46, GlobalEnv.SimulationWeeks, 56, False)
        
    End With
    
End Function

Public Function DebugDashboard(thisWeek As Integer)
    If Not g_SimulDebug Then
       Exit Function
    End If
    Call PrintSimulationResults(thisWeek)
End Function



Private Function GetAllProjectInfo(projectCnt As Integer)

    Dim index As Integer, j As Integer
    Dim strHeader As String
    Dim tempArray As Variant
    Dim posY As Integer
    Dim cProject        As clsProject
    
        
    ' 1번열
    posY = 1
    strHeader = "pType,PRJ_ID,기간,시작가능,끝,발주일,총수익,경험,성공%,CF갯수,CF1%,CF2%,CF3%,선금,중도,잔금"
    tempArray = Split(strHeader, ",") ' 0번 인덱스 임을 주의 하자!!!
    For index = 1 To PRJ_SHEET_HEADER_W  ' 0번 인덱스 임을 주의 하자!!!
        m_AllProjectInfo(posY, index) = tempArray(index - 1) ' 0번 인덱스 임을 주의 하자!!!
    Next
    
    ' 2번열
    posY = posY + 1
    strHeader = "act갯수,,Dur,start,end,,HR_H,HR_M,HR_L,,mon_cf1,mon_cf2,mon_cf3,,,,"
    tempArray = Split(strHeader, ",")
    For index = 1 To PRJ_SHEET_HEADER_W  ' 0번 인덱스 임을 주의 하자!!!
        m_AllProjectInfo(posY, index) = tempArray(index - 1) ' 0번 인덱스 임을 주의 하자!!!
    Next
    
    ' 프로젝트 정보
    For index = 1 To projectCnt '
    
        posY = PROJECT_TABLE_INDEX + (index - 1) * PRJ_SHEET_HEADER_H + 1
        Set cProject = m_projectTable(index)
        Call GetProjectInfomation(posY, cProject)
        
    Next
        
End Function


' 프로젝트의 정보들을 출력을 위한 배열에 옮겨 적는다.
Public Function GetProjectInfomation(posY As Integer, cProject As clsProject)

    Dim index As Integer
    Dim posX As Integer
    Dim cashFlow(MAX_N_CF) As Integer
    Dim activity As Activity_

    posX = 1: m_AllProjectInfo(posY, posX) = cProject.projectType
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.projectNum
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.projectDuration
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.possiblestartDate
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.endDate
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.orderDate
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.profit
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.experience
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.successProbability
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.numCashFlows
    
    Call cProject.CopyFromCashFlows(cashFlow)
    For index = 1 To MAX_N_CF
        posX = posX + 1: m_AllProjectInfo(posY, posX) = cashFlow(index)
    Next index

    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.firstPayment
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.middlePayment
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.finalPayment


    ' 줄바꿈
    posX = 1: posY = posY + 1: m_AllProjectInfo(posY, posX) = cProject.numActivities
    posX = posX + 10: m_AllProjectInfo(posY, posX) = cProject.firstPaymentMonth
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.middlePaymentMonth
    posX = posX + 1: m_AllProjectInfo(posY, posX) = cProject.finalPaymentMonth
       
    ' 2열 부터
    For index = 1 To cProject.numActivities
        Dim strAct As String, strNum As String
        strNum = Format(index, "#0")
        strAct = "Activity" & strNum

        activity = GetProjectActivity(cProject, index)
        
        posX = 2:        m_AllProjectInfo(posY, posX) = strAct
        posX = posX + 1: m_AllProjectInfo(posY, posX) = activity.duration
        posX = posX + 1: m_AllProjectInfo(posY, posX) = activity.startDate
        posX = posX + 1: m_AllProjectInfo(posY, posX) = activity.endDate
        
        posX = posX + 2: m_AllProjectInfo(posY, posX) = activity.highSkill
        posX = posX + 1: m_AllProjectInfo(posY, posX) = activity.midSkill
        posX = posX + 1: m_AllProjectInfo(posY, posX) = activity.lowSkill
        
        posY = posY + 1
    Next index
    
    GetProjectInfomation = posY
      
End Function


Public Function PrintCreateProjectResults(projectCnt As Integer)
    
    Dim index As Integer, j As Integer
    Dim printRow As Integer
    
    printRow = PROJECT_TABLE_INDEX + m_totalProjectNum * PRJ_SHEET_HEADER_H
    
    ReDim m_AllProjectInfo(printRow, PRJ_SHEET_HEADER_W)
                
    ' 정보를 모을 배열을 초기화 한다.
    For index = 1 To printRow
        For j = 1 To PRJ_SHEET_HEADER_W
            m_AllProjectInfo(index, j) = ""
        Next
    Next
        
    Call GetAllProjectInfo(projectCnt) ' 정보를 모은다
        
    Call ClearSheet(g_WsProject) ' 시트를 지운다.
        
    With g_WsProject
        
        ' 모은 정보를 시트에 출력한다.
        .Range(.Cells(1, 1), .Cells(printRow, PRJ_SHEET_HEADER_W)).value = m_AllProjectInfo
        
        ' 헤더의 테두리를 그린다.
        Call PrintLine(g_WsProject, 1, 1, PRJ_SHEET_HEADER_W, 2, False)
        
        Dim projectRow As Integer
        For index = 1 To m_totalProjectNum
            projectRow = PROJECT_TABLE_INDEX + 1 + PRJ_SHEET_HEADER_H * (index - 1) ' 3칸+공간1 + 정보의 크기
            Call PrintLine(g_WsProject, 1, projectRow, PRJ_SHEET_HEADER_W, projectRow + PRJ_SHEET_HEADER_H - 1, True)
        Next
        
    End With
    
End Function


Public Function DebugProjectSheet(projectCnt As Integer)

    If Not g_ProjDebug Then
       Exit Function
    End If
    
    Call PrintCreateProjectResults(projectCnt)
    
End Function
