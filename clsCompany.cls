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

Private doingTable()     As Integer  ' 진행중인 프로젝트들의 ID를 관리하는 테이블
Private doneTable()      As Integer  ' 완료된 프로젝트들의 ID를 관리하는 테이블
Private defferTable()    As Integer  ' 보류된 프로젝트들의 ID를 관리하는 테이블

Private candidateTable() As Integer   ' 후보　프로젝트들의 ID를 관리하는 테이블

Public doingTableSize    As Integer      ' comDoneTable의 현재크기 song ==> 동적으로 수정하자.
Public doneTableSize     As Integer      ' comDoneTable의 현재크기 song ==> 동적으로 수정하자.
Public defferTableSize   As Integer      ' comDoneTable의 현재크기 song ==> 동적으로 수정하자.

Private projectTable()   As clsProject

Private tempHRTable()       As Integer
Private doingHR_HTable()  As Integer  ' 진행중인 프로젝트에 투입된 상급자
Private doingHR_MTable()  As Integer  ' 진행중인 프로젝트에 투입된 중급자
Private doingHR_LTable()  As Integer  ' 진행중인 프로젝트에 투입된 초급자

Private totalHR_HTable()  As Integer  ' 회사가 보유한 상급자
Private totalHR_MTable()  As Integer  ' 회사가 보유한 중급자
Private totalHR_LTable()  As Integer  ' 회사가 보유한 초급자
 
Public companyID As Integer ' 회사 이름


Private Const MAX_CANDIDATES As Integer = 50 ' SONG 임의로 지정. 동적으로 변하게 할 필요까지는 없어 보임.

' Public Property Get PropertyDoingTable()  As Variant
'     PropertyDoingTable = comDoingTable
' End Property

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



' 생성자와 소멸자
Private Sub Class_Initialize()



End Sub

' company ID와 발생할 프로젝트 테이블을 저장한다.
' 프로젝트, 인력, 자금을 관리할 테이블 생성
Public Sub Init(Id As Integer)

    companyID = Id
    projectTable = gProjectTable
    
    doingTableSize = 10
    doneTableSize = 10
    defferTableSize = 10

    ReDim doingTable(1 To doingTableSize, 1 To GlobalEnv.SimulationWeeks)     ' 진행중인 프로젝트들의 ID를 관리하는 테이블
    ReDim doneTable(1 To doneTableSize, 1 To GlobalEnv.SimulationWeeks)       ' 완료된 프로젝트들의 ID를 관리하는 테이블
    ReDim defferTable(1 To defferTableSize, 1 To GlobalEnv.SimulationWeeks)   ' 보류된 프로젝트들의 ID를 관리하는 테이블
    
    ReDim candidateTable(1 To MAX_CANDIDATES) ' 후보 프로젝트들의 ID를 관리하는 테이블

    ReDim tempHRTable(1 To GlobalEnv.SimulationWeeks)       ' 인원 계산을 위한 임시 테이블
    ReDim doingHR_HTable(1 To GlobalEnv.SimulationWeeks)  ' 프로젝트에 투입된 상급자
    ReDim doingHR_MTable(1 To GlobalEnv.SimulationWeeks)  ' 프로젝트에 투입된 중급자
    ReDim doingHR_LTable(1 To GlobalEnv.SimulationWeeks)  ' 프로젝트에 투입된 초급자
    
    ReDim totalHR_HTable(1 To GlobalEnv.SimulationWeeks)  ' 회사가 보유한 모든 상급자
    ReDim totalHR_MTable(1 To GlobalEnv.SimulationWeeks)  ' 회사가 보유한 모든 중급자
    ReDim totalHR_LTable(1 To GlobalEnv.SimulationWeeks)  ' 회사가 보유한 모든 초급자

    totalHR_HTable(1) = GlobalEnv.Hr_Init_H    ' 회사가 보유한 상급자 초기화
    totalHR_MTable(1) = GlobalEnv.Hr_Init_M    ' 회사가 보유한 중급자 초기화
    totalHR_LTable(1) = GlobalEnv.Hr_Init_L    ' 회사가 보유한 초급자 초기화

End Sub


Private Sub Class_Terminate()
    ' 종료할때 버퍼의 크기를 확인해서 가까운 값으로 초기값을 결정하자.
    Debug.Print "Class is being terminated"; doingTableSize, doneTableSize, defferTableSize
End Sub



'' 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
Public Function Decision(thisWeek As Integer)

    ' 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
    Call CheckLastWeek(thisWeek)
    
    '2. 진행 가능한 후보프로젝트들 선택. 이번주 차수를 전달 받아 다음주에 진행 가능한 후보들을 선정한다.
    Call SelectionOfCandidates(thisWeek)

    ' 2. 신규 프로젝트 선택 및 진행프로젝트 업데이트
    '  이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
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
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '기간을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doingTable)       ' 내용을 적는다.

    startRow = startRow + doingTableSize + 2
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '기간을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doneTable)    ' 내용을 적는다.

    startRow = startRow + doneTableSize + 2
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '기간을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, defferTable)      ' 내용을 적는다.
    
End Function

Private Function SelectNewProject(thisWeek As Integer)

    Dim cProject        As clsProject
    Dim cnt             As Integer
    Dim Id              As Integer

    Dim startProjectNum As Integer
    Dim endProjectNum   As Integer


    If thisWeek = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
        Exit Function
    End If
    
    cnt = gOrderTable(2, thisWeek)
    If cnt = 0 Then ' 이번주에 발주된 프로젝트가 없다.
        Exit Function
    End If

    startProjectNum = gOrderTable(1, thisWeek) + 1  ' 이번기간에 처음 프로젝트
    endProjectNum = gOrderTable(1, thisWeek) + cnt  ' 이번기간의 마지막 프로젝트

    ' song ==> 일단 모든 프로젝트를 다 하는것으로 하자. 추후 조건을 추가한다. 인원, 정책
    For Id = startProjectNum To endProjectNum

        Set cProject = projectTable(Id)
        
        If cProject.projectType = P_TYPE_INTERNAL Then   ' 내부프로젝트이면

        ' song ==> 프로젝트를 선택할지 말지 결정하는 루틴이 필요함. 현재는 모두 선택하는 것으로
        Else
            cProject.startDate = thisWeek
            
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
    Dim i       As Integer
    Dim Id      As Integer


    If thisWeek = 1 Then ' song ==> 현재까지 코드에서는 처음 시작에는 지난주가 없다
        Exit Function
    End If
    
    cnt = doingTable(1, thisWeek - 1) ' 지난주에 진행 중이던 프로젝트의 갯수
    If cnt = 0 Then ' song ==> 지난주에 진행중이던 프로젝트가 없다.
        Exit Function
    End If

    ' 지난 기간의 정보를 이번기간에 복사하고
    Call CopyLastToCurrent(doingTable, thisWeek, cnt)

    '복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
    For i = 1 To cnt

        Id = doingTable(i + 1, thisWeek) ' 프로젝트 객체의 ID를 가져온다

        If Id = 0 Then
            Exit For
        End If

        Set cProject = projectTable(Id) ' ID에 해당하는 projct 객체를 가져온다.
        
        If cProject.projectType = P_TYPE_EXTERNAL Then   ' 외내부프로젝트이면
        
            ' 프로젝트 Object 검사
            If thisWeek = (cProject.startDate + cProject.projectDuration - 1) Then ' 지난번에 끝났으면
                Call RemoveFromTable(doingTable, doingTableSize, thisWeek, cProject.projectNum) ' doing table에서 삭제
                i = i - 1 ' 이번 행을 삭제하고 아래것을 가져왔으므로 이번 행부터 이어서 검사해야 한다.
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
    Dim i As Integer
    
    For i = 1 To (lastProjectCnt + 1)
        prjTable(i, thisWeek) = prjTable(i, thisWeek - 1)
    Next
    
End Function




Public Function Resize2DArray(arr As Variant, newRows As Integer, newCols As Integer)

    Dim oldRows As Integer
    Dim oldCols As Integer
    Dim i As Integer
    Dim j As Integer
    
    oldRows = UBound(arr, 1)
    oldCols = UBound(arr, 2)
    
    Dim tempArray() As Variant
    ReDim tempArray(1 To newRows, 1 To newCols)
    
    For i = 1 To oldRows
        For j = 1 To oldCols
            tempArray(i, j) = arr(i, j)
        Next j
    Next i
    
    ReDim arr(1 To newRows, 1 To newCols)

      For i = 1 To oldRows
        For j = 1 To oldCols
            arr(i, j) = tempArray(i, j)
        Next j
    Next i
    
End Function


Public Function AddProjectEntry(prjTable As Variant, tableSize As Integer, week As Integer, projectID As Integer) As Integer

    Dim currentRow As Integer
    
    ' 현재 주(week)의 첫 번째 행 값(값이 들어 있는 행의 갯수)을 가져옴
    currentRow = prjTable(1, week) + 2
    
    ' 만약 현재 행의 값이 tableSize를 초과하면 배열을 다시 할당
    ' currentRow는 ID 갯수 , tableSize는 currentRow + sum 을 기록하는 추가 행
    If currentRow = (tableSize - 1) Then
        tableSize = tableSize + 1
        Call Resize2DArray(prjTable, tableSize, GlobalEnv.SimulationWeeks)
    End If
    
    ' 프로젝트 ID를 prjTable의 적절한 위치에 추가
    prjTable(currentRow, week) = projectID
    
    ' 첫 번째 행의 값을 1 증가시켜 행의 갯수를 업데이트
    prjTable(1, week) = prjTable(1, week) + 1
    
    ' 최종 변경된 행의 크기를 반환
    AddProjectEntry = tableSize
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

    Dim i       As Integer
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
    For i = (sumRow + 1) To (sumRow + idCount)
        If prjTable(i, week) = projectID Then
            found = True
            Exit For
        End If
    Next i

    ' ID를 발견하면 삭제 및 아래 행들 이동
    If found Then

        For i = i To tableSize

            If i = tableSize Then 'ID가 마지막 행에 들어 있으면
                prjTable(i, week) = Empty
            Else
                prjTable(i, week) = prjTable(i + 1, week)
            End If

        Next i

        ' 합계 및 프로젝트 수 감소
        prjTable(sumRow, week) = prjTable(sumRow, week) - 1
        
    Else
        MsgBox "Project ID " & projectID & " not found in week " & week, vbExclamation
    End If

End Sub


' 진행 가능한 후보프로젝트들 선택
' 이번주를 전달 받아 다음주에 진행 가능한 후보들을 선정한다.
' 리드타임 적용부터 먼저 정리하자.
Function SelectionOfCandidates(thisWeek As Integer)

    Dim cProject        As clsProject
    Dim cnt             As Integer
    Dim Id              As Integer

    Dim startProjectNum As Integer
    Dim endProjectNum   As Integer


    If thisWeek = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
        Exit Function
    End If
    
    '　comCandidateTable　을　＇０＇으로　초기화　한다．
    For cnt = 1 To MAX_CANDIDATES
        candidateTable(cnt) = 0
    Next
    
    cnt = gOrderTable(2, thisWeek)
    If cnt = 0 Then ' 이번주에 발주된 프로젝트가 없다.
        Exit Function
    End If

    startProjectNum = gOrderTable(1, thisWeek) + 1  ' 이번기간에 처음 프로젝트
    endProjectNum = gOrderTable(1, thisWeek) + cnt  ' 이번기간의 마지막 프로젝트

    ' song ==> 일단 모든 프로젝트를 다 하는것으로 하자. 추후 조건을 추가한다. 인원, 정책
    For Id = startProjectNum To endProjectNum

        Set cProject = projectTable(Id)
        
        If cProject.projectType = P_TYPE_INTERNAL Then   ' 내부프로젝트이면

        ' song ==> 프로젝트를 선택할지 말지 결정하는 루틴이 필요함. 현재는 모두 선택하는 것으로
        Else
            cProject.startDate = thisWeek
            Call AddProjectEntry(doingTable, doingTableSize, thisWeek, cProject.projectNum) ' Doing table에 추가
        End If
    Next
    
End Function


' 주어진 주간부터 Project 를 진행할 인적자원이 있는지 체크하는 함수
Function CheckHRAvailability(ByVal thisWeek As Integer, cProject As clsProject) As Boolean

    Dim index As Integer
    Dim startWeek As Integer, endWeek As Integer, totalHR As Integer
    Dim duration As Integer
        
    ' 인력자원이 운영이 가능한 프로젝트이면 True, 아니면 False 리턴
    CheckHRAvailability = True
    
    
    ' 프로젝트 시작주와 종료주를 구한다.
    startWeek = cProject.startDate
    endWeek = cProject.startDate + cProject.projectDuration - 1
    
    ' Activity 별로 프로젝트 기간동안 매주 추가되는 인원을 계산한다. 인원이 초과되면 프로젝트 진행 불가
    
    For index = startWeek To endWeek
        totalHR = tempHRTable(index)
        'totalHR = totalHR + cProject.PrjActivities
        
    Next
    
    
    ' 진행중인 프로젝트에 투입된 상급자
    For index = 1 To GlobalEnv.SimulationWeeks
        tempHRTable(index) = doingHR_HTable(index)
    Next
    
    
    ' 진행중인 프로젝트에 투입된 중급자
    For index = 1 To GlobalEnv.SimulationWeeks
        tempHRTable(index) = doingHR_MTable(index)
    Next
    
    ' 진행중인 프로젝트에 투입된 초급자
    For index = 1 To GlobalEnv.SimulationWeeks
        tempHRTable(index) = doingHR_LTable(index)
    Next
    

    Dim oldRows As Integer
    Dim oldCols As Integer
    Dim i As Integer
    Dim j As Integer
    
    
    
    
    
    
    ' 프로젝트 시작주와 종료주를 구한다.
    'startWeek =
    
    ' 프로젝트 기간동안 매주 추가되는 인원을 계산한다. 인원이 초과되면 프로젝트 진행 불가
    For index = startWeek To endWeek
        
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
    Dim startWeek As Integer, endWeek As Integer ' 프로젝트이 시작과 끝
    Dim actStart As Integer, actEnd As Integer ' activity 의 시작과 끝
    Dim activity As Activity_
    
    Dim index As Integer, j As Integer, k As Integer
    
    startWeek = cProject.startDate
    endWeek = cProject.startDate + cProject.projectDuration - 1 ' song activity의 마지막과 동일한지 검증용으로 사용하자
    activityCnt = cProject.numActivities ' activity 갯수
    
    
    '2중 루프 activity->기간-> 등급업데이트 순서로
    ' activity들을 순서대로 가져온다.
    For index = 1 To activityCnt
    
        ' Activity 테이블에서 정보를 가져온다.
        activity = GetProjectActivity(cProject, index)
        
        'activity 시작시간을 시뮬레이션의 절대 시간으로 변경
        actStart = activity.startDate + startWeek - 1
        actEnd = actStart + activity.duration - 1
        
        ' 프로젝트 시작 시점부터 activity의 수행 기간동안
        For j = startWeek To (startWeek + activity.duration)
            
            ' activity의 필요한 인력을 테이블에 추가한다.
            Dim tempCnt As Integer
            
            tempCnt = doingHR_HTable(thisWeek) + activity.HighSkill
            doingHR_HTable(thisWeek) = tempCnt
            
            tempCnt = doingHR_MTable(thisWeek) + activity.MidSkill
            doingHR_MTable(thisWeek) = tempCnt
            
            tempCnt = doingHR_LTable(thisWeek) + activity.LowSkill
            doingHR_LTable(thisWeek) = tempCnt
            
        Next
        
    Next
    
        
    ' HR_H의 테이블을 갱신한다.
    ' 정보 테이블에 정보를 기록한다.
End Function

' 프로젝트의 Activities 배열에서 index 에 해당하는 activity를 가져온다.
' 사용자 정의 type 구조체는 직접 접근이 안되어서 개별로 데이터를 가져와야 한다.
Private Function GetProjectActivity(cProject As clsProject, index As Integer) As Activity_

    Dim activity As Activity_
    
    ' 개별 필드를 가져와서 Activity 구조체에 할당
    activity.ActivityType = cProject.GetActivityActivityType(index)
    activity.duration = cProject.GetActivityDuration(index)
    activity.startDate = cProject.GetActivitystartDate(index)
    activity.EndDate = cProject.GetActivityEndDate(index)
    activity.HighSkill = cProject.GetActivityHighSkill(index)
    activity.MidSkill = cProject.GetActivityMidSkill(index)
    activity.LowSkill = cProject.GetActivityLowSkill(index)
    
    ' 완성된 Activity 구조체를 반환
    GetProjectActivity = activity
End Function



