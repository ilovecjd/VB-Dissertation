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

Private doingTable()     As Integer  ' 기간별로 진행중인 프로젝트들의 ID를 관리하는 테이블
Private doneTable()      As Integer  ' 기간별로 완료된 프로젝트들의 ID를 관리하는 테이블
Private defferTable()    As Integer  ' 기간별로 보류된 프로젝트들의 ID를 관리하는 테이블

Private candidateTable() As Integer   ' 후보　프로젝트들의 ID를 관리하는 테이블

Public doingTableSize    As Integer      ' DoneTable의 현재크기 song ==> 동적으로 수정하자.
Public doneTableSize     As Integer      ' DoneTable의 현재크기 song ==> 동적으로 수정하자.
Public defferTableSize   As Integer      ' DoneTable의 현재크기 song ==> 동적으로 수정하자.

Private projectTable()   As clsProject  ' 생성된(발주된) 프로젝트를 관리하는 테이블

Private doingHR_H_Table()  As Integer  ' 기간별로 진행중인 프로젝트에 투입된 상급자
Private doingHR_M_Table()  As Integer  ' 기간별로 진행중인 프로젝트에 투입된 중급자
Private doingHR_L_Table()  As Integer  ' 기간별로 진행중인 프로젝트에 투입된 초급자

Private totalHR_H_Table()  As Integer  ' 기간별로 회사가 보유한 상급자
Private totalHR_M_Table()  As Integer  ' 기간별로 회사가 보유한 중급자
Private totalHR_L_Table()  As Integer  ' 기간별로 회사가 보유한 초급자
 
Public companyID As Integer ' 회사 이름


Private Const MAX_CANDIDATES As Integer = 50 ' SONG 임의로 지정. 동적으로 변하게 할 필요까지는 없어 보임.


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
      
    projectTable = gProjectTable
    
    companyID = Id
    
    doingTableSize = 10
    doneTableSize = 10
    defferTableSize = 10

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
    
    ReDim totalHR_HTable(1 To GlobalEnv.Hr_TableSize)  ' 회사가 보유한 모든 상급자
    ReDim totalHR_MTable(1 To GlobalEnv.Hr_TableSize)  ' 회사가 보유한 모든 중급자
    ReDim totalHR_LTable(1 To GlobalEnv.Hr_TableSize)  ' 회사가 보유한 모든 초급자

    totalHR_HTable(1) = GlobalEnv.Hr_Init_H    ' 회사가 보유한 상급자 초기화
    totalHR_MTable(1) = GlobalEnv.Hr_Init_M    ' 회사가 보유한 중급자 초기화
    totalHR_LTable(1) = GlobalEnv.Hr_Init_L    ' 회사가 보유한 초급자 초기화

End Sub


' 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
Public Function Decision(ThisWeek As Integer)

    ' 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
    Call CheckLastWeek(ThisWeek)
    
    ' 2. 진행 가능한 후보프로젝트들을  candidateTable에 모은다
    Call SelectionOfCandidates(ThisWeek)

    ' 3. 신규 프로젝트 선택 및 진행프로젝트 업데이트
    ' 이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
    Call SelectNewProject(ThisWeek)

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

' 후보 프로젝트들(candidateTable)중에서 가능한것들을 찾는다.
' 가능한 프로젝트들 중에서 정책에 따라서 진행할 프로젝트를 선택한다.
Private Function SelectNewProject(ThisWeek As Integer)

    Dim cProject        As clsProject
    Dim cnt             As Integer
    Dim Id              As Integer

    Dim startProjectNum As Integer
    Dim endProjectNum   As Integer


    If ThisWeek = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
        Exit Function
    End If
    
    cnt = gOrderTable(2, ThisWeek)
    If cnt = 0 Then ' 이번주에 발주된 프로젝트가 없다.
        Exit Function
    End If

    startProjectNum = gOrderTable(1, ThisWeek) + 1  ' 이번기간에 처음 프로젝트
    endProjectNum = gOrderTable(1, ThisWeek) + cnt  ' 이번기간의 마지막 프로젝트

    ' song ==> 일단 모든 프로젝트를 다 하는것으로 하자. 추후 조건을 추가한다. 인원, 정책
    For Id = startProjectNum To endProjectNum

        Set cProject = projectTable(Id)
        
        If cProject.projectType = P_TYPE_INTERNAL Then   ' 내부프로젝트이면

        ' song ==> 프로젝트를 선택할지 말지 결정하는 루틴이 필요함. 현재는 모두 선택하는 것으로
        Else
            cProject.startDate = ThisWeek
            
            ' song 프로젝트 추가시 Call AddWorkProject(cProject) '프로젝트 관리 테이블들, 프로젝트와 관련된 테이블들 모두 업데이트 하는 코드로 변경예정
            Call AddProjectEntry(doingTable, doingTableSize, cProject.startDate, cProject.projectNum) ' Doing table에 추가
            Call AddProjectHR(ThisWeek, cProject) ' 이번주부터 해당 프로젝트에 투입되는 인력정보 갱신
            
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
Private Function CheckLastWeek(ThisWeek As Integer)

    Dim cProject As clsProject
    Dim cnt     As Integer
    Dim I       As Integer
    Dim Id      As Integer


    If ThisWeek = 1 Then ' song ==> 현재까지 코드에서는 처음 시작에는 지난주가 없다
        Exit Function
    End If
    
    cnt = doingTable(1, ThisWeek - 1) ' 지난주에 진행 중이던 프로젝트의 갯수
    If cnt = 0 Then ' song ==> 지난주에 진행중이던 프로젝트가 없다.
        Exit Function
    End If

    ' 지난 기간의 정보를 이번기간에 복사하고
    Call CopyLastToCurrent(doingTable, ThisWeek, cnt)

    '복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
    For I = 1 To cnt

        Id = doingTable(I + 1, ThisWeek) ' 프로젝트 객체의 ID를 가져온다

        If Id = 0 Then
            Exit For
        End If

        Set cProject = projectTable(Id) ' ID에 해당하는 projct 객체를 가져온다.
        
        If cProject.projectType = P_TYPE_EXTERNAL Then   ' 외내부프로젝트이면
        
            ' 프로젝트 Object 검사
            If ThisWeek = (cProject.startDate + cProject.projectDuration - 1) Then ' 지난번에 끝났으면
                Call RemoveFromTable(doingTable, doingTableSize, ThisWeek, cProject.projectNum) ' doing table에서 삭제
                I = I - 1 ' 이번 행을 삭제하고 아래것을 가져왔으므로 이번 행부터 이어서 검사해야 한다.
            End If
            
        Else
            Debug.Print "내부 프로젝트"
            
        End If
    Next

End Function

' 지난 기간의 정보를 이번기간에 복사하고 복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
' 복사할 테이블, 현재기간, 이전기간의 프로젝트 갯수
Private Function CopyLastToCurrent(prjTable As Variant, ByVal ThisWeek As Integer, ByVal lastProjectCnt As Integer)

    ' song ==> 예외 처리 필요
    Dim index As Integer
    
    For index = 1 To (lastProjectCnt + 1)
        prjTable(index, ThisWeek) = prjTable(index, ThisWeek - 1)
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
Public Function AddProjectEntry(Table As Variant, TableRow As Integer, ThisWeek As Integer, ProjectID As Integer) As Integer

    Dim addRow As Integer ' ProjectId를 기록할 행번호
    
    ' 1번행은 해더(누계)   addRow = 현재 주(week)의 누계 + 헤더(누계) 1행 + 다음 1행
    addRow = Table(1, ThisWeek) + 2
    
    ' 추가할 행이 테이블의 끝을 벗어나면 배열을 다시 할당
    If addRow >= TableRow Then
        TableRow = addRow '  리턴해주고 있지만 안전하게 원본인 멤버변수도 변경한다. ByRef
        Call Resize2DArray(Table, TableRow, GlobalEnv.SimulationWeeks)
    End If
    
    ' 프로젝트 ID를 Table의 next 위치에 추가
    Table(addRow, ThisWeek) = ProjectID
    
    ' 첫 번째 행의 값을 1 증가시켜 누계를 업데이트
    Table(1, ThisWeek) = Table(1, ThisWeek) + 1
    
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

Sub RemoveFromTable(prjTable As Variant, tableSize As Integer, ByVal Week As Integer, ByVal ProjectID As Integer)

    Dim I       As Integer
    Dim sumRow  As Integer ' sum 이 들어있는 위치
    Dim idCount As Integer  ' id들의 갯수
    Dim found   As Boolean
        
    sumRow = 1
    found = False
    idCount = prjTable(sumRow, Week)

    ' 배열의 크기 확인
    If UBound(prjTable, 2) < Week Or Week < 1 Or Week > GlobalEnv.SimulationWeeks Then
        MsgBox "Week is out of range", vbExclamation
        Exit Sub
    End If
    ' song ==> tableSize의 크기도 검증하자.
    
    ' 해당 주에서 projectID 찾기
    For I = (sumRow + 1) To (sumRow + idCount)
        If prjTable(I, Week) = ProjectID Then
            found = True
            Exit For
        End If
    Next I

    ' ID를 발견하면 삭제 및 아래 행들 이동
    If found Then

        For I = I To tableSize

            If I = tableSize Then 'ID가 마지막 행에 들어 있으면
                prjTable(I, Week) = Empty
            Else
                prjTable(I, Week) = prjTable(I + 1, Week)
            End If

        Next I

        ' 합계 및 프로젝트 수 감소
        prjTable(sumRow, Week) = prjTable(sumRow, Week) - 1
        
    Else
        MsgBox "Project ID " & ProjectID & " not found in week " & Week, vbExclamation
    End If

End Sub


' 진행 가능한 모든 후보프로젝트들을 모은다.
Function SelectionOfCandidates(ThisWeek As Integer)

    Dim cProject        As clsProject
    Dim cnt             As Integer
    Dim Id              As Integer
    Dim index           As Integer
    

    Dim startProjectNum As Integer
    Dim endProjectNum   As Integer


    If ThisWeek = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
        Exit Function
    End If
    
    '　candidateTable　을　＇０＇으로　초기화　한다．
    For cnt = 1 To MAX_CANDIDATES
        candidateTable(cnt) = 0
    Next
    
    cnt = gOrderTable(2, ThisWeek)
    If cnt = 0 Then ' 이번주에 발주된 프로젝트가 없다.
        Exit Function
    End If

    startProjectNum = gOrderTable(1, ThisWeek) + 1  ' 이번기간의 처음 프로젝트
    endProjectNum = gOrderTable(1, ThisWeek) + cnt  ' 이번기간의 마지막 프로젝트

    ' 이번기간에 발주된 프로젝트을 모두 모은다.
    index = 1
    For Id = startProjectNum To endProjectNum

        Set cProject = projectTable(Id)
        

        If cProject.projectType = P_TYPE_INTERNAL Then   ' 내부프로젝트이면
        ' song ==> 내부 프로젝트는 별도 테이블로 관리하는것이 맞을것 같다. 추후 결정
        
        Else '
            If IsEnoughHR(ThisWeek, cProject) = True Then
                candidateTable(index) = cProject.projectNum
            End If
        End If
        
        index = index + 1
    Next
    
End Function





Function IsEnoughHR(ThisWeek As Integer, cProject As clsProject) As Boolean
        
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
        actStart = ThisWeek + activity.startDate - 1
        actEnd = actStart + activity.duration - 1
        
        ' 프로젝트 시작 시점부터 activity의 수행 기간동안
        For j = actStart To actEnd
            
            ' activity의 필요한 인력을 임시 테이블에 추가한다.
            Dim tempCnt As Integer
            
            tempCnt = tempH_Table(j) + activity.HighSkill
            tempH_Table(j) = tempCnt
            
            tempCnt = tempM_Table(j) + activity.MidSkill
            tempM_Table(j) = tempCnt
            
            tempCnt = tempL_Table(j) + activity.LowSkill
            tempL_Table(j) = tempCnt
            
        Next j
        
    Next index
        
    
    ' 인력 풀과 비교해서 투입에 필요한 인원이 총 인원보다 크면 False 리턴
    ' 코딩 편의상 모든 기간을 체크한다.
    For index = 1 To GlobalEnv.Hr_TableSize
    
        If tempH_Table(index) > doingHR_H_Table(index) Then
            IsEnoughHR = False
            Exit Function
        End If
        
        If tempM_Table(index) > doingHR_M_Table(index) Then
            IsEnoughHR = False
            Exit Function
        End If
        
        If tempL_Table(index) > doingHR_L_Table(index) Then
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
Function AddProjectHR(ThisWeek As Integer, cProject As clsProject)
    
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
        actStart = ThisWeek + activity.startDate - 1
        actEnd = actStart + activity.duration - 1
        
        ' 프로젝트 시작 시점부터 activity의 수행 기간동안
        For j = actStart To actEnd
            
            ' activity의 필요한 인력을 테이블에 추가한다.
            Dim tempCnt As Integer
            
            tempCnt = doingHR_H_Table(j) + activity.HighSkill
            doingHR_H_Table(j) = tempCnt
            
            tempCnt = doingHR_M_Table(j) + activity.MidSkill
            doingHR_M_Table(j) = tempCnt
            
            tempCnt = doingHR_L_Table(j) + activity.LowSkill
            doingHR_L_Table(j) = tempCnt
            
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



Public Function PrintSimulationResults()
    
    'Call ClearSheet(gWsDashboard)          '시트의 모든 내용을 지우고 셀 병합 해제

    Dim startRow    As Long
    Dim arrHeader   As Variant
    Dim tempArray() As Integer
        
    
    ' 인력테이블을 기록
    ' 상급자 투입 테이블
    startRow = 6
    arrHeader = Array("투입", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
        
    Call PrintArrayWithLine(gWsDashboard, startRow, 1, arrHeader)        ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, doingHR_H_Table)      ' 상급자 투입 테이블
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doingHR_M_Table)      ' 중급자 투입 테이블
    Call PrintArrayWithLine(gWsDashboard, startRow + 3, 2, doingHR_L_Table)      ' 초급자 투입 테이블
    
    'song ==> 일단은 doingHR_H_Table로 넣어 놓음. 추후 필요시 여유 테이브로, 필요없으면 삭제
    startRow = startRow + 5
    arrHeader = Array("여유", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(gWsDashboard, startRow, 1, arrHeader)        ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, doingHR_H_Table)      ' 상급자 여유 테이블
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doingHR_M_Table)      ' 중급자 여유 테이블
    Call PrintArrayWithLine(gWsDashboard, startRow + 3, 2, doingHR_L_Table)      ' 초급자 여유 테이블
    
    
    startRow = startRow + 5
    arrHeader = Array("총원", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(gWsDashboard, startRow, 1, arrHeader)        ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, doingHR_H_Table)      ' 상급자 총원 테이블
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doingHR_M_Table)      ' 중급자 총원 테이블
    Call PrintArrayWithLine(gWsDashboard, startRow + 3, 2, doingHR_L_Table)      ' 초급자 총원 테이블
    
    
    'song 제목을 적자
    arrHeader = Array("주", "누계", "prjNum")
    arrHeader = PivotArray(arrHeader)
    
    startRow = DONG_TABLE_INDEX
    
    ' 진행중인 프로젝트
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '기간을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doingTable)      ' 내용을 적는다.

    ' 완료된 프로젝트
    startRow = startRow + doingTableSize + 2
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '기간을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, doneTable)       ' 내용을 적는다.

    ' 미루어진 프로젝트
    startRow = startRow + doneTableSize + 2
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1, arrHeader)       ' 세로항목을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2, gPrintDurationTable) '기간을 적고
    Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2, defferTable)     ' 내용을 적는다.


    Exit Function

    
End Function

