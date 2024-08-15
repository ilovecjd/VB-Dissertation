VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1


' 프로젝트 속성
Public projectType              As Integer  ' 프로젝트 타입 (0: 외부 / 1: 내부)
Public projectNum               As Integer  ' 프로젝트의 번호

Public orderDate                As Integer  ' 발주일
Public possiblestartDate        As Integer  ' 시작 가능일
Public endDate                  As Integer  ' 프로젝트 종료일
Public projectDuration          As Integer  ' 프로젝트의 총 기간
Public isStart                  As Integer  ' 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
Public profit                   As Double   ' 총 기대 수익 (HR 종속)
Public experience               As Integer  ' 경험 (0: 무경험 1: 유경험)
Public successProbability       As Integer  ' 성공 확률

' 현금 흐름
Public numCashFlows             As Integer  ' 비용 지급 횟수
Private m_cashFlows(MAX_N_CF)   As Integer  ' 용역비를 받는 비율을 기록하는 배열
Public firstPayment             As Long     ' 선금 액수
Public middlePayment            As Long     ' 2차 지급 액수
Public finalPayment             As Long     ' 3차 지급 액수
Public firstPaymentMonth        As Integer  ' 선금 지급일
Public middlePaymentMonth       As Integer  ' 2차 지급일
Public finalPaymentMonth        As Integer  ' 3차 지급일`

' 활동
Public numActivities            As Integer ' 총 활동 수
Private m_activities(MAX_ACT)   As Activity_  ' 활동에 관한 정보를 기록하는 배열
Private activityAttribute()     As Variant ' activity_struct 시트에서 프로젝트 속성을 읽어올 배열
Private activityPattern()       As Variant ' activity_struct 시트에서 패턴 속성을 읽어올 배열
'


' 활동의 ActivityType 반환하는 함수
'Public Function SetActivity(index As Integer, activity As Activity_)
'
'    m_activities(index).activityType = activityType
'    m_activities(index).duration = activity.duration
'    m_activities(index).startDate = activity.startDate
'    m_activities(index).endDate = activity.endDate
'    m_activities(index).highSkill = activity.highSkill
'    m_activities(index).midSkill = activity.midSkill
'    m_activities(index).lowSkill = activity.lowSkill
'
'End Function



Public Function GetActivityActivityType(index As Integer) As Integer
    GetActivityActivityType = m_activities(index).activityType
End Function

' 활동의 Duration을 반환하는 함수
Public Function GetActivityDuration(index As Integer) As Integer
    GetActivityDuration = m_activities(index).duration
End Function

' 활동의 startDate를 반환하는 함수
Public Function GetActivitystartDate(index As Integer) As Integer
    GetActivitystartDate = m_activities(index).startDate
End Function

' 활동의 EndDate를 반환하는 함수
Public Function GetActivityEndDate(index As Integer) As Integer
    GetActivityEndDate = m_activities(index).endDate
End Function

' 활동의 HighSkill을 반환하는 함수
Public Function GetActivityHighSkill(index As Integer) As Integer
    GetActivityHighSkill = m_activities(index).highSkill
End Function

' 활동의 MidSkill을 반환하는 함수
Public Function GetActivityMidSkill(index As Integer) As Integer
    GetActivityMidSkill = m_activities(index).midSkill
End Function

' 활동의 LowSkill을 반환하는 함수
Public Function GetActivityLowSkill(index As Integer) As Integer
    GetActivityLowSkill = m_activities(index).lowSkill
End Function


' clsProject.cls

' 활동의 ActivityType 설정
Public Sub SetActivityActivityType(index As Integer, value As Integer)
    m_activities(index).activityType = value
End Sub
 
' 활동의 Duration 설정
Public Sub SetActivityDuration(index As Integer, value As Integer)
    m_activities(index).duration = value
End Sub

' 활동의 startDate 설정
Public Sub SetActivitystartDate(index As Integer, value As Integer)
    m_activities(index).startDate = value
End Sub

' 활동의 EndDate 설정
Public Sub SetActivityEndDate(index As Integer, value As Integer)
    m_activities(index).endDate = value
End Sub

' 활동의 HighSkill 설정
Public Sub SetActivityHighSkill(index As Integer, value As Integer)
    m_activities(index).highSkill = value
End Sub

' 활동의 MidSkill 설정
Public Sub SetActivityMidSkill(index As Integer, value As Integer)
    m_activities(index).midSkill = value
End Sub

' 활동의 LowSkill 설정
Public Sub SetActivityLowSkill(index As Integer, value As Integer)
    m_activities(index).lowSkill = value
End Sub



Public Function CopyToCashFlows(value() As Integer)

    Dim index As Integer
    
    For index = 1 To MAX_N_CF
        m_cashFlows(index) = value(index)
    Next
    
End Function


Public Function CopyFromCashFlows(value() As Integer)

    Dim index As Integer
    
    For index = 1 To MAX_N_CF
        value(index) = m_cashFlows(index)
    Next
    
End Function


'Public Function CopyToActivities(actNum As Integer, value As Variant)
'
'    m_activities(actNum).activityType = value.activityType
'
'    m_activities(actNum).duration = value.duration
'    m_activities(actNum).startDate = value.startDate
'    m_activities(actNum).endDate = value.endDate
'
'    m_activities(actNum).highSkill = value.highSkill
'    m_activities(actNum).midSkill = value.midSkill
'    m_activities(actNum).lowSkill = value.lowSkill
'
'End Function

'
'Public Function CopyFromActivities(actNum As Integer, value As Activity_)
'
'    value.activityType = m_activities(actNum).activityType
'
'    value.duration = m_activities(actNum).duration
'    value.startDate = m_activities(actNum).startDate
'    value.endDate = m_activities(actNum).endDate
'
'    value.highSkill = m_activities(actNum).highSkill
'    value.midSkill = m_activities(actNum).midSkill
'    value.lowSkill = m_activities(actNum).lowSkill
'
'End Function


' Desc     : Project 클래스의 초기화
' initType : 프로젝트 타입 (0: 외부 / 1: 내부)
' initNum  : 프로젝트의 관리 번호, 프로젝트마다 유니크함
' ODate    : 프로젝트가 발생한 달
Public Sub Init(initType As Integer, initNum As Integer, ODate As Integer)

    ' 전달받은 값들로 초기화
    projectType = initType
    projectNum = initNum

    ' 프로젝트 생성을 위해 필요한 값들을 생성한다.
    orderDate = ODate ' 발주일
    possiblestartDate = ODate + Int(Rnd() * 4) ' 시작 가능일
    isStart = 0 ' 시작여부 (0: 미시작) ' song 변수명 변경하자
    'profit                 = 100   ' 총 기대 수익 (HR 종속 MakeProjectHRandprofit 함수에서 생성)
    experience = GenerateWithProbability(95) ' 경험 (0: 무경험 1: 유경험)
    successProbability = 100 '   song 수주 성공 확률
    numCashFlows = MAX_N_CF ' 비용 지급 횟수 (규모에 따라 변경 가능)

    Call CreateActivities
    Call CalculateHRandprofit
    Call CalculatePaymentSchedule
    
End Sub

' 프로젝트와 패턴을 배열에 담아 온다.
Public Sub SetProjectEnv()
    ReDim activityAttribute(1 To 5, 1 To 14)
    ReDim activityPattern(1 To 6, 1 To 27)

    With g_WsActivity_Struct
        activityAttribute = .Range(.Cells(3, 1), .Cells(7, 14)).value
        activityPattern = .Range(.Cells(15, 1), .Cells(20, 27)).value
    End With
End Sub

' 활동 생성, 활동 선후 관계, 프로젝트 전체 기간 갱신
Public Sub CreateActivities()
    Randomize

    Call SetProjectEnv

    Dim prjType As Integer
    Dim patternType As Integer
    Dim index As Integer
    Dim probability As Integer
    Dim Lb As Integer
    Dim UB As Integer
    Dim maxLoop As Integer
    Dim totalDuration As Integer
    Dim tempDuration As Integer
    
    prjType = 0
    Lb = 0
    UB = 0

    probability = Int(Rnd() * 100)
    maxLoop = MAX_PRJ_TYPE

    For index = 1 To maxLoop
        UB = UB + activityAttribute(index, 1)

        If Lb <= probability And probability < UB Then
            prjType = index
            Exit For
        End If

        Lb = UB
    Next index

    Lb = activityAttribute(prjType, 4)
    UB = activityAttribute(prjType, 5)
    totalDuration = RandomBetween(Lb, UB)
    projectDuration = totalDuration
    endDate = possiblestartDate + totalDuration - 1

    Lb = 0
    UB = 0
    maxLoop = activityAttribute(prjType, 6)

    For index = 1 To maxLoop
        UB = UB + activityAttribute(prjType, 8 + ((index - 1) * 2))

        If Lb <= probability And probability < UB Then
            patternType = activityAttribute(prjType, 7 + (index - 1) * 2)
            Exit For
        End If

        Lb = UB
    Next index

    Lb = 0
    UB = 0
    maxLoop = activityPattern(patternType, 2)
    numActivities = maxLoop

    For index = 1 To maxLoop
        Lb = Lb + activityPattern(patternType, 3 + (index - 1) * 5)
        UB = UB + activityPattern(patternType, 4 + (index - 1) * 5)
        probability = RandomBetween(Lb, UB)
        tempDuration = totalDuration * probability / 100

        If tempDuration = 0 Then
            tempDuration = 1
        End If

        If index = 1 Then
            m_activities(index).duration = tempDuration
            m_activities(index).startDate = possiblestartDate '1
            m_activities(index).endDate = possiblestartDate - 1 + tempDuration
        ElseIf index = 2 Then
            m_activities(index).duration = totalDuration - m_activities(1).duration
            m_activities(index).startDate = m_activities(1).endDate + 1
            m_activities(index).endDate = possiblestartDate - 1 + totalDuration
        ElseIf index = 3 Then
            m_activities(index).duration = tempDuration '3
            m_activities(index).startDate = possiblestartDate - 1 + totalDuration - tempDuration + 1
            m_activities(index).endDate = possiblestartDate - 1 + totalDuration
        Else
            m_activities(index).duration = tempDuration
            m_activities(index).startDate = m_activities(3).startDate - tempDuration
            m_activities(index).endDate = m_activities(3).startDate - 1
        End If
    Next index
End Sub

Public Function RandomBetween(minValue As Integer, maxValue As Integer) As Integer
    Dim lowerBound As Integer
    Dim upperBound As Integer
    
    Randomize
    
    lowerBound = Application.WorksheetFunction.Min(minValue, maxValue)
    upperBound = Application.WorksheetFunction.Max(minValue, maxValue)
    
    RandomBetween = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
End Function

' 활동별 투입 인력 생성, 프로젝트 전체 기대 수익 생성 및 기록
Public Sub CalculateHRandprofit()
    Dim index As Integer
    Dim j As Integer
    Dim high As Integer, mid As Integer, low As Integer
    high = 0: mid = 0: low = 0

    For index = 1 To numActivities
        j = Int(Rnd() * 100)
        If (0 < j) And (j <= RND_HR_H) Then
            m_activities(index).highSkill = 1
        ElseIf (RND_HR_H < j) And (j <= RND_HR_M) Then
            m_activities(index).midSkill = 1
        Else
            m_activities(index).lowSkill = 1
        End If
    Next index

    For index = 1 To numActivities
        high = high + m_activities(index).highSkill * m_activities(index).duration
        mid = mid + m_activities(index).midSkill * m_activities(index).duration
        low = low + m_activities(index).lowSkill * m_activities(index).duration
    Next index

    profit = CalculateTotalLaborCost(high, mid, low)
End Sub

' 등급별 투입인력 계산 및 프로젝트의 수익 생성
Public Function CalculateTotalLaborCost(highCount As Integer, midCount As Integer, lowCount As Integer) As Integer
    Dim highLaborCost As Integer
    Dim midLaborCost As Integer
    Dim lowLaborCost As Integer

    highLaborCost = CalculateLaborCost("H") * highCount
    midLaborCost = CalculateLaborCost("M") * midCount
    lowLaborCost = CalculateLaborCost("L") * lowCount

    CalculateTotalLaborCost = highLaborCost + midLaborCost + lowLaborCost
End Function

' 등급별 투입인력에 따른 수익 계산
Private Function CalculateLaborCost(grade As String) As Integer
    Dim directLaborCost As Integer
    Dim overheadCost As Integer
    Dim technicalFee As Integer
    Dim totalLaborCost As Integer

    grade = UCase(grade)

    Select Case grade
        Case "H"
            directLaborCost = 50
        Case "M"
            directLaborCost = 39
        Case "L"
            directLaborCost = 25
        Case Else
            MsgBox "잘못된 등급입니다. 'H', 'M', 'L' 중 하나를 입력하세요."
            Exit Function
    End Select

    overheadCost = directLaborCost * 0.6
    technicalFee = (directLaborCost + overheadCost) * 0.2
    totalLaborCost = directLaborCost + overheadCost + technicalFee

    CalculateLaborCost = totalLaborCost
End Function

' 대금 지급 조건 생성
Public Sub CalculatePaymentSchedule()
    Dim paymentType As Integer
    Dim paymentRatio As Integer
    Dim totalPayments As Integer

    firstPaymentMonth = 1

    Randomize

    If projectDuration < 6 Then
        paymentType = Int(3 * Rnd + 1)

        Select Case paymentType
            Case 1
                firstPayment = profit * 0.3: m_cashFlows(1) = 30: m_cashFlows(2) = 70
            Case 2
                firstPayment = profit * 0.4: m_cashFlows(1) = 40: m_cashFlows(2) = 60
            Case 3
                firstPayment = profit * 0.5: m_cashFlows(1) = 50: m_cashFlows(2) = 50
        End Select

        middlePayment = profit - firstPayment
        totalPayments = 2
        middlePaymentMonth = projectDuration
    ElseIf projectDuration < 12 Then
        paymentType = Int(10 * Rnd + 1)

        If paymentType <= 3 Then
            paymentRatio = Int(3 * Rnd + 1)
            Select Case paymentRatio
                Case 1
                    firstPayment = profit * 0.3: m_cashFlows(1) = 30: m_cashFlows(2) = 70
                Case 2
                    firstPayment = profit * 0.4: m_cashFlows(1) = 40: m_cashFlows(2) = 60
                Case 3
                    firstPayment = profit * 0.5: m_cashFlows(1) = 50: m_cashFlows(2) = 50
            End Select
            middlePayment = profit - firstPayment
            totalPayments = 2
            middlePaymentMonth = projectDuration
        Else
            paymentRatio = Int(10 * Rnd + 1)
            If paymentRatio <= 6 Then
                firstPayment = profit * 0.3
                middlePayment = profit * 0.3
                m_cashFlows(1) = 30
                m_cashFlows(2) = 30
                m_cashFlows(3) = 40
            Else
                firstPayment = profit * 0.3
                middlePayment = profit * 0.4
                m_cashFlows(1) = 30
                m_cashFlows(2) = 40
                m_cashFlows(3) = 30
            End If

            finalPayment = profit - firstPayment - middlePayment
            totalPayments = 3
            middlePaymentMonth = WorksheetFunction.Ceiling(projectDuration / 2, 1)
            finalPaymentMonth = projectDuration
        End If
    Else
        firstPayment = profit * 0.3
        middlePayment = profit * 0.4
        finalPayment = profit - firstPayment - middlePayment

        m_cashFlows(1) = 30
        m_cashFlows(2) = 40
        m_cashFlows(3) = 30

        totalPayments = 3
        middlePaymentMonth = WorksheetFunction.Ceiling(projectDuration / 2, 1)
        finalPaymentMonth = projectDuration
    End If

    numCashFlows = totalPayments
End Sub

' 주어진 확률에 따라서 1 또는 0을 반환하는 함수
Public Function GenerateWithProbability(probability As Integer) As Integer
    Dim randomProb As Double
    Randomize
    randomProb = Rnd
    If randomProb <= probability / 100 Then
        GenerateWithProbability = 1
    Else
        GenerateWithProbability = 0
    End If
End Function


' 프로젝트의 모든 정보를 엑셀에 출력한다.
' song 출력속도 개선 필요함. 현재는 하나의 프로젝트 정보를 모아서 시트의 특정 부분에 출력하고있다.
Public Sub PrintProjectInfomation()

    Dim index As Integer
    Dim posX As Integer
    Dim posY As Integer

    Dim projectInfo() As Variant
    ReDim projectInfo(1 To PRJ_SHEET_HEADER_H, 1 To PRJ_SHEET_HEADER_W)
    
    posX = 1: posY = 1: projectInfo(posY, posX) = projectType
    posX = posX + 1: projectInfo(posY, posX) = projectNum
    posX = posX + 1: projectInfo(posY, posX) = projectDuration
    posX = posX + 1: projectInfo(posY, posX) = possiblestartDate
    posX = posX + 1: projectInfo(posY, posX) = endDate
    posX = posX + 1: projectInfo(posY, posX) = orderDate
    posX = posX + 1: projectInfo(posY, posX) = profit
    posX = posX + 1: projectInfo(posY, posX) = experience
    posX = posX + 1: projectInfo(posY, posX) = successProbability
    posX = posX + 1: projectInfo(posY, posX) = numCashFlows
    
    For index = 1 To MAX_N_CF
        posX = posX + 1: projectInfo(posY, posX) = m_cashFlows(index)
    Next index

    posX = posX + 1: projectInfo(posY, posX) = firstPayment
    posX = posX + 1: projectInfo(posY, posX) = middlePayment
    posX = posX + 1: projectInfo(posY, posX) = finalPayment


    posX = 1: posY = 2: projectInfo(posY, posX) = numActivities
    posX = posX + 10: projectInfo(posY, posX) = firstPaymentMonth
    posX = posX + 1: projectInfo(posY, posX) = middlePaymentMonth
    posX = posX + 1: projectInfo(posY, posX) = finalPaymentMonth
    
'    posX = 1: posY = 2: projectInfo(posY, posX) = "ActivityCnt"
'    posX = posX + 1: projectInfo(posY, posX) = numActivities
'    posX = posX + 9: projectInfo(posY, posX) = firstPaymentMonth
'    posX = posX + 1: projectInfo(posY, posX) = middlePaymentMonth
'    posX = posX + 1: projectInfo(posY, posX) = finalPaymentMonth

    
    For index = 1 To numActivities
        Dim strAct As String, strNum As String
        strNum = Format(index, "#0")
        strAct = "Activity" & strNum

        posX = 2:        projectInfo(posY, posX) = strAct
        posX = posX + 1: projectInfo(posY, posX) = m_activities(index).duration
        posX = posX + 1: projectInfo(posY, posX) = m_activities(index).startDate
        posX = posX + 1: projectInfo(posY, posX) = m_activities(index).endDate
        
        posX = posX + 2: projectInfo(posY, posX) = m_activities(index).highSkill
        posX = posX + 1: projectInfo(posY, posX) = m_activities(index).midSkill
        posX = posX + 1: projectInfo(posY, posX) = m_activities(index).lowSkill
        
        posY = posY + 1
    Next index

    posY = (projectNum - 1) * PRJ_SHEET_HEADER_H + 4

    With g_WsProject
    
        Dim RNG1 As Object
        Set RNG1 = .Range(.Cells(posY, 1), .Cells(posY + PRJ_SHEET_HEADER_H - 1, PRJ_SHEET_HEADER_W))
        
        RNG1.UnMerge
        RNG1.Clear
        RNG1.ClearContents
        RNG1.value = projectInfo
        RNG1.BorderAround LineStyle:=1, Weight:=xlThin, ColorIndex:=1
        
    End With
End Sub

