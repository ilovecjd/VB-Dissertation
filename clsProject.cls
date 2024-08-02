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

' 출력에 필요한 환경 변수들
Private ProjectInfo()       As Variant ' 출력할 헤더를 기록하는 배열


' 프로젝트 속성
Public ProjectType As Integer ' 프로젝트 타입 (0: 외부 / 1: 내부)
Public ProjectNum As Integer ' 프로젝트의 번호

Public OrderDate As Integer ' 발주일
Public PossibleStartDate As Integer ' 시작 가능일
Public ProjectDuration As Integer ' 프로젝트의 총 기간
Public StartDate As Integer ' 시작일 (0: 미시작)
Public Profit As Double ' 총 기대 수익 (HR 종속)
Public Experience As Integer ' 경험 (0: 무경험 1: 유경험)
Public SuccessProbability As Integer ' 성공 확률

' 현금 흐름
Public NumCashFlows As Integer ' 비용 지급 횟수
Private CashFlows() As Integer ' 용역비를 받는 비율을 기록하는 배열
Public FirstPayment As Long ' 선금 액수
Public MiddlePayment As Long ' 2차 지급 액수
Public FinalPayment As Long ' 3차 지급 액수
Public FirstPaymentMonth As Integer ' 선금 지급 월
Public MiddlePaymentMonth As Integer ' 2차 지급 월
Public FinalPaymentMonth As Integer ' 3차 지급 월

' 활동
Public NumActivities As Integer ' 총 활동 수
Private Activities() As Activity ' 활동에 관한 정보를 기록하는 배열
Private PrjAttribute() As Variant ' activity_struct 시트에서 프로젝트 속성을 읽어올 배열
Private PatAttribute() As Variant ' activity_struct 시트에서 패턴 속성을 읽어올 배열


Public Property Get PrjCashFlows() As Integer()
    PrjCashFlows = CashFlows
End Property

Public Property Get PrjActivities() As Integer()
    'PrjActivities = Activities
End Property

'Public Sub SetPrjActivities(i As Integer, Value As Activity)
 '   Activities(i) = Value
'End Sub

Public Sub SetPrjCashFlows(value() As Integer)
    CashFlows = value
End Sub

Private Sub Class_Initialize()
    
    ReDim ProjectInfo(1 To PRJ_SHEET_HEADER_H, 1 To PRJ_SHEET_HEADER_W)
    ReDim CashFlows(1 To MAX_N_CF)
    ReDim Activities(1 To MAX_ACT)
End Sub

' Desc     : Project 클래스의 초기화
' initType : 프로젝트 타입 (0: 외부 / 1: 내부)
' initNum  : 프로젝트의 관리 번호, 프로젝트마다 유니크함
' ODate    : 프로젝트가 발생한 달
Public Sub Init(initType As Integer, initNum As Integer, ODate As Integer)

    ' 전달받은 값들로 초기화
    ProjectType = initType
    ProjectNum = initNum

    ' 프로젝트 생성을 위해 필요한 값들을 생성한다.
    OrderDate = ODate ' 발주월
    PossibleStartDate = ODate + Int(Rnd() * 4) ' 시작 가능일
    StartDate = 0 ' 시작일 (0: 미시작)
    'Profit                 = 100   ' 총 기대 수익 (HR 종속 MakeProjectHRandProfit 함수에서 생성)
    Experience = GenerateWithProbability(95) ' 경험 (0: 무경험 1: 유경험)
    SuccessProbability = 100 '   수주 성공 확률
    NumCashFlows = MAX_N_CF ' 비용 지급 횟수 (규모에 따라 변경 가능)

    Call CreateActivities
    Call CalculateHRandProfit
    Call CalculatePaymentSchedule
    
End Sub

' 프로젝트와 패턴을 배열에 담아 온다.
Public Sub SetProjectEnv()
    ReDim PrjAttribute(1 To 5, 1 To 14)
    ReDim PatAttribute(1 To 6, 1 To 27)

    With gWsActivity_Struct
        PrjAttribute = .Range(.Cells(3, 1), .Cells(7, 14)).value
        PatAttribute = .Range(.Cells(15, 1), .Cells(20, 27)).value
    End With
End Sub

' 활동 생성, 활동 선후 관계, 프로젝트 전체 기간 갱신
Public Sub CreateActivities()
    Randomize

    Call SetProjectEnv

    Dim prjType As Integer
    Dim patternType As Integer
    Dim I As Integer
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

    For I = 1 To maxLoop
        UB = UB + PrjAttribute(I, 1)

        If Lb <= probability And probability < UB Then
            prjType = I
            Exit For
        End If

        Lb = UB
    Next I

    Lb = PrjAttribute(prjType, 4)
    UB = PrjAttribute(prjType, 5)
    totalDuration = RandomBetween(Lb, UB)
    ProjectDuration = totalDuration

    Lb = 0
    UB = 0
    maxLoop = PrjAttribute(prjType, 6)

    For I = 1 To maxLoop
        UB = UB + PrjAttribute(prjType, 8 + ((I - 1) * 2))

        If Lb <= probability And probability < UB Then
            patternType = PrjAttribute(prjType, 7 + (I - 1) * 2)
            Exit For
        End If

        Lb = UB
    Next I

    Lb = 0
    UB = 0
    maxLoop = PatAttribute(patternType, 2)
    NumActivities = maxLoop

    For I = 1 To maxLoop
        Lb = Lb + PatAttribute(patternType, 3 + (I - 1) * 5)
        UB = UB + PatAttribute(patternType, 4 + (I - 1) * 5)
        probability = RandomBetween(Lb, UB)
        tempDuration = totalDuration * probability / 100

        If tempDuration = 0 Then
            tempDuration = 1
        End If

        If I = 1 Then
            Activities(I).Duration = tempDuration
            Activities(I).StartDate = 1
            Activities(I).EndDate = tempDuration
        ElseIf I = 2 Then
            Activities(I).Duration = totalDuration - Activities(1).Duration
            Activities(I).StartDate = Activities(1).EndDate + 1
            Activities(I).EndDate = totalDuration
        ElseIf I = 3 Then
            Activities(I).Duration = tempDuration
            Activities(I).StartDate = totalDuration - tempDuration
            Activities(I).EndDate = totalDuration
        Else
            Activities(I).Duration = tempDuration
            Activities(I).StartDate = Activities(3).StartDate - tempDuration - 1
            Activities(I).EndDate = Activities(3).StartDate - 1
        End If
    Next I
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
Public Sub CalculateHRandProfit()
    Dim I As Integer
    Dim j As Integer
    Dim high As Integer, mid As Integer, low As Integer
    high = 0: mid = 0: low = 0

    For I = 1 To NumActivities
        j = Int(Rnd() * 100)
        If (0 < j) And (j <= RND_HR_H) Then
            Activities(I).HighSkill = 1
        ElseIf (RND_HR_H < j) And (j <= RND_HR_M) Then
            Activities(I).MidSkill = 1
        Else
            Activities(I).LowSkill = 1
        End If
    Next I

    For I = 1 To NumActivities
        high = high + Activities(I).HighSkill * Activities(I).Duration
        mid = mid + Activities(I).MidSkill * Activities(I).Duration
        low = low + Activities(I).LowSkill * Activities(I).Duration
    Next I

    Profit = CalculateTotalLaborCost(high, mid, low)
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

    FirstPaymentMonth = 1

    Randomize

    If ProjectDuration < 6 Then
        paymentType = Int(3 * Rnd + 1)

        Select Case paymentType
            Case 1
                FirstPayment = Profit * 0.3: CashFlows(1) = 30: CashFlows(2) = 70
            Case 2
                FirstPayment = Profit * 0.4: CashFlows(1) = 40: CashFlows(2) = 60
            Case 3
                FirstPayment = Profit * 0.5: CashFlows(1) = 50: CashFlows(2) = 50
        End Select

        MiddlePayment = Profit - FirstPayment
        totalPayments = 2
        MiddlePaymentMonth = ProjectDuration
    ElseIf ProjectDuration < 12 Then
        paymentType = Int(10 * Rnd + 1)

        If paymentType <= 3 Then
            paymentRatio = Int(3 * Rnd + 1)
            Select Case paymentRatio
                Case 1
                    FirstPayment = Profit * 0.3: CashFlows(1) = 30: CashFlows(2) = 70
                Case 2
                    FirstPayment = Profit * 0.4: CashFlows(1) = 40: CashFlows(2) = 60
                Case 3
                    FirstPayment = Profit * 0.5: CashFlows(1) = 50: CashFlows(2) = 50
            End Select
            MiddlePayment = Profit - FirstPayment
            totalPayments = 2
            MiddlePaymentMonth = ProjectDuration
        Else
            paymentRatio = Int(10 * Rnd + 1)
            If paymentRatio <= 6 Then
                FirstPayment = Profit * 0.3
                MiddlePayment = Profit * 0.3
                CashFlows(1) = 30
                CashFlows(2) = 30
                CashFlows(3) = 40
            Else
                FirstPayment = Profit * 0.3
                MiddlePayment = Profit * 0.4
                CashFlows(1) = 30
                CashFlows(2) = 40
                CashFlows(3) = 30
            End If

            FinalPayment = Profit - FirstPayment - MiddlePayment
            totalPayments = 3
            MiddlePaymentMonth = WorksheetFunction.Ceiling(ProjectDuration / 2, 1)
            FinalPaymentMonth = ProjectDuration
        End If
    Else
        FirstPayment = Profit * 0.3
        MiddlePayment = Profit * 0.4
        FinalPayment = Profit - FirstPayment - MiddlePayment

        CashFlows(1) = 30
        CashFlows(2) = 40
        CashFlows(3) = 30

        totalPayments = 3
        MiddlePaymentMonth = WorksheetFunction.Ceiling(ProjectDuration / 2, 1)
        FinalPaymentMonth = ProjectDuration
    End If

    NumCashFlows = totalPayments
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
Public Sub PrintInfo()
    Dim I As Integer
    Dim posX As Integer
    Dim posY As Integer

    posX = 1: posY = 1: ProjectInfo(posY, posX) = ProjectType
    posX = posX + 1: ProjectInfo(posY, posX) = ProjectNum
    posX = posX + 1: ProjectInfo(posY, posX) = OrderDate
    posX = posX + 1: ProjectInfo(posY, posX) = PossibleStartDate
    posX = posX + 1: ProjectInfo(posY, posX) = ProjectDuration
    posX = posX + 1: ProjectInfo(posY, posX) = StartDate
    posX = posX + 1: ProjectInfo(posY, posX) = Profit
    posX = posX + 1: ProjectInfo(posY, posX) = Experience
    posX = posX + 1: ProjectInfo(posY, posX) = SuccessProbability
    posX = posX + 1: ProjectInfo(posY, posX) = NumCashFlows

    For I = 1 To MAX_N_CF
        posX = posX + 1: ProjectInfo(posY, posX) = CashFlows(I)
    Next I

    posX = posX + 1: ProjectInfo(posY, posX) = FirstPayment
    posX = posX + 1: ProjectInfo(posY, posX) = MiddlePayment
    posX = posX + 1: ProjectInfo(posY, posX) = FinalPayment

    posX = 1: posY = 2: ProjectInfo(posY, posX) = "ActivityCnt"
    posX = posX + 1: ProjectInfo(posY, posX) = NumActivities
    posX = posX + 9: ProjectInfo(posY, posX) = FirstPaymentMonth
    posX = posX + 1: ProjectInfo(posY, posX) = MiddlePaymentMonth
    posX = posX + 1: ProjectInfo(posY, posX) = FinalPaymentMonth

    For I = 1 To MAX_ACT
        Dim strAct As String, strNum As String
        strNum = Format(I, "#0")
        strAct = "Activity" & strNum

        posX = 1: posY = posY + 1: ProjectInfo(posY, posX) = strAct
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(I).Duration
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(I).StartDate
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(I).EndDate
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(I).HighSkill
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(I).MidSkill
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(I).LowSkill
    Next I

    posY = (ProjectNum - 1) * PRJ_SHEET_HEADER_H + 4

    With gWsProject
        Dim RNG1 As Object
        Set RNG1 = .Range(.Cells(posY, 1), .Cells(posY + PRJ_SHEET_HEADER_H - 1, PRJ_SHEET_HEADER_W))
        RNG1.value = ProjectInfo
        RNG1.BorderAround LineStyle:=1, Weight:=xlThin, ColorIndex:=1
    End With
End Sub

