Attribute VB_Name = "typedef"
Option Explicit
Option Base 1

'WorkBook ��ü�� ���������� ���� ����ü
Type EnvExcel
        SimulationDuration              As Integer  ' �ùķ��̼��� ���� ��ų �Ⱓ(��)
        avgProjects                             As Double       ' �ִ� �߻��ϴ� ��� ���� ������Ʈ ��
        Hr_Init_H                       As Integer  ' ���ʿ� ������ ��� �η�
        Hr_Init_M                       As Integer  ' ���ʿ� ������ �߱� �η�
        Hr_Init_L                       As Integer  ' ���ʿ� ������ �ʱ� �η�
        Hr_LeadTime                     As Integer  ' �η� ����� �ɸ��� �ð�
        Cash_Init                       As Integer  ' ���� ���� ����
        Problem                         As Integer  ' ������Ʈ ���� ���� (= ������ ����) / MakePrj �Լ��� ����
End Type

' Ȱ���� ������ ��� ����ü
Type Activity
    ActivityType    As Integer  ' 1-�м�����/2-����/3-����/4-����/5-��������
    Duration        As Integer  ' Ȱ���� �Ⱓ
    StartDate       As Integer  ' Ȱ���� ����
    EndDate         As Integer  ' Ȱ���� ��
    HighSkill       As Integer  ' �ʿ��� ��� �η� ��
    MidSkill        As Integer  ' �ʿ��� �߱� �η� ��
    LowSkill        As Integer  ' �ʿ��� �ʱ� �η� ��
End Type
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Define Global Variable
' Const Start
' sheet name
Public Const PARAMETER_SHEET_NAME = "GenDBoard"
Public Const DBOARD_SHEET_NAME = "dashboard"
Public Const PROJECT_SHEET_NAME = "project"
Public Const ACTIVITY_SHEET_NAME = "activity_struct"

' �ֿ� ���̺��� ����
Public Const ORDER_PROJECT_TITLE = "���� ������Ʈ ��Ȳ"

Public Const P_TYPE_EXTERNAL = 0  ' �ܺ�(����)������Ʈ
Public Const P_TYPE_INTERNAL = 1  ' ���� ������Ʈ

''''''''''''''''''''
' ������Ʈ ������ ���õ� �����
Public Const MAX_ACT            As Integer = 4           ' �ִ� Ȱ���� ��
Public Const MAX_N_CF           As Integer = 3   ' �ִ� CF�� ���� (���ߺ� �ִ�� ������ �޴� Ƚ��)
Public Const W_INFO                     As Integer = 16      ' ����� ������ ũ��
Public Const H_INFO             As Integer = 8       ' ����� ������ ũ��

Public Const RND_HR_H = 20      ' ��� �η��� �ʿ��� Ȯ��
Public Const RND_HR_M = 70      ' �߱� �η��� �ʿ��� Ȯ��

' 1: 2~4 / 2:5~12 3:13~26 4:27~52 5:53~80
Public Const MAX_PRJ_TYPE       As Integer = 5                  ' ������Ʈ �Ⱓ���� Ÿ���� �����Ѵ�.
Public Const RND_PRJ_TYPE1      As Integer = 20         ' 1�� Ÿ���� Ȯ�� 1:  2~4 ��
Public Const RND_PRJ_TYPE2      As Integer = 70         ' 2�� Ÿ���� Ȯ�� 2:  5~12��
Public Const RND_PRJ_TYPE3      As Integer = 20         ' 3�� Ÿ���� Ȯ�� 3: 13~26��
Public Const RND_PRJ_TYPE4      As Integer = 70         ' 4�� Ÿ���� Ȯ�� 4: 27~52��
Public Const RND_PRJ_TYPE5      As Integer = 20         ' 5�� Ÿ���� Ȯ�� 5: 53~80��

''''''''''''''''''''
'' ��°� �ε带 ���� �����
Public Const ORDER_TABLE_INDEX                  As Long = 1             '
Public Const DONG_TABLE_INDEX = 6                               '
Public Const PROJECT_TABLE_INDEX        As Long = 3             '
' Const End

Private gExcelInitialized       As Boolean      ' ���� �������� �ʱ�ȭ �Ǿ����� Ȯ���ϴ� �÷���. �ʱ�ȭ �Ǹ� 1
Private gTableInitialized       As Boolean      ' ���� ���̺��� �ʱ�ȭ �Ǿ����� Ȯ���ϴ� �÷���. �ʱ�ȭ �Ǹ� 1
Public gTotalProjectNum         As Integer       ' �߻��� ������Ʈ�� �� ���� (����)

Public gWsGenDBoard                     As excel.Worksheet    ' ��ũ��Ʈ���� �������� �̸� ���� ���´�.
Public gWsDashboard                     As Worksheet
Public gWsProject                       As Worksheet
Public gWsActivity_Struct       As Worksheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ' ���α׷� ������ ���� �⺻ ������.
Public gExcelEnv                As EnvExcel
Public gOrderTable()    As Variant              ' ���ֵ� ������Ʈ���� �����ϴ� ���̺�
'Public gProjectTable()  As clsProject   ' ��� ������Ʈ���� ��� �ִ� ���̺�


Public gPrintDurationTable()    As Variant              ' ����ϱ� ���ϰ� ��� ���� �־� ���´�.



' #define end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' Public functions
Public Property Get GetExcelEnv() As EnvExcel
        GetExcelEnv = gExcelEnv
End Property

Public Property Get GetExcelInitialized() As Boolean
        GetExcelInitialized = gExcelInitialized
End Property

Public Property Let LetExcelInitialized(Value As Boolean)
        gExcelInitialized = Value
End Property


Public Property Get GetTableInitialized() As Boolean
        GetTableInitialized = gTableInitialized
End Property

Public Property Let LetTableInitialized(Value As Boolean)
        gTableInitialized = Value
End Property


Public Property Get GetTotalProjectNum() As Integer
        GetTotalProjectNum = gTotalProjectNum
End Property

Public Property Get GetOrderTable() As Variant
        GetOrderTable = gOrderTable
End Property

Public Property Get GetProjectTable() As Variant
        'GetProjectTable = gProjectTable
End Property



' utility functions

' desc      : ���α׷� ������ ���� �⺻���� ������ �����Ѵ�. ��� ���ν������� ���۽� ȣ�� �Ͽ��� �Ѵ�.
' return    : none
Sub Prologue(TableInit As Integer)
On Error GoTo ErrorHandler

        Dim i As Integer
        
        If gExcelInitialized = 0 Then           ' ���� �������� �ʱ�ȭ �Ǿ����� Ȯ���ϴ� �÷���. �ʱ�ȭ �Ǹ� 1
                ' �ѹ��� �ϸ� �Ǵ� �͵��� ���⿡
                Call LoadExcelEnv

                ReDim gPrintDurationTable(1, gExcelEnv.SimulationDuration)
                For i = 1 To (gExcelEnv.SimulationDuration)
                        gPrintDurationTable(1, i) = i
                Next

                gExcelInitialized = 1           ' ���� �������� �ʱ�ȭ �Ǿ����� Ȯ���ϴ� �÷���. �ʱ�ȭ �Ǹ� 1
        End If

        ' ���̺���� ���� �����ϰų� �������� �ε��ϰų�.
        ' ���� ó���� �������� ����ڰ� �����ؼ� ����ϵ��� ����.
        gTableInitialized = TableInit
        If gTableInitialized = 0 Then ' Table ���� ��������� �ʾ����� ���̺� ����
                Call BuildTables
        Else
                Call LoadTablesFromExcel   ' ������� ������ ������ ���� ��Ʈ���� ������ �ε�
        End If

        gTableInitialized = 1  'Prologue()�� ȣ���ϱ����� �����ϰ� ȣ���Ѵ�.

        ' �ӵ� ����� ���ؼ�
        ' Application.ScreenUpdating = False
        ' Application.Calculation = xlCalculationManual
        ' Application.EnableEvents = False
        ' ActiveSheet.DisplayPageBreaks = False

        Exit Sub

ErrorHandler:
    Call HandleError("Prologue", Err.Description)

End Sub

Sub BuildTables()

        Call CreateOrderTable
        Call CreateProjects

End Sub

Sub LoadTablesFromExcel()

        Call LoadOrderTable
        Call LoadProjects

End Sub

Private Function LoadOrderTable() As Boolean

        ReDim gOrderTable(2, gExcelEnv.SimulationDuration)

        Dim startIndex As Long
        startIndex = ORDER_TABLE_INDEX
        startIndex = startIndex + 2

        With gWsDashboard
                gOrderTable = .Range(.Cells(startIndex, 2), .Cells(startIndex + 1, gExcelEnv.SimulationDuration + 1)).Value
        End With

        gTotalProjectNum = gOrderTable(1, gExcelEnv.SimulationDuration) + gOrderTable(2, gExcelEnv.SimulationDuration)

End Function

Private Function LoadProjects() As Boolean

        Dim prjID               As Integer
        Dim startRow    As Long
        Dim endRow              As Long
        Dim prjInfo     As Variant
        Dim iTemp               As Integer '
        Dim tempPrj     As clsProject


                                '������Ʈ���� �����Ѵ�.
        ReDim gProjectTable(gTotalProjectNum)






        For prjID = 1 To gTotalProjectNum

                Set tempPrj = New clsProject
                startRow = PROJECT_TABLE_INDEX + (prjID - 1) * H_INFO + 1
                endRow = startRow + H_INFO - 1

                With gWsProject
                        prjInfo = .Range(.Cells(startRow, 1), .Cells(endRow, W_INFO)).Value
                End With


                Dim i As Integer
                Dim j As Integer
                Dim k As Integer
                Dim tempCF(1 To MAX_N_CF) As Integer

                i = 1: j = 1
                tempPrj.ProjectType = prjInfo(i, j): j = j + 1
                tempPrj.ProjectNum = prjInfo(i, j): j = j + 1
                tempPrj.OrderDate = prjInfo(i, j): j = j + 1
                tempPrj.PossibleStartDate = prjInfo(i, j): j = j + 1
                tempPrj.ProjectDuration = prjInfo(i, j): j = j + 1
                tempPrj.StartDate = prjInfo(i, j): j = j + 1
                tempPrj.Profit = prjInfo(i, j): j = j + 1
                tempPrj.Experience = prjInfo(i, j): j = j + 1
                tempPrj.SuccessProbability = prjInfo(i, j): j = j + 1
                tempPrj.NumCashFlows = MAX_N_CF
                For k = 1 To MAX_N_CF
                        tempCF(k) = prjInfo(i, j): j = j + 1
                Next
                'tempPrj.SetPrjCashFlows        = tempCF
                Call tempPrj.SetPrjCashFlows(tempCF)
                

                tempPrj.FirstPayment = prjInfo(i, j): j = j + 1
                tempPrj.MiddlePayment = prjInfo(i, j): j = j + 1
                tempPrj.FinalPayment = prjInfo(i, j): j = j + 1


                i = 2: j = 2
                tempPrj.NumActivities = prjInfo(i, j)
                
                j = j + 9 ' ����� �� ��������
                tempPrj.FirstPaymentMonth = prjInfo(i, j): j = j + 1
                tempPrj.MiddlePaymentMonth = prjInfo(i, j): j = j + 1
                tempPrj.FinalPaymentMonth = prjInfo(i, j): j = j + 1
                
                Dim tempAct     As Activity
                For i = 3 To (tempPrj.NumActivities + i - 1)
                        j = 2
                        tempAct.Duration = prjInfo(i, j): j = j + 1
                        tempAct.StartDate = prjInfo(i, j): j = j + 1
                        tempAct.EndDate = prjInfo(i, j): j = j + 1
                        tempAct.HighSkill = prjInfo(i, j): j = j + 1
                        tempAct.MidSkill = prjInfo(i, j): j = j + 1
                        tempAct.LowSkill = prjInfo(i, j): j = j + 1
                        'tempPrj.Activities(i-2)        = tempAct
                        Call tempPrj.SetPrjActivities(i - 2, tempAct)
                Next

                Set gProjectTable(prjID) = tempPrj

        Next
        
End Function

Sub LoadExcelEnv() ' ���� ��ũ�� ��ü���� �������� ����ϴ� ȯ�� ���� �ε�

        ' ���� ����ϴ� ��Ʈ�� �������� ������ ����. (�ӵ� ����� ����)
        Set gWsGenDBoard = ThisWorkbook.Sheets(PARAMETER_SHEET_NAME)
        Set gWsDashboard = ThisWorkbook.Sheets(DBOARD_SHEET_NAME)
        Set gWsProject = ThisWorkbook.Sheets(PROJECT_SHEET_NAME)
        Set gWsActivity_Struct = ThisWorkbook.Sheets(ACTIVITY_SHEET_NAME)

        ' ���� ���� ȯ�� �������� �����´�.
        Dim rng         As Range
        Set rng = gWsGenDBoard.Range("b:c")

        gExcelEnv.SimulationDuration = GetVariableValue(rng, "SimulTerm")
        gExcelEnv.avgProjects = GetVariableValue(rng, "avgProjects")
        gExcelEnv.Hr_Init_H = GetVariableValue(rng, "Hr_Init_H")
        gExcelEnv.Hr_Init_M = GetVariableValue(rng, "Hr_Init_M")
        gExcelEnv.Hr_Init_L = GetVariableValue(rng, "Hr_Init_L")
        gExcelEnv.Hr_LeadTime = GetVariableValue(rng, "Hr_LeadTime")
        gExcelEnv.Cash_Init = GetVariableValue(rng, "Cash_Init")
        gExcelEnv.Problem = GetVariableValue(rng, "ProblemCnt")

End Sub


' �Ⱓ������ ��� ���� ������Ʈ�� �̸� ���ؼ� �־���´�.
Private Function CreateOrderTable()

        Dim week                        As Integer
        Dim projectCount        As Integer
        Dim sum                         As Integer
                
        ReDim gOrderTable(2, gExcelEnv.SimulationDuration)

        For week = 1 To gExcelEnv.SimulationDuration
                projectCount = PoissonRandom(gExcelEnv.avgProjects)            ' �̹��� �߻��ϴ� ������Ʈ ����
                gOrderTable(1, week) = sum
                gOrderTable(2, week) = projectCount

                ' �̹��� ���� �߻��� ������Ʈ ����. �����ֿ� ��ϵȴ�. ==> �����ֱ��� �߻��� ������Ʈ ������������. vba���� do while ���� ��... ����
                sum = sum + projectCount
        Next

        gTotalProjectNum = sum
        gTableInitialized = 1
        
End Function

Private Function CreateProjects()

        Dim week                        As Integer
        Dim id                          As Integer
        Dim startPrjNum         As Integer
        Dim endPrjNum           As Integer
        Dim preTotal            As Integer
        Dim tempPrj             As clsProject

        If gTotalProjectNum <= 0 Then
                MsgBox "gTotalProjectNum is 0", vbExclamation
                Exit Function
        End If

        '������Ʈ���� �����Ѵ�.
        ReDim gProjectTable(gTotalProjectNum)

        For week = 1 To gExcelEnv.SimulationDuration
                
                preTotal = gOrderTable(1, week)                         ' ���� �Ⱓ ���� �߻��� ������Ʈ ���� ����
                startPrjNum = preTotal + 1                                      ' �̹� �Ⱓ ����������Ʈ ��ȣ
                endPrjNum = gOrderTable(2, week) + preTotal             ' �̹� �Ⱓ ������ ������Ʈ ��ȣ
                
                If startPrjNum = 0 Then
                        GoTo Continue
                End If

                If startPrjNum > endPrjNum Then
                        GoTo Continue
                End If

                ' �̹� �ֿ� �߻��� ������Ʈ���� �����Ѵ�.
                For id = startPrjNum To endPrjNum '
                        Set tempPrj = New clsProject
                        Call tempPrj.Init(P_TYPE_EXTERNAL, id, week)
                        Set gProjectTable(id) = tempPrj
                        'Call tempPrj.PrintInfo()
                Next

Continue:

        Next
        
End Function

Public Function Epilogue()

        ' Application.ScreenUpdating = True
        ' Application.Calculation = xlAutomatic
        ' Application.EnableEvents = True
        ' �� �׸��� ���� �ٽ� ���� ����. ActiveSheet.DisplayPageBreaks = True

End Function


'' �־��� Range ���� �ش� ��Ʈ���� �������� ���� �����´�
Public Function GetVariableValue(rng As Range, variableName As String) As Variant
    Dim dataArray As Variant
    Dim matchIndex As Variant

    ' ������ �迭�� ��ȯ
    dataArray = rng.Value

    ' ���� �̸��� �ִ� ��ġ�� ã��
    matchIndex = Application.Match(variableName, Application.Index(dataArray, 0, 1), 0)
    
    ' ���� �̸��� �ִ� ��� �� ��ȯ
    If Not IsError(matchIndex) Then
        GetVariableValue = dataArray(matchIndex, 2)
    Else
        GetVariableValue = "Variable not found" 'song ==> ���� ó���� ���߿� ����.
    End If

End Function

Sub PrintArrayWithLine(ws As Worksheet, startRow As Long, startCol As Long, dataArray As Variant)

    Dim startRange As Range
    Dim endRange As Range
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long
    
    Set startRange = ws.Cells(startRow, startCol) ' ���� �� ����
    
    ' �迭�� ���� Ȯ��
    Dim dimensions As Integer
    dimensions = GetArrayDimensions(dataArray)
    
    If dimensions = 1 Then
        ' 1���� �迭 ó��
        numRows = UBound(dataArray) - LBound(dataArray) + 1
        numCols = 1 ' 1���� �迭�̹Ƿ� ���� ���� 1
        
        Set endRange = startRange.Resize(numRows, numCols) ' ����� ���� ����
        
        ' 1���� �迭�� 2���� ������ ���
        For i = 1 To numRows
            endRange.Cells(i, 1).Value = dataArray(i)
        Next i
        
    ElseIf dimensions = 2 Then
        ' 2���� �迭 ó��
        numRows = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
        numCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
        
        Set endRange = startRange.Resize(numRows, numCols) ' ����� ���� ����
        endRange.Value = dataArray ' �迭�� ��Ʈ�� ���
    End If
    
    ' �׵θ� �׸���
    With endRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

End Sub

' �迭�� ������ ���ϴ� �Լ�
Function GetArrayDimensions(arr As Variant) As Integer

    Dim dimCount As Integer
    Dim currentDim As Integer
    
    On Error GoTo ErrHandler
    dimCount = 0
    currentDim = 0
    
    Do While True
        currentDim = currentDim + 1
        ' �迭�� �� ������ Ȯ��
        Dim temp As Long
        temp = LBound(arr, currentDim)
        dimCount = currentDim
    Loop
    
ErrHandler:
    If Err.Number <> 0 Then
        GetArrayDimensions = dimCount
    End If
    On Error GoTo 0
End Function




Function PrintProjectHeader()

        Call ClearSheet(gWsProject)                     '��Ʈ�� ��� ������ ����� �� ���� ����

        Dim arrHeader As Variant
    Dim strHeader As String

        ' ù ��° �� ���
    strHeader = "Ÿ��,����,������,���۰���,�Ⱓ,����,����,����,����%,����Ƚ��,CF1%,CF2%,CF3%,����,�ߵ���,�ܱ�"
    arrHeader = Split(strHeader, ",")
    arrHeader = ConvertToBase1(arrHeader)
        arrHeader = ConvertTo1xN(arrHeader)
        Call PrintArrayWithLine(gWsProject, 2, 1, arrHeader)

    
    ' �� ��° �� ���
    strHeader = ",Dur,start,end,HR_H,HR_M,HR_L,,,,mon_cf1,mon_cf2,mon_cf3"
        arrHeader = Split(strHeader, ",")
    arrHeader = ConvertToBase1(arrHeader)
        arrHeader = ConvertTo1xN(arrHeader)
        Call PrintArrayWithLine(gWsProject, 3, 1, arrHeader)
        
End Function




Function PrintProjectAll()

        Dim temp As clsProject
        Dim i As Integer

        For i = 1 To gTotalProjectNum
                Set temp = gProjectTable(i)
                Call temp.PrintInfo
        Next
        
End Function


' 0 ��� �迭�� 1 ��� �迭�� ��ȯ�ϴ� �Լ�
Function ConvertToBase1(arr As Variant) As Variant
    Dim i As Integer
    Dim newArr() As Variant
    
    ReDim newArr(1 To UBound(arr) - LBound(arr) + 1)
    For i = LBound(arr) To UBound(arr)
        newArr(i - LBound(arr) + 1) = arr(i)
    Next i
    
    ConvertToBase1 = newArr
End Function

Function ConvertTo1xN(arr As Variant) As Variant
    Dim i As Integer
    Dim newArr() As Variant
    Dim numCols As Integer
    
    numCols = UBound(arr) - LBound(arr) + 1
    ReDim newArr(1 To 1, 1 To numCols)
    
    For i = LBound(arr) To UBound(arr)
        newArr(1, i - LBound(arr) + 1) = arr(i)
    Next i
    
    ConvertTo1xN = newArr
End Function

Function PrintDashboard()       ' ������ ��ú��带 ��Ʈ�� ����Ѵ�

        On Error GoTo ErrorHandler

        Call ClearSheet(gWsDashboard)                   '��Ʈ�� ��� ������ ����� �� ���� ����

        Dim arrHeader As Variant
    arrHeader = Array("��", "����", "����")

        Call PrintArrayWithLine(gWsDashboard, 2, 1, arrHeader)          ' �����׸��� ����
        Call PrintArrayWithLine(gWsDashboard, 2, 2, gPrintDurationTable) '�Ⱓ�� ����
        Call PrintArrayWithLine(gWsDashboard, 3, 2, gOrderTable)        ' ������ ���´�.

        ' Set myArray = GetPrintHeaderTable
        ' PrintArrayWithLine(ws, 1, 1,myArray)

        Exit Function

        ' Set myArray = GetProjectInfoTable
        ' PrintArrayWithLine(DBOARD_SHEET_NAME, 2, 2,myArray)

ErrorHandler:
                Call HandleError("PrintDashboard", Err.Description)

End Function


Function ClearSheet(ws As Worksheet)

        With ws
                Dim endRow As Long ' ��������
        Dim endCol As Long ' ��������
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' ���� ������ ������ �����Ѵ�.
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).UnMerge
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).Clear
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).ClearContents

        End With
        
End Function


' lambda(��� �߻���)�� ���ڷ� �޾� ���Ƽ� ������ ������ ���� ���� ��ȯ�մϴ�.
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



' On Error GoTo ErrorHandler
' ErrorHandler:
'     Call HandleError("ExampleFunction", Err.Description)


Sub HandleError(funcName As String, errMsg As String)
    MsgBox "Error in Sub " & funcName & ": " & errMsg, vbExclamation
End Sub
