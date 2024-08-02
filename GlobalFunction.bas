Attribute VB_Name = "GlobalFunction"
Option Explicit
Option Base 1

' Define Global Variables
Public xlApp                As Object   ' ���� ���ø����̼� ��ü
Public xlWb                 As Object   ' ���� ��ũ�� ��ü

Public gWsParameters        As Object   ' Parameters ��Ʈ ��ü
Public gWsDashboard         As Object   ' Dashboard ��Ʈ ��ü
Public gWsProject           As Object   ' Project ��Ʈ ��ü
Public gWsActivity_Struct   As Object   ' Activity_Struct ��Ʈ ��ü

Public Const PARAMETERS_SHEET_NAME = "parameters"
Public Const DBOARD_SHEET_NAME = "dashboard"
Public Const PROJECT_SHEET_NAME = "project"
Public Const ACTIVITY_SHEET_NAME = "activity_struct"

Public GCurrentPath As String
Public gProjectLoadOrCreate As Integer ' ���α׷� ���۽� ������Ʈ�� �������� ���� �����͸� �ε����� �����ϴ� ����

Public Const P_TYPE_EXTERNAL = 0
Public Const P_TYPE_INTERNAL = 1

Public Const MAX_ACT As Integer = 4
Public Const MAX_N_CF As Integer = 3
Public Const PRJ_SHEET_HEADER_W As Integer = 16
Public Const PRJ_SHEET_HEADER_H As Integer = 8
Public Const RND_HR_H = 20
Public Const RND_HR_M = 70
Public Const MAX_PRJ_TYPE As Integer = 5
Public Const RND_PRJ_TYPE1 As Integer = 20
Public Const RND_PRJ_TYPE2 As Integer = 70
Public Const RND_PRJ_TYPE3 As Integer = 20
Public Const RND_PRJ_TYPE4 As Integer = 70
Public Const RND_PRJ_TYPE5 As Integer = 20
Public Const ORDER_TABLE_INDEX As Long = 1
Public Const DONG_TABLE_INDEX = 6
Public Const PROJECT_TABLE_INDEX As Long = 3

Private gExcelInitialized As Boolean
Private gTableInitialized As Boolean
Public gTotalProjectNum As Integer
Public GlobalEnv As Environment
Public gOrderTable() As Variant
Public gProjectTable() As clsProject
Public gPrintDurationTable() As Variant

Type Environment
    SimulationWeeks As Integer
    WeeklyProb As Double
    Hr_Init_H As Integer
    Hr_Init_M As Integer
    Hr_Init_L As Integer
    Hr_LeadTime As Integer
    Cash_Init As Integer
    Problem As Integer
End Type

Type Activity
    ActivityType As Integer
    Duration As Integer
    StartDate As Integer
    EndDate As Integer
    HighSkill As Integer
    MidSkill As Integer
    LowSkill As Integer
End Type

Public Property Get GetExcelEnv() As Environment
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

' ���������� ����ϴ� ���̺���� ä���.
Sub Prologue(TableInit As Integer)

    Dim i As Integer
    
    If gExcelInitialized = False Then

        ReDim gPrintDurationTable(1 To GlobalEnv.SimulationWeeks)
        For i = 1 To GlobalEnv.SimulationWeeks
            gPrintDurationTable(i) = i
        Next i

        gExcelInitialized = True
    End If
    
    ' ���̺���� ���� �����ϰų� �������� �ε��ϰų�.
    ' ���� ó���� �������� ����ڰ� �����ؼ� ����ϵ��� ����.
    gTableInitialized = (TableInit = 1)
    
    If gTableInitialized = False Then
        'Call BuildTables ' ���̺��� ���� �����ϰ� ������ ����Ѵ�.
        
        Call PrintProjectHeader ' Project ��Ʈ�� ����� ����Ѵ�.
        Call CreateProjects     ' ������Ʈ�� �����Ѵ�.
        
    Else
        Call LoadTablesFromExcel ' ������ ��ϵ� ����� ���̺��� ä���.
    End If

    gTableInitialized = True
    
End Sub


Sub LoadTablesFromExcel()
    Call LoadOrderTable
    Call LoadProjects
End Sub


Private Function LoadOrderTable() As Boolean
    ReDim gOrderTable(2, GlobalEnv.SimulationWeeks)

    Dim startIndex As Long
    startIndex = ORDER_TABLE_INDEX + 2

    With gWsDashboard
        gOrderTable = .Range(.Cells(startIndex, 2), .Cells(startIndex + 1, GlobalEnv.SimulationWeeks + 1)).value
    End With

    gTotalProjectNum = gOrderTable(1, GlobalEnv.SimulationWeeks) + gOrderTable(2, GlobalEnv.SimulationWeeks)
End Function

Private Function LoadProjects() As Boolean
    Dim prjID As Integer
    Dim startRow As Long
    Dim endRow As Long
    Dim prjInfo As Variant
    Dim tempPrj As clsProject

    ReDim gProjectTable(gTotalProjectNum)

    For prjID = 1 To gTotalProjectNum
        Set tempPrj = New clsProject
        startRow = PROJECT_TABLE_INDEX + (prjID - 1) * PRJ_SHEET_HEADER_H + 1
        endRow = startRow + PRJ_SHEET_HEADER_H - 1

        With gWsProject
            prjInfo = .Range(.Cells(startRow, 1), .Cells(endRow, PRJ_SHEET_HEADER_W)).value
        End With

        tempPrj.ProjectType = prjInfo(1, 1)
        tempPrj.ProjectNum = prjInfo(1, 2)
        tempPrj.OrderDate = prjInfo(1, 3)
        tempPrj.PossibleStartDate = prjInfo(1, 4)
        tempPrj.ProjectDuration = prjInfo(1, 5)
        tempPrj.StartDate = prjInfo(1, 6)
        tempPrj.Profit = prjInfo(1, 7)
        tempPrj.Experience = prjInfo(1, 8)
        tempPrj.SuccessProbability = prjInfo(1, 9)
        
        Dim tempCF(1 To MAX_N_CF) As Integer
        Dim i As Integer
        For i = 1 To MAX_N_CF
            tempCF(i) = prjInfo(1, 10 + i)
        Next i
        tempPrj.SetPrjCashFlows tempCF

        tempPrj.FirstPayment = prjInfo(1, 14)
        tempPrj.MiddlePayment = prjInfo(1, 15)
        tempPrj.FinalPayment = prjInfo(1, 16)

        tempPrj.NumActivities = prjInfo(2, 2)
        tempPrj.FirstPaymentMonth = prjInfo(2, 11)
        tempPrj.MiddlePaymentMonth = prjInfo(2, 12)
        tempPrj.FinalPaymentMonth = prjInfo(2, 13)
        
        Dim tempAct As Activity
        For i = 1 To tempPrj.NumActivities
            tempAct.Duration = prjInfo(2 + i, 2)
            tempAct.StartDate = prjInfo(2 + i, 3)
            tempAct.EndDate = prjInfo(2 + i, 4)
            tempAct.HighSkill = prjInfo(2 + i, 5)
            tempAct.MidSkill = prjInfo(2 + i, 6)
            tempAct.LowSkill = prjInfo(2 + i, 7)
            'tempPrj.SetPrjActivities i, tempAct
        Next i

        Set gProjectTable(prjID) = tempPrj
    Next prjID
End Function

' �߻��� ������Ʈ�� ������ ���̺� ����Ѵ�.
Public Function CreateOrderTable()
    Dim week As Integer
    Dim projectCount As Integer
    Dim sum As Integer
    
    ReDim gOrderTable(2, GlobalEnv.SimulationWeeks)

    For week = 1 To GlobalEnv.SimulationWeeks
        projectCount = PoissonRandom(GlobalEnv.WeeklyProb)        ' �̹��� �߻��ϴ� ������Ʈ ����
        gOrderTable(1, week) = sum
        gOrderTable(2, week) = projectCount
        
        ' �̹��� ���� �߻��� ������Ʈ ����. �����ֿ� ��ϵȴ�. ==> �����ֱ��� �߻��� ������Ʈ ������������. vba���� do while ���� ��... ����
        sum = sum + projectCount
    Next week

    gTotalProjectNum = sum
    
End Function




' ������Ʈ�� �����Ѵ�.
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
    
    ' ������Ʈ ������ �����ϴ� ���̺� �߻��� ������Ʈ�� ������ ����Ѵ�.
    ReDim gProjectTable(gTotalProjectNum)

    MainForm.ProgressBar1.Max = GlobalEnv.SimulationWeeks
    MainForm.ProgressBar1.Min = 0
    MainForm.ProgressBar1.Text = "������Ʈ ������"
    
    For week = 1 To GlobalEnv.SimulationWeeks
        preTotal = gOrderTable(1, week)
        startPrjNum = preTotal + 1
        endPrjNum = gOrderTable(2, week) + preTotal

        If startPrjNum = 0 Then GoTo Continue
        If startPrjNum > endPrjNum Then GoTo Continue

        ' �̹� �ֿ� �߻��� ������Ʈ���� �����ϰ� �ʱ�ȭ �Ѵ�.
        For Id = startPrjNum To endPrjNum
            Set tempPrj = New clsProject
            Call tempPrj.Init(P_TYPE_EXTERNAL, Id, week)
            Set gProjectTable(Id) = tempPrj
        Next Id

Continue:
        MainForm.ProgressBar1.value = week
    Next week
End Function



' ������Ʈ���� ����� ���� ��Ʈ�� ����� �����Ѵ�.
' VB 6.0���� Option Base 1�� ����ߴ���, Split �Լ��� �⺻������ 0 ��� �迭�� ��ȯ
' ���� 1 ��� �迭�� ��ȯ�ϴ� �ڵ带 �߰� ��
Public Sub PrintProjectHeader()

    Dim MyArray() As Variant, TempArray() As String, strHeader As String
    
    Call ClearSheet(gWsProject) '��Ʈ�� ��� ������ ����� �� ���� ����
    
    strHeader = "Ÿ��,����,������,���۰���,�Ⱓ,����,����,����,����%,nCF,CF1%,CF2%,CF3%,����,�ߵ�,�ܱ�"
    TempArray = Split(strHeader, ",")
    MyArray = ConvertToBase1(TempArray) ' TempArray�� 1 ��� �迭�� ��ȯ
    Call PrintArrayWithLine(gWsProject, 1, 1, MyArray)
    
    strHeader = ",Dur,start,end,HR_H,HR_M,HR_L,,,mon_cf1,mon_cf2,mon_cf3,,,,"
    TempArray = Split(strHeader, ",")
    MyArray = ConvertToBase1(TempArray)
    Call PrintArrayWithLine(gWsProject, 2, 1, MyArray)
    
End Sub


Public Function GetVariableValue(rng As Object, variableName As String) As Variant
    Dim dataArray As Variant
    Dim matchIndex As Variant

    dataArray = rng.value
    matchIndex = Application.Match(variableName, Application.Index(dataArray, 0, 1), 0)
    
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
        ' �ʱ�ȭ
        rowCount = 0
        colCount = 0

        ' 1���� �迭�� ���
        On Error Resume Next
        rowCount = UBound(arr, 2)
        If Err.Number <> 0 Then
            ' 1���� �迭
            rowCount = 1
            colCount = UBound(arr, 1) - LBound(arr, 1) + 1
            Err.Clear
        Else
            ' 2���� �迭
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
'    strHeader = "Ÿ��,����,������,���۰���,�Ⱓ,����,����,����,����%,����Ƚ��,CF1%,CF2%,CF3%,����,�ߵ���,�ܱ�"
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

Function PrintProjectAll()
    Dim temp As clsProject
    Dim i As Integer

    For i = 1 To gTotalProjectNum
        Set temp = gProjectTable(i)
        Call temp.PrintInfo
    Next i
End Function

Function ConvertToBase1(arr As Variant) As Variant
    Dim i As Integer
    Dim newArr() As Variant
    ReDim newArr(1 To UBound(arr) - LBound(arr) + 1)
    For i = LBound(arr) To UBound(arr)
        newArr(i - LBound(arr) + 1) = arr(i)
    Next i
    ConvertToBase1 = newArr
End Function

Function PivotArray(arr As Variant) As Variant
    Dim i As Integer
    Dim rowCount As Integer
    Dim result() As Variant
    
    ' 1���� �迭�� ũ�� ���ϱ�
    rowCount = UBound(arr) - LBound(arr) + 1
    
    ' 2���� �迭 ũ�� ���� (rowCount x 1)
    ReDim result(1 To rowCount, 1 To 1)
    
    ' 1���� �迭�� 2���� �迭�� ��ȯ
    For i = LBound(arr) To UBound(arr)
        result(i, 1) = arr(i)
    Next i
    
    PivotArray = result
End Function


Function PrintDashboard()
    Call ClearSheet(gWsDashboard)

    Dim arrHeader As Variant
    arrHeader = Array("��", "����", "����")
    arrHeader = PivotArray(arrHeader)

    Call PrintArrayWithLine(gWsDashboard, 2, 1, arrHeader)
    Call PrintArrayWithLine(gWsDashboard, 2, 2, gPrintDurationTable)
    Call PrintArrayWithLine(gWsDashboard, 3, 2, gOrderTable)
    
    arrHeader = Array("����", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(gWsDashboard, 6, 1, arrHeader)
    
    arrHeader = Array("����", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(gWsDashboard, 11, 1, arrHeader)
    
    arrHeader = Array("�ѿ�", "HR_H", "HR_M", "HR_L")
    arrHeader = PivotArray(arrHeader)
    Call PrintArrayWithLine(gWsDashboard, 16, 1, arrHeader)
    
End Function

Function ClearSheet(ws As Object)
    With ws
        Dim endRow As Long
        Dim endCol As Long
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
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
    Dim i As Long
    
    ' ���� ��ũ��Ʈ�� ������ �� ���ϱ�
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' xlUp = -4162

    ' 1�� ���� ��ȸ�ϸ� Ű���� ã��
    For i = 1 To lastRow
        If InStr(1, ws.Cells(i, 1).value, keyword, vbTextCompare) > 0 Then
            FindRowWithKeyword = i
            Exit Function
        End If
    Next i

    ' Ű���带 ã�� ���� ���
    FindRowWithKeyword = 0
End Function

Function GetLastColumnValue(ws As Object, rowNumber As Long) As Variant
    Dim lastCol As Long
    
    ' Ư�� ���� ������ �� ��ȣ ���ϱ�
    lastCol = ws.Cells(rowNumber, ws.Columns.Count).End(-4159).Column ' xlToLeft = -4159

    ' ������ ���� �� ��ȯ
    GetLastColumnValue = ws.Cells(rowNumber, lastCol).value
End Function
