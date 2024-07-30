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

' 회사 속성
Public comID As Integer ' 회사 ID
Public comName As String ' 회사 이름

' 프로젝트 테이블
Private comProjectTable() As clsProject
Private comDoingTable() As Variant
Private comDoneTable() As Variant
Private comDefferTable() As Variant

Public Property Get PropertyDoingTable() As Variant
    PropertyDoingTable = comDoingTable
End Property

Public Property Get PropertyDoneTable() As Variant
    PropertyDoneTable = comDoneTable
End Property

Public Property Get PropertyDefferTable() As Variant
    PropertyDefferTable = comDefferTable
End Property

Public Property Get comDoingTableSize() As Integer
    comDoingTableSize = UBound(comDoingTable, 2)
End Property

Public Property Get comDoneTableSize() As Integer
    comDoneTableSize = UBound(comDoneTable, 2)
End Property

Public Property Get comDefferTableSize() As Integer
    comDefferTableSize = UBound(comDefferTable, 2)
End Property

' 초기화
Public Sub Init()
    Dim totalProjects As Integer
    totalProjects = gTotalProjectNum
    
    ReDim comProjectTable(1 To totalProjects)
    ReDim comDoingTable(1 To totalProjects)
    ReDim comDoneTable(1 To totalProjects)
    ReDim comDefferTable(1 To totalProjects)
End Sub

' 의사결정
Public Sub Decision(period As Integer)
    Dim i As Integer
    Dim project As clsProject
    
    For i = 1 To gTotalProjectNum
        Set project = gProjectTable(i)
        If project.StartDate = 0 Then
            If CanStartProject(project, period) Then
                StartProject project, period
            End If
        ElseIf project.StartDate > 0 And project.StartDate <= period Then
            If IsProjectDone(project, period) Then
                CompleteProject project, period
            End If
        End If
    Next i
End Sub

' 프로젝트 시작 가능 여부
Private Function CanStartProject(project As clsProject, period As Integer) As Boolean
    If period >= project.PossibleStartDate Then
        CanStartProject = True
    Else
        CanStartProject = False
    End If
End Function

' 프로젝트 시작
Private Sub StartProject(project As clsProject, period As Integer)
    project.StartDate = period
    comDoingTable(project.ProjectNum) = project
End Sub

' 프로젝트 완료 여부
Private Function IsProjectDone(project As clsProject, period As Integer) As Boolean
    If period >= project.StartDate + project.ProjectDuration Then
        IsProjectDone = True
    Else
        IsProjectDone = False
    End If
End Function

' 프로젝트 완료 처리
Private Sub CompleteProject(project As clsProject, period As Integer)
    Dim idx As Integer
    For idx = LBound(comDoingTable) To UBound(comDoingTable)
        If comDoingTable(idx).ProjectNum = project.ProjectNum Then
            comDoneTable(idx) = comDoingTable(idx)
            comDoingTable(idx) = Nothing
            Exit For
        End If
    Next idx
End Sub

