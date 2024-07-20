VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Simulator"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   LinkTopic       =   "MainForm"
   ScaleHeight     =   7785
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton GenBoard 
      Caption         =   "프로젝트생성"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox Hr_Init_M 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Hr_Init_H 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox avgProjects 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Term 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "보유인력(중급)"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "보유인력(고급)"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1515
      Width           =   1935
   End
   Begin VB.Label 프로젝트발생빈 
      Caption         =   "프로젝트발생빈도"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   885
      Width           =   1935
   End
   Begin VB.Label SimulTearm 
      Caption         =   "시간(주)"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Base 1

' 대시보드를 생성한다.
Private Sub GenDBoard_Click()

    LetExcelInitialized = 0 ' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경
    LetTableInitialized = 0 ' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경

    'Call Prologue(0)        ' 전체 파라메터 로드-> 대시보드 생성 -> 프로젝트 생성
    'Call PrintDashboard     ' 대시보드를 시트에 출력한다
    'Call PrintProjectHeader         ' 프로젝트를 시트에 출력한다
    'Call PrintProjectAll
    
    ' Call CreateDashboard()     ' 대시보드를 생성하고 전체 프로젝트의 갯수를 구한다.
    ' Call Epilogue()
    Dim i As Integer
       
    
    

End Sub



Private Sub Form_Load()
    'Set xlwbook = xl.Workbooks.Open("c:\book1.xls")
    'Set xlsheet = xlwbook.Sheets.Item(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set xlwbook = Nothing
    'Set xl = Nothing
End Sub

