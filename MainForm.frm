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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton GenBoard 
      Caption         =   "������Ʈ����"
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
      Caption         =   "�����η�(�߱�)"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "�����η�(���)"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1515
      Width           =   1935
   End
   Begin VB.Label ������Ʈ�߻��� 
      Caption         =   "������Ʈ�߻���"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   885
      Width           =   1935
   End
   Begin VB.Label SimulTearm 
      Caption         =   "�ð�(��)"
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

' ��ú��带 �����Ѵ�.
Private Sub GenDBoard_Click()

    LetExcelInitialized = 0 ' ���ο� ������Ʈ���� �����ϱ� ���� �ʱ�ȭ �÷��� ����
    LetTableInitialized = 0 ' ���ο� ������Ʈ���� �����ϱ� ���� �ʱ�ȭ �÷��� ����

    'Call Prologue(0)        ' ��ü �Ķ���� �ε�-> ��ú��� ���� -> ������Ʈ ����
    'Call PrintDashboard     ' ��ú��带 ��Ʈ�� ����Ѵ�
    'Call PrintProjectHeader         ' ������Ʈ�� ��Ʈ�� ����Ѵ�
    'Call PrintProjectAll
    
    ' Call CreateDashboard()     ' ��ú��带 �����ϰ� ��ü ������Ʈ�� ������ ���Ѵ�.
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

