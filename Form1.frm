VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "FirstClockLite"
   ClientHeight    =   495
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   7530
   StartUpPosition =   3  '�t�ιw�]��
   Begin MSComDlg.CommonDialog Com2 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7080
      Top             =   0
   End
   Begin MSComDlg.CommonDialog Com1 
      Left            =   6600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu setting 
      Caption         =   "�]�w(&S)"
      Begin VB.Menu font 
         Caption         =   "�r��(&F)"
      End
      Begin VB.Menu color 
         Caption         =   "�C��(&C)"
      End
      Begin VB.Menu backcolor 
         Caption         =   "�I���C��(&B)"
      End
      Begin VB.Menu UITop 
         Caption         =   "�w��b�̤W�h(&U)"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim X

Private Sub backcolor_Click()
Com1.ShowColor
Form1.backcolor = Com1.color
End Sub

Private Sub color_Click()
Com1.ShowColor
Label2.ForeColor = Com1.color
Label1.ForeColor = Com1.color
End Sub

Private Sub Command1_Click()
'If X = 0 Then    '�u�ŧiX�ܼƹw�]�ȬO0�A�Ĥ@�����U���s��X��0
'IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)   '�̤W�h���
'X = X + 1      '�O���w���L�@�����s
'Command1.Caption = "�����̤W�h���(&T)"
'ElseIf X = 1 Then    '�ĤG�������s
'IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3)   '�����̤W�h���
'X = 0   '�NX�ܦ^0
'Command1.Caption = "�̤W�h���(&T)"
'End If
End Sub

Private Sub font_Click()
Com2.ShowFont
Label2.FontBold = Com2.FontBold
Label2.FontItalic = Com2.FontItalic
Label2.FontStrikethru = Com2.FontStrikethru
Label2.FontUnderline = Com2.FontUnderline
Label2.FontName = Com2.FontName
Label2.FontSize = Com2.FontSize
Label1.FontBold = Com2.FontBold
Label1.FontItalic = Com2.FontItalic
Label1.FontStrikethru = Com2.FontStrikethru
Label1.FontUnderline = Com2.FontUnderline
Label1.FontName = Com2.FontName
Label1.FontSize = Com2.FontSize
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time()
Label2.Caption = Now()
End Sub

Private Sub UITop_Click()
If UITop.Checked Then    '�u�ŧiX�ܼƹw�]�ȬO0�A�Ĥ@�����U���s��X��0
IntR = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)   '�̤W�h���
UITop.Caption = "�����̤W�h���(&T)"
UITop.Checked = False
Else
IntR = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3)   '�����̤W�h���
UITop.Caption = "�̤W�h���(&T)"
UITop.Checked = True
End If
End Sub
