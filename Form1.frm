VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "�������ֻ�"
   ClientHeight    =   3825
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7875
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�����ʶ�"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3240
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "˫�����"
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YouGet"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7455
   End
   Begin VB.Menu �˵� 
      Caption         =   "�˵�"
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu �ظ���ϰ 
         Caption         =   "�ظ���ϰ"
      End
      Begin VB.Menu �ո��� 
         Caption         =   "�ո���"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ��һ�� 
         Caption         =   "��һ��"
         Shortcut        =   ^Z
      End
      Begin VB.Menu ��һ�� 
         Caption         =   "��һ��"
         Shortcut        =   ^X
      End
      Begin VB.Menu �ٶ�һ�� 
         Caption         =   "�ٶ�һ��"
         Shortcut        =   ^W
      End
      Begin VB.Menu ������ȡ 
         Caption         =   "������ȡ"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu ��ϵ���� 
         Caption         =   "��ϵ����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1, s2, sc1, sc2, fr1, fr2 As String
Dim sj(99999999) As String
Dim k, sum, zongshu As Long
Dim cs, cf As Integer '�������ظ�
Dim yy As Byte
Dim kg As Integer '�ո�
 
 
 
 
Private Sub Form_Load()
    sum = 0
    kg = 0
    cs = 1
    cf = 1
    zongshu = 0
    
    yy = 0
    
    
    
    If Dir("wz.txt") <> "" Then
   Open "wz.txt" For Input As #1
    
        Line Input #1, s1
        Line Input #1, s2
      
       Form1.Left = s1
       Form1.Top = s2
       
'    Do While EOF(0)
'        Line Input #1, s1
'
'    Loop
    
    Close #1
    End If
    
     If Dir("sj.txt") <> "" Then
     
    Open "sj.txt" For Input As #1
    
       k = 0
    
    If FileLen("sj.txt") > -1 Then
    'MsgBox "1"
    
    Do While Not EOF(1)
        Line Input #1, s1
        
        If s1 <> "" Then
        sj(k) = Trim(s1)      '����  ��������
        k = k + 1
        
        End If
       ' MsgBox s1
    Loop
    
    Else
    End If
    Close #1
    
    End If
    
    If sj(sum) <> "" Then
    Label1.Caption = sj(sum)
    End If
    
    
    'MsgBox sj(0) & "/" & sj(1)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fr1 = Form1.Left
    fr2 = Form1.Top
    Open "wz.txt" For Output As #1
    
    Print #1, fr1
    Print #1, fr2
        
    
    Close #1
End Sub

Private Sub Text1_DblClick()
Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   
   'MsgBox cs & "/" & cf
    If Label1.Caption = Trim(Text1.Text) Then
    'MsgBox "1"
        If cs <= cf And cs > 0 Then
    'MsgBox "2"
        'MsgBox cs & "/" & cf
   ' MsgBox sum & "/" & sj(sum)
    
        If sj(sum + 1) <> "" Then
    'MsgBox "3"
        sum = sum + 1
        Label1.Caption = sj(sum)
        Text1.Text = ""
        cf = 0
        If yy = 1 Then
        CreateObject("SAPI.SpVoice").Speak sj(sum)
        End If
        
        Else
            sum = 0
            If sj(sum) <> "" Then
            Label1.Caption = sj(sum)
            Text1.Text = ""
            cf = 0
            
             If yy = 1 Then
            CreateObject("SAPI.SpVoice").Speak sj(sum)
             End If
             
            End If
            
        End If
        
    
     End If
     cf = cf + 1
    End If

Text1.Text = ""
    
    
     'MsgBox sum
End If




If kg = 1 And KeyCode = 32 Then
    
    Open "sj.txt" For Append As #1
    
    
        Print #1, Trim(Text1.Text)
    
    
    Close #1
    
    
    Text1.Text = ""
    
    
End If















End Sub

Private Sub Timer1_Timer()
Label2.Caption = "��ţ�" & sum + 1
Label3.Caption = "������" & k
Label4.Caption = "�ظ���" & cf


If Check1.Value = vbChecked Then
    yy = 1
Else
    yy = 0

End If

End Sub

Private Sub ��������_Click()
   ' MsgBox (App.Path & "\sj.txt")
If Dir("sj.txt") <> "" Then
    Shell ("explorer.exe " & App.Path & "\sj.txt")
Else
    Open "sj.txt" For Output As #1
    
    Close #1
    
End If






End Sub

Private Sub �ո���_Click()
   If kg = 0 Then
        If MsgBox("�Ƿ����ù��ܣ����Կո��β���Զ������ı������ݣ���", 36, "YouGet") = vbYes Then
            kg = 1
        End If
    Else
        If MsgBox("�Ƿ�رոù��ܣ���֮ǰ¼������ݣ����ڲ˵����������ݡ��в鿴����", 36, "YouGet") = vbYes Then
            kg = 0
        End If
    End If
    
End Sub

Private Sub ��ϵ����_Click()
MsgBox "      ΢�Ź��ںţ��и�ͷ������" & vbCrLf & "      QQ:1377351008", 64, "YouGet"
End Sub

Private Sub ��һ��_Click()

    If sum > 0 Then
    'MsgBox sum
    sum = sum - 1
    'MsgBox sum
    Label1.Caption = sj(sum)
    Text1.Text = ""
    
     If yy = 1 Then
    CreateObject("SAPI.SpVoice").Speak sj(sum)
     End If
    'sum = sum + 1
    Else
    
   
   
    sum = k - 1
    ' MsgBox sum & "/" & k
    'MsgBox sum
    Label1.Caption = sj(sum)
    Text1.Text = ""
    
     If yy = 1 Then
    CreateObject("SAPI.SpVoice").Speak sj(sum)
     End If
    
    'MsgBox "��ǰΪ��һ�����޷�������һ����", 48, "YouGet"
    
    End If
    
    
'        If sj(sum) <> "" Then
'
'        Label1.Caption = sj(sum)
'        Text1.Text = ""
'        CreateObject("SAPI.SpVoice").Speak sj(sum)
'        sum = sum + 1
'        Else
'            sum = 0
'            If sj(sum) <> "" Then
'            Label1.Caption = sj(sum)
'            Text1.Text = ""
'            CreateObject("SAPI.SpVoice").Speak sj(sum)
'            End If
'
'        End If
        
    
   
End Sub

Private Sub ��һ��_Click()
  
            sum = sum + 1
        If sj(sum) <> "" Then

        Label1.Caption = sj(sum)
        Text1.Text = ""
         If yy = 1 Then
        CreateObject("SAPI.SpVoice").Speak sj(sum)
         End If
        'sum = sum + 1
        Else
            sum = 0
            If sj(sum) <> "" Then
            Label1.Caption = sj(sum)
            Text1.Text = ""
             If yy = 1 Then
            CreateObject("SAPI.SpVoice").Speak sj(sum)
             End If
            End If

        End If
        
    
End Sub

Private Sub ������ȡ_Click()
    If Text1.Text <> "" Then
        CreateObject("SAPI.SpVoice").Speak Text1.Text
    Else
    
        MsgBox "�����������������Ҫ�ʶ������ݣ�"
    End If
End Sub

Private Sub �ٶ�һ��_Click()
    
    CreateObject("SAPI.SpVoice").Speak Label1.Caption
    
    
End Sub

Private Sub ��������_Click()
    MsgBox "�ڼ�������������--ʱ�������--����--����--����ϵͳ���������С����壬���ġ����ɡ�", , "YouGet"
End Sub

Private Sub �ظ���ϰ_Click()
    If MsgBox("�Ƿ����ظ���ϰ", 36, "YouGet") = vbYes Then
    
       
       Do
        cs = InputBox("�������ظ��Ĵ�������0Ϊ���޴Σ�", "YouGet", "1")
       Loop Until IsNumeric(cs) = True
       
       cf = 1
       
    Else
    
        cf = 1
        cs = 1
    
    End If
    
End Sub
