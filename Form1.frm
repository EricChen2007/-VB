VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ܽ���"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   3135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "����"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "����"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtEncrypted 
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtPlain 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "���ܣ�"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "�ı���"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "�򵥵ļ��ܽ��ܳ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ���ܺͽ��ܳ����VB����ʾ��

' ����ȫ�ֱ���
Dim key As String

Private Sub Form_Load()
    ' �������ʱ������
    Me.Caption = "�򵥵ļ��ܽ��ܳ���"
    key = "123456" ' ����һ���򵥵ļ�����Կ��ʵ��Ӧ����Ӧ�ø�����
End Sub

Private Sub cmdEncrypt_Click()
    ' ���ܰ�ť����¼�
    If txtPlain.Text <> "" Then
        txtEncrypted.Text = EncryptString(txtPlain.Text, key)
    Else
        MsgBox "��������Ҫ���ܵ��ı�", vbExclamation, "��ʾ"
    End If
End Sub

Private Sub cmdDecrypt_Click()
    ' ���ܰ�ť����¼�
    If txtEncrypted.Text <> "" Then
        txtPlain.Text = DecryptString(txtEncrypted.Text, key)
    Else
        MsgBox "��������Ҫ���ܵ��ı�", vbExclamation, "��ʾ"
    End If
End Sub

Private Function EncryptString(ByVal strText As String, ByVal strKey As String) As String
    ' �ַ������ܺ���
    ' ����ֻ��һ��ʾ����ʵ�ʼ��ܹ���Ӧ�ø�����
    Dim i As Integer
    For i = 1 To Len(strText)
        EncryptString = EncryptString & Chr(Asc(Mid(strText, i, 1)) + Asc(Mid(strKey, (i Mod Len(strKey)) + 1, 1)))
    Next i
End Function

Private Function DecryptString(ByVal strText As String, ByVal strKey As String) As String
    ' �ַ������ܺ���
    ' ����ֻ��һ��ʾ����ʵ�ʽ��ܹ���Ӧ�ø�����
    Dim i As Integer
    For i = 1 To Len(strText)
        DecryptString = DecryptString & Chr(Asc(Mid(strText, i, 1)) - Asc(Mid(strKey, (i Mod Len(strKey)) + 1, 1)))
    Next i
End Function
