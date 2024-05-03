VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "加密解密"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   3135
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "解密"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "加密"
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
      Caption         =   "加密："
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "文本："
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
      Caption         =   "简单的加密解密程序"
      BeginProperty Font 
         Name            =   "宋体"
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
' 加密和解密程序的VB代码示例

' 定义全局变量
Dim key As String

Private Sub Form_Load()
    ' 窗体加载时的设置
    Me.Caption = "简单的加密解密程序"
    key = "123456" ' 这是一个简单的加密密钥，实际应用中应该更复杂
End Sub

Private Sub cmdEncrypt_Click()
    ' 加密按钮点击事件
    If txtPlain.Text <> "" Then
        txtEncrypted.Text = EncryptString(txtPlain.Text, key)
    Else
        MsgBox "请输入需要加密的文本", vbExclamation, "提示"
    End If
End Sub

Private Sub cmdDecrypt_Click()
    ' 解密按钮点击事件
    If txtEncrypted.Text <> "" Then
        txtPlain.Text = DecryptString(txtEncrypted.Text, key)
    Else
        MsgBox "请输入需要解密的文本", vbExclamation, "提示"
    End If
End Sub

Private Function EncryptString(ByVal strText As String, ByVal strKey As String) As String
    ' 字符串加密函数
    ' 这里只是一个示例，实际加密过程应该更复杂
    Dim i As Integer
    For i = 1 To Len(strText)
        EncryptString = EncryptString & Chr(Asc(Mid(strText, i, 1)) + Asc(Mid(strKey, (i Mod Len(strKey)) + 1, 1)))
    Next i
End Function

Private Function DecryptString(ByVal strText As String, ByVal strKey As String) As String
    ' 字符串解密函数
    ' 这里只是一个示例，实际解密过程应该更复杂
    Dim i As Integer
    For i = 1 To Len(strText)
        DecryptString = DecryptString & Chr(Asc(Mid(strText, i, 1)) - Asc(Mid(strKey, (i Mod Len(strKey)) + 1, 1)))
    Next i
End Function
