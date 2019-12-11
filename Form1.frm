VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ModBus CRC16"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option1 
      Caption         =   "字符"
      Height          =   285
      Index           =   1
      Left            =   2070
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Caption         =   "十六进制"
      Height          =   285
      Index           =   0
      Left            =   630
      TabIndex        =   3
      Top             =   1890
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算ModBus CRC16"
      Height          =   420
      Left            =   4410
      TabIndex        =   2
      Top             =   1755
      Width           =   2310
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   1
      Left            =   1350
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2475
      Width           =   6810
   End
   Begin VB.TextBox Text1 
      Height          =   1140
      Index           =   0
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   270
      Width           =   7755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "[LoHi]"
      Height          =   510
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long
        Dim j1() As Byte
    If Option1(0).Value = True Then
        Dim k1() As String


        Text1(0) = Replace(Text1(0), ",", " ")
        Text1(0) = Replace(Text1(0), "  ", " ")
        
        k1() = Split(Text1(0))
        
        ReDim j1(UBound(k1))
        
        For i = 0 To UBound(k1)
            j1(i) = CByte(k1(i))
        Next
        

    ElseIf Option1(1).Value = True Then

        Dim j As Long
        Dim k As String

        Dim m As Long
        m = 0
        i = Len(Text1(0)) '获取文本总长度,
        ReDim j1(LenB(Text1(0))) '定义数组维数,汉字2字节
        Debug.Print LenB(Text1(0))
        For j = 1 To i
            k = Mid(Text1(0), j, 1)
            k = Right("0000" & CStr(Hex(Asc(k))), 4)
            
            If LenB(k) = 2 Then
                j1(m) = CByte("&h" & Left(CInt(k), 2))
                m = m + 1
            End If
                
            j1(m) = CByte("&h" & Right(k, 2))
            m = m + 1
        Next
    End If
    Text1(1) = CRC16(j1())
End Sub

Function CRC16(data() As Byte) As String 'CRC计算函数
    Dim CRC16Lo As Byte, CRC16Hi As Byte   'CRC寄存器
    Dim CL As Byte, CH As Byte            '多项式码&HA001
    Dim SaveHi As Byte, SaveLo As Byte
    Dim i As Integer
    Dim Flag As Integer
    CRC16Lo = &HFF
    CRC16Hi = &HFF
    CL = &H1
    CH = &HA0
    For i = 0 To UBound(data)
        CRC16Lo = CRC16Lo Xor data(i) '每一个数据与CRC寄存器进行异或
        For Flag = 0 To 7
            SaveHi = CRC16Hi
            SaveLo = CRC16Lo
            CRC16Hi = CRC16Hi \ 2            '高位右移一位
            CRC16Lo = CRC16Lo \ 2            '低位右移一位
            If ((SaveHi And &H1) = &H1) Then '如果高位字节最后一位为1
                CRC16Lo = CRC16Lo Or &H80      '则低位字节右移后前面补1
            End If                           '否则自动补0
            If ((SaveLo And &H1) = &H1) Then '如果LSB为1，则与多项式码进行异或
                CRC16Hi = CRC16Hi Xor CH
                CRC16Lo = CRC16Lo Xor CL
            End If
        Next
    Next
    CRC16 = "0x" & Right("00" + Hex(CRC16Lo), 2) + Right("00" + Hex(CRC16Hi), 2)
End Function
