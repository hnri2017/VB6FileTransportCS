VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFTServer 
   Caption         =   "FTServer"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   9450
   StartUpPosition =   3  '窗口缺省
   Begin FTServer.ProgressLabel ProgressLabel1 
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   7215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "游览…"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1440
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件另存为"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   435
      Width           =   900
   End
End
Attribute VB_Name = "frmFTServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'''Private Type mtypeTransmitReceiveVariant
'''    ReceivedFileState As Boolean
'''    ReceivedFileNumber As Integer
'''    ReceivedFileName As String
'''    ReceivedFileFolder As String
'''    ReceivedFilePath As String
'''    ReceivedFileTotalSize As Double
'''    ReceivedFileCompletedSize As Double
'''    ReceivedFileOverFlag As Boolean
'''End Type
'''为简化编程，接收与发送文件的一些变量不分开定义，程序中其它地方也注释掉
'''Private Type mtypeTransmitSendVariant
'''    SendFileState As Boolean
'''    SendFileName As String
'''    SendFileFolder As String
'''    SendFileTotalSize As Double
'''    SendFilePath As String
'''    SendFileNumber As Integer
'''    SendFileCompletedSize As Double
'''    SendFileOverFlag As Boolean
'''End Type

'这些自定义Type变量在客户端与服务端应保持一致，
'即客户端与嗠端的这些Type申明相同

'部分常量定义成Type类型，以便编程时不用记忆常量名
Private Type mtypeModelVariant
    DefaultIP As String '默认IP
    BufferSize As Long  '超大文件传输时分块的大小
    DefaultPort As Long '默认侦听端口号
    ConnectMax As Long  '最大连接数
    
    '自定义协议都是内定的一些字符串
    PtConnectedString As String '协议：连接
    PtFileName As String        '协议：文件名
    PtFileSize As String        '协议：文件大小（单位字节）
    PtFileSavePath As String    '协议：保存文件的文件夹名称，位于App.Path下面，且仅有一级目录
    PtFileTransmitReady As String   '协议：文件传输前准备指令
    PtFileTransmitStart As String   '协议：文件传输开始标识
    PtFileTransmitEndError As String    '协议：文件传输异常结束标识
    PtFileTransmitEndSuccess As String  '协议：文件传输正常结束标识
    
    PtErrFilePath As String     '协议：文件路径异常
    PtErrFileFolder As String   '协议：文件夹路径异常
    PtErrOverConnect As String  '协议：超出连接最大数
End Type

'定义文件传输时所用的Type变量
'注意，服务端接收则表示客户端是发送，
'反之，服务端是发送则客户端是接收
'客户端每次发送或接收文件之前都应将文件传输Type变量初始化，
'且发送与接收文件不可同时进行，因此发送文件时客户端软件的整个窗
'口（包括MDI与MDI子窗口）可用Enabled=False来控制,发送完毕后解除
Private Type mtypeTransmitVariant
    ReceivedFileState As Boolean    '文件开始接收标志
    ReceivedFileNumber As Integer   '文件操纵号
    ReceivedFileName As String      '文件名（不含有存储路径的）
    ReceivedFileFolder As String    '文件保存的文件夹名（不含路径）
    ReceivedFilePath As String      '文件名（这才是编程里的文件名，含全路径）
    ReceivedFileTotalSize As Double '文件大小，单位字节
    ReceivedFileCompletedSize As Double '文件已接收量
    ReceivedFileOverFlag As Boolean     '文件接收完成标识
    
    '以下文件发送变量的注释参考上面文件接收变量的注释，可以看出是重复的
    SendFileState As Boolean
    SendFileName As String
    SendFileFolder As String
    SendFileTotalSize As Double
    SendFilePath As String
    SendFileNumber As Integer
    SendFileCompletedSize As Double
    SendFileOverFlag As Boolean
End Type


Private gVar As mtypeModelVariant       '常量Type引用
Private gArr() As mtypeTransmitVariant '为每一个客户端连接建立的文件传输Type变量，下标为0的数组元素可专门用来初始化其它元素
'''Private gArrR() As mtypeTransmitReceiveVariant
'''Private gArrS() As mtypeTransmitSendVariant


Private Function mfDirFileKill(ByVal strPath As String) As Boolean
    '检测接收的文件是否存在同名，存在则删除
    Dim strFile As String
    
    On Error GoTo LineErr
        
    strFile = Dir(strPath, vbArchive + vbHidden + vbReadOnly + vbSystem + vbVolume)
    If Len(strFile) > 0 Then
        SetAttr strPath, vbNormal
        If FileLen(strPath) > 0 Then
            Kill strPath
        End If
    End If
    
    mfDirFileKill = True
    Exit Function
    
LineErr:
    Debug.Print strPath & "(" & Err.Number & ")" & Err.Description
End Function

Private Function mfDirFolder(ByVal strPath As String) As Boolean
    '检测文件夹是否存在，不存在则创建
    Dim strFolder As String
    
    On Error GoTo LineErr
        
    strFolder = Dir(strPath, vbArchive + vbDirectory + vbHidden + vbReadOnly + vbSystem + vbVolume)
    If Len(strFolder) = 0 Then
        MkDir strPath
    Else
        SetAttr strPath, vbNormal
    End If
    
    mfDirFolder = True
    Exit Function
    
LineErr:
    Debug.Print strPath & "(" & Err.Number & ")" & Err.Description
End Function

Private Sub msSendMessage(ByVal ctlIndex As Integer, ByVal strMsg As String)
    '发送消息的处理速度 好像 要 控制一下 ，延迟一下
    Winsock1.Item(ctlIndex).SendData strMsg
    DoEvents
'    Sleep 200
End Sub


Private Sub Command2_Click()
    '发送按钮
    
End Sub

Private Sub Form_Load()
    '初始化Type类型的通用常量
    With gVar
        .BufferSize = 5734  '好像没多大用，Winsock好像自动分块接收一次最多接8192B，无论你每次发多少数据过来
        .DefaultPort = 1361
        .DefaultIP = "127.0.0.1"
        .ConnectMax = 20
        
        .PtConnectedString = "[ConnectedIndex] = "
        .PtFileName = "[FileName] = "
        .PtFileSavePath = "[FilePath] = "
        .PtFileSize = "[FileSize] = "
        .PtFileTransmitStart = " [Start] "
        .PtFileTransmitReady = " [ReadyGo] "
        .PtFileTransmitEndError = " [EndError] "
        .PtFileTransmitEndSuccess = " [EndSuccess] "
        
        .PtErrFileFolder = "[Folder Error]="
        .PtErrFilePath = "[Path Error] = "
        .PtErrOverConnect = "[Connect Error] = "
    End With
    
    Text1.Text = ""
    Winsock1.Item(0).LocalPort = gVar.DefaultPort
    Winsock1.Item(0).Listen '0元素侦听，其它元素建立连接
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Winsock1.Item(0).State <> sckClosed Then
        Winsock1.Item(0).Close
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
    If Index <> 0 Then
'Debug.Print "Try To Close Connect Winsock1.Item(" & Index & ") " & Winsock1.Item(Index).RemoteHostIP & "."
        Unload Winsock1.Item(Index) '连接关闭后清除对应的Winsock实例
        gArr(Index) = gArr(0)     '清空对应的文件传输Type变量
'''        gArrR(Index) = gArrR(0)
'''        gArrS(Index) = gArrS(0)
Debug.Print "Close Connect Winsock1.Item(" & Index & ") Finished."
    End If
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer, U As Integer
    Dim ctlSock As MSWinsockLib.Winsock
    
    If Index = 0 Then
        
'Debug.Print "Now The Count OF Connect Is " & Winsock1.Count
        '用循环找出当前Winsock实例中可以用来建立连接的数组元素最小下标值
        For Each ctlSock In Winsock1    '此处Winsock1是控件集合
            If ctlSock.Index = K Then
                K = K + 1
            Else
                Exit For
            End If
        Next
        Load Winsock1(K)    '加载可以建立连接的实例
        With Winsock1.Item(K)
            .LocalPort = gVar.DefaultPort
            .Accept requestID   '建立连接
            Call msSendMessage(K, gVar.PtConnectedString & K)    '发送连接成功标识
Debug.Print "Apply IP " & .RemoteHostIP & " ,RequestID:" & requestID, K
        End With
        
        U = Winsock1.UBound
        ReDim Preserve gArr(0 To U) '生成对应的文件传输Type数组元素
        gArr(K) = gArr(0)           '初始化元素，防止已被用过
'''        ReDim Preserve gArrR(0 To U)
'''        gArrR(K) = gArrR(0)
'''        ReDim Preserve gArrS(0 To U)
'''        gArrS(K) = gArrS(0)
        
        '限制连接数量。不知道能承受多少？
        If U > gVar.ConnectMax Then
            Call msSendMessage(K, gVar.PtErrOverConnect & "建立连接失败，最多只能连接" & gVar.ConnectMax & "人！")
            Call Winsock1_Close(K)
        End If
        
    End If
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, strTemp As String
    Dim lngInStr As Long
    Static lngByteOver As Long
    
'    On Error Resume Next    '服务端有错应怎么处理？
    
    If Index = 0 Then Exit Sub
    If Winsock1.Item(Index).State <> sckConnected Then Exit Sub
    
    With gArr(Index)
        '接收类型：要么字符串、要么文件
        If Not .ReceivedFileState Then  '非接收文件状态，即协议字符串 + 对应值
            
            If bytesTotal = 8192 Then   '测试时发现传输文件时也会进入此处，不知是何原因，暂如此处理下
                lngByteOver = lngByteOver + 1
                If lngByteOver > 3 Then
                    lngByteOver = 0
                    Exit Sub
                End If
            Else
                lngByteOver = 0
            End If
            
            With Winsock1.Item(Index)
                .GetData strData
Debug.Print strData, Index, bytesTotal, .RemoteHost, .RemoteHostIP, .RemotePort
            End With
            
            lngInStr = InStr(strData, gVar.PtFileTransmitReady) '文件接收准备
            If lngInStr = 1 Then
                gArr(Index) = gArr(0)   '初始化
                Exit Sub
            End If
            
            lngInStr = InStr(strData, gVar.PtFileName)  '接收到的文件名
            If lngInStr = 1 Then
                .ReceivedFileName = Mid(strData, lngInStr + Len(gVar.PtFileName))
                Exit Sub
            End If
            
            lngInStr = InStr(strData, gVar.PtFileSize)  '接收到的文件大小
            If lngInStr = 1 Then
                strTemp = Mid(strData, lngInStr + Len(gVar.PtFileSize))
                If IsNumeric(strTemp) Then .ReceivedFileTotalSize = CDbl(strTemp)
                Exit Sub
            End If
            
            '接收到的文件夹名称。此程序只在接收到文件夹后发送接收标识，可以更改到别处。
            lngInStr = InStr(strData, gVar.PtFileSavePath)
            If lngInStr = 1 Then
                .ReceivedFileFolder = Mid(strData, lngInStr + Len(gVar.PtFileSavePath))
                
                '以下为几处检查，OK则发送接收标识
                If Not (Len(.ReceivedFileName) > 0 And Len(.ReceivedFileFolder) > 0) Then
                    Exit Sub    '文件名或文件夹至少有一个未告知
                End If
                
                strTemp = App.Path & "\" & .ReceivedFileFolder
                If Not mfDirFolder(strTemp) Then    '发送文件夹异常
                    Call msSendMessage(Index, gVar.PtErrFileFolder & .ReceivedFileFolder)
                    Exit Sub
                End If
                
                .ReceivedFilePath = strTemp & "\" & .ReceivedFileName
                If Not mfDirFileKill(.ReceivedFilePath) Then    '发送文件异常
                    Call msSendMessage(Index, gVar.PtErrFilePath & .ReceivedFileName)
                    Exit Sub
                End If

                .ReceivedFileState = True
                With ProgressLabel1
                    .Value = 0
                    .Min = 0
                    .Max = gArr(Index).ReceivedFileTotalSize
                End With
                Call msSendMessage(Index, gVar.PtFileTransmitStart)   '发送接收标识
                Exit Sub
                
            End If

        Else
        
            On Error Resume Next
            
            If .ReceivedFileNumber = 0 Then     '文件第一次接收应打开文件，分块接收的第二次起就不用再打开了
                .ReceivedFileNumber = FreeFile
                Open .ReceivedFilePath For Binary As #.ReceivedFileNumber
            End If
            
            Dim byteGotData() As Byte
            
            ReDim byteGotData(0 To bytesTotal - 1)                      '定义接收文件的字节数组大小
            Winsock1.Item(Index).GetData byteGotData, vbArray + vbByte  '接收文件（块）
            Put #.ReceivedFileNumber, , byteGotData                     '缓存到打开的文件号中
            .ReceivedFileCompletedSize = .ReceivedFileCompletedSize + bytesTotal    '记录已接收到的文件大小
            ProgressLabel1.Value = .ReceivedFileCompletedSize
            
            If .ReceivedFileCompletedSize >= .ReceivedFileTotalSize Then    '检测是否接收完成
                Close #.ReceivedFileNumber
Debug.Print .ReceivedFileName & " Received Over."
                '发送文件接收结束指令
                If Err.Number = 0 Then
                    ProgressLabel1.Value = gArr(Index).ReceivedFileTotalSize
                    Call msSendMessage(Index, gVar.PtFileTransmitEndSuccess)
                Else
                    Call msSendMessage(Index, gVar.PtFileTransmitEndError & "(" & Err.Number & ")" & Err.Description)
                End If
                .ReceivedFileOverFlag = True
                .ReceivedFileState = False
            End If
        
        End If
        
    End With
    
End Sub

