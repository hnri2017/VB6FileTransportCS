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
   StartUpPosition =   3  '����ȱʡ
   Begin FTServer.ProgressLabel ProgressLabel1 
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "������"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
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
      Caption         =   "�ļ����Ϊ"
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
'''Ϊ�򻯱�̣������뷢���ļ���һЩ�������ֿ����壬�����������ط�Ҳע�͵�
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

'��Щ�Զ���Type�����ڿͻ���������Ӧ����һ�£�
'���ͻ������˵���ЩType������ͬ

'���ֳ��������Type���ͣ��Ա���ʱ���ü��䳣����
Private Type mtypeModelVariant
    DefaultIP As String 'Ĭ��IP
    BufferSize As Long  '�����ļ�����ʱ�ֿ�Ĵ�С
    DefaultPort As Long 'Ĭ�������˿ں�
    ConnectMax As Long  '���������
    
    '�Զ���Э�鶼���ڶ���һЩ�ַ���
    PtConnectedString As String 'Э�飺����
    PtFileName As String        'Э�飺�ļ���
    PtFileSize As String        'Э�飺�ļ���С����λ�ֽڣ�
    PtFileSavePath As String    'Э�飺�����ļ����ļ������ƣ�λ��App.Path���棬�ҽ���һ��Ŀ¼
    PtFileTransmitReady As String   'Э�飺�ļ�����ǰ׼��ָ��
    PtFileTransmitStart As String   'Э�飺�ļ����俪ʼ��ʶ
    PtFileTransmitEndError As String    'Э�飺�ļ������쳣������ʶ
    PtFileTransmitEndSuccess As String  'Э�飺�ļ���������������ʶ
    
    PtErrFilePath As String     'Э�飺�ļ�·���쳣
    PtErrFileFolder As String   'Э�飺�ļ���·���쳣
    PtErrOverConnect As String  'Э�飺�������������
End Type

'�����ļ�����ʱ���õ�Type����
'ע�⣬����˽������ʾ�ͻ����Ƿ��ͣ�
'��֮��������Ƿ�����ͻ����ǽ���
'�ͻ���ÿ�η��ͻ�����ļ�֮ǰ��Ӧ���ļ�����Type������ʼ����
'�ҷ���������ļ�����ͬʱ���У���˷����ļ�ʱ�ͻ��������������
'�ڣ�����MDI��MDI�Ӵ��ڣ�����Enabled=False������,������Ϻ���
Private Type mtypeTransmitVariant
    ReceivedFileState As Boolean    '�ļ���ʼ���ձ�־
    ReceivedFileNumber As Integer   '�ļ����ݺ�
    ReceivedFileName As String      '�ļ����������д洢·���ģ�
    ReceivedFileFolder As String    '�ļ�������ļ�����������·����
    ReceivedFilePath As String      '�ļ���������Ǳ������ļ�������ȫ·����
    ReceivedFileTotalSize As Double '�ļ���С����λ�ֽ�
    ReceivedFileCompletedSize As Double '�ļ��ѽ�����
    ReceivedFileOverFlag As Boolean     '�ļ�������ɱ�ʶ
    
    '�����ļ����ͱ�����ע�Ͳο������ļ����ձ�����ע�ͣ����Կ������ظ���
    SendFileState As Boolean
    SendFileName As String
    SendFileFolder As String
    SendFileTotalSize As Double
    SendFilePath As String
    SendFileNumber As Integer
    SendFileCompletedSize As Double
    SendFileOverFlag As Boolean
End Type


Private gVar As mtypeModelVariant       '����Type����
Private gArr() As mtypeTransmitVariant 'Ϊÿһ���ͻ������ӽ������ļ�����Type�������±�Ϊ0������Ԫ�ؿ�ר��������ʼ������Ԫ��
'''Private gArrR() As mtypeTransmitReceiveVariant
'''Private gArrS() As mtypeTransmitSendVariant


Private Function mfDirFileKill(ByVal strPath As String) As Boolean
    '�����յ��ļ��Ƿ����ͬ����������ɾ��
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
    '����ļ����Ƿ���ڣ��������򴴽�
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
    '������Ϣ�Ĵ����ٶ� ���� Ҫ ����һ�� ���ӳ�һ��
    Winsock1.Item(ctlIndex).SendData strMsg
    DoEvents
'    Sleep 200
End Sub


Private Sub Command2_Click()
    '���Ͱ�ť
    
End Sub

Private Sub Form_Load()
    '��ʼ��Type���͵�ͨ�ó���
    With gVar
        .BufferSize = 5734  '����û����ã�Winsock�����Զ��ֿ����һ������8192B��������ÿ�η��������ݹ���
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
    Winsock1.Item(0).Listen '0Ԫ������������Ԫ�ؽ�������
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Winsock1.Item(0).State <> sckClosed Then
        Winsock1.Item(0).Close
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
    If Index <> 0 Then
'Debug.Print "Try To Close Connect Winsock1.Item(" & Index & ") " & Winsock1.Item(Index).RemoteHostIP & "."
        Unload Winsock1.Item(Index) '���ӹرպ������Ӧ��Winsockʵ��
        gArr(Index) = gArr(0)     '��ն�Ӧ���ļ�����Type����
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
        '��ѭ���ҳ���ǰWinsockʵ���п��������������ӵ�����Ԫ����С�±�ֵ
        For Each ctlSock In Winsock1    '�˴�Winsock1�ǿؼ�����
            If ctlSock.Index = K Then
                K = K + 1
            Else
                Exit For
            End If
        Next
        Load Winsock1(K)    '���ؿ��Խ������ӵ�ʵ��
        With Winsock1.Item(K)
            .LocalPort = gVar.DefaultPort
            .Accept requestID   '��������
            Call msSendMessage(K, gVar.PtConnectedString & K)    '�������ӳɹ���ʶ
Debug.Print "Apply IP " & .RemoteHostIP & " ,RequestID:" & requestID, K
        End With
        
        U = Winsock1.UBound
        ReDim Preserve gArr(0 To U) '���ɶ�Ӧ���ļ�����Type����Ԫ��
        gArr(K) = gArr(0)           '��ʼ��Ԫ�أ���ֹ�ѱ��ù�
'''        ReDim Preserve gArrR(0 To U)
'''        gArrR(K) = gArrR(0)
'''        ReDim Preserve gArrS(0 To U)
'''        gArrS(K) = gArrS(0)
        
        '����������������֪���ܳ��ܶ��٣�
        If U > gVar.ConnectMax Then
            Call msSendMessage(K, gVar.PtErrOverConnect & "��������ʧ�ܣ����ֻ������" & gVar.ConnectMax & "�ˣ�")
            Call Winsock1_Close(K)
        End If
        
    End If
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, strTemp As String
    Dim lngInStr As Long
    Static lngByteOver As Long
    
'    On Error Resume Next    '������д�Ӧ��ô����
    
    If Index = 0 Then Exit Sub
    If Winsock1.Item(Index).State <> sckConnected Then Exit Sub
    
    With gArr(Index)
        '�������ͣ�Ҫô�ַ�����Ҫô�ļ�
        If Not .ReceivedFileState Then  '�ǽ����ļ�״̬����Э���ַ��� + ��Ӧֵ
            
            If bytesTotal = 8192 Then   '����ʱ���ִ����ļ�ʱҲ�����˴�����֪�Ǻ�ԭ������˴�����
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
            
            lngInStr = InStr(strData, gVar.PtFileTransmitReady) '�ļ�����׼��
            If lngInStr = 1 Then
                gArr(Index) = gArr(0)   '��ʼ��
                Exit Sub
            End If
            
            lngInStr = InStr(strData, gVar.PtFileName)  '���յ����ļ���
            If lngInStr = 1 Then
                .ReceivedFileName = Mid(strData, lngInStr + Len(gVar.PtFileName))
                Exit Sub
            End If
            
            lngInStr = InStr(strData, gVar.PtFileSize)  '���յ����ļ���С
            If lngInStr = 1 Then
                strTemp = Mid(strData, lngInStr + Len(gVar.PtFileSize))
                If IsNumeric(strTemp) Then .ReceivedFileTotalSize = CDbl(strTemp)
                Exit Sub
            End If
            
            '���յ����ļ������ơ��˳���ֻ�ڽ��յ��ļ��к��ͽ��ձ�ʶ�����Ը��ĵ��𴦡�
            lngInStr = InStr(strData, gVar.PtFileSavePath)
            If lngInStr = 1 Then
                .ReceivedFileFolder = Mid(strData, lngInStr + Len(gVar.PtFileSavePath))
                
                '����Ϊ������飬OK���ͽ��ձ�ʶ
                If Not (Len(.ReceivedFileName) > 0 And Len(.ReceivedFileFolder) > 0) Then
                    Exit Sub    '�ļ������ļ���������һ��δ��֪
                End If
                
                strTemp = App.Path & "\" & .ReceivedFileFolder
                If Not mfDirFolder(strTemp) Then    '�����ļ����쳣
                    Call msSendMessage(Index, gVar.PtErrFileFolder & .ReceivedFileFolder)
                    Exit Sub
                End If
                
                .ReceivedFilePath = strTemp & "\" & .ReceivedFileName
                If Not mfDirFileKill(.ReceivedFilePath) Then    '�����ļ��쳣
                    Call msSendMessage(Index, gVar.PtErrFilePath & .ReceivedFileName)
                    Exit Sub
                End If

                .ReceivedFileState = True
                With ProgressLabel1
                    .Value = 0
                    .Min = 0
                    .Max = gArr(Index).ReceivedFileTotalSize
                End With
                Call msSendMessage(Index, gVar.PtFileTransmitStart)   '���ͽ��ձ�ʶ
                Exit Sub
                
            End If

        Else
        
            On Error Resume Next
            
            If .ReceivedFileNumber = 0 Then     '�ļ���һ�ν���Ӧ���ļ����ֿ���յĵڶ�����Ͳ����ٴ���
                .ReceivedFileNumber = FreeFile
                Open .ReceivedFilePath For Binary As #.ReceivedFileNumber
            End If
            
            Dim byteGotData() As Byte
            
            ReDim byteGotData(0 To bytesTotal - 1)                      '��������ļ����ֽ������С
            Winsock1.Item(Index).GetData byteGotData, vbArray + vbByte  '�����ļ����飩
            Put #.ReceivedFileNumber, , byteGotData                     '���浽�򿪵��ļ�����
            .ReceivedFileCompletedSize = .ReceivedFileCompletedSize + bytesTotal    '��¼�ѽ��յ����ļ���С
            ProgressLabel1.Value = .ReceivedFileCompletedSize
            
            If .ReceivedFileCompletedSize >= .ReceivedFileTotalSize Then    '����Ƿ�������
                Close #.ReceivedFileNumber
Debug.Print .ReceivedFileName & " Received Over."
                '�����ļ����ս���ָ��
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

