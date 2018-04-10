VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFTClient 
   Caption         =   "FTClient"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   10350
   StartUpPosition =   3  '����ȱʡ
   Begin FTClient.ProgressLabel ProgressLabel1 
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
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
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1920
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���ն�IP"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   435
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�ѡ��"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   1275
      Width           =   720
   End
End
Attribute VB_Name = "frmFTClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type mtypeModelVariant
    DefaultIP As String
    BufferSize As Long
    DefaultPort As Long
    MyConnectedIndex As Long
    
    PtConnectedString As String
    PtFileName As String
    PtFileSize As String
    PtFileSavePath As String
    PtFileTransmitReady As String   'Э�飺�ļ�����ǰ׼��ָ��
    PtFileTransmitStart As String
    PtFileTransmitEndError As String    'Э�飺�ļ������쳣������ʶ
    PtFileTransmitEndSuccess As String  'Э�飺�ļ���������������ʶ
    
    PtErrFilePath As String     'Э�飺�ļ�·���쳣
    PtErrFileFolder As String   'Э�飺�ļ���·���쳣
    PtErrOverConnect As String  'Э�飺�������������
End Type

Private Type mtypeFileTransmitVariant
    SendFileState As Boolean
    SendFileName As String
    SendFileFolder As String
    SendFileTotalSize As Double
    SendFilePath As String
    SendFileNumber As Integer
    SendFileCompletedSize As Double
    SendFileOverFlag As Boolean
    
    ReceivedFileState As Boolean
    ReceivedFileNumber As Integer
    ReceivedFileName As String
    ReceivedFileFolder As String
    ReceivedFilePath As String
    ReceivedFileTotalSize As Double
    ReceivedFileCompletedSize As Double
    ReceivedFileOverFlag As Boolean
End Type

Private gVar As mtypeModelVariant
Private gFile As mtypeFileTransmitVariant
Dim lngVal As Long

Private Function mfDirFile(ByVal strPath As String) As Boolean
    Dim strTemp As String
    
    On Error GoTo LineErr
    
    strTemp = Dir(strPath, vbHidden + vbReadOnly)
    If Len(strTemp) = 0 Then Exit Function
    
    If GetAttr(strPath) = vbHidden Then SetAttr strPath, vbNormal
    mfDirFile = True
    Exit Function
    
LineErr:
    Debug.Print strPath & "(" & Err.Number & ")" & Err.Description
End Function

Private Function mfInStrData(ByVal strDt As String, ByVal strIn As String) As Boolean
    '���յ��ַ���ָ��Э��Ĵ���
    Dim lngInStr As Long
    Dim strTemp As String
    
    lngInStr = InStr(strDt, strIn)
    If lngInStr = 1 Then
        With gVar
            Select Case strIn
                Case .PtConnectedString
                    strTemp = Mid(strDt, lngInStr + Len(strIn))
                    If IsNumeric(strTemp) Then
                        .MyConnectedIndex = CLng(strTemp)
                    Else
                        Exit Function
                    End If
                    
                Case .PtErrFileFolder, .PtErrFilePath, .PtErrOverConnect
                    MsgBox strDt, vbCritical
                    
                Case .PtFileName
                
                Case .PtFileSavePath
                
                Case .PtFileSize
                
                Case .PtFileTransmitEndError
                    MsgBox "�ļ��ϴ�ʧ�ܣ�", vbCritical
                    
                Case .PtFileTransmitEndSuccess
                    ProgressLabel1.Value = ProgressLabel1.Max
                    MsgBox "�ļ��ϴ��ɹ���", vbInformation
                    
                Case .PtFileTransmitReady
                
                Case .PtFileTransmitStart
                    Me.Enabled = False
                    Command2.Visible = False
                    gFile.SendFileState = True
                    With ProgressLabel1
                        .Value = 0
                        .Max = gFile.SendFileTotalSize
                        .Min = 0
                    End With
                    Call msSendFileChunk
                Case Else
                    Exit Function
            End Select
        End With
        mfInStrData = True
    End If
End Function

Private Sub msSendFileChunk()
    If Winsock1.State <> sckConnected Then Exit Sub
    
    With gFile
        If .SendFileNumber = 0 Then
            .SendFileNumber = FreeFile
            Open .SendFilePath For Binary As #.SendFileNumber
        End If
        
        Dim lngChunkSize As Long
        Dim byteSendData() As Byte
        Dim lngLocation As Double
        
        lngChunkSize = gVar.BufferSize
        lngLocation = LOF(.SendFileNumber) - Loc(.SendFileNumber)
        If lngLocation < lngChunkSize Then lngChunkSize = lngLocation
        ReDim byteSendData(0 To lngChunkSize - 1)
        
        Get #.SendFileNumber, , byteSendData
        Winsock1.SendData byteSendData
        .SendFileCompletedSize = .SendFileCompletedSize + lngChunkSize
        ProgressLabel1.Value = gFile.SendFileCompletedSize
        
        If .SendFileCompletedSize >= .SendFileTotalSize Then
            Close #.SendFileNumber
            .SendFileOverFlag = True
        End If
        
    End With
    
End Sub

Private Sub msSendMessage(ByVal strMsg As String)
    '�������Ͷ������ַ�����Ϣʱ�ٶȺ���Ҫ����һ�£��������ʱ������ͬʱ����
    Winsock1.SendData strMsg
    DoEvents
    Sleep 300
End Sub



Private Sub Command1_Click()

    With CommonDialog1
        .DialogTitle = "ѡ���ļ�"
        .Filter = "All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End With
    
    Text2.Text = CommonDialog1.FileName
    
End Sub

Private Sub Command2_Click()
    With Winsock1
        If .State <> sckConnected Then
            MsgBox "���ӷ����ʧ�ܣ��޷�" & Command2.Caption & "��", vbCritical
        Else
            
            Dim typeFile As mtypeFileTransmitVariant
            Dim strFile As String
            
            strFile = Trim(Text2.Text)
            If Not mfDirFile(strFile) Then Exit Sub
            
            gFile = typeFile    '��ʼ���Զ�����ļ���Ϣ����
            With gFile
                .SendFilePath = strFile
                .SendFileName = Mid(strFile, 1 + InStrRev(strFile, "\"))
                .SendFileTotalSize = FileLen(strFile)
                .SendFileFolder = "TempFile"
            End With
'            If gFile.SendFileTotalSize > 52428800 Then
'                MsgBox "�����ϴ�����50M���ļ���", vbExclamation
'                gFile = typeFile
'                Exit Sub
'            End If
            Call msSendMessage(gVar.PtFileTransmitReady)
            Call msSendMessage(gVar.PtFileName & gFile.SendFileName)
            Call msSendMessage(gVar.PtFileSize & gFile.SendFileTotalSize)
            Call msSendMessage(gVar.PtFileSavePath & gFile.SendFileFolder)

        End If
    End With
End Sub

Private Sub Command3_Click()
    With Winsock1
        If .State <> sckClosed Then .Close
        .RemoteHost = Text1.Text
        .RemotePort = gVar.DefaultPort
        .Connect
        If .State = sckError Then
            MsgBox "Connect Server Fail!"
            Exit Sub
        End If
        
'Debug.Print .LocalHostName & "--" & .DefaultIP & "--" & .LocalPort & "--" & .Name
'Debug.Print .RemoteHost & "--" & .RemoteHostIP & "--" & .RemotePort & "--" & .State & "--" & .Tag
    End With
    
End Sub

Private Sub Form_Load()
    
    With gVar
        .BufferSize = 5734
        .DefaultPort = 1361
        .DefaultIP = "192.168.2.108"
        .PtConnectedString = "[ConnectedIndex] = "
        .PtFileName = "[FileName] = "
        .PtFileSavePath = "[FilePath] = "
        .PtFileSize = "[FileSize] = "
        .PtFileTransmitStart = " [Start] "
        .PtFileTransmitReady = " [ReadyGo] "
        .PtFileTransmitEndError = " [EndError] "
        .PtFileTransmitEndSuccess = " [EndSuccess] "
        
        .PtErrFileFolder = "[Folder Error] = "
        .PtErrFilePath = "[Path Error] = "
        .PtErrOverConnect = "[Connect Error] = "
    End With
    
    Text1.Text = gVar.DefaultIP
    Text2.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Winsock1_Close
End Sub

Private Sub Winsock1_Close()
    With Winsock1
        If .State <> sckClosed Then
            .Close
            .LocalPort = 0
Debug.Print "Winsock1 Close"
        End If
    End With

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strTemp As String
    Dim lngInStr As Long
    
    If Not gFile.ReceivedFileState Then
        '�ַ�������
        
        Winsock1.GetData strData
Debug.Print strData, bytesTotal
        
        If mfInStrData(strData, gVar.PtFileTransmitEndSuccess) Then Exit Sub
        If mfInStrData(strData, gVar.PtFileTransmitEndError) Then Exit Sub
        If mfInStrData(strData, gVar.PtConnectedString) Then Exit Sub
        If mfInStrData(strData, gVar.PtFileTransmitStart) Then Exit Sub
        If mfInStrData(strData, gVar.PtErrFileFolder) Then Exit Sub
        If mfInStrData(strData, gVar.PtErrFilePath) Then Exit Sub
        If mfInStrData(strData, gVar.PtErrOverConnect) Then Exit Sub
        
    Else
        '�ļ�����
        
    End If
    
End Sub

Private Sub Winsock1_SendComplete()
    DoEvents
    
    If gFile.SendFileState Then
        With gFile
            If .SendFileOverFlag Then
Debug.Print .SendFileName & " Send Over."
                Dim typeFile As mtypeFileTransmitVariant
                gFile = typeFile
                Command2.Visible = True
                Me.Enabled = True
'                Timer1.Enabled = False
            Else
                Call msSendFileChunk
            End If
        End With
    End If
    
End Sub
