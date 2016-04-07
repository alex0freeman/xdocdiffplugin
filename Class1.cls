VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinMergeScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const MOVEFILE_REPLACE_EXISTING = &H1

Private Declare Function GetTempPath Lib "kernel32" _
    Alias "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function OpenProcess Lib "kernel32" ( _
                                             ByVal dwDesiredAccess As Long, _
                                             ByVal bInheritHandle As Long, _
                                             ByVal dwProcessId As Long) As Long


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal FileName As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long

Public Property Get PluginEvent() As String
    PluginEvent = "FILE_PACK_UNPACK"
End Property

Public Property Get PluginDescription() As String
    PluginDescription = "Display the text content of MS Word, Excel, PowerPoint and pdf files."
End Property

Public Property Get PluginFileFilters() As String
    PluginFileFilters = "\.sxw;\.sxc;\.sxi;\.sxd;\.odt;\.ods;\.odp;\.odg;\.docx;\.docm;\.xlsx;\.xlsm;\.pptx;\.pptm;\.doc;\.xls;\.ppt;\.rtf;\.jaw;\.jtw;\.jbw;\.juw;\.jfw;\.jvw;\.jtd;\.jtt;\.oas;\.oa2;\.oa3;\.bun;\.wj2;\.wj3;\.wk3;\.wk4;\.123;\.wri;\.pdf;\.mht;\.eml$"
End Property

Public Property Get PluginIsAutomatic() As Boolean
    PluginIsAutomatic = True
End Property

Public Function UnpackFile(fileSrc As String, fileDst As String, ByRef bChanged As Boolean, ByRef subcode As Long) As Boolean
'    MsgBox "fileSrc: " & fileSrc & vbCrLf & "fileDst: " & fileDst
    
    ' １．対象ファイルをテンポラリディレクトリにコピーする
    Dim TempPath As String * 1000
    Dim fileCopied As Boolean
    fileCopied = False

    'テンポラリディレクトリの取得
    GetTempPath 1000, TempPath
    
    Dim i As Integer
    Dim FileName As String

    i = InStrRev(fileSrc, "\")
    If i > 0 Then
        FileName = Right(fileSrc, Len(fileSrc) - i)
    Else
        FileName = fileSrc
    End If
    
    Dim tempFilename As String
    tempFilename = Left(TempPath, InStr(1, TempPath, vbNullChar) - 1)
    tempFilename = tempFilename & FileName
    
' MsgBox "ファイルをコピーします。" & vbCrLf & "From: " & fileSrc & vbCrLf & "Dest: " & tempFilename

    ' ファイルをテンポラリディレクトリにコピーするかどうかを判定
    Dim Fso As New FileSystemObject
            
    If Fso.FileExists(fileSrc) And Fso.FileExists(tempFilename) Then
        ' 一時的にコピーする先のファイルが、すでに存在している
    
        Dim FileObjSrc As File
        Dim FileObjDest As File
        
        Set FileObjSrc = Fso.GetFile(fileSrc)
        Set FileObjDest = Fso.GetFile(tempFilename)
    
        If FileObjSrc.ShortPath = FileObjDest.ShortPath Then
'            MsgBox "ファイルをコピーする必要はありません。"
        Else
 '           MsgBox "ファイルをコピーします。"
            
            fileCopied = True
        End If
    Else
        fileCopied = True
    End If
    
    If fileCopied Then
        Call Fso.CopyFile(fileSrc, tempFilename)
    
       ' 元ファイルが読み取り専用の場合は解除する
        If Fso.GetFile(tempFilename).Attributes And 1 Then
            Fso.GetFile(tempFilename).Attributes = 0
        End If
    End If
    
    Dim xdoc2txtParam As String
    
    xdoc2txtParam = tempFilename
    
    i = InStr(1, xdoc2txtParam, " ")
    If i > 0 Then
        xdoc2txtParam = """" & xdoc2txtParam & """"
    End If
    
    ' ２．xdoc2txt でテキストファイルに変換
    Dim pid As Long
    Dim ph As Long
    Dim ecode As Long
    
    pid = Shell("xdoc2txt -f " & xdoc2txtParam, vbHide)
    ph = OpenProcess(SYNCHRONIZE Or PROCESS_QUERY_INFORMATION, True, pid)
    WaitForSingleObject ph, 100000

    GetExitCodeProcess ph, ecode  ' 終了コード取得
    CloseHandle ph ' プロセスハンドルを閉じる
    
    ' ３．変換されたテキストファイルを、WinMergeが指定しているファイルに移動する
    Dim temporaryTextFile As String
    temporaryTextFile = tempFilename
    temporaryTextFile = Left(temporaryTextFile, InStrRev(temporaryTextFile, ".")) & "txt"
    
    Call MoveFileEx(temporaryTextFile, fileDst, MOVEFILE_REPLACE_EXISTING)
    
    If fileCopied Then
        Call DeleteFile(tempFilename)
    End If
    
    bChanged = True
    UnpackFile = True
    subcode = 1
    
End Function
 
Public Function PackFile(fileSrc As String, fileDst As String, ByRef bChanged As Boolean, subcode As Long) As Boolean
    ' We can't repack files
    bChanged = False
    PackFile = False
    subcode = 1
End Function

