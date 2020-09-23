VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Files As New Collection
Dim totalFiles As Integer

Dim Downloading As Boolean
Dim Uploading As Boolean

Dim x As Integer

Public Event DownloadsComplete(NumFiles As Integer)
Public Event UploadsComplete(NumFiles As Integer)


Public Property Get IsDownloading() As Boolean
    IsDownloading = Downloading
End Property


Public Property Get IsUploading() As Boolean
    IsUploading = Uploading
End Property



Public Function BeginUploading()
    CancelUpload = False
    Dim x As Integer
    x = 0

    
    With frmUpload.Winsock
        .Close
        .Connect CurrentServer.ServerIP, 6971
        Do While .State <> sckConnected: DoEvents
            If CancelUpload = True Then Exit Function
        Loop
    End With
    
    Uploading = True
    frmUpload.Show
    frmUpload.Counter.Enabled = True
    
    Do While Files.Count <> 0: DoEvents
        If CancelUpload = True Then Exit Function
        
        Col = InStr(File(1), ">")
        If Col Then
            filen$ = Left(File(1), Col - 1)
            dest$ = Right(File(1), Len(File(1)) - Col)
            frmUpload.NumFiles = "Uploading File " & (x + 1) & " of " & Files.Count
            FileDone = False
            Call SendFileToServer(filen$, dest$)
            Do While FileDone = False: DoEvents
                If CancelUpload = True Then Exit Function
            Loop 'wait for server to recieve the file before asking for another
        End If
        
        Call RemoveFirstFile
        x = x + 1
    Loop
    
    frmUpload.Winsock.Close
    RaiseEvent UploadsComplete(x)
    Uploading = False
    frmUpload.Counter.Enabled = False
    totalFiles = 0
End Function

Sub SendFileToServer(xFileName As String, destination As String)
Dim Buffer As String
Dim BufferSize As Integer
Dim Fiz As File
Dim pinche As ListItem
Dim FizObj As Scripting.FileSystemObject
Dim fileLength As Long, SuperBuffer As Long
Dim PercentDone As Long, b As Integer

    Set FizObj = CreateObject("Scripting.FileSystemObject")
    Set Fiz = FizObj.GetFile(xFileName)
    
    BufferSize = 2048
    
    i = FreeFile 'Find free file
    Open xFileName For Binary Access Read As #i 'open the file to read
        Debug.Print "--------Opening " + xFileName
        fileLength = LOF(i)
        
        StartSending = False
        frmUpload.Winsock.SendData ("FILE=" + destination + Fiz.Name + ":" & LOF(i))
        Debug.Print "Sending 'FILE='" + Fiz.Name
        
        If LOF(i) <> 0 Then
            Do While StartSending <> True: DoEvents
                If CancelUpload = True Then Exit Sub
            Loop
          
            Do While Not EOF(i): DoEvents
                If CancelUpload = True Then Exit Sub
                If fileLength - Loc(i) < BufferSize Then
                    Let BufferSize = fileLength - Loc(i)
                    If BufferSize = 0 Then GoTo done
                End If
                
                Buffer = Space(BufferSize)
           
                Get #i, , Buffer
                
                WaitForServerRecieve = True
                frmUpload.Winsock.SendData Buffer
                Do While WaitForServerRecieve = True: DoEvents
                    If CancelUpload = True Then Exit Sub
                Loop 'wait for server to recieve packet
                
                SuperBuffer = SuperBuffer + Len(Buffer)
                Debug.Print "BufferSize=" & BufferSize & " SuperBuffer=" & SuperBuffer
                
                frmUpload.BytesSent = FormatFileSize(SuperBuffer) + " of " + FormatFileSize(fileLength) + " sent"
                If SuperBuffer = 0 Then GoTo skipPercent 'Don't want division by zero
                
                PercentDone = SuperBuffer / fileLength * 100
                On Error Resume Next
                frmUpload.Progress.Value = PercentDone
                On Error GoTo 0
                DoEvents
skipPercent:
    
            Loop
        End If
done:
    Close #i
    Debug.Print "--------Closing " + xFileName
    
    StartSending = False
End Sub
Public Property Get Count() As Integer
    Count = Files.Count
End Property

Public Property Get File(index As Integer) As String
    File = Files(index)
End Property



Public Function AddFile(Path As String)

    Call Files.Add(Path)
    totalFiles = totalFiles + 1
    
End Function

Private Function RemoveFirstFile()
    Call Files.Remove(1)
End Function

Public Function BeginDownloading()
    CancelDownload = False
    Dim x As Integer
    x = 0
    
    frmMain.Winsock.SendData ("BEGIN_DOWNLOAD")
    
    With frmDownload.Winsock
        .LocalPort = 109
        .Listen
        Do While .State <> sckConnected: DoEvents
            'wait for server to connect
            If CancelDownload = True Then Exit Function
        Loop
    End With
    
    
    Downloading = True
    frmDownload.Show
    frmDownload.Timer1.Enabled = True
    
    Do While Files.Count <> 0: DoEvents
        If CancelDownload = True Then Exit Function
        
        Col = InStr(File(1), ">")
        If Col Then
            filen$ = Left(File(1), Col - 1)
            dest$ = Right(File(1), Len(File(1)) - Col)
            If filen$ = "" Or dest$ = "" Then GoTo next1
        End If
        
        CurrentFile = objFso.GetFileName(filen$)
        
        If Dir(dest$ + CurrentFile) <> "" Then
            'delete the file if it exists
            If MsgBox(dest$ + CurrentFile + " already exists. Do you want to replace it?", vbQuestion + vbYesNo, "Replace File?") = vbYes Then
                Kill dest$ + CurrentFile
            Else
                GoTo next1
            End If
        End If
        
        fComplete = False
        Call frmDownload.Winsock.SendData("GET " + filen$ + ":" + frmMain.Winsock.LocalIP)
        frmDownload.NumFiles = "Downloading File " & (x + 1) & " of " & totalFiles
        
        FileNum = FreeFile
        
        frmDownload.Progress.Value = 0
        
        Open dest$ + CurrentFile For Binary Access Write As #FileNum
        
        'wait for file to download
        Do While fComplete = False: DoEvents
            If CancelDownload = True Then Exit Function
        Loop
        
        
next1:
        'remove file from que
        RemoveFirstFile
        x = x + 1
    Loop
    totalFiles = 0
    frmDownload.Winsock.Close
    RaiseEvent DownloadsComplete(x)
    Downloading = False
    frmDownload.Timer1.Enabled = False
End Function

