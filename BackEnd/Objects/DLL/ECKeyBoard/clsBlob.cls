VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BLOBToFile(ByVal strFullPath As String, ByRef objField As ADODB.Field, Optional ByVal bUseStream As Boolean = True, Optional ByVal lngChunkSize As Long = 8192) As Boolean
    On Error Resume Next
    Dim objStream As ADODB.Stream
    Dim intFreeFile As Integer
    Dim lngBytesLeft As Long
    Dim lngReadBytes As Long
    Dim byBuffer() As Byte
    
    If bUseStream Then
        Set objStream = New ADODB.Stream
        With objStream
            .Type = adTypeBinary
            .Open
            .Write objField.Value
            .SaveToFile strFullPath, adSaveCreateOverWrite
        End With
        DoEvents
    Else
        If goUtil.utFileExists(strFullPath) Then
            goUtil.utDeleteFile (strFullPath)
        End If
        lngBytesLeft = objField.ActualSize
        intFreeFile = FreeFile
        Open strFullPath For Binary As #intFreeFile
        
        Do Until lngBytesLeft <= 0
            lngReadBytes = lngBytesLeft
            
            If lngReadBytes > lngChunkSize Then
                lngReadBytes = lngChunkSize
            End If
            
            byBuffer = objField.GetChunk(lngReadBytes)
            Put #intFreeFile, , byBuffer
            lngBytesLeft = lngBytesLeft - lngReadBytes
            
            DoEvents
        Loop
            Close #intFreeFile
    End If

    If Err.Number <> 0 Or Err.LastDllError <> 0 Then
        BLOBToFile = False
    Else
        BLOBToFile = True
    End If
    
     Set objStream = Nothing
     
End Function

Public Function FileToBLOB(ByVal strFullPath As String, ByRef objField As ADODB.Field, Optional ByVal bUseStream As Boolean = True, Optional ByVal lngChunkSize As Long = 8192) As Boolean
    On Error Resume Next
    Dim objStream As ADODB.Stream
    Dim intFreeFile As Integer
    Dim lngBytesLeft As Long
    Dim lngReadBytes As Long
    Dim byBuffer() As Byte
    Dim varChunk As Variant

    If bUseStream Then
        Set objStream = New ADODB.Stream
        With objStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile strFullPath
            objField.Value = .Read(adReadAll)
        End With
    Else
        With objField
            'If the field does not support Long  Binary data
            'then cannot load the data into the field.
            If (.Attributes And adFldLong) <> 0 Then
                intFreeFile = FreeFile
                Open strFullPath For Binary Access Read As #intFreeFile
                lngBytesLeft = LOF(intFreeFile)
                Do Until lngBytesLeft <= 0
                    If lngBytesLeft > lngChunkSize Then
                        lngReadBytes = lngChunkSize
                    Else
                        lngReadBytes = lngBytesLeft
                    End If
                    ReDim byBuffer(lngReadBytes)
                    Get #intFreeFile, , byBuffer()
                    objField.AppendChunk byBuffer()
                    lngBytesLeft = lngBytesLeft - lngReadBytes
                    DoEvents
                    Loop
                    Close #intFreeFile
                Else
                    Err.Raise -10000, "FileToBLOB", "The Database Field does not support Long Binary Data."
                End If
            End With
        End If
        
        If Err.Number <> 0 Or Err.LastDllError <> 0 Then
            FileToBLOB = False
        Else
            FileToBLOB = True
        End If
        
        Set objStream = Nothing
        
End Function

Public Function SetUtilObject(oUtilObject As Object) As Boolean
    On Error Resume Next
    Set goUtil = oUtilObject
    If Err.Number = 0 Then
        SetUtilObject = True
    Else
        SetUtilObject = False
    End If
End Function
