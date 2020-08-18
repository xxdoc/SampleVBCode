Attribute VB_Name = "modSpool"
' *************************************************************
'  Copyright ©1994-99, Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb/
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit
'
' Win32 API Calls
'
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrn As Long, ByVal Level As Long, pDocInfo As DOC_INFO_1) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrn As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'
' Structure required by StartDocPrinter
'
Private Type DOC_INFO_1
   pDocName As String
   pOutputFile As String
   pDatatype As String
End Type

Private mhPrn As Long 'Active printer pointer

 
Public Sub SelectDefaultPrinter(Lst As ComboBox, Optional psLastSelectedPrinter As String)
   Dim sRet As String
   Dim nRet As Integer
   Dim i As Integer
   '
   ' Look for default printer in WIN.INI
   '
   sRet = Space(255)
   nRet = GetProfileString("Windows", ByVal "device", "", _
                           sRet, Len(sRet))
   '
   ' Truncate default printer name.
   '
   If nRet Then
      sRet = UCase(left(sRet, InStr(sRet, ",") - 1))
      '
      ' Cycle list looking for matching entry.
      '
      'Look for last selected first. If can't find it then
      'look for default printer
        If psLastSelectedPrinter <> vbNullString Then
            For i = 0 To Lst.ListCount
                If left(UCase(Lst.List(i)), Len(psLastSelectedPrinter)) = UCase(psLastSelectedPrinter) Then
                    Lst.ListIndex = i
                    Exit Sub
                End If
            Next
        End If
        
      For i = 0 To Lst.ListCount
         If left(UCase(Lst.List(i)), Len(sRet)) = sRet Then
            '
            ' Found it. Set index and bail.
            '
            Lst.ListIndex = i
            Exit For
         End If
      Next i
   End If
End Sub

Public Sub SpoolFile(sFile As String, Optional AppName As String = "")
   Dim Buffer() As Byte
   Dim hFile As Integer
   Dim Written As Long
   Dim di As DOC_INFO_1
   Dim i As Long
   Dim oUT As New V2ECKeyBoard.clsUtil
   
   Const BufSize As Long = &H4000
   '
   ' Extract filename from passed spec, and build job name.
   ' Fill remainder of DOC_INFO_1 structure.
   '
   If InStr(sFile, "\") Then
      For i = Len(sFile) To 1 Step -1
         If Mid(sFile, i, 1) = "\" Then Exit For
         di.pDocName = Mid(sFile, i, 1) & di.pDocName
      Next i
   Else
      di.pDocName = sFile
   End If
   If Len(AppName) Then
      di.pDocName = AppName & ": " & di.pDocName
   End If
   di.pOutputFile = vbNullString
   di.pDatatype = "RAW"
   '
   ' Open printer for output to obtain handle.
   ' Set it up to begin recieving raw data.
   '
   Call StartDocPrinter(mhPrn, 1, di)
   Call StartPagePrinter(mhPrn)
   '
   ' Open file and pump it to the printer.
   '
   hFile = FreeFile
   Open sFile For Binary Access Read As hFile
      '
      ' Read in 16K buffers and spool.
      '
      ReDim Buffer(1 To BufSize) As Byte
      For i = 1 To LOF(hFile) \ BufSize
         Get #hFile, , Buffer
         Call WritePrinter(mhPrn, Buffer(1), BufSize, Written)
      Next i
      '
      ' Get last chunk of file if it doesn't
      ' fit evenly into a 16K buffer.
      '
      If LOF(hFile) Mod BufSize Then
         ReDim Buffer(1 To (LOF(hFile) Mod BufSize)) As Byte
         Get #hFile, , Buffer
         Call WritePrinter(mhPrn, Buffer(1), UBound(Buffer), Written)
      End If
   Close #hFile
   '
   ' Shut down spooling process.
   '
    Call EndPagePrinter(mhPrn)
    Call EndDocPrinter(mhPrn)
    
End Sub

Public Function IsFile(SpecIn As String) As Boolean
   Dim Attr As Byte
   '
   ' Guard against bad SpecIn by ignoring errors.
   '
   On Error Resume Next
   '
   ' Get attribute of SpecIn.
   '
   Attr = GetAttr(SpecIn)
   If Err = 0 Then
      '
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      '
      If (Attr And vbDirectory) = vbDirectory Then
         IsFile = False
      Else
         IsFile = True
      End If
   End If
End Function

Public Function SpoolClosePrn() As Long
    SpoolClosePrn = ClosePrinter(mhPrn)
End Function

Public Function SpoolOpenPrn(psPrnName As String) As Long
    mhPrn = 0
    SpoolOpenPrn = OpenPrinter(psPrnName, mhPrn, ByVal 0&)
End Function
    
