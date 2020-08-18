Attribute VB_Name = "basMain"
Public mbshowing               As Boolean
Public miPtCnt                 As Integer
Public mPts(2)                 As PointDbl
Public mptTemp                 As PointDbl
Public mArc                    As ArcStruct
Public cntr                    As Integer

Public mBox                    As Boolean
Public mFilled                 As Boolean

Public ForeColor               As Double
Public FillColor               As Double

Public mHistory()              As crvHistory

Public mbBalloonClick          As Boolean
Public mbEscapeBalloon         As Boolean
Public mbBalloonMode           As Boolean

Public iHISTORYCOUNT           As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'wddx components

Public Deser As WDDXDeserializer  'Allaire's WDDX deserializer

Public Ser As WDDXSerializer      'Allaire's WDDX serializer

Public mLineWidth              As Integer
