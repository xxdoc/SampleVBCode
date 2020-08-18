Attribute VB_Name = "modSafeco"
Option Explicit

'BGS 9.27.2001 Safeco Tables
Public Const DB_SAFECO_CHECKS As String = "Checks"
Public Const DB_SAFECO_INDEM As String = "Indemnity"
Public Const DB_ACTIVITY As String = "ActivityLog"
Public Const DB_PHOTO As String = "PhotoLog"
Public Const DB_ASSIGNMENTS As String = "Assignments"

'BGS 11.08.2001 Constants used for SAFECO Check Class
Public Const FC_BUILD As String = "01"
Public Const FC_COMBUILD As String = "15"
Public Const FC_CONTENTS As String = "02"
'BGS 2.28.2002 143  Class of Loss - 88 s/b ALE
'Todd let me know that 88 should be the ALE code, not 89
'Its been 89 since forever so we will change it to 88 but we still
'need to account for 89 being used previously.
Public Const FC_ALE As String = "03"
