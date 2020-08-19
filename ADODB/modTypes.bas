Attribute VB_Name = "modTypes"
Option Explicit
Option Base 0

Public Type typDistEmpties
    '1. pTAPAppVersion --Versioning, just in case [different strokes for different folks]
    '"declare @pTAPAppVersion varchar(50) "
    '"TAP2.5"'
    strTAPAppVersion As String
    '2. pRegion --dbo.[Users].Region
    '"declare @pRegion nvarchar(50) "
    '"Region"
    strRegion As String
    '3. pOrderBy --Fields to order by e.g. [ShipState], [Name]
    '"declare @pOrderBy nvarchar(100) "
    '"OrderBy"
    strOrderBy As String
    '4. pCustSel --dbo.[INVENTORYREPORT].[UserID]
    '"declare @pCustSel nvarchar(50) "
    '"UserID"
    strCustSel As String
    '5. pCompany --Set @pCompany = -1 for ALL Companies in results.  Otherwise, ALL other companies besides @pCompany value will be r
    '"declare @pCompany bigint "
    '"Company"
    lngCompany As Long
    '6. pDebugOn --Debugging?  SET @pDebugOn = 1 IF NOT SET @pDebugOn = 0  Will return a robust set of queries in order to interrogate/reconcile the sp results.
    '"declare @pDebugOn bit "
    'Debug OFF
    '"@pDebugOn
    intDebugOn As Integer
End Type






