VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'Property Elements
'SET @SPVersion          =(SELECT MAX(SPVersion)As SPVErsion FROM SoftwarePackage)
'SET @CLIENTNUM          = LTrim(RTrim(@cli06_SALN))
'SET @PolicyNo           = LTrim(RTrim(@ali0054_PolicyNum))
'SET @PolicyDescription      = LTrim(RTrim(@ali0069_PolicyDescription))
'IF LTrim(RTrim(@ali0065_MainFileInsuredName)) = ''
'BEGIN
'    SET @Insured        = LTrim(RTrim(@ali0062_NamedInsured))
'End
'Else
'BEGIN
'    SET @Insured        = LTrim(RTrim(@ali0065_MainFileInsuredName))
'End
'
'SET @MortgageeName      = LTrim(RTrim(@ali0066_MortgageHolder))
'SET @AgentNo            = LTrim(RTrim(@ali0056_AgentNum))
'SET @ReportedBy         = LTrim(RTrim(@ali0052_ReportedBy))
'SET @ReportedByPhone        = LTrim(RTrim(@ali0053_RBPhone))
'SET @Deductible         = Convert(money,LTrim(RTrim(@ali0072_Deductible1)))
'SET @AppDedClassTypeIDOrder     =   (   SELECT  DefaultAppDedClassTypeIDOrder
'                        From FeeSchedule
'                        WHERE   FeeScheduleID = @FeeScheduleID
'                    )
'
'SET @LRFormat           = @LRFormatClassName
'--Set the Status to PENDING !!! for All Assignments that have yet to be downloaded!!!
'SET @StatusID           =   (   SELECT  StatusID
'                        From Status
'                        WHERE   Status = 'PENDING'
'                    )
'--SET @TypeOfLossID         Already Set
'SET @XactTypeOfLoss         = ''
'SET @SentToXact         =0
'IF ISDate(LTrim(RTrim(@cli02_LossDate))) = 0
'BEGIN
'    SET @LossDate       =Null
'End
'Else
'BEGIN
'    SET @LossDate       =Convert(DateTime,LTrim(RTrim(@cli02_LossDate)))
'End
'
'IF ISDate(LTrim(RTrim(@cli04_DateAsgn))) = 0
'BEGIN
'    SET @AssignedDate   =Null
'End
'Else
'BEGIN
'    SET @AssignedDate   =Convert(DateTime,LTrim(RTrim(@cli04_DateAsgn)))
'End
'
'-- SET @ReceivedDate        =null
'-- SET @ContactDate         =null
'-- SET @InspectedDate       =null
'-- SET @CloseDate       =null
'-- SET @Reassigned      =0
'-- SET @DateReassigned      =null
'-- SET @RAAdjusterSpecID    =null
'-- SET @IsLocked        =0
'-- SET @IsDeleted       =0
'SET @AdminComments      =null
'--Build the Policy LImits List
'SET @COLUMN_DELIM =     dbo.GetCOLUMN_DELIM()
'SET @RECORD_DELIM =     dbo.GetRECORD_DELIM()
'SET @PLClassTypeID =    (SELECT ClassTypeID FROM ClassType WHERE Class = 'A')
'SET @PLLimitAmount =    @ali0070_BldgLimit
'SET @PLRCSaidProp = '0'
'SET @PLReserves =   '0'
'SET @PLAdminComments =  'XML01 Admin Loss Info ali0070_BldgLimit'--Must at least a Space in there!!!
'SET @PolicyLimitsList = '0' + @COLUMN_DELIM +               --PolicyLimitsID        int,
'            '0' + @COLUMN_DELIM +               --AssignmentsID     int,
'            '0' + @COLUMN_DELIM +               --[ID]          int,
'            '0' + @COLUMN_DELIM +               --IDAssignments     int,
'            @PLClassTypeID + @COLUMN_DELIM +        --ClassTypeID       int,
'            @PLLimitAmount + @COLUMN_DELIM +        --LimitAmount       money,
'            @PLRCSaidProp + @COLUMN_DELIM +         --RCSaidProp        money,
'            @PLReserves + @COLUMN_DELIM +           --Reserves      money,
'            '0' + @COLUMN_DELIM +               --IsDeleted     bit,
'            '0' + @COLUMN_DELIM +               --DownLoadMe        bit,
'            '0' + @COLUMN_DELIM +               --UpLoadMe      bit,
'            @PLAdminComments + @COLUMN_DELIM +      --AdminComments     Varchar(1000),
'            Cast(GetDate() As varchar(50))+ @COLUMN_DELIM + -- DateLastUpdated  datetime,
'            Cast(@UID as Varchar(15)) + @RECORD_DELIM   --UpdateByUserID        int
'SET @PLClassTypeID =    (SELECT ClassTypeID FROM ClassType WHERE Class = 'C')
'SET @PLLimitAmount =    @ali0071_ContLimit
'SET @PLRCSaidProp = '0'
'SET @PLReserves =   '0'
'SET @PLAdminComments =  'XML01 Admin Loss Info ali0071_ContLimit'--Must at least a Space in there!!!
'SET @PolicyLimitsList = @PolicyLimitsList +
'            '0' + @COLUMN_DELIM +               --PolicyLimitsID        int,
'            '0' + @COLUMN_DELIM +               --AssignmentsID     int,
'            '0' + @COLUMN_DELIM +               --[ID]          int,
'            '0' + @COLUMN_DELIM +               --IDAssignments     int,
'            @PLClassTypeID + @COLUMN_DELIM +        --ClassTypeID       int,
'            @PLLimitAmount + @COLUMN_DELIM +        --LimitAmount       money,
'            @PLRCSaidProp + @COLUMN_DELIM +         --RCSaidProp        money,
'            @PLReserves + @COLUMN_DELIM +           --Reserves      money,
'            '0' + @COLUMN_DELIM +               --IsDeleted     bit,
'            '0' + @COLUMN_DELIM +               --DownLoadMe        bit,
'            '0' + @COLUMN_DELIM +               --UpLoadMe      bit,
'            @PLAdminComments + @COLUMN_DELIM +      --AdminComments     Varchar(1000),
'            Cast(GetDate() As varchar(50))+ @COLUMN_DELIM + -- DateLastUpdated  datetime,
'            Cast(@UID as Varchar(15)) + @RECORD_DELIM   --UpdateByUserID        int
'
'SET @DateLastUpdated        =GetDate()
'SET @UpdateByUserID         = @UID

'End Property Elements




'Auto Elements



'End Auto Elements
