Option Strict On
Option Explicit On
Imports System.Reflection

Public Class TaxData


    Private emplIDValue As String
    ''' <summary>
    ''' Gets or sets the employee ID (EEID less last char)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPLID() As String
        Get
            Return emplIDValue
        End Get
        Set(ByVal value As String)
            emplIDValue = value
        End Set
    End Property


    Private payGroupValue As String
    ''' <summary>
    ''' Gets or sets the pay group (company code)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PAYGROUP() As String
        Get
            Return payGroupValue
        End Get
        Set(ByVal value As String)
            payGroupValue = value
        End Set
    End Property


    Private pay_FrequencyValue As String
    ''' <summary>
    ''' Gets or sets the pay frequency
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PAY_FREQUENCY() As String
        Get
            Return pay_FrequencyValue
        End Get
        Set(ByVal value As String)
            pay_FrequencyValue = value
        End Set
    End Property


    Private federal_Mar_StatusValue As String
    ''' <summary>
    ''' Gets or sets the federal marital status
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FEDERAL_MAR_STATUS() As String
        Get
            Return federal_Mar_StatusValue
        End Get
        Set(ByVal value As String)
            federal_Mar_StatusValue = value
        End Set
    End Property


    Private fed_AllowancesValue As String
    ''' <summary>
    ''' Gets or sets the federal exemptions
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FED_ALLOWANCES() As String
        Get
            Return fed_AllowancesValue
        End Get
        Set(ByVal value As String)
            fed_AllowancesValue = value
        End Set
    End Property


    Private federal_Tax_BlockValue As String
    ''' <summary>
    ''' Gets or sets the federal tax block
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FEDERAL_TAX_BLOCK() As String
        Get
            Return federal_Tax_BlockValue
        End Get
        Set(ByVal value As String)
            federal_Tax_BlockValue = value
        End Set
    End Property


    Private schdist_Tax_BlockValue As String
    ''' <summary>
    ''' Gets or sets the school district tax block
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SCHDIST_TAX_BLOCK() As String
        Get
            Return schdist_Tax_BlockValue
        End Get
        Set(ByVal value As String)
            schdist_Tax_BlockValue = value
        End Set
    End Property


    Private suisdi_Tax_BlockValue As String
    ''' <summary>
    ''' Gets or sets the SUISDI tax block
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SUISDI_TAX_BLOCK() As String
        Get
            Return suisdi_Tax_BlockValue
        End Get
        Set(ByVal value As String)
            suisdi_Tax_BlockValue = value
        End Set
    End Property


    Private ssmed_Tax_BlockValue As String
    ''' <summary>
    ''' Gets or sets the social security/medicare tax block
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SSMED_TAX_BLOCK() As String
        Get
            Return ssmed_Tax_BlockValue
        End Get
        Set(ByVal value As String)
            ssmed_Tax_BlockValue = value
        End Set
    End Property


    Private federal_Addl_AmtValue As Int32
    ''' <summary>
    ''' Gets or sets the additional federal tax dollar amount
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FEDERAL_ADDL_AMT() As Int32
        Get
            Return federal_Addl_AmtValue
        End Get
        Set(ByVal value As Int32)
            federal_Addl_AmtValue = value
        End Set
    End Property


    Private state_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the work in state code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property STATE_TAX_CD() As String
        Get
            Return state_Tax_CdValue
        End Get
        Set(ByVal value As String)
            state_Tax_CdValue = value
        End Set
    End Property


    Private state2_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the resident state code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property STATE2_TAX_CD() As String
        Get
            Return state2_Tax_CdValue
        End Get
        Set(ByVal value As String)
            state2_Tax_CdValue = value
        End Set
    End Property


    Private local_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the local work in state code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LOCAL_TAX_CD() As String
        Get
            Return local_Tax_CdValue
        End Get
        Set(ByVal value As String)
            local_Tax_CdValue = value
        End Set
    End Property


    Private local2_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the local resident state code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LOCAL2_TAX_CD() As String
        Get
            Return local2_Tax_CdValue
        End Get
        Set(ByVal value As String)
            local2_Tax_CdValue = value
        End Set
    End Property


    Private school_DistrictValue As String
    ''' <summary>
    ''' Gets or sets the school district code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SCHOOL_DISTRICT() As String
        Get
            Return school_DistrictValue
        End Get
        Set(ByVal value As String)
            school_DistrictValue = value
        End Set
    End Property


    Private sui_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the SUI tax code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SUI_TAX_CD() As String
        Get
            Return sui_Tax_CdValue
        End Get
        Set(ByVal value As String)
            sui_Tax_CdValue = value
        End Set
    End Property


    Private tax_Loack_End_DtValue As Date
    ''' <summary>
    ''' Gets or sets the tax lock end date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TAX_LOCK_END_DT() As Date
        Get
            Return tax_Loack_End_DtValue
        End Get
        Set(ByVal value As Date)
            tax_Loack_End_DtValue = value
        End Set
    End Property


    Private tax_Lck_Fed_Mar_StValue As String
    ''' <summary>
    ''' Gets or sets the tax lock federal marital status
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TAX_LCK_FED_MAR_ST() As String
        Get
            Return tax_Lck_Fed_Mar_StValue
        End Get
        Set(ByVal value As String)
            tax_Lck_Fed_Mar_StValue = value
        End Set
    End Property


    Private tax_Lock_Fed_AllowValue As String
    ''' <summary>
    ''' Gets or sets the tax lock federal allowance
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TAX_LOCK_FED_ALLOW() As String
        Get
            Return tax_Lock_Fed_AllowValue
        End Get
        Set(ByVal value As String)
            tax_Lock_Fed_AllowValue = value
        End Set
    End Property


    Private local4_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the local4 tax code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LOCAL4_TAX_CD() As String
        Get
            Return local4_Tax_CdValue
        End Get
        Set(ByVal value As String)
            local4_Tax_CdValue = value
        End Set
    End Property


    Private w4IsLockedValue As String
    ''' <summary>
    ''' Gets or sets the W4 is locked 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property W4IsLocked() As String
        Get
            Return w4IsLockedValue
        End Get
        Set(ByVal value As String)
            w4IsLockedValue = value
        End Set
    End Property


    ' This property is only used for conversion
    Private file_NbrValue As String
    ''' <summary>
    ''' Gets or sets the employee number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FILE_NBR() As String
        Get
            Return file_NbrValue
        End Get
        Set(ByVal value As String)
            file_NbrValue = value
        End Set
    End Property

    'New 2020 W4 Fields
    'Private use_Old_W4Value As String
    'Public Property USE_OLD_W4() As String
    '    Get
    '        Return use_Old_W4Value
    '    End Get
    '    Set(value As String)
    '        use_Old_W4Value = value
    '    End Set
    'End Property




    Private w4_FORM_YEARValue As String
    Public Property W4_FORM_YEAR() As String
        Get
            Return w4_FORM_YEARValue
        End Get
        Set(value As String)
            w4_FORM_YEARValue = value
        End Set
    End Property

    Private oth_IncomeValue As String
    Public Property OTH_INCOME() As String
        Get
            Return oth_IncomeValue
        End Get
        Set(value As String)
            oth_IncomeValue = value
        End Set
    End Property

    Private oth_DeductionsValue As String
    Public Property OTH_DEDUCTIONS() As String
        Get
            Return oth_DeductionsValue
        End Get
        Set(value As String)
            oth_DeductionsValue = value
        End Set
    End Property

    Private dependants_AmtValue As String
    Public Property DEPENDENTS_AMT() As String
        Get
            Return dependants_AmtValue
        End Get
        Set(value As String)
            dependants_AmtValue = value
        End Set
    End Property

    Private Property multiple_JobsValue As String
    Public Property MULTIPLE_JOBS() As String
        Get
            Return multiple_JobsValue
        End Get
        Set(value As String)
            multiple_JobsValue = value
        End Set
    End Property

    Private Property _long_Term_Care_Ins_Status As String
    Public Property Long_Term_Care_Ins_Status() As String
        Get
            Return _long_Term_Care_Ins_Status
        End Get
        Set(value As String)
            _long_Term_Care_Ins_Status = value
        End Set
    End Property

    Private SEND_NOTFN_FOR_W4_FORM_YEARValue As String
    Public Property SEND_NOTFN_FOR_W4_FORM_YEAR() As String
        Get
            Return SEND_NOTFN_FOR_W4_FORM_YEARValue
        End Get
        Set(value As String)
            SEND_NOTFN_FOR_W4_FORM_YEARValue = value
        End Set
    End Property

    Private EMP_NoValue As String
    Public Property EMP_No() As String
        Get
            Return EMP_NoValue
        End Get
        Set(value As String)
            EMP_NoValue = value
        End Set
    End Property
    Private EMP_CoId As String
    Public Property EecCoID() As String
        Get
            Return EMP_CoId
        End Get
        Set(value As String)
            EMP_CoId = value
        End Set
    End Property
    Private Employee_Name As String
    Public Property EmployeeName() As String
        Get
            Return Employee_Name
        End Get
        Set(value As String)
            Employee_Name = value
        End Set
    End Property

    ''' <summary>
    ''' Gets/Sets a property
    ''' </summary>
    ''' <param name="name">Property name</param>
    ''' <returns></returns>
    <CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="1")>
    Public Property ByName(ByVal name As String) As Object
        Get
            Dim oType As Type = GetType(TaxData)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)
            Return propInfo.GetValue(Me, Nothing)
        End Get

        Set(value As Object)
            Dim oType As Type = GetType(TaxData)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)

            If propInfo.PropertyType.Name.ToUpper = "INT32" Then
                propInfo.SetValue(Me, CInt(value), Nothing)
            ElseIf propInfo.PropertyType.Name.ToUpper = "DATETIME" And value IsNot Nothing Then
                If value.ToString.Trim.Length > 0 Then
                    propInfo.SetValue(Me, CDate(value), Nothing)
                End If
            Else
                propInfo.SetValue(Me, value, Nothing)
            End If

        End Set
    End Property

End Class
