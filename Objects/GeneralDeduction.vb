Option Strict On
Option Explicit On


Public Class GeneralDeduction


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


    Private dedcdValue As String
    ''' <summary>
    ''' Gets or sets the deduction code name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DEDCD() As String
        Get
            Return dedcdValue
        End Get
        Set(ByVal value As String)
            dedcdValue = value
        End Set
    End Property


    Private ded_Addl_AmtValue As Decimal
    ''' <summary>
    ''' Gets or sets the deduction dollar amount
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DED_ADDL_AMT() As Decimal
        Get
            Return ded_Addl_AmtValue
        End Get
        Set(ByVal value As Decimal)
            ded_Addl_AmtValue = value
        End Set
    End Property


    Private ded_Rate_PctValue As Decimal
    ''' <summary>
    ''' Gets or sets the deduction percentage
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DED_RATE_PCT() As Decimal
        Get
            Return ded_Rate_PctValue
        End Get
        Set(ByVal value As Decimal)
            ded_Rate_PctValue = value
        End Set
    End Property


    Private goal_AmtValue As Decimal
    ''' <summary>
    ''' Gets or sets the deduction goal amount
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GOAL_AMT() As Decimal
        Get
            Return goal_AmtValue
        End Get
        Set(ByVal value As Decimal)
            goal_AmtValue = value
        End Set
    End Property


    Private end_DtValue As Date
    ''' <summary>
    ''' Gets or sets the end date (stop date)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property END_DT() As Date
        Get
            Return end_DtValue
        End Get
        Set(ByVal value As Date)
            end_DtValue = value
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


    Private skipDeductionValue As String
    ''' <summary>
    ''' Gets or sets the skip deduction flag
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SkipDeduction() As String
        Get
            Return skipDeductionValue
        End Get
        Set(ByVal value As String)
            skipDeductionValue = value
        End Set
    End Property


    Private skipDeductionAccForValue As String
    ''' <summary>
    ''' Gets or sets the skipped deduction accounted for flag
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SkipDeductionAccFor() As String
        Get
            Return skipDeductionAccForValue
        End Get
        Set(ByVal value As String)
            skipDeductionAccForValue = value
        End Set
    End Property

    Private skipZeroGoalAmountValue As String
    ''' <summary>
    ''' Gets or sets the skipped zero goal amount flag
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SkipZeroGoalAmount() As String
        Get
            Return skipZeroGoalAmountValue
        End Get
        Set(ByVal value As String)
            skipZeroGoalAmountValue = value
        End Set
    End Property

    Private eecCOIDValue As String
    ''' <summary>
    ''' Gets or sets the company id
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EECCoid() As String
        Get
            Return eecCOIDValue
        End Get
        Set(ByVal value As String)
            eecCOIDValue = value
        End Set
    End Property


    Private employeeNumberValue As String
    ''' <summary>
    ''' Gets or sets the employee number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EmployeeNumber() As String
        Get
            Return employeeNumberValue
        End Get
        Set(ByVal value As String)
            employeeNumberValue = value
        End Set
    End Property


    Private employeeNameValue As String
    ''' <summary>
    ''' Gets or sets the employee name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EmployeeName() As String
        Get
            Return employeeNameValue
        End Get
        Set(ByVal value As String)
            employeeNameValue = value
        End Set
    End Property

    Private _gtd_amt As Decimal
    ''' <summary>
    ''' Gets or sets the GTD taken for the deduction amount.
    ''' </summary>
    ''' <returns></returns>
    Public Property GTD_Amt() As Decimal
        Get
            Return _gtd_amt
        End Get
        Set(value As Decimal)
            _gtd_amt = value
        End Set
    End Property

    Private start_DtValue As Date
    ''' <summary>
    ''' Gets or sets the start date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property START_DT() As Date
        Get
            Return start_DtValue
        End Get
        Set(ByVal value As Date)
            start_DtValue = value
        End Set
    End Property

End Class
