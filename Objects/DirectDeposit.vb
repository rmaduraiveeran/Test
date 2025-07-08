Option Strict On
Option Explicit On


Public Class DirectDeposit


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
    ''' Gets or sets the account type
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


    Private full_DepositValue As String
    ''' <summary>
    ''' Gets or sets the full or partial deposit indicator
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FULL_DEPOSIT() As String
        Get
            Return full_DepositValue
        End Get
        Set(ByVal value As String)
            full_DepositValue = value
        End Set
    End Property


    Private transit_NbrValue As String
    ''' <summary>
    ''' Gets or sets the routing number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TRANSIT_NBR() As String
        Get
            Return transit_NbrValue
        End Get
        Set(ByVal value As String)
            transit_NbrValue = value
        End Set
    End Property


    Private account_NbrValue As String
    ''' <summary>
    ''' Gets or sets the account number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ACCOUNT_NBR() As String
        Get
            Return account_NbrValue
        End Get
        Set(ByVal value As String)
            account_NbrValue = value
        End Set
    End Property


    Private deposit_AmtValue As Decimal
    ''' <summary>
    ''' Gets or sets the deposit amount in dollars
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DEPOSIT_AMT() As Decimal
        Get
            Return deposit_AmtValue
        End Get
        Set(ByVal value As Decimal)
            deposit_AmtValue = value
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


    ''' <summary>
    ''' Gets or sets the account inactive flag
    ''' </summary>
    ''' <remarks></remarks>
    Private accountIsInactiveValue As String
    Public Property AccountIsInactive() As String
        Get
            Return accountIsInactiveValue
        End Get
        Set(ByVal value As String)
            accountIsInactiveValue = value
        End Set
    End Property


    ' This property is only used for conversion
    Private file_NumberValue As String
    ''' <summary>
    ''' Gets or sets the employee number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FILE_NBR() As String
        Get
            Return file_NumberValue
        End Get
        Set(ByVal value As String)
            file_NumberValue = value
        End Set
    End Property


End Class
