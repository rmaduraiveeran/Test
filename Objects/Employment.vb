Option Strict On
Option Explicit On


Public Class Employment


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


    Private hire_DtValue As Date
    ''' <summary>
    ''' Gets or sets the original hire date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HIRE_DT() As Date
        Get
            Return hire_DtValue
        End Get
        Set(ByVal value As Date)
            hire_DtValue = value
        End Set
    End Property


    Private rehire_DtValue As Date
    ''' <summary>
    ''' Gets or sets the rehire date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property REHIRE_DT() As Date
        Get
            Return rehire_DtValue
        End Get
        Set(ByVal value As Date)
            rehire_DtValue = value
        End Set
    End Property


    Private cmpny_Seniority_DtValue As Date
    ''' <summary>
    ''' Gets or sets the seniority date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CMPNY_SENIORITY_DT() As Date
        Get
            Return cmpny_Seniority_DtValue
        End Get
        Set(ByVal value As Date)
            cmpny_Seniority_DtValue = value
        End Set
    End Property


    Private termination_DtValue As Date
    ''' <summary>
    ''' Gets or sets the termination date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TERMINATION_DT() As Date
        Get
            Return termination_DtValue
        End Get
        Set(ByVal value As Date)
            termination_DtValue = value
        End Set
    End Property


    Private last_Date_WorkedValue As Date
    ''' <summary>
    ''' Gets or sets the last date worked
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LAST_DATE_WORKED() As Date
        Get
            Return last_Date_WorkedValue
        End Get
        Set(ByVal value As Date)
            last_Date_WorkedValue = value
        End Set
    End Property


    Private business_TitleValue As String
    ''' <summary>
    ''' Gets or sets the employee's job title
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BUSINESS_TITLE() As String
        Get
            Return business_TitleValue
        End Get
        Set(ByVal value As String)
            business_TitleValue = value
        End Set
    End Property


    Private supervisor_IDValue As String
    ''' <summary>
    ''' Gets or sets the supervisor's EMPLID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SUPERVISOR_ID() As String
        Get
            Return supervisor_IDValue
        End Get
        Set(ByVal value As String)
            supervisor_IDValue = value
        End Set
    End Property


End Class
