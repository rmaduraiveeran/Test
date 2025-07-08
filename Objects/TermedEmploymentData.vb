Option Strict On
Option Explicit On


Public Class TermedEmploymentData

    Private empNumberValue As String
    ''' <summary>
    ''' Gets or sets the Employee Number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPLOYEE_NUMBER() As String
        Get
            Return empNumberValue
        End Get
        Set(ByVal value As String)
            empNumberValue = value
        End Set
    End Property

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

    Private emplNameValue As String
    ''' <summary>
    ''' Gets or sets the Employee's Name 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPLOYEE_NAME() As String
        Get
            Return emplNameValue
        End Get
        Set(ByVal value As String)
            emplNameValue = value
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


    Private hire_DtValue As String
    ''' <summary>
    ''' Gets or sets the original hire date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HIRE_DT() As String
        Get
            Return hire_DtValue
        End Get
        Set(ByVal value As String)
            hire_DtValue = value
        End Set
    End Property


    Private rehire_DtValue As String
    ''' <summary>
    ''' Gets or sets the rehire date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property REHIRE_DT() As String
        Get
            Return rehire_DtValue
        End Get
        Set(ByVal value As String)
            rehire_DtValue = value
        End Set
    End Property


    Private cmpny_Seniority_DtValue As String
    ''' <summary>
    ''' Gets or sets the seniority date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CMPNY_SENIORITY_DT() As String
        Get
            Return cmpny_Seniority_DtValue
        End Get
        Set(ByVal value As String)
            cmpny_Seniority_DtValue = value
        End Set
    End Property


    Private termination_DtValue As String
    ''' <summary>
    ''' Gets or sets the termination date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TERMINATION_DT() As String
        Get
            Return termination_DtValue
        End Get
        Set(ByVal value As String)
            termination_DtValue = value
        End Set
    End Property


    Private last_Date_WorkedValue As String
    ''' <summary>
    ''' Gets or sets the last date worked
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LAST_DATE_WORKED() As String
        Get
            Return last_Date_WorkedValue
        End Get
        Set(ByVal value As String)
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

    Private supNameValue As String
    ''' <summary>
    ''' Gets or sets the Supervisor's Name 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SUPERVISOR_NAME() As String
        Get
            Return supNameValue
        End Get
        Set(ByVal value As String)
            supNameValue = value
        End Set
    End Property

End Class
