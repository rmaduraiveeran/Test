Option Strict On
Option Explicit On


Public Class Job


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
    ''' Gets or sets teh pay group (company code)
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


    Private empl_StatusValue As String
    ''' <summary>
    ''' Gets or sets the employment status (active, leave, terminated)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPL_STATUS() As String
        Get
            Return empl_StatusValue
        End Get
        Set(ByVal value As String)
            empl_StatusValue = value
        End Set
    End Property


    Private action_ReasonValue As String
    ''' <summary>
    ''' Gets or sets the job change reason
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ACTION_REASON() As String
        Get
            Return action_ReasonValue
        End Get
        Set(ByVal value As String)
            action_ReasonValue = value
        End Set
    End Property


    Private locationValue As String
    ''' <summary>
    ''' Gets or sets the location
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LOCATION() As String
        Get
            Return locationValue
        End Get
        Set(ByVal value As String)
            locationValue = value
        End Set
    End Property


    Private full_Part_TimeValue As String
    ''' <summary>
    ''' Gets or sets full time or part time
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FULL_PART_TIME() As String
        Get
            Return full_Part_TimeValue
        End Get
        Set(ByVal value As String)
            full_Part_TimeValue = value
        End Set
    End Property


    Private companyValue As String
    ''' <summary>
    ''' Gets or sets the company code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property COMPANY() As String
        Get
            Return companyValue
        End Get
        Set(ByVal value As String)
            companyValue = value
        End Set
    End Property


    Private empl_TypeValue As String
    ''' <summary>
    ''' Gets or sets the salary or hourly flag
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPL_TYPE() As String
        Get
            Return empl_TypeValue
        End Get
        Set(ByVal value As String)
            empl_TypeValue = value
        End Set
    End Property


    Private empl_ClassValue As String
    ''' <summary>
    ''' Gets or sets the ex-patriot flag
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPL_CLASS() As String
        Get
            Return empl_ClassValue
        End Get
        Set(ByVal value As String)
            empl_ClassValue = value
        End Set
    End Property


    Private data_ControlValue As String
    ''' <summary>
    ''' Gets or sets the nature
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DATA_CONTROL() As String
        Get
            Return data_ControlValue
        End Get
        Set(ByVal value As String)
            data_ControlValue = value
        End Set
    End Property

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


    Private home_DepartmentValue As String
    ''' <summary>
    ''' Gets or sets the department
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HOME_DEPARTMENT() As String
        Get
            Return home_DepartmentValue
        End Get
        Set(ByVal value As String)
            home_DepartmentValue = value
        End Set
    End Property


    Private titleValue As String
    ''' <summary>
    ''' Gets or sets the group number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TITLE() As String
        Get
            Return titleValue
        End Get
        Set(ByVal value As String)
            titleValue = value
        End Set
    End Property


    Private workers_Comp_CdValue As String
    ''' <summary>
    ''' Gets or sets the worker's comp code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WORKERS_COMP_CD() As String
        Get
            Return workers_Comp_CdValue
        End Get
        Set(ByVal value As String)
            workers_Comp_CdValue = value
        End Set
    End Property


End Class
