Option Strict On
Option Explicit On


Public Class PersonalData


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


    Private first_NameValue As String
    ''' <summary>
    ''' Gets or sets the first name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FIRST_NAME() As String
        Get
            Return first_NameValue
        End Get
        Set(ByVal value As String)
            first_NameValue = value
        End Set
    End Property


    Private middle_NameValue As String
    ''' <summary>
    ''' Gets or sets the middle name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MIDDLE_NAME() As String
        Get
            Return middle_NameValue
        End Get
        Set(ByVal value As String)
            middle_NameValue = value
        End Set
    End Property


    Private last_NameValue As String
    ''' <summary>
    ''' Gets or sets the last name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LAST_NAME() As String
        Get
            Return last_NameValue
        End Get
        Set(ByVal value As String)
            last_NameValue = value
        End Set
    End Property


    Private street1Value As String
    ''' <summary>
    ''' Gets or sets street address 1
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property STREET1() As String
        Get
            Return street1Value
        End Get
        Set(ByVal value As String)
            street1Value = value
        End Set
    End Property


    Private street2Value As String
    ''' <summary>
    ''' Gets or sets street adress 2
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property STREET2() As String
        Get
            Return street2Value
        End Get
        Set(ByVal value As String)
            street2Value = value
        End Set
    End Property


    Private cityValue As String
    ''' <summary>
    ''' Gets or sets the city
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CITY() As String
        Get
            Return cityValue
        End Get
        Set(ByVal value As String)
            cityValue = value
        End Set
    End Property


    Private stateValue As String
    ''' <summary>
    ''' Gets or sets the state code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property STATE() As String
        Get
            Return stateValue
        End Get
        Set(ByVal value As String)
            stateValue = value
        End Set
    End Property


    Private zipValue As String
    ''' <summary>
    ''' Gets or sets the zip code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ZIP() As String
        Get
            Return zipValue
        End Get
        Set(ByVal value As String)
            zipValue = value
        End Set
    End Property


    Private home_PhoneValue As String
    ''' <summary>
    ''' Gets or sets the home phone number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HOME_PHONE() As String
        Get
            Return home_PhoneValue
        End Get
        Set(ByVal value As String)
            home_PhoneValue = value
        End Set
    End Property


    Private ssnValue As String
    ''' <summary>
    ''' Gets or sets the SSN
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SSN() As String
        Get
            Return ssnValue
        End Get
        Set(ByVal value As String)
            ssnValue = value
        End Set
    End Property


    Private orig_Hire_DtValue As Date
    ''' <summary>
    ''' Gets or sets the original hire date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ORIG_HIRE_DT() As Date
        Get
            Return orig_Hire_DtValue
        End Get
        Set(ByVal value As Date)
            orig_Hire_DtValue = value
        End Set
    End Property


    Private sexValue As String
    ''' <summary>
    ''' Gets or sets the gender
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SEX() As String
        Get
            Return sexValue
        End Get
        Set(ByVal value As String)
            sexValue = value
        End Set
    End Property


    Private birthDateValue As Date
    ''' <summary>
    ''' Gets or sets the birth date
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BIRTHDATE() As Date
        Get
            Return birthDateValue
        End Get
        Set(ByVal value As Date)
            birthDateValue = value
        End Set
    End Property
End Class
