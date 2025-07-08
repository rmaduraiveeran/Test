Option Strict On
Option Explicit On


Public Class VIPData
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

    Private empNoValue As String
    ''' <summary>
    ''' Gets or sets the employee number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EMPNO() As String
        Get
            Return empNoValue
        End Get
        Set(ByVal value As String)
            empNoValue = value
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

    Private fileNames As String
    ''' <summary>
    ''' Gets or sets the change in data file names 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FILE_NAMES() As String
        Get
            Return fileNames
        End Get
        Set(ByVal value As String)
            fileNames = value
        End Set
    End Property
End Class
