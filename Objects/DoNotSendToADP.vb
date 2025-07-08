Option Strict On
Option Explicit On


Public Class DoNotSendToADP

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

End Class
