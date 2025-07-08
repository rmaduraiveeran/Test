Option Strict On
Option Explicit On
Imports System.Reflection

Public Class W4Data


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


    Private state_Tax_CdValue As String
    ''' <summary>
    ''' Gets or sets the state state tax code
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


    Private tax_BlockValue As String
    ''' <summary>
    ''' Gets or sets the tax block
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TAX_BLOCK() As String
        Get
            Return tax_BlockValue
        End Get
        Set(ByVal value As String)
            tax_BlockValue = value
        End Set
    End Property


    Private marital_StatusValue As String
    ''' <summary>
    ''' Gets or sets the state marital status
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MARITAL_STATUS() As String
        Get
            Return marital_StatusValue
        End Get
        Set(ByVal value As String)
            marital_StatusValue = value
        End Set
    End Property


    Private exemptionsValue As String
    ''' <summary>
    ''' Gets or sets the state exemptions
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EXEMPTIONS() As String
        Get
            Return exemptionsValue
        End Get
        Set(ByVal value As String)
            exemptionsValue = value
        End Set
    End Property


    Private exempt_DollarsyValue As String
    ''' <summary>
    ''' Gets or sets the additional exemptions
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EXEMPT_DOLLARS() As String
        Get
            Return exempt_DollarsyValue
        End Get
        Set(ByVal value As String)
            exempt_DollarsyValue = value
        End Set
    End Property


    Private addl_Tax_AmtValue As Int32
    ''' <summary>
    ''' Gets or sets the additional state tax amount in dollars (whole number)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ADDL_TAX_AMT() As Int32
        Get
            Return addl_Tax_AmtValue
        End Get
        Set(ByVal value As Int32)
            addl_Tax_AmtValue = value
        End Set
    End Property


    Private state_Wh_TableValue As String
    ''' <summary>
    ''' Gets or sets the state withholding value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property STATE_WH_TABLE() As String
        Get
            Return state_Wh_TableValue
        End Get
        Set(ByVal value As String)
            state_Wh_TableValue = value
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

    ''' <summary>
    ''' Gets/Sets a property
    ''' </summary>
    ''' <param name="name">Property name</param>
    ''' <returns></returns>
    Public Property ByName(ByVal name As String) As Object
        Get
            Dim oType As Type = GetType(W4Data)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)
            Return propInfo.GetValue(Me, Nothing)
        End Get

        Set(value As Object)
            Dim oType As Type = GetType(W4Data)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)

            If propInfo.PropertyType.Name.ToUpper = "INT32" Then
                propInfo.SetValue(Me, CInt(value), Nothing)
            Else
                propInfo.SetValue(Me, value, Nothing)
            End If

        End Set
    End Property

End Class
