Option Strict On
Option Explicit On
Imports System.Reflection

Public Class JMSContingentData

    Private jmsEEIDValue As String
    Public Property jmsEEID() As String
        Get
            Return jmsEEIDValue
        End Get
        Set(ByVal value As String)
            jmsEEIDValue = value
        End Set
    End Property

    Private jmsCOIDValue As String
    Public Property jmsCOID() As String
        Get
            Return jmsCOIDValue
        End Get
        Set(ByVal value As String)
            jmsCOIDValue = value
        End Set
    End Property
    Private jmsLocationValue As String
    Public Property jmsLocation() As String
        Get
            Return jmsLocationValue
        End Get
        Set(ByVal value As String)
            jmsLocationValue = value
        End Set
    End Property
    Private jmPayGroupValue As String
    Public Property jmsPayGroup() As String
        Get
            Return jmPayGroupValue
        End Get
        Set(ByVal value As String)
            jmPayGroupValue = value
        End Set
    End Property
    Private jmsEmpNoValue As String
    Public Property jmsEmpNo() As String
        Get
            Return jmsEmpNoValue
        End Get
        Set(ByVal value As String)
            jmsEmpNoValue = value
        End Set
    End Property
    Private jmsEmplStatusValue As String
    Public Property jmsEmplStatus() As String
        Get
            Return jmsEmplStatusValue
        End Get
        Set(ByVal value As String)
            jmsEmplStatusValue = value
        End Set
    End Property
    Private jmsJobCodeValue As String
    Public Property jmsJobCode() As String
        Get
            Return jmsJobCodeValue
        End Get
        Set(ByVal value As String)
            jmsJobCodeValue = value
        End Set
    End Property
    Private jmsFullTimeOrPartTimeValue As String
    Public Property jmsFullTimeOrPartTime() As String
        Get
            Return jmsFullTimeOrPartTimeValue
        End Get
        Set(ByVal value As String)
            jmsFullTimeOrPartTimeValue = value
        End Set
    End Property
    Private jmsOrgLvl1Value As String
    Public Property jmsOrgLvl1() As String
        Get
            Return jmsOrgLvl1Value
        End Get
        Set(ByVal value As String)
            jmsOrgLvl1Value = value
        End Set
    End Property
    Private jmsOrgLvl3Value As String
    Public Property jmsOrgLvl3() As String
        Get
            Return jmsOrgLvl3Value
        End Get
        Set(ByVal value As String)
            jmsOrgLvl3Value = value
        End Set
    End Property
    Private jmsOrgLvl4Value As String
    Public Property jmsOrgLvl4() As String
        Get
            Return jmsOrgLvl4Value
        End Get
        Set(ByVal value As String)
            jmsOrgLvl4Value = value
        End Set
    End Property
    Private jmsSalaryOrHourlyValue As String
    Public Property jmsSalaryOrHourly() As String
        Get
            Return jmsSalaryOrHourlyValue
        End Get
        Set(ByVal value As String)
            jmsSalaryOrHourlyValue = value
        End Set
    End Property
    Private jmsJobChangeReasonValue As String
    Public Property jmsJobChangeReason() As String
        Get
            Return jmsJobChangeReasonValue
        End Get
        Set(ByVal value As String)
            jmsJobChangeReasonValue = value
        End Set
    End Property
    Private jmsSupervisorIDValue As String
    Public Property jmsSupervisorID() As String
        Get
            Return jmsSupervisorIDValue
        End Get
        Set(ByVal value As String)
            jmsSupervisorIDValue = value
        End Set
    End Property
    Private jmsPendingEffDateValue As DateTime
    Public Property jmsPendingEffDate() As DateTime
        Get
            Return jmsPendingEffDateValue
        End Get
        Set(ByVal value As DateTime)
            jmsPendingEffDateValue = value
        End Set
    End Property
    Private jmsCompanyCodeValue As String
    Public Property CompanyCode() As String
        Get
            Return jmsCompanyCodeValue
        End Get
        Set(ByVal value As String)
            jmsCompanyCodeValue = value
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
            Dim oType As Type = GetType(JMSContingentData)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)
            Return propInfo.GetValue(Me, Nothing)
        End Get

        Set(value As Object)
            Dim oType As Type = GetType(JMSContingentData)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)

            If Not propInfo Is Nothing Then
                If propInfo.PropertyType.Name.ToUpper = "INT32" Then
                    propInfo.SetValue(Me, CInt(value), Nothing)
                ElseIf propInfo.PropertyType.Name.ToUpper = "DATETIME" And value IsNot Nothing Then
                    If value.ToString.Trim.Length > 0 Then
                        propInfo.SetValue(Me, CDate(value), Nothing)
                    End If
                Else
                    propInfo.SetValue(Me, value, Nothing)
                End If
            End If


        End Set
    End Property
End Class
