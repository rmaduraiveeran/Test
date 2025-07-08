Option Strict On
Option Explicit On
Imports System.Reflection

Public Class DedCode

    Public Property DedDedCode As String
    Public Property DedLongDesc As String
    Public Property DedDedType As String
    Public Property DedReportCategory As String
    Public Property DedTaxCategory As String
    Public Property DedDedEffStartDate As String
    Public Property DedDedEffStopDate As String
    Public Property DedDatetimeCreated As String

    ''' <summary>
    ''' Gets/Sets a property
    ''' </summary>
    ''' <param name="name">Property name</param>
    ''' <returns></returns>
    Public Property ByName(ByVal name As String) As Object
        Get
            Dim oType As Type = GetType(DedCode)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)
            Return propInfo.GetValue(Me, Nothing)
        End Get

        Set(value As Object)
            Dim oType As Type = GetType(DedCode)
            Dim propInfo As PropertyInfo = oType.GetProperty(name)

            If propInfo.PropertyType.Name.ToUpper = "INT32" Then
                propInfo.SetValue(Me, CInt(value), Nothing)
            Else
                propInfo.SetValue(Me, value, Nothing)
            End If

        End Set
    End Property

End Class
