Option Strict On
Option Explicit On

Imports System.Collections.ObjectModel

Public Class PayControlValue

    Public Property Key As String
    Public Property Value1 As String
    Public Property Value2 As String
End Class

Public Module PayControlValueHandler
    ''' <summary>
    ''' Loads a collection of PayControlValue by Key
    ''' </summary>
    ''' <param name="pcv"></param>
    ''' <param name="key"></param>
    ''' <returns></returns>
    Public Function GetPayControlValuesByKey(ByRef pcv As Collection(Of PayControlValue), ByVal key As String) As Collection(Of PayControlValue)
        Try
            Return New Collection(Of PayControlValue)((From i In pcv
                                                       Where i.Key = key
                                                       Select New PayControlValue With {
                                                           .Key = i.Key,
                                                           .Value1 = i.Value1,
                                                           .Value2 = i.Value2
                                                        }).ToList())
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Loads a collection of PayControlValue by Key
    ''' </summary>
    ''' <param name="pcv"></param>
    ''' <param name="key"></param>
    ''' <param name="value1"></param>
    ''' <returns></returns>
    Public Function GetPayControlValuesByKeyandvalue(ByRef pcv As Collection(Of PayControlValue), ByVal key As String, ByVal value1 As String) As Collection(Of PayControlValue)
        Try
            Return New Collection(Of PayControlValue)((From i In pcv
                                                       Where i.Key = key And
                                                           i.Value1 = value1
                                                       Select New PayControlValue With {
                                                           .Key = i.Key,
                                                           .Value1 = i.Value1,
                                                           .Value2 = i.Value2
                                                        }).ToList())
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Returns a delimited string of value1 by key from a collection of PayControlvalue
    ''' </summary>
    ''' <param name="pcv"></param>
    ''' <param name="key"></param>
    ''' <param name="delimiter"></param>
    ''' <returns></returns>
    Public Function ConcatPayControlValue1(ByRef pcv As Collection(Of PayControlValue), ByVal key As String, delimiter As String) As String
        Dim cs As String
        Try
            Dim cv As Collection(Of PayControlValue) = GetPayControlValuesByKey(pcv, key)

            cs = String.Join(delimiter, cv.Select(Function(c) c.Value1))

            Return cs
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Module
