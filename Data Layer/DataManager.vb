Option Strict On
Option Explicit On


Imports System.Data.SqlClient
Imports System.Collections.ObjectModel
Imports System.Xml
Imports Ultipro.RaaS
Imports System.Net.Mail

Public Class DataManager
    Public Shared SendStandardNotificationDelegate As Action(Of String, String)

    ''' <summary>
    ''' Indicates the type of data to return
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TypeOfData
        Before = 0
        Current = 1
        Deleted = 2
        History = 3
    End Enum
    ''' <summary>
    ''' Loads a PayControlValue collection with Control Values
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function LoadPayControlValues() As Collection(Of PayControlValue)
        Dim cv As Collection(Of PayControlValue) = Nothing
        Dim params As Collection = Nothing
        Try
            params = New Collection()

            Using tblCV As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure("Payroll.dbo.usp_PayGetControlTableValues", DataAccess.StoredProcedureReturnType.DataTable, params), DataTable)
                cv = New Collection(Of PayControlValue)((From rec In tblCV.AsEnumerable()
                                                         Select New PayControlValue With {
                                                             .Key = rec.Field(Of String)("ctlKey").Trim,
                                                             .Value1 = rec.Field(Of String)("ctlValue1"),
                                                             .Value2 = rec.Field(Of String)("ctlValue2")
                                                             }).ToList())

            End Using

            Return cv

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' This method gets the latest contingent data from SaaS and insert into tblJmsContingentEmpChanges
    ''' </summary>
    Public Shared Sub InsertJMSContingentData()
        Try
            Dim raasParam As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String)) 'As there is no param is needed to get contingent data we are sending empty param

            Dim reportRaasName As String = PayControlValueHandler.
                                            GetPayControlValuesByKey(payControlValuesCollection, "RAAS_JMS_CONTINGENT").
                                            Select(Function(cv) cv.Value1).FirstOrDefault()
            Dim reportXml As XmlDocument = RaaS_Service.
                                            GetReportResults(reportRaasName, raasParam)

            TruncateInsertContingentEmpChanges(
                ConvertCollectionToDataTable(
                ParseXmlToJMSContingentData(reportXml)))

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' This method parse the RaaS report data and convert them to JMSContingentData model
    ''' </summary>
    ''' <param name="xml"></param>
    ''' <returns>Returns collection of JMSContingentData</returns>
    Private Shared Function ParseXmlToJMSContingentData(xml As XmlDocument) As Collection(Of JMSContingentData)
        Try
            Dim xmlNodeList As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'metadata']/*[local-name() = 'item']")
            Dim dataRows As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'data']/*[local-name() = 'row']")

            Dim jmsContingentsData As Collection(Of JMSContingentData) = New Collection(Of JMSContingentData)()

            For Each node As XmlNode In dataRows
                Dim jmsContingentData As JMSContingentData = New JMSContingentData()
                Dim idx As Integer = 0

                For idx = 0 To xmlNodeList.Count - 1
                    Dim nodeName As String = xmlNodeList.Item(idx).SelectSingleNode("./@name").Value
                    Dim nodeValue As String = node.SelectNodes("./*[local-name() = 'value']").Item(idx).InnerText.Trim()

                    'NOTE: The nodeName has to match the property name for this to work.
                    jmsContingentData.ByName(nodeName) = nodeValue

                Next

                jmsContingentsData.Add(jmsContingentData)
            Next

            Return jmsContingentsData
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Reuturns a collection of type PersonalData
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadPersonalData(returnTypeOfData As TypeOfData) As Collection(Of PersonalData)
        Dim personalDataCollection As Collection(Of PersonalData) = Nothing
        Dim params As Collection = Nothing
        Dim storedProcName As String = String.Empty

        Try
            params = New Collection
            personalDataCollection = New Collection(Of PersonalData)

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    storedProcName = "Payroll.dbo.usp_PayGetPersonalDataBefore"
                Case TypeOfData.Current
                    storedProcName = "Payroll.dbo.usp_PayGetPersonalDataCurrent"
            End Select

            'Datatable has been used instead of SqlDataReader
            Using tblPersonalData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                storedProcName,
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))


                If tblPersonalData.Rows.Count > 0 Then
                    personalDataCollection = New Collection(Of PersonalData)((From rec In tblPersonalData.AsEnumerable()
                                                                              Select New PersonalData With {
                                                                                      .EMPLID = rec.Field(Of String)("EMPLID").Trim,
                                                                                      .PAYGROUP = rec.Field(Of String)("PAYGROUP").Trim,
                                                                                      .PAY_FREQUENCY = rec.Field(Of String)("PAY_FREQUENCY").Trim,
                                                                                      .FIRST_NAME = rec.Field(Of String)("FIRST_NAME").Trim,
                                                                                      .MIDDLE_NAME = rec.Field(Of String)("MIDDLE_NAME").Trim,
                                                                                      .LAST_NAME = rec.Field(Of String)("LAST_NAME").Trim,
                                                                                      .STREET1 = rec.Field(Of String)("STREET1").Trim,
                                                                                      .STREET2 = rec.Field(Of String)("STREET2").Trim,
                                                                                      .CITY = rec.Field(Of String)("CITY").Trim,
                                                                                      .STATE = rec.Field(Of String)("STATE").Trim,
                                                                                      .ZIP = rec.Field(Of String)("ZIP").Trim,
                                                                                      .HOME_PHONE = rec.Field(Of String)("HOME_PHONE").Trim,
                                                                                      .SSN = rec.Field(Of String)("SSN").Trim,
                                                                                      .ORIG_HIRE_DT = rec.Field(Of Date)("ORIG_HIRE_DT"),
                                                                                      .SEX = rec.Field(Of String)("SEX").Trim,
                                                                                      .BIRTHDATE = rec.Field(Of Date)("BIRTHDATE")
                                                                                   }).ToList())
                End If
                Return personalDataCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns a collection of type Employment
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadEmploymentData(returnTypeOfData As TypeOfData) As Collection(Of Employment)
        Dim employmentDataCollection As Collection(Of Employment) = Nothing
        Dim params As Collection = Nothing
        Dim storedProcName As String = String.Empty

        Try
            params = New Collection
            employmentDataCollection = New Collection(Of Employment)

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    storedProcName = "Payroll.dbo.usp_PayGetEmploymentBefore"
                Case TypeOfData.Current
                    storedProcName = "Payroll.dbo.usp_PayGetEmploymentCurrent"
            End Select
            'Datatable has been used instead of SqlDataReader
            Using tblEmploymentlData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                storedProcName,
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))

                If tblEmploymentlData.Rows.Count > 0 Then
                    For Each rec As DataRow In tblEmploymentlData.Rows
                        Dim employment As New Employment

                        employment.EMPLID = rec("EMPLID").ToString.Trim
                        employment.PAYGROUP = rec("PAYGROUP").ToString.Trim
                        employment.PAY_FREQUENCY = rec("PAY_FREQUENCY").ToString.Trim
                        employment.HIRE_DT = DirectCast(rec("HIRE_DT"), Date)
                        If Not IsDBNull(rec("REHIRE_DT")) Then
                            employment.REHIRE_DT = DirectCast(rec("REHIRE_DT"), Date)
                        Else
                            employment.REHIRE_DT = Nothing
                        End If
                        employment.CMPNY_SENIORITY_DT = DirectCast(rec("CMPNY_SENIORITY_DT"), Date)
                        If Not IsDBNull(rec("TERMINATION_DT")) Then
                            employment.TERMINATION_DT = DirectCast(rec("TERMINATION_DT"), Date)
                        Else
                            employment.TERMINATION_DT = Nothing
                        End If
                        If Not IsDBNull(rec("LAST_DATE_WORKED")) Then
                            employment.LAST_DATE_WORKED = DirectCast(rec("LAST_DATE_WORKED"), Date)
                        Else
                            employment.LAST_DATE_WORKED = Nothing
                        End If
                        employment.BUSINESS_TITLE = rec("BUSINESS_TITLE").ToString.Trim
                        employment.SUPERVISOR_ID = rec("SUPERVISOR_ID").ToString.Trim

                        employmentDataCollection.Add(employment)
                    Next
                End If
                Return employmentDataCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns a collection of type Job
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadJobData(returnTypeOfData As TypeOfData) As Collection(Of Job)
        Dim jobDataCollection As Collection(Of Job) = Nothing
        Dim params As Collection = Nothing
        Dim storedProcName As String = String.Empty

        Try
            params = New Collection
            jobDataCollection = New Collection(Of Job)

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    storedProcName = "Payroll.dbo.usp_PayGetJobBefore"
                Case TypeOfData.Current
                    storedProcName = "Payroll.dbo.usp_PayGetJobCurrent"
            End Select
            'Datatable has been used instead of SqlDataReader
            Using tbljobData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                storedProcName,
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))

                If tbljobData.Rows.Count > 0 Then
                    jobDataCollection = New Collection(Of Job)((From rec In tbljobData.AsEnumerable()
                                                                Select New Job With {
                                                                        .EMPLID = rec.Field(Of String)("EMPLID").Trim,
                                                                        .PAYGROUP = rec.Field(Of String)("PAYGROUP").Trim,
                                                                        .PAY_FREQUENCY = rec.Field(Of String)("PAY_FREQUENCY").Trim,
                                                                        .EMPL_STATUS = rec.Field(Of String)("EMPL_STATUS").Trim,
                                                                        .ACTION_REASON = rec.Field(Of String)("ACTION_REASON").Trim,
                                                                        .LOCATION = rec.Field(Of String)("LOCATION").Trim,
                                                                        .FULL_PART_TIME = rec.Field(Of String)("FULL_PART_TIME").Trim,
                                                                        .COMPANY = rec.Field(Of String)("COMPANY").Trim,
                                                                        .EMPL_TYPE = rec.Field(Of String)("EMPL_TYPE").Trim,
                                                                        .EMPL_CLASS = rec.Field(Of String)("EMPL_CLASS").Trim,
                                                                        .DATA_CONTROL = rec.Field(Of String)("DATA_CONTROL").Trim,
                                                                        .FILE_NBR = rec.Field(Of String)("FILE_NBR").Trim,
                                                                        .HOME_DEPARTMENT = rec.Field(Of String)("HOME_DEPARTMENT").Trim,
                                                                        .TITLE = rec.Field(Of String)("TITLE").Trim,
                                                .WORKERS_COMP_CD = rec.Field(Of String)("WORKERS_COMP_CD").Trim
                                                                     }).ToList())
                End If

                Return jobDataCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Loads the generalDeductionsCurrentCollection and generalDeductionsCurrentAllCollection
    ''' </summary>
    Public Shared Sub LoadGeneralDeductionsCurrent()
        'Dim params As Collection(Of SqlParameter) = Nothing
        Dim params As List(Of SqlParameter) = Nothing

        Try
            Using ds As DataSet = RemoveEmployeesWhoHasDataErrorsFromDataset(Ashley.DataAccess.SQLDataAccess.GetDataSet("UltiPro", "Payroll.dbo.usp_PayGetGeneralDeductionsCurrent", params, 600))
                If ds.Tables.Count = 0 Then
                    Return
                End If


                ' Initialize the collections.
                generalDeductionsCurrentCollection = New Collection(Of GeneralDeduction)
                generalDeductionsCurrentAllCollection = New Collection(Of GeneralDeduction)

                ' This is our general deductions current with the IGNORE deduction filters applied. 
                If ds.Tables.Item(0).Rows.Count > 0 Then

                    For Each rec As DataRow In ds.Tables.Item(0).Rows
                        Dim generalDeductions As New GeneralDeduction

                        generalDeductions.EMPLID = rec("EMPLID").ToString.Trim
                        generalDeductions.PAYGROUP = rec("PAYGROUP").ToString.Trim
                        generalDeductions.PAY_FREQUENCY = rec("PAY_FREQUENCY").ToString.Trim
                        generalDeductions.DEDCD = rec("DEDCD").ToString.Trim

                        If rec("CalcRuleAccountedFor").ToString = "NO" Then
                            ' Throw Error
                            Throw New InvalidOperationException("Calculation Rule Unaccounted For on Employee:" & Environment.NewLine &
                                                "EEID: " & rec("EecEEID").ToString & Environment.NewLine &
                                                "CoID: " & rec("EecCoID").ToString & Environment.NewLine &
                                                "Calc Rule: " & rec("DedEECalcRule").ToString & Environment.NewLine &
                                                "Ded Code: " & rec("DEDCD").ToString)
                        End If

                        generalDeductions.DED_ADDL_AMT = DirectCast(rec("DED_ADDL_AMT"), Decimal)
                        generalDeductions.DED_RATE_PCT = DirectCast(rec("DED_RATE_PCT"), Decimal)
                        generalDeductions.GOAL_AMT = DirectCast(rec("GOAL_AMT"), Decimal)
                        generalDeductions.END_DT = DirectCast(rec("END_DT"), Date)
                        generalDeductions.EmployeeName = rec("EMP_NAME").ToString.Trim
                        generalDeductions.SkipDeduction = rec("SkipDed").ToString.Trim
                        generalDeductions.SkipDeductionAccFor = rec("SkipDedAccFor").ToString.Trim
                        generalDeductions.EmployeeNumber = rec("EecEmpNo").ToString.Trim
                        generalDeductions.EECCoid = rec("EecCoID").ToString
                        generalDeductions.SkipZeroGoalAmount = rec("SkipZeroGoal").ToString

                        generalDeductions.GTD_Amt = DirectCast(rec("GTD_Amt"), Decimal)
                        generalDeductions.START_DT = DirectCast(rec("START_DT"), Date)

                        generalDeductionsCurrentCollection.Add(generalDeductions)
                    Next

                End If

                'This is our general deductions current without the filter applied.
                If ds.Tables.Item(1).Rows.Count > 0 Then

                    For Each rec As DataRow In ds.Tables.Item(1).Rows
                        Dim generalDeductions As New GeneralDeduction

                        generalDeductions.EMPLID = rec("EMPLID").ToString.Trim
                        generalDeductions.PAYGROUP = rec("PAYGROUP").ToString.Trim
                        generalDeductions.PAY_FREQUENCY = rec("PAY_FREQUENCY").ToString.Trim
                        generalDeductions.DEDCD = rec("DEDCD").ToString.Trim

                        If rec("DED_ADDL_AMT") Is System.DBNull.Value Then
                            generalDeductions.DED_ADDL_AMT = CType(0.0, Decimal)
                        Else
                            generalDeductions.DED_ADDL_AMT = DirectCast(rec("DED_ADDL_AMT"), Decimal)
                        End If

                        generalDeductions.DED_RATE_PCT = DirectCast(rec("DED_RATE_PCT"), Decimal)
                        generalDeductions.GOAL_AMT = DirectCast(rec("GOAL_AMT"), Decimal)
                        generalDeductions.END_DT = DirectCast(rec("END_DT"), Date)
                        generalDeductions.EmployeeName = rec("EMP_NAME").ToString.Trim
                        generalDeductions.SkipDeduction = rec("SkipDed").ToString.Trim
                        generalDeductions.SkipDeductionAccFor = rec("SkipDedAccFor").ToString.Trim
                        generalDeductions.EmployeeNumber = rec("EecEmpNo").ToString.Trim
                        generalDeductions.EECCoid = rec("EecCoID").ToString
                        generalDeductions.SkipZeroGoalAmount = rec("SkipZeroGoal").ToString

                        generalDeductions.GTD_Amt = DirectCast(rec("GTD_Amt"), Decimal)
                        generalDeductions.START_DT = DirectCast(rec("START_DT"), Date)

                        generalDeductionsCurrentAllCollection.Add(generalDeductions)
                    Next

                End If

            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Returns a collection of type GeneralDeduction
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadGeneralDeductionsData(returnTypeOfData As TypeOfData) As Collection(Of GeneralDeduction)
        Dim generalDeductionsCollection As Collection(Of GeneralDeduction) = Nothing
        Dim params As Collection = Nothing
        Dim storedProcName As String = String.Empty

        Try
            params = New Collection
            generalDeductionsCollection = New Collection(Of GeneralDeduction)

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    storedProcName = "Payroll.dbo.usp_PayGetGeneralDeductionsBefore"
                'Case TypeOfData.Current
                '    storedProcName = "Payroll.dbo.usp_PayGetGeneralDeductionsCurrent"
                Case TypeOfData.Deleted
                    storedProcName = "Payroll.dbo.usp_PayGetGeneralDeductionsDeleted"

                Case Else
                    Throw New InvalidOperationException("Function LoadGeneralDeductionsData: Invalid returnTypeOfData.")
            End Select
            'Datatable has been used instead of SqlDataReader
            Using tblGeneralDeductionsData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                storedProcName,
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))

                If tblGeneralDeductionsData.Rows.Count > 0 Then
                    For Each rec As DataRow In tblGeneralDeductionsData.Rows
                        Dim generalDeductions As New GeneralDeduction

                        generalDeductions.EMPLID = rec("EMPLID").ToString.Trim
                        generalDeductions.PAYGROUP = rec("PAYGROUP").ToString.Trim
                        generalDeductions.PAY_FREQUENCY = rec("PAY_FREQUENCY").ToString.Trim
                        generalDeductions.DEDCD = rec("DEDCD").ToString.Trim

                        'If returnTypeOfData = TypeOfData.Current Then
                        '    If rec("CalcRuleAccountedFor").ToString = "NO" Then
                        '        ' Throw Error
                        '        Throw New InvalidOperationException("Calculation Rule Unaccounted For on Employee:" & Environment.NewLine &
                        '                            "EEID: " & rec("EecEEID").ToString & Environment.NewLine &
                        '                            "CoID: " & rec("EecCoID").ToString & Environment.NewLine &
                        '                            "Calc Rule: " & rec("DedEECalcRule").ToString & Environment.NewLine &
                        '                            "Ded Code: " & rec("DEDCD").ToString)
                        '    End If
                        'End If

                        generalDeductions.DED_ADDL_AMT = DirectCast(rec("DED_ADDL_AMT"), Decimal)
                        generalDeductions.DED_RATE_PCT = DirectCast(rec("DED_RATE_PCT"), Decimal)
                        generalDeductions.GOAL_AMT = DirectCast(rec("GOAL_AMT"), Decimal)
                        generalDeductions.END_DT = DirectCast(rec("END_DT"), Date)
                        generalDeductions.EmployeeName = rec("EMP_NAME").ToString.Trim

                        'If returnTypeOfData = TypeOfData.Current Then
                        '    generalDeductions.SkipDeduction = rec("SkipDed").ToString.Trim
                        '    generalDeductions.SkipDeductionAccFor = rec("SkipDedAccFor").ToString.Trim
                        '    generalDeductions.EmployeeNumber = rec("EecEmpNo").ToString.Trim
                        'Else
                        generalDeductions.SkipDeduction = String.Empty
                        generalDeductions.SkipDeductionAccFor = String.Empty
                        generalDeductions.EmployeeNumber = rec("EecEmpNo").ToString.Trim
                        'End If

                        'If returnTypeOfData = TypeOfData.Before Then
                        'generalDeductions.GTD_Amt = DirectCast(rec("GTD_Amt"), Decimal)
                        'generalDeductions.START_DT = DirectCast(rec("START_DT"), Date)
                        'End If

                        generalDeductionsCollection.Add(generalDeductions)

                    Next
                End If


                Return generalDeductionsCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function LoadAcceptableDeductionsFromTo() As Collection(Of GeneralDeduction)
        Dim generalDeductionsCollection As Collection(Of GeneralDeduction) = Nothing
        Dim params As Collection = Nothing
        Try
            params = New Collection
            generalDeductionsCollection = New Collection(Of GeneralDeduction)

            Using tbl As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure("Payroll.dbo.usp_PayGetCompanyTransferFromToDeductions",
                                                                                    DataAccess.StoredProcedureReturnType.DataTable,
                                                                                    params), DataTable))

                If tbl.Rows.Count > 0 Then

                    ' map data to collection
                    For Each rec As DataRow In tbl.Rows

                        If AcceptableDedCodes.Where(Function(w) w.DedDedCode.Equals(rec("EedDedCode").ToString.Trim)).Select(Function(s) s).Count = 0 Then
                            Continue For
                        End If

                        Dim generalDeductions As GeneralDeduction = New GeneralDeduction
                        generalDeductions.EmployeeNumber = rec("EmpNo").ToString.Trim
                        generalDeductions.PAYGROUP = rec("CmpCompanyCode").ToString.Trim
                        generalDeductions.DEDCD = rec("EedDedCode").ToString.Trim

                        generalDeductions.START_DT = DirectCast(rec("EedStartDate"), Date)
                        generalDeductions.END_DT = DirectCast(rec("EedStopDate"), Date)

                        generalDeductions.DED_RATE_PCT = DirectCast(IIf(rec("EedEECalcRateOrPct").Equals(System.DBNull.Value), 0.00, rec("EedEECalcRateOrPct")), Decimal)
                        generalDeductions.DED_ADDL_AMT = DirectCast(IIf(rec("DED_ADDL_AMT").Equals(System.DBNull.Value), 0.00, rec("DED_ADDL_AMT")), Decimal)
                        generalDeductions.GOAL_AMT = DirectCast(IIf(rec("GOAL_AMT").Equals(System.DBNull.Value), 0.00, rec("GOAL_AMT")), Decimal)
                        generalDeductions.GTD_Amt = DirectCast(IIf(rec("EedEEGTDAmt").Equals(System.DBNull.Value), 0.00, rec("EedEEGTDAmt")), Decimal)

                        generalDeductionsCollection.Add(generalDeductions)

                    Next

                End If
            End Using

            Return generalDeductionsCollection

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Returns a collection of type DirectDeposit
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadDirectDepositData(returnTypeOfData As TypeOfData) As Collection(Of DirectDeposit)
        Dim directDepositCollection As Collection(Of DirectDeposit) = Nothing
        Dim params As Collection = Nothing
        Dim storedProcName As String = String.Empty

        Try
            params = New Collection
            directDepositCollection = New Collection(Of DirectDeposit)

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    storedProcName = "Payroll.dbo.usp_PayGetDirectDepositBefore"
                Case TypeOfData.Current
                    storedProcName = "Payroll.dbo.usp_PayGetDirectDepositCurrent"
            End Select
            'Datatable has been used instead of SqlDataReader
            Using tblDirectDepositData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                storedProcName,
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))

                If tblDirectDepositData.Rows.Count > 0 Then
                    directDepositCollection = New Collection(Of DirectDeposit)((From rec In tblDirectDepositData.AsEnumerable()
                                                                                Select New DirectDeposit With {
                                                                                    .EMPLID = rec.Field(Of String)("EMPLID").Trim,
                                                                                    .PAYGROUP = rec.Field(Of String)("PAYGROUP").Trim,
                                                                                    .PAY_FREQUENCY = rec.Field(Of String)("PAY_FREQUENCY").Trim,
                                                                                    .DEDCD = rec.Field(Of String)("DEDCD").Trim,
                                                                                    .FULL_DEPOSIT = rec.Field(Of String)("FULL_DEPOSIT").Trim,
                                                                                    .TRANSIT_NBR = rec.Field(Of String)("TRANSIT_NBR").Trim,
                                                                                    .ACCOUNT_NBR = rec.Field(Of String)("ACCOUNT_NBR").Trim,
                                                                                    .DEPOSIT_AMT = rec.Field(Of Decimal)("DEPOSIT_AMT"),
                                                                                    .END_DT = rec.Field(Of Date)("END_DT"),
                                                                                    .AccountIsInactive = rec.Field(Of String)("AccountIsInactive").Trim
                                                                                     }).ToList())
                    '  .FILE_NBR = rec.Field(Of String)("FILE_NBR").Trim
                End If

                Return directDepositCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns a collection of type TaxData
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadTaxData(returnTypeOfData As TypeOfData) As Collection(Of TaxData)
        Try

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    Return GetTaxDataBefore()
                Case TypeOfData.Current
                    Return GetTaxDataCurrent()
            End Select

            Return Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Shared Function GetTaxDataCurrent() As Collection(Of TaxData)
        Try
            Dim endPayPeriod,
                termCutOff,
                lodeCompSession,
                sendNonUsaEmpCodes,
                multpleJobFlagsCodes As String

            endPayPeriod = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "END_PAY_PERIOD").Select(Function(cv) cv.Value1).FirstOrDefault()
            termCutOff = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "TERM_CUT_OFF").Select(Function(cv) cv.Value1).FirstOrDefault()
            lodeCompSession = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "LODECOMP_SESSION", "|")
            sendNonUsaEmpCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "SEND_NON_USA_EMP", String.Empty)
            multpleJobFlagsCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "MULTIPLE_JOB_FLAGS", "|")


            Dim raasParam As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))

            raasParam.Add(New KeyValuePair(Of String, String)("END_PAY_PERIOD", endPayPeriod))
            raasParam.Add(New KeyValuePair(Of String, String)("TERM_CUT_OFF", termCutOff))
            raasParam.Add(New KeyValuePair(Of String, String)("LODECOMP_SESSION", lodeCompSession))
            raasParam.Add(New KeyValuePair(Of String, String)("SEND_NON_USA_EMP", sendNonUsaEmpCodes))
            raasParam.Add(New KeyValuePair(Of String, String)("MULTIPLE_JOB_FLAGS", multpleJobFlagsCodes))

            Dim reportRaasName As String = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "RAAS_TAX_GETCURRENT").Select(Function(cv) cv.Value1).FirstOrDefault()
            Dim reportXml As XmlDocument = RaaS_Service.GetReportResults(reportRaasName, raasParam)

            Return ParseXmlToTaxCollection(reportXml)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Shared Function ParseXmlToTaxCollection(xml As XmlDocument) As Collection(Of TaxData)
        Try
            Dim lstMetaDataItems As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'metadata']/*[local-name() = 'item']")
            Dim lstDataRows As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'data']/*[local-name() = 'row']")

            Dim taxData As Collection(Of TaxData) = New Collection(Of TaxData)()

            For Each node As XmlNode In lstDataRows
                Dim tmpTaxData As TaxData = New TaxData()
                Dim idx As Integer = 0

                For idx = 0 To lstMetaDataItems.Count - 1
                    Dim nodeName As String = lstMetaDataItems.Item(idx).SelectSingleNode("./@name").Value
                    Dim nodeValue As String = node.SelectNodes("./*[local-name() = 'value']").Item(idx).InnerText.Trim()

                    'Skip the xml nodenames here
                    If nodeName.Equals("FederalFilingStatus") Then
                        Continue For
                    End If
                    'NOTE: The nodeName has to match the property name for this to work.
                    tmpTaxData.ByName(nodeName) = nodeValue

                Next

                taxData.Add(tmpTaxData)
            Next

            Return taxData

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' Gets the compare table data for the Tax data that has already been sent to ADP.
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function GetTaxDataBefore() As Collection(Of TaxData)
        Dim params As Collection = Nothing
        Dim taxDataCollection As Collection(Of TaxData) = Nothing

        Try
            Using tblTaxData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                "Payroll.dbo.usp_PayGetTaxDataBefore",
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))

                If tblTaxData.Rows.Count > 0 Then
                    taxDataCollection = New Collection(Of TaxData)

                    For Each rec As DataRow In tblTaxData.Rows
                        Dim taxData As New TaxData

                        taxData.EMPLID = rec("EMPLID").ToString.Trim
                        taxData.PAYGROUP = rec("PAYGROUP").ToString.Trim
                        taxData.PAY_FREQUENCY = rec("PAY_FREQUENCY").ToString.Trim
                        taxData.FEDERAL_MAR_STATUS = rec("FEDERAL_MAR_STATUS").ToString.Trim
                        taxData.FED_ALLOWANCES = rec("FED_ALLOWANCES").ToString.Trim
                        taxData.FEDERAL_TAX_BLOCK = rec("FEDERAL_TAX_BLOCK").ToString.Trim
                        taxData.SCHDIST_TAX_BLOCK = rec("SCHDIST_TAX_BLOCK").ToString.Trim
                        taxData.SUISDI_TAX_BLOCK = rec("SUISDI_TAX_BLOCK").ToString.Trim
                        taxData.SSMED_TAX_BLOCK = rec("SSMED_TAX_BLOCK").ToString.Trim
                        taxData.FEDERAL_ADDL_AMT = DirectCast(rec("FEDERAL_ADDL_AMT"), Int32)
                        taxData.STATE_TAX_CD = rec("STATE_TAX_CD").ToString.Trim
                        taxData.STATE2_TAX_CD = rec("STATE2_TAX_CD").ToString.Trim
                        taxData.LOCAL_TAX_CD = rec("LOCAL_TAX_CD").ToString.Trim
                        taxData.LOCAL2_TAX_CD = rec("LOCAL2_TAX_CD").ToString.Trim
                        taxData.SCHOOL_DISTRICT = rec("SCHOOL_DISTRICT").ToString.Trim
                        taxData.SUI_TAX_CD = rec("SUI_TAX_CD").ToString.Trim
                        If Not IsDBNull(rec("TAX_LOCK_END_DT")) Then
                            taxData.TAX_LOCK_END_DT = DirectCast(rec("TAX_LOCK_END_DT"), Date)
                        Else
                            taxData.TAX_LOCK_END_DT = Nothing
                        End If
                        taxData.TAX_LCK_FED_MAR_ST = rec("TAX_LCK_FED_MAR_ST").ToString.Trim
                        taxData.TAX_LOCK_FED_ALLOW = rec("TAX_LOCK_FED_ALLOW").ToString.Trim
                        taxData.LOCAL4_TAX_CD = rec("LOCAL4_TAX_CD").ToString.Trim
                        taxData.W4IsLocked = rec("W4IsLocked").ToString.Trim
                        'taxData.USE_OLD_W4 = rec.Field(Of String)("USE_OLD_W4")
                        taxData.W4_FORM_YEAR = rec.Field(Of String)("W4_FORM_YEAR")
                        taxData.OTH_INCOME = rec.Field(Of String)("OTH_INCOME")
                        taxData.OTH_DEDUCTIONS = rec.Field(Of String)("OTH_DEDUCTIONS")
                        taxData.DEPENDENTS_AMT = rec.Field(Of String)("DEPENDENTS_AMT")
                        taxData.MULTIPLE_JOBS = rec.Field(Of String)("MULTIPLE_JOBS")
                        taxData.Long_Term_Care_Ins_Status = rec.Field(Of String)("LONG_TERM_STATUS")
                        'taxData.FILE_NBR = dr.Item("FILE_NBR").ToString.Trim ' !!! REMOVE after conversion is complete !!!
                        taxData.SEND_NOTFN_FOR_W4_FORM_YEAR = ""
                        taxData.EMP_No = ""
                        taxDataCollection.Add(taxData)
                    Next
                End If

                Return taxDataCollection
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Returns a collection of type W4Data
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadW4Data(returnTypeOfData As TypeOfData) As Collection(Of W4Data)

        Try

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    Return GetW4DateBefore()
                Case TypeOfData.Current
                    Return GetW4DataCurrent()
            End Select

            Return Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Gets the Current w4 Data from Ultipro
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function GetW4DataCurrent() As Collection(Of W4Data)
        Try
            'Get the parameters we need to pass to Raas
            Dim endPayPeriod, termCutOff,
            lodeCompSession, stateWhCodes,
            stateAddExceptionsCodes, w4ExemptStateCodes,
            sendNonUsaEmpCodes, multpleJobFlagsCodes,
            useOldW4FlagCodes As String

            endPayPeriod = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "END_PAY_PERIOD").Select(Function(cv) cv.Value1).FirstOrDefault()
            termCutOff = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "TERM_CUT_OFF").Select(Function(cv) cv.Value1).FirstOrDefault()
            lodeCompSession = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "LODECOMP_SESSION", "|")
            stateWhCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "STATE_WH_TABLE", "|")

            stateAddExceptionsCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "STATE_ADD_EXEMPTIONS", "|")
            w4ExemptStateCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "W4_EXEMPT_STATE", "|")
            sendNonUsaEmpCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "SEND_NON_USA_EMP", String.Empty)
            multpleJobFlagsCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "MULTIPLE_JOB_FLAGS", "|")
            useOldW4FlagCodes = PayControlValueHandler.ConcatPayControlValue1(payControlValuesCollection, "USE_OLD_W4_FLAG", "|")

            Dim raasParam As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))

            raasParam.Add(New KeyValuePair(Of String, String)("END_PAY_PERIOD", endPayPeriod))
            raasParam.Add(New KeyValuePair(Of String, String)("TERM_CUT_OFF", termCutOff))
            raasParam.Add(New KeyValuePair(Of String, String)("LODECOMP_SESSION", lodeCompSession))
            raasParam.Add(New KeyValuePair(Of String, String)("STATE_WH_TABLE", stateWhCodes))

            raasParam.Add(New KeyValuePair(Of String, String)("STATE_ADD_EXEMPTIONS", stateAddExceptionsCodes))
            raasParam.Add(New KeyValuePair(Of String, String)("W4_EXEMPT_STATE", w4ExemptStateCodes))
            raasParam.Add(New KeyValuePair(Of String, String)("SEND_NON_USA_EMP", sendNonUsaEmpCodes))
            raasParam.Add(New KeyValuePair(Of String, String)("MULTIPLE_JOB_FLAGS", multpleJobFlagsCodes))
            raasParam.Add(New KeyValuePair(Of String, String)("USE_OLD_W4_FLAGS", useOldW4FlagCodes))


            'Call the Raas Report service
            Dim reportRaasName As String = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "RAAS_W4_GETCURRENT").Select(Function(cv) cv.Value1).FirstOrDefault()
            Dim reportXml As XmlDocument = RaaS_Service.GetReportResults(reportRaasName, raasParam)

            'Parse the Xml to a collection
            Return RemoveEmployeesWhoHasDataErrorsFromCollection(ParseXmlToW4Collection(reportXml))
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' Parses the returned raas xml into a collection
    ''' </summary>
    ''' <param name="xml"></param>
    ''' <returns></returns>
    Private Shared Function ParseXmlToW4Collection(xml As XmlDocument) As Collection(Of W4Data)
        Try
            Dim lstMetaDataItems As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'metadata']/*[local-name() = 'item']")
            Dim lstDataRows As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'data']/*[local-name() = 'row']")

            Dim w4Data As Collection(Of W4Data) = New Collection(Of W4Data)()

            For Each node As XmlNode In lstDataRows
                Dim tmpW4Data As W4Data = New W4Data()
                Dim idx As Integer = 0

                For idx = 0 To lstMetaDataItems.Count - 1
                    Dim nodeName As String = lstMetaDataItems.Item(idx).SelectSingleNode("./@name").Value
                    Dim nodeValue As String = node.SelectNodes("./*[local-name() = 'value']").Item(idx).InnerText.Trim()

                    'Skip the xml nodenames here
                    If nodeName.Equals("FederalFilingStatus") Then
                        Continue For
                    End If
                    'NOTE: The nodeName has to match the property name for this to work.
                    tmpW4Data.ByName(nodeName) = nodeValue

                Next

                w4Data.Add(tmpW4Data)
            Next

            Return w4Data
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' Gets the compare table data of records that have already been sent to ADP.
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function GetW4DateBefore() As Collection(Of W4Data)
        Dim w4DataCollection As Collection(Of W4Data) = Nothing
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'w4DataCollection = New Collection(Of W4Data)

            Using tblW4Data As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                "Payroll.dbo.usp_PayGetW4DataBefore",
                DataAccess.StoredProcedureReturnType.DataTable,
                params), DataTable))

                If tblW4Data.Rows.Count > 0 Then
                    w4DataCollection = New Collection(Of W4Data)((From rec In tblW4Data.AsEnumerable()
                                                                  Select New W4Data With {
                                                                      .EMPLID = rec.Field(Of String)("EMPLID").Trim,
                                                                      .PAYGROUP = rec.Field(Of String)("PAYGROUP").Trim,
                                                                      .PAY_FREQUENCY = rec.Field(Of String)("PAY_FREQUENCY").Trim,
                                                                      .STATE_TAX_CD = rec.Field(Of String)("STATE_TAX_CD").Trim,
                                                                      .TAX_BLOCK = rec.Field(Of String)("TAX_BLOCK").Trim,
                                                                      .MARITAL_STATUS = rec.Field(Of String)("MARITAL_STATUS").Trim,
                                                                      .EXEMPTIONS = rec.Field(Of String)("EXEMPTIONS").Trim,
                                                                      .EXEMPT_DOLLARS = rec.Field(Of String)("EXEMPT_DOLLARS").Trim,
                                                                      .ADDL_TAX_AMT = rec.Field(Of Int32)("ADDL_TAX_AMT"),
                                                                      .STATE_WH_TABLE = rec.Field(Of String)("STATE_WH_TABLE").Trim
                                                                       }).ToList())
                    '  .FILE_NBR = rec.Field(Of String)("FILE_NBR").Trim
                End If

                Return w4DataCollection
            End Using


        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' Returns a collection of employees to not send to ADP
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadDoNotSendToADP() As Collection(Of DoNotSendToADP)
        Dim doNotSendToADPCollection As Collection(Of DoNotSendToADP) = Nothing

        Try
            doNotSendToADPCollection = New Collection(Of DoNotSendToADP)
            'Datatable has been used instead of SqlDataReader
            Using tblDoNotSendToADP As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                "Payroll.dbo.usp_PayGetDoNotSendToADPEmployees",
                DataAccess.StoredProcedureReturnType.DataTable,
                Nothing), DataTable)

                If tblDoNotSendToADP.Rows.Count > 0 Then
                    doNotSendToADPCollection = New Collection(Of DoNotSendToADP)((From rec In tblDoNotSendToADP.AsEnumerable()
                                                                                  Select New DoNotSendToADP With {
                                                                                    .EMPLID = rec.Field(Of String)("EMPLID").Trim,
                                                                                    .PAYGROUP = rec.Field(Of String)("PAYGROUP").Trim
                                                                                        }).ToList())
                End If

                Return doNotSendToADPCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns all the employees where the EecSalaryOrHourly is not in sync with their paygroup - if valid should return 0 results
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateEecSalaryOrHourly() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using eecSalaryOrHourlyValidationResults As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateEecSalaryOrHourly",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return eecSalaryOrHourlyValidationResults
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns all the employees that have multiple pending changes in LodEComp for a give pay period - if valid should return 0 results
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateLodEComp() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using loadECompValidationResults As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateLodEComp",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return loadECompValidationResults
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Inserts a PersonalData record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="first_Name"></param>
    ''' <param name="middle_Name"></param>
    ''' <param name="last_Name"></param>
    ''' <param name="street1"></param>
    ''' <param name="street2"></param>
    ''' <param name="city"></param>
    ''' <param name="state"></param>
    ''' <param name="zip"></param>
    ''' <param name="home_Phone"></param>
    ''' <param name="ssn"></param>
    ''' <param name="orig_Hire_Dt"></param>
    ''' <param name="sex"></param>
    ''' <param name="birthDate"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertPersonalData(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                         ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                         ByVal first_Name As String, ByVal middle_Name As String, ByVal last_Name As String,
                                         ByVal street1 As String, ByVal street2 As String, ByVal city As String,
                                         ByVal state As String, ByVal zip As String, ByVal home_Phone As String,
                                         ByVal ssn As String, ByVal orig_Hire_Dt As Date, ByVal sex As String,
                                         ByVal birthDate As Date, ByVal runID As Int32)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@FIRST_NAME", DbType.String, first_Name))
            params.Add(DataAccess.SetSQLParameterProperties("@MIDDLE_NAME", DbType.String, middle_Name))
            params.Add(DataAccess.SetSQLParameterProperties("@LAST_NAME", DbType.String, last_Name))
            params.Add(DataAccess.SetSQLParameterProperties("@STREET1", DbType.String, street1))
            params.Add(DataAccess.SetSQLParameterProperties("@STREET2", DbType.String, street2))
            params.Add(DataAccess.SetSQLParameterProperties("@CITY", DbType.String, city))
            params.Add(DataAccess.SetSQLParameterProperties("@STATE", DbType.String, state))
            params.Add(DataAccess.SetSQLParameterProperties("@ZIP", DbType.String, zip))
            params.Add(DataAccess.SetSQLParameterProperties("@HOME_PHONE", DbType.String, home_Phone))
            params.Add(DataAccess.SetSQLParameterProperties("@SSN", DbType.String, ssn))
            params.Add(DataAccess.SetSQLParameterProperties("@ORIG_HIRE_DT", DbType.DateTime, orig_Hire_Dt))
            params.Add(DataAccess.SetSQLParameterProperties("@SEX", DbType.String, sex))
            params.Add(DataAccess.SetSQLParameterProperties("@BIRTHDATE", DbType.DateTime, birthDate))
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertPersonalData",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts an Employment record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="hire_Dt"></param>
    ''' <param name="rehire_Dt"></param>
    ''' <param name="cmpny_Seniority_Dt"></param>
    ''' <param name="termination_Dt"></param>
    ''' <param name="last_Date_Worked"></param>
    ''' <param name="business_Title"></param>
    ''' <param name="supervisor_ID"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertEmployment(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                       ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                       ByVal hire_Dt As Date, ByVal rehire_Dt As Date, ByVal cmpny_Seniority_Dt As Date,
                                       ByVal termination_Dt As Date, ByVal last_Date_Worked As Date, business_Title As String,
                                       ByVal supervisor_ID As String, ByVal runID As Int32)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@HIRE_DT", DbType.DateTime, hire_Dt))
            If rehire_Dt <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@REHIRE_DT", DbType.DateTime, rehire_Dt))
            End If
            params.Add(DataAccess.SetSQLParameterProperties("@CMPNY_SENIORITY_DATE", DbType.DateTime, cmpny_Seniority_Dt))
            If termination_Dt <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@TERMINATION_DT", DbType.DateTime, termination_Dt))
            End If
            If last_Date_Worked <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@LAST_DATE_WORKED", DbType.DateTime, last_Date_Worked))
            End If
            params.Add(DataAccess.SetSQLParameterProperties("@BUSINESS_TITLE", DbType.String, business_Title))
            params.Add(DataAccess.SetSQLParameterProperties("@SUPERVISOR_ID", DbType.String, supervisor_ID))
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertEmployment",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts a Job record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="empl_Status"></param>
    ''' <param name="action_Reason"></param>
    ''' <param name="location"></param>
    ''' <param name="full_Part_Time"></param>
    ''' <param name="company"></param>
    ''' <param name="empl_Type"></param>
    ''' <param name="empl_Class"></param>
    ''' <param name="data_Control"></param>
    ''' <param name="file_Nbr"></param>
    ''' <param name="home_Department"></param>
    ''' <param name="title"></param>
    ''' <param name="workers_Cmp_Cd"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertJob(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                ByVal empl_Status As String, ByVal action_Reason As String, ByVal location As String,
                                ByVal full_Part_Time As String, ByVal company As String, ByVal empl_Type As String,
                                ByVal empl_Class As String, ByVal data_Control As String, ByVal file_Nbr As String,
                                ByVal home_Department As String, ByVal title As String, ByVal workers_Cmp_Cd As String,
                                ByVal runID As Int32)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPL_STATUS", DbType.String, empl_Status))
            params.Add(DataAccess.SetSQLParameterProperties("@ACTION_REASON", DbType.String, action_Reason))
            params.Add(DataAccess.SetSQLParameterProperties("@LOCATION", DbType.String, location))
            params.Add(DataAccess.SetSQLParameterProperties("@FULL_PART_TIME", DbType.String, full_Part_Time))
            params.Add(DataAccess.SetSQLParameterProperties("@COMPANY", DbType.String, company))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPL_TYPE", DbType.String, empl_Type))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPL_CLASS", DbType.String, empl_Class))
            params.Add(DataAccess.SetSQLParameterProperties("@DATA_CONTROL", DbType.String, data_Control))
            params.Add(DataAccess.SetSQLParameterProperties("@FILE_NBR", DbType.String, file_Nbr))
            params.Add(DataAccess.SetSQLParameterProperties("@HOME_DEPARTMENT", DbType.String, home_Department))
            params.Add(DataAccess.SetSQLParameterProperties("@TITLE", DbType.String, title))
            params.Add(DataAccess.SetSQLParameterProperties("@WORKERS_COMP_CD", DbType.String, workers_Cmp_Cd))
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertJob",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts a GernalDeduction record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="dedcd"></param>
    ''' <param name="ded_Addl_Amt"></param>
    ''' <param name="ded_Rate_Pct"></param>
    ''' <param name="goal_Amt"></param>
    ''' <param name="end_Dt"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertGeneralDeduction(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                             ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                             ByVal dedcd As String, ByVal ded_Addl_Amt As Decimal, ByVal ded_Rate_Pct As Decimal,
                                             ByVal goal_Amt As Decimal, ByVal end_Dt As Date, ByVal runID As Int32)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@DEDCD", DbType.String, dedcd))
            params.Add(DataAccess.SetSQLParameterProperties("@DED_ADDL_AMT", DbType.Decimal, ded_Addl_Amt))
            params.Add(DataAccess.SetSQLParameterProperties("@DED_RATE_PCT", DbType.Decimal, ded_Rate_Pct))
            params.Add(DataAccess.SetSQLParameterProperties("@GOAL_AMT", DbType.Decimal, goal_Amt))
            params.Add(DataAccess.SetSQLParameterProperties("@END_DT", DbType.DateTime, end_Dt))
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertGeneralDeduction",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts a DirectDeposit record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="dedcd"></param>
    ''' <param name="full_Deposit"></param>
    ''' <param name="transit_Nbr"></param>
    ''' <param name="account_Nbr"></param>
    ''' <param name="deposit_Amt"></param>
    ''' <param name="end_Dt"></param>
    ''' <param name="accountIsInactive"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertDirectDeposit(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                          ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                          ByVal dedcd As String, ByVal full_Deposit As String, ByVal transit_Nbr As String,
                                          ByVal account_Nbr As String, ByVal deposit_Amt As Decimal, ByVal end_Dt As Date,
                                          ByVal accountIsInactive As String, ByVal runID As Int32)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@DEDCD", DbType.String, dedcd))
            params.Add(DataAccess.SetSQLParameterProperties("@FULL_DEPOSIT", DbType.String, full_Deposit))
            params.Add(DataAccess.SetSQLParameterProperties("@TRANSIT_NBR", DbType.String, transit_Nbr))
            params.Add(DataAccess.SetSQLParameterProperties("@ACCOUNT_NBR", DbType.String, account_Nbr))
            params.Add(DataAccess.SetSQLParameterProperties("@DEPOSIT_AMT", DbType.Decimal, deposit_Amt))
            params.Add(DataAccess.SetSQLParameterProperties("@END_DT", DbType.DateTime, end_Dt))
            params.Add(DataAccess.SetSQLParameterProperties("@AccountIsInactive", DbType.String, accountIsInactive))
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertDirectDeposit",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts a TaxData record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="federal_Mar_Status"></param>
    ''' <param name="fed_Allowances"></param>
    ''' <param name="federal_Tax_Block"></param>
    ''' <param name="Schdist_Tax_Block"></param>
    ''' <param name="suisdi_Tax_Block"></param>
    ''' <param name="ssmed_Tax_Block"></param>
    ''' <param name="federal_Addl_Amt"></param>
    ''' <param name="state_Tax_Cd"></param>
    ''' <param name="state2_Tax_Cd"></param>
    ''' <param name="local_Tax_Cd"></param>
    ''' <param name="local2_Tax_Cd"></param>
    ''' <param name="school_District"></param>
    ''' <param name="sui_Tax_Cd"></param>
    ''' <param name="tax_Lock_End_Dt"></param>
    ''' <param name="tax_Lck_Fed_Mar_St"></param>
    ''' <param name="tax_Lock_Fed_Allow"></param>
    ''' <param name="local4_Tax_Cd"></param>
    ''' <param name="w4IsLocked"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    <CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="27")>
    <CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="26")>
    <CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId:="25")>
    Public Shared Sub InsertTaxData(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                    ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                    ByVal federal_Mar_Status As String, ByVal fed_Allowances As String, ByVal federal_Tax_Block As String,
                                    ByVal Schdist_Tax_Block As String, ByVal suisdi_Tax_Block As String, ByVal ssmed_Tax_Block As String,
                                    ByVal federal_Addl_Amt As Int32, ByVal state_Tax_Cd As String, ByVal state2_Tax_Cd As String,
                                    ByVal local_Tax_Cd As String, ByVal local2_Tax_Cd As String, ByVal school_District As String,
                                    ByVal sui_Tax_Cd As String, ByVal tax_Lock_End_Dt As DateTime, ByVal tax_Lck_Fed_Mar_St As String,
                                    ByVal tax_Lock_Fed_Allow As String, ByVal local4_Tax_Cd As String, ByVal w4IsLocked As String,
                                    ByVal runID As Int32,
                                   ByVal w4_FORM_YEAR As String,
                                   ByVal Oth_Income As String,
                                   ByVal oth_Deductions As String,
                                   ByVal dependents_Amt As String,
                                   ByVal multiple_Jobs As String,
                                   ByVal Long_Term_Care_Ins_Status As String)

        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@FEDERAL_MAR_STATUS", DbType.String, federal_Mar_Status))
            params.Add(DataAccess.SetSQLParameterProperties("@FED_ALLOWANCES", DbType.String, fed_Allowances))
            params.Add(DataAccess.SetSQLParameterProperties("@FEDERAL_TAX_BLOCK", DbType.String, federal_Tax_Block))
            params.Add(DataAccess.SetSQLParameterProperties("@SCHDIST_TAX_BLOCK", DbType.String, Schdist_Tax_Block))
            params.Add(DataAccess.SetSQLParameterProperties("@SUISDI_TAX_BLOCK", DbType.String, suisdi_Tax_Block))
            params.Add(DataAccess.SetSQLParameterProperties("@SSMED_TAX_BLOCK", DbType.String, ssmed_Tax_Block))
            params.Add(DataAccess.SetSQLParameterProperties("@FEDERAL_ADDL_AMT", DbType.Int32, federal_Addl_Amt))
            params.Add(DataAccess.SetSQLParameterProperties("@STATE_TAX_CD", DbType.String, state_Tax_Cd))
            params.Add(DataAccess.SetSQLParameterProperties("@STATE2_TAX_CD", DbType.String, state2_Tax_Cd))
            params.Add(DataAccess.SetSQLParameterProperties("@LOCAL_TAX_CD", DbType.String, local_Tax_Cd))
            params.Add(DataAccess.SetSQLParameterProperties("@LOCAL2_TAX_CD", DbType.String, local2_Tax_Cd))
            params.Add(DataAccess.SetSQLParameterProperties("@SCHOOL_DISTRICT", DbType.String, school_District))
            params.Add(DataAccess.SetSQLParameterProperties("@SUI_TAX_CD", DbType.String, sui_Tax_Cd))
            If tax_Lock_End_Dt <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@TAX_LOCK_END_DT", DbType.DateTime, tax_Lock_End_Dt))
            End If
            params.Add(DataAccess.SetSQLParameterProperties("@TAX_LCK_FED_MAR_ST", DbType.String, tax_Lck_Fed_Mar_St))
            params.Add(DataAccess.SetSQLParameterProperties("@TAX_LOCK_FED_ALLOW", DbType.String, tax_Lock_Fed_Allow))
            params.Add(DataAccess.SetSQLParameterProperties("@LOCAL4_TAX_CD", DbType.String, local4_Tax_Cd))
            params.Add(DataAccess.SetSQLParameterProperties("@W4IsLocked", DbType.String, w4IsLocked))

            'params.Add(DataAccess.SetSQLParameterProperties("@USE_OLD_W4", DbType.String, use_Old_W4))

            params.Add(DataAccess.SetSQLParameterProperties("@txdW4_FORM_YEAR", DbType.String, w4_FORM_YEAR))


            If Oth_Income.Length > 0 Then
                params.Add(DataAccess.SetSQLParameterProperties("@OTH_INCOME", DbType.Int32, CInt(Oth_Income)))
            End If
            If oth_Deductions.Length > 0 Then
                params.Add(DataAccess.SetSQLParameterProperties("@OTH_DEDUCTIONS", DbType.Int32, CInt(oth_Deductions)))
            End If
            If dependents_Amt.Length > 0 Then
                params.Add(DataAccess.SetSQLParameterProperties("@DEPENDENTS_AMT", DbType.Int32, CInt(dependents_Amt)))
            End If

            params.Add(DataAccess.SetSQLParameterProperties("@MULTIPLE_JOBS", DbType.String, multiple_Jobs))

            params.Add(DataAccess.SetSQLParameterProperties("@WACRFEE_TAX_BLOCK", DbType.String, Long_Term_Care_Ins_Status))

            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertTaxData",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts a W4Data record into the database
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="emplID"></param>
    ''' <param name="payGroup"></param>
    ''' <param name="pay_Frequency"></param>
    ''' <param name="state_Tax_Cd"></param>
    ''' <param name="tax_Block"></param>
    ''' <param name="marital_Status"></param>
    ''' <param name="exemptions"></param>
    ''' <param name="exempt_Dollars"></param>
    ''' <param name="addl_Tax_Amt"></param>
    ''' <param name="state_Wh_Table"></param>
    ''' <param name="runID"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertW4Data(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                   ByVal emplID As String, ByVal payGroup As String, ByVal pay_Frequency As String,
                                   ByVal state_Tax_Cd As String, ByVal tax_Block As String, ByVal marital_Status As String,
                                   ByVal exemptions As String, ByVal exempt_Dollars As String, ByVal addl_Tax_Amt As Int32,
                                   ByVal state_Wh_Table As String, ByVal runID As Int32)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@EMPLID", DbType.String, emplID))
            params.Add(DataAccess.SetSQLParameterProperties("@PAYGROUP", DbType.String, payGroup))
            params.Add(DataAccess.SetSQLParameterProperties("@PAY_FREQUENCY", DbType.String, pay_Frequency))
            params.Add(DataAccess.SetSQLParameterProperties("@STATE_TAX_CD", DbType.String, state_Tax_Cd))
            params.Add(DataAccess.SetSQLParameterProperties("@TAX_BLOCK", DbType.String, tax_Block))
            params.Add(DataAccess.SetSQLParameterProperties("@MARITAL_STATUS", DbType.String, marital_Status))
            params.Add(DataAccess.SetSQLParameterProperties("@EXEMPTIONS", DbType.String, exemptions))
            params.Add(DataAccess.SetSQLParameterProperties("@EXEMPT_DOLLARS", DbType.String, exempt_Dollars))
            params.Add(DataAccess.SetSQLParameterProperties("@ADDL_TAX_AMT", DbType.Int32, addl_Tax_Amt))
            params.Add(DataAccess.SetSQLParameterProperties("@STATE_WH_TABLE", DbType.String, state_Wh_Table))

            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertW4Data",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Inserts a run log record into the database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertRunLogRecord() As Int32
        Dim params As Collection = Nothing
        Dim runID As Int32 = Nothing

        Try
            params = New Collection
            runID = CInt(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayInsertFileCreationRunLog",
                                DataAccess.StoredProcedureReturnType.Scalar,
                                params))

            Return runID
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Updates the run log record with the final completion information (Non-Transaction Type)
    ''' </summary>
    ''' <param name="runID"></param>
    ''' <param name="persDataRecordCountSent"></param>
    ''' <param name="employmentRecordCountSent"></param>
    ''' <param name="jobRecordCountSent"></param>
    ''' <param name="genDedRecordCountSent"></param>
    ''' <param name="dirDepRecordCountSent"></param>
    ''' <param name="taxDataRecordCountSent"></param>
    ''' <param name="w4DataRecordCountSent"></param>
    ''' <param name="completionStatus"></param>
    ''' <param name="errorMessage"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateRunLogRecord(ByVal runID As Int32, ByVal persDataRecordCountSent As Int32,
                                         ByVal employmentRecordCountSent As Int32, ByVal jobRecordCountSent As Int32,
                                         ByVal genDedRecordCountSent As Int32, ByVal dirDepRecordCountSent As Int32,
                                         ByVal taxDataRecordCountSent As Int32, ByVal w4DataRecordCountSent As Int32,
                                         ByVal completionStatus As String, ByVal errorMessage As String)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@fileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@persDataRecordCountSent", DbType.Int32, persDataRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@employmentRecordCountSent", DbType.Int32, employmentRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@jobRecordCountSent", DbType.Int32, jobRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@genDedRecordCountSent", DbType.Int32, genDedRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@dirDepRecordCountSent", DbType.Int32, dirDepRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@taxDataRecordCountSent", DbType.Int32, taxDataRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@w4DataRecordCountSent", DbType.Int32, w4DataRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@completionStatus", DbType.String, completionStatus))
            If errorMessage <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@errorMessage", DbType.String, errorMessage))
            End If
            DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayUpdateFileCreationRunLog",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Updates the run log record with the final completion information (Transaction Type)
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="runID"></param>
    ''' <param name="persDataRecordCountSent"></param>
    ''' <param name="employmentRecordCountSent"></param>
    ''' <param name="jobRecordCountSent"></param>
    ''' <param name="genDedRecordCountSent"></param>
    ''' <param name="dirDepRecordCountSent"></param>
    ''' <param name="taxDataRecordCountSent"></param>
    ''' <param name="w4DataRecordCountSent"></param>
    ''' <param name="completionStatus"></param>
    ''' <param name="errorMessage"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateRunLogRecord(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction,
                                         ByVal runID As Int32, ByVal persDataRecordCountSent As Int32,
                                         ByVal employmentRecordCountSent As Int32, ByVal jobRecordCountSent As Int32,
                                         ByVal genDedRecordCountSent As Int32, ByVal dirDepRecordCountSent As Int32,
                                         ByVal taxDataRecordCountSent As Int32, ByVal w4DataRecordCountSent As Int32,
                                         ByVal completionStatus As String, ByVal errorMessage As String)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@fileCreateRunID", DbType.Int32, runID))
            params.Add(DataAccess.SetSQLParameterProperties("@persDataRecordCountSent", DbType.Int32, persDataRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@employmentRecordCountSent", DbType.Int32, employmentRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@jobRecordCountSent", DbType.Int32, jobRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@genDedRecordCountSent", DbType.Int32, genDedRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@dirDepRecordCountSent", DbType.Int32, dirDepRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@taxDataRecordCountSent", DbType.Int32, taxDataRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@w4DataRecordCountSent", DbType.Int32, w4DataRecordCountSent))
            params.Add(DataAccess.SetSQLParameterProperties("@completionStatus", DbType.String, completionStatus))
            If errorMessage <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@errorMessage", DbType.String, errorMessage))
            End If
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayUpdateFileCreationRunLog",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Sends an email notification to recipients defined in database when an error occurs
    ''' </summary>
    ''' <param name="errorMessage"></param>
    ''' <remarks></remarks>
    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="mode")>
    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="params")>
    Public Shared Sub SendErrorNotification(ByVal errorMessage As String, Optional ByVal ovrSubject As String = "")
        Dim params As Collection = Nothing
        Dim mode As String = String.Empty
        Dim cv As List(Of String) = Nothing
        Try
            If String.IsNullOrWhiteSpace(ovrSubject) Then
                ovrSubject = String.Empty
            End If

            ' Set the mode the application is going to run in.
            mode = System.Configuration.ConfigurationManager.AppSettings.Get("EmailMode")
            If String.IsNullOrEmpty(mode) Or (Not mode.Equals("PROD")) Then
                mode = "TEST"
            End If

            ' if Test mode override
            If (mode.Equals("TEST")) Then
                errorMessage = String.Concat("This email would have been sent to payroll group.", Environment.NewLine, Environment.NewLine, Environment.NewLine, errorMessage)
            End If
            params = New Collection
            'params.Add(DataAccess.SetSQLParameterProperties("@errorMessage", DbType.String, errorMessage))

            'If ovrSubject.Length > 0 Then
            '    params.Add(DataAccess.SetSQLParameterProperties("@ovrSubject", DbType.String, ovrSubject))
            'End If

            'DataAccess.ExecuteStoredProcedure(
            '                    "Payroll.dbo.usp_PaySendErrorNotification",
            '                    DataAccess.StoredProcedureReturnType.RowsAffected,
            '                    params)

            Dim modelEmail As ModelEmail = New ModelEmail
            modelEmail.MailBody = errorMessage
            Using tblCV As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure("Payroll.dbo.usp_PayGetErrorEmailRecipients", DataAccess.StoredProcedureReturnType.DataTable, params), DataTable)
                cv = New List(Of String)((From rec In tblCV.AsEnumerable()
                                          Select rec.Field(Of String)("ErrorMailRecipients")
                                                             ).ToList())

                modelEmail.MailRecipients = cv

                If String.IsNullOrWhiteSpace(ovrSubject) Then
                    modelEmail.MailSubject = ((From rec In tblCV.AsEnumerable()
                                               Select rec.Field(Of String)("DefaultErrorEmailSubject")
                                                                 ).ToList().FirstOrDefault())
                Else
                    modelEmail.MailSubject = ovrSubject
                End If

                modelEmail.SmtpClient = ((From rec In tblCV.AsEnumerable()
                                          Select rec.Field(Of String)("EmailSMTPHost")
                                                                 ).ToList().FirstOrDefault())
                modelEmail.MailSender = ((From rec In tblCV.AsEnumerable()
                                          Select rec.Field(Of String)("FromMailAddress")
                                                                 ).ToList().FirstOrDefault())
            End Using


            SendEmail(modelEmail, True)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub SendEmail(ByVal modelEmail As ModelEmail, ByVal bIsFormatHtml As Boolean)
        Try

            If modelEmail IsNot Nothing Then

                Using smtpClient As SmtpClient = New SmtpClient(modelEmail.SmtpClient)

                    Using email As MailMessage = New MailMessage()
                        email.Subject = modelEmail.MailSubject
                        email.[To].Add(String.Join(",", modelEmail.MailRecipients))

                        email.From = New MailAddress(modelEmail.MailSender)
                        email.Body = modelEmail.MailBody
                        email.IsBodyHtml = bIsFormatHtml
                        email.Priority = MailPriority.High

                        If modelEmail.MailAttachments IsNot Nothing Then

                            For Each attachment As String In modelEmail.MailAttachments
                                email.Attachments.Add(New Attachment(attachment))
                            Next
                        End If

                        smtpClient.Send(email)
                    End Using
                End Using
            End If

        Catch
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Sets the pay period dates in the database to use for all processing
    ''' </summary>
    ''' <param name="beginPayPeriodDate"></param>
    ''' <param name="endPayPeriodDate"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetPayPeriodDates(ByVal beginPayPeriodDate As Date, ByVal endPayPeriodDate As Date)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            If beginPayPeriodDate <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@beginPayPeriodDate", DbType.Date, beginPayPeriodDate))
            End If
            If endPayPeriodDate <> Nothing Then
                params.Add(DataAccess.SetSQLParameterProperties("@endPayPeriodDate", DbType.Date, endPayPeriodDate))
            End If
            DataAccess.ExecuteStoredProcedure(
                                            "Payroll.dbo.usp_PaySetPayPeriodDates",
                                            DataAccess.StoredProcedureReturnType.RowsAffected,
                                            params)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Returns the default file location where the files should be written to
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileLocation(ByVal locationType As String) As String
        Dim fileLocation As String = String.Empty
        Dim params As Collection = Nothing
        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@LocType", DbType.String, locationType))
            fileLocation = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayGetFileLocations",
                                DataAccess.StoredProcedureReturnType.Scalar,
                                params), String)

            Return fileLocation
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that there are no futerm terminations entered in Ultipro
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoFutureTermDates() As DataTable
        Dim params As Collection = Nothing
        Try
            params = New Collection
            Return DirectCast(DataAccess.ExecuteStoredProcedure(
                               "Payroll.dbo.usp_PayValidateNoFutureTermDates",
                               DataAccess.StoredProcedureReturnType.DataTable,
                               params), DataTable)
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that there are no unaccounted for salary employees reporting to an hourly employee
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoSalaryReprtingToHourly() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using futureTermDates As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateNoSalaryRepToHourly",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return futureTermDates
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates for employees that live in a different state than what they work in that either 'Not subject to work in state tax' or 
    ''' 'Not subject to resident state tax' option is selected.  It also validates that both options are not selected at the same time.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateWorkInLiveInStateTax() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using invalidWorkInLiveInStateTaxes As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateWorkInLiveInStateTax",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return invalidWorkInLiveInStateTaxes
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that there are no employees with a DISAL deduction and a benefit amount not equal to 0.00
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoBenAmtForDISAL() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using hasBenAmtForDISALDed As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                    "Payroll.dbo.usp_PayValidateNoBenAmtForDISAL",
                                    DataAccess.StoredProcedureReturnType.DataTable,
                                    params), DataTable)

                Return hasBenAmtForDISALDed
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that there are no employees with multiples of School, Occ, WC, or Other local taxes.  There can only be one
    ''' populated for any given employee.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoMultipleLocalTaxes() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using multipleLocalTaxes As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateNotMultipleLocalTaxes",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return multipleLocalTaxes
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that there are no employees with extra tax dollars that have cents populated in the dollar amount
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoCentsInExtraTaxDollars() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using extraTaxDollarsWithCents As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateNoCentsInExtraTaxAmount",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return extraTaxDollarsWithCents
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns identified employees if their tax records have changed
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetEmployeesWithTaxChanges(ByVal runID As Int32) As DataTable
        Dim params As Collection = Nothing
        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            'Datatable has been used instead of SqlDataReader
            Using employeesWithTaxChanges As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                    "Payroll.dbo.usp_PayGetTaxChangeEmployees",
                                    DataAccess.StoredProcedureReturnType.DataTable,
                                    params), DataTable)

                Return employeesWithTaxChanges
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Sends an email notification to recipients defined in database
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="mode")>
    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="params")>
    Public Shared Sub SendStandardNotification(ByVal type As String, ByVal message As String)
        Dim params As Collection = Nothing
        Dim mode As String = String.Empty
        Dim cv As List(Of String) = Nothing
        Try
            ' Set the mode the application is going to run in.
            mode = System.Configuration.ConfigurationManager.AppSettings.Get("EmailMode")
            If String.IsNullOrEmpty(mode) Or (Not mode.Equals("PROD")) Then
                mode = "TEST"
            End If

            ' if Test mode override
            If (mode.Equals("TEST")) Then
                message = String.Concat("This email would have been sent to payroll group.<br/><br/>", Environment.NewLine, Environment.NewLine, Environment.NewLine, message)
            End If

            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@type", DbType.String, type))

            'params.Add(DataAccess.SetSQLParameterProperties("@message", DbType.String, message))

            'DataAccess.ExecuteStoredProcedure(
            '                    "Payroll.dbo.usp_PaySendStandardNotification",
            '                    DataAccess.StoredProcedureReturnType.RowsAffected,
            '                    params)


            Dim modelEmail As ModelEmail = New ModelEmail
            modelEmail.MailBody = message
            Using tblCV As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure("Payroll.dbo.usp_PayStandardEmailRecipients", DataAccess.StoredProcedureReturnType.DataTable, params), DataTable)
                cv = New List(Of String)((From rec In tblCV.AsEnumerable()
                                          Select rec.Field(Of String)("MailRecipients")
                                                             ).ToList())

                modelEmail.MailRecipients = cv
                modelEmail.MailSubject = ((From rec In tblCV.AsEnumerable()
                                           Select rec.Field(Of String)("EmailSubject")
                                                             ).ToList().FirstOrDefault())

                modelEmail.SmtpClient = ((From rec In tblCV.AsEnumerable()
                                          Select rec.Field(Of String)("EmailSMTPHost")
                                                                 ).ToList().FirstOrDefault())
                modelEmail.MailSender = ((From rec In tblCV.AsEnumerable()
                                          Select rec.Field(Of String)("FromMailAddress")
                                                                 ).ToList().FirstOrDefault())
            End Using


            SendEmail(modelEmail, True)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Validates that there are no partial direct deposit accounts with a zero dollar amount
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoPartialDirDepAcctWithZeroDolAmt() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using partialDirDepWithZeroDollar As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidatePartialDirDepHasDolAmt",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return partialDirDepWithZeroDollar
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that there are no percent direct deposit rules for salary employees
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateNoPercentDirDepRules() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using percentDirDepRules As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayValidateNoPercentDirDepRules",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return percentDirDepRules
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that dependents are entered on all applicable deductions based on calc rule
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateDeductionDependentsExist() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using empsMissingDedDependents As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                    "Payroll.dbo.usp_PayValidateDedDependentPopulated",
                                    DataAccess.StoredProcedureReturnType.DataTable,
                                    params), DataTable)

                Return empsMissingDedDependents
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Validates that dependents tied to deductions have their DOB entered
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateDeductionDependentsDOB() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using empsMissingDependentDOB As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                    "Payroll.dbo.usp_PayValidateDedDependentDOB",
                                    DataAccess.StoredProcedureReturnType.DataTable,
                                    params), DataTable)

                Return empsMissingDependentDOB
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns the employees that were salary and moved to hourly or were hourly and moved to salary
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSalToHourOrHourToSalTransfers(ByVal runID As Int32) As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            'Datatable has been used instead of SqlDataReader
            Using salToHourOrHourToSalTransfers As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                    "Payroll.dbo.usp_PayGetSalToHourOrHourToSalEmployees",
                                    DataAccess.StoredProcedureReturnType.DataTable,
                                    params), DataTable)

                Return salToHourOrHourToSalTransfers
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Truncates the tblPayHourlyEmployees table then insertes the current hourly employee list
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateHourlyEmployees(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            Dim sqlParameter As New SqlParameter
            sqlParameter = New SqlParameter("@ErroredEmployees", dataValidationErrorsAsDataTable)
            sqlParameter.SqlDbType = SqlDbType.Structured
            params.Add(sqlParameter)
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayTruncateInsertHourlyEmployees",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub



    ''' <summary>
    ''' Returns Duplicate Employees from vewPayLatestEmpCompData table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDuplicateEmployees() As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using dtDuplicateEmployees As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_GetPayDuplicateEmployees",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return dtDuplicateEmployees
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Truncates the tblPaySalaryEmployees table then insertes the current salary employee list
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateSalaryEmployees(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            Dim sqlParameter As New SqlParameter
            sqlParameter = New SqlParameter("@ErroredEmployees", dataValidationErrorsAsDataTable)
            sqlParameter.SqlDbType = SqlDbType.Structured
            params.Add(sqlParameter)
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayTruncateInsertSalaryEmployees",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Returns the deductions that have changed for the specified employees and deductions
    ''' </summary>
    ''' <param name="runID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSpecificDeductionChanges(ByVal runID As Int32) As DataTable
        Dim params As Collection = Nothing

        Try
            params = New Collection

            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            'Datatable has been used instead of SqlDataReader
            Using secifiedDeductionChanges As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayGetDedChangeEmployees",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                Return secifiedDeductionChanges
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns the employees that have transferred companies but have retained the same employee number
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSameEmpNumberCompanyTransfers() As DataTable
        Dim params As Collection = Nothing
        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using sameEmpNumberEmployees As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                               "Payroll.dbo.usp_PayGetSameEmpNumCompTransfers",
                               DataAccess.StoredProcedureReturnType.DataTable,
                               params), DataTable)

                Return sameEmpNumberEmployees
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' Returns the employees who had company transfers to find the employee's state for “Worked in” or “Residence” tax changes.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetEmpsCompanyTransfers() As DataTable
        Dim params As Collection = Nothing
        Try
            params = New Collection
            'Datatable has been used instead of SqlDataReader
            Using sameEmpNumberEmployees As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                               "Payroll.dbo.usp_PayGetEmpsCompanyTransfers",
                               DataAccess.StoredProcedureReturnType.DataTable,
                               params), DataTable)

                Return sameEmpNumberEmployees
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Returns data of VIP control values from tblPayControlValues tables for the provided key
    ''' </summary>
    ''' <param name="controlkey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetVIPDetails(ByVal controlKey As String) As Collection(Of VIPData)
        Dim params As Collection = Nothing
        Dim vipDetails As Collection(Of VIPData) = Nothing
        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@controlkey", DbType.String, controlKey))

            vipDetails = New Collection(Of VIPData)
            'Datatable has been used instead of SqlDataReader
            Using tblVIPDetails As DataTable = DirectCast(DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayGetVipDetails",
                                DataAccess.StoredProcedureReturnType.DataTable,
                                params), DataTable)

                If tblVIPDetails.Rows.Count > 0 Then
                    vipDetails = New Collection(Of VIPData)((From rec In tblVIPDetails.AsEnumerable()
                                                             Select New VIPData With {
                                                                 .EMPLID = rec.Field(Of String)("EmpId").Trim,
                                                                 .FIRST_NAME = rec.Field(Of String)("FirstName").Trim,
                                                                 .LAST_NAME = rec.Field(Of String)("LastName").Trim,
                                                                 .EMPNO = rec.Field(Of String)("EmployeeNumber").Trim
                                                                  }).ToList())
                End If

            End Using
            Return vipDetails
        Catch ex As Exception
            Throw ex
        Finally
            If vipDetails Is Nothing Then
                vipDetails = Nothing
            End If

        End Try
    End Function

    ''' <summary>
    ''' Returns the Terminated Employee List from Employment Table
    ''' </summary>
    ''' <param name="returnTypeOfData"></param>
    ''' <returns></returns>
    Public Shared Function GetAllTerminatedEmployeeData(returnTypeOfData As TypeOfData) As Collection(Of TermedEmploymentData)
        Dim terminatedEmployeeDataCollection As Collection(Of TermedEmploymentData) = Nothing
        Dim params As Collection = Nothing
        Dim storedProcName As String = String.Empty
        Try
            params = New Collection
            terminatedEmployeeDataCollection = New Collection(Of TermedEmploymentData)

            Select Case returnTypeOfData
                Case TypeOfData.Before
                    storedProcName = "Payroll.dbo.usp_PayGetAllTerminatedDataBefore"
                Case TypeOfData.History
                    storedProcName = "Payroll.dbo.usp_PayGetTermedEmploymentHistory"
            End Select

            'Datatable has been used instead of SqlDataReader
            Using terminatedEmployementData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                               storedProcName,
                               DataAccess.StoredProcedureReturnType.DataTable,
                               params), DataTable), "epyEMPLID")

                If terminatedEmployementData.Rows.Count > 0 Then
                    terminatedEmployeeDataCollection = New Collection(Of TermedEmploymentData)((From rec In terminatedEmployementData.AsEnumerable()
                                                                                                Select New TermedEmploymentData With {
                                                                                                    .EMPLOYEE_NUMBER = rec.Field(Of String)("EecEmpNo").Trim,
                                                                                                    .EMPLID = rec.Field(Of String)("epyEMPLID").Trim,
                                                                                                    .EMPLOYEE_NAME = rec.Field(Of String)("Employee_Name").Trim,
                                                                                                    .PAYGROUP = rec.Field(Of String)("epyPAYGROUP").Trim,
                                                                                                    .PAY_FREQUENCY = rec.Field(Of String)("epyPAY_FREQUENCY").Trim,
                                                                                                    .BUSINESS_TITLE = rec.Field(Of String)("epyBUSINESS_TITLE").Trim,
                                                                                                    .CMPNY_SENIORITY_DT = rec.Field(Of String)("epyCMPNY_SENIORITY_DT"),
                                                                                                    .HIRE_DT = rec.Field(Of String)("epyHIRE_DT"),
                                                                                                    .LAST_DATE_WORKED = rec.Field(Of String)("epyLAST_DATE_WORKED"),
                                                                                                    .REHIRE_DT = rec.Field(Of String)("epyREHIRE_DT"),
                                                                                                    .SUPERVISOR_ID = rec.Field(Of String)("epySUPERVISOR_ID").Trim,
                                                                                                    .SUPERVISOR_NAME = rec.Field(Of String)("Supervisor_Name").Trim,
                                                                                                    .TERMINATION_DT = rec.Field(Of String)("epyTERMINATION_DT")
                                                                                                 }).ToList())
                End If
                Return terminatedEmployeeDataCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Inserts the Termed Employment logs
    ''' </summary>
    ''' <param name="sqlConn"></param>
    ''' <param name="sqlTrans"></param>
    ''' <param name="terminatedEmployeeCollection"></param>
    ''' <param name="runID"></param>
    Public Shared Sub InsertTermedEmploymentLogs(ByVal sqlConn As SqlConnection, ByVal sqlTrans As SqlTransaction, ByVal terminatedEmployeeCollection As IEnumerable(Of TermedEmploymentData), ByVal runID As Int32)
        Dim params As Collection = Nothing
        Try
            params = New Collection
            params.Add(DataAccess.SetSQLParameterProperties("@termedEmpInsertXml", DbType.Xml, BuildPayrollTermedEmploymentInsertXml(terminatedEmployeeCollection).ToString(SaveOptions.DisableFormatting)))
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))
            DataAccess.ExecuteStoredProcedureCommon(
                                sqlConn,
                                "Payroll.dbo.usp_PayInsertTermedEmploymentHistory",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params,
                                sqlTrans)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Returns the xml of Termed Employees
    ''' </summary>
    ''' <param name="terminatedEmployeeCollection"></param>
    ''' <returns></returns>
    Public Shared Function BuildPayrollTermedEmploymentInsertXml(terminatedEmployeeCollection As IEnumerable(Of TermedEmploymentData)) As XElement
        Try
            Return New XElement("PayrollTermedEmpDetails",
                                        From rec In terminatedEmployeeCollection
                                        Select New XElement("Record",
                                            New XElement("EMPLID", rec.EMPLID),
                                            New XElement("PAYGROUP", rec.PAYGROUP),
                                            New XElement("BUSINESS_TITLE", rec.BUSINESS_TITLE),
                                            New XElement("SUPERVISOR_ID", rec.SUPERVISOR_ID)))
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Shared Function GetAllPreviouslyTerminatedCurrentlyActive() As Collection(Of TermedEmploymentData)
        Dim terminatedEmployeeDataCollection As Collection(Of TermedEmploymentData) = Nothing
        Dim params As Collection = Nothing
        Try
            params = New Collection
            terminatedEmployeeDataCollection = New Collection(Of TermedEmploymentData)

            'Datatable has been used instead of SqlDataReader
            Using terminatedEmployementData As DataTable = RemoveEmployeesWhoHasDataErrorsFromDatatable(DirectCast(DataAccess.ExecuteStoredProcedure(
                               "Payroll.dbo.usp_PayGetAllPreviouslyTerminatedCurrentlyActive",
                               DataAccess.StoredProcedureReturnType.DataTable,
                               params), DataTable), "epyEMPLID")

                If terminatedEmployementData.Rows.Count > 0 Then
                    terminatedEmployeeDataCollection = New Collection(Of TermedEmploymentData)((From rec In terminatedEmployementData.AsEnumerable()
                                                                                                Select New TermedEmploymentData With {
                                                                                                    .EMPLOYEE_NUMBER = rec.Field(Of String)("EecEmpNo").Trim,
                                                                                                    .EMPLID = rec.Field(Of String)("epyEMPLID").Trim,
                                                                                                    .EMPLOYEE_NAME = rec.Field(Of String)("Employee_Name").Trim,
                                                                                                    .PAYGROUP = rec.Field(Of String)("epyPAYGROUP").Trim
                                                                                                 }).ToList())
                End If
                Return terminatedEmployeeDataCollection
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Truncates the tblJmsContingentEmpChanges table then insertes the current contingent data
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub TruncateInsertContingentEmpChanges(ByVal jmsContingentData As DataTable)
        Dim params As Collection = Nothing

        Try
            params = New Collection
            Dim sqlParameter As New SqlParameter
            sqlParameter = New SqlParameter("@tblJMSType", jmsContingentData)
            sqlParameter.SqlDbType = SqlDbType.Structured
            params.Add(sqlParameter)
            DataAccess.ExecuteStoredProcedure(
                                "Payroll.dbo.usp_PayTruncateInsertContingentEmpChanges",
                                DataAccess.StoredProcedureReturnType.RowsAffected,
                                params)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Function GetAcceptableDedCodes() As Collection(Of DedCode)
        Try
            Dim raasParam As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))

            Dim reportRaasName As String = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "RAAS_GETDEDCODES").Select(Function(cv) cv.Value1).FirstOrDefault()
            Dim reportXml As XmlDocument = RaaS_Service.GetReportResults(reportRaasName, raasParam)

            ' Parse the report into a collection of DedCodes
            Dim DedCdData As Collection(Of DedCode) = ParseXmlToDedCodesCollection(reportXml)

            ' Filter the decodes need for the compares using the control table category combos (ADPC_Feed_Ded_Cat_Rv) and single deductions (ADPC_Feed_Ded_Review).
            Return FilteredDedCodes(DedCdData)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Shared Function FilteredDedCodes(DedCdData As Collection(Of DedCode)) As Collection(Of DedCode)
        Try
            Dim DedCategoriesRv As Collection(Of PayControlValue) = GetPayControlValuesByKey(payControlValuesCollection, "ADPC_Feed_Ded_Cat_Rv")
            Dim DedCodesRv As Collection(Of PayControlValue) = GetPayControlValuesByKey(payControlValuesCollection, "ADPC_Feed_Ded_Review")

            Dim AcceptableDedCodes As Collection(Of DedCode) = New Collection(Of DedCode)

            ' Parse out the deductions by category
            For Each cat As PayControlValue In DedCategoriesRv.Where(Function(w) w.Value2.ToUpper.Equals("INCLUDE"))
                'split each cat into an string array
                Dim category() As String = cat.Value1.Split(New Char() {" "c})
                ' use the array to set the category types to get the deduction information
                Dim DedType As String = category(0)
                Dim ReportCat As String = category(1)
                Dim TaxCat As String = category(2)

                Dim DedCodesInfo = From s In DedCdData
                                   Where s.DedDedType.Equals(DedType) And s.DedReportCategory.Equals(ReportCat) And s.DedTaxCategory.Equals(TaxCat)
                                   Select s

                'DedCdData.Where(Function(w) w.DedDedType.Equals(DedType) And w.DedReportCategory.Equals(ReportCat) And w.DedTaxCategory.Equals(TaxCat)).Select(Function(s) s)

                For Each dedcd In DedCodesInfo
                    AcceptableDedCodes.Add(dedcd)
                Next
            Next

            ' Parse out the deductions by code
            For Each ded As PayControlValue In DedCodesRv.Where(Function(w) w.Value2.ToUpper.Equals("INCLUDE"))
                'check to see if a single dedcode is in the AcceptableDedCodes collection, if so skip adding
                If AcceptableDedCodes.Where(Function(w) w.DedDedCode.Equals(ded.Value1)).Select(Function(s) s).SingleOrDefault IsNot Nothing Then
                    Continue For
                End If

                ' Add the deduction info to the AcceptableDedCodes
                AcceptableDedCodes.Add(DedCdData.Where(Function(w) w.DedDedCode.Equals(ded.Value1)).Select(Function(s) s).SingleOrDefault)
            Next

            Return AcceptableDedCodes

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Shared Function ParseXmlToDedCodesCollection(xml As XmlDocument) As Collection(Of DedCode)
        Try
            Dim lstMetaDataItems As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'metadata']/*[local-name() = 'item']")
            Dim lstDataRows As XmlNodeList = xml.DocumentElement.SelectNodes("//*[local-name() = 'data']/*[local-name() = 'row']")

            Dim DedCdData As Collection(Of DedCode) = New Collection(Of DedCode)()

            For Each node As XmlNode In lstDataRows
                Dim tmpDedCdData As DedCode = New DedCode()
                Dim idx As Integer = 0

                For idx = 0 To lstMetaDataItems.Count - 1
                    Dim nodeName As String = lstMetaDataItems.Item(idx).SelectSingleNode("./@name").Value
                    Dim nodeValue As String = node.SelectNodes("./*[local-name() = 'value']").Item(idx).InnerText.Trim()

                    'NOTE: The nodeName has to match the property name for this to work.
                    tmpDedCdData.ByName(nodeName) = nodeValue

                Next

                DedCdData.Add(tmpDedCdData)
            Next

            Return DedCdData

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' This method eliminates the employee data for any individuals who have data validation errors.
    ''' </summary>
    ''' <param name="employeesData"></param>
    ''' <param name="fieldName"></param>
    ''' <returns>This method returns the employee data as a datatable for individuals who do not have any validation errors.</returns>
    Public Shared Function RemoveEmployeesWhoHasDataErrorsFromDatatable(ByVal employeesData As DataTable, Optional ByVal fieldName As String = "") As DataTable
        Try
            fieldName = If(fieldName = "", "EMPLID", fieldName)

            If employeesData IsNot Nothing AndAlso employeesData.Rows.Count > 0 AndAlso dataValidationErrors IsNot Nothing AndAlso dataValidationErrors.Any() Then
                'Checking if the errored emplids are present in the source datatable
                If employeesData.AsEnumerable().Where(Function(emp) dataValidationErrors.Select(Function(err) err.EmpEEID).ToList().Contains(emp.Field(Of String)(fieldName))).Any() Then
                    'Checking if there is any data present in source datatable other than errored emplids
                    If employeesData.AsEnumerable().Where(Function(emp) Not dataValidationErrors.Select(Function(err) err.EmpEEID).ToList().Contains(emp.Field(Of String)(fieldName))).Any() Then
                        'Copying the non errored emplid data from source datatable
                        employeesData = employeesData.AsEnumerable().Where(Function(emp) Not dataValidationErrors.Select(Function(err) err.EmpEEID).ToList().Contains(emp.Field(Of String)(fieldName))).CopyToDataTable()
                    Else
                        'Making the source datatable nothing if all the data are errored emplids 
                        employeesData = New DataTable
                    End If
                End If
            End If

            Return employeesData

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' This method eliminates the employee data for any individuals who have data validation errors.
    ''' </summary>
    ''' <param name="employeesData"></param>
    ''' <param name="fieldName"></param>
    ''' <returns>This method returns the employee data as a dataset for individuals who do not have any validation errors.</returns>
    Public Shared Function RemoveEmployeesWhoHasDataErrorsFromDataset(ByVal employeesData As DataSet, Optional ByVal fieldName As String = "") As DataSet
        Try
            Dim resultDataSet As New DataSet

            For Each employeesDataTable As DataTable In employeesData.Tables
                'To avoid the error of adding the same table to a dataset when there are no changes from the calling method, we are using the "CopyToDataTable" method to pass the datatable as a copy.
                resultDataSet.Tables.Add(RemoveEmployeesWhoHasDataErrorsFromDatatable(employeesDataTable.AsEnumerable().CopyToDataTable(), fieldName))
            Next

            Return resultDataSet

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' This method eliminates the employee data for any individuals who have data validation errors.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="employeesData"></param>
    ''' <param name="fieldName"></param>
    ''' <returns>This method returns the employee data as a generic collection for individuals who do not have any validation errors.</returns>
    Public Shared Function RemoveEmployeesWhoHasDataErrorsFromCollection(Of T)(ByRef employeesData As Collection(Of T), Optional ByVal fieldName As String = "") As Collection(Of T)
        Try
            Dim itemType As Type
            Dim propValue As String

            fieldName = If(fieldName = "", "EMPLID", fieldName)

            If dataValidationErrors IsNot Nothing AndAlso dataValidationErrors.Any() Then
                Dim i As Integer = 0
                Do While i <= employeesData.Count - 1
                    Dim empData = employeesData(i)
                    itemType = empData.GetType()
                    propValue = itemType.GetProperty(fieldName).GetValue(empData).ToString()

                    If propValue IsNot Nothing AndAlso dataValidationErrors.Select(Function(err) err.EmpEEID).ToList().Contains(propValue) Then
                        employeesData.Remove(empData)
                    Else
                        i += 1
                    End If
                Loop
            End If

            Return employeesData

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' This method converts the data validation error collection into datatable
    ''' </summary>
    Public Shared Sub ConvertErrorCollectionToDatatable()

        dataValidationErrorsAsDataTable = New DataTable()
        dataValidationErrorsAsDataTable.Columns.Add("EecEEID")
        dataValidationErrorsAsDataTable.Columns.Add("EecCOID")
        dataValidationErrorsAsDataTable.Columns.Add("EecEmpNo")

        If dataValidationErrors IsNot Nothing AndAlso dataValidationErrors.Any() Then
            For i As Integer = 0 To dataValidationErrors.Count - 1
                dataValidationErrorsAsDataTable.Rows.Add(dataValidationErrors(i).EEID, dataValidationErrors(i).CompanyCode, dataValidationErrors(i).EmployeeNumber)
            Next i
        End If
    End Sub
    ''' <summary>
    ''' Convert the generic collection to datatable
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="collection"></param>
    ''' <returns></returns>
    Public Shared Function ConvertCollectionToDataTable(Of T)(collection As IEnumerable(Of T)) As DataTable
        ' Create new DataTable
        Dim dataTable As New DataTable()

        ' Get the type information of the generic class
        Dim type As Type = GetType(T)
        Dim properties = type.GetProperties().Where(Function(p) p.Name <> "ByName")


        ' Add columns based on properties
        For Each prop In properties
            ' Add column with property name and type
            dataTable.Columns.Add(prop.Name, If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType))

        Next

        ' Add rows
        If collection IsNot Nothing Then
            For Each item In collection
                Dim row As DataRow = dataTable.NewRow()

                ' Set values for each property
                For Each prop In properties
                    row(prop.Name) = If(prop.GetValue(item), DBNull.Value)
                Next

                dataTable.Rows.Add(row)
            Next
        End If


        Return dataTable
    End Function


End Class

