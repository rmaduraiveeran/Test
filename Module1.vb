Option Strict On
Option Explicit On

Imports System.Collections.ObjectModel
Imports System.Configuration.ConfigurationManager
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports Ionic.Zip


Public Module Module1

    Private personalDataBeforeCollection As Collection(Of PersonalData) = Nothing
    Private personalDataCurrentCollection As Collection(Of PersonalData) = Nothing
    Private employmentBeforeCollection As Collection(Of Employment) = Nothing
    Private employmentCurrentCollection As Collection(Of Employment) = Nothing
    Private terminatedEmployeeCollection As Collection(Of TermedEmploymentData) = Nothing
    Private previouslyTerminatedCurrentlyActiveCollection As Collection(Of TermedEmploymentData) = Nothing
    Private terminatedEmployeeHistoryCollection As Collection(Of TermedEmploymentData) = Nothing
    Private jobBeforeCollection As Collection(Of Job) = Nothing
    Private jobCurrentCollection As Collection(Of Job) = Nothing
    Public generalDeductionsBeforeCollection As Collection(Of GeneralDeduction) = Nothing
    Private generalDeductionsDeletedCollection As Collection(Of GeneralDeduction) = Nothing
    Private directDepositBeforeCollection As Collection(Of DirectDeposit) = Nothing
    Private directDepositCurrentCollection As Collection(Of DirectDeposit) = Nothing
    Private taxDataBeforeCollection As Collection(Of TaxData) = Nothing
    Private taxDataCurrentCollection As Collection(Of TaxData) = Nothing
    Private erroredTaxDataCurrentCollection As List(Of TaxData) = Nothing
    Private goalAmountMissingOnDeductions As List(Of GeneralDeduction) = Nothing
    Private w4DataBeforeCollection As Collection(Of W4Data) = Nothing
    Private w4DataCurrentCollection As Collection(Of W4Data) = Nothing
    Private doNotSendToADPCollection As Collection(Of DoNotSendToADP) = Nothing
    Private runID As Int32 = Nothing
    Private ultiproDataValidationError As String = String.Empty
    Private collectionsNotPopulatedError As String = String.Empty

    Public payControlValuesCollection As Collection(Of PayControlValue) = Nothing

    Public AcceptableDedCodes As Collection(Of DedCode) = Nothing
    Private acceptableDeductionsFromToCompanyTransfers As Collection(Of GeneralDeduction) = Nothing

    Public generalDeductionsCurrentCollection As Collection(Of GeneralDeduction) = Nothing
    Public generalDeductionsCurrentAllCollection As Collection(Of GeneralDeduction) = Nothing
    Public generalDeductionFinalCollections As Collection(Of GeneralDeduction) = Nothing
    Public dataValidationErrors As Collection(Of DataValidationError) = Nothing
    Public dataValidationErrorsAsDataTable As DataTable = Nothing
    Public changedVIPDeductionData As IEnumerable(Of VIPData) = Nothing

    Sub Main()
        InitiateFileCreationProcess(Nothing, Nothing)
    End Sub


    ' Main subroutine
    Private Sub InitiateFileCreationProcess(ByVal beginPayPeriodDate As Date, ByVal endPayPeriodDate As Date)
        Dim fileCreateLocation As String = String.Empty
        Dim fileOutboundLocation As String = String.Empty
        Dim fileArchiveLocation As String = String.Empty

        Try
            ' Insert the file creation run log record and return the ID
            runID = DataManager.InsertRunLogRecord()

            'Load the control values
            payControlValuesCollection = DataManager.LoadPayControlValues()

            'Insert the contingent data into tblJmsContingentEmpChanges
            DataManager.InsertJMSContingentData()
            ' Set the pay period dates to use
            DataManager.SetPayPeriodDates(beginPayPeriodDate, endPayPeriodDate)

            'Load Tax and Personal data
            LoadTaxAndPersonalData()

            'Load all deductions related data
            LoadDeductions()

            'Build deduction data
            BuildDeductionToSendADP()

            ' Validate Ultipro Data before continueing
            ValidUltiproData()

            'Remove Errored Data From Tax And Personal data
            RemoveErroredDataFromTaxAndPersonalData()

            'Remove Errored Data From all deductions related data
            RemoveErroredDataFromDeductionData()

            ' *** LOAD CREATE LOCATION AND VALIDATE ***
            ' -------------------------------------------------------------------------------------
            ' Set the file create location
            fileCreateLocation = DataManager.GetFileLocation("CREATE")

            ' Validate the file create location is valid
            If Not Directory.Exists(fileCreateLocation) Then
                Throw New InvalidOperationException(String.Concat("File location '", IIf(fileCreateLocation <> String.Empty, fileCreateLocation, "Does not exist").ToString, "' is not valid.  ",
                                    "Please enter a valid file create location before running again."))
            End If

            ' Add last backslash to create location if not present
            If fileCreateLocation(fileCreateLocation.Length - 1) <> "\" Then
                fileCreateLocation = String.Concat(fileCreateLocation, "\")
            End If

            ' Validate there are currently no files in the file location specified
            ' This is to prevent the newer file of any given file being loaded before the old file does or one getting overwritten.
            If New DirectoryInfo(fileCreateLocation).GetFiles("*.csv").Count > 0 Then
                Throw New InvalidOperationException(String.Concat("There currently are .csv files in the location: ", fileCreateLocation, ".  Please remove the existing .csv files before ",
                                    "running the ADPC employee feed again."))
            End If


            ' *** LOAD OUTBOUND LOCATION AND VALIDATE ***
            ' -------------------------------------------------------------------------------------
            ' Set the file outbound location
            fileOutboundLocation = DataManager.GetFileLocation("OUTBOUND")

            ' Validate the file outbound location is valid
            If Not Directory.Exists(fileOutboundLocation) Then
                Throw New InvalidOperationException(String.Concat("File location '", IIf(fileOutboundLocation <> String.Empty, fileOutboundLocation, "Does not exist").ToString, "' is not valid.  ",
                                    "Please enter a valid file outbound location before running again."))
            End If

            ' Add last backslash to outbound location if not present
            If fileOutboundLocation(fileOutboundLocation.Length - 1) <> "\" Then
                fileOutboundLocation = String.Concat(fileOutboundLocation, "\")
            End If


            ' *** LOAD ARCHIVE LOCATION AND VALIDATE ***
            ' -------------------------------------------------------------------------------------
            ' Set the file archive location
            fileArchiveLocation = DataManager.GetFileLocation("ARCHIVE")

            ' Validate the file archive location is valid
            If Not Directory.Exists(fileArchiveLocation) Then
                Throw New InvalidOperationException(String.Concat("File location '", IIf(fileArchiveLocation <> String.Empty, fileArchiveLocation, "Does not exist").ToString, "' is not valid.  ",
                                    "Please enter a valid file archive location before running again."))
            End If

            ' Add last backslash to outbound location if not present
            If fileArchiveLocation(fileArchiveLocation.Length - 1) <> "\" Then
                fileArchiveLocation = String.Concat(fileArchiveLocation, "\")
            End If


            ' *** LOAD OBJECT COLLECTIONS ***
            ' -------------------------------------------------------------------------------------
            ' Load All the Object Collections
            If LoadObjectCollections() Then
                ' *** CREATE CSV FILES ***
                ' -------------------------------------------------------------------------------------
                ' Check to see if CSV files are to be created based on flag in app config file
                Select Case AppSettings.Get("CreateCSVfiles")
                    Case "TRUE"

                        ' Create the CSV Files
                        CreateFiles(fileCreateLocation, fileOutboundLocation, fileArchiveLocation)
                    Case "FALSE"

                        ' Update run log record with blanks for counts as nothing was sent and indicate a PASS status for all the validations
                        DataManager.UpdateRunLogRecord(runID, 0, 0, 0, 0, 0, 0, 0, "PASS", Nothing)
                    Case Else

                        ' Throw error as value should always be set to True or False
                        Throw New InvalidOperationException("Please check the application config file and make sure the 'CreateCSVfiles' setting has a value of 'True' or 'False'")
                End Select
            Else

                ' ******** ONE OR MORE COLLECTIONS WERE NOT POPULATED *********
                Throw New InvalidOperationException(collectionsNotPopulatedError)
            End If

            '----Added InvalidTaxDataException handling Begin------
        Catch ex As InvalidTaxDataException
            ' Log the error in the database on the run log record
            If Not IsNothing(runID) Then
                DataManager.UpdateRunLogRecord(runID, 0, 0, 0, 0, 0, 0, 0, "FAIL", ex.ToString)
            End If

            ' Email the error recipients
            Try
                DataManager.SendErrorNotification(ex.ToString, "ERROR - ADPC Employee Feed - URGENT ATTENTION NEEDED")
            Catch ex2 As Exception
                EventLog.WriteEntry("ADPC File Creation Error Notification", ex2.ToString(), EventLogEntryType.Error)
            End Try
            ' Log the error in the event log
            EventLog.WriteEntry("ADPC File Creation", "The employees list sent through email are either rehired or changed to salaried position with an old W4 form but their hire date is 2020 or later. Please review and correct W4 information to reflect 2020 or later W4 form to align with hire date.", EventLogEntryType.Error)
            '----Added InvalidTaxDataException handling End-------
            '----Added SQL Exception handling Begin------
        Catch ex As SqlException
            ' Log the error in the database on the run log record
            If Not IsNothing(runID) Then
                DataManager.UpdateRunLogRecord(runID, 0, 0, 0, 0, 0, 0, 0, "FAIL", ex.ToString)
            End If
            ' Email the error recipients
            Try
                If DirectCast(ex, System.Data.SqlClient.SqlException).Number.Equals(2627) Then
                    SendDuplicateNotifications(ex.StackTrace, "ERROR - ADPC Employee Feed - URGENT ATTENTION NEEDED", ex.Message)
                Else
                    DataManager.SendErrorNotification(String.Concat("<b>IT will review and determine to resubmit the feed or if other action is needed.</b> <br/><br/>", Environment.NewLine,
                                                                    "Error in ADP File Creation Process:<br/>", Environment.NewLine,
                                             Environment.NewLine, ex.ToString))
                End If
            Catch ex2 As Exception
                EventLog.WriteEntry("ADPC File Creation Error Notification", ex2.ToString(), EventLogEntryType.Error)
            End Try
            ' Log the error in the event log
            EventLog.WriteEntry("ADPC File Creation", ex.ToString(), EventLogEntryType.Error)
            '----Added SQL Exception handling End-------
        Catch ex As Exception

            ' ********* ERROR!!!!!! Logs error to database, notifies users via email, logs error to event log on pc which it is ran *********

            ' For testing, write error out to screen
            'Console.WriteLine(ex.ToString)

            ' Log the error in the database on the run log record
            If Not IsNothing(runID) Then
                DataManager.UpdateRunLogRecord(runID, 0, 0, 0, 0, 0, 0, 0, "FAIL", ex.ToString)
            End If

            ' Email the error recipients
            Try

                DataManager.SendErrorNotification(String.Concat("<b>IT will review and determine to resubmit the feed or if other action is needed.</b> <br/><br/>", Environment.NewLine,
                                                                    "Error in ADP File Creation Process:<br/>", Environment.NewLine,
                                             Environment.NewLine, ex.ToString))
            Catch ex2 As Exception
                EventLog.WriteEntry("ADPC File Creation Error Notification", ex2.ToString(), EventLogEntryType.Error)
            End Try
            ' Log the error in the event log
            EventLog.WriteEntry("ADPC File Creation", ex.ToString().Substring(0, 5000), EventLogEntryType.Error)

        End Try
    End Sub


    ' Sends duplicate notifications to defined recipients 
    Private Sub SendDuplicateNotifications(ByVal errorMessage As String, Optional ByVal ovrSubject As String = "", Optional ByVal strheader As String = "")

        Dim dtDuplicateEmployees As DataTable = Nothing
        Dim changeMessage As String = String.Empty
        Dim empCounter As Int32 = 0
        Try
            Dim headertext As String() = strheader.Split(New Char() {"."c})

            ' ***** Get the recognized employees that have had changes to specific deductions *****
            dtDuplicateEmployees = DataManager.GetDuplicateEmployees()

            changeMessage = String.Concat("Error in ADP File Creation Process: ", headertext(0), Environment.NewLine,
                                             Environment.NewLine)

            changeMessage = String.Concat(changeMessage, "<br/><br/> Ultipro Data Is Not Valid. <br/><br/>")
            ' Additional message body text requested by Business - CJG 04/05/2022
            changeMessage = String.Concat(changeMessage,
                                                           "<span style=""font-weight: bold; font-style:italic;"">HR – Please correct below errors as quickly as possible and respond back to all on this email once complete so that the feed to ADP can be reattempted.  The feed is very time sensitive, so your prompt attention is appreciated.  Thank you.</span><br/><br/>")


            If dtDuplicateEmployees.Rows.Count > 0 Then
                changeMessage = String.Concat(changeMessage, "Please review below employees has multiple job change records in LodeComp.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                ' Add each employee to the list
                'For Each row As DataRow In employeesWithChangedDeductions.Rows
                '    changeMessage = String.Concat(changeMessage, "EE#: ", row.Item("Employee Number").ToString, ", Company: ", row.Item("Company").ToString, _
                '              ", Deduction: ", row.Item("Deduction").ToString, Environment.NewLine, Environment.NewLine)
                'Next

                '' Send the tax change notification
                'DataManager.SendStandardNotification("DEDCHANGE", changeMessage)

                empCounter = 1
                'changeMessage = String.Empty
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>EE#</th><th class='tdborder'>Name</th></tr>")
                For Each row As DataRow In dtDuplicateEmployees.Rows
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("EmpNumber").ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Name").ToString & "</td>")
                    empCounter += 1
                Next


                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"))

                changeMessage = String.Concat(changeMessage, Environment.NewLine, errorMessage)
                ' Send the tax change notification
                DataManager.SendErrorNotification(changeMessage, ovrSubject)
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Load the collections of objects
    Private Function LoadObjectCollections() As Boolean
        Dim collectionsPopulated As Boolean = True

        Try
            ' Acceptable deductions for notifications
            AcceptableDedCodes = DataManager.GetAcceptableDedCodes()

            ' Personal Data Before
            personalDataBeforeCollection = DataManager.LoadPersonalData(DataManager.TypeOfData.Before)

            ' Employment Before
            employmentBeforeCollection = DataManager.LoadEmploymentData(DataManager.TypeOfData.Before)

            ' Employment Current
            employmentCurrentCollection = DataManager.LoadEmploymentData(DataManager.TypeOfData.Current)

            'Terminated Employees List
            terminatedEmployeeCollection = DataManager.GetAllTerminatedEmployeeData(DataManager.TypeOfData.Before)

            'Terminated Employees History 
            terminatedEmployeeHistoryCollection = DataManager.GetAllTerminatedEmployeeData(DataManager.TypeOfData.History)

            'Previously Terminated and currently active Employee list
            previouslyTerminatedCurrentlyActiveCollection = DataManager.GetAllPreviouslyTerminatedCurrentlyActive()

            ' Job Before
            jobBeforeCollection = DataManager.LoadJobData(DataManager.TypeOfData.Before)

            ' Job Current
            jobCurrentCollection = DataManager.LoadJobData(DataManager.TypeOfData.Current)

            'Acceptable deductions for the Company Transfer message.
            acceptableDeductionsFromToCompanyTransfers = DataManager.LoadAcceptableDeductionsFromTo()

            ' Direct Deposit Before
            directDepositBeforeCollection = DataManager.LoadDirectDepositData(DataManager.TypeOfData.Before)

            ' Direct Deposit Current
            directDepositCurrentCollection = DataManager.LoadDirectDepositData(DataManager.TypeOfData.Current)

            ' Tax Data Before
            taxDataBeforeCollection = DataManager.LoadTaxData(DataManager.TypeOfData.Before)

            ' W4 Data Before
            w4DataBeforeCollection = DataManager.LoadW4Data(DataManager.TypeOfData.Before)

            ' W4 Data Current
            w4DataCurrentCollection = DataManager.LoadW4Data(DataManager.TypeOfData.Current)

            ' Validate items are present within all required collections
            If (IsNothing(personalDataBeforeCollection) OrElse personalDataBeforeCollection.Count = 0) Or
               (IsNothing(personalDataCurrentCollection) OrElse personalDataCurrentCollection.Count = 0) Or
               (IsNothing(employmentBeforeCollection) OrElse employmentBeforeCollection.Count = 0) Or
               (IsNothing(employmentCurrentCollection) OrElse employmentCurrentCollection.Count = 0) Or
               (IsNothing(jobBeforeCollection) OrElse jobBeforeCollection.Count = 0) Or
               (IsNothing(jobCurrentCollection) OrElse jobCurrentCollection.Count = 0) Or
               (IsNothing(generalDeductionsBeforeCollection) OrElse generalDeductionsBeforeCollection.Count = 0) Or
               (IsNothing(generalDeductionsCurrentCollection) OrElse generalDeductionsCurrentCollection.Count = 0) Or
               (IsNothing(generalDeductionsCurrentAllCollection) OrElse generalDeductionsCurrentAllCollection.Count = 0) Or
               (IsNothing(directDepositBeforeCollection) OrElse directDepositBeforeCollection.Count = 0) Or
               (IsNothing(directDepositCurrentCollection) OrElse directDepositCurrentCollection.Count = 0) Or
               (IsNothing(taxDataBeforeCollection) OrElse taxDataBeforeCollection.Count = 0) Or
               (IsNothing(taxDataCurrentCollection) OrElse taxDataCurrentCollection.Count = 0) Or
               (IsNothing(w4DataBeforeCollection) OrElse w4DataBeforeCollection.Count = 0) Or
               (IsNothing(w4DataCurrentCollection) OrElse w4DataCurrentCollection.Count = 0) Then

                collectionsPopulated = False

                collectionsNotPopulatedError = String.Format("The following collections were not populated:{0}", Environment.NewLine)

                If (IsNothing(personalDataBeforeCollection) OrElse personalDataBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Personal Data Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(personalDataCurrentCollection) OrElse personalDataCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Personal Data Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(employmentBeforeCollection) OrElse employmentBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Employment Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(employmentCurrentCollection) OrElse employmentCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Employment Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(jobBeforeCollection) OrElse jobBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Job Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(jobCurrentCollection) OrElse jobCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Job Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(generalDeductionsBeforeCollection) OrElse generalDeductionsBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Deductions Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(generalDeductionsCurrentCollection) OrElse generalDeductionsCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Deductions Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If

                If (IsNothing(generalDeductionsCurrentAllCollection) OrElse generalDeductionsCurrentAllCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Deductions Current All{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If

                If (IsNothing(directDepositBeforeCollection) OrElse directDepositBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Direct Deposit Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(directDepositCurrentCollection) OrElse directDepositCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Direct Deposit Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(taxDataBeforeCollection) OrElse taxDataBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Tax Data Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(taxDataCurrentCollection) OrElse taxDataCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          Tax Data Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(w4DataBeforeCollection) OrElse w4DataBeforeCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          W4 Data Before{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If
                If (IsNothing(w4DataCurrentCollection) OrElse w4DataCurrentCollection.Count = 0) Then
                    collectionsNotPopulatedError = String.Format("{0}          W4 Data Current{1}", collectionsNotPopulatedError, Environment.NewLine)
                End If

                collectionsNotPopulatedError = String.Format("{0}{1}", collectionsNotPopulatedError, Environment.NewLine)
            End If

            Return collectionsPopulated
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    ' Creates files based on employee information that has changed from what was last sent
    Private Sub CreateFiles(ByVal createFileLocation As String, ByVal outboundFileLocation As String, ByVal archiveFileLocation As String)
        Dim sqlConn As SqlConnection = Nothing
        Dim sqlTrans As SqlTransaction = Nothing
        Dim dataRow As StringBuilder = Nothing
        Dim addTaxDataRecordsSent As Integer = 0
        Dim doNotSendPersDataRecs As Integer = 0
        Dim doNotSendEmployRecs As Integer = 0
        Dim doNotSendJobRecs As Integer = 0
        Dim doNotSendGenDedRecs As Integer = 0
        Dim doNotSendDirDepRecs As Integer = 0
        Dim doNotSendW4DataRecs As Integer = 0
        Dim doNotSendTaxDataRecs As Integer = 0
        Dim vipDetails As Collection(Of VIPData) = Nothing
        Dim deductionsValue As String = String.Empty
        Dim directdepValue As String = String.Empty
        Dim employmentValue As String = String.Empty
        Dim jobValue As String = String.Empty
        Dim personaldataValue As String = String.Empty
        Dim taxdataValue As String = String.Empty
        Dim w4dataValue As String = String.Empty
        Try

            ' Create an open connection to the sql server
            DataAccess.CreateOpenConnection(sqlConn)

            ' Begin the sql transaction
            sqlTrans = sqlConn.BeginTransaction()

            vipDetails = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(DataManager.GetVIPDetails("EMPLOYEE_CHNG"))

            ' ********************************* PERSONAL DATA FILE *********************************
            ' **************************************************************************************

            deductionsValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "deductions.csv").Select(Function(cv) cv.Value2).FirstOrDefault()
            directdepValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "direct dep.csv").Select(Function(cv) cv.Value2).FirstOrDefault()
            employmentValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "employment.csv").Select(Function(cv) cv.Value2).FirstOrDefault()
            jobValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "job.csv").Select(Function(cv) cv.Value2).FirstOrDefault()
            personaldataValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "personal data.csv").Select(Function(cv) cv.Value2).FirstOrDefault()
            taxdataValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "tax data.csv").Select(Function(cv) cv.Value2).FirstOrDefault()
            w4dataValue = PayControlValueHandler.GetPayControlValuesByKeyandvalue(payControlValuesCollection, "ADPC_Feed_Prefix", "w4 data.csv").Select(Function(cv) cv.Value2).FirstOrDefault()

            Dim newPersonalData1 = From pdc In personalDataCurrentCollection
                                   Group Join pdb In personalDataBeforeCollection
                                    On pdc.EMPLID Equals pdb.EMPLID And pdc.PAYGROUP Equals pdb.PAYGROUP Into g = Group
                                   From pdb In g.DefaultIfEmpty()
                                   Where IsNothing(pdb)
                                   Select pdc

            ' Get records that have changed
            Dim changedPersonalData = From pdc In personalDataCurrentCollection
                                      Join pdb In personalDataBeforeCollection
                                        On pdc.EMPLID Equals pdb.EMPLID And pdc.PAYGROUP Equals pdb.PAYGROUP
                                      Where pdc.PAY_FREQUENCY <> pdb.PAY_FREQUENCY Or
                                        pdc.FIRST_NAME <> pdb.FIRST_NAME Or
                                        pdc.MIDDLE_NAME <> pdb.MIDDLE_NAME Or
                                        pdc.LAST_NAME <> pdb.LAST_NAME Or
                                        pdc.STREET1 <> pdb.STREET1 Or
                                        pdc.STREET2 <> pdb.STREET2 Or
                                        pdc.CITY <> pdb.CITY Or
                                        pdc.STATE <> pdb.STATE Or
                                        pdc.ZIP <> pdb.ZIP Or
                                        pdc.HOME_PHONE <> pdb.HOME_PHONE Or
                                        pdc.SSN <> pdb.SSN Or
                                        pdc.ORIG_HIRE_DT <> pdb.ORIG_HIRE_DT Or
                                        pdc.SEX <> pdb.SEX Or
                                        pdc.BIRTHDATE <> pdb.BIRTHDATE
                                      Select pdc

            Dim newPersDataMissing = From pdc In personalDataCurrentCollection
                                     Where pdc.FIRST_NAME = "" Or
                                     pdc.LAST_NAME = "" Or
                                     pdc.SEX = "" Or
                                    (pdc.STREET1 = "" And pdc.STREET2 = "") Or
                                     pdc.SSN = "" Or
                                     pdc.STATE = "" Or
                                     pdc.CITY = "" Or
                                     pdc.PAYGROUP = "" Or
                                     pdc.PAY_FREQUENCY = "" Or
                                     pdc.ZIP = "" Or
                                     (pdc.ORIG_HIRE_DT.ToShortDateString = "" Or pdc.ORIG_HIRE_DT.ToShortDateString = "1/1/1900") Or
                                     (pdc.BIRTHDATE.ToShortDateString = "" Or pdc.BIRTHDATE.ToShortDateString = "1/1/1900")
                                     Select pdc

            ' Get new records
            Dim newPersonalData = From pdc In newPersonalData1
                                  Group Join pdb In newPersDataMissing
                                    On pdc.EMPLID Equals pdb.EMPLID And pdc.PAYGROUP Equals pdb.PAYGROUP Into g = Group
                                  From pdb In g.DefaultIfEmpty()
                                  Where IsNothing(pdb)
                                  Select pdc


            ' Combine the result sets from both the new and changed data
            Dim personalDataToUpdate = Enumerable.Union(newPersonalData, changedPersonalData)

            ' Get records that have changed for VIPs
            Dim changedVIPPersonalData = (From pdc In personalDataToUpdate
                                          Join vip In vipDetails
                                         On pdc.EMPLID Equals vip.EMPLID
                                          Select vip).Distinct()

            SendMissingEmployeeDataNotification(newPersDataMissing)

            ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
            If personalDataToUpdate.Count > 0 Then

                Using sw As StreamWriter = New StreamWriter(createFileLocation & "Personal Data.csv")

                    ' Add the header row
                    sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,FIRST_NAME,MIDDLE_NAME,LAST_NAME,STREET1,STREET2,CITY,STATE,ZIP,HOME_PHONE,SSN," &
                                 "ORIG_HIRE_DT,SEX,BIRTHDATE")

                    ' Add the data rows
                    For Each pd In personalDataToUpdate

                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                Where dns.EMPLID = pd.EMPLID And
                                                  dns.PAYGROUP = pd.PAYGROUP
                                                Select dns

                        ' Make sure the emp is not in the do not send to ADP list before writing to file
                        If doNotSendEmpToADP.Count = 0 Then

                            ' Write object to csv file
                            dataRow = New StringBuilder
                            dataRow.Append("""" & personaldataValue & """")
                            dataRow.Append(",""" & pd.EMPLID & """")
                            dataRow.Append(",""" & pd.PAYGROUP & """")
                            dataRow.Append(",""" & pd.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & pd.FIRST_NAME & """")
                            dataRow.Append(",""" & pd.MIDDLE_NAME & """")
                            dataRow.Append(",""" & pd.LAST_NAME & """")
                            dataRow.Append(",""" & pd.STREET1 & """")
                            dataRow.Append(",""" & pd.STREET2 & """")
                            dataRow.Append(",""" & pd.CITY & """")
                            dataRow.Append(",""" & pd.STATE & """")
                            dataRow.Append(",""" & pd.ZIP & """")
                            dataRow.Append(",""" & pd.HOME_PHONE & """")
                            dataRow.Append(",""" & pd.SSN & """")
                            dataRow.Append(",""" & pd.ORIG_HIRE_DT.ToShortDateString & """")
                            dataRow.Append(",""" & pd.SEX & """")
                            dataRow.Append(",""" & pd.BIRTHDATE.ToShortDateString & """")
                            sw.WriteLine(dataRow)
                        Else
                            doNotSendPersDataRecs += 1
                        End If

                        ' Insert object to database
                        DataManager.InsertPersonalData(sqlConn, sqlTrans, pd.EMPLID, pd.PAYGROUP, pd.PAY_FREQUENCY, pd.FIRST_NAME, pd.MIDDLE_NAME,
                                                       pd.LAST_NAME, pd.STREET1, pd.STREET2, pd.CITY, pd.STATE, pd.ZIP, pd.HOME_PHONE,
                                                       pd.SSN, pd.ORIG_HIRE_DT, pd.SEX, pd.BIRTHDATE, runID)

                    Next

                    ' Close and dispose the streamwriter to finish writing out file
                    sw.Close()
                End Using
            End If
            ' **************************************************************************************
            ' **************************************************************************************


            ' ********************************** EMPLOYMENT FILE ***********************************
            ' **************************************************************************************
            ' Get new records
            Dim newEmploymentData1 = From emc In employmentCurrentCollection
                                     Group Join emb In employmentBeforeCollection
                                  On emc.EMPLID Equals emb.EMPLID And emc.PAYGROUP Equals emb.PAYGROUP Into g = Group
                                     From emb In g.DefaultIfEmpty()
                                     Where IsNothing(emb)
                                     Select emc

            ' Get records that have changed
            Dim changedEmploymentData = From emc In employmentCurrentCollection
                                        Join emb In employmentBeforeCollection
                                      On emc.EMPLID Equals emb.EMPLID And emc.PAYGROUP Equals emb.PAYGROUP
                                        Where emc.PAY_FREQUENCY <> emb.PAY_FREQUENCY Or
                                      emc.HIRE_DT <> emb.HIRE_DT Or
                                      emc.REHIRE_DT <> emb.REHIRE_DT Or
                                      emc.CMPNY_SENIORITY_DT <> emb.CMPNY_SENIORITY_DT Or
                                      emc.TERMINATION_DT <> emb.TERMINATION_DT Or
                                      emc.LAST_DATE_WORKED <> emb.LAST_DATE_WORKED Or
                                      emc.BUSINESS_TITLE <> emb.BUSINESS_TITLE Or
                                      emc.SUPERVISOR_ID <> emb.SUPERVISOR_ID
                                        Select emc
            ' Final resultant set of New Employement Data
            Dim newEmploymentData = From emc In newEmploymentData1
                                    Group Join emb In newPersDataMissing
                                    On emc.EMPLID Equals emb.EMPLID And emc.PAYGROUP Equals emb.PAYGROUP Into g = Group
                                    From emb In g.DefaultIfEmpty()
                                    Where IsNothing(emb)
                                    Select emc

            ' Combine the result sets from both the new and changed data
            Dim employmentDataToUpdate = Enumerable.Union(newEmploymentData, changedEmploymentData)

            'Get the modified data of Terminated Users 
            Dim termedEmploymentData = (From emc In employmentDataToUpdate
                                        Join emt In terminatedEmployeeCollection
                                            On emc.EMPLID Equals emt.EMPLID And emc.PAYGROUP Equals emt.PAYGROUP
                                        Select emt).Distinct()

            ' Final result set of Termed Employement Data after removing currently Active Employees
            Dim termedEmploymentDataOthers = From emc In termedEmploymentData
                                             Group Join emb In previouslyTerminatedCurrentlyActiveCollection
                                                On emc.EMPLID Equals emb.EMPLID And emc.PAYGROUP Equals emb.PAYGROUP Into g = Group
                                             From emb In g.DefaultIfEmpty()
                                             Where IsNothing(emb)
                                             Select emc

            ' Notify the business with the modified data of Terminated Users
            Dim termedEmploymentDataToUpdate = SendTerminatedEmployeeChangeNotification(employmentDataToUpdate, termedEmploymentDataOthers)
            If termedEmploymentDataToUpdate.Count > 0 Then
                ' Insert the Termed Employment Logs
                DataManager.InsertTermedEmploymentLogs(sqlConn, sqlTrans, termedEmploymentDataToUpdate, runID)
            End If

            ' Get records that have changed for VIPs
            Dim changedVIPEmploymentData = (From pdc In employmentDataToUpdate
                                            Join vip In vipDetails
                                            On pdc.EMPLID Equals vip.EMPLID
                                            Select vip).Distinct()

            ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
            If employmentDataToUpdate.Count > 0 Then

                Using sw As StreamWriter = New StreamWriter(createFileLocation & "Employment.csv")

                    If Not employmentDataToUpdate.Count = termedEmploymentDataToUpdate.Count Then
                        ' Add the header row
                        sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,HIRE_DT,REHIRE_DT,CMPNY_SENIORITY_DT,TERMINATION_DT,LAST_DATE_WORKED,BUSINESS_TITLE," &
                                 "SUPERVISOR_ID")
                    Else
                        sw.Close()
                        'To remove the empty csv file
                        File.Delete(createFileLocation & "Employment.csv")
                    End If

                    ' Add the data rows
                    For Each em In employmentDataToUpdate

                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                Where dns.EMPLID = em.EMPLID And
                                                  dns.PAYGROUP = em.PAYGROUP
                                                Select dns

                        ' Check if termed emp is in do not send to ADP list
                        Dim doNotSendTermedEmpToADP = From dns In termedEmploymentDataToUpdate
                                                      Where dns.EMPLID = em.EMPLID And
                                                        dns.PAYGROUP = em.PAYGROUP
                                                      Select dns

                        ' Make sure the emp is not in the do not send to ADP list before writing to file
                        If doNotSendEmpToADP.Count = 0 And doNotSendTermedEmpToADP.Count = 0 Then

                            ' Write object to csv file
                            dataRow = New StringBuilder
                            dataRow.Append("""" & employmentValue & """")
                            dataRow.Append(",""" & em.EMPLID & """")
                            dataRow.Append(",""" & em.PAYGROUP & """")
                            dataRow.Append(",""" & em.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & em.HIRE_DT.ToShortDateString & """")
                            If em.REHIRE_DT <> Nothing Then
                                dataRow.Append(",""" & em.REHIRE_DT.ToShortDateString & """")
                            Else
                                dataRow.Append(",""""")
                            End If
                            dataRow.Append(",""" & em.CMPNY_SENIORITY_DT.ToShortDateString & """")
                            If em.TERMINATION_DT <> Nothing Then
                                dataRow.Append(",""" & em.TERMINATION_DT.ToShortDateString & """")
                            Else
                                dataRow.Append(",""""")
                            End If
                            If em.LAST_DATE_WORKED <> Nothing Then
                                dataRow.Append(",""" & em.LAST_DATE_WORKED.ToShortDateString & """")
                            Else
                                dataRow.Append(",""""")
                            End If
                            dataRow.Append(",""" & em.BUSINESS_TITLE & """")
                            dataRow.Append(",""" & em.SUPERVISOR_ID & """")
                            sw.WriteLine(dataRow)
                        Else
                            doNotSendEmployRecs += 1
                        End If

                        ' Insert object to database
                        DataManager.InsertEmployment(sqlConn, sqlTrans, em.EMPLID, em.PAYGROUP, em.PAY_FREQUENCY, em.HIRE_DT, em.REHIRE_DT, em.CMPNY_SENIORITY_DT,
                                                     em.TERMINATION_DT, em.LAST_DATE_WORKED, em.BUSINESS_TITLE, em.SUPERVISOR_ID, runID)

                    Next

                    ' Close and dispose the streamwriter to finish writing out file
                    sw.Close()
                End Using
            End If
            ' **************************************************************************************
            ' **************************************************************************************


            ' ************************************* JOB FILE ***************************************
            ' **************************************************************************************
            ' Get new records
            Dim newJobData1 = From jbc In jobCurrentCollection
                              Group Join jbb In jobBeforeCollection
                           On jbc.EMPLID Equals jbb.EMPLID And jbc.PAYGROUP Equals jbb.PAYGROUP Into g = Group
                              From jbb In g.DefaultIfEmpty()
                              Where IsNothing(jbb)
                              Select jbc

            ' Get records that have changed
            Dim changedJobData = From jbc In jobCurrentCollection
                                 Join jbb In jobBeforeCollection
                               On jbc.EMPLID Equals jbb.EMPLID And jbc.PAYGROUP Equals jbb.PAYGROUP
                                 Where jbc.PAY_FREQUENCY <> jbb.PAY_FREQUENCY Or
                               jbc.EMPL_STATUS <> jbb.EMPL_STATUS Or
                               jbc.LOCATION <> jbb.LOCATION Or
                               jbc.FULL_PART_TIME <> jbb.FULL_PART_TIME Or
                               jbc.COMPANY <> jbb.COMPANY Or
                               jbc.EMPL_TYPE <> jbb.EMPL_TYPE Or
                               jbc.EMPL_CLASS <> jbb.EMPL_CLASS Or
                               jbc.DATA_CONTROL <> jbb.DATA_CONTROL Or
                               jbc.FILE_NBR <> jbb.FILE_NBR Or
                               jbc.HOME_DEPARTMENT <> jbb.HOME_DEPARTMENT Or
                               jbc.TITLE <> jbb.TITLE Or
                               jbc.WORKERS_COMP_CD <> jbb.WORKERS_COMP_CD
                                 Select jbc

            Dim newJobData = From jbc In newJobData1
                             Group Join jbb In newPersDataMissing
                           On jbc.EMPLID Equals jbb.EMPLID And jbc.PAYGROUP Equals jbb.PAYGROUP Into g = Group
                             From jbb In g.DefaultIfEmpty()
                             Where IsNothing(jbb)
                             Select jbc

            ' Combine the result sets from both the new and changed data
            Dim jobDataToUpdate = Enumerable.Union(newJobData, changedJobData)

            ' Get records that have changed for VIPs
            Dim changedVIPJobData = (From pdc In jobDataToUpdate
                                     Join vip In vipDetails
                                    On pdc.EMPLID Equals vip.EMPLID
                                     Select vip).Distinct()

            Dim jobTermedEmpData = (From jbd In jobDataToUpdate
                                    Join emp In termedEmploymentDataToUpdate
                                        On jbd.EMPLID Equals emp.EMPLID And jbd.PAYGROUP Equals emp.PAYGROUP
                                    Select jbd).Distinct()

            ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
            If jobDataToUpdate.Count > 0 Then

                Using sw As StreamWriter = New StreamWriter(createFileLocation & "Job.csv")

                    If Not jobDataToUpdate.Count = jobTermedEmpData.Count Then
                        ' Add the header row
                        sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,EMPL_STATUS,ACTION_REASON,LOCATION,FULL_PART_TIME,COMPANY,EMPL_TYPE,EMPL_CLASS," &
                                 "DATA_CONTROL,FILE_NBR,HOME_DEPARTMENT,TITLE,WORKERS_COMP_CD")
                    Else
                        sw.Close()
                        'To remove the empty csv file
                        File.Delete(createFileLocation & "Job.csv")
                    End If

                    ' Add the data rows
                    For Each jb In jobDataToUpdate

                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                Where dns.EMPLID = jb.EMPLID And
                                                  dns.PAYGROUP = jb.PAYGROUP
                                                Select dns

                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendTermedEmpToADP = From dns In jobTermedEmpData
                                                      Where dns.EMPLID = jb.EMPLID And
                                                        dns.PAYGROUP = jb.PAYGROUP
                                                      Select dns

                        ' Make sure the emp is not in the do not send to ADP list before writing to file
                        If doNotSendEmpToADP.Count = 0 And doNotSendTermedEmpToADP.Count = 0 Then
                            ' Write object to csv file
                            dataRow = New StringBuilder
                            dataRow.Append("""" & jobValue & """")
                            dataRow.Append(",""" & jb.EMPLID & """")
                            dataRow.Append(",""" & jb.PAYGROUP & """")
                            dataRow.Append(",""" & jb.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & jb.EMPL_STATUS & """")
                            dataRow.Append(",""" & jb.ACTION_REASON & """")
                            dataRow.Append(",""" & jb.LOCATION & """")
                            dataRow.Append(",""" & jb.FULL_PART_TIME & """")
                            dataRow.Append(",""" & jb.COMPANY & """")
                            dataRow.Append(",""" & jb.EMPL_TYPE & """")
                            dataRow.Append(",""" & jb.EMPL_CLASS & """")
                            dataRow.Append(",""" & jb.DATA_CONTROL & """")
                            dataRow.Append(",""" & jb.FILE_NBR & """")
                            dataRow.Append(",""" & jb.HOME_DEPARTMENT & """")
                            dataRow.Append(",""" & jb.TITLE & """")
                            dataRow.Append(",""" & jb.WORKERS_COMP_CD & """")
                            sw.WriteLine(dataRow)
                        Else
                            doNotSendJobRecs += 1
                        End If

                        ' Insert object to database
                        DataManager.InsertJob(sqlConn, sqlTrans, jb.EMPLID, jb.PAYGROUP, jb.PAY_FREQUENCY, jb.EMPL_STATUS, jb.ACTION_REASON, jb.LOCATION,
                                              jb.FULL_PART_TIME, jb.COMPANY, jb.EMPL_TYPE, jb.EMPL_CLASS, jb.DATA_CONTROL, jb.FILE_NBR,
                                              jb.HOME_DEPARTMENT, jb.TITLE, jb.WORKERS_COMP_CD, runID)

                    Next
                    ' Close and dispose the streamwriter
                    sw.Close()
                End Using
            End If
            ' **************************************************************************************
            ' **************************************************************************************


            ' ******************************* GENERAL DEDUCTIONS FILE ******************************
            ' **************************************************************************************

            If generalDeductionFinalCollections.Count > 0 Then
                Using sw As StreamWriter = New StreamWriter(createFileLocation & "Deductions.csv")

                    ' Add the header row
                    sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,DEDCD,DED_ADDL_AMT,DED_RATE_PCT,GOAL_AMT,END_DT") ' !!! REMOVE the FILE_NBR header after conversion is complete !!!

                    ' Add the data rows
                    For Each de In generalDeductionFinalCollections

                        ' Write object to csv file
                        dataRow = New StringBuilder
                        dataRow.Append("""" & deductionsValue & """")
                        dataRow.Append(",""" & de.EMPLID & """")
                        dataRow.Append(",""" & de.PAYGROUP & """")
                        dataRow.Append(",""" & de.PAY_FREQUENCY & """")
                        dataRow.Append(",""" & de.DEDCD & """")
                        dataRow.Append(",""" & de.DED_ADDL_AMT.ToString & """")
                        dataRow.Append(",""" & de.DED_RATE_PCT.ToString & """")
                        dataRow.Append(",""" & de.GOAL_AMT.ToString & """")
                        dataRow.Append(",""" & de.END_DT.ToShortDateString & """")
                        'dataRow.Append(",""" & de.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!
                        sw.WriteLine(dataRow)

                        ' Insert object to database
                        DataManager.InsertGeneralDeduction(sqlConn, sqlTrans, de.EMPLID, de.PAYGROUP, de.PAY_FREQUENCY, de.DEDCD, de.DED_ADDL_AMT,
                                                               de.DED_RATE_PCT, de.GOAL_AMT, de.END_DT, runID)
                    Next

                    ' Close and dispose the streamwriter
                    sw.Close()
                End Using
            End If
            ' **************************************************************************************
            ' **************************************************************************************


            ' ******************************** DIRECT DEPOSIT FILE *********************************
            ' **************************************************************************************
            ' Get new records
            Dim newDirectDepositData1 = From ddc In directDepositCurrentCollection
                                        Group Join ddb In directDepositBeforeCollection
                                       On ddc.EMPLID Equals ddb.EMPLID And
                                       ddc.PAYGROUP Equals ddb.PAYGROUP And
                                       ddc.TRANSIT_NBR Equals ddb.TRANSIT_NBR And
                                       ddc.ACCOUNT_NBR Equals ddb.ACCOUNT_NBR And
                                       ddc.DEDCD Equals ddb.DEDCD Into g = Group
                                        From ddb In g.DefaultIfEmpty()
                                        Where IsNothing(ddb) And ddc.AccountIsInactive = "N"
                                        Select ddc

            ' Get records that have changed
            Dim changedDirectDepositData = From ddc In directDepositCurrentCollection
                                           Join ddb In directDepositBeforeCollection
                                           On ddc.EMPLID Equals ddb.EMPLID And
                                           ddc.PAYGROUP Equals ddb.PAYGROUP And
                                           ddc.TRANSIT_NBR Equals ddb.TRANSIT_NBR And
                                           ddc.ACCOUNT_NBR Equals ddb.ACCOUNT_NBR And
                                           ddc.DEDCD Equals ddb.DEDCD
                                           Where ddc.PAY_FREQUENCY <> ddb.PAY_FREQUENCY Or
                                           ddc.FULL_DEPOSIT <> ddb.FULL_DEPOSIT Or
                                           ddc.DEPOSIT_AMT <> ddb.DEPOSIT_AMT Or
                                           ddc.AccountIsInactive <> ddb.AccountIsInactive
                                           Select ddc

            Dim newDirectDepositData = From ddc In newDirectDepositData1
                                       Group Join ddb In newPersDataMissing
                                       On ddc.EMPLID Equals ddb.EMPLID And
                                       ddc.PAYGROUP Equals ddb.PAYGROUP Into g = Group
                                       From ddb In g.DefaultIfEmpty()
                                       Where IsNothing(ddb)
                                       Select ddc

            ' Combine the result sets from both the new and changed data
            Dim directDepositDataToUpdate = Enumerable.Union(newDirectDepositData, changedDirectDepositData)

            ' Get records that have changed for VIPs
            Dim changedVIPDirectDepositData = (From pdc In directDepositDataToUpdate
                                               Join vip In vipDetails
                                            On pdc.EMPLID Equals vip.EMPLID
                                               Select vip).Distinct()

            ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
            If directDepositDataToUpdate.Count > 0 Then

                Using sw As StreamWriter = New StreamWriter(createFileLocation & "Direct Dep.csv")

                    ' Add the header row
                    sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,DEDCD,FULL_DEPOSIT,TRANSIT_NBR,ACCOUNT_NBR,DEPOSIT_AMT,END_DT") ' !!! REMOVE the FILE_NBR header after conversion is complete !!!

                    ' Add the data rows
                    For Each dd In directDepositDataToUpdate

                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                Where dns.EMPLID = dd.EMPLID And
                                                  dns.PAYGROUP = dd.PAYGROUP
                                                Select dns

                        ' Make sure the emp is not in the do not send to ADP list before writing to file
                        If doNotSendEmpToADP.Count = 0 Then

                            ' Write object to csv file
                            dataRow = New StringBuilder
                            dataRow.Append("""" & directdepValue & """")
                            dataRow.Append(",""" & dd.EMPLID & """")
                            dataRow.Append(",""" & dd.PAYGROUP & """")
                            dataRow.Append(",""" & dd.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & dd.DEDCD & """")
                            dataRow.Append(",""" & dd.FULL_DEPOSIT & """")
                            dataRow.Append(",""" & dd.TRANSIT_NBR & """")
                            dataRow.Append(",""" & dd.ACCOUNT_NBR & """")
                            dataRow.Append(",""" & dd.DEPOSIT_AMT.ToString & """")
                            dataRow.Append(",""" & dd.END_DT.ToShortDateString & """")
                            'dataRow.Append(",""" & dd.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!
                            sw.WriteLine(dataRow)
                        Else
                            doNotSendDirDepRecs += 1
                        End If

                        ' Insert object to database
                        DataManager.InsertDirectDeposit(sqlConn, sqlTrans, dd.EMPLID, dd.PAYGROUP, dd.PAY_FREQUENCY, dd.DEDCD, dd.FULL_DEPOSIT,
                                                        dd.TRANSIT_NBR, dd.ACCOUNT_NBR, dd.DEPOSIT_AMT, dd.END_DT, dd.AccountIsInactive, runID)

                    Next

                    ' Close and dispose the streamwriter
                    sw.Close()
                End Using
            End If
            ' **************************************************************************************
            ' **************************************************************************************


            ' *********************************** W4 DATA FILE *************************************
            ' **************************************************************************************
            ' Get new records
            Dim newW4Data1 = From w4c In w4DataCurrentCollection
                             Group Join w4b In w4DataBeforeCollection
                            On w4c.EMPLID Equals w4b.EMPLID And w4c.PAYGROUP Equals w4b.PAYGROUP Into g = Group
                             From w4b In g.DefaultIfEmpty()
                             Where IsNothing(w4b)
                             Select w4c

            ' Get records that have changed
            Dim changedW4Data = From w4c In w4DataCurrentCollection
                                Join w4b In w4DataBeforeCollection
                                On w4c.EMPLID Equals w4b.EMPLID And w4c.PAYGROUP Equals w4b.PAYGROUP
                                Where w4c.PAY_FREQUENCY <> w4b.PAY_FREQUENCY Or
                                    w4c.STATE_TAX_CD <> w4b.STATE_TAX_CD Or
                                    w4c.TAX_BLOCK <> w4b.TAX_BLOCK Or
                                    w4c.MARITAL_STATUS <> w4b.MARITAL_STATUS Or
                                    w4c.EXEMPTIONS <> w4b.EXEMPTIONS Or
                                    w4c.EXEMPT_DOLLARS <> w4b.EXEMPT_DOLLARS Or
                                    w4c.ADDL_TAX_AMT <> w4b.ADDL_TAX_AMT Or
                                    w4c.STATE_WH_TABLE <> w4b.STATE_WH_TABLE
                                Select w4c

            Dim newW4Data = From w4c In newW4Data1
                            Group Join w4b In newPersDataMissing
                           On w4c.EMPLID Equals w4b.EMPLID And w4c.PAYGROUP Equals w4b.PAYGROUP Into g = Group
                            From w4b In g.DefaultIfEmpty()
                            Where IsNothing(w4b)
                            Select w4c

            ' Combine the result sets from both the new and changed data
            Dim w4DataToupdate = Enumerable.Union(newW4Data, changedW4Data)

            ' Get records that have changed for VIPs
            Dim changedVIPW4Data = (From pdc In w4DataToupdate
                                    Join vip In vipDetails
                                    On pdc.EMPLID Equals vip.EMPLID
                                    Select vip).Distinct()

            ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
            If w4DataToupdate.Count > 0 Then

                Using sw As StreamWriter = New StreamWriter(createFileLocation & "W4 Data.csv")

                    ' Add the header row
                    'sw.WriteLine("EMPLID,PAYGROUP,PAY_FREQUENCY,STATE_TAX_CD,TAX_BLOCK,MARITAL_STATUS,EXEMPTIONS,EXEMPT_DOLLARS,ADDL_TAX_AMT,STATE_WH_TABLE") ' !!! REMOVE the FILE_NBR header after conversion is complete !!!
                    'New 2020 fields added to the end.
                    sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,STATE_TAX_CD,TAX_BLOCK,MARITAL_STATUS,EXEMPTIONS,EXEMPT_DOLLARS,ADDL_TAX_AMT,STATE_WH_TABLE") ' !!! REMOVE the FILE_NBR header after conversion is complete !!!

                    ' Add the data rows
                    For Each wf In w4DataToupdate

                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                Where dns.EMPLID = wf.EMPLID And
                                                  dns.PAYGROUP = wf.PAYGROUP
                                                Select dns

                        ' Make sure the emp is not in the do not send to ADP list before writing to file
                        If doNotSendEmpToADP.Count = 0 Then

                            ' Write object to csv file
                            dataRow = New StringBuilder
                            dataRow.Append("""" & w4dataValue & """")
                            dataRow.Append(",""" & wf.EMPLID & """")
                            dataRow.Append(",""" & wf.PAYGROUP & """")
                            dataRow.Append(",""" & wf.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & wf.STATE_TAX_CD & """")
                            dataRow.Append(",""" & wf.TAX_BLOCK & """")
                            dataRow.Append(",""" & wf.MARITAL_STATUS & """")
                            dataRow.Append(",""" & wf.EXEMPTIONS & """")
                            dataRow.Append(",""" & wf.EXEMPT_DOLLARS & """")
                            dataRow.Append(",""" & wf.ADDL_TAX_AMT.ToString & """")
                            dataRow.Append(",""" & wf.STATE_WH_TABLE & """")
                            'dataRow.Append(",""" & wf.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!

                            sw.WriteLine(dataRow)
                        Else
                            doNotSendW4DataRecs += 1
                        End If

                        ' Insert object to database
                        DataManager.InsertW4Data(sqlConn, sqlTrans, wf.EMPLID, wf.PAYGROUP, wf.PAY_FREQUENCY, wf.STATE_TAX_CD, wf.TAX_BLOCK, wf.MARITAL_STATUS, wf.EXEMPTIONS,
                                                 wf.EXEMPT_DOLLARS, wf.ADDL_TAX_AMT, wf.STATE_WH_TABLE, runID)

                    Next

                    ' Close and dispose the streamwriter
                    sw.Close()
                End Using
            End If
            ' **************************************************************************************
            ' **************************************************************************************


            ' *********************************** TAX DATA FILE ************************************
            ' **************************************************************************************

            'Update the TAX_LOCK_END_DT  
            'This Rule was In the usp_PayGetTaxDataCurrent, since we are no longer using the OnPremise SP
            'we replicate the rule here.
            For Each item As TaxData In taxDataCurrentCollection.Where(Function(w) w.W4IsLocked = "N").ToList()
                Dim bfTaxData As TaxData = taxDataBeforeCollection.Where(Function(w) w.EMPLID = item.EMPLID And w.PAYGROUP = item.PAYGROUP And w.W4IsLocked = "Y").FirstOrDefault()

                If bfTaxData IsNot Nothing Then
                    item.TAX_LOCK_END_DT = Date.Today
                End If
            Next

            ' Get new records
            Dim newTaxData1 = From txc In taxDataCurrentCollection
                              Group Join txb In taxDataBeforeCollection
                            On txc.EMPLID Equals txb.EMPLID And txc.PAYGROUP Equals txb.PAYGROUP Into g = Group
                              From txb In g.DefaultIfEmpty()
                              Where IsNothing(txb)
                              Select txc
            'Select Case txc.EMPLID, txc.FEDERAL_ADDL_AMT, txc.DEPENDENTS_AMT, txc.FEDERAL_MAR_STATUS, txc.FEDERAL_TAX_BLOCK, txc.FED_ALLOWANCES, txc.FILE_NBR, txc.LOCAL2_TAX_CD, txc.LOCAL4_TAX_CD, txc.LOCAL_TAX_CD, txc.Long_Term_Care_Ins_Status, txc.MULTIPLE_JOBS, txc.OTH_DEDUCTIONS, txc.OTH_INCOME, txc.PAYGROUP, txc.PAY_FREQUENCY, txc.SCHDIST_TAX_BLOCK, txc.SCHOOL_DISTRICT, txc.SSMED_TAX_BLOCK, txc.STATE2_TAX_CD, txc.STATE_TAX_CD, txc.SUISDI_TAX_BLOCK, txc.SUI_TAX_CD, txc.TAX_LCK_FED_MAR_ST, txc.TAX_LOCK_END_DT, txc.TAX_LOCK_FED_ALLOW, txc.W4IsLocked, txc.W4_FORM_YEAR


            ' Get records that have changed
            Dim changedTaxData = From txc In taxDataCurrentCollection
                                 Join txb In taxDataBeforeCollection
                                On txc.EMPLID Equals txb.EMPLID And txc.PAYGROUP Equals txb.PAYGROUP
                                 Where txc.PAY_FREQUENCY <> txb.PAY_FREQUENCY Or
                                    txc.FEDERAL_MAR_STATUS <> txb.FEDERAL_MAR_STATUS Or
                                    txc.FED_ALLOWANCES <> txb.FED_ALLOWANCES Or
                                    txc.FEDERAL_TAX_BLOCK <> txb.FEDERAL_TAX_BLOCK Or
                                    txc.SCHDIST_TAX_BLOCK <> txb.SCHDIST_TAX_BLOCK Or
                                    txc.SUISDI_TAX_BLOCK <> txb.SUISDI_TAX_BLOCK Or
                                    txc.SSMED_TAX_BLOCK <> txb.SSMED_TAX_BLOCK Or
                                    txc.FEDERAL_ADDL_AMT <> txb.FEDERAL_ADDL_AMT Or
                                    txc.STATE_TAX_CD <> txb.STATE_TAX_CD Or
                                    txc.STATE2_TAX_CD <> txb.STATE2_TAX_CD Or
                                    txc.LOCAL_TAX_CD <> txb.LOCAL_TAX_CD Or
                                    txc.LOCAL2_TAX_CD <> txb.LOCAL2_TAX_CD Or
                                    txc.SCHOOL_DISTRICT <> txb.SCHOOL_DISTRICT Or
                                    txc.SUI_TAX_CD <> txb.SUI_TAX_CD Or
                                    txc.TAX_LCK_FED_MAR_ST <> txb.TAX_LCK_FED_MAR_ST Or
                                    txc.TAX_LOCK_FED_ALLOW <> txb.TAX_LOCK_FED_ALLOW Or
                                    txc.LOCAL4_TAX_CD <> txb.LOCAL4_TAX_CD Or
                                    txc.W4IsLocked <> txb.W4IsLocked Or
                                    txc.W4_FORM_YEAR <> txb.W4_FORM_YEAR Or
                                    txc.OTH_INCOME <> txb.OTH_INCOME Or
                                    txc.OTH_DEDUCTIONS <> txb.OTH_DEDUCTIONS Or
                                    txc.DEPENDENTS_AMT <> txb.DEPENDENTS_AMT Or
                                    txc.MULTIPLE_JOBS <> txb.MULTIPLE_JOBS Or
                                    txc.Long_Term_Care_Ins_Status <> txb.Long_Term_Care_Ins_Status
                                 Select txc

            Dim newTaxData = From txc In newTaxData1
                             Group Join txb In newPersDataMissing
                            On txc.EMPLID Equals txb.EMPLID And txc.PAYGROUP Equals txb.PAYGROUP Into g = Group
                             From txb In g.DefaultIfEmpty()
                             Where IsNothing(txb)
                             Select txc

            ' Combine the result sets from both the new and changed data
            Dim taxDataToupdate = Enumerable.Union(newTaxData, changedTaxData)

            ' Get records that have changed for VIPs
            Dim changedVIPTaxData = (From pdc In taxDataToupdate
                                     Join vip In vipDetails
                                    On pdc.EMPLID Equals vip.EMPLID
                                     Select vip).Distinct()

            ' Get the employee W4 records that did not have a tax data file sent
            Dim w4SentButNoTax = From w4s In w4DataToupdate
                                 Group Join txu In taxDataToupdate
                                 On w4s.EMPLID Equals txu.EMPLID And w4s.PAYGROUP Equals txu.PAYGROUP Into g = Group
                                 From txu In g.DefaultIfEmpty()
                                 Where IsNothing(txu)
                                 Select w4s

            ' Make sure the emps are not in the do not send to ADP list before continueing
            Dim w4MissingTaxData = From w4u In w4SentButNoTax
                                   Group Join dns In doNotSendToADPCollection
                                   On w4u.EMPLID Equals dns.EMPLID And w4u.PAYGROUP Equals dns.PAYGROUP Into g = Group
                                   From dns In g.DefaultIfEmpty()
                                   Where IsNothing(dns)
                                   Select w4u

            If taxDataToupdate.Count + w4MissingTaxData.Count > 0 Then
                Using sw As StreamWriter = New StreamWriter(createFileLocation & "Tax Data.csv")

                    ' Add the header row
                    sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,FEDERAL_MAR_STATUS,FED_ALLOWANCES,FEDERAL_TAX_BLOCK,SCHDIST_TAX_BLOCK,SUISDI_TAX_BLOCK," &
                                 "SSMED_TAX_BLOCK,FEDERAL_ADDL_AMT,STATE_TAX_CD,STATE2_TAX_CD,LOCAL_TAX_CD,LOCAL2_TAX_CD,SCHOOL_DISTRICT,SUI_TAX_CD," &
                                 "TAX_LOCK_END_DT,TAX_LCK_FED_MAR_ST,TAX_LOCK_FED_ALLOW,LOCAL4_TAX_CD,W4_FORM_YEAR,OTH_INCOME,OTH_DEDUCTIONS,DEPENDENTS_AMT,MULTIPLE_JOBS,Long_Term_Care_Ins_Status") ' !!! REMOVE the FILE_NBR header after conversion is complete !!!

                    ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
                    If newTaxData.Count > 0 Then
                        ' Add the data rows
                        For Each tx In newTaxData

                            ' Check if emp is in do not send to ADP list
                            Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                    Where dns.EMPLID = tx.EMPLID And
                                                      dns.PAYGROUP = tx.PAYGROUP
                                                    Select dns

                            ' Make sure the emp is not in the do not send to ADP list before writing to file
                            If doNotSendEmpToADP.Count = 0 Then

                                ' Write object to csv file
                                dataRow = New StringBuilder
                                dataRow.Append("""" & taxdataValue & """")
                                dataRow.Append(",""" & tx.EMPLID & """")
                                dataRow.Append(",""" & tx.PAYGROUP & """")
                                dataRow.Append(",""" & tx.PAY_FREQUENCY & """")
                                dataRow.Append(",""" & tx.FEDERAL_MAR_STATUS & """")
                                dataRow.Append(",""" & tx.FED_ALLOWANCES & """")
                                dataRow.Append(",""" & tx.FEDERAL_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.SCHDIST_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.SUISDI_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.SSMED_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.FEDERAL_ADDL_AMT.ToString & """")
                                dataRow.Append(",""" & tx.STATE_TAX_CD & """")
                                dataRow.Append(",""" & tx.STATE2_TAX_CD & """")
                                dataRow.Append(",""" & tx.LOCAL_TAX_CD & """")
                                dataRow.Append(",""" & tx.LOCAL2_TAX_CD & """")
                                dataRow.Append(",""" & tx.SCHOOL_DISTRICT & """")
                                dataRow.Append(",""" & tx.SUI_TAX_CD & """")
                                If tx.TAX_LOCK_END_DT <> Nothing Then
                                    dataRow.Append(",""" & tx.TAX_LOCK_END_DT.ToShortDateString & """")
                                Else
                                    dataRow.Append(",""""")
                                End If
                                dataRow.Append(",""" & tx.TAX_LCK_FED_MAR_ST & """")
                                dataRow.Append(",""" & tx.TAX_LOCK_FED_ALLOW & """")
                                dataRow.Append(",""" & tx.LOCAL4_TAX_CD & """")
                                'dataRow.Append(",""" & tx.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!
                                'New 2020 W4 Fields
                                'dataRow.Append(",""" & tx.USE_OLD_W4 & """")
                                dataRow.Append(",""" & tx.W4_FORM_YEAR & """")

                                dataRow.Append(",""" & IIf(tx.OTH_INCOME IsNot Nothing AndAlso tx.OTH_INCOME.Length > 0, tx.OTH_INCOME, "0").ToString & """")
                                dataRow.Append(",""" & IIf(tx.OTH_DEDUCTIONS IsNot Nothing AndAlso tx.OTH_DEDUCTIONS.Length > 0, tx.OTH_DEDUCTIONS, "0").ToString & """")
                                dataRow.Append(",""" & IIf(tx.DEPENDENTS_AMT IsNot Nothing AndAlso tx.DEPENDENTS_AMT.Length > 0, tx.DEPENDENTS_AMT, "0").ToString & """")

                                dataRow.Append(",""" & tx.MULTIPLE_JOBS & """")
                                dataRow.Append(",""" & IIf(tx.Long_Term_Care_Ins_Status = "N", "", "E").ToString & """")
                                sw.WriteLine(dataRow)
                            Else
                                doNotSendTaxDataRecs += 1
                            End If

                            ' Insert object to database
                            DataManager.InsertTaxData(sqlConn, sqlTrans, tx.EMPLID, tx.PAYGROUP, tx.PAY_FREQUENCY, tx.FEDERAL_MAR_STATUS, tx.FED_ALLOWANCES, tx.FEDERAL_TAX_BLOCK,
                                                      tx.SCHDIST_TAX_BLOCK, tx.SUISDI_TAX_BLOCK, tx.SSMED_TAX_BLOCK, tx.FEDERAL_ADDL_AMT, tx.STATE_TAX_CD,
                                                      tx.STATE2_TAX_CD, tx.LOCAL_TAX_CD, tx.LOCAL2_TAX_CD, tx.SCHOOL_DISTRICT, tx.SUI_TAX_CD,
                                                      tx.TAX_LOCK_END_DT, tx.TAX_LCK_FED_MAR_ST, tx.TAX_LOCK_FED_ALLOW, tx.LOCAL4_TAX_CD, tx.W4IsLocked, runID,
                                                      tx.W4_FORM_YEAR, tx.OTH_INCOME, tx.OTH_DEDUCTIONS, tx.DEPENDENTS_AMT, tx.MULTIPLE_JOBS, tx.Long_Term_Care_Ins_Status)

                        Next
                    End If

                    If changedTaxData.Count > 0 Then
                        ' Add the data rows
                        For Each tx In changedTaxData

                            ' Check if emp is in do not send to ADP list
                            Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                    Where dns.EMPLID = tx.EMPLID And
                                                      dns.PAYGROUP = tx.PAYGROUP
                                                    Select dns

                            ' Make sure the emp is not in the do not send to ADP list before writing to file
                            If doNotSendEmpToADP.Count = 0 Then

                                ' Write object to csv file
                                dataRow = New StringBuilder
                                dataRow.Append("""" & taxdataValue & """")
                                dataRow.Append(",""" & tx.EMPLID & """")
                                dataRow.Append(",""" & tx.PAYGROUP & """")
                                dataRow.Append(",""" & tx.PAY_FREQUENCY & """")
                                dataRow.Append(",""" & tx.FEDERAL_MAR_STATUS & """")
                                dataRow.Append(",""" & tx.FED_ALLOWANCES & """")
                                dataRow.Append(",""" & tx.FEDERAL_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.SCHDIST_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.SUISDI_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.SSMED_TAX_BLOCK & """")
                                dataRow.Append(",""" & tx.FEDERAL_ADDL_AMT.ToString & """")
                                dataRow.Append(",""" & tx.STATE_TAX_CD & """")
                                dataRow.Append(",""" & tx.STATE2_TAX_CD & """")
                                dataRow.Append(",""" & tx.LOCAL_TAX_CD & """")
                                dataRow.Append(",""" & tx.LOCAL2_TAX_CD & """")
                                dataRow.Append(",""" & tx.SCHOOL_DISTRICT & """")
                                dataRow.Append(",""" & tx.SUI_TAX_CD & """")
                                If tx.TAX_LOCK_END_DT <> Nothing Then
                                    dataRow.Append(",""" & tx.TAX_LOCK_END_DT.ToShortDateString & """")
                                Else
                                    dataRow.Append(",""""")
                                End If
                                dataRow.Append(",""" & tx.TAX_LCK_FED_MAR_ST & """")
                                dataRow.Append(",""" & tx.TAX_LOCK_FED_ALLOW & """")
                                dataRow.Append(",""" & tx.LOCAL4_TAX_CD & """")
                                'dataRow.Append(",""" & tx.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!
                                'New 2020 W4 Fields
                                'dataRow.Append(",""" & tx.USE_OLD_W4 & """")
                                dataRow.Append(",""" & tx.W4_FORM_YEAR & """")

                                dataRow.Append(",""" & IIf(tx.OTH_INCOME IsNot Nothing AndAlso tx.OTH_INCOME.Length > 0, tx.OTH_INCOME, "0").ToString & """")
                                dataRow.Append(",""" & IIf(tx.OTH_DEDUCTIONS IsNot Nothing AndAlso tx.OTH_DEDUCTIONS.Length > 0, tx.OTH_DEDUCTIONS, "0").ToString & """")
                                dataRow.Append(",""" & IIf(tx.DEPENDENTS_AMT IsNot Nothing AndAlso tx.DEPENDENTS_AMT.Length > 0, tx.DEPENDENTS_AMT, "0").ToString & """")

                                dataRow.Append(",""" & tx.MULTIPLE_JOBS & """")
                                dataRow.Append(",""" & IIf(tx.Long_Term_Care_Ins_Status = "N", "C", "E").ToString & """")
                                sw.WriteLine(dataRow)
                            Else
                                doNotSendTaxDataRecs += 1
                            End If

                            ' Insert object to database
                            DataManager.InsertTaxData(sqlConn, sqlTrans, tx.EMPLID, tx.PAYGROUP, tx.PAY_FREQUENCY, tx.FEDERAL_MAR_STATUS, tx.FED_ALLOWANCES, tx.FEDERAL_TAX_BLOCK,
                                                      tx.SCHDIST_TAX_BLOCK, tx.SUISDI_TAX_BLOCK, tx.SSMED_TAX_BLOCK, tx.FEDERAL_ADDL_AMT, tx.STATE_TAX_CD,
                                                      tx.STATE2_TAX_CD, tx.LOCAL_TAX_CD, tx.LOCAL2_TAX_CD, tx.SCHOOL_DISTRICT, tx.SUI_TAX_CD,
                                                      tx.TAX_LOCK_END_DT, tx.TAX_LCK_FED_MAR_ST, tx.TAX_LOCK_FED_ALLOW, tx.LOCAL4_TAX_CD, tx.W4IsLocked, runID,
                                                      tx.W4_FORM_YEAR, tx.OTH_INCOME, tx.OTH_DEDUCTIONS, tx.DEPENDENTS_AMT, tx.MULTIPLE_JOBS, tx.Long_Term_Care_Ins_Status)

                        Next
                    End If

                    ' ***** Check if W4 Data files were sent on each employee but no federal files sent *****
                    ' ADP requires that if w4 data was sent, that employee's respective federal data needs to be sent as well.                 

                    If w4MissingTaxData.Count > 0 Then

                        ' Get the additional tax records to send over to ADP
                        Dim addTaxDataToSend = From w4u In w4MissingTaxData
                                               Group Join txb In taxDataBeforeCollection
                                               On w4u.EMPLID Equals txb.EMPLID And w4u.PAYGROUP Equals txb.PAYGROUP Into g = Group
                                               From txb In g.DefaultIfEmpty()

                        ' Iterate through the additional employee federal tax records to send
                        ' holds info about any employees missing tax info that should have some.
                        Dim sbMissTaxData As StringBuilder = New StringBuilder()

                        For Each add In addTaxDataToSend

                            ' First, check to see if we have federal (tax data) information to send
                            If IsNothing(add.txb) Then
                                ' capture all emps missing tax info and throw exception below.
                                If sbMissTaxData.Length = 0 Then
                                    sbMissTaxData.AppendLine("Employees are missing his/her federal tax information.")
                                End If

                                sbMissTaxData.AppendLine(String.Format("EMPLID: {0}  CoID: {1}", add.w4u.EMPLID, add.w4u.PAYGROUP))
                                Continue For
                            End If

                            ' Write object to csv file
                            dataRow = New StringBuilder
                            'dataRow.Append("""" & w4dataValue & """")
                            dataRow.Append("""" & taxdataValue & """")
                            dataRow.Append(",""" & add.txb.EMPLID & """")
                            dataRow.Append(",""" & add.txb.PAYGROUP & """")
                            dataRow.Append(",""" & add.txb.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & add.txb.FEDERAL_MAR_STATUS & """")
                            dataRow.Append(",""" & add.txb.FED_ALLOWANCES & """")
                            dataRow.Append(",""" & add.txb.FEDERAL_TAX_BLOCK & """")
                            dataRow.Append(",""" & add.txb.SCHDIST_TAX_BLOCK & """")
                            dataRow.Append(",""" & add.txb.SUISDI_TAX_BLOCK & """")
                            dataRow.Append(",""" & add.txb.SSMED_TAX_BLOCK & """")
                            dataRow.Append(",""" & add.txb.FEDERAL_ADDL_AMT.ToString & """")
                            dataRow.Append(",""" & add.txb.STATE_TAX_CD & """")
                            dataRow.Append(",""" & add.txb.STATE2_TAX_CD & """")
                            dataRow.Append(",""" & add.txb.LOCAL_TAX_CD & """")
                            dataRow.Append(",""" & add.txb.LOCAL2_TAX_CD & """")
                            dataRow.Append(",""" & add.txb.SCHOOL_DISTRICT & """")
                            dataRow.Append(",""" & add.txb.SUI_TAX_CD & """")
                            If add.txb.TAX_LOCK_END_DT <> Nothing Then
                                dataRow.Append(",""" & add.txb.TAX_LOCK_END_DT.ToShortDateString & """")
                            Else
                                dataRow.Append(",""""")
                            End If
                            dataRow.Append(",""" & add.txb.TAX_LCK_FED_MAR_ST & """")
                            dataRow.Append(",""" & add.txb.TAX_LOCK_FED_ALLOW & """")
                            dataRow.Append(",""" & add.txb.LOCAL4_TAX_CD & """")
                            'New 2020 W4 Fields
                            'dataRow.Append(",""" & add.txb.USE_OLD_W4 & """")
                            dataRow.Append(",""" & add.txb.W4_FORM_YEAR & """")
                            dataRow.Append(",""" & IIf(add.txb.OTH_INCOME IsNot Nothing AndAlso add.txb.OTH_INCOME.Length > 0, add.txb.OTH_INCOME, "0").ToString & """")
                            dataRow.Append(",""" & IIf(add.txb.OTH_DEDUCTIONS IsNot Nothing AndAlso add.txb.OTH_DEDUCTIONS.Length > 0, add.txb.OTH_DEDUCTIONS, "0").ToString & """")
                            dataRow.Append(",""" & IIf(add.txb.DEPENDENTS_AMT IsNot Nothing AndAlso add.txb.DEPENDENTS_AMT.Length > 0, add.txb.DEPENDENTS_AMT, "0").ToString & """")

                            dataRow.Append(",""" & add.txb.MULTIPLE_JOBS & """")
                            dataRow.Append(",""" & IIf(add.txb.Long_Term_Care_Ins_Status = "N", "", "E").ToString & """")
                            'dataRow.Append(",""" & add.txb.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!
                            sw.WriteLine(dataRow)

                            addTaxDataRecordsSent += 1

                        Next
                        'Move here for missing tax data. That way we can capture all at once.
                        If sbMissTaxData.Length > 0 Then
                            Throw New InvalidOperationException(sbMissTaxData.ToString())
                        End If

                    End If
                    ' Close and dispose the streamwriter if started for tax data
                    sw.Close()
                End Using
            End If

            ' ************ ADD W4 DATA FOR THOSE EMPLOYEES THAT HAVE TAX RECORD SENT ***************
            ' Product Backlog Item 399881: SD - ADPC File Feed - Send W4 Tax Data anytime a tax change is sent
            ' **************************************************************************************
            ' Get the employee Tax records that did not have W4 data to be sent.
            Dim TaxSentButNoW4 = From txu In taxDataToupdate
                                 Group Join w4s In w4DataToupdate
                                On txu.EMPLID Equals w4s.EMPLID And txu.PAYGROUP Equals w4s.PAYGROUP Into g = Group
                                 From w4s In g.DefaultIfEmpty()
                                 Where IsNothing(w4s)
                                 Select txu

            ' Make sure the emps are not in the do not send to ADP list before continueing
            Dim TaxMissingW4Data = From tsn4 In TaxSentButNoW4
                                   Group Join dns In doNotSendToADPCollection
                                   On tsn4.EMPLID Equals dns.EMPLID And tsn4.PAYGROUP Equals dns.PAYGROUP Into g = Group
                                   From dns In g.DefaultIfEmpty()
                                   Where IsNothing(dns)
                                   Select tsn4

            ' Append/Add the current w4 data to output file and update the w4 table. 
            Dim w4DataToAdd = From w4c In w4DataCurrentCollection
                              Join tmw4 In TaxMissingW4Data
                                      On w4c.EMPLID Equals tmw4.EMPLID And w4c.PAYGROUP Equals tmw4.PAYGROUP
                              Select w4c

            If w4DataToAdd.Count > 0 Then

                ' set append to file true as the w4 May exist. If not create the file.
                Dim fileExists As Boolean = File.Exists(createFileLocation & "W4 Data.csv")
                Using sw As StreamWriter = New StreamWriter(createFileLocation & "W4 Data.csv", fileExists)
                    'file doesn't exist add header
                    If fileExists = False Then
                        sw.WriteLine("REC_ID,EMPLID,PAYGROUP,PAY_FREQUENCY,STATE_TAX_CD,TAX_BLOCK,MARITAL_STATUS,EXEMPTIONS,EXEMPT_DOLLARS,ADDL_TAX_AMT,STATE_WH_TABLE") ' !!! REMOVE the FILE_NBR header after conversion is complete !!!
                    End If

                    For Each wf In w4DataToAdd
                        ' Check if emp is in do not send to ADP list
                        Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                                Where dns.EMPLID = wf.EMPLID And
                                                  dns.PAYGROUP = wf.PAYGROUP
                                                Select dns

                        ' Make sure the emp is not in the do not send to ADP list before writing to file
                        If doNotSendEmpToADP.Count = 0 Then
                            ' Write object to csv file
                            dataRow = New StringBuilder
                            dataRow.Append("""" & w4dataValue & """")
                            dataRow.Append(",""" & wf.EMPLID & """")
                            dataRow.Append(",""" & wf.PAYGROUP & """")
                            dataRow.Append(",""" & wf.PAY_FREQUENCY & """")
                            dataRow.Append(",""" & wf.STATE_TAX_CD & """")
                            dataRow.Append(",""" & wf.TAX_BLOCK & """")
                            dataRow.Append(",""" & wf.MARITAL_STATUS & """")
                            dataRow.Append(",""" & wf.EXEMPTIONS & """")
                            dataRow.Append(",""" & wf.EXEMPT_DOLLARS & """")
                            dataRow.Append(",""" & wf.ADDL_TAX_AMT.ToString & """")
                            dataRow.Append(",""" & wf.STATE_WH_TABLE & """")
                            'dataRow.Append(",""" & wf.FILE_NBR & """") ' !!! REMOVE after conversion is complete !!!

                            sw.WriteLine(dataRow)
                        Else
                            doNotSendW4DataRecs += 1
                        End If

                        ' Insert object to database
                        DataManager.InsertW4Data(sqlConn, sqlTrans, wf.EMPLID, wf.PAYGROUP, wf.PAY_FREQUENCY, wf.STATE_TAX_CD, wf.TAX_BLOCK, wf.MARITAL_STATUS, wf.EXEMPTIONS,
                                                 wf.EXEMPT_DOLLARS, wf.ADDL_TAX_AMT, wf.STATE_WH_TABLE, runID)

                    Next

                    ' Close and dispose the streamwriter
                    sw.Close()
                End Using
            End If

            'This method validates partially inserted employees.  If any found send error and roll back the insertes.
            'ValidatePartialEmployees(runID) ' Commented checking partial employees as rollback the deployment.

            ' **************************************************************************************
            ' **************************************************************************************

            ' ********************** NOTIFY UNACCOUNTED FOR SKIPPED DEDUCTIONS *********************
            SendUnaccountedForSkippedDeductions()
            ' **************************************************************************************

            ' *************************** SEND CHANGE NOTIFICATIONS ********************************
            SendChangeNotifications()
            ' **************************************************************************************

            ' ********************** NOTIFY VIP User's Data Changes *********************
            SendVIPChangeNotifications(vipDetails, changedVIPPersonalData, changedVIPEmploymentData, changedVIPJobData, changedVIPDeductionData, changedVIPDirectDepositData, changedVIPTaxData, changedVIPW4Data)
            ' **************************************************************************************

            ' ********************* UPDATE HOURLY AND SALARY EMPLOYEE TABLES ***********************
            DataManager.ConvertErrorCollectionToDatatable()
            DataManager.UpdateHourlyEmployees(sqlConn, sqlTrans)
            DataManager.UpdateSalaryEmployees(sqlConn, sqlTrans)
            ' **************************************************************************************


            ' ***************************** UPDATE RUN LOG RECORD **********************************
            ' Update the file creation run log record
            DataManager.UpdateRunLogRecord(sqlConn,
                                           sqlTrans,
                                           runID,
                                           personalDataToUpdate.Count - doNotSendPersDataRecs,
                                           employmentDataToUpdate.Count - doNotSendEmployRecs,
                                           jobDataToUpdate.Count - doNotSendJobRecs,
                                           generalDeductionFinalCollections.Count,
                                           directDepositDataToUpdate.Count - doNotSendDirDepRecs,
                                           (taxDataToupdate.Count + addTaxDataRecordsSent) - doNotSendTaxDataRecs,
                                           (w4DataToupdate.Count + w4DataToAdd.Count) - doNotSendW4DataRecs,
                                           "PASS",
                                           Nothing)
            ' **************************************************************************************



            ' ********************************** ZIP .CSV FILES ************************************
            ' Check to see if CSV files are to be zipped up and sent to outbound directory
            Select Case AppSettings.Get("ZipFilesMoveToOutbound")
                Case "TRUE"

                    ' Zip, Move, Archive files
                    ZipFiles_MoveZip_ArchiveOriginalFiles(createFileLocation, outboundFileLocation, archiveFileLocation)
                Case "FALSE"

                    ' Don't do anything then as we just leave the .csv files in the create location
                Case Else

                    ' Throw error as value should always be set to True or False
                    Throw New InvalidOperationException("Please check the application config file and make sure the 'ZipFilesMoveToOutbound' setting has a value of 'True' or 'False'")
            End Select
            ' **************************************************************************************



            ' ****************************** COMIT THE TRANSACTION *********************************
            ' Commit the changes to the database
            sqlTrans.Commit()
            ' **************************************************************************************


        Catch ex As Exception

            ' Rollback the changes to the database
            sqlTrans.Rollback()

            ' Delete any csv files already created in the specified location
            For Each file As FileInfo In New DirectoryInfo(createFileLocation).GetFiles("*.csv")
                file.Delete()
            Next

            ' Throw exception
            Throw ex
        Finally
            If Not IsNothing(sqlTrans) Then
                sqlTrans.Dispose()
            End If
            If Not IsNothing(sqlConn) Then
                sqlConn.Close()
                sqlConn.Dispose()
            End If
        End Try
    End Sub
    ''' <summary>
    ''' 'This method validates partially inserted employees.  If any found send error and roll back the insertes.
    ''' </summary>
    ''' <param name="runID"></param>
    Private Sub ValidatePartialEmployees(runID As Int32)
        Dim partialEmployees As DataTable = ValidatePartialEmps(runID)
        Dim counter As Integer = 0

        Dim partialEmp As StringBuilder = New StringBuilder()
        For Each partialEmployee As DataRow In partialEmployees.Rows
            counter += 1
            If partialEmp.Length = 0 Then
                partialEmp.AppendLine("The following employee(s) are missing one or more required records.  The ADPC feed did not send any record changes for any employees. <br/> " & GetCssForEmailNotifications("tblheader") & "
                                 <th class='tdborder'>S.No</th><th class='tdborder'>Empl Id</th><th class='tdborder'>Paygroup</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Missing Record(s)</th></tr>")
            End If

            partialEmp.AppendLine(String.Format("<tr><td class='tdborder'>{0} </td><td class='tdborder'>{1} </td><td class='tdborder'>{2} </td><td class='tdborder'>{3} </td><td class='tdborder'>{4} </td><td class='tdborder'>{5} </td></tr>", counter, partialEmployee("EmplId"), partialEmployee("Paygroup"), partialEmployee("EecEmpNo"), partialEmployee("EmpName"), partialEmployee("MissingData")))
        Next

        If partialEmp.Length > 0 Then
            partialEmp.AppendLine(GetCssForEmailNotifications("tblfooter") & " <br/> <br/>")
            Throw New InvalidOperationException(partialEmp.ToString())
        End If

    End Sub

    ''' <summary>
    ''' 'This method calls stored procedure and return partially inserted employees.
    ''' </summary>
    ''' <param name="runID"></param>
    ''' <returns></returns>
    Private Function ValidatePartialEmps(runID As Int32) As DataTable

        Try
            Dim params = New Collection()
            params.Add(DataAccess.SetSQLParameterProperties("@FileCreateRunID", DbType.Int32, runID))

            'Datatable has been used instead of SqlDataReader
            Using partiallyEmployees As DataTable = DataManager.RemoveEmployeesWhoHasDataErrorsFromDatatable(
                                                            DirectCast(DataAccess.ExecuteStoredProcedure(
                                                           "Payroll.dbo.usp_PayValidatePartialEmployees",
                                                           DataAccess.StoredProcedureReturnType.DataTable,
                                                           params), DataTable))


                Return partiallyEmployees
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ' Zips the .csv files, then moves the .zip file to outbound directory, then archives the original .csv files
    Private Sub ZipFiles_MoveZip_ArchiveOriginalFiles(ByVal createFileLocation As String, ByVal outboundFileLocation As String, ByVal archiveFileLocation As String)
        Dim dateTimeNow As Date = Nothing
        Dim archiveFolderName As String = String.Empty
        Dim zipFileName As String = String.Empty
        Dim zipFile As String = String.Empty

        Try

            ' Set variable values
            dateTimeNow = Date.Now
            archiveFolderName = dateTimeNow.Year.ToString & Right("0" & dateTimeNow.Month.ToString, 2) &
                                Right("0" & dateTimeNow.Day.ToString, 2) & Right("0" & dateTimeNow.Hour.ToString, 2) &
                                Right("0" & dateTimeNow.Minute.ToString, 2) & Right("0" & dateTimeNow.Second.ToString, 2)
            zipFileName = "AshleyADPC_" & archiveFolderName & ".zip"
            zipFile = createFileLocation & zipFileName


            If Directory.GetFiles(createFileLocation, "*.csv").Count > 0 Then
                ' Zip the files
                Using zip As ZipFile = New ZipFile
                    For Each filename As String In Directory.GetFiles(createFileLocation, "*.csv")
                        zip.AddFile(filename, String.Empty)
                    Next
                    zip.Save(zipFile)
                End Using

                ' Move the zipped file to the outbound location
                If Directory.GetFiles(createFileLocation, "*.zip").Count > 0 Then
                    For Each file As String In Directory.GetFiles(createFileLocation, "*.zip")
                        System.IO.File.Move(file, outboundFileLocation & System.IO.Path.GetFileName(file))
                    Next
                Else
                    Throw New InvalidOperationException("Error creating a .zip file for the outbound FTP directory.  Please review prior to running the ADPC employee feed again.")
                End If

                ' Notify users that files were sent to ADP
                DataManager.SendStandardNotification("FILESSENT", "Files were sent to ADP.")

            End If

            ' Archive original files
            Directory.CreateDirectory(archiveFileLocation & archiveFolderName)

            For Each file As String In Directory.GetFiles(createFileLocation, "*.csv")
                System.IO.File.Move(file, archiveFileLocation & archiveFolderName & "\" & System.IO.Path.GetFileName(file))
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to remove the errored emplids from tax and personal data
    ''' </summary>
    Private Sub RemoveErroredDataFromTaxAndPersonalData()
        taxDataCurrentCollection = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(taxDataCurrentCollection)
        personalDataCurrentCollection = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(personalDataCurrentCollection)
    End Sub

    Private Sub RemoveErroredDataFromDeductionData()
        generalDeductionsBeforeCollection = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(generalDeductionsBeforeCollection)
        generalDeductionsCurrentCollection = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(generalDeductionsCurrentCollection)
        generalDeductionsCurrentAllCollection = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(generalDeductionsCurrentAllCollection)
        generalDeductionsDeletedCollection = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(generalDeductionsDeletedCollection)
        generalDeductionFinalCollections = DataManager.RemoveEmployeesWhoHasDataErrorsFromCollection(generalDeductionFinalCollections)
    End Sub
    ''' <summary>
    ''' This method is used to add the tax and personal data
    ''' </summary>
    Private Sub LoadTaxAndPersonalData()
        ' Tax Data Current
        taxDataCurrentCollection = DataManager.LoadTaxData(DataManager.TypeOfData.Current)
        ' Personal Data Current
        personalDataCurrentCollection = DataManager.LoadPersonalData(DataManager.TypeOfData.Current)
    End Sub

    ''' <summary>
    ''' This method is used to add the tax and personal data
    ''' </summary>
    Private Sub LoadDeductions()
        ' General Deductions Before
        generalDeductionsBeforeCollection = DataManager.LoadGeneralDeductionsData(DataManager.TypeOfData.Before)

        ' General Deductions Current
        DataManager.LoadGeneralDeductionsCurrent()

        doNotSendToADPCollection = DataManager.LoadDoNotSendToADP()

        ' General Deductions Deleted
        generalDeductionsDeletedCollection = DataManager.LoadGeneralDeductionsData(DataManager.TypeOfData.Deleted)

        Dim employeesWithGTD = generalDeductionsBeforeCollection.Where(Function(d) d.GTD_Amt <> 0).ToList()

    End Sub

    ''' <summary>
    ''' This method helps to buid deduction data
    ''' </summary>
    Private Sub BuildDeductionToSendADP()
        Dim futureEndDate As Date = Convert.ToDateTime(PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "FUTURE_END_DATE").Select(Function(cv) cv.Value1).FirstOrDefault()).Date

        generalDeductionFinalCollections = New Collection(Of GeneralDeduction)
        Dim vipDetails = DataManager.GetVIPDetails("EMPLOYEE_CHNG")

        Dim newPersDataMissing = From pdc In personalDataCurrentCollection
                                 Where pdc.FIRST_NAME = "" Or
                                 pdc.LAST_NAME = "" Or
                                 pdc.SEX = "" Or
                                (pdc.STREET1 = "" And pdc.STREET2 = "") Or
                                 pdc.SSN = "" Or
                                 pdc.STATE = "" Or
                                 pdc.CITY = "" Or
                                 pdc.PAYGROUP = "" Or
                                 pdc.PAY_FREQUENCY = "" Or
                                 pdc.ZIP = "" Or
                                 (pdc.ORIG_HIRE_DT.ToShortDateString = "" Or pdc.ORIG_HIRE_DT.ToShortDateString = "1/1/1900") Or
                                 (pdc.BIRTHDATE.ToShortDateString = "" Or pdc.BIRTHDATE.ToShortDateString = "1/1/1900")
                                 Select pdc

        Dim newDeductionData1 = From dec In generalDeductionsCurrentCollection
                                Group Join deb In generalDeductionsBeforeCollection
                                   On dec.EMPLID Equals deb.EMPLID And dec.PAYGROUP Equals deb.PAYGROUP And dec.DEDCD Equals deb.DEDCD Into g = Group
                                From deb In g.DefaultIfEmpty()
                                Where IsNothing(deb) And
                                   dec.SkipDeduction <> "Y" And (dec.END_DT = futureEndDate Or dec.END_DT > Date.Today)
                                Select dec

        ' Get records that have changed
        Dim changedDeductionData = From dec In generalDeductionsCurrentCollection
                                   Join deb In generalDeductionsBeforeCollection
                                       On dec.EMPLID Equals deb.EMPLID And dec.PAYGROUP Equals deb.PAYGROUP And dec.DEDCD Equals deb.DEDCD
                                   Where (dec.PAY_FREQUENCY <> deb.PAY_FREQUENCY Or
                                       dec.DED_ADDL_AMT <> deb.DED_ADDL_AMT Or
                                       dec.DED_RATE_PCT <> deb.DED_RATE_PCT Or
                                       dec.GOAL_AMT <> deb.GOAL_AMT Or
                                       dec.END_DT <> deb.END_DT) And
                                       dec.SkipDeduction <> "Y"
                                   Select dec

        Dim newDeductionData = From dec In newDeductionData1
                               Group Join deb In newPersDataMissing
                                      On dec.EMPLID Equals deb.EMPLID And dec.PAYGROUP Equals deb.PAYGROUP Into g = Group
                               From deb In g.DefaultIfEmpty()
                               Where IsNothing(deb)
                               Select dec

        ' Combine the result sets from both the new and changed data
        Dim deductionDataToupdate = Enumerable.Union(newDeductionData, changedDeductionData)

        ' Get records that have changed for VIPs
        changedVIPDeductionData = (From pdc In deductionDataToupdate
                                   Join vip In vipDetails
                                            On pdc.EMPLID Equals vip.EMPLID
                                   Select vip).Distinct()

        If deductionDataToupdate.Count + generalDeductionsDeletedCollection.Count > 0 Then

            ' Iterate through the personal data objects that are new or have changed and update database and write out to file to send to ADP
            If deductionDataToupdate.Count > 0 Then
                ' Add the data rows
                For Each deductionData In deductionDataToupdate

                    ' Check if emp is in do not send to ADP list
                    Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                            Where dns.EMPLID = deductionData.EMPLID And
                                                      dns.PAYGROUP = deductionData.PAYGROUP
                                            Select dns

                    ' Make sure the emp is not in the do not send to ADP list before writing to file
                    If doNotSendEmpToADP.Count = 0 Then
                        generalDeductionFinalCollections.Add(deductionData)
                    End If
                Next
            End If

            ' *** Check to see if a deduction was permanently deleted ***
            ' if this does exist, we need to send over the start date as the STOP date for the deduction - Per Brenda Pronschinske
            If generalDeductionsDeletedCollection.Count > 0 Then

                ' Add the data rows
                For Each deductionData In generalDeductionsDeletedCollection

                    ' Check if emp is in do not send to ADP list
                    Dim doNotSendEmpToADP = From dns In doNotSendToADPCollection
                                            Where dns.EMPLID = deductionData.EMPLID And
                                                      dns.PAYGROUP = deductionData.PAYGROUP
                                            Select dns

                    ' Make sure the emp is not in the do not send to ADP list before writing to file
                    If doNotSendEmpToADP.Count = 0 Then
                        generalDeductionFinalCollections.Add(deductionData)
                    End If

                Next
            End If
            ' Close and dispose the streamwriter
        End If

    End Sub

    ' Validates that the Ultipro data is correct before starting the file creation
    Private Sub ValidUltiproData()
        Dim validData As Boolean = True
        Dim empCounter As Integer = Nothing
        Dim eecSalaryOrHourlyValidationResults As DataTable = Nothing
        Dim loadECompValidationResults As DataTable = Nothing
        Dim futureTerminationResults As DataTable = Nothing
        Dim salaryEmpsReportingToHourlyEmps As DataTable = Nothing
        Dim invalidWorkInLiveInStateTaxes As DataTable = Nothing
        Dim invalidBenAmtForDISALDed As DataTable = Nothing
        Dim multipleLocalTaxes As DataTable = Nothing
        Dim extraTaxDollarsWithCents As DataTable = Nothing
        Dim partialDirDepAccountWithNoDolAmt As DataTable = Nothing
        Dim dirDepPercentRules As DataTable = Nothing
        Dim dedDependentMissing As DataTable = Nothing
        Dim dedDependentDOBMissing As DataTable = Nothing

        erroredTaxDataCurrentCollection = New List(Of TaxData)
        Dim validationMode As String = ""
        dataValidationErrors = New Collection(Of DataValidationError)

        Try

            ' Verify that the EecSalaryOrHourly field is correctly set for salary pay groups
            eecSalaryOrHourlyValidationResults = DataManager.ValidateEecSalaryOrHourly()
            If eecSalaryOrHourlyValidationResults.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify only one record per employee in LodEComp for the given pay period
            loadECompValidationResults = DataManager.ValidateLodEComp()
            If loadECompValidationResults.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify no employees with future termination dates
            futureTerminationResults = DataManager.ValidateNoFutureTermDates()
            If futureTerminationResults.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify no unaccounted for slaray employees are reporting to an hourly employee
            salaryEmpsReportingToHourlyEmps = DataManager.ValidateNoSalaryReprtingToHourly
            If salaryEmpsReportingToHourlyEmps.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify work in and live in tax options are correct
            invalidWorkInLiveInStateTaxes = DataManager.ValidateWorkInLiveInStateTax
            If invalidWorkInLiveInStateTaxes.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify all DISAL deductions are valid
            invalidBenAmtForDISALDed = DataManager.ValidateNoBenAmtForDISAL()
            If invalidBenAmtForDISALDed.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify no multiple local taxes
            multipleLocalTaxes = DataManager.ValidateNoMultipleLocalTaxes()
            If multipleLocalTaxes.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify no cents in the extra tax dollars field
            extraTaxDollarsWithCents = DataManager.ValidateNoCentsInExtraTaxDollars()
            If extraTaxDollarsWithCents.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify no partial direct deposit records with a zero dollar amount
            partialDirDepAccountWithNoDolAmt = DataManager.ValidateNoPartialDirDepAcctWithZeroDolAmt()
            If partialDirDepAccountWithNoDolAmt.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify no direct deposit percent rules for salary employees
            dirDepPercentRules = DataManager.ValidateNoPercentDirDepRules()
            If dirDepPercentRules.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify dependents are entered for all aplicable deductions
            dedDependentMissing = DataManager.ValidateDeductionDependentsExist()
            If dedDependentMissing.Rows.Count > 0 Then
                validData = False
            End If

            ' Verify DOB is entered on dependents for all applicable deductions
            dedDependentDOBMissing = DataManager.ValidateDeductionDependentsDOB()
            If dedDependentDOBMissing.Rows.Count > 0 Then
                validData = False
            End If

            'Verify SEND_NOTFN_FOR_W4_FORM_YEAR is 'N'
            Dim erroredTaxData = From txc In taxDataCurrentCollection
                                 Where txc.SEND_NOTFN_FOR_W4_FORM_YEAR = "Y"
                                 Select txc
            If erroredTaxData.Count > 0 Then
                validData = False
            End If

            erroredTaxDataCurrentCollection = (From txc In taxDataCurrentCollection
                                               Where
                                       txc.LOCAL_TAX_CD.Contains("_IS_INVALID_TAX_CODE") Or
                                       txc.LOCAL2_TAX_CD.Contains("_IS_INVALID_TAX_CODE") Or
                                       txc.LOCAL4_TAX_CD.Contains("_IS_INVALID_TAX_CODE")
                                               Select txc).ToList()


            If erroredTaxDataCurrentCollection.Count > 0 Then
                For Each tx In erroredTaxDataCurrentCollection
                    PopulateDataValidationErrors(tx.EMPLID + "0", tx.EecCoID, tx.PAYGROUP, tx.EMP_No)
                Next
            End If

            goalAmountMissingOnDeductions = (From ded In generalDeductionFinalCollections
                                             Where ded.SkipZeroGoalAmount = "Y").ToList()
            If goalAmountMissingOnDeductions.Count > 0 Then
                For Each goalAmountMissingOnDeduction In goalAmountMissingOnDeductions
                    PopulateDataValidationErrors(goalAmountMissingOnDeduction.EMPLID + "0",
                                                 goalAmountMissingOnDeduction.EECCoid,
                                                 goalAmountMissingOnDeduction.PAYGROUP,
                                                 goalAmountMissingOnDeduction.EmployeeNumber)
                Next
            End If

            validationMode = PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "ADPC_Feed_Test").Select(Function(cv) cv.Value1).FirstOrDefault()
            If (validationMode = "Test") Then
                validData = True
            End If

            ' If not valid data, format the error message to notify which employees caused the error
            If Not validData Then

                ultiproDataValidationError = "<br/><br/> Ultipro Data Is Not Valid. <br/><br/>"
                ' Additional message body text requested by Business - CJG 04/05/2022
                ultiproDataValidationError = String.Concat(ultiproDataValidationError,
                                                           "<span style="" font-weight:bold; font-style:italic;"">HR – Please correct below errors as quickly as possible and respond back to all on this email once complete so that the feed to ADP can be reattempted.  The feed is very time sensitive, so your prompt attention is appreciated.  Thank you.</span><br/><br/>")

                ' Check if EecSalaryOrHourly is out of sync with paygroup pay frequency
                If eecSalaryOrHourlyValidationResults.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) need their pay group to be aligned with their EecSalaryOrHourly field... <br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In eecSalaryOrHourlyValidationResults.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, ", Pay Group: ", row.Item("EecPayGroup").ToString, _
                    '    ", Salary Or Hourly: ", row.Item("EecSalaryOrHourly").ToString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th>" &
                                       "<th class='tdborder'>Pay Group</th><th class='trbtmborder'>Salary Or Hourly</th></tr>")
                    For Each row As DataRow In eecSalaryOrHourlyValidationResults.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecPayGroup").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='trbtmborder'>" & row.Item("EecSalaryOrHourly").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check if any employees have multiple pending changes in LodEComp
                If loadECompValidationResults.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have multiple Contingent changes for the pay period ran... <br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In loadECompValidationResults.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EecEEID: ", row.Item("EecEEID").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, ", Count of Pending Changes: ", row.Item("CountOfPendingChanges").ToString, _
                    '    Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EecEEID #</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th>" &
                                       "<th class='tdborder'>Count of Contingent Changes</th></tr>")
                    For Each row As DataRow In loadECompValidationResults.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEEID").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CountOfContingentChanges").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check if any employees have a termination date in the future
                If futureTerminationResults.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have termination dates in the future...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In futureTerminationResults.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, ", Termination Date: ", _
                    '    DirectCast(row.Item("EecDateOfTermination"), Date).ToShortDateString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th>" &
                                       "<th class='tdborder'>Termination Date</th></tr>")
                    For Each row As DataRow In futureTerminationResults.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & DirectCast(row.Item("EecDateOfTermination"), Date).ToShortDateString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check if there are unaccounted for salary employees reporting to hourly employees
                If salaryEmpsReportingToHourlyEmps.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following salary employee(s) are reporting to an hourly employee...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In salaryEmpsReportingToHourlyEmps.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                    For Each row As DataRow In salaryEmpsReportingToHourlyEmps.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check if there are invalid work in or live in state tax options
                If invalidWorkInLiveInStateTaxes.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) either have both 'Not subject to work in state tax' or 'Not subject to resident state tax' options ",
                        "selected or neither selected for those that live in a different state than where they work...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In invalidWorkInLiveInStateTaxes.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: " , row.Item("CmpCompanyCode").ToString , Environment.NewLine , Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                    For Each row As DataRow In invalidWorkInLiveInStateTaxes.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check if thare are invalid DISAL deduction records
                If invalidBenAmtForDISALDed.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have a benefit amount entered for a DISAL deduction...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In invalidBenAmtForDISALDed.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                    For Each row As DataRow In invalidBenAmtForDISALDed.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check to see if there are multiple local tax codes
                If multipleLocalTaxes.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have at least two of School, Occ, WC, or Other local tax codes populated...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In multipleLocalTaxes.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                    For Each row As DataRow In multipleLocalTaxes.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check if there are cents in the extra tax dollars
                If extraTaxDollarsWithCents.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have a cents entered in the extra tax dollars field...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In extraTaxDollarsWithCents.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, ", Extra Tax Dollars: ", row.Item("EetExtraTaxDollars").ToString, _
                    '    Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th><th class='tdborder'>Extra Tax Dollars</th></tr>")
                    For Each row As DataRow In extraTaxDollarsWithCents.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EetExtraTaxDollars").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check for any partial direct deposit accounts with zero dollar amounts
                If partialDirDepAccountWithNoDolAmt.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have a partial direct deposit account with a zero dollar amount...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In partialDirDepAccountWithNoDolAmt.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                    For Each row As DataRow In partialDirDepAccountWithNoDolAmt.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check for percent direct deposit rules
                If dirDepPercentRules.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have percent direct deposit rules...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In dirDepPercentRules.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". EE #: ", row.Item("EecEmpNo").ToString, _
                    '    ", Company: ", row.Item("CmpCompanyCode").ToString, Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                    For Each row As DataRow In dirDepPercentRules.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check for missing dependents on deductions
                If dedDependentMissing.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) are missing dependents associated with deductions...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In dedDependentMissing.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". Name: ", row.Item("Employee Name").ToString, _
                    '    ", EE #: ", row.Item("Employee Number").ToString, ", Company: ", row.Item("Company").ToString, _
                    '    ", Deduction: ", row.Item("Deduction").ToString, _
                    '    Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>Name</th><th class='tdborder'>EE #</th><th class='tdborder'>Company</th><th class='tdborder'>Deduction</th></tr>")
                    For Each row As DataRow In dedDependentMissing.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Employee Name").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Employee Number").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Company").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Deduction").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row, True)

                        empCounter += 1
                    Next
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                ' Check for deduction dependents missing their DOB
                If dedDependentDOBMissing.Rows.Count > 0 Then
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, Environment.NewLine, Environment.NewLine,
                        "The following employee(s) have deductions where the associated dependent is missing his/her DOB...<br/><br/>", Environment.NewLine)

                    'empCounter = 1
                    'For Each row As DataRow In dedDependentDOBMissing.Rows
                    '    ultiproDataValidationError = String.Concat(ultiproDataValidationError, empCounter.ToString, ". Name: ", row.Item("Employee Name").ToString, _
                    '    ", EE #: ", row.Item("Employee Number").ToString, ", Company: ", row.Item("Company").ToString, _
                    '    ", Deduction: ", row.Item("Deduction").ToString, ", Dependent: ", row.Item("Dependent Name").ToString, _
                    '    Environment.NewLine, Environment.NewLine)

                    '    empCounter += 1
                    'Next

                    empCounter = 1
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>Name</th><th class='tdborder'>EE #</th><th class='tdborder'>Company</th><th class='tdborder'>Deduction</th><th class='tdborder'>Dependent</th></tr>")
                    For Each row As DataRow In dedDependentDOBMissing.Rows
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Employee Name").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Employee Number").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Company").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Deduction").ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & row.Item("Dependent Name").ToString & "</td></tr>")

                        PopulateDataValidationErrors(row, True)

                        empCounter += 1
                    Next

                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                End If

                If erroredTaxData.Count > 0 Then
                    empCounter = 1

                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, "The following employees are either rehired or changed to salaried position with an old W4 form but their hire date is 2020 or later. Please review and correct W4 information to reflect 2020 or later W4 form to align with hire date.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                    'Create the table header to be rendered in the notification email
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>Employee ID</th><th class='tdborder'>EE#</th><th class='tdborder'>First Name</th>" &
                                       "<th class='tdborder'>Last Name</th><th class='tdborder'>Company Code</th><th class='tdborder'>W4 Form Year</th></tr>")


                    Dim employeesWithtaxDataWithWarning = From txc In erroredTaxData
                                                          Join pdc In personalDataCurrentCollection
                                                          On pdc.EMPLID Equals txc.EMPLID
                                                          Select txc, pdc
                    'For Each row As TaxData In taxDataWithWarning
                    For Each tx In employeesWithtaxDataWithWarning
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & tx.txc.EMPLID & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & tx.txc.EMP_No & "</td>")
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & tx.pdc.FIRST_NAME & "</td>") 'FirstName
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & tx.pdc.LAST_NAME & "</td>")  'LastName
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & tx.txc.PAYGROUP & "</td>") 'CompanyCode
                        ultiproDataValidationError = String.Concat(ultiproDataValidationError, "<td class='tdborder'>" & tx.txc.W4_FORM_YEAR & "</td></tr>") 'W4FormYear 

                        PopulateDataValidationErrors(tx.txc.EMPLID + "0", tx.txc.EecCoID, tx.txc.PAYGROUP, tx.txc.EMP_No)

                        empCounter += 1
                    Next

                    'Create the table footer to be rendered in the notification email
                    ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))

                End If

                DataManager.SendErrorNotification(String.Concat("Error in ADP File Creation Process:<br/>", Environment.NewLine,
                             Environment.NewLine, ultiproDataValidationError), "ERROR - ADPC Employee Feed - URGENT ATTENTION NEEDED")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Sends specific change notifications to defined recipients 
    Private Sub SendChangeNotifications()
        Dim salToHourOrHourToSalChanges As DataTable = Nothing
        Dim employeesWithChangedTaxData As DataTable = Nothing
        Dim employeesWithChangedDeductions As DataTable = Nothing
        Dim sameEmpNumberCompanyTransfers As DataTable = Nothing
        Dim changeMessage As String = String.Empty
        Dim empCounter As Int32 = 0
        Try

            ' ***** Get employees that that went from salary to hourly or from hourly to salary *****
            salToHourOrHourToSalChanges = DataManager.RemoveEmployeesWhoHasDataErrorsFromDatatable(
                                          DataManager.GetSalToHourOrHourToSalTransfers(runID)
                                          )

            Dim salToHourlyChanges As DataTable = Nothing
            If (From row In salToHourOrHourToSalChanges.AsEnumerable()
                Where row.Item("MovedTo").ToString.Equals("Moved to Hourly")
                Select row).FirstOrDefault IsNot Nothing Then

                salToHourlyChanges = (From row In salToHourOrHourToSalChanges.AsEnumerable()
                                      Where row.Item("MovedTo").ToString.Equals("Moved to Hourly")
                                      Select row).CopyToDataTable()
            End If

            Dim hourToSalaryChanges As DataTable = Nothing
            If (From row In salToHourOrHourToSalChanges.AsEnumerable()
                Where row.Item("MovedTo").ToString.Equals("Moved to Salary")
                Select row).FirstOrDefault IsNot Nothing Then

                hourToSalaryChanges = (From row In salToHourOrHourToSalChanges.AsEnumerable()
                                       Where row.Item("MovedTo").ToString.Equals("Moved to Salary")
                                       Select row).CopyToDataTable()

            End If


            If salToHourOrHourToSalChanges.Rows.Count > 0 Then

                ' Moved to Hourly
                If salToHourlyChanges IsNot Nothing Then
                    changeMessage = String.Concat("The following employees have moved from salary to hourly within the same company.  Please review the changes.<br/><br/>",
                              Environment.NewLine, Environment.NewLine)

                    empCounter = 1

                    For Each row As DataRow In salToHourlyChanges.Rows
                        ' skip move to salary handled below

                        If empCounter = 1 Then
                            changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>EE #</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Company</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Move</th></tr>")
                        End If

                        changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("EecEmpNo").ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("EmpName").ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("CmpCompanyCode").ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("MovedTo").ToString & "</td></tr>")
                        empCounter += 1
                    Next

                    If empCounter > 1 Then
                        changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"),
                                                        Environment.NewLine, Environment.NewLine)
                    End If
                End If ' End salHr changes


                'Moved to Salary
                If hourToSalaryChanges IsNot Nothing Then

                    changeMessage = String.Concat(changeMessage, Environment.NewLine, "<br/><br/>" &
                                    "The following employees have moved from hourly to salary within the same company.  Please review the changes.<br/><br/>",
                                    Environment.NewLine, Environment.NewLine)

                    ' Get a list of employees that have deductions that need to be in the message
                    Dim empWithDedsToInclude = (From emp In hourToSalaryChanges.AsEnumerable()
                                                Join ded In generalDeductionsCurrentAllCollection
                                                   On emp.Item("EecEmpNo") Equals ded.EmployeeNumber
                                                Join acd In AcceptableDedCodes
                                                   On ded.DEDCD Equals acd.DedDedCode
                                                Select EmpNo = emp.Item("EecEmpNo"), EmpName = emp.Item("EmpName"), CompanyCode = emp.Item("CmpCompanyCode"), MovedTo = emp.Item("MovedTo")
                                                ).Distinct().ToList()

                    'Get the list of deductions for the employees above
                    Dim empWithDedsToIncludeDeductions = (From emp In empWithDedsToInclude
                                                          Join ded In generalDeductionsCurrentAllCollection
                                                            On emp.EmpNo Equals ded.EmployeeNumber
                                                          Join acd In AcceptableDedCodes
                                                            On ded.DEDCD Equals acd.DedDedCode
                                                          Select emp.EmpNo, emp.CompanyCode, ded.DEDCD, cmpycd = ded.PAYGROUP, ded.DED_ADDL_AMT, ded.GOAL_AMT, ded.GTD_Amt, ded.DED_RATE_PCT, ded.START_DT, ded.END_DT
                                                        )


                    ' Get a list of employees that do not have deductions to be included in the message
                    Dim empsWithoutDedsToInclude = (From emp In hourToSalaryChanges.AsEnumerable()
                                                    Group Join ewd In empWithDedsToInclude
                                                       On emp.Item("EecEmpNo") Equals ewd.EmpNo
                                                       Into gl = Group
                                                    From g In gl.DefaultIfEmpty()
                                                    Where g Is Nothing
                                                    Select EmpNo = emp.Item("EecEmpNo"), EmpName = emp.Item("EmpName"), CompanyCode = emp.Item("CmpCompanyCode"), MovedTo = emp.Item("MovedTo")
                                                    ) 'REMOVE: Take(5)

                    empCounter = 1

                    'Emps without deduction info
                    For Each emp In empsWithoutDedsToInclude
                        ' Create the table and header
                        If empCounter = 1 Then
                            If Not changeMessage.Contains("<style") Then
                                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>EE #</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Company</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Move</th></tr>")
                            Else
                                changeMessage = String.Concat(changeMessage, "<br/><table class='tbl' cellspacing='0'><tr>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>EE #</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Company</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Move</th></tr>")
                            End If
                        End If

                        ' add the employee info
                        changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpNo.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpName.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.CompanyCode.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.MovedTo.ToString & "</td></tr>")

                        empCounter += 1
                    Next
                    ' close the above table
                    If empCounter > 1 Then
                        changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)
                    End If

                    'Emps with deduction info
                    empCounter = 1

                    For Each emp In empWithDedsToInclude
                        'new table for each employee with deductions
                        If Not changeMessage.Contains("<style") Then
                            changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader"))
                        Else
                            changeMessage = String.Concat(changeMessage, "<br/><table class='tbl' cellspacing='0'><tr>")
                        End If
                        'Add the employee info headers
                        changeMessage = String.Concat(changeMessage, "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>EE #</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Company</th>" &
                                       "<th class='tdborder' style='background-color: lightgray;'>Move</th>" &
                                       "<th class='tdborder' style='border-style:none; background-color: lightgray;'>&nbsp;</th>" &
                                       "<th class='tdborder' style='border-style:none; background-color: lightgray;'>&nbsp;</th>" &
                                       "<th class='tdborder' style='border-style:none; background-color: lightgray;'>&nbsp;</th></tr>")

                        'Add the employee Info
                        changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpNo.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpName.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.CompanyCode.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.MovedTo.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td></tr>")

                        'Add the deduction Info
                        Dim deductions = empWithDedsToIncludeDeductions.Where(Function(w) w.EmpNo.Equals(emp.EmpNo)).Select(Function(s) s)
                        'Header for deductions
                        changeMessage = String.Concat(changeMessage,
                                                      "<tr><th class='tdborder' style='background-color: lightgray;'>Company</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Cd.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Amt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Start Dt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Stop Dt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Goal Amt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>GTD Amt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Goal Bal.</th>" &
                                                      "</tr>"
                                                      )

                        For Each ded In deductions
                            changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'>" & ded.cmpycd.ToString & "</td>")
                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & ded.DEDCD.ToString & "</td>") 'Ded code

                            Dim formattedDedAmt As String = String.Empty
                            If ded.DED_ADDL_AMT > 0.0 Then
                                formattedDedAmt = ded.DED_ADDL_AMT.ToString("#####0.00")
                            ElseIf ded.DED_RATE_PCT > 0.0 Then
                                formattedDedAmt = ded.DED_RATE_PCT.ToString("P")
                            End If

                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & formattedDedAmt & "</td>") 'Ded Amt

                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & ded.START_DT & "</td>") 'Ded Start dt

                            If ded.END_DT = #1/1/2200 12:00:00 AM# Then
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & "&nbsp" & "</td>") 'Ded Stop Dt
                            Else
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & ded.END_DT & "</td>") 'Ded Stop Dt
                            End If

                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & ded.GOAL_AMT.ToString("#####0.00") & "</td>") 'Goal Amt
                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & ded.GTD_Amt.ToString("#####0.00") & "</td>") 'Goal YTD Amt


                            Dim goalBalance As Decimal = ded.GOAL_AMT - ded.GTD_Amt
                            If goalBalance <= 0 Then
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & "0.00" & "</td></tr>") 'Goal balance
                            Else
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & goalBalance.ToString("#####0.00") & "</td></tr>") 'Goal balance
                            End If

                        Next

                        ' close the above table
                        'changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)
                        changeMessage = String.Concat(changeMessage, "</table><br/>", System.Environment.NewLine)

                        empCounter += 1
                    Next ' End with deductions

                End If 'End Hourly to Sal changes

                ' Send the SaltoHourly HourlytoSal notification
                DataManager.SendStandardNotification("SALARYHOURLYTRANSFER", changeMessage)
            End If 'End Both SaltoHourly HourlytoSal

            ' ***** Get the employees that have transferred companies but have kept the same employee number *****
            ' This will only pertain to salary to salary and hourly to salary transfers
            sameEmpNumberCompanyTransfers = DataManager.RemoveEmployeesWhoHasDataErrorsFromDatatable(DataManager.GetSameEmpNumberCompanyTransfers())

            If sameEmpNumberCompanyTransfers.Rows.Count > 0 Then
                'List of employees that have deductions that need to be included in the message
                Dim empsWithDedsToInclude = (From emp In sameEmpNumberCompanyTransfers.AsEnumerable()
                                             Join ded In acceptableDeductionsFromToCompanyTransfers
                                                 On emp.Item("Employee Number") Equals ded.EmployeeNumber
                                             Join acd In AcceptableDedCodes
                                                 On ded.DEDCD Equals acd.DedDedCode
                                             Select EmpNo = emp.Item("Employee Number"), EmpName = emp.Item("EmpName"), OldCmpy = emp.Item("Old Company"), NewCmpy = emp.Item("New Company")
                                            ).Distinct().ToList()


                'List of employees that do not have deductions to be included
                Dim empsWithoutDeductionsToInclude = (From emp In sameEmpNumberCompanyTransfers.AsEnumerable()
                                                      Group Join ewd In empsWithDedsToInclude
                                                        On emp.Item("Employee Number") Equals ewd.EmpNo
                                                        Into gl = Group
                                                      From g In gl.DefaultIfEmpty()
                                                      Where g Is Nothing
                                                      Select EmpNo = emp.Item("Employee Number"), EmpName = emp.Item("EmpName"), OldCmpy = emp.Item("Old Company"), NewCmpy = emp.Item("New Company")
                                                     )


                changeMessage = String.Concat("The following employees transferred companies.  Please review the changes.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                If empsWithoutDeductionsToInclude.Count > 0 Then
                    empCounter = 1
                    changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                                "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                                "<th class='tdborder' style='background-color: lightgray;'>EE#</th>" &
                                                "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                                "<th class='tdborder' style='background-color: lightgray;'>Old Company</th>" &
                                                "<th class='tdborder' style='background-color: lightgray;'>New Company</th></tr>")

                    For Each emp In empsWithoutDeductionsToInclude
                        changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpNo.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpName.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.OldCmpy.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.NewCmpy.ToString & "</td></tr>")
                        empCounter += 1
                    Next

                    ' close the above table
                    changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)


                End If

                ' Employees That have the acceptable Deductions
                If empsWithDedsToInclude.Count > 0 Then
                    empCounter = 1

                    For Each emp In empsWithDedsToInclude
                        If Not changeMessage.Contains("<style") Then
                            changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader"))
                        Else
                            changeMessage = String.Concat(changeMessage, "<table class='tbl' cellspacing='0'><tr>")
                        End If

                        changeMessage = String.Concat(changeMessage, "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                                                     "<th class='tdborder' style='background-color: lightgray;'>EE#</th>" &
                                                                     "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                                                     "<th class='tdborder' style='background-color: lightgray;'>Old Company</th>" &
                                                                     "<th class='tdborder' style='background-color: lightgray;'>New Company</th>")
                        ' add the filler
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td></tr>")
                        'add the employee info
                        changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpNo.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpName.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.OldCmpy.ToString & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.NewCmpy.ToString & "</td>")
                        ' add the filler
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td>")
                        changeMessage = String.Concat(changeMessage, "<td style='border-width: 1px; border-style: none none solid none; background-color: lightgray;'>" & "&nbsp;" & "</td></tr>")


                        Dim acDeductions = From acd In acceptableDeductionsFromToCompanyTransfers
                                           Where acd.EmployeeNumber.Equals(emp.EmpNo)

                        If acDeductions.Count > 0 Then
                            'add the header for the deductions
                            changeMessage = String.Concat(changeMessage,
                                                      "<tr><th class='tdborder' style='background-color: lightgray;'>Company</th><th class='tdborder' style='background-color: lightgray;'>Ded. Cd.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Amt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Start Dt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Ded. Stop Dt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Goal Amt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>GTD Amt.</th>" &
                                                      "<th class='tdborder' style='background-color: lightgray;'>Goal Bal.</th>" &
                                                      "</tr>")

                        End If

                        For Each ded In acDeductions
                            ' Look into pulling the to company from the general deductions collection.
                            Dim toDed As GeneralDeduction = generalDeductionsCurrentAllCollection.
                                                            Where(Function(w) w.DEDCD.Equals(ded.DEDCD) And w.EmployeeNumber.Equals(ded.EmployeeNumber) And w.PAYGROUP.Equals(ded.PAYGROUP)).
                                                            Select(Function(s) s).FirstOrDefault

                            Dim wrkDed As GeneralDeduction = Nothing

                            If toDed Is Nothing Then
                                wrkDed = ded
                            Else
                                wrkDed = toDed
                            End If

                            'add the deduction info
                            changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'>" & wrkDed.PAYGROUP.ToString & "</td>")
                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & wrkDed.DEDCD.ToString & "</td>") 'Ded code


                            Dim formattedDedAmt As String = String.Empty
                            If wrkDed.DED_ADDL_AMT > 0.0 Then
                                formattedDedAmt = wrkDed.DED_ADDL_AMT.ToString("#####0.00")
                            ElseIf wrkDed.DED_RATE_PCT > 0.0 Then
                                formattedDedAmt = wrkDed.DED_RATE_PCT.ToString("P")
                            End If

                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & formattedDedAmt & "</td>") 'Ded Amt



                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & wrkDed.START_DT & "</td>") 'Ded Start dt

                            If wrkDed.END_DT = #1/1/2200 12:00:00 AM# Then
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & "&nbsp" & "</td>") 'Ded Stop Dt
                            Else
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & wrkDed.END_DT & "</td>") 'Ded Stop Dt
                            End If

                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & wrkDed.GOAL_AMT.ToString("#####0.00") & "</td>") 'Goal Amt
                            changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & wrkDed.GTD_Amt.ToString("#####0.00") & "</td>") 'Goal YTD Amt

                            Dim goalBalance As Decimal = wrkDed.GOAL_AMT - wrkDed.GTD_Amt
                            If goalBalance <= 0 Then
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & "0.00" & "</td></tr>") 'Goal balance
                            Else
                                changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & goalBalance.ToString("#####0.00") & "</td></tr>") 'Goal balance
                            End If


                        Next
                        'close the table
                        'changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)
                        changeMessage = String.Concat(changeMessage, "</table><br/>", System.Environment.NewLine)

                    Next
                End If
                ' Send the same employee number company transfer change notification
                DataManager.SendStandardNotification("SAMEEMPNUMTRANS", changeMessage)
            End If

            ' ***** Get the recognized employees that have had tax changes *****
            employeesWithChangedTaxData = DataManager.RemoveEmployeesWhoHasDataErrorsFromDatatable(DataManager.GetEmployeesWithTaxChanges(runID))

            If employeesWithChangedTaxData.Rows.Count > 0 Then
                changeMessage = String.Concat("The following employees have had tax changes sent to ADP and currently has manual updates to their local tax setup on ADP.  Please review the changes and determine if local tax setup needs to be updated.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                ' Add each employee to the list
                'For Each row As DataRow In employeesWithChangedTaxData.Rows
                '    changeMessage = String.Concat(changeMessage, row.Item("Name").ToString, " (", row.Item("Employee Number").ToString, ") - ", _
                '              row.Item("Company").ToString, Environment.NewLine, Environment.NewLine)
                'Next

                '' Send the tax change notification
                'DataManager.SendStandardNotification("TAXCHANGE", changeMessage)

                empCounter = 1
                'changeMessage = String.Empty
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>Employee Number</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                For Each row As DataRow In employeesWithChangedTaxData.Rows
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Employee Number").ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Name").ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Company").ToString & "</td></tr>")
                    empCounter += 1
                Next
                ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                ' Send the tax change notification
                DataManager.SendStandardNotification("TAXCHANGE", changeMessage)
            End If


            ' ***** Get the recognized employees that have had changes to specific deductions *****
            employeesWithChangedDeductions = DataManager.RemoveEmployeesWhoHasDataErrorsFromDatatable(DataManager.GetSpecificDeductionChanges(runID))

            If employeesWithChangedDeductions.Rows.Count > 0 Then
                changeMessage = String.Concat("The following employees have had deduction changes sent to ADP.  Please review the changes.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                ' Add each employee to the list
                'For Each row As DataRow In employeesWithChangedDeductions.Rows
                '    changeMessage = String.Concat(changeMessage, "EE#: ", row.Item("Employee Number").ToString, ", Company: ", row.Item("Company").ToString, _
                '              ", Deduction: ", row.Item("Deduction").ToString, Environment.NewLine, Environment.NewLine)
                'Next

                '' Send the tax change notification
                'DataManager.SendStandardNotification("DEDCHANGE", changeMessage)

                empCounter = 1
                'changeMessage = String.Empty
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>EE#</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th><th class='tdborder'>Deduction</th></tr>")
                For Each row As DataRow In employeesWithChangedDeductions.Rows
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Employee Number").ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Name").ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Company").ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & row.Item("Deduction").ToString & "</td></tr>")
                    empCounter += 1
                Next
                ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                ' Send the tax change notification
                DataManager.SendStandardNotification("DEDCHANGE", changeMessage)
            End If

            ' ***** Get employees  have changes in W4 Locked flag *****
            ' Get records that have changed
            Dim employeesWithChangedW4Lock = From txc In taxDataCurrentCollection
                                             Join txb In taxDataBeforeCollection
                                             On txc.EMPLID Equals txb.EMPLID And txc.PAYGROUP Equals txb.PAYGROUP
                                             Join pdc In personalDataCurrentCollection
                                             On pdc.EMPLID Equals txc.EMPLID
                                             Join jcc In jobCurrentCollection
                                             On jcc.EMPLID Equals txc.EMPLID
                                             Where txc.W4IsLocked <> txb.W4IsLocked
                                             Select txc, txb, pdc, jcc

            If employeesWithChangedW4Lock.Count > 0 Then
                changeMessage = String.Concat("The following employees have had changes to their W4 lock status. Please review the changes make necessary modifications in Vantage.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)
                empCounter = 1
                'changeMessage = String.Empty
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>Employee ID</th><th class='tdborder'>EE#</th><th class='tdborder'>First Name</th><th class='tdborder'>Last Name</th><th class='tdborder'>Company Code</th><th class='tdborder'>Old flag</th><th class='tdborder'>New flag</th></tr>")
                For Each tx In employeesWithChangedW4Lock
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.txb.EMPLID.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.jcc.FILE_NBR.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.pdc.FIRST_NAME.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.pdc.LAST_NAME.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.jcc.COMPANY.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.txb.W4IsLocked.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.txc.W4IsLocked.ToString & "</td></tr>")
                    empCounter += 1
                Next
                ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                ' Send the tax change notification
                DataManager.SendStandardNotification("W4LOCKCHANGE", changeMessage)
            End If

            ' Get records that have changed
            Dim changedPersonalData = From pdc In personalDataCurrentCollection
                                      Join pdb In personalDataBeforeCollection
                                        On pdc.EMPLID Equals pdb.EMPLID And pdc.PAYGROUP Equals pdb.PAYGROUP
                                      Join jcc In jobCurrentCollection
                                             On jcc.EMPLID Equals pdc.EMPLID
                                      Where
                                        pdc.SSN <> pdb.SSN
                                      Select pdc, pdb, jcc

            If changedPersonalData.Count > 0 Then
                changeMessage = String.Concat("The following employees have had SSN changes sent to Vantage. Please update it manually in Vantage then rerun the employee in the data bridge. <br/> Make the changes via Employee ID maintenance (by going to Vantage>People>Person Information>Employee ID Maintenance) <br/><br/>",
                          Environment.NewLine, Environment.NewLine)
                empCounter = 1
                'changeMessage = String.Empty
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>Employee ID</th><th class='tdborder'>EE#</th><th class='tdborder'>Employee Name</th><th class='tdborder'>Company Code</th><th class='tdborder'>Old SSN</th><th class='tdborder'>New SSN</th></tr>")
                For Each tx In changedPersonalData
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.pdc.EMPLID.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.jcc.FILE_NBR.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & String.Concat(tx.pdc.FIRST_NAME.ToString, " ", tx.pdc.LAST_NAME) & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.jcc.COMPANY.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & String.Concat("*****", tx.pdb.SSN.ToString.Substring(Math.Max(0, tx.pdb.SSN.ToString.Length - 4))) & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & String.Concat("*****", tx.pdc.SSN.ToString.Substring(Math.Max(0, tx.pdc.SSN.ToString.Length - 4))) & "</td></tr>")
                    empCounter += 1
                Next
                ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                ' Send the tax change notification
                DataManager.SendStandardNotification("SSNCHANGE", changeMessage)
            End If


            ' Get records that have changed
            Dim changedWACRFEE_TAX = From tdc In taxDataCurrentCollection
                                     Join tdb In taxDataBeforeCollection
                                        On tdc.EMPLID Equals tdb.EMPLID And tdc.PAYGROUP Equals tdb.PAYGROUP
                                     Join pdc In personalDataCurrentCollection
                                         On pdc.EMPLID Equals tdc.EMPLID
                                     Join jcc In jobCurrentCollection
                                             On jcc.EMPLID Equals pdc.EMPLID
                                     Where
                                        tdc.Long_Term_Care_Ins_Status <> tdb.Long_Term_Care_Ins_Status
                                     Select tdc, tdb, pdc, jcc

            If changedWACRFEE_TAX.Count > 0 Then
                changeMessage = String.Concat("The following salaried employees have had changes to their WA cares exemption status. Please review the changes. <br/><br/>",
                          Environment.NewLine, Environment.NewLine)
                empCounter = 1
                'changeMessage = String.Empty
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>Employee ID</th><th class='tdborder'>EE#</th><th class='tdborder'>First Name</th> <th class='tdborder'>Last Name</th> <th class='tdborder'>Company Code</th><th class='tdborder'>Old flag</th><th class='tdborder'>New flag</th></tr>")
                For Each tx In changedWACRFEE_TAX
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.pdc.EMPLID.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.jcc.FILE_NBR.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.pdc.FIRST_NAME.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.pdc.LAST_NAME & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.jcc.COMPANY.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.tdb.Long_Term_Care_Ins_Status.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & tx.tdc.Long_Term_Care_Ins_Status.ToString & "</td></tr>")
                    empCounter += 1
                Next
                ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                ' Send the tax change notification
                DataManager.SendStandardNotification("WACareCHANGE", changeMessage)
            End If

            'Send Email Notification when employee transfer companies and State Change in WorkedIn or Residence tax code change
            SendCompanyTransferStateChangeNotification()

            'Send Email notification when the given employee's(Remote Worker) have changed their Primary location or the Local Residence Tax codes.
            SendLocalTaxCodeChangeNotification()

            SendNewLocalTaxCodeNotification()

            SendRequiredGoalAmountMissingNotification()

            SendGoalAmountAndBalanceChangeNotifications()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to send notification if required goal amount is zero.
    ''' </summary>
    Private Sub SendRequiredGoalAmountMissingNotification()
        Dim changeMessage As String = String.Empty
        Dim localTaxCode1 As String = String.Empty
        Dim localTaxCode2 As String = String.Empty
        Dim localTaxCode4 As String = String.Empty
        Dim taxCodes As String = String.Empty
        Dim empCounter As Int32 = 0
        Try

            changeMessage = String.Concat("There is no goal amount for the deduction below.  This deduction requires a goal amount to successfully send to ADP.  Enter a goal amount in UKG for this employee and deduction code so changes will flow over on the next feed to ADP.<br/><br/> These employees will not feed over to ADP until the goal amount has been added. <br/>",
                          Environment.NewLine, Environment.NewLine)

            If goalAmountMissingOnDeductions.Count > 0 Then


                empCounter = 1
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                            "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Company Code</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Employee Name</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Employee Number</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Deduction Code</th></tr>")

                For Each goalAmountMissingOnDeduction In goalAmountMissingOnDeductions


                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & goalAmountMissingOnDeduction.PAYGROUP & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & goalAmountMissingOnDeduction.EmployeeName & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & goalAmountMissingOnDeduction.EmployeeNumber & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & goalAmountMissingOnDeduction.DEDCD & "</td></tr>")
                    empCounter += 1
                Next

                ' close the above table
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)


                DataManager.SendStandardNotification("MISSINGGOALAMT", changeMessage)
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' This method is used to send notification if New local tax code is identified.
    ''' </summary>
    Private Sub SendNewLocalTaxCodeNotification()
        Dim changeMessage As String = String.Empty
        Dim localTaxCode1 As String = String.Empty
        Dim localTaxCode2 As String = String.Empty
        Dim localTaxCode4 As String = String.Empty
        Dim taxCodes As String = String.Empty
        Dim empCounter As Int32 = 0
        Try

            changeMessage = String.Concat("A new Local Tax Code has been detected. The new code should be added to the cross-reference table located in AshleyNet with the corresponding ADP code. You may also need to setup this new code on ADP prior to the ADPC feed being sent.<br/><br/> These employees will not feed over to ADP until the new code has been added to the cross-reference table. <br/>",
                          Environment.NewLine, Environment.NewLine)

            If erroredTaxDataCurrentCollection.Count > 0 Then


                empCounter = 1
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                            "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Company Code</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Employee Name</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Employee Number</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Missing Local Tax Code</th></tr>")

                For Each emp In erroredTaxDataCurrentCollection

                    localTaxCode1 = If(emp.LOCAL_TAX_CD.Contains("_IS_INVALID_TAX_CODE"), emp.LOCAL_TAX_CD.Replace("_IS_INVALID_TAX_CODE", String.Empty), String.Empty)
                    localTaxCode2 = If(emp.LOCAL2_TAX_CD.Contains("_IS_INVALID_TAX_CODE"), emp.LOCAL2_TAX_CD.Replace("_IS_INVALID_TAX_CODE", String.Empty), String.Empty)
                    localTaxCode4 = If(emp.LOCAL4_TAX_CD.Contains("_IS_INVALID_TAX_CODE"), emp.LOCAL4_TAX_CD.Replace("_IS_INVALID_TAX_CODE", String.Empty), String.Empty)

                    taxCodes = String.Join(", ", {localTaxCode1, localTaxCode2, localTaxCode4}.Where(Function(s) Not String.IsNullOrEmpty(s)).Select(Function(s) s.Trim()) _
                             .Distinct())

                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.PAYGROUP.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmployeeName.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EMP_No.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & taxCodes.ToString & "</td></tr>")
                    empCounter += 1
                Next

                ' close the above table
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)


                DataManager.SendStandardNotification("NEWLOCTAXCD", changeMessage)
            End If


        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SendCompanyTransferStateChangeNotification()

        Dim changeMessage As String = String.Empty
        Dim empCounter As Int32 = 0
        Try
            ' Get records that have changed
            Dim companyChangedData = DataManager.GetEmpsCompanyTransfers()

            ' Get records that have changed
            Dim empsWithCompStateChange = (From pdc In companyChangedData.AsEnumerable()
                                           Join txc In taxDataCurrentCollection
                                             On pdc.Item("EecEEID") Equals String.Concat(txc.EMPLID, "0") And pdc.Item("New Company") Equals txc.PAYGROUP
                                           Join txb In taxDataBeforeCollection
                                             On txc.EMPLID Equals txb.EMPLID 'And txc.PAYGROUP Equals txb.PAYGROUP
                                           Where txc.STATE_TAX_CD <> txb.STATE_TAX_CD Or txc.STATE2_TAX_CD <> txb.STATE2_TAX_CD
                                           Select newWorkedIn = txc.STATE_TAX_CD, newResidenceIn = txc.STATE2_TAX_CD, oldWorkedIn = txb.STATE_TAX_CD, OldResidenceIn = txb.STATE2_TAX_CD,
                                                 EmpNo = pdc.Item("Employee Number"), EmpName = pdc.Item("EmpName"), HireDate = pdc.Item("Hire Date"), OldCompany = pdc.Item("Old Company"), NewCompany = pdc.Item("New Company")).Distinct()

            changeMessage = String.Concat("The following employees who had a company transfers and employee's state for 'Worked in' or 'Residence' tax changes. Due to known Vantage Data Bridge defect, please review that the marital status and the new state tax record updated correctly along with all other changes.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

            If empsWithCompStateChange.Count > 0 Then
                empCounter = 1
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                            "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>EE#</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Hire Date</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Old Company</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>New Company</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Old Worked In Tax</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>New Worked In Tax</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Old Resident Tax</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>New Resident Tax</th></tr>")

                For Each emp In empsWithCompStateChange
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpNo.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpName.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.HireDate.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.OldCompany.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.NewCompany.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.oldWorkedIn.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.newWorkedIn.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.OldResidenceIn.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.newResidenceIn.ToString & "</td></tr>")
                    empCounter += 1
                Next

                ' close the above table
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)


                DataManager.SendStandardNotification("COMPTRNSSTCHG", changeMessage)
            End If


        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub SendLocalTaxCodeChangeNotification()

        Dim changeMessage As String = String.Empty
        Dim empCounter As Int32 = 0
        Try

            Dim lstRemoteWorking = (From rmt In PayControlValueHandler.GetPayControlValuesByKey(payControlValuesCollection, "REMOTE_WRK_EMP")
                                    Select EmpNo = rmt.Value1, Company = rmt.Value2).Distinct()

            ' Get records that have changed Local Residence Tax Code or Work-In location
            Dim empsWithLocalTaxChange = (From txc In taxDataCurrentCollection
                                          Join txb In taxDataBeforeCollection
                                             On txc.EMPLID Equals txb.EMPLID And txc.PAYGROUP Equals txb.PAYGROUP
                                          Join jcc In jobCurrentCollection
                                             On txc.EMPLID Equals jcc.EMPLID And txc.PAYGROUP Equals jcc.PAYGROUP
                                          Join jbc In jobBeforeCollection
                                             On jbc.EMPLID Equals jcc.EMPLID And jbc.PAYGROUP Equals jcc.PAYGROUP
                                          Join rmt In lstRemoteWorking
                                             On txc.EMP_No Equals rmt.EmpNo
                                          Where txc.LOCAL2_TAX_CD <> txb.LOCAL2_TAX_CD Or jcc.LOCATION <> jbc.LOCATION
                                          Select EmpNo = txc.EMP_No, EMPLID = txc.EMPLID, PAYGROUP = txc.PAYGROUP,
                                                 NewResidenceIn = txc.LOCAL2_TAX_CD, OldResidenceIn = txb.LOCAL2_TAX_CD,
                                                 OldWorkinLocation = jbc.LOCATION, NewWorkInLocation = jcc.LOCATION).Distinct()

            Dim empsWithLocalTaxCodeChange = (From ltc In empsWithLocalTaxChange
                                              Join per In personalDataCurrentCollection
                                                 On ltc.EMPLID Equals per.EMPLID And ltc.PAYGROUP Equals per.PAYGROUP
                                              Select EmpNo = ltc.EmpNo,
                                                     EmpName = String.Concat(per.LAST_NAME, ",", per.FIRST_NAME),
                                                     OldWorkInLocation = IIf(String.IsNullOrWhiteSpace(ltc.OldWorkinLocation), "-", ltc.OldWorkinLocation),
                                                     NewWorkInLocation = IIf(String.IsNullOrWhiteSpace(ltc.NewWorkInLocation), "-", ltc.NewWorkInLocation),
                                                     OldResidenceIn = IIf(String.IsNullOrWhiteSpace(ltc.OldResidenceIn), "-", ltc.OldResidenceIn),
                                                     NewResidenceIn = IIf(String.IsNullOrWhiteSpace(ltc.NewResidenceIn), "-", ltc.NewResidenceIn)).Distinct()

            changeMessage = String.Concat("This is a remote worker who had a different tax setup in ADP than how it was displayed in UKG previously and could impact local tax codes, now residence tax codes have changed as mentioned below. Please review/manually update work-in local tax information in ADP if necessary.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

            If empsWithLocalTaxCodeChange.Count > 0 Then
                empCounter = 1
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                                            "<th class='tdborder' style='background-color: lightgray;'>S.No</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>EE#</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Name</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Old Work-In Location</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>New Work-In Location</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>Old Local Resident Tax Code</th>" &
                                            "<th class='tdborder' style='background-color: lightgray;'>New Local Resident Tax Code</th></tr>")

                For Each emp In empsWithLocalTaxCodeChange
                    changeMessage = String.Concat(changeMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpNo.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.EmpName.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.OldWorkInLocation.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.NewWorkInLocation.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.OldResidenceIn.ToString & "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>" & emp.NewResidenceIn.ToString & "</td></tr>")
                    empCounter += 1
                Next

                ' close the above table
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)


                DataManager.SendStandardNotification("LOCTAXCDCHG", changeMessage)
            End If


        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' This method is used to send notification if an employee's deduction goal amount or goal balance has changed.
    ''' </summary>
    Public Sub SendGoalAmountAndBalanceChangeNotifications()
        Dim changeMessage As String = String.Empty
        Dim empCounter As Int32 = 0
        Try
            ' Find deductions where GOAL_AMT or (GOAL_AMT - GTD_Amt) has changed
            Dim goalAmtChanges = From curr In generalDeductionsCurrentCollection
                                 Join prev In generalDeductionsBeforeCollection
                                 On curr.EMPLID Equals prev.EMPLID And curr.PAYGROUP Equals prev.PAYGROUP And curr.DEDCD Equals prev.DEDCD
                                 Where curr.GOAL_AMT <> prev.GOAL_AMT OrElse
                                   (curr.GOAL_AMT - curr.GTD_Amt) <> (prev.GOAL_AMT - prev.GTD_Amt)
                                 Select curr, prev

            If goalAmtChanges.Any() Then
                changeMessage = String.Concat("The following employees had changes in the 401LP Deduction Goal Amounts.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                empCounter = 1
                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblheader") &
                "<th class='tdborder' style='background-color: lightgray;'>Employee Name</th>" &
                "<th class='tdborder' style='background-color: lightgray;'>Employee Number</th>" &
                "<th class='tdborder' style='background-color: lightgray;'>Previous Total Goal</th>" &
                "<th class='tdborder' style='background-color: lightgray;'>New Total Goal</th>" &
                "<th class='tdborder' style='background-color: lightgray;'>Previous Goal Balance</th>" &
                "<th class='tdborder' style='background-color: lightgray;'>New Goal Balance</th></tr>")

                For Each change In goalAmtChanges
                    Dim oldGoalAmt As Decimal = change.prev.GOAL_AMT
                    Dim newGoalAmt As Decimal = change.curr.GOAL_AMT
                    Dim oldGoalBal As Decimal = change.prev.GOAL_AMT - change.prev.GTD_Amt
                    Dim newGoalBal As Decimal = change.curr.GOAL_AMT - change.curr.GTD_Amt

                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>", change.curr.EmployeeName, "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>", change.curr.EmployeeNumber, "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>", oldGoalAmt.ToString("#####0.00"), "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>", newGoalAmt.ToString("#####0.00"), "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>", oldGoalBal.ToString("#####0.00"), "</td>")
                    changeMessage = String.Concat(changeMessage, "<td class='tdborder'>", newGoalBal.ToString("#####0.00"), "</td></tr>")
                    empCounter += 1
                Next

                changeMessage = String.Concat(changeMessage, GetCssForEmailNotifications("tblfooter"), System.Environment.NewLine)
                DataManager.SendStandardNotification("GOALAMTCHG", changeMessage)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ' Notifies users for any skipped deductions that are not accounted for
    Private Sub SendUnaccountedForSkippedDeductions()
        Dim emailMessage As String = String.Empty
        Dim empCounter As Int32 = 0
        Try

            ' Get the skipped deductions that are unaccounted for
            Dim unaccountedForSkippedDed = (From dec In generalDeductionsCurrentCollection
                                            Where dec.SkipDeduction = "Y" And
                                               dec.SkipDeductionAccFor = "N"
                                            Select dec.EmployeeNumber, dec.EmployeeName, dec.PAYGROUP).Distinct()

            ' Check if unaccounted for skipped deductions exist
            If unaccountedForSkippedDed.Count > 0 Then
                emailMessage = String.Concat("The following employees contained an EX* deduction which was Not accounted for. The EX* deductions were skipped over And Not sent to ADP; please review.<br/><br/>",
                          Environment.NewLine, Environment.NewLine)

                ' Iterate through and form email message
                'For Each emp In unaccountedForSkippedDed
                '    emailMessage = String.Concat(emailMessage, "EE#:  ", emp.EmployeeNumber, ", Company: ", emp.PAYGROUP, _
                '                   Environment.NewLine, Environment.NewLine)
                'Next

                '' Send the notification
                'DataManager.SendStandardNotification("DEDSKIP", emailMessage)

                empCounter = 1
                'emailMessage = String.Empty
                emailMessage = String.Concat(emailMessage, GetCssForEmailNotifications("tblheader") &
                                   "<th class='tdborder'>S.No</th><th class='tdborder'>EE#</th><th class='tdborder'>Name</th><th class='tdborder'>Company</th></tr>")
                For Each emp In unaccountedForSkippedDed
                    emailMessage = String.Concat(emailMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                    emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & emp.EmployeeNumber & "</td>")
                    emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & emp.EmployeeName.ToString & "</td>")
                    emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & emp.PAYGROUP & "</td></tr>")
                    empCounter += 1
                Next
                ultiproDataValidationError = String.Concat(ultiproDataValidationError, GetCssForEmailNotifications("tblfooter"))
                ' Send the notification
                DataManager.SendStandardNotification("DEDSKIP", emailMessage)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Sends VIP data change notifications to the defined recipients 
    Private Sub SendVIPChangeNotifications(vipDetails As IEnumerable(Of VIPData), changedVIPPersonalData As IEnumerable(Of VIPData),
                                           changedVIPEmploymentData As IEnumerable(Of VIPData), changedVIPJobData As IEnumerable(Of VIPData),
                                           changedVIPDeductionData As IEnumerable(Of VIPData), changedVIPDirectDepositData As IEnumerable(Of VIPData),
                                           changedVIPTaxData As IEnumerable(Of VIPData), changedVIPW4Data As IEnumerable(Of VIPData))
        Try

            Dim emailMessage As String = String.Empty
            Dim empCounter = 1

            If vipDetails.Count > 0 Then
                For Each vipData As VIPData In vipDetails
                    If changedVIPPersonalData.Count > 0 Then
                        'Get changed file names of each VIP employee in the Personal data list
                        For Each row As VIPData In changedVIPPersonalData
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = "Personal Data.csv"
                            End If
                        Next
                    End If
                    If changedVIPEmploymentData.Count > 0 Then
                        'Get changed file names of each VIP employee in the Employment data list
                        For Each row As VIPData In changedVIPEmploymentData
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = If(vipData.FILE_NAMES = String.Empty, "Employment.csv", String.Concat(vipData.FILE_NAMES, ",", "Employment.csv"))
                            End If
                        Next
                    End If
                    If changedVIPJobData.Count > 0 Then
                        'Get changed file names of each VIP employee in the Job data list
                        For Each row As VIPData In changedVIPJobData
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = If(vipData.FILE_NAMES = String.Empty, "Job.csv", String.Concat(vipData.FILE_NAMES, ",", "Job.csv"))
                            End If
                        Next
                    End If
                    If changedVIPDeductionData.Count > 0 Then
                        'Get changed file names of each VIP employee in the General Deduction data list
                        For Each row As VIPData In changedVIPDeductionData
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = If(vipData.FILE_NAMES = String.Empty, "Deductions.csv", String.Concat(vipData.FILE_NAMES, ",", "Deductions.csv"))
                            End If
                        Next
                    End If
                    If changedVIPDirectDepositData.Count > 0 Then
                        'Get changed file names of each VIP employee in the Direct Deposit data list
                        For Each row As VIPData In changedVIPDirectDepositData
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = If(vipData.FILE_NAMES = String.Empty, "Direct Dep.csv", String.Concat(vipData.FILE_NAMES, ",", "Direct Dep.csv"))
                            End If
                        Next
                    End If
                    If changedVIPTaxData.Count > 0 Then
                        'Get changed file names of each VIP employee in the Tax data list
                        For Each row As VIPData In changedVIPTaxData
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = If(vipData.FILE_NAMES = String.Empty, "Tax Data.csv", String.Concat(vipData.FILE_NAMES, ",", "Tax Data.csv"))
                            End If
                        Next
                    End If
                    If changedVIPW4Data.Count > 0 Then
                        'Get changed file names of each VIP employee in the W4 data list
                        For Each row As VIPData In changedVIPW4Data
                            If (row.EMPLID = vipData.EMPLID) Then
                                vipData.FILE_NAMES = If(vipData.FILE_NAMES = String.Empty, "W4 Data.csv", String.Concat(vipData.FILE_NAMES, ",", "W4 Data.csv"))
                            End If
                        Next
                    End If

                    ' Verify and construct the email body
                    If (Not String.IsNullOrEmpty(vipData.FILE_NAMES)) Then
                        'changeMessage = String.Concat(changeMessage, empCounter, ". Employee ID: ", vipData.EMPLID, ", First Name: ", vipData.FIRST_NAME, _
                        '                ", Last Name: ", vipData.LAST_NAME, " – Changes sent on the following files: {", vipData.FILE_NAMES, "}", _
                        '                Environment.NewLine, Environment.NewLine)

                        emailMessage = String.Concat(emailMessage, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                        emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & vipData.EMPLID & "</td>")
                        emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & vipData.EMPNO & "</td>")
                        emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & vipData.FIRST_NAME & "</td>")
                        emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & vipData.LAST_NAME & "</td>")
                        emailMessage = String.Concat(emailMessage, "<td class='tdborder'>" & vipData.FILE_NAMES & "</td></tr>")

                        empCounter += 1
                    End If
                Next
                ' Send the VIP's change notification to the defined recipients
                If (Not String.IsNullOrEmpty(emailMessage)) Then

                    emailMessage = String.Concat(GetCssForEmailNotifications("tblheader") &
                                                 "<th class='tdborder'>S.No</th><th class='tdborder'>Employee ID</th><th class='tdborder'>EE #</th><th class='tdborder'>First Name</th><th class='tdborder'>Last Name</th><th class='tdborder'>Changes sent on the following files</th></tr>",
                                                 emailMessage,
                                                 GetCssForEmailNotifications("tblfooter"))
                    DataManager.SendStandardNotification("VIPNOTIFICATION", String.Concat("The following employees have had changes sent to ADP.  Please review the changes.<br/><br/>",
                          Environment.NewLine, Environment.NewLine, emailMessage))
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' Returns the css styles of email notification table formats
    Private Function GetCssForEmailNotifications(type As String) As String
        If type.Equals("tblheader") Then

            Return "<style type='text/css'>.tbl{ float:left; margin-top:10px; border-top:1px Solid gray; border-left:1px Solid gray; border-right:1px Solid gray; border-bottom:0px Solid transparent !important; font-size: 15px;font-family: Times New Roman; } " &
                                          ".tdborder,.tderror{ border-right:1px Solid gray; } .tdborder,.trbtmborder,.tderror{ padding:3px !important; border-bottom:1px Solid gray; vertical-align: middle;line-height: 20px;} .tderror{ background-color:yellow; }</style>" &
                                           "<br/><table class='tbl' cellspacing='0'><tr>"
        Else
            Return "</table><br/><br/>"
        End If
    End Function

    ''' <summary>
    ''' Send Terminated data change notification
    ''' </summary>
    ''' <param name="employmentData"></param>
    ''' <param name="terminatedEmployeeCollection"></param>
    Private Function SendTerminatedEmployeeChangeNotification(employmentData As IEnumerable(Of Employment), terminatedEmployeeCollection As IEnumerable(Of TermedEmploymentData)) As IEnumerable(Of TermedEmploymentData)
        Dim terminatedEmpDataChange = String.Empty
        Try
            'Check whether Terminated Employees available in the collection or not
            If terminatedEmployeeCollection.Count > 0 Then
                Dim empCounter = 1
                'Create the table header to be rendered in the notification email
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EE #</th><th class='tdborder'>EMPL ID</th><th class='tdborder'>Employee Name</th>" &
                                       "<th class='tdborder'>Company</th><th class='tdborder'>Pay Frequency</th><th class='tdborder'>Hire Date</th>" &
                                       "<th class='tdborder'>Rehire Date</th><th class='tdborder'>Company Seniority Date</th>" &
                                       "<th class='tdborder'>Business Title</th><th class='tdborder'>Supervisor ID</th><th class='tdborder'>Supervisor Name</th>" &
                                       "<th class='tdborder'>Termination Date</th><th class='trbtmborder'>Last Date Worked</th></tr>")

                ' Get the Terminated Employees Data - changes in both BUSINESS_TITLE and SUPERVISOR_ID 
                Dim termedEmpChng = ((From emc In employmentData
                                      Join emt In terminatedEmployeeCollection
                                            On emc.EMPLID Equals emt.EMPLID And emc.PAYGROUP Equals emt.PAYGROUP
                                      Where emc.BUSINESS_TITLE <> emt.BUSINESS_TITLE And emc.SUPERVISOR_ID <> emt.SUPERVISOR_ID
                                      Select emt).Distinct())

                'Bind the above found employee list in email - changes in both BUSINESS_TITLE and SUPERVISOR_ID 
                If termedEmpChng.Count > 0 Then
                    ' Remove all the Terminated Employees whose data change has been already intimated to business
                    Dim terminatedEmployeeEmailCollection = (From emt In termedEmpChng
                                                             Group Join emth In terminatedEmployeeHistoryCollection
                                                                On emth.EMPLID Equals emt.EMPLID And
                                                             emth.PAYGROUP Equals emt.PAYGROUP Into g = Group
                                                             From emth In g.DefaultIfEmpty()
                                                             Where IsNothing(emth)
                                                             Select emt)
                    terminatedEmpDataChange = BindTermedEmpChangeNotificationBodyText(terminatedEmployeeEmailCollection, terminatedEmpDataChange, "BOTH", empCounter)

                    'Remove the data with both BUSINESS_TITLE and SUPERVISOR_ID changes to avoid duplicates in email 
                    terminatedEmployeeCollection = From emc In terminatedEmployeeCollection
                                                   Group Join emt In termedEmpChng
                                                    On emc.EMPLID Equals emt.EMPLID And emc.PAYGROUP Equals emt.PAYGROUP Into g = Group
                                                   From emt In g.DefaultIfEmpty()
                                                   Where IsNothing(emt)
                                                   Select emc
                End If

                ' Get the Terminated Employees Data - changes only in BUSINESS_TITLE 
                Dim termedEmpBusTitleChng = (From emc In employmentData
                                             Join emt In terminatedEmployeeCollection
                                            On emc.EMPLID Equals emt.EMPLID And emc.PAYGROUP Equals emt.PAYGROUP
                                             Where emc.BUSINESS_TITLE <> emt.BUSINESS_TITLE
                                             Select emt).Distinct()
                'Bind the above found employee list in email - changes only in BUSINESS_TITLE  
                If termedEmpBusTitleChng.Count > 0 Then
                    ' Remove all the Terminated Employees whose data change has been already intimated to business
                    Dim terminatedEmployeeEmailCollection = (From emt In termedEmpBusTitleChng
                                                             Group Join emth In terminatedEmployeeHistoryCollection
                                                                On emth.EMPLID Equals emt.EMPLID And
                                                             emth.PAYGROUP Equals emt.PAYGROUP Into g = Group
                                                             From emth In g.DefaultIfEmpty()
                                                             Where IsNothing(emth)
                                                             Select emt)
                    terminatedEmpDataChange = BindTermedEmpChangeNotificationBodyText(terminatedEmployeeEmailCollection, terminatedEmpDataChange, "BUSINESS_TITLE", empCounter)
                End If

                ' Get the Terminated Employees Data - changes only in SUPERVISOR_ID 
                Dim termedEmpSupIdChng = (From emc In employmentData
                                          Join emt In terminatedEmployeeCollection
                                            On emc.EMPLID Equals emt.EMPLID And emc.PAYGROUP Equals emt.PAYGROUP
                                          Where emc.SUPERVISOR_ID <> emt.SUPERVISOR_ID
                                          Select emt).Distinct()
                'Bind the above found employee list in email - changes only in SUPERVISOR_ID 
                If termedEmpSupIdChng.Count > 0 Then
                    ' Remove all the Terminated Employees whose data change has been already intimated to business
                    Dim terminatedEmployeeEmailCollection = (From emt In termedEmpSupIdChng
                                                             Group Join emth In terminatedEmployeeHistoryCollection
                                                                On emth.EMPLID Equals emt.EMPLID And
                                                             emth.PAYGROUP Equals emt.PAYGROUP Into g = Group
                                                             From emth In g.DefaultIfEmpty()
                                                             Where IsNothing(emth)
                                                             Select emt)
                    terminatedEmpDataChange = BindTermedEmpChangeNotificationBodyText(terminatedEmployeeEmailCollection, terminatedEmpDataChange, "SUPERVISOR_ID", empCounter)
                End If

                terminatedEmployeeCollection = Enumerable.Union(termedEmpChng, Enumerable.Union(termedEmpBusTitleChng, termedEmpSupIdChng))

                'Create the table footer to be rendered in the notification email
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, GetCssForEmailNotifications("tblfooter"))

                'Send email notifications to business

                If empCounter > 1 Then
                    DataManager.SendStandardNotification("TERMEDEMPCHNGDATA", String.Concat("Below is the terminated employee data showing the updated Business title or Supervisor ID or Both. Please review the data and confirm. <br/><br/>", terminatedEmpDataChange))
                End If

            End If
            Return terminatedEmployeeCollection
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Bind Table body for Terminated employees data change
    ''' </summary>
    ''' <param name="terminatedEmployee"></param>
    ''' <param name="terminatedEmpDataChange"></param>
    ''' <param name="cndnType"></param>
    ''' <param name="empCounter"></param>
    ''' <returns></returns>
    Private Function BindTermedEmpChangeNotificationBodyText(terminatedEmployee As IEnumerable(Of TermedEmploymentData), terminatedEmpDataChange As String, cndnType As String, ByRef empCounter As Int32) As String
        Try
            For Each row As TermedEmploymentData In terminatedEmployee
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<tr><td class='tdborder'> " & empCounter.ToString & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.EMPLOYEE_NUMBER & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.EMPLID & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.EMPLOYEE_NAME & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.PAYGROUP & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.PAY_FREQUENCY & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.HIRE_DT & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.REHIRE_DT & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.CMPNY_SENIORITY_DT & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='" & If(cndnType = "BUSINESS_TITLE" Or cndnType = "BOTH", "tderror", "tdborder") & "'>" & row.BUSINESS_TITLE & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='" & If(cndnType = "SUPERVISOR_ID" Or cndnType = "BOTH", "tderror", "tdborder") & "'>" & row.SUPERVISOR_ID & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='" & If(cndnType = "SUPERVISOR_ID" Or cndnType = "BOTH", "tderror", "tdborder") & "'>" & row.SUPERVISOR_NAME & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='tdborder'>" & row.TERMINATION_DT & "</td>")
                terminatedEmpDataChange = String.Concat(terminatedEmpDataChange, "<td class='trbtmborder'>" & row.LAST_DATE_WORKED & "</td></tr>")
                empCounter += 1
            Next
            Return terminatedEmpDataChange
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' Send Missing Employee data notification 
    ''' </summary>
    ''' <param name="newPersDataMissing"></param>
    Private Sub SendMissingEmployeeDataNotification(newPersDataMissing As IEnumerable(Of PersonalData))
        Try
            If newPersDataMissing.Count > 0 Then
                Dim missingEmpInfo = New StringBuilder
                Dim missingEmpCount = 1
                ' Add the data rows
                missingEmpInfo.Append(GetCssForEmailNotifications("tblheader") &
                                       "<th class='tdborder'>S.No</th><th class='tdborder'>EMPLID</th><th class='tdborder'>PAYGROUP</th>" &
                                       "<th class='tdborder'>PAY_FREQUENCY</th><th class='tdborder'>FIRST_NAME</th>" &
                                       "<th class='tdborder'>MIDDLE_NAME</th><th class='tdborder'>LAST_NAME</th>" &
                                       "<th class='tdborder'>SEX</th><th class='tdborder'>BIRTHDATE</th>" &
                                       "<th class='tdborder'>STREET1</th><th class='tdborder'>STREET2</th>" &
                                       "<th class='tdborder'>CITY</th><th class='tdborder'>STATE</th>" &
                                       "<th class='tdborder'>ZIP</th><th class='tdborder'>SSN</th>" &
                                       "<th class='tdborder'>ORIG_HIRE_DT</th><th class='trbtmborder'>HOME_PHONE</th></tr>")
                For Each pd In newPersDataMissing
                    missingEmpInfo.Append("<tr><td class='tdborder'> " & missingEmpCount & "</td>")
                    missingEmpInfo.Append("<td class='tdborder'>" & pd.EMPLID & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.PAYGROUP = "", "tderror", "tdborder") & "'>" & pd.PAYGROUP & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.PAY_FREQUENCY = "", "tderror", "tdborder") & "'>" & pd.PAY_FREQUENCY & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.FIRST_NAME = "", "tderror", "tdborder") & "'>" & pd.FIRST_NAME & "</td>")
                    missingEmpInfo.Append("<td class='tdborder'>" & pd.MIDDLE_NAME & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.LAST_NAME = "", "tderror", "tdborder") & "'>" & pd.LAST_NAME & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.SEX = "", "tderror", "tdborder") & "'>" & pd.SEX & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.BIRTHDATE.ToShortDateString = "1/1/1900", "tderror", "tdborder") & "'>" & If(pd.BIRTHDATE.ToShortDateString = "1/1/1900", "", pd.BIRTHDATE.ToShortDateString) & "</td>") 'String.Concat(Environment.NewLine, Environment.NewLine))
                    missingEmpInfo.Append("<td class='" & If(String.Concat(pd.STREET1, pd.STREET2) = "", "tderror", "tdborder") & "'>" & pd.STREET1 & "</td>")
                    missingEmpInfo.Append("<td class='" & If(String.Concat(pd.STREET1, pd.STREET2) = "", "tderror", "tdborder") & "'>" & pd.STREET2 & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.CITY = "", "tderror", "tdborder") & "'>" & pd.CITY & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.STATE = "", "tderror", "tdborder") & "'>" & pd.STATE & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.ZIP = "", "tderror", "tdborder") & "'>" & pd.ZIP & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.SSN = "", "tderror", "tdborder") & "'>" & pd.SSN & "</td>")
                    missingEmpInfo.Append("<td class='" & If(pd.ORIG_HIRE_DT.ToShortDateString = "1/1/1900", "tderror", "tdborder") & "'>" & If(pd.ORIG_HIRE_DT.ToShortDateString = "1/1/1900", "", pd.ORIG_HIRE_DT.ToShortDateString) & "</td>")
                    missingEmpInfo.Append("<td class='trbtmborder'>" & pd.HOME_PHONE & "</td></tr>")
                    missingEmpCount += 1
                Next
                missingEmpInfo.Append(GetCssForEmailNotifications("tblfooter"))
                'DataManager.SendStandardNotification("EMP_CHNG_EMAIL", String.Concat("Below new employee(s) missing personal information in database. Please review the data.", Environment.NewLine, _
                '                                 Environment.NewLine, missingEmpInfo))
                DataManager.SendStandardNotification("MISSINGEMPLOYEEDATA", String.Concat("The new employee(s) below are missing personal information and have not been included in the ADP files. Please review the data so they may be sent on the next file feed.", Environment.NewLine, Environment.NewLine, Environment.NewLine,
                                                 Environment.NewLine, missingEmpInfo))
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub



    ''' <summary>
    ''' This method populates distinct employee data validation error from different results
    ''' </summary>
    ''' <param name="eecEEID"></param>
    ''' <param name="eecCoID"></param>
    ''' <param name="empCompanyCode"></param>
    ''' <param name="eecEmpNo"></param>
    Private Sub PopulateDataValidationErrors(ByVal eecEEID As String,
                                             ByVal eecCoID As String,
                                             ByVal empCompanyCode As String,
                                             ByVal eecEmpNo As String)


        If Not dataValidationErrors.Where(Function(errors) errors.EEID = eecEEID And errors.COID = eecCoID).Any() Then

            Dim emplId As String = eecEEID.Trim().
                                   Substring(0, eecEEID.Length - 1)

            dataValidationErrors.Add(New DataValidationError() With {
                                       .EEID = eecEEID,
                                       .EmpEEID = emplId,
                                       .COID = eecCoID,
                                       .CompanyCode = empCompanyCode,
                                       .EmployeeNumber = eecEmpNo
                                     })
        End If
    End Sub


    ''' <summary>
    ''' This method populates distinct employee data validation error from different results
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="isColumnNameVaried"></param>
    Private Sub PopulateDataValidationErrors(ByVal row As DataRow, Optional ByVal isColumnNameVaried As Boolean = False)
        Dim eecEEID As String = String.Empty
        Dim eecCoID As String = String.Empty
        Dim empCompanyCode As String = String.Empty
        Dim eecEmpNo As String = String.Empty

        If isColumnNameVaried = False Then
            eecEEID = "EecEEID"
            eecCoID = "EecCoID"
            empCompanyCode = "CmpCompanyCode"
            eecEmpNo = "EecEmpNo"
        Else
            eecEEID = "EecEEID"
            eecCoID = "EecCoID"
            empCompanyCode = "Company"
            eecEmpNo = "Employee Number"
        End If

        eecEEID = row.Item(eecEEID).ToString().Trim()
        eecCoID = row.Item(eecCoID).ToString().Trim()
        empCompanyCode = row.Item(empCompanyCode).ToString().Trim()
        eecEmpNo = row.Item(eecEmpNo).ToString().Trim()

        PopulateDataValidationErrors(eecEEID, eecCoID, empCompanyCode, eecEmpNo)
    End Sub
End Module
