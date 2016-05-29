Imports Dexcom
Imports System.Net.Http
Imports System.Threading
Imports Windows.Devices.Power
Imports System.Text.UTF8Encoding
Imports Windows.Foundation.Metadata
Imports Windows.Networking.Connectivity
Imports System.Runtime.Serialization.Json

''' <summary>
''' Author: Jay Lagorio
''' Date: May 29, 2016
''' Summary: A singleton class that syncs all connected and enrolled devices with Nightscout.
''' </summary>

Public Class Synchronizer
    ' The timer that fires automatically to start the synchronization process.
    Private Shared pSyncTimer As Timer = Nothing

    ' The instance of MainPage with the UI thread that will be used to update controls.
    Private Shared pMainPage As MainPage = Nothing

    ' The ProgressRing shown/hidden depending on whether synchronization is in progress.
    Private Shared pProgressRing As ProgressRing = Nothing

    ' The TextBlock where status messages are displayed.
    Private Shared pStatusTextBlock As TextBlock = Nothing

    ' Indicates whether the synchronization process is running or not.
    Private Shared pSyncRunning As Boolean = False

    ' Allows another thread to cancel the sync operation in 
    ' an Awaitable fashion
    Private Const CancelDelayMilliseconds As Integer = 100
    Private Shared pCancelSync As Boolean = False       ' Cancellation has been requested
    Private Shared pSyncCancelled As Boolean = False    ' The request has been fulfilled

    ' Events indicating that synchronization has started or stopped.
    Public Shared Event SynchronizationStarted()
    Public Shared Event SynchronizationStopped()

    ''' <summary>
    ''' Initializes user controls so they can be updated when synchronization is running
    ''' </summary>
    ''' <param name="MainPage">An instance of MainPage to run user component update operations on</param>
    ''' <param name="ProgressRing">A ProgressRing to show or hide depending on sync status</param>
    ''' <param name="StatusTextBlock">A TextBlock to update when different stages of sync are reached</param>
    ''' <param name="LastSyncTime">A DateTime representing the last time a sync attempt was made</param>
    Shared Sub Initialize(MainPage As MainPage, ProgressRing As ProgressRing, StatusTextBlock As TextBlock, ByVal LastSyncTime As DateTime)
        pMainPage = MainPage
        pProgressRing = ProgressRing
        pStatusTextBlock = StatusTextBlock
        If LastSyncTime = DateTime.MinValue Then
            pStatusTextBlock.Text = "Last sync time: Never"
        Else
            pStatusTextBlock.Text = "Last sync time: " & LastSyncTime.ToString
        End If
    End Sub

    ''' <summary>
    ''' Creates a timer that fires at Settings.UploadInterval in minutes
    ''' </summary>
    Public Shared Sub StartTimer()
        If Not pSyncTimer Is Nothing Then
            Call pSyncTimer.Dispose()
            pSyncTimer = Nothing
        End If

        pSyncTimer = New Timer(AddressOf Synchronize, Nothing, Settings.UploadInterval * 60 * 1000, Settings.UploadInterval * 60 * 1000)
    End Sub

    ''' <summary>
    ''' Destroys the timer if one has been created.
    ''' </summary>
    Public Shared Sub StopTimer()
        If Not pSyncTimer Is Nothing Then
            Call pSyncTimer.Dispose()
            pSyncTimer = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Cancels the synchronization process.
    ''' </summary>
    ''' <returns>True when the synchronization process has been successfully stopped, False otherwise</returns>
    Public Shared Async Function CancelSync() As Task(Of Boolean)
        ' Indicate that a cancellation has been requested
        Call Volatile.Write(pCancelSync, True)

        Call UpdateSyncStatus("Cancelling Synchronization...")

        ' Wait to see whether the cancellation has succeeded
        While Not Volatile.Read(pSyncCancelled)
            Await Task.Yield()
            Await Task.Delay(CancelDelayMilliseconds)
        End While

        ' Reset all of the synchronization and cancellation state variables
        Call Volatile.Write(pSyncRunning, False)
        Call Volatile.Write(pSyncCancelled, False)
        Call Volatile.Write(pCancelSync, False)
        Return True
    End Function

    ''' <summary>
    ''' Returns whether a cancellation has been requested.
    ''' </summary>
    ''' <returns>True if a cancellation has been requested, False otherwise</returns>
    Private Shared Function IsSyncCancelled() As Boolean
        Return Volatile.Read(pCancelSync)
    End Function

    ''' <summary>
    ''' Indicates that cancelling the sync process has completed.
    ''' </summary>
    Private Shared Sub ConfirmSyncCancelled()
        Call Volatile.Write(pSyncCancelled, True)
    End Sub

    ''' <summary>
    ''' Resets the status of synchronization to False.
    ''' </summary>
    Public Shared Sub ResetSyncStatus()
        Call Volatile.Write(pSyncRunning, False)
    End Sub

    ''' <summary>
    ''' Indicates whether a synchronization process is currently running
    ''' </summary>
    ''' <returns>True if synchronizing, False otherwise</returns>
    Public Shared ReadOnly Property IsSynchronizing() As Boolean
        Get
            Return Volatile.Read(pSyncRunning)
        End Get
    End Property

    ''' <summary>
    ''' Updates the synchronization status message on the MainPage.
    ''' </summary>
    ''' <param name="StatusMessage">The message to display on the MainPage</param>
    Private Shared Sub UpdateSyncStatus(ByVal StatusMessage As String)
        Call UpdateSyncStatus(True, StatusMessage)
    End Sub

    ''' <summary>
    ''' Updates the controls on MainPage that indicate the synchronization status.
    ''' </summary>
    ''' <param name="SyncInProgress">True if synchronization is in progress, False otherwise</param>
    ''' <param name="StatusMessage">The status message to set, if any</param>
    Private Shared Async Sub UpdateSyncStatus(ByVal SyncInProgress As Boolean, ByVal StatusMessage As String)
        ' Run the following code on the UI thread to indicate changes in synchronization status.
        Await pMainPage.Dispatcher.RunAsync _
        (Windows.UI.Core.CoreDispatcherPriority.Normal, Sub()
                                                            If (Not pProgressRing Is Nothing) And (Not pStatusTextBlock Is Nothing) Then
                                                                If SyncInProgress Then
                                                                    pProgressRing.Visibility = Visibility.Visible
                                                                Else
                                                                    pProgressRing.Visibility = Visibility.Collapsed
                                                                End If
                                                                pProgressRing.IsActive = SyncInProgress
                                                                pStatusTextBlock.Text = StatusMessage
                                                            End If
                                                        End Sub)
    End Sub

    ''' <summary>
    ''' Thread procedure that runs the synchronization process.
    ''' </summary>
    Public Shared Async Sub Synchronize()
        ' Ensure the synchronization process doesn't run concurrently if the 
        ' sync process takes a long time and runs over.
        Dim ContinueSync As Boolean = False
        If Not Volatile.Read(pSyncRunning) Then
            ' Check Nightscout API status, can't sync without a key or an address
            If Settings.NightscoutAPIKey <> "" And Settings.NightscoutURL <> "" Then
                ' Check to see if we have devices to sync
                If Settings.EnrolledDevices.Count > 0 Then
                    ' Check Internet connectivity
                    Dim NetworkConnection As ConnectionProfile = NetworkInformation.GetInternetConnectionProfile()
                    If Not NetworkConnection Is Nothing Then
                        If NetworkConnection.GetNetworkConnectivityLevel() = NetworkConnectivityLevel.InternetAccess Then
                            ' Indicate the sync process should continue, set the flag to stop additional sync events, and stop the timer
                            ContinueSync = True
                            Call Volatile.Write(pSyncRunning, True)
                            Call StopTimer()
                        End If
                    End If
                End If
            End If
        Else
            Return
        End If

        ' If every condition above isn't acceptable then don't continue
        If Not ContinueSync Then
            Return
        End If

        ' Notify clients that a sync attempt has started
        RaiseEvent SynchronizationStarted()

        ' Attempt to get the last time Nightscout received an entry
        Call UpdateSyncStatus(True, "Connecting to Nightscout...")
        Dim LastNightscoutEntryTime As DateTime = Await GetLastNightscoutSyncTime()

        ' Get all records from each device one by one and aggregate them together
        Dim NightscoutRecords As New Collection(Of NightscoutEntry)
        Dim GlucoseEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutGlucoseEntry()))
        For i = 0 To Settings.EnrolledDevices.Count - 1
            ' If the device has never been sync'd go back only as far as an hour
            If (Settings.EnrolledDevices(i).LastSyncTime = DateTime.MinValue) Or (LastNightscoutEntryTime = DateTime.MinValue) Then
                Settings.EnrolledDevices(i).LastSyncTime = DateTime.Now.Subtract(New TimeSpan(1, 0, 0, 0))
            Else
                ' If the device's last sync time is after the most recent entry in Nightscout then go from the Nightscout time
                If Settings.EnrolledDevices(i).LastSyncTime > LastNightscoutEntryTime Then
                    Settings.EnrolledDevices(i).LastSyncTime = LastNightscoutEntryTime
                End If
            End If

            ' Read each device's records
            Call UpdateSyncStatus("Synchronizing " & Settings.EnrolledDevices(i).Manufacturer & " " & Settings.EnrolledDevices(i).Model & " (" & Settings.EnrolledDevices(i).SerialNumber & ")...")
            Dim DeviceDownloadStartTime As DateTime = DateTime.Now
            If Settings.EnrolledDevices(i).GetType Is GetType(DexcomDevice) Then
                NightscoutRecords = Await ReadDexcomDeviceData(Settings.EnrolledDevices(i), Settings.EnrolledDevices(i).LastSyncTime, NightscoutRecords)
            End If

            ' If a sync in progress was cancelled don't move onto the next device
            ' and don't set this device's last sync time
            If IsSyncCancelled() Then
                Exit For
            End If

            ' Update the device's last sync time
            Call Settings.SetEnrolledDeviceLastSyncTime(Settings.EnrolledDevices(i), DeviceDownloadStartTime)
        Next

        ' Check to see if the sync was cancelled
        If Not IsSyncCancelled() Then
            ' Attempt to upload all records to Nightscout. 
            If Await UploadRecordsToNightscout(NightscoutRecords) Then
                Settings.LastSyncTime = DateTime.Now
            End If
        Else
            ' Indicate the cancellation was successful.
            Call ConfirmSyncCancelled()
        End If

        ' Update the UI, indicate we're no longer syncing, restart the timer event, and notify that we're not syncing anymore
        Call UpdateSyncStatus(False, "Last Sync: " & Settings.LastSyncTime)
        Call Volatile.Write(pSyncRunning, False)
        Call StartTimer()
        RaiseEvent SynchronizationStopped()

        Return
    End Sub

    ''' <summary>
    ''' Connects to the Nightscout server and attempts to get the most recent entry, then gets the time that entry was uploaded
    ''' </summary>
    ''' <returns>A DateTime representing the time the last entry was uploaded, or DateTime.MinValue on error</returns>
    Private Shared Async Function GetLastNightscoutSyncTime() As Task(Of DateTime)
        ' Attempt to get the latest entry entered into Nightscout
        Dim WebClient As New HttpClient()
        Dim SingleEntryUri As Uri = New Uri("http://" & Settings.NightscoutURL & "/api/v1/entries.json?count=1")
        Dim EntriesString As String = ""
        Try
            EntriesString = Await WebClient.GetStringAsync(SingleEntryUri)
        Catch ex As Exception
            ' If we can't get the entries fail out
            Return DateTime.MinValue
        End Try

        Dim GlucoseEntries() As NightscoutGlucoseEntry = Nothing
        Dim LastEntryTime As DateTime = DateTime.MinValue
        Dim GlucoseEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutGlucoseEntry()))

        If EntriesString <> "" Then
            Dim JsonStream As New MemoryStream(UTF8.GetBytes(EntriesString))

            Try
                ' Serialize the listing from JSON
                GlucoseEntries = GlucoseEntrySerializer.ReadObject(JsonStream)
            Catch Ex As Exception
                ' If something goes wrong then fail out
                Return DateTime.MinValue
            End Try

            ' Get the date and time from the item and attempt to calculate the date from epoch
            Try
                If Not GlucoseEntries Is Nothing Then
                    If GlucoseEntries.Count > 0 Then
                        LastEntryTime = DateTimeFromUnixEpoch(CULng(CStr(GlucoseEntries(0).date).Substring(0, CStr(GlucoseEntries(0).date).Length - 3))).ToLocalTime
                    End If
                End If
            Catch Ex As Exception
                Return DateTime.MinValue
            End Try
        End If

        Return LastEntryTime
    End Function

    ''' <summary>
    ''' Reads all entry data from Dexcom Receiver devices.
    ''' </summary>
    ''' <param name="SyncDevice">The connected DexcomDevice to read</param>
    ''' <param name="LastEntryTime">The farthest time to read behind</param>
    ''' <param name="NightscoutRecords">A Collection of NightscoutGlucoseEntry records, an empty Collection, or Nothing to serve as the start of the array of records to collect</param>
    ''' <returns>A Collection of NightscoutGlucoseEntries created from data read from the device</returns>
    Private Shared Async Function ReadDexcomDeviceData(SyncDevice As DexcomDevice, ByVal LastEntryTime As DateTime, ByVal NightscoutRecords As Collection(Of NightscoutEntry)) As Task(Of Collection(Of NightscoutEntry))
        Call UpdateSyncStatus("Connecting to " & SyncDevice.Manufacturer & " " & SyncDevice.Model & " (" & SyncDevice.SerialNumber & ")...")

        ' Attempt to connect to the device
        Try
            If Not Await SyncDevice.Connect() Then
                Return NightscoutRecords
            End If
        Catch Ex As Exception
            Return NightscoutRecords
        End Try

        ' If the Collection we were passed as a basis of results to return was Nothing, create the Collection
        If NightscoutRecords Is Nothing Then
            NightscoutRecords = New Collection(Of NightscoutEntry)
        End If

        ' Check to see if the sync was cancelled.
        If IsSyncCancelled() Then
            Await SyncDevice.Disconnect()
            Return NightscoutRecords
        End If

        ' Get all of the records from the last entry time, convert each record to a Nightscout record
        Try
            Call UpdateSyncStatus("Reading " & SyncDevice.Manufacturer & " " & SyncDevice.Model & " Sensor Data...")
            Dim DeviceEGVRecords As Collection(Of DatabaseRecord) = Await SyncDevice.DexcomReceiver.GetDatabaseContents(DatabasePartitions.EGVData, LastEntryTime)

            ' Check to see if the sync was cancelled.
            If IsSyncCancelled() Then
                Await SyncDevice.Disconnect()
                Return NightscoutRecords
            End If

            Dim DeviceSensorRecords As Collection(Of DatabaseRecord) = Await SyncDevice.DexcomReceiver.GetDatabaseContents(DatabasePartitions.SensorData, LastEntryTime)

            ' Check to see if the sync was cancelled.
            If IsSyncCancelled() Then
                Await SyncDevice.Disconnect()
                Return NightscoutRecords
            End If

            Call UpdateSyncStatus("Reading " & SyncDevice.Manufacturer & " " & SyncDevice.Model & " Meter Data...")
            Dim DeviceMeterRecords As Collection(Of DatabaseRecord) = Await SyncDevice.DexcomReceiver.GetDatabaseContents(DatabasePartitions.MeterData, LastEntryTime)

            ' Check to see if the sync was cancelled.
            If IsSyncCancelled() Then
                Await SyncDevice.Disconnect()
                Return NightscoutRecords
            End If

            Call UpdateSyncStatus("Reading " & SyncDevice.Manufacturer & " " & SyncDevice.Model & " Insertion Data...")
            Dim DeviceInsertionRecords As Collection(Of DatabaseRecord) = Await SyncDevice.DexcomReceiver.GetDatabaseContents(DatabasePartitions.InsertionTime, LastEntryTime)

            ' Check to see if the sync was cancelled.
            If IsSyncCancelled() Then
                Await SyncDevice.Disconnect()
                Return NightscoutRecords
            End If

            ' Look through the different record types and correlate the EGV and Sensor records based on time.
            ' Match based on whether the EGV record and Sensor records are within 10 seconds of each other.
            For i = 0 To DeviceEGVRecords.Count - 1
                Dim MatchFound As Boolean = False
                For j = 0 To DeviceSensorRecords.Count - 1
                    If (DeviceEGVRecords(i).SystemTime - DeviceSensorRecords(j).SystemTime).Duration() <= New TimeSpan(0, 0, 10) Then
                        Call NightscoutRecords.Add(ConvertToNightscoutGlucoseEntry(DeviceEGVRecords(i), DeviceSensorRecords(j)))
                        MatchFound = True
                        Exit For
                    End If
                Next

                ' If an EGV record didn't have a matching Sensor record for some reason add just that part of the data.
                If Not MatchFound Then
                    Call NightscoutRecords.Add(ConvertToNightscoutGlucoseEntry(DeviceEGVRecords(i), Nothing))
                End If
            Next

            ' Add Meter records
            For i = 0 To DeviceMeterRecords.Count - 1
                Call NightscoutRecords.Add(ConvertToNightscoutGlucoseEntry(DeviceMeterRecords(i)))
            Next

            ' Add sensor insertion events
            For i = 0 To DeviceInsertionRecords.Count - 1
                Dim TreatmentEntry As NightscoutTreatmentEntry = ConvertToNightscoutTreatmentEntry(DeviceInsertionRecords(i))
                If Not TreatmentEntry Is Nothing Then
                    Call NightscoutRecords.Add(TreatmentEntry)
                End If
            Next

            ' If there is uploader device status to report create the record
            Dim DeviceStatus As NightscoutDeviceStatusEntry = Await CreateDeviceStatusEntry()
            If Not DeviceStatus Is Nothing Then
                Call NightscoutRecords.Add(DeviceStatus)
            End If
        Catch Ex As Exception
            ' Don't do anything special, continue on to disconnect and return the results thus far
        End Try

        ' Disconnect from the device so it can be connected to later.
        Try
            Await SyncDevice.Disconnect()
        Catch ex As Exception
            ' Don't do anything special, just continue through
        End Try

        ' Return any results
        Return NightscoutRecords
    End Function

    ''' <summary>
    ''' Converts a Dexcom EGVDatabaseRecord and SensorDatabaseRecord pair to a NightscoutGlucoseEntry record.
    ''' </summary>
    ''' <param name="EGVRecord">An EGVDatabaseRecord read from a Dexcom device</param>
    ''' <returns>A NightscoutGlucoseEntry record to upload to Nightscout</returns>
    Private Shared Function ConvertToNightscoutGlucoseEntry(ByRef EGVRecord As EGVDatabaseRecord, ByRef SensorRecord As SensorDatabaseRecord) As NightscoutGlucoseEntry
        Dim NightscoutEntry As New NightscoutGlucoseEntry

        ' Deal with dates
        Dim Offset As New DateTimeOffset(EGVRecord.DisplayTime, TimeZoneInfo.Local.GetUtcOffset(EGVRecord.DisplayTime))
        NightscoutEntry.date = Offset.ToUnixTimeMilliseconds
        NightscoutEntry.dateString = ISO8601TimeFromDateTime(EGVRecord.DisplayTime)

        ' Find the trend arrow
        Select Case EGVRecord.TrendArrow
            Case EGVDatabaseRecord.TrendArrows.DoubleDown
                NightscoutEntry.direction = "DoubleDown"
            Case EGVDatabaseRecord.TrendArrows.DoubleUp
                NightscoutEntry.direction = "DoubleUp"
            Case EGVDatabaseRecord.TrendArrows.Flat
                NightscoutEntry.direction = "Flat"
            Case EGVDatabaseRecord.TrendArrows.FortyFiveDown
                NightscoutEntry.direction = "FortyFiveDown"
            Case EGVDatabaseRecord.TrendArrows.FortyFiveUp
                NightscoutEntry.direction = "FortyFiveUp"
            Case EGVDatabaseRecord.TrendArrows.SingleDown
                NightscoutEntry.direction = "SingleDown"
            Case EGVDatabaseRecord.TrendArrows.SingleUp
                NightscoutEntry.direction = "SingleUp"

            Case Else
                NightscoutEntry.direction = ""
        End Select

        NightscoutEntry.[type] = "sgv"
        NightscoutEntry.sgv = EGVRecord.GlucoseLevel
        NightscoutEntry.noise = EGVRecord.Noise
        NightscoutEntry.device = "WindowsUploader-DexcomShare"

        If Not SensorRecord Is Nothing Then
            NightscoutEntry.rssi = SensorRecord.RSSI
            NightscoutEntry.filtered = SensorRecord.Filtered
            NightscoutEntry.unfiltered = SensorRecord.Unfiltered
        End If

        Return NightscoutEntry
    End Function

    ''' <summary>
    ''' Converts a Dexcom MeterDatabaseRecord to a NightscoutGlucoseEntry record.
    ''' </summary>
    ''' <param name="MeterRecord">A MeterDatabaseRecord read from a Dexcom device</param>
    ''' <returns>A NightscoutGlucoseEntry record to upload to Nightscout</returns>
    Private Shared Function ConvertToNightscoutGlucoseEntry(ByRef MeterRecord As MeterDatabaseRecord) As NightscoutGlucoseEntry
        Dim NightscoutEntry As New NightscoutGlucoseEntry

        ' Deal with dates
        Dim Offset As New DateTimeOffset(MeterRecord.DisplayTime, TimeZoneInfo.Local.GetUtcOffset(MeterRecord.DisplayTime))
        NightscoutEntry.date = Offset.ToUnixTimeMilliseconds
        NightscoutEntry.dateString = ISO8601TimeFromDateTime(MeterRecord.DisplayTime)

        NightscoutEntry.[type] = "mbg"
        NightscoutEntry.mbg = MeterRecord.MeterGlucose
        NightscoutEntry.device = "WindowsUploader-DexcomShare"

        Return NightscoutEntry
    End Function

    ''' <summary>
    ''' Creates an instance of a class that represents the status of the upload device if the device is a phone.
    ''' </summary>
    ''' <returns>An instance of NightscoutDeviceStatusEntry that represents the status of the upload device</returns>
    Private Shared Async Function CreateDeviceStatusEntry() As Task(Of NightscoutDeviceStatusEntry)
        ' Get battery data but only if the device is a phone. This doesn't work for laptops for some reason.
        If ApiInformation.IsApiContractPresent("Windows.Phone.PhoneContract", 1) Then
            Dim NightscoutEntry As New NightscoutDeviceStatusEntry

            ' Get the battery report
            Dim Report As BatteryReport = (Await Battery.FromIdAsync(Windows.Devices.Power.Battery.AggregateBattery.DeviceId)).GetReport()
            NightscoutEntry.created_at = ISO8601TimeFromDateTime(DateTime.Now)
            NightscoutEntry.uploaderBattery = CInt(CDbl(Report.RemainingCapacityInMilliwattHours.Value) / CDbl(Report.FullChargeCapacityInMilliwattHours.Value) * 100)

            Return NightscoutEntry
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Converts a Dexcom InsertionDatabaseRecord to a NightscoutTreatmentEntry record.
    ''' </summary>
    ''' <param name="InsertionRecord">A TreatmentDatabaseRecord read from a Dexcom device</param>
    ''' <returns>A NightscoutGlucoseEntry record to upload to Nightscout</returns>
    Private Shared Function ConvertToNightscoutTreatmentEntry(ByRef InsertionRecord As InsertionDatabaseRecord) As NightscoutTreatmentEntry
        Dim NightscoutEntry As New NightscoutTreatmentEntry

        ' Deal with dates
        NightscoutEntry.eventTime = ISO8601TimeFromDateTime(InsertionRecord.DisplayTime)
        NightscoutEntry.created_at = ISO8601TimeFromDateTime(InsertionRecord.DisplayTime)

        If InsertionRecord.InsertionState = InsertionDatabaseRecord.InsertionStates.Started Then
            NightscoutEntry.eventType = "Sensor Start"
        Else
            NightscoutEntry = Nothing
        End If

        Return NightscoutEntry
    End Function

    ''' <summary>
    ''' Uploads each record to the Nightscout server.
    ''' </summary>
    ''' <param name="NightscoutRecords">A Collection of NightscoutGlucoseEntry records to upload</param>
    ''' <returns>True if the upload was successful, False otherwise</returns>
    Private Shared Async Function UploadRecordsToNightscout(ByVal NightscoutRecords As Collection(Of NightscoutEntry)) As Task(Of Boolean)
        Dim Success As Boolean = False
        Dim GlucoseEntries As New Collection(Of NightscoutGlucoseEntry)
        Dim DeviceStatusEntries As New Collection(Of NightscoutDeviceStatusEntry)
        Dim TreatmentEntries As New Collection(Of NightscoutTreatmentEntry)

        Dim GlucoseEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutGlucoseEntry()))
        Dim DeviceStatusEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutDeviceStatusEntry()))
        Dim TreatmentEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutTreatmentEntry()))

        ' Divide up the records based on type so they get uploaded to the right endpoint
        If NightscoutRecords.Count > 0 Then
            For i = 0 To NightscoutRecords.Count - 1
                Select Case NightscoutRecords(i).EntryType
                    Case NightscoutEntry.EntryTypes.GlucoseEntry
                        GlucoseEntries.Add(NightscoutRecords(i))
                    Case NightscoutEntry.EntryTypes.DeviceStatusEntry
                        DeviceStatusEntries.Add(NightscoutRecords(i))
                    Case NightscoutEntry.EntryTypes.TreatmentEntry
                        TreatmentEntries.Add(NightscoutRecords(i))
                End Select
            Next
        End If

        If GlucoseEntries.Count > 0 Then
            ' Serialize to JSON
            Dim UploadRecordStream As New MemoryStream
            Dim NightscoutRecordsArray() As NightscoutGlucoseEntry = GlucoseEntries.ToArray()
            GlucoseEntrySerializer.WriteObject(UploadRecordStream, NightscoutRecordsArray)
            Dim UploadRecordString As String = UTF8.GetString(UploadRecordStream.ToArray())

            Call UpdateSyncStatus("Uploading device status to Nightscout...")
            Success = Await PostStringToNightscout(UploadRecordString, "/api/v1/entries.json")
        Else
            Success = True
        End If

        ' If the previous data upload was successful (or not attempted) attempt to upload device status
        If DeviceStatusEntries.Count > 0 And Success = True Then
            ' Serialize to JSON
            Dim UploadRecordStream As New MemoryStream
            Dim NightscoutRecordsArray() As NightscoutDeviceStatusEntry = DeviceStatusEntries.ToArray()
            DeviceStatusEntrySerializer.WriteObject(UploadRecordStream, DeviceStatusEntries)
            Dim UploadRecordString As String = UTF8.GetString(UploadRecordStream.ToArray())

            Call UpdateSyncStatus("Uploading device status to Nightscout...")
            Success = Await PostStringToNightscout(UploadRecordString, "/api/v1/devicestatus.json")
        Else
            Success = True
        End If

        If TreatmentEntries.Count > 0 And Success = True Then
            ' Serialize to JSON
            Dim UploadRecordStream As New MemoryStream
            Dim NightscoutRecordsArray() As NightscoutTreatmentEntry = TreatmentEntries.ToArray()
            TreatmentEntrySerializer.WriteObject(UploadRecordStream, NightscoutRecordsArray)
            Dim UploadRecordString As String = UTF8.GetString(UploadRecordStream.ToArray())

            Call UpdateSyncStatus("Uploading treatments to Nightscout...")
            Success = Await PostStringToNightscout(UploadRecordString, "/api/v1/treatments.json")
        Else
            Success = True
        End If

        Return Success
    End Function

    ''' <summary>
    ''' Posts the passed string to the passed endpoint. It makes three attempts in case of bad connectivity.
    ''' </summary>
    ''' <param name="UploadRecordString">A string to be uploaded to the Nightscout server</param>
    ''' <param name="NightscoutAPIEndpoint">The endpoint on the Nightscout server where the string is uploaded</param>
    ''' <returns>True of the POST is successful, False otherwise</returns>
    Private Shared Async Function PostStringToNightscout(ByVal UploadRecordString As String, ByVal NightscoutAPIEndpoint As String) As Task(Of Boolean)
        Dim HttpClient As New HttpClient
        Dim SyncResponse As HttpResponseMessage = Nothing

        ' Prepare the POST reqeuest and upload to Nightscout
        Dim SyncRequest As New HttpRequestMessage
        Call SyncRequest.Headers.Add("API-SECRET", Settings.NightscoutAPIKey)
        SyncRequest.RequestUri = New Uri("http://" & Settings.NightscoutURL & NightscoutAPIEndpoint)
        SyncRequest.Method = HttpMethod.Post
        SyncRequest.Content = New StringContent(UploadRecordString)
        SyncRequest.Content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")

        Call UpdateSyncStatus("Uploading treatment entries to Nightscout...")

        ' Do the upload, try three times
        For i = 1 To 3
            Try
                SyncResponse = Await HttpClient.SendAsync(SyncRequest)
                Return SyncResponse.IsSuccessStatusCode
            Catch ex As Exception
                Return False
            End Try
        Next

        Return False
    End Function

    ''' <summary>
    ''' Converts an integer representing Unix epoch time to a DateTime.
    ''' </summary>
    ''' <param name="UnixEpochTimeSeconds">Seconds since January 1, 1970 00:00:00</param>
    ''' <returns>A DateTime at GMT</returns>
    Private Shared Function DateTimeFromUnixEpoch(ByVal UnixEpochTimeSeconds As Double) As DateTime
        Return New DateTime((New DateTime(1970, 1, 1)).Ticks + (UnixEpochTimeSeconds * TimeSpan.TicksPerSecond))
    End Function

    ''' <summary>
    ''' Converts a DateTime to an integer representing Unix epoch time.
    ''' </summary>
    ''' <param name="RealTime">The DateTime to convert</param>
    ''' <returns>The number of seconds since January 1, 1970 00:00:00 at GMT</returns>
    Private Shared Function UnixEpochFromDateTime(ByVal RealTime As DateTime) As Double
        Return ((RealTime.Ticks / TimeSpan.TicksPerSecond) - ((New DateTime(1970, 1, 1)).Ticks / TimeSpan.TicksPerSecond))
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="RealTime"></param>
    ''' <returns></returns>
    Private Shared Function ISO8601TimeFromDateTime(ByVal RealTime As DateTime) As String
        ' Deal with dates
        Dim Offset As New DateTimeOffset(RealTime, TimeZoneInfo.Local.GetUtcOffset(RealTime))
        Dim Formats() As String = DateTime.Now.ToUniversalTime.GetDateTimeFormats("o".ToCharArray()(0))
        Return Formats(0)
    End Function
End Class
