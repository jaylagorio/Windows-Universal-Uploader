Imports Dexcom
Imports System.Net.Http
Imports System.Threading
Imports System.Text.UTF8Encoding
Imports Windows.Networking.Connectivity
Imports System.Runtime.Serialization.Json

''' <summary>
''' Author: Jay Lagorio
''' Date: May 15, 2016
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
        If pSyncTimer Is Nothing Then
            pSyncTimer = New Timer(AddressOf Synchronize, Nothing, Settings.UploadInterval * 60 * 1000, Settings.UploadInterval * 60 * 1000)
        End If
    End Sub

    ''' <summary>
    ''' Destroys the timer if one has been created.
    ''' </summary>
    Public Shared Sub StopTimer()
        If Not pSyncTimer Is Nothing Then
            pSyncTimer = Nothing
        End If
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
        Dim NightscoutRecords As New Collection(Of NightscoutGlucoseEntry)
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
            If Settings.EnrolledDevices(i).GetType Is GetType(DexcomDevice) Then
                NightscoutRecords = Await ReadDexcomDeviceData(Settings.EnrolledDevices(i), Settings.EnrolledDevices(i).LastSyncTime, NightscoutRecords)

                ' Update the device's last sync time
                Call Settings.SetEnrolledDeviceLastSyncTime(Settings.EnrolledDevices(i), DateTime.Now)
            End If
        Next

        ' Attempt to upload all records to Nightscout. If successful within three tries update the last sync time setting.
        For i = 1 To 3
            If Await UploadRecordsToNightscout(NightscoutRecords) Then
                Settings.LastSyncTime = DateTime.Now
                Exit For
            End If
        Next

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
    Private Shared Async Function ReadDexcomDeviceData(SyncDevice As DexcomDevice, ByVal LastEntryTime As DateTime, ByVal NightscoutRecords As Collection(Of NightscoutGlucoseEntry)) As Task(Of Collection(Of NightscoutGlucoseEntry))
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
            NightscoutRecords = New Collection(Of NightscoutGlucoseEntry)
        End If

        ' Get all of the records from the last entry time, convert each record to a Nightscout record
        Call UpdateSyncStatus("Reading " & SyncDevice.Manufacturer & " " & SyncDevice.Model & " (" & SyncDevice.SerialNumber & ")...")
        Try
            Dim DeviceEGVRecords As Collection(Of DatabaseRecord) = Await SyncDevice.DexcomReceiver.GetDatabaseContents(DatabasePartitions.EGVData, LastEntryTime)
            For i = 0 To DeviceEGVRecords.Count - 1
                Call NightscoutRecords.Add(ConvertToNightscoutGlucoseEntry(DeviceEGVRecords(i)))
            Next
        Catch Ex As Exception
            ' Don't do anything special, continue on to disconnect and return the results thus far
        End Try

        ' Disconnect from the device so it can be connected to later.
        Await SyncDevice.Disconnect()

        ' Return any results
        Return NightscoutRecords
    End Function

    ''' <summary>
    ''' Converts a Dexcom EGVDatabaseRecord to a NightscoutGlucoseEntry record.
    ''' </summary>
    ''' <param name="EGVRecord">An EGVDatabaseRecord read from a Dexcom device</param>
    ''' <returns>A NightscoutGlucoseEntry record to upload to Nightscout</returns>
    Private Shared Function ConvertToNightscoutGlucoseEntry(ByRef EGVRecord As EGVDatabaseRecord) As NightscoutGlucoseEntry
        Dim NightscoutEntry As New NightscoutGlucoseEntry

        ' Deal with dates
        Dim Offset As New DateTimeOffset(EGVRecord.DisplayTime, TimeZoneInfo.Local.GetUtcOffset(EGVRecord.DisplayTime))
        Dim Formats() As String = EGVRecord.DisplayTime.ToUniversalTime.GetDateTimeFormats("r".ToCharArray()(0))
        NightscoutEntry.date = Offset.ToUnixTimeMilliseconds
        NightscoutEntry.dateString = Formats(0)

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
        NightscoutEntry.device = "WindowsUploader-DexcomShare"

        Return NightscoutEntry
    End Function

    ''' <summary>
    ''' Uploads each record to the Nightscout server.
    ''' </summary>
    ''' <param name="NightscoutRecords">A Collection of NightscoutGlucoseEntry records to upload</param>
    ''' <returns>True if the upload was successful, False otherwise</returns>
    Private Shared Async Function UploadRecordsToNightscout(ByVal NightscoutRecords As Collection(Of NightscoutGlucoseEntry)) As Task(Of Boolean)
        If NightscoutRecords.Count > 0 Then
            ' Serialize to JSON
            Dim UploadRecordStream As New MemoryStream
            Dim NightscoutRecordsArray() As NightscoutGlucoseEntry = NightscoutRecords.ToArray()
            Dim GlucoseEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutGlucoseEntry()))
            GlucoseEntrySerializer.WriteObject(UploadRecordStream, NightscoutRecordsArray)
            Dim UploadRecordString As String = UTF8.GetString(UploadRecordStream.ToArray())

            ' Prepare the POST reqeuest and upload to Nightscout
            Dim HttpClient As New HttpClient
            Dim SyncResponse As HttpResponseMessage = Nothing

            Dim SyncRequest As New HttpRequestMessage
            Call SyncRequest.Headers.Add("API-SECRET", Settings.NightscoutAPIKey)
            SyncRequest.RequestUri = New Uri("http://" & Settings.NightscoutURL & "/api/v1/entries.json")
            SyncRequest.Method = HttpMethod.Post
            SyncRequest.Content = New StringContent(UploadRecordString)
            SyncRequest.Content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")

            ' Do the upload
            Call UpdateSyncStatus("Uploading results to Nightscout...")
            Try
                SyncResponse = Await HttpClient.SendAsync(SyncRequest)
            Catch ex As Exception
                Return False
            End Try
        End If

        Return True
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
End Class
