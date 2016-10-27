Imports System.Net.Http
Imports System.Text.UTF8Encoding
Imports System.Runtime.Serialization.Json

''' <summary>
''' Author: Jay Lagorio
''' Date: October 30, 2016
''' Summary: A class used to interact with Nightscout servers.
''' </summary>

Public Class Server
    ' Data available from the Nightscout dashboard
    Public Structure NightscoutDashboardData
        Dim TrendArrow As String            ' The trend arrow
        Dim CGMUpdateDelta As Integer       ' The time in minutes since the last sensor reading
        Dim CurrentSGV As Integer           ' The current sensor value
        Dim DifferentialSGV As Integer      ' The difference between the current sensor value and the previous
        Dim ResultsLoaded As Boolean        ' Indicates the structure is informationally complete
    End Structure

    ' The server name of the Nightscout host
    Private pNightscoutURL As String

    ' Switches from HTTP to HTTPS
    Private pUseSSL As Boolean

    ' The API key used to submit data to the server if that's ever necessary
    Private pNightscoutAPIKey As String

    ''' <summary>
    ''' Creates the Server object targetted at the described Nightscout server.
    ''' </summary>
    Sub New(ByVal NightscoutURL As String, ByVal UseSSL As Boolean, ByVal NightscoutAPIKey As String)
        pNightscoutURL = NightscoutURL
        pUseSSL = UseSSL
        pNightscoutAPIKey = NightscoutAPIKey
    End Sub

    ''' <summary>
    ''' Returns data currently displayed on the Nightscout screen depending on what features are turned on.
    ''' </summary>
    ''' <returns>A structure containing data as Nightscout currently understands it. If the structure's ResultsLoaded value is True the data should be considered valid.</returns>
    Public Async Function GetDashboardData() As Task(Of NightscoutDashboardData)
        ' Attempt to get the latest entry entered into Nightscout
        Dim WebClient As New HttpClient()
        Dim DashboardData As New NightscoutDashboardData

        ' Setup the SGV data URI
        Dim LastEntriesUri As Uri
        If pUseSSL Then
            LastEntriesUri = New Uri("https://" & pNightscoutURL & "/api/v1/entries.json?type=sgv&count=2")
        Else
            LastEntriesUri = New Uri("http://" & pNightscoutURL & "/api/v1/entries.json?type=sgv&count=2")
        End If

        ' Get the last two SGV values from the Nightscout server
        Dim EntriesString As String = ""
        Try
            EntriesString = Await WebClient.GetStringAsync(LastEntriesUri)
        Catch ex As Exception
            ' If we can't get the entries fail out
            Return Nothing
        End Try

        Dim GlucoseEntries() As NightscoutGlucoseEntry = Nothing
        Dim LastEntryTime As DateTime = DateTime.MinValue
        Dim GlucoseEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutGlucoseEntry()))

        If EntriesString <> "" Then
            Dim JsonStream As New MemoryStream(UTF8.GetBytes(EntriesString))
            ' Serialize the JSON into something usable
            Try
                ' Serialize the listing from JSON
                GlucoseEntries = GlucoseEntrySerializer.ReadObject(JsonStream)
            Catch Ex As Exception
                ' If something goes wrong then fail out
                Return Nothing
            End Try

            Try
                If Not GlucoseEntries Is Nothing Then
                    If GlucoseEntries.Count > 0 Then
                        DashboardData.CurrentSGV = GlucoseEntries(0).sgv
                        DashboardData.TrendArrow = GlucoseEntries(0).direction.ToUpper

                        ' If there's more than one entry get the two differential values
                        If GlucoseEntries.Count > 1 Then
                            DashboardData.DifferentialSGV = GlucoseEntries(0).sgv - GlucoseEntries(1).sgv
                            DashboardData.CGMUpdateDelta = (DateTime.Now.ToLocalTime - DateTimeFromUnixEpoch(CULng(CStr(GlucoseEntries(0).date).Substring(0, CStr(GlucoseEntries(0).date).Length - 3))).ToLocalTime).TotalMinutes
                        End If

                        DashboardData.ResultsLoaded = True
                    End If
                End If
            Catch Ex As Exception
                Return Nothing
            End Try
        End If

        Return DashboardData
    End Function

    ''' <summary>
    ''' Posts the passed string to the passed endpoint. It makes three attempts in case of bad connectivity.
    ''' </summary>
    ''' <param name="UploadRecordString">A string to be uploaded to the Nightscout server</param>
    ''' <param name="NightscoutAPIEndpoint">The endpoint on the Nightscout server where the string is uploaded</param>
    ''' <returns>True of the POST is successful, False otherwise</returns>
    Public Async Function PostStringToServer(ByVal UploadRecordString As String, ByVal NightscoutAPIEndpoint As String) As Task(Of Boolean)
        Dim HttpClient As New HttpClient
        Dim SyncResponse As HttpResponseMessage = Nothing

        ' Prepare the POST reqeuest and upload to Nightscout
        Dim SyncRequest As New HttpRequestMessage
        Call SyncRequest.Headers.Add("API-SECRET", pNightscoutAPIKey)
        If pUseSSL Then
            SyncRequest.RequestUri = New Uri("https://" & pNightscoutURL & NightscoutAPIEndpoint)
        Else
            SyncRequest.RequestUri = New Uri("http://" & pNightscoutURL & NightscoutAPIEndpoint)
        End If
        SyncRequest.Method = HttpMethod.Post
        SyncRequest.Content = New StringContent(UploadRecordString)
        SyncRequest.Content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")

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
    ''' Connects to the Nightscout server and attempts to get the most recent entry, then gets the time that entry was uploaded
    ''' </summary>
    ''' <returns>A DateTime representing the time the last entry was uploaded, or DateTime.MinValue on error</returns>
    Public Async Function GetLastSyncTime() As Task(Of DateTime)
        ' Attempt to get the latest entry entered into Nightscout
        Dim WebClient As New HttpClient()

        Dim SingleEntryUri As Uri
        If pUseSSL Then
            SingleEntryUri = New Uri("https://" & pNightscoutURL & "/api/v1/entries.json?count=1")
        Else
            SingleEntryUri = New Uri("http://" & pNightscoutURL & "/api/v1/entries.json?count=1")
        End If

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
    ''' Uploads each record to the Nightscout server.
    ''' </summary>
    ''' <param name="NightscoutRecords">A Collection of NightscoutGlucoseEntry records to upload</param>
    ''' <returns>True if the upload was successful, False otherwise</returns>
    Public Async Function UploadRecords(ByVal NightscoutRecords As Collection(Of NightscoutEntry)) As Task(Of Boolean)
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

            Success = Await PostStringToServer(UploadRecordString, "/api/v1/entries.json")
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

            Success = Await PostStringToServer(UploadRecordString, "/api/v1/devicestatus.json")
        Else
            Success = True
        End If

        If TreatmentEntries.Count > 0 And Success = True Then
            ' Serialize to JSON
            Dim UploadRecordStream As New MemoryStream
            Dim NightscoutRecordsArray() As NightscoutTreatmentEntry = TreatmentEntries.ToArray()
            TreatmentEntrySerializer.WriteObject(UploadRecordStream, NightscoutRecordsArray)
            Dim UploadRecordString As String = UTF8.GetString(UploadRecordStream.ToArray())

            Success = Await PostStringToServer(UploadRecordString, "/api/v1/treatments.json")
        Else
            Success = True
        End If

        Return Success
    End Function

    ''' <summary>
    ''' Converts an integer representing Unix epoch time to a DateTime.
    ''' </summary>
    ''' <param name="UnixEpochTimeSeconds">Seconds since January 1, 1970 00:00:00</param>
    ''' <returns>A DateTime at GMT</returns>
    Private Function DateTimeFromUnixEpoch(ByVal UnixEpochTimeSeconds As Double) As DateTime
        Return New DateTime((New DateTime(1970, 1, 1)).Ticks + (UnixEpochTimeSeconds * TimeSpan.TicksPerSecond))
    End Function
End Class
