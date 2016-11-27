Imports System.Net.Http
Imports System.Text.UTF8Encoding
Imports System.Runtime.Serialization.Json

''' <summary>
''' Author: Jay Lagorio
''' Date: November 6, 2016
''' Summary: A class used to interact with Nightscout servers.
''' </summary>

Public Class Server
    ' Types of data that can be retrieved from a Nightscout server.
    Public Enum NightscoutDataTypes
        CGM = 1                 ' Retrieve CGM data
        OpenAPS = 2             ' Retrieve OpenAPS loop data
        Pump = 4                ' Retrieve insulin pump status
        FoodEntry = 8           ' Retrieve food entry data
        TreatmentEntry = 16     ' Retrieve bolus entry data
    End Enum

    ' Possible states for the pump
    Public Enum PumpStatus
        Normal = 0      ' Normal state
        Bolusing = 1    ' The pump is bolusing
        Suspended = 2   ' The pump is suspended
    End Enum

    ' Nightscout system status
    Public Structure NightscoutData
        Dim NightscoutCGMData As NightscoutCGMData                      ' CGM data returned when NightscoutDataTypes.CGM is set
        Dim NightscoutOpenAPSData As NightscoutOpenAPSData              ' OpenAPS data returned when NightscoutDataTypes.OpenAPS is set
        Dim NightscoutPumpData As NightscoutPumpData                    ' Pump data returned when NightscoutDataTypes.Pump is set
        Dim NightscoutFoodTreatmentData As NightscoutTreatmentData      ' Meal/snack data returned when NightscoutDataTypes.FoodEntry is set
        Dim NightscoutBolusTreatmentData As NightscoutTreatmentData     ' Insulin data returned when NightscoutDataTypes.TreatmentEntry is set
    End Structure

    ' Nightscout CGM data
    Public Structure NightscoutCGMData
        Dim UpdateDelta As Integer      ' The time in minutes since the last sensor reading
        Dim TrendArrow As String        ' The trend arrow
        Dim CurrentSGV As Integer       ' The current sensor value
        Dim DifferentialSGV As Integer  ' The difference between the current sensor value and the previous
        Dim ResultsLoaded As Boolean    ' Indicates the structure is informationally complete
    End Structure

    ' Nightscout OpenAPS data
    Public Structure NightscoutOpenAPSData
        Dim UpdateDelta As Integer   ' The time in minutes since the last loop execution
        Dim IOB As Double            ' Current insulin on board
        Dim CurrentBasal As Double   ' Current basal rate (temp or otherwise)
        Dim Duration As Integer      ' The duration of any temporary basal rate
        Dim CurrentSGV As Integer    ' The sensor glucose value used to calculate this state
        Dim Reason As String         ' The OpenAPS verbose reason that got to this state
        Dim Enacted As Boolean       ' Whether this data has been enacted on the user
        Dim ResultsLoaded As Boolean ' Indicates the structure is informationally complete
    End Structure

    ' Nightscout Pump data
    Public Structure NightscoutPumpData
        Dim UpdateDelta As Integer      ' The time in minutes since the last loop execution
        Dim BatteryVoltage As Double    ' Voltage for the insulin pump battery
        Dim Status As PumpStatus        ' The state of the pump
        Dim ReservoirLevel As Double    ' The amount of insulin in the reservoir
        Dim ResultsLoaded As Boolean    ' Indicates the structure is informationally complete
    End Structure

    ' Nightscout Treatment data
    Public Structure NightscoutTreatmentData
        Dim TreatmentTime As DateTime                               ' The time this treatment occurred
        Dim TreatmentType As NightscoutTreatmentEntry.EntryTypes    ' The type of treatment
        Dim Units As Double                                         ' The number of units in the treatment, if applicable
        Dim Carbohydrates As Integer                                ' The number of carbs in the treatment, if applicable
        Dim ResultsLoaded As Boolean                                ' Indicates the structure is informationally complete
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
    ''' <param name="DataType">An enum specifying the type of data to retrieve. These values can be added together to get multiple data types in a single call.</param>
    ''' <returns>A structure containing system data as Nightscout currently understands it. If the substructures' ResultsLoaded value is True the data should be considered valid.</returns>
    Public Async Function GetCurrentData(ByVal DataType As NightscoutDataTypes) As Task(Of NightscoutData)
        ' Get a method for making HTTP requests
        Dim WebClient As New HttpClient

        ' Initialize the return data structure
        Dim SystemData As New NightscoutData
        SystemData.NightscoutCGMData = New NightscoutCGMData
        SystemData.NightscoutOpenAPSData = New NightscoutOpenAPSData
        SystemData.NightscoutPumpData = New NightscoutPumpData
        SystemData.NightscoutFoodTreatmentData = New NightscoutTreatmentData
        SystemData.NightscoutBolusTreatmentData = New NightscoutTreatmentData

        ' Check to see if CGM data is being requested
        If (DataType And NightscoutDataTypes.CGM) = NightscoutDataTypes.CGM Then
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
                EntriesString = ""
            End Try

            ' Setup the array and serializer
            Dim GlucoseEntries() As NightscoutGlucoseEntry = Nothing
            Dim GlucoseEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutGlucoseEntry()))

            If EntriesString <> "" Then
                Dim SerializationFailure As Boolean = False
                Dim JsonStream As New MemoryStream(UTF8.GetBytes(EntriesString))
                ' Serialize the JSON into something usable
                Try
                    ' Serialize the listing from JSON
                    GlucoseEntries = GlucoseEntrySerializer.ReadObject(JsonStream)
                Catch Ex As Exception
                    ' If something goes wrong then fail out
                    SerializationFailure = True
                End Try

                If Not SerializationFailure Then
                    Try
                        If Not GlucoseEntries Is Nothing Then
                            ' If there is glucose data, add it to the structure
                            If GlucoseEntries.Count > 0 Then
                                SystemData.NightscoutCGMData.CurrentSGV = GlucoseEntries(0).sgv
                                SystemData.NightscoutCGMData.TrendArrow = GlucoseEntries(0).direction.ToUpper

                                ' If there's more than one entry get the two differential values
                                If GlucoseEntries.Count > 1 Then
                                    SystemData.NightscoutCGMData.DifferentialSGV = GlucoseEntries(0).sgv - GlucoseEntries(1).sgv
                                    SystemData.NightscoutCGMData.UpdateDelta = (DateTime.Now.ToLocalTime - DateTimeFromUnixEpoch(CULng(CStr(GlucoseEntries(0).date).Substring(0, CStr(GlucoseEntries(0).date).Length - 3))).ToLocalTime).TotalMinutes
                                End If

                                SystemData.NightscoutCGMData.ResultsLoaded = True
                            End If
                        End If
                    Catch Ex As Exception
                        SystemData.NightscoutCGMData.ResultsLoaded = False
                    End Try
                End If
            End If
        End If

        ' Check to see if OpenAPS or pump data is being requested
        If ((DataType And NightscoutDataTypes.OpenAPS) = NightscoutDataTypes.OpenAPS) Or ((DataType And NightscoutDataTypes.Pump) = NightscoutDataTypes.Pump) Then
            ' Setup the OpenAPS data URI
            Dim LastOpenAPSUri As Uri
            If pUseSSL Then
                LastOpenAPSUri = New Uri("https://" & pNightscoutURL & "/api/v1/devicestatus.json?count=1")
            Else
                LastOpenAPSUri = New Uri("http://" & pNightscoutURL & "/api/v1/devicestatus.json?count=1")
            End If

            ' Get the last two SGV values from the Nightscout server
            Dim EntriesString As String = ""
            Try
                EntriesString = Await WebClient.GetStringAsync(LastOpenAPSUri)
            Catch ex As Exception
                ' If we can't get the entries fail out
                EntriesString = ""
            End Try

            ' Setup the array and serializer
            Dim OpenAPSEntries() As NightscoutDeviceStatusEntry = Nothing
            Dim OpenAPSEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutDeviceStatusEntry()))

            If EntriesString <> "" Then
                Dim SerializationFailure As Boolean = False
                Dim JsonStream As New MemoryStream(UTF8.GetBytes(EntriesString))
                ' Serialize the JSON into something usable
                Try
                    ' Serialize the listing from JSON
                    OpenAPSEntries = OpenAPSEntrySerializer.ReadObject(JsonStream)
                Catch Ex As Exception
                    ' If something goes wrong then fail out
                    SerializationFailure = True
                End Try

                If Not SerializationFailure Then
                    Try
                        If Not OpenAPSEntries Is Nothing Then
                            If OpenAPSEntries.Count > 0 Then
                                ' If there was OpenAPS data, add it to the structure
                                Try
                                    If Not OpenAPSEntries(0).openaps.enacted Is Nothing Then
                                        SystemData.NightscoutOpenAPSData.CurrentBasal = OpenAPSEntries(0).openaps.enacted.rate
                                        SystemData.NightscoutOpenAPSData.Enacted = OpenAPSEntries(0).openaps.enacted.received
                                        SystemData.NightscoutOpenAPSData.IOB = OpenAPSEntries(0).openaps.enacted.IOB
                                        SystemData.NightscoutOpenAPSData.Duration = OpenAPSEntries(0).openaps.enacted.duration
                                        SystemData.NightscoutOpenAPSData.CurrentSGV = OpenAPSEntries(0).openaps.enacted.bg
                                        SystemData.NightscoutOpenAPSData.Reason = OpenAPSEntries(0).openaps.enacted.reason
                                        SystemData.NightscoutOpenAPSData.UpdateDelta = (DateTime.Now.ToLocalTime - DateTime.Parse(OpenAPSEntries(0).openaps.enacted.timestamp).ToLocalTime).TotalMinutes
                                        SystemData.NightscoutOpenAPSData.ResultsLoaded = True
                                    End If
                                Catch ex As Exception
                                    SystemData.NightscoutOpenAPSData.ResultsLoaded = False
                                End Try

                                ' If there was pump status data, add it to the structure
                                Try
                                    If Not OpenAPSEntries(0).pump Is Nothing Then
                                        SystemData.NightscoutPumpData.UpdateDelta = (DateTime.Now.ToLocalTime - DateTime.Parse(OpenAPSEntries(0).pump.status.timestamp).ToLocalTime).TotalMinutes
                                        SystemData.NightscoutPumpData.BatteryVoltage = OpenAPSEntries(0).pump.battery.voltage
                                        SystemData.NightscoutPumpData.ReservoirLevel = OpenAPSEntries(0).pump.reservoir

                                        If OpenAPSEntries(0).pump.status.suspended Then
                                            SystemData.NightscoutPumpData.Status = PumpStatus.Suspended
                                        ElseIf OpenAPSEntries(0).pump.status.bolusing Then
                                            SystemData.NightscoutPumpData.Status = PumpStatus.Bolusing
                                        Else
                                            SystemData.NightscoutPumpData.Status = PumpStatus.Normal
                                        End If

                                        SystemData.NightscoutPumpData.ResultsLoaded = True
                                    End If
                                Catch ex As Exception
                                    SystemData.NightscoutPumpData.ResultsLoaded = False
                                End Try
                            End If
                        End If
                    Catch Ex As Exception
                        SystemData.NightscoutOpenAPSData.ResultsLoaded = False
                        SystemData.NightscoutPumpData.ResultsLoaded = False
                    End Try
                End If
            End If
        End If

        ' Check to see if food entry data is being requested
        If (DataType And NightscoutDataTypes.FoodEntry) = NightscoutDataTypes.FoodEntry Then
            ' Setup the food treatment data URI
            Dim LastTreatmentUri As Uri
            If pUseSSL Then
                LastTreatmentUri = New Uri("https://" & pNightscoutURL & "/api/v1/treatments.json?count=1&find[carbs][$gt]=0")
            Else
                LastTreatmentUri = New Uri("http://" & pNightscoutURL & "/api/v1/treatments.json?count=1&find[carbs][$gt]=0")
            End If

            ' Get the last two SGV values from the Nightscout server
            Dim EntriesString As String = ""
            Try
                EntriesString = Await WebClient.GetStringAsync(LastTreatmentUri)
            Catch ex As Exception
                ' If we can't get the entries fail out
                EntriesString = ""
            End Try

            ' Setup the array and serializer
            Dim TreatmentEntries() As NightscoutTreatmentEntry = Nothing
            Dim TreatmentEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutTreatmentEntry()))

            If EntriesString <> "" Then
                Dim SerializationFailure As Boolean = False
                Dim JsonStream As New MemoryStream(UTF8.GetBytes(EntriesString))
                ' Serialize the JSON into something usable
                Try
                    ' Serialize the listing from JSON
                    TreatmentEntries = TreatmentEntrySerializer.ReadObject(JsonStream)
                Catch Ex As Exception
                    ' If something goes wrong then fail out
                    SerializationFailure = True
                End Try

                If Not SerializationFailure Then
                    Try
                        If Not TreatmentEntries Is Nothing Then
                            If TreatmentEntries.Count > 0 Then
                                SystemData.NightscoutFoodTreatmentData.Units = TreatmentEntries(0).insulin
                                SystemData.NightscoutFoodTreatmentData.Carbohydrates = TreatmentEntries(0).carbs
                                SystemData.NightscoutFoodTreatmentData.TreatmentType = NightscoutEntry.EntryTypes.TreatmentEntry
                                SystemData.NightscoutFoodTreatmentData.TreatmentTime = DateTime.Parse(TreatmentEntries(0).timestamp)

                                SystemData.NightscoutFoodTreatmentData.ResultsLoaded = True
                            End If
                        End If
                    Catch Ex As Exception
                        SystemData.NightscoutFoodTreatmentData.ResultsLoaded = False
                    End Try
                End If
            End If
        End If

        ' Check to see if bolus data is being requested
        If (DataType And NightscoutDataTypes.TreatmentEntry) = NightscoutDataTypes.TreatmentEntry Then
            ' Setup the bolus treatment data URI
            Dim LastTreatmentUri As Uri
            If pUseSSL Then
                LastTreatmentUri = New Uri("https://" & pNightscoutURL & "/api/v1/treatments.json?count=1&find[insulin][$gt]=0")
            Else
                LastTreatmentUri = New Uri("http://" & pNightscoutURL & "/api/v1/treatments.json?count=1&find[insulin][$gt]=0")
            End If

            ' Get the last two SGV values from the Nightscout server
            Dim EntriesString As String = ""
            Try
                EntriesString = Await WebClient.GetStringAsync(LastTreatmentUri)
            Catch ex As Exception
                ' If we can't get the entries fail out
                EntriesString = ""
            End Try

            ' Setup the array and serializer
            Dim TreatmentEntries() As NightscoutTreatmentEntry = Nothing
            Dim TreatmentEntrySerializer As New DataContractJsonSerializer(GetType(NightscoutTreatmentEntry()))

            If EntriesString <> "" Then
                Dim SerializationFailure As Boolean = False
                Dim JsonStream As New MemoryStream(UTF8.GetBytes(EntriesString))
                ' Serialize the JSON into something usable
                Try
                    ' Serialize the listing from JSON
                    TreatmentEntries = TreatmentEntrySerializer.ReadObject(JsonStream)
                Catch Ex As Exception
                    ' If something goes wrong then fail out
                    SerializationFailure = True
                End Try

                If Not SerializationFailure Then
                    Try
                        If Not TreatmentEntries Is Nothing Then
                            If TreatmentEntries.Count > 0 Then
                                SystemData.NightscoutBolusTreatmentData.Units = TreatmentEntries(0).insulin
                                SystemData.NightscoutBolusTreatmentData.Carbohydrates = TreatmentEntries(0).carbs
                                SystemData.NightscoutBolusTreatmentData.TreatmentType = NightscoutEntry.EntryTypes.TreatmentEntry
                                SystemData.NightscoutBolusTreatmentData.TreatmentTime = DateTime.Parse(TreatmentEntries(0).timestamp)

                                SystemData.NightscoutBolusTreatmentData.ResultsLoaded = True
                            End If
                        End If
                    Catch Ex As Exception
                        SystemData.NightscoutBolusTreatmentData.ResultsLoaded = False
                    End Try
                End If
            End If
        End If

        Return SystemData
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
