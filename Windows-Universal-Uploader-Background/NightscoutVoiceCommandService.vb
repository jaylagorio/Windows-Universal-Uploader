Imports Nightscout
Imports Windows.Storage
Imports Windows.Storage.ApplicationData
Imports Windows.ApplicationModel.AppService
Imports Windows.ApplicationModel.Background
Imports Windows.ApplicationModel.VoiceCommands

''' <summary>
''' Author: Jay Lagorio
''' Date: November 6, 2016
''' Summary: This service handles interactions from incomming Cortana commands.
''' </summary>

Public NotInheritable Class NightscoutVoiceCommandService
    Implements IBackgroundTask

    ' Indicates whether the First Run setup process was completed
    Dim pFirstRunSetupDone As Boolean

    ' The URL to the Nightscout server
    Dim pNightscoutURL As String

    ' Whether to connect to Nightscout using SSL
    Dim pUseSSL As Boolean

    ' The API key to use when adding data to Nightscout
    Dim pNightscoutAPIKey As String

    ' Keys for the settings Key/Value pairs we need for background operations
    Private Const UseSecureUploadConnectionKey As String = "UseSecureUploadConnection"
    Private Const FirstRunSetupDoneKey As String = "FirstRunSetupDone"
    Private Const NightscoutURLKey As String = "NightscoutURL"
    Private Const NightscoutSecretKey As String = "NightscoutSecret"
    Private Const UseRoamingSettingsKey As String = "UseRoamingSettings"

    ' The maximum size of a string that Cortana will tolerate for a response
    Private Const CortanaMaxDataLength As Integer = 256

    ''' <summary>
    ''' This routine is called by Cortana when a voice command is sent to the app.
    ''' </summary>
    ''' <param name="taskInstance">Information about the command given by the user</param>
    Public Async Sub Run(taskInstance As IBackgroundTaskInstance) Implements IBackgroundTask.Run
        ' Load settings as configured by the user
        Call LoadSettings()

        Dim TriggerDetails As AppServiceTriggerDetails = taskInstance.TriggerDetails
        If (Not TriggerDetails Is Nothing) Then
            If TriggerDetails.Name = "NightscoutVoiceCommandService" Then
                ' Take a differal, connect to the voice service, and get the voice command that got us here
                Dim Differal As BackgroundTaskDeferral = taskInstance.GetDeferral

                Dim VoiceServiceConnection As VoiceCommandServiceConnection = VoiceCommandServiceConnection.FromAppServiceTriggerDetails(TriggerDetails)
                Dim VoiceCommand As VoiceCommand = Await VoiceServiceConnection.GetVoiceCommandAsync

                ' Add completion handlers
                AddHandler taskInstance.Canceled, AddressOf Differal.Complete
                AddHandler VoiceServiceConnection.VoiceCommandCompleted, AddressOf Differal.Complete

                ' Check to make sure the First Run configuration has taken place
                If pFirstRunSetupDone Then
                    ' Make sure the Nightscout server URL is configured
                    If pNightscoutURL <> "" Then
                        ' Figure out which command was fired
                        Select Case VoiceCommand.CommandName
                            Case "GetCGM"
                                Await RespondToCGMVoiceCommand(VoiceServiceConnection)
                            Case "GetOpenAPS"
                                Await RespondToOpenAPSVoiceCommand(VoiceServiceConnection)
                            Case "GetPump"
                                Await RespondToPumpVoiceCommand(VoiceServiceConnection)
                            Case "GetLastEntryFood"
                                Await RespondToLastEntryFoodVoiceCommand(VoiceServiceConnection)
                            Case "GetLastEntryTreatment"
                                Await RespondToLastEntryTreatmentVoiceCommand(VoiceServiceConnection)
                            Case "GetFullSystemStatus"
                                Await RespondToFullSystemStatusVoiceCommand(VoiceServiceConnection)
                            Case Else
                                Call VoiceServiceConnection.RequestAppLaunchAsync(CreateCortanaResponse("Launching Nightscout"))
                        End Select
                    Else
                        ' Failure notification for if the Nightscout URL isn't configured
                        Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("Please configure the address to the Nightscout server before using this command."))
                    End If
                Else
                    ' Failure notification for if the user hasn't run the app yet
                    Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("Please configure the Nightscout Uploader app before using this command."))
                End If

                Call Differal.Complete()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Builds and executes a response to the user command to announce CGM data
    ''' </summary>
    ''' <param name="VoiceServiceConnection">An active connection to the Cortana service.</param>
    ''' <returns>Does not return anything.</returns>
    Private Async Function RespondToCGMVoiceCommand(ByVal VoiceServiceConnection As VoiceCommandServiceConnection) As Task
        ' Initialize the connection settings for the server and request CGM data
        Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
        Dim CurrentData As Nightscout.Server.NightscoutData = Await Server.GetCurrentData(Server.NightscoutDataTypes.CGM)

        ' Make sure the structure was completely loaded
        If CurrentData.NightscoutCGMData.ResultsLoaded Then
            ' Interpret the trend arrow for voice
            Dim Trend As String = ""
            Select Case CurrentData.NightscoutCGMData.TrendArrow
                Case "DOUBLEUP"
                    CurrentData.NightscoutCGMData.TrendArrow = "up very quickly"
                Case "SINGLEUP"
                    CurrentData.NightscoutCGMData.TrendArrow = "up quickly"
                Case "FORTYFIVEUP"
                    CurrentData.NightscoutCGMData.TrendArrow = "up"
                Case "FLAT"
                    CurrentData.NightscoutCGMData.TrendArrow = "flat"
                Case "FORTYFIVEDOWN"
                    CurrentData.NightscoutCGMData.TrendArrow = "low"
                Case "SINGLEDOWN"
                    CurrentData.NightscoutCGMData.TrendArrow = "low quickly"
                Case "DOUBLEDOWN"
                    CurrentData.NightscoutCGMData.TrendArrow = "low very quickly"
                Case Else
                    CurrentData.NightscoutCGMData.TrendArrow = "unknown"
            End Select

            ' State the direction the delta changed. Intentionally left out 0.
            Dim DeltaSign As String
            If CurrentData.NightscoutCGMData.DifferentialSGV > 0 Then
                DeltaSign = "positive"
            Else
                DeltaSign = "negative"
            End If

            Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("Your last sensor reading was " & InterpretEntryTime(CurrentData.NightscoutCGMData.UpdateDelta, False) & " ago at " & CurrentData.NightscoutCGMData.CurrentSGV & " mg/dl, a " & DeltaSign & " " & Math.Abs(CurrentData.NightscoutCGMData.DifferentialSGV) & " point difference from the last reading and trending " & CurrentData.NightscoutCGMData.TrendArrow & "."))
        Else
            Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The Nightscout server could not be contacted."))
        End If
    End Function

    ''' <summary>
    ''' Builds and executes a response to the user command for OpenAPS loop status.
    ''' </summary>
    ''' <param name="VoiceServiceConnection">An active connection to the Cortana service.</param>
    ''' <returns>Does not return anything.</returns>
    Private Async Function RespondToOpenAPSVoiceCommand(ByVal VoiceServiceConnection As VoiceCommandServiceConnection) As Task
        ' Initialize the connection settings for the server and request OpenAPS data
        Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
        Dim CurrentData As Nightscout.Server.NightscoutData = Await Server.GetCurrentData(Server.NightscoutDataTypes.OpenAPS)

        ' Make sure the structure was completely loaded
        If CurrentData.NightscoutOpenAPSData.ResultsLoaded Then
            ' If the data is more than 15 minutes old it doesn't represent the current state
            ' of the loop enough to be trustworthy without more information.
            If CurrentData.NightscoutOpenAPSData.UpdateDelta > 15 Then
                Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The OpenAPS data obtained from Nightscout is too out of date. Please consult Nightscout for more information."))
            Else
                If CurrentData.NightscoutOpenAPSData.Enacted Then
                    ' Full and complete enacted loop data
                    Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse(InterpretEntryTime(CurrentData.NightscoutOpenAPSData.UpdateDelta, False) & " ago the loop detected " & CurrentData.NightscoutOpenAPSData.IOB & " units on board at a blood glucose level of " & CurrentData.NightscoutOpenAPSData.CurrentSGV & " mg/dl. It set a temporary basal of " & CurrentData.NightscoutOpenAPSData.CurrentBasal & " units which will run for the next " & (CurrentData.NightscoutOpenAPSData.Duration - CurrentData.NightscoutOpenAPSData.UpdateDelta) & " minutes."))
                Else
                    ' This covers the case when the loop was last run but the calculation couldn't be enacted because of missing data.
                    Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("Loop data is available but incomplete. " & InterpretEntryTime(CurrentData.NightscoutOpenAPSData.UpdateDelta, False) & "  ago the loop detected " & CurrentData.NightscoutOpenAPSData.IOB & " units on board at a blood glucose level of " & CurrentData.NightscoutOpenAPSData.CurrentSGV & " mg/dl. It set a temporary basal of " & CurrentData.NightscoutOpenAPSData.CurrentBasal & " units which will run for the next " & (CurrentData.NightscoutOpenAPSData.Duration - CurrentData.NightscoutOpenAPSData.UpdateDelta) & " minutes."))
                End If
            End If
        Else
            Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The Nightscout server could not be contacted."))
        End If
    End Function

    ''' <summary>
    ''' Builds and executes a response to the user command for pump status.
    ''' </summary>
    ''' <param name="VoiceServiceConnection">An active connection to the Cortana service.</param>
    ''' <returns>Does not return anything.</returns>
    Private Async Function RespondToPumpVoiceCommand(ByVal VoiceServiceConnection As VoiceCommandServiceConnection) As Task
        ' Initialize the connection settings for the server and request OpenAPS data
        Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
        Dim CurrentData As Nightscout.Server.NightscoutData = Await Server.GetCurrentData(Server.NightscoutDataTypes.Pump)

        ' Make sure the structure was completely loaded
        If CurrentData.NightscoutPumpData.ResultsLoaded Then
            ' If the data is more than 15 minutes old it doesn't represent the current state
            ' of the loop enough to be trustworthy without more information.
            If CurrentData.NightscoutPumpData.UpdateDelta > 15 Then
                Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The pump data obtained from Nightscout is too out of date. Please consult Nightscout for more information."))
            Else
                ' Add current basal rate if available
                If CurrentData.NightscoutOpenAPSData.ResultsLoaded Then
                    Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("The current basal rate is " & CurrentData.NightscoutOpenAPSData.CurrentBasal & ". The pump reservoir contains " & CurrentData.NightscoutPumpData.ReservoirLevel & " units. The battery level is " & CurrentData.NightscoutPumpData.BatteryVoltage & " volts."))
                Else
                    Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("The pump reservoir contains " & CurrentData.NightscoutPumpData.ReservoirLevel & " units. The battery level is " & CurrentData.NightscoutPumpData.BatteryVoltage & " volts."))
                End If
            End If
        Else
            Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The Nightscout server could not be contacted."))
        End If
    End Function

    ''' <summary>
    ''' Builds and executes a response to the user command for the last food entry.
    ''' </summary>
    ''' <param name="VoiceServiceConnection"></param>
    ''' <returns></returns>
    Private Async Function RespondToLastEntryFoodVoiceCommand(ByVal VoiceServiceConnection As VoiceCommandServiceConnection) As Task
        ' Initialize the connection settings for the server and request last food entry data
        Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
        Dim CurrentData As Nightscout.Server.NightscoutData = Await Server.GetCurrentData(Server.NightscoutDataTypes.FoodEntry)

        ' Make sure the structure was completely loaded
        If CurrentData.NightscoutFoodTreatmentData.ResultsLoaded Then
            Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("The last food entry, " & InterpretEntryTime((DateTime.Now.ToLocalTime - CurrentData.NightscoutFoodTreatmentData.TreatmentTime.ToLocalTime).TotalMinutes, False) & " ago, consisted of " & CurrentData.NightscoutFoodTreatmentData.Carbohydrates & " carbohydrates and " & CurrentData.NightscoutFoodTreatmentData.Units & " units."))
        Else
            Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The Nightscout server could not be contacted."))
        End If
    End Function

    ''' <summary>
    ''' Builds and executes a response to the user command for the last bolus entry.
    ''' </summary>
    ''' <param name="VoiceServiceConnection">An active connection to the Cortana service.</param>
    ''' <returns>Does not return anything.</returns>
    Private Async Function RespondToLastEntryTreatmentVoiceCommand(ByVal VoiceServiceConnection As VoiceCommandServiceConnection) As Task
        ' Initialize the connection settings for the server and request last treatment entry data
        Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
        Dim CurrentData As Nightscout.Server.NightscoutData = Await Server.GetCurrentData(Server.NightscoutDataTypes.TreatmentEntry)

        ' Make sure the structure was completely loaded
        If CurrentData.NightscoutBolusTreatmentData.ResultsLoaded Then
            If CurrentData.NightscoutBolusTreatmentData.Carbohydrates > 0 Then
                Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("The last bolus, " & InterpretEntryTime((DateTime.Now.ToLocalTime - CurrentData.NightscoutBolusTreatmentData.TreatmentTime.ToLocalTime).TotalMinutes, False) & " ago, consisted of " & CurrentData.NightscoutBolusTreatmentData.Units & " units in addition to " & CurrentData.NightscoutBolusTreatmentData.Carbohydrates & " carbohydrates."))
            Else
                Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("The last bolus, " & InterpretEntryTime((DateTime.Now.ToLocalTime - CurrentData.NightscoutBolusTreatmentData.TreatmentTime.ToLocalTime).TotalMinutes, False) & " ago, consisted of " & CurrentData.NightscoutBolusTreatmentData.Units & " units."))
            End If
        Else
            Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The Nightscout server could not be contacted."))
        End If
    End Function

    ''' <summary>
    ''' Builds and executes a response to the user command for full and detailed status of all treatment and monitoring systems.
    ''' </summary>
    ''' <param name="VoiceServiceConnection">An active connection to the Cortana service.</param>
    ''' <returns>Does not return anything.</returns>
    Private Async Function RespondToFullSystemStatusVoiceCommand(ByVal VoiceServiceConnection As VoiceCommandServiceConnection) As Task
        ' Initialize the connection settings for the server and request all available data
        Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
        Dim CurrentData As Nightscout.Server.NightscoutData = Await Server.GetCurrentData(Server.NightscoutDataTypes.CGM Or Server.NightscoutDataTypes.OpenAPS Or Server.NightscoutDataTypes.Pump Or Server.NightscoutDataTypes.FoodEntry Or Server.NightscoutDataTypes.TreatmentEntry)

        Dim RetrievedData As String = ""
        Dim UnretrievedData As String = ""

        If CurrentData.NightscoutCGMData.ResultsLoaded Then
            ' Interpret the trend arrow for voice
            Dim Trend As String = ""
            Select Case CurrentData.NightscoutCGMData.TrendArrow
                Case "DOUBLEUP"
                    CurrentData.NightscoutCGMData.TrendArrow = "up very quickly"
                Case "SINGLEUP"
                    CurrentData.NightscoutCGMData.TrendArrow = "up quickly"
                Case "FORTYFIVEUP"
                    CurrentData.NightscoutCGMData.TrendArrow = "up"
                Case "FLAT"
                    CurrentData.NightscoutCGMData.TrendArrow = "flat"
                Case "FORTYFIVEDOWN"
                    CurrentData.NightscoutCGMData.TrendArrow = "low"
                Case "SINGLEDOWN"
                    CurrentData.NightscoutCGMData.TrendArrow = "low quickly"
                Case "DOUBLEDOWN"
                    CurrentData.NightscoutCGMData.TrendArrow = "low very quickly"
                Case Else
                    CurrentData.NightscoutCGMData.TrendArrow = "unknown"
            End Select

            ' State the direction the delta changed. Intentionally left out 0.
            Dim DeltaSign As String
            If CurrentData.NightscoutCGMData.DifferentialSGV > 0 Then
                DeltaSign = "positive"
            Else
                DeltaSign = "negative"
            End If

            ' Synthesize the returned data
            Dim Results As String = "Sensor: " & InterpretEntryTime(CurrentData.NightscoutCGMData.UpdateDelta, True) & " ago, " & CurrentData.NightscoutCGMData.CurrentSGV & " mg/dl, " & DeltaSign & " " & Math.Abs(CurrentData.NightscoutCGMData.DifferentialSGV) & " difference, trending " & CurrentData.NightscoutCGMData.TrendArrow & "."
            If RetrievedData = "" Then
                RetrievedData = Results
            Else
                RetrievedData &= " " & Results
            End If
        Else
            UnretrievedData = "CGM data"
        End If

        If CurrentData.NightscoutOpenAPSData.ResultsLoaded Then
            Dim Results As String = ""
            ' If the data is more than 15 minutes old it doesn't represent the current state
            ' of the loop enough to be trustworthy without more information.
            If CurrentData.NightscoutOpenAPSData.UpdateDelta > 15 Then
                Results = "OpenAPS data is too out of date."
            Else
                If CurrentData.NightscoutOpenAPSData.Enacted Then
                    ' Full and complete enacted loop data
                    Results = CurrentData.NightscoutOpenAPSData.IOB & " IOB " & InterpretEntryTime(CurrentData.NightscoutOpenAPSData.UpdateDelta, True) & " ago. Temp basal: " & CurrentData.NightscoutOpenAPSData.CurrentBasal & " units for " & (CurrentData.NightscoutOpenAPSData.Duration - CurrentData.NightscoutOpenAPSData.UpdateDelta) & " mins."
                Else
                    ' This covers the case when the loop was last run but the calculation couldn't be enacted because of missing data.
                    Results = "Loop incomplete: " & CurrentData.NightscoutOpenAPSData.IOB & " IOB " & InterpretEntryTime(CurrentData.NightscoutOpenAPSData.UpdateDelta, True) & " ago. Temp basal: " & CurrentData.NightscoutOpenAPSData.CurrentBasal & " units for " & (CurrentData.NightscoutOpenAPSData.Duration - CurrentData.NightscoutOpenAPSData.UpdateDelta) & " mins."
                End If
            End If

            ' Synthesize the returned data
            If RetrievedData = "" Then
                RetrievedData = Results
            Else
                RetrievedData &= " " & Results
            End If
        Else
            If UnretrievedData = "" Then
                UnretrievedData = "OpenAPS data"
            Else
                UnretrievedData &= ", OpenAPS data"
            End If
        End If

        If CurrentData.NightscoutFoodTreatmentData.ResultsLoaded Then
            ' Synthesize the returned data
            Dim Results As String = "Last food: " & InterpretEntryTime((DateTime.Now.ToLocalTime - CurrentData.NightscoutFoodTreatmentData.TreatmentTime.ToLocalTime).TotalMinutes, True) & " ago, " & CurrentData.NightscoutFoodTreatmentData.Carbohydrates & " carbs, " & CurrentData.NightscoutFoodTreatmentData.Units & " units."
            If RetrievedData = "" Then
                RetrievedData = Results
            Else
                RetrievedData &= " " & Results
            End If
        Else
            If UnretrievedData = "" Then
                UnretrievedData = "food data"
            Else
                UnretrievedData &= ", food data"
            End If
        End If

        If CurrentData.NightscoutBolusTreatmentData.ResultsLoaded Then
            ' Synthesize the returned data
            Dim Results As String = ""
            If CurrentData.NightscoutBolusTreatmentData.Carbohydrates > 0 Then
                Results = "Last bolus: " & InterpretEntryTime((DateTime.Now.ToLocalTime - CurrentData.NightscoutBolusTreatmentData.TreatmentTime.ToLocalTime).TotalMinutes, True) & " ago, " & CurrentData.NightscoutBolusTreatmentData.Units & " units, " & CurrentData.NightscoutBolusTreatmentData.Carbohydrates & " carbs."
            Else
                Results = "Last bolus: " & InterpretEntryTime((DateTime.Now.ToLocalTime - CurrentData.NightscoutBolusTreatmentData.TreatmentTime.ToLocalTime).TotalMinutes, True) & " ago, " & CurrentData.NightscoutBolusTreatmentData.Units & " units."
            End If

            If RetrievedData = "" Then
                RetrievedData = Results
            Else
                RetrievedData &= " " & Results
            End If
        Else
            If UnretrievedData = "" Then
                UnretrievedData = "bolus data"
            Else
                UnretrievedData &= ", bolus data"
            End If
        End If

        ' First make sure the status string isn't longer than what Cortana can handle
        If RetrievedData.Length > CortanaMaxDataLength Then
            Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("Nightscout returned too much data to detail all at once. Please ask seperate questions about different aspects of the system."))
        Else
            ' If there was no unretrivable data then speak the response. Otherwise speak the response along with which data was unretrievable.
            If UnretrievedData = "" Then
                Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse(RetrievedData))
            Else
                Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse(RetrievedData & " Unretrievable data: " & UnretrievedData))
            End If
        End If
    End Function

    ''' <summary>
    ''' Loads settings for Cortana command processing
    ''' </summary>
    Private Sub LoadSettings()
        Dim UseLocalSettings As Boolean = False
        Dim pLocalSettingsContainer As ApplicationDataContainer = Current.LocalSettings

        ' Attempt to see whether settings should be synced remotely or stored locally
        Try
            UseLocalSettings = CBool(pLocalSettingsContainer.Values(UseRoamingSettingsKey))
        Catch ex As Exception
            UseLocalSettings = False
        End Try

        ' Pick the right container
        Dim pSettingsContainer As ApplicationDataContainer
        If UseLocalSettings Then
            pSettingsContainer = Current.LocalSettings
        Else
            pSettingsContainer = Current.RoamingSettings
        End If

        ' Load the settings needed in the background service
        Try
            pFirstRunSetupDone = CBool(pSettingsContainer.Values(FirstRunSetupDoneKey))
        Catch ex As Exception
            pFirstRunSetupDone = False
        End Try

        Try
            pNightscoutURL = pSettingsContainer.Values(NightscoutURLKey)
        Catch ex As Exception
            pNightscoutURL = ""
        End Try

        Try
            pNightscoutAPIKey = pSettingsContainer.Values(NightscoutSecretKey)
        Catch ex As KeyNotFoundException
            pNightscoutAPIKey = ""
        End Try

        Try
            pUseSSL = CBool(pSettingsContainer.Values(UseSecureUploadConnectionKey))
        Catch ex As Exception
            pUseSSL = False
        End Try
    End Sub

    ''' <summary>
    ''' Takes an amount of minutes and interprets it into a string containing the number of hours and minutes represented.
    ''' </summary>
    ''' <param name="EntryTime">The number of minutes to interpret</param>
    ''' <param name="Shorten">Whether the returned string should use abreviations</param>
    ''' <returns></returns>
    Private Function InterpretEntryTime(ByVal EntryTime As Integer, ByVal Shorten As Boolean) As String
        Dim Interpretation As String = ""
        Dim Hours As Integer = 0
        Dim Minutes As Integer = 0

        ' See if more than an hours worth of minutes was passed
        If EntryTime > 60 Then
            ' Get the number of hours
            Hours = EntryTime \ 60

            ' Get the number of minutes less than the a full hour
            Minutes = EntryTime - (Hours * 60)
            If Shorten Then
                If Hours = 1 Then
                    Interpretation = Hours & " hour, "
                Else
                    Interpretation = Hours & " hours, "
                End If

                If Minutes = 1 Then
                    Interpretation &= Minutes & " min"
                Else
                    Interpretation &= Minutes & " mins"
                End If
            Else
                If Hours = 1 Then
                    Interpretation = Hours & " hour and "
                Else
                    Interpretation = Hours & " hours and "
                End If

                If Minutes = 1 Then
                    Interpretation &= Minutes & " minute"
                Else
                    Interpretation &= Minutes & " minutes"
                End If
            End If
        Else
            ' If the number of minutes passed is less than an hour then
            ' just report those minutes in the long or short form as
            ' indicated by Shorten.
            Minutes = EntryTime

            If Shorten Then
                If Minutes = 1 Then
                    Interpretation = Minutes & " min"
                Else
                    Interpretation = Minutes & " mins"
                End If
            Else
                If Minutes = 1 Then
                    Interpretation = Minutes & " minute"
                Else
                    Interpretation = Minutes & " minutes"
                End If
            End If
        End If

        Return Interpretation
    End Function

    ''' <summary>
    ''' Creates a visual and auditory response for Cortana.
    ''' </summary>
    ''' <param name="ResponseString">The string to display and speak</param>
    ''' <returns>A VoiceCommandResponse to be processed by Cortana in response to the user.</returns>
    Private Function CreateCortanaResponse(ByVal ResponseString As String) As VoiceCommandResponse
        Dim UserMessage As New VoiceCommandUserMessage
        UserMessage.DisplayMessage = ResponseString
        UserMessage.SpokenMessage = ResponseString

        Return VoiceCommandResponse.CreateResponse(UserMessage)
    End Function
End Class
