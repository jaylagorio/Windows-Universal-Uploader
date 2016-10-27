Imports Nightscout
Imports Windows.Storage
Imports Windows.Storage.ApplicationData
Imports Windows.ApplicationModel.AppService
Imports Windows.ApplicationModel.Background
Imports Windows.ApplicationModel.VoiceCommands

''' <summary>
''' Author: Jay Lagorio
''' Date: October 30, 2016
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
                            Case "getDashboard"
                                Dim Server As New Nightscout.Server(pNightscoutURL, pUseSSL, pNightscoutAPIKey)
                                Dim Dashboard As Nightscout.Server.NightscoutDashboardData = Await Server.GetDashboardData()

                                ' Make sure the structure was completely loaded
                                If Dashboard.ResultsLoaded Then
                                    ' Interpret the trend arrow for voice
                                    Dim Trend As String = ""
                                    Select Case Dashboard.TrendArrow
                                        Case "DOUBLEUP"
                                            Dashboard.TrendArrow = "up very quickly"
                                        Case "SINGLEUP"
                                            Dashboard.TrendArrow = "up quickly"
                                        Case "FORTYFIVEUP"
                                            Dashboard.TrendArrow = "up"
                                        Case "FLAT"
                                            Dashboard.TrendArrow = "flat"
                                        Case "FORTYFIVEDOWN"
                                            Dashboard.TrendArrow = "low"
                                        Case "SINGLEDOWN"
                                            Dashboard.TrendArrow = "low quickly"
                                        Case "DOUBLEDOWN"
                                            Dashboard.TrendArrow = "low very quickly"
                                        Case Else
                                            Dashboard.TrendArrow = "unknown"
                                    End Select

                                    Call VoiceServiceConnection.ReportSuccessAsync(CreateCortanaResponse("Your last sensor reading was " & Dashboard.CGMUpdateDelta & " minutes ago at " & Dashboard.CurrentSGV & " mg/dl, a " & Dashboard.DifferentialSGV & " point difference from the last reading and trending " & Dashboard.TrendArrow & "."))
                                Else
                                    Call VoiceServiceConnection.ReportFailureAsync(CreateCortanaResponse("The Nightscout server could not be contacted."))
                                End If
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
