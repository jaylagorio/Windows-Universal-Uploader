Imports Windows.System.Display
Imports Windows.Devices.Power

''' <summary>
''' Author: Jay Lagorio
''' Date: November 6, 2016
''' Summary: MainPage serves as the primary display window for the app.
''' </summary>

Public NotInheritable Class MainPage
    Inherits Page

    ' Supports screen behavior for when the app is in the foreground
    Dim pDisplayRequest As New DisplayRequest
    Dim pDisplayRequestActive As Boolean = False
    Dim pDisplayACStatus As Boolean = False

    ' Used to see if the upload interval changes between showings of the Settings window
    Dim pPreviousUploadInterval As Integer

    ''' <summary>
    ''' Fires when the MainPage is loaded.
    ''' </summary>
    Private Async Sub MainPage_Loaded(sender As Object, e As RoutedEventArgs)
        ' Run the out of box experience if the app has never been run before. This
        ' gives the user the opportunity to configure the Nightscout URL and API Secret
        ' as well as add sync devices.
        If Not Settings.FirstRunSetupDone Then
            ' Check to see if at least one round in the Settings window has occured
            If Not Settings.FirstRunSettingsSaved Then
                ' Show the Settings page, pass True to activate Welcome Mode
                Frame.Navigate(GetType(SettingsPage), True)
            Else
                ' If the first round of the Settings window resulted in saved settings then
                ' continue on the OOBE and prompt the user to enroll a device

                ' Set the first run setting
                Settings.FirstRunSetupDone = True

                ' If the user didn't enter a Nightscout API key they can't sync
                ' so we won't show them the enrollment page
                If Settings.NightscoutAPIKey = "" Then
                    Await (New Windows.UI.Popups.MessageDialog("Before enrolling devices to upload your data to Nightscout you will need to enter your API Secret. Return to the Settings window when you're ready to do that and then add your devices.", "Nightscout API Secret")).ShowAsync
                Else
                    ' Show the enrollment page
                    Dim WelcomeEnrollPage As New EnrollDevicePage
                    Call WelcomeEnrollPage.ShowWelcomeMode()
                    Await WelcomeEnrollPage.ShowAsync
                    Call SetDeviceButtonState()
                End If
            End If
        End If

        ' Initialize the main window and the synchronizer only if the entire First Run process has been completed.
        If Settings.FirstRunSetupDone And Settings.FirstRunSettingsSaved Then
            ' Depending on settings entered by the user initialize the main window layout
            Call InitializeLayout()

            ' Load Cortana voice commands
            Await InitializeCortana()

            ' Add event handler and start the timer for Synchronizer
            AddHandler Synchronizer.SynchronizationStarted, AddressOf SynchronizationStatusChanged
            AddHandler Synchronizer.SynchronizationStopped, AddressOf SynchronizationStatusChanged

            ' Add event handler to track AC power vs. battery power for screen activity and call
            ' the screen configuration function
            AddHandler Battery.AggregateBattery.ReportUpdated, AddressOf OnBatteryReportUpdated
            Call SetScreenBehavior()

            ' Start the timer for the synchronizer, if there is one
            Call Synchronizer.StartTimer()

            Call CenterWebView.Focus(FocusState.Pointer)
            Call Me.Focus(FocusState.Pointer)
        End If
    End Sub

    ''' <summary>
    ''' Fires when the MainPage gets focus, mainly serves to set the CommandBar button state.
    ''' </summary>
    Private Sub MainPage_GotFocus(sender As Object, e As RoutedEventArgs)
        ' If the UploadInterval changed because the user went to the SettingsPage, recalibrate the timer
        If pPreviousUploadInterval <> Settings.UploadInterval Then
            Call Synchronizer.StopTimer()
        End If

        Call SetDeviceButtonState()
        Call Synchronizer.StartTimer()
    End Sub

    ''' <summary>
    ''' Fires when MainPage loses focus, mainly serves to set the CommandBar button state.
    ''' </summary>
    Private Sub MainPage_LostFocus(sender As Object, e As RoutedEventArgs)
        Call SetDeviceButtonState()
    End Sub

    ''' <summary>
    ''' Fires when MainPage is tapped, mainly serves to set the CommandBar button state.
    ''' </summary>
    Private Sub MainPage_Tapped(sender As Object, e As TappedRoutedEventArgs)
        Call SetDeviceButtonState()
    End Sub

    ''' <summary>
    ''' Initializes the layout of the MainPage depending on what settings the user has entered.
    ''' </summary>
    Private Sub InitializeLayout()
        ' Initialize the synchronization engine so it can find the UI thread and the sync progress controls
        Call Synchronizer.Initialize(Me, {prgTopSyncing, prgBottomSyncing}, {lblTopSyncStatus, lblBottomSyncStatus}, Settings.LastSyncTime, Settings.NightscoutURL, Settings.UseSecureUploadConnection, Settings.NightscoutAPIKey)

        ' Set the CommandBar button state based on settings
        Call SetDeviceButtonState()

        If Settings.NightscoutURL <> "" Then
            ' The Nightscout URL being entered means the prompt images aren't
            ' necessary. Collapse them and show the center WebView, then navigate to
            ' the Nightscout site.
            PlaceHolderImage.Visibility = Visibility.Collapsed
            SettingsDirectionImage.Visibility = Visibility.Collapsed
            CenterWebView.HorizontalAlignment = HorizontalAlignment.Stretch
            CenterWebView.HorizontalAlignment = HorizontalAlignment.Stretch
            CenterWebView.Visibility = Visibility.Visible

            If Settings.UseSecureUploadConnection Then
                Call CenterWebView.Navigate(New Uri("https://" & Settings.NightscoutURL, UriKind.Absolute))
            Else
                Call CenterWebView.Navigate(New Uri("http://" & Settings.NightscoutURL, UriKind.Absolute))
            End If
        Else
            ' If the Nightscout URL hasn't been enetered the primary WebView is pretty
            ' useless. Collapse the WebView and show some image assets to prompt
            ' the user to configure the app.
            PlaceHolderImage.Visibility = Visibility.Visible
            SettingsDirectionImage.Visibility = Visibility.Visible
            CenterWebView.Visibility = Visibility.Collapsed
        End If
    End Sub

    ''' <summary>
    ''' Register the latest version of the voice commands VCD to Cortana
    ''' </summary>
    ''' <returns>Nothing</returns>
    Private Async Function InitializeCortana() As Task
        Await VoiceCommands.VoiceCommandDefinitionManager.InstallCommandDefinitionsFromStorageFileAsync(
            Await Windows.Storage.StorageFile.GetFileFromApplicationUriAsync(New Uri("ms-appx:///VoiceCommands.xml"))
        )
    End Function

    Protected Overrides Sub OnNavigatedTo(e As NavigationEventArgs)
    End Sub

    ''' <summary>
    ''' Requests or releases control of whether the screen goes to sleep at its normal timeout
    ''' depending on the user setting and whether the device is plugged into AC power or not.
    ''' </summary>
    Private Sub SetScreenBehavior()
        Select Case Settings.ScreenSleepBehavior
            Case Settings.ScreenBehavior.NormalScreenBehavior
                If pDisplayRequestActive Then
                    Call pDisplayRequest.RequestRelease()
                End If
            Case Settings.ScreenBehavior.KeepScreenOnWhenPluggedIn
                If pDisplayACStatus Then
                    If Not pDisplayRequestActive Then
                        ' Device is plugged in and we don't have a request - get one
                        Call pDisplayRequest.RequestActive()
                    Else
                        ' Device is plugged in and we have a request - do nothing
                    End If
                Else
                    If Not pDisplayRequestActive Then
                        ' Device is not plugged in and we don't have a request - do nothing
                    Else
                        ' Device is not plugged in and we have a request - release it
                        Call pDisplayRequest.RequestRelease()
                    End If
                End If
            Case Settings.ScreenBehavior.KeepScreenOnAlways
                If Not pDisplayRequestActive Then
                    Call pDisplayRequest.RequestActive()
                End If
        End Select
    End Sub

    ''' <summary>
    ''' The Add Device button in the CommandBar.
    ''' </summary>
    Private Async Sub cmdAddDevice_Click(sender As Object, e As RoutedEventArgs) Handles cmdAddDevice.Click
        Call Synchronizer.StopTimer()

        If Settings.NightscoutAPIKey = "" Then
            Await (New Windows.UI.Popups.MessageDialog("Please configure your API Secret before enrolling devices. Devices will not be able to sync with Nightscout without this setting.")).ShowAsync
        Else
            ' Show the window and if the user clicked OK set the CommandBar button state
            If Await (New EnrollDevicePage).ShowAsync() = ContentDialogResult.None Or ContentDialogResult.Primary Or ContentDialogResult.Secondary Then
                Call CenterWebView.Focus(FocusState.Pointer)
                Call Me.Focus(FocusState.Pointer)
                Call SetDeviceButtonState()
            End If
        End If

        Call Synchronizer.StartTimer()
    End Sub

    ''' <summary>
    ''' The Device List button in the CommandBar.
    ''' </summary>
    Private Async Sub cmdDeviceList_Click(sender As Object, e As RoutedEventArgs) Handles cmdDeviceList.Click
        Call Synchronizer.StopTimer()

        ' Show the DeviceListPage to view or remove sync devices
        Await (New DeviceListPage).ShowAsync

        Call SetDeviceButtonState()
        Call Synchronizer.StartTimer()
    End Sub

    ''' <summary>
    ''' The Stop Sync button in the CommandBar.
    ''' </summary>
    Private Async Sub cmdCancelSync_Click(sender As Object, e As RoutedEventArgs) Handles cmdCancelSync.Click
        ' Only let the button be pressed once
        cmdCancelSync.IsEnabled = False

        ' Cancel the sync
        Await Synchronizer.CancelSync()
    End Sub

    ''' <summary>
    ''' The Synchronize button in the CommandBar.
    ''' </summary>
    Private Async Sub cmdSyncNow_Click(sender As Object, e As RoutedEventArgs) Handles cmdSyncNow.Click
        Dim ErrorOccurred As Boolean = False

        ' Kick the synchronization engine into action manually but on a different thread
        Try
            Await Task.Run(Sub()
                               Call Synchronizer.Synchronize()
                           End Sub)
        Catch ex As Exception
            Call Synchronizer.ResetSyncStatus()
            Call SetDeviceButtonState()
            ErrorOccurred = True
        End Try

        ' Display an error message when an error occurs
        If ErrorOccurred Then
            Await (New Windows.UI.Popups.MessageDialog("An error occurred while synchronizing with Nightscout. Please try again.", "Nightscout")).ShowAsync()
        End If
    End Sub

    ''' <summary>
    ''' The Settings button in the CommandBar.
    ''' </summary>
    Private Sub cmdSettings_Click(sender As Object, e As RoutedEventArgs) Handles cmdPrimarySettings.Click, cmdSecondarySettings.Click
        Call Synchronizer.StopTimer()

        'Save the current upload interval to see if it changes later
        pPreviousUploadInterval = Settings.UploadInterval

        ' Show the SettingsPage to change settings
        Frame.Navigate(GetType(SettingsPage))
    End Sub

    ''' <summary>
    ''' Allows the user to refresh the Nightscout web view
    ''' </summary>
    Private Sub cmdRefreshNightscout_Tapped(sender As Object, e As TappedRoutedEventArgs) Handles cmdRefreshNightscout.Tapped
        If Settings.UseSecureUploadConnection Then
            Call CenterWebView.Navigate(New Uri("https://" & Settings.NightscoutURL, UriKind.Absolute))
            Call CenterWebView.Refresh()
        Else
            Call CenterWebView.Navigate(New Uri("http://" & Settings.NightscoutURL, UriKind.Absolute))
            Call CenterWebView.Refresh()
        End If
    End Sub

    ''' <summary>
    ''' Determine the enabled state of the CommandBar buttons.
    ''' </summary>
    Private Sub SetDeviceButtonState()
        ' Only enable the Device List button if there are devices enrolled and we're not syncing
        cmdDeviceList.IsEnabled = (Settings.EnrolledDevices.Count > 0) And (Not Synchronizer.IsSynchronizing)

        cmdAddDevice.IsEnabled = Not Synchronizer.IsSynchronizing
        cmdPrimarySettings.IsEnabled = Not Synchronizer.IsSynchronizing
        cmdSecondarySettings.IsEnabled = Not Synchronizer.IsSynchronizing

        ' Alternate Sync Now/Cancel Sync button visibility
        If Synchronizer.IsSynchronizing Then
            cmdSyncNow.Visibility = Visibility.Collapsed
            cmdCancelSync.Visibility = Visibility.Visible
        Else
            cmdSyncNow.Visibility = Visibility.Visible
            cmdCancelSync.Visibility = Visibility.Collapsed
        End If

        ' Only enable the Sync button if there are devices enrolled and we're
        ' not syncing right now. Only show the Refresh Nightscout option if the
        ' address has been set.
        If Settings.NightscoutAPIKey <> "" And Settings.NightscoutURL <> "" Then
            cmdRefreshNightscout.Visibility = Visibility.Visible
            cmdSyncNow.IsEnabled = (Settings.EnrolledDevices.Count > 0) And (Not Synchronizer.IsSynchronizing)
        Else
            If Settings.NightscoutURL <> "" Then
                cmdRefreshNightscout.Visibility = Visibility.Visible
            Else
                cmdRefreshNightscout.Visibility = Visibility.Collapsed
            End If
            cmdSyncNow.IsEnabled = False
        End If
    End Sub

    ''' <summary>
    ''' Event gets called when the synchronization process starts or stops. It sets
    ''' the state of the buttons to be enabled or disabled in relation to sync state.
    ''' </summary>
    Private Async Sub SynchronizationStatusChanged()
        Await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, AddressOf SetDeviceButtonState)
    End Sub

    ''' <summary>
    ''' Event gets called when the user unplugs or plugs in their mobile device and changes
    ''' the behavior of the screen when the app is in the foreground based on that setting.
    ''' </summary>
    ''' <param name="sender">Sent with the event</param>
    ''' <param name="args">Sent with the event</param>
    Private Async Sub OnBatteryReportUpdated(sender As Windows.Devices.Power.Battery, args As Object)
        Dim Report As BatteryReport = sender.GetReport()

        ' Get a battery report and see if we're on battery. If the battery is discharging then we
        ' can safely say we're not plugged in, but if there isn't a battery, it's idle, or it's charging
        ' then we can assume we're plugged in.
        If Report.Status = Windows.System.Power.BatteryStatus.Discharging Then
            pDisplayACStatus = False
        Else
            pDisplayACStatus = True
        End If

        ' Set the screen behavior accordingly using the main UI thread
        Await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, Sub()
                                                                                     Call SetScreenBehavior()
                                                                                 End Sub)
    End Sub
End Class
