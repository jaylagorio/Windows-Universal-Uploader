''' <summary>
''' Author: Jay Lagorio
''' Date: June 5, 2016
''' Summary: MainPage serves as the primary display window for the app.
''' </summary>

Public NotInheritable Class MainPage
    Inherits Page

    ''' <summary>
    ''' Fires when the MainPage is loaded.
    ''' </summary>
    Private Async Sub MainPage_Loaded(sender As Object, e As RoutedEventArgs)
        ' Run the out of box experience if the app has never been run before. This
        ' gives the user the opportunity to configure the Nightscout URL and API Secret
        ' as well as add sync devices.
        If Not Settings.FirstRunSetupDone Then
            Await RunOobe()
        End If

        ' Depending on settings entered by the user initialize the main window layout
        Call InitializeLayout()

        ' Add event handler and start the timer for Synchronizer
        AddHandler Synchronizer.SynchronizationStarted, AddressOf SynchronizationStatusChanged
        AddHandler Synchronizer.SynchronizationStopped, AddressOf SynchronizationStatusChanged
        Call Synchronizer.StartTimer()
    End Sub

    ''' <summary>
    ''' Fires when the MainPage gets focus, mainly serves to set the CommandBar button state.
    ''' </summary>
    Private Sub MainPage_GotFocus(sender As Object, e As RoutedEventArgs)
        Call SetDeviceButtonState()
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
        Call Synchronizer.Initialize(Me, {prgTopSyncing, prgBottomSyncing}, {lblTopSyncStatus, lblBottomSyncStatus}, Settings.LastSyncTime)

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

            CenterWebView.Navigate(New Uri("http://" & Settings.NightscoutURL, UriKind.Absolute))
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
    ''' Runs the Out-of-Box Experience. This amounts to showing extra labeling on the
    ''' Settings and New Device windows and displaying those windows in order.
    ''' </summary>
    ''' <returns>No value returned</returns>
    Private Async Function RunOobe() As Task
        ' Show the Settings page
        Dim WelcomeSettingsPage As New SettingsPage
        Call WelcomeSettingsPage.ShowWelcomeMode()
        Await WelcomeSettingsPage.ShowAsync

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

        Call CenterWebView.Focus(FocusState.Pointer)
        Call Me.Focus(FocusState.Pointer)

        Return
    End Function

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
    Private Async Sub cmdSettings_Click(sender As Object, e As RoutedEventArgs) Handles cmdPrimarySettings.Click, cmdSecondarySettings.Click
        Call Synchronizer.StopTimer()

        Dim PreviousUploadInterval As Integer = Settings.UploadInterval
        ' Show the SettingsListPage to change settings, then reset the upload interval if it changed
        Await (New SettingsPage).ShowAsync

        ' If the UploadInterval changed, recalibrate the timer
        If PreviousUploadInterval <> Settings.UploadInterval Then
            Call Synchronizer.StopTimer()
            Call Synchronizer.StartTimer()
        End If

        Call SetDeviceButtonState()
        Call Synchronizer.StartTimer()
    End Sub

    ''' <summary>
    ''' Determine the enabled state of the CommandBar buttons.
    ''' </summary>
    Private Sub SetDeviceButtonState()
        ' Only enable the Device List button if there are devices enrolled
        cmdAddDevice.IsEnabled = Not Synchronizer.IsSynchronizing
        cmdDeviceList.IsEnabled = (Settings.EnrolledDevices.Count > 0) And (Not Synchronizer.IsSynchronizing)
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
        ' not syncing right now
        If Settings.NightscoutAPIKey <> "" And Settings.NightscoutURL <> "" Then
            cmdSyncNow.IsEnabled = (Settings.EnrolledDevices.Count > 0) And (Not Synchronizer.IsSynchronizing)
        Else
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
End Class
