Imports Microsoft.Band
Imports Microsoft.Band.Tiles
Imports Windows_Universal_Uploader_Background

''' <summary>
''' Author: Jay Lagorio
''' Date: March 19, 2017
''' Summary: Allows the user to change app settings.
''' </summary>

Public NotInheritable Class SettingsPage
    Inherits Page

    ' Indicates we're in OOBE mode
    Private pWelcomeMode As Boolean = False

    Private Const NightscoutAPITextBoxPlaceholder As String = ""
    Private Const NightscoutAPITextBoxObfuscated As String = "**********"

    ''' <summary>
    ''' Loads the dialog with preexisting settings values if set.
    ''' </summary>
    Private Sub SettingsPages_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If Settings.NightscoutURL <> "" Then
            txtNightscoutDomain.Text = Settings.NightscoutURL
            txtNightscoutDomain.FontStyle = Windows.UI.Text.FontStyle.Normal
        End If

        If Settings.NightscoutAPIKey <> "" Then
            txtNightscoutSecret.Text = NightscoutAPITextBoxObfuscated
            txtNightscoutSecret.FontStyle = Windows.UI.Text.FontStyle.Normal
        End If

        Select Case Settings.UploadInterval
            Case 0
                lstSyncInterval.SelectedIndex = 0
            Case 1
                lstSyncInterval.SelectedIndex = 1
            Case 3
                lstSyncInterval.SelectedIndex = 2
            Case 5
                lstSyncInterval.SelectedIndex = 3
            Case 10
                lstSyncInterval.SelectedIndex = 4
            Case 15
                lstSyncInterval.SelectedIndex = 5
            Case 20
                lstSyncInterval.SelectedIndex = 6
            Case Else
                lstSyncInterval.SelectedIndex = 0
        End Select

        lstScreenBehavior.SelectedIndex = CInt(Settings.ScreenSleepBehavior)

        chkSecureConnection.IsOn = Settings.UseSecureUploadConnection
        chkUseAudibleAlarms.IsOn = Settings.UseAudibleAlarms
        chkAlwaysShowLastSyncTime.IsOn = Settings.AlwaysShowLastSyncTime
        chkUseRoamingSettings.IsOn = Settings.UseRoamingSettings
    End Sub

    Private Sub SettingsPages_KeyUp(sender As Object, e As KeyRoutedEventArgs) Handles Me.KeyUp
        If e.Key = Windows.System.VirtualKey.Enter Then
            Call Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, Sub()
                                                                                        cmdOK_Click(Nothing, Nothing)
                                                                                    End Sub)
        ElseIf e.Key = Windows.System.VirtualKey.Escape Then
            Call Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, Sub()
                                                                                        cmdCancel_Click(Nothing, Nothing)
                                                                                    End Sub)
        End If
    End Sub

    ''' <summary>
    ''' Allows the calling page to specify whether to use Welcome Mode.
    ''' </summary>
    ''' <param name="e">Passed by the navigation manager.</param>
    Protected Overrides Sub OnNavigatedTo(e As NavigationEventArgs)
        MyBase.OnNavigatedTo(e)

        ' Check that a parameter was passed
        If Not e.Parameter Is Nothing Then
            ' Check that the parameter was True
            If CBool(e.Parameter) Then
                Call ShowWelcomeMode()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Fires when the user clicks the OK button. Checks the settings the user indicated for validity and stores the new settings.
    ''' </summary>
    Private Async Sub cmdOK_Click(sender As Object, e As RoutedEventArgs) Handles cmdOK.Click
        'Check to make sure the domain name wasn't left blank
        If txtNightscoutDomain.Text.Trim() = "" Or txtNightscoutDomain.Text = "yoursitehere.azurewebsites.net" Then
            Await (New Windows.UI.Popups.MessageDialog("Please enter the domain name for your Nightscout host.", "Nightscout")).ShowAsync
            Return
        End If

        ' Check to make sure the API Secret wasn't left blank if we're not in welcome more
        If Not pWelcomeMode Then
            If txtNightscoutSecret.Text.Trim() = "" Or txtNightscoutSecret.Text = "YOURAPISECRET" Then
                Await (New Windows.UI.Popups.MessageDialog("Please enter the API Secret for your Nightscout host.", "Nightscout")).ShowAsync
                Return
            End If
        End If

        ' Settings were valid, save and close. Watch out for the API Secret in case it was loaded as obfuscated text.
        Settings.NightscoutURL = txtNightscoutDomain.Text
        If txtNightscoutSecret.Text.Trim() <> "" And txtNightscoutSecret.Text <> NightscoutAPITextBoxObfuscated And txtNightscoutSecret.Text <> NightscoutAPITextBoxPlaceholder Then
            Settings.NightscoutAPIKey = txtNightscoutSecret.Text
        End If
        Select Case lstSyncInterval.SelectedIndex
            Case 0
                Settings.UploadInterval = 0
            Case 1
                Settings.UploadInterval = 1
            Case 2
                Settings.UploadInterval = 3
            Case 3
                Settings.UploadInterval = 5
            Case 4
                Settings.UploadInterval = 10
            Case 5
                Settings.UploadInterval = 15
            Case 6
                Settings.UploadInterval = 20
        End Select

        Settings.UseSecureUploadConnection = chkSecureConnection.IsOn
        Settings.AlwaysShowLastSyncTime = chkAlwaysShowLastSyncTime.IsOn
        Settings.UseAudibleAlarms = chkUseAudibleAlarms.IsOn
        Settings.UseRoamingSettings = chkUseRoamingSettings.IsOn
        Settings.ScreenSleepBehavior = CInt(lstScreenBehavior.SelectedIndex)
        Settings.FirstRunSettingsSaved = True

        Call Frame.GoBack()
    End Sub

    ''' <summary>
    ''' Closes the dialog without saving any changes.
    ''' </summary>
    Private Sub cmdCancel_Click(sender As Object, e As RoutedEventArgs) Handles cmdCancel.Click
        Call Frame.GoBack()
    End Sub

    ''' <summary>
    ''' Fires when the Nightscout Domain TextBox receives focus to remove the background label and normalize formatting.
    ''' </summary>
    Private Sub txtNightscoutDomain_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtNightscoutDomain.GotFocus
        If txtNightscoutDomain.Text = "yoursitehere.azurewebsites.net" Then
            txtNightscoutDomain.Text = ""
            txtNightscoutDomain.SelectionStart = 0
            txtNightscoutDomain.FontStyle = Windows.UI.Text.FontStyle.Normal
        End If
    End Sub

    ''' <summary>
    ''' Fires when the Nightscout Domain TextBox loses focus to readd the background label and formatting if the box is empty.
    ''' </summary>
    Private Sub txtNightscoutDomain_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtNightscoutDomain.LostFocus
        If txtNightscoutDomain.Text = "" Then
            txtNightscoutDomain.Text = "yoursitehere.azurewebsites.net"
            txtNightscoutDomain.FontStyle = Windows.UI.Text.FontStyle.Italic
        End If
    End Sub

    ''' <summary>
    ''' Fires when the Nightscout Secret TextBox receives focus to remove the background label and normalize formatting.
    ''' </summary>
    Private Sub txtNightscoutSecret_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtNightscoutSecret.GotFocus
        If txtNightscoutSecret.Text = "YOURAPISECRET" Then
            txtNightscoutSecret.Text = ""
            txtNightscoutSecret.SelectionStart = 0
            txtNightscoutSecret.FontStyle = Windows.UI.Text.FontStyle.Normal
        End If
    End Sub

    ''' <summary>
    ''' Fires when the Nightscout Secret TextBox loses focus to readd the background label and formatting if the box is empty.
    ''' </summary>
    Private Sub txtNightscoutSecret_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtNightscoutSecret.LostFocus
        If txtNightscoutSecret.Text = "" Then
            txtNightscoutSecret.Text = "YOURAPISECRET"
            txtNightscoutSecret.FontStyle = Windows.UI.Text.FontStyle.Italic
        End If
    End Sub

    ''' <summary>
    ''' Makes a TextBlock with a welcome message visible to the user and hides the Cancel button.
    ''' </summary>
    Public Sub ShowWelcomeMode()
        pWelcomeMode = True
        lblWelcomeText.Visibility = Visibility.Visible
        lblSettings.Visibility = Visibility.Collapsed
        cmdCancel.Visibility = Visibility.Collapsed
    End Sub
End Class
