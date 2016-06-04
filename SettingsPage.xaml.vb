''' <summary>
''' Author: Jay Lagorio
''' Date: June 5, 2016
''' Summary: Allows the user to change the Nighscout host, API Secret, and sync timer interval.
''' </summary>

Public NotInheritable Class SettingsPage
    Inherits ContentDialog

    ' Indicates we're in OOBE mode
    Private pWelcomeMode As Boolean = False

    Private Const NightscoutAPITextBoxPlaceholder As String = ""
    Private Const NightscoutAPITextBoxObfuscated As String = "**********"

    ''' <summary>
    ''' Loads the dialog with preexisting settings values if set.
    ''' </summary>
    Private Sub ContentDialog_Loaded(sender As Object, e As RoutedEventArgs)
        If Settings.NightscoutURL <> "" Then
            txtNightscoutDomain.Text = Settings.NightscoutURL
            txtNightscoutDomain.FontStyle = Windows.UI.Text.FontStyle.Normal
        End If

        If Settings.NightscoutAPIKey <> "" Then
            txtNightscoutSecret.Text = NightscoutAPITextBoxObfuscated
            txtNightscoutSecret.FontStyle = Windows.UI.Text.FontStyle.Normal
        End If

        For i = 0 To lstSyncInterval.Items.Count - 1
            If lstSyncInterval.Items(i).Tag = Settings.UploadInterval Then
                lstSyncInterval.SelectedIndex = i
            End If
        Next

        chkSecureConnection.IsChecked = Settings.UseSecureUploadConnection
    End Sub

    ''' <summary>
    ''' Fires when the user clicks the OK button. Checks the settings the user indicated for validity and stores the new settings.
    ''' </summary>
    Private Async Sub ContentDialog_PrimaryButtonClick(sender As ContentDialog, args As ContentDialogButtonClickEventArgs)
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
        Settings.UploadInterval = lstSyncInterval.SelectedItem.Tag
        Settings.UseSecureUploadConnection = chkSecureConnection.IsChecked

        Call Me.Hide()
    End Sub

    ''' <summary>
    ''' Closes the dialog without saving any changes.
    ''' </summary>
    Private Sub ContentDialog_SecondaryButtonClick(sender As ContentDialog, args As ContentDialogButtonClickEventArgs)
        Call Me.Hide()
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
    ''' Makes a TextBlock with a welcome message visible to the user.
    ''' </summary>
    Public Sub ShowWelcomeMode()
        pWelcomeMode = True
        lblWelcomeText.Visibility = Visibility.Visible
    End Sub
End Class
