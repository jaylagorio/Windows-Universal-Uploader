Imports Windows.UI.Popups

''' <summary>
''' Author: Jay Lagorio
''' Date: May 22, 2016
''' Summary: Shows the user a list of currently enrolled devices and allows them to remove any that are no longer needed.
''' </summary>

Public NotInheritable Class DeviceListPage
    Inherits ContentDialog

    ''' <summary>
    ''' Causes the list of devices to be loaded when the dialog opens.
    ''' </summary>
    Private Sub ContentDialog_Loaded(sender As Object, e As RoutedEventArgs)
        Call ListEnrolledDevices()
    End Sub

    ''' <summary>
    ''' Closes the dialog.
    ''' </summary>
    Private Sub ContentDialog_PrimaryButtonClick(sender As ContentDialog, args As ContentDialogButtonClickEventArgs)
        Me.Hide()
    End Sub

    ''' <summary>
    ''' Sets the listbox to the list of enrolled devices in Settings.
    ''' </summary>
    Private Sub ListEnrolledDevices()
        lstDevices.DataContext = Nothing
        Dim DeviceList As Collection(Of Device) = Settings.EnrolledDevices
        lstDevices.DataContext = DeviceList
    End Sub

    ''' <summary>
    ''' Removes a device from the sync process.
    ''' </summary>
    Private Async Sub RemoveButton_Click(sender As Object, e As RoutedEventArgs)
        Dim ButtonPressed As Button = sender
        Dim SerialNumber As String = ButtonPressed.Tag

        ' Ask the user to make sure they want to get rid of the device
        Dim ConfirmDialog As New MessageDialog("Are you sure you want to remove the device with serial number " & SerialNumber & "?", "Nightscout")
        ConfirmDialog.Commands.Add(New UICommand("No"))
        ConfirmDialog.Commands.Add(New UICommand("Yes"))
        ConfirmDialog.CancelCommandIndex = 0
        ConfirmDialog.DefaultCommandIndex = 1
        Dim ConfirmDialogResult As UICommand = Await ConfirmDialog.ShowAsync()

        ' Find the device in the list and remove it
        If ConfirmDialogResult.Label = "Yes" Then
            Dim DeviceList As Collection(Of Device) = Settings.EnrolledDevices
            For i = 0 To DeviceList.Count - 1
                If DeviceList(i).SerialNumber = SerialNumber Then
                    Settings.RemoveEnrolledDevice(DeviceList(i))
                    Call ListEnrolledDevices()
                    Exit For
                End If
            Next
        End If

        ' If there aren't any more devices to remove then close the dialog
        If Settings.EnrolledDevices.Count = 0 Then
            Me.Hide()
        End If
    End Sub
End Class
