Imports Windows.UI.Popups
Imports Microsoft.ApplicationInsights

''' <summary>
''' Author: Jay Lagorio
''' Date: June 12, 2016
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
                    If Not Await DeviceList(i).IsConnected Then
                        If Await DeviceList(i).Connect() Then
                            Call ReportDeviceRemovalTelemetry(DeviceList(i))
                            Await DeviceList(i).Disconnect()
                        End If
                    Else
                        Call ReportDeviceRemovalTelemetry(DeviceList(i))
                    End If

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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="NewDevice"></param>
    Private Sub ReportDeviceRemovalTelemetry(ByRef NewDevice As DexcomDevice)
        Dim ApplicationInsights As New TelemetryClient
        Dim EventProperties As New Dictionary(Of String, String)

        ' Gather device firmware data. The first call will cause the device to be queried.
        Call EventProperties.Add("SerialNumber", NewDevice.SerialNumber)
        Call EventProperties.Add("InterfaceName", NewDevice.InterfaceName)
        If Not NewDevice.DexcomReceiver Is Nothing Then
            Call EventProperties.Add("SchemaVersion", NewDevice.DexcomReceiver.SchemaVersion)
            Call EventProperties.Add("ProductID", NewDevice.DexcomReceiver.ProductId)
            Call EventProperties.Add("ProductName", NewDevice.DexcomReceiver.ProductName)
            Call EventProperties.Add("SoftwareNumber", NewDevice.DexcomReceiver.SoftwareNumber)
            Call EventProperties.Add("FirmwareVersion", NewDevice.DexcomReceiver.FirmwareVersion)
        End If

        ' Send the event to the Insights platform
        Call ApplicationInsights.TrackEvent("Unenrolled Dexcom Device", EventProperties)
    End Sub
End Class
