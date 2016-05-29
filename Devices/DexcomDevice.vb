Imports Dexcom
Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: May 29, 2016
''' Summary: Provides communication with Dexcom monitor devices.
''' </summary>

<DataContract> Public Class DexcomDevice
    Inherits Device

    ' Instance of the Dexcom.Receiver
    Dim pDevice As Receiver

    ''' <summary>
    ''' Creates an instance of a Dexcom monitor device on the specified interface. One or both of the SerialNumber and
    ''' DeviceId fields must be provided. If both are provided FindDevice() will use DeviceId first to find the
    ''' device, then will use the SerialNumber. For some interfaces SerialNumber is required to authenticate to the device.
    ''' </summary>
    ''' <param name="InterfaceName">The name of the Interface to use when connecting a device</param>
    ''' <param name="SerialNumber">The serial number of the device, if known</param>
    ''' <param name="DeviceId">The DeviceId provided from Windows, if known</param>
    Sub New(ByVal InterfaceName As String, ByVal SerialNumber As String, ByVal DeviceId As String)
        MyBase.New(InterfaceName, SerialNumber, DeviceId)
        pDisplayName = "Dexcom Receiver"
        Me.SerialNumber = SerialNumber
    End Sub

    ''' <summary>
    ''' Creates an instance of a Dexcom monitor device from an existing Generic instance.
    ''' </summary>
    ''' <param name="GenericDevice">The Generic instance of the device from which to derive the Dexcom receiver</param>
    Sub New(ByRef GenericDevice As Generic)
        Me.New(GenericDevice.InterfaceName, GenericDevice.SerialNumber, GenericDevice.DeviceId)
        LastSyncTime = GenericDevice.LastSyncTime
    End Sub

    ''' <summary>
    ''' Returns DeviceTypes.CGMData. Setting this property does not change its value.
    ''' </summary>
    ''' <returns>DeviceTypes.CGMData</returns>
    <DataMember> Public Overrides Property DeviceType As DeviceTypes
        Get
            Return DeviceTypes.CGMData
        End Get
        Set(value As DeviceTypes)

        End Set
    End Property

    ''' <summary>
    ''' Returns the display name of the device.
    ''' </summary>
    ''' <returns>The display name</returns>
    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return pDisplayName
        End Get
    End Property

    ''' <summary>
    ''' Returns the manufacturer of the device.
    ''' </summary>
    ''' <returns>The string "Dexcom"</returns>
    <DataMember> Public Overrides Property Manufacturer As String
        Get
            Return "Dexcom"
        End Get
        Set(value As String)

        End Set
    End Property

    ''' <summary>
    ''' Returns the model of the device. Setting this property does not change its value.
    ''' </summary>
    ''' <returns>Returns "G4 Receiver"</returns>
    <DataMember> Public Overrides Property Model As String
        Get
            Return "G4 Receiver"
        End Get
        Set(value As String)

        End Set
    End Property

    ''' <summary>
    ''' This property gives the application access to the Dexcom.Receiver object to query for data.
    ''' </summary>
    ''' <returns>The current instance of the Dexcom.Receiver</returns>
    Public ReadOnly Property DexcomReceiver As Receiver
        Get
            Return pDevice
        End Get
    End Property

    ''' <summary>
    ''' Returns a string suitable for use with image display controls to represent the device.
    ''' </summary>
    ''' <returns>A string to an image asset in the application</returns>
    Public Overrides ReadOnly Property ThumbnailAssetSource As String
        Get
            Return "ms-appx:///Assets/DexcomReceiverIcon.png"
        End Get
    End Property

    ''' <summary>
    ''' Attempts to locate and connect to the device specified when the instance of this class was constructed. Connected
    ''' devices are considered found when the DeviceId parameter matches if that parameter was passed or
    ''' if the SerialNumber parameter matches the serial number of a connected device.
    ''' </summary>
    ''' <returns>True if the device is successfully connected, False otherwise</returns>
    Public Overrides Async Function Connect() As Task(Of Boolean)
        ' Check to see if pDevice is already something and if so try to revive
        ' the connection by calling Connect() so that uses the existing interface data
        If Not pDevice Is Nothing Then
            ' If the device is already connected don't look for it again.
            If Await IsConnected() Then
                Return True
            End If

            ' Try to reestablish any previously existing connection
            If Await pDevice.Connect() Then
                Return True
            End If
        End If

        ' Create a new interface to use to connect to the device
        Dim USBInterface As New USBInterface
        Dim BLEInterface As New BLEInterface
        Dim ConnectionInterface As DeviceInterface = Nothing
        If InterfaceName = USBInterface.InterfaceName Then
            ConnectionInterface = USBInterface
        ElseIf InterfaceName = BLEInterface.InterfaceName Then
            BLEInterface.SerialNumber = Me.SerialNumber
            ConnectionInterface = BLEInterface
        End If

        ' If an interface wasn't settled then fail out
        If ConnectionInterface Is Nothing Then
            Return False
        End If

        ' Search for all available Dexcom receivers on the selected interface.
        Dim DevicesFound As Collection(Of DeviceInterface.DeviceConnection) = Await ConnectionInterface.GetAvailableDevices()

        For i = 0 To DevicesFound.Count - 1
            Dim NewReceiver As New Dexcom.Receiver(ConnectionInterface)

            ' Check the DeviceID parameter if passed.
            If DeviceId <> "" Then
                If DevicesFound(i).DeviceId = DeviceId Then
                    ' Connect to the device
                    If Await NewReceiver.Connect(DevicesFound(i)) Then
                        pDevice = NewReceiver
                        pDeviceId = DevicesFound(i).DeviceId
                        SerialNumber = NewReceiver.SerialNumber
                        pDisplayName = DevicesFound(i).DisplayName
                        Return True
                    Else
                        ' Fail out if an error occurs
                        Return False
                    End If
                End If
            ElseIf SerialNumber <> "" Then
                ' Connect to the receiver
                If Await NewReceiver.Connect(DevicesFound(i)) Then
                    ' Check the serial number, but don't return an error if it doesn't match
                    If NewReceiver.SerialNumber = Me.SerialNumber Then
                        pDevice = NewReceiver
                        pDeviceId = DevicesFound(i).DeviceId
                        pDisplayName = DevicesFound(i).DisplayName
                        Return True
                    End If
                Else
                    ' Fail out if an error occurs
                    Return False
                End If
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Sends a Ping command and waits for a proper response or a timeout to indicate
    ''' whether the device is connected.
    ''' </summary>
    ''' <returns>True if the device is connected, False otherwise</returns>
    Public Overrides Async Function IsConnected() As Task(Of Boolean)
        If Not pDevice Is Nothing Then
            If Await pDevice.Ping Then
                Return True
            End If
        End If

        Return False
    End Function

    ''' <summary>
    ''' Disconnects the device.
    ''' </summary>
    ''' <returns>True if the device is disconnected, False otherwise.</returns>
    Public Overrides Async Function Disconnect() As Task(Of Boolean)
        ' Disconnect the interface and destroy the underlying object
        Try
            Await pDevice.Disconnect()
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function
End Class
