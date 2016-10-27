Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: October 30, 2016
''' Summary: Device represents a generic device from which the user will read and upload data to Nightscout.
''' </summary>

<DataContract> Public MustInherit Class Device

    ' The DeviceId provided by Windows for the device.
    Friend pDeviceId As String

    ' The friendly name of the device provided by the inheriting class.
    Friend pDisplayName As String

    ' A variable used to indicate whether the last sync attempt was successful. Initialized to True because
    ' a failure hasn't occured if a sync hasn't been attempted.
    Private pLastSyncSuccess As Boolean = True

    ' The type of data the device provides. One or many options can be set.
    Public Enum DeviceTypes
        CGMData = 1
        PumpData = 2
    End Enum

    ''' <summary>
    ''' Creates the Device object specifying the interface name, serial number, and DeviceId
    ''' to connect to the device.
    ''' </summary>
    ''' <param name="InterfaceName">The name of the interface to search for the device</param>
    ''' <param name="SerialNumber">The serial number of the device to search for</param>
    ''' <param name="DeviceId">The DeviceId provided by Windows used to connect to the device</param>
    Sub New(ByVal InterfaceName As String, ByVal SerialNumber As String, ByVal DeviceId As String)
        Me.InterfaceName = InterfaceName
        Me.SerialNumber = SerialNumber
        pDeviceId = DeviceId
    End Sub

    ''' <summary>
    ''' Creates the Device object specifying the interface name to search and the serial
    ''' number of the device to search for.
    ''' </summary>
    ''' <param name="InterfaceName">The interface to search for the device</param>
    ''' <param name="SerialNumber">The serial number of the device to search for</param>
    Sub New(ByVal InterfaceName As String, ByVal SerialNumber As String)
        Me.New(InterfaceName, SerialNumber, "")
    End Sub

    ''' <summary>
    ''' The serial number of the device.
    ''' </summary>
    ''' <returns>A string representing the serial number of the device, sometimes required to authenticate to the device</returns>
    <DataMember> Public Property SerialNumber As String

    ''' <summary>
    ''' The manufacturer of the device.
    ''' </summary>
    ''' <returns>A string representing the device manufacturer</returns>
    <DataMember> Public MustOverride Property Manufacturer As String

    ''' <summary>
    ''' The model of the device.
    ''' </summary>
    ''' <returns>A string representing the device model</returns>
    <DataMember> Public MustOverride Property Model As String

    ''' <summary>
    ''' The name of the Interface used to connect to the device.
    ''' </summary>
    ''' <returns>A string representing the interface used to connect to the device ("USB", "BLE", etc.)</returns>
    <DataMember> Public Property InterfaceName As String

    ''' <summary>
    ''' One or more types of data the device provides.
    ''' </summary>
    ''' <returns>A DeviceTypes value indicating the types of data the device provides.</returns>
    <DataMember> Public MustOverride Property DeviceType As DeviceTypes

    ''' <summary>
    ''' The last time the device successfully synchronized with Nighscout.
    ''' </summary>
    <DataMember> Public LastSyncTime As DateTime

    ''' <summary>
    ''' The friendly name of the device.
    ''' </summary>
    ''' <returns>A string representing the display name of the device</returns>
    Public MustOverride ReadOnly Property DisplayName As String

    ''' <summary>
    ''' The DeviceId provided by Windows to connected to the device.
    ''' </summary>
    ''' <returns>A string used to connect to the device with FromIdAsync() functions</returns>
    Public ReadOnly Property DeviceId As String
        Get
            Return pDeviceId
        End Get
    End Property

    ''' <summary>
    ''' Used to indicate whether the last sync attempt failed. This is initialized to True
    ''' because never having attempted a sync means it couldn't have failed. This property is
    ''' entirely up to the calling app to manage and doesn't in any way influence the operation
    ''' of a Device.
    ''' </summary>
    ''' <returns>True if never synced or the last sync attempt was successful, False if the last sync attempt failed.</returns>
    Public Property LastSyncSuccess As Boolean
        Get
            Return pLastSyncSuccess
        End Get
        Set(value As Boolean)
            pLastSyncSuccess = value
        End Set
    End Property

    ''' <summary>
    ''' The path to a thumbnail image of the device.
    ''' </summary>
    ''' <returns>A string with a ms-appx:/// path to a thumbnail asset</returns>
    Public MustOverride ReadOnly Property ThumbnailAssetSource() As String

    ' Devices will be looked for via DeviceId, then via SerialNumber if both have been provided
    ''' <summary>
    ''' Required to attempt to find the device on the system.
    ''' </summary>
    ''' <returns>True if the device is found, False otherwise.</returns>
    Public MustOverride Async Function Connect() As Task(Of Boolean)

    ''' <summary>
    ''' Required to see if the device is connected to and authenticated with the system.
    ''' </summary>
    ''' <returns>True if the device is connected and authenticated, False otherwise.</returns>
    Public MustOverride Async Function IsConnected() As Task(Of Boolean)

    ''' <summary>
    ''' Required to disconnect the device and destroy the underlying interface connection.
    ''' </summary>
    ''' <returns>True if the device is disconnected, False otherwise</returns>
    Public MustOverride Async Function Disconnect() As Task(Of Boolean)
End Class
