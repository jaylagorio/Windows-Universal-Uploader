Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: May 29, 2016
''' Summary: A Generic device class used to load devices from settings storage.
''' </summary>

<DataContract> Public Class Generic
    Inherits Device

    ''' <summary>
    ''' Creates an instance of a device with the specified interface name, serial number, and DeviceId. All values
    ''' are accepted but must conform to the required values if converting this Device into another Device.
    ''' </summary>
    ''' <param name="InterfaceName">Interface to use to connect to the device</param>
    ''' <param name="SerialNumber">Serial number of the required device</param>
    ''' <param name="DeviceId">A DeviceId provided by Windows to connect to the device</param>
    Public Sub New(InterfaceName As String, SerialNumber As String, DeviceId As String)
        MyBase.New(InterfaceName, SerialNumber, DeviceId)
        pDisplayName = "Generic Device"
    End Sub

    ''' <summary>
    ''' Returns (DeviceTypes.CGMData And DeviceTypes.PumpData). Setting this property does not change its value.
    ''' </summary>
    ''' <returns></returns>
    <DataMember> Public Overrides Property DeviceType As DeviceTypes
        Get
            Return DeviceTypes.CGMData And DeviceTypes.PumpData
        End Get
        Set(value As DeviceTypes)

        End Set
    End Property

    ''' <summary>
    ''' The display name of the device.
    ''' </summary>
    ''' <returns>Returns a string containing the display name for the device</returns>
    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return "Generic"
        End Get
    End Property

    ''' <summary>
    ''' Returns the manufacturer of the device
    ''' </summary>
    ''' <returns>A string representing the manufacturer of the device</returns>
    <DataMember> Public Overrides Property Manufacturer As String

    ''' <summary>
    ''' Returns the model of the device
    ''' </summary>
    ''' <returns>A string representing the model of the device</returns>
    <DataMember> Public Overrides Property Model As String

    ''' <summary>
    ''' An empty string signifying that there is no thumbnail to represent a Generic device.
    ''' </summary>
    ''' <returns>An empty string</returns>
    Public Overrides ReadOnly Property ThumbnailAssetSource As String
        Get
            Return ""
        End Get
    End Property

    ''' <summary>
    ''' This function cannot be called for a Generic device.
    ''' </summary>
    ''' <returns>False</returns>
    Public Overrides Async Function Connect() As Task(Of Boolean)
        ' Make an Await call to defeat the compiler warning about the Async marking on the function.
        Await Task.Yield()
        Return False
    End Function

    ''' <summary>
    ''' This function cannot be called for a Generic device.
    ''' </summary>
    ''' <returns>False</returns>
    Public Overrides Async Function Disconnect() As Task(Of Boolean)
        ' Make an Await call to defeat the compiler warning about the Async marking on the function.
        Await Task.Yield()
        Return False
    End Function

    ''' <summary>
    ''' This function cannot be called for a Generic device.
    ''' </summary>
    ''' <returns>False</returns>
    Public Overrides Async Function IsConnected() As Task(Of Boolean)
        ' Make an Await call to defeat the compiler warning about the Async marking on the function.
        Await Task.Yield()
        Return False
    End Function
End Class
