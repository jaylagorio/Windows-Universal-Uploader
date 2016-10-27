Imports Windows.Storage
Imports Windows.Storage.Streams
Imports System.Text.UTF8Encoding
Imports Windows.Storage.ApplicationData
Imports System.Runtime.Serialization.Json
Imports Windows.Security.Cryptography.Core

''' <summary>
''' Author: Jay Lagorio
''' Date: October 30, 2016
''' Summary: Exposes a quick and easy way to save and retrieve settings. Data is saved as Roaming Settings, allowing the 
''' user to move from system to system while being able to sync devices that have been enrolled on any other system.
''' </summary>

Public Class Settings
    ' Settings storage container
    Private Shared pSettingsContainer As ApplicationDataContainer

    ' Collection of Devices currently enrolled
    Private Shared pDevices As New Collection(Of Device)

    ' Keys for the settings Key/Value pairs
    Private Const FirstRunSetupDoneKey As String = "FirstRunSetupDone"
    Private Const FirstRunSettingsSavedKey As String = "FirstRunSettingsSaved"
    Private Const NightscoutURLKey As String = "NightscoutURL"
    Private Const NightscoutSecretKey As String = "NightscoutSecret"
    Private Const LastSyncTimeKey As String = "LastSyncTime"
    Private Const LastRecordTimestampKey As String = "LastRecordTimestamp"
    Private Const UploadIntervalKey As String = "UploadInterval"
    Private Const EnrolledDevicesKey As String = "EnrolledDevices"
    Private Const UseSecureUploadConnectionKey As String = "UseSecureUploadConnection"
    Private Const UseRoamingSettingsKey As String = "UseRoamingSettings"
    Private Const ScreenBehaviorKey As String = "ScreenBehavior"

    ' Behavior for the screen depending on battery/power situation
    Public Enum ScreenBehavior
        NormalScreenBehavior = 0        ' Follows the device's power settings
        KeepScreenOnWhenPluggedIn = 1   ' Keeps the device screen on when the app is active but only when plugged in
        KeepScreenOnAlways = 2          ' Keeps the device screen on when the app is active whether on battery or plugged in
    End Enum

    ''' <summary>
    ''' Initializes the class to store/retrieve settings.
    ''' </summary>
    Shared Sub New()
        Dim UseLocalSettings As Boolean = False
        Dim pLocalSettingsContainer As ApplicationDataContainer = Current.LocalSettings

        ' Attempt to see whether settings should be synced remotely or stored locally
        Try
            UseLocalSettings = CBool(pLocalSettingsContainer.Values(UseRoamingSettingsKey))
        Catch ex As Exception
            UseLocalSettings = False
        End Try

        ' Pick the right container
        If UseLocalSettings Then
            pSettingsContainer = Current.LocalSettings
        Else
            pSettingsContainer = Current.RoamingSettings
        End If

        Call DeserializeEnrolledDevices()
    End Sub

    ''' <summary>
    ''' Clears all settings for the user.
    ''' </summary>
    Public Shared Sub ClearSettings()
        pSettingsContainer.Values.Remove(FirstRunSetupDoneKey)
        pSettingsContainer.Values.Remove(FirstRunSettingsSavedKey)
        pSettingsContainer.Values.Remove(NightscoutURLKey)
        pSettingsContainer.Values.Remove(NightscoutSecretKey)
        pSettingsContainer.Values.Remove(LastSyncTimeKey)
        pSettingsContainer.Values.Remove(LastRecordTimestampKey)
        pSettingsContainer.Values.Remove(UploadIntervalKey)
        pSettingsContainer.Values.Remove(EnrolledDevicesKey)
        pSettingsContainer.Values.Remove(UseSecureUploadConnectionKey)
        pSettingsContainer.Values.Remove(UseRoamingSettingsKey)
        pSettingsContainer.Values.Remove(ScreenBehaviorKey)
        pDevices = New Collection(Of Device)
    End Sub

    ''' <summary>
    ''' Sets the OOBE property to reflect the welcome prompts being shown on first run.
    ''' </summary>
    ''' <returns>True if the prompts have already been run, False otherwise</returns>
    Public Shared Property FirstRunSetupDone As Boolean
        Get
            Try
                Return CBool(pSettingsContainer.Values(FirstRunSetupDoneKey))
            Catch ex As Exception
                Return False
            End Try
        End Get
        Set(value As Boolean)
            pSettingsContainer.Values(FirstRunSetupDoneKey) = value
        End Set
    End Property

    ''' <summary>
    ''' Indicates that the first run settings have been saved.
    ''' </summary>
    ''' <returns>True if the prompts have already been run, False otherwise</returns>
    Public Shared Property FirstRunSettingsSaved As Boolean
        Get
            Try
                Return CBool(pSettingsContainer.Values(FirstRunSettingsSavedKey))
            Catch ex As Exception
                Return False
            End Try
        End Get
        Set(value As Boolean)
            pSettingsContainer.Values(FirstRunSettingsSavedKey) = value
        End Set
    End Property

    ''' <summary>
    ''' Records the last time the app attempted to sync with Nightscout.
    ''' </summary>
    ''' <returns>A DateTime representing the last time the synchronization process was run</returns>
    Public Shared Property LastSyncTime As DateTime
        Get
            Try
                Return DateTime.Parse(pSettingsContainer.Values(LastSyncTimeKey))
            Catch ex As Exception
                Return DateTime.MinValue
            End Try
        End Get
        Set(value As DateTime)
            pSettingsContainer.Values(LastSyncTimeKey) = value.ToString()
        End Set
    End Property

    ''' <summary>
    ''' Records the URL to the Nightscout server not including the protocol or any path component.
    ''' </summary>
    ''' <returns>The Nightscout server address</returns>
    Public Shared Property NightscoutURL As String
        Get
            Try
                Return pSettingsContainer.Values(NightscoutURLKey)
            Catch ex As Exception
                Return ""
            End Try
        End Get
        Set(value As String)
            pSettingsContainer.Values(NightscoutURLKey) = value
        End Set
    End Property

    ''' <summary>
    ''' Sets the Nightscout API Secret. The property is set with the plain-text
    ''' secret but the SHA1 hash of the secret is the only thing stored and returned.
    ''' </summary>
    ''' <returns>A String containing the SHA1 of the assigned API Secret</returns>
    Public Shared Property NightscoutAPIKey As String
        Get
            Try
                Return pSettingsContainer.Values(NightscoutSecretKey)
            Catch ex As KeyNotFoundException
                Return ""
            End Try
        End Get
        Set(value As String)
            ' Turn the string into bytes and hash with SHA1
            Dim Bytes() As Byte = UTF8.GetBytes(value)
            Dim SHA1 As HashAlgorithmProvider = HashAlgorithmProvider.OpenAlgorithm("SHA1")
            Dim HashBuffer As IBuffer = SHA1.HashData(Bytes.AsBuffer())
            Dim HashBytes() As Byte = HashBuffer.ToArray()

            ' Convert the hash into a hex string
            Dim SHA1String As String = ""
            For i = 0 To HashBytes.Count - 1
                Dim Hex As String = DecimalToHex(HashBytes(i))
                If Hex.Length < 2 Then Hex = "0" & Hex
                SHA1String &= Hex
            Next

            pSettingsContainer.Values(NightscoutSecretKey) = SHA1String.ToLower()
        End Set
    End Property

    ''' <summary>
    ''' Specifies the sync interval in minutes. The default is 5.
    ''' </summary>
    ''' <returns>An integer representing the number of minutes between automatic sync attempts</returns>
    Public Shared Property UploadInterval As Integer
        Get
            Try
                Return pSettingsContainer.Values(UploadIntervalKey)
            Catch ex As KeyNotFoundException
                Return 5
            End Try
        End Get
        Set(value As Integer)
            pSettingsContainer.Values(UploadIntervalKey) = value
        End Set
    End Property

    ''' <summary>
    ''' Sets whether to use HTTP or HTTPS when uploading to Nightscout
    ''' </summary>
    ''' <returns>True makes the sync process use HTTPS, False uses HTTP</returns>
    Public Shared Property UseSecureUploadConnection As Boolean
        Get
            Try
                Return CBool(pSettingsContainer.Values(UseSecureUploadConnectionKey))
            Catch ex As Exception
                Return False
            End Try
        End Get
        Set(value As Boolean)
            pSettingsContainer.Values(UseSecureUploadConnectionKey) = value
        End Set
    End Property

    ''' <summary>
    ''' Sets whether to use/store settings in the Roaming store or the Local store
    ''' </summary>
    ''' <returns>True means settings are sync'd remotely, False means they're stored locally</returns>
    Public Shared Property UseRoamingSettings As Boolean
        Get
            Try
                Return CBool(pSettingsContainer.Values(UseRoamingSettingsKey))
            Catch ex As Exception
                Return True
            End Try
        End Get
        Set(value As Boolean)
            ' Only go through the following if the setting has changed
            If value <> UseRoamingSettings Then
                Dim CurrentContainer As ApplicationDataContainer = pSettingsContainer
                Dim NewContainer As ApplicationDataContainer = Nothing
                Dim EnrolledDevicesJSON As String = ""
                Try
                    EnrolledDevicesJSON = pSettingsContainer.Values(EnrolledDevicesKey).ToString.Trim()
                Catch ex As Exception
                    EnrolledDevicesJSON = ""
                End Try

                If value = True Then
                    ' Move all the values from the local store to the remote store
                    NewContainer = Current.RoamingSettings
                Else
                    ' Move all the values from the remote store to the local store
                    NewContainer = Current.LocalSettings
                End If

                ' Move all the current settings to the new container
                NewContainer.Values(FirstRunSetupDoneKey) = FirstRunSetupDone
                NewContainer.Values(NightscoutURLKey) = NightscoutURL
                NewContainer.Values(NightscoutSecretKey) = NightscoutAPIKey
                NewContainer.Values(LastSyncTimeKey) = LastSyncTime.ToString()
                NewContainer.Values(LastRecordTimestampKey) = LastSyncTime.ToString()
                NewContainer.Values(UploadIntervalKey) = UploadInterval
                NewContainer.Values(UseSecureUploadConnectionKey) = UseSecureUploadConnection
                NewContainer.Values(ScreenBehaviorKey) = ScreenBehaviorKey
                NewContainer.Values(UseRoamingSettingsKey) = value
                NewContainer.Values(EnrolledDevicesKey) = EnrolledDevicesJSON

                ' Switch out the container storage location
                pSettingsContainer = NewContainer
            End If
        End Set
    End Property

    ''' <summary>
    ''' Specifies how the screen should behave when the app is in the foreground.
    ''' </summary>
    ''' <returns>A ScreenBehavior enum value indicating whether to use the screen normally,
    ''' keep the screen on when the app is in the foreground and plugged into power, or keep
    ''' the screen on always even when on battery.</returns>
    Public Shared Property ScreenSleepBehavior As ScreenBehavior
        Get
            Try
                Return pSettingsContainer.Values(ScreenBehaviorKey)
            Catch ex As KeyNotFoundException
                Return ScreenBehavior.NormalScreenBehavior
            End Try
        End Get
        Set(value As ScreenBehavior)
            pSettingsContainer.Values(ScreenBehaviorKey) = CInt(value)
        End Set
    End Property

    ''' <summary>
    ''' Retrieves a Collection of Devices enrolled in the application that are saved between sessions.
    ''' </summary>
    ''' <returns>A collection of devices enrolled in the application.</returns>
    Public Shared ReadOnly Property EnrolledDevices() As Collection(Of Device)
        Get
            Try
                Return pDevices
            Catch ex As KeyNotFoundException
                Return New Collection(Of Device)
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Saves the sync time of the passed Device.
    ''' </summary>
    ''' <param name="EnrolledDevice">The Device representing the sync'd device</param>
    ''' <param name="LastSyncTime">A DateTime representing the last time a sync attempt was successful</param>
    Public Shared Sub SetEnrolledDeviceLastSyncTime(ByVal EnrolledDevice As Device, ByVal LastSyncTime As DateTime)
        Call RemoveEnrolledDevice(EnrolledDevice)
        EnrolledDevice.LastSyncTime = LastSyncTime
        Call AddEnrolledDevice(EnrolledDevice)
    End Sub

    ''' <summary>
    ''' Adds a new Device to the list of enrolled devices for the app to track across sessions
    ''' </summary>
    ''' <param name="NewDevice"></param>
    Public Shared Sub AddEnrolledDevice(ByVal NewDevice As Device)
        ' Add to the Device collection
        Call pDevices.Add(NewDevice)

        ' Serialize the Device into JSON
        Dim DeviceSerializer As DataContractJsonSerializer = Nothing
        If NewDevice.GetType Is GetType(DexcomDevice) Then
            DeviceSerializer = New DataContractJsonSerializer(GetType(DexcomDevice))
        End If
        Dim DeviceJsonStream As New MemoryStream
        DeviceSerializer.WriteObject(DeviceJsonStream, NewDevice)
        Dim DeviceJson As String = UTF8.GetString(DeviceJsonStream.ToArray())

        ' Store new JSON
        pSettingsContainer.Values(EnrolledDevicesKey) &= DeviceJson
    End Sub

    ''' <summary>
    ''' Removes a device from the list of enrolled devices for this and future sessions.
    ''' </summary>
    ''' <param name="Device">The device to remove</param>
    Public Shared Sub RemoveEnrolledDevice(ByVal Device As Device)
        Dim DeviceIndex As Integer = -1

        ' Look for the index of the device to remove
        For i = 0 To pDevices.Count - 1
            ' Try by Device ID first, then serial number
            If Device.DeviceId <> "" And pDevices(i).DeviceId <> "" Then
                If pDevices(i).DeviceId = Device.DeviceId Then
                    DeviceIndex = i
                    Exit For
                End If
            ElseIf Device.SerialNumber <> "" And pDevices(i).SerialNumber <> "" Then
                If pDevices(i).SerialNumber = Device.SerialNumber Then
                    DeviceIndex = i
                    Exit For
                End If
            End If
        Next

        If DeviceIndex >= 0 Then
            ' Remove it from the collection of current devices
            Call pDevices.RemoveAt(DeviceIndex)

            ' Take all of the devices remaining in the collection, serialize them, and
            ' rewrite the serialized setting to save them for future sessions.
            Dim DevicesJson As String = ""
            For i = 0 To pDevices.Count - 1
                Dim DeviceSerializer As DataContractJsonSerializer = Nothing
                If Device.GetType Is GetType(DexcomDevice) Then
                    DeviceSerializer = New DataContractJsonSerializer(GetType(DexcomDevice))
                End If

                Dim DeviceJsonStream As New MemoryStream
                DeviceSerializer.WriteObject(DeviceJsonStream, Device)
                DevicesJson &= UTF8.GetString(DeviceJsonStream.ToArray())
            Next

            pSettingsContainer.Values(EnrolledDevicesKey) = DevicesJson
        End If
    End Sub

    ''' <summary>
    ''' Load the serialized devices and deserialize them into a collection of devices that can be accessed
    ''' </summary>
    Private Shared Sub DeserializeEnrolledDevices()
        ' If the OOBE steps haven't been run then we don't have anything to load
        If Not FirstRunSetupDone Then Return

        Dim DeviceJson As String = ""
        Try
            DeviceJson = pSettingsContainer.Values(EnrolledDevicesKey).ToString.Trim()
        Catch ex As Exception
            DeviceJson = ""
        End Try

        If DeviceJson <> "" Then
            ' Load and deserialize the stream and pull the first Device object off
            Dim DeviceJsonStream As New MemoryStream(UTF8.GetBytes(DeviceJson))
            Dim DeviceSerializer As New DataContractJsonSerializer(GetType(Generic))
            Dim CurrentDevice As Device = Nothing

            ' Read the first object in the stream
            Try
                CurrentDevice = DeviceSerializer.ReadObject(DeviceJsonStream)
            Catch ex As Exception
                CurrentDevice = Nothing
            End Try

            Do
                If Not CurrentDevice Is Nothing Then
                    ' Check the manufacturer and create the specifc device (DexcomDevice, etc) from the Device
                    If CurrentDevice.Manufacturer = "Dexcom" Then
                        Call pDevices.Add(New DexcomDevice(CurrentDevice))
                    End If
                End If

                ' Try to pull another object off the stream and exit the loop if there isn't one
                Try
                    CurrentDevice = DeviceSerializer.ReadObject(DeviceJsonStream)
                Catch Ex As Exception
                    CurrentDevice = Nothing
                End Try
            Loop While Not CurrentDevice Is Nothing
        End If
    End Sub

    ''' <summary>
    ''' Encodes decimal values into hexidecimal strings.
    ''' </summary>
    ''' <param name="Dec">The decimal value to encode.</param>
    ''' <returns>A string representing the hexidecimal value of the passed byte.</returns>
    Private Shared Function DecimalToHex(ByVal Dec As Byte) As String
        Dim HexString As String

        If Dec < 10 Then
            ' Return the decimal value as the hex value
            HexString = Dec
        ElseIf Dec >= 10 And Dec < 16 Then
            ' Return the hex value A - F
            HexString = ChrW(Dec + AscW("A") - 10)
        Else
            ' Return two hex digits representing this one byte number
            HexString = DecimalToHex(Dec \ 16) & DecimalToHex(Dec Mod 16)
        End If

        Return HexString
    End Function
End Class
