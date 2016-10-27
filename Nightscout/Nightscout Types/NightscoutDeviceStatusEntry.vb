Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: October 30, 2016
''' Summary: A helper class that allows quick and easy JSON de/serialization to and from Nightscout.
''' </summary>

<DataContract> Public Class NightscoutDeviceStatusEntry
    Inherits NightscoutEntry

    Sub New()
        pEntryType = EntryTypes.DeviceStatusEntry
    End Sub

    <DataMember(EmitDefaultValue:=False)> Public Property _id As String
    <DataMember(EmitDefaultValue:=False)> Public Property created_at As String
    <DataMember(EmitDefaultValue:=False)> Public Property uploaderBattery As Integer
End Class
