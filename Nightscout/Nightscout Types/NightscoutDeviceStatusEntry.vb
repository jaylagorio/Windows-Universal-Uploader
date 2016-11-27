Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: November 6, 2016
''' Summary: A helper class that allows quick and easy JSON de/serialization to and from Nightscout.
''' </summary>

<DataContract> Public Class NightscoutDeviceStatusEntry
    Inherits NightscoutEntry

    <DataContract> Public Class PumpData
        <DataContract> Public Class BatteryData
            <DataMember(EmitDefaultValue:=False)> Public Property status As String
            <DataMember(EmitDefaultValue:=False)> Public Property voltage As Double
        End Class
        <DataContract> Public Class StatusData
            <DataMember(EmitDefaultValue:=False)> Public Property status As String
            <DataMember(EmitDefaultValue:=False)> Public Property timestamp As String
            <DataMember(EmitDefaultValue:=False)> Public Property bolusing As Boolean
            <DataMember(EmitDefaultValue:=False)> Public Property suspended As Boolean
        End Class
        <DataMember(EmitDefaultValue:=False)> Public Property battery As BatteryData
        <DataMember(EmitDefaultValue:=False)> Public Property status As StatusData
        <DataMember(EmitDefaultValue:=False)> Public Property reservoir As Double
        <DataMember(EmitDefaultValue:=False)> Public Property clock As String
    End Class

    <DataContract> Public Class OpenAPSData
        <DataContract> Public Class SuggestedEnactedData
            <DataMember(EmitDefaultValue:=False)> Public Property bg As Integer
            <DataMember(EmitDefaultValue:=False)> Public Property temp As String
            <DataMember(EmitDefaultValue:=False)> Public Property received As Boolean
            <DataMember(EmitDefaultValue:=False)> Public Property timestamp As String
            <DataMember(EmitDefaultValue:=False)> Public Property rate As Double
            <DataMember(EmitDefaultValue:=False)> Public Property reason As String
            <DataMember(EmitDefaultValue:=False)> Public Property duration As Integer
            <DataMember(EmitDefaultValue:=False)> Public Property IOB As Double
        End Class
        <DataMember(EmitDefaultValue:=False)> Public Property suggested As SuggestedEnactedData
        <DataMember(EmitDefaultValue:=False)> Public Property enacted As SuggestedEnactedData
    End Class

    Sub New()
        pEntryType = EntryTypes.DeviceStatusEntry
    End Sub

    <DataMember(EmitDefaultValue:=False)> Public Property _id As String
    <DataMember(EmitDefaultValue:=False)> Public Property device As String
    <DataMember(EmitDefaultValue:=False)> Public Property pump As PumpData
    <DataMember(EmitDefaultValue:=False)> Public Property openaps As OpenAPSData
    <DataMember(EmitDefaultValue:=False)> Public Property uploaderBattery As Integer
    <DataMember(EmitDefaultValue:=False)> Public Property created_at As String
End Class
