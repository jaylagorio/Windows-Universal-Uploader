Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: May 15, 2016
''' Summary: A helper class that allows quick and easy JSON de/serialization to and from Nightscout.
''' </summary>

<DataContract> Public Class NightscoutGlucoseEntry
    <DataMember(EmitDefaultValue:=False)> Public Property _Id As String
    <DataMember(EmitDefaultValue:=False)> Public Property dateString As String
    <DataMember(EmitDefaultValue:=False)> Public Property [date] As ULong
    <DataMember(EmitDefaultValue:=False)> Public Property type As String
    <DataMember(EmitDefaultValue:=False)> Public Property device As String
    <DataMember(EmitDefaultValue:=False)> Public Property enteredBy As String
    <DataMember(EmitDefaultValue:=False)> Public Property direction As String
    <DataMember(EmitDefaultValue:=False)> Public Property sgv As Integer
    <DataMember(EmitDefaultValue:=False)> Public Property rssi As Integer
    <DataMember(EmitDefaultValue:=False)> Public Property noise As Integer
    <DataMember(EmitDefaultValue:=False)> Public Property mbg As Integer
    <DataMember(EmitDefaultValue:=False)> Public Property cal As Integer
End Class
