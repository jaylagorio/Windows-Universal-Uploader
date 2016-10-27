Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: October 30, 2016
''' Summary: A helper class that allows quick and easy JSON de/serialization to and from Nightscout.
''' </summary>

<DataContract> Public Class NightscoutTreatmentEntry
    Inherits NightscoutEntry

    Sub New()
        pEntryType = EntryTypes.TreatmentEntry
    End Sub

    <DataMember(EmitDefaultValue:=False)> Public Property _id As String
    <DataMember(EmitDefaultValue:=False)> Public Property eventTime As String
    <DataMember(EmitDefaultValue:=False)> Public Property created_at As String
    <DataMember(EmitDefaultValue:=False)> Public Property eventType As String
    <DataMember(EmitDefaultValue:=False)> Public Property insulin As Double
    <DataMember(EmitDefaultValue:=False)> Public Property carbs As Double
    <DataMember(EmitDefaultValue:=False)> Public Property glucose As Double
    <DataMember(EmitDefaultValue:=False)> Public Property enteredBy As String
    <DataMember(EmitDefaultValue:=False)> Public Property reason As String
    <DataMember(EmitDefaultValue:=False)> Public Property duration As Integer
End Class
