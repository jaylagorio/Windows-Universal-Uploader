Imports System.Runtime.Serialization

''' <summary>
''' Author: Jay Lagorio
''' Date: May 22, 2016
''' Summary: A helper class that allows quick and easy JSON de/serialization to and from Nightscout.
''' </summary>

<DataContract> Public MustInherit Class NightscoutEntry
    Friend pEntryType As EntryTypes

    Public Enum EntryTypes
        GlucoseEntry = 0
        DeviceStatusEntry = 1
        TreatmentEntry = 2
    End Enum

    ''' <summary>
    ''' Indicates which kind of NightscoutEntry this instance represents.
    ''' </summary>
    ''' <returns>A value from EntryTypes describing the type of entry</returns>
    Public ReadOnly Property EntryType As EntryTypes
        Get
            Return pEntryType
        End Get
    End Property
End Class
