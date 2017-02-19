Imports Microsoft.Band
Imports Microsoft.Band.Tiles
Imports Microsoft.Band.Tiles.Pages

Public Class BandApp
    Dim PageLayout As PageLayout
    Dim PageLayoutData As PageLayoutData
    Dim FilledPanel As FilledPanel

    Private lblCurrentSGV As TextBlock
    Private lblUpdateDelta As TextBlock
    Private lblCurrentSGVData As New TextBlockData(1, "---")
    Private lblUpdateDeltaData As New TextBlockData(2, "Loading...")

    Public Sub New()
        FilledPanel = New FilledPanel

    End Sub

    Public ReadOnly Property Layout As PageLayout
        Get
            Return PageLayout
        End Get
    End Property

    Public ReadOnly Property Data As PageLayoutData
        Get
            Return PageLayoutData
        End Get
    End Property

    Public Async Function LoadIconsAsync(ByVal Tile As BandTile) As Task

    End Function
End Class

Public Class PageLayoutData
    Private array(4) As PageElementData
    Private clone As PageElementData()

    Sub New(ByVal PageElementDataArray As PageElementData())
        array = PageElementDataArray
    End Sub

    Public ReadOnly Property All As PageElementData()
        Get
            Return array.Clone
        End Get
    End Property
End Class