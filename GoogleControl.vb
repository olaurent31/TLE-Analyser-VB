<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
Public Class GoogleControl : Inherits UserControl

    '----------------------------------------------------------------------------------------------------------------------------
    'GOOGLE MAP API
    '----------------------------------------------------------------------------------------------------------------------------

    'GOOGLE MAP VARIABLES

    Private WebBrowser1 As WebBrowser
    Private StatusStrip1 As StatusStrip
    Private StatusButtonCenter As ToolStripButton
    Private KeepSatCentered As ToolStripButton
    Private SetInitialParameters As ToolStripButton
    Private ZoomSat As ToolStripButton
    Private StatusLabelLatLng As ToolStripStatusLabel

    Private KeepSatCenteredOpt As Boolean

    Private InitialZoom As Integer
    Private InitialLatitude As Double
    Private InitialLongitude As Double
    Private InitialMapType As GoogleMapType

    'GOOGLE MAP INIT

    Private Sub GoogleControl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        WebBrowser1.DocumentText = My.Computer.FileSystem.ReadAllText(Home.GOOGLEMAPPath & "GoogleMap.htm")
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        WebBrowser1.Document.InvokeScript("Initialize", New Object() {InitialZoom, InitialLatitude, InitialLongitude, CInt(InitialMapType)})
    End Sub

    Public Enum GoogleMapType
        RoadMap
        Terrain
        Hybrid
        Satellite
    End Enum

    Sub New()
        MyBase.New()

        WebBrowser1 = New WebBrowser
        StatusStrip1 = New StatusStrip
        SetInitialParameters = New ToolStripButton
        StatusButtonCenter = New ToolStripButton
        KeepSatCentered = New ToolStripButton
        ZoomSat = New ToolStripButton
        StatusLabelLatLng = New ToolStripStatusLabel

        WebBrowser1.Dock = DockStyle.Fill
        WebBrowser1.AllowWebBrowserDrop = False
        WebBrowser1.IsWebBrowserContextMenuEnabled = False
        WebBrowser1.WebBrowserShortcutsEnabled = False
        WebBrowser1.ObjectForScripting = Me
        WebBrowser1.ScriptErrorsSuppressed = True
        AddHandler WebBrowser1.DocumentCompleted, AddressOf WebBrowser1_DocumentCompleted

        StatusStrip1.Dock = DockStyle.Bottom
        StatusStrip1.Items.Add(SetInitialParameters)
        StatusStrip1.Items.Add(New ToolStripSeparator)
        StatusStrip1.Items.Add(StatusButtonCenter)
        StatusStrip1.Items.Add(New ToolStripSeparator)
        StatusStrip1.Items.Add(ZoomSat)
        StatusStrip1.Items.Add(New ToolStripSeparator)
        StatusStrip1.Items.Add(KeepSatCentered)
        StatusStrip1.Items.Add(New ToolStripSeparator)
        StatusStrip1.Items.Add(StatusLabelLatLng)

        SetInitialParameters.Text = "Set Initial Map Parameters"
        AddHandler SetInitialParameters.Click, AddressOf SetInitialParameters_Click

        StatusButtonCenter.Text = "Center Map on Satellite"
        AddHandler StatusButtonCenter.Click, AddressOf StatusButtonCenter_Click

        ZoomSat.Text = "Zoom In"
        AddHandler ZoomSat.Click, AddressOf ZoomSat_Click

        KeepSatCentered.Text = "Keep Satellite Centered"
        KeepSatCenteredOpt = False
        AddHandler KeepSatCentered.Click, AddressOf KeepSatCentered_Click

        Me.Controls.Add(WebBrowser1)
        Me.Controls.Add(StatusStrip1)

        InitialZoom = 2
        InitialLatitude = 0
        InitialLongitude = 0
        InitialMapType = GoogleMapType.Terrain

    End Sub

    Sub New(ByVal zoom As Integer, ByVal lat As Double, ByVal lng As Double, ByVal mapType As Integer) ' As GoogleMapType)
        Me.New()
        InitialZoom = zoom
        InitialLatitude = lat
        InitialLongitude = lng
        InitialMapType = mapType
    End Sub

    'GOOGLE MAP FUNCTIONS

    Sub GGMapWeb()
        If Home.MapsTab.SelectedIndex = 1 Then
            WebBrowser1.Document.InvokeScript("AddSunMarker", New Object() {Home.LATS, Home.LONGIS})
            'WebBrowser1.Document.InvokeScript("dayNightSun", New Object() {Home.LATS, Home.LONGIS})
            WebBrowser1.Document.InvokeScript("AddMarker", New Object() {Home.SATNAME, Home.LAT, Home.LONGI, Home.SatVisual3})
            If KeepSatCenteredOpt = True Then WebBrowser1.Document.InvokeScript("CenterMap")
            WebBrowser1.Document.InvokeScript("DeleteTrace")
            If Home.MapShowTrack.Checked = True Then WebBrowser1.Document.InvokeScript("AddTrace")
        End If
    End Sub

    Sub InitOrb()

        'Initialise l'orbite (vide la précédente si existante)
        WebBrowser1.Document.InvokeScript("InitOrb")
    End Sub

    'Sub CreateTrace(ByVal latitude, ByVal longitude)

    '    'Ajoute les points de la trace à la Polyline Google Maps
    '    WebBrowser1.Document.InvokeScript("CreateTrace", New Object() {latitude, longitude})
    'End Sub

    Sub CreateTrace()

        'Ajoute les points de la trace à la Polyline Google Maps
        For i = 0 To 360 * Home.MapPeriodNbr.SelectedItem
            WebBrowser1.Document.InvokeScript("CreateTrace", New Object() {Home.SatTrace(i, 2), Home.SatTrace(i, 3)})
            Home.ProgressBar.Value = i * 100 / (360 * Home.MapPeriodNbr.SelectedItem)
        Next
    End Sub

    Sub DeleteTrace()
        If Home.CheckNW = False Then Exit Sub

        WebBrowser1.Document.InvokeScript("DeleteTrace")
    End Sub

    'GOOGLE MAP OPTIONS

    Public Sub Map_MouseMove(ByVal lat As Double, ByVal lng As Double)
        StatusLabelLatLng.Text = "Latitude/Longitude: " & CStr(Math.Round(lat, 4)) & " , " & CStr(Math.Round(lng, 4))
    End Sub

    'Public Sub Map_Click(ByVal lat As Double, ByVal lng As Double)
    '    Dim MarkerName As String = InputBox("Enter a Marker Name", "New Marker")

    '    If Not String.IsNullOrEmpty(MarkerName) Then
    '        WebBrowser1.Document.InvokeScript("AddMarker", New Object() {MarkerName, lat, lng})
    '    End If
    'End Sub

    Private Sub SetInitialParameters_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Home.CheckNW = False Then Exit Sub

        WebBrowser1.Document.InvokeScript("SetInitialParameters")
    End Sub

    Private Sub StatusButtonCenter_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Home.CheckNW = False Then Exit Sub

        If Home.LAT = Nothing Or Home.LONGI = Nothing Then
            Exit Sub
        Else
            WebBrowser1.Document.InvokeScript("CenterMap")
        End If
    End Sub

    Private Sub KeepSatCentered_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Home.CheckNW = False Then Exit Sub

        If KeepSatCentered.Text = "Keep Satellite Centered" Then
            KeepSatCenteredOpt = True
            KeepSatCentered.Text = "Don't Keep Satellite Centered"
        Else
            KeepSatCentered.Text = "Keep Satellite Centered"
            KeepSatCenteredOpt = False
        End If
    End Sub

    Private Sub ZoomSat_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Home.CheckNW = False Then Exit Sub
        If Home.TLELoaded = True Then
            WebBrowser1.Document.InvokeScript("ZoomSat", New Object() {7, Home.LAT, Home.LONGI})
        End If
    End Sub

End Class
