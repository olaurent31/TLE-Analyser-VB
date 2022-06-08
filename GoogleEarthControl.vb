<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
Public Class GoogleEarthControl : Inherits UserControl

    '----------------------------------------------------------------------------------------------------------------------------
    'GOOGLE EARTH API
    '----------------------------------------------------------------------------------------------------------------------------

    'GOOGLE EARTH VARIABLES

    Private WebBrowser2 As WebBrowser
    Private StatusStrip2 As StatusStrip
    Private GoToSat As ToolStripButton
    Private GoToGround As ToolStripButton
    Private InitialView As ToolStripButton
    Private EarthGrid As ToolStripButton
    Private Borders As ToolStripButton
    Private Nav As ToolStripButton

    'GOOGLE EARTH INIT

    Private Sub GoogleEarthControl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        WebBrowser2.DocumentText = My.Computer.FileSystem.ReadAllText(Home.GOOGLEPath & "GoogleEarth.htm")
    End Sub

    'Private Sub WebBrowser2_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
    '    'WebBrowser2.Document.InvokeScript("init")
    'End Sub

    Sub New()
        MyBase.New()

        WebBrowser2 = New WebBrowser
        StatusStrip2 = New StatusStrip
        GoToSat = New ToolStripButton
        GoToGround = New ToolStripButton
        InitialView = New ToolStripButton
        EarthGrid = New ToolStripButton
        Borders = New ToolStripButton
        Nav = New ToolStripButton

        WebBrowser2.Dock = DockStyle.Fill
        WebBrowser2.AllowWebBrowserDrop = False
        WebBrowser2.IsWebBrowserContextMenuEnabled = False
        WebBrowser2.WebBrowserShortcutsEnabled = False
        WebBrowser2.ObjectForScripting = Me
        WebBrowser2.ScriptErrorsSuppressed = True
        'AddHandler WebBrowser2.DocumentCompleted, AddressOf WebBrowser2_DocumentCompleted

        StatusStrip2.Dock = DockStyle.Bottom
        StatusStrip2.Items.Add(InitialView)
        StatusStrip2.Items.Add(New ToolStripSeparator)
        StatusStrip2.Items.Add(GoToSat)
        StatusStrip2.Items.Add(New ToolStripSeparator)
        StatusStrip2.Items.Add(GoToGround)
        StatusStrip2.Items.Add(New ToolStripSeparator)
        StatusStrip2.Items.Add(EarthGrid)
        StatusStrip2.Items.Add(New ToolStripSeparator)
        StatusStrip2.Items.Add(Borders)
        StatusStrip2.Items.Add(New ToolStripSeparator)
        StatusStrip2.Items.Add(Nav)
        StatusStrip2.Items.Add(New ToolStripSeparator)

        GoToSat.Text = "Go to Satellite"
        AddHandler GoToSat.Click, AddressOf GoToSat_Click

        GoToGround.Text = "Go to Ground Point"
        AddHandler GoToGround.Click, AddressOf GoToGround_Click

        InitialView.Text = "Initial View"
        AddHandler InitialView.Click, AddressOf InitialView_Click

        EarthGrid.Text = "Grid On"
        AddHandler EarthGrid.Click, AddressOf EarthGrid_Click

        Borders.Text = "Borders On"
        AddHandler Borders.Click, AddressOf Borders_Click

        Nav.Text = "Show Nav"
        AddHandler Nav.Click, AddressOf Nav_Click

        Me.Controls.Add(WebBrowser2)
        Me.Controls.Add(StatusStrip2)

    End Sub

    'GOOGLE MAP FUNCTIONS

    Sub addSat()
        Dim geEpoch
        geEpoch = MJDGGEDate(Home.EPOCH) & "T" & MJDGGEHour(Home.EPOCH) & "Z"

        If Home.MapsTab.SelectedIndex = 2 Then
            WebBrowser2.Document.InvokeScript("DeleteTrace")
            WebBrowser2.Document.InvokeScript("addSat", New Object() {Home.SATNAME, Home.LAT, Home.LONGI, Home.ALT * 1000, geEpoch})
            If Home.MapShowTrack.Checked = True Then WebBrowser2.Document.InvokeScript("AddTrace")
        End If
    End Sub

    Sub InitOrb()

        'Initialise l'orbite (vide la précédente si existante)
        WebBrowser2.Document.InvokeScript("InitOrb")
    End Sub

    'Sub CreateTrace(ByVal latitude, ByVal longitude, ByVal altitude)

    '    'Ajoute les points de la trace à la Polyline Google Maps
    '    WebBrowser2.Document.InvokeScript("CreateTrace", New Object() {latitude, longitude, altitude})
    'End Sub

    Sub CreateTrace()
        'Ajoute les points de la trace à la Polyline Google Maps
        For i = 0 To 360 * Home.MapPeriodNbr.SelectedItem
            WebBrowser2.Document.InvokeScript("CreateTrace", New Object() {Home.SatTrace(i, 2), Home.SatTrace(i, 3), Home.SatTrace(i, 4)})
            Home.ProgressBar.Value = i * 100 / (360 * Home.MapPeriodNbr.SelectedItem)
        Next
    End Sub

    Sub DeleteTrace()
        If Home.CheckNW = False Then Exit Sub

        WebBrowser2.Document.InvokeScript("DeleteTrace")
    End Sub

    'GOOGLE MAP OPTIONS

    Private Sub GoToSat_Click(ByVal sender As Object, ByVal e As EventArgs)
        GoToSatF()
    End Sub

    Sub GoToSatF()
        Dim TILT, HEADING
        If Home.CheckNW = False Or Home.TLELoaded = False Then Exit Sub

        If Home.ALT < 1000 Then
            TILT = 70
        ElseIf Home.ALT > 1000 And Home.ALT < 20000 Then
            TILT = 30
        ElseIf Home.ALT > 20000 Then
            TILT = 0
        End If
        If Home.INC > 90 Then
            HEADING = (Home.INC - 90) + 20
        ElseIf Home.INC < 90 Then
            HEADING = -((90 - Home.INC) - 20)
        End If
        WebBrowser2.Document.InvokeScript("GoToSat", New Object() {Home.LAT, Home.LONGI, Home.ALT * 1000, TILT, HEADING})
    End Sub

    Private Sub GoToGround_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Home.CheckNW = False Or Home.TLELoaded = False Then Exit Sub

        WebBrowser2.Document.InvokeScript("GoToSat", New Object() {Home.LAT, Home.LONGI, 10000, 0, 0})
    End Sub

    Private Sub InitialView_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Home.CheckNW = False Then Exit Sub

        InitialViewF()
    End Sub

    Function InitialViewF()
        Dim ALT
        If Home.CheckNW = False Then Exit Function

        If Home.ALT < 2000 Then
            ALT = 13000000
        ElseIf Home.ALT > 2000 Then
            ALT = 63000000
        End If
        WebBrowser2.Document.InvokeScript("InitialView", New Object() {Home.LAT, Home.LONGI, Home.ALT, ALT})
    End Function

    Private Sub EarthGrid_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim Status As Boolean
        If Home.CheckNW = False Then Exit Sub

        If EarthGrid.Text = "Grid OFF" Then
            Status = False
            EarthGrid.Text = "Grid ON"
        Else
            Status = True
            EarthGrid.Text = "Grid OFF"
        End If

        WebBrowser2.Document.InvokeScript("EarthGrid", New Object() {Status})
    End Sub

    Private Sub Borders_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim Status As Boolean
        If Home.CheckNW = False Then Exit Sub

        If Borders.Text = "Borders OFF" Then
            Status = False
            Borders.Text = "Borders ON"
        Else
            Status = True
            Borders.Text = "Borders OFF"
        End If

        WebBrowser2.Document.InvokeScript("Borders", New Object() {Status})
    End Sub

    Private Sub Nav_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim Status As Boolean
        If Home.CheckNW = False Then Exit Sub

        If Nav.Text = "Show Nav" Then
            Status = True
            Nav.Text = "Hide Nav"
        Else
            Status = False
            Nav.Text = "Show Nav"
        End If

        WebBrowser2.Document.InvokeScript("Navigation", New Object() {Status})
    End Sub

    Sub DrawSun()
        Dim Status As Boolean
        Dim geEpoch
        If Home.CheckNW = False Then Exit Sub

        geEpoch = MJDGGEDate(Home.EPOCH) & "T" & MJDGGEHour(Home.EPOCH) & "Z"

        If Home.SunShadowCB.Checked = True Then
            Status = True
        Else
            Status = False
        End If

        WebBrowser2.Document.InvokeScript("Sun", New Object() {Status, geEpoch})
    End Sub

End Class
