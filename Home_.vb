'
'    TLE ANALYSER: orbital datas, position of artificial satellites, prediction of their passes, ground track
'    Copyright (C) 2012-2013  mailto: olaurent31@gmail.com
'
'    This program is a free software
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'_______________________________________________________________________________________________________
'
' Description
' > orbital datas, position of artificial satellites, prediction of their passes, ground track, export to 3D softawres
'
' Auteur
' > Olivier LAURENT
'
' Date de creation
' > Octobre 2012
'
' Date de revision
' > 13/01/2013

Imports System.IO
Imports System.Globalization
Imports System.Math
Imports GoogleControl

Public Class Home

    'DECLARATION DES VARIABLES PUBLIQUES

    'Parametres
    Public SMA, ECC, INC, RAAN, AOP, MA, TA, EA, ETFP, ALPHA, LAT, LONGI, APA, PEA, APR, PER, ALT, MALT, APV, PEV, VEL, BSTAR, _
        LATS, LONGIS, LATTLE, LONGITLE, DREL, ETFE As Double
    Public X, Y, Z, VX, VY, VZ, XS, YS, ZS, XN, YN, ZN, VXN, VYN, VZN, XTLE, YTLE, ZTLE, VXTLE, VYTLE, VZTLE, lsol, bsol As Double
    Public XC, YC, ZC, VXC, VYC, VZC, SMAC, ECCC, INCC, MALTC, XSC, YSC, ZSC, LNGC As Double
    Public APER, DPER, LANO, LAN, MM, NP, AP, DL As Double
    Public EPOCH, EPOCH0, GST, GSTTLE, LST, TLEPOCH, NT0 As Double
    Public LIGNE1, LIGNE2, CATNUMB, TLEFILE, SATNAME, TLETXT, DELTAD, LTAN, TLEONAE, ONAE As String
    Public OffsetUTC As String
    Public DeleteFavMode, TLELoaded, TLETime, CheckNW, OptionsSaved, TrackModeTrace, ResizeMode As Boolean

    'Fichiers
    Public AppPath = "C:\TLEAnalyser\"
    Public FavAdress = AppPath & "FAV\favorites.txt"
    Public FavPath = AppPath & "FAV\"
    Public TLEPath = AppPath & "TLE\"
    Public PlotPath = AppPath & "PLOT\"
    Public GMATPath = AppPath & "GMAT\"
    Public CELESTIAPath = AppPath & "CELESTIA\"
    Public GOOGLEPath = AppPath & "GOOGLEEARTH\"
    Public GOOGLEMAPPath = AppPath & "GOOGLEMAP\"
    Public DLLPath = AppPath & "DLL\"

    'Google API
    Public GoogleControl1 As GoogleControl
    Public GoogleEarthControl1 As GoogleEarthControl

    'Graphiques et Traces
    Public MapW As Integer
    Public MapH As Integer
    Public MapW1 As Integer
    Public MapH1 As Integer
    Public MapW2 As Integer
    Public MapH2 As Integer
    Public Trace
    Public g As Graphics
    Public SunZone(360) As PointF
    Public SatTrace(360, 4) As Double

    'Options
    Public TLEAIniVersion As String
    Public TleUpdateDate As String
    Public TrackMode1 As String
    Public TrackMode2 As String
    Public Speed As String
    Public SatVisual1 As String
    Public SatVisual2 As String
    Public SatVisual3 As String
    Public SatVisual4 As String
    Public ShowTrack As String
    Public OptionGmatModel1 As String
    Public OptionGmatModel2 As String
    Public Propagate As String
    Public LngWin As String
    Public LatWin As String
    Public REFTLE As String
    Public REFRM As String
    Public LatMis As String
    Public LngMis As String
    Public SimulOn As String

    'Version du logiciel
    Public TLEAVersion As String = "1.10"

    'ACTIVATE

    Sub New()
        InitializeComponent()

        GoogleControl1 = New GoogleControl(1, 0, 0, 4)
        GoogleControl1.Dock = DockStyle.Fill
        MapsTab.TabPages(1).Controls.Add(GoogleControl1)

        GoogleEarthControl1 = New GoogleEarthControl
        GoogleEarthControl1.Dock = DockStyle.Fill
        MapsTab.TabPages(2).Controls.Add(GoogleEarthControl1)

    End Sub

    Private Sub Load_App_1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim TLE
        Dim istat As Integer

        'Installation des TLE si non-existantes
        CreateTLEFolder()
        'Cration des autres fichiers et dossiers
        CreateFolders()

        SetCulture()                            'langue et séparateur de décimale
        CheckNetwork()                          'Vérifie le réseau
        ReadInit()                              'Récupère les données options
        SaveInit()                              'Sauvegarde les données options
        If CheckNW = True Then CheckVersion() 'Vérifie si une version plus récente existe
        TrackMode()
        OptionsSaved = False

        'Affiche le n° de version en cours (issue de CheckVersion())
        Text = "TLE ANALYSER V" & TLEAVersion

        'Remplissage des listes de Sat.
        TLE_ListBox.Items.Clear()
        FileOpen(1, "C:\TLEAnalyser\TLE\TLEList.txt", OpenMode.Input)
        While Not EOF(1)
            TLE = LineInput(1)
            TLE_ListBox.Items.Add(TLE & ".txt")
        End While
        FileClose(1)

        'Icones Sat et Sun pour les LatLng de la map 2D std
        Dim SatPict As New Bitmap(My.Resources.sat2)
        Dim SunPictSize As New Size(16, 16)
        Dim SunPict As New Bitmap(My.Resources.sun, SunPictSize)
        SatPict.MakeTransparent(SatPict.GetPixel(0, 0))
        SunPict.MakeTransparent(SunPict.GetPixel(0, 0))
        MapSatPict.Image = SatPict
        MapSunPict.Image = SunPict

        'Compte le nombre de Satellite dans la biliothèque
        FindSCLabel.Text = "Find a Satellite (" & SatCount() & " TLE)"

        'Gestion des Onglets
        If CheckNW = False Then
            MessageBox.Show("You must be connected to access to Google Earth/Maps functions", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Else
            'Charge les Tabs "GOOGLE MAP"
            MapsTab.SelectTab(1)
            'Affiche la vue Google Earth
            MapsTab.SelectTab(2)
        End If

    End Sub

    Private Sub Load_App_2(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.FormClosing

        Try

            If MessageBox.Show("Exit TLE Analyser ?", "TLE Analyser", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) = DialogResult.No Then
                e.Cancel = True
            Else
                Dim Destination As String
                Destination = "C:\TLEAnalyser\UM.pdf"
                System.IO.File.Delete(Destination)

                If My.Computer.FileSystem.FileExists(GOOGLEMAPPath & "GoogleMap.htm") = True Then File.Delete(GOOGLEMAPPath & "GoogleMap.htm")
                If My.Computer.FileSystem.FileExists(GOOGLEPath & "GoogleEarth.htm") = True Then File.Delete(GOOGLEPath & "GoogleEarth.htm")
            End If

        Catch Ex As Exception
            MessageBox.Show("An error as occured at TLE Analyser closing." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            ProgressBar.Visible = False
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub GoogleMapTab_Selected_2(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles MapsTab.Selecting

        If Timer1.Enabled = True Then Timer1.Enabled = False
        If Timer2.Enabled = True Then Timer2.Enabled = False
        If CheckNW = False Then 'Si aucun réseau, on empeche la sélection des onglets google
            e.Cancel = True
            MessageBox.Show("You must be connected to access to Google Earth functions", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End If

    End Sub

    Private Sub GoogleMapTab_Selected_1(ByVal sender As Object, ByVal e As System.EventArgs) Handles MapsTab.Click

        'MAP TAB
        If MapsTab.SelectedIndex = 0 Then
            If TLELoaded = True Then ResizeWindow()
        End If

        'GOOGLE MAP TAB
        If MapsTab.SelectedIndex = 1 Then
            If TLELoaded = True Then
                GoogleControl1.GGMapWeb()
                ResizeWindow()
            End If
        End If

        'GOOGLE EARTH TAB
        If MapsTab.SelectedIndex = 2 Then
            If TLELoaded = True Then
                GoogleEarthControl1.addSat()
                ResizeWindow()
            End If
        End If
    End Sub

    'DIMENSIONNEMENT

    Private Sub home_Resize() Handles Me.Resize

        If TrackMode1 <> "" Or TrackMode2 <> "" Or TrackMode1 <> Nothing Or TrackMode2 <> Nothing Then
            ResizeMode = True

            MapW1 = MapPanel.Width
            MapH1 = MapPanel.Height

            'MapsTab
            MapsTab.Width = Width - MapsTab.Location.X - 20
            If MapsTab.SelectedIndex = 0 Then
                MapsTab.Height = 577
            Else
                MapsTab.Height = Height - MapsTab.Location.Y - 50
            End If
            MapsTab.Refresh()

            'MapPanel/Tracepicture
            If TABMAP.Width / 2 < TABMAP.Height - 20 Then
                MapPanel.Width = TABMAP.Width
                MapPanel.Height = MapPanel.Width / 2
            ElseIf TABMAP.Width / 2 > TABMAP.Height - 20 Then
                MapPanel.Height = TABMAP.Height - 20
                MapPanel.Width = MapPanel.Height * 2
            End If
            MapPanel.Refresh()

            MapW = MapPanel.Width
            MapH = MapPanel.Height

            TracePicture.Width = MapW
            TracePicture.Height = MapH
            TracePicture.MinimumSize = New Point(MapW, MapH)
            TracePicture.MaximumSize = New Point(MapW, MapH)
            TracePicture.Refresh()

            InitGraphics()

            If TLELoaded = True Then ActualiseGraphics(False)

            MapW2 = MapW
            MapH2 = MapH

        End If
    End Sub

    'MENU / ABOUT

    Private Sub AboutMenuStrip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutMenu.Click
        AboutAppli()
    End Sub

    Private Sub HelpMenuStrip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpMenu.Click
        HelpAppli()
    End Sub

    Private Sub ExitMenuStrip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitMenu.Click
        ExitAppli()
    End Sub

    'TOOLS MENU (ICONES)

    Private Sub ButtonMenu_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Open.Click
        OpenFavs()
    End Sub

    Private Sub ButtonMenu_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Save.Click
        SaveFavs()
    End Sub

    Private Sub ButtonMenu_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Delete.Click
        DeleteFavs()
    End Sub

    Private Sub ButtonMenu_Reload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Reload.Click
        ReloadTLE()
    End Sub

    Private Sub ButtonMenu_GMAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_GMAT.Click
        ExportToGMAT()
    End Sub

    Private Sub ButtonMenu_Celes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Celes.Click
        ExportToCelestia()
    End Sub

    Private Sub ButtonMenu_GE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_GE.Click
        ExportToGoogle()
    End Sub

    Private Sub ButtonMenu_Summ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Summ.Click
        Summary()
    End Sub

    Private Sub ButtonMenu_About_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_About.Click
        AboutAppli()
    End Sub

    Private Sub ButtonMenu_Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Help.Click
        HelpAppli()
    End Sub

    Private Sub ButtonMenu_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Exit.Click
        ExitAppli()
    End Sub

    Private Sub ButtonMenu_Options_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Options.Click
        Options_Show()
    End Sub

    Private Sub ButtonMenu_Mode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Mode.Click

        If TrackMode1 = True Then
            TrackMode1 = False
            TrackMode2 = True
        ElseIf TrackMode2 = True Then
            TrackMode1 = True
            TrackMode2 = False
        End If

        'Changement de Mode
        TrackMode()

        If TLELoaded = True Then ActualiseGraphics(False)

    End Sub

    Private Sub ButtonMenu_OpenF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_OpenF.Click

        Process.Start(AppPath)

    End Sub

    Private Sub ButtonMenu_Charts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMenu_Charts.Click

        ChartForm.Show()

        If MM > 0.99 AndAlso MM < 1.01 AndAlso ECC < 0.01 AndAlso INC < 1 Then
            ChartForm.CHART_LNG_CB.Enabled = True
        Else
            ChartForm.CHART_LNG_CB.Enabled = False
        End If

        ChartForm.ChartXvalue.SelectedIndex = 0


    End Sub

    'TOOLS MENU (TEXT)

    Private Sub Reload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReloadTLEMenu.Click
        ReloadTLE()
    End Sub

    Private Sub ExportToGMAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportGmatMenu.Click
        ExportToGMAT()
    End Sub

    Private Sub Summary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummaryMenu.Click

        Summary()

    End Sub

    Private Sub SaveToFavourites_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveFavMenu.Click
        SaveFavs()
    End Sub

    Private Sub OpenFavourites_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenFavMenu.Click
        OpenFavs()
    End Sub

    Private Sub CelestiaMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CelestiaMenu.Click

        ExportToCelestia()

    End Sub

    Private Sub GoogleMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoogleMenu.Click

        ExportToGoogle()

    End Sub

    Sub OpenFavs()
        Dim Chaine, TLE, SAT
        Dim FindFav As Boolean
        FindFav = False

        'On vérifie que le fichier n'est pas dans la liste des TLE
        For i = 0 To TLE_ListBox.Items.Count - 1
            If TLE_ListBox.Items.Item(i) = "favorites.txt" Then FindFav = True 'le fichier existe
        Next

        If FindFav = False Then
            TLE_ListBox.Items.Add("favorites.txt")
            TLE_ListBox.Sorted = True
            TLE_ListBox.SelectedItem = "favorites.txt"
        End If

        TLE_ListBox.SelectedItem = "favorites.txt"
    End Sub

    Sub SaveFavs()
        Dim TLE, SAT
        Dim FindFav As Boolean

        FindFav = False

        'On verifie si un satellite est sélectionné
        If Sat_ListBox.SelectedItem = "" Then
            MessageBox.Show("You must select a satellite in the list !", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        Try

            'On vérifie que le SAT n'est pas déjà sauvegardé dans le fichier
            If My.Computer.FileSystem.FileExists(FavAdress) = True Then 'Si le fichier existe
                FileOpen(1, FavAdress, OpenMode.Input)
                While Not EOF(1)
                    TLE = LineInput(1)
                    SAT = LineInput(1)

                    If SAT = Sat_ListBox.SelectedItem Then
                        MessageBox.Show(SAT & " is already saved as favourite", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        FindFav = True
                    End If
                End While
                FileClose(1)
            End If

            'On ajoute la TLE et le SAT en cours
            If FindFav = False Then
                TLE = TLE_ListBox.SelectedItem
                SAT = Sat_ListBox.SelectedItem
                File.AppendAllText(FavAdress, TLE & vbCrLf & SAT & vbCrLf)
                MessageBox.Show(SAT & " has been added into favorites", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            End If

        Catch Ex As Exception
            MessageBox.Show("An error has occured." & vbCrLf & vbCrLf & "Error : " + Ex.Message, "TLE ANALYSER - Error")
        End Try
    End Sub

    Sub DeleteFavs()
        Dim TLE(2000), SAT(2000) As String
        Dim FindFav As Boolean
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim FindSat, FindTLE
        Dim NewFavFile = "favorites_.txt"
        File.AppendAllText(FavPath & "\" & NewFavFile, "")

        FindFav = False
        DeleteFavMode = True

        'On verifie si un satellite est sélectionné
        If Sat_ListBox.SelectedItem = "" Then
            MessageBox.Show("You must select a satellite in the list !", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        Try

            'On vérifie que le SAT est présent dans le fichier
            If My.Computer.FileSystem.FileExists(FavAdress) = True Then 'Si le fichier existe
                FileOpen(1, FavAdress, OpenMode.Input)
                While Not EOF(1)
                    TLE(i) = LineInput(1)
                    SAT(i) = LineInput(1)

                    'SAT présent dans les favoris
                    If SAT(i) = Sat_ListBox.SelectedItem Then
                        FindFav = True
                        FindTLE = TLE(i)
                        FindSat = SAT(i)
                    Else
                        File.AppendAllText(FavPath & "\" & NewFavFile, TLE(i) & vbCrLf & SAT(i) & vbCrLf)
                    End If

                    i = i + 1
                End While
                FileClose(1)

                'SAT non présent dans les favoris, on supprime le nouveau fichier
                If FindFav = False Then
                    MessageBox.Show(Sat_ListBox.SelectedItem & " wasn't find in favorites.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    My.Computer.FileSystem.DeleteFile(FavPath & "\" & NewFavFile)
                End If

            End If

            'SAT présent dans les favoris, on supprime l'ancien fichier puis on renomme le nouveau
            If FindFav = True Then
                My.Computer.FileSystem.DeleteFile(FavAdress)
                My.Computer.FileSystem.RenameFile(FavPath & "\" & NewFavFile, "favorites.txt")
                If TLE_ListBox.SelectedItem = "favorites.txt" Then
                    Sat_ListBox.Items.Remove(FindSat)
                End If
                MessageBox.Show(FindSat & " has been deleted from favorites", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            End If

            DeleteFavMode = False

        Catch Ex As Exception
            MessageBox.Show("An error has occured." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
        End Try

    End Sub

    Sub ReloadTLE()
        TLETime = True
        LoadTLE()
        TLETime = False
        TLELoaded = True
        ActualiseModule(True)
        'TLELoaded = False
    End Sub

    Sub Summary()
        OrbitSummary.Show()

        On Error Resume Next

        'Le ListView1 existe
        OrbitSummary.ListSummary.View = View.Details
        Dim MyLine As ListViewItem = New ListViewItem(New String() {"Lasserre", "Philippe", "1951"})
        OrbitSummary.ListSummary.Items.Add(MyLine)

        Dim Epoch
        If Me.EpochFormat.SelectedIndex = 0 Then
            Epoch = MJDtoGreg(Val(Me.EPOCHBox.Text))
        Else
            Epoch = Me.EPOCHBox.Text
        End If

        'LAN W/E
        Dim LAN2 As String
        If Me.LANType.Text = "[0;360]" Then
            If Val(Me.LANBox.Text) < 360 AndAlso Val(Me.LANBox.Text) > 180 Then
                LAN2 = Round(360 - Val(Me.LANBox.Text), 2) & " W"
            ElseIf Val(Me.LANBox.Text) < 180 AndAlso Val(Me.LANBox.Text) > 0 Then
                LAN2 = Round(Val(Me.LANBox.Text), 2) & " E"
            End If
        ElseIf Me.LANType.Text = "[-180;180]" Then
            If Val(Me.LANBox.Text) < 0 Then
                LAN2 = Abs(Round(Val(Me.LANBox.Text), 2)) & " W"
            ElseIf Val(Me.LANBox.Text) > 0 Then
                LAN2 = Round(Val(Me.LANBox.Text), 2) & " E"
            End If
        End If

        OrbitSummary.OrbSum_TextBox.Clear()
        OrbitSummary.OrbSum_TextBox.Text = _
            Me.SCName.Text & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Epoch" & Chr(9) & Epoch & vbCrLf & _
            "GST" & Chr(9) & GST & vbCrLf & _
            "LST" & Chr(9) & LST & vbCrLf & _
            "ETFE" & Chr(9) & ETFE & vbCrLf & _
            "ONAE" & Chr(9) & ONAE & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Keplerian Elements:" & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "SMA" & Chr(9) & SMA & vbCrLf & _
            "ECC" & Chr(9) & ECC & vbCrLf & _
            "INC" & Chr(9) & INC & vbCrLf & _
            "RAAN" & Chr(9) & RAAN & vbCrLf & _
            "AOP" & Chr(9) & AOP & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Sat. Position:" & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "MA" & Chr(9) & MA & vbCrLf & _
            "TA" & Chr(9) & Maths.RAD2DEG * (TA) & vbCrLf & _
            "EA" & Chr(9) & Maths.RAD2DEG * (EA) & vbCrLf & _
            "ETFP" & Chr(9) & ETFP & vbCrLf & _
            "α" & Chr(9) & ALPHA & vbCrLf & _
            "Lat" & Chr(9) & LAT & vbCrLf & _
            "Long" & Chr(9) & LONGI & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Phasing" & Chr(9) & Me.phase.Text & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Altitudes:" & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "ApA" & Chr(9) & APA & vbCrLf & _
            "PeA" & Chr(9) & PEA & vbCrLf & _
            "ApR" & Chr(9) & APR & vbCrLf & _
            "PeR" & Chr(9) & PER & vbCrLf & _
            "ALT" & Chr(9) & ALT & vbCrLf & _
            "MALT" & Chr(9) & MALT & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Velocities:" & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "ApV" & Chr(9) & APV & vbCrLf & _
            "PeV" & Chr(9) & PEV & vbCrLf & _
            "Vel" & Chr(9) & VEL & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Periods:" & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Ta" & Chr(9) & APER & vbCrLf & _
            "Td" & Chr(9) & DPER & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "Mvmts, Precessions:" & vbCrLf & _
            "--------------------------" & vbCrLf & _
            "MM" & Chr(9) & MM & vbCrLf & _
            "NP" & Chr(9) & NP & vbCrLf & _
            "AP" & Chr(9) & AP & vbCrLf & _
            "DL" & Chr(9) & Me.DLBox.Text & vbCrLf & _
            "-------------------" & vbCrLf & _
            "Cartesian State:" & vbCrLf & _
            "-------------------" & vbCrLf & _
            "X" & Chr(9) & X & vbCrLf & _
            "Y" & Chr(9) & Y & vbCrLf & _
            "Z" & Chr(9) & Z & vbCrLf & _
            "VX" & Chr(9) & VX & vbCrLf & _
            "VY" & Chr(9) & VY & vbCrLf & _
            "VZ" & Chr(9) & VZ & vbCrLf & _
            "-------------------" & vbCrLf & _
            "SUN:" & vbCrLf & _
            "-------------------" & vbCrLf & _
            "Lat" & Chr(9) & LATS & vbCrLf & _
            "Long" & Chr(9) & LONGIS & vbCrLf & _
            "Eclipse" & Chr(9) & EclipseBox.Text

        'Ascending Node
        OrbitSummary.OrbSum_TextBox.Text = OrbitSummary.OrbSum_TextBox.Text & vbCrLf & _
        "-------------------" & vbCrLf & _
        "Ascending Node:" & vbCrLf & _
        "-------------------" & vbCrLf & _
        "LAN" & Chr(9) & LAN & vbCrLf & _
        "LTAN" & Chr(9) & LTAN

        'Adapted Parameters
        If AP_GroupBox.Visible = True Then
            If Me.CircularPanel.Visible = True Then
                OrbitSummary.OrbSum_TextBox.Text = OrbitSummary.OrbSum_TextBox.Text & vbCrLf & _
                "-------------------" & vbCrLf & _
                "Adapted Parameters:" & vbCrLf & _
                "-------------------" & vbCrLf & _
                "SMA" & Chr(9) & Me.SMA_AP_ECC.Text & vbCrLf & _
                "ex" & Chr(9) & Me.ex.Text & vbCrLf & _
                "ey" & Chr(9) & Me.ey.Text & vbCrLf & _
                "INC" & Chr(9) & Me.INC_AP_ECC.Text & vbCrLf & _
                "RAAN" & Chr(9) & Me.RAAN_AP_ECC.Text & vbCrLf & _
                "α'" & Chr(9) & Me.AlphaPrime.Text
            ElseIf Me.EquPanel.Visible = True Then
                OrbitSummary.OrbSum_TextBox.Text = OrbitSummary.OrbSum_TextBox.Text & vbCrLf & _
                "-------------------" & vbCrLf & _
                "Adapted Parameters:" & vbCrLf & _
                "-------------------" & vbCrLf & _
                "SMA" & Chr(9) & Me.SMA_AP_INC.Text & vbCrLf & _
                "ECC" & Chr(9) & Me.ECC_AP_INC.Text & vbCrLf & _
                "ix" & Chr(9) & Me.ix.Text & vbCrLf & _
                "iy" & Chr(9) & Me.iy.Text & vbCrLf & _
                "ω'" & Chr(9) & Me.AOP_AP_INC.Text & vbCrLf & _
                "MA" & Chr(9) & Me.MA_AP.Text
            ElseIf Me.CircEquPanel.Visible = True Then
                OrbitSummary.OrbSum_TextBox.Text = OrbitSummary.OrbSum_TextBox.Text & vbCrLf & _
                "-------------------" & vbCrLf & _
                "Adapted Parameters:" & vbCrLf & _
                "-------------------" & vbCrLf & _
                "SMA" & Chr(9) & Me.SMA_AP_ECCINC.Text & vbCrLf & _
                "ex" & Chr(9) & Me.ex_ECCINC.Text & vbCrLf & _
                "ey" & Chr(9) & Me.ey_ECCINC.Text & vbCrLf & _
                "ix" & Chr(9) & Me.ix_ECCINC.Text & vbCrLf & _
                "iy" & Chr(9) & Me.iy_ECCINC.Text & vbCrLf & _
                "λ'" & Chr(9) & Me.MeanL_ECCINC.Text
            End If
        End If
    End Sub

    Sub AboutAppli()
        Dim Destination As String
        Destination = "C:\TLEAnalyser\README.txt"

        If My.Computer.FileSystem.FileExists(Destination) = True Then My.Computer.FileSystem.DeleteFile(Destination)
        System.IO.File.WriteAllText(Destination, My.Resources.README)

        System.Diagnostics.Process.Start(Destination)
    End Sub

    Sub HelpAppli()
        Dim Destination As String
        Destination = "C:\TLEAnalyser\UM.pdf"

        If My.Computer.FileSystem.FileExists(Destination) = True Then My.Computer.FileSystem.DeleteFile(Destination)
        System.IO.File.WriteAllBytes(Destination, My.Resources.TAUM)

        System.Diagnostics.Process.Start(Destination)
    End Sub

    Sub ExitAppli()
        Close()
    End Sub

    Sub Options_Show()
        Options.Show()
    End Sub

    'IMPORT TLE

    Private Sub ImportTLEButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportTLEButton.Click
        If TLETextBox.ForeColor = Color.Black Then SATNAME = InputBox("Please enter a Satellite Name:")
        ImportTLE()
        ImportTLEButton.Enabled = False
    End Sub

    Private Sub DetailsTLEButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailsTLEButton.Click
        DetailTLE()
    End Sub

    Private Sub TLETextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TLETextBox.TextChanged
        ImportTLEButton.Enabled = True
        DetailsTLEButton.Enabled = True
    End Sub

    Private Sub NewButton_Click_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewButton.Click
        Me.TLETextBox.Enabled = True
        Me.TLETextBox.Text = ""
        Me.MapsTab.Focus()

        'Erase Datas
        Me.Value1.Text = ""
        Me.Value2.Text = ""
        Me.Value3.Text = ""
        Me.Value4.Text = ""
        Me.Value5.Text = ""
        Me.Value6.Text = ""
        Me.Value7.Text = ""
        Me.Value8.Text = ""
        Me.Value9.Text = ""
        Me.Value10.Text = ""
        Me.Value11.Text = ""
        Me.Value12.Text = ""
        Me.Value13.Text = ""
        Me.Value14.Text = ""
        Me.Value15.Text = ""
        Me.Value16.Text = ""
        Me.Value17.Text = ""
        Me.Value18.Text = ""
        Me.Value19.Text = ""
        Me.Value20.Text = ""
        Me.Value21.Text = ""
        Me.Value22.Text = ""
        Me.Value23.Text = ""
        Me.Value24.Text = ""

        Me.Sat_ListBox.Items.Clear()
        Me.TLETextBox.ReadOnly = False
        Me.DetailsTLEButton.Enabled = False
        Me.ImportTLEButton.Enabled = False

        Me.GroupBox1.Height = 111
        Me.DetailsTLEButton.Text = "SHOW DETAILS"

    End Sub

    Private Sub TLE_ListBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TLE_ListBox.SelectedIndexChanged

        Dim Fichier As String
        Dim i

        For i = 0 To Me.TLE_ListBox.Items.Count
            If Me.TLE_ListBox.SelectedIndex = i Then
                Fichier = TLEPath & TLE_ListBox.SelectedItem
                Exit For
            End If
        Next

        If Fichier = Nothing Then Exit Sub
        OpenTleFile(Fichier)

    End Sub

    Private Sub Sat_ListBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sat_ListBox.SelectedIndexChanged

        Dim line, ligne, SatNum, i, s As Integer
        Dim sat, sat1, sat2 As String
        Dim FindSat As Boolean

        FindSat = False

        If TLE_ListBox.SelectedItem = "favorites.txt" Then 'dans le cas de fichier TLE "FAVORIS"
            Try

                'Gère l'action après suppression d'un sat. dans les favoris
                If Sat_ListBox.SelectedItem = Nothing Then Exit Sub
                FileOpen(1, FavAdress, OpenMode.Input)
                While Not EOF(1)
                    sat1 = LineInput(1)
                    sat2 = LineInput(1)
                    If RTrim(Me.Sat_ListBox.SelectedItem) = sat2 Then
                        TLEFILE = TLEPath & sat1
                        TLETXT = sat1
                        FindSat = True
                        Exit While
                    End If
                End While
                FileClose(1)

                If FindSat = False AndAlso DeleteFavMode = False Then MessageBox.Show(sat2 & " wasn't found in TLE lists.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

            Catch Ex As Exception
                FileClose(1)
                MessageBox.Show(FavAdress & " seems to be corruped :" & vbCrLf & "Please read User Manual for help." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            End Try

        End If

        'On récupère les 2 lignes en fonction du satellite choisi
        Try

            FileOpen(1, TLEFILE, OpenMode.Input)
            While Not EOF(1)
                sat = CheckCaracts(RTrim(LineInput(1)))
                sat1 = LineInput(1)
                sat2 = LineInput(1)
                SATNAME = sat

                If Sat_ListBox.SelectedItem = sat Then Exit While
            End While

            Me.TLETextBox.Text = sat1 & vbCrLf & sat2
            FileClose(1)

            Me.TLETextBox.ReadOnly = True

            'DetailTLE()
            Me.ImportTLEButton.Enabled = False
            Me.DetailsTLEButton.Enabled = True
            ImportTLE()

        Catch ex As Exception
            MessageBox.Show(TLEFILE & " seems to be corruped :" & vbCrLf & "Please use TLE Updater." & vbCrLf & vbCrLf & ex.Message, "TLE ANALYSER - Error")
        End Try

    End Sub

    Private Sub SATSEARCH_TextChanged(ByVal sender As System.Object, ByVal e As KeyEventArgs) Handles SATSEARCH.KeyUp

        If e.KeyCode = Keys.Enter Then SATCHECK_module()

    End Sub

    Private Sub SATCHECK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SATCHECK.Click

        SATCHECK_module()

    End Sub

    Private Sub ModifTLEButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModifTLEButton.Click
        TLETextBox.ReadOnly = False
        TLETextBox.ForeColor = Color.Red
        ImportTLEButton.Enabled = True
    End Sub

    Sub ImportTLE()

        If CheckTLE() = False Then Exit Sub 'Vérifie si le TLE est conforme
        TLETextBox.ForeColor = Color.Black
        TLETextBox.ReadOnly = True

        TLETime = False
        LoadTLE()                           'Charge les paramètres du TLE
        TLELoaded = True
        ActualiseModule(True)               'Affiche les éléments 2D, 3D, GGMap
        'TLELoaded = False

        'Paramètres du menu (texte)
        ReloadTLEMenu.Enabled = True
        SaveFavMenu.Enabled = True
        ExportGmatMenu.Enabled = True
        GoogleMenu.Enabled = True
        CelestiaMenu.Enabled = True

        'Paramètres du menu (icones)
        ButtonMenu_Save.Enabled = True
        ButtonMenu_GMAT.Enabled = True
        ButtonMenu_Celes.Enabled = True
        ButtonMenu_GE.Enabled = True
        ButtonMenu_Reload.Enabled = True
        ButtonMenu_Delete.Enabled = True
        ButtonMenu_Summ.Enabled = True
        ButtonMenu_Charts.Enabled = True

        ModifTLEButton.Enabled = True

        If CBool(SimulOn) = True Then
            SimuListe1.SelectedIndex = 2 '1'
            SimuListe2.SelectedIndex = 0 'seconde'
            SimuSBSP.Enabled = True
            Timer2.Enabled = True
        End If

    End Sub

    Sub DetailTLE()
        Dim i As Integer
        Dim LineNum1, SatNum, Classif, ID1, ID2, ID3, EpoquYear, EpoquDay, MeanMotionDer1, MeanMotionDer2, _
            DragTerm, EphType, ElNum, CheckSum1, LineNum2, Inc, Raan, Ecc, ArgP, MeanA, MeanMotion, RevNum, CheckSum2, SMAval As String

        'Check TLE
        'Line Integrity
        If TLETextBox.Text = "" OrElse TLETextBox.Text.Length < 69 Then
            MessageBox.Show("Please check TLE:" & vbLf & "Line 1 andalso/orelse Line 2 seems to be empty orelse incorect.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Lines Beginning
        If TLETextBox.Lines(0).StartsWith("1 ") = False OrElse TLETextBox.Lines(1).StartsWith("2 ") = False Then
            MessageBox.Show("Please check TLE:" & vbLf & "Line 1 functions.Must begin with '1 '" & vbLf & "Line 2 functions.Must begin with '2 ' ", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Characters Numbers
        Dim Line0, Line1 As Integer
        Line0 = TLETextBox.Lines(0).Length
        Line1 = TLETextBox.Lines(1).Length
        If Line0 < 69 OrElse Line1 < 69 Then
            MessageBox.Show("Please check TLE: each line functions.Must contain 69 characters." & vbLf & vbLf & _
                   "Actually: " & vbLf & Line0 & " characters on line 1" & vbLf & Line1 & " characters on line 2", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        If Me.DetailsTLEButton.Text = "HIDE DETAILS" Then
            Me.GroupBox1.Height = 111
            Me.DetailsTLEButton.Text = "SHOW DETAILS"
        ElseIf Me.DetailsTLEButton.Text = "SHOW DETAILS" Then
            Me.GroupBox1.Height = 530
            Me.DetailsTLEButton.Text = "HIDE DETAILS"
        End If

        'Extract from TLE
        'Line Number
        LineNum1 = ""
        For i = 0 To 0
            LineNum1 = LineNum1 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value1.Text = LineNum1

        'Satellite number
        SatNum = ""
        For i = 2 To 6
            SatNum = SatNum & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value2.Text = SatNum

        'Classification
        Classif = ""
        i = 7
        Classif = Classif & TLETextBox.Lines(0).Chars(i)
        Me.Value3.Text = Classif

        'ID 1
        ID1 = ""
        For i = 9 To 10
            ID1 = ID1 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value4.Text = ID1

        'ID 2
        ID2 = ""
        For i = 11 To 13
            ID2 = ID2 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value5.Text = ID2

        'ID 3
        ID3 = ""
        For i = 14 To 16
            ID3 = ID3 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value6.Text = ID3

        'Epoqu Year
        EpoquYear = ""
        For i = 18 To 19
            EpoquYear = EpoquYear & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value7.Text = EpoquYear

        'Epoqu Day
        EpoquDay = ""
        For i = 20 To 31
            EpoquDay = EpoquDay & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value8.Text = EpoquDay

        'Mean Motion 1
        MeanMotionDer1 = ""
        For i = 33 To 42
            MeanMotionDer1 = MeanMotionDer1 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value9.Text = MeanMotionDer1

        'Mean Motion 2
        MeanMotionDer2 = ""
        For i = 44 To 51
            MeanMotionDer2 = MeanMotionDer2 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value10.Text = MeanMotionDer2

        'Drag Term
        DragTerm = ""
        For i = 53 To 60
            DragTerm = DragTerm & TLETextBox.Lines(0).Chars(i)
        Next
        DragTerm = 0.00001 * Double.Parse(Mid(DragTerm, 1, 6)) * 10 ^ Double.Parse(Mid(DragTerm, 7, 2))
        Me.Value11.Text = DragTerm

        'Ephemeris Type
        EphType = ""
        For i = 62 To 62
            EphType = EphType & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value12.Text = EphType

        'Element Number
        ElNum = ""
        For i = 64 To 67
            ElNum = ElNum & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value13.Text = ElNum

        'Check Sum
        CheckSum1 = ""
        For i = 68 To 68
            CheckSum1 = CheckSum1 & TLETextBox.Lines(0).Chars(i)
        Next
        Me.Value14.Text = CheckSum1

        'Line Number
        LineNum2 = ""
        For i = 0 To 0
            LineNum2 = LineNum2 & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value15.Text = LineNum2

        Me.Value16.Text = SatNum

        'Inclinaison
        Inc = ""
        For i = 8 To 15
            Inc = Inc & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value17.Text = Inc

        'RAAN
        Raan = ""
        For i = 17 To 24
            Raan = Raan & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value18.Text = Raan

        'Eccentricity
        Ecc = "0."
        For i = 26 To 32
            Ecc = Ecc & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value19.Text = Ecc

        'Argument of Perigee
        ArgP = ""
        For i = 34 To 41
            ArgP = ArgP & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value20.Text = ArgP

        'Mean Anomaly
        MeanA = ""
        For i = 43 To 50
            MeanA = MeanA & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value21.Text = MeanA

        'Mean Motion
        MeanMotion = ""
        For i = 52 To 62
            MeanMotion = MeanMotion & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value22.Text = MeanMotion

        'Revolution Number
        RevNum = ""
        For i = 63 To 67
            RevNum = RevNum & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value23.Text = RevNum

        'Check Sum
        CheckSum2 = ""
        For i = 68 To 68
            CheckSum2 = CheckSum2 & TLETextBox.Lines(1).Chars(i)
        Next
        Me.Value24.Text = CheckSum2

    End Sub

    Sub LoadTLE()

        '----------------------------------------------------------------------------------------------------------------------------
        'DECLARATION DES VARIABLES
        '----------------------------------------------------------------------------------------------------------------------------
        Dim i As Integer
        Dim SatNum, EpoquYear, EpoquDay, MeanMotionDer1, MeanMotionDer2, DragTerm, RevNum, ID1, ID2, ID3 As String
        Dim SMAKep, MeanMouvSecVar, AopSecVar, KepPeriod As Double
        Dim GST2

        '----------------------------------------------------------------------------------------------------------------------------
        'Check TLE
        '----------------------------------------------------------------------------------------------------------------------------
        If CheckTLE() = False Then Exit Sub 'Vérifie si le TLE est conforme

        '----------------------------------------------------------------------------------------------------------------------------
        'Extract LINES from TLE
        '----------------------------------------------------------------------------------------------------------------------------
        LIGNE1 = TLETextBox.Lines(0)
        LIGNE2 = TLETextBox.Lines(1)

        SatNum = "" 'Satellite number
        For i = 2 To 6
            SatNum = SatNum & TLETextBox.Lines(0).Chars(i)
        Next
        SatNum = Val(SatNum)
        Me.Spacecraft.Text = SatNum

        'ID 1
        ID1 = ""
        For i = 9 To 10
            ID1 = ID1 & TLETextBox.Lines(0).Chars(i)
        Next
        If ID1 > 0 AndAlso ID1 < 30 Then
            ID1 = "20" & ID1
        Else
            ID1 = "19" & ID1
        End If

        'ID 2
        ID2 = ""
        For i = 11 To 13
            ID2 = ID2 & TLETextBox.Lines(0).Chars(i)
        Next

        'ID 3
        ID3 = ""
        For i = 14 To 16
            ID3 = ID3 & TLETextBox.Lines(0).Chars(i)
        Next

        'Catalog Number
        CATNUMB = ID1 & "-" & ID2 & ID3
        CATNbre.Text = RTrim(CATNUMB)

        'Orbit Number
        TLEONAE = ""
        For i = 63 To 67
            TLEONAE = TLEONAE & TLETextBox.Lines(1).Chars(i)
        Next

        '----------------------------------------------------------------------------------------------------------------------------
        'EPOCH DETERMINATION
        '----------------------------------------------------------------------------------------------------------------------------
        EpoquYear = ""
        For i = 18 To 19
            EpoquYear = EpoquYear & TLETextBox.Lines(0).Chars(i)
        Next
        If EpoquYear > 0 AndAlso EpoquYear < 30 Then
            EpoquYear = "20" & EpoquYear
        Else
            EpoquYear = "19" & EpoquYear
        End If
        EpoquYear = Val(EpoquYear)

        EpoquDay = ""
        For i = 20 To 31
            EpoquDay = EpoquDay & TLETextBox.Lines(0).Chars(i)
        Next
        EpoquDay = Val(EpoquDay)

        DELTAD = ""
        For i = 23 To 31
            DELTAD = DELTAD & TLETextBox.Lines(0).Chars(i)
        Next

        EPOCH = ((1721424.5 - Int((EpoquYear - 1) / 100) + Int((EpoquYear - 1) / 400) + Int(365.25 * (EpoquYear - 1)) + EpoquDay) - 2430000).ToString
        EPOCH0 = 2430000 + (1721424.5 - Int((EpoquYear - 1) / 100) + Int((EpoquYear - 1) / 400) + Int(365.25 * (EpoquYear - 1)) + (EpoquDay - DELTAD) - 2430000)

        TLEPOCH = EPOCH

        'Current Date or TLE Date
        If TLETime = False Then
            EPOCH = GregtoMJD2(CurrentDateToGreg())
            DELTAD = (CDbl(EPOCH + 0.5)) - Truncate(CDbl(EPOCH))
            EPOCH0 = 2430000 + EPOCH - DELTAD
        Else
            EPOCH = TLEPOCH
        End If

        EpochFormat.Enabled = True
        EpochFormat.SelectedIndex = 1
        EPOCHBox.Text = MJDtoGreg2(EPOCH)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS CARTESIENS ET KEPLERIENS VIA SGP4 (CLASSES PREVISAT)
        'A UNE EPOCH DONNEE (TLE ou CURRENT DATE)
        '----------------------------------------------------------------------------------------------------------------------------

        SGP4(EPOCH, TLEPOCH)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DE LA LATITUDE/LONGITUDE DU SATELLITE A LA DATE TLE
        '----------------------------------------------------------------------------------------------------------------------------
        GSTTLE = GSTCalc(TLEPOCH)
        SGP4_TLE_EPOCH(TLEPOCH, TLEPOCH)
        LATLONG(XTLE, YTLE, ZTLE, GSTTLE)
        LATTLE = LAT
        LONGITLE = LONGI

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS SECONDAIRES
        '----------------------------------------------------------------------------------------------------------------------------
        ElementSecondaires()

        '----------------------------------------------------------------------------------------------------------------------------
        'REMPLISSAGE DES BOXES (en mode classic uniquement)
        '----------------------------------------------------------------------------------------------------------------------------
        'If TrackMode1 = True Then 
        Boxes()

        '----------------------------------------------------------------------------------------------------------------------------
        'PARAMETRAGE DU MODULE
        '----------------------------------------------------------------------------------------------------------------------------
        SCName.Text = SATNAME
        EpochFormat.SelectedIndex = "1"

        'Paramétrage des boutons
        NowDate.Enabled = True
        EPOCHBox.Enabled = True

        'Paramétrage de la simulation
        SimuListe1.SelectedIndex = 2 '1
        SimuListe2.SelectedIndex = 1 'min

        'Paramètrage Map
        TrackOptionPanel.Enabled = True

        MapPeriodNbr.Enabled = True
        MapPeriodNbr.SelectedItem = "1"
        MapShowTrack.Enabled = True
        MapShowTrack.Checked = True

        FullTrackCB.Checked = False

        'MapShowGrid.Enabled = True
        'MapShowGrid.Checked = False
        TracePicture.Visible = True

        MapOptions.Enabled = True

    End Sub

    Sub ElementSecondaires()

        '----------------------------------------------------------------------------------------------------------------------------
        'DECLARATION DES VARIABLES
        '----------------------------------------------------------------------------------------------------------------------------
        Dim i As Integer
        Dim SMAKep, MeanMouvSecVar, AopSecVar, KepPeriod As Double
        Dim GST2

        '----------------------------------------------------------------------------------------------------------------------------
        'VARIATIONS SECULAIRES
        '----------------------------------------------------------------------------------------------------------------------------
        MM = Maths.RAD2DEG * (Sqrt(Mu / SMA ^ 3)) * 240
        APER = 1440 / MM
        SMAKep = (((APER * 60) ^ 2 * Mu) / (4 * PI ^ 2)) ^ (1 / 3)
        MeanMouvSecVar = (3 / (4 * (1 - ECC ^ 2) ^ (3 / 2))) * J2 * (EarthEquRad / SMAKep) ^ 2 * (3 * Cos(Maths.DEG2RAD * (INC)) ^ 2 - 1) 'Dn/n
        AopSecVar = (3 / (4 * (1 - ECC ^ 2) ^ 2)) * J2 * (EarthEquRad / SMAKep) ^ 2 * (5 * Cos(Maths.DEG2RAD * (INC)) ^ 2 - 1) 'Dw/n

        '----------------------------------------------------------------------------------------------------------------------------
        'ELEMENTS SECONDAIRES
        '----------------------------------------------------------------------------------------------------------------------------
        'Position on Orbit vs AN
        ALPHA = (AOP + Maths.RAD2DEG * (TA)) Mod 360

        'Elapsed Time Form Periapsis
        ETFP = (Maths.DEG2RAD * (MA) / (Sqrt(Mu / SMA ^ 3))) / 60
        ETFP_Label.Text = "min"

        'Precession/Mvmt
        KepPeriod = 2 * PI * Sqrt(SMA ^ 3 / Mu)
        NP = (Maths.RAD2DEG * (-(3 / 2) * ((EarthEquRad ^ 2) / (SMA * (1 - ECC ^ 2)) ^ 2) * J2 * (2 * PI / KepPeriod) * Cos(Maths.DEG2RAD * (INC)))) * 86400
        AP = (3 / 4) * (Maths.RAD2DEG * (2 * PI / KepPeriod) * 3600 * 24) * J2 * (EarthEquRad / SMA) ^ (2) * ((5 * (Cos(Maths.DEG2RAD * (INC)) ^ 2) - 1))

        'Delta Longitude au NA
        Dim FreqPhas, MeanMotionNA As Double
        MeanMotionNA = MM * 360 + (AP)
        FreqPhas = MeanMotionNA / (EarthNodPrec * 360 - NP)
        DL = (360 / FreqPhas)
        DLong_Label.Text = "deg"

        'Periodes
        APER = 1440 / MM
        DPER = (APER / (1 + (AopSecVar)))

        AnoPeriod_Label.Text = "min"
        DracoPeriod_Label.Text = "min"

        'Distance relative (altitude du satellite exprimé en rayon terrestre)
        DREL = Sqrt(X ^ 2 + Y ^ 2 + Z ^ 2) / EarthEquRad

        'Temps écoulé depuis TLE Epoque
        ETFE = EPOCH - TLEPOCH

        'Détermination du Nombre d'orbite à current Epoch
        ONAE = Round(TLEONAE + ((EPOCH - TLEPOCH) * MM))

        'Adapted Parameters determination
        AdpatedParameters()

        '----------------------------------------------------------------------------------------------------------------------------
        'Gestion du Phasage
        '----------------------------------------------------------------------------------------------------------------------------
        Dim K0 = (EarthNodPrec * 360 - NP) / 360
        Dim nd = Val(MM) * 360 + AP
        Dim k = nd / (EarthNodPrec * 360 - NP)
        Dim v0 = Round(Round(k, 1), 0)
        Dim v = 1440 / Val(DPER)
        Dim CT0, DT0 As Double
        Dim iter As Integer
        Dim Spec As Double
        Dim FindNT0 As Boolean

        If NP < 0.975 OrElse NP > 0.995 Then
            'Dans le cas de satellitie NON Heliosynchrone
            '   Determination de NT0 par itération sur CT0
            FindNT0 = False
            For Spec = 0.01 To 0.05 Step 0.0001
                For iter = 1 To 38 'on itère sur 38 jours uniquement
                    NT0 = iter * k
                    If NT0 - Round(NT0, 0) < Spec AndAlso NT0 - Round(NT0, 0) > -Spec Then
                        CT0 = iter
                        FindNT0 = True
                        Exit For
                    End If
                Next
                If FindNT0 = True Then Exit For
            Next
            DT0 = Round(NT0 - v0 * CT0)
        ElseIf NP > 0.975 AndAlso NP < 0.995 Then
            'Dans le cas de satellitie Heliosynchrone
            '   Determination de NT0 par itération sur CT0
            FindNT0 = False
            For Spec = 0.01 To 0.05 Step 0.001
                For iter = 1 To 38
                    NT0 = iter * v
                    If NT0 - Round(NT0, 0) < Spec AndAlso NT0 - Round(NT0, 0) > -Spec Then
                        CT0 = iter
                        FindNT0 = True
                        Exit For
                    End If
                Next
                If FindNT0 = True Then Exit For
            Next

            DT0 = Round((v - v0) * CT0, 0)

        End If
        Me.phase.Text = "[ " & v0 & " ; " & DT0 & " ; " & CT0 & " ] " & Round(NT0, 0)

        'Gestion du Full Track
        If CT0 < 15 Then
            Me.FullTrackCB.Enabled = True
            NT0 = Round(NT0, 0)
        Else
            Me.FullTrackCB.Checked = False
            Me.FullTrackCB.Enabled = False
            NT0 = Nothing
        End If

        '----------------------------------------------------------------------------------------------------------------------------
        'Calcul de RAGreenwich ou GW Sideral Time (GST)
        '----------------------------------------------------------------------------------------------------------------------------
        GST = GSTCalc(Val(EPOCH))

        '----------------------------------------------------------------------------------------------------------------------------
        'Latitude/Longitude du satellite
        '----------------------------------------------------------------------------------------------------------------------------
        LATLONG(X, Y, Z, GST)

        '----------------------------------------------------------------------------------------------------------------------------
        'ALTITUDE REELLE DU SATELLITE
        '----------------------------------------------------------------------------------------------------------------------------
        Dim SatRad, Rt

        SatRad = Norme(X, Y, Z)
        Rt = EarthEquRad / (Sqrt((Cos(Maths.DEG2RAD * (LAT)) ^ 2) + ((Sin(Maths.DEG2RAD * (LAT)) ^ 2) / (1 - EarthFlat) ^ 2)))

        'Altitudes
        APR = SMA * (1 + ECC) 'ApA + EarthEquRad
        PER = SMA * (1 - ECC) 'PeA + EarthEquRad
        APA = APR - Rt
        PEA = PER - Rt
        ALT = SatRad - Rt
        MALT = SMA - Rt

        'Vitesses
        APV = Sqrt(Mu * (2 / APR - 1 / SMA))
        PEV = Sqrt(Mu * (2 / PER - 1 / SMA))
        VEL = (Sqrt(Mu * ((2 / (ALT + Rt)) - (1 / SMA))))

        '----------------------------------------------------------------------------------------------------------------------------
        'POSITION DU SOLEIL
        '----------------------------------------------------------------------------------------------------------------------------
        Dim Soleil As New Soleil
        Dim OffsetUTCJR = Val(OffsetUTC) / 24

        'Converti la date actuelle en JJ
        Dim ChaineEpoch, EpochDate, EpochTime, ChaineDate, ChaineTime, Jour, Mois, Année, Hre, Min, Sec

        ChaineEpoch = CStr(MJDtoGreg2(EPOCH)).Split(" ")
        EpochDate = ChaineEpoch(0)
        EpochTime = ChaineEpoch(1)

        ChaineDate = EpochDate.split("/")
        Jour = Val(ChaineDate(0))
        Mois = Val(ChaineDate(1))
        Année = Val(ChaineDate(2))

        ChaineTime = EpochTime.split(":")
        Hre = Val(ChaineTime(0))
        Min = Val(ChaineTime(1))
        Sec = Val(ChaineTime(2))

        Dim dat1 As New Dates(Année, Mois, Jour, Hre, Min, Sec, OffsetUTCJR)
        Dim dat As New Dates(dat1, )

        'Calcul la position du Soleil
        Soleil.CalculPosition(dat)
        XS = Soleil.Position.X
        YS = Soleil.Position.Y
        ZS = Soleil.Position.Z

        '----------------------------------------------------------------------------------------------------------------------------
        'LONGITUDE ET LATITUDE DU SOLEIL
        '----------------------------------------------------------------------------------------------------------------------------
        Dim r0, Lats0, c, sph
        'Longitude
        LONGIS = Maths.RAD2DEG * ((Atan2(YS, XS) - Maths.DEG2RAD * (GST)) Mod Maths.DEUX_PI)
        If LONGIS > 180 Then LONGIS -= 360
        If LONGIS < -180 Then LONGIS += 360

        'Latitude
        r0 = Sqrt(XS * XS + YS * YS)
        Lats0 = Atan2(ZS, r0)
        Do
            LATS = Lats0
            sph = Sin(LATS)
            c = 1.0 / Sqrt(1.0 - Terre.E2 * sph * sph)
            Lats0 = Atan((ZS + Terre.RAYON * c * Terre.E2 * sph) / r0)
        Loop While Abs(Lats0 - LATS) > 0.0000001
        LATS = Maths.RAD2DEG * (LATS)

        '----------------------------------------------------------------------------------------------------------------------------
        'POSITION DU SOLEIL SUR LA MAP
        '----------------------------------------------------------------------------------------------------------------------------

        Dim latspix, longspix
        Dim MapWS2 = MapW / 2
        Dim MapHS2 = MapH / 2

        'Convertir les données lat/long en pixel
        latspix = (Round(MapHS2 - (LATS * MapH / 180), 0))
        longspix = (Round(MapWS2 + (LONGIS * MapW / 360), 0))

        lsol = longspix
        bsol = latspix

        '----------------------------------------------------------------------------------------------------------------------------
        'LAT/LNG SUR LA MAP 2D STD
        '----------------------------------------------------------------------------------------------------------------------------
        Dim LongDirection, LatDirection
        Dim SunLongDirection, SunLatDirection

        'Satellite
        If LONGI > 0 Then
            LongDirection = " E"
            LONGBox.Text = Round(LONGI, 3) & LongDirection
        ElseIf LONGI < 0 Then
            LongDirection = " W"
            LONGBox.Text = Round(LONGI, 3) & LongDirection
        End If

        If LAT > 0 Then
            LatDirection = " N"
            LATBox.Text = Round(LAT, 3) & LatDirection
        ElseIf LAT < 0 Then
            LatDirection = " S"
            LATBox.Text = Round(LAT, 3) & " S"
        End If

        LNGLATBOX.Text = Round(LONGI, 3).ToString("000.000") & LongDirection & " | " & Round(LAT, 3).ToString("000.000") & LatDirection

        'Soleil
        If LATS < 0 Then
            SunLatDirection = " S"
        ElseIf LATS >= 0 Then
            SunLatDirection = " N"
        End If
        If LONGIS < 0 Then
            SunLongDirection = " W"
        ElseIf LONGIS >= 0 Then
            SunLongDirection = " E"
        End If

        LNGLATSUNBOX.Text = Round(LONGIS, 3).ToString("000.000") & SunLongDirection & " | " & Round(LATS, 3).ToString("000.000") & SunLatDirection

        '----------------------------------------------------------------------------------------------------------------------------
        'Ascending node
        '----------------------------------------------------------------------------------------------------------------------------
        Dim LANValue As Double

        LANValue = LAN_F(ECC, MA, APER, RAAN, AOP, DELTAD, EPOCH0) Mod 360
        If LANValue > 180 Then
            LAN = LANValue - 360
        ElseIf LANValue < -180 Then
            LAN = 360 + LANValue
        Else
            LAN = LANValue
        End If
        Me.LANType.Text = "[-180;180]"

        'Local Time of AN
        LTAN = LTAN_F()

        '----------------------------------------------------------------------------------------------------------------------------
        'Calcul de LST (Local Sidereal Time)
        '----------------------------------------------------------------------------------------------------------------------------
        If LONGI < 0 Then
            LST = ((LONGI + 360) + GST) Mod 360
        Else
            LST = (LONGI + GST) Mod 360
        End If

        '----------------------------------------------------------------------------------------------------------------------------
        'GESTION DE L'ECLIPSE
        '----------------------------------------------------------------------------------------------------------------------------
        Dim ANGLEX, ANGLEX0

        ANGLEX = Acos(ProduitScalaire(Normalise(X, Y, Z), Normalise(-XS, -YS, -ZS)))
        ANGLEX0 = Asin(1 / DREL)

        If ANGLEX < ANGLEX0 Then
            EclipseBox.Text = "YES"
        Else
            EclipseBox.Text = "NO"
        End If

        '----------------------------------------------------------------------------------------------------------------------------
        ''Gestion de la dérive en longitude et Latitude
        '----------------------------------------------------------------------------------------------------------------------------
        Dim Beta = 0.001561 'deg.j-2
        Dim LONGID = LONGI
        Dim LongAcc, NDEG2, NDEG

        Dim LATREF, LONGREF
        If CBool(REFTLE) = True Then
            LATREF = LATTLE
            LONGREF = LONGITLE
        ElseIf CBool(REFRM) = True Then
            LATREF = LatMis
            LONGREF = LngMis
        End If

        'If LONGREF < 0 Then LONGREF = 360 + LONGREF
        Dim LONGP = LONGREF + Val(LngWin)
        Dim LONGM = LONGREF - Val(LngWin)
        Dim LATP = LATREF + Val(LatWin)
        Dim LATM = LATREF - Val(LatWin)

        If MM > 0.99 AndAlso MM < 1.01 AndAlso ECC < 0.01 AndAlso INC < 1 Then
            LongDriftPanel.Visible = True

            'Acceleration Longitudinale (sous l'effet des termes tesséraux du potentiel terrestre > modification des SMA)
            If LONGI < 0 Then LONGID = 360 + LONGI
            LongAcc = Round(Beta * ((1.087 * Sin(2 * Maths.DEG2RAD * (LONGID - L22))) - (0.05 * Sin(LONGID - L31)) + (0.15 * Sin(3 * (LONGID - L33)))), 5)

            'Nombre de deg (LONG) avant de sortir de la fenêtre
            LONGACCBox.Text = LongAcc
            If LongAcc < 0 Then
                NDEG = LONGI - LONGM
            ElseIf LongAcc > 0 Then
                NDEG = LONGP - LONGI
            End If
            If NDEG < 0 Then
                ExitWinDeg.ForeColor = Color.Red
                NDEG = "OUT"
                ExitWinDeg.Text = NDEG
            Else
                ExitWinDeg.ForeColor = Color.Black
                ExitWinDeg.Text = Round(NDEG, 2)
            End If

            'Nombre de deg (LAT) avant de sortir de la fenêtre
            If LATP - LAT < LAT - LATM Then
                NDEG2 = Abs(LATP - LAT)
            ElseIf LATP - LAT > LAT - LATM Then
                NDEG2 = Abs(LAT - LATM)
            Else
                NDEG2 = Abs(LATP - LAT)
            End If
            If LAT > LATP OrElse LAT < LATM Then
                ExitWinDeg2.ForeColor = Color.Red
                NDEG2 = "OUT"
                ExitWinDeg2.Text = NDEG2
            Else
                ExitWinDeg2.ForeColor = Color.Black
                ExitWinDeg2.Text = Round(NDEG2, 2)
            End If

        Else
            LongDriftPanel.Visible = False
        End If

    End Sub

    Sub ActualiseModule(ByVal AfficheTrace As Boolean)

        ActualiseGraphics(AfficheTrace)

        If MapsTab.SelectedIndex = 1 Then GoogleControl1.GGMapWeb() 'Affiche la vue Google Map
        If MapsTab.SelectedIndex = 2 Then
            GoogleEarthControl1.addSat()            'Affiche la vue Google Earth
            If TLELoaded = True Then GoogleEarthControl1.InitialViewF() 'Actualise la vue Earth sur le dernier satellite chargé
        End If

    End Sub

    Sub OpenTleFile(ByVal Fichier As String)

        Dim TLE, SAT

        Dim line, sat1, sat2 As String
        Dim ligne, SatNum, i As Integer

        'on vide la liste Sat
        Sat_ListBox.Items.Clear()
        'on sauvegarde le nom l'adresse du fichier
        TLEFILE = Fichier

        If TLE_ListBox.SelectedItem = "favorites.txt" Then 'dans le cas de fichier TLE "FAVORIS"

            'On créer le fichier favoris si besoin
            If My.Computer.FileSystem.FileExists(FavAdress) = False Then File.AppendAllText(FavAdress, "")

            Try
                'On récupère le nom de chaque Satellite puis on l'ajoute dans la liste Sat_ListBox
                FileOpen(1, FavAdress, OpenMode.Input)
                While Not EOF(1)
                    TLE = LineInput(1)
                    SAT = LineInput(1)
                    Sat_ListBox.Items.Add(SAT)
                End While
                FileClose(1)

            Catch Ex As Exception
                FileClose(1)
                MessageBox.Show(FavAdress & " seems to be corruped :" & vbCrLf & "Please read User Manual for Help." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            End Try

        Else 'dans le cas des fichiers TLE standard

            Try
                'On récupère le nom de chaque Satellite puis on l'ajoute dans la liste Sat_ListBox
                FileOpen(1, Fichier, OpenMode.Input)
                While Not EOF(1)
                    SAT = CheckCaracts(RTrim(LineInput(1)))
                    sat1 = LineInput(1)
                    sat2 = LineInput(1)
                    Sat_ListBox.Items.Add(SAT)
                End While
                FileClose(1)

            Catch Ex As Exception
                FileClose(1)
                MessageBox.Show(Fichier & " seems to be corruped :" & vbCrLf & "&a TLE Updater." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            End Try
        End If

        Sat_ListBox.Sorted = True


    End Sub

    Sub SATCHECK_module()

        Dim TLEList
        Dim TLEName
        Dim SATName

        'On vérifie l'existance des dossiers/fichiers
        If My.Computer.FileSystem.DirectoryExists("C:\TLEAnalyser\TLE") Then _
            TLEList = My.Computer.FileSystem.FindInFiles("C:\TLEAnalyser\TLE", Me.SATSEARCH.Text, True, FileIO.SearchOption.SearchTopLevelOnly)

        If TLEList.count = 0 Then MessageBox.Show("No TLE was found", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        If TLEList.count >= 1 Then
            'Selectionne la bonne liste de TLE
            TLEName = My.Computer.FileSystem.GetName(TLEList(0))
            If TLEName = "favorites.txt" Then
                TLEName = My.Computer.FileSystem.GetName(TLEList(1))
            End If

            TLE_ListBox.SelectedItem = TLEName

            'Recherche le satellite dans la liste
            For i = 0 To Me.Sat_ListBox.Items.Count - 1 'boucle pour chaque nom de satellite
                For j = 1 To Len(Sat_ListBox.Items(i)) 'recherche de mot dans le nom du satellite
                    If Mid(Sat_ListBox.Items(i), j, Len(SATSEARCH.Text)) = SATSEARCH.Text Then
                        Sat_ListBox.SelectedItem = Sat_ListBox.Items(i)
                        Exit Sub
                    End If
                Next
            Next
        End If

    End Sub

    Sub Boxes()

        '----------------------------------------------------------------------------------------------------------------------------
        'REMPLISSAGE DES DATAS
        '----------------------------------------------------------------------------------------------------------------------------

        SMABox.Text = Round(SMA, 4)
        ECCBox.Text = Round(ECC, 4)
        INCBox.Text = Round(INC, 4)
        RAANBox.Text = Round(RAAN, 4)
        AOPBox.Text = Round(AOP, 4)

        MABox.Text = Round(MA, 4)
        TABox.Text = Round(Maths.RAD2DEG * (TA), 4)
        EABox.Text = Round(Maths.RAD2DEG * (EA), 4)
        ETFPBox.Text = Round(ETFP, 4)
        ALPHABox.Text = Round(ALPHA, 4)

        XBox.Text = Round(X, 4)
        YBox.Text = Round(Y, 4)
        ZBox.Text = Round(Z, 4)
        VXBox.Text = Round(VX, 4)
        VYBox.Text = Round(VY, 4)
        VZBox.Text = Round(VZ, 4)

        APABox.Text = Round(APA, 4)
        PEABox.Text = Round(PEA, 4)
        APRBox.Text = Round(APR, 4)
        PERBox.Text = Round(PER, 4)
        ALTBox.Text = Round(ALT, 4)
        MALTBox.Text = Round(MALT, 4)
        APVBox.Text = Round(APV, 4)
        PEVBox.Text = Round(PEV, 4)
        VELBox.Text = Round(VEL, 4)

        LANBox.Text = Round(LAN, 4)
        LTANBox.Text = LTAN

        APERBox.Text = Round(APER, 4)
        DPERBox.Text = Round(DPER, 4)

        MMBox.Text = Round(MM, 4)
        NPBox.Text = Round(NP, 4)
        APBox.Text = Round(AP, 4)
        DLBox.Text = Round(DL, 4)

        RelDist.Text = Round(DREL, 4)
        ETFEBox.Text = Round(ETFE, 4)
        OrbNumBox.Text = ONAE

        GSTBox.Text = Round(GST, 4)
        LSTBox.Text = Round(LST, 4)

    End Sub

    'LINKS

    Private Sub ETFP_Label_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ETFP_Label.LinkClicked
        On Error GoTo Err
        If Me.ETFP_Label.Text = "min" Then
            Me.ETFPBox.Text = Round4(ETFP * 60)
            Me.ETFP_Label.Text = "sec"
        ElseIf Me.ETFP_Label.Text = "sec" Then
            Me.ETFPBox.Text = Round4(ETFP)
            Me.ETFP_Label.Text = "min"
        End If
Err:
        Exit Sub
    End Sub

    Private Sub AnoPeriod_Label_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles AnoPeriod_Label.LinkClicked
        On Error GoTo Err
        If Me.AnoPeriod_Label.Text = "min" Then
            Me.APERBox.Text = Round4(APER * 60)
            Me.AnoPeriod_Label.Text = "sec"
        ElseIf Me.AnoPeriod_Label.Text = "sec" Then
            Me.APERBox.Text = Round4(APER)
            Me.AnoPeriod_Label.Text = "min"
        End If
Err:
        Exit Sub
    End Sub

    Private Sub DracoPeriod_Label_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles DracoPeriod_Label.LinkClicked
        On Error GoTo Err
        If Me.DracoPeriod_Label.Text = "min" Then
            Me.DPERBox.Text = Round4(DPER * 60)
            Me.DracoPeriod_Label.Text = "sec"
        ElseIf Me.DracoPeriod_Label.Text = "sec" Then
            Me.DPERBox.Text = Round4(DPER)
            Me.DracoPeriod_Label.Text = "min"
        End If
Err:
        Exit Sub
    End Sub

    Private Sub DLong_Label_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles DLong_Label.LinkClicked
        On Error GoTo Err
        If Me.DLong_Label.Text = "deg" Then
            Me.DLBox.Text = Round4(Maths.DEG2RAD * (DL) * EarthEquRad)
            Me.DLong_Label.Text = "km"
        ElseIf Me.DLong_Label.Text = "km" Then
            Me.DLBox.Text = Round4(DL)
            Me.DLong_Label.Text = "deg"
        End If
Err:
        Exit Sub
    End Sub

    Private Sub LAN_Label_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LAN_Label.LinkClicked
        On Error GoTo Err
        If Me.LANType.Text = "[-180;180]" Then
            If Val(Me.LANBox.Text) > 0 AndAlso Val(Me.LANBox.Text) < 180 Then
                Me.LANType.Text = "[0;360]"
            Else
                Me.LANBox.Text = 360 + Val(Me.LANBox.Text)
                Me.LANType.Text = "[0;360]"
            End If
        ElseIf Me.LANType.Text = "[0;360]" Then
            If Val(Me.LANBox.Text) > 180 Then
                Me.LANBox.Text = Val(Me.LANBox.Text) - 360
                Me.LANType.Text = "[-180;180]"
            Else
                Me.LANType.Text = "[-180;180]"
            End If
        End If
Err:
        Exit Sub
    End Sub

    'EPOCH

    Private Sub EpochFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EpochFormat.SelectedIndexChanged

        If CheckEpoch() = False Then Exit Sub

        If Me.EpochFormat.SelectedIndex = 1 Then

            Me.EPOCHBox.Text = MJDtoGreg2(EPOCH)

        ElseIf Me.EpochFormat.SelectedIndex = 0 Then

            Me.EPOCHBox.Text = EPOCH

        End If

    End Sub

    Private Sub Epoch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EPOCHBox.GotFocus
        Me.TrackOptionPanel.Enabled = False

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False
        Me.EpochFormat.Enabled = True
    End Sub

    Private Sub Epoch_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EPOCHBox.LostFocus

        TrackOptionPanel.Enabled = True

        If EpochFormat.SelectedIndex = 0 Then
            EPOCH = CDbl(EPOCHBox.Text)
        ElseIf EpochFormat.SelectedIndex = 1 Then
            EPOCH = GregtoMJD2(EPOCHBox.Text)
        End If
        Prediction()

        '----------------------------------------------------------------------------------------------------------------------------
        '2D/3D
        '----------------------------------------------------------------------------------------------------------------------------
        ActualiseModule(True)

    End Sub

    Private Sub Epoch_TextChanged_2(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles EPOCHBox.KeyUp

        If e.KeyCode = Keys.Enter Then
            SATSEARCH.Focus()
        End If

    End Sub

    Private Sub NowDate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles NowDate.LinkClicked

        NowDateF()

    End Sub

    Sub NowDateF()

        Dim MJDDate
        MJDDate = GregtoMJD2(CurrentDateToGreg())

        If Me.EpochFormat.SelectedIndex = "1" Then
            Me.EPOCHBox.Text = MJDtoGreg2(MJDDate)
        Else
            Me.EPOCHBox.Text = MJDDate
        End If

        EPOCH = CDbl(MJDDate)

        Prediction()

        '----------------------------------------------------------------------------------------------------------------------------
        '2D/3D
        '----------------------------------------------------------------------------------------------------------------------------
        TLELoaded = True
        ActualiseModule(True)

    End Sub

    '2D MAP / 2D TRACK

    Private Sub SiMuBefore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimuBefore.Click

        Dim Liste1 = Me.SimuListe1.SelectedItem
        Dim Liste2 = Me.SimuListe2.SelectedItem

        Me.Timer2.Enabled = False
        Me.EpochFormat.Enabled = False

        If Liste1 = "" OrElse Liste2 = "" Then
            MessageBox.Show("Please select prediction parameters.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Si tout est OK
        Me.Timer1.Enabled = True
    End Sub

    Private Sub SiMuAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimuAfter.Click

        Dim Liste1 = Me.SimuListe1.SelectedItem
        Dim Liste2 = Me.SimuListe2.SelectedItem

        Me.Timer1.Enabled = False
        Me.EpochFormat.Enabled = False

        If Liste1 = "" OrElse Liste2 = "" Then
            MessageBox.Show("Please select prediction parameters.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Si tout est OK
        Me.Timer2.Enabled = True
    End Sub

    Private Sub SimuPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimuPause.Click

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False
        Me.EpochFormat.Enabled = True

    End Sub

    Private Sub SimuSBSP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimuSBSP.Click

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        Dim Liste1 = Me.SimuListe1.SelectedItem
        Dim Liste2 = Me.SimuListe2.SelectedItem

        Me.EpochFormat.Enabled = False

        If Liste1 = "" OrElse Liste2 = "" Then
            MessageBox.Show("Please select prediction parameters.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Si tout est OK
        SimuSBSP.Enabled = False
        Timer2.Enabled = True

    End Sub

    Private Sub SimuSBSM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimuSBSM.Click

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        Dim Liste1 = Me.SimuListe1.SelectedItem
        Dim Liste2 = Me.SimuListe2.SelectedItem

        Me.EpochFormat.Enabled = False

        If Liste1 = "" OrElse Liste2 = "" Then
            MessageBox.Show("Please select prediction parameters.", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        'Si tout est OK
        SimuSBSM.Enabled = False
        Me.Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim par1, DracoPeriod
        Dim Liste1 = Me.SimuListe1.SelectedItem
        Dim Liste2 = Me.SimuListe2.SelectedItem

        Select Case CStr(Speed)
            Case "1"
                Timer1.Interval = 1000
            Case "10"
                Timer1.Interval = 100
            Case "100"
                Timer1.Interval = 10
        End Select

        'On converti la période draconitique en min
        If Me.DracoPeriod_Label.Text = "sec" Then
            DracoPeriod = DPER / 60
        ElseIf Me.DracoPeriod_Label.Text = "min" Then
            DracoPeriod = DPER
        End If

        Select Case Liste2
            Case "sec"
                par1 = 86400
            Case "min"
                par1 = 1440
            Case "hr"
                par1 = 24
            Case "day"
                par1 = 1
            Case "period"
                par1 = 1440 / DracoPeriod
        End Select

        EPOCH = EPOCH - (Liste1 * (1 / par1))

        If Me.EpochFormat.SelectedIndex = "0" Then
            Me.EPOCHBox.Text = EPOCH
        Else
            Me.EPOCHBox.Text = MJDtoGreg2(EPOCH)
        End If

        Prediction()
        TLELoaded = False
        ResizeMode = False
        ActualiseModule(False)
        TLELoaded = True

        'Gestion du Step par step
        If SimuSBSM.Enabled = False Then
            SimuSBSM.Enabled = True
            Timer1.Enabled = False
            EpochFormat.Enabled = True
        End If

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Dim par1, DracoPeriod
        Dim Liste1 = Me.SimuListe1.SelectedItem
        Dim Liste2 = Me.SimuListe2.SelectedItem
        Dim EventSBSP As System.EventArgs

        Select Case CStr(Speed)
            Case "1"
                Timer2.Interval = 1000
            Case "10"
                Timer2.Interval = 100
            Case "100"
                Timer2.Interval = 10
        End Select

        'On converti la période draconitique en min
        If Me.DracoPeriod_Label.Text = "sec" Then
            DracoPeriod = DPER / 60
        ElseIf Me.DracoPeriod_Label.Text = "min" Then
            DracoPeriod = DPER
        End If

        Select Case Liste2
            Case "sec"
                par1 = 86400
            Case "min"
                par1 = 1440
            Case "hr"
                par1 = 24
            Case "day"
                par1 = 1
            Case "period"
                par1 = 1440 / DracoPeriod
        End Select

        EPOCH = EPOCH + (Liste1 * (1 / par1))

        If Me.EpochFormat.SelectedIndex = "0" Then
            Me.EPOCHBox.Text = EPOCH
        Else
            Me.EPOCHBox.Text = MJDtoGreg2(EPOCH)
        End If

        Prediction()
        TLELoaded = False
        ResizeMode = False
        ActualiseModule(False)
        TLELoaded = True

        'Gestion du Step par step
        If SimuSBSP.Enabled = False Then
            SimuSBSP.Enabled = True
            Timer2.Enabled = False
            EpochFormat.Enabled = True
        End If

    End Sub

    Private Sub SatPicture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SatPicture.Click

        'GoogleControl1.GGMapWeb()
        'GoogleEarthControl1.addSat()
        'TabControl2.SelectedIndex = 2

    End Sub

    Private Sub MapPeriodNbr_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MapPeriodNbr.SelectedIndexChanged
        If TLELoaded = True Then

            FullTrackCB.Checked = False
            ActualiseModule(False)

            'ActualiseGraphics(False)

            'GoogleControl1.GGMapWeb()
            'GoogleEarthControl1.addSat()
        End If

    End Sub

    Private Sub MapShowTrack_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MapShowTrack.Click

        'If MapShowTrack.Checked = False Then
        '    MapShowTrack.Checked = True
        'End If

        If Me.MapShowTrack.Enabled = True AndAlso Me.MapShowTrack.Checked = True Then
            Me.MapPeriodNbr.Enabled = True
        ElseIf Me.MapShowTrack.Enabled = True AndAlso Me.MapShowTrack.Checked = False Then
            Me.MapPeriodNbr.Enabled = False
        End If

        Me.FullTrackCB.Checked = False

        ActualiseModule(False)

        'ActualiseGraphics(False)

    End Sub

    Private Sub MapShowGrid_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MapShowGrid.CheckedChanged

        'If Me.MapShowGrid.Enabled = True AndAlso Me.MapShowGrid.Checked = True Then

        '    Dim BGImg As New Bitmap(My.Resources.EarthMap2)
        '    Me.MapPanel.BackgroundImage = BGImg

        '    Me.MapLongScale.Visible = True
        '    Me.MapLatScale.Visible = True

        'ElseIf Me.MapShowGrid.Enabled = True AndAlso Me.MapShowGrid.Checked = False Then

        '    Dim BGImg As New Bitmap(My.Resources.EarthMap)
        '    Me.MapPanel.BackgroundImage = BGImg

        '    Me.MapLongScale.Visible = False
        '    Me.MapLatScale.Visible = False

        'End If

    End Sub

    Private Sub ActuTrack_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ActuTrack.LinkClicked
        ActualizeTrack()
    End Sub

    Private Sub FullTrackCB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FullTrackCB.CheckedChanged

        If Me.FullTrackCB.Enabled = True AndAlso Me.FullTrackCB.Checked = True Then

            SunShadowCB.Enabled = False

            g.Clear(TracePicture.BackColor)
            DrawSunShadow()
            FullTraceSC(NT0)
            ExportToMap()
            DrawSun()
            TracePicture.Image = Trace

        ElseIf Me.MapShowTrack.Enabled = True AndAlso Me.FullTrackCB.Checked = False Then
            SunShadowCB.Enabled = True
            ActualiseGraphics(True)
        End If

    End Sub

    Private Sub SunShadowCB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SunShadowCB.Click

        ActualiseGraphics(False)
        GoogleEarthControl1.DrawSun()

    End Sub

    Private Sub MapMouseLocal(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TracePicture.MouseMove

        Dim Xmouse, Ymouse, LongDirection, LatDirection
        Dim LongM, LatM As Double

        Dim MapW = MapPanel.Width
        Dim MapWS2 = MapW / 2
        Dim MapH = MapPanel.Height
        Dim MapHS2 = MapH / 2
        Dim OffsetY
        Dim Mouse As Point = PointToClient(MousePosition)

        'PIXEL
        If TrackMode1 = True Then OffsetY = 23
        If TrackMode2 = True Then OffsetY = 22.5

        Xmouse = Mouse.X - MapsTab.Location.X - MapPanel.Location.X - TracePicture.Location.X - 5
        Ymouse = Mouse.Y - MapsTab.Location.Y - MapPanel.Location.Y - TracePicture.Location.Y - OffsetY

        'LONGITUDE
        LongM = (Xmouse * 360) / MapW
        If Xmouse < MapWS2 Then
            LongM = -(180 - LongM)
            LongDirection = " W"
        ElseIf Xmouse >= MapWS2 Then
            LongM = LongM - 180
            LongDirection = " E"
        End If

        'LATITUDE
        LatM = (Ymouse * 180) / MapH
        If Ymouse < MapHS2 Then
            LatM = 90 - LatM
            LatDirection = " N"
        ElseIf Ymouse >= MapHS2 Then
            LatM = -(LatM - 90)
            LatDirection = " S"
        End If

        LNGLATBOX.Text = Round(LongM, 3).ToString("000.000") & LongDirection & " | " & Round(LatM, 3).ToString("000.000") & LatDirection

    End Sub

    Private Sub MapMouseLocal_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TABMAP.MouseEnter

        LNGLATBOX.Text = LONGBox.Text & " | " & LATBox.Text

    End Sub

    Sub ActualizeTrack()
        Timer1.Enabled = False
        Timer2.Enabled = False
        EpochFormat.Enabled = True
        If FullTrackCB.Checked = True Then
            FullTrackCB.Checked = False
            MapShowTrack.Checked = True
        End If
        Me.MapPeriodNbr.SelectedIndex = 0

        ActualiseModule(True)

        'ActualiseGraphics(True)

        'If MapsTab.SelectedIndex = 1 Then GoogleControl1.GGMapWeb()
        'If MapsTab.SelectedIndex = 2 Then GoogleEarthControl1.addSat()
    End Sub

    'PREVISIONS - TRACKING

    Sub SGP4(ByVal EPOCH As Double, ByVal TLEPOCH As Double)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS CARTESIENS ET KEPLERIENS VIA SGP (CLASSES PREVISAT)
        '----------------------------------------------------------------------------------------------------------------------------
        Dim TLE As New TLE(LIGNE1, LIGNE2)
        Dim Sat As New Satellite(TLE)

        Sat.CalculPosVit(EPOCH, TLEPOCH)

        X = Sat.Position.X
        Y = Sat.Position.Y
        Z = Sat.Position.Z
        VX = Sat.Vitesse.X
        VY = Sat.Vitesse.Y
        VZ = Sat.Vitesse.Z

        CartesianToKeplerian(X, Y, Z, VX, VY, VZ)

    End Sub

    Sub SGP4Trace(ByVal EPOCH_NEW As Double, ByVal TLEPOCH As Double)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS CARTESIENS VIA SGP POUR DETERMINER LA TRACE (CLASSES PREVISAT)
        '----------------------------------------------------------------------------------------------------------------------------
        Dim TLE As New TLE(LIGNE1, LIGNE2)
        Dim Sat As New Satellite(TLE)

        Sat.CalculPosVit(EPOCH_NEW, TLEPOCH)

        XN = Sat.Position.X
        YN = Sat.Position.Y
        ZN = Sat.Position.Z
        VXN = Sat.Vitesse.X
        VYN = Sat.Vitesse.Y
        VZN = Sat.Vitesse.Z

    End Sub

    Sub SGP4_TLE_EPOCH(ByVal EPOCH_NEW As Double, ByVal TLEPOCH As Double)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS CARTESIENS VIA SGP POUR DETERMINER LA TRACE (CLASSES PREVISAT)
        'POUR LA DETERMINATION DE LA LATITUDE/LONGITUDE A L'EPOCH TLE
        '----------------------------------------------------------------------------------------------------------------------------
        Dim TLE As New TLE(LIGNE1, LIGNE2)
        Dim Sat As New Satellite(TLE)

        Sat.CalculPosVit(EPOCH_NEW, TLEPOCH)

        XTLE = Sat.Position.X
        YTLE = Sat.Position.Y
        ZTLE = Sat.Position.Z
        VXTLE = Sat.Vitesse.X
        VYTLE = Sat.Vitesse.Y
        VZTLE = Sat.Vitesse.Z

    End Sub

    Sub ActualiseGraphics(ByVal ActuTrack As Boolean)
        '----------------------------------------------------------------------------------------------------------------------------
        'ACTUALISE LA MAP 2D STANDARD
        '----------------------------------------------------------------------------------------------------------------------------

        g.Clear(TracePicture.BackColor)
        DrawSunShadow()
        If ActuTrack = True Then
            TraceSC()
        ElseIf (SatTrace.Length / 5) < 360 * MapPeriodNbr.SelectedItem Then
            TraceSC()
        Else
            If ResizeMode = True Then
                ActualizeTrack() 'actualise la trace
                ResizeMode = False
            Else
                TraceNom()
            End If
        End If
        ExportToMap()
        DrawSun()
        TracePicture.Image = Trace

    End Sub

    Sub Prediction()

        'Check Epoch Format / Delta Day
        Dim DeltaEpoch, DracoPeriod, AnoPeriod, GST2 As Double

        DELTAD = (CDbl(EPOCH + 0.5)) - Truncate(CDbl(EPOCH))
        EPOCH0 = 2430000 + EPOCH - DELTAD

        'On converti la période draconitique en min
        If Me.DracoPeriod_Label.Text = "sec" Then
            DracoPeriod = DPER / 60
        ElseIf Me.DracoPeriod_Label.Text = "min" Then
            DracoPeriod = DPER
        End If

        'On converti la période Anomalistique en min
        If Me.AnoPeriod_Label.Text = "sec" Then
            AnoPeriod = APER / 60
        ElseIf Me.AnoPeriod_Label.Text = "min" Then
            AnoPeriod = APER
        End If

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS CARTESIENS ET KEPLERIENS VIA SGP4 (CLASSES PREVISAT)
        '----------------------------------------------------------------------------------------------------------------------------
        SGP4(EPOCH, TLEPOCH)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS SECONDAIRES
        '----------------------------------------------------------------------------------------------------------------------------
        ElementSecondaires()

        '----------------------------------------------------------------------------------------------------------------------------
        'REMPLISSAGE DES BOXES
        '----------------------------------------------------------------------------------------------------------------------------
        Boxes()

    End Sub

    Sub ExportToMap()

        '----------------------------------------------------------------------------------------------------------------------------
        'DESSIN DU SATELLITE SUR LA MAP
        '----------------------------------------------------------------------------------------------------------------------------

        'Convertir les données lat/long en pixel

        Dim lat0, lat1, lat2, lat3, longi0, longi1, longi2, longi3, latpix, longpix, visuel, SatPWS2, SatPHS2
        Dim SatP As Bitmap

        Dim MapWS2 = MapW / 2
        Dim MapHS2 = MapH / 2

        If CBool(SatVisual1) = True Then
            SatP = My.Resources.Sat
            SatPicture.Width = 13
            SatPicture.Height = 11
        ElseIf CBool(SatVisual2) = True Then
            SatP = My.Resources.sat2
            SatPicture.Width = 16
            SatPicture.Height = 16
        End If

        Dim SPHS2 = SatPicture.Height / 2 + 1
        Dim SPWS2 = SatPicture.Width / 2 + 1

        latpix = Round(MapHS2 - (LAT * (MapH - 1) / 180), 0)
        longpix = Round(MapWS2 + (LONGI * (MapW - 1) / 360), 0)

        'Gestion de la transparence
        SatP.MakeTransparent(SatP.GetPixel(0, 0))

        g.DrawImage(SatP, CInt(longpix - SPHS2), CInt(latpix - SPHS2))

    End Sub

    Sub DrawSun()

        Dim SunP = My.Resources.sun

        Dim MapWS2 = MapW / 2
        Dim MapHS2 = MapH / 2

        bsol = Round(MapHS2 - (LATS * MapH / 180), 0)
        lsol = Round(MapWS2 + (LONGIS * MapW / 360), 0)

        'Gestion de la transparence
        SunP.MakeTransparent(SunP.GetPixel(0, 0))
        g.DrawImage(SunP, CInt(lsol - 11), CInt(bsol - 11))


    End Sub

    Sub DrawSunShadow()

        Dim DEG2PXHZ = MapW / Maths.T360
        Dim DEG2PXVT = MapH * 2.0 / Maths.T360

        Dim hcarte = MapH
        Dim lcarte = MapW

        If SunShadowCB.Checked = True Then

            '----------------------------------------------------------------------------------------------------------------------------
            'ZONE D'OMBRE DU SOLEIL
            '----------------------------------------------------------------------------------------------------------------------------
            Dim Corps As New Corps
            Dim pinceau As New Pen(Color.Brown)
            Dim AlphaZone As SolidBrush = New SolidBrush(Color.FromArgb(CInt(2.55 * 50), 0, 0, 0))

            Corps.CalculZoneVisibilite()

            '--------------'
            ' Zone d'ombre '
            '--------------

            Dim jmin, xmin As Integer
            Dim zone(360) As Point

            xmin = MapW - 3

            For j = 0 To 360
                zone(j).X = CInt(SunZone(j).X * DEG2PXHZ)
                zone(j).Y = CInt(SunZone(j).Y * DEG2PXVT)

                If LATS < 0.0 Then
                    If zone(j).X <= xmin Then
                        xmin = zone(j).X
                        jmin = j
                    End If
                Else
                    If zone(j).X < xmin Then
                        xmin = zone(j).X
                        jmin = j
                    End If
                End If
            Next

            If Abs(LATS) > 0.002449 * Maths.DEG2RAD Then

                ReDim zone(365)

                If LATS < 0.0 Then

                    For j = 3 To 362
                        zone(j).X = CInt(SunZone((j + jmin - 2) Mod 360).X * DEG2PXHZ)
                        zone(j).Y = CInt(SunZone((j + jmin - 2) Mod 360).Y * DEG2PXVT)
                    Next
                    zone(0) = New Point(MapW - 1, 0)
                    zone(1) = New Point(MapW - 1, hcarte + 1)
                    zone(2) = New Point(MapW - 1, CInt(0.5 * (zone(3).Y + zone(362).Y)))

                    zone(363) = New Point(0, CInt(0.5 * (zone(3).Y + zone(362).Y)))
                    zone(364) = New Point(0, hcarte + 1)
                    zone(365) = New Point(0, 0)

                Else

                    For j = 2 To 361
                        zone(j).X = CInt(SunZone((j + jmin - 2) Mod 360).X * DEG2PXHZ)
                        zone(j).Y = CInt(SunZone((j + jmin - 2) Mod 360).Y * DEG2PXVT)
                    Next

                    zone(0) = New Point(0, 0)
                    zone(1) = New Point(0, CInt(0.5 * (zone(2).Y + zone(361).Y)))

                    zone(362) = New Point(MapW - 1, CInt(0.5 * (zone(2).Y + zone(361).Y)))
                    zone(363) = New Point(MapW - 1, 0)
                    zone(364) = New Point(MapW - 1, hcarte + 1)
                    zone(365) = New Point(0, hcarte + 1)

                End If

                g.FillPolygon(AlphaZone, zone)
            Else

                Dim zone1(3), zone2(3) As Point

                If lsol > lcarte \ 4 AndAlso lsol < (4 * lcarte) \ 3 Then
                    zone1(0) = New Point(0, 0)
                    zone1(1).X = CInt(Min(SunZone(90).X, SunZone(270).X) * DEG2PXHZ)
                    zone1(1).Y = 0
                    zone1(2).X = CInt(Min(SunZone(90).X, SunZone(270).X) * DEG2PXHZ)
                    zone1(2).Y = hcarte + 1
                    zone1(3) = New Point(0, hcarte + 1)

                    zone2(0) = New Point(MapW - 1, 0)
                    zone2(1).X = CInt(Max(SunZone(90).X, SunZone(270).X) * DEG2PXHZ)
                    zone2(1).Y = 0
                    zone2(2).X = CInt(Max(SunZone(90).X, SunZone(270).X) * DEG2PXHZ)
                    zone2(2).Y = hcarte + 1
                    zone2(3) = New Point(lcarte + 1, hcarte + 1)

                    g.FillPolygon(AlphaZone, zone1)
                    g.FillPolygon(AlphaZone, zone2)
                Else
                    zone1(0) = New Point(CInt(Min(SunZone(90).X, SunZone(270).X) * DEG2PXHZ), 0)
                    zone1(1) = New Point(CInt(Min(SunZone(90).X, SunZone(270).X) * DEG2PXHZ), hcarte + 1)

                    zone1(2) = New Point(CInt(Max(SunZone(90).X, SunZone(270).X) * DEG2PXHZ), hcarte + 1)
                    zone1(3) = New Point(CInt(Max(SunZone(90).X, SunZone(270).X) * DEG2PXHZ), 0)
                    g.FillPolygon(AlphaZone, zone1)
                End If
            End If

        End If

    End Sub

    'CREE LA TRACE COMPLETE DU SATELLITE AU CHARGEMENT

    Sub TraceSC()

        'Sauvegarde le mode dans lequel la trace est générée
        If TrackMode1 = True Then
            TrackModeTrace = True 'Mode Classic
        ElseIf TrackMode2 = True Then
            TrackModeTrace = False 'Mode Graphic
        End If

        If Me.MapShowTrack.Checked = False Then Exit Sub

        Me.Cursor = Cursors.WaitCursor

        Dim MapWS2 = MapW / 2
        Dim MapHS2 = MapH / 2

        Dim pinceau As New Pen(Color.Brown)

        Dim j, SemiLatus, Radius
        Dim a, P, lati, lati1, e2, l, N, h, Ls As Double

        Dim lat0, lat1, lat2, lat3, longi0, longi1, longi2, longi3 As Double
        Dim latpix1, latpix2, longpix1, longpix2
        Dim GSTNew, EPOCH_NEW, NEW_ALT

        Dim mapPeriod As Integer = MapPeriodNbr.SelectedItem
        ReDim SatTrace(360 * mapPeriod, 4)

        EPOCH_NEW = EPOCH

        'PINT DE DEPART DE LA TRACE
        latpix2 = CSng(Round(MapHS2 - (LAT * MapH / 180), 0))
        longpix2 = CSng(Round(MapWS2 + (LONGI * MapW / 360), 0))

        SatTrace(0, 0) = CDbl(longpix2)
        SatTrace(0, 1) = CDbl(latpix2)
        SatTrace(0, 2) = Round(LAT, 2)
        SatTrace(0, 3) = Round(LONGI, 2)
        SatTrace(0, 4) = Round(ALT * 1000, 2)

        If CheckNW = True Then
            'Initialise l'orbite (vide la précédente si existante)
            GoogleControl1.InitOrb()
            GoogleEarthControl1.InitOrb()
        End If

        'Initialisation
        Dim st = 1 / (MM * 360)

        ProgressBar.Value = 0
        ProgressBar.Visible = True

        Try

            'Export PLOT
            Dim fichier = PlotPath & Me.SCName.Text & ".plot"
            Dim BeginDateDay = MJDGGEDate(EPOCH_NEW)
            Dim BeginDateHour = MJDGGEHour(EPOCH_NEW)
            If My.Computer.FileSystem.FileExists(fichier) = True Then File.Delete(fichier)
            'Write in Plot file
            File.AppendAllText(fichier, X & "," & Y & "," & Z & "," & BeginDateDay & "," & BeginDateHour & "," & LONGI & "," & LAT & "," & ALT * 1000 & vbCrLf)

            For i = 1 To 360 * mapPeriod

                latpix1 = latpix2
                longpix1 = longpix2

                'On détermine la nouvelle époque
                EPOCH_NEW = EPOCH + CDbl(i) * st

                'On détermine les données cartésiennes de la nouvelle époque
                SGP4Trace(EPOCH_NEW, TLEPOCH)

                'GST
                GSTNew = GSTCalc(EPOCH_NEW)

                ' Latitude - Méthode PREVISAT
                Dim r, phi, sph, ct
                r = Sqrt(XN * XN + YN * YN)
                lat0 = Atan(ZN / r)
                phi = 7.0
                While Abs(lat0 - phi) > 0.0000001
                    phi = lat0
                    sph = Sin(phi)
                    ct = 1.0 / Sqrt(1.0 - Terre.E2 * sph * sph)
                    lat0 = Atan((ZN + EarthEquRad * ct * Terre.E2 * sph) / r)
                End While
                lat0 = Maths.RAD2DEG * (lat0)

                'Longitude
                Ls = Maths.RAD2DEG * ((Atan2(YN, XN) - (Maths.DEG2RAD * GSTNew)) Mod Maths.DEUX_PI)
                If Ls > 180 Then Ls -= 360
                If Ls < -180 Then Ls += 360

                'Write in Plot file
                BeginDateDay = MJDGGEDate(EPOCH_NEW)
                BeginDateHour = MJDGGEHour(EPOCH_NEW)
                NEW_ALT = (Norme(XN, YN, ZN) - EarthEquRad) * 1000
                File.AppendAllText(fichier, XN & "," & YN & "," & ZN & "," & BeginDateDay & "," & BeginDateHour & "," & Ls & "," & lat0 & "," & NEW_ALT & vbCrLf)

                SatTrace(i, 0) = CDbl(longpix1)
                SatTrace(i, 1) = CDbl(latpix1)
                SatTrace(i, 2) = Round(lat0, 2)
                SatTrace(i, 3) = Round(Ls, 2)
                SatTrace(i, 4) = Round(NEW_ALT, 2)

                'Convertir les données lat/long en pixel
                latpix2 = CSng(Round(MapHS2 - (lat0 * MapH / 180), 0))
                longpix2 = CSng(Round(MapWS2 + (Ls * MapW / 360), 0))

                'Gestion des butées image en fonction du sens de propagation
                'Direction W > E (INC < 90)
                If INC < 90 Then
                    'Gestion de la butée image droite
                    If longpix1 > (MapWS2 + MapWS2 / 2) AndAlso longpix2 < (MapWS2 - MapWS2 / 2) Then
                        g.DrawLine(pinceau, longpix1, latpix1, MapW, (latpix2 + latpix1) / 2)
                        longpix1 = 1
                        latpix1 = (latpix2 + latpix1) / 2
                    End If
                    'Direction E > W (INC > 90)
                ElseIf INC > 90 Then
                    'Gestion de la butée image gauche
                    If longpix1 < (MapWS2 - MapWS2 / 2) AndAlso longpix2 > (MapWS2 + MapWS2 / 2) Then
                        g.DrawLine(pinceau, longpix1, latpix1, 0, (latpix2 + latpix1) / 2)
                        longpix1 = MapW
                        latpix1 = (latpix2 + latpix1) / 2
                    End If
                End If

                If Abs(longpix1 - longpix2) < MapWS2 Then g.DrawLine(pinceau, longpix1, latpix1, longpix2, latpix2)
                ProgressBar.Value = i * 100 / (360 * mapPeriod)

            Next

            ProgressBar.Visible = False

            'Create Google Orbits
            If CheckNW = True Then
                ProgressBar.Visible = True

                ProgressBar.Value = 0
                GoogleControl1.CreateTrace()

                ProgressBar.Value = 0
                GoogleEarthControl1.CreateTrace()

                ProgressBar.Visible = False
            End If

            Cursor = Cursors.Default

        Catch Ex As Exception
            MessageBox.Show("An error as occured in Track Module." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            ProgressBar.Visible = False
            Cursor = Cursors.Default
        End Try

    End Sub

    'DESSINE LA TRACE D'ORIGINE DU SATELLITE

    Sub TraceNom()

        If Me.MapShowTrack.Checked = False Then Exit Sub

        Cursor = Cursors.WaitCursor

        Dim MaxLng = Max(MapW1, MapW2)
        Dim MinLng = Min(MapW1, MapW2)
        Dim MaxLat = Max(MapH1, MapH2)
        Dim MinLat = Min(MapH1, MapH2)

        Dim LNGLO2HI = MaxLng / MinLng
        Dim LATLO2HI = MaxLat / MinLat
        Dim LNGHI2LO = MinLng / MaxLng
        Dim LATHI2LO = MinLat / MaxLat

        Dim MapWS2 = MapW / 2
        Dim MapHS2 = MapH / 2
        Dim latpix2, longpix2, latpix1, longpix1
        Dim pinceau As New Pen(Color.Brown)

        Dim mapPeriod As Integer = MapPeriodNbr.SelectedItem

        longpix2 = CSng(SatTrace(0, 0))
        latpix2 = CSng(SatTrace(0, 1))

        'Cas d'une passage entre 2 modes
        If TrackModeTrace = True And CBool(TrackMode2) = True Then 'Si Trace générée en Classic et redim > Graphic
            longpix2 = CSng(longpix2 * LNGLO2HI)
            latpix2 = CSng(latpix2 * LATLO2HI)
        ElseIf TrackModeTrace = False And CBool(TrackMode1) = True Then 'Si Trace générée en Graphic et redim > classic
            longpix2 = CSng(longpix2 * LNGHI2LO)
            latpix2 = CSng(latpix2 * LATHI2LO)
        End If

        For i = 1 To 360 * mapPeriod

            latpix1 = latpix2
            longpix1 = longpix2

            longpix2 = CSng(SatTrace(i, 0))
            latpix2 = CSng(SatTrace(i, 1))

            If TrackModeTrace = True And CBool(TrackMode2) = True Then 'Si Trace générée en Classic et redim > Graphic
                longpix2 = CSng(longpix2 * LNGLO2HI)
                latpix2 = CSng(latpix2 * LNGLO2HI)
            ElseIf TrackModeTrace = False And CBool(TrackMode1) = True Then 'Si Trace générée en Graphic et redim > classic
                longpix2 = CSng(longpix2 * LATHI2LO)
                latpix2 = CSng(latpix2 * LATHI2LO)
            End If

            'Gestion des butées image en fonction du sens de propagation
            'Direction W > E (INC < 90)
            If INC < 90 Then
                'Gestion de la butée image droite
                If longpix1 > (MapWS2 + MapWS2 / 2) AndAlso longpix2 < (MapWS2 - MapWS2 / 2) Then
                    g.DrawLine(pinceau, longpix1, latpix1, MapW, (latpix2 + latpix1) / 2)
                    longpix1 = 1
                    latpix1 = (latpix2 + latpix1) / 2
                End If
                'Direction E > W (INC > 90)
            ElseIf INC > 90 Then
                'Gestion de la butée image gauche
                If longpix1 < (MapWS2 - MapWS2 / 2) AndAlso longpix2 > (MapWS2 + MapWS2 / 2) Then
                    g.DrawLine(pinceau, longpix1, latpix1, 0, (latpix2 + latpix1) / 2)
                    longpix1 = MapW
                    latpix1 = (latpix2 + latpix1) / 2
                End If
            End If

            If Abs(longpix1 - longpix2) < MapWS2 Then g.DrawLine(pinceau, longpix1, latpix1, longpix2, latpix2)

        Next

        'Create Google Orbits
        If CheckNW = True And MapPeriodNbr.Focused = True Then
            'Initialise l'orbite (vide la précédente si existante)
            GoogleControl1.InitOrb()
            GoogleEarthControl1.InitOrb()
            'Crée l'orbite
            GoogleControl1.CreateTrace()
            GoogleEarthControl1.CreateTrace()
        End If

        Cursor = Cursors.Default

    End Sub

    Sub FullTraceSC(ByVal Period As Long)

        Me.Cursor = Cursors.WaitCursor

        Dim MapWS2 = MapW / 2
        Dim MapHS2 = MapH / 2

        Dim DracoPeriod, AnoPeriod, min
        Dim pinceau As New Pen(Color.Brown)
        Dim RAANNew, AOPNew, XN, YN, ZN, VXN, VYN, VZN, GSTNew, MAN, TANew, EAN, EPOCH_NEW
        Dim lat0, lat1, lat2, lat3, longi0, longi1, longi2, longi3
        Dim latpix1, latpix2, longpix1, longpix2

        'On converti la période draconitique en min
        If Me.DracoPeriod_Label.Text = "sec" Then
            DracoPeriod = DPER / 60
        ElseIf Me.DracoPeriod_Label.Text = "min" Then
            DracoPeriod = DPER
        End If

        'On converti la période Anomalistique en min
        If Me.AnoPeriod_Label.Text = "sec" Then
            AnoPeriod = APER / 60
        ElseIf Me.AnoPeriod_Label.Text = "min" Then
            AnoPeriod = APER
        End If

        'On recherche la point de départ sur le graphe
        '> Convertir les données lat/long actuelles en pixel

        latpix2 = CSng(Round(MapHS2 - (LAT * MapH / 180), 0))
        longpix2 = CSng(Round(MapWS2 + (LONGI * MapW / 360), 0))

        'on démarre la boucle
        Dim dep = DracoPeriod - Truncate(DracoPeriod)
        Dim TotalRev, ETFP0, ETFP_NEW As Double
        Dim Step0 = 0.25 'en min
        Dim Step1 = Step0 / 1440 ' en jour

        'Epoque de départ
        Dim EpochD = EPOCH

        'Epoque d'arrivée
        Dim EpochA = ((DracoPeriod * Period) / 1440) + EpochD

        If Me.ETFP_Label.Text = "min" Then ETFP0 = ETFP
        If Me.ETFP_Label.Text = "sec" Then ETFP0 = ETFP / 60
        ETFP_NEW = ETFP0

        'keplerian Parameters
        RAANNew = RAAN
        AOPNew = AOP

        Me.ProgressBar.Value = 0
        Me.ProgressBar.Visible = True

        For EPOCH_NEW = EpochD + Step1 To EpochA Step Step1

            latpix1 = latpix2
            longpix1 = longpix2

            'keplerian Parameters
            RAANNew = (RAANNew + (Step1 * NP) Mod 360) Mod 360
            AOPNew = (AOPNew + (Step1 * AP) Mod 360) Mod 360

            'Anomalies
            TotalRev = Step1 * MM
            ETFP_NEW = (ETFP_NEW + ((TotalRev * AnoPeriod) Mod AnoPeriod)) Mod (AnoPeriod)
            If ETFP_NEW < 0 Then ETFP_NEW = AnoPeriod + ETFP_NEW

            'Mean Anomaly
            MAN = Round(Maths.RAD2DEG * ((ETFP_NEW * 60) * Sqrt(Mu / SMA ^ 3)), 3)
            If MAN < 0 Then MAN = 360 + MAN

            'Eccentric Anomaly
            Dim j
            EAN = PI
            For j = 0 To 4
                EAN = ((MAN * PI / 180) - (ECC * (EAN * Cos(EAN) - Sin(EAN)))) / (1 - ECC * Cos(EAN))
            Next

            'True Anomaly
            TANew = (Atan(Sqrt((1 + ECC) / (1 - ECC)) * Tan((EAN / 2))) * 2)
            If TANew > 0 Then
                TANew = TANew
            Else
                TANew = (2 * PI) - Abs(TANew)
            End If
            TANew = Maths.RAD2DEG * (TANew)

            'On converti en données cartésiennes
            Dim SemiLatus = Val(SMA * (1 - ECC ^ 2))
            Dim Radius = Val(SemiLatus / (1 + (ECC * Cos(Maths.DEG2RAD * (TANew)))))

            XN = Radius * (Cos(Maths.DEG2RAD * (AOPNew + TANew)) * Cos(Maths.DEG2RAD * (RAANNew)) - Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (AOPNew + TANew)) * Sin(Maths.DEG2RAD * (RAANNew)))
            YN = Radius * (Cos(Maths.DEG2RAD * (AOPNew + TANew)) * Sin(Maths.DEG2RAD * (RAANNew)) + Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (AOPNew + TANew)) * Cos(Maths.DEG2RAD * (RAANNew)))
            ZN = Radius * (Sin(Maths.DEG2RAD * (AOPNew + TANew)) * Sin(Maths.DEG2RAD * (INC)))

            VX = Sqrt(Mu / SemiLatus) * (Cos(Maths.DEG2RAD * (TANew)) + ECC) * (-Sin(Maths.DEG2RAD * (AOPNew)) * Cos(Maths.DEG2RAD * (RAANNew)) - Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (RAANNew)) * Cos(Maths.DEG2RAD * (AOPNew))) - Sqrt(Mu / SemiLatus) * (Sin(Maths.DEG2RAD * (TANew))) * (Cos(Maths.DEG2RAD * (AOPNew)) * Cos(Maths.DEG2RAD * (RAANNew)) - Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (RAANNew)) * Sin(Maths.DEG2RAD * (AOPNew)))
            VY = Sqrt(Mu / SemiLatus) * (Cos(Maths.DEG2RAD * (TANew)) + ECC) * (-Sin(Maths.DEG2RAD * (AOPNew)) * Sin(Maths.DEG2RAD * (RAANNew)) + Cos(Maths.DEG2RAD * (INC)) * Cos(Maths.DEG2RAD * (RAANNew)) * Cos(Maths.DEG2RAD * (AOPNew))) - Sqrt(Mu / SemiLatus) * (Sin(Maths.DEG2RAD * (TANew))) * (Cos(Maths.DEG2RAD * (AOPNew)) * Sin(Maths.DEG2RAD * (RAANNew)) + Cos(Maths.DEG2RAD * (INC)) * Cos(Maths.DEG2RAD * (RAANNew)) * Sin(Maths.DEG2RAD * (AOPNew)))
            VZ = Sqrt(Mu / SemiLatus) * ((Cos(Maths.DEG2RAD * (TANew)) + ECC) * Sin(Maths.DEG2RAD * (INC)) * Cos(Maths.DEG2RAD * (AOPNew)) - Sin(Maths.DEG2RAD * (TANew)) * Sin(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (AOPNew)))

            'GSTNew
            GSTNew = GSTCalc(Val(EPOCH_NEW))

            'Latitude/Longitude
            Dim a, P, lati, lati1, e2, l, N, h, Lon, Ls As Double

            'Latitude
            'Ellipsoide
            a = EarthEquRad
            e2 = EarthFlat * (2 - EarthFlat)
            'Données
            P = Sqrt(XN ^ 2 + YN ^ 2)
            lati = Atan(ZN / P)
            'Init
            lati1 = lati
            'Boucle
            Do While l > 0.0000001
                N = a * (1 - e2 * Sin(lati1)) ^ (-1 / 2)
                h = (P / Cos(lati1)) - N
                lat1 = Atan((ZN / P) * (1 - (e2 / (1 + (h / N)))) ^ (-1))
                l = Abs(lati1 - lati)
            Loop
            lat0 = Maths.RAD2DEG * (lati1)

            'Longitude
            Dim AbsX = Abs(XN)
            Dim AbsY = Abs(YN)
            'Longitude fonction des coordonnées cartésiennes XN et yN
            If XN > 0 AndAlso YN > 0 AndAlso XN > YN Then Lon = Atan(AbsY / AbsX)
            If XN > 0 AndAlso YN > 0 AndAlso XN < YN Then Lon = (PI / 2) - Atan(AbsX / AbsY)
            If XN < 0 AndAlso YN > 0 AndAlso Abs(XN) < YN Then Lon = (PI / 2) + Atan(AbsX / AbsY)
            If XN < 0 AndAlso YN > 0 AndAlso Abs(XN) > YN Then Lon = PI - Atan(AbsY / AbsX)
            If XN < 0 AndAlso YN < 0 AndAlso Abs(XN) > Abs(YN) Then Lon = PI + Atan(AbsY / AbsX)
            If XN < 0 AndAlso YN < 0 AndAlso Abs(XN) < Abs(YN) Then Lon = (PI * 3 / 2) - Atan(AbsX / AbsY)
            If XN > 0 AndAlso YN < 0 AndAlso Abs(XN) < Abs(YN) Then Lon = (PI * 3 / 2) + Atan(AbsX / AbsY)
            If XN > 0 AndAlso YN < 0 AndAlso Abs(XN) > Abs(YN) Then Lon = (PI * 2) - Atan(AbsY / AbsX)
            Lon = Maths.RAD2DEG * (Lon)

            If GSTNew > 180 Then GSTNew -= 360

            If GSTNew > 0 Then
                If Lon < 0 Then Ls = -(GSTNew - Lon)
                If Lon > 0 AndAlso GSTNew > Lon Then Ls = -(GSTNew - Lon)
                If Lon > 0 AndAlso GSTNew < Lon Then Ls = Lon - GSTNew
            End If
            If GSTNew < 0 Then
                If Lon > 0 Then Ls = Abs(GSTNew) + Lon
                If Lon < 0 AndAlso Abs(Lon) < Abs(GSTNew) Then Ls = Abs(GSTNew) - Abs(Lon)
                If Lon < 0 AndAlso Abs(Lon) > Abs(GSTNew) Then Ls = -(Abs(Lon) - Abs(GSTNew))
            End If

            If Ls > 180 Then Ls -= 360
            If Ls < -180 Then Ls += Ls

            'Convertir les données lat/long en pixel
            latpix2 = CSng(Round(MapHS2 - (lat0 * MapH / 180), 0))
            longpix2 = CSng(Round(MapWS2 + (Ls * MapW / 360), 0))

            'Gestion des butées image en fonction du sens de propagation
            'Direction W > E (INC < 90)
            If INC < 90 Then
                'Gestion de la butée image droite
                If longpix1 > MapWS2 AndAlso longpix2 < MapWS2 Then
                    g.DrawLine(pinceau, longpix1, latpix1, MapW, (latpix2 + latpix1) / 2)
                    longpix1 = 1
                    latpix1 = (latpix2 + latpix1) / 2
                End If
                'Direction E > W (INC > 90)
            ElseIf INC > 90 Then
                'Gestion de la butée image gauche
                If longpix1 < MapWS2 AndAlso longpix2 > MapWS2 Then
                    g.DrawLine(pinceau, longpix1, latpix1, 0, (latpix2 + latpix1) / 2)
                    longpix1 = MapW
                    latpix1 = (latpix2 + latpix1) / 2
                End If
            End If

            If Abs(longpix1 - longpix2) < MapWS2 Then g.DrawLine(pinceau, longpix1, latpix1, longpix2, latpix2)

            ProgressBar.Value = 100 - ((100 * (EpochA - EPOCH_NEW)) / (EpochA - EpochD))
        Next

        Me.ProgressBar.Visible = False
        Me.Cursor = Cursors.Default
    End Sub

End Class

