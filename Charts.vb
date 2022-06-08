Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Math
Imports System.IO

Public Class ChartForm

    'CHART GENERATOR

    Sub SGP4Chart(ByVal EPOCH_NEW As Double, ByVal TLEPOCH As Double)

        '----------------------------------------------------------------------------------------------------------------------------
        'DETERMINATION DES ELEMENTS CARTESIENS VIA SGP POUR DETERMINER LA TRACE (CLASSES PREVISAT)
        '----------------------------------------------------------------------------------------------------------------------------
        Dim TLE As New TLE(Home.LIGNE1, Home.LIGNE2)
        Dim Sat As New Satellite(TLE)

        Sat.CalculPosVit(EPOCH_NEW, TLEPOCH)

        Home.XC = Sat.Position.X
        Home.YC = Sat.Position.Y
        Home.ZC = Sat.Position.Z
        Home.VXC = Sat.Vitesse.X
        Home.VYC = Sat.Vitesse.Y
        Home.VZC = Sat.Vitesse.Z

        CartesianToKeplerianChart(Home.XC, Home.YC, Home.ZC, Home.VXC, Home.VYC, Home.VZC)

    End Sub

    Private Sub ChartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChartButton.Click

        'Charts
        Dim ChartSMA As New Chart
        Dim ChartECC As New Chart
        Dim ChartINC As New Chart
        Dim ChartMALT As New Chart
        Dim ChartLNG As New Chart
        Dim ChartAPA As New Chart
        Dim ChartPEA As New Chart

        Cursor = Cursors.WaitCursor

        Try

            'Option Satellite Géostationnaire
            If Home.MM > 0.99 AndAlso Home.MM < 1.01 AndAlso Home.ECC < 0.01 AndAlso Home.INC < 1 Then
                CHART_LNG_CB.Checked = True
                CHART_LNG_CB.Enabled = True
            Else
                CHART_LNG_CB.Checked = False
                CHART_LNG_CB.Enabled = False
            End If

            'Sécurise le form
            If CHART_SMA_CB.Checked = False And CHART_ECC_CB.Checked = False And _
                CHART_INC_CB.Checked = False And CHART_MALT_CB.Checked = False And _
                CHART_APA_CB.Checked = False And CHART_PEA_CB.Checked = False And CHART_LNG_CB.Checked = False Then Exit Sub

            'Variables
            Dim EPOCH_NEW As Double = Home.EPOCH
            Dim GSTNew As Double
            Dim XPlage As Single = Chart_Days.Text
            Dim SMAMin, SMAMax, ECCMin, ECCMAx, INCMin, INCMax, MALTMin, MALTMax, LNGMin, LNGMAx, APAMin, APAMax, PEAMin, PEAMax As Double
            Dim SMAPlage, ECCPlage, INCPlage, MALTPlage, LNGPlage, APAPlage, PEAPlage
            Dim APRC, PERC, APAC, PEAC As Double
            Dim Ex, Ey As Double

            'Paramètre X
            Dim Int, Title
            If ChartXvalue.SelectedIndex = 0 Then 'Option "MIN"
                Int = 1 / 1440
                Title = "Minutes"
            ElseIf ChartXvalue.SelectedIndex = 1 Then 'Option "HRE"
                Int = 1 / 24
                Title = "Hours"
            ElseIf ChartXvalue.SelectedIndex = 2 Then 'Option "PERIODE"
                Int = 1 / 1440
                Title = "Periods"
                XPlage = CInt(CDbl(Chart_Days.Text) * Home.DPER)
            End If

            'Interdit le calcul à plus de 10 jours
            If (Int * XPlage) > 10 Then
                MessageBox.Show("Range must be lower than 10 days.", "TLE ANALYSER - Error")
                Cursor = Cursors.Default
                Exit Sub
            End If

            'Tableaux
            Dim SMAT(XPlage) As Double
            Dim ECCT(XPlage) As Double
            Dim INCT(XPlage) As Double
            Dim MALTT(XPlage) As Double
            Dim LNGT(XPlage) As Double
            Dim APAT(XPlage) As Double
            Dim PEAT(XPlage) As Double

            Dim Exy(XPlage, 1)

            ' Créer ChartArea (zone graphique)
            Dim ChartAreaSMA As New ChartArea()
            Dim ChartAreaECC As New ChartArea()
            Dim ChartAreaINC As New ChartArea()
            Dim ChartAreaMALT As New ChartArea()
            Dim ChartAreaLNG As New ChartArea()
            Dim ChartAreaAPA As New ChartArea()
            Dim ChartAreaPEA As New ChartArea()

            ' Ajouter le Chart Area à la Collection ChartAreas du Chart
            ChartSMA.ChartAreas.Add(ChartAreaSMA)
            ChartECC.ChartAreas.Add(ChartAreaECC)
            ChartINC.ChartAreas.Add(ChartAreaINC)
            ChartMALT.ChartAreas.Add(ChartAreaMALT)
            ChartLNG.ChartAreas.Add(ChartAreaLNG)
            ChartAPA.ChartAreas.Add(ChartAreaAPA)
            ChartPEA.ChartAreas.Add(ChartAreaPEA)

            ' Créer les data series (qui contiendront les DataPoint)
            Dim SMASerie As New Series()
            Dim ECCSerie As New Series()
            Dim INCSerie As New Series()
            Dim MALTSerie As New Series()
            Dim LNGSerie As New Series()
            Dim APASerie As New Series()
            Dim PEASerie As New Series()

            CHART_SMATAB.Controls.Clear()
            CHART_ECCTAB.Controls.Clear()
            CHART_INCTAB.Controls.Clear()
            CHART_MALTTAB.Controls.Clear()
            CHART_LNGTAB.Controls.Clear()
            CHART_APATAB.Controls.Clear()
            CHART_PEATAB.Controls.Clear()

            'Inialisation
            Dim r, phi, sph, ct, lat0, SatRad, Rt, Ls

            'Boucle
            For i = 0 To XPlage

                'Calcul des Elements
                SGP4Chart(EPOCH_NEW, Home.TLEPOCH)

                'GST
                GSTNew = GSTCalc(EPOCH_NEW)

                'Latitude - Méthode PREVISAT
                r = Sqrt(Home.XC * Home.XC + Home.YC * Home.YC)
                lat0 = Atan(Home.ZC / r)
                phi = 7.0
                While Abs(lat0 - phi) > 0.0000001
                    phi = lat0
                    sph = Sin(phi)
                    ct = 1.0 / Sqrt(1.0 - Terre.E2 * sph * sph)
                    lat0 = Atan((Home.ZC + EarthEquRad * ct * Terre.E2 * sph) / r)
                End While
                lat0 = Maths.RAD2DEG * (lat0)

                SatRad = Norme(Home.XC, Home.YC, Home.ZC)
                Rt = EarthEquRad / (Sqrt((Cos(Maths.DEG2RAD * (lat0)) ^ 2) + ((Sin(Maths.DEG2RAD * (lat0)) ^ 2) / (1 - EarthFlat) ^ 2)))
                Home.MALTC = Home.SMAC - Rt
                'Longitude
                Ls = Maths.RAD2DEG * ((Atan2(Home.YC, Home.XC) - (Maths.DEG2RAD * GSTNew)) Mod Maths.DEUX_PI)
                If Ls > 180 Then Ls -= 360
                If Ls < -180 Then Ls += 360

                'Exy(i, 0) = Home.ECCC * Cos(Maths.DEG2RAD * Home.AOPC)
                'Exy(i, 1) = Home.ECCC * Sin(Maths.DEG2RAD * Home.AOPC)
                'File.AppendAllText("c:\TLEAnalyser\testecc.txt", Exy(i, 0) & " " & Exy(i, 1) & vbCrLf)

                'Enregistrement des points dans les Series()
                If CHART_SMA_CB.Checked = True Then
                    SMASerie.Points.AddXY(i, Home.SMAC)
                    SMAT(i) = Home.SMAC
                End If

                If CHART_ECC_CB.Checked = True Then
                    ECCSerie.Points.AddXY(i, Home.ECCC)
                    ECCT(i) = Home.ECCC
                End If

                If CHART_INC_CB.Checked = True Then
                    INCSerie.Points.AddXY(i, Home.INCC)
                    INCT(i) = Home.INCC
                End If

                If CHART_MALT_CB.Checked = True Then
                    MALTSerie.Points.AddXY(i, Home.MALTC)
                    MALTT(i) = Home.MALTC
                End If

                If CHART_LNG_CB.Checked = True Then
                    LNGSerie.Points.AddXY(i, Ls)
                    LNGT(i) = Ls
                End If

                If CHART_APA_CB.Checked = True Then
                    APRC = Home.SMAC * (1 + Home.ECCC)
                    APAC = APRC - Rt
                    APASerie.Points.AddXY(i, APAC)
                    APAT(i) = APAC
                End If

                If CHART_PEA_CB.Checked = True Then
                    PERC = Home.SMAC * (1 - Home.ECCC)
                    PEAC = PERC - Rt
                    PEASerie.Points.AddXY(i, PEAC)
                    PEAT(i) = PEAC
                End If

                'Incrément de la date
                EPOCH_NEW = EPOCH_NEW + Int
            Next

            Array.Sort(SMAT)
            Array.Sort(ECCT)
            Array.Sort(INCT)
            Array.Sort(MALTT)
            Array.Sort(LNGT)
            Array.Sort(APAT)
            Array.Sort(PEAT)

            'On indique d'afficher ces Series sur les ChartArea concernées
            SMASerie.ChartArea = "ChartArea1"
            ECCSerie.ChartArea = "ChartArea1"
            INCSerie.ChartArea = "ChartArea1"
            MALTSerie.ChartArea = "ChartArea1"
            LNGSerie.ChartArea = "ChartArea1"
            APASerie.ChartArea = "ChartArea1"
            PEASerie.ChartArea = "ChartArea1"

            ' Ajouter les series à la collection Series du chart
            ChartSMA.Series.Add(SMASerie)
            ChartECC.Series.Add(ECCSerie)
            ChartINC.Series.Add(INCSerie)
            ChartMALT.Series.Add(MALTSerie)
            ChartLNG.Series.Add(LNGSerie)
            ChartAPA.Series.Add(APASerie)
            ChartPEA.Series.Add(PEASerie)

            '--------------------------------------------------
            'SMA
            '--------------------------------------------------
            If CHART_SMA_CB.Checked = True Then
                SMAMin = Min(SMAT(0), SMAT(XPlage))
                If SMAMin < EarthEquRad Then SMAMin = EarthEquRad
                SMAMax = Max(SMAT(0), SMAT(XPlage))
                If SMAMin = SMAMax Then
                    SMAPlage = Home.SMAC / 10
                Else
                    SMAPlage = SMAMax - SMAMin
                End If
                With ChartSMA
                    .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                    .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                    .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                    .ChartAreas(0).AxisX.Title = Title
                    .ChartAreas(0).AxisX.Minimum = 0
                    .ChartAreas(0).AxisX.RoundAxisValues()
                    .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                    .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                    If SMAMin - SMAPlage / 4 < EarthEquRad Then
                        .ChartAreas(0).AxisY.Minimum = EarthEquRad
                    Else
                        .ChartAreas(0).AxisY.Minimum = SMAMin - SMAPlage / 4
                    End If

                    .ChartAreas(0).AxisY.Maximum = SMAMax + SMAPlage / 4
                    .ChartAreas(0).AxisY.Interval = SMAPlage / 10
                    .Series(0).ChartType = SeriesChartType.Spline
                    .Series(0).BorderWidth = 2

                    .Location = New System.Drawing.Point(0, 0)
                    .Size = New System.Drawing.Size(625, 424)
                    CHART_SMATAB.Controls.Add(ChartSMA)

                End With

            End If
            '--------------------------------------------------
            'ECC
            '--------------------------------------------------
            If CHART_ECC_CB.Checked = True Then
                ECCMin = Min(ECCT(0), ECCT(XPlage))
                ECCMAx = Max(ECCT(0), ECCT(XPlage))
                If ECCMin = ECCMAx Then
                    ECCPlage = Home.ECCC / 10
                Else
                    ECCPlage = ECCMAx - ECCMin
                End If
                With ChartECC
                    .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                    .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                    .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                    .ChartAreas(0).AxisX.Title = Title
                    .ChartAreas(0).AxisX.Minimum = 0
                    .ChartAreas(0).AxisX.RoundAxisValues()
                    .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                    .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                    If ECCMin - ECCPlage / 4 < 0 Then
                        .ChartAreas(0).AxisY.Minimum = 0
                    Else
                        .ChartAreas(0).AxisY.Minimum = ECCMin - ECCPlage / 4
                    End If

                    .ChartAreas(0).AxisY.Maximum = ECCMAx + ECCPlage / 4
                    .ChartAreas(0).AxisY.Interval = ECCPlage / 5
                    .ChartAreas(0).AxisY.RoundAxisValues()
                    .Series("Series1").ChartType = SeriesChartType.Spline
                    .Series(0).BorderWidth = 2

                    .Location = New System.Drawing.Point(0, 0)
                    .Size = New System.Drawing.Size(625, 424)
                    CHART_ECCTAB.Controls.Add(ChartECC)
                End With
            End If
            '--------------------------------------------------
            'INC
            '--------------------------------------------------
            If CHART_INC_CB.Checked = True Then
                INCMin = Min(INCT(0), INCT(XPlage))
                INCMax = Max(INCT(0), INCT(XPlage))
                If INCMin = INCMax Then
                    INCPlage = Home.INCC / 10
                Else
                    INCPlage = INCMax - INCMin
                End If
                With ChartINC
                    .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                    .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                    .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                    .ChartAreas(0).AxisX.Title = Title
                    .ChartAreas(0).AxisX.Minimum = 0
                    .ChartAreas(0).AxisX.RoundAxisValues()
                    .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                    .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                    If INCMin - INCPlage / 4 < 0 Then
                        .ChartAreas(0).AxisY.Minimum = 0
                    Else
                        .ChartAreas(0).AxisY.Minimum = INCMin - INCPlage / 4
                    End If

                    .ChartAreas(0).AxisY.Maximum = INCMax + INCPlage / 4
                    .ChartAreas(0).AxisY.Interval = INCPlage / 5
                    .ChartAreas(0).AxisY.RoundAxisValues()
                    .Series("Series1").ChartType = SeriesChartType.Spline
                    .Series(0).BorderWidth = 2

                    .Location = New System.Drawing.Point(0, 0)
                    .Size = New System.Drawing.Size(625, 424)
                    CHART_INCTAB.Controls.Add(ChartINC)
                End With
            End If
            '--------------------------------------------------
            'MALT
            '--------------------------------------------------
            If CHART_MALT_CB.Checked = True Then
                MALTMin = Min(MALTT(0), MALTT(XPlage))
                If MALTMin < 0 Then MALTMin = 0
                MALTMax = Max(MALTT(0), MALTT(XPlage))
                If MALTMin = MALTMax Then
                    MALTPlage = Home.MALTC / 10
                Else
                    MALTPlage = MALTMax - MALTMin
                End If
                With ChartMALT
                    .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                    .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                    .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                    .ChartAreas(0).AxisX.Title = Title
                    .ChartAreas(0).AxisX.Minimum = 0
                    .ChartAreas(0).AxisX.RoundAxisValues()
                    .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                    .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                    If MALTMin - MALTPlage / 4 < 0 Then
                        .ChartAreas(0).AxisY.Minimum = 0
                    Else
                        .ChartAreas(0).AxisY.Minimum = MALTMin - MALTPlage / 4
                    End If

                    .ChartAreas(0).AxisY.Maximum = MALTMax + MALTPlage / 4
                    .ChartAreas(0).AxisY.Interval = MALTPlage / 10
                    .ChartAreas(0).AxisY.RoundAxisValues()
                    .Series("Series1").ChartType = SeriesChartType.Spline
                    .Series(0).BorderWidth = 2

                    .Location = New System.Drawing.Point(0, 0)
                    .Size = New System.Drawing.Size(625, 424)
                    CHART_MALTTAB.Controls.Add(ChartMALT)
                End With
            End If
            '--------------------------------------------------
            'LNG
            '--------------------------------------------------
            If Home.MM > 0.99 AndAlso Home.MM < 1.01 AndAlso Home.ECC < 0.01 AndAlso Home.INC < 1 Then
                If CHART_LNG_CB.Checked = True Then
                    LNGMin = Min(LNGT(0), LNGT(XPlage))
                    LNGMAx = Max(LNGT(0), LNGT(XPlage))
                    If LNGMin = LNGMAx Then
                        LNGPlage = Home.LNGC / 10
                    Else
                        LNGPlage = LNGMAx - LNGMin
                    End If
                    With ChartLNG
                        .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                        .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                        .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                        .ChartAreas(0).AxisX.Title = Title
                        .ChartAreas(0).AxisX.Minimum = 0
                        .ChartAreas(0).AxisX.RoundAxisValues()
                        .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                        .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                        If LNGMin - LNGPlage / 4 < -180 Then
                            .ChartAreas(0).AxisY.Minimum = -180
                        Else
                            .ChartAreas(0).AxisY.Minimum = LNGMin - LNGPlage / 4
                        End If
                        If LNGMAx + LNGPlage / 4 > 180 Then
                            .ChartAreas(0).AxisY.Maximum = 180
                        Else
                            .ChartAreas(0).AxisY.Maximum = LNGMAx + LNGPlage / 4
                        End If

                        .ChartAreas(0).AxisY.Interval = LNGPlage / 5
                        .ChartAreas(0).AxisY.RoundAxisValues()
                        .Series("Series1").ChartType = SeriesChartType.Spline
                        .Series(0).BorderWidth = 2

                        .Location = New System.Drawing.Point(0, 0)
                        .Size = New System.Drawing.Size(625, 424)
                        CHART_LNGTAB.Controls.Add(ChartLNG)
                    End With
                End If
            End If
            '--------------------------------------------------
            'APA
            '--------------------------------------------------
            If CHART_APA_CB.Checked = True Then
                APAMin = Min(APAT(0), APAT(XPlage))
                If APAMin < 0 Then APAMin = 0
                APAMax = Max(APAT(0), APAT(XPlage))
                If APAMin = APAMax Then
                    APAPlage = APAC / 10
                Else
                    APAPlage = APAMax - APAMin
                End If
                With ChartAPA
                    .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                    .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                    .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                    .ChartAreas(0).AxisX.Title = Title
                    .ChartAreas(0).AxisX.Minimum = 0
                    .ChartAreas(0).AxisX.RoundAxisValues()
                    .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                    .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                    If APAMin - APAPlage / 4 < 0 Then
                        .ChartAreas(0).AxisY.Minimum = 0
                        'Dim LimitLine As New StripLine()
                        'LimitLine.BackColor = Color.FromArgb(120, Color.Red)
                        'LimitLine.IntervalOffset = 200
                        'LimitLine.Interval = 0
                        'LimitLine.StripWidth = 2
                        'LimitLine.Text = "200 kms Limit"
                        '.ChartAreas(0).AxisY.StripLines.Add(LimitLine)
                    Else
                        .ChartAreas(0).AxisY.Minimum = APAMin - APAPlage / 4
                    End If

                    .ChartAreas(0).AxisY.Maximum = APAMax + APAPlage / 4
                    .ChartAreas(0).AxisY.Interval = APAPlage / 10
                    .ChartAreas(0).AxisY.RoundAxisValues()
                    .Series("Series1").ChartType = SeriesChartType.Spline
                    .Series(0).BorderWidth = 2

                    .Location = New System.Drawing.Point(0, 0)
                    .Size = New System.Drawing.Size(625, 424)
                    CHART_APATAB.Controls.Add(ChartAPA)
                End With
            End If
            '--------------------------------------------------
            'PEA
            '--------------------------------------------------
            If CHART_PEA_CB.Checked = True Then
                PEAMin = Min(PEAT(0), PEAT(XPlage))
                If PEAMin < 0 Then PEAMin = 0
                PEAMax = Max(PEAT(0), PEAT(XPlage))
                If PEAMin = PEAMax Then
                    PEAPlage = PEAC / 10
                Else
                    PEAPlage = PEAMax - PEAMin
                End If
                With ChartPEA
                    .ChartAreas("ChartArea1").BackColor = Color.DarkGray
                    .ChartAreas("ChartArea1").BackSecondaryColor = Color.White
                    .ChartAreas("ChartArea1").BackGradientStyle = GradientStyle.TopBottom

                    .ChartAreas(0).AxisX.Title = Title
                    .ChartAreas(0).AxisX.Minimum = 0
                    .ChartAreas(0).AxisX.RoundAxisValues()
                    .ChartAreas(0).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
                    .ChartAreas(0).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot

                    If PEAMin - PEAPlage / 4 < 0 Then
                        .ChartAreas(0).AxisY.Minimum = 0
                    Else
                        .ChartAreas(0).AxisY.Minimum = PEAMin - PEAPlage / 4
                    End If

                    .ChartAreas(0).AxisY.Maximum = PEAMax + PEAPlage / 4
                    .ChartAreas(0).AxisY.Interval = PEAPlage / 10
                    .ChartAreas(0).AxisY.RoundAxisValues()
                    .Series("Series1").ChartType = SeriesChartType.Spline
                    .Series(0).BorderWidth = 2

                    .Location = New System.Drawing.Point(0, 0)
                    .Size = New System.Drawing.Size(625, 424)
                    CHART_PEATAB.Controls.Add(ChartPEA)
                End With
            End If

            Cursor = Cursors.Default

        Catch Ex As Exception
            MessageBox.Show("An error as occured during Chart generation." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub Chart_Days_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Chart_Days.KeyPress
        If IsNumeric(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = ControlChars.Back Then
            e.KeyChar = ControlChars.Back
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

End Class