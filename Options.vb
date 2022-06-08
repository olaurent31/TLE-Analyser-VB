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
' > Gestion du module OPTIONS
'
' Auteur
' > Olivier LAURENT
'
' Date de creation
' > 02/01/2013
'
' Date de revision
' > 03/01/2013

Public Class Options

    'CHARGEMENT

    Private Sub Load_App_1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Home.OptionsSaved = False
        ReadInit()

        'Gestion des fuseaux horaires
        TimeZoneDef()
    End Sub

    Private Sub Load_App_2(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.FormClosing

        If Home.OptionsSaved = False Then

            If MessageBox.Show("Quit Options without saving ?", "TLE Analyser", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) = DialogResult.No Then
                e.Cancel = True
            End If

        End If

    End Sub

    'SAUVEGARDE

    Private Sub ACTU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTU.Click

        Home.OptionsSaved = True

        'Sauvegarde
        If CheckTLE() = False Then
            SaveInit()
            TrackMode()
            Me.Close()
            Exit Sub
        End If

        'Sauvegarde + Actualisation
        SaveInit()
        TrackMode()
        Home.Prediction()
        Home.ActualiseModule(True)
        Me.Close()

        Home.OptionsSaved = False

    End Sub

    Private Sub DefOptButt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefOptButt.Click
        'GMAT
        OptionGmatShowTrack.Checked = True
        OptionGmatModel2.Checked = True
        OptionGmatPropPer.SelectedItem = "1"

        'Tracking Options
        TrackMode1.Checked = True
        SimulOnCB.Checked = False
        TrackSpeed.SelectedItem = "1"

        'Station Keeping
        LongWin.SelectedIndex = 0
        LatWin.SelectedIndex = 0
        LLREFTLE.Checked = True

        'Satellite Visual
        SatVisual2.Checked = True
        SatVisual4.Checked = True

    End Sub

    Private Sub CloseButt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseButt.Click
        Me.Close()
    End Sub

    'OPTIONS

    Private Sub DownloadTLE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DownloadTLE.Click

        Dim Link1 = "http://www.celestrak.com/NORAD/elements/"

        If Home.CheckNW = False Then
            MessageBox.Show("You must be connected to upload TLE lists", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Else

            Try

                Dim TLE

                FileOpen(1, "C:\TLEAnalyser\TLE\TLEList.txt", OpenMode.Input)

                ProgressBar1.Visible = True
                ProgressBar1.Value = 0

                While Not EOF(1)
                    TLE = LineInput(1) & ".txt"
                    My.Computer.Network.DownloadFile(Link1 & TLE, Home.TLEPath & TLE, "", "", False, 100000, True)

                    ProgressBar1.Value = ProgressBar1.Value + 2
                End While
                FileClose(1)

                ProgressBar1.Value = 0
                ProgressBar1.Visible = False

                MessageBox.Show("Update Successfull", "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                'Mise à jour du fichier INI
                Home.TleUpdateDate = Date.Now
                SaveInit()
                ReadInit()

            Catch Ex As Exception
                MessageBox.Show("Your Network seems to be restricted:" & vbCrLf & "TLE ANALYSER can't dowload updated TLE." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
            End Try

        End If

    End Sub

    Private Sub LLREFRM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LLREFRM.CheckedChanged
        MissionLatBox.Text = Round4(Home.LATTLE)
        MissionLongBox.Text = Round4(Home.LONGITLE)
    End Sub

    'WEB LINKS

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim MonURL As String = "http://nssdc.gsfc.nasa.gov/nmc/masterCatalog.do?sc=" & Home.CATNUMB
        Process.Start(MonURL)
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim MonURL As String = "http://gmat.gsfc.nasa.gov/"
        Process.Start(MonURL)
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim MonURL As String = "http://www.celestrak.com/NORAD/elements/"
        Process.Start(MonURL)
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Dim MonURL As String = "http://tleanalyser.blogspot.fr"
        Process.Start(MonURL)
    End Sub

    Private Sub RTST_Link_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles RTST_Link.LinkClicked
        Dim MonURL As String = "http://www.n2yo.com/?s=" & Home.Spacecraft.Text
        Process.Start(MonURL)
    End Sub

    'TIME ZONE

    Private Sub TimeZone_Timer(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click

        If TabControl1.SelectedIndex = "4" Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        TimeZoneDef()
    End Sub

    
End Class