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
' > 07/12/2012

Public Class OrbitSummary

    Private Sub OrbSum_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrbSum_OK.Click
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.OrbSum_TextBox.SelectAll()
    End Sub

    Private Sub CopySummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopySummary.Click
        Me.OrbSum_TextBox.SelectAll()
        Me.OrbSum_TextBox.Copy()
        Me.OrbSum_TextBox.DeselectAll()
    End Sub
End Class