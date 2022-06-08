'
'    PreviSat, position of artificial satellites, prediction of their passes, Iridium flares
'    Copyright (C) 2005-2011  Astropedia web: http://astropedia.free.fr  -  mailto: astropedia@free.fr
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'_______________________________________________________________________________________________________
'
' Description
' > Gestion de l'affichage des boites de messages
'
' Auteur
' > Astropedia
'
' Date de creation
' > 11/04/2009
'
' Date de revision
' > 17/07/2009
'
' Revisions
'

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text

Public NotInheritable Class Messages

    '----------------------'
    ' Constantes publiques '
    '----------------------'
    Public Const INFO As Integer = 0
    Public Const ERREUR As Integer = -1
    Public Const WARNING As Integer = 1

    '----------------------'
    ' Constantes protegees '
    '----------------------'

    '--------------------'
    ' Constantes privees '
    '--------------------'

    '---------------------'
    ' Variables publiques '
    '---------------------'

    '---------------------'
    ' Variables protegees '
    '---------------------'

    '-------------------'
    ' Variables privees '
    '-------------------'

    '---------------'
    ' Constructeurs '
    '---------------'
    Private Sub New()
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Construction d'un message
    ''' </summary>
    ''' <param name="msgvar">Partie variable du message</param>
    ''' <param name="msgfix">Partie fixe du message</param>
    ''' <returns>Message construit</returns>
    ''' <remarks></remarks>
    Public Shared Function Construire(ByVal msgvar As String, ByVal msgfix As String) As String

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i As Integer
        Dim res As String

        '-----------------'
        ' Initialisations '
        '-----------------'
        res = msgfix

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        If msgfix.Contains("#") Then

            ' Construction du message avec partie variable
            Dim tab() As String
            tab = msgvar.Split("#"c)

            Dim nb As Integer
            nb = msgfix.Split("#"c).Length - 1

            If nb = tab.Length Then
                Dim j As Integer

                j = 0
                For i = 0 To tab.Length - 1
                    j = res.IndexOf("#", j, StringComparison.CurrentCulture)

                    res = res.Insert(j, tab(i))
                    j += tab(i).Length + 1
                Next
                res = res.Replace("#"c, "").Replace("&"c, Environment.NewLine).Replace("$"c, ControlChars.Tab)
            End If
        Else
            res = res.Replace("&"c, Environment.NewLine).Replace("$"c, ControlChars.Tab)
        End If

        '--------'
        ' Retour '
        '--------'
        Return (res)
    End Function

    ''' <summary>
    ''' Affichage des boites de message d'erreur ou d'information
    ''' </summary>
    ''' <param name="msgvar">Partie variable du message apparaissant dans le corps de la boite de message</param>
    ''' <param name="msgfix">Partie fixe du message apparaissant dans le corps de la boite de message</param>
    ''' <param name="ierr">Type de l'erreur</param>
    ''' <remarks></remarks>
    Public Shared Sub Afficher(ByVal msgvar As String, ByVal msgfix As String, ByVal ierr As Integer)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim titre As String
        Dim typ As MessageBoxIcon

        '-----------------'
        ' Initialisations '
        '-----------------'
        titre = ""

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        If ierr = INFO Then
            ' Message informatif
            typ = MessageBoxIcon.Information
            titre = My.Resources.titre1

        ElseIf ierr = ERREUR Then
            ' Cas des erreurs fatales
            typ = MessageBoxIcon.Error
            titre = My.Resources.titre2

        ElseIf ierr = WARNING Then
            ' Cas des erreurs graves (mais non fatales)
            typ = MessageBoxIcon.Exclamation
            titre = My.Resources.titre3

        Else
        End If

        ' Affichage du message
        MessageBox.Show(Construire(msgvar, msgfix), titre, MessageBoxButtons.OK, typ, MessageBoxDefaultButton.Button1, 0, False)

        If ierr = ERREUR Then

            ' Suppression du repertoire d'execution a la fermeture du programme
            Dim di As DirectoryInfo
            di = New DirectoryInfo(My.Application.Info.DirectoryPath & "\out")
            If di.Exists Then di.Delete(True)

            ' Fin du programme
            End
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'

End Class
