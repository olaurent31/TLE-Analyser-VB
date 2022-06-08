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
' > Utilitaires sur la manipulation des TLE
'
' Auteur
' > Astropedia
'
' Date de creation
' > 11/04/2009
'
' Date de revision
' > 19/02/2011
'
' Revisions
' > 19/02/2011 : Modification de la gestion des dates
'

Option Explicit On
Option Strict On

Imports System.Globalization
Imports System.Math
Imports System.IO

Public Class TLE

    '----------------------'
    ' Constantes publiques '
    '----------------------'

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
    Private _nb0 As Integer

    Private _argpo As Double
    Private _bstar As Double
    Private _ecco As Double
    Private _inclo As Double
    Private _mo As Double
    Private _no As Double
    Private _omegao As Double
    Private _epoque As Dates

    Private _nom As String
    Private _ligne1 As String
    Private _ligne2 As String
    Private _norad As String

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur a partir des 2 lignes du TLE
    ''' </summary>
    ''' <param name="li1">Ligne 1 du TLE</param>
    ''' <param name="li2">Ligne 2 du TLE</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal li1 As String, ByVal li2 As String)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim an, ibe As Integer
        Dim jrs As Double
        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat

        '-----------------'
        ' Initialisations '
        '-----------------'
        _ligne1 = li1
        _ligne2 = li2

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Numero NORAD
        _norad = _ligne1.Substring(2, 5)

        ' Epoque
        an = Integer.Parse(_ligne1.Substring(18, 2), nfi)
        jrs = Double.Parse(_ligne1.Substring(20, 12), nfi)
        an = CInt(IIf(an < 57, an + Dates.AN2000, an + 1900))
        Dim dat As New Dates(an, 1, 1.0, 0.0)
        _epoque = New Dates(dat.JourJulien + jrs - 1.0, 0.0, True)

        ' Coefficient pseudo-balistique
        ibe = Integer.Parse(_ligne1.Substring(59, 2), nfi)
        _bstar = 0.00001 * Double.Parse(_ligne1.Substring(53, 6), nfi) * Pow(10.0, ibe)

        ' Elements orbitaux moyens
        ' Inclinaison
        _inclo = Double.Parse(_ligne2.Substring(8, 8), nfi)

        ' Ascension droite du noeud ascendant
        _omegao = Double.Parse(_ligne2.Substring(17, 8), nfi)

        ' Excentricite
        _ecco = 0.0000001 * Double.Parse(_ligne2.Substring(26, 7), nfi)

        ' Argument du perigee
        _argpo = Double.Parse(_ligne2.Substring(34, 8), nfi)

        ' Anomalie moyenne
        _mo = Double.Parse(_ligne2.Substring(43, 8), nfi)

        ' Moyen mouvement
        _no = Double.Parse(_ligne2.Substring(52, 11), nfi)
        Home.N0 = _no
        ' Nombre d'orbites a l'epoque
        _nb0 = Integer.Parse(_ligne2.Substring(63, 5), nfi)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Verification d'un fichier TLE
    ''' </summary>
    ''' <param name="nomFichier">Nom du fichier TLE</param>
    ''' <param name="alarm">Affichage du message d'erreur</param>
    ''' <returns>Nombre de satellites contenus dans le fichier</returns>
    ''' <remarks></remarks>
    Public Shared Function VerifieFichier(ByVal nomFichier As String, ByVal alarm As Boolean) As Integer

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim ierr, itle, nb As Integer
        Dim li1, li2, ligne, nomsat As String

        '-----------------'
        ' Initialisations '
        '-----------------'
        ierr = 0
        itle = 0
        nb = 0
        nomsat = "---"
        li1 = ""
        li2 = ""

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Dim sr As New StreamReader(nomFichier)

        Try
            Do Until sr.Peek = -1
                ligne = sr.ReadLine
                If nomsat = "---" OrElse nomsat.Trim = "" Then nomsat = ligne.Trim

                If ligne.Length > 0 Then
                    If ligne.Chars(0) = "1"c Then
                        li1 = ligne

                        Do
                            li2 = sr.ReadLine
                        Loop While li2.Trim.Length = 0

                        VerifieLignes(li1, li2)
                        If (li1 = nomsat AndAlso itle = 3) OrElse (li1 <> nomsat AndAlso itle = 2) Then Throw New PreviSatException(8)

                        If li1 = nomsat Then itle = 2 Else itle = 3
                        nb += 1
                        nomsat = "---"
                    End If
                End If
            Loop
            sr.Close()

            If nb = 0 OrElse nomsat <> "---" Then Throw New PreviSatException(8)

        Catch ex As PreviSatException
            Dim msg As String
            msg = ""
            nb = 0

            ' Construction du message
            ierr = Integer.Parse(ex.Source)
            Select Case ierr
                Case Is <= 4
                    msg = nomsat & "#" & li2.Substring(1, 6).Trim
                Case 5
                    msg = nomsat & "#" & li1.Substring(2, 5) & "#" & li2.Substring(2, 5)
                Case 6, 7
                    msg = (ierr - 5).ToString(CultureInfo.CurrentCulture) & "#" & nomsat & "#" & li1.Substring(2, 5)
                    ierr = 6
                Case Else
                    ierr = 0
                    msg = nomFichier
            End Select

            If alarm Then Throw New PreviSatException(msg, My.Resources.ResourceManager.GetString("mess" & (ierr + 5).ToString("00")), Messages.WARNING)

        End Try

        '--------'
        ' Retour '
        '--------'
        Return (nb)
    End Function

    ''' <summary>
    ''' Lecture d'un fichier TLE
    ''' </summary>
    ''' <param name="nomFichier">Fichier TLE</param>
    ''' <param name="listeSatellites">Liste de satellites (numeros NORAD)</param>
    ''' <param name="tle">Ensemble des TLEs a initialiser</param>
    ''' <remarks></remarks>
    Public Shared Sub LectureFichier(ByVal nomFichier As String, ByVal listeSatellites() As String, ByRef tle() As TLE)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i, j, jmax, k As Integer
        Dim fichier, li1, li2, ligne, nomsat As String

        '-----------------'
        ' Initialisations '
        '-----------------'
        If listeSatellites Is Nothing Then
            jmax = tle.Length
        Else
            jmax = listeSatellites.Length
        End If

        Dim sr2 As New StreamReader(My.Application.Info.DirectoryPath & "\data\magnitude")
        fichier = sr2.ReadToEnd
        sr2.Close()

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Try

            Dim sr As New StreamReader(nomFichier)
            j = 0
            k = 0

            nomsat = "---"
            Do Until sr.Peek = -1 OrElse j > jmax

                ligne = sr.ReadLine
                If ligne.Chars(0) = "1"c Then
                    li1 = ligne
                    Do
                        li2 = sr.ReadLine
                    Loop While li2.Trim.Length = 0

                    If nomsat.Chars(0) = "1"c OrElse nomsat = "---" Then

                        Dim indx1, indx2 As Integer
                        indx1 = fichier.IndexOf(li1.Substring(2, 5))
                        If indx1 >= 0 Then
                            indx2 = fichier.IndexOf(Environment.NewLine, indx1) - indx1
                            nomsat = fichier.Substring(indx1 + 36, indx2 - 36).Trim
                        Else
                            nomsat = li1.Substring(2, 5)
                        End If
                    End If

                    If nomsat.Length > 25 AndAlso nomsat.Substring(25).Contains(".") Then nomsat = nomsat.Substring(0, 15).Trim
                    If nomsat.ToLower(CultureInfo.CurrentCulture) = "iss (zarya)" Then nomsat = "ISS"
                    If nomsat.ToLower(CultureInfo.CurrentCulture).Contains("iridium") AndAlso nomsat.Contains("[") Then nomsat = nomsat.Substring(0, nomsat.IndexOf("[", StringComparison.CurrentCulture)).Trim

                    If listeSatellites Is Nothing Then
                        tle(k) = New TLE(li1, li2)
                        tle(k)._nom = nomsat.Trim
                        k += 1
                    Else
                        For i = 0 To listeSatellites.Length - 1
                            If listeSatellites(i) = ligne.Substring(2, 5) Then
                                tle(i) = New TLE(li1, li2)
                                tle(i)._nom = nomsat.Trim
                                j += 1
                                Exit For
                            End If
                        Next
                    End If
                End If
                nomsat = ligne.Trim
            Loop
            sr.Close()
        Catch ex As FileNotFoundException

        End Try

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Mise a jour d'un fichier TLE
    ''' </summary>
    ''' <param name="fic_old">Nom du fichier a mettre a jour</param>
    ''' <param name="fic_new">Nom du fichier a lire</param>
    ''' <param name="compteRendu">Compte-rendu de la mise a jour</param>
    ''' <remarks></remarks>
    Public Shared Sub MiseAJourFichier(ByVal fic_old As String, ByVal fic_new As String, ByRef compteRendu As ArrayList)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim isat, nb_old, nb_new As Integer

        '-----------------'
        ' Initialisations '
        '-----------------'
        ' Verification du fichier contenant les anciens TLE
        nb_old = VerifieFichier(fic_old, False)
        If nb_old = 0 Then Throw New PreviSatException(fic_old, My.Resources.mess14, Messages.WARNING)

        ' Lecture du TLE
        Dim tle_old(nb_old - 1) As TLE
        LectureFichier(fic_old, Nothing, tle_old)

        ' Verification du fichier contenant les TLE recents
        nb_new = VerifieFichier(fic_new, False)
        If nb_new = 0 Then Throw New PreviSatException(fic_new, My.Resources.mess14, Messages.WARNING)

        ' Lecture du TLE
        Dim tle_new(nb_new - 1) As TLE
        LectureFichier(fic_new, Nothing, tle_new)

        If nb_new < nb_old Then Messages.Afficher(fic_new & "#" & fic_old, My.Resources.mess15, Messages.INFO)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Mise a jour
        Dim j As Integer = 0
        Dim nbMaj As Integer = 0
        For isat = 0 To nb_old - 1

            Dim nor1 As String = tle_old(isat)._norad
            Dim nor2 As String = ""

            nor2 = tle_new(j)._norad
            While String.Compare(nor2, nor1, StringComparison.CurrentCulture) < 0 AndAlso j < nb_new - 1
                j += 1
                nor2 = tle_new(j)._norad
            End While

            If nor1.Equals(nor2) Then
                If tle_old(isat)._epoque.JourJulien < tle_new(j)._epoque.JourJulien Then
                    Dim nomsat As String = IIf(tle_new(j)._nom.Equals(nor2), tle_old(isat)._nom, tle_new(j)._nom).ToString
                    tle_old(isat) = tle_new(j)
                    tle_old(isat)._nom = nomsat.Trim
                    nbMaj += 1
                Else
                    compteRendu.Add(tle_old(isat)._nom & Space(70 - tle_old(isat)._nom.Length) & tle_old(isat)._norad)
                End If
            Else
                compteRendu.Add(tle_old(isat)._nom & Space(70 - tle_old(isat)._nom.Length) & tle_old(isat)._norad)
            End If
        Next

        compteRendu.Add(nbMaj.ToString(CultureInfo.CurrentCulture))
        compteRendu.Add(nb_old.ToString(CultureInfo.CurrentCulture))

        ' Copie des TLE dans le fichier
        If nbMaj > 0 Then
            Try
                Dim sw As New StreamWriter(fic_old)

                For isat = 0 To nb_old - 1
                    sw.WriteLine(tle_old(isat)._nom)
                    sw.WriteLine(tle_old(isat)._ligne1)
                    sw.WriteLine(tle_old(isat)._ligne2)
                Next
                sw.Close()

            Catch ex As UnauthorizedAccessException
                Throw New PreviSatException(My.Resources.update01 & "#" & fic_old, My.Resources.mess21, Messages.WARNING)
            End Try
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Verification du checksum d'une ligne
    ''' </summary>
    ''' <param name="ligne">Ligne du TLE dont il faut verifier le checksum</param>
    ''' <returns>Resultat de la verification</returns>
    ''' <remarks></remarks>
    Private Shared Function CheckSum(ByVal ligne As String) As Boolean

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim check, i As Integer

        '-----------------'
        ' Initialisations '
        '-----------------'
        check = 0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        For i = 0 To 67
            check = CInt(IIf(ligne.Chars(i) = "-"c, check + 1, check + Val(ligne.Chars(i))))
        Next

        '--------'
        ' Retour '
        '--------'
        Return (check Mod 10 = Integer.Parse(ligne.Substring(68), CultureInfo.CurrentCulture))
    End Function

    ''' <summary>
    ''' Verification des lignes d'un TLE
    ''' </summary>
    ''' <param name="li1">Ligne 1 du TLE</param>
    ''' <param name="li2">Ligne 2 du TLE</param>
    ''' <remarks></remarks>
    Private Shared Sub VerifieLignes(ByVal li1 As String, ByVal li2 As String)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim exc, ierr As Integer

        '-----------------'
        ' Initialisations '
        '-----------------'
        exc = 0
        ierr = 0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Verification de la longueur des lignes
        If li1.Length <> 69 OrElse li2.Length <> 69 Then
            exc = 1
            ierr += 1
        End If

        ' Verification du numero des lignes
        If li1.Chars(0) <> "1"c OrElse li2.Chars(0) <> "2"c Then
            exc = 2
            ierr += 1
        End If

        If ierr = 1 Then
            Throw New PreviSatException(exc)
        ElseIf ierr > 1 Then
            Throw New PreviSatException(8)
        End If

        ' Verification des espaces dans les lignes
        If li1.Chars(1) <> " "c OrElse li1.Chars(8) <> " "c OrElse li1.Chars(17) <> " "c OrElse li1.Chars(32) <> " "c OrElse li1.Chars(43) <> " "c OrElse li1.Chars(52) <> " "c OrElse li1.Chars(61) <> " "c OrElse li1.Chars(63) <> " "c OrElse li2.Chars(1) <> " "c OrElse li2.Chars(7) <> " "c OrElse li2.Chars(16) <> " "c OrElse li2.Chars(25) <> " "c OrElse li2.Chars(33) <> " "c OrElse li2.Chars(42) <> " "c OrElse li2.Chars(51) <> " "c Then Throw New PreviSatException(3)

        ' Verification de la ponctuation des lignes
        If li1.Chars(23) <> "."c OrElse li1.Chars(34) <> "."c OrElse li2.Chars(11) <> "."c OrElse li2.Chars(20) <> "."c OrElse li2.Chars(37) <> "."c OrElse li2.Chars(46) <> "."c OrElse li2.Chars(54) <> "."c Then Throw New PreviSatException(4)

        ' Verification du numero NORAD
        If Not li1.Substring(2, 5).Equals(li2.Substring(2, 5)) Then Throw New PreviSatException(5)

        ' Verification des checksums
        If Not CheckSum(li1) Then Throw New PreviSatException(6)

        If Not CheckSum(li2) Then Throw New PreviSatException(7)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Argpo() As Double
        Get
            Return _argpo
        End Get
    End Property

    Public ReadOnly Property Bstar() As Double
        Get
            Return _bstar
        End Get
    End Property

    Public ReadOnly Property Ecco() As Double
        Get
            Return _ecco
        End Get
    End Property

    Public ReadOnly Property Epoque() As Dates
        Get
            Return _epoque
        End Get
    End Property

    Public ReadOnly Property Inclo() As Double
        Get
            Return _inclo
        End Get
    End Property

    Public ReadOnly Property Ligne1() As String
        Get
            Return _ligne1
        End Get
    End Property

    Public ReadOnly Property Ligne2() As String
        Get
            Return _ligne2
        End Get
    End Property

    Public ReadOnly Property Mo() As Double
        Get
            Return _mo
        End Get
    End Property

    Public ReadOnly Property Nb0() As Integer
        Get
            Return _nb0
        End Get
    End Property

    Public ReadOnly Property No() As Double
        Get
            Return _no
        End Get
    End Property

    Public ReadOnly Property Nom() As String
        Get
            Return _nom
        End Get
    End Property

    Public ReadOnly Property Norad() As String
        Get
            Return _norad
        End Get
    End Property

    Public ReadOnly Property Omegao() As Double
        Get
            Return _omegao
        End Get
    End Property
End Class
