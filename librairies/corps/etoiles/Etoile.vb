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
' > Utilitaires pour le calcul de la position des etoiles
'
' Auteur
' > Astropedia
'
' Date de creation
' > 19/02/2011
'
' Date de revision
' >
'
' Revisions
' >

Option Explicit On
Option Strict On

Imports System.Math
Imports System.Globalization
Imports System.IO

Public Class Etoile
    Inherits Corps

    '----------------------'
    ' Constantes publiques '
    '----------------------'

    '----------------------'
    ' Constantes protegees '
    '----------------------'

    '--------------------'
    ' Constantes privees '
    '--------------------'
    Private Const TABMAX As Integer = 9110

    '---------------------'
    ' Variables publiques '
    '---------------------'

    '---------------------'
    ' Variables protegees '
    '---------------------'

    '-------------------'
    ' Variables privees '
    '-------------------'
    Private _magnitude As Double
    Private _nom As String

    Private Shared _initStar As Boolean

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur par defaut
    ''' </summary>
    ''' <param name="nom">Nom de l'etoile</param>
    ''' <param name="ascensionDroite">Ascension droite (en heures)</param>
    ''' <param name="declinaison">Declinaison (en degres)</param>
    ''' <param name="magnitude">Magnitude visuelle</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal nom As String, ByVal ascensionDroite As Double, ByVal declinaison As Double, ByVal magnitude As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _nom = nom
        _ascensionDroite = ascensionDroite * Maths.HEUR2RAD
        _declinaison = declinaison * Maths.DEG2RAD
        _magnitude = magnitude

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul de la position du catalogue d'etoiles
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <param name="etoiles">Catalogue d'etoiles</param>
    ''' <remarks></remarks>
    Public Shared Sub CalculPositionEtoiles(ByVal observateur As Observateur, ByRef etoiles() As Etoile)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i As Integer

        '-----------------'
        ' Initialisations '
        '-----------------'
        If Not _initStar Then
            InitTabEtoiles(etoiles)
            _initStar = True
        End If

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        For i = 0 To TABMAX - 1
            etoiles(i).CalculCoordHoriz(observateur)
        Next

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul des coordonnees horizontales du corps
    ''' Le calcul de la refraction est issu de l'Astronomical Algorithms 2nd edition de Jean Meeus, p106.
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <remarks></remarks>
    Public Overloads Sub CalculCoordHoriz(ByVal observateur As Observateur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim cd, ht, htd, r As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        _visible = False
        _hauteur = -Math.PI
        cd = Cos(_declinaison)
        Dim vec1 As New Vecteur(Cos(_ascensionDroite) * cd, Sin(_ascensionDroite) * cd, Sin(_declinaison))
        Dim vec2 As New Vecteur(observateur.RotHz.MultVec(vec1))

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Hauteur
        ht = Asin(vec2.Z)

        If ht > -Maths.DEG2RAD Then

            ' Prise en compte de la refraction atmospherique
            htd = Maths.RAD2DEG * ht
            r = Maths.DEG2RAD * 1.02 / (Maths.ARCMIN_PAR_DEG * Tan(Maths.DEG2RAD * (htd + 10.3 / (htd + 5.11))))

            _hauteur = ht + r

            If _hauteur >= 0.0 Then
                ' Azimut
                _azimut = Atan2(vec2.Y, -vec2.X)
                If _azimut < 0.0 Then _azimut += Maths.DEUX_PI
                _visible = True
            Else
                _visible = False
                _hauteur = ht
            End If
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Lecture du catalogue d'etoiles
    ''' </summary>
    ''' <param name="etoiles">Catalogue d'etoiles</param>
    ''' <remarks></remarks>
    Private Shared Sub InitTabEtoiles(ByRef etoiles() As Etoile)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i, sgn, x1, x2 As Integer
        Dim ascDte, dec, mag, x3 As Double
        Dim ligne, nom As String
        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat

        '-----------------'
        ' Initialisations '
        '-----------------'
        i = 0
        ReDim etoiles(TABMAX - 1)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Dim sr As New StreamReader(My.Application.Info.DirectoryPath & "\data\etoiles.str")
        Do Until sr.Peek = -1
            ligne = sr.ReadLine

            nom = ""
            ascDte = 0.0
            dec = 0.0
            mag = 99.0
            If ligne.Length > 34 Then
                ' Ascension droite
                x1 = Integer.Parse(ligne.Substring(0, 2), nfi)
                x2 = Integer.Parse(ligne.Substring(2, 2), nfi)
                x3 = Double.Parse(ligne.Substring(4, 4), nfi)
                ascDte = x1 + x2 * Maths.DEG_PAR_ARCMIN + x3 * Maths.DEG_PAR_ARCSEC

                ' Declinaison
                sgn = CInt(IIf(ligne.Substring(9, 1) = "-", -1, 1))
                x1 = Integer.Parse(ligne.Substring(10, 2), nfi)
                x2 = Integer.Parse(ligne.Substring(12, 2), nfi)
                x3 = Double.Parse(ligne.Substring(14, 2), nfi)
                dec = sgn * (x1 + x2 * Maths.DEG_PAR_ARCMIN + x3 * Maths.DEG_PAR_ARCSEC)

                ' Magnitude visuelle
                mag = Double.Parse(ligne.Substring(31, 5), nfi)

                ' Mouvement propre
                'mvtAD = Double.Parse(ligne.Substring(17, 6), nfi) * Maths.DEG_PAR_ARCSEC
                'mvtDec = Double.Parse(ligne.Substring(24, 6), nfi) * Maths.DEG_PAR_ARCSEC

                ' Nom de l'etoile
                If ligne.Length > 37 Then
                    nom = ligne.Substring(37)
                End If

            End If
            etoiles(i) = New Etoile(nom, ascDte, dec, mag)

            i += 1
        Loop
        sr.Close()
        If i = 0 Then Throw New PreviSatException

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Magnitude As Double
        Get
            Return _magnitude
        End Get
    End Property

    Public ReadOnly Property Nom As String
        Get
            Return _nom
        End Get
    End Property
End Class
