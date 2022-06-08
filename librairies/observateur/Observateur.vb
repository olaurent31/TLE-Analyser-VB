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
' > Utilitaires lies au lieu d'observation
'
' Auteur
' > Astropedia
'
' Date de creation
' > 11/04/2009
'
' Date de revision
' > 06/03/2010
'
' Revisions
' > 06/03/2010 : Modification de la gestion des coordonnees cartesiennes du lieu d'observation

Option Explicit On
Option Strict On

Imports System.Math

Public Class Observateur

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
    ' Coordonnees geographiques du lieu d'observation
    Private _longitude As Double
    Private _latitude As Double
    Private _altitude As Double
    Private _nomlieu As String

    Private _coslat As Double
    Private _sinlat As Double
    Private _rayon As Double
    Private _posZ As Double

    Private _aaer As Double
    Private _aray As Double

    Private _tempsSideralGreenwich As Double

    Private _position As New Vecteur
    Private _vitesse As New Vecteur

    Private _rotHz As New Matrice
    Private _rotEq As New Matrice

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur d'un lieu d'observation a partir de ses coordonnees
    ''' </summary>
    ''' <param name="lieu">Nom du lieu d'observation</param>
    ''' <param name="longitude">Longitude (en degres, negative a l'est)</param>
    ''' <param name="latitude">Latitude (en degres)</param>
    ''' <param name="altitude">Altitude (en metres)</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lieu As String, ByVal longitude As Double, ByVal latitude As Double, ByVal altitude As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim coster, sinter As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        _nomlieu = lieu
        _longitude = Maths.DEG2RAD * longitude
        _latitude = Maths.DEG2RAD * latitude
        _altitude = altitude * 0.001

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _coslat = Cos(_latitude)
        _sinlat = Sin(_latitude)

        coster = 1.0 / Sqrt(1.0 - Terre.E2 * _sinlat * _sinlat)
        sinter = (1.0 - Terre.APLA) * (1.0 - Terre.APLA) * coster

        _rayon = (Terre.RAYON * coster + _altitude) * _coslat
        _posZ = (Terre.RAYON * sinter + _altitude) * _sinlat

        ' Pour l'extinction atmospherique
        _aray = 0.1451 * Exp(-_altitude / 7.996)
        _aaer = 0.12 * Exp(-_altitude / 1.5)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir d'un lieu d'observation
    ''' </summary>
    ''' <param name="observateur">Lieu d'observation</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal observateur As Observateur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _nomlieu = observateur._nomlieu
        _longitude = observateur._longitude
        _latitude = observateur._latitude
        _altitude = observateur._altitude

        _coslat = observateur._coslat
        _sinlat = observateur._sinlat

        _rayon = observateur._rayon
        _posZ = observateur._posZ

        ' Pour l'extinction atmospherique
        _aray = observateur._aray
        _aaer = observateur._aaer

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul de la position et de la vitesse ECI de l'observateur a une date donnee
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <remarks></remarks>
    Public Sub CalculPosVit(ByVal dat As Dates)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim costsl, sintsl, tsl As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        _tempsSideralGreenwich = TempsSideralDeGreenwich(dat)
        ' Temps sideral local
        tsl = _tempsSideralGreenwich - _longitude
        costsl = Cos(tsl)
        sintsl = Sin(tsl)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Position de l'observateur
        _position = New Vecteur(_rayon * costsl, _rayon * sintsl, _posZ)

        ' Vitesse de l'observateur
        _vitesse = New Vecteur(-Terre.OMEGA * _position.Y, Terre.OMEGA * _position.X, 0.0)

        ' Matrice utile pour le calcul des coordonnees horizontales
        Dim v1 As New Vecteur(_sinlat * costsl, -sintsl, _coslat * costsl)
        Dim v2 As New Vecteur(_sinlat * sintsl, costsl, _coslat * sintsl)
        Dim v3 As New Vecteur(-_coslat, 0.0, _sinlat)
        _rotHz = New Matrice(v1, v2, v3)

        ' Matrice utile pour le calcul des coordonnees equatoriales
        _rotEq = New Matrice(_rotHz.Transpose)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul du temps sideral de Greenwich.
    ''' D'apres la formule donnee dans l'Astronomical Algorithms 2nd edition de Jean Meeus, p88
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function TempsSideralDeGreenwich(ByVal dat As Dates) As Double

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim tu, tu2 As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        tu = dat.JourJulienUTC * Dates.NB_SIECJ_PAR_JOURS
        tu2 = tu * tu

        '---------------------'
        ' Corps de la methode '
        '---------------------'

        '--------'
        ' Retour '
        '--------'
        Return Maths.DEG2RAD * ((280.46061837 + 360.98564736629 * dat.JourJulienUTC + 0.000387933 * tu2 - tu2 * tu / 38710000.0) Mod Maths.T360)
    End Function

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Aaer() As Double
        Get
            Return _aaer
        End Get
    End Property

    Public ReadOnly Property Altitude() As Double
        Get
            Return _altitude
        End Get
    End Property

    Public ReadOnly Property Aray() As Double
        Get
            Return _aray
        End Get
    End Property

    Public ReadOnly Property Latitude() As Double
        Get
            Return _latitude
        End Get
    End Property

    Public ReadOnly Property Longitude() As Double
        Get
            Return _longitude
        End Get
    End Property

    Public ReadOnly Property Nomlieu() As String
        Get
            Return _nomlieu
        End Get
    End Property

    Public ReadOnly Property Position() As Vecteur
        Get
            Return _position
        End Get
    End Property

    Public ReadOnly Property RotEq() As Matrice
        Get
            Return _rotEq
        End Get
    End Property

    Public ReadOnly Property RotHz() As Matrice
        Get
            Return _rotHz
        End Get
    End Property

    Public ReadOnly Property TempsSideralGreenwich() As Double
        Get
            Return _tempsSideralGreenwich
        End Get
    End Property

    Public ReadOnly Property Vitesse() As Vecteur
        Get
            Return _vitesse
        End Get
    End Property
End Class
