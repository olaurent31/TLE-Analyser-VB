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
' > Utilitaires lies a la position du Soleil
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

Imports System.Math

Public Class Soleil
    Inherits Corps

    '----------------------'
    ' Constantes publiques '
    '----------------------'
    Public Const RAYON As Double = 696000.0
    Public Const UA As Double = 149597870.0
    Public Const MAGNITUDE As Double = -26.98

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
    Private _distanceUA As Double

    '---------------'
    ' Constructeurs '
    '---------------'

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul de la position du Soleil avec le modele simplifie issu de
    ''' l'Astronomical Algorithms 2nd edition de Jean Meeus, pp163-164
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <remarks></remarks>
    Public Sub CalculPosition(ByVal dat As Dates)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim e, ms, lp, ls, lv, obliquite, rp, tu, tu2, u, u1, v, xx As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        tu = dat.JourJulienUTC * Dates.NB_SIECJ_PAR_JOURS 'code d'origine
        'tu = (dat + 2430000) * Dates.NB_SIECJ_PAR_JOURS 'conversion en JJ
        tu2 = tu * tu

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Longitude moyenne
        ls = Maths.DEG2RAD * ((280.466457 + 36000.7698278 * tu + 0.00030322 * tu2) Mod Maths.T360)

        ' Longitude du perihelie
        lp = Maths.DEG2RAD * ((282.937348 + 1.7195366 * tu + 0.00045688 * tu2) Mod Maths.T360)

        ' Excentricite
        e = 0.01670843 - 0.000042037 * tu - 0.0000001267 * tu2

        ' Anomalie moyenne
        ms = ls - lp

        ' Resolution de l'equation de Kepler : u = ms + e * sin(u)
        u = ms
        Do
            u1 = u
            u = u1 + (ms + e * SIN(u1) - u1) / (1.0 - e * COS(u1))
        Loop While (Abs(u - u1) > 0.000000001)

        ' Anomalie vraie
        v = 2.0 * ATAN(Sqrt((1.0 + e) / (1.0 - e)) * TAN(0.5 * u))

        ' Longitude vraie
        lv = lp + v

        ' Rayon vecteur
        _distanceUA = 1.000001018 * (1.0 - e * COS(u))
        rp = _distanceUA * UA

        obliquite = Maths.DEG2RAD * (84381.448 - 46.815 * tu - 0.00059 * tu2 + 0.001813 * tu * tu2) * Maths.DEG_PAR_ARCSEC

        xx = rp * SIN(lv)
        _position = New Vecteur(rp * COS(lv), xx * COS(obliquite), xx * SIN(obliquite))

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property DistanceUA As Double
        Get
            Return _distanceUA
        End Get
    End Property

End Class
