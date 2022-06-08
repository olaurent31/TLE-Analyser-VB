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
' > Utilitaires lies a la position de la Lune
'
' Auteur
' > Astropedia
'
' Date de creation
' > 02/10/2009
'
' Date de revision
' > 19/02/2011
'
' Revisions
' > 13/12/2009 : Optimisations dans le calcul de la position de la Lune
' > 30/01/2011 : Suppression du francais
' > 19/02/2011 : Modification de la gestion des dates
'

Option Explicit On
Option Strict On

Imports System.Math

Public Class Lune
    Inherits Corps

    '----------------------'
    ' Constantes publiques '
    '----------------------'
    Public Const RAYON As Double = 1738.0

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
    Private _elongation As Double
    Private _fractionIlluminee As Double
    Private _phase As String

    ' Pour le calcul de la position
    Private Shared tabLon() As Double = {6288774.0, 1274027.0, 658314.0, 213618.0, -185116.0, -114332.0, 58793.0, 57066.0, 53322.0, 45758.0, -40923.0, -34720.0, -30383.0, 15327.0, -12528.0, 10980.0, 10675.0, 10034.0, 8548.0, -7888.0, -6766.0, -5163.0, 4987.0, 4036.0, 3994.0, 3861.0, 3665.0, -2689.0, -2602.0, 2390.0, -2348.0, 2236.0, -2120.0, -2069.0, 2048.0, -1773.0, -1595.0, 1215.0, -1110.0, -892.0, -810.0, 759.0, -713.0, -700.0, 691.0, 596.0, 549.0, 537.0, 520.0, -487.0, -399.0, -381.0, 351.0, -340.0, 330.0, 327.0, -323.0, 299.0, 294.0, 0.0}
    Private Shared tabDist() As Double = {-20905355.0, -3699111.0, -2955968.0, -569925.0, 48888.0, -3149.0, 246158.0, -152138.0, -170733.0, -204586.0, -129620.0, 108743.0, 104755.0, 10321.0, 0.0, 79661.0, -34782.0, -23210.0, -21636.0, 24208.0, 30824.0, -8379.0, -16675.0, -12831.0, -10445.0, -11650.0, 14403.0, -7003.0, 0.0, 10056.0, 6322.0, -9884.0, 5751.0, 0.0, -4950.0, 4130.0, 0.0, -3958.0, 0.0, 3258.0, 2616.0, -1897.0, -2117.0, 2354.0, 0.0, 0.0, -1423.0, -1117.0, -1571.0, -1739.0, 0.0, -4421.0, 0.0, 0.0, 0.0, 0.0, 1165.0, 0.0, 0.0, 8752.0}
    Private Shared tabLat() As Double = {5128122.0, 280602.0, 277693.0, 173237.0, 55413.0, 46271.0, 32573.0, 17198.0, 9266.0, 8822.0, 8216.0, 4324.0, 4200.0, -3359.0, 2463.0, 2211.0, 2065.0, -1870.0, 1828.0, -1794.0, -1749.0, -1565.0, -1491.0, -1475.0, -1410.0, -1344.0, -1335.0, 1107.0, 1021.0, 833.0, 777.0, 671.0, 607.0, 596.0, 491.0, -451.0, 439.0, 422.0, 421.0, -366.0, -351.0, 331.0, 315.0, 302.0, -283.0, -229.0, 223.0, 223.0, -220.0, -220.0, -185.0, 181.0, -177.0, 176.0, 166.0, -164.0, 132.0, -119.0, 115.0, 107.0}

    Private Shared tabCoef1(,) As Integer = {{0, 0, 1, 0}, {2, 0, -1, 0}, {2, 0, 0, 0}, {0, 0, 2, 0}, {0, 1, 0, 0}, {0, 0, 0, 2}, {2, 0, -2, 0}, {2, -1, -1, 0}, {2, 0, 1, 0}, {2, -1, 0, 0}, {0, 1, -1, 0}, {1, 0, 0, 0}, {0, 1, 1, 0}, {2, 0, 0, -2}, {0, 0, 1, 2}, {0, 0, 1, -2}, {4, 0, -1, 0}, {0, 0, 3, 0}, {4, 0, -2, 0}, {2, 1, -1, 0}, {2, 1, 0, 0}, {1, 0, -1, 0}, {1, 1, 0, 0}, {2, -1, 1, 0}, {2, 0, 2, 0}, {4, 0, 0, 0}, {2, 0, -3, 0}, {0, 1, -2, 0}, {2, 0, -1, 2}, {2, -1, -2, 0}, {1, 0, 1, 0}, {2, -2, 0, 0}, {0, 1, 2, 0}, {0, 2, 0, 0}, {2, -2, -1, 0}, {2, 0, 1, -2}, {2, 0, 0, 2}, {4, -1, -1, 0}, {0, 0, 2, 2}, {3, 0, -1, 0}, {2, 1, 1, 0}, {4, -1, -2, 0}, {0, 2, -1, 0}, {2, 2, -1, 0}, {2, 1, -2, 0}, {2, -1, 0, -2}, {4, 0, 1, 0}, {0, 0, 4, 0}, {4, -1, 0, 0}, {1, 0, -2, 0}, {2, 1, 0, -2}, {0, 0, 2, -2}, {1, 1, 1, 0}, {3, 0, -2, 0}, {4, 0, -3, 0}, {2, -1, 2, 0}, {0, 2, 1, 0}, {1, 1, -1, 0}, {2, 0, 3, 0}, {2, 0, -1, -2}}
    Private Shared tabCoef2(,) As Integer = {{0, 0, 0, 1}, {0, 0, 1, 1}, {0, 0, 1, -1}, {2, 0, 0, -1}, {2, 0, -1, 1}, {2, 0, -1, -1}, {2, 0, 0, 1}, {0, 0, 2, 1}, {2, 0, 1, -1}, {0, 0, 2, -1}, {2, -1, 0, -1}, {2, 0, -2, -1}, {2, 0, 1, 1}, {2, 1, 0, -1}, {2, -1, -1, 1}, {2, -1, 0, 1}, {2, -1, -1, -1}, {0, 1, -1, -1}, {4, 0, -1, -1}, {0, 1, 0, 1}, {0, 0, 0, 3}, {0, 1, -1, 1}, {1, 0, 0, 1}, {0, 1, 1, 1}, {0, 1, 1, -1}, {0, 1, 0, -1}, {1, 0, 0, -1}, {0, 0, 3, 1}, {4, 0, 0, -1}, {4, 0, -1, 1}, {0, 0, 1, -3}, {4, 0, -2, 1}, {2, 0, 0, -3}, {2, 0, 2, -1}, {2, -1, 1, -1}, {2, 0, -2, 1}, {0, 0, 3, -1}, {2, 0, 2, 1}, {2, 0, -3, -1}, {2, 1, -1, 1}, {2, 1, 0, 1}, {4, 0, 0, 1}, {2, -1, 1, 1}, {2, -2, 0, -1}, {0, 0, 1, 3}, {2, 1, 1, -1}, {1, 1, 0, -1}, {1, 1, 0, 1}, {0, 1, -2, -1}, {2, 1, -1, -1}, {1, 0, 1, 1}, {2, -1, -2, -1}, {0, 1, 2, 1}, {4, 0, -2, -1}, {4, -1, -1, -1}, {1, 0, 1, -1}, {4, 0, 1, -1}, {1, 0, -1, -1}, {4, -1, 0, -1}, {2, -2, 0, 1}}

    '---------------'
    ' Constructeurs '
    '---------------'

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul de la position de la Lune avec le modele simplifie issu de
    ''' l'Astronomical Algorithms 2nd edition de Jean Meeus, pp337-342
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <remarks></remarks>
    Public Sub CalculPosition(ByVal dat As Dates)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i, j As Integer
        Dim a1, a2, a3, ang1, ang2, b0, bt, cb, ce, fact1, fact2, l0, ll, lv, obliquite, r0, rp, sb, se, t, t2, t3, t4, xx As Double
        Dim coef(4) As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        b0 = 0.0
        l0 = 0.0
        r0 = 0.0
        t = dat.JourJulienUTC * Dates.NB_SIECJ_PAR_JOURS
        t2 = t * t
        t3 = t2 * t
        t4 = t3 * t

        ' Longitude moyenne de la Lune
        ll = Maths.DEG2RAD * ((218.3164477 + 481267.88123421 * t - 0.0015786 * t2 + t3 / 538841.0 - t4 / 65194000.0) Mod Maths.T360)

        ' Elongation moyenne de la Lune
        coef(0) = Maths.DEG2RAD * ((297.8501921 + 445267.1114034 * t - 0.0018819 * t2 + t3 / 545868.0 - t4 / 113065000.0) Mod Maths.T360)

        ' Anomalie moyenne du Soleil
        coef(1) = Maths.DEG2RAD * ((357.5291092 + 35999.0502909 * t - 0.0001536 * t2 + t3 / 24490000.0) Mod Maths.T360)

        ' Anomalie moyenne de la Lune
        coef(2) = Maths.DEG2RAD * ((134.9633964 + 477198.8675055 * t + 0.0087414 * t2 + t3 / 69699.0 - t4 / 14712000.0) Mod Maths.T360)

        ' Argument de latitude de la Lune
        coef(3) = Maths.DEG2RAD * ((93.272095 + 483202.0175233 * t - 0.0036539 * t2 - t3 / 3526000.0 + t4 / 863310000.0) Mod Maths.T360)

        coef(4) = 1.0 - 0.002516 * t - 0.0000074 * t2

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        For i = 0 To 59

            ang1 = 0.0
            ang2 = 0.0
            fact1 = 1.0
            fact2 = 1.0
            For j = 0 To 3
                ang1 += coef(j) * tabCoef1(i, j)
                ang2 += coef(j) * tabCoef2(i, j)
            Next
            If tabCoef1(i, 1) <> 0 Then fact1 = Pow(coef(4), Abs(tabCoef1(i, 1)))
            If tabCoef2(i, 1) <> 0 Then fact2 = Pow(coef(4), Abs(tabCoef2(i, 1)))

            ' Terme en longitude
            l0 += tabLon(i) * fact1 * Sin(ang1)

            ' Terme en distance
            r0 += tabDist(i) * fact1 * Cos(ang1)

            ' Terme en latitude
            b0 += tabLat(i) * fact2 * Sin(ang2)

        Next

        ' Principaux termes planetaires
        a1 = Maths.DEG2RAD * ((119.75 + 131.849 * t) Mod Maths.T360)
        a2 = Maths.DEG2RAD * ((53.09 + 479264.29 * t) Mod Maths.T360)
        a3 = Maths.DEG2RAD * ((313.45 + 481266.484 * t) Mod Maths.T360)
        l0 += 3958.0 * Sin(a1) + 1962.0 * Sin(ll - coef(3)) + 318.0 * Sin(a2)
        b0 += -2235.0 * Sin(ll) + 382.0 * Sin(a3) + 175.0 * (Sin(a1 - coef(3)) + Sin(a1 + coef(3))) + 127.0 * Sin(ll - coef(2)) - 115.0 * Sin(ll + coef(2))

        ' Coordonnees ecliptiques en repere spherique
        lv = ll + Maths.DEG2RAD * l0 * 0.000001
        bt = Maths.DEG2RAD * b0 * 0.000001
        rp = 385000.56 + r0 * 0.001

        cb = Cos(bt)
        sb = Sin(bt)
        obliquite = Maths.DEG2RAD * (84381.448 - 46.815 * t - 0.00059 * t2 + 0.001813 * t3) * Maths.DEG_PAR_ARCSEC
        ce = Cos(obliquite)
        se = Sin(obliquite)
        xx = rp * cb * Sin(lv)

        _position = New Vecteur(rp * cb * Cos(lv), xx * ce - rp * se * sb, xx * se + rp * ce * sb)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul de la phase de la Lune
    ''' </summary>
    ''' <param name="sun">Soleil</param>
    ''' <remarks></remarks>
    Public Sub CalculPhase(ByVal sun As Soleil)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim sgn As Double
        Dim w As New Vecteur(0.0, 0.0, 1.0)

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Elongation (ou angle de phase)
        _elongation = sun.Position.Angle(_position.Oppose)

        sgn = sun.Position.ProduitVectoriel(_position).ProduitScalaire(w)

        ' Fraction illuminee
        _fractionIlluminee = 0.5 * (1.0 + Cos(_elongation))

        If _fractionIlluminee >= 0.0 AndAlso _fractionIlluminee < 0.03 Then _phase = My.Resources.phase_lune1
        If _fractionIlluminee >= 0.03 AndAlso _fractionIlluminee < 0.31 Then
            _phase = IIf(sgn > 0.0, My.Resources.phase_lune2, My.Resources.phase_lune3).ToString
        End If
        If _fractionIlluminee >= 0.31 AndAlso _fractionIlluminee < 0.69 Then
            _phase = IIf(sgn > 0.0, My.Resources.phase_lune4, My.Resources.phase_lune5).ToString
        End If
        If _fractionIlluminee >= 0.69 AndAlso _fractionIlluminee < 0.97 Then
            _phase = IIf(sgn > 0.0, My.Resources.phase_lune6, My.Resources.phase_lune7).ToString
        End If
        If _fractionIlluminee >= 0.97 AndAlso _fractionIlluminee <= 1.0 Then _phase = My.Resources.phase_lune8

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Elongation() As Double
        Get
            Return _elongation
        End Get
    End Property

    Public ReadOnly Property FractionIlluminee() As Double
        Get
            Return _fractionIlluminee
        End Get
    End Property

    Public ReadOnly Property Phase() As String
        Get
            Return _phase
        End Get
    End Property
End Class
