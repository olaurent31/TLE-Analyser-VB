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
' > Utilitaires mathematiques
'
' Auteur
' > Astropedia
'
' Date de creation
' > 11/04/2009
'
' Date de revision
' > 06/11/2009
'
' Revisions
' > 06/11/2009 : Correction de la dimension d'un tableau (reduit au necessaire et suffisant)

Option Explicit On
Option Strict On

Imports System.Math

Public NotInheritable Class Maths

    '----------------------'
    ' Constantes publiques '
    '----------------------'
    Public Const EPSDBL As Double = 0.000000000001
    Public Const EPSDBL100 As Double = 0.0000000001

    Public Const DEUX_TIERS As Double = 2.0 / 3.0
    Public Const DEUX_PI As Double = 2.0 * PI
    Public Const PI_SUR_DEUX As Double = 0.5 * PI
    Public Const T360 As Double = 360.0
    Public Const DEG2RAD As Double = PI / 180.0

    Public Const ARCSEC_PAR_DEG As Double = 3600.0
    Public Const ARCMIN_PAR_DEG As Double = 60.0

    Public Const DEUX_SUR_PI As Double = 1.0 / PI_SUR_DEUX
    Public Const RAD2DEG As Double = 1.0 / DEG2RAD
    Public Const DEG_PAR_ARCSEC As Double = 1.0 / ARCSEC_PAR_DEG
    Public Const DEG_PAR_ARCMIN As Double = 1.0 / ARCMIN_PAR_DEG
    Public Const HEUR2RAD As Double = 15.0 * DEG2RAD
    Public Const RAD2HEUR As Double = 1.0 / HEUR2RAD

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
    ''' Reduction des valeurs decimales en valeurs sexagesimales
    ''' </summary>
    ''' <param name="xdec">Valeur decimale</param>
    ''' <param name="deg">Degres</param>
    ''' <param name="min">Minutes</param>
    ''' <param name="sec">Secondes</param>
    ''' <remarks></remarks>
    Public Shared Sub Deg2DMS(ByVal xdec As Double, ByRef deg As Integer, ByRef min As Integer, ByRef sec As Integer)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim z As Integer
        Dim y As Double

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        y = Abs(xdec)
        z = Sign(xdec)
        deg = CInt(Floor(y))
        min = CInt(Floor(ARCMIN_PAR_DEG * (y - deg)))
        sec = CInt(Floor(ARCSEC_PAR_DEG * (y - deg) - ARCMIN_PAR_DEG * min + 0.5))
        If sec = ARCMIN_PAR_DEG Then sec = 0 : min += 1
        If min = ARCMIN_PAR_DEG Then min = 0 : deg += 1
        deg *= z

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul d'un extremum par interpolation a l'ordre 3,
    ''' issu de l'Astronomical Algorithms 2nd edition, de Jean Meeus, pp23-25
    ''' </summary>
    ''' <param name="xtab">Tableau des abscisses</param>
    ''' <param name="ytab">Tableau des ordonnees</param>
    ''' <param name="extremum">Extremum local</param>
    ''' <remarks></remarks>
    Public Shared Sub CalculExtremumInterpolation3(ByVal xtab() As Double, ByVal ytab() As Double, ByRef extremum() As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim a, b, ci As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        a = ytab(1) - ytab(0)
        b = ytab(2) - ytab(1)
        ci = (a + b) / (b - a)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        extremum(0) = xtab(1) - 0.5 * (xtab(1) - xtab(0)) * ci
        extremum(1) = ytab(1) - 0.125 * (a + b) * ci

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul d'une valeur x pour une valeur y donnee, par interpolation a l'ordre 3,
    ''' issu de l'Astronomical Algorithms 2nd edition, de Jean Meeus, pp23-27
    ''' </summary>
    ''' <param name="xtab">Tableau des abscisses</param>
    ''' <param name="ytab">Tableau des ordonnees</param>
    ''' <param name="epsilon">Precision pour laquelle la valeur x doit etre obtenue</param>
    ''' <param name="yval">Valeur y donnee</param>
    ''' <param name="xval">Valeur x calculee par interpolation</param>
    ''' <remarks></remarks>
    Public Shared Sub CalculValeurXInterpolation3(ByVal xtab() As Double, ByVal ytab() As Double, ByVal epsilon As Double, ByVal yval As Double, ByRef xval As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i, iter As Integer
        Dim a, b, c, dn0, dy, n0 As Double
        Dim yy(2) As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        iter = 0
        dn0 = 100000.0
        n0 = 0.0
        For i = 0 To 2
            yy(i) = ytab(i) - yval
        Next

        a = yy(1) - yy(0)
        b = yy(2) - yy(1)
        c = b - a

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        dy = 2.0 * yy(1)
        While Abs(dn0) >= epsilon AndAlso iter < 10000
            Dim tmp1, tmp2 As Double
            tmp1 = c * n0
            tmp2 = a + b + tmp1
            dn0 = -(dy + n0 * tmp2) / (tmp1 + tmp2)
            n0 += dn0
            iter += 1
        End While
        xval = xtab(1) + n0 * (xtab(1) - xtab(0))

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'

End Class
