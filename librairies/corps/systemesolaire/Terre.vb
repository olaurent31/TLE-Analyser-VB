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
' > Donnees relatives a la Terre (WGS 72)
'
' Auteur
' > Astropedia
'
' Date de creation
' > 11/04/2009
'
' Date de revision
' > 13/11/2010
'
' Revisions
' > 13/11/2010 : Ajout des constantes MILE_PAR_KM et PIED_PAR_METRE
'

Option Explicit On
Option Strict On

Imports System.Math

Public NotInheritable Class Terre

    '----------------------'
    ' Constantes publiques '
    '----------------------'
    ' Rayon equatorial
    Public Const RAYON As Double = 6378.135

    ' Constante geocentrique de la gravitation (km^3 s^-2)
    Public Const GE As Double = 398600.8
    Public Shared ReadOnly KE As Double = Dates.NB_SEC_PAR_MIN * Sqrt(GE / (RAYON * RAYON * RAYON))

    ' Premieres harmoniques zonales
    Public Const J2 As Double = 0.0010826158
    Public Const J3 As Double = -0.00000253881
    Public Const J4 As Double = -0.00000165597

    ' Rapport du jour solaire moyen sur le jour sideral
    Public Const OMEGA0 As Double = 1.0027379093507951

    ' Vitesse de rotation de la Terre (rad s^-1)
    Public Const OMEGA As Double = Maths.DEUX_PI * OMEGA0 * Dates.NB_JOUR_PAR_SEC

    ' Aplatissement de la Terre
    Public Const APLA As Double = 1.0 / 298.26
    Public Const E2 As Double = APLA * (2.0 - APLA)

    ' Nombre de miles par kilometre
    Public Const MILE_PAR_KM As Double = 1.0 / 1.609344

    ' Nombre de pieds par metre
    Public Const PIED_PAR_METRE As Double = 1.0 / 0.3048

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

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'

End Class
