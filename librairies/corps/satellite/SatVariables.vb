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
' > Variables du modele SGP4
'
' Auteur
' > Astropedia
'
' Date de creation
' > 30/01/2011
'
' Date de revision
' >
'
' Revisions
'

Option Explicit On
Option Strict On

Public Class SatVariables

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
    ' Elements orbitaux moyens
    Protected Friend argpo As Double
    Protected Friend bstar As Double
    Protected Friend ecco As Double
    Protected Friend inclo As Double
    Protected Friend mo As Double
    Protected Friend no As Double
    Protected Friend omegao As Double
    Protected Friend epoque As Dates

    ' Variables du modele SGP4
    Protected Friend init As Boolean
    Protected Friend isimp As Boolean
    Protected Friend irez As Integer
    Protected Friend ao As Double
    Protected Friend argpdot As Double
    Protected Friend argpm As Double
    Protected Friend argpp As Double
    Protected Friend atime As Double
    Protected Friend aycof As Double
    Protected Friend cc1 As Double
    Protected Friend cc4 As Double
    Protected Friend cc5 As Double
    Protected Friend cnodm As Double
    Protected Friend con41 As Double
    Protected Friend con42 As Double
    Protected Friend cosim As Double
    Protected Friend cosio As Double
    Protected Friend cosio2 As Double
    Protected Friend cosomm As Double
    Protected Friend d2 As Double
    Protected Friend d2201 As Double
    Protected Friend d2211 As Double
    Protected Friend d3 As Double
    Protected Friend d3210 As Double
    Protected Friend d3222 As Double
    Protected Friend d4 As Double
    Protected Friend d4410 As Double
    Protected Friend d4422 As Double
    Protected Friend d5220 As Double
    Protected Friend d5232 As Double
    Protected Friend d5421 As Double
    Protected Friend d5433 As Double
    Protected Friend day As Double
    Protected Friend dedt As Double
    Protected Friend del1 As Double
    Protected Friend del2 As Double
    Protected Friend del3 As Double
    Protected Friend delmo As Double
    Protected Friend delt As Double
    Protected Friend didt As Double
    Protected Friend dmdt As Double
    Protected Friend domdt As Double
    Protected Friend dndt As Double
    Protected Friend dnodt As Double
    Protected Friend eccsq As Double
    Protected Friend ee2 As Double
    Protected Friend e3 As Double
    Protected Friend em As Double
    Protected Friend emsq As Double
    Protected Friend ep As Double
    Protected Friend eta As Double
    Protected Friend f220 As Double
    Protected Friend f221 As Double
    Protected Friend f311 As Double
    Protected Friend f321 As Double
    Protected Friend f322 As Double
    Protected Friend f330 As Double
    Protected Friend f441 As Double
    Protected Friend f442 As Double
    Protected Friend f522 As Double
    Protected Friend f523 As Double
    Protected Friend f542 As Double
    Protected Friend f543 As Double
    Protected Friend g200 As Double
    Protected Friend g201 As Double
    Protected Friend g211 As Double
    Protected Friend g300 As Double
    Protected Friend g310 As Double
    Protected Friend g322 As Double
    Protected Friend g410 As Double
    Protected Friend g422 As Double
    Protected Friend g520 As Double
    Protected Friend g521 As Double
    Protected Friend g532 As Double
    Protected Friend g533 As Double
    Protected Friend gam As Double
    Protected Friend gsto As Double
    Protected Friend inclm As Double
    Protected Friend mdot As Double
    Protected Friend mm As Double
    Protected Friend mp As Double
    Protected Friend nm As Double
    Protected Friend nodecf As Double
    Protected Friend nodedot As Double
    Protected Friend nodem As Double
    Protected Friend omegap As Double
    Protected Friend omeosq As Double
    Protected Friend omgcof As Double
    Protected Friend posq As Double
    Protected Friend rp As Double
    Protected Friend rtemsq As Double
    Protected Friend rteosq As Double
    Protected Friend s1 As Double
    Protected Friend s2 As Double
    Protected Friend s3 As Double
    Protected Friend s4 As Double
    Protected Friend s5 As Double
    Protected Friend s6 As Double
    Protected Friend s7 As Double
    Protected Friend se2 As Double
    Protected Friend se3 As Double
    Protected Friend sgh2 As Double
    Protected Friend sgh3 As Double
    Protected Friend sgh4 As Double
    Protected Friend sh2 As Double
    Protected Friend sh3 As Double
    Protected Friend si2 As Double
    Protected Friend si3 As Double
    Protected Friend sinim As Double
    Protected Friend sinio As Double
    Protected Friend sinmao As Double
    Protected Friend sinomm As Double
    Protected Friend sl2 As Double
    Protected Friend sl3 As Double
    Protected Friend sl4 As Double
    Protected Friend snodm As Double
    Protected Friend ss1 As Double
    Protected Friend ss2 As Double
    Protected Friend ss3 As Double
    Protected Friend ss4 As Double
    Protected Friend ss5 As Double
    Protected Friend ss6 As Double
    Protected Friend ss7 As Double
    Protected Friend sz1 As Double
    Protected Friend sz2 As Double
    Protected Friend sz3 As Double
    Protected Friend sz11 As Double
    Protected Friend sz12 As Double
    Protected Friend sz13 As Double
    Protected Friend sz21 As Double
    Protected Friend sz22 As Double
    Protected Friend sz23 As Double
    Protected Friend sz31 As Double
    Protected Friend sz32 As Double
    Protected Friend sz33 As Double
    Protected Friend t As Double
    Protected Friend t2cof As Double
    Protected Friend t3cof As Double
    Protected Friend t4cof As Double
    Protected Friend t5cof As Double
    Protected Friend x1mth2 As Double
    Protected Friend x7thm1 As Double
    Protected Friend xfact As Double
    Protected Friend xi2 As Double
    Protected Friend xi3 As Double
    Protected Friend xincp As Double
    Protected Friend xgh2 As Double
    Protected Friend xgh3 As Double
    Protected Friend xgh4 As Double
    Protected Friend xh2 As Double
    Protected Friend xh3 As Double
    Protected Friend xl2 As Double
    Protected Friend xl3 As Double
    Protected Friend xl4 As Double
    Protected Friend xlamo As Double
    Protected Friend xlcof As Double
    Protected Friend xli As Double
    Protected Friend xmcof As Double
    Protected Friend xni As Double
    Protected Friend xpidot As Double
    Protected Friend z1 As Double
    Protected Friend z2 As Double
    Protected Friend z3 As Double
    Protected Friend z11 As Double
    Protected Friend z12 As Double
    Protected Friend z13 As Double
    Protected Friend z21 As Double
    Protected Friend z22 As Double
    Protected Friend z23 As Double
    Protected Friend z31 As Double
    Protected Friend z32 As Double
    Protected Friend z33 As Double
    Protected Friend zmol As Double
    Protected Friend zmos As Double

    '-------------------'
    ' Variables privees '
    '-------------------'

    '---------------'
    ' Constructeurs '
    '---------------'

    '----------'
    ' Methodes '
    '----------'

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'

End Class
