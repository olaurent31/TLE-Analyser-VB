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
' > Definition des elements osculateurs
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

Public Class ElementsOsculateurs

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
    Protected Friend _demiGrandAxe As Double
    Protected Friend _excentricite As Double
    Protected Friend _inclinaison As Double
    Protected Friend _ascDroiteNA As Double
    Protected Friend _argumentPerigee As Double
    Protected Friend _anomalieVraie As Double
    Protected Friend _anomalieMoyenne As Double
    Protected Friend _anomalieExcentrique As Double
    Protected Friend _apogee As Double
    Protected Friend _perigee As Double
    Protected Friend _periode As Double

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
    Public ReadOnly Property AnomalieExcentrique() As Double
        Get
            Return _anomalieExcentrique
        End Get
    End Property

    Public ReadOnly Property AnomalieMoyenne() As Double
        Get
            Return _anomalieMoyenne
        End Get
    End Property

    Public ReadOnly Property AnomalieVraie() As Double
        Get
            Return _anomalieVraie
        End Get
    End Property

    Public ReadOnly Property Apogee() As Double
        Get
            Return _apogee
        End Get
    End Property

    Public ReadOnly Property ArgumentPerigee() As Double
        Get
            Return _argumentPerigee
        End Get
    End Property

    Public ReadOnly Property AscDroiteNA() As Double
        Get
            Return _ascDroiteNA
        End Get
    End Property

    Public ReadOnly Property DemiGrandAxe() As Double
        Get
            Return _demiGrandAxe
        End Get
    End Property

    Public ReadOnly Property Excentricite() As Double
        Get
            Return _excentricite
        End Get
    End Property

    Public ReadOnly Property Inclinaison() As Double
        Get
            Return _inclinaison
        End Get
    End Property

    Public ReadOnly Property Perigee() As Double
        Get
            Return _perigee
        End Get
    End Property

    Public ReadOnly Property Periode() As Double
        Get
            Return _periode
        End Get
    End Property
End Class
