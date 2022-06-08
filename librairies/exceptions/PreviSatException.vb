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
' > Gestion des exceptions
'
' Auteur
' > Astropedia
'
' Date de creation
' > 21/05/2009
'
' Date de revision
' > 19/06/2009
'
' Revisions
'

Option Explicit On
Option Strict On

Imports System.Globalization
Imports PreviSat.Messages

Public Class PreviSatException
    Inherits Exception

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

    '---------------'
    ' Constructeurs '
    '---------------'
    Public Sub New()
    End Sub

    Public Sub New(ByRef ierr As Integer)
        Me.Source = ierr.ToString(CultureInfo.CurrentCulture)
    End Sub

    Public Sub New(ByVal msgvar As String, ByVal msgfix As String, ByVal ierr As Integer)
        Messages.Afficher(msgvar, msgfix, ierr)
    End Sub

    '----------'
    ' Methodes '
    '----------'

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'

End Class
