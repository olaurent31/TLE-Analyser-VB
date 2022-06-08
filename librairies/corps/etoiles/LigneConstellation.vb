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
' > Definition des lignes des constellations
'
' Auteur
' > Astropedia
'
' Date de creation
' > 20/02/2011
'
' Date de revision
' >
'
' Revisions
' >

Option Explicit On
Option Strict On

Imports System.Globalization
Imports System.IO

Public Class LigneConstellation

    '----------------------'
    ' Constantes publiques '
    '----------------------'

    '----------------------'
    ' Constantes protegees '
    '----------------------'

    '--------------------'
    ' Constantes privees '
    '--------------------'
    Private Const MAXTAB As Integer = 646

    '---------------------'
    ' Variables publiques '
    '---------------------'

    '---------------------'
    ' Variables protegees '
    '---------------------'

    '-------------------'
    ' Variables privees '
    '-------------------'
    Private _dessin As Boolean
    Private Shared _initLig As Boolean
    Private _etoile1 As Etoile
    Private _etoile2 As Etoile
    Private Shared _tabLigCst(MAXTAB, 1) As Integer

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur d'une ligne de constellation
    ''' </summary>
    ''' <param name="etoile1">Etoile 1</param>
    ''' <param name="etoile2">Etoile 2</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal etoile1 As Etoile, ByVal etoile2 As Etoile)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _etoile1 = etoile1
        _etoile2 = etoile2
        _dessin = _etoile1.Visible AndAlso _etoile2.Visible

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul des lignes de constellation
    ''' </summary>
    ''' <param name="etoiles">Catalogue d'etoiles</param>
    ''' <param name="lignesCst">Lignes de constellation</param>
    ''' <remarks></remarks>
    Public Shared Sub CalculLignesCst(ByVal etoiles() As Etoile, ByRef lignesCst() As LigneConstellation)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i, ind1, ind2 As Integer

        '-----------------'
        ' Initialisations '
        '-----------------'
        If Not _initLig Then
            ReDim lignesCst(MAXTAB - 1)
            InitTabLignesCst()
            _initLig = True
        End If

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        For i = 0 To MAXTAB - 1
            ind1 = _tabLigCst(i, 0) - 1
            ind2 = _tabLigCst(i, 1) - 1
            lignesCst(i) = New LigneConstellation(etoiles(ind1), etoiles(ind2))
        Next

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Lecture du fichier de constellations
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub InitTabLignesCst()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i, j As Integer
        Dim ligne() As String
        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat

        '-----------------'
        ' Initialisations '
        '-----------------'
        i = 0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Dim sr As New StreamReader(My.Application.Info.DirectoryPath & "\data\constlines.cst")
        Do Until sr.Peek = -1
            ligne = sr.ReadLine.Split(" "c)
            For j = 0 To 1
                _tabLigCst(i, j) = Integer.Parse(ligne(j), nfi)
            Next
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
    Public ReadOnly Property Dessin As Boolean
        Get
            Return _dessin
        End Get
    End Property

    Public ReadOnly Property Etoile1 As Etoile
        Get
            Return _etoile1
        End Get
    End Property

    Public ReadOnly Property Etoile2 As Etoile
        Get
            Return _etoile2
        End Get
    End Property
End Class
