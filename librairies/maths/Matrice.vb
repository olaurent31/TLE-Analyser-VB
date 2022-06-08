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
' > Utilitaires sur les matrices
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
' > 06/03/2010 : Simplification de la classe

Option Explicit On
Option Strict On

Public Class Matrice

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
    Private _vec1 As Vecteur
    Private _vec2 As Vecteur
    Private _vec3 As Vecteur

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur par defaut
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir d'une matrice
    ''' </summary>
    ''' <param name="mat">Matrice</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal mat As Matrice)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _vec1 = mat._vec1
        _vec2 = mat._vec2
        _vec3 = mat._vec3

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir de 3 vecteurs colonnes
    ''' </summary>
    ''' <param name="vec1">Vecteur colonne 1</param>
    ''' <param name="vec2">Vecteur colonne 2</param>
    ''' <param name="vec3">Vecteur colonne 3</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal vec1 As Vecteur, ByVal vec2 As Vecteur, ByVal vec3 As Vecteur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _vec1 = vec1
        _vec2 = vec2
        _vec3 = vec3

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Transposition d'une matrice
    ''' </summary>
    ''' <returns>Matrice transposee</returns>
    ''' <remarks></remarks>
    Public Function Transpose() As Matrice

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Dim l As New Vecteur(_vec1.X, _vec2.X, _vec3.X)
        Dim m As New Vecteur(_vec1.Y, _vec2.Y, _vec3.Y)
        Dim n As New Vecteur(_vec1.Z, _vec2.Z, _vec3.Z)

        '--------'
        ' Retour '
        '--------'
        Return New Matrice(l, m, n)
    End Function

    ''' <summary>
    ''' Multiplication d'un vecteur 3D par une matrice 3x3
    ''' </summary>
    ''' <param name="vec">Vecteur</param>
    ''' <returns>Vecteur multiplie par la matrice</returns>
    ''' <remarks></remarks>
    Public Function MultVec(ByVal vec As Vecteur) As Vecteur

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim x, y, z As Double

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        x = _vec1.X * vec.X + _vec2.X * vec.Y + _vec3.X * vec.Z
        y = _vec1.Y * vec.X + _vec2.Y * vec.Y + _vec3.Y * vec.Z
        z = _vec1.Z * vec.X + _vec2.Z * vec.Y + _vec3.Z * vec.Z

        '--------'
        ' Retour '
        '--------'
        Return (New Vecteur(x, y, z))
    End Function

    ''' <summary>
    ''' Multiplication d'une matrice 3x3 par une matrice 3x3
    ''' </summary>
    ''' <param name="mat">Matrice</param>
    ''' <returns>Produit des deux matrices</returns>
    ''' <remarks></remarks>
    Public Function MatMult(ByVal mat As Matrice) As Matrice

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'

        '--------'
        ' Retour '
        '--------'
        Return (New Matrice(MultVec(mat._vec1), MultVec(mat._vec2), MultVec(mat._vec3)))
    End Function

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'

End Class
