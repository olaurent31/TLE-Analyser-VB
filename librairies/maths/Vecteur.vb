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
' > Utilitaires sur les vecteurs 3D
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
' > 06/03/2010 : suite a la modification de la classe Observateur, les modificateurs sont devenus obsoletes

Option Explicit On
Option Strict On

Imports System.Math

Public Class Vecteur

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
    Private _x As Double
    Private _y As Double
    Private _z As Double

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
    ''' Constructeur a partir d'un vecteur
    ''' </summary>
    ''' <param name="vec">Vecteur</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal vec As Vecteur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _x = vec._x
        _y = vec._y
        _z = vec._z

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir de 3 reels
    ''' </summary>
    ''' <param name="x">Composante x du vecteur</param>
    ''' <param name="y">Composante y du vecteur</param>
    ''' <param name="z">Composante z du vecteur</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal x As Double, ByVal y As Double, ByVal z As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _x = x
        _y = y
        _z = z

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Norme d'un vecteur 3D
    ''' </summary>
    ''' <returns>Norme du vecteur</returns>
    ''' <remarks></remarks>
    Public Function Norme() As Double

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
        Return Sqrt(_x * _x + _y * _y + _z * _z)
    End Function

    ''' <summary>
    ''' Oppose d'un vecteur
    ''' </summary>
    ''' <returns>Oppose du vecteur</returns>
    ''' <remarks></remarks>
    Public Function Oppose() As Vecteur

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
        Return New Vecteur(-_x, -_y, -_z)
    End Function

    ''' <summary>
    ''' Soustraction de 2 vecteurs 3D
    ''' </summary>
    ''' <param name="vec">Vecteur soustractif</param>
    ''' <returns>Vecteur soustrait</returns>
    ''' <remarks></remarks>
    Public Function Moins(ByVal vec As Vecteur) As Vecteur

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
        Return New Vecteur(_x - vec._x, _y - vec._y, _z - vec._z)
    End Function

    ''' <summary>
    ''' Multiplication d'un vecteur 3D par un scalaire
    ''' </summary>
    ''' <param name="xval">Valeur scalaire</param>
    ''' <returns>Vecteur multiplie par le scalaire</returns>
    ''' <remarks></remarks>
    Public Function MultScal(ByVal xval As Double) As Vecteur

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
        Return New Vecteur(_x * xval, _y * xval, _z * xval)
    End Function

    ''' <summary>
    ''' Calcul du vecteur 3D unitaire
    ''' </summary>
    ''' <returns>Vecteur 3D unitaire</returns>
    ''' <remarks></remarks>
    Public Function Normalise() As Vecteur

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim norm, val As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        norm = Norme()

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        val = CDbl(IIf(norm < Maths.EPSDBL100, 1.0, 1.0 / norm))

        '--------'
        ' Retour '
        '--------'
        Return New Vecteur(MultScal(val))
    End Function

    ''' <summary>
    ''' Produit scalaire de 2 vecteurs 3D
    ''' </summary>
    ''' <param name="vec">Vecteur</param>
    ''' <returns>Produit scalaire</returns>
    ''' <remarks></remarks>
    Public Function ProduitScalaire(ByVal vec As Vecteur) As Double

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
        Return (_x * vec._x + _y * vec._y + _z * vec._z)
    End Function

    ''' <summary>
    ''' Produit vectoriel de 2 vecteurs 3D
    ''' </summary>
    ''' <param name="vec">Vecteur</param>
    ''' <returns>Produit vectoriel</returns>
    ''' <remarks></remarks>
    Public Function ProduitVectoriel(ByVal vec As Vecteur) As Vecteur

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
        Return New Vecteur(_y * vec._z - _z * vec._y, _z * vec._x - _x * vec._z, _x * vec._y - _y * vec._x)
    End Function

    ''' <summary>
    ''' Angle entre 2 vecteurs 3D
    ''' </summary>
    ''' <param name="vec">Vecteur</param>
    ''' <returns>Angle en radians</returns>
    ''' <remarks></remarks>
    Public Function Angle(ByVal vec As Vecteur) As Double

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim ang As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        Dim norme1 As Double = Norme()
        Dim norme2 As Double = vec.Norme()

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        If norme1 < Maths.EPSDBL OrElse norme2 < Maths.EPSDBL Then
            ang = 0.0
        Else
            ang = Acos(ProduitScalaire(vec) / (norme1 * norme2))
        End If

        '--------'
        ' Retour '
        '--------'
        Return (ang)
    End Function

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property X() As Double
        Get
            Return _x
        End Get
    End Property

    Public ReadOnly Property Y() As Double
        Get
            Return _y
        End Get
    End Property

    Public ReadOnly Property Z() As Double
        Get
            Return _z
        End Get
    End Property
End Class
