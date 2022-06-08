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
' > Donnees relatives a un corps (satellite, Soleil, ...)
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
' > 31/01/2010 : Correction de la valeur de l'altitude du corps (cas rares)
' > 19/02/2011 : Modification de la gestion des dates
'

Option Explicit On
'Option Strict On

Imports System.Globalization
Imports System.IO
Imports System.Math

Public Class Corps

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
    ' Coordonnees horizontales
    Protected _azimut As Double
    Protected _hauteur As Double
    Protected _distance As Double

    ' Coordonnees equatoriales
    Protected _ascensionDroite As Double
    Protected _declinaison As Double
    Protected _constellation As String

    ' Coordonnees terrestres
    Protected _longitude As Double
    Protected _latitude As Double
    Protected _altitude As Double

    Protected _visible As Boolean
    Protected _rangeRate As Double

    ' Coordonnees cartesiennes
    Protected _position As New Vecteur
    Protected _vitesse As New Vecteur
    Protected _dist As New Vecteur

    ' Zone de visibilite
    Protected _zone(360) As PointF

    '-------------------'
    ' Variables privees '
    '-------------------'
    Private Shared _initCst As Boolean
    Private Shared _tabConst(358) As String
    Private Shared _tabCstCoord(358, 2) As Double

    '---------------'
    ' Constructeurs '
    '---------------'

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul des coordonnees equatoriales du corps (satellite, Soleil ...)
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <remarks></remarks>
    Public Sub CalculCoordEquat(ByVal observateur As Observateur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim atrouve As Boolean
        Dim i As Integer
        Dim ch As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        atrouve = False
        i = 0
        ch = Cos(_hauteur)
        Dim vec1 As New Vecteur(-Cos(_azimut) * ch, Sin(_azimut) * ch, Sin(_hauteur))
        Dim vec2 As New Vecteur(observateur.RotEq.MultVec(vec1))

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Declinaison
        _declinaison = Asin(vec2.Z)

        ' Ascension droite
        _ascensionDroite = Atan2(vec2.Y, vec2.X)
        If _ascensionDroite < 0.0 Then _ascensionDroite += Maths.DEUX_PI

        ' Determination de la constellation
        Try

            If Not _initCst Then
                InitTabConstellations()
                _initCst = True
            End If

            While Not atrouve
                If _declinaison >= _tabCstCoord(i, 2) Then
                    If _ascensionDroite < _tabCstCoord(i, 1) Then
                        If _ascensionDroite >= _tabCstCoord(i, 0) Then
                            atrouve = True
                            _constellation = _tabConst(i)
                        End If
                    End If
                End If
                i += 1
            End While

        Catch ex As PreviSatException

        End Try

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul des coordonnees horizontales du corps
    ''' Le calcul de la refraction est issu de l'Astronomical Algorithms 2nd edition de Jean Meeus, p106.
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <remarks></remarks>
    Public Sub CalculCoordHoriz(ByVal observateur As Observateur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim ht, htd, r As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        _dist = _position.Moins(observateur.Position)
        _distance = _dist.Norme
        Dim distp As New Vecteur(_vitesse.Moins(observateur.Vitesse))

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Taux de variation de la distance a l'observateur
        _rangeRate = _dist.ProduitScalaire(distp) / _distance

        Dim vec As New Vecteur(observateur.RotHz.MultVec(_dist))

        ' Azimut
        _azimut = Atan2(vec.Y, -vec.X)
        If _azimut < 0.0 Then _azimut += Maths.DEUX_PI

        ' Hauteur
        ht = Asin(vec.Z / _distance)

        ' Prise en compte de la refraction atmospherique
        htd = Maths.RAD2DEG * ht
        r = Maths.DEG2RAD * 1.02 / (Maths.ARCMIN_PAR_DEG * Tan(Maths.DEG2RAD * (htd + 10.3 / (htd + 5.11))))

        _hauteur = ht + r

        ' Visibilite du corps
        If _hauteur >= 0.0 Then
            _visible = True
        Else
            _visible = False
            _hauteur = ht
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul des coordonnees terrestres du corps a la date courante
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <remarks></remarks>
    Public Sub CalculCoordTerrestres(ByVal observateur As Observateur)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Longitude
        _longitude = (observateur.TempsSideralGreenwich - Atan2(_position.Y, _position.X)) Mod Maths.DEUX_PI
        If _longitude > PI Then _longitude -= Maths.DEUX_PI
        If _longitude < -PI Then _longitude += Maths.DEUX_PI

        CalculLatitudeAltitude()
       
        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul des coordonnees terrestres du corps a une date donnee
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <remarks></remarks>
    Public Sub CalculCoordTerrestres(ByVal dat As Dates)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Longitude
        _longitude = (Observateur.TempsSideralDeGreenwich(dat) - Atan2(_position.Y, _position.X)) Mod Maths.DEUX_PI
        If _longitude > PI Then _longitude -= Maths.DEUX_PI
        If _longitude < -PI Then _longitude += Maths.DEUX_PI

        'CalculLatitudeAltitude()

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul de la zone de visibilite
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CalculZoneVisibilite()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i As Integer
        Dim az, beta, cb, cl0, la1, lo0, lo1, numden, sb, sl0 As Double
        Dim SunPos As New Vecteur(Home.XS, Home.YS, Home.ZS)
        Dim LATSRAD = Maths.DEG2RAD * (Home.LATS)
        Dim LONGISRAD = Maths.DEG2RAD * (Home.LONGIS)

        '-----------------'
        ' Initialisations '
        '-----------------'
        lo0 = -LONGISRAD '_longitude
        If lo0 > 0.0 Then lo0 -= Maths.DEUX_PI
        cl0 = Cos(LATSRAD) '_latitude)
        sl0 = Sin(LATSRAD) '_latitude)
        beta = Acos((EarthEquRad - 15.0) / SunPos.Norme) '_position.Norme)
        If Double.IsNaN(beta) Then beta = 0.0
        cb = Cos(beta)
        sb = Sin(beta)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        For i = 0 To 359
            az = Maths.DEG2RAD * CDbl(i)
            la1 = Asin(sl0 * cb + Cos(az) * sb * cl0)

            numden = (cb - sl0 * Sin(la1)) / (cl0 * Cos(la1))

            If i = 0 AndAlso beta > Maths.PI_SUR_DEUX - LATSRAD Then
                lo1 = lo0 + PI
            Else
                If i = 180 AndAlso beta > Maths.PI_SUR_DEUX + LATSRAD Then
                    lo1 = lo0 + PI
                Else
                    If Abs(numden) > 1.0 Then
                        lo1 = lo0
                    Else
                        If i <= 180 Then
                            lo1 = lo0 + Acos(numden)
                        Else
                            lo1 = lo0 - Acos(numden)
                        End If
                    End If
                End If
            End If
            Home.SunZone(i).X = CSng(((PI - lo1) Mod Maths.DEUX_PI) * Maths.RAD2DEG)
            Home.SunZone(i).Y = CSng((Maths.PI_SUR_DEUX - la1) * Maths.RAD2DEG)
        Next
        Home.SunZone(360).X = Home.SunZone(0).X
        Home.SunZone(360).Y = Home.SunZone(0).Y

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Formattage des angles
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub FormatteAngles()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        _azimut *= Maths.RAD2DEG
        _hauteur *= Maths.RAD2DEG
        _declinaison *= Maths.RAD2DEG
        _ascensionDroite *= Maths.RAD2HEUR

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Initialisation du tableau de constellations (un seul acces fichier)
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub InitTabConstellations()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i As Integer
        Dim ligne As String
        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat

        '-----------------'
        ' Initialisations '
        '-----------------'
        i = 0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Dim sr As New StreamReader(My.Application.Info.DirectoryPath & "\data\constellations.cst")
        Do Until sr.Peek = -1
            ligne = sr.ReadLine
            _tabConst(i) = ligne.Substring(0, 3)
            _tabCstCoord(i, 0) = Double.Parse(ligne.Substring(4, 7), nfi) * Maths.HEUR2RAD
            _tabCstCoord(i, 1) = Double.Parse(ligne.Substring(12, 7), nfi) * Maths.HEUR2RAD
            _tabCstCoord(i, 2) = Maths.DEG2RAD * Double.Parse(ligne.Substring(20, 8), nfi)
            i += 1
        Loop
        sr.Close()
        If i = 0 Then Throw New PreviSatException

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul de la latitude et de l'altitude survolees par le corps
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CalculLatitudeAltitude()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim c, lat, r0, sph As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        c = 1.0
        _latitude = PI

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Latitude
        r0 = Sqrt(_position.X * _position.X + _position.Y * _position.Y)
        _latitude = Atan2(_position.Z, r0)
        Do
            lat = _latitude
            sph = Sin(lat)
            c = 1.0 / Sqrt(1.0 - Terre.E2 * sph * sph)
            _latitude = Atan((_position.Z + Terre.RAYON * c * Terre.E2 * sph) / r0)
        Loop While Abs(_latitude - lat) > 0.0000001

        ' Altitude
        _altitude = CDbl(IIf(r0 < 0.001, Abs(_position.Z) - Terre.RAYON * (1.0 - Terre.APLA), r0 / Cos(_latitude) - Terre.RAYON * c))

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Altitude() As Double
        Get
            Return _altitude
        End Get
    End Property

    Public ReadOnly Property AscensionDroite() As Double
        Get
            Return _ascensionDroite
        End Get
    End Property

    Public ReadOnly Property Azimut() As Double
        Get
            Return _azimut
        End Get
    End Property

    Public ReadOnly Property Constellation() As String
        Get
            Return _constellation
        End Get
    End Property

    Public ReadOnly Property Declinaison() As Double
        Get
            Return _declinaison
        End Get
    End Property

    Public ReadOnly Property Distance() As Double
        Get
            Return _distance
        End Get
    End Property

    Public ReadOnly Property Dist() As Vecteur
        Get
            Return _dist
        End Get
    End Property

    Public ReadOnly Property Hauteur() As Double
        Get
            Return _hauteur
        End Get
    End Property

    Public ReadOnly Property Latitude() As Double
        Get
            Return _latitude
        End Get
    End Property

    Public ReadOnly Property Longitude() As Double
        Get
            Return _longitude
        End Get
    End Property

    Public Property Position() As Vecteur
        Get
            Return _position
        End Get
        Set(ByVal position As Vecteur)
            _position = position
        End Set
    End Property

    Public ReadOnly Property RangeRate() As Double
        Get
            Return _rangeRate
        End Get
    End Property

    Public ReadOnly Property Visible() As Boolean
        Get
            Return _visible
        End Get
    End Property

    Public ReadOnly Property Vitesse() As Vecteur
        Get
            Return _vitesse
        End Get
    End Property

    Public ReadOnly Property Zone() As PointF()
        Get
            Return _zone
        End Get
    End Property
End Class
