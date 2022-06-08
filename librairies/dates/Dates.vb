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
' > Utilitaires sur les dates
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
' > 19/02/2011 : Modification de la gestion des dates
'

Option Explicit On
Option Strict On

Imports System.Globalization
Imports System.Math

Public Class Dates

    '----------------------'
    ' Constantes publiques '
    '----------------------'
    Public Const AN2000 As Integer = 2000
    Public Const NB_MIN_PAR_HEUR As Double = 60.0
    Public Const NB_HEUR_PAR_JOUR As Double = 24.0
    Public Const NB_MIN_PAR_JOUR As Double = 1440.0
    Public Const NB_SEC_PAR_MIN As Double = 60.0
    Public Const NB_SEC_PAR_HEUR As Double = 3600.0
    Public Const NB_SEC_PAR_JOUR As Double = 86400.0

    Public Const NB_JOURS_PAR_ANJ As Double = 365.25
    Public Const NB_JOURS_PAR_SIECJ As Double = 36525.0

    Public Const DATE_INFINIE As Double = 9999999.0
    Public Const TJ2000 As Double = 2451545.0
    Public Const EPS_DATES As Double = 0.0000001

    Public Const NB_JOUR_PAR_HEUR As Double = 1.0 / NB_HEUR_PAR_JOUR
    Public Const NB_JOUR_PAR_MIN As Double = 1.0 / NB_MIN_PAR_JOUR
    Public Const NB_JOUR_PAR_SEC As Double = 1.0 / NB_SEC_PAR_JOUR
    Public Const NB_HEUR_PAR_MIN As Double = 1.0 / NB_MIN_PAR_HEUR
    Public Const NB_HEUR_PAR_SEC As Double = 1.0 / NB_SEC_PAR_HEUR
    Public Const NB_MIN_PAR_SEC As Double = 1.0 / NB_SEC_PAR_MIN
    Public Const NB_ANJ_PAR_JOURS As Double = 1.0 / NB_JOURS_PAR_ANJ
    Public Const NB_SIECJ_PAR_JOURS As Double = 1.0 / NB_JOURS_PAR_SIECJ

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
    Private _annee As Integer
    Private _mois As Integer
    Private _jour As Integer
    Private _heure As Integer
    Private _minutes As Integer
    Private _secondes As Double

    Private _jourJulien As Double
    Private _jourJulienUTC As Double
    Private _offsetUTC As Double

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur par defaut (obtention de la date systeme)
    ''' </summary>
    ''' <param name="offsetUTC">Ecart heure legale - UTC</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal offsetUTC As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        With Now
            _annee = .Year
            _mois = .Month
            _jour = .Day
            _heure = .Hour
            _minutes = .Minute
            _secondes = .Second
        End With

        _offsetUTC = offsetUTC

        CalculJourJulien()

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir d'une date
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <param name="offset">Ecart heure legale - UTC</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dat As Dates, Optional ByVal offset As Double = 0.0)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _annee = dat._annee
        _mois = dat._mois
        _jour = dat._jour
        _heure = dat._heure
        _minutes = dat._minutes
        _secondes = dat._secondes

        _offsetUTC = offset

        _jourJulien = dat._jourJulien
        _jourJulienUTC = _jourJulien - offset

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir d'un jour julien 2000
    ''' </summary>
    ''' <param name="jourJulien">Jour julien 2000</param>
    ''' <param name="offsetUTC">Ecart heure legale - UTC</param>
    ''' <param name="acalc">Calcul de la date calendaire</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal jourJulien As Double, ByVal offsetUTC As Double, Optional ByVal acalc As Boolean = True)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim aa, b, c, d, e, z As Integer
        Dim f, j0, j1 As Double

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        If acalc Then
            j1 = jourJulien + 0.5
            z = CInt(Floor(j1))
            f = j1 - z
            z += CInt(TJ2000 + Maths.EPSDBL100)

            If z < 2299161 Then
                aa = z
            Else
                Dim al As Integer
                al = CInt(Floor((z - 1867216.25) / 36524.25))
                aa = z + 1 + al - al \ 4
            End If

            b = aa + 1524
            c = CInt(Floor((b - 122.1) * NB_ANJ_PAR_JOURS))
            d = CInt(Floor(NB_JOURS_PAR_ANJ * c))
            e = CInt(Floor((b - d) / 30.6001))
            j0 = b - d - Floor(30.6001 * e) + f

            _mois = CInt(IIf(e < 14, e - 1, e - 13))
            _annee = CInt(IIf(_mois > 2, c - 4716, c - 4715))
            _jour = CInt(Floor(j0))

            _heure = CInt(Floor(NB_HEUR_PAR_JOUR * (j0 - _jour)))
            _minutes = CInt(Floor(NB_MIN_PAR_JOUR * (j0 - _jour) - NB_MIN_PAR_HEUR * _heure))
            _secondes = NB_SEC_PAR_JOUR * (j0 - _jour) - NB_SEC_PAR_HEUR * _heure - NB_SEC_PAR_MIN * _minutes
        End If

        _offsetUTC = offsetUTC
        _jourJulien = jourJulien
        _jourJulienUTC = _jourJulien - _offsetUTC

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir d'une date calendaire, ou le jour est exprime sous forme decimale
    ''' </summary>
    ''' <param name="annee">Annee (1957-2056)</param>
    ''' <param name="mois">Mois</param>
    ''' <param name="xj">Jour (decimal)</param>
    ''' <param name="offsetUTC">Ecart heure legale - UTC</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal annee As Integer, ByVal mois As Integer, ByVal xj As Double, ByVal offsetUTC As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _annee = annee
        _mois = mois
        _jour = CInt(Floor(xj))
        _heure = CInt(Floor(NB_HEUR_PAR_JOUR * (xj - _jour)))
        _minutes = CInt(Floor(NB_MIN_PAR_JOUR * (xj - _jour) - NB_MIN_PAR_HEUR * _heure))
        _secondes = NB_SEC_PAR_JOUR * (xj - _jour) - NB_SEC_PAR_HEUR * _heure - NB_SEC_PAR_MIN * _minutes

        _offsetUTC = offsetUTC

        CalculJourJulien()

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Constructeur a partir d'une date calendaire
    ''' </summary>
    ''' <param name="annee">Annee (1957-2056)</param>
    ''' <param name="mois">Mois</param>
    ''' <param name="jour">Jour</param>
    ''' <param name="heure">Heure</param>
    ''' <param name="minutes">Minutes</param>
    ''' <param name="secondes">Secondes</param>
    ''' <param name="offsetUTC">Ecart heure legale - UTC</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal annee As Integer, ByVal mois As Integer, ByVal jour As Integer, ByVal heure As Integer, ByVal minutes As Integer, ByVal secondes As Double, ByVal offsetUTC As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _annee = annee
        _mois = mois
        _jour = jour
        _heure = heure
        _minutes = minutes
        _secondes = secondes

        _offsetUTC = offsetUTC

        CalculJourJulien()

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Conversion d'une date de type Dates au type DateTime
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns>Date au format DateTime</returns>
    ''' <remarks></remarks>
    Public Function ToDate(ByVal type As Integer) As DateTime

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim sec As Integer
        Dim dat As Dates

        '-----------------'
        ' Initialisations '
        '-----------------'
        sec = CInt(IIf(type = 0, 0.0, Round(_secondes)))

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        dat = New Dates(_annee, _mois, _jour, _heure, _minutes, sec, _offsetUTC)

        '--------'
        ' Retour '
        '--------'
        With dat
            Return New DateTime(._annee, ._mois, ._jour, ._heure, ._minutes, CInt(._secondes))
        End With
    End Function

    ''' <summary>
    ''' Conversion de la date en date locale
    ''' </summary>
    ''' <returns>Date locale</returns>
    ''' <remarks></remarks>
    Public Function ToLocalDate() As Dates

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
        Return New Dates(_jourJulienUTC + _offsetUTC)
    End Function

    ''' <summary>
    ''' Conversion de la date en date locale
    ''' </summary>
    ''' <param name="offsetUTC">Ecart date locale - UTC</param>
    ''' <returns>Date locale</returns>
    ''' <remarks></remarks>
    Public Function ToLocalDate(ByVal offsetUTC As Double) As Dates

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
        Return New Dates(_jourJulienUTC + offsetUTC)
    End Function

    ''' <summary>
    ''' Obtention de la date sous forme de chaine de caracteres (type ISO)
    ''' </summary>
    ''' <param name="type">Nombre de decimales pour les secondes (0 ou 1)</param>
    ''' <returns>Date sous forme de chaine de caracteres</returns>
    ''' <remarks></remarks>
    Public Function ToShortDate(ByVal type As Integer) As String

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim jjsec, tmp As Double
        Dim fmt As String
        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat
        Dim dat As Dates

        '-----------------'
        ' Initialisations '
        '-----------------'
        fmt = IIf(type = 0, "00", "00.0").ToString
        tmp = Floor(_jourJulien)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        'jjsec = Round(NB_SEC_PAR_JOUR * (_jourJulien - tmp), CInt(IIf(type = 0, 0, 1))) * NB_JOUR_PAR_SEC + tmp + Maths.EPSDBL100
        'dat = New Dates(jjsec, _offsetUTC)

        '--------'
        ' Retour '
        '--------'
        With dat
            Return String.Format(nfi, "{0:d4}/{1:d2}/{2:d2} {3:d2}:{4:d2}:{5}", ._annee, ._mois, ._jour, ._heure, ._minutes, ._secondes.ToString(fmt, nfi))
        End With
    End Function

    ''' <summary>
    ''' Obtention de la date sous forme de chaine de caracteres
    ''' </summary>
    ''' <returns>Date sous forme de chaine de caracteres</returns>
    ''' <remarks></remarks>
    Public Function ToLongDate() As String

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim sdate As Date

        '-----------------'
        ' Initialisations '
        '-----------------'
        sdate = New DateTime(_annee, _mois, _jour, _heure, _minutes, CInt(_secondes))

        '---------------------'
        ' Corps de la methode '
        '---------------------'

        '--------'
        ' Retour '
        '--------'
        Return sdate.ToLongDateString.Substring(0, 1).ToUpper(CultureInfo.CurrentCulture) & sdate.ToLongDateString.Substring(1) & "  " & sdate.ToLongTimeString
    End Function

    ''' <summary>
    ''' Calcul du jour julien 2000 a partir de la date calendaire. Le jour julien
    ''' est nul pour le 1er janvier 2000 12h. La date est supposee correcte et
    ''' appartenir au calendrier gregorien.
    ''' Inspire du calcul du jour julien dans l'Astronomical Algorithms
    ''' 2nd edition de Jean Meeus, p61
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CalculJourJulien()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim b, c, d, n As Integer
        Dim xj As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        d = _annee
        n = _mois
        xj = _jour + _heure * NB_JOUR_PAR_HEUR + _minutes * NB_JOUR_PAR_MIN + _secondes * NB_JOUR_PAR_SEC

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        If n < 3 Then
            d -= 1
            n += 12
        End If

        c = d \ 100
        b = 2 - c + c \ 4
        d -= AN2000

        _jourJulien = Floor(NB_JOURS_PAR_ANJ * d) + Floor(30.6001 * (n + 1)) + xj + b - 50.5
        _jourJulienUTC = _jourJulien - _offsetUTC

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Annee() As Integer
        Get
            Return _annee
        End Get
    End Property

    Public ReadOnly Property JourJulien() As Double
        Get
            Return _jourJulien
        End Get
    End Property

    Public ReadOnly Property JourJulienUTC() As Double
        Get
            Return _jourJulienUTC
        End Get
    End Property

    Public ReadOnly Property OffsetUTC() As Double
        Get
            Return _offsetUTC
        End Get
    End Property
End Class
