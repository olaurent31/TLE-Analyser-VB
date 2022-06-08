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
' > Utilitaires pour le calcul de la position des satellites
'
' Auteur
' > Astropedia
'
' Date de creation
' > 12/04/2009
'
' Date de revision
' > 19/02/2011
'
' Revisions
' > 13/11/2010 : Ajout de l'accesseur AnomalieExcentrique
' > 19/02/2011 : - Modification de la gestion de la date (arguments des methodes)
'                - Modification de la gestion du calcul de la position des satellites
'

Option Explicit On
'Option Strict On

Imports System.Globalization
Imports System.IO
Imports System.Math

Public Class Satellite
    Inherits Corps

    '----------------------'
    ' Constantes publiques '
    '----------------------'

    '----------------------'
    ' Constantes protegees '
    '----------------------'

    '--------------------'
    ' Constantes privees '
    '--------------------'
    Private Const J3SJ2 As Double = Terre.J3 / Terre.J2
    Private Const RPTIM As Double = Terre.OMEGA * Dates.NB_SEC_PAR_MIN

    ' Pour le modele haute orbite
    Private Const ZEL As Double = 0.0549
    Private Const ZES As Double = 0.01675
    Private Const ZNL As Double = 0.00015835218
    Private Const ZNS As Double = 0.0000119459
    Private Const C1SS As Double = 0.0000029864797
    Private Const C1L As Double = 0.00000047968065
    Private Const ZSINIS As Double = 0.39785416
    Private Const ZCOSIS As Double = 0.91744867
    Private Const ZCOSGS As Double = 0.1945905
    Private Const ZSINGS As Double = -0.98088458
    Private Const Q22 As Double = 0.0000017891679
    Private Const Q31 As Double = 0.0000021460748
    Private Const Q33 As Double = 0.00000022123015
    Private Const ROOT22 As Double = 0.0000017891679
    Private Const ROOT44 As Double = 0.0000000073636953
    Private Const ROOT54 As Double = 0.0000000021765803
    Private Const ROOT32 As Double = 0.00000037393792
    Private Const ROOT52 As Double = 0.00000011428639
    Private Const FASX2 As Double = 0.13130908
    Private Const FASX4 As Double = 2.8843198
    Private Const FASX6 As Double = 0.37448087
    Private Const G22 As Double = 5.7686396
    Private Const G32 As Double = 0.95240898
    Private Const G44 As Double = 1.8014998
    Private Const G52 As Double = 1.050833
    Private Const G54 As Double = 4.4108898
    Private Const STEPP As Double = 720.0
    Private Const STEPN As Double = -720.0
    Private Const STEP2 As Double = 259200.0

    '---------------------'
    ' Variables publiques '
    '---------------------'
    Public Shared initCalcul As Boolean

    '---------------------'
    ' Variables protegees '
    '---------------------'

    '-------------------'
    ' Variables privees '
    '-------------------'
    ' Donnees sur le satellite
    Private _eclipse As Boolean = True
    Private _penombre As Boolean
    Private _methMagnitude As Char
    Private _nbOrbites As Integer
    Private _elongation As Double
    Private _fractionIlluminee As Double
    Private _magnitude As Double
    Private _magnitudeStandard As Double
    Private _rayonApparentSoleil As Double
    Private _rayonApparentTerre As Double
    Private _section As Double
    Private _t1 As Double
    Private _t2 As Double
    Private _t3 As Double

    Private _ieralt As Boolean
    Private _method As Char

    Private _tle As TLE
    Private _sat As New SatVariables

    ' Trajectoire
    Private _trajectoire(360, 2) As Double

    ' Elements osculateurs
    Private _elements As New ElementsOsculateurs

    '---------------'
    ' Constructeurs '
    '---------------'
    ''' <summary>
    ''' Constructeur par defaut
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal tle As TLE)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'
        If tle Is Nothing Then
            Throw New PreviSatException
        End If

        '-----------------------'
        ' Corps du constructeur '
        '-----------------------'
        _tle = tle
        SGP4Init()

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '----------'
    ' Methodes '
    '----------'
    ''' <summary>
    ''' Calcul de la position et de la vitesse d'un satellite
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <remarks>Modele SGP4 : d'apres l'article "Revisiting Spacetrack Report #3: Rev 1" de David Vallado (2006)</remarks>
    Public Sub CalculPosVit(ByVal dat As Double, ByVal dat2 As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim ktr As Integer
        Dim am, argpdf, axnl, aynl, coseo1, cosip, ecose, el2, eo1, esine, nodedf, pl, sineo1, sinip, t2, tc, tem5, temp, tempa, tempe, templ, tsince, u, xl, xlm, xmdf As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        ' Calcul du temps ecoule depuis l'epoque (en minutes)
        tsince = Dates.NB_MIN_PAR_JOUR * (dat - dat2) '(dat.JourJulienUTC - _tle.Epoque.JourJulienUTC)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Try

            _sat.t = tsince

            ' Prise en compte des termes seculaires de la gravite et du freinage atmospherique
            xmdf = (_sat.mo + _sat.mdot * _sat.t) Mod Maths.DEUX_PI
            argpdf = _sat.argpo + _sat.argpdot * _sat.t
            nodedf = _sat.omegao + _sat.nodedot * _sat.t
            _sat.argpm = argpdf
            _sat.mm = xmdf
            t2 = _sat.t * _sat.t
            _sat.nodem = nodedf + _sat.nodecf * t2
            tempa = 1.0 - _sat.cc1 * _sat.t
            tempe = _sat.bstar * _sat.cc4 * _sat.t
            templ = _sat.t2cof * t2

            If Not _sat.isimp Then
                Dim delm, delomg, t3, t4 As Double

                delomg = _sat.omgcof * _sat.t
                delm = _sat.xmcof * (Pow((1.0 + _sat.eta * COS(xmdf)), 3.0) - _sat.delmo)
                temp = delomg + delm
                _sat.mm = xmdf + temp
                _sat.argpm = argpdf - temp
                t3 = t2 * _sat.t
                t4 = t3 * _sat.t
                tempa += -_sat.d2 * t2 - _sat.d3 * t3 - _sat.d4 * t4
                tempe += _sat.bstar * _sat.cc5 * (SIN(_sat.mm) - _sat.sinmao)
                templ += _sat.t3cof * t3 + t4 * (_sat.t4cof + _sat.t * _sat.t5cof)
            End If

            _sat.nm = _sat.no
            _sat.em = _sat.ecco
            _sat.inclm = _sat.inclo

            If _method = "d"c Then
                tc = _sat.t
                Dspace(tc)
            End If

            am = Pow((Terre.KE / _sat.nm), Maths.DEUX_TIERS) * tempa * tempa
            _sat.nm = Terre.KE * Pow(am, -1.5)
            _sat.em -= tempe

            If _sat.em < 0.000001 Then _sat.em = 0.000001

            _sat.mm += _sat.no * templ
            xlm = _sat.mm + _sat.argpm + _sat.nodem
            _sat.emsq = _sat.em * _sat.em
            temp = 1.0 - _sat.emsq

            _sat.nodem = _sat.nodem Mod Maths.DEUX_PI
            _sat.argpm = _sat.argpm Mod Maths.DEUX_PI
            xlm = xlm Mod Maths.DEUX_PI
            _sat.mm = (xlm - _sat.argpm - _sat.nodem) Mod Maths.DEUX_PI

            _sat.sinim = SIN(_sat.inclm)
            _sat.cosim = COS(_sat.inclm)

            ' Prise en compte des termes periodiques luni-solaires
            _sat.ep = _sat.em
            _sat.xincp = _sat.inclm
            _sat.argpp = _sat.argpm
            _sat.omegap = _sat.nodem
            _sat.mp = _sat.mm
            sinip = _sat.sinim
            cosip = _sat.cosim
            If _method = "d"c Then
                Dpper()
                If _sat.xincp < 0.0 Then
                    _sat.xincp = -_sat.xincp
                    _sat.omegap += PI
                    _sat.argpp -= PI
                End If
            End If

            ' Termes longue periode
            If _method = "d"c Then
                sinip = SIN(_sat.xincp)
                cosip = COS(_sat.xincp)
                _sat.aycof = -0.5 * J3SJ2 * sinip

                If (Abs(cosip + 1.0) > 0.0000000000015) Then
                    _sat.xlcof = -0.25 * J3SJ2 * sinip * (3.0 + 5.0 * cosip) / (1.0 + cosip)
                Else
                    _sat.xlcof = -0.25 * J3SJ2 * sinip * (3.0 + 5.0 * cosip) / 0.0000000000015
                End If
            End If

            axnl = _sat.ep * COS(_sat.argpp)
            temp = 1.0 / (am * (1.0 - _sat.ep * _sat.ep))
            aynl = _sat.ep * SIN(_sat.argpp) + temp * _sat.aycof
            xl = _sat.mp + _sat.argpp + _sat.omegap + temp * _sat.xlcof * axnl

            ' Resolution de l'equation de Kepler
            u = (xl - _sat.omegap) Mod Maths.DEUX_PI
            eo1 = u
            tem5 = 9999.9
            ktr = 1

            While ((Abs(tem5) >= 0.000000000001) AndAlso (ktr <= 10))
                sineo1 = SIN(eo1)
                coseo1 = COS(eo1)
                tem5 = 1.0 - coseo1 * axnl - sineo1 * aynl
                tem5 = (u - aynl * coseo1 + axnl * sineo1 - eo1) / tem5
                If (Abs(tem5) >= 0.95) Then tem5 = CDbl(IIf(tem5 > 0.0, 0.95, -0.95))
                eo1 += tem5
                ktr += 1
            End While

            ' Termes courte periode
            ecose = axnl * coseo1 + aynl * sineo1
            esine = axnl * sineo1 - aynl * coseo1
            el2 = axnl * axnl + aynl * aynl
            pl = am * (1.0 - el2)
            If pl >= 0.0 Then
                Dim betal, cnod, cos2u, cosi, cossu, cosu, mrt, mvt, rdotl, rl, rvdotl, rvdot, sin2u, sinsu, sinu, snod, su, temp1, temp2, xinc, xnode As Double

                rl = am * (1.0 - ecose)
                rdotl = Sqrt(am) * esine / rl
                rvdotl = Sqrt(pl) / rl
                betal = Sqrt(1.0 - el2)
                temp = esine / (1.0 + betal)
                sinu = am / rl * (sineo1 - aynl - axnl * temp)
                cosu = am / rl * (coseo1 - axnl + aynl * temp)
                su = Atan2(sinu, cosu)
                sin2u = (cosu + cosu) * sinu
                cos2u = 1.0 - 2.0 * sinu * sinu
                temp = 1.0 / pl
                temp1 = 0.5 * Terre.J2 * temp
                temp2 = temp1 * temp

                ' Prise en compte des termes courte periode
                If _method = "d"c Then
                    Dim cosisq As Double

                    cosisq = cosip * cosip
                    _sat.con41 = 3.0 * cosisq - 1.0
                    _sat.x1mth2 = 1.0 - cosisq
                    _sat.x7thm1 = 7.0 * cosisq - 1.0
                End If

                mrt = Terre.RAYON * (rl * (1.0 - 1.5 * temp2 * betal * _sat.con41) + 0.5 * temp1 * _sat.x1mth2 * cos2u)
                su -= 0.25 * temp2 * _sat.x7thm1 * sin2u
                xnode = _sat.omegap + 1.5 * temp2 * cosip * sin2u
                xinc = _sat.xincp + 1.5 * temp2 * cosip * sinip * cos2u
                mvt = Terre.RAYON * Dates.NB_MIN_PAR_SEC * (Terre.KE * rdotl - _sat.nm * temp1 * _sat.x1mth2 * sin2u)
                rvdot = Terre.RAYON * Dates.NB_MIN_PAR_SEC * (Terre.KE * rvdotl + _sat.nm * temp1 * (_sat.x1mth2 * cos2u + 1.5 * _sat.con41))

                ' Vecteurs directeurs
                snod = SIN(xnode)
                cnod = COS(xnode)
                cosi = COS(xinc)
                sinsu = SIN(su)
                cossu = COS(su)

                Dim mm As New Vecteur(-snod * cosi, cnod * cosi, SIN(xinc))
                Dim nn As New Vecteur(cnod, snod, 0.0)

                Dim uu As New Vecteur(mm.MultScal(sinsu).Moins(nn.MultScal(-cossu)))
                Dim vv As New Vecteur(mm.MultScal(cossu).Moins(nn.MultScal(sinsu)))

                ' Position et vitesse
                _position = uu.MultScal(mrt)
                _vitesse = uu.MultScal(mvt).Moins(vv.MultScal(-rvdot))

            End If

        Catch ex As PreviSatException

        End Try


        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Determination de la condition d'eclipse du satellite
    ''' </summary>
    ''' <param name="sun">Soleil</param>
    ''' <remarks></remarks>
    Public Sub CalculEclipse(ByVal sun As Soleil)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        _rayonApparentTerre = Asin(Terre.RAYON / _position.Norme)
        If Double.IsNaN(_rayonApparentTerre) Then _rayonApparentTerre = Maths.PI_SUR_DEUX
        Dim rho As New Vecteur(sun.Position.Moins(_position))
        _rayonApparentSoleil = Asin(Soleil.RAYON / rho.Norme)

        _elongation = sun.Position.Angle(_position.Oppose)

        ' Test si le satellite est en phase d'eclipse
        _eclipse = CBool(IIf(_rayonApparentTerre > _rayonApparentSoleil AndAlso _elongation < _rayonApparentTerre - _rayonApparentSoleil, True, False))

        ' Test si le satellite est dans la penombre
        _penombre = CBool(IIf(Abs(_rayonApparentTerre - _rayonApparentSoleil) < _elongation AndAlso _elongation < _rayonApparentTerre + _rayonApparentSoleil, True, False))

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul de la magnitude du satellite
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <param name="extinction">Prise en compte de l'extinction atmospherique</param>
    ''' <remarks></remarks>
    Public Sub CalculMagnitude(ByVal observateur As Observateur, ByVal extinction As Boolean)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'

        '-----------------'
        ' Initialisations '
        '-----------------'
        _magnitude = 99.0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        If Not _eclipse Then

            ' Fraction illuminee
            _fractionIlluminee = 0.5 * (1.0 + Cos(_elongation))

            ' Magnitude
            If _magnitudeStandard < 99.0 Then
                _magnitude = _magnitudeStandard - 15.75 + 2.5 * Log10(_distance * _distance / _fractionIlluminee)
                If extinction Then _magnitude += ExtinctionAtmospherique(observateur)
            End If
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Determination de l'extinction atmospherique, issu de l'article
    ''' "Magnitude corrections for atmospheric extinction" de Daniel Green, 1992
    ''' </summary>
    ''' <param name="observateur">Observateur</param>
    ''' <returns>Correction due a l'extinction atmospherique</returns>
    ''' <remarks></remarks>
    Public Function ExtinctionAtmospherique(ByVal observateur As Observateur) As Double

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim corr As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        corr = 0.0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        If _hauteur > 0.0 Then
            Dim cosz, x As Double

            cosz = Cos(Maths.PI_SUR_DEUX - _hauteur)
            x = 1.0 / (cosz + 0.025 * Exp(-11.0 * cosz))
            corr = x * (0.016 + observateur.Aray + observateur.Aaer)
        End If

        '--------'
        ' Retour '
        '--------'
        Return (corr)
    End Function

    ''' <summary>
    ''' Calcul des elements osculateurs du satellite
    ''' </summary>
    ''' <param name="dat">Dates</param>
    ''' <param name="tle">TLE du satellite</param>
    ''' <remarks></remarks>
    Public Sub CalculElementsOsculateurs(ByVal dat As Dates, ByVal tle As TLE)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim age, er, ne, nn0, rv, alpha, alpha2, beta, gamma, gamma2, norme, temp1, temp2, v, v2 As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        temp1 = 1.0 / Terre.GE
        norme = _position.Norme
        If norme < Maths.EPSDBL100 Then Exit Sub
        temp2 = 1.0 / norme
        v = _vitesse.Norme
        v2 = v * v
        rv = _position.ProduitScalaire(_vitesse)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Demi-grand axe
        _elements._demiGrandAxe = 1.0 / (2.0 * temp2 - v2 * temp1)

        ' Excentricite
        Dim exc As New Vecteur(_position.MultScal(v2 * temp1 - temp2).Moins(_vitesse.MultScal(rv * temp1)))
        _elements._excentricite = exc.Norme

        ' Inclinaison
        Dim h As New Vecteur(_position.ProduitVectoriel(_vitesse))
        _elements._inclinaison = Acos(h.Z / h.Norme)

        ' Ascension droite du noeud ascendant
        Dim n As New Vecteur(-h.Y, h.X, 0.0)
        _elements._ascDroiteNA = Acos(n.X / n.Norme)
        If n.Y < 0.0 Then _elements._ascDroiteNA = Maths.DEUX_PI - _elements._ascDroiteNA

        ' Argument du perigee
        ne = n.ProduitScalaire(exc)
        _elements._argumentPerigee = Acos(ne / (n.Norme * _elements._excentricite))
        If exc.Z < 0.0 Then _elements._argumentPerigee = Maths.DEUX_PI - _elements._argumentPerigee

        ' Anomalie vraie
        er = exc.ProduitScalaire(_position)
        _elements._anomalieVraie = Acos(er * temp2 / _elements._excentricite)
        If rv < 0.0 Then _elements._anomalieVraie = Maths.DEUX_PI - _elements._anomalieVraie

        ' Anomalie excentrique
        _elements._anomalieExcentrique = 2.0 * Atan(Sqrt((1.0 - _elements._excentricite) / (1.0 + _elements._excentricite)) * Tan(0.5 * _elements._anomalieVraie))
        If _elements._anomalieExcentrique < 0.0 Then _elements._anomalieExcentrique += Maths.DEUX_PI

        ' Anomalie moyenne
        _elements._anomalieMoyenne = _elements._anomalieExcentrique - _elements._excentricite * Sin(_elements._anomalieExcentrique)
        If _elements._anomalieMoyenne < 0.0 Then _elements._anomalieMoyenne += Maths.DEUX_PI

        alpha = _elements._demiGrandAxe / Terre.RAYON
        alpha2 = alpha * alpha
        beta = 1.0 - _elements._excentricite * _elements._excentricite
        gamma = 1.0 + _elements._excentricite * Cos(_elements._argumentPerigee)
        gamma2 = gamma * gamma
        temp1 = Sin(_elements._inclinaison)
        nn0 = 1.0 - 1.5 * Terre.J2 * ((2.0 - 2.5 * temp1 * temp1) / (alpha2 * Sqrt(beta) * gamma2) + gamma2 * gamma / (alpha2 * beta * beta * beta))

        ' Apogee, perigee, periode orbitale
        _elements._apogee = _elements._demiGrandAxe * (1.0 + _elements._excentricite)
        _elements._perigee = _elements._demiGrandAxe * (1.0 - _elements._excentricite)
        _elements._periode = nn0 * Maths.DEUX_PI * Sqrt(_elements._demiGrandAxe * _elements._demiGrandAxe * _elements._demiGrandAxe / Terre.GE) * Dates.NB_HEUR_PAR_SEC

        ' Nombre d'orbites
        age = dat.JourJulienUTC - tle.Epoque.JourJulienUTC
        _nbOrbites = tle.Nb0 + CInt(Floor((tle.No + age * tle.Bstar) * age + ((tle.Omegao + tle.Mo) Mod Maths.DEUX_PI) / Maths.T360 - ((_elements._argumentPerigee + _elements._anomalieVraie) Mod Maths.DEUX_PI) / Maths.DEUX_PI + 0.5))

        ' Formattage des elements
        _elements._inclinaison *= Maths.RAD2DEG
        _elements._ascDroiteNA *= Maths.RAD2DEG
        _elements._argumentPerigee *= Maths.RAD2DEG
        _elements._anomalieVraie *= Maths.RAD2DEG
        _elements._anomalieMoyenne *= Maths.RAD2DEG
        _elements._anomalieExcentrique *= Maths.RAD2DEG

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul de la position d'une liste de satellites
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <param name="tle">Ensemble des TLE de la liste de satellites</param>
    ''' <param name="observateur">Observateur</param>
    ''' <param name="soleil">Soleil</param>
    ''' <param name="nbTrajectoires">Nombre de trajectoires a calculer</param>
    ''' <param name="visibilite">Calcul de la zone de visibilite des satellites</param>
    ''' <param name="extinction">Prise en compte de l'extinction atmospherique</param>
    ''' <param name="satellites">Liste de satellites</param>
    ''' <remarks></remarks>
    Public Shared Sub CalculPositionListeSatellites(ByVal dat As Dates, ByVal tle() As TLE, ByVal observateur As Observateur, ByVal soleil As Soleil, ByVal nbTrajectoires As Integer, ByVal visibilite As Boolean, ByVal extinction As Boolean, ByRef satellites() As Satellite)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim isat As Integer

        '-----------------'
        ' Initialisations '
        '-----------------'
        If Not initCalcul Then
            ReDim satellites(tle.Length - 1)
            For isat = 0 To tle.Length - 1
                satellites(isat) = New Satellite(tle(isat))
            Next
            If tle.Length > 0 Then initCalcul = True
        End If

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Calcul de la position des satellites de la liste 
        For isat = 0 To tle.Length - 1

            ' Calcul de la position du satellite courant
            satellites(isat).CalculPosVit(Home.EPOCH, Home.TLEPOCH)

            ' Calcul des coordonnees horizontales
            satellites(isat).CalculCoordHoriz(observateur)

            ' Calcul des conditions d'eclipse
            satellites(isat).CalculEclipse(soleil)

            ' Calcul des coordonnees terrestres
            'satellites(isat).CalculCoordTerrestres(obs(0))
            If satellites(isat).Altitude < 0.0 Then satellites(isat)._ieralt = True Else satellites(isat)._ieralt = False

            If visibilite Then
                ' Calcul de la zone de visibilite des satellites
                satellites(isat).CalculZoneVisibilite()
            End If

            If isat = 0 Then

                ' Calcul des coordonnees equatoriales
                satellites(isat).CalculCoordEquat(observateur)

                ' Calcul de la magnitude
                satellites(isat).CalculMagnitude(observateur, extinction)

                ' Calcul des elements osculateurs et du nombre d'orbites
                satellites(isat).CalculElementsOsculateurs(dat, tle(isat))

                ' Calcul de la trajectoire
                satellites(isat).CalculTrajectoire(dat, tle(isat), nbTrajectoires)

            End If
        Next

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Lecture du fichier de magnitude standard
    ''' </summary>
    ''' <param name="listeSatellites">Liste de satellites</param>
    ''' <param name="sat">Ensemble des satellites</param>
    ''' <remarks></remarks>
    Public Shared Sub LectureMagnitudeStandard(ByVal listeSatellites() As String, ByVal tabtle() As TLE, ByRef sat() As Satellite)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim isat, j As Integer
        Dim ligne As String
        Dim nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat

        '-----------------'
        ' Initialisations '
        '-----------------'
        j = 0
        If Not initCalcul Then
            ReDim sat(listeSatellites.Length - 1)
            For isat = 0 To listeSatellites.Length - 1
                sat(isat) = New Satellite(tabtle(isat))
            Next
            If listeSatellites.Length > 0 Then initCalcul = True
        End If
        For isat = 0 To listeSatellites.Length - 1
            sat(isat)._magnitudeStandard = 99.0
        Next

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        Dim sr As New StreamReader(My.Application.Info.DirectoryPath & "\data\magnitude")
        Do Until sr.Peek = -1 OrElse j > listeSatellites.Length - 1

            ligne = sr.ReadLine
            For isat = 0 To listeSatellites.Length - 1
                If listeSatellites(isat).Equals(ligne.Substring(0, 5)) Then
                    sat(isat)._t1 = Double.Parse(ligne.Substring(6, 5), nfi)
                    sat(isat)._t2 = Double.Parse(ligne.Substring(12, 4), nfi)
                    sat(isat)._t3 = Double.Parse(ligne.Substring(17, 4), nfi)
                    sat(isat)._magnitudeStandard = Double.Parse(ligne.Substring(22, 4), nfi)
                    sat(isat)._methMagnitude = ligne.Chars(27)
                    sat(isat)._section = Double.Parse(ligne.Substring(29, 6), nfi)
                    j += 1
                    Exit For
                End If
            Next
        Loop
        sr.Close()

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    Public Sub SGP4Init()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim qzms2t, ss As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        _method = "n"c

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Recuperation des elements du TLE et formattage
        _sat.argpo = Maths.DEG2RAD * _tle.Argpo
        _sat.bstar = _tle.Bstar
        _sat.ecco = _tle.Ecco
        _sat.epoque = _tle.Epoque
        _sat.inclo = Maths.DEG2RAD * _tle.Inclo
        _sat.mo = Maths.DEG2RAD * _tle.Mo
        _sat.no = _tle.No * Maths.DEUX_PI * Dates.NB_JOUR_PAR_MIN
        _sat.omegao = Maths.DEG2RAD * _tle.Omegao

        ss = 78.0 / Terre.RAYON + 1.0
        qzms2t = Pow((120.0 - 78.0) / Terre.RAYON, 4.0)
        _sat.t = 0.0

        Initl()

        If _sat.omeosq >= 0.0 OrElse _sat.no > 0.0 Then
            Dim cc2, cc3, coef, coef1, cosio4, eeta, etasq, perige, pinvsq, psisq, qzms24, sfour, temp1, temp2, temp3, tsi, xhdot1 As Double

            _sat.isimp = False
            If _sat.rp < 220.0 / Terre.RAYON + 1.0 Then _sat.isimp = True

            sfour = ss
            qzms24 = qzms2t

            If _sat.rp < 1.0 Then Throw New PreviSatException

            perige = (_sat.rp - 1.0) * Terre.RAYON
            If perige < 156.0 Then
                sfour = perige - 78.0
                If perige < 98.0 Then sfour = 20.0

                qzms24 = Pow((120.0 - sfour) / Terre.RAYON, 4.0)
                sfour /= Terre.RAYON + 1.0
            End If

            pinvsq = 1.0 / _sat.posq
            tsi = 1.0 / (_sat.ao - sfour)
            _sat.eta = _sat.ao * _sat.ecco * tsi

            etasq = _sat.eta * _sat.eta
            eeta = _sat.ecco * _sat.eta
            psisq = Abs(1.0 - etasq)
            coef = qzms24 * Pow(tsi, 4.0)
            coef1 = coef * Pow(psisq, -3.5)
            cc2 = coef1 * _sat.no * (_sat.ao * (1.0 + 1.5 * etasq + eeta * (4.0 + etasq)) + 0.375 * Terre.J2 * tsi / psisq * _sat.con41 * (8.0 + 3.0 * etasq * (8.0 + etasq)))
            _sat.cc1 = _sat.bstar * cc2
            cc3 = 0.0
            If _sat.ecco > 0.0001 Then cc3 = -2.0 * coef * tsi * J3SJ2 * _sat.no * _sat.sinio / _sat.ecco

            _sat.x1mth2 = 1.0 - _sat.cosio2
            _sat.cc4 = 2.0 * _sat.no * coef1 * _sat.ao * _sat.omeosq * (_sat.eta * (2.0 + 0.5 * etasq) + _sat.ecco * (0.5 + 2.0 * etasq) - Terre.J2 * tsi / (_sat.ao * psisq) * (-3.0 * _sat.con41 * (1.0 - 2.0 * eeta + etasq * (1.5 - 0.5 * eeta)) + 0.75 * _sat.x1mth2 * (2.0 * etasq - eeta * (1.0 + etasq)) * COS(2.0 * _sat.argpo)))
            _sat.cc5 = 2.0 * coef1 * _sat.ao * _sat.omeosq * (1.0 + 2.75 * (etasq + eeta) + eeta * etasq)

            cosio4 = _sat.cosio2 * _sat.cosio2
            temp1 = 1.5 * Terre.J2 * pinvsq * _sat.no
            temp2 = 0.5 * temp1 * Terre.J2 * pinvsq
            temp3 = -0.46875 * Terre.J4 * pinvsq * pinvsq * _sat.no
            _sat.mdot = _sat.no + 0.5 * temp1 * _sat.rteosq * _sat.con41 + 0.0625 * temp2 * _sat.rteosq * (13.0 - 78.0 * _sat.cosio2 + 137.0 * cosio4)
            _sat.argpdot = -0.5 * temp1 * _sat.con42 + 0.0625 * temp2 * (7.0 - 114.0 * _sat.cosio2 + 395.0 * cosio4) + temp3 * (3.0 - 36.0 * _sat.cosio2 + 49.0 * cosio4)
            xhdot1 = -temp1 * _sat.cosio
            _sat.nodedot = xhdot1 + (0.5 * temp2 * (4.0 - 19.0 * _sat.cosio2) + 2.0 * temp3 * (3.0 - 7.0 * _sat.cosio2)) * _sat.cosio
            _sat.xpidot = _sat.argpdot + _sat.nodedot
            _sat.omgcof = _sat.bstar * cc3 * COS(_sat.argpo)
            _sat.xmcof = 0.0
            If _sat.ecco > 0.0001 Then _sat.xmcof = -Maths.DEUX_TIERS * coef * _sat.bstar / eeta

            _sat.nodecf = 3.5 * _sat.omeosq * xhdot1 * _sat.cc1
            _sat.t2cof = 1.5 * _sat.cc1

            If (Abs(_sat.cosio + 1.0) > 0.0000000000015) Then
                _sat.xlcof = -0.25 * J3SJ2 * _sat.sinio * (3.0 + 5.0 * _sat.cosio) / (1.0 + _sat.cosio)
            Else
                _sat.xlcof = -0.25 * J3SJ2 * _sat.sinio * (3.0 + 5.0 * _sat.cosio) / 0.0000000000015
            End If

            _sat.aycof = -0.5 * J3SJ2 * _sat.sinio
            _sat.delmo = Pow((1.0 + _sat.eta * COS(_sat.mo)), 3.0)
            _sat.sinmao = SIN(_sat.mo)
            _sat.x7thm1 = 7.0 * _sat.cosio2 - 1.0

            ' Initialisation du modele haute orbite
            If (Maths.DEUX_PI / _sat.no) >= 225.0 Then
                _method = "d"c
                _sat.isimp = True
                Dim tc As Double = 0.0
                _sat.inclm = _sat.inclo

                Dscom(tc)

                _sat.mp = _sat.mo
                _sat.argpp = _sat.argpo
                _sat.ep = _sat.ecco
                _sat.omegap = _sat.omegao
                _sat.xincp = _sat.inclo

                Dpper()

                _sat.argpm = 0.0
                _sat.nodem = 0.0
                _sat.mm = 0.0

                Dsinit(tc)
            End If

            If Not _sat.isimp Then
                Dim cc1sq, temp As Double

                cc1sq = _sat.cc1 * _sat.cc1
                _sat.d2 = 4.0 * _sat.ao * tsi * cc1sq
                temp = _sat.d2 * tsi * _sat.cc1 / 3.0
                _sat.d3 = (17.0 * _sat.ao + sfour) * temp
                _sat.d4 = 0.5 * temp * _sat.ao * tsi * (221.0 * _sat.ao + 31.0 * sfour) * _sat.cc1
                _sat.t3cof = _sat.d2 + 2.0 * cc1sq
                _sat.t4cof = 0.25 * (3.0 * _sat.d3 + _sat.cc1 * (12.0 * _sat.d2 + 10.0 * cc1sq))
                _sat.t5cof = 0.2 * (3.0 * _sat.d4 + 12.0 * _sat.cc1 * _sat.d3 + 6.0 * _sat.d2 * _sat.d2 + 15.0 * cc1sq * (2.0 * _sat.d2 + cc1sq))
            End If
            _sat.init = True
        Else
            Throw New PreviSatException
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    Private Sub Initl()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim adel, ak, d1, del, po As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        If _sat.no < Maths.EPSDBL100 Then Throw New PreviSatException

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        _sat.eccsq = _sat.ecco * _sat.ecco
        _sat.omeosq = 1.0 - _sat.eccsq
        _sat.rteosq = Sqrt(_sat.omeosq)
        _sat.cosio = COS(_sat.inclo)
        _sat.cosio2 = _sat.cosio * _sat.cosio

        ak = Pow(Terre.KE / _sat.no, Maths.DEUX_TIERS)
        d1 = 0.75 * Terre.J2 * (3.0 * _sat.cosio2 - 1.0) / (_sat.rteosq * _sat.omeosq)
        del = d1 / (ak * ak)
        adel = ak * (1.0 - del * del - del * (0.5 * Maths.DEUX_TIERS + 134.0 * del * del / 81.0))
        del = d1 / (adel * adel)
        _sat.no /= (1.0 + del)

        _sat.ao = Pow(Terre.KE / _sat.no, Maths.DEUX_TIERS)
        _sat.sinio = SIN(_sat.inclo)
        po = _sat.ao * _sat.omeosq
        _sat.con42 = 1.0 - 5.0 * _sat.cosio2
        _sat.con41 = -_sat.con42 - _sat.cosio2 - _sat.cosio2
        _sat.posq = po * po
        _sat.rp = _sat.ao * (1.0 - _sat.ecco)
        _method = "n"c

        _sat.gsto = Observateur.TempsSideralDeGreenwich(_sat.epoque)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    Private Sub Dscom(ByVal tc As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim lsflg As Integer
        Dim betasq, ctem, stem, xnodce, zcosgl, zcoshl, zcosil, zsingl, zsinhl, zsinil, zx, zy As Double

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        _sat.nm = _sat.no
        _sat.em = _sat.ecco
        _sat.snodm = SIN(_sat.omegao)
        _sat.cnodm = COS(_sat.omegao)
        _sat.sinomm = SIN(_sat.argpo)
        _sat.cosomm = COS(_sat.argpo)
        _sat.sinim = SIN(_sat.inclo)
        _sat.cosim = COS(_sat.inclo)
        _sat.emsq = _sat.em * _sat.em
        betasq = 1.0 - _sat.emsq
        _sat.rtemsq = Sqrt(betasq)

        ' Initialisation des termes luni-solaires
        _sat.day = _sat.epoque.JourJulienUTC + Dates.NB_JOURS_PAR_SIECJ + tc * Dates.NB_JOUR_PAR_MIN
        xnodce = (4.523602 - 0.00092422029 * _sat.day) Mod Maths.DEUX_PI
        stem = SIN(xnodce)
        ctem = COS(xnodce)
        zcosil = 0.91375164 - 0.03568096 * ctem
        zsinil = Sqrt(1.0 - zcosil * zcosil)
        zsinhl = 0.089683511 * stem / zsinil
        zcoshl = Sqrt(1.0 - zsinhl * zsinhl)
        _sat.gam = 5.8351514 + 0.001944368 * _sat.day
        zx = 0.39785416 * stem / zsinil
        zy = zcoshl * ctem + 0.91744867 * zsinhl * stem
        zx = Atan2(zx, zy)
        zx = _sat.gam + zx - xnodce
        zcosgl = COS(zx)
        zsingl = SIN(zx)

        ' Termes solaires
        Dim zcosg As Double = ZCOSGS
        Dim zsing As Double = ZSINGS
        Dim zcosi As Double = ZCOSIS
        Dim zsini As Double = ZSINIS
        Dim zcosh As Double = _sat.cnodm
        Dim zsinh As Double = _sat.snodm
        Dim cc As Double = C1SS
        Dim xnoi As Double = 1.0 / _sat.nm

        For lsflg = 1 To 2

            Dim a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, x1, x2, x3, x4, x5, x6, x7, x8 As Double

            a1 = zcosg * zcosh + zsing * zcosi * zsinh
            a3 = -zsing * zcosh + zcosg * zcosi * zsinh
            a7 = -zcosg * zsinh + zsing * zcosi * zcosh
            a8 = zsing * zsini
            a9 = zsing * zsinh + zcosg * zcosi * zcosh
            a10 = zcosg * zsini
            a2 = _sat.cosim * a7 + _sat.sinim * a8
            a4 = _sat.cosim * a9 + _sat.sinim * a10
            a5 = -_sat.sinim * a7 + _sat.cosim * a8
            a6 = -_sat.sinim * a9 + _sat.cosim * a10

            x1 = a1 * _sat.cosomm + a2 * _sat.sinomm
            x2 = a3 * _sat.cosomm + a4 * _sat.sinomm
            x3 = -a1 * _sat.sinomm + a2 * _sat.cosomm
            x4 = -a3 * _sat.sinomm + a4 * _sat.cosomm
            x5 = a5 * _sat.sinomm
            x6 = a6 * _sat.sinomm
            x7 = a5 * _sat.cosomm
            x8 = a6 * _sat.cosomm

            _sat.z31 = 12.0 * x1 * x1 - 3.0 * x3 * x3
            _sat.z32 = 24.0 * x1 * x2 - 6.0 * x3 * x4
            _sat.z33 = 12.0 * x2 * x2 - 3.0 * x4 * x4
            _sat.z1 = 3.0 * (a1 * a1 + a2 * a2) + _sat.z31 * _sat.emsq
            _sat.z2 = 6.0 * (a1 * a3 + a2 * a4) + _sat.z32 * _sat.emsq
            _sat.z3 = 3.0 * (a3 * a3 + a4 * a4) + _sat.z33 * _sat.emsq
            _sat.z11 = -6.0 * a1 * a5 + _sat.emsq * (-24.0 * x1 * x7 - 6.0 * x3 * x5)
            _sat.z12 = -6.0 * (a1 * a6 + a3 * a5) + _sat.emsq * (-24.0 * (x2 * x7 + x1 * x8) - 6.0 * (x3 * x6 + x4 * x5))
            _sat.z13 = -6.0 * a3 * a6 + _sat.emsq * (-24.0 * x2 * x8 - 6.0 * x4 * x6)
            _sat.z21 = 6.0 * a2 * a5 + _sat.emsq * (24.0 * x1 * x5 - 6.0 * x3 * x7)
            _sat.z22 = 6.0 * (a4 * a5 + a2 * a6) + _sat.emsq * (24.0 * (x2 * x5 + x1 * x6) - 6.0 * (x4 * x7 + x3 * x8))
            _sat.z23 = 6.0 * a4 * a6 + _sat.emsq * (24.0 * x2 * x6 - 6.0 * x4 * x8)
            _sat.z1 = _sat.z1 + _sat.z1 + betasq * _sat.z31
            _sat.z2 = _sat.z2 + _sat.z2 + betasq * _sat.z32
            _sat.z3 = _sat.z3 + _sat.z3 + betasq * _sat.z33
            _sat.s3 = cc * xnoi
            _sat.s2 = -0.5 * _sat.s3 / _sat.rtemsq
            _sat.s4 = _sat.s3 * _sat.rtemsq
            _sat.s1 = -15.0 * _sat.em * _sat.s4
            _sat.s5 = x1 * x3 + x2 * x4
            _sat.s6 = x2 * x3 + x1 * x4
            _sat.s7 = x2 * x4 - x1 * x3

            ' Termes lunaires
            If lsflg = 1 Then
                _sat.ss1 = _sat.s1
                _sat.ss2 = _sat.s2
                _sat.ss3 = _sat.s3
                _sat.ss4 = _sat.s4
                _sat.ss5 = _sat.s5
                _sat.ss6 = _sat.s6
                _sat.ss7 = _sat.s7
                _sat.sz1 = _sat.z1
                _sat.sz2 = _sat.z2
                _sat.sz3 = _sat.z3
                _sat.sz11 = _sat.z11
                _sat.sz12 = _sat.z12
                _sat.sz13 = _sat.z13
                _sat.sz21 = _sat.z21
                _sat.sz22 = _sat.z22
                _sat.sz23 = _sat.z23
                _sat.sz31 = _sat.z31
                _sat.sz32 = _sat.z32
                _sat.sz33 = _sat.z33
                zcosg = zcosgl
                zsing = zsingl
                zcosi = zcosil
                zsini = zsinil
                zcosh = zcoshl * _sat.cnodm + zsinhl * _sat.snodm
                zsinh = _sat.snodm * zcoshl - _sat.cnodm * zsinhl
                cc = C1L
            End If
        Next

        _sat.zmol = (4.7199672 + 0.2299715 * _sat.day - _sat.gam) Mod Maths.DEUX_PI
        _sat.zmos = (6.2565837 + 0.017201977 * _sat.day) Mod Maths.DEUX_PI

        ' Termes solaires
        _sat.se2 = 2.0 * _sat.ss1 * _sat.ss6
        _sat.se3 = 2.0 * _sat.ss1 * _sat.ss7
        _sat.si2 = 2.0 * _sat.ss2 * _sat.sz12
        _sat.si3 = 2.0 * _sat.ss2 * (_sat.sz13 - _sat.sz11)
        _sat.sl2 = -2.0 * _sat.ss3 * _sat.sz2
        _sat.sl3 = -2.0 * _sat.ss3 * (_sat.sz3 - _sat.sz1)
        _sat.sl4 = -2.0 * _sat.ss3 * (-21.0 - 9.0 * _sat.emsq) * ZES
        _sat.sgh2 = 2.0 * _sat.ss4 * _sat.sz32
        _sat.sgh3 = 2.0 * _sat.ss4 * (_sat.sz33 - _sat.sz31)
        _sat.sgh4 = -18.0 * _sat.ss4 * ZES
        _sat.sh2 = -2.0 * _sat.ss2 * _sat.sz22
        _sat.sh3 = -2.0 * _sat.ss2 * (_sat.sz23 - _sat.sz21)

        ' Termes lunaires
        _sat.ee2 = 2.0 * _sat.s1 * _sat.s6
        _sat.e3 = 2.0 * _sat.s1 * _sat.s7
        _sat.xi2 = 2.0 * _sat.s2 * _sat.z12
        _sat.xi3 = 2.0 * _sat.s2 * (_sat.z13 - _sat.z11)
        _sat.xl2 = -2.0 * _sat.s3 * _sat.z2
        _sat.xl3 = -2.0 * _sat.s3 * (_sat.z3 - _sat.z1)
        _sat.xl4 = -2.0 * _sat.s3 * (-21.0 - 9.0 * _sat.emsq) * ZEL
        _sat.xgh2 = 2.0 * _sat.s4 * _sat.z32
        _sat.xgh3 = 2.0 * _sat.s4 * (_sat.z33 - _sat.z31)
        _sat.xgh4 = -18.0 * _sat.s4 * ZEL
        _sat.xh2 = -2.0 * _sat.s2 * _sat.z22
        _sat.xh3 = -2.0 * _sat.s2 * (_sat.z23 - _sat.z21)

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    Private Sub Dpper()

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim f2, f3, pe, pgh, ph, pinc, pl, sel, ses, sghl, sghs, shll, shs, sil, sinzf, sis, sll, sls, zf, zm As Double

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        zm = _sat.zmos + ZNS * _sat.t
        If Not _sat.init Then zm = _sat.zmos

        zf = zm + 2.0 * ZES * SIN(zm)
        sinzf = SIN(zf)
        f2 = 0.5 * sinzf * sinzf - 0.25
        f3 = -0.5 * sinzf * COS(zf)
        ses = _sat.se2 * f2 + _sat.se3 * f3
        sis = _sat.si2 * f2 + _sat.si3 * f3
        sls = _sat.sl2 * f2 + _sat.sl3 * f3 + _sat.sl4 * sinzf
        sghs = _sat.sgh2 * f2 + _sat.sgh3 * f3 + _sat.sgh4 * sinzf
        shs = _sat.sh2 * f2 + _sat.sh3 * f3
        zm = _sat.zmol + ZNL * _sat.t
        If Not _sat.init Then zm = _sat.zmol

        zf = zm + 2.0 * ZEL * SIN(zm)
        sinzf = SIN(zf)
        f2 = 0.5 * sinzf * sinzf - 0.25
        f3 = -0.5 * sinzf * COS(zf)

        sel = _sat.ee2 * f2 + _sat.e3 * f3
        sil = _sat.xi2 * f2 + _sat.xi3 * f3
        sll = _sat.xl2 * f2 + _sat.xl3 * f3 + _sat.xl4 * sinzf
        sghl = _sat.xgh2 * f2 + _sat.xgh3 * f3 + _sat.xgh4 * sinzf
        shll = _sat.xh2 * f2 + _sat.xh3 * f3
        pe = ses + sel
        pinc = sis + sil
        pl = sls + sll
        pgh = sghs + sghl
        ph = shs + shll

        If _sat.init Then
            Dim cosip, sinip As Double

            _sat.xincp += pinc
            _sat.ep += pe
            sinip = SIN(_sat.xincp)
            cosip = COS(_sat.xincp)

            ' Application directe des termes periodiques
            If _sat.xincp >= 0.2 Then
                ph /= sinip
                pgh -= cosip * ph
                _sat.argpp += pgh
                _sat.omegap += ph
                _sat.mp += pl
            Else

                ' Application des termes periodiques avec la modification de Lyddane
                Dim alfdp, betdp, cosop, dalf, dbet, dls, sinop, xls, xnoh As Double
                sinop = SIN(_sat.omegap)
                cosop = COS(_sat.omegap)
                alfdp = sinip * sinop
                betdp = sinip * cosop
                dalf = ph * cosop + pinc * cosip * sinop
                dbet = -ph * sinop + pinc * cosip * cosop
                alfdp += dalf
                betdp += dbet
                _sat.omegap = _sat.omegap Mod Maths.DEUX_PI
                If _sat.omegap < 0.0 Then _sat.omegap += Maths.DEUX_PI

                xls = _sat.mp + _sat.argpp + cosip * _sat.omegap
                dls = pl + pgh - pinc * _sat.omegap * sinip
                xls += dls
                xnoh = _sat.omegap
                _sat.omegap = Atan2(alfdp, betdp)
                If _sat.omegap < 0.0 Then _sat.omegap += Maths.DEUX_PI

                If (Abs(xnoh - _sat.omegap) > PI) Then
                    _sat.omegap = CDbl(IIf(_sat.omegap < xnoh, _sat.omegap + Maths.DEUX_PI, _sat.omegap - Maths.DEUX_PI))
                End If

                _sat.mp += pl
                _sat.argpp = xls - _sat.mp - cosip * _sat.omegap
            End If
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    Private Sub Dsinit(ByVal tc As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim ses, sghl, sghs, sgs, shll, shs, sis, sls, theta As Double

        '-----------------'
        ' Initialisations '
        '-----------------'

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        _sat.irez = 0
        If _sat.nm < 0.0052359877 AndAlso _sat.nm > 0.0034906585 Then _sat.irez = 1
        If _sat.nm >= 0.00826 AndAlso _sat.nm <= 0.00924 AndAlso _sat.em >= 0.5 Then _sat.irez = 2

        ' Termes solaires
        ses = _sat.ss1 * ZNS * _sat.ss5
        sis = _sat.ss2 * ZNS * (_sat.sz11 + _sat.sz13)
        sls = -ZNS * _sat.ss3 * (_sat.sz1 + _sat.sz3 - 14.0 - 6.0 * _sat.emsq)
        sghs = _sat.ss4 * ZNS * (_sat.sz31 + _sat.sz33 - 6.0)
        shs = -ZNS * _sat.ss2 * (_sat.sz21 + _sat.sz23)

        If _sat.inclm < 0.052359877 OrElse _sat.inclm > PI - 0.052359877 Then shs = 0.0
        If _sat.sinim <> 0.0 Then shs /= _sat.sinim
        sgs = sghs - _sat.cosim * shs

        ' Termes lunaires
        _sat.dedt = ses + _sat.s1 * ZNL * _sat.s5
        _sat.didt = sis + _sat.s2 * ZNL * (_sat.z11 + _sat.z13)
        _sat.dmdt = sls - ZNL * _sat.s3 * (_sat.z1 + _sat.z3 - 14.0 - 6.0 * _sat.emsq)
        sghl = _sat.s4 * ZNL * (_sat.z31 + _sat.z33 - 6.0)
        shll = -ZNL * _sat.s2 * (_sat.z21 + _sat.z23)

        If _sat.inclm < 0.052359877 OrElse _sat.inclm > PI - 0.052359877 Then shll = 0.0

        _sat.domdt = sgs + sghl
        _sat.dnodt = shs
        If _sat.sinim <> 0.0 Then
            _sat.domdt -= (_sat.cosim / _sat.sinim * shll)
            _sat.dnodt += (shll / _sat.sinim)
        End If

        ' Calcul des effets de resonance haute orbite
        _sat.dndt = 0.0
        theta = (_sat.gsto + tc * RPTIM) Mod Maths.DEUX_PI
        _sat.em += _sat.dedt * _sat.t
        _sat.inclm += _sat.didt * _sat.t
        _sat.argpm += _sat.domdt * _sat.t
        _sat.nodem += _sat.dnodt * _sat.t
        _sat.mm += _sat.dmdt * _sat.t
        ' sgp4fix for negative inclinations
        ' the following if statement should be commented out
        ' If inclm < 0.0 Then
        '     sat.inclm = -sat.inclm
        '     sat.argpm = sat.argpm - PI
        '     sat.nodem = sat.nodem + PI
        ' End If

        ' Initialisation des termes de resonance
        If _sat.irez <> 0 Then

            Dim aonv As Double

            aonv = Pow(_sat.nm / Terre.KE, Maths.DEUX_TIERS)

            ' Resonance geopotentielle pour les orbites de 12h
            If _sat.irez = 2 Then
                Dim ainv2, cosisq, emo, emsqo, eoc, sini2, temp, temp1, xno2 As Double

                cosisq = _sat.cosim * _sat.cosim
                emo = _sat.em
                _sat.em = _sat.ecco
                emsqo = _sat.emsq
                _sat.emsq = _sat.eccsq
                eoc = _sat.em * _sat.emsq
                _sat.g201 = -0.306 - (_sat.em - 0.64) * 0.44

                If _sat.em <= 0.65 Then
                    _sat.g211 = 3.616 - 13.247 * _sat.em + 16.29 * _sat.emsq
                    _sat.g310 = -19.302 + 117.39 * _sat.em - 228.419 * _sat.emsq + 156.591 * eoc
                    _sat.g322 = -18.9068 + 109.7927 * _sat.em - 214.6334 * _sat.emsq + 146.5816 * eoc
                    _sat.g410 = -41.122 + 242.694 * _sat.em - 471.094 * _sat.emsq + 313.953 * eoc
                    _sat.g422 = -146.407 + 841.88 * _sat.em - 1629.014 * _sat.emsq + 1083.435 * eoc
                    _sat.g520 = -532.114 + 3017.977 * _sat.em - 5740.032 * _sat.emsq + 3708.276 * eoc
                Else
                    _sat.g211 = -72.099 + 331.819 * _sat.em - 508.738 * _sat.emsq + 266.724 * eoc
                    _sat.g310 = -346.844 + 1582.851 * _sat.em - 2415.925 * _sat.emsq + 1246.113 * eoc
                    _sat.g322 = -342.585 + 1554.908 * _sat.em - 2366.899 * _sat.emsq + 1215.972 * eoc
                    _sat.g410 = -1052.797 + 4758.686 * _sat.em - 7193.992 * _sat.emsq + 3651.957 * eoc
                    _sat.g422 = -3581.69 + 16178.11 * _sat.em - 24462.77 * _sat.emsq + 12422.52 * eoc
                    If _sat.em > 0.715 Then
                        _sat.g520 = -5149.66 + 29936.92 * _sat.em - 54087.36 * _sat.emsq + 31324.56 * eoc
                    Else
                        _sat.g520 = 1464.74 - 4664.75 * _sat.em + 3763.64 * _sat.emsq
                    End If
                End If
                If _sat.em < 0.7 Then
                    _sat.g533 = -919.2277 + 4988.61 * _sat.em - 9064.77 * _sat.emsq + 5542.21 * eoc
                    _sat.g521 = -822.71072 + 4568.6173 * _sat.em - 8491.4146 * _sat.emsq + 5337.524 * eoc
                    _sat.g532 = -853.666 + 4690.25 * _sat.em - 8624.77 * _sat.emsq + 5341.4 * eoc
                Else
                    _sat.g533 = -37995.78 + 161616.52 * _sat.em - 229838.2 * _sat.emsq + 109377.94 * eoc
                    _sat.g521 = -51752.104 + 218913.95 * _sat.em - 309468.16 * _sat.emsq + 146349.42 * eoc
                    _sat.g532 = -40023.88 + 170470.89 * _sat.em - 242699.48 * _sat.emsq + 115605.82 * eoc
                End If

                sini2 = _sat.sinim * _sat.sinim
                _sat.f220 = 0.75 * (1.0 + 2.0 * _sat.cosim + cosisq)
                _sat.f221 = 1.5 * sini2
                _sat.f321 = 1.875 * _sat.sinim * (1.0 - 2.0 * _sat.cosim - 3.0 * cosisq)
                _sat.f322 = -1.875 * _sat.sinim * (1.0 + 2.0 * _sat.cosim - 3.0 * cosisq)
                _sat.f441 = 35.0 * sini2 * _sat.f220
                _sat.f442 = 39.375 * sini2 * sini2
                _sat.f522 = 9.84375 * _sat.sinim * (sini2 * (1.0 - 2.0 * _sat.cosim - 5.0 * cosisq) + 0.33333333 * (-2.0 + 4.0 * _sat.cosim + 6.0 * cosisq))
                _sat.f523 = _sat.sinim * (4.92187512 * sini2 * (-2.0 - 4.0 * _sat.cosim + 10.0 * cosisq) + 6.56250012 * (1.0 + 2.0 * _sat.cosim - 3.0 * cosisq))
                _sat.f542 = 29.53125 * _sat.sinim * (2.0 - 8.0 * _sat.cosim + cosisq * (-12.0 + 8.0 * _sat.cosim + 10.0 * cosisq))
                _sat.f543 = 29.53125 * _sat.sinim * (-2.0 - 8.0 * _sat.cosim + cosisq * (12.0 + 8.0 * _sat.cosim - 10.0 * cosisq))

                xno2 = _sat.nm * _sat.nm
                ainv2 = aonv * aonv
                temp1 = 3.0 * xno2 * ainv2
                temp = temp1 * ROOT22
                _sat.d2201 = temp * _sat.f220 * _sat.g201
                _sat.d2211 = temp * _sat.f221 * _sat.g211
                temp1 *= aonv
                temp = temp1 * ROOT32
                _sat.d3210 = temp * _sat.f321 * _sat.g310
                _sat.d3222 = temp * _sat.f322 * _sat.g322
                temp1 *= aonv
                temp = 2.0 * temp1 * ROOT44
                _sat.d4410 = temp * _sat.f441 * _sat.g410
                _sat.d4422 = temp * _sat.f442 * _sat.g422
                temp1 *= aonv
                temp = temp1 * ROOT52
                _sat.d5220 = temp * _sat.f522 * _sat.g520
                _sat.d5232 = temp * _sat.f523 * _sat.g532
                temp = 2.0 * temp1 * ROOT54
                _sat.d5421 = temp * _sat.f542 * _sat.g521
                _sat.d5433 = temp * _sat.f543 * _sat.g533
                _sat.xlamo = (_sat.mo + _sat.omegao + _sat.omegao - theta - theta) Mod Maths.DEUX_PI
                _sat.xfact = _sat.mdot + _sat.dmdt + 2.0 * (_sat.nodedot + _sat.dnodt - RPTIM) - _sat.no
                _sat.em = emo
                _sat.emsq = emsqo
            End If

            ' Termes de resonance synchrones
            If _sat.irez = 1 Then
                _sat.g200 = 1.0 + _sat.emsq * (-2.5 + 0.8125 * _sat.emsq)
                _sat.g310 = 1.0 + 2.0 * _sat.emsq
                _sat.g300 = 1.0 + _sat.emsq * (-6.0 + 6.60937 * _sat.emsq)
                _sat.f220 = 0.75 * (1.0 + _sat.cosim) * (1.0 + _sat.cosim)
                _sat.f311 = 0.9375 * _sat.sinim * _sat.sinim * (1.0 + 3.0 * _sat.cosim) - 0.75 * (1.0 + _sat.cosim)
                _sat.f330 = 1.0 + _sat.cosim
                _sat.f330 = 1.875 * _sat.f330 * _sat.f330 * _sat.f330
                _sat.del1 = 3.0 * _sat.nm * _sat.nm * aonv * aonv
                _sat.del2 = 2.0 * _sat.del1 * _sat.f220 * _sat.g200 * Q22
                _sat.del3 = 3.0 * _sat.del1 * _sat.f330 * _sat.g300 * Q33 * aonv
                _sat.del1 *= _sat.f311 * _sat.g310 * Q31 * aonv
                _sat.xlamo = (_sat.mo + _sat.omegao + _sat.argpo - theta) Mod Maths.DEUX_PI
                _sat.xfact = _sat.mdot + _sat.xpidot - RPTIM + _sat.dmdt + _sat.domdt + _sat.dnodt - _sat.no
            End If

            ' Initialisation de l'integrateur
            _sat.xli = _sat.xlamo
            _sat.xni = _sat.no
            _sat.atime = 0.0
            _sat.nm = _sat.no + _sat.dndt
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    Private Sub Dspace(ByVal tc As Double)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim theta, xldot, xnddt, xndt As Double

        '-----------------'
        ' Initialisations '
        '-----------------'
        xldot = 0.0
        xnddt = 0.0
        xndt = 0.0

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        ' Calcul des effets de resonance haute orbite
        _sat.dndt = 0.0
        theta = (_sat.gsto + tc * RPTIM) Mod Maths.DEUX_PI
        _sat.em += _sat.dedt * _sat.t
        _sat.inclm += _sat.didt * _sat.t
        _sat.argpm += _sat.domdt * _sat.t
        _sat.nodem += _sat.dnodt * _sat.t
        _sat.mm += _sat.dmdt * _sat.t

        ' sgp4fix for negative inclinations
        ' the following if statement should be commented out
        ' If sat.inclm < 0.0 Then
        '     sat.inclm = -sat.inclm
        '     sat.argpm = sat.argpm - PI
        '     sat.nodem = sat.nodem + PI
        ' End If

        ' Integration numerique (Euler-MacLaurin)
        Dim ft As Double = 0.0
        If _sat.irez <> 0 Then
            If _sat.atime = 0.0 OrElse _sat.t * _sat.atime <= 0.0 OrElse Abs(_sat.t) < Abs(_sat.atime) Then
                _sat.atime = 0.0
                _sat.xni = _sat.no
                _sat.xli = _sat.xlamo
            End If
            _sat.delt = CDbl(IIf(_sat.t > 0.0, STEPP, STEPN))

            Dim iretn As Integer = 381
            While iretn = 381
                ' Calculs des termes derives
                ' Termes de resonance quasi-synchrones
                If _sat.irez <> 2 Then
                    xndt = _sat.del1 * SIN(_sat.xli - FASX2) + _sat.del2 * SIN(2.0 * (_sat.xli - FASX4)) + _sat.del3 * SIN(3.0 * (_sat.xli - FASX6))
                    xldot = _sat.xni + _sat.xfact
                    xnddt = _sat.del1 * COS(_sat.xli - FASX2) + 2.0 * _sat.del2 * COS(2.0 * (_sat.xli - FASX4)) + 3.0 * _sat.del3 * COS(3.0 * (_sat.xli - FASX6))
                    xnddt *= xldot
                Else
                    ' Termes de resonance d'environ 12h
                    Dim x2li, x2omi, xomi As Double

                    xomi = _sat.argpo + _sat.argpdot * _sat.atime
                    x2omi = xomi + xomi
                    x2li = _sat.xli + _sat.xli
                    xndt = _sat.d2201 * SIN(x2omi + _sat.xli - G22) + _sat.d2211 * SIN(_sat.xli - G22) + _sat.d3210 * SIN(xomi + _sat.xli - G32) + _sat.d3222 * SIN(-xomi + _sat.xli - G32) + _sat.d4410 * SIN(x2omi + x2li - G44) + _sat.d4422 * SIN(x2li - G44) + _sat.d5220 * SIN(xomi + _sat.xli - G52) + _sat.d5232 * SIN(-xomi + _sat.xli - G52) + _sat.d5421 * SIN(xomi + x2li - G54) + _sat.d5433 * SIN(-xomi + x2li - G54)
                    xldot = _sat.xni + _sat.xfact
                    xnddt = _sat.d2201 * COS(x2omi + _sat.xli - G22) + _sat.d2211 * COS(_sat.xli - G22) + _sat.d3210 * COS(xomi + _sat.xli - G32) + _sat.d3222 * COS(-xomi + _sat.xli - G32) + _sat.d5220 * COS(xomi + _sat.xli - G52) + _sat.d5232 * COS(-xomi + _sat.xli - G52) + 2.0 * (_sat.d4410 * COS(x2omi + x2li - G44) + _sat.d4422 * COS(x2li - G44) + _sat.d5421 * COS(xomi + x2li - G54) + _sat.d5433 * COS(-xomi + x2li - G54))
                    xnddt *= xldot
                End If

                ' Integrateur
                If Abs(_sat.t - _sat.atime) >= STEPP Then
                    iretn = 381
                Else
                    ft = _sat.t - _sat.atime
                    iretn = 0
                End If

                If iretn = 381 Then
                    _sat.xli += xldot * _sat.delt + xndt * STEP2
                    _sat.xni += xndt * _sat.delt + xnddt * STEP2
                    _sat.atime += _sat.delt
                End If
            End While

            Dim xl As Double
            _sat.nm = _sat.xni + xndt * ft + xnddt * ft * ft * 0.5
            xl = _sat.xli + xldot * ft + xndt * ft * ft * 0.5
            If _sat.irez <> 1 Then
                _sat.mm = xl - 2.0 * _sat.nodem + 2.0 * theta
                _sat.dndt = _sat.nm - _sat.no
            Else
                _sat.mm = xl - _sat.nodem - _sat.argpm + theta
                _sat.dndt = _sat.nm - _sat.no
            End If
            _sat.nm = _sat.no + _sat.dndt
        End If

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    ''' <summary>
    ''' Calcul de la trajectoire du satellite
    ''' </summary>
    ''' <param name="dat">Date</param>
    ''' <param name="tle">TLE du satellite</param>
    ''' <param name="nbTrajectoires">Nombre de trajectoires</param>
    ''' <remarks></remarks>
    Private Sub CalculTrajectoire(ByVal dat As Dates, ByVal tle As TLE, ByVal nbTrajectoires As Integer)

        '------------------------------------'
        ' Declarations des variables locales '
        '------------------------------------'
        Dim i As Integer
        Dim ct, lat, lon, phi, r, sph, st As Double
        Dim j0 As Dates
        Dim sat As New Satellite(tle)
        Dim soleil As New Soleil

        '-----------------'
        ' Initialisations '
        '-----------------'
        st = 1.0 / (tle.No * Maths.T360)
        ReDim _trajectoire(360 * nbTrajectoires, 2)

        '---------------------'
        ' Corps de la methode '
        '---------------------'
        For i = 0 To 360 * nbTrajectoires - 1
            j0 = New Dates(dat.JourJulienUTC + CDbl(i) * st, 0.0, False)

            ' Calcul de la position du satellite
            'sat.CalculPosVit(j0)

            ' Longitude
            lon = Maths.RAD2DEG * ((PI + Atan2(sat._position.Y, sat._position.X) - Observateur.TempsSideralDeGreenwich(j0)) Mod Maths.DEUX_PI)
            If lon < 0.0 Then lon += Maths.T360

            ' Latitude
            r = Sqrt(sat._position.X * sat._position.X + sat._position.Y * sat._position.Y)
            lat = ATAN(sat._position.Z / r)
            phi = 7.0
            While Abs(lat - phi) > 0.0000001
                phi = lat
                sph = SIN(phi)
                ct = 1.0 / Sqrt(1.0 - Terre.E2 * sph * sph)
                lat = ATAN((sat._position.Z + Terre.RAYON * ct * Terre.E2 * sph) / r)
            End While
            lat = Maths.RAD2DEG * (Maths.PI_SUR_DEUX - lat)

            ' Position du Soleil
            'soleil.CalculPosition(j0)

            ' Conditions d'eclipse
            sat.CalculEclipse(soleil)

            _trajectoire(i, 0) = lon
            _trajectoire(i, 1) = lat
            _trajectoire(i, 2) = CDbl(IIf(sat._eclipse, 1, 0))
        Next

        '--------'
        ' Retour '
        '--------'
        Return
    End Sub

    '-----------------------------'
    ' Accesseurs et modificateurs '
    '-----------------------------'
    Public ReadOnly Property Eclipse() As Boolean
        Get
            Return _eclipse
        End Get
    End Property

    Public ReadOnly Property Elements As ElementsOsculateurs
        Get
            Return _elements
        End Get
    End Property

    Public ReadOnly Property Elongation() As Double
        Get
            Return _elongation
        End Get
    End Property

    Public ReadOnly Property FractionIlluminee() As Double
        Get
            Return _fractionIlluminee
        End Get
    End Property

    Public ReadOnly Property Ieralt() As Boolean
        Get
            Return _ieralt
        End Get
    End Property

    Public ReadOnly Property Magnitude() As Double
        Get
            Return _magnitude
        End Get
    End Property

    Public ReadOnly Property MagnitudeStandard() As Double
        Get
            Return _magnitudeStandard
        End Get
    End Property

    Public ReadOnly Property MethMagnitude() As Char
        Get
            Return _methMagnitude
        End Get
    End Property

    Public ReadOnly Property Method() As Char
        Get
            Return _method
        End Get
    End Property

    Public ReadOnly Property NbOrbites() As Integer
        Get
            Return _nbOrbites
        End Get
    End Property

    Public ReadOnly Property Penombre() As Boolean
        Get
            Return _penombre
        End Get
    End Property

    Public ReadOnly Property RayonApparentTerre() As Double
        Get
            Return _rayonApparentTerre
        End Get
    End Property

    Public ReadOnly Property RayonApparentSoleil() As Double
        Get
            Return _rayonApparentSoleil
        End Get
    End Property

    Public ReadOnly Property Section() As Double
        Get
            Return _section
        End Get
    End Property

    Public ReadOnly Property T1() As Double
        Get
            Return _t1
        End Get
    End Property

    Public ReadOnly Property T2() As Double
        Get
            Return _t2
        End Get
    End Property

    Public ReadOnly Property T3() As Double
        Get
            Return _t3
        End Get
    End Property

    Public ReadOnly Property Trajectoire() As Double(,)
        Get
            Return _trajectoire
        End Get
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
