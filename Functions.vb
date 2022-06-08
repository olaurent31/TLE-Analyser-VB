'
'    TLE ANALYSER: orbital datas, position of artificial satellites, prediction of their passes, ground track
'    Copyright (C) 2012-2013  mailto: olaurent31@gmail.com
'
'    This program is a free software
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'_______________________________________________________________________________________________________
'
' Description
' > orbital datas, position of artificial satellites, prediction of their passes, ground track, export to 3D softawres
'
' Auteur
' > Olivier LAURENT
'
' Date de creation
' > Octobre 2012
'
' Date de revision
' > 03/01/2013


Imports System.IO
'Imports ICSharpCode.SharpZipLib.Zip
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Math
Imports System.Xml
Imports System.Threading



Module Functions

    'DECLARATIONS DES FONCTIONS POUR LA LECTURE/ECRITURE DANS FICHIER INI

    'Fonction lisant un Integer
    Declare Function GetPrivateProfileInt Lib "Kernel32" Alias "GetPrivateProfileIntA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

    'Il existe aussi une fonction écrivant un Entier
    Declare Function WritePrivateProfileInt Lib "Kernel32" Alias "WritePrivateProfileIntA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lplFileName As String) As Long

    'Fonction lisant une string
    Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Short, _
    ByVal lpFileName As String) As Integer

    'Fonction écrivant une string
    Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Long

    'PUBLIC CONST
    Public Const Mu As Double = 398600.4415 'km3.s-2
    Public Const EarthMeanRad As Double = 6368.77132337461 '<From IXION Software; Old: 6371.01 'kms
    Public Const EarthPolRad As Double = 6356.751847 '<From IXION Software; Old: 6356.7523 'kms
    Public Const EarthEquRad As Double = 6378.136658 '<From IXION Software; Old: 6378.1363 'kms
    Public Const EarthFlat As Double = 0.003352806
    'Earth Porential
    Public Const J2 As Double = 0.0010826298 '<From IXION Software; Old: 0.001082626
    Public Const J3 As Double = -0.000002536151
    Public Const J4 As Double = -0.00000165597
    'Earth Porential (EIGEN Model)
    Public Const J31 As Double = 0.0000022095
    Public Const J33 As Double = 0.0000002214
    Public Const L22 As Double = 345.07 '-14.93
    Public Const L31 As Double = 6.98
    Public Const L33 As Double = 20.99
    'Stable longitudes (Geostationary Satellites)
    Public Const LSE As Double = 75.059
    Public Const LSW As Double = 348.577
    'Time Localization
    Public DS As String = DateTimeFormatInfo.CurrentInfo.DateSeparator
    Public TS As String = DateTimeFormatInfo.CurrentInfo.TimeSeparator
    Public FMT_dd_MM_yyyy As String = "dd" & DS & "MM" & DS & "yyyy"
    Public FMT_HH_mm_ss As String = "HH" & TS & "mm" & TS & "ss"
    Public FMT_HH_mm_ss_fff As String = FMT_HH_mm_ss & "ss.fff"
    'Other Parameters
    Public ReadOnly KE As Double = 60 * Sqrt(Mu / (EarthEquRad * EarthEquRad * EarthEquRad))
    Public Const Ssgp As Double = 1.01222928 / EarthEquRad
    Public Const DJ2000 As Double = 2451545
    Public Const EarthNodPrec As Double = 1.00273790934
    Public Const JDREF = 2451545 'JD de référence 01/01/2000 à 12h
    Public Const ConsGravUniv As Double = 0.000000000066726 'N.m2.kg-2
    Public Const EarthMass As Double = 5.9737E+24 'kg
    Public Const SideralTime As Double = 365.256348 'jrs '<IXION '365.25636
    Public Const TropicalTime As Double = 365.24219 'jrs
    Public Const EarthPeriod As Double = 23.934471 'hre '<IXION '23.9342 
    Public Const IhsMin As Double = 95.6768017282494 'Inclinaison Heliosynchrone Mini
    Public Const IhsMax As Double = 180 'Inclinaison Heliosynchrone Maxi

    'Adresse du fichier version
    Public CheckVersionFile = "http://www.loylart.com/tleanalyser/version.txt"

    'COMPRESION TO ZIP FILES .MKZ (NOT USED IN THIS VERSION)

    'Sub CompToZip(ByVal ZipFileName As String, ByVal PngFile As String, ByVal KmlFile As String)
    '    Dim Files(1) As String
    '    Files(0) = PngFile
    '    Files(1) = KmlFile

    '    Try
    '        '--- Definit et exécute
    '        Dim nomzip As String = ZipFileName
    '        Dim ZipStream As ZipOutputStream = New ZipOutputStream(File.Create(nomzip))
    '        ZipStream.SetLevel(9)
    '        Dim i As Integer = 0

    '        '--- Lecture des fichiers dans la liste
    '        While i <= 1
    '            Dim fichier As String = Files(i)
    '            Dim fs As FileStream = File.OpenRead(fichier)
    '            Dim buffer(fs.Length) As Byte
    '            fs.Read(buffer, 0, buffer.Length)
    '            Dim entry As ZipEntry = New ZipEntry(Path.GetFileName(fichier))
    '            ZipStream.PutNextEntry(entry)
    '            ZipStream.Write(buffer, 0, buffer.Length)
    '            i = i + 1
    '        End While
    '        '--- Termine la procédure de compression
    '        ZipStream.Finish()
    '        '--- ferme le fichier de compression
    '        ZipStream.Close()
    '        '--- Avertit l'utilisateur que la compression s'est bien passée
    '        MessageBox.Show("Fichier " + nomzip + " créé avec succès", "Succès")
    '    Catch Ex As Exception
    '        '--- La compression ne s'est pas bien passée, une erreur est survenue
    '        MessageBox.Show("Erreur lors de la création de l'archive" & Microsoft.VisualBasic.Chr(10) & "Erreur : " + Ex.Message, "Erreur")
    '    End Try

    'End Sub

    'CHECK NETWORK / TLE ANALYSER VERSION

    Function CheckNetwork() As Boolean

        If My.Computer.Network.IsAvailable = True Then Home.CheckNW = True

    End Function

    Function CheckVersion() As Boolean
        Dim Ver As Single

        CheckVersion = True

        Try
            'Download la fichier Version
            My.Computer.Network.DownloadFile(CheckVersionFile, Home.AppPath & "version.txt", "", "", False, 100000, True)

            'Récupère le n° de version
            FileOpen(1, Home.AppPath & "version.txt", OpenMode.Input)
            Ver = CSng(LineInput(1))
            FileClose(1)
            File.Delete(Home.AppPath & "version.txt")

            'Vérifie le n° avec celui du fichier .ini
            If Ver > CSng(Home.TLEAVersion) Then
                If MessageBox.Show("A new version of TLE ANALYSER is available at SourceForge.net." & vbCrLf & _
                                    "Click 'OK' to go to the Download page.", "TLE ANALYSER", _
                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) _
                                    = DialogResult.OK Then
                    Process.Start("https://sourceforge.net/projects/tleanalyser/")
                End If
            End If

        Catch Ex As Exception
            MessageBox.Show("Your Network seems to be restricted:" & vbCrLf & "TLE ANALYSER can't check new version availability." & vbCrLf & vbCrLf & Ex.Message, "TLE ANALYSER - Error")
        End Try

    End Function

    'MATH FUNCTION

    Function Round4(ByVal PARAM As Double)
        PARAM = Round(PARAM, 4)
        Return PARAM
    End Function

    'VECTORS FUNCTIONS

    Public Function Norme(ByVal Xval, ByVal Yval, ByVal Zval) As Double

        Return Sqrt(Xval * Xval + Yval * Yval + Zval * Zval)

    End Function

    Public Function ProduitVectoriel(ByVal vec1 As Vecteur, ByVal vec2 As Vecteur) As Vecteur

        Return New Vecteur(vec1.Y * vec2.Z - vec1.Z * vec2.Y, vec1.Z * vec2.X - vec1.X * vec2.Z, vec1.X * vec2.Y - vec1.Y * vec2.X)

    End Function

    Public Function ProduitScalaire(ByVal vec1 As Vecteur, ByVal vec2 As Vecteur) As Double

        Return (vec1.X * vec2.X + vec1.Y * vec2.Y + vec1.Z * vec2.Z)

    End Function

    Public Function Normalise(ByVal Xval, ByVal Yval, ByVal Zval) As Vecteur

        Dim norm, val As Double
        norm = Norme(Xval, Yval, Zval)

        val = CDbl(IIf(norm < 0.0000000001, 1.0, 1.0 / norm))

        Return New Vecteur(Xval * val, Yval * val, Zval * val)

    End Function

    'DATES/TIME FUNCTIONS

    Function CheckEpoch() As Boolean
        CheckEpoch = True
        If Home.EPOCHBox.Text = "" OrElse Home.EPOCHBox.Text = Nothing Then CheckEpoch = False

    End Function

    Function MJDtoGreg(ByRef epoch As String) As String
        'Converti une date du format MJD vers le format Gregorian JJ/MM/AAAA hh:mm:ss,fff
        Dim Z, F, a, S1, S2, B, C, D, E, Q1, Q2, JJ, MM, AAAA, HH, MM1, MM2, MM3, SS1, SS2, SS3, EpochDate, EpochTime
        epoch = Val(epoch)

        Z = Truncate(epoch + 2430000)
        F = epoch + 2430000 - Z + 0.5
        a = Truncate((Z - 1867216.25) / 36524.25)
        Dim a1 = ((Z - 1867216.25) \ 36524.25)
        S1 = Truncate((a / 4))
        S2 = Z + 1 + a - S1
        B = S2 + 1524
        C = Truncate((B - 122.1) / 365.25)
        D = Truncate(365.25 * C)
        E = Truncate((B - D) / 30.6001)
        Q1 = Truncate(30.6001 * E)
        Q2 = B - D - Q1 + F
        JJ = Truncate(Q2)
        If JJ > 31 Then
            JJ = JJ - 31
            E = E + 1
        End If
        If E < 14 Then
            MM = E - 1
        Else
            If E > 14 Or E = 14 Then
                MM = E - 13
            End If
        End If
        If MM > 2 Then
            AAAA = C - 4716
        Else
            If MM = 1 Or MM = 2 Then
                AAAA = C - 4715
            End If
        End If

        'Vérif. année bisextile
        Dim i As Integer
        Dim BisYear As Boolean
        BisYear = False

        For i = 1912 To 2112 Step 4
            If AAAA = i Then BisYear = True
        Next

        'Prise en compte des mois à 30 jours
        If JJ = "31" And MM = "02" Then
            JJ = "01"
            MM = "03"
        ElseIf JJ = "31" And MM = "04" Then
            JJ = "01"
            MM = "05"
        ElseIf JJ = "31" And MM = "06" Then
            JJ = "01"
            MM = "07"
        ElseIf JJ = "31" And MM = "09" Then
            JJ = "01"
            MM = "10"
        ElseIf JJ = "31" And MM = "11" Then
            JJ = "01"
            MM = "12"
        End If

        EpochDate = CDate(JJ & DS & MM & DS & AAAA).ToString(FMT_dd_MM_yyyy)

        HH = Truncate((Q2 Mod 1) * 24)

        MM1 = ((Q2 Mod 1) * 24)
        MM2 = MM1 Mod 1
        MM3 = Truncate(MM2 * 60)

        SS1 = (Q2 Mod 1) * 24
        SS2 = (SS1 Mod 1) * 60
        SS3 = Round((SS2 Mod 1) * 60, 3)
        EpochTime = CDate(HH & DS & MM3 & DS & SS3).ToString(FMT_HH_mm_ss_fff)

        epoch = EpochDate & " " & EpochTime

        Return epoch
    End Function

    Function MJDtoGreg2(ByVal epoch As String)
        'Converti une date du format MJD vers le format Gregorian JJ/MM/AAAA hh:mm:ss,fff
        Dim Z, F, a, S1, S2, B, C, D, E, Q1, Q2, JJ, MM, AAAA, HH, MM1, MM2, MM3, SS1, SS2, SS3, EpochDate, EpochTime
        epoch = Val(epoch)

        Z = Truncate(epoch + 2430000)
        F = epoch + 2430000 - Z + 0.5
        a = Truncate((Z - 1867216.25) / 36524.25)
        S1 = Truncate((a / 4))
        S2 = Z + 1 + a - S1
        B = S2 + 1524
        C = Truncate((B - 122.1) / 365.25)
        D = Truncate(365.25 * C)
        E = Truncate((B - D) / 30.6001)
        Q1 = Truncate(30.6001 * E)
        Q2 = B - D - Q1 + F
        JJ = Truncate(Q2)
        If JJ > 31 Then
            JJ = JJ - 31
            E = E + 1
        End If
        If E < 14 Then
            MM = E - 1
        Else
            If E > 14 Or E = 14 Then
                MM = E - 13
            End If
        End If
        If MM > 2 Then
            AAAA = C - 4716
        Else
            If MM = 1 Or MM = 2 Then
                AAAA = C - 4715
            End If
        End If

        'Vérif. année bisextile
        Dim i As Integer
        Dim BisYear As Boolean
        BisYear = False

        For i = 1912 To 2112 Step 4
            If AAAA = i Then BisYear = True
        Next

        'Prise en compte des mois à 30 jours
        'Prise en compte des mois à 30 jours
        If JJ = "31" And MM = "02" Then
            JJ = "01"
            MM = "03"
        ElseIf JJ = "31" And MM = "04" Then
            JJ = "01"
            MM = "05"
        ElseIf JJ = "31" And MM = "06" Then
            JJ = "01"
            MM = "07"
        ElseIf JJ = "31" And MM = "09" Then
            JJ = "01"
            MM = "10"
        ElseIf JJ = "31" And MM = "11" Then
            JJ = "01"
            MM = "12"
        End If

        EpochDate = CDate(JJ & DS & MM & DS & AAAA).ToString(FMT_dd_MM_yyyy)

        HH = Truncate((Q2 Mod 1) * 24)

        MM1 = ((Q2 Mod 1) * 24)
        MM2 = MM1 Mod 1
        MM3 = Truncate(MM2 * 60)

        SS1 = (Q2 Mod 1) * 24
        SS2 = (SS1 Mod 1) * 60
        SS3 = Round((SS2 Mod 1) * 60, 3)
        If SS3 = 60 Then SS3 = 59.999

        EpochTime = CDate(HH & TS & MM3 & TS & SS3).ToString(FMT_HH_mm_ss_fff)

        epoch = EpochDate & " " & EpochTime

        Return epoch
    End Function

    Function GregtoMJD(ByRef epoch As String)
        'Converti une date du format Gregorian (JJ/MM/AAAA hh:mm:ss,fff) vers le format MJD

        Dim ChaineEpoch, EpochDate, EpochTime, ChaineDate, ChaineTime, JJ, MM, AAAA, HH, MN, SS

        ChaineEpoch = epoch.Split(" ")
        EpochDate = ChaineEpoch(0)
        EpochTime = ChaineEpoch(1)

        ChaineDate = EpochDate.split(DS)
        JJ = Val(ChaineDate(0))
        MM = Val(ChaineDate(1))
        AAAA = Val(ChaineDate(2))

        ChaineTime = EpochTime.split(TS)
        HH = Val(ChaineTime(0))
        MN = Val(ChaineTime(1))
        SS = Val(ChaineTime(2))

        epoch = (367 * AAAA - Int((7 * (AAAA + Int((MM + 9) / 12))) / 4) + Int((275 * MM) / 9) + JJ + 1721013.5) - 2430000 + (((((SS / 60) + MN) / 60) + HH) / 24)
        Return epoch

    End Function

    Function GregtoMJD2(ByVal epoch As String)
        'Converti une date du format Gregorian (JJ/MM/AAAA hh:mm:ss,fff) vers le format MJD

        Dim ChaineEpoch, EpochDate, EpochTime, ChaineDate, ChaineTime, JJ, MM, AAAA, HH, MN, SS

        ChaineEpoch = epoch.Split(" ")
        EpochDate = ChaineEpoch(0)
        EpochTime = ChaineEpoch(1)

        ChaineDate = EpochDate.split(DS)
        JJ = Val(ChaineDate(0))
        MM = Val(ChaineDate(1))
        AAAA = Val(ChaineDate(2))

        ChaineTime = EpochTime.split(TS)
        HH = Val(ChaineTime(0))
        MN = Val(ChaineTime(1))
        SS = Val(ChaineTime(2))

        epoch = (367 * AAAA - Int((7 * (AAAA + Int((MM + 9) / 12))) / 4) + Int((275 * MM) / 9) + JJ + 1721013.5) - 2430000 + (((((SS / 60) + MN) / 60) + HH) / 24)
        Return epoch

    End Function

    Function GregtoTime(ByVal epoch As String)
        'Récupère la partie Heure (hh:mm:ss,fff) d'un date au format Gregorian (JJ/MM/AAAA hh:mm:ss,fff)

        Dim ChaineEpoch, ChaineDate

        ChaineEpoch = epoch.Split(" ")
        ChaineDate = ChaineEpoch(0)
        epoch = ChaineEpoch(1)

        Return epoch

    End Function

    Function GregtoYear(ByVal epoch As String) As Integer
        'Récupère l'année (AAAA) d'un date au format Gregorian (JJ/MM/AAAA hh:mm:ss,fff)

        Dim ChaineEpoch, ChaineDate1, ChaineDate2, ChaineTime, JJ, MM, AAAA

        ChaineEpoch = epoch.Split(" ")
        ChaineDate1 = ChaineEpoch(0)
        ChaineTime = ChaineEpoch(1)

        ChaineDate2 = ChaineDate1.split(DS)
        JJ = ChaineDate2(0)
        MM = ChaineDate2(1)
        GregtoYear = ChaineDate2(2)

    End Function

    Function HMStoH(ByVal Time As String)
        'Converti une heure (hh:mm:ss.fff) en Heure décimale

        Dim ChaineTime, HH, MM, SS

        ChaineTime = Time.Split(TS)
        HH = CInt(ChaineTime(0))
        MM = CInt(ChaineTime(1))
        SS = CInt(ChaineTime(2))

        Time = (HH + MM / 60 + SS / 3600)
        Return Time

    End Function

    Function HtoHMS(ByVal Time As String)
        'Converti une heure décimale en Heure Min Sec (hh:mm:ss.fff) 

        Dim ChaineTime, HH, MM, SS

        HH = Int(Time)
        MM = Int((Time - HH) * 60)
        SS = ((((Time - HH) * 60) - MM) * 60)

        ChaineTime = CDate(HH & TS & MM & TS & SS).ToString(FMT_HH_mm_ss_fff)
        Time = ChaineTime
        Return Time

    End Function

    Function MJDGGEDate(ByVal MJDate As Double) As String

        'Fonction qui converti la date MJD en format Google Earth (JJ MM AAAA)
        Dim Z, F, a, S1, S2, B, C, D, E, Q1, Q2, JJ, MM, AAAA, HH, MM1, MM2, MM3, SS1, SS2, SS3, EpochDate

        Z = Truncate(MJDate + 2430000)
        F = MJDate + 2430000 - Z + 0.5
        a = Truncate((Z - 1867216.25) / 36524.25)
        S1 = Truncate((a / 4))
        S2 = Z + 1 + a - S1
        B = S2 + 1524
        C = Truncate((B - 122.1) / 365.25)
        D = Truncate(365.25 * C)
        E = Truncate((B - D) / 30.6001)
        Q1 = Truncate(30.6001 * E)
        Q2 = B - D - Q1 + F
        JJ = Truncate(Q2)
        If JJ > 31 Then
            JJ = JJ - 31
            E = E + 1
        End If
        If E < 14 Then
            MM = E - 1
        Else
            If E > 14 Or E = 14 Then
                MM = E - 13
            End If
        End If
        If MM > 2 Then
            AAAA = C - 4716
        Else
            If MM = 1 Or MM = 2 Then
                AAAA = C - 4715
            End If
        End If

        'Vérif. année bisextile
        Dim i As Integer
        Dim BisYear As Boolean
        BisYear = False

        For i = 1912 To 2112 Step 4
            If AAAA = i Then BisYear = True
        Next

        'Prise en compte des mois à 30 jours
        If JJ = "31" And MM = "02" Then
            JJ = "01"
            MM = "03"
        ElseIf JJ = "31" And MM = "04" Then
            JJ = "01"
            MM = "05"
        ElseIf JJ = "31" And MM = "06" Then
            JJ = "01"
            MM = "07"
        ElseIf JJ = "31" And MM = "09" Then
            JJ = "01"
            MM = "10"
        ElseIf JJ = "31" And MM = "11" Then
            JJ = "01"
            MM = "12"
        End If

        EpochDate = CDate(AAAA & "-" & MM & "-" & JJ).ToString("yyyy-MM-dd")

        MJDGGEDate = EpochDate

        Return MJDGGEDate

    End Function

    Function MJDGGEHour(ByVal MJDate As Double) As String

        'Fonction qui converti la date MJD en format Google Earth (HH MM SS)
        Dim Z, F, a, S1, S2, B, C, D, E, Q1, Q2, JJ, MM, AAAA, HH, MM1, MM2, MM3, SS1, SS2, SS3, EpochTime

        Z = Truncate(MJDate + 2430000)
        F = MJDate + 2430000 - Z + 0.5
        a = Truncate((Z - 1867216.25) / 36524.25)
        S1 = Truncate((a / 4))
        S2 = Z + 1 + a - S1
        B = S2 + 1524
        C = Truncate((B - 122.1) / 365.25)
        D = Truncate(365.25 * C)
        E = Truncate((B - D) / 30.6001)
        Q1 = Truncate(30.6001 * E)
        Q2 = B - D - Q1 + F

        HH = Truncate((Q2 Mod 1) * 24)

        MM1 = ((Q2 Mod 1) * 24)
        MM2 = MM1 Mod 1
        MM3 = Truncate(MM2 * 60)

        SS1 = (Q2 Mod 1) * 24
        SS2 = (SS1 Mod 1) * 60
        SS3 = Round((SS2 Mod 1) * 60, 0)
        If SS3 = 60 Then SS3 = 59
        EpochTime = CDate(HH & TS & MM3 & TS & SS3).ToString(FMT_HH_mm_ss)

        MJDGGEHour = EpochTime

        Return MJDGGEHour

    End Function

    Function CheckCaracts(ByRef StringToCheck As String)

        Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11

        n1 = InStr(1, StringToCheck, "/")
        If n1 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n1, 1) = " "
        End If

        n2 = InStr(1, StringToCheck, "\")
        If n2 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n2, 1) = " "
        End If

        n3 = InStr(1, StringToCheck, ":")
        If n3 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n3, 1) = " "
        End If

        n4 = InStr(1, StringToCheck, "*")
        If n4 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n4, 1) = " "
        End If

        n5 = InStr(1, StringToCheck, "?")
        If n5 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n5, 1) = " "
        End If

        n6 = InStr(1, StringToCheck, "<")
        If n6 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n6, 1) = " "
        End If

        n7 = InStr(1, StringToCheck, ">")
        If n7 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n7, 1) = " "
        End If

        n8 = InStr(1, StringToCheck, "|")
        If n8 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n8, 1) = " "
        End If

        n9 = InStr(1, StringToCheck, "!")
        If n9 = 0 Then
            StringToCheck = StringToCheck
        Else
            Mid(StringToCheck, n9, 1) = " "
        End If

        'n10 = InStr(1, StringToCheck, "[")
        'If n10 = 0 Then
        '    StringToCheck = StringToCheck
        'Else
        '    Mid(StringToCheck, n10, 1) = " "
        'End If

        'n11 = InStr(1, StringToCheck, "]")
        'If n11 = 0 Then
        '    StringToCheck = StringToCheck
        'Else
        '    Mid(StringToCheck, n11, 1) = " "
        'End If

        Return StringToCheck

    End Function

    Function CurrentDateToGreg()

        Dim AA, MM, DD, HH, MI, SS

        Dim localZone As TimeZone = TimeZone.CurrentTimeZone
        Dim currentDate As DateTime = DateTime.Now
        Dim currentUTC As DateTime = localZone.ToUniversalTime(currentDate)

        AA = currentUTC.Year
        MM = currentUTC.Month
        DD = currentUTC.Day
        HH = currentUTC.Hour
        MI = currentUTC.Minute
        SS = Round(currentUTC.Second, 0)

        CurrentDateToGreg = DD & DS & MM & DS & AA & " " & HH & TS & MI & TS & SS

    End Function

    'FUNCTIONS

    Sub KepleriantoCartesian(ByVal SMA As Double, ByVal ECC As Double, ByVal INC As Double, ByVal RAAN As Double, ByVal AOP As Double, ByVal TA As Double)

        On Error Resume Next

        Dim SemiLatus = Val(SMA * (1 - ECC ^ 2))
        Dim Radius = Val(SemiLatus / (1 + (ECC * Cos(TA * PI / 180))))

        Home.X = Radius * (Cos(Maths.DEG2RAD * (AOP + TA)) * Cos(Maths.DEG2RAD * (RAAN)) - Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (AOP + TA)) * Sin(Maths.DEG2RAD * (RAAN)))
        Home.Y = Radius * (Cos(Maths.DEG2RAD * (AOP + TA)) * Sin(Maths.DEG2RAD * (RAAN)) + Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (AOP + TA)) * Cos(Maths.DEG2RAD * (RAAN)))
        Home.Z = Radius * (Sin(Maths.DEG2RAD * (AOP + TA)) * Sin(Maths.DEG2RAD * (INC)))

        Home.VX = Sqrt(Mu / SemiLatus) * (Cos(Maths.DEG2RAD * (TA)) + ECC) * (-Sin(Maths.DEG2RAD * (AOP)) * Cos(Maths.DEG2RAD * (RAAN)) - Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (RAAN)) * Cos(Maths.DEG2RAD * (AOP))) - Sqrt(Mu / SemiLatus) * (Sin(Maths.DEG2RAD * (TA))) * (Cos(Maths.DEG2RAD * (AOP)) * Cos(Maths.DEG2RAD * (RAAN)) - Cos(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (RAAN)) * Sin(Maths.DEG2RAD * (AOP)))
        Home.VY = Sqrt(Mu / SemiLatus) * (Cos(Maths.DEG2RAD * (TA)) + ECC) * (-Sin(Maths.DEG2RAD * (AOP)) * Sin(Maths.DEG2RAD * (RAAN)) + Cos(Maths.DEG2RAD * (INC)) * Cos(Maths.DEG2RAD * (RAAN)) * Cos(Maths.DEG2RAD * (AOP))) - Sqrt(Mu / SemiLatus) * (Sin(Maths.DEG2RAD * (TA))) * (Cos(Maths.DEG2RAD * (AOP)) * Sin(Maths.DEG2RAD * (RAAN)) + Cos(Maths.DEG2RAD * (INC)) * Cos(Maths.DEG2RAD * (RAAN)) * Sin(Maths.DEG2RAD * (AOP)))
        Home.VZ = Sqrt(Mu / SemiLatus) * ((Cos(Maths.DEG2RAD * (TA)) + ECC) * Sin(Maths.DEG2RAD * (INC)) * Cos(Maths.DEG2RAD * (AOP)) - Sin(Maths.DEG2RAD * (TA)) * Sin(Maths.DEG2RAD * (INC)) * Sin(Maths.DEG2RAD * (AOP)))

    End Sub

    Sub CartesianToKeplerian(ByVal Xval As Double, ByVal Yval As Double, ByVal Zval As Double, ByVal VXval As Double, ByVal VYval As Double, ByVal VZval As Double)
        Dim en, Hnorm, Rnorm, NSE, ESR, Nnorm, RSV As Double
        Dim H, VH, VHSMU, RSNR, E, K, N As Vecteur
        Dim R As New Vecteur(Home.X, Home.Y, Home.Z)
        Dim V As New Vecteur(Home.VX, Home.VY, Home.VZ)

        'ENERGY
        en = Norme(VXval, VYval, VZval) ^ 2 / 2 - Mu / Norme(Xval, Yval, Zval)

        'SMA
        Home.SMA = -Mu / 2 / en

        'ECC
        Rnorm = Norme(R.X, R.Y, R.Z)
        H = ProduitVectoriel(R, V)
        Hnorm = Norme(H.X, H.Y, H.Z)
        VH = ProduitVectoriel(V, H)
        VHSMU = New Vecteur(VH.X / Mu, VH.Y / Mu, VH.Z / Mu)
        RSNR = New Vecteur(R.X / Rnorm, R.Y / Rnorm, R.Z / Rnorm)
        E = New Vecteur(VHSMU.X - RSNR.X, VHSMU.Y - RSNR.Y, VHSMU.Z - RSNR.Z)
        Home.ECC = Norme(E.X, E.Y, E.Z)

        'INC
        Home.INC = Maths.RAD2DEG * (Acos(H.Z / Hnorm))

        'RAAN
        K = New Vecteur(0, 0, 1)
        N = ProduitVectoriel(K, H)
        ESR = ProduitScalaire(E, R)
        Nnorm = Norme(N.X, N.Y, N.Z)
        Home.RAAN = Maths.RAD2DEG * (Acos(N.X / Nnorm))
        If N.Y < 0 Then
            Home.RAAN = 360 - Home.RAAN
        End If

        'AOP
        NSE = ProduitScalaire(E, N)
        Home.AOP = Maths.RAD2DEG * (Acos(NSE / (Nnorm * Home.ECC)))
        If E.Z < 0 Then Home.AOP = 360 - Home.AOP

        'TA, EA, MA
        Home.TA = Acos(ESR / (Home.ECC * Rnorm))
        RSV = ProduitScalaire(R, V)
        If RSV < 0 Then Home.TA = 2 * PI - Home.TA
        Home.EA = Acos((Home.ECC + Cos(Home.TA)) / (1 + Home.ECC * Cos(Home.TA)))
        If Home.TA > PI Then Home.EA = 2 * PI - Home.EA
        Home.MA = Maths.RAD2DEG * (Home.EA - Home.ECC * Sin(Home.EA))

    End Sub

    Sub CartesianToKeplerianChart(ByVal Xval As Double, ByVal Yval As Double, ByVal Zval As Double, ByVal VXval As Double, ByVal VYval As Double, ByVal VZval As Double)
        Dim en, Hnorm, Rnorm, NSE, ESR, Nnorm, RSV As Double
        Dim H, VH, VHSMU, RSNR, E, K, N As Vecteur
        Dim R As New Vecteur(Home.XC, Home.YC, Home.ZC)
        Dim V As New Vecteur(Home.VXC, Home.VYC, Home.VZC)

        'ENERGY
        en = Norme(VXval, VYval, VZval) ^ 2 / 2 - Mu / Norme(Xval, Yval, Zval)

        'SMA
        Home.SMAC = -Mu / 2 / en

        'ECC
        Rnorm = Norme(R.X, R.Y, R.Z)
        H = ProduitVectoriel(R, V)
        Hnorm = Norme(H.X, H.Y, H.Z)
        VH = ProduitVectoriel(V, H)
        VHSMU = New Vecteur(VH.X / Mu, VH.Y / Mu, VH.Z / Mu)
        RSNR = New Vecteur(R.X / Rnorm, R.Y / Rnorm, R.Z / Rnorm)
        E = New Vecteur(VHSMU.X - RSNR.X, VHSMU.Y - RSNR.Y, VHSMU.Z - RSNR.Z)
        Home.ECCC = Norme(E.X, E.Y, E.Z)

        'INC
        Home.INCC = Maths.RAD2DEG * (Acos(H.Z / Hnorm))

        'RAAN
        K = New Vecteur(0, 0, 1)
        N = ProduitVectoriel(K, H)
        ESR = ProduitScalaire(E, R)
        Nnorm = Norme(N.X, N.Y, N.Z)
        Home.RAANC = Maths.RAD2DEG * (Acos(N.X / Nnorm))
        If N.Y < 0 Then
            Home.RAANC = 360 - Home.RAANC
        End If

        'AOP
        NSE = ProduitScalaire(E, N)
        Home.AOPC = Maths.RAD2DEG * (Acos(NSE / (Nnorm * Home.ECCC)))
        If E.Z < 0 Then Home.AOPC = 360 - Home.AOPC

    End Sub

    Sub AdpatedParameters()

        'Adpapted Parameters Option
        Dim wprime As Double

        If Home.ECC < 0.1 AndAlso Home.INC < 1 AndAlso Home.INC > -1 Then
            'Circular & Equatorial

            Home.CircularPanel.Visible = False
            Home.EquPanel.Visible = False
            Home.CircEquPanel.Visible = True

            'Home.AP_GroupBox.Text = "Adapted Parameters (Geostationnary)"

            'Parameters
            wprime = Round4((Home.AOP + Home.RAAN) Mod 360)
            Home.SMA_AP_ECCINC.Text = Round4(Home.SMA)
            Home.ex_ECCINC.Text = Round4(Home.ECC * Cos(Maths.DEG2RAD * (wprime)))
            Home.ey_ECCINC.Text = Round4(Home.ECC * Sin(Maths.DEG2RAD * (wprime)))
            Home.ix_ECCINC.Text = Round4(Home.INC * Cos(Maths.DEG2RAD * (Home.RAAN)))
            Home.iy_ECCINC.Text = Round4(Home.INC * Sin(Maths.DEG2RAD * (Home.RAAN)))
            Home.MeanL_ECCINC.Text = Round4((Home.AOP + Home.RAAN + Home.MA - Home.GST) Mod 360)

        ElseIf Home.ECC < 0.1 Then
            'Circular

            Home.CircEquPanel.Visible = False
            Home.EquPanel.Visible = False
            Home.CircularPanel.Visible = True

            'Home.AP_GroupBox.Text = "Adapted Parameters (Cirular Orbit)"

            'Parameters
            Home.SMA_AP_ECC.Text = Round4(Home.SMA)
            Home.ex.Text = Round4(Home.ECC * Cos(Maths.DEG2RAD * (Home.AOP)))
            Home.ey.Text = Round4(Home.ECC * Sin(Maths.DEG2RAD * (Home.AOP)))
            Home.INC_AP_ECC.Text = Round4(Home.INC)
            Home.RAAN_AP_ECC.Text = Round4(Home.RAAN)
            Home.AlphaPrime.Text = Round4(Home.AOL)

        ElseIf Home.INC < 1 AndAlso Home.INC > -1 Then
            'Equatorial

            Home.CircularPanel.Visible = False
            Home.CircEquPanel.Visible = False
            Home.EquPanel.Visible = True

            'Home.AP_GroupBox.Text = "Adapted Parameters (Equatorial Orbit)"

            'Parameters
            Home.SMA_AP_INC.Text = Round4(Home.SMA)
            Home.ECC_AP_INC.Text = Round4(Home.ECC)
            Home.AOP_AP_INC.Text = Round4((Home.AOP + Home.RAAN) Mod 360)
            Home.ix.Text = Round4(Home.INC * Cos(Maths.DEG2RAD * (Home.RAAN)))
            Home.iy.Text = Round4(Home.INC * Sin(Maths.DEG2RAD * (Home.RAAN)))
            Home.MA_AP.Text = Round4(Home.MA)

        Else

            Home.CircularPanel.Visible = False
            Home.CircEquPanel.Visible = False
            Home.EquPanel.Visible = False

            Home.AP_GroupBox.Text = "Adapted Parameters"

        End If

    End Sub

    Sub LATLONG(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal GST As Double)
        'DETERMINE LA LATITUDE ET LONGITUDE DU SATELLITE - METHODE PREVISAT

        Dim r, Lat0, e2, Ls, phi, ct, sph As Double

        ' Latitude
        r = Sqrt(X * X + Y * Y)
        Lat0 = Atan(Z / r)
        phi = 7.0
        While Abs(Lat0 - phi) > 0.0000001
            phi = Lat0
            sph = Sin(phi)
            ct = 1.0 / Sqrt(1.0 - Terre.E2 * sph * sph)
            Lat0 = Atan((Z + EarthEquRad * ct * Terre.E2 * sph) / r)
        End While
        Home.LAT = Maths.RAD2DEG * (Lat0)

        ' Longitude 
        Ls = Maths.RAD2DEG * ((Atan2(Y, X) - (Maths.DEG2RAD * GST)) Mod Maths.DEUX_PI)
        If Ls > 180 Then
            Home.LONGI = Ls - 360
        ElseIf Ls < -180 Then
            Home.LONGI = 360 + Ls
        Else : Home.LONGI = Ls
        End If

    End Sub

    Function LAN_F(ByVal ECC As Double, ByVal MA As Double, ByVal AnoPeriod As Double, _
                         ByVal RAAN As Double, ByVal ArgP As Double, ByVal DeltaD As Double, ByVal Epoch0 As Double)
        'Longitude of Ascending Node

        'Anomaly Delta
        Dim DeltaM As Double
        Dim AOP = Maths.DEG2RAD * (-ArgP)
        DeltaM = (MA - Maths.RAD2DEG * (2 * Atan(Sqrt((1 - ECC) / (1 + ECC)) * Tan(AOP / 2)) - (ECC * Sqrt(1 - ECC ^ 2) * Sin(AOP)) / (1 + ECC * Cos(AOP)))) Mod 360

        'Time Delta
        Dim DeltaT As Double
        DeltaT = (DeltaM / 360) * (AnoPeriod * 60)

        'Delta RAAN
        Dim Du = Epoch0 - DJ2000
        Dim Tu = Du / 36525
        Dim qG00 = (24110.54841 + (((86400.0 * 36525.0) / 365.2421897) * Tu) + (0.093104 * (Tu ^ 2)) - (0.0000062 * (Tu ^ 3))) Mod 86400
        Dim qGt = (qG00 + 86400 * EarthNodPrec * DeltaD) Mod 86400
        Dim OhmGt = qGt / 240

        'Longitude
        Dim LAN0 = RAAN - OhmGt
        Dim DeltaLAN = -DeltaT * 360 / 86400
        LAN_F = LAN0 - DeltaLAN

    End Function

    Function Saison(ByVal YY) As Double
        'Trouve la date de l'Equinoxe Vernal en fonction de l'année
        Dim k, n As Integer
        Dim dk, T, D, TETUJ
        Dim JJD As Double

        k = YY - 2000 - 1

        For n = 0 To 4

            dk = k + 0.25 * n

            T = 0.21451814 + 0.99997862442 * dk _
            + 0.00642125 * Sin(1.580244 + 0.0001621008 * dk) + 0.0031065 * Sin(4.143931 + 6.2829005032 * dk) _
            + 0.00190024 * Sin(5.604775 + 6.2829478479 * dk) + 0.00178801 * Sin(3.987335 + 6.2828291282 * dk) _
            + 0.00004981 * Sin(1.507976 + 6.283109952 * dk) + 0.00006264 * Sin(5.723365 + 6.283062603 * dk) _
            + 0.00006262 * Sin(5.702396 + 6.2827383999 * dk) + 0.00003833 * Sin(7.166906 + 6.2827857489 * dk) _
            + 0.00003616 * Sin(5.58175 + 6.2829912245 * dk) + 0.00003597 * Sin(5.591081 + 6.2826670315 * dk) _
            + 0.00003744 * Sin(4.3918 + 12.5657883 * dk) + 0.00001827 * Sin(8.3129 + 12.56582984 * dk) _
            + 0.00003482 * Sin(8.1219 + 12.56572963 * dk) - 0.00001327 * Sin(-2.1076 + 0.33756278 * dk) _
            - 0.00000557 * Sin(5.549 + 5.753262 * dk) + 0.00000537 * Sin(1.255 + 0.003393 * dk) _
            + 0.00000486 * Sin(19.268 + 77.7121103 * dk) - 0.00000426 * Sin(7.675 + 7.8602511 * dk) _
            - 0.00000385 * Sin(2.911 + 0.0005412 * dk) - 0.00000372 * Sin(2.266 + 3.9301258 * dk) _
            - 0.0000021 * Sin(4.785 + 11.5065238 * dk) + 0.0000019 * Sin(6.158 + 1.5774 * dk) _
            + 0.00000204 * Sin(0.582 + 0.5296557 * dk) _
            - 0.00000157 * Sin(1.782 + 5.8848012 * dk) + 0.00000137 * Sin(-4.265 + 0.3980615 * dk) _
            - 0.00000124 * Sin(3.871 + 5.2236573 * dk) _
            + 0.00000119 * Sin(2.145 + 5.5075293 * dk) + 0.00000144 * Sin(0.476 + 0.0261074 * dk) _
            + 0.00000038 * Sin(6.45 + 18.848689 * dk) + 0.00000078 * Sin(2.8 + 0.775638 * dk) _
            - 0.00000051 * Sin(3.67 + 11.790375 * dk) + 0.00000045 * Sin(-5.79 + 0.796122 * dk) _
            + 0.00000024 * Sin(5.61 + 0.213214 * dk) + 0.00000043 * Sin(7.39 + 10.976868 * dk) - 0.00000038 * Sin(3.1 + 5.486739 * dk) _
            - 0.00000033 * Sin(0.64 + 2.544339 * dk) _
            + 0.00000033 * Sin(-4.78 + 5.573024 * dk) - 0.00000032 * Sin(5.33 + 6.069644 * dk) - 0.00000021 * Sin(2.65 + 0.020781 * dk) _
            - 0.00000021 * Sin(5.61 + 2.9424 * dk) + 0.00000019 * Sin(-0.93 + 0.000799 * dk) - 0.00000016 * Sin(3.22 + 4.694014 * dk) + 0.00000016 * Sin(-3.59 + 0.006829 * dk) _
            - 0.00000016 * Sin(1.96 + 2.146279 * dk) - 0.00000016 * Sin(5.92 + 15.720504 * dk) + 0.00000115 * Sin(23.671 + 83.9950108 * dk) + 0.00000115 * Sin(17.845 + 71.4292098 * dk)

        Next

        JJD = 2451545 + T * 365.25

        D = YY / 100
        TETUJ = (32.23 * (D - 18.3) * (D - 18.3) - 15) / 86400
        JJD = JJD - TETUJ     'moins TE-TU avant conversion en date.
        Saison = JJD + 0.0003472222 'ajout de 30s pour arrondi sur la minute avant troncature lors de l'affichage */

    End Function

    Function GSTCalc(ByVal Epoch) As Double
        'Permet de calculer l'angle entre le Méridien de Greenwich et l'Axe Vernal

        Dim Du, Tu, QG0, DJ, QGT, RAGT As Double

        Epoch = Epoch + 2430000

        If Epoch >= (Int(Epoch) + 0.5) Then
            Du = Int(Epoch) + 0.5 - JDREF
        ElseIf Epoch < (Int(Epoch) + 0.5) Then
            Du = Int(Epoch) - 0.5 - JDREF
        End If

        Tu = Du / 36525
        QG0 = (24110.54841 + (86400.0 * 36525 / 365.2421897 * Tu) + (0.093104 * (Tu ^ 2)) - (0.0000062 * (Tu ^ 3))) Mod 86400

        If Epoch < (Int(Epoch) + 0.5) Then
            DJ = Epoch - (Int(Epoch) - 0.5)
        ElseIf Epoch > (Int(Epoch) + 0.5) Then
            DJ = Epoch - (Int(Epoch) + 0.5)
        End If

        QGT = (QG0 + (86400 * DJ * EarthNodPrec)) Mod 86400

        GSTCalc = (QGT / 240) Mod 360


    End Function

    Function LTAN_F()

        'Fonction de calcul du Local Time of Ascending Node
        'Gestion des options Epoch, RAAN et Local Sid.Time
        Dim EquinoxRef As Double ' Equinoxe de printemps de l'année considérée
        Dim Year
        Dim VarEpoch = 0.986301369863014 'variation angle/jour = 360/365
        Dim RASUN
        Dim AngleH
        Dim HHL
        Dim LST As Double
        Dim JDREF = 2451545 'JD de référence 01/01/2000 à 12h

        Year = GregtoYear(MJDtoGreg2(Home.EPOCH)) 'récupère l'année en cours
        EquinoxRef = Saison(Year) - 2430000

        If Home.EPOCH > EquinoxRef Then
            RASUN = (((Home.EPOCH - EquinoxRef) * VarEpoch) - (2 * VarEpoch)) Mod 360
        Else
            RASUN = 360 + ((((Home.EPOCH - EquinoxRef) * VarEpoch) - (2 * VarEpoch)) Mod 360)
        End If

        'Détermination de LTAN
        'Cacul de l'angle horaire
        If (Home.RAAN - RASUN) < 0 Then
            AngleH = 360 + (Home.RAAN - RASUN)
        Else
            AngleH = (Home.RAAN - RASUN)
        End If
        If AngleH = Nothing Then AngleH = 0
        HHL = (12 + (AngleH / 15)) Mod 24
        LTAN_F = HtoHMS(HHL)

    End Function

    Sub CreateTLEFolder()
        Dim TLE, TLE_, TLER

        'On créer le dossier TLE si besoin
        If My.Computer.FileSystem.DirectoryExists("C:\TLEAnalyser") = False Then My.Computer.FileSystem.CreateDirectory("C:\TLEAnalyser")
        If My.Computer.FileSystem.DirectoryExists("C:\TLEAnalyser\TLE") = False Then My.Computer.FileSystem.CreateDirectory("C:\TLEAnalyser\TLE")

        'On récupère la liste des TLE dans le fichier TLEListe de Resources
        If My.Computer.FileSystem.FileExists("C:\TLEAnalyser\TLE\TLEList.txt") = True Then My.Computer.FileSystem.DeleteFile("C:\TLEAnalyser\TLE\TLEList.txt")
        File.AppendAllText("C:\TLEAnalyser\TLE\TLEList.txt", My.Resources.TLEList)
        FileOpen(1, "C:\TLEAnalyser\TLE\TLEList.txt", OpenMode.Input)
        While Not EOF(1)
            TLE = LineInput(1)
            TLE_ = TLE.Replace("-", "_")
            TLER = CType(My.Resources.ResourceManager.GetObject(TLE_), String)
            'On colle la TLE de base si besoin dans le dossier TLE
            If My.Computer.FileSystem.FileExists(Home.TLEPath & TLE & ".txt") = False Then _
            File.WriteAllText(Home.TLEPath & TLE & ".txt", TLER)
        End While
        FileClose(1)

    End Sub

    Sub CreateFolders()

        'GMAT Folder
        If My.Computer.FileSystem.DirectoryExists(Home.GMATPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.GMATPath)
        'PLOT Folder
        If My.Computer.FileSystem.DirectoryExists(Home.PlotPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.PlotPath)
        'FAVOURITES Folder
        If My.Computer.FileSystem.DirectoryExists(Home.FavPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.FavPath)
        If My.Computer.FileSystem.FileExists(Home.FavPath & "\favourites.txt") = True Then My.Computer.FileSystem.RenameFile(Home.FavPath & "\favourites.txt", "Favorites.txt")
        'CELESTIA Folder
        If My.Computer.FileSystem.DirectoryExists(Home.CELESTIAPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.CELESTIAPath)
        'GOOGLE EARTH Folder
        If My.Computer.FileSystem.DirectoryExists(Home.GOOGLEPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.GOOGLEPath)
        If My.Computer.FileSystem.FileExists(Home.GOOGLEPath & "GoogleEarth.htm") = False Then File.AppendAllText(Home.GOOGLEPath & "GoogleEarth.htm", My.Resources.GoogleEarth)
        'GOOGLE Map Folder
        If My.Computer.FileSystem.DirectoryExists(Home.GOOGLEMAPPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.GOOGLEMAPPath)
        If My.Computer.FileSystem.FileExists(Home.GOOGLEMAPPath & "GoogleMap.htm") = False Then File.AppendAllText(Home.GOOGLEMAPPath & "GoogleMap.htm", My.Resources.GoogleMap)
        'Init
        If My.Computer.FileSystem.FileExists(Home.AppPath & "tlea.ini") = False Then File.AppendAllText(Home.AppPath & "tlea.ini", My.Resources.tlea)

        ''DLL Folder
        'If My.Computer.FileSystem.DirectoryExists(Home.DLLPath) = False Then My.Computer.FileSystem.CreateDirectory(Home.DLLPath)
        'Dim dll As String
        'dll = Home.DLLPath & "ICSharpCode.SharpZipLib.dll"
        'If My.Computer.FileSystem.FileExists(dll) = False Then
        '    System.IO.File.WriteAllBytes(dll, My.Resources.ICSharpCode_SharpZipLib)
        'End If

    End Sub

    Function SatCount()

        Dim Ligne, SatNum, Line
        Dim fichier As String

        SatNum = 0

        For i = 0 To Home.TLE_ListBox.Items.Count - 1

            fichier = Home.TLEPath & Home.TLE_ListBox.Items(i)
            FileOpen(i + 1, fichier, OpenMode.Input)
            Ligne = 0
            While Not EOF(i + 1)
                Line = LineInput(i + 1)
                Ligne = Ligne + 1
            End While

            SatNum = SatNum + (Ligne / 3)
            FileClose(i + 1)

        Next

        SatCount = CInt(SatNum)

    End Function

    Sub Pause(ByVal t As Integer)

        System.Threading.Thread.Sleep(t * 1000)

    End Sub

    Function CheckTLE() As Boolean
        Dim Message As String

        Message = "Check TLE :" & vbCrLf & vbCrLf
        CheckTLE = True

        'Check TLE
        'Line Integrity
        If Home.TLETextBox.Text = "" Or Home.TLETextBox.Text.Length < 69 Then
            CheckTLE = False
            Message = Message & "Line 1 and/or Line 2 seems to be empty or incorect."
        End If

        On Error Resume Next
        'Lines
        If Home.TLETextBox.Lines(0) = "" Then
            CheckTLE = False
            Message = Message & vbCrLf & "Line 1 seems to be empty"
        Else
            If Home.TLETextBox.Lines(0).StartsWith("1 ") = False Then
                CheckTLE = False
                Message = Message & vbCrLf & "Line 1 must begin with '1 '"
            End If
        End If

        If Home.TLETextBox.Lines(1) = "" Then
            CheckTLE = False
            Message = Message & vbCrLf & "Line 2 seems to be empty"
        Else
            If Home.TLETextBox.Lines(1).StartsWith("2 ") = False Then
                CheckTLE = False
                Message = Message & vbCrLf & "Line 2 must begin with '2 ' "
            End If
        End If

        'Characters Numbers
        Dim Line0, Line1 As Integer
        Line0 = Home.TLETextBox.Lines(0).Length
        Line1 = Home.TLETextBox.Lines(1).Length
        If Line0 < 69 Or Line1 < 69 Then
            CheckTLE = False
            Message = Message & vbCrLf & vbCrLf & "Each line must contain 69 characters." & vbLf & _
                   "Actually: " & vbLf & Line0 & " characters on line 1" & vbLf & Line1 & " characters on line 2"
        End If

        If Home.OptionsSaved = False Then
            If CheckTLE = False Then MessageBox.Show(Message, "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End If

    End Function

    Public Sub DownloadFile( _
     ByVal address As Uri, _
     ByVal destinationFileName As String, _
     ByVal networkCredentials As Net.ICredentials, _
     ByVal showUI As Boolean, _
     ByVal connectionTimeout As Integer, _
     ByVal overwrite As Boolean _
    )
    End Sub

    'REGIONAL PARAMETERS

    Sub SetCulture()

        'Séparateur de décimal
        Dim c As CultureInfo

        'Langue
        Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("fr-FR")

        c = New CultureInfo(Thread.CurrentThread.CurrentCulture.Name)
        c.NumberFormat.NumberDecimalSeparator = "."
        c.NumberFormat.CurrencyDecimalSeparator = "."

        Thread.CurrentThread.CurrentCulture = c
        Thread.CurrentThread.CurrentUICulture = c

    End Sub

    Sub TimeZoneDef()

        ' Prendre le time zone local , la datetime de maintenant.
        'l'année actuelle
        Dim localZone As TimeZone = TimeZone.CurrentTimeZone
        Dim currentDate As DateTime = DateTime.Now
        Dim currentYear As Integer = currentDate.Year

        ' Affiche le fuseau local
        ' Affiche le fuseau de l'heure d'été.
        Options.TimeZoneBox.Text = localZone.StandardName

        ' Affiche date et heure courantes
        Options.CurrentDateBox.Text = currentDate

        ' Affiche date et heure courantes UTC + OffsetUTC
        Dim Offset, Sign
        Dim currentUTC As DateTime = localZone.ToUniversalTime(currentDate)
        Home.OffsetUTC = localZone.GetUtcOffset(currentDate).Hours.ToString("00")

        Options.CurrentDateUTCBox.Text = currentUTC
        Options.OffsetUTCBox.Text = Home.OffsetUTC & ":00"

    End Sub

    'EXPORTS MODULES

    Sub ExportToGMAT()

        'Définition des paramètres
        Dim SAT, PERIOD, TrackPlot, ForceModel

        If Home.EpochFormat.SelectedIndex = 0 Then Home.EPOCH = Home.EPOCHBox.Text
        If Home.EpochFormat.SelectedIndex = 1 Then Home.EPOCH = GregtoMJD2(Home.EPOCHBox.Text)

        SAT = Home.SCName.Text

        'Options
        'Periodes
        If Home.DracoPeriod_Label.Text = "sec" Then PERIOD = Home.DPER * Home.Propagate
        If Home.DracoPeriod_Label.Text = "min" Then PERIOD = Home.DPER * 60 * Home.Propagate

        'Track Plot
        If CBool(Home.ShowTrack) = True Then
            TrackPlot = _
"Create GroundTrackPlot DefaultGroundTrackPlot;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.SolverIterations = Current;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.UpperLeft = [ 0.004716981132075472 0.4570446735395189 ];" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.Size = [ 0.5009433962264152 0.4501718213058419 ];" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.RelativeZOrder = 17;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.Add = {DefaultSC, Earth};" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.DataCollectFrequency = 1;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.UpdatePlotFrequency = 50;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.NumPointsToRedraw = 0;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.ShowPlot = true;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.CentralBody = Earth;" & vbCrLf & _
"GMAT DefaultGroundTrackPlot.TextureMap = '../data/graphics/texture/ModifiedBlueMarble.jpg';"
        Else
            TrackPlot = ""
        End If

        'Force Model
        If CBool(Home.OptionGmatModel1) = True Then
            ForceModel = _
"%----------------------------------------" & vbCrLf & _
"%---------- ForceModels" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"Create ForceModel DefaultProp_ForceModel;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.CentralBody = Earth;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.PrimaryBodies = {Earth};" & vbCrLf & _
"GMAT DefaultProp_ForceModel.Drag = None;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.SRP = Off;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.RelativisticCorrection = Off;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.ErrorControl = RSSStep;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.Degree = 2;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.Order = 0;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.PotentialFile = 'JGM3.cof';" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.EarthTideModel = 'None';"
        ElseIf CBool(Home.OptionGmatModel2) = True Then
            ForceModel = _
"%----------------------------------------" & vbCrLf & _
"%---------- ForceModels" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"Create ForceModel DefaultProp_ForceModel;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.CentralBody = Earth;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.PrimaryBodies = {Earth};" & vbCrLf & _
"GMAT DefaultProp_ForceModel.PointMasses = {Luna, Sun};" & vbCrLf & _
"GMAT DefaultProp_ForceModel.Drag = None;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.SRP = On;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.RelativisticCorrection = Off;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.ErrorControl = RSSStep;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.Degree = 4;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.Order = 4;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.PotentialFile = 'JGM3.cof';" & vbCrLf & _
"GMAT DefaultProp_ForceModel.GravityField.Earth.EarthTideModel = 'None';" & vbCrLf & _
"GMAT DefaultProp_ForceModel.Drag.AtmosphereModel = MSISE90;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.Drag.F107 = 150;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.Drag.F107A = 150;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.Drag.MagneticIndex = 3;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.SRP.Flux = 1367;" & vbCrLf & _
"GMAT DefaultProp_ForceModel.SRP.Nominal_Sun = 149597870.691;"
        End If

        Dim GMATCode As String
        GMATCode = "%General Mission Analysis Tool(GMAT) Script" & vbCrLf & _
"%Created from TLE Analyser:" & Date.Now & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"%----------" & SAT & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"Create Spacecraft DefaultSC;" & vbCrLf & _
"GMAT DefaultSC.DateFormat = UTCModJulian;" & vbCrLf & _
"GMAT DefaultSC.Epoch = '" & Home.EPOCH & "';" & vbCrLf & _
"GMAT DefaultSC.CoordinateSystem = EarthMJ2000Eq;" & vbCrLf & _
"GMAT DefaultSC.DisplayStateType = Keplerian;" & vbCrLf & _
"GMAT DefaultSC.SMA = " & Home.SMA & ";" & vbCrLf & _
"GMAT DefaultSC.ECC = " & Home.ECC & ";" & vbCrLf & _
"GMAT DefaultSC.INC = " & Home.INC & ";" & vbCrLf & _
"GMAT DefaultSC.RAAN = " & Home.RAAN & ";" & vbCrLf & _
"GMAT DefaultSC.AOP = " & Home.AOP & ";" & vbCrLf & _
"GMAT DefaultSC.TA = " & Maths.RAD2DEG * (Home.TA) & ";" & vbCrLf & _
"GMAT DefaultSC.DryMass = 850;" & vbCrLf & _
"GMAT DefaultSC.Cd = 2.2;" & vbCrLf & _
"GMAT DefaultSC.Cr = 1.8;" & vbCrLf & _
"GMAT DefaultSC.DragArea = 15;" & vbCrLf & _
"GMAT DefaultSC.SRPArea = 1;" & vbCrLf & _
"GMAT DefaultSC.NAIFId = -123456789;" & vbCrLf & _
"GMAT DefaultSC.NAIFIdReferenceFrame = -123456789;" & vbCrLf & _
"GMAT DefaultSC.Id = 'SatId';" & vbCrLf & _
"GMAT DefaultSC.Attitude = CoordinateSystemFixed;" & vbCrLf & _
"GMAT DefaultSC.ModelFile = '../data/vehicle/models/aura.3ds';" & vbCrLf & _
"GMAT DefaultSC.ModelOffsetX = 0;" & vbCrLf & _
"GMAT DefaultSC.ModelOffsetY = 0;" & vbCrLf & _
"GMAT DefaultSC.ModelOffsetZ = 0;" & vbCrLf & _
"GMAT DefaultSC.ModelRotationX = 0;" & vbCrLf & _
"GMAT DefaultSC.ModelRotationY = 0;" & vbCrLf & _
"GMAT DefaultSC.ModelRotationZ = 0;" & vbCrLf & _
"GMAT DefaultSC.ModelScale = 3;" & vbCrLf & _
"GMAT DefaultSC.AttitudeDisplayStateType = 'Quaternion';" & vbCrLf & _
"GMAT DefaultSC.AttitudeRateDisplayStateType = 'AngularVelocity';" & vbCrLf & _
"GMAT DefaultSC.AttitudeCoordinateSystem = 'EarthMJ2000Eq';" & vbCrLf & _
"GMAT DefaultSC.Q1 = 0;" & vbCrLf & _
"GMAT DefaultSC.Q2 = 0;" & vbCrLf & _
"GMAT DefaultSC.Q3 = 0;" & vbCrLf & _
"GMAT DefaultSC.Q4 = 1;" & vbCrLf & _
"GMAT DefaultSC.EulerAngleSequence = '321';" & vbCrLf & _
"GMAT DefaultSC.AngularVelocityX = 0;" & vbCrLf & _
"GMAT DefaultSC.AngularVelocityY = 0;" & vbCrLf & _
"GMAT DefaultSC.AngularVelocityZ = 0;" & vbCrLf & _
ForceModel & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"%---------- Propagators" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"Create Propagator DefaultProp;" & vbCrLf & _
"GMAT DefaultProp.FM = DefaultProp_ForceModel;" & vbCrLf & _
"GMAT DefaultProp.Type = RungeKutta89;" & vbCrLf & _
"GMAT DefaultProp.InitialStepSize = 60;" & vbCrLf & _
"GMAT DefaultProp.Accuracy = 9.999999999999999e-012;" & vbCrLf & _
"GMAT DefaultProp.MinStep = 0.001;" & vbCrLf & _
"GMAT DefaultProp.MaxStep = 2700;" & vbCrLf & _
"GMAT DefaultProp.MaxStepAttempts = 50;" & vbCrLf & _
"GMAT DefaultProp.StopIfAccuracyIsViolated = true;" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"%---------- Burns" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"Create ImpulsiveBurn DefaultIB;" & vbCrLf & _
"GMAT DefaultIB.CoordinateSystem = Local;" & vbCrLf & _
"GMAT DefaultIB.Origin = Earth;" & vbCrLf & _
"GMAT DefaultIB.Axes = VNB;" & vbCrLf & _
"GMAT DefaultIB.Element1 = 0;" & vbCrLf & _
"GMAT DefaultIB.Element2 = 0;" & vbCrLf & _
"GMAT DefaultIB.Element3 = 0;" & vbCrLf & _
"GMAT DefaultIB.DecrementMass = false;" & vbCrLf & _
"GMAT DefaultIB.Isp = 300;" & vbCrLf & _
"GMAT DefaultIB.GravitationalAccel = 9.810000000000001;" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"%---------- Subscribers" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"Create OrbitView DefaultOrbitView;" & vbCrLf & _
"GMAT DefaultOrbitView.SolverIterations = Current;" & vbCrLf & _
"GMAT DefaultOrbitView.UpperLeft = [ 0.004716981132075472 0 ];" & vbCrLf & _
"GMAT DefaultOrbitView.Size = [ 0.5009433962264152 0.4501718213058419 ];" & vbCrLf & _
"GMAT DefaultOrbitView.RelativeZOrder = 5;" & vbCrLf & _
"GMAT DefaultOrbitView.Add = {DefaultSC, Earth};" & vbCrLf & _
"GMAT DefaultOrbitView.CoordinateSystem = EarthMJ2000Eq;" & vbCrLf & _
"GMAT DefaultOrbitView.DrawObject = [ true true ];" & vbCrLf & _
"GMAT DefaultOrbitView.OrbitColor = [ 255 32768 ];" & vbCrLf & _
"GMAT DefaultOrbitView.TargetColor = [ 8421440 0 ];" & vbCrLf & _
"GMAT DefaultOrbitView.DataCollectFrequency = 1;" & vbCrLf & _
"GMAT DefaultOrbitView.UpdatePlotFrequency = 50;" & vbCrLf & _
"GMAT DefaultOrbitView.NumPointsToRedraw = 0;" & vbCrLf & _
"GMAT DefaultOrbitView.ShowPlot = true;" & vbCrLf & _
"GMAT DefaultOrbitView.ViewPointReference = Earth;" & vbCrLf & _
"GMAT DefaultOrbitView.ViewPointVector = [ 30000 0 0 ];" & vbCrLf & _
"GMAT DefaultOrbitView.ViewDirection = Earth;" & vbCrLf & _
"GMAT DefaultOrbitView.ViewScaleFactor = 1;" & vbCrLf & _
"GMAT DefaultOrbitView.ViewUpCoordinateSystem = EarthMJ2000Eq;" & vbCrLf & _
"GMAT DefaultOrbitView.ViewUpAxis = Z;" & vbCrLf & _
"GMAT DefaultOrbitView.CelestialPlane = Off;" & vbCrLf & _
"GMAT DefaultOrbitView.XYPlane = Off;" & vbCrLf & _
"GMAT DefaultOrbitView.WireFrame = Off;" & vbCrLf & _
"GMAT DefaultOrbitView.Axes = On;" & vbCrLf & _
"GMAT DefaultOrbitView.Grid = Off;" & vbCrLf & _
"GMAT DefaultOrbitView.SunLine = On;" & vbCrLf & _
"GMAT DefaultOrbitView.UseInitialView = On;" & vbCrLf & _
"GMAT DefaultOrbitView.StarCount = 7000;" & vbCrLf & _
"GMAT DefaultOrbitView.EnableStars = On;" & vbCrLf & _
"GMAT DefaultOrbitView.EnableConstellations = Off;" & vbCrLf & _
 TrackPlot & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"%---------- Mission Sequence" & vbCrLf & _
"%----------------------------------------" & vbCrLf & _
"BeginMissionSequence;" & vbCrLf & _
"Propagate DefaultProp(DefaultSC) {DefaultSC.ElapsedSecs = " & PERIOD & "};"

        File.WriteAllText(Home.GMATPath & SAT & ".script", GMATCode)

        MessageBox.Show(SAT & ".script has been created in " & vbCrLf & Home.GMATPath, "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

    End Sub

    Sub ExportToCelestia()

        'Déclaration des variables
        Dim SatPath = Home.CELESTIAPath & Home.SATNAME & "\"
        Dim DataPath = SatPath & "data\"
        Dim ModelPath = SatPath & "models\"
        Dim TexturePath = SatPath & "textures\"
        Dim FileContent As String

        'Création des dossiers et copie des fichiers de base
        'Dossier MODEL
        If My.Computer.FileSystem.DirectoryExists(ModelPath) = False Then My.Computer.FileSystem.CreateDirectory(ModelPath)
        My.Computer.FileSystem.WriteAllBytes(ModelPath & Home.SATNAME & ".3ds", My.Resources.satellite, False)

        'Dossier TEXTURES
        If My.Computer.FileSystem.DirectoryExists(TexturePath) = False Then My.Computer.FileSystem.CreateDirectory(TexturePath)

        If My.Computer.FileSystem.FileExists(TexturePath & "IRID_OR.jpg") = False Then
            My.Computer.FileSystem.WriteAllBytes(TexturePath & "IRID_OR", My.Resources.IRID_OR, False)
            My.Computer.FileSystem.RenameFile(TexturePath & "IRID_OR", "IRID_OR.jpg")
        End If

        If My.Computer.FileSystem.FileExists(TexturePath & "SAT_01.jpg") = False Then
            My.Computer.FileSystem.WriteAllBytes(TexturePath & "SAT_01", My.Resources.SAT_01, False)
            My.Computer.FileSystem.RenameFile(TexturePath & "SAT_01", "SAT_01.jpg")
        End If

        If My.Computer.FileSystem.FileExists(TexturePath & "SAT_03.jpg") = False Then
            My.Computer.FileSystem.WriteAllBytes(TexturePath & "SAT_03", My.Resources.SAT_03, False)
            My.Computer.FileSystem.RenameFile(TexturePath & "SAT_03", "SAT_03.jpg")
        End If

        If My.Computer.FileSystem.FileExists(TexturePath & "SAT_04.jpg") = False Then
            My.Computer.FileSystem.WriteAllBytes(TexturePath & "SAT_04", My.Resources.SAT_04, False)
            My.Computer.FileSystem.RenameFile(TexturePath & "SAT_04", "SAT_04.jpg")
        End If

        'Dossier EXTRAS
        Dim apost As Char = Convert.ToChar(34)
        FileContent = _
            "# Generated by TLE ANALYSER  https://sourceforge.net/projects/tleanalyser/" & vbCrLf & _
            "#" & vbCrLf & _
            "# This file contains the orbital elements for" & Home.SATNAME & "satellite" & vbCrLf & _
            "# The datas come from the Celestrak site : http://celestrak.com/NORAD/elements/" & vbCrLf & _
            "#" & vbCrLf & _
            apost & Home.SATNAME & apost & " " & apost & "Sol/Earth" & apost & vbCrLf & _
            "{" & vbCrLf & _
            "Class " & apost & "spacecraft" & apost & vbCrLf & _
            "Mesh " & apost & Home.SATNAME & ".3ds" & apost & vbCrLf & _
        "Radius 0.005 " & vbCrLf & _
        "EllipticalOrbit" & vbCrLf & _
        "{" & vbCrLf & _
        "Period " & Home.APER / 1440 & vbCrLf & _
        "SemiMajorAxis " & Home.SMA & vbCrLf & _
        "Eccentricity " & Home.ECC & vbCrLf & _
        "Inclination " & Home.INC & vbCrLf & _
        "AscendingNode " & Home.RAAN & vbCrLf & _
        "ArgOfPericenter " & Home.AOP & vbCrLf & _
        "MeanAnomaly " & Home.MA & vbCrLf & _
        "Epoch " & Home.EPOCH + 2430000 & vbCrLf & _
        "}" & vbCrLf & _
        "Albedo 10" & vbCrLf & _
        "}"
        If My.Computer.FileSystem.FileExists(SatPath & Home.SATNAME & ".ssc") = True Then File.Delete(SatPath & Home.SATNAME & ".ssc")
        File.AppendAllText(SatPath & Home.SATNAME & ".ssc", FileContent)
        MessageBox.Show(Home.SATNAME & ".ssc" & " has been created in " & vbCrLf & SatPath, "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

    End Sub

    Sub ExportToGoogle()

        Home.Cursor = Cursors.WaitCursor

        'Déclaration des variables
        Dim PlaceMark_Anime, Coordinates, Satellite, LookAt, FileContent, TimeStamp As String
        Dim FichierGE = Home.PlotPath & Home.SATNAME & ".plot"
        Dim i As Integer
        Dim k As Integer = 0
        Dim L1, L2, L10, L20
        Dim JJ1, HH1, lat1, long1, alt1
        Dim JJ2, HH2, lat2, long2, alt2
        PlaceMark_Anime = ""
        Coordinates = ""
        Satellite = ""

        Home.ProgressBar.Value = 0
        Home.ProgressBar.Visible = True

        '----------------------------------------GENERATION DU CODE----------------------------------------------
        'on compte les lignes
        i = 360 * Home.MapPeriodNbr.SelectedItem
        Dim Ligne(i) As String
        Dim Limite = i

        'On génère les lignes PlaceMark pour le section ANIMEE

        FileOpen(1, FichierGE, OpenMode.Input)
        For i = 0 To Limite
            Ligne(i) = LineInput(1)
        Next

        k = 0
        For j = 0 To i - 1
            L10 = Split(Ligne(k), ",")
            JJ1 = L10(3)
            HH1 = L10(4)
            long1 = L10(5)
            lat1 = L10(6)
            alt1 = L10(7)
            k = k + 1
            If k = Limite Then Exit For
            L20 = Split(Ligne(k), ",")
            JJ2 = L20(3)
            HH2 = L20(4)
            long2 = L20(5)
            lat2 = L20(6)
            alt2 = L20(7)

            PlaceMark_Anime = PlaceMark_Anime & _
    "<Placemark><name>" & JJ2 & " " & HH2 & "</name>" & vbCrLf & _
    "<TimeSpan><begin>" & JJ1 & "T" & HH1 & "Z</begin></TimeSpan>" & vbCrLf & _
    "<styleUrl>#redLine</styleUrl>" & vbCrLf & _
    "<LineString><tessellate>0</tessellate>" & vbCrLf & _
    "<extrude>0</extrude>" & vbCrLf & _
    "<altitudeMode>absolute</altitudeMode>" & vbCrLf & _
    "<coordinates>" & vbCrLf & _
    long1 & "," & lat1 & "," & alt1 & vbCrLf & _
    long2 & "," & lat2 & "," & alt2 & vbCrLf & _
    "</coordinates>" & vbCrLf & _
    "</LineString>" & vbCrLf & _
    "</Placemark>" & vbCrLf & vbCrLf

            Coordinates = Coordinates & _
    long1 & "," & lat1 & "," & alt1 & vbCrLf

            Satellite = Satellite & _
"<Placemark><name>" & JJ2 & " " & HH2 & "</name>" & vbCrLf & _
"<TimeSpan><begin>" & JJ1 & "T" & HH1 & "Z</begin><end>" & JJ2 & "T" & HH2 & "Z</end></TimeSpan>" & vbCrLf & _
"<styleUrl>#satellite2</styleUrl>" & vbCrLf & _
"<Point><altitudeMode>absolute</altitudeMode>" & vbCrLf & _
"<coordinates>" & vbCrLf & _
long2 & "," & lat2 & "," & alt2 & vbCrLf & _
"</coordinates>" & vbCrLf & _
"</Point>" & vbCrLf & _
"</Placemark>"

            Home.ProgressBar.Value = j * 100 / (360 * Home.MapPeriodNbr.SelectedItem)

        Next j

        'TimeStamp
        L10 = Split(Ligne(0), ",")
        JJ1 = L10(3)
        HH1 = L10(4)
        TimeStamp = JJ1 & "T" & HH1 & "Z"

        LookAt = _
    "<LookAt>" & vbCrLf & _
    "<gx:TimeStamp><when>" & TimeStamp & "</when></gx:TimeStamp>" & vbCrLf & _
    "<longitude>0</longitude><latitude>0</latitude><altitude>0</altitude><heading>0</heading><tilt>0</tilt><range>15312547.31385426</range>" & vbCrLf & _
    "</LookAt>" & vbCrLf & vbCrLf

        FileClose(1)
        '----------------------------------------FIN DE GENERATION DU CODDE-----------------------------------------

        '----------------------------------------------COPIE DE FICHIERS--------------------------------------------
        If My.Computer.FileSystem.DirectoryExists(Home.GOOGLEPath & Home.SATNAME) = False Then My.Computer.FileSystem.CreateDirectory(Home.GOOGLEPath & Home.SATNAME)
        Dim SatPath = Home.GOOGLEPath & Home.SATNAME & "\"

        'Image du satellite
        Dim PngFile = SatPath & "satellite.png"
        If My.Computer.FileSystem.FileExists(PngFile) = False Then
            My.Computer.FileSystem.WriteAllBytes(SatPath & "satellite", My.Resources.satellite_PNG, False)
            My.Computer.FileSystem.RenameFile(SatPath & "satellite", "satellite.png")
        End If

        'Fichier KML
        Dim KMLFile = SatPath & Home.SATNAME & ".txt"
        Dim KMLFileKML = SatPath & Home.SATNAME & ".kml"

        'If My.Computer.FileSystem.FileExists(KMLFile) = True Then File.Delete(KMLFile)
        'If My.Computer.FileSystem.FileExists(KMLFileKML) = True Then File.Delete(KMLFileKML)
        My.Computer.FileSystem.WriteAllBytes(KMLFileKML, My.Resources.TemplateKML, False)
        'File.AppendAllText(KMLFileKML, PlaceMark_Anime)

        'Modification du fichier KML
        FileContent = File.ReadAllText(KMLFileKML)
        FileContent = FileContent.Replace("#SATNAME#", Home.SATNAME)
        FileContent = FileContent.Replace("#LOOKAT#", LookAt)
        FileContent = FileContent.Replace("#ANIME#", PlaceMark_Anime)
        FileContent = FileContent.Replace("#COORDINATES#", Coordinates)
        FileContent = FileContent.Replace("#SATELLITE#", Satellite)
        File.Delete(KMLFileKML)
        File.AppendAllText(KMLFileKML, FileContent)

        '----------------------------------------------COMPRESSIONS DES FICHIERS--------------------------------------------
        'CompToZip(ZipFileName, PngFile, KMLFileKML)
        'Dim ZipFIleName = SatPath & Home.SATNAME & ".zip"

        Home.ProgressBar.Visible = False
        Home.Cursor = Cursors.Default
        MessageBox.Show(Home.SATNAME & ".kml has been created in " & vbCrLf & SatPath, "TLE Analyser", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)


    End Sub

    'INIT FUNCTIONS

    Function Get_Private_Profile_Int(ByVal cAppName As String, ByVal cKeyName As String, ByVal nKeyDefault As Integer, ByVal cProfName As String) As Long
        ' LIRE UN ENTIER
        ' Parametres:
        ' cAppName Correspond à [Rubrique]
        ' cKeyName Nom de l'entrée, de la clé
        ' nKeyDefault Valeur par défaut de la chaîne cherchée
        ' cProfName Nom du Fichier "INI" Privé
        ' Sortie: La fonction retourne une valeur numérique entière
        Get_Private_Profile_Int = GetPrivateProfileInt(cAppName, cKeyName, nKeyDefault, cProfName)
    End Function

    Function Get_Private_Profile_String(ByVal cAppName As String, ByVal cKeyName As String, ByVal cKeyDefault As String, ByRef cKeyValue As String, ByVal cProfName As String) As Integer
        ' LIRE UNE STRING
        ' Parametres:
        ' cAppName Correspond à [Rubrique]
        ' cKeyName Nom de l'entrée, de la clé
        ' cKeyDefault Valeur par défaut de la chaîne cherchée
        ' cKeyValue Valeur lue en face de l'Entrée ou cKeyDefault si l'Entrée est vide
        ' cProfName Nom du Fichier "INI" Privé
        '
        ' Sortie:
        ' Valeur lue dans cKeyValue
        ' La fonction retourne le nombre de caractères dans cKeyValue
        Dim iReaded As Integer
        Const sLongueur As Short = 255
        If cKeyName = "" Then
            cKeyValue = Space$(1025)
            iReaded = GetPrivateProfileString(cAppName, "", "", cKeyValue, 1024, cProfName)
        Else
            cKeyValue = Space$(255)
            iReaded = GetPrivateProfileString(cAppName, cKeyName, cKeyDefault, cKeyValue, sLongueur, cProfName)
        End If
        cKeyValue = Trim$(cKeyValue)
        'Enlever le dernier caractère?
        'If Len(cKeyValue) <> 0 Then
        ' cKeyValue = Mid$(cKeyValue, 1, Len(cKeyValue) - 1)
        'End If
        Get_Private_Profile_String = iReaded
    End Function

    Function Put_Private_Profile_String(ByVal cAppName As String, ByVal cKeyName As String, ByVal cKeyValue As String, ByVal cProfName As String) As Boolean
        ' ECRIRE UNE STRING
        ' Parametres:
        ' cAppName Correspond à [Rubrique]
        ' cKeyName Nom de l'entrée de la clé
        ' cKeyValue Valeur lue en face de l'Entrée ou cKeyDefault si l'Entrée est vide
        ' cProfName Nom du Fichier "INI" Privé
        ' Sortie:
        ' La fonction retourne True si cela a marché
        Dim Status As Long
        Status = WritePrivateProfileString(cAppName, cKeyName, cKeyValue, cProfName)
        If (Status <> 0) Then
            Put_Private_Profile_String = True
        Else
            Put_Private_Profile_String = False
        End If
    End Function

    Sub ReadInit()

        'INIT VIRIABLES
        Home.TleUpdateDate = Space(255)
        Home.Speed = Space(255)
        Home.TrackMode1 = Space(255)
        Home.TrackMode2 = Space(255)
        Home.SatVisual1 = Space(255)
        Home.SatVisual2 = Space(255)
        Home.SatVisual3 = Space(255)
        Home.SatVisual4 = Space(255)
        Home.ShowTrack = Space(255)
        Home.OptionGmatModel1 = Space(255)
        Home.OptionGmatModel2 = Space(255)
        Home.Propagate = Space(255)
        Home.LngWin = Space(255)
        Home.LatWin = Space(255)
        Home.REFTLE = Space(255)
        Home.REFRM = Space(255)
        Home.LatMis = Space(255)
        Home.LngMis = Space(255)
        Home.SimulOn = Space(255)

        Dim cIniFile As String = Home.AppPath & "tlea.ini" 'Nom du fichier Ini
        Dim istat As Integer

        On Error Resume Next

        istat = Get_Private_Profile_String("TLE", "date", "", Home.TleUpdateDate, cIniFile)

        istat = Get_Private_Profile_String("Tracking Options", "TrackMode1", "", Home.TrackMode1, cIniFile)
        istat = Get_Private_Profile_String("Tracking Options", "TrackMode2", "", Home.TrackMode2, cIniFile)
        istat = Get_Private_Profile_String("Tracking Options", "speed", "", Home.Speed, cIniFile)
        istat = Get_Private_Profile_String("Tracking Options", "SatVisual1", "", Home.SatVisual1, cIniFile)
        istat = Get_Private_Profile_String("Tracking Options", "SatVisual2", "", Home.SatVisual2, cIniFile)
        istat = Get_Private_Profile_String("Tracking Options", "SatVisual3", "", Home.SatVisual3, cIniFile)
        istat = Get_Private_Profile_String("Tracking Options", "SatVisual4", "", Home.SatVisual4, cIniFile)

        istat = Get_Private_Profile_String("Tracking Options", "Simulation On", "false", Home.SimulOn, cIniFile)

        istat = Get_Private_Profile_String("GMAT Export", "show track", "", Home.ShowTrack, cIniFile)
        istat = Get_Private_Profile_String("GMAT Export", "OptionGmatModel1", "", Home.OptionGmatModel1, cIniFile)
        istat = Get_Private_Profile_String("GMAT Export", "OptionGmatModel2", "", Home.OptionGmatModel2, cIniFile)
        istat = Get_Private_Profile_String("GMAT Export", "propagate", "", Home.Propagate, cIniFile)

        istat = Get_Private_Profile_String("Station Keeping", "lng window", "", Home.LngWin, cIniFile)
        istat = Get_Private_Profile_String("Station Keeping", "lat window", "", Home.LatWin, cIniFile)
        istat = Get_Private_Profile_String("Station Keeping", "LLREFTLE", "", Home.REFTLE, cIniFile)
        istat = Get_Private_Profile_String("Station Keeping", "LLREFRM", "", Home.REFRM, cIniFile)
        istat = Get_Private_Profile_String("Station Keeping", "lat mission", "", Home.LatMis, cIniFile)
        istat = Get_Private_Profile_String("Station Keeping", "lng mission", "", Home.LngMis, cIniFile)

        Options.TleLastUpdate.Text = Home.TleUpdateDate

        Options.OptionGmatShowTrack.Checked = CBool(Home.ShowTrack)
        Options.OptionGmatModel1.Checked = CBool(Home.OptionGmatModel1)
        Options.OptionGmatModel2.Checked = CBool(Home.OptionGmatModel2)
        Options.OptionGmatPropPer.SelectedItem = CStr(CInt(Home.Propagate))

        Options.TrackMode1.Checked = CBool(Home.TrackMode1)
        Options.TrackMode2.Checked = CBool(Home.TrackMode2)
        Options.TrackSpeed.SelectedItem = CStr(CInt(Home.Speed))
        Options.SatVisual1.Checked = CBool(Home.SatVisual1)
        Options.SatVisual2.Checked = CBool(Home.SatVisual2)
        Options.SatVisual3.Checked = CBool(Home.SatVisual3)
        Options.SatVisual4.Checked = CBool(Home.SatVisual4)
        Options.SimulOnCB.Checked = CBool(Home.SimulOn)
        Options.LongWin.SelectedItem = CDbl(Home.LngWin).ToString("###0.00")
        Options.LatWin.SelectedItem = CDbl(Home.LatWin).ToString("###0.00")
        Options.LLREFTLE.Checked = CBool(Home.REFTLE)
        Options.LLREFRM.Checked = CBool(Home.REFRM)
        Options.MissionLatBox.Text = Home.LatMis
        Options.MissionLongBox.Text = Home.LngMis

        'Option Sat Visual
        Dim SatP1 As New Bitmap(My.Resources.Sat)
        Dim SatP2 As New Bitmap(My.Resources.sat2)
        Dim SatP3 As New Bitmap(My.Resources.gmico1)
        Dim SatP4 As New Bitmap(My.Resources.gmico2)
        SatP1.MakeTransparent(SatP1.GetPixel(0, 0))
        SatP2.MakeTransparent(SatP2.GetPixel(0, 0))
        SatP3.MakeTransparent(SatP1.GetPixel(0, 0))
        SatP4.MakeTransparent(SatP2.GetPixel(0, 0))
        Options.SatPictOp1.Image = SatP1
        Options.SatPictOp2.Image = SatP2
        Options.SatPictOp3.Image = SatP3
        Options.SatPictOp4.Image = SatP4

    End Sub

    Sub SaveInit()

        Dim cIniFile As String = Home.AppPath & "tlea.ini" 'Nom du fichier Ini
        Dim bOk As Boolean

        Home.TleUpdateDate = CStr(Home.TleUpdateDate)

        Home.ShowTrack = CStr(Options.OptionGmatShowTrack.Checked)
        Home.OptionGmatModel1 = CStr(Options.OptionGmatModel1.Checked)
        Home.OptionGmatModel2 = CStr(Options.OptionGmatModel2.Checked)
        Home.Propagate = Options.OptionGmatPropPer.SelectedItem

        Home.TrackMode1 = CStr(Options.TrackMode1.Checked)
        Home.TrackMode2 = CStr(Options.TrackMode2.Checked)
        Home.Speed = Options.TrackSpeed.SelectedItem
        Home.SatVisual1 = CStr(Options.SatVisual1.Checked)
        Home.SatVisual2 = CStr(Options.SatVisual2.Checked)
        Home.SatVisual3 = CStr(Options.SatVisual3.Checked)
        Home.SatVisual4 = CStr(Options.SatVisual4.Checked)
        Home.SimulOn = CStr(Options.SimulOnCB.Checked)

        Home.LngWin = Options.LongWin.SelectedItem
        Home.LatWin = Options.LatWin.SelectedItem
        Home.REFTLE = CStr(Options.LLREFTLE.Checked)
        Home.REFRM = CStr(Options.LLREFRM.Checked)
        Home.LatMis = Options.MissionLatBox.Text
        Home.LngMis = Options.MissionLongBox.Text

        bOk = Put_Private_Profile_String("TLE", "date", Home.TleUpdateDate, cIniFile)

        bOk = Put_Private_Profile_String("Tracking Options", "TrackMode1", Home.TrackMode1, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "TrackMode2", Home.TrackMode2, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "speed", Home.Speed, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "SatVisual1", Home.SatVisual1, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "SatVisual2", Home.SatVisual2, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "SatVisual3", Home.SatVisual3, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "SatVisual4", Home.SatVisual4, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "Simulation On", Home.SimulOn, cIniFile)

        bOk = Put_Private_Profile_String("GMAT Export", "show track", Home.ShowTrack, cIniFile)
        bOk = Put_Private_Profile_String("GMAT Export", "OptionGmatModel1", Home.OptionGmatModel1, cIniFile)
        bOk = Put_Private_Profile_String("GMAT Export", "OptionGmatModel2", Home.OptionGmatModel2, cIniFile)
        bOk = Put_Private_Profile_String("GMAT Export", "propagate", Home.Propagate, cIniFile)

        bOk = Put_Private_Profile_String("Station Keeping", "lng window", Home.LngWin, cIniFile)
        bOk = Put_Private_Profile_String("Station Keeping", "lat window", Home.LatWin, cIniFile)
        bOk = Put_Private_Profile_String("Station Keeping", "LLREFTLE", Home.REFTLE, cIniFile)
        bOk = Put_Private_Profile_String("Station Keeping", "LLREFRM", Home.REFRM, cIniFile)
        bOk = Put_Private_Profile_String("Station Keeping", "lat mission", Home.LatMis, cIniFile)
        bOk = Put_Private_Profile_String("Station Keeping", "lng mission", Home.LngMis, cIniFile)

    End Sub

    'TRACK MODE

    Sub TrackMode() 'Deplace/cache les cadres 

        If Home.TrackMode1 = True Then

            Home.CalcPanel.Visible = True
            Home.CalcPanel.Refresh()

            Home.MapShowGrid.Checked = False
            Home.MapShowGrid.Enabled = True

            Home.MapsTab.Left = 515
            Home.MapsTab.Refresh()
            Home.MapPanel.Refresh()

        ElseIf Home.TrackMode1 = False Then

            Home.CalcPanel.Visible = False
            Home.CalcPanel.Refresh()

            Home.MapShowGrid.Checked = False
            Home.MapShowGrid.Enabled = False

            Home.MapsTab.Left = 165
            Home.MapsTab.Refresh()
            Home.MapPanel.Refresh()

        End If

        ResizeWindow()
        InitGraphics()

        'Sauvegarde dans ini
        Dim cIniFile As String = Home.AppPath & "tlea.ini" 'Nom du fichier Ini
        Dim bOk As Boolean
        bOk = Put_Private_Profile_String("Tracking Options", "TrackMode1", Home.TrackMode1, cIniFile)
        bOk = Put_Private_Profile_String("Tracking Options", "TrackMode2", Home.TrackMode2, cIniFile)

    End Sub

    Sub ResizeWindow() 'Redimensionne les cadres

        If Home.TrackMode1 <> "" Or Home.TrackMode2 <> "" Or Home.TrackMode1 <> Nothing Or Home.TrackMode2 <> Nothing Then
            Home.ResizeMode = True

            Home.MapW1 = Home.MapPanel.Width
            Home.MapH1 = Home.MapPanel.Height

            'MapsTab
            Home.MapsTab.Width = Home.Width - Home.MapsTab.Location.X - 20
            If Home.MapsTab.SelectedIndex = 0 Then
                    Home.MapsTab.Height = 577
                Else
                    Home.MapsTab.Height = Home.Height - Home.MapsTab.Location.Y - 50
                End If
            Home.MapsTab.Refresh()

            'MapPanel/Tracepicture
            If Home.TABMAP.Width / 2 < Home.TABMAP.Height - 20 Then
                Home.MapPanel.Width = Home.TABMAP.Width
                Home.MapPanel.Height = Home.MapPanel.Width / 2
            ElseIf Home.TABMAP.Width / 2 > Home.TABMAP.Height - 20 Then
                Home.MapPanel.Height = Home.TABMAP.Height - 20
                Home.MapPanel.Width = Home.MapPanel.Height * 2
            End If
            Home.MapPanel.Refresh()

            Home.MapW = Home.MapPanel.Width
            Home.MapH = Home.MapPanel.Height

            Home.TracePicture.Width = Home.MapW
            Home.TracePicture.Height = Home.MapH
            Home.TracePicture.MinimumSize = New Point(Home.MapW, Home.MapH)
            Home.TracePicture.MaximumSize = New Point(Home.MapW, Home.MapH)
            Home.TracePicture.Refresh()

            InitGraphics()

            Home.MapW2 = Home.MapW
            Home.MapH2 = Home.MapH

        End If
    End Sub
    
    Sub InitGraphics()
        'Re-Init les données Graphique

        Home.MapW = Home.MapPanel.Width
        Home.MapH = Home.MapPanel.Height

        Home.Trace = New Bitmap(Home.MapW, Home.MapH)
        Home.g = Graphics.FromImage(Home.Trace)

    End Sub

End Module
