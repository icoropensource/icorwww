<%
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2007
'All Right Reserved
'http://www.kmksoftware.prv.pl
'*
'Z funkcji mo¿na korzystaæ za darmo do celów prywatnych i zawodowych.
'Jeœli jesteœ bardzo zadowolony z dzia³ania funkcji i znalaz³a ona zastosowanie w Twojej pracy,
'to proszê o dowolne datki na konto: 78  1140 2004  0000 3602 3243 9571
'*
'2007-05-15 - Przetestowane dzia³anie funkcji z MS Excel 2007. Wszystko OK.
'2005-06-08 - Zmieniono nazwê funkcji, ze "S³ownie" na "Slownie"
'   - bez polskiej litery, bo Excel 2003 jej nie rozpoznawa³.
'2005-08-07 - Poprawiono b³¹d odmiany koñcówek przy x01 tys.; x01 mil.; x01 mld.
'   Usuniêto zbêdne spacje przy równych liczbach z zakresu od 10 tys. do 999 tys.; np. 100 000

Dim IDziesiatka
Dim Przerwa
Public Function Slownie(ByVal Numer, Format)
'Public Function Slownie(ByVal Numer, Format, Spacja)
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2005
'All Right Reserved
'http://www.kmksoftware.prv.pl
'Stop
    Dim TempS, TempD, TempJ, Zlote, Grosze, Licznik, Grd, Grj, MiejsceDz, T_S
    Dim Znak
    Przerwa = True
'    Przerwa = Spacja
    If Numer < 0 Then
        Znak = True
        Numer = Abs(Numer)
    Else
        Znak = False
    End If
    If Numer = 0 Then
        If Format = True Then
            If Przerwa = True Then
                Slownie = "zero kropka zero"
            Else
                Slownie = "zerokropkazero"
            End If
        Else
            If Przerwa = True Then
                Slownie = "zero z³otych zero groszy"
            Else
                Slownie = "zeroz³otychzerogroszy"
            End If
        End If
        Exit Function
    End If
    If Numer > 9999999999999.99 Then
        Slownie = "Z³y zakres danych."
        Exit Function
    End If
    Numer = Trim(CStr(Numer))
    MiejsceDz = InStr(Numer, ".")
    If MiejsceDz > 0 Then
        Grosze = Left(Mid(Numer, MiejsceDz + 1) & "00", 2)
        Grd = Dziesiatki(Right(Grosze, 2))
        If IDziesiatka <> 1 Then Grj = Jednostki(Right(Grosze, 1))
        If Format = True Then
            If IDziesiatka = 0 Then
                If Przerwa = True Then
                    Grosze = " " & Grd & "zero " & Grj
                Else
                    Grosze = Grd & "zero" & Grj
                End If
            Else
                If Przerwa = True Then
                    Grosze = " " & Grd & Grj
                Else
                    Grosze = Grd & Grj
                End If
            End If
        Else
            If Przerwa = True Then
                Grosze = " " & Grd & Grj & "gr"
            Else
                Grosze = Grd & Grj & "gr"
            End If
        End If
        Numer = Trim(Left(Numer, MiejsceDz - 1))
    Else
        If Format = True Then
            If Przerwa = True Then
                Grosze = " zero"
            Else
                Grosze = "zero"
            End If
        Else
            If Przerwa = True Then
                Grosze = " zero gr"
            Else
                Grosze = "zerogr"
            End If
        End If
    End If
    If Numer <> "" Then
        Licznik = 1
        Do While Numer <> ""
            TempS = ""
            TempD = ""
            TempJ = ""
            TempS = Setki(Right("000" & Numer, 3))
            TempD = Dziesiatki(Right("00" & Numer, 2))
            If IDziesiatka <> 1 Then TempJ = Jednostki(Right(Numer, 1))
            Select Case Licznik
                Case 1
                    T_S = TempS & TempD & TempJ
                Case 2
                    T_S = TempS & TempD & TempJ & KTys(Numer)
                Case 3
                    T_S = TempS & TempD & TempJ & KMil(Numer, Licznik)
                Case 4
                    T_S = TempS & TempD & TempJ & KMil(Numer, Licznik)
                Case 5
                    T_S = TempS & TempD & TempJ & KMil(Numer, Licznik)
            End Select
            Zlote = T_S & Zlote
            If Len(Numer) > 3 Then
                Numer = Left(Numer, Len(Numer) - 3)
                Licznik = Licznik + 1
            Else
                Numer = ""
            End If
        Loop
    Else
        If Przerwa = True Then
            Zlote = "zero "
        Else
            Zlote = "zero"
        End If
    End If
    If Format = True Then
        Slownie = Trim(Zlote & "kropka" & Grosze)
    Else
        Slownie = Trim(Zlote & "z³" & Grosze)
    End If
    If Znak = True Then
        If Przerwa = True Then
            Slownie = "minus " + Slownie
        Else
            Slownie = "minus" + Slownie
        End If
    End If
    Slownie = Replace(Slownie, " zero groszy", "")
    Slownie = Replace(Slownie, " zero gr", "")
    Slownie = Replace(Slownie, " kropka zero", "")
End Function

Private Function KTys(ByVal Numer)
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2005
'All Right Reserved
'http://www.kmksoftware.prv.pl
    Dim Tys
    Tys = CDbl(Right("000" & Numer, 3))
    If Tys = 0 Then
        KTys = ""
    Else
        Tys = CDbl(Right(Numer, 2))
        If Tys = 1 Then
            If CDbl(Right("000" & Numer, 3)) > 1 Then
                KTys = "êcy"
            Else
                KTys = "¹c"
            End If
        Else
            If Tys = 12 Or Tys = 13 Or Tys = 14 Then
                KTys = "êcy"
            Else
            Tys = CDbl(Right(Numer, 1))
            Select Case Tys
                Case 2, 3, 4
                    KTys = "¹ce"
                Case Else
                    KTys = "êcy"
            End Select
            End If
        End If
        If Przerwa = True Then KTys = KTys + " "
        KTys = "tysi" & KTys
    End If
End Function

Private Function KMil(ByVal Numer, L)
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2005
'All Right Reserved
'http://www.kmksoftware.prv.pl
    Dim Mil
    Dim RzadW(5)
    RzadW(3) = "milion"
    RzadW(4) = "miliard"
    RzadW(5) = "bilion"
    Mil = CDbl(Right("000" & Numer, 3))
    If Mil = 0 Then
        KMil = ""
    Else
        Mil = CDbl(Right(Numer, 2))
        If Mil = 1 Then
            If CDbl(Right("000" & Numer, 3)) > 1 Then
                KMil = "ów"
            Else
                KMil = ""
            End If
        Else
            If Mil = 12 Or Mil = 13 Or Mil = 14 Then
                KMil = "ów"
            Else
                Mil = CDbl(Right(Numer, 1))
                Select Case Mil
                    Case 2, 3, 4
                        KMil = "y"
                    Case Else
                        KMil = "ów"
                End Select
            End If
        End If
        If Przerwa = True Then KMil = KMil + " "
        KMil = RzadW(L) & KMil
    End If
End Function

Private Function Setki(ByVal Numer)
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2005
'All Right Reserved
'http://www.kmksoftware.prv.pl
    Dim Setka, Wynik
    Setka = CDbl(Left(Numer, 1))
    Select Case Setka
        Case 1
            Wynik = "sto"
        Case 2
            Wynik = "dwieœcie"
        Case 3
            Wynik = "trzysta"
        Case 4
            Wynik = "czterysta"
        Case 5
            Wynik = "piêæset"
        Case 6
            Wynik = "szeœæset"
        Case 7
            Wynik = "siedemset"
        Case 8
            Wynik = "osiemset"
        Case 9
            Wynik = "dziewiêæset"
    End Select
    If (Przerwa = True) And (Setka <> 0) Then Wynik = Wynik + " "
    Setki = Wynik
End Function

Private Function Dziesiatki(ByVal Number)
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2005
'All Right Reserved
'http://www.kmksoftware.prv.pl
    Dim Wynik
    IDziesiatka = CDbl(Left(Number, 1))
    If IDziesiatka = 1 Then
        Select Case CDbl(Number)
            Case 10
                Wynik = "dziesiêæ"
            Case 11
                Wynik = "jedenaœcie"
            Case 12
                Wynik = "dwanaœcie"
            Case 13
                Wynik = "trzynaœcie"
            Case 14
                Wynik = "czternaœcie"
            Case 15
                Wynik = "piêtnaœcie"
            Case 16
                Wynik = "szesnaœcie"
            Case 17
                Wynik = "siedemnaœcie"
            Case 18
                Wynik = "osiemnaœcie"
            Case 19
                Wynik = "dziewiêtnaœcie"
        End Select
    Else
        Select Case IDziesiatka
            Case 2
                Wynik = "dwadzieœcia"
            Case 3
                Wynik = "trzydzieœci"
            Case 4
                Wynik = "czterdzieœci"
            Case 5
                Wynik = "piêædziesi¹t"
            Case 6
                Wynik = "szeœædziesi¹t"
            Case 7
                Wynik = "siedemdziesi¹t"
            Case 8
                Wynik = "osiemdziesi¹t"
            Case 9
                Wynik = "dziewiêædziesi¹t"
        End Select
    End If
    If (Przerwa = True) And (IDziesiatka <> 0) Then Wynik = Wynik + " "
    Dziesiatki = Wynik
End Function

Private Function Jednostki(ByVal Numer)
'Studio Us³ug Komputerowych - KMK Software, Katowice, 2003 - 2005
'All Right Reserved
'http://www.kmksoftware.prv.pl
    Dim Jedst, Wynik
    Jedst = CDbl(Right(Numer, 1))
    Select Case Jedst
        Case 1
            Wynik = "jeden"
        Case 2
            Wynik = "dwa"
        Case 3
            Wynik = "trzy"
        Case 4
            Wynik = "cztery"
        Case 5
            Wynik = "piêæ"
        Case 6
            Wynik = "szeœæ"
        Case 7
            Wynik = "siedem"
        Case 8
            Wynik = "osiem"
        Case 9
            Wynik = "dziewiêæ"
    End Select
    If (Przerwa = True) And (Jedst <> 0) Then Wynik = Wynik + " "
    Jednostki = Wynik
End Function
%>