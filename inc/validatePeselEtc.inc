<%

Function CleanString(ByVal s, alen)
   Dim i, s1
   CleanString = ""
   For i = 1 To Len(s)
      s1 = Mid(s, i, 1)
      If (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Then
         CleanString = CleanString & s1
      End If
   Next
   If (alen > 0) And (Len(CleanString) <> alen) Then
      CleanString = ""
   End If
End Function

Function CleanStringBook(ByVal s, alen)
   Dim i, s1, s2
   CleanStringBook = ""
   s2 = ""
   For i = 1 To Len(s)
      s1 = Mid(s, i, 1)
      If (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Or s1 = "x" Or s1 = "X" Then
         s2 = s2 & s1
      End If
   Next
   CleanStringBook = Mid(s2, Len(s2), 1)
   For i = Len(s2) - 1 To 1 Step -1
      s1 = Mid(s2, i, 1)
      If (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Then
         CleanStringBook = s1 & CleanStringBook
      End If
   Next
   If (alen > 0) And (Len(CleanStringBook) <> alen) Then
      CleanStringBook = ""
   End If
End Function

Function CleanStringMusic(ByVal s, alen)
   Dim i, s1
   CleanStringMusic = ""
   For i = 1 To Len(s)
      s1 = Mid(s, i, 1)
      If (CleanStringMusic = "") And (s1 = "m" Or s1 = "M") Then
         CleanStringMusic = CleanStringMusic & "3"
      ElseIf (Asc(s1) >= Asc("0")) And (Asc(s1) <= Asc("9")) Then
         CleanStringMusic = CleanStringMusic & s1
      End If
   Next
   If (alen > 0) And (Len(CleanStringMusic) <> alen) Then
      CleanStringMusic = ""
   End If
End Function

Function ValidateREGON9(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateREGON9 = 0
   wagi = Array(8, 9, 2, 3, 4, 5, 6, 7)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (suma Mod 11) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateREGON9 = 1
   End If
End Function

Function ValidateREGON14(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateREGON14 = 0
   wagi = Array(2, 4, 8, 5, 0, 9, 7, 3, 6, 1, 2, 4, 8)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   If ValidateREGON9(Mid(s, 1, 9)) = 0 Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (suma Mod 11) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateREGON14 = 1
   End If
End Function

Function ValidateREGON(anumer)
   Dim s
   s = CleanString(anumer, 0)
   ValidateREGON = 0
   If Len(s) = 9 Then
      ValidateREGON = ValidateREGON9(s)
   ElseIf Len(s) = 14 Then
      ValidateREGON = ValidateREGON14(s)
   End If
End Function

Function ValidatePESEL(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidatePESEL = 0
   wagi = Array(1, 3, 7, 9, 1, 3, 7, 9, 1, 3)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidatePESEL = 1
   End If
End Function

Function ValidateNIP(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateNIP = 0
   wagi = Array(6, 5, 7, 2, 3, 4, 5, 6, 7)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = suma Mod 11
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateNIP = 1
   End If
End Function

Function ValidateISBN(anumer)
   Dim s, i, wagi, anum, suma, aarr, acheck
   ValidateISBN = 0
   wagi = Array(10, 9, 8, 7, 6, 5, 4, 3, 2)
   s = CleanStringBook(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (11 - (suma Mod 11)) Mod 11
   If Mid(s, i, 1) = "x" Or Mid(s, i, 1) = "X" Then
      acheck = 10
   Else
      acheck = CLng(Mid(s, i, 1))
   End If
   If suma = acheck Then
      ValidateISBN = 1
   End If
End Function

Function ValidateISSN(anumer)
   Dim s, i, wagi, anum, suma, aarr, acheck
   ValidateISSN = 0
   wagi = Array(8, 7, 6, 5, 4, 3, 2)
   s = CleanStringBook(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (11 - (suma Mod 11)) Mod 11
   If Mid(s, i, 1) = "x" Or Mid(s, i, 1) = "X" Then
      acheck = 10
   Else
      acheck = CLng(Mid(s, i, 1))
   End If
   If suma = acheck Then
      ValidateISSN = 1
   End If
End Function

Function ValidateISMN(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateISMN = 0
   wagi = Array(3, 1, 3, 1, 3, 1, 3, 1, 3)
   s = CleanStringMusic(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateISMN = 1
   End If
End Function

Function ValidateEAN13(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateEAN13 = 0
   wagi = Array(1, 3, 1, 3, 1, 3, 1, 3, 1, 3, 1, 3)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateEAN13 = 1
   End If
End Function

Function ValidateCPV(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateCPV = 0
   wagi = Array(3, 7, 1, 3, 7, 1, 3, 7)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = suma Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateCPV = 1
   End If
End Function

Function ValidateEAN8(anumer)
   Dim s, i, wagi, anum, suma, aarr
   ValidateEAN8 = 0
   wagi = Array(3, 1, 3, 1, 3, 1, 3)
   s = CleanString(anumer, UBound(wagi) + 2)
   If s = "" Then
      Exit Function
   End If
   suma = 0
   For i = 1 To Len(s) - 1
      suma = suma + wagi(i - 1) * CLng(Mid(s, i, 1))
   Next
   suma = (10 - (suma Mod 10)) Mod 10
   If suma = CLng(Mid(s, i, 1)) Then
      ValidateEAN8 = 1
   End If
End Function

%>
