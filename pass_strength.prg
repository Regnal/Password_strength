parameters pwd,lvystup

If Parameters()<2
	lvystup=.F.
Endif

nScore = 0
sComplexity = ""
iUpperCase = 0
iLowerCase = 0
iDigit = 0
iSymbol = 0
iRepeated = 1
iMiddle = 0
iMiddleEx = 1
ConsecutiveMode = 0
iConsecutiveUpper = 0
iConsecutiveLower = 0
iConsecutiveDigit = 0
iLevel = 0
sAlphas = "abcdefghijklmnopqrstuvwxyz"
sNumerics = "01234567890"
nSeqAlpha = 0
nSeqChar = 0
nSeqNumber = 0
cpwd=pwd

If Empty(pwd)
	Return "Without a password!!!   0%"
Endif

For Cnt=1 To Len(pwd)
	ch=Substr(pwd,Cnt,1)

	If Isdigit(ch)
		iDigit=iDigit+1
		If ConsecutiveMode = 3
			iConsecutiveDigit=iConsecutiveDigit+1
		Endif
		ConsecutiveMode = 3
	Endif

	If Isupper(ch) And At(LOWER(ch),sAlphas)>0
		iUpperCase=iUpperCase+1
		If ConsecutiveMode = 1
			iConsecutiveUpper=iConsecutiveUpper+1
		Endif
		ConsecutiveMode = 1
	Endif

	If Islower(ch) And At(ch,sAlphas)>0
		iLowerCase=iLowerCase+1
		If ConsecutiveMode = 2
			iConsecutiveLower=iConsecutiveLower+1
		Endif
		ConsecutiveMode = 2
	Endif

	If !(At(LOWER(ch),sAlphas)>0 Or Isdigit(ch))
		iSymbol=iSymbol+1
		ConsecutiveMode = 0
	Endif

	If At(LOWER(ch),sAlphas)>0 Or Isdigit(ch)
		If Occurs(ch,cpwd)>1
			iRepeated=iRepeated+1
			cpwd=Strtran(cpwd,ch,":")
		Endif
	Endif
	If (iMiddleEx > 1)
		iMiddle = iMiddleEx - 1
	Endif

	If Cnt>1
		If Isdigit(ch) Or !At(LOWER(ch),sAlphas)>0
			iMiddleEx=iMiddleEx+1
		Endif
	Endif
Endfor

For S = 1 To 38
	sFwd = Substr(sAlphas,S, 3)
	sRev = strReverse(sFwd)
	If At(sFwd,Lower(pwd))>0 Or At(sRev,pwd)>0
		nSeqAlpha=nSeqAlpha+1
		nSeqChar=nSeqChar+1
	Endif
Endfor

For S = 1 To 8
	sFwd = Substr(sNumerics,S,3)
	sRev = strReverse(sFwd)
	If At(sFwd,Lower(pwd))>0 Or At(sRev,pwd)>0
		nSeqNumber=nSeqNumber+1
		nSeqChar=nSeqChar+1
	Endif
Endfor

nScore = 4 * Len(pwd)
If lvystup
	? "Password length "+Str(Len(pwd))+","+Str(4 * Len(pwd))
Endif
If iUpperCase > 0
	nScore = nScore+((Len(pwd) - iUpperCase) * 2)
	If lvystup
		? "Uppercase Letters "+Str(iUpperCase)+","+Str(((Len(pwd) - iUpperCase) * 2))
	Endif
Endif
If iLowerCase > 0
	nScore = nScore+((Len(pwd) - iLowerCase) * 2)
	If lvystup
		? "Lowercase Letters "+Str(iLowerCase)+","+Str(((Len(pwd) - iLowerCase) * 2))
	Endif
Endif
If iDigit < Len(pwd)
	nScore = nScore+(iDigit * 4)
	If lvystup
		? "Numbers "+Str(iDigit)+","+Str((iDigit * 4))
	Endif
Endif
nScore = nScore+(iSymbol * 6)
If lvystup
	? "Symbols "+Str(iSymbol)+","+Str((iSymbol * 6))
Endif

nScore = nScore+(iMiddle * 2)
If lvystup
	? "Middle Numbers or Symbols "+Str(iMiddle)+","+Str((iMiddle * 2))
Endif

requirments = 0
If (Len(pwd) >= 8) Then
	requirments=requirments+1     && Min password length
Endif
If iUpperCase > 0
	requirments=requirments+1      && Uppercase letters
Endif
If iLowerCase > 0
	requirments=requirments+1     && Lowercase letters
Endif
If iDigit > 0
	requirments=requirments+1         && Digits
Endif
If iSymbol > 0
	requirments=requirments+1         && Symbols
Endif

&& If we have more than 3 requirments Then
If requirments > 3
	nScore = nScore+(requirments * 2)
	If lvystup
		? "Requirments "+Str(requirments)+","+Str((requirments * 2))
	Endif
Endif

If iDigit = 0 And iSymbol = 0
	nScore=nScore- Len(pwd)
	If lvystup
		? "Letters only "+Str(Len(pwd))+","+Str((Len(pwd)*-1))
	Endif
Endif

If iDigit = Len(pwd)
	nScore=nScore- Len(pwd)
	If lvystup
		? "Numbers only "+Str(Len(pwd))+","+Str((Len(pwd)*-1))
	Endif
Endif

If iRepeated > 1
	nScore=nScore-(iRepeated * (iRepeated - 1))
	If lvystup
		? "Repeat Characters "+Str(iRepeated)+","+Str(-(iRepeated * (iRepeated - 1)))
	Endif
Endif

nScore=nScore-(iConsecutiveUpper * 2)
If lvystup
	? "Consecutive Uppercase Letters "+Str(iConsecutiveUpper)+","+Str(-(iConsecutiveUpper * 2))
Endif

nScore=nScore-(iConsecutiveLower * 2)
If lvystup
	? "Consecutive Lowercase Letters "+Str(iConsecutiveLower)+","+Str(-(iConsecutiveLower * 2))
Endif
nScore=nScore-(iConsecutiveDigit * 2)
If lvystup
	? "Consecutive Numbers "+Str(iConsecutiveDigit)+","+Str(-(iConsecutiveDigit * 2))
Endif
nScore=nScore-(nSeqAlpha * 3)
If lvystup
	? "Sequential Letters (3+) "+Str(nSeqAlpha)+","+Str(-(nSeqAlpha * 3))
Endif
nScore=nScore-(nSeqNumber * 3)
If lvystup
	? "Sequential Numbers (3+) "+Str(nSeqNumber)+","+Str(-(nSeqNumber * 3))
Endif

If nScore > 100 Then
	nScore = 100
Else
	If nScore < 0
		nScore = 0
	Endif
Endif
Do Case
Case (nScore >= 0 And nScore < 20)
	sComplexity = "Very weak "
Case(nScore >= 20 And nScore < 40)
	sComplexity = "Weak "
Case (nScore >= 40 And nScore < 60)
	sComplexity = "Good"
Case (nScore >= 60 And nScore < 80)
	sComplexity = "Strong "
Case (nScore >= 80 And nScore <= 100)
	sComplexity = "Very strong "
Endcase

If lvystup
	?sComplexity+Str(nScore,3)+"% "
Endif

Return sComplexity+Str(nScore,3)+"%"

Function strReverse
Lparameters str_a
Local newstring
newstring = ""
For i = 1 To  Len(str_a)
	newstring = Substr(str_a,i,1) + newstring
Endfor
Return newstring
Endfunc
