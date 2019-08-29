<%

Dim Com_pSelected	'selected������
Dim Com_pNow		'�ꎞ����
Dim Com_pStr		'�ꎞ������
Dim Com_pInt		'�ꎞ����
Dim DEBUG_ : DEBUG_ = 1

'=====================================================================
' [Function		] <select>�^�O�쐬
' [Contents     ] �z�񂩂�select�^�O���쐬����
'---------------------------------------------------------------------
' [Return       ] <select...������
' [Input        ] (integer)mode:�쐬���[�h
'				  (string)selname:selection��
'				  (string array)strarray:�z��P
'				  (string array)strarray02:�z��Q
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_MakeSelectionByStrArray(mode,selname,strarray,strarray02)

	Dim i
	Dim strtmp
	Dim selected

	If(mode = 1) Then

		For i = 0 to UBound(strarray)
			strtmp = strtmp & "<option value='" & i & "'>" & strarray(i) & "</option>" & vbCrLf
		Next

		Com_MakeSelectionByStrArray = "<select name='" & selname & "'>" & vbCrLf & strtmp & "</select>" & vbCrLf

	ElseIf(mode = 2) Then

		For i = 0 to UBound(strarray)

			If(Com_pSelected = strarray02(i)) Then
				selected = "selected"
			Else
				selected = ""
			End If

			strtmp = strtmp & "<option value='" & strarray(i) & "' " & selected & ">" & strarray02(i) & "</option>" & vbCrLf
		Next

		Com_MakeSelectionByStrArray = "<select name='" & selname & "'>" & vbCrLf & strtmp & "</select>" & vbCrLf


	End If

End Function

'=====================================================================
' [Function		] �d�����
' [Contents     ] �z�񂩂�d�����Ȃ���
'---------------------------------------------------------------------
' [Return       ] (string array)�d�����Ȃ������z��
' [Input        ] (string array)strarray:�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function TAKO_SortInOne(strarray)

	Dim i
	Dim j
	Dim pcnt
	Dim flag
	Dim tmpArray()

	ReDim tmpArray(UBound(strarray))
	flag = 0

	For i = 0 to UBound(strarray)
		flag = 0
		If(tmpArray(0) = "") Then
		ElsIf
			For j = 0 to UBound(tmparray)
				If(strarray(i) = tmparray(j)) Then
					flag = 1
					break
				End If
			Next

			If(flag = 0) Then
				tmparray(j) = strarray(i)
			End If

		End If
	Next

End Function

'=====================================================================
' [Function		] Unshift
' [Contents     ] �z��̐擪�ɒl��ǉ�����
'---------------------------------------------------------------------
' [Return       ] (string array)�ǉ������z��
' [Input        ] (string array)arr:�z��
'				  (string)val:�ǉ�����l
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function ArrayUnshift(arr(),val)

	Dim i
	Dim j
	Dim n
	Dim ar

	ar = arr

	ReDim Preserve ar(UBound(ar) + 1)

	n = UBound(ar)

	For i = 0 to n

		j = n - i

		If(j <> 0) Then
			ar(j) = ar(j - 1)

		Else		
			ar(0) = val
		End If
	Next

	arrayUnshift = ar

End Function


'=====================================================================
' [Function		] �A���_�[�o�[�u��
' [Contents     ] "_������_"��Session���̒l�ɕϊ�����
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (string)�ϊ�����O�̕�����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ConvVal(strtmp)

	Dim pVal
	Dim i
	Dim i2
	Dim l
	Dim word
	Dim s

	l = Len(strtmp)

	i = 1
	Do While(1)

		'�J�E���^�������ʒu���z������I���
		If(i > Len(strtmp)) Then Exit Do

		'�ꕶ���Ƃ�
		s = Mid(strtmp,i,1)

		'������ "_" �������ꍇ�A���̃A���_�[�o�[�܂ŕ����� pVal �Ɋi�[����
		If(s = "_") Then

			'�i�[���[�v --- ��������

			pVal = ""
			For i2 = i + 1 to l
				s = Mid(strtmp,i2,1)

				If(s = "_") Then
					'Response.Write(pVal & "�ł���<br>")
					strtmp = Replace(strtmp,"_" & pVal & "_",Session(pVal))
					l = Len(strtmp)
					'i = i2
					Exit For
				ElseIf( _
					(Asc(s) >= 65 AND Asc(s) <= 90) OR _
					(Asc(s) >= 97 AND Asc(s) <= 122) OR _
					(Asc(s) >= 45 AND Asc(s) <= 57) _
				      ) Then
					pVal = pVal & Mid(strtmp,i2,1)
 				Else

					'Response.Write("[" & s & "]�Ȃ̂ł����<br>")
					Exit For

				End If
			Next

			'�i�[���[�v --- �����܂�

		End If

		i = i + 1

	Loop

	Com_ConvVal = strtmp

End Function

'=====================================================================
' [Function		] ��������w�蕶���ŋ�؂��Ĕz��ŕԂ�
' [Contents     ] ��������w�蕶���ŋ�؂��Ĕz��ŕԂ�
'---------------------------------------------------------------------
' [Return       ] (string array)��؂����z��
' [Input        ] (string)strtmp:��؂�O�̕�����
'				  (string)delimiter):��؂蕶��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_KSplit(strtmp,delimiter)

	Dim i
	Dim l
	Dim c
	Dim cnt
	Dim str
	Dim ar_str()

	l = Len(strtmp)
	cnt = -1
	For i = 1 to l

		c = Mid(strtmp,i,1)

		If(c = delimiter Or i = l Or delimiter = "") Then
			If(i = l Or  delimiter = "") Then
				str = str & c
			End If
			cnt = cnt + 1
			ReDim Preserve ar_str(cnt)
			ar_str(cnt) = str
			str = ""
		Else
			str = str & c
		End If

	Next

	Com_KSplit = ar_str

End Function

'=====================================================================
' [Function		] �p���p��
' [Contents     ] �F�����݂ɂ���
'---------------------------------------------------------------------
' [Return       ] (string)���݂̐F
' [Input        ] (string)�W���ȊO�̐F
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Dim Com_Parapara_i : Com_Parapara_i = 0
Function Com_Parapara(col)
	if(Com_Parapara_i = 1) then
		Com_Parapara = col
		Com_Parapara_i = 0
	else
		Com_Parapara = "#FFFFFF"
		Com_Parapara_i = 1
	end if
End Function

'=====================================================================
' [Function		] �z��\��
' [Contents     ] �z������s���ŕ\������
'---------------------------------------------------------------------
' [Return       ] 
' [Input        ] (integer)mode:�\�����[�h
'				  (string array)strarray:�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_PrintArray(mode,strarray())
	Dim i
	Dim imax

	imax = UBound(strarray)
	For i = 0 to imax
		Com_Debug(strarray(i))
	Next

End Function

'=====================================================================
' [Function		] �f�o�O�\��
' [Contents     ] �f�o�O���鎞�Ɏg���֐� DEBUG_�ϐ����P�̎��ɕ\��
'---------------------------------------------------------------------
' [Return       ] 
' [Input        ] (string)str:�\��������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_Debug(str)
	If(DEBUG_) Then Response.Write(str & "<br>" & vbCrLf)
End Function

'=====================================================================
' [Function		] �G���[�\��
' [Contents     ] �G���[���b�Z�[�W��\�����ďI������
'---------------------------------------------------------------------
' [Return       ] 
' [Input        ] (string)str:�\��������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_Err001(str)
	Response.Write(str)
	Response.End
End Function

'=====================================================================
' [Function		] ���t�擾
' [Contents     ] YYYY:MM:DD:HH:MM:SS�`���Ŏ擾����
'---------------------------------------------------------------------
' [Return       ] (string) YYYY:MM:DD:HH:MM:SS������
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetDateString()

	Dim str
	Dim d
	Dim t
	Dim ar

	Dim year
	Dim mon
	Dim mday
	Dim hour
	Dim min
	Dim sec

	str = Now

	ar = Split(str," ")
	d = ar(0)
	t = ar(1)

	ar = Split(d,"/")
	year = 2000 + Cint(ar(0))
	mon = ar(1)
	mday = ar(2)

	ar = Split(t,":")
	hour = ar(0)
	min = ar(1)
	sec = ar(2)

	Com_GetDateString = year  & ":" & mon & ":" & mday & ":" & hour & ":" & min & ":" & sec

End Function

'=====================================================================
' [Function		] �w�b�_�[�o��
' [Contents     ] Perl�̎��悭�g�����֐�
'---------------------------------------------------------------------
' [Return       ] (String)�w�b�_�[������
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function PH()

	PH = ""

End Function

'=====================================================================
' [Function		] Refresh��ʕ\��
' [Contents     ] META��Refresh�̓������y�[�W��\������
'---------------------------------------------------------------------
' [Return       ] (string)HTML������
' [Input        ] (string)url:�����N��URL
'				  (string)sec:�b��
'				  (string)str:html������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_MakeRefresh(url,sec,str)

	Com_MakeRefresh = "<html><head><title>" & str & "</title><META HTTP-EQUIV='Content-Type' Content=""text/html; charset=Shift_JIS""><META HTTP-EQUIV='Refresh' Content=""" & sec & " ; url='" & url & "'""></head><body>" & str & "</body></html>"

End Function

'=====================================================================
' [Function		] HTML�֎~�^�O�̒u��
' [Contents     ] <��>�Ȃǂ̃^�O���������
'---------------------------------------------------------------------
' [Return       ] (string)�u������������
' [Input        ] (string)�u�����镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ConvHTMLTag(strtmp)

	strtmp = Relace(strtmp," ","&nbsp;")
	strtmp = Relace(strtmp,"<","&lt;")
	strtmp = Relace(strtmp,">","&gt;")
	strtmp = Relace(strtmp,vbCr + vbLf,vbLf)
	strtmp = Relace(strtmp,vbCr,"<br>")
	strtmp = Relace(strtmp,"<br><br>","<br> <br>")
	strtmp = Relace(strtmp,"_N_","<br>")

	Com_ConvHTMLTag

End Function

'=====================================================================
' [Function		] ���t�ϊ�
' [Contents     ] ���t���w�肵���t�H�[�}�b�g�ɕϊ�����
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ConvDateString(strtmp,way,sp)

	Dim year
	Dim mon
	Dim date
	Dim hour
	Dim min
	Dim sec
	Dim ar

	If(sp = "") Then
		sp = "_"
	End If

	If(way = 1) Then
		ar = Split(strtmp,sp)
		year 	= ar(0)
		mon 	= ar(1)
		date 	= ar(2)
		hour 	= ar(3)
		min 	= ar(4)
		sec 	= ar(5)
		Com_ConvDateString = year & "�N" & mon & "��" & date & "��" & hour & "��"
	ElseIf(way = 2) Then
		ar = Split(strtmp,sp)
		year 	= ar(0)
		mon 	= ar(1)
		date 	= ar(2)
		hour 	= ar(3)
		min 	= ar(4)
		sec 	= ar(5)
		Com_ConvDateString = year & "�N" & mon & "��" & date & "��" & hour & "��" & min & "��"
	End If

End Function

'=====================================================================
' [Function		] META�^�O������擾
' [Contents     ] META�^�O���쐬����
'---------------------------------------------------------------------
' [Return       ] (string)<META...������
' [Input        ] (string)way:META�̎��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_PrintMETA(way)

	If(way = "euc") Then
		Com_PrintMETA =  "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=x-euc-jp"">"
	End If

End Function

'=====================================================================
' [Function		] ������t�H�[�}�b�g
' [Contents     ] ��������w��̃t�H�[�}�b�g�ɕϊ�����
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (integer)mode:���[�h
'				  (string)str:�ϊ����镶����
'				  (integer)cmd01:�R�}���h�P�i�����j
'				  (integer)cmd02:�R�}���h�Q�i���p���Ă܂���j
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_Sprint(mode,str,cmd01,cmd02)

	Dim Flg
	Dim i
	Dim j
	Dim k
	Dim strtmp
	Dim strarray

	' �S������������̐����� (��F0234 -> 234) 
	if(mode = 1) Then

		Com_Sprint = CInt(str)

	
	' ������������w�肵�����̕�����ɕϊ�
	ElseIf(mode = 2) Then
		j = cmd01 - Len(str)
		
		strtmp = str
		
		For i = 1 to j
			strtmp = "0" & strtmp
		Next

		Com_Sprint = strtmp

	End If

End Function

'=====================================================================
' [Function		] �T���h�C�b�`
' [Contents     ] ��������w�肵��������ł͂���
'---------------------------------------------------------------------
' [Return       ] (string)�͂��񂾕�����
' [Input        ] (integer)mode:���[�h
'				  (string)str:�͂��܂�镶����
'				  (string)substr:�͂��ޕ�����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_SandWitch(mode,str,substr)

	If(mode = 1) Then
		Com_SandWitch = substr &  str & substr
	End If

End Function

'=====================================================================
' [Function		] input�p�ɕ������ϊ�����
' [Contents     ] input�^�O�p�ɕ������ϊ�����
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (string)strtmp:�ϊ����镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ConvInputString(strtmp)

	Com_ConvInputString = Replace(strtmp,"_N_",vbCrLf)

End Function

'=====================================================================
' [Function		] �z�������w�蕶���񂪂��邩���肷��
' [Contents     ] �z�������w�蕶���񂪂��邩���肷��
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_StringMatchToArray(strF,strarray)

	Dim i
	Dim imax

	imax = UBound(strarray)

	For i = 0 to imax
		If((strarray(i) = strF) and (strF <> "")) Then
			Com_StringMatchToArray = 1
			Exit Function
		End If
	Next

	Com_StringMatchToArray = 0

End Function

'=====================================================================
' [Function		] �����̌��𐔎��̌��ɕϊ�����
' [Contents     ] Jan->1,Feb->2....
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (string)�ϊ����镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function ConvNumMon(strtmp)

	If(strtmp = "Jan") Then
		strtmp = "1"
	ElseIf(strtmp = "Feb") Then
		strtmp = "2"
	ElseIf(strtmp = "Mar") Then
		strtmp = "3"
	ElseIf(strtmp = "Apr") Then
		strtmp = "4"
	ElseIf(strtmp = "May") Then
		strtmp = "5"
	ElseIf(strtmp = "Jun") Then
		strtmp = "6"
	ElseIf(strtmp = "Jul") Then
		strtmp = "7"
	ElseIf(strtmp = "Aug") Then
		strtmp = "8"
	ElseIf(strtmp = "Sep") Then
		strtmp = "9"
	ElseIf(strtmp = "Oct") Then
		strtmp = "10"
	ElseIf(strtmp = "Nov") Then
		strtmp = "11"
	ElseIf(strtmp = "Dec") Then
		strtmp = "12"
	End If

	ConvNumMon = strtmp

End Function

'=====================================================================
' [Function		] �w��Z�N�V�����̍ő�L�[�̐������擾����
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string)�w��L�[�̍ő�l
' [Input        ] (integer)mode:���[�h
'				  (string)section:�Z�N�V����
'				  (string)hash:�A�z�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetMaxNumberFromINIHash(mode,section,hash)

	Dim i
	Dim tmp01
	Dim strtmp

	If(mode = 1) Then
		For i = 1 to 1000
			tmp01 = hash & "_" & section & "_" & i
			If(Session(tmp01) = "") Then
				Com_GetMaxNumberFromINIHash = i - 1
				Exit For
			End If
		Next
	End If

End Function

'=====================================================================
' [Function		] ���p�����ϊ�
' [Contents     ] �P�Q�R�S���T�U -> 123456
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (string)�ϊ����镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ConvPrice(str)

	Dim i
	Dim imax
	Dim ar
	Dim p

	str = Replace(str,"�O",0)
	str = Replace(str,"�P",1)
	str = Replace(str,"�Q",2)
	str = Replace(str,"�R",3)
	str = Replace(str,"�S",4)
	str = Replace(str,"�T",5)
	str = Replace(str,"�U",6)
	str = Replace(str,"�V",7)
	str = Replace(str,"�W",8)
	str = Replace(str,"�X",9)

	ar = Com_KSplit(str,"")
	imax = UBound(ar)

	p = ""
	For i = 0 to imax

		If(i > imax) Then Exit For

		If(IsNumeric(ar(i))) Then

			p = p & ar(i)

		End If
	Next

	Com_ConvPrice = p

End Function

'=====================================================================
' [Function		] �A�z�z�񂩂�hidden�^�O���쐬����
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string)hidden�^�O������
' [Input        ] (string)remove:hidden�ɓ���Ȃ��A�z�z��
'			      (string)hname:�A�z�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_MakeHiddenValueWithFilter(remove,hname)

	Dim i
	Dim imax
	Dim key
	Dim k
	Dim v
	Dim p
	Dim rem_flag
	Dim str
	Dim ar_remove

	If(Len(remove)) Then
		ar_remove = Split(remove,":")
		imax = UBound(ar_remove)
	Else
		imax = -1
	End If

	For Each key In Session.Contents

		If(InStr(key,hname & "_")) Then

			rem_flag = 0

			For i = 0 to imax
				If(InStr(key,ar_remove(i))) Then
					rem_flag = 1
					Exit For
				End If
			Next

			If(rem_flag = 0) Then

				p = Len(hname & "_") + 1

				k = Mid(key,p,Len(key) - p + 1)
				v = Session(key)

				str = str & "<input type=""hidden"" name=""" & k & """ value=""" & v & """>" & vbCrLf

			End If

		End If

	Next

	Com_MakeHiddenValueWithFilter = str

End Function

'=====================================================================
' [Function		] �����񔽓]
' [Contents     ] ��������t���܂ɂ���
'---------------------------------------------------------------------
' [Return       ] (string)���]����������
' [Input        ] (string)���]���镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ReverseString(str)

	Dim ar

	ar = Com_KSplit(str,"")
	ar = Com_ReverseArray(ar)

	Com_ReverseString = Join(ar,"")

End Function

'=====================================================================
' [Function		] �z�񔽓]
' [Contents     ] �z����t���܂ɂ���
'---------------------------------------------------------------------
' [Return       ] (string array)���]�����z��
' [Input        ] (string array)���]����z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ReverseArray(ar)

	Dim i
	Dim imax
	Dim ar2

	imax = UBound(ar)

	For i = imax to 0 Step - 1
		ar2(imax - i) = ar(i)
	Next

	Com_ReverseArray = ar2

End Function

'=====================================================================
' [Function		] URL�G���R�[�h
' [Contents     ] �������URL�G���R�[�h����
'---------------------------------------------------------------------
' [Return       ] (string)�G���R�[�h����������
' [Input        ] (string)�G���R�[�h���镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_URLEncode(str)

	Com_URLEncode = Server.URLEncode(str)

End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function URLDecode()
End Function

'=====================================================================
' [Function		] �A�z�z��\��
' [Contents     ] �w�肵���A�z�z����f�o�O�\������
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (string)hname:�A�z�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ListHash(hname)

	Dim key

	For Each key In Session.Contents
		If(InStr(key,hname & "_")) Then
			Com_Debug(Session(key))
		End If 
	Next

End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function GetCookie()
End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function PutCookie()
End Function

'=====================================================================
' [Function		] �w����t�l�擾
' [Contents     ] year->�N,mon->���̂悤�Ɉ����œ��t�̈ꕔ�����擾����
'---------------------------------------------------------------------
' [Return       ] (string)key�Ȃ�->time��ϊ�����������,����->���t�ꕔ��
' [Input        ] (now)time:���t
'				  (string)key:�擾������t�L�[
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetSpecDateString(time,key)

	Dim sec
	Dim min
	Dim hour
	Dim mday
	Dim mon
	Dim year
	Dim wday
	Dim yday
	Dim isdst
	Dim datestr
	Dim ar
	
	ar = Split(time,":")

	Session("f_year")	= ar(0)
	Session("f_mon")	= ar(1)
	Session("f_mday")	= ar(2)
	Session("f_hour")	= ar(3)
	Session("f_min")	= ar(4)
	Session("f_sec")	= ar(5)

	If(Session("sec") < 10) Then
		Session("sec") = "0" & Session("sec")
	End If

	If(Session("min") < 10) Then
		Session("min") = "0" & Session("min")
	End If

	If(Session("hour") < 10) Then
		Session("hour") = "0" & Session("hour")
	End If

	If(Session("mday") < 10) Then
		Session("mday") = "0" & Session("mday")
	End If

	If(Session("mon") < 10) Then
		Session("mon") = "0" & Session("mon")
	End If

	If(Session("year") < 99) Then
		Session("year") = Session("year") + 2000
	Else
		Session("year") = Session("year") + 1900
	End If

	datestr = Session("year") & ":" & Session("mon") & ":" & Session("mday") & ":" & Session("hour") & ":" & Session("min") & ":" & Session("sec")

	If(key = "") Then
		Com_GetSpecDateString = datestr
	Else
		Com_GetSpecDateString = Session("f_" & key)
	End If

End Function

'=====================================================================
' [Function		] "k=v"�̔z�񂩂�A�z�z����쐬����
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (string array)strarray:"k=v"�̔z��
'				  (string)hash:�A�z�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function MakeHashFromStrArray(strarray,hash)

	Dim i
	Dim imax
	Dim k
	Dim v
	Dim ar

	imax = UBound(strarray)

	For i = 0 to imax

		ar = Split(strarray(i),"=")
		k = ar(0)
		v = ar(1)

		Session(hash & "_" & k) = v

	Next

End Function

'=====================================================================
' [Function		] ���tSelection�쐬
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string)selection������
' [Input        ] (integer)mode:���[�h
'				  (now)date:�I���ς݂̓��t
'				  (string)sname:select��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_MakeDateSelection(mode, date, sname)

	Dim y_i
	Dim m_i
	Dim d_i
	Dim cur_y
	Dim cur_m
	Dim cur_d
	Dim y_max
	Dim m_max
	Dim d_max
	Dim pdate
	Dim py
	Dim pm
	Dim pd
	Dim selecteddate
	Dim selected
	Dim strtmp

	cur_y = Com_GetSpecDateString(date,"year")
	cur_m = Com_GetSpecDateString(date,"mon")
	cur_d = Com_GetSpecDateString(date,"mday")

	y_max = cur_y + 1
	m_max = 12
	d_max = 31

	if(Com_pStr = "") Then
		selecteddate = cur_y & cur_m & cur_d
	Else
		selecteddate = Com_pStr
	End If

	If(mode = 1) Then
		For y_i = cur_y to y_max
			For m_i = 1 to m_max
				For d_i = 1 to d_max

					py = Com_Sprint(2,y_i,4,"")
					pm = Com_Sprint(2,m_i,2,"")
					pd = Com_Sprint(2,d_i,2,"")

					If(selecteddate = py & pm & pd) Then
						selected = " selected"
					Else
						selected = ""
					End If

					pdate = py & "/" & pm & "/" & pd
					strtmp = strtmp & "<option value=""" & pdate & """" & selected & ">" & pdate & "</option>" & vbCrLf
				Next
			Next
		Next

		strtmp = "<select name=""sname"">" & strtmp & "</select>"

		 Com_MakeDateSelection = strtmp
	End If

End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function CheckFileDateAndDelete()
End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function KstdSendMail()
End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function CheckIsNULLInHashWithFilter()
End Function

'=====================================================================
' [Function		] �l�i������ϊ�
' [Contents     ] 1000000 -> 1,000,000�ɕϊ�����
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (string)num:�ϊ����镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ConvPriceComma(num)

	Dim i
	Dim l
	Dim str
	Dim imax
	Dim ar
	Dim cnt

	cnt = 0
	If(IsNumeric(num)) Then
		num = CStr(num)
		ar = Com_KSplit(num,"")
		imax = UBound(ar)

		For i = imax to 0 Step -1

			If((cnt Mod 3) = 0 And (cnt <> 0)) Then
				str = "," & str
			End If

			str = ar(i) & str
			cnt = cnt + 1

		Next

		Com_ConvPriceComma = str

	End If

End Function

'=====================================================================
' [Function		] 
' [Contents     ] 
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function GetDltDateString()
End Function

'=====================================================================
' [Function		] ���~
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string)�� or �~
' [Input        ] (integer or string) 1 or 0 �܂��͕����񂪂��邩�Ȃ���
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_MaruBatu(q)

	If(q) Then
		Com_MaruBatu = "��"
	Else
		Com_MaruBatu = "�~"
	End If

End Function

'=====================================================================
' [Function		] 01
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string) 0 or 1
' [Input        ] (integer or string) 1 or 0 �܂��͕����񂪂��邩�Ȃ���
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_IchiZero(q)

	If(q) Then
		Com_IchiZero = "0"
	Else
		Com_IchiZero = "1"
	End If

End Function

'=====================================================================
' [Function		] �z�������������
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string array)�����������z��
' [Input        ] (string array)����������z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_SuffleArray(oldarray)

	Dim i
	Dim imax
	Dim newarray
	Dim tmparray


	newarray = Array()
	imax = UBound(oldarray)

	Randomize
	For i = 0 to imax
		tmparray = Com_ArraySplice(oldarray, Int(Rnd * UBound(oldarray)), 1)
		Com_ArrayPush newarray , tmparray
	Next

	Com_SuffleArray = newarray

End Function

'=====================================================================
' [Function		] �t�@�C��������content-type���擾����
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string) content-type�̕�����
' [Input        ] (string)fname:�t�@�C����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_DetectFileType(fname)

	Dim fext

	fext = Com_GetFileExtension(fname)

	If(fext = "gif") Then

		Com_DetectFileType = "image/gif"

	ElseIf((fext = "jpeg") or (f_type = "jpg")) Then

		Com_DetectFileType = "image/jpeg"

	ElseIf(fext = "bmp") Then

		Com_DetectFileType = "image/bmp"

	Else

		Com_DetectFileType = "application/octet-stream"

	End If

End Function

'=====================================================================
' [Function		] �t�@�C��������g���q���擾����
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string)�g���q
' [Input        ] (string)fname:�t�@�C����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetFileExtension(fname)


	Dim pFobj
	Dim strtmp

	Com_GetFileExtension = ""

	Set pFobj = Server.CreateObject("Scripting.FileSystemObject")

	Com_GetFileExtension = pFobj.GetExtensionName(fname)

	Set pFobj = Nothing

End Function

'=====================================================================
' [Function		] �X�g�b�v�E�H�b�`
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] �Q��ڂ���͑O��Ăяo�������Ƃ̍��i�b�P�ʁj
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_StopWatchVer1()

	If(IsEmpty(Com_pNow)) Then
		Com_pNow = Now
	Else
		Com_StopWatchVer1 = DateDiff("s",Com_pNow,Now)
	End If

End Function

'=====================================================================
' [Function		] �z��𑫂�
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (byref string array)ar:�ǉ�����Ĕz��
'				  (byref string array)str:�ǉ����镶����܂��͔z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ArrayPush(ByRef ar, ByRef str)

	Dim i
	Dim j
	Dim jmax

	If(IsEmpty(ar)) Then
		i = 0
	Else
		i = UBound(ar) + 1
	End If

	If(IsArray(str)) Then

		jmax = UBound(str)
		For j = 0 to jmax
			ReDim Preserve ar(i)
			ar(i) = str(j)
			i = i + 1
		Next

	Else

		ReDim Preserve ar(i)
		ar(i) = str

	End If

End Function

'=====================================================================
' [Function		] �z�񂩂�v�f����菜��
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (byref string array)ar:��菜�����z��
'				  (integer)pos:��菜���J�n�ʒu
'				  (integer)len:��菜����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_ArraySplice(ByRef ar, pos, len)

	Dim i
	Dim imax
	Dim ar2cnt
	Dim ar3cnt
	Dim ar2()
	Dim ar3()

	imax = UBound(ar)

	ar2cnt = 0
	ar3cnt = 0

	'��
	For i = 0 to pos - 1
		ReDim Preserve ar2(ar2cnt)
		ar2(ar2cnt) = ar(i)
		ar2cnt = ar2cnt + 1
	Next

	'��镔��
	For i = pos to pos + len - 1
		ReDim Preserve ar3(ar3cnt)
		ar3(ar3cnt) = ar(i)
		ar3cnt = ar3cnt + 1
	Next

	'�E
	For i = pos + len to imax
		ReDim Preserve ar2(ar2cnt)
		ar2(ar2cnt) = ar(i)
		ar2cnt = ar2cnt + 1
	Next

	If(ar2cnt > 0) Then
		ReDim ar(UBound(ar2))
		ar = ar2
	End If

	Com_ArraySplice = ar3

End Function

'=====================================================================
' [Function		] �����z��擾
' [Contents     ] ����
'---------------------------------------------------------------------
' [Return       ] (string array)�z��
' [Input        ] (integer)num:��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetRandomArray(num)

	Dim i
	Dim ar

	ReDim ar(num - 1)

	For i = 0 to num - 1
		ar(i) = CStr(i + 1)
	Next

	Com_GetRandomArray = ar

End Function

%>