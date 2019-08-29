<%

Dim Com_pSelected	'selected文字列
Dim Com_pNow		'一時時間
Dim Com_pStr		'一時文字列
Dim Com_pInt		'一時整数
Dim DEBUG_ : DEBUG_ = 1

'=====================================================================
' [Function		] <select>タグ作成
' [Contents     ] 配列からselectタグを作成する
'---------------------------------------------------------------------
' [Return       ] <select...文字列
' [Input        ] (integer)mode:作成モード
'				  (string)selname:selection名
'				  (string array)strarray:配列１
'				  (string array)strarray02:配列２
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 重複回避
' [Contents     ] 配列から重複をなくす
'---------------------------------------------------------------------
' [Return       ] (string array)重複をなくした配列
' [Input        ] (string array)strarray:配列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
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
' [Contents     ] 配列の先頭に値を追加する
'---------------------------------------------------------------------
' [Return       ] (string array)追加した配列
' [Input        ] (string array)arr:配列
'				  (string)val:追加する値
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] アンダーバー置換
' [Contents     ] "_文字列_"をSession内の値に変換する
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (string)変換する前の文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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

		'カウンタが文字位置を越えたら終わる
		If(i > Len(strtmp)) Then Exit Do

		'一文字とる
		s = Mid(strtmp,i,1)

		'文字が "_" だった場合、次のアンダーバーまで文字を pVal に格納する
		If(s = "_") Then

			'格納ループ --- ここから

			pVal = ""
			For i2 = i + 1 to l
				s = Mid(strtmp,i2,1)

				If(s = "_") Then
					'Response.Write(pVal & "でした<br>")
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

					'Response.Write("[" & s & "]なのでおわり<br>")
					Exit For

				End If
			Next

			'格納ループ --- ここまで

		End If

		i = i + 1

	Loop

	Com_ConvVal = strtmp

End Function

'=====================================================================
' [Function		] 文字列を指定文字で区切って配列で返す
' [Contents     ] 文字列を指定文字で区切って配列で返す
'---------------------------------------------------------------------
' [Return       ] (string array)区切った配列
' [Input        ] (string)strtmp:区切る前の文字列
'				  (string)delimiter):区切り文字
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] パラパラ
' [Contents     ] 色を交互にする
'---------------------------------------------------------------------
' [Return       ] (string)現在の色
' [Input        ] (string)標準以外の色
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 配列表示
' [Contents     ] 配列を改行つきで表示する
'---------------------------------------------------------------------
' [Return       ] 
' [Input        ] (integer)mode:表示モード
'				  (string array)strarray:配列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] デバグ表示
' [Contents     ] デバグする時に使う関数 DEBUG_変数が１の時に表示
'---------------------------------------------------------------------
' [Return       ] 
' [Input        ] (string)str:表示文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_Debug(str)
	If(DEBUG_) Then Response.Write(str & "<br>" & vbCrLf)
End Function

'=====================================================================
' [Function		] エラー表示
' [Contents     ] エラーメッセージを表示して終了する
'---------------------------------------------------------------------
' [Return       ] 
' [Input        ] (string)str:表示文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_Err001(str)
	Response.Write(str)
	Response.End
End Function

'=====================================================================
' [Function		] 日付取得
' [Contents     ] YYYY:MM:DD:HH:MM:SS形式で取得する
'---------------------------------------------------------------------
' [Return       ] (string) YYYY:MM:DD:HH:MM:SS文字列
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] ヘッダー出力
' [Contents     ] Perlの時よく使った関数
'---------------------------------------------------------------------
' [Return       ] (String)ヘッダー文字列
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function PH()

	PH = ""

End Function

'=====================================================================
' [Function		] Refresh画面表示
' [Contents     ] METAにRefreshの入ったページを表示する
'---------------------------------------------------------------------
' [Return       ] (string)HTML文字列
' [Input        ] (string)url:リンク先URL
'				  (string)sec:秒数
'				  (string)str:html文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_MakeRefresh(url,sec,str)

	Com_MakeRefresh = "<html><head><title>" & str & "</title><META HTTP-EQUIV='Content-Type' Content=""text/html; charset=Shift_JIS""><META HTTP-EQUIV='Refresh' Content=""" & sec & " ; url='" & url & "'""></head><body>" & str & "</body></html>"

End Function

'=====================================================================
' [Function		] HTML禁止タグの置換
' [Contents     ] <や>などのタグを回避する
'---------------------------------------------------------------------
' [Return       ] (string)置換した文字列
' [Input        ] (string)置換する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 日付変換
' [Contents     ] 日付を指定したフォーマットに変換する
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
		Com_ConvDateString = year & "年" & mon & "月" & date & "日" & hour & "時"
	ElseIf(way = 2) Then
		ar = Split(strtmp,sp)
		year 	= ar(0)
		mon 	= ar(1)
		date 	= ar(2)
		hour 	= ar(3)
		min 	= ar(4)
		sec 	= ar(5)
		Com_ConvDateString = year & "年" & mon & "月" & date & "日" & hour & "時" & min & "分"
	End If

End Function

'=====================================================================
' [Function		] METAタグ文字列取得
' [Contents     ] METAタグを作成する
'---------------------------------------------------------------------
' [Return       ] (string)<META...文字列
' [Input        ] (string)way:METAの種類
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_PrintMETA(way)

	If(way = "euc") Then
		Com_PrintMETA =  "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=x-euc-jp"">"
	End If

End Function

'=====================================================================
' [Function		] 文字列フォーマット
' [Contents     ] 文字列を指定のフォーマットに変換する
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (integer)mode:モード
'				  (string)str:変換する文字列
'				  (integer)cmd01:コマンド１（桁数）
'				  (integer)cmd02:コマンド２（利用してません）
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_Sprint(mode,str,cmd01,cmd02)

	Dim Flg
	Dim i
	Dim j
	Dim k
	Dim strtmp
	Dim strarray

	' ４桁数字文字列の数字化 (例：0234 -> 234) 
	if(mode = 1) Then

		Com_Sprint = CInt(str)

	
	' 数字文字列を指定した桁の文字列に変換
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
' [Function		] サンドイッチ
' [Contents     ] 文字列を指定した文字列ではさむ
'---------------------------------------------------------------------
' [Return       ] (string)はさんだ文字列
' [Input        ] (integer)mode:モード
'				  (string)str:はさまれる文字列
'				  (string)substr:はさむ文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_SandWitch(mode,str,substr)

	If(mode = 1) Then
		Com_SandWitch = substr &  str & substr
	End If

End Function

'=====================================================================
' [Function		] input用に文字列を変換する
' [Contents     ] inputタグ用に文字列を変換する
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (string)strtmp:変換する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ConvInputString(strtmp)

	Com_ConvInputString = Replace(strtmp,"_N_",vbCrLf)

End Function

'=====================================================================
' [Function		] 配列内から指定文字列があるか判定する
' [Contents     ] 配列内から指定文字列があるか判定する
'---------------------------------------------------------------------
' [Return       ] ()
' [Input        ] ()
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 文字の月を数字の月に変換する
' [Contents     ] Jan->1,Feb->2....
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (string)変換する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 指定セクションの最大キーの数字を取得する
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string)指定キーの最大値
' [Input        ] (integer)mode:モード
'				  (string)section:セクション
'				  (string)hash:連想配列名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 半角数字変換
' [Contents     ] １２３４あ５６ -> 123456
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (string)変換する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ConvPrice(str)

	Dim i
	Dim imax
	Dim ar
	Dim p

	str = Replace(str,"０",0)
	str = Replace(str,"１",1)
	str = Replace(str,"２",2)
	str = Replace(str,"３",3)
	str = Replace(str,"４",4)
	str = Replace(str,"５",5)
	str = Replace(str,"６",6)
	str = Replace(str,"７",7)
	str = Replace(str,"８",8)
	str = Replace(str,"９",9)

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
' [Function		] 連想配列からhiddenタグを作成する
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string)hiddenタグ文字列
' [Input        ] (string)remove:hiddenに入れない連想配列名
'			      (string)hname:連想配列名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 文字列反転
' [Contents     ] 文字列を逆さまにする
'---------------------------------------------------------------------
' [Return       ] (string)反転した文字列
' [Input        ] (string)反転する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ReverseString(str)

	Dim ar

	ar = Com_KSplit(str,"")
	ar = Com_ReverseArray(ar)

	Com_ReverseString = Join(ar,"")

End Function

'=====================================================================
' [Function		] 配列反転
' [Contents     ] 配列を逆さまにする
'---------------------------------------------------------------------
' [Return       ] (string array)反転した配列
' [Input        ] (string array)反転する配列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] URLエンコード
' [Contents     ] 文字列をURLエンコードする
'---------------------------------------------------------------------
' [Return       ] (string)エンコードした文字列
' [Input        ] (string)エンコードする文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function URLDecode()
End Function

'=====================================================================
' [Function		] 連想配列表示
' [Contents     ] 指定した連想配列をデバグ表示する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (string)hname:連想配列名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function PutCookie()
End Function

'=====================================================================
' [Function		] 指定日付値取得
' [Contents     ] year->年,mon->月のように引数で日付の一部分を取得する
'---------------------------------------------------------------------
' [Return       ] (string)keyなし->timeを変換した文字列,あり->日付一部分
' [Input        ] (now)time:日付
'				  (string)key:取得する日付キー
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] "k=v"の配列から連想配列を作成する
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (string array)strarray:"k=v"の配列
'				  (string)hash:連想配列名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 日付Selection作成
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string)selection文字列
' [Input        ] (integer)mode:モード
'				  (now)date:選択済みの日付
'				  (string)sname:select名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function CheckIsNULLInHashWithFilter()
End Function

'=====================================================================
' [Function		] 値段文字列変換
' [Contents     ] 1000000 -> 1,000,000に変換する
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (string)num:変換する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function GetDltDateString()
End Function

'=====================================================================
' [Function		] ○×
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string)○ or ×
' [Input        ] (integer or string) 1 or 0 または文字列があるかないか
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_MaruBatu(q)

	If(q) Then
		Com_MaruBatu = "○"
	Else
		Com_MaruBatu = "×"
	End If

End Function

'=====================================================================
' [Function		] 01
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string) 0 or 1
' [Input        ] (integer or string) 1 or 0 または文字列があるかないか
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 配列をかき混ぜる
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string array)かき混ぜた配列
' [Input        ] (string array)かき混ぜる配列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] ファイル名からcontent-typeを取得する
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string) content-typeの文字列
' [Input        ] (string)fname:ファイル名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] ファイル名から拡張子を取得する
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string)拡張子
' [Input        ] (string)fname:ファイル名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] ストップウォッチ
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] ２回目からは前回呼び出した時との差（秒単位）
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 配列を足す
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (byref string array)ar:追加されて配列
'				  (byref string array)str:追加する文字列または配列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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
' [Function		] 配列から要素を取り除く
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (byref string array)ar:取り除かれる配列
'				  (integer)pos:取り除く開始位置
'				  (integer)len:取り除く個数
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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

	'左
	For i = 0 to pos - 1
		ReDim Preserve ar2(ar2cnt)
		ar2(ar2cnt) = ar(i)
		ar2cnt = ar2cnt + 1
	Next

	'取る部分
	For i = pos to pos + len - 1
		ReDim Preserve ar3(ar3cnt)
		ar3(ar3cnt) = ar(i)
		ar3cnt = ar3cnt + 1
	Next

	'右
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
' [Function		] 乱数配列取得
' [Contents     ] 同じ
'---------------------------------------------------------------------
' [Return       ] (string array)配列
' [Input        ] (integer)num:個数
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
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