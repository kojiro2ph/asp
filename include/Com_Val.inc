<%

Dim ThisFile : ThisFile = Request.ServerVariables("SCRIPT_NAME")

'=====================================================================
' [Function		] フォームデータ初期化
' [Contents     ] Request.FomまたはRequest.QueryString変数をSessionに格納する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_Init_Form()

	'フォームデータごとの処理
	If(Request.ServerVariables("REQUEST_METHOD") = "POST") Then
		For Each header In Request.Form
			Session("f_" & header) = Request.Form(header)
		Next
	Else
		For Each header In Request.QueryString
			Session("f_" & header) = Request.QueryString(header)
		Next
	End If

End Function

'=====================================================================
' [Function		] フォームSession変数クリア
' [Contents     ] Sessionに格納したForm変数を空にする
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_f_SessionClear()

	Dim i
	Dim j
	Dim str

	For i = 1 to Session.Contents.Count

		'F_をf_に変換する（iis対応）
		str = Session.Contents.Key(i)
		str = Replace(str,"F_","f_")

		j = InStr(1,str,"f_") 'ここのf_に注意！

		If(j = 1) Then
			Session(Session.Contents.Key(i)) = ""
		End If

	Next

End Function

'=====================================================================
' [Function		] 設定ファイル初期化
' [Contents     ] 設定ファイルをSession変数に格納する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (integer) mode : 初期化モード
'				  (string) inifile : 設定ファイルパス
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_InitINIData(mode, inifile)

	Dim f,t,b,ar,k,v,n,sec

	'ノーマルモード■■■■■■■■■■■■■■■■■■■■■■■
	If(mode = 1) Then

		Set f = Server.CreateObject("Scripting.FileSystemObject")
		Set t = f.OpenTextFile(inifile)

		'行ループ
		Do Until t.AtEndOfStream

			'一行読み込み ----
			b = t.ReadLine

			'"[" を探す ----
			n = InStr(1, b, "[")
			'Response.Write(n & "dao<br>" & vbCrLf)
			'Response.Write(Server.HTMLEncode(b) & "をチェックします<br>" & vbCrLf)

			If(b = "") Then
				n = -1
			End If

			'セクション区切りの場合 ===========================
			If(n = 1) Then
				'Response.Write("セクションです<br>" & vbCrLf)
				b = Right(b,len(b) - 1)
				b = Left(b,len(b) - 1)

				sec = b
				'Response.Write(sec & "セクションです<br>" & vbCrLf)

			ElseIf(n = -1) Then

				'Response.Write("空です<br>" & vbCrLf)

			'セクション内部の場合 ===========================
			Else
				'[=]で区切る ----
				ar = Split(b,"=")

				'両端の空白をとる ----
				k = Trim(ar(0))
				k = Replace(k,vbTab,"")
				v = Trim(ar(1))

				'両端に["]があった場合削除する ----
				n = InStr(1, v, """")
				If(n = 1) Then
					v = Right(v,len(v) - 1)
				End If
				n = InStrRev(v, """")
				If(n <> 0) Then
					v = Left(v,len(v) - 1)
				End If

				Session("INI_" & sec & "_" & k) = v

				'Response.Write(n & "<br>" & vbCrLf)
				'Response.Write(sec & "_" & k & "の値は" & v & "です<br>" & vbCrLf)
			End If
		Loop

		'ファイルを閉じる ----
		t.close

		'オブジェクトの参照をなくす ----
		Set f = Nothing
		Set t = Nothing
	End If

End Function

'=====================================================================
' [Function		] 設定ファイル更新
' [Contents     ] 指定した設定ファイルの指定セクションの指定キーを更新する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (integer)mode:更新モード
'				  (string)fname:ファイル名
'				  (string)section:更新するセクション名
'				  (string)strFkey:更新するキー名
'				  (string)strCval:更新する値
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_UpdateINIFile(mode,fname,section,strFkey,strCval)

	Dim i
	Dim imax
	Dim StrCurSection
	Dim strline
	Dim strarray

	strline = Com_ReadFileData(1,fname)
	strarray = Split(strline,vbCrLf)

	imax = UBound(strarray)
	For i = 0 to imax
		If(InStr(strarray(i),"[") and InStr(strarray(i),"]")) Then
			StrCurSection = strarray(i)
			StrCurSection = Replace(StrCurSection,"[","")
			StrCurSection = Replace(StrCurSection,"]","")
		ElseIf(InStr(strarray(i),strFkey) and StrCurSection = section) Then
			strFval = section & "-" & strFkey
			strarray(i) = Replace(strarray(i),"= " & Session("INI_" & strFval),"= " & strCval)
		End If
	Next
	
	strline = Join(strarray,vbCrLf)
	Com_RecordFileData 1,fname,strline

End Function

'=====================================================================
' [Function		] NTユーザーログイン名取得
' [Contents     ] アクセスしてきたLogin名を取得する
'---------------------------------------------------------------------
' [Return       ] (string)ログイン名
' [Input        ] 
'				  
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function GetMyName()

	Dim name

	if(Session("f_MyName") = "") then
		name = Request.ServerVariables("LOGON_USER")
		Session("f_MyName") = Mid(name,InStr(1,name,"\") + 1)
		GetMyName = Mid(name,InStr(1,name,"\") + 1)
	else
		GetMyName = Session("f_MyName")
	end if

End Function

'=====================================================================
' [Function		] タグ初期化
' [Contents     ] よく使うタグをSessionに格納する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_Init_Tag()

	Session("f_f1") 				= "<font size=""1"">"
	Session("f2") 					= "<font size=""2"">"
	Session("f_f3") 				= "<font size=""3"">"
	Session("f_f4") 				= "<font size=""4"">"
	Session("f_f5") 				= "<font size=""5"">"
	Session("f_f6") 				= "<font size=""6"">"
	Session("f_f7") 				= "<font size=""7"">"
	Session("f_f8") 				= "<font size=""8"">"
	Session("f_fred")  				= "<font color=""red"">"
	Session("f_fpink")  			= "<font color=""pink"">"
	Session("f_fblue")  			= "<font color=""blue"">"
	Session("f_fwhite")  			= "<font color=""white"">"
	Session("f_fblack")  			= "<font color=""black"">"
	Session("f_flblue")  			= "<font color=""lblue"">"
	Session("f_fb")  				= "<b>"
	Session("f_fi")  				= "<i>"
	Session("f_fc")  				= "</font>"

	Session("f_al")  				= "align=""left"""
	Session("f_ar")  				= "align=""right"""
	Session("f_am")  				= "align=""center"""
	Session("f_ac")  				= "align=""center"""
	Session("f_vab") 				= "valign=""bottom"""
	Session("f_vat") 				= "valign=""top"""
	Session("f_vam") 				= "valign=""middle"""

	Session("f_t")					= "<table>"
	Session("f_tc") 				= "</table>"

	Session("f_htdoc")				= "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2 Final//EN"">"

End Function

%>


