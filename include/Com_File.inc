<%

'=====================================================================
' [Function		] ファイル読み込み
' [Contents     ] ファイル文字列を返す
'---------------------------------------------------------------------
' [Return       ] (string)ファイル文字列
' [Input        ] (int) mode : 取得モード
'				  (string) fpath : ファイルパス
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ReadFileData(mode,fpath)

	Dim pFobj
	Dim pFhdl
	Dim strtmp

	Com_ReadFileData = ""

	Set pFobj = Server.CreateObject("Scripting.FileSystemObject")

	If Not(pFobj.FileExists(fpath)) Then Exit Function

	Set pFhdl = pFobj.OpenTextFile(fpath,1)

	If(mode = 1) Then
		strtmp = pFhdl.ReadAll
	End If

	pFhdl.close

	Set pFobj = Nothing
	Set pFhdl = Nothing

	Com_ReadFileData = strtmp

End Function

'=====================================================================
' [Function		] ファイル書き込み
' [Contents     ] ファイル書き込み
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] (int) mode : 書き込みモード
'				  (string) fname : ファイル名
'				  (string) str : 文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_RecordFileData(mode,fname,str)

	Dim ObjFSO				'ファイルオブジェクト
	Dim ObjTS				'テキストオブジェクト

	If(mode = 1) Then

		Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set ObjTS = ObjFSO.OpenTextFile(fname,2,True)
		ObjTS.WriteLine(str)
		ObjTS.Close

		Set ObjFSO = Nothing
		Set ObjTS = Nothing

	End If

End Function

'=====================================================================
' [Function		] ディレクトリ一覧取得
' [Contents     ] 指定ディレクトリの一覧を配列で返す
'---------------------------------------------------------------------
' [Return       ] (string array)ファイル配列
' [Input        ] (int) mode : 取得モード
'				  (string) path : ディレクトリパス
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetFileArray(mode,path)

	Dim i
	Dim n
	Dim ar()
	Dim ObjFSO
	Dim ObjFolder
	Dim ColFiles

	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set ObjFolder = ObjFSO.GetFolder(path)
	Set ColFiles = ObjFolder.Files

	i = 0
	If(mode = 1) Then
		For Each n in ColFiles
			ReDim Preserve ar(i)
			ar(i) = ObjFSO.GetFileName(n)
			i = i + 1
		Next
	End If

	Com_GetFileArray = ar

	Set ColFiles = Nothing
	Set ObjFolder = Nothing
	Set ObjFSO = Nothing

End Function

'=====================================================================
' [Function		] ＣＳＶファイルからＨＴＭＬテーブルを作成する
' [Contents     ] ＣＳＶファイルからＨＴＭＬテーブルを作成する
'---------------------------------------------------------------------
' [Return       ] (string)テーブル
' [Input        ] (integer)mode:表示モード
'				  (string)fname:ファイルパス
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Dim Com_MakeCSVHtmlTable_border 	: Com_MakeCSVHtmlTable_border 	= "1"
Dim Com_MakeCSVHtmlTable_width 		: Com_MakeCSVHtmlTable_width 	= "100%"
Dim Com_MakeCSVHtmlTable_ColTr 		: Com_MakeCSVHtmlTable_ColTr 	= "#EFE8EF"
Dim Com_MakeCSVHtmlTable_ColHeader 	: Com_MakeCSVHtmlTable_ColHeader = "pink"
Dim Com_MakeCSVHtmlTable_UseColor 	: Com_MakeCSVHtmlTable_UseColor = 1
Dim Com_MakeCSVHtmlTable_FontSize 	: Com_MakeCSVHtmlTable_FontSize = "12px"
Dim Com_MakeCSVHtmlTable_NoWrap 	: Com_MakeCSVHtmlTable_NoWrap 	= "nowrap"
Dim Com_MakeCSVHtmlTable_TdNoData 	: Com_MakeCSVHtmlTable_TdNoData = "&nbsp;"
Function Com_MakeCSVHtmlTable(mode,fname)

	Dim str
	Dim strlines
	Dim strline
	Dim i
	Dim imax
	Dim j
	Dim jmax
	Dim col
	Dim colmax
	Dim ar_str
	Dim strtable
	Dim coltr

	'通常モード（テーブル出力）
	If(mode = 1) Then

		'ファイルデータ読み込み
		strlines = Com_ReadFileData(1,fname)

		'一行ごとに区切る
		strline = Split(strlines,vbCrLf)

		'行数を取得する
		imax = UBound(strline)


		'最大項目数を取得する
		colmax = 0
		For i = 0 to imax
			If(strline(i) <> "") Then
				ar_str = Com_ConvCSVtoNormal(strline(i))
				col = UBound(ar_str)
				If(col > colmax) Then
					colmax = col
				End If
			End If
		Next

		'<table>タグ作成
		strtable = "<table border='" & Com_MakeCSVHtmlTable_border & "' width='" & Com_MakeCSVHtmlTable_width & "' style='font-size:" & Com_MakeCSVHtmlTable_FontSize & ";'>" & vbCrLf

		'一行ごとのループ（<tr><td>..</td></tr>の作成）
		For i = 0 to imax
			If(strline(i) <> "") Then

				'行の色設定
				If(Com_MakeCSVHtmlTable_UseColor) Then
					If(i = 0) Then
						coltr = "bgcolor='" & Com_MakeCSVHtmlTable_ColHeader & "'"
					Else
						coltr = "bgcolor='" & Com_Parapara(Com_MakeCSVHtmlTable_ColTr) & "'"
					End If
				End If

				strtable = strtable & "<tr " & coltr & ">" & vbCrLf

				ar_str = Com_ConvCSVtoNormal(strline(i))

				jmax = UBound(ar_str)
				For j = 0 to colmax
					If(j <= jmax) Then
						str = ar_str(j)
					Else
						str = Com_MakeCSVHtmlTable_TdNoData
					End If

					strtable = strtable & "<td " & Com_MakeCSVHtmlTable_NoWrap & ">" & str & "</td>" & vbCrLf
				Next

				strtable = strtable & "</tr>" & vbCrLf

			End If
		Next

		strtable = strtable & "</table>" & vbCrLf


		Com_MakeCSVHtmlTable = strtable

	End If

End Function

'=====================================================================
' [Function		] xls からＨＴＭＬテーブルを作成する
' [Contents     ] xls からＨＴＭＬテーブルを作成する
'---------------------------------------------------------------------
' [Return       ] (string)テーブル
' [Input        ] (integer)mode:表示モード
'				  (string)fname:ファイルパス
'				  (tname)tname:エリア名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Dim Com_MakeExcelHtmlTable_border : Com_MakeExcelHtmlTable_border = "1"
Dim Com_MakeExcelHtmlTable_width : Com_MakeExcelHtmlTable_width = "100%"
Dim Com_MakeExcelHtmlTable_ColTr : Com_MakeExcelHtmlTable_ColTr = "#EFE8EF"
Dim Com_MakeExcelHtmlTable_ColHeader : Com_MakeExcelHtmlTable_ColHeader = "pink"
Dim Com_MakeExcelHtmlTable_UseColor : Com_MakeExcelHtmlTable_UseColor = 1
Dim Com_MakeExcelHtmlTable_FontSize : Com_MakeExcelHtmlTable_FontSize = "12px"
Dim Com_MakeExcelHtmlTable_NoWrap : Com_MakeExcelHtmlTable_NoWrap = "nowrap"
Dim Com_MakeExcelHtmlTable_TdNoData : Com_MakeExcelHtmlTable_TdNoData = "&nbsp;"
Function Com_MakeExcelHtmlTable(mode,fname,tname)

	Dim i
	Dim coltr
	Dim strDSN
	Dim objConn
	Dim objRst
	Dim objField
	Dim strtable

	If(mode = 1) Then

		strDSN = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & fname

		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open strDSN

		Set objRst = Server.CreateObject("ADODB.Recordset")
		objRst.Open tname, objConn

		'<table>タグ作成
		strtable = "<table border='" & Com_MakeExcelHtmlTable_border & "' width='" & Com_MakeExcelHtmlTable_width & "' style='font-size:" & Com_MakeExcelHtmlTable_FontSize & ";'>" & vbCrLf

		' 各行の表示
		i = 0
		Do While Not objRst.EOF

			'行の色設定
			If(Com_MakeExcelHtmlTable_UseColor) Then
				If(i = 0) Then
					coltr = "bgcolor='" & Com_MakeExcelHtmlTable_ColHeader & "'"
				Else
					coltr = "bgcolor='" & Com_Parapara(Com_MakeExcelHtmlTable_ColTr) & "'"
				End If
			End If

			strtable = strtable & "<tr " & coltr & ">" & vbCrLf


			'項目ごとのループ
			For Each objField In objRst.Fields
					strtable = strtable & "<td " & Com_MakeExcelHtmlTable_NoWrap & ">" & objField & "</td>" & vbCrLf
			Next

			strtable = strtable & "</tr>" & vbCrLf

			i = i + 1
			objRst.MoveNext
		Loop

		strtable = strtable & "</table>" & vbCrLf



		Com_MakeExcelHtmlTable = strtable


		' クローズ
		objRst.Close
		objConn.Close

	End If

End Function

'=====================================================================
' [Function		] 一行のＣＳＶデータをカンマ区切りで分けて配列で返す
' [Contents     ] 一行のＣＳＶデータをカンマ区切りで分けて配列で返す
'---------------------------------------------------------------------
' [Return       ] (string array)配列
' [Input        ] (string)ＣＳＶ一行
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Dim Com_ConvCSVtoNormal_Delimiter : Com_ConvCSVtoNormal_Delimiter = ","
Function Com_ConvCSVtoNormal(strtmp)

	Dim k
	Dim kmax
	Dim j
	Dim jmax
	Dim l
	Dim c
	Dim InValFlag
	Dim strarray()
	Dim chararray
	Dim valarray

	InValFlag = 0

	strtmp = Replace(strtmp,"""""","_DBLAP_")
	chararray = Com_KSplit(strtmp,"")
	jmax = UBound(chararray)
	For j = 0 to jmax
		c = chararray(j)
		If(c = """") Then
			If(InValFlag = 0) Then
				InValFlag = 1
			ElseIf(InValFlag = 1) Then
				InValFlag = 0
			End If
		Else
			If(c = ",") Then
				If(InValFlag = 0) Then
				ElseIf(InValFlag = 1) Then
					chararray(j) = "_VALKAM_"
				End If
			End If
		End If
	Next

	strtmp = Join(chararray,"")
	strtmp = Replace(strtmp,"""","")


	valarray = Split(strtmp,Com_ConvCSVtoNormal_Delimiter)
	kmax = UBound(valarray)
	For k = 0 to kmax
		valarray(k) = Replace(valarray(k),"_DBLAP_","""")
		valarray(k) = Replace(valarray(k),"_VALKAM_",",")

		ReDim Preserve strarray(k)
		strarray(k) = valarray(k)
	Next

	Com_ConvCSVtoNormal = strarray

End Function

'=====================================================================
' [Function		] SessionからCSV行を作成
' [Contents     ] Session連想配列からCSVを一行作成する
'---------------------------------------------------------------------
' [Return       ] (string)CSV一行
' [Input        ] (string)mode:モード
'				  (string)csv:連想配列の名前
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_MakeCSVLineFromHash(mode,csv)

	Dim strtmp
	Dim i
	Dim tmp01
	Dim tmp02

	If(mode = 1) Then
		strtmp = ""
		strtmp = "<table border='1'>"

		For i = 0 to Session(csv & "_" & "MAXRECORD")
			tmp01 = "MAX_" & i
			strtmp = strtmp & "<tr>"
			For j = 0 to Session(csv & "_" & tmp01)
				tmp02 = i & "_" & j
				Session(csv & "_" & tmp02) = Com_ConvHTMLTag(Session(csv & "_" & tmp02))
				strtmp = strtmp & "<td>" & Session(csv & "_" & tmp02) & "</td>"
			Next
			strtmp = strtmp & "</tr>"
		Next
		strtmp = strtmp & "</table>"

		Com_MakeCSVLineFromHash = strtmp

	End If

End Function

'=====================================================================
' [Function		] CSV形式に変換
' [Contents     ] 文字列をCSV形式に変換する
'---------------------------------------------------------------------
' [Return       ] (string)変換した文字列
' [Input        ] (string)strtmp:変換する文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ConvCSVString(strtmp)

	strtmp = Replace(strtmp,"""","""""")
	If(InStr(strtmp,",")) Then
		strtmp = Com_SandWitch(1,strtmp,"""")
	End If

	'strtmp = Replace(strtmp,vbCr + vbLf,vbLf)
	strtmp = Replace(strtmp,vbCrLf,"_N_")

	Com_ConvCSVString = strtmp

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
Function OpenBinaryFileData()
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
Function RecordBinaryFileData()
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
Function Com_FileTransferVer1()
End Function

'=====================================================================
' [Function		] INI連想配列から配列を取得する
' [Contents     ] INI連想配列から配列を取得する
'---------------------------------------------------------------------
' [Return       ] (string array)配列
' [Input        ] (integer)mode:モード
'				  (string)section:セクション名
'				  (string)hash:連想配列名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetArrayFromINIHash(mode, section, hash)

	Dim i
	Dim tmp01
	Dim strtmp
	Dim tmparray

	tmparray = Array("")

	If(mode = 1) Then

		For i = 1 to 1000
			tmp01 = hash & "_" & section & "_" & i

			If(Session(tmp01) = "") Then
			Else
				Com_ArrayPush tmparray,Session(tmp01)
			End If

		Next

		Com_GetArrayFromINIHash = tmparray

	End If

End Function

'=====================================================================
' [Function		] オリジナルＳＱＬ
' [Contents     ] テキストベースの簡単な値取得関数
'---------------------------------------------------------------------
' [Return       ] (string)mode=1:要求された文字列,2=配列
' [Input        ] (integer)mode:モード
'				  (string)fname:ファイル名
'				  (integer)lsp:１データごとの区切り文字
'				  (integer)dsp:１項目ごとの区切り文字
'				  (string)id:...???
'				  (integer)idx:項目の位置
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_FunnySQL(mode,fname,lsp,dsp,id,idx)

	Dim i
	Dim imax
	Dim strtmp
	Dim fdata
	Dim flines
	Dim dlines
	Dim strarray

	strarray = Array("")

	If(mode = "1") Then
		If(fname = "") Then
			fdata = Com_pStr
		Else
			fdata = Com_ReadFileData(1,fname)
		End If

		flines = Split(fdata,lsp)
		imax = UBound(flines)

		For i = 0 to imax
			If(InStr(flines(i),dsp)) Then
				dlines = Split(flines(i),dsp)
				If(dlines(0) = id) Then
					strtmp = dlines(idx)
					Exit For
				End If
			End If
		Next

		Com_FunnySQL = strtmp

	ElseIf(mode = "2") Then

		If(fname = "") Then
			fdata = Com_pStr
		Else
			fdata = Com_ReadFileData(1,fname)
		End If

		flines = Split(fdata,lsp)
		imax = UBound(flines)


		For i = 0 to imax
			If(InStr(flines(i),dsp)) Then
				dlines = Split(flines(i),dsp)
				Com_ArrayPush strarray , dlines(idx)
			End If
		Next

		Com_FunnySQL = strarray
	End If

End Function

'=====================================================================
' [Function		] 確認付きフォルダ作成
' [Contents     ] フォルダが存在してなかったら作成する
'---------------------------------------------------------------------
' [Return       ] (boolean)作成したか、しないか
' [Input        ] (string)path:パス
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_CheckAndMakeFile(path)


	Dim pFobj

	Set pFobj = Server.CreateObject("Scripting.FileSystemObject")

	Com_CheckAndMakeFile = False

	If Not(pFobj.FolderExists(path)) Then Exit Function
	pFobj.CreateFolder(path)

	Com_CheckAndMakeFile = True

	Set pFobj = Nothing

End Function

%>