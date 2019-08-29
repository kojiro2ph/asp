<%

Public Com_Sql_cn		'Connectionオブジェクト
Public Com_Sql_rs		'Recordsetオブジェクト
Public Com_Sql_cmd		'Commandオブジェクト
Public Com_Sql_prm		'Parameterオブジェクト
Public Com_Sql_cs		'Sql用接続先文字列情報
Public Com_Sql_uid		'Sql用接続UserID
Public Com_Sql_pwd		'Sql用Password
Public Com_Sql_ds		'Sql用DataSource

'=====================================================================
' [Function		] SQL初期化
' [Contents     ] SQL接続関係の変数などを初期化する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_Init_SQL()

	On Error Resume Next

	'接続用の変数が設定されてない場合はここで標準の設定にする
	If(Com_Sql_ds = "" and Com_Sql_pwd = "" and Com_Sql_uid = "") Then
		Com_Sql_ds 		= "TAKO"
		Com_Sql_uid 	= "takoadmin"
		Com_Sql_pwd 	= "takoadmin"
	End If

	'接続先文字列情報作成
	Com_Sql_cs = "User ID=" 	& Com_Sql_uid
	Com_Sql_cs = Com_Sql_cs & 	";Password=" & Com_Sql_pwd
	Com_Sql_cs = Com_Sql_cs & 	";Data Source=" & Com_Sql_ds

End Function

'=====================================================================
' [Function		] SQL接続
' [Contents     ] SQLに接続する
'---------------------------------------------------------------------
' [Return       ] True または False
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ConnectSQL()

	On Error Resume Next

	Com_ConnectSQL = True
	Set Com_Sql_cn = Server.CreateObject("ADODB.Connection")
	Com_Sql_cn.Open Com_Sql_cs
	If(Err.Number <> 0) Then Com_ConnectSQL = False
	Set Com_Sql_rs = Server.CreateObject("ADODB.Recordset")
	If(Err.Number <> 0) Then Com_ConnectSQL = False

End Function

'=====================================================================
' [Function		] SQL切断
' [Contents     ] SQLを切断する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_DisconnectSQL()

	On Error Resume Next

	If(Com_Sql_rs.State) Then Com_Sql_rs.Close
	Set Com_Sql_rs = Nothing
	Com_Sql_cn.Close
	Set Com_Sql_cn = Nothing

End Function

'=====================================================================
' [Function		] SQL実行
' [Contents     ] SQLを実行する
'---------------------------------------------------------------------
' [Return       ] True または False
' [Input        ] strsql : SQL文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ExecSQL(strsql)

	On Error Resume Next

	Com_ExecSQL = True
	If(Com_Sql_rs.State <> 0) Then Com_Sql_rs.Close
	Com_Sql_rs.Open strsql,Com_Sql_cn
	If(Err.Number <> 0) Then Com_ExecSQL = False
End Function

'=====================================================================
' [Function		] クエリ行数取得
' [Contents     ] クエリーの行数を取得する
'---------------------------------------------------------------------
' [Return       ] (integer)行数
' [Input        ] string strsql : SQL文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetRows(strsql)

	On Error Resume Next

	Com_Sql_rs.Open strsql, Com_Sql_cn, 3, 1
	Com_GetRows = Com_Sql_rs.RecordCount
	Com_Sql_rs.Close

End Function

'=====================================================================
' [Function		] 最初の値を取得する
' [Contents     ] 最初の値を取得する
'---------------------------------------------------------------------
' [Return       ] (string)値
' [Input        ] string strsql : SQL文字列
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetValue(strsql)

	On Error Resume Next

	Com_Sql_rs.Open strsql, Com_Sql_cn, 3, 1
	If(Com_Sql_rs.EOF) Then
		Com_GetValue = ""
	Else
		Com_GetValue = Com_Sql_rs(0).Value
	End If
	Com_Sql_rs.Close

End Function

'=====================================================================
' [Function		] SP用のオブジェクトをスタンバイする
' [Contents     ] SP用のオブジェクトをスタンバイする
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_SetSPObject() 

	On Error Resume Next

	'Commandｵﾌﾞｼﾞｪｸﾄの作成
	Set Com_Sql_cmd = Server.CreateObject("ADODB.Command")
	Set Com_Sql_cmd.ActiveConnection = Com_Sql_cn

	'Parameterｵﾌﾞｼﾞｪｸﾄの作成
	Set Com_Sql_prm = Server.CreateObject("ADODB.Parameter")

End Function

'=====================================================================
' [Function		] SP用のオブジェクトを開放する
' [Contents     ] SP用のオブジェクトを開放する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_CloseSPObject() 

	On Error Resume Next

	'オブジェクトの開放
	Set Com_Sql_cmd = Nothing
	Set Com_Sql_prm = Nothing

End Function

'=====================================================================
' [Function		] SP実行
' [Contents     ] SPを実行する
'---------------------------------------------------------------------
' [Return       ] なし
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_ExecSP(strspname,intspmode,strprm,strprmd1,strprmd2)

	Dim i
	Dim imax
	Dim strprmarray
	Dim strprmarray2

	On Error Resume Next

	Com_SetSPObject()

	Com_Sql_cmd.CommandText = strspname
	Com_Sql_cmd.CommandType = intspmode

	If(strprm <> "") Then
		strprmarray = Split(strprmd1,strprm)
		imax = UBound(strprmarray)

		For i = 0 to imax
			strprmarray2 = Split(strprmd2,strprmarray(i))
			'パラメータ追加
			Set Com_Sql_prm = Com_Sql_cmd.CreateParameter(strprmarray2(0),strprmarray2(1),strprmarray2(2),strprmarray2(3),strprmarray2(4))
			Com_Sql_cmd.Parameters.Append Com_Sql_prm
			'★
		Next

	End If

	Com_Sql_cmd.Execute

	Com_CloseSPObject()

End Function

'=====================================================================
' [Function		] テーブル項目配列取得
' [Contents     ] 指定のテーブルから項目を配列で取得する
'---------------------------------------------------------------------
' [Return       ] (string array)項目
' [Input        ] (string)tname:テーブル名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetFldArrayFromTable(tname)

	Dim i
	Dim ar()
	Dim strsql

	On Error Resume Next

	strsql = ""
	strsql = "sp_columns " & tname

	'ＳＱＬ文実行
	Com_ExecSQL(strsql)

	i = 0
	Do While Not (Com_Sql_rs.EOF)
		ReDim Preserve ar(i)
		ar(i) = Com_Sql_rs("COLUMN_NAME")
		Com_Sql_rs.MoveNext
		i = i + 1
	Loop

	Com_GetFldArrayFromTable = ar

End Function

'=====================================================================
' [Function		] テーブル配列取得
' [Contents     ] User Table 配列取得
'---------------------------------------------------------------------
' [Return       ] (string array):テーブル名
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetTableArray()
	Debug Com_ExecSQL("SELECT name FROM sysobjects WHERE category = 512")
	Com_GetTableArray = Com_GetSpecFieldArray("name")
End Function

'=====================================================================
' [Function		] 指定項目配列取得
' [Contents     ] Com_Sql_rsに入っている指定項目配列を取得する
'---------------------------------------------------------------------
' [Return       ] (string array):テーブル名
' [Input        ] なし
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetSpecFieldArray(fldname)

	Dim i
	Dim ar()

	i = 0
	Do While Not(Com_Sql_rs.EOF)
		ReDim Preserve ar(i)
		ar(i) = Com_Sql_rs(fldname)
		Com_Sql_rs.MoveNext
		i = i + 1
	Loop

	Com_GetSpecFieldArray = ar

End Function

'=====================================================================
' [Function		] 指定テーブル内の指定項目配列を取得する
' [Contents     ] 指定テーブル内の指定項目配列を取得する
'---------------------------------------------------------------------
' [Return       ] (string array):項目配列
' [Input        ] (string)tname:テーブル名
'				  (string)fldname:項目名
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP用
' [Update       ] 
'=====================================================================
Function Com_GetSpecFldArrayFromSpecTable(tname,fldname)

	Com_ExecSQL("SELECT " & fldname & " FROM " & tname)
	Com_GetSpecFldArrayFromSpecTable = Com_GetSpecFieldArray(fldname)

End Function

%>