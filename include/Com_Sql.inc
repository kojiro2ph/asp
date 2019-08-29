<%

Public Com_Sql_cn		'Connection�I�u�W�F�N�g
Public Com_Sql_rs		'Recordset�I�u�W�F�N�g
Public Com_Sql_cmd		'Command�I�u�W�F�N�g
Public Com_Sql_prm		'Parameter�I�u�W�F�N�g
Public Com_Sql_cs		'Sql�p�ڑ��敶������
Public Com_Sql_uid		'Sql�p�ڑ�UserID
Public Com_Sql_pwd		'Sql�pPassword
Public Com_Sql_ds		'Sql�pDataSource

'=====================================================================
' [Function		] SQL������
' [Contents     ] SQL�ڑ��֌W�̕ϐ��Ȃǂ�����������
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_Init_SQL()

	On Error Resume Next

	'�ڑ��p�̕ϐ����ݒ肳��ĂȂ��ꍇ�͂����ŕW���̐ݒ�ɂ���
	If(Com_Sql_ds = "" and Com_Sql_pwd = "" and Com_Sql_uid = "") Then
		Com_Sql_ds 		= "TAKO"
		Com_Sql_uid 	= "takoadmin"
		Com_Sql_pwd 	= "takoadmin"
	End If

	'�ڑ��敶������쐬
	Com_Sql_cs = "User ID=" 	& Com_Sql_uid
	Com_Sql_cs = Com_Sql_cs & 	";Password=" & Com_Sql_pwd
	Com_Sql_cs = Com_Sql_cs & 	";Data Source=" & Com_Sql_ds

End Function

'=====================================================================
' [Function		] SQL�ڑ�
' [Contents     ] SQL�ɐڑ�����
'---------------------------------------------------------------------
' [Return       ] True �܂��� False
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] SQL�ؒf
' [Contents     ] SQL��ؒf����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] SQL���s
' [Contents     ] SQL�����s����
'---------------------------------------------------------------------
' [Return       ] True �܂��� False
' [Input        ] strsql : SQL������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �N�G���s���擾
' [Contents     ] �N�G���[�̍s�����擾����
'---------------------------------------------------------------------
' [Return       ] (integer)�s��
' [Input        ] string strsql : SQL������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetRows(strsql)

	On Error Resume Next

	Com_Sql_rs.Open strsql, Com_Sql_cn, 3, 1
	Com_GetRows = Com_Sql_rs.RecordCount
	Com_Sql_rs.Close

End Function

'=====================================================================
' [Function		] �ŏ��̒l���擾����
' [Contents     ] �ŏ��̒l���擾����
'---------------------------------------------------------------------
' [Return       ] (string)�l
' [Input        ] string strsql : SQL������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] SP�p�̃I�u�W�F�N�g���X�^���o�C����
' [Contents     ] SP�p�̃I�u�W�F�N�g���X�^���o�C����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_SetSPObject() 

	On Error Resume Next

	'Command��޼ު�Ă̍쐬
	Set Com_Sql_cmd = Server.CreateObject("ADODB.Command")
	Set Com_Sql_cmd.ActiveConnection = Com_Sql_cn

	'Parameter��޼ު�Ă̍쐬
	Set Com_Sql_prm = Server.CreateObject("ADODB.Parameter")

End Function

'=====================================================================
' [Function		] SP�p�̃I�u�W�F�N�g���J������
' [Contents     ] SP�p�̃I�u�W�F�N�g���J������
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_CloseSPObject() 

	On Error Resume Next

	'�I�u�W�F�N�g�̊J��
	Set Com_Sql_cmd = Nothing
	Set Com_Sql_prm = Nothing

End Function

'=====================================================================
' [Function		] SP���s
' [Contents     ] SP�����s����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
			'�p�����[�^�ǉ�
			Set Com_Sql_prm = Com_Sql_cmd.CreateParameter(strprmarray2(0),strprmarray2(1),strprmarray2(2),strprmarray2(3),strprmarray2(4))
			Com_Sql_cmd.Parameters.Append Com_Sql_prm
			'��
		Next

	End If

	Com_Sql_cmd.Execute

	Com_CloseSPObject()

End Function

'=====================================================================
' [Function		] �e�[�u�����ڔz��擾
' [Contents     ] �w��̃e�[�u�����獀�ڂ�z��Ŏ擾����
'---------------------------------------------------------------------
' [Return       ] (string array)����
' [Input        ] (string)tname:�e�[�u����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetFldArrayFromTable(tname)

	Dim i
	Dim ar()
	Dim strsql

	On Error Resume Next

	strsql = ""
	strsql = "sp_columns " & tname

	'�r�p�k�����s
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
' [Function		] �e�[�u���z��擾
' [Contents     ] User Table �z��擾
'---------------------------------------------------------------------
' [Return       ] (string array):�e�[�u����
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetTableArray()
	Debug Com_ExecSQL("SELECT name FROM sysobjects WHERE category = 512")
	Com_GetTableArray = Com_GetSpecFieldArray("name")
End Function

'=====================================================================
' [Function		] �w�荀�ڔz��擾
' [Contents     ] Com_Sql_rs�ɓ����Ă���w�荀�ڔz����擾����
'---------------------------------------------------------------------
' [Return       ] (string array):�e�[�u����
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �w��e�[�u�����̎w�荀�ڔz����擾����
' [Contents     ] �w��e�[�u�����̎w�荀�ڔz����擾����
'---------------------------------------------------------------------
' [Return       ] (string array):���ڔz��
' [Input        ] (string)tname:�e�[�u����
'				  (string)fldname:���ږ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_GetSpecFldArrayFromSpecTable(tname,fldname)

	Com_ExecSQL("SELECT " & fldname & " FROM " & tname)
	Com_GetSpecFldArrayFromSpecTable = Com_GetSpecFieldArray(fldname)

End Function

%>