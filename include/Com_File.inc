<%

'=====================================================================
' [Function		] �t�@�C���ǂݍ���
' [Contents     ] �t�@�C���������Ԃ�
'---------------------------------------------------------------------
' [Return       ] (string)�t�@�C��������
' [Input        ] (int) mode : �擾���[�h
'				  (string) fpath : �t�@�C���p�X
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �t�@�C����������
' [Contents     ] �t�@�C����������
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (int) mode : �������݃��[�h
'				  (string) fname : �t�@�C����
'				  (string) str : ������
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_RecordFileData(mode,fname,str)

	Dim ObjFSO				'�t�@�C���I�u�W�F�N�g
	Dim ObjTS				'�e�L�X�g�I�u�W�F�N�g

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
' [Function		] �f�B���N�g���ꗗ�擾
' [Contents     ] �w��f�B���N�g���̈ꗗ��z��ŕԂ�
'---------------------------------------------------------------------
' [Return       ] (string array)�t�@�C���z��
' [Input        ] (int) mode : �擾���[�h
'				  (string) path : �f�B���N�g���p�X
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �b�r�u�t�@�C������g�s�l�k�e�[�u�����쐬����
' [Contents     ] �b�r�u�t�@�C������g�s�l�k�e�[�u�����쐬����
'---------------------------------------------------------------------
' [Return       ] (string)�e�[�u��
' [Input        ] (integer)mode:�\�����[�h
'				  (string)fname:�t�@�C���p�X
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
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

	'�ʏ탂�[�h�i�e�[�u���o�́j
	If(mode = 1) Then

		'�t�@�C���f�[�^�ǂݍ���
		strlines = Com_ReadFileData(1,fname)

		'��s���Ƃɋ�؂�
		strline = Split(strlines,vbCrLf)

		'�s�����擾����
		imax = UBound(strline)


		'�ő區�ڐ����擾����
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

		'<table>�^�O�쐬
		strtable = "<table border='" & Com_MakeCSVHtmlTable_border & "' width='" & Com_MakeCSVHtmlTable_width & "' style='font-size:" & Com_MakeCSVHtmlTable_FontSize & ";'>" & vbCrLf

		'��s���Ƃ̃��[�v�i<tr><td>..</td></tr>�̍쐬�j
		For i = 0 to imax
			If(strline(i) <> "") Then

				'�s�̐F�ݒ�
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
' [Function		] xls ����g�s�l�k�e�[�u�����쐬����
' [Contents     ] xls ����g�s�l�k�e�[�u�����쐬����
'---------------------------------------------------------------------
' [Return       ] (string)�e�[�u��
' [Input        ] (integer)mode:�\�����[�h
'				  (string)fname:�t�@�C���p�X
'				  (tname)tname:�G���A��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
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

		'<table>�^�O�쐬
		strtable = "<table border='" & Com_MakeExcelHtmlTable_border & "' width='" & Com_MakeExcelHtmlTable_width & "' style='font-size:" & Com_MakeExcelHtmlTable_FontSize & ";'>" & vbCrLf

		' �e�s�̕\��
		i = 0
		Do While Not objRst.EOF

			'�s�̐F�ݒ�
			If(Com_MakeExcelHtmlTable_UseColor) Then
				If(i = 0) Then
					coltr = "bgcolor='" & Com_MakeExcelHtmlTable_ColHeader & "'"
				Else
					coltr = "bgcolor='" & Com_Parapara(Com_MakeExcelHtmlTable_ColTr) & "'"
				End If
			End If

			strtable = strtable & "<tr " & coltr & ">" & vbCrLf


			'���ڂ��Ƃ̃��[�v
			For Each objField In objRst.Fields
					strtable = strtable & "<td " & Com_MakeExcelHtmlTable_NoWrap & ">" & objField & "</td>" & vbCrLf
			Next

			strtable = strtable & "</tr>" & vbCrLf

			i = i + 1
			objRst.MoveNext
		Loop

		strtable = strtable & "</table>" & vbCrLf



		Com_MakeExcelHtmlTable = strtable


		' �N���[�Y
		objRst.Close
		objConn.Close

	End If

End Function

'=====================================================================
' [Function		] ��s�̂b�r�u�f�[�^���J���}��؂�ŕ����Ĕz��ŕԂ�
' [Contents     ] ��s�̂b�r�u�f�[�^���J���}��؂�ŕ����Ĕz��ŕԂ�
'---------------------------------------------------------------------
' [Return       ] (string array)�z��
' [Input        ] (string)�b�r�u��s
'---------------------------------------------------------------------
' [Create       ] 2002. 1.11  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] Session����CSV�s���쐬
' [Contents     ] Session�A�z�z�񂩂�CSV����s�쐬����
'---------------------------------------------------------------------
' [Return       ] (string)CSV��s
' [Input        ] (string)mode:���[�h
'				  (string)csv:�A�z�z��̖��O
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] CSV�`���ɕϊ�
' [Contents     ] �������CSV�`���ɕϊ�����
'---------------------------------------------------------------------
' [Return       ] (string)�ϊ�����������
' [Input        ] (string)strtmp:�ϊ����镶����
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP�p
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
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
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_FileTransferVer1()
End Function

'=====================================================================
' [Function		] INI�A�z�z�񂩂�z����擾����
' [Contents     ] INI�A�z�z�񂩂�z����擾����
'---------------------------------------------------------------------
' [Return       ] (string array)�z��
' [Input        ] (integer)mode:���[�h
'				  (string)section:�Z�N�V������
'				  (string)hash:�A�z�z��
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �I���W�i���r�p�k
' [Contents     ] �e�L�X�g�x�[�X�̊ȒP�Ȓl�擾�֐�
'---------------------------------------------------------------------
' [Return       ] (string)mode=1:�v�����ꂽ������,2=�z��
' [Input        ] (integer)mode:���[�h
'				  (string)fname:�t�@�C����
'				  (integer)lsp:�P�f�[�^���Ƃ̋�؂蕶��
'				  (integer)dsp:�P���ڂ��Ƃ̋�؂蕶��
'				  (string)id:...???
'				  (integer)idx:���ڂ̈ʒu
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �m�F�t���t�H���_�쐬
' [Contents     ] �t�H���_�����݂��ĂȂ�������쐬����
'---------------------------------------------------------------------
' [Return       ] (boolean)�쐬�������A���Ȃ���
' [Input        ] (string)path:�p�X
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
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