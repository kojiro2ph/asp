<%

Dim ThisFile : ThisFile = Request.ServerVariables("SCRIPT_NAME")

'=====================================================================
' [Function		] �t�H�[���f�[�^������
' [Contents     ] Request.Fom�܂���Request.QueryString�ϐ���Session�Ɋi�[����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_Init_Form()

	'�t�H�[���f�[�^���Ƃ̏���
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
' [Function		] �t�H�[��Session�ϐ��N���A
' [Contents     ] Session�Ɋi�[����Form�ϐ�����ɂ���
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_f_SessionClear()

	Dim i
	Dim j
	Dim str

	For i = 1 to Session.Contents.Count

		'F_��f_�ɕϊ�����iiis�Ή��j
		str = Session.Contents.Key(i)
		str = Replace(str,"F_","f_")

		j = InStr(1,str,"f_") '������f_�ɒ��ӁI

		If(j = 1) Then
			Session(Session.Contents.Key(i)) = ""
		End If

	Next

End Function

'=====================================================================
' [Function		] �ݒ�t�@�C��������
' [Contents     ] �ݒ�t�@�C����Session�ϐ��Ɋi�[����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (integer) mode : ���������[�h
'				  (string) inifile : �ݒ�t�@�C���p�X
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
' [Update       ] 
'=====================================================================
Function Com_InitINIData(mode, inifile)

	Dim f,t,b,ar,k,v,n,sec

	'�m�[�}�����[�h����������������������������������������������
	If(mode = 1) Then

		Set f = Server.CreateObject("Scripting.FileSystemObject")
		Set t = f.OpenTextFile(inifile)

		'�s���[�v
		Do Until t.AtEndOfStream

			'��s�ǂݍ��� ----
			b = t.ReadLine

			'"[" ��T�� ----
			n = InStr(1, b, "[")
			'Response.Write(n & "dao<br>" & vbCrLf)
			'Response.Write(Server.HTMLEncode(b) & "���`�F�b�N���܂�<br>" & vbCrLf)

			If(b = "") Then
				n = -1
			End If

			'�Z�N�V������؂�̏ꍇ ===========================
			If(n = 1) Then
				'Response.Write("�Z�N�V�����ł�<br>" & vbCrLf)
				b = Right(b,len(b) - 1)
				b = Left(b,len(b) - 1)

				sec = b
				'Response.Write(sec & "�Z�N�V�����ł�<br>" & vbCrLf)

			ElseIf(n = -1) Then

				'Response.Write("��ł�<br>" & vbCrLf)

			'�Z�N�V���������̏ꍇ ===========================
			Else
				'[=]�ŋ�؂� ----
				ar = Split(b,"=")

				'���[�̋󔒂��Ƃ� ----
				k = Trim(ar(0))
				k = Replace(k,vbTab,"")
				v = Trim(ar(1))

				'���[��["]���������ꍇ�폜���� ----
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
				'Response.Write(sec & "_" & k & "�̒l��" & v & "�ł�<br>" & vbCrLf)
			End If
		Loop

		'�t�@�C������� ----
		t.close

		'�I�u�W�F�N�g�̎Q�Ƃ��Ȃ��� ----
		Set f = Nothing
		Set t = Nothing
	End If

End Function

'=====================================================================
' [Function		] �ݒ�t�@�C���X�V
' [Contents     ] �w�肵���ݒ�t�@�C���̎w��Z�N�V�����̎w��L�[���X�V����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] (integer)mode:�X�V���[�h
'				  (string)fname:�t�@�C����
'				  (string)section:�X�V����Z�N�V������
'				  (string)strFkey:�X�V����L�[��
'				  (string)strCval:�X�V����l
'---------------------------------------------------------------------
' [Create       ] 2002. 1.16  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] NT���[�U�[���O�C�����擾
' [Contents     ] �A�N�Z�X���Ă���Login�����擾����
'---------------------------------------------------------------------
' [Return       ] (string)���O�C����
' [Input        ] 
'				  
'---------------------------------------------------------------------
' [Create       ] 2002. 1.10  by K.Hamada (TechnoVan) - ASP�p
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
' [Function		] �^�O������
' [Contents     ] �悭�g���^�O��Session�Ɋi�[����
'---------------------------------------------------------------------
' [Return       ] �Ȃ�
' [Input        ] �Ȃ�
'---------------------------------------------------------------------
' [Create       ] 2002. 1.15  by K.Hamada (TechnoVan) - ASP�p
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

