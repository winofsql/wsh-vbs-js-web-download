if Wscript.Arguments.Count = 0 then
    Wscript.Echo "httpget url [savepath]"
    Wscript.Quit
end if

' *****************************
' �_�E�����[�h�p�̃I�u�W�F�N�g
' *****************************
Set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")

' *****************************
' ��1������ URL
' *****************************
strUrl = Wscript.Arguments(0)
if Wscript.Arguments.Count = 1 then
    ' ��2�����������ꍇ�́AURL �̍Ō�̕���
    ' ( �J�����g�Ƀ_�E�����[�h )
    aData = Split(strUrl,"/")
    strFile = aData(Ubound(aData))
else
    ' ��2����������ꍇ�͂�������[�J���t�@�C���Ƃ���
    strFile = Wscript.Arguments(1)
end if

' *****************************
' �_�E�����[�h�v��
' *****************************
on error resume next
Call objSrvHTTP.Open("GET", strUrl, False )
if Err.Number <> 0 then
    Wscript.Echo Err.Description
    Wscript.Quit
end if
objSrvHTTP.Send
if Err.Number <> 0 then
    ' �����炭�T�[�o�[�̎w�肪�Ԉ���Ă���
    Wscript.Echo Err.Description
    Wscript.Quit
end if
on error goto 0

if objSrvHTTP.status = 404 then
    Wscript.Echo "URL ������������܂���(404)"
    Wscript.Quit
end if

' *****************************
' �o�C�i���f�[�^�ۑ��p�I�u�W�F�N�g
' *****************************
Set Stream = Wscript.CreateObject("ADODB.Stream")
Stream.Open
Stream.Type = 1	' �o�C�i��
' �߂��ꂽ�o�C�i�����t�@�C���Ƃ��ăX�g���[���ɏ�������
Stream.Write objSrvHTTP.responseBody
' �t�@�C���Ƃ��ĕۑ�
Stream.SaveToFile strFile, 2
Stream.Close

