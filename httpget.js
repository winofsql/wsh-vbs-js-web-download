// ****************************
// ��������
// ****************************
if ( WScript.Arguments.length == 0 ) {
    WScript.Echo( "httpget url [savepath]" );
    WScript.Quit();
}

// ****************************
// �_�E�����[�h�p�̃I�u�W�F�N�g
// ****************************
var http = new ActiveXObject("Msxml2.ServerXMLHTTP")

// ****************************
// ��1������ URL
// ****************************
var file;
var url = WScript.Arguments(0);
if ( WScript.Arguments.length == 1 ) {
    // ��2�����������ꍇ�́AURL �̍Ō�̕���
    // ( �J�����g�Ƀ_�E�����[�h )
    var aData = url.split("/");
    file = aData[aData.length-1];
}
else {
    // ��2����������ꍇ�͂�������[�J���t�@�C���Ƃ���
    file = WScript.Arguments(1);
}

// ****************************
// �_�E�����[�h�v��
// ****************************
WScript.Echo( url );
http.open("GET", url, false );
http.send();
try {
} catch (error) {
    WScript.Echo( error.description );
    WScript.Quit();
}

if ( http.status == 404  ) {
    WScript.Echo( "URL ������������܂���(404)" );
    WScript.Quit();
}

// ****************************
// �o�C�i���f�[�^�ۑ��p�I�u�W�F�N�g
// ****************************
var stream = new ActiveXObject("ADODB.Stream");
stream.Open();
stream.Type = 1	// �o�C�i��
// �߂��ꂽ�o�C�i�����t�@�C���Ƃ��ăX�g���[���ɏ�������
stream.Write( http.responseBody );
// �t�@�C���Ƃ��ĕۑ�
stream.SaveToFile( file, 2 );
stream.Close

// ****************************
// �t�@�C���̍Ō�
// ****************************

