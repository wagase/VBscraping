# scraping
## Windows IIS Classic ASP
�N���V�b�NASP��Instagram�̃E�G�u�X�N���C�s���O���s��  
�D���ȃ��[�U�[�̍ŐV�摜���}�X�N����Ă��Ȃ������₷���`���ŕ\������  
ID������������Ɠ���  

  
# VBscript��Instagram��WEB�X�N���C�s���O�����b
Instagram�̉摜���ȒP�ɕۑ�������@  
## �ړI
Instagram�̉摜�������₷���`���ɂ���localhost�ŎQ�Ƃ���

## �o��
Instagram�̉摜��ۑ�������  
��  
Instagram�̉摜��div�^�O�Ń}�X�N����Ă��ĉE�N���b�N�ŊȒP�ɕۑ�������ł��Ȃ��悤�ɂȂ��Ă���  
��  
�摜��ۑ����������F12�L�[�������ĊJ���҃c�[���𗧂��グ�邩�O���T�[�r�X��URL��\��t����̂��嗬  
��  
����͖ʓ|������  
  
WEB�Ō��J����Ă�����̂����E�F�u�X�N���C�s���O�iWeb scraping�j���܂�  
  
## ��
����FVBScript  
�T�[�o�[�FWindows IIS ASP  

## �\�[�X
�S�̂�github�Q�Ƃ�������  
�ʉ��  
### WEB�T�C�g�փA�N�Z�X  

	Private Function getXMLHTTP(byval url)

		Dim xmlhttp,resText
		Set xmlhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
		xmlhttp.Open "GET",url, False
		xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlhttp.Send ""
		resText = xmlhttp.responseText
		Set xmlhttp = Nothing

		getXMLHTTP = resText
	End Function`
VBS��WEB�T�C�g�փA�N�Z�X����ɂ�  
Server.Createobject("MSXML2.ServerXMLHTTP")  
���g���܂�  
xmlhttp.responseText�Ō��ʂ����̏ꍇHTML���Ԃ��Ă���̂ł������͂��܂�  

### HTML�̉��
���ۂɂ��ꂩ��Instagram�ɃA�N�Z�X���ă\�[�X���݂��script�^�O��javascript�̋L�q�����邱�Ƃ��킩��܂��B������e�L�X�g�`���Ŏ擾���邱�ƂɂȂ�̂Ő��K�\���ŉ�͂��܂��B  
`"https://scontent-nrt1-1\.cdninstagram\.com.*?\.jpg"`  
�擾�������̂�HTML��<img src=�`�`�Ə��������Ă��Ή摜����邾���Ȃ�I���ł�  

### ������ق���
������Ƃ�ɂ͍H�v���K�v��Instagram�̋L���ŗL�y�[�W�ɍs���K�v������܂���  

�L���ŗL�y�[�W�ւ͏�L��javascript��code�Ə�����Ă���P�P���̉p�����̕������  
`https://www.instagram.com/p/(�P�P���̉p����)`  
�̂悤�ɂ���ƍs����悤�ł�  
���������Ă܂��͌ʃy�[�W��code���擾���܂�  
���K�\����  
`"""code"":.*?"", ""date"""`  
�Ƃ���Mid�֐��łP�O�����ڂ���P�P���������Ɨǂ������ł�  

	Set regEx = CreateObject("VBScript.RegExp")
	regEx.Pattern = """code"":.*?"", ""date"""
	regEx.IgnoreCase = False ' �啶���Ə���������ʂ��Ȃ�
	regEx.Global = True ' ������S�̂�����

���Ƃ�  
`Set matches = regEx.Execute()`  
���g���ď��ԂɌʃy�[�W�ɃA�N�Z�X���s�����ꂼ�ꌋ�ʂ��擾����  
�摜�Ȃ�  
`<img src="(�摜URL)">`  
����Ȃ�  
`<video autoplay loop muted controls><source src="(����URL)" type=""video/mp4"" /></video>`  
�Ə�����response.Write���Ă��ƃ}�X�N�̂������Ă��Ȃ��摜�̈ꗗ���擾�ł��܂�  

