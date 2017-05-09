/*
Copyright (c) 2003-2011, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
	// Define changes to default configuration here. For example:
	// config.language = 'fr';
	// config.uiColor = '#AADC6E';
	config.filebrowserBrowseUrl = '../ckfinder/ckfinder.html'; 
    config.filebrowserUploadUrl = '../ckfinder/connector.asp?command=QuickUpload&type=Files'; 
    config.filebrowserImageBrowseUrl = '../ckfinder/ckfinder.html?Type=Images'; 
    config.filebrowserImageUploadUrl = '../ckfinder/connector.asp?command=QuickUpload&type=Images'; 
    config.filebrowserFlashBrowseUrl = '../ckfinder/ckfinder.html?Type=Flash'; 
    config.filebrowserFlashUploadUrl = '../ckfinder/connector.asp?command=QuickUpload&type=Flash';
	config.enterMode = CKEDITOR.ENTER_BR;
	config.shiftEnterMode = CKEDITOR.ENTER_P;
};
