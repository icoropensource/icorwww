/*
Copyright (c) 2003-2009, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
   config.enterMode = CKEDITOR.ENTER_P;
   config.shiftEnterMode = CKEDITOR.ENTER_BR;
   config.forcePasteAsPlainText = false;
   config.language = 'pl';
   config.skin = 'v2';
   config.startupFocus = true;
   
   config.icorspell_path='/icormanager/inc/spellpageck.asp';
   config.extraPlugins = 'icorspell';
  
   config.toolbar = 'UMToolbar';
   config.toolbar_UMToolbar =
[
    ['Source','-','NewPage','Preview','-','Templates'],
    ['Cut','Copy','Paste','PasteText','PasteFromWord','-','Print','IcorSpell'],
    ['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
    ['Form', 'Checkbox', 'Radio', 'TextField', 'Textarea', 'Select', 'Button', 'ImageButton', 'HiddenField'],
    '/',
    ['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
    ['NumberedList','BulletedList','-','Outdent','Indent','Blockquote'],
    ['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
    ['Link','Unlink','Anchor'],
    ['Image','Flash','Table','HorizontalRule','SpecialChar','PageBreak'],
    '/',
    ['Styles','Format','Font','FontSize'],
    ['TextColor','BGColor'],
    ['Maximize', 'ShowBlocks'], ['IcorFiles','IcorTable']
];

   config.filebrowserBrowseUrl = '/icormanager/inc/ckuploader/browser.html?Connector=connectors/asp/connector.asp';
   config.filebrowserImageBrowseUrl = '/icormanager/inc/ckuploader/browser.html?Connector=connectors/asp/connector.asp&Type=Image';
   config.filebrowserFlashBrowseUrl = '/icormanager/inc/ckuploader/browser.html?Connector=connectors/asp/connector.asp&Type=Flash';
   config.filebrowserUploadUrl = '/icormanager/inc/ckuploader/connectors/asp/upload.asp?CurrentFolder=file';
   config.filebrowserImageUploadUrl = '/icormanager/inc/ckuploader/connectors/asp/upload.asp?CurrentFolder=image';
   config.filebrowserFlashUploadUrl = '/icormanager/inc/ckuploader/connectors/asp/upload.asp?CurrentFolder=flash';

};
CKEDITOR.plugins.addExternal( 'icorspell', '/icormanager/inc/ckplugins/icorspell/');
