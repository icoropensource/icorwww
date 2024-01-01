// Register a plugin named "icorspell".
CKEDITOR.plugins.add( 'icorspell',
{
	init : function( editor )
	{
		editor.addCommand( 'icorspell',
			{
				modes : { wysiwyg:1, source:1 },

				exec : function( editor )
				{
					var command = this;
					function afterCommand()
					{
						// Defer to happen after 'selectionChange'.
						setTimeout( function()
						{
							editor.fire( 'afterCommandExec',
							{
								name: command.name,
								command: command
							} );
						}, 500 );
					}
					if ( editor.mode == 'wysiwyg')
						editor.on( 'contentDom', function( evt ){

							evt.removeListener();
	                        afterCommand();
						} );

               var aparams=new Object();
               aparams.html=editor.getData();
               aparams.lang=editor.config.language;
               ret=showModalDialog(editor.config.icorspell_path,aparams,'dialogWidth:700px;dialogHeight:550px;help:No;Status:No;resizable:Yes;scroll:No');
               if (ret!='') {
                  editor.setData(ret);
               }
					editor.focus();

					if( editor.mode == 'source' )
						afterCommand();

				},
				async : true
			});

		editor.ui.addButton( 'IcorSpell',
			{
				label : editor.lang.spellCheck.title,
				command : 'icorspell',
            icon:this.path+'skins/'+editor.config.skin+'/icons.png'
			});
	}
});

CKEDITOR.config.icorspell_path = "";
