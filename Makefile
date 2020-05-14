
f2e_templates.py: sheet1_template.xml.bottom sheet1_template.xml.row sheet1_template.xml.top Makefile
	( echo '###START_OF_F2E_TEMPLATES_PY' \
	\
	; echo -n 'sheet1_template_xml_bottom="""' \
	; cat sheet1_template.xml.bottom \
	; echo '"""' \
	\
	; echo -n 'sheet1_template_xml_row="""' \
	; cat sheet1_template.xml.row \
	; echo '"""' \
	\
	; echo -n 'sheet1_template_xml_top="""' \
	; cat sheet1_template.xml.top \
	; echo '"""' \
	\
	; ( echo 'try:' \
	  ; echo '  no_sheet1_template_xlsx_zip=bytes(' \
	  ; echo "print((open('no_sheet1_template_xlsx.zip','rb').read(10000),))" | python | sed "s/^(b'/('/" \
	  ; echo "[0],encoding='latin1')" \
	  ; echo 'except:' \
	  ; echo '  no_sheet1_template_xlsx_zip=(' \
	  ; echo "print((open('no_sheet1_template_xlsx.zip','rb').read(10000),))" | python | sed "s/^(b'/('/" \
	  ; echo '[0])' \
	  ) \
	\
	; echo '###END_OF_F2E_TEMPLATES_PY' \
	) > $@

f2e.py: f2e_templates.py
	sed -i -e '/^###START_OF_F2E_TEMPLATES_PY/,/^###END_OF_F2E_TEMPLATES_PY/d' \
	       -e '/^###INSERT_F2E_TEMPLATES_PY_HERE/r$<' $@
