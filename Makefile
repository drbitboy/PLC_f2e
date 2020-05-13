
f2e_templates.py: sheet1_template.xml.bottom sheet1_template.xml.row sheet1_template.xml.top
	( echo -n 'sheet1_template_xml_bottom="""' \
	; cat sheet1_template.xml.bottom \
	; echo '"""' \
	; echo -n 'sheet1_template_xml_row="""' \
	; cat sheet1_template.xml.row \
	; echo '"""' \
	; echo -n 'sheet1_template_xml_top="""' \
	; cat sheet1_template.xml.top \
	; echo '"""' \
	) > $@
