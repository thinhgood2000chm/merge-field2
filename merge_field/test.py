import subprocess
from io import BytesIO
from subprocess import Popen

from lxml import etree

from custom_mailmerge import MergeField, MERGE_FIELD_TYPE_CHECKBOX


def docx_to_pdf(input_docx_path, output_folder_path):
    import sys
    # if sys.platform.startswith('win'):
    #     process = f'start /wait soffice --headless --convert-to pdf --outdir "{output_folder_path}" "{input_docx_path}"'
    #     subprocess.call(process, shell=True)
    # else:
    #     process = Popen(
    #         ["libreoffice", "--headless", "--convert-to", "pdf",
    #             "--outdir", output_folder_path, input_docx_path]
    #     )
    #     process.communicate()


if __name__ == "__main__":
    # docx_file_url = './test.docx'
    docx_file_url = './an hien dong.docx'
    # docx_to_pdf('./filecanfix.docx', './output.pdf')
    # docx_file_url = './MAU_BAOCAO_TONGHOP_FULL (copy).docx'

    # namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    # file = etree.parse(docx_file_url)
    # for index, si in enumerate(file.findall(f"{{{namespace}}}si")):
    #     if index == 60:
    #         ts = si.findall(f'.//{{{namespace}}}t')
    #         for t in ts:
    #             print(t.text)

    with MergeField(file=docx_file_url) as document:
        document.merge(not_in_group_replacements={
            # "nhomlon1": {'value': [], 'type': 'show_row_in_table'},
            "nhomlon2=hhihihi": {'value': [], 'type': 'show_row_in_table'}
        }, in_group_replacements=

            {
            #
            # "S1.A.V.2.1.10.1.14.12a":[{
            #     'S1.A.V.2.1.10.1.14.12a':123123
            # },
            #     {
            #         'S1.A.V.2.1.10.1.14.12a': 12222
            #     }
            # ]
            }


        )

        for key, value in document.merge_fields.items():
            print(key, value)
        # print(document.merge_fields)
        document.write('exampledoc.docx')

    # docx_to_pdf('example.docx', '')

    # with MergeField(docx_file_url, is_remove_empty_table=True) as delete_space_mergefield:
    #     delete_space_mergefield.change_merge_field_have_space()
    #     delete_space_mergefield.write('new_example.docx')

    # print(test)




