import subprocess
from io import BytesIO
from subprocess import Popen

from lxml import etree

from custom_mailmerge import MergeField


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
    docx_file_url = './BM_DEMO_v1 dùng de test du th .docx'
    # docx_to_pdf('./filecanfix.docx', './output.pdf')
    # docx_file_url = './MAU_BAOCAO_TONGHOP_FULL (copy).docx'

    # namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    # file = etree.parse(docx_file_url)
    # for index, si in enumerate(file.findall(f"{{{namespace}}}si")):
    #     if index == 60:
    #         ts = si.findall(f'.//{{{namespace}}}t')
    #         for t in ts:
    #             print(t.text)

    with MergeField(docx_file_url) as document:
        document.merge({
            "ho":"",
            "ho2": '',
        "ten":"123123213",
            "ho23":'',

            "hohohoho": "",
            "ho3":"",

            "ten2":"",
            "ten3": ""
        },
            {
        "nhom.a333":[
            {
                "nhom.a333": "",
            }
                ],
            "nhom.a4444":[{
                "nhom.a4444": "",
                "nhom.a444224":""
            }

            ],

            }
        )
        # for key, value in document.merge_fields.items():
        #     print(key, value)
        # print(document.merge_fields)
        document.write('exampledoc.docx')

    # docx_to_pdf('example.docx', '')

    # with MergeField(docx_file_url, is_remove_empty_table=True) as delete_space_mergefield:
    #     delete_space_mergefield.change_merge_field_have_space()
    #     delete_space_mergefield.write('new_example.docx')

    # print(test)
