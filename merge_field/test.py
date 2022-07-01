import subprocess
from io import BytesIO
from subprocess import Popen

from lxml import etree

# from custom_mailmerge import MergeField, MERGE_FIELD_TYPE_CHECKBOX
# from custome_mailmeger_v2 import MergeField, MERGE_FIELD_TYPE_CHECKBOX
from custom_mailmergev3 import MergeField, MERGE_FIELD_TYPE_CHECKBOX
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

    with MergeField(file=docx_file_url) as document:
        document.merge(not_in_group_replacements={
            # "nhomlon1": {'value': [], 'type': 'show_row_in_table'},
            "nhomlon2": {'value': [1], 'type': 'show_row_in_table'}

        }, in_group_replacements={
            # "bay2":[
            #     {
            #         "bay4":444,
            #         "bay2":123123,
            #         "bay1":55555,
            #         "bay1a":[
            #             {
            #                 "bay1a":11,
            #             },
            #             {
            #                 "bay1a": 22,
            #             }
            #         ]
            #     },
            #       {
            #         "bay2":55555,
            #         "bay1":66666
            #     }
            # ],
            # 'mot1':[],
            # 'hai1':[
            #     {
            #         "hai1":123123,
            #         "hai2":"hahahah",
            #         "hai2a":[
            #             {
            #                 "hai2a":1232323
            #             },
            #             {
            #                 "hai2a": 'ggggggggg'
            #             }
            #         ],
            #         "hai5":945958
            #     },
            #     {
            #         "hai1": 4444,
            #         "hai2": "hihihi"
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




