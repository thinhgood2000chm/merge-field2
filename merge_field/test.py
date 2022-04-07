import subprocess
from io import BytesIO
from subprocess import Popen
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
    docx_file_url = './response.docx'
    # docx_to_pdf('./filecanfix.docx', './output.pdf')
    # docx_file_url = './MAU_BAOCAO_TONGHOP_FULL (copy).docx'

    with MergeField(docx_file_url, is_remove_empty_table=True) as document:
        # document.merge({
        #     "S1.A.1.5.4":['test'],
        #     "S1.A.V.2.3.4.18.20":"123123123",
        #     "S1.A.1.12": {
        #         "value":
        #             [
        #                 {
        #                     "option": "Đăng ký thông tin/Register for information",
        #                     "extend_data": "4567"
        #                 }
        #             ],
        #         "type": "extend_checkbox"
        #     },
        #     "S1.A.IV.8": "11111111111",
        #     "S1.A.1.10.27": "aaaaaaaaaaa1",
        #     'S1.A.1.12.1': {
        #         'value': {
        #             "option": True,
        #             "extend_data": "221122005000050000"
        #         },
        #         "type": "extend_checkbox"
        #     },
        #     'S1.A.1.2.41': True,
        #     "S1.A.1.10.1": {
        #         'value': {
        #             "option": True,
        #             "extend_data": "day laf du lieu fill vao extend "
        #         },
        #         "type": "extend_checkbox"
        #     },
        #     'S1.A.1.10.3': ['Nhanh/Instant'],
        #     "S1.A.1.10.14": {
        #
        #         "value": "NGUYEN PH THAO QUYEN",
        #         "type": "embossed_table"
        #     },
        #     # "S1.A.1.10.28":{
        #     #     "value": "aaaaaaaaaaaa",
        #     #     "type":"embossed_table"
        #     # },
        #     # "S1.A.IV.8":[
        #     #     {
        #     #         "S1.A.IV.8": "hahaha",
        #     #         "S1.A.IV.3": "123123",
        #     #         "S1.A.IV.1": "11111111"
        #     #     },
        #     #     {
        #     #         "S1.A.IV.8": "55555",
        #     #         "S1.A.IV.3": "7777",
        #     #         "S1.A.IV.1": "99999"
        #     #     },
        #     #     {
        #     #         "S1.A.IV.8": "666666",
        #     #         "S1.A.IV.3": "888888",
        #     #         "S1.A.IV.1": "0000000"
        #     #     },
        #     # ],
        #     # "S1.A.IV.2":[
        #     #     {
        #     #         "S1.A.IV.2": ["Bố - mẹ"]
        #     #     },
        #     #     {
        #     #         "S1.A.IV.2": ["Bố - mẹ"]
        #     #     }
        #     # ],
        #     # "S1.A.III.2.6": "Cho vay"
        #
        #     # "S1.A.III.2.14": [
        #     #     "Hàng quý"
        #     # ],
        #     # "SO_UY_QUYEN3": [
        #     #     {
        #     #         "SO_UY_QUYEN2": "SO_UY_QUYEN2",
        #     #         "SO_UY_QUYEN4": "SO_UY_QUYEN4 da thay the ",
        #     #         "SO_UY_QUYEN3": "Thử table 1"
        #     #     },
        #     #
        #     #     {
        #     #         "SO_UY_QUYEN2": "SO_UY_QUYEN2 1",
        #     #         "SO_UY_QUYEN3": "Thử table 2",
        #     #         "SO_UY_QUYEN4": "Thử table 3"
        #     #     },
        #     #     {
        #     #         "SO_UY_QUYEN2": "SO_UY_QUYEN23",
        #     #         "SO_UY_QUYEN3": "Thử table 4",
        #     #         "SO_UY_QUYEN4": "Thử table 5"
        #     #     }
        #     # ],
        #
        # })

        print(document.merge_fields)
        # document.write('example.docx')
        document.write('example.docx')

    # docx_to_pdf('example.docx', '')

    # with MergeField(docx_file_url, is_remove_empty_table=True) as delete_space_mergefield:
    #     delete_space_mergefield.change_merge_field_have_space()
    #     delete_space_mergefield.write('new_example.docx')

    # print(test)
