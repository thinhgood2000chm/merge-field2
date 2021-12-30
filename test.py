import subprocess
from subprocess import Popen
from custom_mailmerge import MergeField


def docx_to_pdf(input_docx_path, output_folder_path):
    import sys
    if sys.platform.startswith('win'):
        process = f'start /wait soffice --headless --convert-to pdf --outdir "{output_folder_path}" "{input_docx_path}"'
        subprocess.call(process, shell=True)
    else:
        process = Popen(
            ["libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", output_folder_path, input_docx_path]
        )
        process.communicate()


if __name__ == "__main__":
    docx_file_url = './BM_01_DE_NGHI_CAP_TIN_DUNG.docx'
    #docx_file_url = './MAU_BAOCAO_TONGHOP_FULL (copy).docx'

    with MergeField(docx_file_url, is_remove_empty_table=True) as document:
        # document.merge({
        #     "BC_sothuadat": "2512542",
        #     "BC_sobando": "BD2512542",
        #     "BC_mucdich": "công nghiệp",
        #     "BC_thoihan": "10 năm",
        #     "BC_diachithuadat": "địa chỉ 100",
        #     "BC_dientichdat": "100",
        #     "BC_dientichchu": "một trăm mét vuông",
        #     "BC_sophathanh": "222513542",
        #     "BC_sovaoso": "SO222513542",
        #     "BC_coquan": "cơ quan",
        #     "BC_ngaycap": "25",
        #     "BC_thang": "01",
        #     "BC_nam": "2000",
        # })

        # document.merge({
        #     "BC_sothuadat": [
        #         {
        #             "BC_sothuadat": "2512542",
        #             "BC_sobando": "BD2512542",
        #             "BC_mucdich": "công nghiệp",
        #             "BC_thoihan": "10 năm",
        #             "BC_diachithuadat": "địa chỉ 100",
        #             "BC_dientichdat": "100",
        #             "BC_dientichchu": "một trăm mét vuông",
        #             "BC_sophathanh": "222513542",
        #             "BC_sovaoso": "SO222513542",
        #             "BC_coquan": "cơ quan",
        #             "BC_ngaycap": "25",
        #             "BC_thang": "01",
        #             "BC_nam": "2000",
        #         },
        #         {
        #             "BC_sothuadat": "□",
        #             "BC_sobando": "3",
        #             "BC_mucdich": "4 công nghiệp",
        #             "BC_thoihan": "510 năm",
        #             "BC_diachithuadat": "6 địa chỉ 100",
        #             "BC_dientichdat": "7 100",
        #             "BC_dientichchu": "8 một trăm mét vuông",
        #             "BC_sophathanh": "9 222513542",
        #             "BC_sovaoso": "10 SO222513542",
        #             "BC_coquan": "11 cơ quan",
        #             "BC_ngaycap": "12 25",
        #             "BC_thang": "13 01",
        #             "BC_nam": "14 2000",
        #         }
        #     ]
        # })

        # document.merge({
        #     "BC_sothuadat": []
        # })
        document.merge({
            # "BC_trangthaiTS": '',
            # "THONG_TIN_HOI_SO_SCB": 'kldfaskljdfaskljdfsklj',
            "THONG_TIN_HOI_SO_SCB2": 'THONG_TIN_HOI_SO_SCB2\n1\n2\n3',
            "THONG_TIN_HOI_SO_SCB": '',
            "TEN_DON_VI_SCB": '927 Tran Hung Dao',
            "SO_UY_QUYEN3": [
                {
                    "SO_UY_QUYEN2": "SO_UY_QUYEN2",
                    "SO_UY_QUYEN4": "SO_UY_QUYEN4 da thay the ",
                    "SO_UY_QUYEN3": "Thử table 1"
                },

                {
                    "SO_UY_QUYEN2": "SO_UY_QUYEN2 1",
                    "SO_UY_QUYEN3": "Thử table 2",
                    "SO_UY_QUYEN4": "Thử table 3"
                },
                {
                    "SO_UY_QUYEN2": "SO_UY_QUYEN23",
                    "SO_UY_QUYEN3": "Thử table 4",
                    "SO_UY_QUYEN4": "Thử table 5"
                }
            ],
            'BC_chinhanh': "927 Tran Hung Dao"
        })

        print(document.merge_fields)
        document.write('example.docx')

    # docx_to_pdf('example.docx', '')
