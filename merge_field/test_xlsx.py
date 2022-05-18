from merge_field.customailmerge_xlsx import MergeField

if __name__ == "__main__":
    xlsx_url = './S2 - VAY - MÃ HÓA BIỂU MẪU.xlsx'
    with MergeField(xlsx_url, is_remove_empty_table=True) as document:
        for i in document.merge_fields():
            print(i)