# docx introduction: https://www.toptal.com/xml/an-informal-introduction-to-docx

# lxml documentation: https://lxml.de/tutorial.html
# https://lxml.de/api/lxml.etree._ElementTree-class.html
# https://lxml.de/api/lxml.etree._Element-class.html
# https://lxml.de/tutorial.html#elementpath

import re
from copy import deepcopy
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from lxml.etree import Element

NAMESPACE_WORDPROCESSINGML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'  # noqa
NAMESPACE_CONTENT_TYPE = 'http://schemas.openxmlformats.org/package/2006/content-types'  # noqa
NAMESPACE_WORDML_2010 = 'http://schemas.microsoft.com/office/word/2010/wordml'  # noqa

ELEMENT_PATH_RECURSIVE_MERGE_FIELD = f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t[@is_merge_field="True"]'
ELEMENT_PATH_RECURSIVE_T = f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t'
ELEMENT_PATH_TR = f'{{{NAMESPACE_WORDPROCESSINGML}}}tr'

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',  # noqa
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',  # noqa
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',  # noqa
)

# https://docs.microsoft.com/en-us/dotnet/desktop/xaml-services/xml-space-handling
# https://www.w3.org/XML/1998/namespace
XML_SPACE_ATTRIBUTE = '{http://www.w3.org/XML/1998/namespace}space'

# open tag should be different from close tag
OPEN_TAG = '«'
CLOSE_TAG = '»'

CHECKBOX_CHECKED_TEXT = '☒'
CHECKBOX_UNCHECKED_TEXT = '☐'

# parse from docx
MERGE_FIELD_TYPE_TEXT = 'text'
MERGE_FIELD_TYPE_CHECKBOX = 'checkbox'

# config in TMS
MERGE_FIELD_TYPE_EMBOSSED_TABLE = "embossed_table"
MERGE_FIELD_TYPE_SHOW_ROW_IN_TABLE = "show_row_in_table"

EMBOSSED_TABLE_BYPASS_SEPARATOR_CHARACTERS = ["-", ":"]

NUMBER_ROW_IN_TABLE_IS_ONE = 1
NUMBER_ROW_IN_TABLE_IS_TWO = 2

FROM_COLOR_VALUE = 'FF0000'
TO_COLOR_VALUE = '000000'


class MergeField:
    def __init__(self, file, is_remove_row_or_table_has_one_row_when_empty=True, change_color_flag=False):
        self.zip: ZipFile = ZipFile(file)
        self.parts: dict = {}

        self.field_name__elements: dict = {}
        self.in_table_field_name__details: dict = {}
        self.checkbox_field_name__list_group_checkbox_details: dict = {}
        self.checkbox_field_name__values: dict = {}
        self.embossed_table_field_name__details: dict = {}

        self.change_color_flag: bool = change_color_flag
        self.is_remove_row_or_table_has_one_row_when_empty: bool = is_remove_row_or_table_has_one_row_when_empty

        try:
            content_types = etree.parse(self.zip.open('[Content_Types].xml'))
            for file in content_types.findall(f'{{{NAMESPACE_CONTENT_TYPE}}}Override'):
                if file.attrib['ContentType'] in CONTENT_TYPES_PARTS:
                    filename = file.attrib['PartName'].split('/', 1)[1]  # remove first /
                    self.parts[filename] = etree.parse(self.zip.open(self.zip.getinfo(filename)))

            for part in self.parts.values():
                # init data for field_name__elements
                self.__parse_merge_fields(part=part)
                # init data for in_table_field_name__details, embossed_table_field_name__details
                self.__parse_tables(part=part)
                # init data for checkbox_field_name__list_group_checkbox_details
                self.__parse_checkboxes(part=part)

        except Exception as ex:
            self.zip.close()
            raise ex

    def __parse_merge_fields(self, part):
        is_found_open_tag = False
        previous_p_element = None
        part_of_field_names = []
        part_of_merge_fields = []
        previous_remainder_field_name = None
        is_append_previous_remainder_field_name_to_part_of_field_names = False
        previous_last_remainder_t_element = None
        previous_p_element_contains_previous_last_remainder_t_element = None

        t_elements = part.findall(ELEMENT_PATH_RECURSIVE_T)

        for t_element in t_elements:

            # Trường hợp thẻ không có text, vd: <w:t xml:space="preserve"/>
            # cần gán lại là string rỗng '' để không chạy lỗi khi duyệt t_element.text
            if t_element.text is None:
                t_element.text = ''

            if not is_found_open_tag and OPEN_TAG not in t_element.text and not previous_remainder_field_name:
                previous_p_element = t_element.getparent().getparent()
                continue

            if not is_found_open_tag and OPEN_TAG in t_element.text:
                is_found_open_tag = True

            current_p_element = t_element.getparent().getparent()
            if previous_remainder_field_name:
                # Trường hợp trong <t> trước có «abc nhưng <t> sau lại nằm trong <p> khác
                # -> Thêm element chứa «abc vào <p> trước đó
                if current_p_element is not previous_p_element_contains_previous_last_remainder_t_element:
                    remainder_after_previous_last_remainder_t_element = deepcopy(previous_last_remainder_t_element)
                    self.__set_text_for_t_element(
                        element=remainder_after_previous_last_remainder_t_element,
                        text=previous_remainder_field_name
                    )
                    previous_last_remainder_t_element.addnext(remainder_after_previous_last_remainder_t_element)

                    previous_remainder_field_name = None
                    is_append_previous_remainder_field_name_to_part_of_field_names = False
                    previous_last_remainder_t_element = None
                    previous_p_element_contains_previous_last_remainder_t_element = None

                    previous_p_element = t_element.getparent().getparent()
                    part_of_field_names = []
                    part_of_merge_fields = []
                    continue

                # Trường hợp trong <t> trước có «abc nhưng <t> cùng nằm trong <p>
                # -> append vào part_of_field_names để phòng <t> sau có thể là def» -> merge field là «abcdef»
                else:
                    if not is_append_previous_remainder_field_name_to_part_of_field_names:
                        part_of_field_names.append(previous_remainder_field_name)
                        is_append_previous_remainder_field_name_to_part_of_field_names = True

            if current_p_element is not previous_p_element:
                part_of_field_names = []
                part_of_merge_fields = []

            previous_p_element = t_element.getparent().getparent()

            part_of_field_names.append(t_element.text)
            # add element between open and close tag to list need to delete
            part_of_merge_fields.append(t_element)

            # handle when found close tag
            if CLOSE_TAG in t_element.text:
                # xóa khoảng trắng bên trong các merge field
                text_contain_list_field_name = re.sub(
                    pattern=rf'{OPEN_TAG}\s*(.*?)\s*{CLOSE_TAG}',
                    repl=rf'{OPEN_TAG}\1{CLOSE_TAG}',
                    string=''.join(part_of_field_names)
                )
                # tìm tất cả các field name bởi vì có thể có nhiều field name trong một <t>
                field_name_contain_open_close_tags = re.findall(
                    pattern=rf'({OPEN_TAG}.*?{CLOSE_TAG})',
                    string=text_contain_list_field_name
                )
                for field_name_contain_open_close_tag in field_name_contain_open_close_tags:
                    start_index_found = text_contain_list_field_name.find(field_name_contain_open_close_tag)

                    front_remainder = text_contain_list_field_name[:start_index_found]
                    if front_remainder:
                        front_remainder_t_element = deepcopy(t_element)
                        self.__set_text_for_t_element(element=front_remainder_t_element, text=front_remainder)
                        t_element.addprevious(front_remainder_t_element)

                    # create merged field
                    new_field_name = field_name_contain_open_close_tag[1:-1]
                    new_merge_field = deepcopy(t_element)
                    new_merge_field.text = field_name_contain_open_close_tag
                    # set new attribute named is_merge_field -> easy find this element by filter by #elementpath
                    new_merge_field.set('is_merge_field', 'True')

                    t_element.addprevious(new_merge_field)

                    # maybe there are some merge field with the same name -> add to list
                    if new_field_name not in self.field_name__elements:
                        self.field_name__elements[new_field_name] = []
                    self.field_name__elements[new_field_name].append(new_merge_field)

                    # xóa phần trước merge field và merge field đã được xử lý
                    text_contain_list_field_name = text_contain_list_field_name[
                        start_index_found + len(field_name_contain_open_close_tag):]

                previous_remainder_field_name = None
                previous_last_remainder_t_element = None
                previous_p_element_contains_previous_last_remainder_t_element = None
                # trường hợp merge field nằm ở nhiều <t>:
                # phía cuối có thể là «abc và <t> tiếp theo là def» -> merge field là «abcdef»
                if OPEN_TAG in text_contain_list_field_name:
                    start_index_found = text_contain_list_field_name.find(OPEN_TAG)

                    last_remainder = text_contain_list_field_name[:start_index_found]
                    last_remainder_t_element = deepcopy(t_element)
                    self.__set_text_for_t_element(element=last_remainder_t_element, text=last_remainder)
                    t_element.addprevious(last_remainder_t_element)

                    previous_remainder_field_name = text_contain_list_field_name[start_index_found:]
                    is_append_previous_remainder_field_name_to_part_of_field_names = False
                    previous_last_remainder_t_element = last_remainder_t_element
                    previous_p_element_contains_previous_last_remainder_t_element = last_remainder_t_element.getparent().getparent()

                # xử lý phần còn lại của text_contain_list_field_name
                else:
                    last_remainder_t_element = deepcopy(t_element)
                    self.__set_text_for_t_element(element=last_remainder_t_element, text=text_contain_list_field_name)
                    t_element.addprevious(last_remainder_t_element)

                # delete all element between open and close tag
                for element in part_of_merge_fields:
                    parent = element.getparent()
                    parent.remove(element)

                is_found_open_tag = False
                part_of_field_names = []
                part_of_merge_fields = []

    def __parse_tables(self, part):
        result = self.__get_in_table_field_name__details__and__embossed_table_field_name__details(part=part)
        self.in_table_field_name__details.update(result['in_table_field_name__details'])
        self.embossed_table_field_name__details.update(result['embossed_table_field_name__details'])

    @staticmethod
    def __get_in_table_field_name__details__and__embossed_table_field_name__details(part,
                                                                                    is_parse_in_table_field_name=True):
        in_table_field_name__details = {}
        embossed_table_field_name__details = {}

        tables = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tbl')
        for table in tables:
            tr__list_tc_index = {}
            if is_parse_in_table_field_name:
                # TH có merge row
                # tr -> tc -> tcPr -> <w:vMerge w:val="restart"/>  (thẻ Merge bắt đầu nhóm Merge)
                # tr -> tc -> tcPr -> <w:vMerge/>  (thẻ Merge con trong nhóm Merge)
                v_merges = table.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}vMerge')
                tr__merged_trs = {}
                tr__index_column_merged_trs = {}
                if v_merges:
                    # Chỉ tìm trong bảng đó, không tìm hết bằng recusive xpath vì có thể có bảng lồng bên trong làm sai index
                    tc_indexes = {}
                    for tr in table.findall(ELEMENT_PATH_TR):
                        column_index = 0
                        for tc in tr.findall(f'{{{NAMESPACE_WORDPROCESSINGML}}}tc'):
                            tc_indexes[tc] = column_index
                            column_index += 1

                    group_trs_by_column_index = {}
                    for v_merge in v_merges:
                        tc = v_merge.getparent().getparent()
                        tr = tc.getparent()

                        # nếu tr không thuộc bảng hiện tại thì bỏ qua vì hàm sẽ xử lý bảng bên trong sau
                        if tr not in table:
                            continue

                        column_index = tc_indexes[tc]
                        # lưu lại trong hàng có các cột merge ở vị trí nào
                        if tr not in tr__list_tc_index:
                            tr__list_tc_index[tr] = []
                        tr__list_tc_index[tr].append({tc: column_index})

                        if column_index not in group_trs_by_column_index:
                            # trong cùng một column_index có thể có nhiều merged row
                            # mỗi phần list trong list là một nhóm các tr bị merge
                            group_trs_by_column_index[column_index] = [[]]

                        # có merged row mới trong cùng một column_index
                        if v_merge.get(f'{{{NAMESPACE_WORDPROCESSINGML}}}val') == 'restart' and len(group_trs_by_column_index[column_index][-1]) != 0:
                            group_trs_by_column_index[column_index].append([])

                        group_trs_by_column_index[column_index][-1].append(tr)

                    for index, group_trs in group_trs_by_column_index.items():
                        for trs in group_trs:
                            for tr in trs:
                                # phải lưu hàng và vị trí cột của hàng merge
                                # tránh th 1 bảng có nhiều cột có ô merge thì sẽ bị sai
                                if tr in tr__index_column_merged_trs:
                                    tr__index_column_merged_trs[tr][index] = trs
                                else:
                                    tr__index_column_merged_trs[tr] = {index: trs}

                # TH chung
                for row in table.findall(ELEMENT_PATH_TR):
                    merge_fields = row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                    current_siblings = [merge_field.text[1:-1] for merge_field in merge_fields]

                    if row in tr__index_column_merged_trs:
                        for index, merged_trs in tr__index_column_merged_trs[row].items():
                            for merged_tr in merged_trs:
                                current_siblings.extend([m.text[1:-1]
                                                         for m in merged_tr.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)])

                    for merge_field in merge_fields:
                        row_merge = None
                        field_name = merge_field.text[1:-1]
                        if not in_table_field_name__details.get(field_name):
                            siblings = list(set(current_siblings))
                        else:
                            siblings = list(
                                set(in_table_field_name__details[field_name]['siblings'] + current_siblings)
                            )
                        # nếu là bảng con nằm trong bảng thì reset contain_merge_field_rows để tránh sai index row
                        if field_name in in_table_field_name__details and in_table_field_name__details[field_name]['table'] is not table:
                            in_table_field_name__details[field_name]['contain_merge_field_rows'] = []

                        contain_merge_field_rows = in_table_field_name__details.get(field_name, {}).get('contain_merge_field_rows', [])

                        # trường hợp nếu 1 field xuất hiện 2 lần trong cùng 1 row
                        # nên phải check rồi append để tránh lỗi chọn sai row index để xóa
                        if row not in tr__index_column_merged_trs:
                            if row not in contain_merge_field_rows:
                                contain_merge_field_rows.append(row)
                        else:
                            if row in tr__list_tc_index:
                                tc = merge_field.getparent().getparent().getparent()
                                for column_index in tr__list_tc_index[row]:
                                    if tc in column_index:
                                        # lấy ra các row merge của field đó dựa theo field đó đang ở cột nào
                                        row_merge = tr__index_column_merged_trs[row][column_index[tc]]
                                        contain_merge_field_rows.append(row_merge)

                        in_table_field_name__details[field_name] = {
                            'table': table,
                            'rows': [row] if row not in tr__index_column_merged_trs else row_merge,
                            'siblings': sorted(siblings),
                            'tr__index_column_merged_trs': tr__index_column_merged_trs,
                            'contain_merge_field_rows': contain_merge_field_rows  # là các row có chứa các field được merge bên trong
                        }

            # Bảng khung dập nổi
            # tbl -> tc -> tbl (bảng dập nổi) + t (merge_field)
            text_element_in_tables = table.findall(ELEMENT_PATH_RECURSIVE_T)
            # chỉ chấp nhận trường hợp các ô trong bảng đều là ô trống hoặc là ký tự phân cách, và có tối thiểu 1
            if (not text_element_in_tables) or \
                    (text_element_in_tables and all(
                        text_element.text in EMBOSSED_TABLE_BYPASS_SEPARATOR_CHARACTERS
                        for text_element in text_element_in_tables)):

                # Quy định bảng dập nổi phải nằm trong ô của table
                tc = table.getparent()
                if tc.tag == f'{{{NAMESPACE_WORDPROCESSINGML}}}tc':
                    merge_fields = tc.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                    for merge_field in merge_fields:
                        field_name = merge_field.text[1:-1]
                        embossed_table_field_name__details[field_name] = {
                            'embossed_table': table,
                            'merge_field': merge_field
                        }

        return {
            'in_table_field_name__details': in_table_field_name__details,
            'embossed_table_field_name__details': embossed_table_field_name__details
        }

    def __parse_checkboxes(self, part):
        self.checkbox_field_name__list_group_checkbox_details.update(
            self.__get_checkbox_field_name__list_group_checkbox_details(part=part)
        )

        for field_name, group_checkbox_details in self.checkbox_field_name__list_group_checkbox_details.items():
            self.checkbox_field_name__values[field_name] = []
            for group_checkbox_detail in group_checkbox_details:
                for value in group_checkbox_detail['values']:
                    if value not in self.checkbox_field_name__values[field_name]:
                        self.checkbox_field_name__values[field_name].append(value)

    def __get_checkbox_field_name__list_group_checkbox_details(self, part):
        # tree order finding checkbox in a tc:
        # tc -> p -> sdt -> sdtPr -> checkbox

        # tree order finding text (checked or unchecked) of checkbox in a tc:
        # tc -> p -> sdt -> sdtContent -> r -> t

        checkbox_field_name__list_group_checkbox_details = {}

        cells = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tc')
        for cell in cells:
            first_matching_checkbox = cell.find(f'.//{{{NAMESPACE_WORDML_2010}}}checkbox')
            if first_matching_checkbox is not None:
                checkbox_infos = []
                checkbox_info = None
                is_found_checkbox = False

                # iterate over all elements in the subtree in Preorder (Root, Left, Right)
                for element in cell.iter():
                    if element.tag == f'{{{NAMESPACE_WORDML_2010}}}checkbox':
                        if not is_found_checkbox:
                            is_found_checkbox = True
                        else:
                            checkbox_infos.append(checkbox_info)

                        # sdt -> sdtPr -> checkbox
                        sdt_element = element.getparent().getparent()
                        checkbox_info = {
                            # Ex. ☒ ABC
                            # checkbox_obj is the only one t element in sdt, contains: ☒ or ☐
                            'checkbox_obj': sdt_element.find(ELEMENT_PATH_RECURSIVE_T),
                            # value_obj is a merged all t element in part_value_objs contains: ABC
                            'value_obj': None,
                            # value is the text of value_obj after strip() -> use to compare with data fill: ABC
                            'value': None,
                            # original_value is the text of value_obj (maybe startswith or endswith space): ABC
                            'original_value': None,
                            # part_value_objs is list of t element after checkbox_obj which contains: ABC
                            'part_value_objs': [],
                            # part_values is list of text value of part_value_objs: ["A", "BC"]
                            'part_values': []
                        }

                    if is_found_checkbox and element.tag == f'{{{NAMESPACE_WORDPROCESSINGML}}}t':
                        if not element.get('is_merge_field'):
                            # t element in sdt of checkbox define status checked or unchecked
                            # -> remove it to keep right values of checkbox
                            if element.text not in [CHECKBOX_CHECKED_TEXT, CHECKBOX_UNCHECKED_TEXT]:
                                checkbox_info['part_value_objs'].append(element)
                                checkbox_info['part_values'].append(element.text)

                        # found merge field which is the nearest with last checkbox
                        # -> group checkboxes together by this merge field
                        else:

                            # append last checkbox_info
                            checkbox_infos.append(checkbox_info)

                            for checkbox_info in checkbox_infos:
                                # join part_values string to one string
                                checkbox_info['original_value'] = ''.join(checkbox_info['part_values'])
                                # strip() -> clean data to use when compare with data fill
                                checkbox_info['value'] = checkbox_info['original_value'].strip()
                                del checkbox_info['part_values']

                                # use first part value object, set text part_values, remove remainder part_value_objs
                                # -> merge all part_value_objs to one element
                                ########################################################################################
                                # handle case: do not have any t element after checkbox_obj
                                if checkbox_info['part_value_objs']:
                                    checkbox_info['value_obj'] = checkbox_info['part_value_objs'][0]

                                    self.__set_text_for_t_element(
                                        element=checkbox_info['value_obj'],
                                        text=checkbox_info['original_value']
                                    )

                                    # remove value_obj from part_value_objs to do not delete it
                                    checkbox_info['part_value_objs'].pop(0)

                                    # remove remainder t_element in part_value_objs
                                    for el in checkbox_info['part_value_objs']:
                                        parent = el.getparent()
                                        parent.remove(el)

                                del checkbox_info['part_value_objs']
                                ########################################################################################

                            field_name = element.text[1:-1]

                            if field_name not in checkbox_field_name__list_group_checkbox_details:
                                checkbox_field_name__list_group_checkbox_details[field_name] = []

                            # IMPORTANT: only append new group checkboxes by this merge field
                            # and do not delete in self.field_name__elements to keep the order of fields
                            checkbox_field_name__list_group_checkbox_details[field_name].append({
                                'checkbox_infos': checkbox_infos,
                                'values': [checkbox_info['value'] for checkbox_info in checkbox_infos],
                                'merge_field': element
                            })

                            # re-assign value for process new group of checkboxes (if exist)
                            # same way with before for element in cell.iter():
                            checkbox_infos = []
                            checkbox_info = None
                            is_found_checkbox = False

        for field_name, group_checkbox_details in checkbox_field_name__list_group_checkbox_details.items():
            new_group_checkbox_details = []
            unique_merge_field_objs = []
            for group_checkbox_detail in group_checkbox_details:
                if group_checkbox_detail['merge_field'] not in unique_merge_field_objs:
                    new_group_checkbox_details.append(group_checkbox_detail)
                    unique_merge_field_objs.append(group_checkbox_detail['merge_field'])

            checkbox_field_name__list_group_checkbox_details[field_name] = new_group_checkbox_details

        return checkbox_field_name__list_group_checkbox_details

    @property
    def merge_fields(self):
        field_name__infos = {}

        for field_name in self.field_name__elements.keys():
            field_name__infos[field_name] = {
                'type': MERGE_FIELD_TYPE_TEXT,
                'values': None,
                'in_table': False,
                'siblings': None
            }

        for checkbox_field_name, group_checkbox_details in self.checkbox_field_name__list_group_checkbox_details.items():
            field_name__infos[checkbox_field_name] = {
                'type': MERGE_FIELD_TYPE_CHECKBOX,
                'values': self.checkbox_field_name__values[checkbox_field_name],
                'in_table': False,
                'siblings': None
            }

        for field_name in field_name__infos:
            if field_name in self.in_table_field_name__details:
                field_name__infos[field_name]['in_table'] = True
                field_name__infos[field_name]['siblings'] = self.in_table_field_name__details[field_name]['siblings']

                # value cho bảng ẩn dòng
                if field_name not in self.checkbox_field_name__list_group_checkbox_details:
                    field_name__infos[field_name]['values'] = [
                        row_index for row_index in range(
                            len(self.in_table_field_name__details[field_name]['contain_merge_field_rows'])
                        )
                    ]

        return field_name__infos

    def write(self, file):
        # clear attribute is_merge_field
        for part in self.parts.values():
            merge_fields = part.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
            for merge_field in merge_fields:
                text = merge_field.text
                merge_field.clear()
                self.__set_text_for_t_element(element=merge_field, text=text)

        with ZipFile(file, 'w', ZIP_DEFLATED) as output:
            for zip_info in self.zip.filelist:
                filename = zip_info.filename
                if filename in self.parts:
                    xml = etree.tostring(self.parts[filename].getroot())
                    output.writestr(filename, xml)
                else:
                    output.writestr(filename, self.zip.read(zip_info))

    def merge(self, not_in_group_replacements: dict, in_group_replacements: dict):
        for field_name, replacement in not_in_group_replacements.items():
            if self.__is_valid_replace_data_for_checkbox(replace_data=replacement) \
                    and field_name in self.checkbox_field_name__list_group_checkbox_details:
                for group_checkbox_detail in self.checkbox_field_name__list_group_checkbox_details[field_name]:
                    self.__fill_checkbox(
                        merge_field=group_checkbox_detail['merge_field'],
                        checkbox_infos=group_checkbox_detail['checkbox_infos'],
                        need_to_checked_values=replacement['value']
                    )

            elif self.__is_valid_replace_data_for_embossed_table(replace_data=replacement) \
                    and field_name in self.embossed_table_field_name__details:
                detail = self.embossed_table_field_name__details[field_name]
                self.__fill_embossed_table(
                    merge_field=detail['merge_field'],
                    embossed_table_obj=detail['embossed_table'],
                    text=replacement['value']
                )

            elif self.__is_valid_replace_data_for_show_row_in_table(replace_data=replacement) \
                    and field_name in self.in_table_field_name__details:
                detail = self.in_table_field_name__details[field_name]
                self.__fill_show_row_in_table(
                    field_name=field_name,
                    table=detail['table'],
                    contain_merge_field_rows=detail['contain_merge_field_rows'],
                    show_row_indexes=replacement['value']
                )

            else:
                for merge_field in self.field_name__elements.get(field_name, []):
                    self.__fill_text(merge_field=merge_field, text=replacement)

        for field_name_anchor, list_row_replacement in in_group_replacements.items():
            if self.__is_valid_replace_data_for_row(replace_data=list_row_replacement):
                self.__merge_rows(
                    anchor=field_name_anchor,
                    rows=list_row_replacement,
                    in_table_field_name__details=self.in_table_field_name__details
                )
            else:
                for merge_field in self.field_name__elements.get(field_name_anchor, []):
                    self.__fill_text(merge_field=merge_field, text=list_row_replacement)

    def __merge_rows(self, anchor, rows, in_table_field_name__details):
        if anchor not in in_table_field_name__details:
            return None

        table = in_table_field_name__details[anchor]['table']
        need_to_duplicate_rows = in_table_field_name__details[anchor]['rows']

        if len(rows) > 0:
            for row_data in rows:
                last_need_to_duplicate_row = need_to_duplicate_rows[-1]
                for need_to_duplicate_row in need_to_duplicate_rows:
                    new_row = deepcopy(need_to_duplicate_row)
                    last_need_to_duplicate_row.addprevious(new_row)

                    in_new_row_checkbox_field_name__list_group_checkbox_details = \
                        self.__get_checkbox_field_name__list_group_checkbox_details(part=new_row)

                    result = self.__get_in_table_field_name__details__and__embossed_table_field_name__details(part=new_row)
                    in_new_row_field_name__details = result['in_table_field_name__details']
                    in_new_row_embossed_table_field_name__details = result['embossed_table_field_name__details']

                    filled_show_row_in_table_field_names = []
                    merge_fields = new_row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                    for merge_field in merge_fields:
                        field_name = merge_field.text[1:-1]
                        need_to_checked_values_or_text = row_data.get(field_name, None)
                        if need_to_checked_values_or_text is not None:
                            if self.__is_valid_replace_data_for_row(replace_data=need_to_checked_values_or_text):
                                self.__merge_rows(
                                    anchor=field_name,
                                    rows=need_to_checked_values_or_text,
                                    in_table_field_name__details=in_new_row_field_name__details
                                )

                            elif self.__is_valid_replace_data_for_checkbox(
                                    replace_data=need_to_checked_values_or_text
                            ) and field_name in in_new_row_checkbox_field_name__list_group_checkbox_details:

                                for group_checkbox_detail \
                                        in in_new_row_checkbox_field_name__list_group_checkbox_details[field_name]:
                                    self.__fill_checkbox(
                                        merge_field=merge_field,
                                        checkbox_infos=group_checkbox_detail['checkbox_infos'],
                                        need_to_checked_values=need_to_checked_values_or_text['value'],
                                    )

                            elif self.__is_valid_replace_data_for_embossed_table(
                                    replace_data=need_to_checked_values_or_text) \
                                    and field_name in in_new_row_embossed_table_field_name__details:
                                detail = in_new_row_embossed_table_field_name__details[field_name]
                                self.__fill_embossed_table(
                                    merge_field=detail['merge_field'],
                                    embossed_table_obj=detail['embossed_table'],
                                    text=need_to_checked_values_or_text['value']
                                )

                            elif self.__is_valid_replace_data_for_show_row_in_table(replace_data=need_to_checked_values_or_text) \
                                    and field_name in in_new_row_field_name__details:
                                if field_name not in filled_show_row_in_table_field_names:
                                    detail = in_new_row_field_name__details[field_name]
                                    self.__fill_show_row_in_table(
                                        field_name=field_name,
                                        table=detail['table'],
                                        contain_merge_field_rows=detail['contain_merge_field_rows'],
                                        show_row_indexes=need_to_checked_values_or_text['value']
                                    )
                                    filled_show_row_in_table_field_names.append(field_name)
                            else:
                                self.__fill_text(merge_field=merge_field, text=need_to_checked_values_or_text)

            for need_to_duplicate_row in need_to_duplicate_rows:
                table.remove(need_to_duplicate_row)
        else:
            if self.is_remove_row_or_table_has_one_row_when_empty:
                tr__index_column_merged_trs = in_table_field_name__details[anchor]['tr__index_column_merged_trs']

                count_row_in_table = 0
                previous_column_index_merged_trs = None

                for row in table.findall(ELEMENT_PATH_TR):
                    if row in tr__index_column_merged_trs:
                        if tr__index_column_merged_trs[row] != previous_column_index_merged_trs:
                            count_row_in_table += 1
                            previous_column_index_merged_trs = tr__index_column_merged_trs[row]
                    else:
                        count_row_in_table += 1

                # TH1: bảng 1 dòng -> xóa bảng
                if count_row_in_table == NUMBER_ROW_IN_TABLE_IS_ONE:
                    self.__remove_table_and_empty_line(tbl_element=table)

                # TH 2: bảng 2 dòng
                # row trên không chứa merge field nào thì xóa bảng, còn không thì chỉ xóa dòng đó trong bảng
                elif count_row_in_table == NUMBER_ROW_IN_TABLE_IS_TWO:
                    # if first row dows not contain merge field
                    if table.find(ELEMENT_PATH_TR).find(ELEMENT_PATH_RECURSIVE_MERGE_FIELD) is None:
                        self.__remove_table_and_empty_line(tbl_element=table)
                    else:
                        for need_to_duplicate_row in need_to_duplicate_rows:
                            table.remove(need_to_duplicate_row)

                # TH 3: bảng có nhiều hơn 2 dòng -> xóa dòng
                else:
                    for need_to_duplicate_row in need_to_duplicate_rows:
                        table.remove(need_to_duplicate_row)

    def __fill_checkbox(self, merge_field, checkbox_infos, need_to_checked_values):
        # uncheck all checkbox in group
        for checkbox_info in checkbox_infos:
            checkbox_info['checkbox_obj'].text = CHECKBOX_UNCHECKED_TEXT

        is_checked_checkbox = False

        # Trường hợp người dùng truyền rỗng để không chọn checkbox nào
        if not need_to_checked_values:
            is_checked_checkbox = True
        else:
            # Trường hợp có extend text cho checkbox:
            for need_to_checked_value in need_to_checked_values:
                for checkbox_info in checkbox_infos:
                    is_checked_checkbox = False
                    if need_to_checked_value['option'].lower() == checkbox_info['value'].lower():
                        # skip if checkbox is checked
                        if checkbox_info['checkbox_obj'].text == CHECKBOX_CHECKED_TEXT:
                            continue

                        checkbox_info['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT
                        is_checked_checkbox = True

                        # checkbox có dữ liệu mở rộng
                        if need_to_checked_value['extend_data'] is not None:
                            self.__set_text_for_t_element(
                                element=checkbox_info['value_obj'],
                                text=checkbox_info['original_value'].replace(
                                    checkbox_info['value'],
                                    f"{checkbox_info['value']} {need_to_checked_value['extend_data']}"
                                )
                            )

                        # break loop -> continue to next need_to_checked_value
                        break

        # if found checkbox has value the same with need to checked values, then replace merge_field with empty text
        if is_checked_checkbox:
            self.__fill_text(merge_field=merge_field, text='')
        # if NOT, then replace merge_field with need_to_checked_values to easily DEBUG
        else:
            self.__fill_text(
                merge_field=merge_field,
                text=', '.join([need_to_checked_value['option']
                                for need_to_checked_value in need_to_checked_values])
            )

    def __fill_embossed_table(self, merge_field, embossed_table_obj, text):
        cells = embossed_table_obj.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tc')

        # Dữ liệu nhập dài hơn số ô đang có sẽ không được điền vào
        if len(text) > len(cells):
            return

        # Điền text vào ô
        # tc(cell) -> p -> r -> t(text)
        characters = [character for character in text]
        for index in range(len(characters)):
            if characters[index] in EMBOSSED_TABLE_BYPASS_SEPARATOR_CHARACTERS:  # Bỏ qua nếu là ký tự phân cách
                continue
            p_node = cells[index].find(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}p')
            r_node = Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}r')
            t_node = Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}t')
            t_node.text = characters[index]

            r_node.append(t_node)
            p_node.append(r_node)

        self.__fill_text(merge_field=merge_field, text='')

    def __fill_show_row_in_table(self, field_name, table, contain_merge_field_rows, show_row_indexes):
        # nếu truyền sai row indexes thì fill merge field để dễ DEBUG
        if show_row_indexes and max(show_row_indexes) >= len(contain_merge_field_rows):
            in_table_merge_fields = table.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
            for in_table_merge_field in in_table_merge_fields:
                if in_table_merge_field.text[1:-1] == field_name:
                    self.__fill_text(merge_field=in_table_merge_field, text=str(show_row_indexes))
            return None

        # xóa các row không cần hiện
        for row_index, contain_merge_field_row in enumerate(contain_merge_field_rows):
            if row_index not in show_row_indexes:
                if isinstance(contain_merge_field_row, list):
                    for merged_contain_merge_field_row in contain_merge_field_row:
                        table.remove(merged_contain_merge_field_row)
                else:
                    table.remove(contain_merge_field_row)

        # fill rỗng cho các field đánh dấu của các row hiện
        in_table_merge_fields = table.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
        for in_table_merge_field in in_table_merge_fields:
            if in_table_merge_field.text[1:-1] == field_name:
                self.__fill_text(merge_field=in_table_merge_field, text='')

        # kiểm tra bảng rỗng
        is_table_empty = True
        for _row in table.getchildren():
            if _row.tag == ELEMENT_PATH_TR:
                is_table_empty = False
                break

        if is_table_empty:
            self.__remove_table_and_empty_line(tbl_element=table)

        return None

    def __fill_text(self, merge_field, text):
        text = str(text) or ''  # text might be None

        # remove blank line: nếu fill rỗng thì xóa p đó nếu p chỉ có 1 t chứa chữ cái và bullet (nếu có)
        if not text:
            # t -> r -> p
            r_element = merge_field.getparent()

            p_element = r_element.getparent()
            t_elements = p_element.findall(ELEMENT_PATH_RECURSIVE_T)

            number_of_t_element_contains_not_blank = 0
            for t_element in t_elements:
                # t_element.text có thể bằng None, '', '  ', 'abc'
                # cần đếm những t_element.text giống 'abc' thì
                if t_element.text != '' and t_element.text and not t_element.text.isspace():
                    number_of_t_element_contains_not_blank += 1

            # t_element đó chỉ chứa merge_field
            if number_of_t_element_contains_not_blank == 1:
                p_parent_element = p_element.getparent()
                # nếu merge field nằm trong 1 table
                if 'tc' in p_parent_element.tag:
                    spacing_tag_in_p = p_element.find(f".//{{{NAMESPACE_WORDPROCESSINGML}}}spacing")

                    p_parent_element.remove(p_element)

                    # TODO: thay vì append thẻ p mới thì xóa bullet ra khỏi p cũ (nếu có)
                    # nếu row ko còn thẻ p thì cần add lại thẻ p nếu ko row sẽ mất border
                    if len(p_parent_element.getchildren()) == 1:  # chỉ còn tcPr
                        new_p_tag = Element(f"{{{NAMESPACE_WORDPROCESSINGML}}}p")
                        new_ppr_tag = Element(f"{{{NAMESPACE_WORDPROCESSINGML}}}pPr")

                        if spacing_tag_in_p is not None:
                            new_spacing_tag_in_p = deepcopy(spacing_tag_in_p)
                            new_ppr_tag.append(new_spacing_tag_in_p)

                        new_p_tag.append(new_ppr_tag)

                        p_parent_element.append(new_p_tag)

                # nếu merge field không nằm trong table
                else:
                    p_parent_element.remove(p_element)
            else:
                merge_field.text = ''

        # fill text to merge field
        else:
            if self.change_color_flag:
                self.__set_color_for_text(
                    element=merge_field,
                    from_color_value=FROM_COLOR_VALUE,
                    to_color_value=TO_COLOR_VALUE
                )

            # preserve new lines in replacement text
            text_parts = str(text).split('\n')

            if len(text_parts) == 1:
                self.__set_text_for_t_element(element=merge_field, text=text_parts[0])
            else:
                p_element = merge_field.getparent().getparent()

                p_element_only_contains_merge_field = deepcopy(p_element)
                t_elements = p_element_only_contains_merge_field.findall(ELEMENT_PATH_RECURSIVE_T)
                for t_element in t_elements:
                    if t_element.text and (not t_element.get('is_merge_field') or t_element.text != merge_field.text):
                        r_element = t_element.getparent()
                        r_element.remove(t_element)

                # TODO: xử lý trường hợp có t_element phía sau merge_field trong p_element

                self.__set_text_for_t_element(element=merge_field, text=text_parts[0])

                for text_part in reversed(text_parts[1:]):
                    new_p_element = deepcopy(p_element_only_contains_merge_field)
                    self.__set_text_for_t_element(
                        element=new_p_element.find(ELEMENT_PATH_RECURSIVE_MERGE_FIELD),
                        text=text_part
                    )
                    p_element.addnext(new_p_element)

    def __remove_table_and_empty_line(self, tbl_element):
        # Kiểm tra, xử lý dòng trống nếu như xóa bảng
        previous_p_element = tbl_element.getprevious()
        next_p_element = tbl_element.getnext()

        # Bảng ngay đầu file
        if previous_p_element is None and self.__is_blank_line(next_p_element):
            next_p_element_parent = next_p_element.getparent()
            next_p_element_parent.remove(next_p_element)

        # Bảng ngay cuối file hoặc TH thường
        elif (next_p_element is None is None and self.__is_blank_line(previous_p_element)) or \
                (self.__is_blank_line(previous_p_element) and self.__is_blank_line(next_p_element)):
            previous_p_element_parent = previous_p_element.getparent()
            previous_p_element_parent.remove(previous_p_element)

        parent = tbl_element.getparent()
        parent.remove(tbl_element)

    @staticmethod
    def __is_blank_line(element):
        if element is not None and element.tag == f'{{{NAMESPACE_WORDPROCESSINGML}}}p':
            text_in_p_element = ''.join(element.itertext())
            if not text_in_p_element or text_in_p_element.isspace():
                return True
        return False

    @staticmethod
    def __set_color_for_text(element, from_color_value, to_color_value):
        parent_element = element.getparent()
        format_color_paragraph = parent_element.find(f".//{{{NAMESPACE_WORDPROCESSINGML}}}color")

        if format_color_paragraph is not None \
                and format_color_paragraph.get(f'{{{NAMESPACE_WORDPROCESSINGML}}}val') == from_color_value:
            format_color_paragraph.set(f"{{{NAMESPACE_WORDPROCESSINGML}}}val", to_color_value)

    @staticmethod
    def __set_text_for_t_element(element, text):
        element.text = text
        if text.startswith(' ') or text.endswith(' '):
            element.set(XML_SPACE_ATTRIBUTE, 'preserve')

    @staticmethod
    def __is_valid_replace_data_for_checkbox(replace_data):
        if isinstance(replace_data, dict) \
                and isinstance(replace_data.get('value'), list) \
                and replace_data.get('type') == MERGE_FIELD_TYPE_CHECKBOX:

            for value in replace_data['value']:
                if not(
                        isinstance(value, dict) and 'option' in value and isinstance(value['option'], str) and 'extend_data' in value and (value['extend_data'] is None or isinstance(value['extend_data'], str))
                ):
                    return False
            return True
        return False

    @staticmethod
    def __is_valid_replace_data_for_embossed_table(replace_data):
        if isinstance(replace_data, dict) \
                and isinstance(replace_data.get('value'), str) \
                and replace_data.get('type') == MERGE_FIELD_TYPE_EMBOSSED_TABLE:
            return True
        return False

    @staticmethod
    def __is_valid_replace_data_for_show_row_in_table(replace_data):
        if isinstance(replace_data, dict) \
                and isinstance(replace_data.get('value'), list) \
                and (len(replace_data['value']) == 0 or all(isinstance(row_index, int) for row_index in replace_data['value'])) \
                and replace_data.get('type') == MERGE_FIELD_TYPE_SHOW_ROW_IN_TABLE:
            return True
        return False

    @staticmethod
    def __is_valid_replace_data_for_row(replace_data):
        if isinstance(replace_data, list) \
                and (len(replace_data) == 0 or all(isinstance(row, dict) for row in replace_data)):
            return True
        return False

    def __enter__(self):
        return self

    def __exit__(self, exc_type, value, traceback):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None
