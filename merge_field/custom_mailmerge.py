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
MERGE_FIELD_TYPE_EXTEND_TEXT_CHECKBOX = "extend_checkbox"

EMBOSSED_TABLE_BYPASS_SEPARATOR_CHARACTERS = ["-", ":"]

NUMBER_ROW_IN_TABLE_IS_ONE = 1
NUMBER_ROW_IN_TABLE_IS_TWO = 2

FROM_COLOR_VALUE = 'FF0000'
TO_COLOR_VALUE = '000000'


class MergeField:
    def __init__(self, file, is_remove_row_or_table_has_one_row_when_empty=True):
        self.zip: ZipFile = ZipFile(file)
        self.parts: dict = {}
        self.field_name__elements: dict = {}
        self.in_table_field_name__details: dict = {}
        self.checkbox_field_name__list_group_checkbox_details: dict = {}
        self.checkbox_field_name__values: dict = {}
        self.embossed_table_field_name__details: dict = {}

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

        t_elements = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')

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
            if is_parse_in_table_field_name:
                for row in table:
                    merge_fields = row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                    for merge_field in merge_fields:
                        current_siblings = [merge_field.text[1:-1] for merge_field in merge_fields]

                        if not in_table_field_name__details.get(merge_field.text[1:-1]):
                            siblings = current_siblings
                        else:
                            siblings = list(
                                set(in_table_field_name__details[merge_field.text[1:-1]]['siblings'] + current_siblings)
                            )

                        in_table_field_name__details[merge_field.text[1:-1]] = {
                            'table': table,
                            'row': row,
                            'siblings': sorted(siblings)
                        }

            # Bảng khung dập nổi
            # tbl -> tc -> tbl (bảng dập nổi) + t (merge_field)
            text_element_in_tables = table.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')
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
                            'checkbox_obj': sdt_element.find(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t'),
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

    def merge(self, not_in_group_replacements: dict, in_group_replacements: dict, change_color_flag: bool = False):
        for field_name, replacement in not_in_group_replacements.items():
            if self.__is_valid_need_to_checked_values_for_checkbox(value_need_to_check=replacement) \
                    and field_name in self.checkbox_field_name__list_group_checkbox_details:
                self.__merge_checkbox(field_name=field_name, need_to_checked_values=replacement, change_color_flag=change_color_flag)

            elif self.__is_valid_need_to_checked_values_for_extend_text_checkbox(value_need_to_check=replacement) \
                    and field_name in self.checkbox_field_name__list_group_checkbox_details:
                self.__merge_checkbox(field_name=field_name, need_to_checked_values=replacement,
                                      is_add_extend_text=True, change_color_flag=change_color_flag)

            elif self.__is_valid_values_for_embossed_table(value_need_to_check=replacement) \
                    and field_name in self.embossed_table_field_name__details:
                self.__merge_embossed_table(field_name=field_name, embossed_table_info=replacement,
                                            change_color_flag=change_color_flag)
            else:
                self.__merge_field(field_name=field_name, text=replacement, change_color_flag=change_color_flag)

        for field_name_anchor, list_row_replacement in in_group_replacements.items():
            if self.__is_valid_values_for_row(value_need_to_check=list_row_replacement):
                self.__merge_rows(anchor=field_name_anchor, rows=list_row_replacement,
                                  change_color_flag=change_color_flag)
            else:
                self.__merge_field(field_name=field_name_anchor, text=list_row_replacement,
                                   change_color_flag=change_color_flag)

    def __merge_rows(self, anchor, rows, change_color_flag=False):
        if anchor not in self.in_table_field_name__details:
            return None

        table = self.in_table_field_name__details[anchor]['table']
        row = self.in_table_field_name__details[anchor]['row']

        if len(rows) > 0:
            for row_data in rows:
                new_row = deepcopy(row)
                row.addprevious(new_row)
                in_new_row_checkbox_field_name__list_group_checkbox_details = \
                    self.__get_checkbox_field_name__list_group_checkbox_details(part=new_row)

                result = self.__get_in_table_field_name__details__and__embossed_table_field_name__details(
                    part=new_row,
                    is_parse_in_table_field_name=False
                )
                in_new_row_embossed_table_field_name__details = result['embossed_table_field_name__details']

                merge_fields = new_row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                for merge_field in merge_fields:
                    field_name = merge_field.text[1:-1]
                    need_to_checked_values_or_text = row_data.get(field_name, None)
                    if need_to_checked_values_or_text is not None:
                        if self.__is_valid_need_to_checked_values_for_checkbox(
                                value_need_to_check=need_to_checked_values_or_text
                        ) and field_name in in_new_row_checkbox_field_name__list_group_checkbox_details:

                            for group_checkbox_detail \
                                    in in_new_row_checkbox_field_name__list_group_checkbox_details[field_name]:
                                self.__fill_checkbox(
                                    field_name=field_name,
                                    merge_field=merge_field,
                                    parent=merge_field.getparent(),
                                    checkbox_infos=group_checkbox_detail['checkbox_infos'],
                                    need_to_checked_values=need_to_checked_values_or_text,
                                    change_color_flag=change_color_flag
                                )

                        elif self.__is_valid_need_to_checked_values_for_extend_text_checkbox(
                                value_need_to_check=need_to_checked_values_or_text
                        ) and field_name in in_new_row_checkbox_field_name__list_group_checkbox_details:

                            for group_checkbox_detail \
                                    in in_new_row_checkbox_field_name__list_group_checkbox_details[field_name]:
                                self.__fill_checkbox(
                                    field_name=field_name,
                                    merge_field=merge_field,
                                    parent=merge_field.getparent(),
                                    checkbox_infos=group_checkbox_detail['checkbox_infos'],
                                    need_to_checked_values=need_to_checked_values_or_text,
                                    is_add_extend_text=True,
                                    change_color_flag=change_color_flag
                                )

                        elif self.__is_valid_values_for_embossed_table(
                                value_need_to_check=need_to_checked_values_or_text) \
                                and field_name in in_new_row_embossed_table_field_name__details:
                            detail = in_new_row_embossed_table_field_name__details[field_name]
                            self.__fill_embossed_table(
                                merge_field=detail['merge_field'],
                                parent=detail['merge_field'].getparent(),
                                embossed_table_obj=detail['embossed_table'],
                                need_to_checked_values=need_to_checked_values_or_text,
                                change_color_flag=change_color_flag
                            )

                        else:
                            # wt -> wr -> wp -> wtc -> wtr
                            wp_of_merge_field = merge_field.getparent().getparent()
                            columns = new_row.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tc')

                            # trong từng cột của 1 hàng cột và trong cột đó chỉ có duy nhất mergefield và nhập vào empty string
                            # thì tiến hành xóa đi đoạn text trong column

                            for column in columns:
                                if wp_of_merge_field in column:
                                    child_t_elements_dont_have_empty_in_p = []
                                    child_elements_in_p = wp_of_merge_field.findall(
                                        f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')
                                    for w_text in child_elements_in_p:
                                        if w_text.text != '':
                                            child_t_elements_dont_have_empty_in_p.append(w_text)
                                    if not need_to_checked_values_or_text and len(
                                            child_t_elements_dont_have_empty_in_p) == 1 \
                                            and child_t_elements_dont_have_empty_in_p[0].get("is_merge_field"):

                                        space_tag_in_p = wp_of_merge_field.find(
                                            f".//{{{NAMESPACE_WORDPROCESSINGML}}}spacing")
                                        new_space_tag_in_p = deepcopy(space_tag_in_p)
                                        list_attr_of_p_tag = space_tag_in_p.items()

                                        column.remove(wp_of_merge_field)
                                        # set lại attribute kích thước của row sau khi xóa đi 1 p trong row
                                        new_row.find(f".//{{{NAMESPACE_WORDPROCESSINGML}}}trHeight").set(
                                            f"{{{NAMESPACE_WORDPROCESSINGML}}}val", "0")
                                        new_row.find(f".//{{{NAMESPACE_WORDPROCESSINGML}}}trHeight").set(
                                            f"{{{NAMESPACE_WORDPROCESSINGML}}}hRule", "auto")

                                        is_exits_p = column.findall(f".//{{{NAMESPACE_WORDPROCESSINGML}}}p")
                                        if not is_exits_p:
                                            # set laị kích thước của row sau khi tạo mới lại thẻ p bằng kích thước
                                            # của thẻ p trước đó

                                            # wppr ->wp-> wtc
                                            p_tag = Element(f"{{{NAMESPACE_WORDPROCESSINGML}}}p")
                                            ppr_tag = Element(f"{{{NAMESPACE_WORDPROCESSINGML}}}pPr")
                                            ppr_tag.append(new_space_tag_in_p)
                                            p_tag.append(ppr_tag)
                                            # nếu row ko còn thẻ p thì cần add lại thẻ p nếu ko row sẽ mất border
                                            column.append(p_tag)
                                            # set lại kích thước của row
                                            new_row.find(f".//{{{NAMESPACE_WORDPROCESSINGML}}}trHeight").set(
                                                f"{{{NAMESPACE_WORDPROCESSINGML}}}hRule", "auto")

                                    else:
                                        self.__fill_text(
                                            merge_field=merge_field,
                                            parent=merge_field.getparent(),
                                            text=need_to_checked_values_or_text,
                                            change_color_flag=change_color_flag
                                        )

            table.remove(row)
        else:
            if self.is_remove_row_or_table_has_one_row_when_empty:
                count_row_in_table = 0
                for _row in table.getchildren():
                    if _row.tag == f'{{{NAMESPACE_WORDPROCESSINGML}}}tr':
                        count_row_in_table += 1

                # TH1: bảng 1 dòng -> xóa bảng
                if count_row_in_table == NUMBER_ROW_IN_TABLE_IS_ONE:
                    parent = table.getparent()
                    parent.remove(table)

                # TH 2: bảng 2 dòng
                # row trên không chứa merge field nào thì xóa bảng, còn không thì chỉ xóa dòng đó trong bảng
                elif count_row_in_table == NUMBER_ROW_IN_TABLE_IS_TWO:
                    first_row_in_table = table.find(f'{{{NAMESPACE_WORDPROCESSINGML}}}tr')
                    merge_field_in_first_row_in_tables = first_row_in_table.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                    if not merge_field_in_first_row_in_tables:
                        parent = table.getparent()
                        parent.remove(table)
                    else:
                        table.remove(row)

                # TH 3: bảng có nhiều hơn 2 dòng -> xóa dòng
                else:
                    table.remove(row)

    def __merge_field(self, field_name, text, change_color_flag=False):
        for merge_field in self.field_name__elements.get(field_name, []):
            parent = merge_field.getparent()
            # remove blank line
            if not text \
                    and merge_field.text[1:-1] not in self.in_table_field_name__details \
                    and len(parent.getchildren()) == 2:
                grand_parent = parent.getparent()

                is_contain_num_format = False

                format_element = parent[0]  # pPr element
                for detail_format_element in format_element:
                    # if have tag bulleted or numbered list format then by pass
                    if 'numPr' in detail_format_element.tag:
                        is_contain_num_format = True
                        break

                if not is_contain_num_format:
                    great_grand_parent = grand_parent.getparent()
                    great_grand_parent.remove(grand_parent)

            else:
                # nếu truyền rỗng thì xóa hàng đó nếu hàng chỉ có 1 field hoặc có dấu "-", "bullet", "number"
                # wt - wr - wp
                grand_parent = parent.getparent()
                child_elements_in_p = grand_parent.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')

                child_t_elements_dont_have_empty_in_p = []
                # không xóa trực tiếp wt = '' trong child_elements_in_p vì sẽ làm thay đổi vị trí các phần tử trong list
                for w_text in child_elements_in_p:
                    if w_text.text != '':
                        child_t_elements_dont_have_empty_in_p.append(w_text)

                if not text and len(child_t_elements_dont_have_empty_in_p) == 1:
                    grand_grand_parent = grand_parent.getparent()
                    # nếu fill data nằm trong 1 table nhưng truyền lên theo cách fill của field ko nằm trong table
                    if 'tc' in grand_grand_parent.tag:
                        spacing_tag_in_p = grand_parent.find(f".//{{{NAMESPACE_WORDPROCESSINGML}}}spacing")
                        new_spacing_tag_in_p = deepcopy(spacing_tag_in_p)

                        grand_grand_parent.remove(grand_parent)

                        p_tag = Element(f"{{{NAMESPACE_WORDPROCESSINGML}}}p")
                        ppr_tag = Element(f"{{{NAMESPACE_WORDPROCESSINGML}}}pPr")
                        ppr_tag.append(new_spacing_tag_in_p)
                        p_tag.append(ppr_tag)
                        # nếu row ko còn thẻ p thì cần add lại thẻ p nếu ko row sẽ mất border
                        grand_grand_parent.append(p_tag)

                    else:
                        # nếu là fill data cho 1 field ko nằm trong table
                        grand_grand_parent.remove(grand_parent)
                else:
                    self.__fill_text(
                        merge_field=merge_field,
                        parent=parent,
                        text=text,
                        change_color_flag=change_color_flag
                    )

    def __merge_embossed_table(self, field_name, embossed_table_info, change_color_flag=False):

        detail = self.embossed_table_field_name__details[field_name]
        self.__fill_embossed_table(
            merge_field=detail['merge_field'],
            parent=detail['merge_field'].getparent(),
            embossed_table_obj=detail['embossed_table'],
            need_to_checked_values=embossed_table_info,
            change_color_flag=change_color_flag
        )

    def __merge_checkbox(self, field_name, need_to_checked_values, is_add_extend_text=False, change_color_flag=False):
        for group_checkbox_detail in self.checkbox_field_name__list_group_checkbox_details[field_name]:
            self.__fill_checkbox(
                field_name=field_name,
                merge_field=group_checkbox_detail['merge_field'],
                parent=group_checkbox_detail['merge_field'].getparent(),
                checkbox_infos=group_checkbox_detail['checkbox_infos'],
                need_to_checked_values=need_to_checked_values,
                is_add_extend_text=is_add_extend_text,
                change_color_flag=change_color_flag
            )

    def __fill_checkbox(self, field_name, merge_field, parent, checkbox_infos, need_to_checked_values,
                        change_color_flag=False, is_add_extend_text=False):
        for checkbox_info in checkbox_infos:
            # uncheck all checkbox in group
            checkbox_info['checkbox_obj'].text = CHECKBOX_UNCHECKED_TEXT

        exclude_values = []  # exclude values for 'Khác' checkbox
        is_checked_checkbox = False

        # Trường hợp merge field chỉ có 1 checkbox thì có thể điền True để tick vào checkbox thay vì điền ['abc']
        if len(checkbox_infos) == 1 and need_to_checked_values is True \
                and checkbox_infos[0]['checkbox_obj'].text == CHECKBOX_UNCHECKED_TEXT:
            checkbox_infos[0]['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT
            is_checked_checkbox = True
            others_need_to_checked_values = None

        # Trường hợp có extend text cho checkbox:
        elif is_add_extend_text:
            # trường hợp extend data cho duy nhất 1 checkbox sử dụng true để điền giá trị
            if len(checkbox_infos) == 1 and \
                    isinstance(need_to_checked_values['value'], dict) and \
                    need_to_checked_values['value']['option'] is True:
                checkbox_infos[0]['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT

                is_checked_checkbox = True
                others_need_to_checked_values = None

                self.__set_text_for_t_element(
                    element=checkbox_infos[0]['value_obj'],
                    text=checkbox_infos[0]['original_value'].replace(
                        checkbox_infos[0]['value'],
                        f"{checkbox_infos[0]['value']} {need_to_checked_values['value']['extend_data']}"
                    )
                )

            else:
                for need_to_checked_value in need_to_checked_values['value']:
                    for checkbox_info in checkbox_infos:
                        is_checked_checkbox = False
                        if need_to_checked_value['option'].lower() == checkbox_info['value'].lower():
                            # skip if checkbox is checked
                            if checkbox_info['checkbox_obj'].text == CHECKBOX_CHECKED_TEXT:
                                continue

                            checkbox_info['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT
                            is_checked_checkbox = True

                            self.__set_text_for_t_element(
                                element=checkbox_info['value_obj'],
                                text=checkbox_info['original_value'].replace(
                                    checkbox_info['value'],
                                    f"{checkbox_info['value']} {need_to_checked_value['extend_data']}"
                                )
                            )

                            exclude_values.append(need_to_checked_value['option'])

                            # break loop -> continue to next need_to_checked_value
                            break

                for value in self.checkbox_field_name__values[field_name]:
                    exclude_values.append(value)

                options = [need_to_checked_value['option'] for need_to_checked_value in need_to_checked_values['value']]

                others_need_to_checked_values = list(
                    set(options).difference(set(exclude_values))
                )

        # Trường hợp chỉ là checkbox bình thường, không có extend text:
        else:
            for need_to_checked_value in need_to_checked_values:
                for checkbox_info in checkbox_infos:
                    is_checked_checkbox = False
                    if need_to_checked_value.lower() == checkbox_info['value'].lower():
                        # skip if checkbox is checked
                        if checkbox_info['checkbox_obj'].text == CHECKBOX_CHECKED_TEXT:
                            continue

                        checkbox_info['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT
                        is_checked_checkbox = True

                        exclude_values.append(need_to_checked_value)

                        # break loop -> continue to next need_to_checked_value
                        break

            for value in self.checkbox_field_name__values[field_name]:
                exclude_values.append(value)

            others_need_to_checked_values = list(
                set(need_to_checked_values).difference(set(exclude_values))
            )

        ################################################################################################################
        # SPECIAL: Fill to checkbox "Khác:"
        ################################################################################################################
        # replace merge_field with need to checked values and check 'Khác' checkbox (if exist)
        if others_need_to_checked_values:
            for checkbox_info in checkbox_infos:
                if checkbox_info['value'].lower() in ['khác:', 'khác/other:']:
                    checkbox_info['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT
                    is_checked_checkbox = True

                    self.__fill_text(
                        merge_field=checkbox_info['value_obj'],
                        parent=checkbox_info['value_obj'].getparent(),
                        text=checkbox_info['original_value'].replace(
                            checkbox_info['value'],
                            f"{checkbox_info['value']} {', '.join(others_need_to_checked_values)}"

                        ),
                        change_color_flag=change_color_flag
                    )
                    # self.__set_text_for_t_element(
                    #     element=checkbox_info['value_obj'],
                    #     text=checkbox_info['original_value'].replace(
                    #         checkbox_info['value'],
                    #         f"{checkbox_info['value']} {', '.join(others_need_to_checked_values)}"
                    #     )
                    # )
        ################################################################################################################
        # if found checkbox has value the same with need to checked values, then replace merge_field with empty text
        if is_checked_checkbox:
            self.__fill_text(
                merge_field=merge_field,
                parent=parent,
                text='',
                change_color_flag=change_color_flag
            )
        # if NOT, then replace merge_field with need_to_checked_values to easily DEBUG
        else:
            self.__fill_text(
                merge_field=merge_field,
                parent=parent,
                text=', '.join(others_need_to_checked_values),
                change_color_flag=change_color_flag
            )

    def __fill_text(self, merge_field, parent, text, change_color_flag=False):
        text = text or ''  # text might be None
        print(change_color_flag)
        if change_color_flag:
            print(" da vao ")
            self.__set_color_for_text(merge_field, FROM_COLOR_VALUE, TO_COLOR_VALUE)
        # preserve new lines in replacement text
        text_parts = str(text).split('\n')
        nodes = []

        for text_part in text_parts:
            text_node = Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}t')
            self.__set_text_for_t_element(element=text_node, text=text_part)
            nodes.append(text_node)

            nodes.append(Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}br'))

        nodes.pop()  # remove last br element

        for node in reversed(nodes):
            merge_field.addnext(node)

        parent.remove(merge_field)  # remove old merge field element due to it is replaced by new text in nodes

    def __fill_embossed_table(self, merge_field, parent, embossed_table_obj, need_to_checked_values,
                              change_color_flag=False):
        text = str(need_to_checked_values['value'])

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

        self.__fill_text(
            merge_field=merge_field,
            parent=parent,
            text="",
            change_color_flag=change_color_flag
        )

    @staticmethod
    def __set_color_for_text(element, from_color_value, to_color_value):
        p_element = element.getparent().getparent()
        # color bao gồm color của bullet(w:numPr) và color của text (w:t)
        list_format_color_pararaph = p_element.findall(f".//{{{NAMESPACE_WORDPROCESSINGML}}}color")
        for format_color_pararaph in list_format_color_pararaph:
            if format_color_pararaph is not None \
                    and format_color_pararaph.get(f'{{{NAMESPACE_WORDPROCESSINGML}}}val') == from_color_value:
                print("da vao ")
                format_color_pararaph.set(f"{{{NAMESPACE_WORDPROCESSINGML}}}val", to_color_value)

    @staticmethod
    def __set_text_for_t_element(element, text):
        element.text = text
        if text.startswith(' ') or text.endswith(' '):
            element.set(XML_SPACE_ATTRIBUTE, 'preserve')

    @staticmethod
    def __is_valid_need_to_checked_values_for_checkbox(value_need_to_check):
        if isinstance(value_need_to_check, bool) \
                or (isinstance(value_need_to_check, list) and (not value_need_to_check or (
                value_need_to_check and all(isinstance(value, str) for value in value_need_to_check)))):
            return True
        return False

    @staticmethod
    def __is_valid_need_to_checked_values_for_extend_text_checkbox(value_need_to_check):
        if isinstance(value_need_to_check, dict) and \
                ('type' in value_need_to_check and value_need_to_check[
                    'type'] == MERGE_FIELD_TYPE_EXTEND_TEXT_CHECKBOX) and 'value' in value_need_to_check:

            # kiểm tra 1 checkbox có extend data
            if isinstance(value_need_to_check['value'], dict) \
                    and isinstance(value_need_to_check['value']['option'], bool) and \
                    ('extend_data' in value_need_to_check['value'] and isinstance(
                        value_need_to_check['value']['extend_data'], str)):
                return True

            # kiểm tra check box có extend data
            if not isinstance(value_need_to_check['value'], list):
                return False
            for value in value_need_to_check['value']:
                if ('option' not in value) or ('option' in value and not isinstance(value['option'], str)) or \
                        ('extend_data' not in value) or (
                        'extend_data' in value and not isinstance(value['extend_data'], str)):
                    return False
            return True
        return False

    @staticmethod
    def __is_valid_values_for_embossed_table(value_need_to_check):
        if isinstance(value_need_to_check, dict) and \
                ('type' in value_need_to_check and value_need_to_check['type'] == MERGE_FIELD_TYPE_EMBOSSED_TABLE):
            return True
        return False

    @staticmethod
    def __is_valid_values_for_row(value_need_to_check):
        if isinstance(value_need_to_check, list) \
                and (not value_need_to_check or (value_need_to_check and all(isinstance(value, dict)
                                                                             for value in value_need_to_check))):
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
