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

OPEN_TAG = '«'
CLOSE_TAG = '»'

CHECKBOX_CHECKED_TEXT = '☒'
CHECKBOX_UNCHECKED_TEXT = '☐'

MERGE_FIELD_TYPE_TEXT = 'text'
MERGE_FIELD_TYPE_CHECKBOX = 'checkbox'
MERGE_FIELD_TYPE_EMBOSSED_TABLE = "embossed_table"

BYPASS_SEPARATOR_CHARACTERS = ["-", ":"]

MERGE_FIELD_TYPE_EXTEND_TEXT_CHECKBOX = "extend_checkbox"


class MergeField:
    def __init__(self, file, is_remove_empty_table=False):
        self.zip: ZipFile = ZipFile(file)
        self.parts: dict = {}
        self.field_name__elements: dict = {}
        self.in_table_field_name__details: dict = {}
        self.checkbox_field_name__list_group_checkbox_details: dict = {}
        self.checkbox_field_name__values: dict = {}
        self.embossed_table_field_name__details: dict = {}

        self.is_remove_empty_table: bool = is_remove_empty_table

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
            # for key, location in self.field_name__elements.items():
            #     print("aaaaaaa", key, etree.tostring(location[0]))
            print("pppppppppppppppppppp", self.field_name__elements)
        except Exception as ex:
            self.zip.close()
            raise ex

    def __parse_merge_fields(self, part):
        is_found_merge_field = False
        is_first_element_contains_open_tag = False
        part_of_field_names = []
        part_of_merge_fields = []
        list_part_of_element_in_two_wr = []

        t_elements = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')

        for t_element in t_elements:
            # print('kkkkkkkkk', t_element.text)
            # Trường hợp thẻ không có text, vd: <w:t xml:space="preserve"/>
            # cần gán lại là string rỗng '' để không chạy lỗi khi duyệt t_element.text
            if t_element.text is None:
                t_element.text = ''


            if CLOSE_TAG in t_element.text and not is_found_merge_field and list_part_of_element_in_two_wr:
                print("lllllllllllllllllll", t_element.text)
                self.__set_text_for_t_element(element= t_element, text= f'{list_part_of_element_in_two_wr[0]}{t_element.text}')
                print("ggggggggggggg",  t_element.text)
                field_name = t_element.text[1:-1]
                if field_name not in  self.field_name__elements:
                    self.field_name__elements[field_name] = []
                self.field_name__elements[field_name].append(t_element)
                list_part_of_element_in_two_wr = []
                continue
            if OPEN_TAG not in t_element.text and not is_found_merge_field:
                continue

            if OPEN_TAG in t_element.text and not is_found_merge_field:
                is_found_merge_field = True
                is_first_element_contains_open_tag = True

                # can not found close tag in previous merge field
                if OPEN_TAG != CLOSE_TAG:
                    part_of_field_names = []
                    part_of_merge_fields = []

            part_of_field_names.append(t_element.text)
            # add element between open and close tag to list need to delete
            part_of_merge_fields.append(t_element)
            print("part_of_field_names", part_of_field_names)
            # useful when open tag = close tag -> prevent handle found close tag when only has open tag
            if is_first_element_contains_open_tag:
                if (OPEN_TAG != CLOSE_TAG and CLOSE_TAG not in t_element.text) \
                        or (OPEN_TAG == CLOSE_TAG and t_element.text.count(CLOSE_TAG) != 2):
                    is_first_element_contains_open_tag = False
                    continue

            print("t_element",t_element.text)
            # handle when found close tag
            if CLOSE_TAG in t_element.text:
                print("da vao")
                # there are some text before OPEN_TAG -> add new element contains text before OPEN_TAG
                if not part_of_field_names[0].startswith(OPEN_TAG):
                    remainder_and_list_first_part_of_field_name = part_of_field_names[0].split(OPEN_TAG)
                    print("remainder_and_list_first_part_of_field_name", remainder_and_list_first_part_of_field_name)
                    # print("remainder_and_list_first_part_of_field_name", remainder_and_list_first_part_of_field_name)
                    remainder = remainder_and_list_first_part_of_field_name[0]
                    print("remainder", remainder)
                    first_part_of_field_name = f'{OPEN_TAG}'.join(
                        remainder_and_list_first_part_of_field_name[1:]
                    )
                    part_of_field_names[0] = f'{OPEN_TAG}{first_part_of_field_name}'

                    ###########################################################################################
                    if CLOSE_TAG  in remainder:
                        pass
                    else:
                        remainder_t_element = deepcopy(t_element)
                        print('remainder_t_element', remainder_t_element.text)
                        self.__set_text_for_t_element(element=remainder_t_element, text=remainder)
                        t_element.addprevious(remainder_t_element)
                    #########################################################################################
                # there are some text after CLOSE_TAG -> add new element contains text after CLOSE_TAG
                if not part_of_field_names[-1].endswith(CLOSE_TAG):
                    list_last_part_of_field_name_and_remainder = part_of_field_names[-1].split(CLOSE_TAG)
                    print("list_last_part_of_field_name_and_remainder", list_last_part_of_field_name_and_remainder)
                    ###########################################################################
                    # if OPEN_TAG in list_last_part_of_field_name_and_remainder[-1]:
                    #     last_part_of_field_name = f'{CLOSE_TAG}'.join(
                    #         list_last_part_of_field_name_and_remainder
                    #     )
                    # else:
                    last_part_of_field_name = f'{CLOSE_TAG}'.join(
                        list_last_part_of_field_name_and_remainder[:-1]
                    )
                        ##########################################################################3
                    print("last_part_of_field_name",last_part_of_field_name)
                    remainder = list_last_part_of_field_name_and_remainder[-1]
                    print("hhhhhhhhhhhh ", f'{last_part_of_field_name}{CLOSE_TAG}')
                    part_of_field_names[-1] = f'{last_part_of_field_name}{CLOSE_TAG}'
                    ###############################################################################
                    if OPEN_TAG in remainder:
                        list_part_of_element_in_two_wr.append(remainder)
                        
                    else:
                        remainder_t_element = deepcopy(t_element)
                        self.__set_text_for_t_element(element=remainder_t_element, text=remainder)
                        t_element.addnext(remainder_t_element)
                        ###########################################################################
                print("ffffffffffff", part_of_field_names)
                print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$",list_part_of_element_in_two_wr)
                # xóa khoảng trắng bên trong các merge field
                field_name_contain_open_close_tag = re.sub(r'«\s*(.*?)\s*»', r'«\1»', ''.join(part_of_field_names))

                # remove open and close tag
                field_name = field_name_contain_open_close_tag[1:-1]
                print(field_name)
                # handle case when multi merge field in one t element
                # EX: «field1» something between «field2».....«field3»
                # after remove open and close tag: field1» something between «field2».....«field3
                if OPEN_TAG in field_name and CLOSE_TAG in field_name:
                    for something_and_field_name in field_name.split(CLOSE_TAG):
                        # if not contain OPEN_TAG -> It is merge field
                        if OPEN_TAG not in something_and_field_name:
                            new_field_name = something_and_field_name
                        else:
                            something_between, new_field_name = something_and_field_name.split(OPEN_TAG)

                            new_t_element = deepcopy(t_element)
                            self.__set_text_for_t_element(element=new_t_element, text=something_between)

                            t_element.addprevious(new_t_element)

                        new_merge_field = deepcopy(t_element)
                        new_merge_field.text = f'{OPEN_TAG}{new_field_name}{CLOSE_TAG}'
                        # set new attribute named is_merge_field -> easy find this element by filter by #elementpath
                        new_merge_field.set('is_merge_field', 'True')

                        t_element.addprevious(new_merge_field)

                        # maybe there are some merge field with the same name -> add to list
                        if new_field_name not in self.field_name__elements:
                            self.field_name__elements[new_field_name] = []
                        self.field_name__elements[new_field_name].append(new_merge_field)

                    # remove current merge field due to contain multi merge field
                    parent = t_element.getparent()
                    parent.remove(t_element)
                else:
                    print("aaaaaaaaaaaaaaaaaaaaaaaa",t_element.text)
                    t_element.text = field_name_contain_open_close_tag
                    # set new attribute named is_merge_field -> easy find this element by filter by #elementpath
                    t_element.set('is_merge_field', 'True')

                    # maybe there are some merge field with the same name -> add to list
                    if field_name not in self.field_name__elements:
                        self.field_name__elements[field_name] = []
                    self.field_name__elements[field_name].append(t_element)

                # remove current merge field
                part_of_merge_fields.pop()

                # delete all element between open and close tag
                for element in part_of_merge_fields:
                    parent = element.getparent()
                    grand_parent = parent.getparent()
                    grand_parent.remove(parent)

                is_found_merge_field = False
                part_of_field_names = []
                part_of_merge_fields = []


    def __parse_tables(self, part):
        tables = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tbl')
        for table in tables:
            for row in table:
                merge_fields = row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                for merge_field in merge_fields:
                    self.in_table_field_name__details[merge_field.text[1:-1]] = {
                        'table': table,
                        'row': row
                    }

            # Bảng khung dập nổi
            text_element_in_tables = table.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')
            # chỉ chấp nhận trường hợp các ô trong bảng đều là ô trống hoặc là ký tự phân cách
            if (not text_element_in_tables) or \
                    (text_element_in_tables and all(
                        text_element.text in BYPASS_SEPARATOR_CHARACTERS for text_element in text_element_in_tables)):

                tc = table.getparent()
                merge_fields = tc.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                for merge_field in merge_fields:
                    field_name = merge_field.text[1:-1]
                    self.embossed_table_field_name__details[field_name] = {
                        'embossed_table': table,
                        'merge_field': merge_field
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
                'values': None
            }

        for checkbox_field_name, group_checkbox_details in self.checkbox_field_name__list_group_checkbox_details.items():
            field_name__infos[checkbox_field_name] = {
                'type': MERGE_FIELD_TYPE_CHECKBOX,
                'values': self.checkbox_field_name__values[checkbox_field_name]
            }

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

    def merge(self, replacements: dict):
        for field_name, replacement in replacements.items():
            if self.__is_valid_need_to_checked_values_for_checkbox(value_need_to_check=replacement) \
                    and field_name in self.checkbox_field_name__list_group_checkbox_details:
                self.__merge_checkbox(field_name=field_name, need_to_checked_values=replacement)

            elif self.__is_valid_need_to_checked_values_for_extend_text_checkbox(value_need_to_check=replacement) \
                    and field_name in self.checkbox_field_name__list_group_checkbox_details:
                self.__merge_checkbox(field_name=field_name, need_to_checked_values=replacement,
                                      is_add_extend_text=True)

            elif self.__is_valid_values_for_embossed_table(value_need_to_check=replacement) \
                    and field_name in self.embossed_table_field_name__details:
                self.__merge_embossed_table(field_name=field_name, embossed_table_info=replacement)

            elif self.__is_valid_values_for_row(value_need_to_check=replacement):
                self.__merge_rows(anchor=field_name, rows=replacement)

            else:
                self.__merge_field(field_name=field_name, text=replacement)

    def __merge_rows(self, anchor, rows):
        if anchor not in self.field_name__elements:
            return None

        table = self.in_table_field_name__details[anchor]['table']
        row = self.in_table_field_name__details[anchor]['row']

        if len(rows) > 0:
            for row_data in rows:
                new_row = deepcopy(row)
                row.addprevious(new_row)

                in_new_row_checkbox_field_name__list_group_checkbox_details = \
                    self.__get_checkbox_field_name__list_group_checkbox_details(part=new_row)

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
                                    need_to_checked_values=need_to_checked_values_or_text
                                )

                        else:
                            self.__fill_text(
                                merge_field=merge_field,
                                parent=merge_field.getparent(),
                                text=need_to_checked_values_or_text
                            )

            table.remove(row)
        else:
            if self.is_remove_empty_table:
                parent = table.getparent()
                parent.remove(table)

    def __merge_field(self, field_name, text):
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
                self.__fill_text(merge_field=merge_field, parent=parent, text=text)

    def __merge_embossed_table(self, field_name, embossed_table_info):

        text = str(embossed_table_info['value'])

        detail = self.embossed_table_field_name__details[field_name]
        table = detail['embossed_table']
        merge_field = detail['merge_field']
        cells = table.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tc')

        # Dữ liệu nhập dài hơn số ô đang có sẽ không được điền vào
        if len(text) > len(cells):
            return

        # Điền text vào ô
        # tc(cell) -> p -> r -> t(text)
        characters = [t for t in text]
        for index in range(len(characters)):
            if characters[index] in BYPASS_SEPARATOR_CHARACTERS:  # Bỏ qua nếu là ký tự phân cách
                continue
            p_node = cells[index].find(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}p')
            r_node = Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}r')
            t_node = Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}t')
            t_node.text = characters[index]

            r_node.append(t_node)
            p_node.append(r_node)

        p = merge_field.getparent()
        p.remove(merge_field)

    def __merge_checkbox(self, field_name, need_to_checked_values, is_add_extend_text=False):
        for group_checkbox_detail in self.checkbox_field_name__list_group_checkbox_details[field_name]:
            self.__fill_checkbox(
                field_name=field_name,
                merge_field=group_checkbox_detail['merge_field'],
                parent=group_checkbox_detail['merge_field'].getparent(),
                checkbox_infos=group_checkbox_detail['checkbox_infos'],
                need_to_checked_values=need_to_checked_values,
                is_add_extend_text=is_add_extend_text
            )

    def __fill_checkbox(self, field_name, merge_field, parent, checkbox_infos, need_to_checked_values,
                        is_add_extend_text=False):
        for checkbox_info in checkbox_infos:
            # uncheck all checkbox in group
            checkbox_info['checkbox_obj'].text = CHECKBOX_UNCHECKED_TEXT

        exclude_values = []  # exclude values for 'Khác' checkbox
        is_checked_checkbox = False

        # Trường hợp có extend text cho checkbox:
        if is_add_extend_text:
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
                if checkbox_info['value'].lower() == 'khác:':
                    checkbox_info['checkbox_obj'].text = CHECKBOX_CHECKED_TEXT
                    is_checked_checkbox = True

                    self.__set_text_for_t_element(
                        element=checkbox_info['value_obj'],
                        text=checkbox_info['original_value'].replace(
                            checkbox_info['value'],
                            f"{checkbox_info['value']} {', '.join(others_need_to_checked_values)}"
                        )
                    )
        ################################################################################################################
        # if found checkbox has value the same with need to checked values, then replace merge_field with empty text
        if is_checked_checkbox:
            self.__fill_text(
                merge_field=merge_field,
                parent=parent,
                text=''
            )
        # if NOT, then replace merge_field with need_to_checked_values to easily DEBUG
        else:
            self.__fill_text(
                merge_field=merge_field,
                parent=parent,
                text=', '.join(others_need_to_checked_values)
            )

    def __fill_text(self, merge_field, parent, text):
        text = text or ''  # text might be None

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

    @staticmethod
    def __set_text_for_t_element(element, text):
        element.text = text
        if text.startswith(' ') or text.endswith(' '):
            element.set(XML_SPACE_ATTRIBUTE, 'preserve')

    @staticmethod
    def __is_valid_need_to_checked_values_for_checkbox(value_need_to_check):
        if isinstance(value_need_to_check, list) \
                and (not value_need_to_check or (value_need_to_check and all(isinstance(value, str)
                                                                             for value in value_need_to_check))):
            return True
        return False

    @staticmethod
    def __is_valid_need_to_checked_values_for_extend_text_checkbox(value_need_to_check):
        if isinstance(value_need_to_check, dict) and \
                ('type' in value_need_to_check and value_need_to_check['type'] == MERGE_FIELD_TYPE_EXTEND_TEXT_CHECKBOX) and \
                ('value' in value_need_to_check and isinstance(value_need_to_check['value'], list)):

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
