# docx introduction: https://www.toptal.com/xml/an-informal-introduction-to-docx

# lxml documentation: https://lxml.de/tutorial.html
# https://lxml.de/api/lxml.etree._ElementTree-class.html
# https://lxml.de/api/lxml.etree._Element-class.html
# https://lxml.de/tutorial.html#elementpath

from copy import deepcopy
from lxml.etree import Element
from lxml import etree
from zipfile import ZipFile, ZIP_DEFLATED

NAMESPACE_WORDPROCESSINGML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'  # noqa
NAMESPACE_CONTENT_TYPE = 'http://schemas.openxmlformats.org/package/2006/content-types'  # noqa

ELEMENT_PATH_RECURSIVE_MERGE_FIELD = f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t[@is_merge_field="True"]'

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',  # noqa
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',  # noqa
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',  # noqa
)

OPEN_TAG = '«'
CLOSE_TAG = '»'

# .// : find element in anywhere in the tree
class MergeField:
    def __init__(self, file, is_remove_empty_table=False):
        self.zip: ZipFile = ZipFile(file)
        self.parts: dict = {} # parts này dùng để lưu lại các phần như main, header, fooder bao gồm toàn bộ xml của từng phần
        self.merge_field_name__elements: dict = {}
        self.merge_field_name_in_table__details: dict = {}

        self.is_remove_empty_table: bool = is_remove_empty_table

        try:
            content_types = etree.parse(self.zip.open('[Content_Types].xml'))
            print(etree.tostring(content_types))
            for file in content_types.findall(f'{{{NAMESPACE_CONTENT_TYPE}}}Override'):
                print(etree.tostring(file))
                print("122222222222222222",file.attrib)
                if file.attrib['ContentType'] in CONTENT_TYPES_PARTS:
                    filename = file.attrib['PartName'].split('/', 1)[1]  # remove first / ==> word/document.xml
                    self.parts[filename] = etree.parse(self.zip.open(self.zip.getinfo(filename)))

            for part in self.parts.values():
                print("part in parts", etree.tostring(part))
                self.__parse_merge_fields(part=part)  # init data for merge_field_name__elements
                self.__parse_tables(part=part)  # init data for merge_field_name_in_table__details

        except Exception as ex:
            self.zip.close()
            raise ex

    def __parse_merge_fields(self, part):
        is_found_merge_field = False
        part_of_name_merge_fields = [] # luưu lại w:t << asdfasdf >>
        part_of_merge_fields = [] # lưu những field mà << , "nội dung" nằm ở vị trí khác nhau nhưng cùng trong 1 w:p

        merge_fields = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}t')# tìm tất cả các w:t có trong xml
        # tìm những vùng có << asdfasdf >> , những vùng nào ko có thì bỏ qua

        for merge_field in merge_fields:
            print("text + ",merge_field.text)
            # .text thì nó sẽ ra được dấu open tag và close tag đang bị mã hóa
            if OPEN_TAG not in merge_field.text and not is_found_merge_field:
                continue

            # can not found close tag in previous merge field
            if OPEN_TAG in merge_field.text and is_found_merge_field: # nếu như gặp 1 open tag khác ==> tạo mới list để tiến hành tìm và add các phần tử của merge_field vào listt này
                part_of_name_merge_fields = []
                part_of_merge_fields = []

            if OPEN_TAG in merge_field.text:
                is_found_merge_field = True

            part_of_name_merge_fields.append(merge_field.text)

            if CLOSE_TAG not in merge_field.text: #  phần này giải quyết vấn đề dấu << ở 1 nơi nội dung 1 nơi và >> ở nơi khác
                print("#############3", merge_field.text)
                # add element between open and close tag to list need to delete
                part_of_merge_fields.append(merge_field)
                for i in range(len(part_of_merge_fields)):
                    print("#############3", part_of_merge_fields[i].text)
            else:
                # there are some text after » -> add new element contains text after »
                print("^^^^^^^^^^^^^^^^^^^^",part_of_name_merge_fields[-1])
                if not part_of_name_merge_fields[-1].endswith(CLOSE_TAG):# kiemr tra xem phần tử add cuối cùng có kết thúc bằng >> hay không
                    last_part_of_name_merge_field, remainder = part_of_name_merge_fields[-1].split(CLOSE_TAG)# nếu không thì tiến hành cắt để lấy ra phần merge_field thôi
                    print("777777777777777777777", last_part_of_name_merge_field, remainder)
                    part_of_name_merge_fields[-1] = f'{last_part_of_name_merge_field}{CLOSE_TAG}' # thêm [-1] để đảm bảo luôn thêm vào vị trí cuối cùng trong list
                    print("333333333333333333",  part_of_name_merge_fields)

                    remainder_element = deepcopy(merge_field)
                    remainder_element.text = remainder
                    parent = merge_field.getparent()
                    parent.append(remainder_element)

                merge_field_name = ''.join(part_of_name_merge_fields)
                print("??????????????", merge_field_name)
                merge_field.text = merge_field_name
                # kết thúc tìm << >>

                # set new attribute named is_merge_field -> easy find this element by filter by #elementpath
                merge_field.set('is_merge_field', 'True') # chỗ này để khi xuống dưới có thể tìm ra những cái cần fill bằng cachs dùng ELEMENT_PATH_RECURSIVE_MERGE_FIELD

                merge_field_name = merge_field_name[1:-1] # lấy ra giá trị bên trong << >>

                # maybe there are some merge field with the same name -> add to list
                if merge_field_name not in self.merge_field_name__elements:
                    self.merge_field_name__elements[merge_field_name] = []
                self.merge_field_name__elements[merge_field_name].append(merge_field) # vd với fieldname = THONG_TIN_HOI_SO_SCB thì giá trị trong list sẽ lfa 1 chuỗi các xml của filed đó

                # delete all element between open and close tag
                # chỗ này xóa thẻ w:r của << và " nọi dung" để khi thay thế nội dung chỉ cần thay thế vào vị trí của thẻ đóng >> là được khoogn
                # cần phải tìm vị trí của "nội dung" để thay
                for element in part_of_merge_fields:
                    print("element ", element.text)
                    parent = element.getparent()# lấy ra các thẻ cha ( thẻ bao bên ngoài )
                    grand_parent = parent.getparent()
                    grand_parent.remove(parent)

                is_found_merge_field = False
                part_of_name_merge_fields = []
                part_of_merge_fields = []

    def __parse_tables(self, part): ## phần này tìm và lưu lại các field của table
        tables = part.findall(f'.//{{{NAMESPACE_WORDPROCESSINGML}}}tbl') # tìm ra các vị trị của table
        for table in tables:
            for row in table: # lấy ra các thẻ w:tr
                merge_fields = row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD) # lấy ra các w:t đã được đánh dấu trong row của table  đó
                for merge_field in merge_fields:
                    print("merge_filed in table : ", merge_field.text[1:-1])
                    print("row", row.tag)
                    self.merge_field_name_in_table__details[merge_field.text[1:-1]] = {
                        'table': table, # table trên này là chỉ add cái khung table (trong này bao gồm các row vs các merge_filed để xuống dưới có thể đưa các giá trị đã fill data vào
                        'row': row
                    }

    @property
    def merge_fields(self):
        return list(self.merge_field_name__elements.keys())

    def write(self, file):
        # clear attribute is_merge_field
        for part in self.parts.values():
            merge_fields = part.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
            for merge_field in merge_fields:
                text = merge_field.text
                # ở đây xóa đi is_merge_field = true đã thiết lập ở ben trên
                # xóa nội dung cũ có is_merge_field sau đó cặng nhật lại filed đó vs text ko có is_merge_field
                print("merge_field.text", merge_field.text)
                #print("ffff", etree.tostring(merge_field))
                merge_field.clear()
                print(etree.tostring(merge_field))
                merge_field.text = text
                print("merge_field.text ....", merge_field.text)
                print(etree.tostring(merge_field))

        with ZipFile(file, 'w', ZIP_DEFLATED) as output:
            for zip_info in self.zip.filelist:
                filename = zip_info.filename # filename : hiển thị tất cả các word/document.xml, word/styles.xml.... giống vs filename ở phần init
                #print(filename)
                if filename in self.parts:
                    #print("fill" , etree.tostring(self.parts[filename].getroot()))
                    xml = etree.tostring(self.parts[filename].getroot())
                    output.writestr(filename, xml)
                else:
                    output.writestr(filename, self.zip.read(zip_info))

    def merge(self, replacements: dict): # dùng để thay thế giá trị vào các field
        for field_name, replacement in replacements.items():
            if isinstance(replacement, list):
                self.__merge_rows(anchor=field_name, rows=replacement)
            else:
                self.__merge_field(field_name=field_name, text=replacement)

    def __merge_rows(self, anchor, rows):
        if anchor not in self.merge_field_name__elements:# kiểm tra xem field này có nằm trong các field bị trùng hay ko ( ko nằm trong table)
            return None

        table = self.merge_field_name_in_table__details[anchor]['table']
        row = self.merge_field_name_in_table__details[anchor]['row']

        if len(rows) > 0:
            for row_data in rows:
                new_row = deepcopy(row) # tạo 1 row mới để tránh việc khi sửa 1 row này thì tất cả các row đều thay đổi 
                table.append(new_row) # đưa dữ liệu row vào tables
                # print("table in __merge_rows", etree.tostring(table) )

                merge_fields = new_row.findall(ELEMENT_PATH_RECURSIVE_MERGE_FIELD)
                for merge_field in merge_fields:
                    text = row_data.get(merge_field.text[1:-1]) # lấy ra gía trị nguoừi dùng gửi lên với key là merge_field.text
                    print("text in __merge_rows ", text)
                    if text:
                        self.__fill_merge_field(merge_field=merge_field, parent=merge_field.getparent(), text=text)
            # sau khi đã có row mới và có data của row đó thì xóa đi row cũ với dòng có << >>
            table.remove(row)
        else:
            if self.is_remove_empty_table:
                parent = table.getparent()
                parent.remove(table)

    def __merge_field(self, field_name, text):
        for merge_field in self.merge_field_name__elements.get(field_name, []):
            parent = merge_field.getparent()

            # remove blank line
            if not text \
                    and merge_field.text[1:-1] not in self.merge_field_name_in_table__details \
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
                self.__fill_merge_field(merge_field=merge_field, parent=parent, text=text)

    @staticmethod
    def __fill_merge_field(merge_field, parent, text):
        text = text or ''  # text might be None

        # preserve new lines in replacement text
        text_parts = str(text).split('\n')# chỗ này dùng để tách ra những cái có cùng mergefile nhưng nhiều dòng
        print("text_parts", text_parts)
        nodes = []

        for text_part in text_parts:
            text_node = Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}t')# chỉ lấy ra những chỗ với thẻ w:t

            text_node.text = text_part # add data vào thẻ đó
            print("text_node", etree.tostring(text_node))
            nodes.append(text_node)

            nodes.append(Element(f'{{{NAMESPACE_WORDPROCESSINGML}}}br'))

        nodes.pop()  # remove last br element

        for node in reversed(nodes): # addnext chỗ này sẽ add theo kiểu first in last out do đó nên đổi thứ tự ngược lại trước rối mới add next thì nó sẽ tự động đúng thứ tự
            merge_field.addnext(node)# sau khi các node con ( các field cần điền data đều có giá trị thì add vào merge_field ( những vị trí đã đành dấu bằng is_merge_field = true theo thứ tự )

        parent.remove(merge_field)  # remove old merge field element due to it is replaced by new text in nodes

    def __enter__(self):
        return self

    def __exit__(self, exc_type, value, traceback):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None
