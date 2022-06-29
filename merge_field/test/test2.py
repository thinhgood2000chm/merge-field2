# import requests
#
# if __name__ =="__main__":
#
#     url = "http://192.168.73.138:9000/service-file/2/23-06-2022/8d9d1319efcb4da7ac73e2ada349e09e?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=service-file%2F20220623%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20220623T084331Z&X-Amz-Expires=3600&X-Amz-SignedHeaders=host&X-Amz-Signature=e2ca868294af1217d6abe115fbba4bcb174e34c286afc3a0b0a86116a50f8ac0"
#
#     payload = {}
#     headers = {}
#
#     response = requests.request("GET", url, headers=headers, data=payload, stream=True)
#     print("###############################################################################################################")
#     for i in response:
#         print(i)
#     print(response.iter_content)
comma = 'phẩy'

LEVEL_N = {
    0: "",
    3: 'nghìn',
    6: 'triệu',
    9: 'tỷ'
}
LESS_THAN_100 = {
    0: 'không',
    1: 'một',
    2: 'hai',
    3: 'ba',
    4: 'bốn',
    5: 'năm',
    6: 'sáu',
    7: 'bảy',
    8: 'tám',
    9: 'chín',
    10: 'mười'
}

INTEGER_PART = 0
DECIMAL_PART = 1

for number in range(11, 20):
    if number == 15:
        LESS_THAN_100[number] = 'mười lăm'
        continue
    LESS_THAN_100[number] = f'mười {LESS_THAN_100[number % 10]}'

for number in range(20, 100):
    if number % 10 == 0:
        LESS_THAN_100[number] = f'{LESS_THAN_100[number // 10]} mươi'
    elif number % 10 == 1:
        LESS_THAN_100[number] = f'{LESS_THAN_100[number // 10]} mươi mốt'
    elif number % 10 == 5:
        LESS_THAN_100[number] = f'{LESS_THAN_100[number // 10]} mươi lăm'
    else:
        LESS_THAN_100[number] = f'{LESS_THAN_100[number // 10]} mươi {LESS_THAN_100[number % 10]}'


def convert_e_number_to_number(input_number:float) -> str:
    # 0.0000058 = 5.8e-06
    num_e_part_and_zero_part = str(input_number).split('-')
    str_input_number_type_float = f'0.{"".join(["0"] * (int(num_e_part_and_zero_part[1])-1))}' \
                                  f'{num_e_part_and_zero_part[0].replace(".","").replace("e","")}'
    return str_input_number_type_float


def number_to_words(input_number: float) -> str:
    if input_number == 0:
        return LESS_THAN_100[0]

    is_positive = True
    if input_number < 0:
        is_positive = False
        input_number = -input_number

    str_input_number = str(input_number)
    if 'e' in str_input_number:
        # nếu số bắt đầu bằng số 0 và có nhiều số 0 sau dấu phẩy sẽ bị máy chuyển thành số e
        # ==> convert số e lại số ban đầu nhưng dưới dạng string
        # ko sử dụng format(float...) vì sẽ làm sai lệch số
        str_input_number = convert_e_number_to_number(input_number)

    integer_and_decimal_parts = str_input_number.split('.')

    output_total = []
    for index, part in enumerate(integer_and_decimal_parts):
        if index == 1:  # có phần thập phân
            output_total.append(comma)

        digits = list(part)
        digits.reverse()

        digits = [int(digit) for digit in digits]
        level = 0
        output = []
        for i in range(0, len(digits), 3):
            unit_digit = digits[i]  # chữ số hàng đơn vị
            ten_digit = digits[i + 1] if i + 1 < len(digits) else None  # chữ số hàng chục
            hundred_digit = digits[i + 2] if i + 2 < len(digits) else None  # chữ số hàng trăm

            if level > 9:
                level = 3

            if level >= 3:
                if level == 9:
                    output.append(LEVEL_N[level])  # Trường hợp tỷ tỷ thì cần phải thêm tỷ vào, không bỏ qua khi 000
                else:
                    if index == INTEGER_PART and ((unit_digit != 0)
                            or (ten_digit is not None and ten_digit != 0)
                            or (hundred_digit is not None and hundred_digit != 0)):
                        # Trường hợp phần số nguyên không phải 000 thì mới thêm level nghìn, triệu, tỷ vào
                        output.append(LEVEL_N[level])
                    elif index == DECIMAL_PART:
                        output.append(LEVEL_N[level])

            # Trường hợp chỉ có chữ số hàng đơn vị
            if unit_digit is not None and ten_digit is None and hundred_digit is None:
                if index == INTEGER_PART and unit_digit != 0:
                    output.append(f'{LESS_THAN_100[unit_digit]}')
                else:
                    output.append(f'{LESS_THAN_100[unit_digit]}')
            # Trường hợp chỉ có chữ số hàng đơn vị và hàng chục
            elif ten_digit is not None and hundred_digit is None:
                # đổi với các số phía sau dấu phẩy
                if ten_digit == 0 and index == DECIMAL_PART:
                    if level >= 3:
                        output.append(
                            f'{LESS_THAN_100[ten_digit]} chục {LESS_THAN_100[unit_digit]}' if unit_digit != 0 else f'{LESS_THAN_100[ten_digit]} chục')
                    else:
                        output.append(f'{LESS_THAN_100[ten_digit]} {LESS_THAN_100[unit_digit]}' if unit_digit != 0 else f'{LESS_THAN_100[ten_digit]}')
                else:
                    output.append(f'{LESS_THAN_100[ten_digit * 10 + unit_digit]}')

                # Trường hợp có cả ba chữ số hàng đơn vị, hàng chục và hàng trăm
            else:
                # NOTE: vẽ bảng chân trị sẽ thấy đã xử lý tất cả trường hợp rõ ràng hơn

                # Nếu cả 3 chữ số đều bằng 0 thì bỏ qua nếu ở phần số nguyên
                if unit_digit == 0 and ten_digit == 0 and hundred_digit == 0:
                    if index == DECIMAL_PART:
                        output.append(f'{LESS_THAN_100[hundred_digit]} trăm')
                    else:
                        pass

                # Nếu chữ số hàng đơn vị và hàng chục bằng 0, chữ số hàng trăm khác 0
                # => x trăm
                elif unit_digit == 0 and ten_digit == 0 and (hundred_digit != 0):
                    print(" vao 2 ")
                    output.append(f'{LESS_THAN_100[hundred_digit]} trăm')

                # Nếu chữ số hàng chục bằng 0, chữ số hàng đơn vị khác 0
                # => x trăm lẻ y
                elif unit_digit != 0 and ten_digit == 0:
                    print(" vao 3 ")
                    output.append(f'{LESS_THAN_100[hundred_digit]} trăm lẻ {LESS_THAN_100[unit_digit]}')

                # Nếu chữ số hàng chục khác 0
                else:
                    output.append(f'{LESS_THAN_100[hundred_digit]} trăm {LESS_THAN_100[ten_digit * 10 + unit_digit]}')

                # Sau 3 chữ số thì tăng level lên để có thể thêm nghìn, triệu tỷ
            level += 3

        if not is_positive:
            output.append('âm')
        output.reverse()
        output_total.extend(output)

    output_total[0] = output_total[0].capitalize()  # viết hoa chữ cái đầu
    return ' '.join(output_total)


if __name__ == '__main__':
    print(number_to_words(0.000005))