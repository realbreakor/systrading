import asyncio
import math

from ls_api.quote import xingCurrentPrice


def find_closest_base_number(number):
    candidates = [0, 2, 5, 7]
    ones_place = int(round(number)) % 10
    closest = min(candidates, key=lambda x: abs(ones_place - x))
    base_number = round(number) - ones_place + closest
    return f"{base_number:03d}"


def format_number(num):
    return f"{num:03d}"


def generate_numbers(cp, base_number):
    base_number = int(base_number)
    numbers = []
    if cp == '2':
        for i in range(-40, 30):
            candidate = base_number + i
            if candidate % 10 in [0, 2, 5, 7]:
                numbers.append(format_number(candidate))
    elif cp == '3':
        for i in range(-40, 30):
            candidate = base_number - i
            if candidate % 10 in [0, 2, 5, 7]:
                numbers.append(format_number(candidate))
    numbers = sorted(set(numbers))
    return numbers if cp == '2' else numbers[::-1]


async def find_opt_order(cp, lst):
    print(lst)
    _lst = []
    for opt_price in lst:
        if len(_lst) == 2:
            return _lst
        opt_code = cp + opt_price
        yeprice = float(await xingCurrentPrice(opt_code, 'yeprice'))
        print(yeprice)
        if 4 <= yeprice < 6:
            _lst.append((opt_code, 6))
        elif 2 <= yeprice < 4:
            _lst.append((opt_code, 12))
        elif 1 <= yeprice < 2:
            _lst.append((opt_code, 20))
        else:
            await asyncio.sleep(1.2)


def get_opt_num(price, jango, cnt_code, cnt):
    if 1 <= price < 7:
        multi = int(math.floor(int(jango) / 5000000 / cnt_code)) * 2
        num = ((15000000 - price * 250000) / (price * 0.31 * 250000)) / 4 / 4 * multi * (cnt / 4)
        return round(num)


def get_opt_num_excess(price, jango, cnt_code, cnt):
    multi = int(math.floor(int(jango) / 5000000 / cnt_code)) * 2
    num = ((15000000 - price * 250000) / (price * 0.31 * 250000)) / 4 / 4 * multi * (cnt / 4)
    return round(num)


def sort_dict_by_list(data_dict, oList):
    return {key: data_dict[key] for key in oList if key in data_dict}
