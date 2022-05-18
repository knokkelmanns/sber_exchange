import urllib.request
import openpyxl


# download a file with bank branches where you can buy money
def get_data():
    try:
        file = "ADD_DIRECT_LINK_TO_XLSX_FROM_BANK"
        urllib.request.urlretrieve(file, "exchange.xlsx")

        return 'XLSX успешно скачан!'

    except Exception as _ex:
        return 'Упс... Перепроверь URL, пожалуйста!'


# parsing and displaying results
def parse_data(city):
    try:
        book = openpyxl.open("exchange.xlsx", read_only=True)
        sheet = book.active

        # cities = sheet['D3:D999']
        address = sheet['F3:F999']

        for row in address:
            for cell in row:
                address_list = cell.value
                if f'г.{city}'.lower().replace('ё', 'е') in address_list.lower().replace('ё', 'е'):
                    print(address_list)

    except Exception as _ex:
        return 'Упс... что-то пошло не так. Попробуй снова'


def main():
    city = input('Введите название города на русском: ')
    # print(get_data())
    print(parse_data(city))


if __name__ == '__main__':
    main()
