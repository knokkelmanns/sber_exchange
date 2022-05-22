import requests
import openpyxl

headers = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,"
              "application/signed-exchange;v=b3;q=0.9",
    "user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.81 "
                  "Safari/537.36",
    "bx-ajax": "true"
}


# download a file with bank branches where you can buy money
def get_data():
    try:
        r = requests.get(
            url='ADD_DIRECT_LINK_TO_XLSX_FROM_BANK',
            headers=headers)
        with open("exchange.xlsx", "wb") as file:
            file.write(r.content)

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

                # for x, y in ('ё', '-'), ('е', ' '):
                # if f'г.{city}'.lower().replace(x, y) in address_list.lower().replace(x, y):
                # print(address_list)

    except Exception as _ex:
        return 'Упс... что-то пошло не так. Попробуй снова'


def main():
    city = input('Введите название города на русском: ')
    print(get_data())
    print(parse_data(city))


if __name__ == '__main__':
    main()
