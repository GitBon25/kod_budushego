import csv
import requests
import openpyxl


api = "https://jsonplaceholder.org/"
resources = {1: "posts",
             2: "users",
             3: "comments",
             4: "todos"
             }


def flatten(item, prefix=""):
    flat = {}
    for key, value in item.items():
        name = prefix + key
        if isinstance(value, dict):
            flat.update(flatten(value, name + "."))
        elif isinstance(value, list):
            flat[name] = ", ".join(str(v) for v in value)
        elif value is None:
            flat[name] = ""
        else:
            flat[name] = value
    return flat


def to_table(data):
    rows = [flatten(item) for item in data if isinstance(item, dict)]
    headers = []
    for row in rows:
        for key in row:
            if key not in headers:
                headers.append(key)
    table = [[row.get(h, "") for h in headers] for row in rows]
    return headers, table


def save_csv(headers, rows, file):
    with open(file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(headers)
        writer.writerows(rows)


def save_xlsx(headers, rows, file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    for col in ws.columns:
        width = max(len(str(c.value))
                    if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
    wb.save(file)


def get_info(n, count, format, file):
    try:
        response = requests.get(api + resources[n], timeout=10)
        response.raise_for_status()
        response = response.json()
    except requests.exceptions.Timeout:
        print("Ошибка: сервер не ответил вовремя.")
        return
    except requests.exceptions.ConnectionError:
        print("Ошибка: нет соединения. Проверьте интернет.")
        return
    except requests.exceptions.HTTPError as e:
        print(f"Ошибка HTTP: {e}")
        return
    except requests.exceptions.RequestException as e:
        print(f"Ошибка запроса: {e}")
        return
    except ValueError:
        print("Ошибка: сервер вернул некорректный JSON.")
        return

    if isinstance(response, dict):
        response = [response]
    if not isinstance(response, list) or not response:
        print("Ошибка: API вернул пустой или неожиданный ответ.")
        return

    if count:
        data = response[:count]
    else:
        data = response

    headers, rows = to_table(data)
    try:
        if format == "csv":
            save_csv(headers, rows, file)
        else:
            save_xlsx(headers, rows, file)
    except PermissionError:
        print("Ошибка: нет доступа к файлу.")
        return
    except OSError as e:
        print(f"Ошибка записи файла: {e}")
        return
    print(f"Данные сохранены в {file} ({len(rows)} записей)")


def main():
    print("Какой ресурс загрузить? \n 1. posts \n 2. users \n 3. comments \n 4. todos")
    while True:
        try:
            n = int(input("Введите номер ресурса: "))
            if n in resources:
                break
        except ValueError:
            pass
        print("Ошибка: введите число от 1 до 4.")

    while True:
        count = input(
            "Сколько записей сохранить? Оставьте пустым, чтобы сохранить все:").strip()
        if not count:
            count = None
            break
        if count.isdigit() and int(count) > 0:
            count = int(count)
            break
        print("Ошибка: введите положительное целое число или оставьте пустым.")

    while True:
        format = input("Выберите формат файла (csv/xlsx): ").strip().lower()
        if format in ("csv", "xlsx"):
            break
        print("Ошибка: допустимы только csv или xlsx.")

    while True:
        file = input("Введите имя файла: ").strip()
        if file and not any(c in file for c in '<>:"/\\|?*'):
            break
        print("Ошибка: пустое имя или недопустимые символы.")
    if not file.lower().endswith("." + format):
        file += "." + format

    get_info(n, count, format, file)


if __name__ == "__main__":
    main()
