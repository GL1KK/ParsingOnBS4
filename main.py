import requests
from bs4 import BeautifulSoup
import time
from pprint import pprint
import openpyxl
from openpyxl.styles import Font

def get_info():
    all_data = [] 
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:138.0) Gecko/20100101 Firefox/138.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        'DNT': '1',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Sec-GPC': '1',
        'Priority': 'u=0, i',
    }

    params = {
        'per-page': '50',
    }

    for i in range(801, 7651):
        try:
            cookies = {
                'PHPSESSID': 'mg0be83s1ot3bmjipct14297qm',
                '_language': '14157e6cc3d06fbac29fe2636f897f417d1e0472027dc8e3f3abcbfc366257e2a%3A2%3A%7Bi%3A0%3Bs%3A9%3A%22_language%22%3Bi%3A1%3Bs%3A2%3A%22ru%22%3B%7D',
                '_csrf': '8ce9526630e730476a7067d58e1fc6d85423514953f5b20b7fe81960c1987279a%3A2%3A%7Bi%3A0%3Bs%3A5%3A%22_csrf%22%3Bi%3A1%3Bs%3A32%3A%22Go_0hZCGm963f7YaMdBwX-7hD8mUvsfN%22%3B%7D',
                'cf_clearance': 'bdUwbEJa7JaNdfzv.vBnWqiV05ojD37ROfdLqVHmsoA-1747992026-1.2.1.1-C928h0s3.EOJUqwgkA8gjFWqK3Lmq5ZmgzyQi5SmireCBNDYN7VnMhK_yeM4zYIImnqZuEDozb7domt9dx8il9hsUFYF2MYsxscCpj6jV.W_ukjqi.vxeucRtvbsgViNbCqfieKqOrPXh7r_I0GK1nPMj7n17Gtl5wC2FR3HigQefzNE2hg24_av8B4weIz3RObyU1P2xcuqrsTq.JBjnZh8ZR0EZEhsxuBUAlSYMD2lRZ6uTiriMu9S7.I7Iq8skHJjxeQ4cv7Wjh4Yt2.8u.MZ8l1bvz6C6GY4RWDpqQf5j8i._pFFVEsvUpkm9Srn4YUykPO4K0dKMfdSh9VeXpnmS2bXhahyFjPTFWlhDjg',
            }

            params['page'] = i
            
            print(f"Запрос страницы {i}")
            
            if i > 1:
                time.sleep(2)  
            
            response = requests.get(
                'https://pricing.parts/ru/spares',
                params=params,
                cookies=cookies,
                headers=headers,
                timeout=30 
            )
            
            if response.status_code != 200:
                print(f"Ошибка: статус код {response.status_code}")
                break
            
            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.find("table", class_="table table-striped table-condensed table-hover grid-parts")
            
            if not table:
                print("Таблица не найдена на странице")
                break
                
            rows = table.find_all('tr')
            
            for row in rows:
                cols = row.find_all("td")
                if cols:
                    brand = cols[0].text.strip()
                    number = cols[1].text.strip()
                    opisanie = cols[2].text.strip()
                    price = cols[3].text.replace('\xa0', ' ').strip()
                    
                    all_data.append({
                        "Марка": brand,
                        "Номер": number,
                        "Описание": opisanie,
                        "Цена": price
                    })
            
            if i % 100 == 0:
                save_to_excel(all_data, f"запчасти_temp_{i}.xlsx")
                print(f"Временное сохранение после {i} страниц")
                
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при запросе страницы {i}: {e}")
            print("Повторная попытка через 10 секунд...")
            time.sleep(10)
            continue
        except Exception as e:
            print(f"Неожиданная ошибка на странице {i}: {e}")
            break
    
    save_to_excel(all_data, "запчасти_final.xlsx")
    print("Парсинг завершен!")

def save_to_excel(data, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Запчасти"

    headers = ["Марка", "Номер", "Описание", "Цена"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for item in data:
        ws.append([
            item["Марка"],
            item["Номер"],
            item["Описание"],
            item["Цена"]
        ])

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(filename)
    print(f"Данные сохранены в файл '{filename}'")

if __name__ == "__main__":
    get_info()