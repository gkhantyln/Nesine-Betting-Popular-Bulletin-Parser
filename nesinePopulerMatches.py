import requests
import xlsxwriter
import random
import string
from datetime import datetime, timedelta

# Kullanıcıdan seçeneği al
choice = input("Bir seçenek seçin (0: Hepsi, 1: Bugün, 2: Yarın, 3: Yarından sonraki gün) : ")

try:
    choice = int(choice)
except ValueError:
    print("Geçersiz bir seçenek girdiniz.")
    exit()

# Tarih seçeneğini oluştur
if choice == 0:
    selected_date = None
elif choice == 1:
    selected_date = datetime.now().strftime('%Y-%m-%d')   
elif choice == 2:
    tomorrow = datetime.now() + timedelta(days=1)
    selected_date = tomorrow.strftime('%Y-%m-%d')
elif choice == 3:
    day_after_tomorrow = datetime.now() + timedelta(days=2)
    selected_date = day_after_tomorrow.strftime('%Y-%m-%d')
else:
    print("Geçersiz bir seçenek girdiniz.")
    exit()

# Önce verileri çekelim
def get_matches(date):
    url = 'https://pb.nesine.com/v1/Bet?eventType=1'
    if date:
        params = {'FilterDates': date}
    else:
        params = {}  # Tarih belirtilmemişse boş bir parametre sözlüğü kullan

    response = requests.get(url, params=params)

    if response.status_code == 200:
        data = response.json()
        all_matches = data['d']['PopularBetList']
        
        if date:
            # Sadece tarih belirtilmişse filtrele
            filtered_matches = [match for match in all_matches if match['MatchTime'].startswith(date)]
        else:
            filtered_matches = all_matches  # Tüm maçları al
            
        return filtered_matches
    else:
        return None

# Şimdi Excel dosyasına yazma işlemini yapalım
def write_matches_to_excel(matches):
    if not matches:
        print("No matches to write to Excel.")
        return

    # Rastgele bir dize oluştur
    random_suffix = ''.join(random.choices(string.ascii_letters + string.digits, k=6))

    # Dosya adına rastgele değeri ekle
    file_name = f'nesine_matches_{random_suffix}.xlsx'

    # Excel dosyasını oluştur
    wb = xlsxwriter.Workbook(file_name)
    
    ws = wb.add_worksheet('Matches')

    # Başlıkları ve verileri yazdırmak için stilleri tanımla
    bold = wb.add_format({'bold': True})
    centered = wb.add_format({'align': 'center', 'valign': 'vcenter'})

    # Sütun başlıkları
    header = ['Match Code', 'Date', 'Time', 'Match Name', 'Market Name', 'Outcome Name', 'Odd', 'Played Count', 'Statistics URL']

    # Sütun başlıklarını yazdır ve ortalama ayarla
    for col, item in enumerate(header):
        ws.write(0, col, item, bold)
        ws.set_column(col, col, len(item) + 2, centered)  # Kolonları ortalama ayarla

    # Tarih formatını tanımla
    date_format = wb.add_format({'num_format': 'dd-mm-yyyy'})
    date_format.set_align('center')  # Tarih sütunu için ortalama ayarı

    # Saat formatını tanımla
    time_format = wb.add_format({'num_format': 'hh:mm:ss'})
    time_format.set_align('center')  # Saat sütunu için ortalama ayarı

    # Verileri yazdır
    for row, match in enumerate(matches, start=1):
        ws.write(row, 0, match['Code'])
        
        match_time = datetime.strptime(match['MatchTime'], '%Y-%m-%dT%H:%M:%S')

        # Tarih ve saat sütunlarını belirli sütunlara yazdırın
        ws.write(row, 1, match_time.strftime('%d-%m-%Y'), date_format)  # Tarih
        ws.write(row, 2, match_time.strftime('%H:%M:%S'), time_format)  # Saat
        
        ws.write(row, 3, match['Name'])
        ws.write(row, 4, match['MarketName'])
        ws.write(row, 5, match['OutcomeName'])
        ws.write(row, 6, match['Odd'])
        ws.write(row, 7, match['PlayedCount'])
        ws.write(row, 8, match['StatisticsUrl'])

    # Sütun genişliklerini ayarlayın
    ws.set_column('A:A', 12)  # 1. sütun
    ws.set_column('B:B', 12)  # 2. ve 3. sütun
    ws.set_column('C:C', 10)  # 2. ve 3. sütun
    ws.set_column('D:D', 32)  # 4. sütun
    ws.set_column('E:E', 20)  # 5. sütun
    ws.set_column('F:F', 15)  # 6. sütun
    ws.set_column('G:G', 10)  # 7. sütun
    ws.set_column('H:H', 15)  # 8. sütun
    ws.set_column('I:I', 36)  # 9. sütun

    wb.close()
    print(f"Matches have been written to {file_name}")

if __name__ == "__main__":
    matches = get_matches(selected_date)
    if matches:
        write_matches_to_excel(matches)
