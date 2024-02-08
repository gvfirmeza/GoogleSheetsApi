import math
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def calculate_and_update_results(spreadsheet_id, sheet_name, start_row, end_row):
    credentials = ServiceAccountCredentials.from_json_keyfile_name('.\\api\\diesel-cat-413500-193f55b18f89.json', ['https://spreadsheets.google.com/feeds'])
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key(spreadsheet_id)
    
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{sheet_name}' not found.")
        return
    
    data = worksheet.get_all_values()[start_row - 1:end_row]

    for i, row in enumerate(data, start=start_row):
        attendance = float(row[2])
        values = [float(cell) for cell in row[3:6]]
        average = sum(values) / 3
        
        total_classes = 60
        if attendance > 0.25 * total_classes:
            result = "Reprovado por Falta"
            worksheet.update_cell(i, 8, 0)
        elif average < 50:
            result = "Reprovado por Nota"
            worksheet.update_cell(i, 8, 0)
        elif 50 <= average < 70:
            result = "Exame Final"
            # Calculate NAF
            naf = max(math.ceil(100 - average), 0) 
            worksheet.update_cell(i, 8, naf)
        else:
            result = "Aprovado"
            worksheet.update_cell(i, 8, 0)
        
        worksheet.update_cell(i, 7, result)

if __name__ == "__main__":
    spreadsheet_id = '1Kjre36osftjSRzaNFNxGWeJ8TtZnuvhf3ydli1fhvds'
    sheet_name = 'engenharia_de_software'
    start_row = 4
    end_row = 27

    calculate_and_update_results(spreadsheet_id, sheet_name, start_row, end_row)
