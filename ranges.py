from src.excel_file import ExcelFile
import src.utils as utl
import tqdm


CONF = utl.get_yaml('config.yml')
test_file = ExcelFile()
test_file.add_workbook(CONF['input_path'])
test_file.select_sheet(CONF['sheet_name'])


column_series = test_file.column_to_series("C")
participating = column_series[column_series == 'Yes']
rows_to_paint = participating.index

print(f'[ranges.py]: Number of rows to paint: {len(rows_to_paint)}')
print(f'[ranges.py]: Painting ...')
for row_number in tqdm.tqdm(rows_to_paint):
    address = f'A{row_number + 2}:H{row_number + 2}'
    test_file.paint_range(address, rgb=(0, 255, 0))



test_file.save_close_wb(CONF['output_path'])
test_file.stop_app()
