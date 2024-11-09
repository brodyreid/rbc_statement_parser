from os import path, makedirs, listdir
from csv import DictWriter
from pdfquery import PDFQuery
from pdfquery.cache import FileCache

COLUMN_NAMES = ['Date', 'Description', 'Withdrawals ($)', 'Deposits ($)', 'Balance ($)']
COLUMN_BUFFER = float(1.0)

def load_pdf(pdf_path, use_cache):
  if not path.exists(pdf_path):
    raise FileNotFoundError(f"The specified PDF file does not exist: {pdf_path}")
  
  pdf = PDFQuery(pdf_path)

  try:
    pdf.load()
    return pdf
  except Exception as e:
    return e

def extract_table_from_page(pdf, page_id):
  # Columns
  label_elements = []

  for name in COLUMN_NAMES:
    label_query = f'LTPage[pageid="{page_id}"] LTTextLineHorizontal:contains("{name}")'
    element = pdf.pq(label_query)

    if not element:
      print(f"Warning: Column '{name}' not found on page {page_id}.")

    label_elements.append(element)

  if len(label_elements) < 5:
    print(f"Error: Found less than 5 labels on {page_id}.")
    return []

  columns = []

  for current, next in zip(label_elements, label_elements[1:]):
    col_start = float(current.attr('x0'))
    col_stop = float(next.attr('x0'))

    if col_start is None or col_stop is None:
      print(f"Error: Column boundaries are None on page {page_id}.")
      return []

    columns.append((col_start - COLUMN_BUFFER, col_stop + COLUMN_BUFFER))

  last_start = float(label_elements[-1].attr('x0')) - COLUMN_BUFFER
  last_stop = float(label_elements[-1].attr('x1')) + COLUMN_BUFFER

  columns.append((last_start, last_stop))

  # Rows
  line_query = f'LTPage[pageid="{page_id}"] LTLine[width="1.0"]'
  lines = pdf.pq(line_query)
  table_top = float(label_elements[0].attr('y0'))
  distinct_lines = list({y for line in lines if (y := float(line.attrib['y0'])) < table_top})
  distinct_lines.append(table_top)
  distinct_lines.sort(reverse=True)

  rows = []

  for current, next in zip(distinct_lines, distinct_lines[1:]):
    rows.append((next - COLUMN_BUFFER, current + COLUMN_BUFFER))

  # Data
  page_data = []

  for row_start, row_end in rows:
    row_contents = {}

    for index, (col_start, col_end) in enumerate(columns):
      cell_query = f"LTTextLineHorizontal:in_bbox('{col_start}, {row_start}, {col_end}, {row_end}')"
      cell = pdf.pq(cell_query)
      row_contents[COLUMN_NAMES[index]] = cell.text()

    page_data.append(row_contents)

  return page_data

def sanitize_cell_value(value):
  if value:
    return value.replace(',', '').strip()
  return 

def export_to_csv(data, output_path):
  with open(output_path, mode='w', newline='', encoding='utf-8') as csv_file:
    writer = DictWriter(csv_file, fieldnames=COLUMN_NAMES)
    writer.writeheader() 
    for row in data:
      cleaned_row = {key: sanitize_cell_value(value) for key, value in row.items()}
      writer.writerow(cleaned_row)

  print(f"Data exported to {output_path}")

def generate_output_filename(pdf_path):
  filename = pdf_path.split("/")[-1].split(" ")
  account_type = filename[0].lower()
  date = filename[-1].split(".")[0].replace(" ", "_").replace("-", "_")
  return "_".join([account_type, date]) + ".csv"

def get_files_in_folder(folder_path):
  files = []

  for filename in listdir(folder_path):
    account_type = filename.split("/")[-1].split(" ")[0]
    full_path = path.join(folder_path, filename)

    if path.isfile(full_path) and filename.endswith('.pdf') and account_type == 'Chequing':
      files.append(full_path)

  files.sort()
  return files

def parse_bank_statement(pdf_path, number_of_pages, use_cache=False):
  pdf, number_of_pages = load_pdf(pdf_path, use_cache)

  if not pdf:
    raise SystemError("Failed to load PDF.")

  # pages = pdf.pq('LTPage')

  data = []

  for i, _ in enumerate(number_of_pages, start=1):
    page_data = extract_table_from_page(pdf, i)

    for rows in page_data:
      data.append(rows)

  if not path.exists('./data'):
    makedirs('./data')
  
  output_filename = './data/' + generate_output_filename(pdf_path)
  export_to_csv(data, output_filename)

def parse_bank_statements_from_folder(folder_path):
  files = get_files_in_folder(folder_path)

  for file in files:
    parse_bank_statement(file)