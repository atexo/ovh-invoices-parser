#!.venv/bin/python
# -*- coding: utf-8 -*-

import csv
import json
import os
import re
import sys
import time

from datetime import datetime
from tika import parser 
from unidecode import unidecode

INPUT_FOLDER = "./input"
OUTPUT_FOLDER = "./output"

def main(argv=None):

    all_files_data = []

    files = [f for f in os.listdir(INPUT_FOLDER) if os.path.isfile(f'%s/%s'%(INPUT_FOLDER, f)) and f.endswith('.pdf')]

    for file in files:
        print(f'Processing file %s'%(file))

        # Parse data from PDF file using tika
        data = parsePdf(f'%s/%s'%(INPUT_FOLDER, file))

        # Sanitize data 
        sanitized_data = sanitizePDFExtraction(data.splitlines())     

        # Read items
        items = extractItems(sanitized_data)
        invoice = OVHInvoice(items)
        all_files_data.extend(invoice.get_items())

        # Write to CSV
        writeToCsv(invoice.get_items(), f'%s/%s'%(OUTPUT_FOLDER, file.replace(".pdf", ".csv")))

    # Write to CSV
    writeToCsv(all_files_data, f'%s/%s'%(OUTPUT_FOLDER, "report.csv"))

    sys.exit(0)


class OVHInvoiceItem:
    """
    Describe a single item in a OVH invoice. This object will stored parsed data, including the period if present.
    """
    def __init__(
        self,
        invoice,
        section,
        description,
        reference,
        unit_count,
        unit_price,
        price,
        period_start,
        period_end
    ):
        self._invoice = unidecode(invoice)
        self._section = unidecode(section)
        self._description = unidecode(description)
        self._reference = unidecode(reference)
        self._unit_count = float(unit_count.replace(",","."))
        self._unit_price = float(unit_price.replace(",","."))
        self._price = float(price.replace(",","."))
        self._period_start = handleDate(period_start)
        self._period_end = handleDate(period_end)

    def __repr__(self):
        return f"OVHInvoiceItem({self._invoice!r}, {self._section!r}, {self._description!r}, {self._reference!r}, {self._unit_count!r}, {self._unit_price!r}, {self._price!r}, {self._period_start!r}, {self._period_end!r})"

    def __iter__(self):
        return iter([self._invoice, self._section, self._description, self._reference, self._unit_count, self._unit_price, self._price, self._period_start, self._period_end])

    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__,
            sort_keys=True, indent=4,  ensure_ascii=False)

    def get_invoice(self):
        return self._invoice
    def get_section(self):
        return self._section
    def get_description(self):
        return self._description
    def get_reference(self):
        return self._reference
    def get_unit_count(self):
        return self._unit_count
    def get_unit_price(self):
        return self._unit_price
    def get_price(self):
        return self._price
    def get_period_start(self):
        return self._period_start
    def get_period_end(self):
        return self._period_end


        
class OVHInvoice:
    """
    Describe a OVH invoice.
    """
    def __init__(
        self,
        items
    ):
        self.items = items

    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__,
            sort_keys=True, indent=4,  ensure_ascii=False)

    def get_items(self):
        return self.items

def parsePdf(input_file):
    
    # opening pdf file
    parsed_pdf = parser.from_file(input_file)
  
    # saving content of pdf
    # you can also bring text only, by parsed_pdf['text'] 
    # parsed_pdf['content'] returns string
    return parsed_pdf['content'] 

def sanitizePDFExtraction(data):
    """
    Clean data extracted from PDF (handle lines breaks).
    Returns an array of sanitized lines.
    """

    sanitized_data = []
    
    buffer = []
    is_item = False

    re_end_of_item = re.compile(r",(\d{2})\s€$")

    for line in data:
        line = line.strip()
        if len(line):
            # Référence de la facture start pattern
            if line.startswith("Référence de la facture"):
                sanitized_data.append(line)
                
            # Rubrique start pattern
            if line.startswith("Rubrique"):
                sanitized_data.append(line)

            # Header start pattern
            if line.startswith("Abonnement Référence Quantité"):
                is_item = False
                buffer = []
                buffer.append(line)

            # Header end pattern
            if line.endswith("Prix HT"):
                # Ends header parsing and start item parsing
                is_item = True
                buffer.append(line)
                sanitized_data.append(" ".join(buffer).strip())
                buffer = []
                continue

            # Page end pattern
            if line.endswith("javascript:history.back()"):
                buffer.append(line)
                sanitized_data.append(" ".join(buffer).strip())
                buffer = []
                continue

            # Handle item
            if is_item:
                buffer.append(line)

                if re_end_of_item.search(line):
                    sanitized_data.append(" ".join(buffer).strip())
                    buffer = []
                    continue
                
    return sanitized_data

def handleDate(date_as_string):
    try:
        date_time = datetime.strptime(date_as_string, '%d/%m/%Y').date()
        return time.mktime(date_time.timetuple())
    except:
        return ""

def extractItems(sanitized_data):
    """
    Parse and extract items from sanitized data.
    Returns an array of OVHInvoiceItem objects
    """

    items = []
    invoice = ""
    rubrique = ""

    for line in sanitized_data:
        
        # Rubrique start pattern
        if line.startswith("Rubrique"):
            rubrique = line.replace("Rubrique ", "")

        if line.startswith("Référence de la facture"):
            invoice = line.replace("Référence de la facture : ", "")

        # Handle spaces in reference
        line = re.sub(r'(\S-)\s', r'\g<1>', line)

        # Handle thousand separator
        prices = re.findall(r"\s\d*\s([\-]?(?:\d+|\d{1,3}(?:\s\d{3})*)(?:\,\d*)?\s€)\s([\-]?(?:\d+|\d{1,3}(?:\s\d{3})*)(?:\,\d*)?\s€)", line)
        for price in prices:
            for data in price:
                line = line.replace(data, data.replace(" ", ""), 1)
  
        re_item_line = re.compile(r"(.*)\s([^ ]*)\s(\d*)\s([\-]?\d*,\d*)€\s([\-]?\d*,\d*)€")
        
        if re_item_line.search(line):

            # Execute regex on processed line
            re_groups = re.search(r"(.*)\s([^ ]*)\s(\d*)\s([\-]?\d*,\d*)€\s([\-]?\d*,\d*)€", line).groups()
            
            # Create item object
            if len(re_groups) == 5:
                try:
                    re_date_groups = re.search(r"(\d{2}\/\d{2}\/\d{4})-(\d{2}\/\d{2}\/\d{4})\)", re_groups[0]).groups()
                except:
                    re_date_groups = ["", ""]

                item = OVHInvoiceItem(
                    invoice,
                    rubrique,
                    re_groups[0],
                    re_groups[1],
                    re_groups[2],
                    re_groups[3],
                    re_groups[4],
                    re_date_groups[0],
                    re_date_groups[1],
                    )
                
                items.append(item)
    return items

def writeToCsv(data, output):

    with open(output, "w") as stream:
        writer = csv.writer(stream, delimiter=';')
        writer.writerow(["invoice", "section", "description", "reference", "unit_count", "unit_price", "price", "period_start", "period_end"])
        writer.writerows(data)
  
if __name__ == "__main__":
    main(sys.argv[1:])
