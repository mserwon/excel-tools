import sys
import os
import json
import re
import argparse
import jinja2
import datetime
from datetime import timedelta
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom


def parse_xml_ff(cdict, xmlfile):

    with open(xmlfile) as ff:
        ffs = ff.read()
    ff.close()

    for key, value in cdict.items():
        ffs, cn = re.subn(key, value, ffs)
        print(key, ' - ', cn)

    ff = open(xmlfile, 'w')
    ff.write(ffs)
    ff.close()

def convert_csv_to_xml(spreadsheet,xmlfile):
    input_csv_file = spreadsheet
    export_xml_file = xmlfile

    if not input_csv_file or not export_xml_file:
        print("Please select both CSV file and export xml file.")
        return

    try:
        df = pd.read_csv(input_csv_file)
        xml_data = df.to_dict(orient='records')
        root = ET.Element("data")
        for item in xml_data:
            record = ET.SubElement(root, "record")
            for key, value in item.items():
                ET.SubElement(record, key).text = str(value)

        xml_string = minidom.parseString(ET.tostring(root)).toprettyxml(indent="   ")
        xml_file_path = f"{export_xml_file}"
        with open(xml_file_path, "w") as xml_file:
            xml_file.write(xml_string)

        print(f"CSV file converted and saved as XML: {xml_file_path}")
    except Exception as e:
        print(f"Error occurred: {str(e)}")

def cmd_line_parser():
    cfg = argparse.ArgumentParser(
        description="Uses spreadsheet files to create xml "
    )
    cfg.add_argument("-o", "--output-type", default="xml",help="Type of the query results brief,json,html")
    cfg.add_argument("-j", "--json-file",default="", help="json input file")
    cfg.add_argument("-s", "--spreadsheet",default="", help="csv or xlsx input file")
    cfg.add_argument("-x", "--xmlfile",default="", help="xml output file")
    cfg.add_argument("searchobj", help="Item type to find, user, callhandler, dirn",default="")
    return cfg

def main():
    cfg = cmd_line_parser()
    opt = cfg.parse_args()
    opt.output_type = opt.output_type.lstrip()
    opt.spreadsheet = opt.spreadsheet.lstrip()
    opt.xmlfile = opt.xmlfile.lstrip()
    opt.searchobj = opt.searchobj.lstrip()
              
    convert_csv_to_xml(opt.spreadsheet,opt.xmlfile)
    qfound = 1
    if opt.searchobj == 'inventory':
        cdict = {
            r'data>': r'INVENTORY>',
            r'record>': r'ITEM>',
            r'<\?xml.+?>\n': r'',
            r'   <QTYFILLED>nan</QTYFILLED>\n   ': r''
        }
        parse_xml_ff(cdict, opt.xmlfile)
    elif opt.searchobj == 'data':
        parse_xml_ff(cdict, opt.xmlfile)
if __name__ == '__main__':
    main()