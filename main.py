"""
    Note: Please manually delete <Data> </Data> tag and <?xml version="1.0" encoding="UTF-8"?> tag
before the <Data> tag from created XML file, after you will be able successfully upload it

"""


import pandas as pd
import xml.etree.ElementTree as ET
from custom import beginning, end
from collections import Counter


# write the 1)name of a source file,  2)name of a new xml file  3)name of a new exel file
df = pd.read_excel('name_of_source_exel_file.xlsx')
name_of_XML = 'name_of_final_xml_file.xml'
name_of_xlsx = 'name_of_final_exel_file.xlsx'

# n is the quantity of digits coming after decimal point for Brutto Weight
n = 2


# names of custom columns from the source file
sh_code = 'TNVEDCode'
name = 'Description'
mn = 'DEIGoodsQuantity'
quantity = 'Quantity'
mb = 'GrossWeight'
price = 'InvoiceCost'

# extract the columns needed from xlsx source file
df2 = df[[sh_code, name, quantity, mb, price]]

# defining an empty dataframe to store the final data in
final_df = pd.DataFrame(columns=[sh_code, name, mn, mb, price])

# group by code 'Код ТН ВЭД'
for group in df2.groupby(sh_code):

    # if sum length of the 'description' of each group is less or equal than 800, format and combine rows of the group into one
    if len(''.join(list(set(list(group[1][name]))))) <= 800:
        names = [f'{key} {value} шт; ' for key, value in Counter(group[1][name]).items()]
        # print(names)
        ser1 = pd.Series({sh_code: int(group[0]),
                         name: ''.join(names)+'(ТОВАРЫ ЕАЭС)',
                         # mn: group[1]['Масса нетто'].sum(),
                         mn: group[1][quantity].sum(),
                         mb: round(group[1]['Масса брутто'].sum(), n),
                         price: round(group[1][price].sum(), 2)
                         })

        final_df = final_df._append(ser1, ignore_index=True)

    else:    # calculate how many times we should split the group
        n = int(len(group[1][name].sum()) / 600) + 1
        num = int(len(group[1])/n)

        # Split the group into multiple rows with a length of num
        rows = [pd.Series({sh_code: int(group[0]),
                           name: ''.join([f'{key} {value}шт; ' for key, value in Counter(group[1][name].iloc[i:i+num]).items()])+'(ТОВАРЫ ЕАЭС)',
                               # mn: group[1]['Масса нетто'].iloc[i:i+num].sum(),
                               mn: group[1][quantity].iloc[i:i+num].sum(),
                               mb: round(group[1]['Масса брутто'].iloc[i:i+num].sum(), 2),
                               price: round(group[1][price].iloc[i:i+num].sum(), 2)})
                            for i in range(0, len(group[1]), num)]

        # append new rows to final dataframe
        for row in rows:
            final_df = final_df._append(row, ignore_index=True)



# define tags needed for creating of an XML file
final_df.columns = ['TNVEDCode', 'Description', 'DEIGoodsQuantity', 'GrossWeight', 'InvoiceCost']

# Export to an XLSX string
exel = final_df.to_excel(name_of_xlsx)

# Export to an XML string
final_df['TNVEDCode'] = final_df['TNVEDCode'] * 10
xmlString = final_df.to_xml(name_of_XML)


#  Adding < cat >   &  < tcat > tags to XML

# Read the existing XML file with the correct encoding
tree = ET.parse(name_of_XML)
root = tree.getroot()

# Find all the <cat:InvoiceCost> tags
invoice_costs = root.findall('.//InvoiceCost')

# Iterate over each <cat:InvoiceCost> tag
for invoice_cost in invoice_costs:
    # Create the <tcat:GoodsCost> element
    goods_cost_element = ET.Element('tcat:GoodsCost')

    # Get the parent of the <cat:InvoiceCost> element
    invoice_cost_parent = None

    # Find the parent element by traversing the XML tree
    for elem in root.iter():
        if invoice_cost in elem:
            invoice_cost_parent = elem
            break

    if invoice_cost_parent is not None:
        # Get the index of the <cat:InvoiceCost> element within its parent
        invoice_cost_index = list(invoice_cost_parent).index(invoice_cost)

        # Remove the <cat:InvoiceCost> element from its parent
        invoice_cost_parent.remove(invoice_cost)

        # Insert the <tcat:GoodsCost> element at the same index
        invoice_cost_parent.insert(invoice_cost_index, goods_cost_element)

        # Append the <cat:InvoiceCost> element to the <tcat:GoodsCost> element
        goods_cost_element.append(invoice_cost)

# Write the modified XML back to a file
tree.write(name_of_XML, encoding='UTF-8', xml_declaration=True)


# Add 'cat' namespace prefix and modify 'code', 'description', and 'weight' tags
cat_namespace = {'cat': 'http://example.com/cat'}
for i in ['TNVEDCode', 'Description', 'GrossWeight', 'InvoiceCost']:
    for element in root.iter(i):
        element.tag = f"cat:{element.tag}"

# Add 'tcat' namespace prefix and modify 'number' and 'gweight' tags
tcat_namespace = {'tcat': 'http://example.com/tcat'}
for element in root.iter('DEIGoodsQuantity'):
    element.tag = f"tcat:{element.tag}"

# Write the modified XML to a new file
tree.write(name_of_XML, encoding='utf-8', xml_declaration=True, default_namespace='', method='xml')


#  Adding   beginning and end to XML

# Read the existing XML file
with open(name_of_XML, 'r', encoding='utf-8') as file1:
    existing_xml = file1.read()

# XML code to add at the beginning
xml_to_add_beginning = beginning

# XML code to add at the end
xml_to_add_end = end

# Concatenate the XML code
modified_xml = xml_to_add_beginning + existing_xml + xml_to_add_end

# Write the modified XML content to a new file or overwrite the existing file
with open(name_of_XML, 'w', encoding='utf-8') as file2:
    file2.write(modified_xml)

# Adding <tcat:Goods> tags
with open(name_of_XML, 'r', encoding='utf-8') as file:
    xml_data = file.read()

# Replace <row> tags with <tcat:Goods> tags
modified_xml_data = xml_data.replace('<row>', '<tcat:Goods>').replace('</row>', '</tcat:Goods>')

with open(name_of_XML, 'w', encoding='utf-8') as file:
    file.write(modified_xml_data)





