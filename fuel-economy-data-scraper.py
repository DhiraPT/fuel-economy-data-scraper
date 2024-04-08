import requests
from bs4 import BeautifulSoup
from datetime import datetime
import xlsxwriter

excel_file_name = input('Input file name (e.g. fuel_economy.xlsx): ')
while excel_file_name == '':
    excel_file_name = input('Input file name (e.g. fuel_economy.xlsx): ')

print('Scraping fuel economy data...')

timestamp = int(datetime.timestamp(datetime.now()) * 1000)
client = requests.session()

# Get makes and make_values
res = client.get('https://vrl.lta.gov.sg/lta/vrl/action/pubfunc?ID=FuelCostCalculator')
soup = BeautifulSoup(res.text, 'html.parser')
soup = soup.find("select", id="carOneMaker")
make_options = soup.find_all("option")

makes = []
for option in range(len(make_options)):
    if option == 0:
        continue
    make = make_options[option].text
    make_value = make_options[option]['value']
    makes.append({"make": make, "make_value": make_value})

for make_dict in makes:
    models = []
    make_value = make_dict['make_value']
    res = client.get(f'https://vrl.lta.gov.sg/vrl/action/ajaxLoadFuelCostCalculatorAction?FUNCTION_ID=F2305001ET&parmVal={make_value}&typ=car1MakeTyp')
    soup = BeautifulSoup(res.text, 'html.parser')
    model_options = soup.find_all("option")

    for option in range(len(model_options)):
        if option == 0:
            continue
        model = model_options[option].text
        model_value = model_options[option]['value']

        res = client.get(f'https://vrl.lta.gov.sg/vrl/action/ajaxLoadFuelCostCalculatorAction?FUNCTION_ID=F2305001ET&_={timestamp}&parmVal={model_value}&typ=car1ModelTyp')
        models.append({"model": model, "model_value": model_value, "data": res.json()})

    make_dict['models'] = models

print(f'Writing data into {excel_file_name}...')

workbook = xlsxwriter.Workbook(excel_file_name)
worksheet = workbook.add_worksheet()

worksheet.write_string(0, 0, 'Make')
worksheet.write_string(0, 1, 'Model')
worksheet.write_string(0, 2, 'Body Type')
worksheet.write_string(0, 3, 'Engine(cc)')
worksheet.write_string(0, 4, 'MPO(kW)')
worksheet.write_string(0, 5, 'Fuel Type')
worksheet.write_string(0, 6, 'Transmission Type')
worksheet.write_string(0, 7, 'Turbo/Supercharged')
worksheet.write_string(0, 8, 'Hybrid')
worksheet.write_string(0, 9, 'VES Band')
worksheet.write_string(0, 10, 'CVES Band')

engine_format  = workbook.add_format({"num_format": "0"})
mpo_format = workbook.add_format({"num_format": "0.0"})

row = 1
for make_index in range(len(makes)):
    make_dict = makes[make_index]
    for model_index in range(len(make_dict['models'])):
        model_dict = make_dict['models'][model_index]
        worksheet.write_string(row, 0, make_dict['make'])
        worksheet.write_string(row, 1, model_dict['model'])
        worksheet.write_string(row, 2, model_dict['data']['carOneBodyTyp'])
        worksheet.write_number(row, 3, int(model_dict['data']['carOneEngine']) if model_dict['data']['carOneEngine'] != '-' else 0, engine_format)
        worksheet.write_number(row, 4, float(model_dict['data']['carOneEnginPower']), mpo_format)
        worksheet.write_string(row, 5, model_dict['data']['carOneFuelTyp'])
        worksheet.write_string(row, 6, model_dict['data']['carOneTransTyp'])
        worksheet.write_string(row, 7, model_dict['data']['carOneTurbo'])
        worksheet.write_string(row, 8, model_dict['data']['carOneHybridSys'])
        worksheet.write_string(row, 9, model_dict['data']['carOneVesBand'])
        worksheet.write_string(row, 10, model_dict['data']['carOneCvesBand'])

        row += 1

workbook.close()

print(f'Data has been successfully written to {excel_file_name}')
