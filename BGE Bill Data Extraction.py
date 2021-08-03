#!/usr/bin/env python
# coding: utf-8

# **Things to extract**
# 
# ~~- Utility Name~~
# ~~- Issued Date~~
# 
# ~~- Company Name~~
# ~~- Street~~
# ~~- City~~
# ~~- State~~
# ~~- Zip~~
# 
# 
# ~~- Electric Choice ID#~~
# ~~- Rate Code~~
# 
# ~~- Usage~~
# ~~- Supplier~~
# 

# In[1]:


import pdfplumber
import numpy as np
import pandas as pd
import openpyxl


# In[2]:


### FIRST PAGE ###

# page - a Page object from the pdfplumber module
# creates coordinates for address bounding box to extract text from
# returns address_bounding_box - tuple of ints/floats containing coordinates of address bounding box 
def setAddressBoundingBox(page):
    
    #get page height and width for box coordinate calculations
    page_height = page.height
    page_width = page.width
    
    #address box is near second top quarter
    top_half = page.height/2
    top_quarter = page.height/4
    second_quarter = top_half - top_quarter
    
    #coordinates of the address box
    address_box_left = (page.width//2) - 20
    address_box_top = second_quarter - 75
    address_box_right = page.width
    address_box_bottom = (page.height//2) -215
    
    address_bounding_box = (address_box_left,address_box_top,address_box_right,address_box_bottom)
    
    return address_bounding_box


# page - a Page object from the pdfplumber module
# address_bounding_box - tuple of coordinates of address box
# crops page to the bounding box and extracts text from it
# returns address_extract_text - str containing extracted text from address box
def extractAddressBoxText(page, address_bounding_box):
    
    address_extract_text = page.crop(address_bounding_box).extract_text(x_tolerance=1)
    
    return address_extract_text


# address_extract_text - str containing extracted text from address box
# returns company_name_text - str containing name of company
def setCompanyName(address_extract_text):
    
    company_name_text = address_extract_text.splitlines()[0]
    
    return company_name_text


# address_extract_text - str containing extracted text from address box
# returns street_text - str containing street name 
def setStreet(address_extract_text):

    street_text = address_extract_text.splitlines()[1]
    
    return street_text


# address_extract_text - str containing extracted text from address box
# returns city_text, state_text, zip_code_text - tuple of str containing city, state, and zip code respectively
def setCityStateZIP(address_extract_text):
    
    city_state_zip_text = address_extract_text.splitlines()[2]
    city_state_zip_text = city_state_zip_text.replace(',', '')
    
    city_text = city_state_zip_text.split(' ')[0]
    state_text =  city_state_zip_text.split(' ')[1]
    zip_code_text =  city_state_zip_text.split(' ')[2]
    
    return city_text, state_text, zip_code_text


# address_extract_text - str containing extracted text from address box
# returns acc_num_text - str containing account number
# NOT USED AS IT TURNS OUT THE ACCOUNT NUMBER IS NOT NEEDED
def setAccNum(address_extract_text): 
    
    acc_num_text = address_extract_text.splitlines()[3]
    acc_num_text = acc_num_text.split('#')[1]
    acc_num_text = acc_num_text.strip()
    
    return acc_num_text

# address_extract_text - str containing extracted text from address box
# returns issued_date_text - str containing issued date 
def setIssuedDate(address_extract_text):
    
    issued_date_text = address_extract_text.splitlines()[4]
    issued_date_text = issued_date_text.split(':')[1]
    issued_date_text = issued_date_text.strip()
    
    return issued_date_text

# page - a Page object from the pdfplumber module
# returns electric_supply_bounding_box - tuple of ints/float containing coordinates of electric supply box
def setElectricSupplyBoundingBox(page):
    
    top_half = page.height/2
    top_quarter = page.height/4
    second_quarter = top_half - top_quarter

    electric_supply_box_top  = second_quarter - 75
    electric_supply_box_bottom = (page.height//2) - 215
    electric_supply_box_left = (page.width//2) - 130
    electric_supply_box_right = (page.width//2)
    
    electric_supply_bounding_box = (electric_supply_box_left,electric_supply_box_top,electric_supply_box_right,electric_supply_box_bottom)
    
    return electric_supply_bounding_box


# page - a Page object from the pdfplumber module
# electric_supply_bounding_box - tuple of coordinates of electric supply box
# crops page to the bounding box and extracts text from it
# returns electric_supply_extract_text - str containing extracted text from address box
def extractElectricSupplyBoxText(page, electric_supply_bounding_box):
    
    electric_supply_extract_text = page.crop(electric_supply_bounding_box).extract_text(x_tolerance=1)
    
    return electric_supply_extract_text

# electric_supplier_extract_text - str containing extracted text from electric supply box
# returns electric_supplier_text - str containing name of the electric supplier
def setElectricSupplier(electric_supply_extract_text):
    
    electric_supplier_text = electric_supply_extract.splitlines()[1]
    
    return electric_supplier_text

# electric_supplier_extract_text - str containing extracted text from electric supply box
# returns electric_choice_id_text - str containing electric choice id number
def setElectricChoiceID(electric_supply_extract_text):
    
    electric_choice_id_text = electric_supply_extract.splitlines()[4]
    electric_choice_id_text = electric_choice_id_text.split('Electric Choice ID:')[1].strip()
    
    return electric_choice_id_text

# page - a Page object from the pdfplumber module
# returns utility_name_text  - str containing name of utility
def setUtilityName(page):
    
    page_extract = page.extract_text(x_tolerance=1, y_tolerance=1)
    utility_name_text = page_extract.splitlines()[-6]
    
    return utility_name_text

### SECOND PAGE ###

# page - a Page object from the pdfplumber module
# creates coordinates for rate bounding box to extract text from
# returns rate_bounding_box - tuple of ints/floats containing coordinates of rate bounding box 
def setElectricRateBoundingBox(page):
    
    rate_box_left = 20
    rate_box_top = (page.width//4) + 5
    rate_box_right = (page.width//2) - 100
    rate_box_bottom = (page.width//4) + 50
    rate_bounding_box = (rate_box_left, rate_box_top, rate_box_right, rate_box_bottom)
    
    return rate_bounding_box

# page - a Page object from the pdfplumber module
# rate_bounding_box - tuple of ints/floats containing coordinates of rate bounding box 
# returns rate_extract_text - str containing extracted text from rate box
def extractRateBoxText(page, rate_bounding_box):
    
    rate_extract_text = second_page.crop(rate_bounding_box).extract_text(x_tolerance = 1)
    
    return rate_extract_text

# rate_extract_text - str containing extracted text from rate box
# returns rate_text - str containing rate code
def setRate(rate_extract_text):

    #gets the first line after the word 'Service' and remove the leading whitespace and '-'
    rate_text = rate_extract_text.split('Service')[1].splitlines()[0].replace('-','',1).strip()
    
    if 'TOU -' in rate_text:
        
        #remove TOU and leading '-' for uniformity
        rate_text = rate_text.split('TOU')[1].replace('-','',1).strip()
    
    return rate_text

# page - a Page object from the pdfplumber module
# creates coordinates for usage bounding box to extract text from
# returns usage_bounding_box - tuple of ints/floats containing coordinates of usage bounding box 
def setElectricUsageBoundingBox(page):
    
    usage_box_left = (page.width//2) - 130
    usage_box_top = (page.height//2) - 180
    usage_box_right = (page.width//2) - 20
    usage_box_bottom = (page.height//2) - 150
    usage_bounding_box = (usage_box_left,usage_box_top,usage_box_right,usage_box_bottom)
    
    return usage_bounding_box

# page - a Page object from the pdfplumber module
# usage_bounding_box - tuple of ints/floats containing coordinates of usage bounding box
# returns usage_extract_text - str containing extracted text from usage box
def extractUsageBoxText(page, usage_bounding_box):
    
    usage_extract_text = page.crop(usage_bounding_box).extract_text()
    
    return usage_extract_text

def setUsage(usage_extract_text):
    
    usage_text = usage_extract_text.splitlines()[0]
    
    return usage_text


    


# In[3]:


# electric_supplier_extract_text - str containing extracted text from electric supply box
# returns electric_choice_id_text - str containing electric choice id number
def setElectricChoiceIDNoSupp(electric_supply_extract_text):
    
    electric_choice_id_text = electric_supply_extract.splitlines()[0]
    electric_choice_id_text = electric_choice_id_text.split('Electric Choice ID:')[1].strip()
    
    return electric_choice_id_text


# In[19]:


# page - a Page object from the pdfplumber module
# returns gas_supply_bounding_box - tuple of ints/float containing coordinates of electric supply box
def setGasSupplyBoundingBox(page):
    
    top_half = page.height/2
    top_quarter = page.height/4
    second_quarter = top_half - top_quarter

    gas_supply_box_top  = second_quarter + 80
    gas_supply_box_bottom = (page.height//2) + 20
    gas_supply_box_left = (page.width//2) - 130
    gas_supply_box_right = (page.width//2)
    
    gas_supply_bounding_box = (gas_supply_box_left,gas_supply_box_top,gas_supply_box_right,gas_supply_box_bottom)
    
    return gas_supply_bounding_box

# page - a Page object from the pdfplumber module
# gas_supply_bounding_box - tuple of coordinates of gas supply box
# crops page to the bounding box and extracts text from it
# returns gas_supply_extract_text - str containing extracted text from gas box
def extractGasSupplyBoxText(page, gas_supply_bounding_box):
    
    gas_supply_extract_text = page.crop(gas_supply_bounding_box).extract_text(x_tolerance=1)
    
    return gas_supply_extract_text

# electric_supplier_extract_text - str containing extracted text from electric supply box
# returns electric_choice_id_text - str containing electric choice id number
def setGasChoiceID(gas_supply_extract_text):
    
    gas_choice_id_text = gas_supply_extract_text#.splitlines()[0]
    gas_choice_id_text = gas_choice_id_text.split('Gas Choice ID:')[1].strip()
    
    return gas_choice_id_text
 


# In[4]:


file_name = input("Name of PDF file: ")

# file_name = 'PII Example 1'
#with pdfplumber.open('Bills/'+ file_name + '.pdf') as pdf:

#initializing variables outside scope of open file
utility_name = ''
company_name = ''
street = ''
city = ''
state = ''
zip_code = ''
issued_date = ''
electric_supplier = ''
electric_choice_id = ''
gas_choice_id = ''
electric_rate = ''
electric_usage = ''
bill_dict = {}


#adding logic for different cases

print('What type of bill is this?')
print('e - electricity only    g - gas only    eg - electricity and gas ')
bill_type = input('bill type: ')

while (bill_type != 'e') and (bill_type != 'g') and (bill_type != 'eg'):
    print('Invalid choice. Try again:')
    bill_type = input('bill type: ')

print('Is there a supplier on the bill?')
print('yes    no')
supplier_present = input().lower()

while (supplier_present != 'yes') and (supplier_present != 'no'):
    print('Invalid choice. Try again:')
    supplier_present = input()



# In[5]:


if (bill_type == 'e') and (supplier_present == 'no'): 
     with pdfplumber.open(f'Bills/{file_name}.pdf') as pdf:
        first_page = pdf.pages[0]
        second_page = pdf.pages[1]

        #extracting from address box
        address_extract = extractAddressBoxText(first_page, setAddressBoundingBox(first_page))

        company_name = setCompanyName(address_extract)
        street = setStreet(address_extract)
        city, state, zip_code = setCityStateZIP(address_extract)
        issued_date = setIssuedDate(address_extract)

        #extracting from electric supply box
        electric_supply_extract = extractElectricSupplyBoxText(first_page, setElectricSupplyBoundingBox(first_page))
        electric_supplier = np.nan
        electric_choice_id = setElectricChoiceIDNoSupp(electric_supply_extract) 
        
        #extracting utility name from the bottom of the page
        utility_name = setUtilityName(first_page)

        #extracting from electric_rate box
        electric_rate_extract = extractRateBoxText(second_page, setElectricRateBoundingBox(second_page))
        electric_rate = setRate(electric_rate_extract)

        #extracting from electric_usage
        electric_usage_extract = extractUsageBoxText(second_page, setElectricUsageBoundingBox(second_page))
        electric_usage = setUsage(electric_usage_extract)

bill_dict = dict(utility = utility_name, issued_date = issued_date, company = company_name, 
                street = street, city = city, state = state, zip_code = zip_code, 
                electric_choice_id = electric_choice_id, 
                electric_rate_code = electric_rate, electric_supplier = electric_supplier, electric_usage = electric_usage)

    #10.22.20 E, 11.18.20 E format looks slightly different so doesn't work

# bill_dict
        
        


# In[6]:


if (bill_type == 'e') and (supplier_present == 'yes'): 
    with pdfplumber.open(f'Bills/{file_name}.pdf') as pdf:
        first_page = pdf.pages[0]
        second_page = pdf.pages[1]

        #extracting from address box
        address_extract = extractAddressBoxText(first_page, setAddressBoundingBox(first_page))

        company_name = setCompanyName(address_extract)
        street = setStreet(address_extract)
        city, state, zip_code = setCityStateZIP(address_extract)
        issued_date = setIssuedDate(address_extract)

        #extracting from electric supply box
        electric_supply_extract = extractElectricSupplyBoxText(first_page, setElectricSupplyBoundingBox(first_page))

        electric_supplier = setElectricSupplier(electric_supply_extract)
        electric_choice_id = setElectricChoiceID(electric_supply_extract)

        #extracting utility name from the bottom of the page
        utility_name = setUtilityName(first_page)

        #extracting from rate box
        rate_extract = extractRateBoxText(second_page, setRateBoundingBox(second_page))
        rate = setRate(rate_extract)

        #extracting from usage
        usage_extract = extractUsageBoxText(second_page, setUsageBoundingBox(second_page))
        usage = setUsage(usage_extract)

bill_dict = dict(utility = utility_name, issued_date = issued_date, company = company_name, 
                street = street, city = city, state = state, zip_code = zip_code, 
                electric_choice_id = electric_choice_id, 
                rate_code = rate, electric_supplier = electric_supplier, usage = usage)

    #10.22.20 E, 11.18.20 E format looks slightly different so doesn't work

# bill_dict


# In[8]:


if (bill_type == 'eg') and (supplier_present == 'no'):
    with pdfplumber.open(f'Bills/{file_name}.pdf') as pdf:
        first_page = pdf.pages[0]
        second_page = pdf.pages[1]
        
        #extracting from address box
        address_extract = extractAddressBoxText(first_page, setAddressBoundingBox(first_page))

        company_name = setCompanyName(address_extract)
        street = setStreet(address_extract)
        city, state, zip_code = setCityStateZIP(address_extract)
        issued_date = setIssuedDate(address_extract)
        
        #extracting from electric supply box
        electric_supply_extract = extractElectricSupplyBoxText(first_page, setElectricSupplyBoundingBox(first_page))
        electric_supplier = np.nan
        electric_choice_id = setElectricChoiceIDNoSupp(electric_supply_extract) 
        
        #extracting from gas supply box
        gas_extract_text = first_page.crop(setGasSupplyBoundingBox(first_page)).extract_text(x_tolerance=1)
        gas_choice_id = setGasChoiceID(gas_extract_text)
        
        
        
        #extracting utility name from the bottom of the page
        utility_name = setUtilityName(first_page)


# In[23]:


bill_df = pd.DataFrame([bill_dict])
bill_df.columns = [col.replace('_',' ') for col in bill_df]
bill_df.columns = [col.title() for col in bill_df]
bill_df.rename(columns = {
    'Electric Choice Id': 'Electric Choice ID',
    'Electric Usage' : 'Usage (in kWh)'
}, inplace = True)

bill_df


# In[11]:


bill_df.to_excel(f'{file_name}.xlsx')

