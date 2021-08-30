# set export DISPLAY=localhost:0.0 before each exec in cmd line
import PySimpleGUI as sg
import extractdata


def collapse(layout, key, visible):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this section visible / invisible
    :param visible: visible determines if section is rendered visible or invisible on initialization
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    # return sg.pin(sg.Column(layout, key=key, visible=visible))
    return sg.Column(layout, key=key, visible=visible)

sg.theme('Dark Blue 3')  # please make your windows colorful

eg_supplier_section = [[sg.Text('Is there a supplier for electricity?', size = (35,1), key = '-CHOOSE_SUPPLIERS_EG-')],
                        [sg.Radio('Yes', 'Supplier Electricity', default = False, key = '-ELECTRIC_SUPPLIER_PRESENT_EG-'), sg.Radio('No', 'Supplier Electricity', default = False)],
                        [sg.Text('Is there a supplier for gas?', size = (35,1))],
                        [sg.Radio('Yes', 'Supplier Gas', default = False, key = '-GAS_SUPPLIER_PRESENT_EG-'), sg.Radio('No', 'Supplier Gas', default = False)],
                        [sg.Text('Are there multiple usage amounts on this bill for electricity?', size = (75,1))],
                        [sg.Radio('Yes', 'Multiple Usage Electricity', default = False, key = '-MULTIPLE_USAGE_E-'), sg.Radio('No', 'Multiple Usage Electricity', default = False)],
                        [sg.Text('Are there multiple usage amounts on this bill for gas?', size = (75,1))],
                        [sg.Radio('Yes', 'Multiple Usage Gas', default = False, key = '-MULTIPLE_USAGE_G-'), sg.Radio('No', 'Multiple Usage Gas', default = False)]]

e_or_g_supplier_section = [[sg.Text('Is there a supplier for this commodity?', size = (75,1))],
                    [sg.Radio('Yes', 'Supplier Electricity or Gas', default = False, key = '-SUPPLIER_PRESENT-'), sg.Radio('No', 'Supplier Electricity or Gas', default = False)],
                    [sg.Text('Are there multiple usage amounts on this bill?', size = (75,1))],
                    [sg.Radio('Yes', 'Multiple Usage Electricity or Gas', default = False, key = '-MULTIPLE_USAGE_E_OR_G-'), sg.Radio('No', 'Multiple Usage Electricity or Gas', default = False)]]

layout = [[sg.Text('Select a bill PDF: ', size = (19,1))],
        
        [sg.Input(), sg.FileBrowse(key = '-BILL_PDF-', file_types = (('PDF files', '*.pdf'),))],
        #   [sg.Text( size=(40,1), key='-OUTPUT_BILL_PDF_FILE_NAME-')],

          [sg.Radio('Electricity Only', 'Bill Type', default = False, key = '-ELECTRICITY_ONLY-', enable_events=True), 
          sg.Radio('Gas Only', 'Bill Type', default = False, key = '-GAS_ONLY-', enable_events=True), 
          sg.Radio('Electricity and Gas ', 'Bill Type', default = False, key = '-ELECTRICITY_GAS-', enable_events=True)],

          [collapse(eg_supplier_section, '-ASK_SUPPLIER_EG-', False), collapse(e_or_g_supplier_section, '-ASK_SUPPLIER_E_OR_G-', False) ],

          [sg.Text('Select the Excel file you would like to output to: ', size = (51,1))],
          [sg.Input(), sg.FileBrowse(key = '-EXCEL_FILE-', file_types = (('Excel files', '*.xlsx'),))],
        #   [sg.Text( size=(40,1), key='-OUTPUT_EXCEL_FILE_NAME-')],
          [sg.Button('OK'), sg.Exit()]]

def drawMainWindow():
    
    window = sg.Window('BGE Bill Extraction', layout)

    
    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        extraction_complete = False
        bill_file_name = values['-BILL_PDF-'].split('/')[-1]
        # window['-OUTPUT_BILL_PDF_FILE_NAME-'].update('Bill selected:  ' + bill_file_name)

        if event == '-ELECTRICITY_GAS-':
            window['-ASK_SUPPLIER_E_OR_G-'].update(visible = False)
            window['-ASK_SUPPLIER_EG-'].update(visible = True)

        elif event == '-GAS_ONLY-' or event == '-ELECTRICITY_ONLY-':
            window['-ASK_SUPPLIER_EG-'].update(visible = False)
            window['-ASK_SUPPLIER_E_OR_G-'].update(visible = True)

        excel_file_name = values['-EXCEL_FILE-'].split('/')[-1]
        # window['-OUTPUT_EXCEL_FILE_NAME-'].update('Excel file selected:  ' + excel_file_name)

        if event == 'OK':

            if values['-ELECTRICITY_ONLY-'] == True:
                
                extractdata.bill_type = 'e'

                if values['-MULTIPLE_USAGE_E_OR_G-'] == True:
                    
                    extractdata.multiple_usage_electricity = 'yes'
                
                else: 

                    extractdata.multiple_usage_electricity = 'no'

            elif values['-GAS_ONLY-'] == True:
                
                extractdata.bill_type = 'g'

                if values['-MULTIPLE_USAGE_E_OR_G-'] == True:

                    extractdata.multiple_usage_gas = 'yes'

                else:

                    extractdata.multiple_usage_gas = 'no'

            elif values['-ELECTRICITY_GAS-'] == True:
                extractdata.bill_type = 'eg'

                if values['-ELECTRIC_SUPPLIER_PRESENT_EG-'] == True:
                
                    extractdata.electric_supplier_present = 'yes'

                else:
                    extractdata.electric_supplier_present = 'no'

                if values['-GAS_SUPPLIER_PRESENT_EG-'] == True:
                    extractdata.gas_supplier_present = 'yes'
                
                else:
                    extractdata.gas_supplier_present = 'no'

            if values['-SUPPLIER_PRESENT-'] == True:
                extractdata.supplier_present = 'yes'
            
            else:
                extractdata.supplier_present = 'no'
            
            # if values['-ELECTRIC_SUPPLIER_PRESENT_EG-'] == True:
            #     extractdata.electric_supplier_present = 'yes'

            # else:
            #     extractdata.electric_supplier_present = 'no'

            # if values['-GAS_SUPPLIER_PRESENT_EG-'] == True:
            #     extractdata.gas_supplier_present = 'yes'
            
            # else:
            #     extractdata.gas_supplier_present = 'no'

            if values['-MULTIPLE_USAGE_E-'] == True:
                extractdata.multiple_usage_electricity = 'yes'

            else:
                extractdata.multiple_usage_electricity = 'no'

            if values['-MULTIPLE_USAGE_G-'] == True:
                extractdata.multiple_usage_gas = 'yes'

            else:
                extractdata.multiple_usage_gas = 'no' 
            

            extraction_complete = extractdata.writeToExcelFile(extractdata.createBillDf(values['-BILL_PDF-']), values['-EXCEL_FILE-'])
            
        if extraction_complete == True:
            sg.Popup('Extraction completed successfully', keep_on_top = True)

        

    window.close()

drawMainWindow()
