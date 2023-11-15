import time
import pandas as pd
from mail import Mail
from automation import Automation
import sys
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


class Matchdata:
    selected_files = eval(sys.argv[1])
    file_path = selected_files['Main file']
    df1= pd.read_excel(file_path)
    main_columns = ['Project',"Long description", "Quantity"]
    main_data = df1[main_columns]
    try:
        file_path = selected_files['Accessories file']
        df=pd.read_csv(file_path)
        columns = ['item_name', 'Quantity']
        accesories_data = df.loc[:, columns]
    except FileNotFoundError as e:
        print("No File provided")
        accesories_data=None
   

    file_path = selected_files['Yarn file']
    df2=pd.read_csv(file_path)
    yarn_columns=['invoice_number','Item Description','Weight']
    yarn_data=df2[yarn_columns]
    yarn_data=yarn_data.iloc[0]
    match_found=False
    def preprocess_value(value):
        return int(value.replace(',', ''))

    for _,values in main_data.iterrows():
        if(values['Project']==yarn_data['invoice_number']):
            values['Long description']=values['Long description'].replace(" ","")
            yarn_data['Item Description']=yarn_data['Item Description'].replace(" ","")

            if(values['Long description']==yarn_data['Item Description']):
                yarn_data['Weight']=preprocess_value(yarn_data['Weight'])
                if(values['Quantity']==yarn_data['Weight']):
                    match_found=True
                    

    @staticmethod
    def accesories():
       
        
        if Matchdata.accesories_data==None:
            return None 
        Matchdata.accesories_data['Size'] = \
        Matchdata.accesories_data['item_name'].str.extract(r'(\d+\.\d+X\d+\.\d+X\d+\.\d+) (\w+)')[0]
        Matchdata.accesories_data.loc[:, 'item_name'] = Matchdata.accesories_data.loc[:, 'item_name'].str.replace(
            r'(\d+\.\d+X\d+\.\d+X\d+\.\d+) ', '')
        Matchdata.accesories_data.loc[:, 'item_name'] = Matchdata.accesories_data.loc[:, 'item_name'].str.strip()
        Matchdata.accesories_data.loc[:, 'item_name'] = Matchdata.accesories_data.loc[:, 'item_name'].str.replace(
            'CARTONBOX', 'CBOX')
        Matchdata.accesories_data.loc[:, 'All Substrings'] = Matchdata.accesories_data['item_name'].apply(
            Matchdata.all_substrings)
        matching_row_numbers = []
        for idx, row in enumerate(Matchdata.main_data['Code 2']):
            if isinstance(row, str):
                match_found = False

                for substrings in Matchdata.accesories_data['All Substrings']:
                    if row in substrings:
                        match_found = True
                        break
                if match_found:
                    matching_row_numbers.append(idx)
        matching_main_data = Matchdata.main_data.loc[matching_row_numbers, :]
        subs = matching_main_data['Long description'].apply(Matchdata.all_substrings)
        matching_row_numbers = []
        for row in Matchdata.accesories_data['Size']:
            for idx, substrings in enumerate(subs):
                if row in substrings:
                    matching_row_numbers.append(idx)
                    break
        matched_data = matching_main_data.loc[matching_row_numbers, :]
        is_quantity_present = Matchdata.accesories_data['Quantity'].isin(matched_data['Quantity'])
        counter = 0
        for i in is_quantity_present:
            if not i:
                counter += 1
        return counter

    # @staticmethod
    # def yarn():
        is_item_present = Matchdata.main_data['IT'].str.contains(Matchdata.split_yarn_values[1])
        is_item_present = list(is_item_present)
        print(is_item_present)
        try:
            index_of_true = is_item_present.index(True)
            if type(index_of_true) == int:
                index_of_true = [index_of_true]
        except ValueError as e:
            index_of_true = None
        if index_of_true==None:
            print("Item not found")
            return False

        matching_item = Matchdata.main_data.loc[index_of_true, :]
        all_codes = False
        if index_of_true is not None:
            for index in index_of_true:
                for i in range(1, 10):
                    if i == 10:
                        all_codes = True
                    try:
                        is_present = Matchdata.yarn_values[i + 1] in Matchdata.main_data.loc[index, f'Code {i}']
                    except KeyError as e:
                        print("breaking")
                        break
                    except TypeError as e:
                        print("all codes do not match")
                        return False
            is_present_2 = False
            if all_codes:
                concatenated_string = ""

                # Start iterating from index 10
                index = 11

                while index < len(Matchdata.split_yarn_values) and not isinstance(Matchdata.split_yarn_values[index],
                                                                                  float):
                    concatenated_string += str(Matchdata.split_yarn_values[index])
                    index += 1
                is_present_2 = concatenated_string in matching_item['Long description']
            final = False
            if is_present_2:
                final = Matchdata.split_yarn_values[-1] in matching_item['Quantity']
            return final


if __name__ == "__main__":
    # mail_sent = False
    # unmatched_accesoories = Matchdata.accesories()
    # if unmatched_accesoories==None:
    #     print("No accesories file provided")

    # elif unmatched_accesoories != 0:
    #     Mail.send_mail(f"Accessories data does not match for {unmatched_accesoories} rows")
    #     mail_sent = True
    # time.sleep(5)
    if Matchdata.match_found:
        print("Running automation")
        Automation.web_automation_jsp(url="http://192.168.0.237:8080/now/MainDesktop.jsp", user="emdx", psw="emdx")
    else:
        Mail.send_mail("Yarn data does not match")
        mail_sent = True

