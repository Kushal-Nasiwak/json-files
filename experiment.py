import pyautogui
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import os
import xlwings as xw
import pandas as pd
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', handlers=[logging.FileHandler('bot_log.log', 'a',), logging.StreamHandler()],encoding= 'utf-8')

class Excel_check:
    
    def condition(self,その他_sheet,野縁_sheet,builder_id):
        
        df_c = pd.read_excel('demo.xlsx',dtype=str)
        logging.info(f"Executing Conditions of Demo Excel for {builder_id} Builder")
        # logging.info(f"{df_c.info()}")
        no_flore = len([f for f in os.listdir(rf'Ankens/{self.bango}')])-3
        for b_id,floor,cellval,con in zip(df_c['Builder Code'],df_c['Floor value'],df_c['Location'],df_c['What to do']):
            if b_id == builder_id:
                cell_val = cellval.split(':')[-1].strip()
                sheet_name = cellval.split('Range:')[0].split(' ')[1].strip()
                if sheet_name == 'その他、':
                    sheet = その他_sheet
                elif sheet_name == '野縁、':
                    sheet = 野縁_sheet
                if cell_val.upper() == 'J16' and con.split()[0] == 'Add':
                    logging.info(f'value of J16 is : {sheet[cell_val].value}')
                    sheet[cell_val].value = sheet[cell_val].value + int(con.split()[1])
                elif 'first free cell available'.upper() in cell_val.upper() and con.split()[0] == 'Put':
                    if floor == 'All floors':
                        column = cell_val.split()[-2]
                        # Get the last cell with data in the column
                        last_cell = sheet.range(column + str(sheet.cells.last_cell.row)).end('up').row
                        # logging.info(f'{column}{last_cell}')
                        sheet[f'{column}{last_cell+1}'].value = int(con.split()[1].strip("'"))
                        logging.info(f"{con.split()[1].strip("'")} Added to {column}{last_cell+1}")
                        
                    elif floor.split()[0] == 'ONLY' and f"{floor.split()[-1]}.xls" in str(sheet):
                        column = cell_val.split()[-2]
                        
                        # Get the last cell with data in the column
                        last_cell = sheet.range(column + str(sheet.cells.last_cell.row)).end('up').row
                        if column == 'M':
                            last_cell = sheet.range('Q' + str(sheet.cells.last_cell.row)).end('up').row
                        # logging.info(f'{column}{last_cell}')
                        
                        sheet[f'{column}{last_cell+1}'].value = con.split()[1].strip("'")
                        logging.info(f"{con.split()[1].strip("'")} Added to {column}{last_cell+1}")
                        
                elif cell_val.upper() == 'J19' :
                    if con.split()[0] == 'Shoumeifukku' and f"{floor.split()[-1]}.xls" in str(sheet):
                        val = con.split()[2].strip("'")
                        logging.info(f'{val} added to J19')
                        sheet['J19'].value =  val
                        
                    elif con.split()[0] == 'Overwrite' and floor == 'All floors':
                        val = con.split()[2].strip("'")
                        logging.info(f'{val} added to J19')
                        sheet['J19'].value =  val
                    
                elif cell_val.upper() == 'J20':
                    if no_flore <= 2 and floor == '1.2F case':
                        sheet['J20'].value =  3
                        logging.info('3 added to J20')
                    elif no_flore >= 3  and floor == '1.2.3F case':
                        sheet['J20'].value =  1
                        logging.info('1 added to J20')
                    elif floor == 'All floors':
                        sheet['J20'].value =  1
                        logging.info('1 added to J20')
                        
                elif cell_val.upper() == 'J12':
                    sheet['J12'].value = 2
                    logging.info('2 added to J12')
                    
                elif cell_val.upper() == 'C21':
                    if no_flore <= 2 and floor == '1.2F case':
                        sheet['C21'].value = '吊元セット【L＝250】'
                        logging.info('吊元セット【L＝250】 added to C21')
                    elif no_flore >= 3  and floor == '1.2.3F case':
                        sheet['C21'].value = '吊元セット【L＝300】'
                        logging.info('吊元セット【L＝300】 added to C21')
                elif cell_val.upper() == 'C22':
                    sheet['C22'].value = '吊元セット【L＝150】'
                    logging.info('吊元セット【L＝150】 added to C22')
    

            
            
        
    def __init__(self,Anken_Bango,Builder_id):
        # try:
            logging.info(f"Starting Excel Check")
            logging.info(f"Anken Bango: {Anken_Bango} Builder ID: {Builder_id.strip(' \n')}")
            url = 'https://webaccess.nsk-cad.com/'
            options = webdriver.ChromeOptions()
            self.driver = webdriver.Chrome(options=options)
            self.bango = Anken_Bango
            self.builder_id = Builder_id
            self.driver.maximize_window()
            time.sleep(2)
            self.driver.get(url)

            logid = self.driver.find_element("name","u")
            logpassword = self.driver.find_element("name","p")

            logid.clear()
            time.sleep(1)
            logpassword.clear()

            logid.send_keys("0618")
            time.sleep(1)
            logpassword.send_keys("0618")
            time.sleep(1)

            logid.submit()
            self.driver.implicitly_wait(10)
            logging.info('Logged in to Webaccess')
            if not os.path.exists('Ankens/'+Anken_Bango):
                self.data_fetching(Anken_Bango,Builder_id)
                # logging.info(f"{Anken_Bango} folder Created")
            else:
                self.Display(f"{Anken_Bango} Folder present in Ankens")
                raise Exception("Exists")
            self.close_web()
            
        # except Exception as E:
        #     logging.info(f"Error: {E}")
        # workbook.save('demo.xlsx')

    def data_fetching(self,id,builder_id):
        self.driver.find_element("xpath",'//*[@id="f-menus"]/ul/li[4]/a').click()
        time.sleep(1)
        self.driver.find_element("xpath",'//*[@id="f-search-box"]/div/button[2]').click()
        time.sleep(1)
        self.driver.execute_script("scroll(0, 0);")
        self.driver.find_element(By.XPATH,'/html/body/div[2]/div[2]/div[2]/form/div/table[5]/tbody/tr[1]/td[3]/input[1]').clear()
        time.sleep(1)
        bango = self.driver.find_element("xpath",'//*[@id="f-search-box"]/table[3]/tbody/tr/td[1]/input')
        bango.send_keys(id)
        time.sleep(1)

        self.driver.find_element('xpath','//*[contains(text(), "検索")]').click()
        time.sleep(1)

        Cust_name=self.driver.find_element(By.XPATH, '//*[@id="orderlist_wrapper"]/div/div[3]/div[2]/div/table/tbody/tr/td[9]')
        obj_name=self.driver.find_element(By.XPATH, '//*[@id="orderlist_wrapper"]/div/div[3]/div[2]/div/table/tbody/tr/td[10]')
        delivery_date=self.driver.find_element(By.XPATH, '//*[@id="orderlist_wrapper"]/div/div[3]/div[2]/div/table/tbody/tr/td[12]')
        address=self.driver.find_element(By.XPATH,'//*[@id="orderlist"]/tbody/tr/td[21]')
        logging.info(f'Customer Name : {Cust_name.text}')
        logging.info(f'Object Name : {obj_name.text}')
        logging.info(f'Delivery Date : {delivery_date.text}')
        logging.info(f'Address : {address.text}')
                
        self.Builder_name = Cust_name.text
        time.sleep(1)
        self.Address = address.text
        time.sleep(1)

        data_list=[Cust_name.text,obj_name.text,delivery_date.text,address.text]

        self.driver.find_element(By.XPATH,'//*[@id="orderlist_wrapper"]/div/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/input').click()
        time.sleep(1)
        url=self.driver.find_element(By.XPATH,'//*[@id="f-search-box"]/table[10]/tbody/tr/td/input').get_attribute('value')
        self.目地=self.driver.find_element(By.NAME,'project_articles[0][project_orders][0][project_products][meji][product_num]').get_attribute('value')
        if self.目地=="": self.目地='0'
        self.入隅=self.driver.find_element(By.NAME,'project_articles[0][project_orders][0][project_products][irisumi][product_num]').get_attribute('value')
        if self.入隅=="": self.入隅='0'
        
        try:
            quantity=self.driver.find_element(By.XPATH,'//*[@id="deliver_225162_0"]/table[3]/tbody/tr[3]/td[2]/input').get_attribute('value')
        except :
            quantity = 0
        data_list.extend([url,quantity])
        self.Display('Fetched the Above details from Web Access')
        self.extract_files(id,url,builder_id)
        return data_list


    def close_web(self):
        self.driver.close()


    
# class 2
    def extract_files(self,id,url,builder_id):
        self.id = id
        self.handle_login()  # Handles the login process    
        self.driver.get(url)
        logging.info(f"Sharepoint link: {url}")
        time.sleep(5)
        self.Download_folder(str(id))
        time.sleep(2)
        self.builder_id_drive(builder_id)
    
    def handle_login(self):
        url='https://nskkogyo.sharepoint.com/sites/2021'
        # Assuming the login page has input fields with IDs 'username' and 'password'
        self.driver.get(url)

        username = "kushalnasiwak@nskkogyo.onmicrosoft.com"
        password = "Vay32135"
        time.sleep(2.5)
        # Find the username input field on the login page
        self.driver.find_element(By.XPATH, '//*[@id="i0116"]').clear()
        time.sleep(1.5)

        self.driver.find_element(By.XPATH, '//*[@id="i0116"]').send_keys(username)
        time.sleep(1.5)
        self.driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
        time.sleep(1.5)

        # Find the password input field on the login page
        if self.driver.find_element(By.XPATH, '//*[@id="i0118"]').text:
            self.driver.find_element(By.XPATH, '//*[@id="i0118"]').clear()
            time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="i0118"]').send_keys(password)
        time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
        time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="KmsiCheckboxField"]').click()
        time.sleep(1.5)
        self.driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
        time.sleep(2)
        self.Display('Logged in to Sharepoint\nPlease wait.....')


    # functions to download the files from the share point, and store it to the respected Ankens Bango
    def Download_folder(self,id):
        if not os.path.exists('Ankens/'+id):
            os.makedirs('Ankens/'+id)
            
        else:
            logging.info("Folder Exists")
       
        # Changing the downloading directory for every Anken bango
        # self.driver.execute_cdp_cmd('Page.setDownloadBehavior',{'behavior':'allow','downloadPath':rf'C:\Users\kusha\Working\Day1\Ankens\{id}'})
        self.driver.execute_cdp_cmd('Page.setDownloadBehavior',{'behavior':'allow','downloadPath':rf'{os.getcwd()}\Ankens\{id}'})
        # C:\Users\kusha\Working\Complete_Project\Ankens

        # finding the folder with pdf and excel files
        folder = self.driver.find_elements(By.XPATH,'//span/button')
        for btn in folder:
            if btn.text in ['割付図・エクセル','割付図・ エクセル']:
                btn.click()
                time.sleep(1)
                break
        else:
            self.Display('割付図・エクセル folder not found' )
            self.delete_folder()
            self.close_web()
            raise Exception('Error')
        time.sleep(5)
        file  = self.driver.find_elements(By.XPATH,"//button")

        # finding the pdf and excel files to download
        for i in  file:
            try:
                if i.text.endswith(('.pdf','階.xls','平屋.xls')):
                    # mouse action to download the respective file
                    actions = ActionChains(self.driver)
                    actions.move_to_element(i)
                    actions.context_click(i)
                    actions.perform()
                    time.sleep(2)
                    # to find the downoad button and select it
                    download_button=self.driver.find_element(By.NAME,'Download')
                    download_button.click()
                    time.sleep(3)
            except Exception as E:
                # logging.info(f"Error: {E}")
                continue
                
        time.sleep(5)
    

# class 3

    def builder_id_drive(self,builder_id):
        drive_url = 'https://nskkogyo.sharepoint.com/sites/DQ/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FDQ%2FShared%20Documents%2F%E3%82%A4%E3%83%B3%E3%83%89%E4%BA%8B%E5%8B%99%E6%89%80%E9%96%A2%E4%BF%82%2FExcel%20Check%20App%20Database&p=true&ga=1'
        self.driver.get(drive_url)
        # logging.info('Got the page')
        time.sleep(2)
        self.driver.implicitly_wait(100)
        self.driver.find_element(By.XPATH,'//*[@id="sbcId"]/form/input').send_keys(builder_id)
        pyautogui.hotkey('enter')
        time.sleep(3)

        
        # self.driver.find_element(By.CLASS_NAME,'submitSearchButton-220').click()
        
        # self.driver.find_element(By.XPATH,'//*[@id="sbcId"]/form/span[6]/button').click()
        self.driver.implicitly_wait(100)
        time.sleep(5)
        files = self.driver.find_elements(By.XPATH,'//button')
        if not any([i.text.startswith(builder_id) for i in files]):
            drive_url = 'https://nskkogyo.sharepoint.com/sites/DQ/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FDQ%2FShared%20Documents%2F%E3%82%A4%E3%83%B3%E3%83%89%E4%BA%8B%E5%8B%99%E6%89%80%E9%96%A2%E4%BF%82%2FExcel%20Check%20App%20Database&p=true&ga=1'
            self.driver.get(drive_url)
            # logging.info('Got the page')
            time.sleep(3)
            self.driver.implicitly_wait(100)
            self.driver.find_element(By.XPATH,'//*[@id="sbcId"]/form/input').send_keys(builder_id)
            pyautogui.hotkey('enter')
            time.sleep(3)
            self.driver.implicitly_wait(100)
            time.sleep(5)
            files = self.driver.find_elements(By.XPATH,'//button')
        for i in files:
            try:
                if i.text.startswith(builder_id):
                    logging.info(f"{i.text}")
                    
                    self.driver.implicitly_wait(30)
                    # mouse action to download the respective file
                    actions = ActionChains(self.driver)
                    actions.move_to_element(i)
                    actions.context_click(i)
                    actions.perform()
                    self.driver.implicitly_wait(10)
                    time.sleep(3)
                    # to find the downoad button amd select it
                    download_button=self.driver.find_element(By.NAME,'Download')
                    download_button.click()
                    self.driver.implicitly_wait(20)
                    # time.sleep(3)
                    self.Display(f'{i.text} file Downloaded')
                    break
            except Exception as E:
                logging.info(f"Error: {E}")
                continue
        else:
            self.Display("Builder ID Invalid: Check builder Id and Try Again")
            self.delete_folder()
            self.close_web()
            raise Exception('Error')
        # time.sleep(5)
        self.builder_copy(self.id,i.text)
        
    def builder_copy(self,id,builder_id):
        self.driver.minimize_window()

        # directory='Ankens/{id}'
        # self.builder_id = builder_id

        def repeat_elements(lengths,room_num, repetitions):
            list_after_count = []
            for length,rno ,repeat_count in zip(lengths,room_num,repetitions):
                if pd.notna(repeat_count):
                    repeat_count = int(repeat_count)  # Convert repeat_count to an integer
                    list_after_count.extend([(length,rno)] * repeat_count)
                else:
                    list_after_count.extend([(length,rno)])
            return list_after_count
        
        

        # Specify the file and macro
        # file_path = r"C:\path\to\your\file.xlsm"
        # macro_name = "YourMacroName"  # Use the exact macro name

        # run_excel_macro(file_path, macro_name)

            
    
        def copy_pasting(file_name,builder_name,address):
            time.sleep(5)
            # logging.info(builder_name,address)
            app=xw.App(visible=False)
            
            
            sorce_wb = app.books.open(rf'Ankens/{id}/{builder_id}')
            sorce_sheet = sorce_wb.sheets[str(len(excel_files))]
            

            destination_wb = app.books.open(rf'Ankens/{id}/{file_name}')
            # macro_click=destination_wb.macro('Macro2')
            # macro_click()
            destination_sheet = destination_wb.sheets['その他']


            sorce_range = 'A1:H19'
            destination_start_cell = 'B5'

            sorce_range_copy = sorce_sheet.range(sorce_range).value

            destination_sheet.range(destination_start_cell).value =sorce_range_copy
            if file_name.endswith('1階.xls'):
                destination_sheet['J9'].value = self.目地
                destination_sheet['J8'].value = self.入隅
            logging.info(f"{self.目地},{self.入隅}")   
            

            # # stacking
            stock_sheet = destination_wb.sheets['野縁']
            pasting_sheet = destination_wb.sheets['製作用データ']
            page1 = destination_wb.sheets['提出書 ']
            name=stock_sheet['E3'].value
            fno=name.split()[1][0]
            stock_sheet_range = 'E11:G56'
            stock_copy = stock_sheet.range(stock_sheet_range).value
            df = pd.DataFrame(stock_copy,columns=stock_sheet.range('E10:G10').value)
            df.dropna(inplace=True)

            lengths = repeat_elements(df['野縁寸法'], df['部屋番号'],df['本数'])
            lengths_df = pd.DataFrame(lengths,columns=['野縁寸法 length','部屋番号 Room Number'])
            # result_info=pd.DataFrame(stack_material(lengths_df),columns=['材長','部屋番号','部屋別本数','stack id'])
            # result_info.insert(2,'階数',value=[fno for i in range(len(result_info))])
            # page1['G33'].value = max(result_info['stack id'])
            pasting_sheet['A1'].value=name
            # pasting_sheet.range('A2').options(index=False).value =result_info
            stock_sheet['AE3'].value = address
            stock_sheet['AE5'].value = builder_name
            try:
                self.condition(destination_sheet,stock_sheet,self.builder_id)
            except Exception as E:
                logging.info(f"{E}")
                destination_wb.close()
                sorce_wb.close()
                os.remove(rf'Ankens/{id}/{builder_id}')
                raise Exception(f'{E}')
            def macro2_recreate():
                    workbook = destination_wb
                    sheet1 = workbook.sheets['野縁']
                    source_range = sheet1.range("Q11:R110")
                    target_range = sheet1.range("U11")
                    target_range.value = source_range.value 
                    data_range = sheet1.range("U11:U110")
                    data = data_range.value 
                    data = sorted(data, reverse=True)  
                    data_range.value = [[item] for item in data]  
                    sheet1.range("A1").select()
                    sheet2 = workbook.sheets["製作用データ"]
                    sheet2.range("A1").clear_contents()
                    sheet2.range("A3:D73").clear_contents()
                    c = int(sheet1.range((111, 41)).value)
                    dynamic_range = sheet1.range((11, 30), (11 + c - 1, 33))
                    sheet2.range("A3").value = dynamic_range.value
                    e3_value = sheet1.range("E3").value
                    sheet2.range("A1").value = e3_value
                    sheet1.range("A1").select()
                    workbook.save()
                    logging.info("Macro2 run Successfully ")
                    
            def run_macro():
                    workbook = destination_wb
                    # Work with the active sheet (or specify the sheet name if needed)
                    sheet = workbook.sheets.active

                    # Step 1: Define the range to sort
                    sort_range = sheet.range("Q4:T32")
                    # sort_key = sheet.range("Q5:Q32")  # Sorting key range (Q5 column)

                    # Read the data
                    data = sort_range.value
                    print(data)

                    # Sort data based on the first column (column Q, starting from row 5)
                    header = data[0] if isinstance(data[0], list) else []  # Handle header if present
                    data_to_sort = data[1:] if header else data  # Exclude header if detected

                    sorted_data = sorted(data_to_sort, key=lambda x: x[0], reverse=True)  # Sort by column Q descending

                    # Reassign sorted data back, including the header if present
                    if header:
                        sheet.range("Q4:T28").value = [header] + sorted_data
                    else:
                        sheet.range("Q4:T28").value = sorted_data

                    # Step 2: Select U4
                    sheet.range("U4").select()

                    # Save and close the workbook
                    workbook.save()
                    logging.info("Ran macro executed successfully")
            try:
                macro2_recreate()
                time.sleep(2)
                run_macro()
                time.sleep(2)
            except Exception as E:
                logging.info(f'{E}')

            destination_wb.save(rf'Ankens/{id}/★{file_name}')
            destination_wb.close()
            sorce_wb.close()
            time.sleep(2)
            os.remove(rf'Ankens/{id}/{file_name}')
            # logging.info('Task Completed')
        
        excel_files = [f for f in os.listdir(rf'Ankens/{id}') if f.endswith('階.xls') or f.endswith('屋.xls')]
        for i in excel_files:
            self.Display(f'Copying the Builder Model to {i} file')
            copy_pasting(i,self.Builder_name,self.Address)
        try:
            os.remove(rf'Ankens/{id}/{builder_id}')
        except:
            logging.info('FileNotFoundError')
        self.Display(f"Excel Check Completed Successfully\n\n")  
        

    def Display(self,info):
        logging.info(f"{info}")

    def delete_folder(self):
        files = [f for f in os.listdir(rf'Ankens/{self.bango}')]
        for file in files:
            os.remove(rf'Ankens/{self.bango}/{file}')

        if os.path.exists('Ankens/'+self.bango):
            os.rmdir(rf'Ankens/{self.bango}')


# if __name__ == '__main__':
#     # Excel_check('390925','007000')
#     # Excel_check('389966','001706')
#     # Excel_check('403758','008001')
#     # Excel_check('428930','060301')
#     # Excel_check('443483','007000')
#     # Excel_check('445069','001705')
#     # Excel_check('447048','027902')
#     # Excel_check('441126','095200')
#     Excel_check('449730','051600')
    
# pyinstaller --onefile --windowed  --add-data="gui:gui"  --icon=appicon.ico Zenbubot.py

Ankens_bango_builder = {}
# Function to clear placeholder text when user focuses on the entry widget
def clear_placeholder(event, entry, placeholder):
    if entry.get() == placeholder:
        entry.delete(0, tk.END)
        entry.config(foreground="black")


# Function to restore placeholder text if entry is empty
def restore_placeholder(event, entry, placeholder):
    if entry.get() == "":
        entry.insert(0, placeholder)
        entry.config(foreground="gray")


# Function to load and display values in a structured format in the text area, then clear the entries
def load_values():
    anken_bango = anken_entry.get()
    builder_id = builder_id_entry.get()

    # Check if placeholders are still present; if so, replace them with an empty string
    if anken_bango == anken_placeholder:
        anken_bango = ""
    if builder_id == builder_placeholder:
        builder_id = ""

    # Append headers if they aren't already present
    if text_area.compare("end-1c", "==", "1.0"):
        text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30}\n")
        text_area.insert(tk.END, "-" * 80 + "\n")  # Divider line

    # Append the entered values in a formatted way
    if anken_bango and builder_id:
        Ankens_bango_builder[str(anken_bango)] = str(builder_id)
        text_area.insert(tk.END, f"{anken_bango:^30} {builder_id:^30}\n")
    

    # Clear the entry fields and reset placeholders
    anken_entry.delete(0, tk.END)
    builder_id_entry.delete(0, tk.END)
    restore_placeholder(None, anken_entry, anken_placeholder)
    restore_placeholder(None, builder_id_entry, builder_placeholder)


# Function to browse and upload an Excel file
def upload_excel():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm"), ("All Files", "*.*")]
    )
    if file_path:
        excel_file_entry.delete(0, tk.END)
        excel_file_entry.insert(0, file_path)
        read_excel_file(file_path)


# Function to read and display Anken Name and Builder ID columns from the Excel file
def read_excel_file(file_path):
    try:
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(file_path,dtype=str)
        # Check if the required columns exist
        if "Anken Number" in df.columns and "Builder Code" in df.columns:
            anken_data = df["Anken Number"]
            builder_data = df["Builder Code"]

            # Display headers if not already present
            if text_area.compare("end-1c", "==", "1.0"):
                text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30}\n")
                text_area.insert(tk.END, "-" * 80 + "\n")  # Divider line

            # Display the data row by row
            for anken, builder in zip(anken_data, builder_data):
                Ankens_bango_builder[anken] = builder
                text_area.insert(tk.END, f"{anken:^30} {builder:^30}\n")

        else:
            messagebox.showerror("Error", "The file does not contain 'Anken Name' and 'Builder ID' columns.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the Excel file:\n{str(e)}")

# Function to start the Excel check
def Start_Excel_check():
    print("Executing the Program")
    global Ankens_bango_builder
    print("Anken Dictionary:")
    print(Ankens_bango_builder)
    text_area.delete(1.0, tk.END)  # Clear the text area
    text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30} {'Status':^20}\n")
    text_area.insert(tk.END, "-" * 80 + "\n")  # Divider line
    for Ankens,Builder in Ankens_bango_builder.items():
        try:
            Excel_check(str(Ankens),str(Builder))
        except Exception as e:
            # If an error occurs, display with a red cross
            text_area.insert(tk.END, f"{Ankens:^30} {Builder:^30} {'❌':^20}\n")
            print(f"error:{e}")
        else:
            text_area.insert(tk.END, f"{Ankens:^30} {Builder:^30} {'✅':^20}\n")
    Ankens_bango_builder.clear()
    
   

# Initialize the main window
root = tk.Tk()
root.title("Excel Check")
root.geometry("1000x700")  # Increased height for extra area
root.configure(bg="#d6e0f0")

# Logo placeholder
logo_frame = tk.Frame(root, bg="#d6e0f0", height=50)
logo_frame.pack(fill='x', padx=10, pady=5)
logo_label = tk.Label(logo_frame, text="Nasiwak", font=("Arial", 12, "bold"), bg="#d6e0f0")
logo_label.pack(side='left')

# Title
title_label = tk.Label(root, text="Excel Check", font=("Arial", 14, "bold"), bg="#d6e0f0", fg="#333333")
title_label.pack(pady=(10, 5))

# Instruction label
instruction_label = tk.Label(
    root, text="Enter Anken Bango and the Builder Id, Click on the START button to start Excel Check.",
    font=("Arial", 10), bg="#d6e0f0", fg="#555555"
)
instruction_label.pack()

# Entry Frame
entry_frame = tk.Frame(root, bg="#d6e0f0")
entry_frame.pack(pady=(10, 10))

# Placeholder texts
anken_placeholder = "Enter Anken Bango..."
builder_placeholder = "Enter Builder Id..."

# Anken Bango entry with placeholder
anken_entry = ttk.Entry(entry_frame, width=30, foreground="gray")
anken_entry.insert(0, anken_placeholder)
anken_entry.grid(row=0, column=0, padx=10, pady=5)

anken_entry.bind("<FocusIn>", lambda event: clear_placeholder(event, anken_entry, anken_placeholder))
anken_entry.bind("<FocusOut>", lambda event: restore_placeholder(event, anken_entry, anken_placeholder))

# Builder Id entry with placeholder
builder_id_entry = ttk.Entry(entry_frame, width=30, foreground="gray")
builder_id_entry.insert(0, builder_placeholder)
builder_id_entry.grid(row=0, column=1, padx=10, pady=5)

builder_id_entry.bind("<FocusIn>", lambda event: clear_placeholder(event, builder_id_entry, builder_placeholder))
builder_id_entry.bind("<FocusOut>", lambda event: restore_placeholder(event, builder_id_entry, builder_placeholder))

# Load button
load_button = ttk.Button(entry_frame, text="LOAD", command=load_values)
load_button.grid(row=0, column=2, padx=10, pady=5)

# Excel file upload frame
excel_frame = tk.Frame(root, bg="#d6e0f0")
excel_frame.pack(pady=(10, 10))

# Upload Excel button
upload_button = ttk.Button(excel_frame, text="Upload Excel File", command=upload_excel)
upload_button.grid(row=0, column=0, padx=10, pady=5)

# Entry box to show the uploaded file path
excel_file_entry = ttk.Entry(excel_frame, width=50, state="normal")
excel_file_entry.grid(row=0, column=1, padx=10, pady=5)

# Main text area for displaying entries in a structured way
text_area = tk.Text(root, height=15, width=80, bg="#e6eefc", font=("Courier", 10))
text_area.pack(pady=10)

# Start button
start_button = ttk.Button(root, text="START", width=10,command=Start_Excel_check)
start_button.pack(pady=10)

# Footer
footer_label = tk.Label(root, text="Nasiwak Services Pvt Ltd     v9.0.0", bg="#d6e0f0", fg="#333333")
footer_label.pack(side='bottom', pady=(5, 0))

# Run the main event loop
root.mainloop()
