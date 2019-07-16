import xlrd
import datetime
import os

class NoDataError(Exception):
    pass

class Quote_to_Policy:

    def __init__(self, filename):
        self.filename = filename
        
    
    
    #Arrange short keys in order of prefernce if conflict occurs
    #or append more (key, expected value type) tuples in the list
    #Make sure that the keys appended do not appear in the values.
    short_keys = [('name', str), ('quoteno', str),('address', str),('code', str),
                  ('code', float),('code', int),('quote', float),
                  ('rate', float), ('no', int), ('holder', str),('policyno', str),
                  ('quoteno', int),('quote', str), ('quote', int),('risk', str),
                  ('write' ,str),('total', float),('qms', float),
                  ('qms', str),
                  ('section', int), ('section', str), ('indust', str), ('indust', int),
                  ('comm', str), ('loc', str), ('holder', str),
                  ('client', str), ('relation', str), ('group', int),('group', str), 
                  ('prem', int), ('comm', float), ('mail', str), ('busi', str)]
    
    
    final_data = {}  
    discard_data = {} #Data which has key value pair but is rejected
    policies = {}   
    tnc = {}
        
    def get_sheetname(self, filename):
        
        try:
            wb = xlrd.open_workbook(filename)
        
        except FileNotFoundError:
            raise FileNotFoundError
        
        sheets = wb.sheet_names()
        sheetname = ''
        for name in sheets:
            temp = name.lower().replace(' ','')
            if 'quote'in temp or 'quotation' in temp:
                sheetname = name
        
        return(sheetname)


    def get_workbook_data(self, filename, sheetname):
        index = 0
        data = []
        #True if file is in .xlsx format False for .xls format
        xlsx_file = True
        
        try:
            try:
                #Open Workbook in .xls mode
                wb = xlrd.open_workbook(self.filename, formatting_info = True)
                xlsx_file = False
            
            except NotImplementedError:
                #Open Workbook in .xlsx mode
                wb = xlrd.open_workbook(self.filename)
            
            try:
                sheet = wb.sheet_by_name(sheetname)
            except xlrd.XLRDError:
                return(xlrd.XLRDError)
            
            #get list of cell indices and the range of merge cells
            #[(rows, cols), merge_length]
            merged_range_list = self.count_merge_range(sheet, self.filename)
            rows = 0
            
            #Read each cell value and copy to raw_data
            # [data, row_no, col_no, raw_data_index, merge_length]
            
            while rows < sheet.nrows:
                
                #Skip the hidden rows
                if not xlsx_file:
                    try:
                        hidden = sheet.rowinfo_map[rows].hidden
                        if hidden == True:
                            rows += 1
                            continue
                    except KeyError:
                        rows += 1
                        continue
                                
                cols = 0
                while cols < sheet.ncols:
                    
                    #Check merged cell range
                    i = (rows, cols)     #i is a temp variable
                    merge_range = 0
                    for j in merged_range_list:
                        if i == j[0]:
                            merge_range = j[1]
                            break
                    
                    temp = sheet.cell_value(rows, cols)
                    data.append((temp, rows, cols, index, merge_range))
                    index += 1                       
                    
                    if not merge_range == 0:
                        cols += merge_range
                    else:
                        cols += 1
                
                rows += 1
            
            if data == []:
                raise NoDataError
                
            return(data, sheet, wb)
        
        except FileNotFoundError:   #File name is incorrect or location incorrect
            raise( FileNotFoundError)
    
    def modify_data(self, data):
        
        modify_data = []
        for word in data:
            temp = word[0]
            
            #Remove Spaces, convert to lower case and strip the spaces at the end or beginning of the word
            if not(type(temp) == int or type(temp) == float):
                temp = temp.replace(' ', '').lower().strip()
            
            modify_data.append((temp, word[3], word[4]))
            
        return(modify_data)
    
    
    def get_data(self, wb, sheet, raw_data, mod_data, short_keys, final_data, discard_data):
        skip = 0
        
        #Values and keys which are already found
        found_values, found_keys = [], []

        #Root words used to identify keywords and values
        keywords = list(set([i[0] for i in short_keys]))    
        
        for num in range(len(raw_data)):
    
            base_key =  ''   
            #base key is the root word that is considered to 
                #diffrentiate between key and value
            
            #key_flag = False means key is not found
            key, key_flag = mod_data[num][0], False
            actual_key = raw_data[num][0]
            
            #keys are only of type string so reject other datatype keys
            if key == '' or type(key) == float or type(key) == int:
                continue
            
            #Separate the sheet in key, value sheet and policy and terms and conditions sheet
            elif key =='suminsured' or key == 's.no':
                break

# If key and value are in same cell then they are separated by a separator

            #Consider the key only upto the index of the separator
#Separator = ':' or '-'           
            
            if  ':' in key:
                part = key.index(':')
                actual_part = actual_key.index(':')
            elif '-' in key:
                part = key.index('-')
                actual_part = actual_key.index('-')
                
            else:
                part = len(key)
                actual_part = len(key)
                
            x = key         #temporary variable x
            key = key[:part]
            
            #get the base key for key
            for keys in keywords:
                if keys in key:
                    #Key Found
                    key_flag = True
                    base_key = keys
            
            if key_flag:
                #List of Expected Values for a key
                #Values added to the list are...
                #value_list = [(row, col), (row, col+1), (row, col+2), (row+1, col)]
                value_list = []
                skip = 0
                value_list.append(x[part+1:])
                for i in range(num+1, num+3):
                    value_list.append(raw_data[i][0])
                    skip += raw_data[i][4]
                
                try:
                    temp = raw_data[num+sheet.ncols-skip][0]
                    if not type(temp) == str:
                        value_list.append(temp)
                    else:
                        value_list.append('')
                #Last row of the sheet and no value_list[4] = ''
                except IndexError:
                    value_list.append('')
                 
                value = ''
                #Value_flag if false when value is not found
                #When the value is found change the value_flag = True
                value_flag = False
                for val in value_list:
                    
                    if val == '':
                        continue
                    
                    #value_list is filled in order of expected probability of answers
                    #Check if base key and type of value selected matches
                    #if matched then that is the value else check the next value
                    elif (base_key, type(val)) in short_keys:
                        #Value Found
                        value = val
                        value_flag = True
                        value_index = value_list.index(val)
                        
                        temp = ''
                        if type(value) == str:
                            temp = value.lower().replace(' ','')
                                        
                            if  ':' in key:
                                part = key.index(':')
                            elif '-' in key:
                                part = key.index('-')
                            else:
                                part = len(value)
                            temp = value[:part]            
                            temp = temp.lower().replace(' ','')
                            for i in keywords:
                                if i in temp:
                                    value_flag = False
                                    
                                    #proper keys can be discarded so add them in discarded data and can be retrieved later
                                    if key not in discard_data:
                                        discard_data[key] = value
                                    value = ''
                                    
                        
                        if value_flag and temp not in found_keys and key not in found_values:
                            
                            if temp not in found_values:
                                found_values.append(temp)
                            if key not in found_keys:
                                found_keys.append(key)
                            
                            if value_index == 0:
                                
                                value = raw_data[num][0]
                                value = value[actual_part+1:]
                                
                            if key == 'quotevalidity':
                                
                                if not type(value) == datetime.datetime:
                                    try:
                                        #returns date object of a float value
                                        temp_2 = raw_data[num+1][0]
                                        if not type(temp_2) == float:
                                            raise xlrd.XLRDError
                                            
                                        date = datetime.datetime(*xlrd.xldate_as_tuple(temp_2, wb.datemode))
                                        date = str(date)
                                        part = date.index(' ')
                                        temp = date[:part+1]
                                    except ValueError:
                                        temp = ''
                                    except TypeError:
                                        temp = ''
                                    except xlrd.xldate.XLDateAmbiguous:
                                        temp = ''
                                    
                                    try:
                                        value += str(temp)
                                    except TypeError:
                                        pass
                                
                            
                            final_data[key] = value                        
                            break
        
        #Retrieve the proper data that was discarded
        #If value doesnt have separator and value is not present in keys 
        #then add the data to the final_data
        for keys in discard_data:
            value =  discard_data[keys]
            try:
                temp = value.lower().replace(' ','')
            
            except ValueError:
                temp = ''
                
            if temp not in found_keys and ':' not in value:
                final_data[keys] = discard_data[keys]
        
        if 'quotevalidity' in final_data:
            value = final_data['quotevalidity']
            try:
                date = datetime.datetime(*xlrd.xldate_as_tuple(value, wb.datemode))
                date = str(date)
                part = date.index(' ')
                temp = date[:part+1]
                final_data['quotevalidity'] = temp
            except xlrd.xldate.XLDateAmbiguous:
                pass
            except TypeError:
                pass
            except ValueError:
                pass
            
        
        return(final_data)
        
        
    def count_merge_range(self, sheet, filename):    
        
        l = sheet.merged_cells
        
        #Temporary List to be used as list of merged cells returns a list of tuple
        x = []
        merged_range = []
        
        for i in l:
            x.append(list(i))
        
        for i in x:
            i[1] -= 1
            r = i[3] - i[2]
            merged_range.append([(i[0], i[2]), r])        #row, col, range of merged cell
        
        return(merged_range)
    
    #Get data of policies
    def get_policy(self, raw_data, mod_data, sheet, final_data, policies, tnc):
        
        tnc_name = ''
        for num in range(len(mod_data)):
            
            name = ''
            data = mod_data[num][0]
            if not type(data) == str:
                continue
            index = mod_data[num][1]
            #Index points at suminsured data
                #index += sheet.ncols - mod_data[2] + 1
            header, policy = [], []
            
            #Sum insured is the keyword that is searched to get to the policy data
            if  data == 'suminsured' :
                index += sheet.ncols - mod_data[num][2] + 1
                i = num - 1
                
                while not raw_data[i][2] == 0:
                    header.append(raw_data[i][0])
                    i -= 1
                header.append(raw_data[i][0])
                header = header[::-1]
                
                
                row = raw_data[num][1]
                count = self.terminate_col_loop(sheet, row)
                
                while count>0:
                    count -= 1
                    header.append(raw_data[num][0])
                    num += 1
                                
                count = 0
                j = sheet.ncols
                while True:
                    
                    while j>0:
                        i = i - mod_data[i-1][2] - 1
                        j = j-mod_data[i-1][2]-1
                    name = raw_data[i-1][0]
                    tnc_name = name
                    count += 1
                    temp = name.lower().replace(' ','')
                    if 'policy' in temp or 'insurance' in temp or count == 3 or 'perils' in temp:
                        if count == 3:
                            name = ''
                            tnc_name = name
                        break
                
                index -= raw_data[index][2]
                start_index, end_index = index, len(header)
                           
                flag = True
                count_row = 0
                while True:            
                    count_row += 1
                    x = 0
                    for i in range(start_index, start_index+end_index):    
                        try:
                            policy.append([header[x], raw_data[i][0]])
                            x += 1
                        except IndexError:
                            flag = False
                        
                        
                    if flag == False:
                        break
                    
                    skip = start_index + sheet.ncols
                    start_index = skip
                    try:
                        if not self.terminate_row_loop(sheet, raw_data[start_index][1], len(header)):
                            break
                    except IndexError:
                        break
                
                num += (count_row+2)*sheet.ncols    #Skip info already read
                
                #Check for blank data and fill them with appropriate data
                policy = self.policy_data(policy, len(header))
                policies[name] = policy
                
            elif data == 'termsandconditions':             
                row = raw_data[num][1] + 1
                tnc[tnc_name] = self.get_tnc(sheet, raw_data, mod_data, num)   
         
        final_data['policy_info'] = policies
        final_data['tnc_info'] = tnc
        return(policies, tnc)  
    
    #Input: list of fields in the policy
    #       count is the length of the header list
    def policy_data(self, policy, count):
        
        index = 0
        count -= 1
        start = True
        
        for data in policy:
            
            if start and index > count :
                start = False
                    
        
            if  index > count:
                if data[1] == '':
                    temp = policy[index-count-1][1]
                    if type(temp) == str:
                        index += 1
                        continue
                    
                    policy[index][1] = temp

            
            elif (not data[1] == '' and type(data[1]) == str )  or start:
                index += 1
                continue
            
            index += 1
            
        return(policy)
    
    #Count number of rows to be read
    def terminate_row_loop(self, sheet, row, col_length):
        row += 1
        blank = 0
        for i in range(col_length):
            if sheet.cell_value(row, i) == '':
                blank += 1
        
        if blank > col_length//2:
            return(False)
        else:
            return(True)
            
    #Count number of cols to be read
    def terminate_col_loop(self, sheet, row):
        
        #possible words that can be present in the header row
        possible_headers = ['no','sum','insur','gst','pay','net','occ','rate',
                            'prem','descrip','loss']
        count = 0
        col = 0
        flag = True
        while True:
            try:
                temp = sheet.cell_value(row, col).lower().replace(' ','')
            except AttributeError:
                return(col)
            flag = True
            if not temp == '':
                for i in possible_headers:
                    if i in temp:
                        flag = False
                        col += 1
                        count += 1
                        break
                if flag:
                    break
                
            elif temp == '':
                temp = sheet.cell_value(row, col+1).lower().replace(' ','')
                if temp == '':
                    break
                for i in possible_headers:
                    if i in temp:
                        flag = False
                        col += 2
                        count += 2
                        break
                if flag:
                    break
            else:
                break
        return(count-2)
    
    
    #pass the index number of the data that has field tnc
    def get_tnc(self, sheet, raw_data, mod_data, index):            
        t = []  #Temporary variable to store the TNC for a particular policy
        
        flag = True
    
        count = 0
        while count <= sheet.ncols:
                count += mod_data[index][2] + 1
                index += 1
       
        while True:
            
            if mod_data[index+1][2] == 0:     #Cell is not Merged
                count = 1
            else:                           #Cell is Merged
                count = 0
                
            sr_no = raw_data[index][0]
            data = raw_data[index+1][0]
            
            #Next info present after 48cols i.e sheet.ncols+1 cols later
            #Skip 1 row
            while count <= sheet.ncols:
                count += mod_data[index][2] + 1
                index += 1
            
            #Last TnC Break is obtained when the word total is seen
            try:
                temp = data.lower()
                if temp == 'total':
                    break
            except AttributeError:
                pass
            
            #Flag = False indicates that the upper row has missing data.
            #If data is missing in the current row as well then terminate the loop
            if not flag:
                flag = True
                #Data and Sr no both are missing
                if data == '' and sr_no == '':        
                    break
                
                #Only Data is missing
                elif data == '':   
                    t.append((sr_no, data))
                
                #Only serial no is missing
                elif sr_no == '':
                    t.append((sr_no, data))
                
            #SrNo is present
            if type(sr_no) == int or type(sr_no) == float:
                
                #Proper data is present in both srno and description
                #col so add it to the list
                if not data == '' and type(data) == str:   
                    t.append((sr_no, data))
                
                #SrNo is present but data is missing so check if the data is present in the
                #next row or not
                elif data == '':
                    flag = False
            
            #SrNo is absent and data is also absent so check the next row
            elif sr_no == '' and data == '':
                flag = False
            
            #Data is present so append the data
            elif sr_no == '' and (not data == ''):
                t.append((sr_no, data))
        
        return(t)
        
        
    def quote_simplify(self):
        
        sheetname = self.get_sheetname(loc+filename)
        try:
            raw_data, sheet, wb = self.get_workbook_data(self.filename, sheetname)
        
    
        except FileNotFoundError:
            return("File Not Found")
        
        except xlrd.XLRDError:
            return("No Sheet Found")
        
        
        except NoDataError:
            return("No data available in the Sheet")
        
        mod_data = self.modify_data(raw_data)
        final_data = self.get_data(wb, sheet, raw_data, mod_data, self.short_keys, self.final_data, self.discard_data)
        policies, tnc = self.get_policy(raw_data, mod_data, sheet, self.final_data, self.policies, self.tnc)
        
        return(final_data)            

#Input from user
loc = ''    #Location of the file that is to be read. End the location with a forward slash and always use forward slash in the Location
filename = '' #EXCEL File name



#Driver Code         
qps = Quote_to_Policy(loc+filename)
final_data = qps.quote_simplify()


