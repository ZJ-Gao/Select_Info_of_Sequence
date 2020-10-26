'''
    Author：Zijie Gao
    Date：3/21/2019
    Function: Read well_name, strata division by 'Petro Exploration Institution'
    and write them in txt documents.

    Version : 5.0
    2.0 : Write bottom_depth in
    3.0 : Add codes which can process many documents
    4.0 : Distinguish sequence 'chang1' with 'chang10'
    5.0 : Extract sequence 'chang9_2'
'''
#encoding=utf-8
import xlrd
import os
def get_filename(path,filetype):
    '''
        Read the name of the files
    '''
    name = []
    for root,dirs,files in os.walk(path):
        for i in files:
            if filetype in i:
                name.append(i.replace(filetype,''))
    return name


def get_keys(d, value):
    '''
        search key by values
    '''
    return [k for k, v in d.items() if v == value]

def select_strata(number, data):

    table = data.sheets()[0]
    # read the data of the column  of layer and institutions
    col2 = table.col_values(2)
    col8 = table.col_values(8)
    col3 = table.col_values(3)
    col4 = table.col_values(4)
    # print(col2,'\n', col8)

    # Merge two lists into one dictionary
    dictionary1 = dict(zip(col2, col8))
    # print(dictionary1)

    # Pick the data of layers provided by the institution 'Petro Exploration Institution'
    list_strata = get_keys(dictionary1, 'Petro Exploration Institution')
    # print(list_strata)
    dictionary2 = dict(zip(col2, col3))
    # Preset the final dictionary
    dict_final = {'1': 'None', '2': 'None', '3': 'None', '4': 'None', '6': 'None', '7': 'None', '8': 'None', '9': 'None', '9_2': 'None', '10': 'None', '11': 'None'}



    list_room_3 = []

    # pick the layers of the 'chang' sequences
    for i in list_strata:
        if i[0] == 'chang':
            if i[1] == str(number):
                list_room_3.append(i)

    # select the minimum data of the same sub-'chang'sequences
    for i in list_room_3:
        return dictionary2.pop(min(list_room_3))
        break

def write_txt(list):
    '''
        Write information into .txt file
    '''

    fileObject = open('strata_plus_9_shiyou.txt', 'a', encoding='utf-8')
    fileObject.write('\n')
    for ip in list:
        fileObject.writelines(ip)
        fileObject.write(' ')
    fileObject.write('\n')
    fileObject.close()

def main():
    # Get the names of the file we need to process
    path = 'C:\\changename'
    filetype = '.xls'
    name = get_filename(path, filetype)
    # print(name)

    # batch processing
    for i in name:
        # batch processing
        data = xlrd.open_workbook(i + '.xls')

        table = data.sheets()[0]

        # read the data of the column  of layer and institutions
        col2 = table.col_values(2)
        col8 = table.col_values(8)
        col3 = table.col_values(3)
        col4 = table.col_values(4)

        # Remove the Chinese strings in the column of depth
        col4.remove(max(col4))

        # remove null character
        while '' in col4:
            col4.remove('')
        # print(col4)

        # float the strings to make sure we get the numerical maximum
        col4_float = []
        for i in col4:
            e = float(i)
            col4_float.append(e)
        max_value = max(col4_float)


        # print(col2,'\n', col8)

        dictionary1 = dict(zip(col2, col8))
        # print(dictionary1)

        # Pick the data of layers provided by the institution 'Petro Exploration Institution'
        list_strata = get_keys(dictionary1, 'Petro Exploration Institution')
        # print(list_strata)
        dictionary2 = dict(zip(col2, col3))

        well_name = table.cell(3, 1).value

        # Default a dictionary which will be used to record information read from Excel documents
        dict_final = {'1': 'None', '2': 'None', '3': 'None', '4': 'None', '6': 'None', '7': 'None', '8': 'None',
                      '9': 'None', '9_2': 'None', '10': 'None', '11': 'None'}

        # pick layers except 'chang1' and 'chang10'
        list_range = ['2', '3', '4', '6', '7', '8', '9']

        # Write the second string of strata into the list
        list_one = []
        for i in list_strata:
            list_one.append(i[1])
        # print(list_one)

        # Write the third string of strata into the list
        list_two = []
        for i in list_strata:
            if len(i) >= 3:
                list_two.append(i[2])

        # Get the intersection of two lists, and pick layers except 'chang1' and 'chang10'
        tmp = [val for val in list_range if val in list_one]
        for i in tmp:
            dict_final[i] = select_strata(i, data)
        dict_final['11'] = max(col4)

        # Pick bottom depth
        list2 = []
        for i in col4:
            if i[0] == '1' or i[0] == '2':
                list2.append(i)
        dict_final['11'] = str(max_value)
        # print(dict_final.values())

        # Determines whether the 'chang1' or 'chang10' exists and whether the data should be written into .txt
        for i in col2:
            if i == 'chang1':
                dict_final['1'] = dictionary2.pop(i)

        list_10 = []
        for i in col2:
            if len(i) > 2:
                if i[0] == 'chang':
                    if i[1] == '1' and i[2] == '0':
                        list_10.append(dictionary2.pop(i))
                        dict_final['10'] = max(list_10)

        # Decide if chang9_2 exist
        for i in col2:
            if len(i) > 2:
                if i[0] == 'chang':
                    if i[1] == '9' and i[3] == '2':
                        dict_final['9_2'] = dictionary2.pop(i)
                        break


        list_write = [well_name]
        for i in dict_final.values():
            list_write.append(str(i))
        print(list_write)
        write_txt(list_write)

if __name__ =='__main__':
    main()