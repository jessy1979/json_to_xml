import json
import pprint

import xlsxwriter

class Filemanagement(object):
    def __init__(self):
        self.json_file_name = 'imput.json'
        self.excel_excel_file = 'output.xlsx'
        self._array_user_id = []
        self._array_heart_rate = []
        self._array_resp_rate = []
        self._array_stress_rate = []

    def read_text_file(self):
        try:
            with open(self.json_file_name)as data_file:
                data = json.load(data_file)
                pprint.pprint(data)
                for x in data["data"]:
                    user_id = x["user_id"]
                    heart_rate = int(x["heart_rate"])
                    resp_rate = int(x["respiratory_rate"])
                    stress_lvl = int(x['stress_level'])
                    self._array_user_id.append(user_id)
                    self._array_heart_rate.append(heart_rate)
                    self._array_resp_rate.append(resp_rate)
                    self._array_stress_rate.append(stress_lvl)
        except IOError:
            print('expected error')
            raise

    def save_to_excel(self):
        workbk = xlsxwriter.Workbook(self.excel_excel_file)
        worksheet = workbk.add_worksheet()
        for index, value in enumerate(self._array_user_id):
            worksheet.write(index, 0, self._array_user_id[index])
            worksheet.write(index, 1, self._array_heart_rate[index])
            worksheet.write(index, 2, self._array_resp_rate[index])
            worksheet.write(index, 3, self._array_stress_rate[index])

        workbk.close()


if __name__ == '__main__':
    fml = Filemanagement()
    fml.read_text_file()
    fml.save_to_excel()
