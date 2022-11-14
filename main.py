from restrictions import DataUpdater
import FAO


if __name__ == '__main__':
    path = "data/accounts.xlsx"
    workbook_obj = FAO.Work_Book(path)
    worksheet_obj = FAO.WorkSheet(workbook_obj)
    data_object = FAO.TableData(worksheet_obj)
    updatet_data_obj = DataUpdater(data_object)
    updated_data = updatet_data_obj.update_data_dict()
    print(updated_data)
