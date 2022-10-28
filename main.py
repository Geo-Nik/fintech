from restrictions import DataUpdater


if __name__ == '__main__':
    pth_to_data_file = 'data/accounts.xlsx'
    data_dict_instance = DataUpdater(pth_to_data_file)
    data_dict = data_dict_instance.update_data_dict()
    print(data_dict)
