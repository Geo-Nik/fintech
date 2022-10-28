from bs4 import BeautifulSoup
import requests
from FAO import Data as FileData
from URL import URL
import logging


class BsDocument:
    def __init__(self, url):
        self.url = url

    def get_request(self):
        try:
            request_result = requests.get(self.url).text
            return request_result
        except requests.exceptions.ConnectionError as err:
            logging.error(
                f"""The url is not correct or connection is failed.
                More deatails: {err}"""
            )
            return None

    def get_bs_document(self):
        request_result = self.get_request()
        if request_result:
            try:
                bs_document = BeautifulSoup(request_result, "html.parser")
                return bs_document
            except TypeError as err:
                logging.error(
                    f"""The type of the argument should be str.
                    More details:{err}"""
                )
                return None


class Restrictions(BsDocument):
    def __init__(self, url):
        super().__init__(url)

    def get_div_with_restrictions(self):
        restrictions_class_name = "BoiRestrictedAccountsRestricted"
        bs_document = self.get_bs_document()
        if bs_document:
            list_of_divs = bs_document.find_all(class_=restrictions_class_name)
            # The documents have only one div with such class.
            # That is why list contains the only element.
            if list_of_divs:
                return list_of_divs[0]

    def get_restricted_date(self):
        div_ = self.get_div_with_restrictions()
        if div_:
            div_elements_list = div_.contents
            # Example of 'restricted_date_element':
            # ' מוגבל עד\r\n                      22/03/2026'
            restricted_date_element = div_elements_list[-1]
            splited_text_and_date = restricted_date_element.split("    ")
            restricted_date_with_trailings = splited_text_and_date[-1]
            restricted_date = restricted_date_with_trailings.strip()
            return restricted_date


class DataUpdater(Restrictions, FileData):
    def __init__(self, path_to_data_file):
        self.path_to_data_file = path_to_data_file
        FileData.__init__(self, path_to_data_file)

    def update_data_dict(self):
        input_account_data_list = self.get_data_from_worksheet()
        output_account_data_list = []
        if input_account_data_list:
            for input_account_data in input_account_data_list:
                account_url_instance = URL(input_account_data)
                account_url = account_url_instance.get_url()
                restrictions_instance = Restrictions(account_url)
                restrictions_date = restrictions_instance.get_restricted_date()
                if restrictions_date:
                    input_account_data["isrestricted"] = True
                    input_account_data["expiration_date"] = restrictions_date

                else:
                    input_account_data["isrestricted"] = False
                    input_account_data["expiration_date"] = None

                output_account_data_list.append(input_account_data)

            return output_account_data_list


if __name__ == "__main__":
    pass