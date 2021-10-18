import gspread
from oauth2client.service_account import ServiceAccountCredentials


def open_sheet(spreadsheetname, sheetname):
    """Logs into Google account and returns sheet object.
        Takes the SpreadSheet name and then Sheet name as arguments.
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('db/reefer.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(spreadsheetname)
    data = sheet.worksheet(sheetname)
    # trucks = data.get_all_values()
    return data


def create_truck(truck_list):
    """Creates a dictionary from list of trucks
    """
    truck_dict = {}
    truck_dict['load_no'] = truck_list[0]
    truck_dict['married_load'] = truck_list[1]
    truck_dict['truck_no'] = truck_list[2]
    truck_dict['carrier_name'] = truck_list[3]
    truck_dict['dispatched'] = truck_list[4]
    return truck_dict


class Truck:
    def __init__(self, truck_list):
        self.load_no = truck_list[0]
        self.married_load = truck_list[1]
        self.truck_no = truck_list[2]
        self.carrier_name = truck_list[3]
        self.dispatched = truck_list[4]


def reefer_list_data():
    """Gets list of POs from current week Boa Warehousing Reefer List Google Sheet.
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('db/reefer.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('Warehousing Sheet Data Feed')
    data = sheet.worksheet('rl_qry')
    reeferlist = data.get_all_values()
    return reeferlist
