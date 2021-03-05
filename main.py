import getopt
import os
import requests
import socket
import sys
from datetime import datetime
from typing import Dict
from dotenv import load_dotenv
from xlwt import Workbook

response_api_data = Dict[str, str]


def get_api_token() -> (str, str, str, str):
    """It reads token from environment file and returns it"""
    load_dotenv()
    return os.getenv("API_TOKEN"), os.getenv("APOD_ENDPOINT"), os.getenv("ASTEROID_NEO_LOOKUP"), \
           os.getenv("ASTEROID_NEO_FEED")


# def multiple_list(sheet1, cell_position, responce_data):
#    """Handling multiple list"""
#    for x in responce_data:
#        cell_position = write_multiple_dict(sheet1, cell_position, x)
#    return cell_position


def write_multiple_dict(sheet1, cell_position, responce_data) -> int:
    """Handling nested dictionary """
    for x in responce_data:
        if isinstance(responce_data[x], dict):
            cell_position = write_multiple_dict(sheet1, cell_position, responce_data[x])
        sheet1.write(0, cell_position, x)
        sheet1.write(1, cell_position, str(responce_data[x]))
        cell_position += 1
    return cell_position


def write_excel(responce_data: response_api_data, call_from: str) -> None:
    """Excel write, call_from will be the sheet name"""
    ## xlwt work book
    wb: Workbook.Workbook = Workbook()
    ## sheet creation
    sheet1: Workbook.Workbook = wb.add_sheet("api_response")
    ## column position
    cell_position: int = 0
    try:
        for x in responce_data:
            """Excel sheet with dict keys will be the first row data and
                its corresponding values will be in the second row"""
            if isinstance(responce_data[x], dict):
                cell_position = write_multiple_dict(sheet1, cell_position, responce_data[x])
                continue
            elif isinstance(responce_data[x], list):
                ###Now code skips list entry
                # cell_position = multiple_list(sheet1, cell_position, responce_data[x], cell_header)
                continue
            else:
                sheet1.write(0, cell_position, x)
                sheet1.write(1, cell_position, str(responce_data[x]))
                cell_position += 1
                wb.save(f'{call_from}.xls')
        print(f'File is created at {os.path.abspath(f"{call_from}.xls")}')
    except Exception as e:
        print(e)


def capture_remote_ip() -> str:
    """capture remote IP address"""
    return socket.gethostbyname(socket.gethostname())


def date_validation(date: str) -> datetime:
    """Validating date"""
    try:
        return datetime.strptime(date, "%Y-%m-%d")
    except Exception as e:
        print(f"Exception occurred in date validation {e}")


def api_responce_validation(api_response: response_api_data, call_from: str) -> None:
    """validating and response data formation"""
    print("Validating response")
    if api_response.status_code == 200:
        api_response_data: response_api_data = api_response.json()
        api_response_data["X-RateLimit-Remaining"] = api_response.headers['X-RateLimit-Remaining']
        api_response_data["remote_ip"] = capture_remote_ip()
        write_excel(responce_data=api_response_data, call_from=call_from)
    else:
        print(f"Bad response code{api_response.status_code}")


def apod_api_call(api_token: str, call_from: str, apod_endpoint: str) -> None:
    """apod api call, call_from is our sheet name and here it is APOD"""
    try:
        print('-' * 50, '\n', ' ' * 10, f'Welcome to {call_from} download\n', '-' * 50)
        print("Loading APOD api...")
        ## fetching urls from .env
        # resp: response_api_data = requests.get(f'https://api.nasa.gov/planetary/apod?api_key={api_token}')
        apod_endpoint = apod_endpoint.format(api_token=api_token)
        resp: response_api_data = requests.get(apod_endpoint)
        api_responce_validation(api_response=resp, call_from=call_from)
        print("-" * 50)
    except Exception as e:
        print(f"Exception occurred {e}")


def asteroids_api_call(api_token: str, call_from: str,asteroid_neo_lookup: str, asteroid_neo_feed: str) -> None:
    """Api to all asteroids, if astroid id then use neo api else feed"""
    print('-' * 50, '\n', ' ' * 10, f'Welcome to {call_from} download\n', '-' * 50)

    ## getting astroid id from user
    astroid_id: str = str(input("Enter you astroid id:"))

    ##If Astroid id is entered use Neo - Lookup api otherwise use Neo - Feed
    if astroid_id:
        print("loading Neo - Lookup call...")
        api_link: str = asteroid_neo_lookup.format(astroid_id=astroid_id, api_token=api_token)# f'https://api.nasa
        # .gov/neo/rest/v1/neo/{astroid_id}?api_key={api_token}'
    else:
        ## get start date and end date
        print("loading Neo - Feed call...")
        start_date: str = date_validation(input("Enter start date in yyyy-mm-dd").strip())
        end_date: str = date_validation(input("Enter end date in yyyy-mm-dd").strip())
        api_link: str = asteroid_neo_feed.format(start_date=start_date, end_date=end_date, api_token=api_token)#f"https://api.nasa.gov/neo/rest/v1/feed?start_date={start_date}&end_date={end_date}&api_key={api_token}"
    try:
        resp = requests.get(api_link)
        api_responce_validation(api_response=resp, call_from=call_from)
    except Exception as e:
        print(f"Exception occurred in asteroid api call and exception is {e}")


def main():
    """main.py -h or --help to request help
       main.py -p or --Apod to run APOD api
       main.py -s or --Asteroids to run Neows api"""
    # token
    api_token, apod_endpoint, asteroid_neo_lookup, asteroid_neo_feed = get_api_token()
    # Remove 1st argument from the
    # list of command line arguments
    argumentList: list[str] = sys.argv[1:]
    # option
    options: str = "hps"
    # long option
    long_options: list[str] = ["Help", "Apod", "Asteroids"]
    try:
        arguments, value = getopt.getopt(argumentList, options, long_options)
        for current_argument, value in arguments:
            print(f"Current option {current_argument}")
            if current_argument in ("-h", "--Help"):
                print("-p or --Apod to run APOD api \n-s or --Asteroids to run Neows API")
            elif current_argument in ("-p", "--Apod"):
                print("Processing APOD")
                apod_api_call(api_token=api_token, call_from="APOD", apod_endpoint=apod_endpoint)
            elif current_argument in ("-s", "--Asteroids"):
                print("Processing Asteroids")
                asteroids_api_call(api_token=api_token, call_from="Asteroids", asteroid_neo_lookup=asteroid_neo_lookup, asteroid_neo_feed=asteroid_neo_feed)
    except getopt.error as err:
        print(str(err))


if __name__ == '__main__':
    main()
