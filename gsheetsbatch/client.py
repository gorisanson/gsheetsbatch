"""
client
~~~~~~~~~~~~~~~

This module contains Client class responsible for communicating with
Google Sheets API v4.

And it also contains functions and classes for checking write quota.
Client class automatically adjust the speed of executing requests.
Write quota is now 100 create and batchUpdate requests executions per 100 seconds.

For more info about quota, visit the following page.

https://developers.google.com/sheets/api/limits

"""
# Todo: add functions for checking read quota

from .api_service import create_service
import time
import math

from pprint import pprint
from functools import wraps

from .models import Spreadsheet


class TimeRequestTable:
    def __init__(self):
        self.time_table = []
        self.end_time = -1

    def register(self, elapsed_time, num_of_executed_requests):
        # 등록은 천천히 하는 것이 쿼터를 넘지 않는데 이로우므로 올림을 한다.(더 빨리 등록되는 건 문제가 될 수 있다)
        elapsed_time = int(math.ceil(elapsed_time))
        if elapsed_time > self.end_time:
            if self.end_time == -1:
                last_num_of_executed_requests = 0
            else:
                last_num_of_executed_requests = self.time_table[self.end_time]
            for i in range(self.end_time + 1, elapsed_time):
                self.time_table.append(last_num_of_executed_requests)

            self.time_table.append(num_of_executed_requests)
            self.end_time = elapsed_time

    def get_requests_num_at(self, elapsed_time):
        # 참조는 앞에 있는 것을 하는 것이 지난 100초 동안 요청 수가 크게 잡히므로 쿼터를 넘지 않는데 이롭다.
        return self.time_table[int(math.floor(elapsed_time))]


def under_write_quota(execute_function):
    """Quota limit test decorator.
    """

    @wraps(execute_function)
    def wrapper(*args, **kwargs):
        Client.test_for_write_quota()
        result = execute_function(*args, **kwargs)
        Client.num_of_executed_write_requests += 1
        elapsed_time = time.time() - Client.time_start
        Client.time_requests_table.register(elapsed_time,
                                            Client.num_of_executed_write_requests)
        return result

    return wrapper


class RequestsContainer:
    """Container for batchUpdate requests_container."""

    def __init__(self):
        self.d = {}

    def deposit(self, spreadsheet_id, request):
        if spreadsheet_id not in self.d:
            self.d[spreadsheet_id] = []
        self.d[spreadsheet_id].append(request)


class Client:
    """An instance of this class communicates with Google Sheets API.

    """
    # Todo: To include class variables below to instance variables.
    service = create_service()
    time_start = time.time()
    num_of_executed_write_requests = 0
    time_requests_table = TimeRequestTable()

    def __init__(self, google_account_email=None):
        """If sheet protection request is to be used, set google_account_email.

        google_account_email: str (ex: 'example@gmail.com')
        spreadsheet_id: Spreadsheet ID
        """
        self.google_account_email = google_account_email
        self.requests_container = RequestsContainer()  # for batchUpdate requests containing

    @staticmethod
    def test_for_write_quota():
        """Test for whether below (write) quota limit. If it's not, sleep for proper time.
        (According to Usage Limits web page of google (https://developers.google.com/sheets/api/limits)
        "Limits for reads and writes are tracked separately.")

        returns: bool : time_start_reset flag (if this is True, should reset cls.time_start after response)
        """
        elapsed_time = time.time() - Client.time_start
        future_executed_requests_num = Client.num_of_executed_write_requests + 1
        # (한 project 의 한 user 마다) 100초당 100개의 요청이 Google Sheets API 의 정해진 quota 다.
        if elapsed_time < 100:
            request_per_second = future_executed_requests_num / elapsed_time
            print("If I execute this requests now, during past {} seconds I execute {} requests. requests/seconds: {}".
                  format(elapsed_time, future_executed_requests_num, request_per_second))
            if future_executed_requests_num > 100:
                print("whoa.... too fast.")
                sleep_time = 100 - elapsed_time
                print("I'll take a nap for {} seconds...".format(sleep_time))
                time.sleep(sleep_time)
        else:
            requests_num_last_100s = future_executed_requests_num \
                                     - Client.time_requests_table.get_requests_num_at(elapsed_time - 100)
            print("If I execute this requests now, during past {} seconds I execute {} requests. requests/seconds: {}".
                  format(100, requests_num_last_100s, requests_num_last_100s / 100))
            if requests_num_last_100s > 100:
                print("whoa.... too fast. It's no good if request execution number during past 100 seconds"
                      "is over 100...")
                for i in range(0, 101):
                    temp = future_executed_requests_num \
                           - Client.time_requests_table.get_requests_num_at(elapsed_time - 100 + i)
                    if temp < 100:
                        break
                sleep_time = i
                print("I'll take a nap for {} seconds...".format(sleep_time))
                time.sleep(sleep_time)

    @under_write_quota
    def create_spreadsheet(self, title):
        """Create a spreadsheet file whose title is title and returns it

        :param title: str
        :returns: Spreadsheet object
        """
        spreadsheet_body = {
            'properties': {
                'title': title
            }
        }
        request = self.service.spreadsheets().create(body=spreadsheet_body)
        response = request.execute()

        pprint(response)
        print()

        spreadsheet_json = response
        return Spreadsheet(self, spreadsheet_json)

    def open_by_id(self, spreadsheet_id):
        """Returns Spreadsheet object corresponding to spreadsheet_id

        :param spreadsheet_id: str
        :returns: Spreadsheet object
        """
        request = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=True)
        spreadsheet_json = request.execute()
        return Spreadsheet(self, spreadsheet_json)

    def execute_all_deposited_requests(self):
        """Executes all deposited (batchUpdate) requests in self.requests_container.

        :returns: None
        """
        # if doesn't change to list, following for loop erase RuntimeError which says
        # "RuntimeError: dictionary changed size during iteration"
        for spreadsheet_id in list(self.requests_container.d.keys()):
            self.execute_deposited_requests_of_the_spreadsheet(spreadsheet_id)

    @under_write_quota
    def execute_deposited_requests_of_the_spreadsheet(self, spreadsheet_id):
        """Executes deposited requests of the spreadsheet whose spreadsheet id is spreadsheet_id.

        :param spreadsheet_id: str
        :returns: None
        """
        assert spreadsheet_id in self.requests_container.d, "No requests deposited for spreadsheet " + spreadsheet_id
        requests = self.requests_container.d.pop(spreadsheet_id)
        requests_body = {
            'requests': requests
        }
        response = self.service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                           body=requests_body).execute()
        pprint(response)
        print()


if __name__ == '__main__':
    pass
