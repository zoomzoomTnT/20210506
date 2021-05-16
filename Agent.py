import json
import os

import pandas as pd


class Agent(object):
    def __init__(
            self,
            agent_id: str,
            agent_name: str,
            agent_level: str,
            agency: str,
            department: str,
            team: str,
            monthly_pre_collected_insurance: float,
            monthly_underwriting_insurance: float,
            monthly_pre_collected_value: float,
            monthly_underwriting_value: float,
            num_3000p: int,
            is_3000p: bool,
            daily_pre_collected1: float,
            daily_underwriting1: float,
            daily_pre_collected2: float,
            daily_underwriting2: float,
            weekly_num_3000p: int,
            advance_commission: float,
            pre_collected_to_be_returned: float,
            underwriting_to_be_returned: float,
            holiday_crawfish_pre_collected_num: int,
            holiday_crawfish_underwriting_num: int,
            holiday_gift_massage_chair_num: int,
            holiday_gift_massage_chair_fee: float
    ) -> None:
        self.agent_id = agent_id
        self.agent_name = agent_name
        self.agent_level = agent_level
        self.agency = agency
        self.department = department
        self.team = team
        self.monthly_pre_collected_insurance = monthly_pre_collected_insurance
        self.monthly_underwriting_insurance = monthly_underwriting_insurance
        self.monthly_pre_collected_value = monthly_pre_collected_value
        self.monthly_underwriting_value = monthly_underwriting_value
        self.num_3000p = num_3000p
        self.is_3000p = is_3000p
        self.daily_pre_collected = daily_pre_collected1
        self.daily_underwriting = daily_underwriting1
        self.weekly_num_3000p = weekly_num_3000p
        self.advance_commission = advance_commission
        self.pre_collected_to_be_returned = pre_collected_to_be_returned
        self.underwriting_to_be_returned = underwriting_to_be_returned
        self.holiday_crawfish_pre_collected_num = holiday_crawfish_pre_collected_num
        self.holiday_crawfish_underwriting_num = holiday_crawfish_underwriting_num
        self.holiday_gift_massage_chair_num = holiday_gift_massage_chair_num
        self.holiday_gift_massage_chair_fee = holiday_gift_massage_chair_fee

    @classmethod
    def from_file(cls, src_file: str, src_sheet: str):
        if src_file is None:
            return

        with open(os.path.join(os.path.dirname(__file__), 'config/column_name_mapping.json'), 'r') as config:
            if src_file.endswith('.xlsx'):
                df = pd.read_excel(src_file, sheet_name=src_sheet)
                json_list_str = df.to_json(orient='records', force_ascii=False)
                for k, v in json.loads(config.read()).items():
                    json_list_str = json_list_str.replace(k, v)

                return cls.from_list(json_list_str)

        return

    @classmethod
    def from_list(cls, json_list_str):
        if json_list_str is None:
            return

        json_list = json.loads(json_list_str)
        return (cls.from_json(json_dict) for json_dict in json_list)

    @classmethod
    def from_json(cls, json_dict):
        if type(json_dict) is not dict:
            json_dict = json.loads(json_dict)
        return cls(**json_dict)

    def to_json(self):
        return json.dumps(self.__dict__)
