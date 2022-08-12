from retrying import retry
from __init__ import Main
from selenium.webdriver.common.by import By
import pandas as pd


pd.read_excel("./unemployedSurveyForm/data/changfeng_data.xlsx", sheet_name=0)




if __name__ == '__main__':
    a = Ala()
    a.test()
