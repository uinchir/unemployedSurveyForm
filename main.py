import pandas as pd
import selenium.common.exceptions as ex
import __init__ as init
number = 112


def main(range_start):
    """range_start:循环开始序号"""
    b = init.Main()  # 实例化对象
    b.load_web()  # 刷新页面
    my_excel = pd.read_excel(b.file_name, sheet_name=0, usecols=[2, 3, 4, 5, 6])  # 读取Excel数据
    for i in range(range_start, len(my_excel) + 1):
        id_name, card_id, processing_ability, willingness_to_work, special = \
            my_excel.loc[i][0], \
            my_excel.loc[i][1], \
            my_excel.loc[i][2], \
            my_excel.loc[i][3], \
            my_excel.loc[i][4]
        b.switch_to_third_frame()
        b.input_data_and_search(card_id)
        b.tick_and_check(card_id)
        try:
            b.jump_to_lay_frame()
        except ex.TimeoutException:
            b.close_abnormal_frame()
            continue
        else:
            b.fill_form(processing_ability, willingness_to_work, special)
            b.close_lay_frame()
            print(f"{i} {id_name} done")
        global number
        number = i


if __name__ == '__main__':
    while True:
        try:
            main(number)
        except ex.ElementClickInterceptedException:
            print("报错了，click错误")
            continue
        except ex.TimeoutException:
            print("报错了，找不到元素")
            continue
        except ex.NoSuchFrameException:
            print("报错了，找不到frame")
            continue
        except ValueError:
            print("搞定了，辛苦了！")
            exit()
    # main(0)
