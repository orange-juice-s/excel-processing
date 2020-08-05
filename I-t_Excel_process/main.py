# -*- coding: utf-8 -*-
from sys import exit
from settings import Settings
from function import print_test, program


def processing_excel():
    """
    测试文件，例如输入以下参数，表示：
    原始数据在当前文件夹的test文件下
    处理过的数据生成在当前文件夹的test_processed目录下
    原始数据是由 PDA 设备采集的，可选另一参数为2636B
    输入原文件路径，和测试机器 PDA or 2636B
    """
    # source_path = "./10-1/405/"
    # test_device = "2636B"
    while True:
        ai_settings = Settings()

        print_test(ai_settings.asterisk)

        program()

        print("-" * ai_settings.asterisk)
        order = input("退出请输入 'q' 来退出小程序~\n继续处理数据就随便输入\n")
        if order == "q":
            exit()


if __name__ == '__main__':
    processing_excel()
