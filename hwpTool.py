import os
import time
from hwp_api_class import HwpApi


def convert_to_hml(hwp_address):
    hwpapi = HwpApi()
    hml_address = hwpapi.hwpOpen(hwp_address)
    hml_address = hwpapi.hwpSaveAs(hwp_address, save_ext=".hml")
    hwpapi.hwpQuit()
    return hml_address


def convert_to_hwps(folder_address):

    hwpapi = HwpApi()
    hml_list = [
        os.path.join(folder_address, file)
        for file in os.listdir(folder_address)
        if file[-3:] == "hml"
    ]

    for hml in hml_list:
        hwp_address = hwpapi.hwpOpen(hml)
        hwpapi.hwpSaveAs(hml, save_ext=".hwp")
        hwpapi.hwpFileClose()
        os.remove(hml)
    hwpapi.hwpQuit()


def check_output_path(data_path):

    output_path = os.path.join(data_path, "output")
    if not os.path.isdir(output_path):
        os.mkdir(os.path.join(output_path, "output"))

    return output_path
