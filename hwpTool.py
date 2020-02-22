import os
from hwp_api_class import HwpApi


def convert_to_hml(hwp_address):
    hwpapi = HwpApi()
    hml_address = hwpapi.hwpOpen(hwp_address)
    hml_address = hwpapi.hwpSaveAs(hwp_address, save_ext=".hml")
    hwpapi.hwpQuit()
    return hml_address


def convert_to_hwp(hml_addr):

    hwpapi = HwpApi()

    print(hml_addr)
    if os.path.exists(hml_addr[:-3] + "hwp"):
        os.remove(hml_addr[:-3] + "hwp")
        print(hml_addr[:-3] + "hwp")
    hwpapi.hwpOpen(hml_addr)
    hwpapi.hwpSaveAs(hml_addr, save_ext=".hwp")
    hwpapi.hwpFileClose()
    os.remove(hml_addr)
    hwpapi.hwpQuit()


def check_output_path(data_path):

    output_path = os.path.join(data_path, "output")
    if not os.path.isdir(output_path):
        os.mkdir(output_path)

    return output_path


def open_hwp_file(hwp_address):
    hwpapi = HwpApi()
    hwpapi.hwpOpen(hwp_address)
