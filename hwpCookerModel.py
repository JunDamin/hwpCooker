from lxml import etree as et
import pandas as pd
from numpy import float64, int64
import re
import os
import copy


def prettify_filename(filename):

    """ Replacing strings that are not allowed in filename """

    return re.sub(r'[\\\/:*?"<>|.%]+', "", filename)


def convert_to_string(num, belowZeroDigit=0, thousand=True):

    """ Change number_format which is for Using in Koica"""

    null_list = [" 00:00:00", "1900-01-01 "]
    num_form = f",.{belowZeroDigit}f" if thousand else f".{belowZeroDigit}f"

    if type(num) in [int, float, float64, int64]:
        if num > 0:
            output = format(num, num_form)
        elif num < 0:
            output = "△" + format(-num, num_form)
        elif num == 0:
            output = "-"
        else:
            output = "  "
    else:
        output = str(num)
        for n in null_list:
            output = output.replace(n, "")
    return output


def replace_text(root, old, new):

    """ find char that match and replace text"""

    for char in root.iter("CHAR"):
        string = char.text
        if not string:
            string = ""
        if string.find(old) != -1:
            p = char.getparent().getparent()

            p_list = p.getparent()

            for i, new_text in enumerate(new.splitlines()):
                new_p = copy.deepcopy(p)
                for char in new_p.iter("CHAR"):
                    string = char.text
                    if not string:
                        string = ""
                    new_text = string
                    print(old, ":", new)
                    new_text = new_text.replace(old, new)
                    char.text = new_text
                p_list.insert(p_list.index(p) + i, new_p)
            p_list.remove(p)

    return root


def replace_doc(root, pd_series, name, belowZeroDigit=0, thousand=True):

    """ Pandas DataSeries Works"""

    print("\n 파일명 :", pd_series[name], "\n")

    for i in pd_series.index:
        old = i
        new = pd_series[i]
        if not new or type(new) in [pd._libs.tslibs.nattype.NaTType]:
            new = ""

        if old.find("__") != -1:
            new = convert_to_string(new, "", False)
        else:
            new = convert_to_string(new, belowZeroDigit, thousand)
        new = str(new)
        root = replace_text(root, old, new)
    return root


def generate_hml(
    hml_address,
    data_series,
    output_path,
    name="__파일명__",
    belowZeroDigit=0,
    thousand=True,
):

    """ generate hml files. Return generated hml address"""

    tree = et.parse(hml_address)
    root = tree.getroot()
    root = replace_doc(root, data_series, name, belowZeroDigit, thousand)
    tree = root.getroottree()

    filename = data_series[name]

    if type(filename) != str:
        filename = str(filename)

    print(filename)
    filename = prettify_filename(filename)
    file_address = os.path.join(output_path, filename + ".hml")
    tree.write(file=file_address, xml_declaration=True, encoding="utf8")

    return file_address
