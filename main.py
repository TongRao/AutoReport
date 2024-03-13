"""
Author: TongRao

I'm done with all these dumb daily reports.
"""
from yaml import load as yaml_load
try:
    from yaml import CLoader as yaml_Loader
except ImportError:
    from yaml import Loader as yaml_Loader

import json
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm


def load_files(template_file, param_file):
    """
    Load docx template and parameters.

    Parameters
    ----------
    template_file <str>: template file name, the template file must be inside the template/ directory
    param_file <str>: parameters file, containing the parameters to fill data into the docx file

    Returns
    -------
    params <dict>: result parameters
    tpl <DocxTemplate>: docx object
    """
    # load params
    with open(f"{param_file}", 'r') as f:
        if param_file.endswith(".json"):
            params = json.load(f)
        else:
            params = yaml_load(stream=f, Loader=yaml_Loader)

    # load template
    tpl = DocxTemplate(f"template/{template_file}")
    return params, tpl


def simple_plot(_image_file):
    """
    plot a simple plot
    :return:
    """
    plt.figure(figsize=(15, 8))
    plt.plot(np.random.randint(0, 100, 10))
    plt.title(f"Sample {_image_file}", fontsize=18)
    plt.savefig(_image_file, dpi=300, bbox_inches='tight')


# **************************************************************************** #
# Simple Table
def simple_df2dict(_df):
    """
    use to convert df into dictionary, the format of result dictionary can be used to render simple docx table directly
    :param _df: source dataframe
    :return:
    """
    _table = _df.to_dict('split')
    return _table


# **************************************************************************** #
# Complex Table
def complex_table_header_depth(_table):
    """
    用来计算一个table多级表头的深度
    :param _table: target table
    :return: depth of target table's headers
    """
    return len(_table.columns[0])


def extract_headers(_table):
    """
    获取多级表头，保存在字典中
    :param _table:
    :return:
    """
    result = {}
    for i in range(complex_table_header_depth(_table)):
        # use dict to keep order
        result[i] = list(dict.fromkeys([h[i] for h in _table.columns]))
    return result


def render_report(_tql, _contents):
    _tql.render(_contents)
    tpl.save("report/{}".format("report.docx"))


# **************************************************************************** #
if __name__ == '__main__':
    # load template docx file and params
    params, tpl = load_files(template_file="template.docx", param_file="params.yaml")

    # table 1
    sample_table_1_content = {
        'table_1_col_labels': ['A', 'B', 'C', 'D', 'E'],
        'table_1_contents': [
            {'cols': ['banana', 'capsicum', 'pyrite', 'taxi', 'a']},
            {'cols': ['apple', 'tomato', 'cinnabar', 'doubledecker', 'b']},
            {'cols': ['guava', 'cucumber', 'aventurine', 'card', 'c']}
        ]
    }

    # table 2
    df_table_2 = pd.read_excel("data/table_2.xlsx")
    table_2 = df_table_2.to_dict('split')

    # table 3 complex table
    table_3 = {
        'header': ['rate1', 'rate2'],
        'header1': ['Nov', 'Dec', 'Change'],
        'header2': ['Nov', 'Dec', 'Change'],
        'index': ['area1', 'area2', 'area3'],
        'data1': [
            [100, 100, 1],
            [100, 100, 0],
            [100, 100, 0]
        ],
        'data2': [
            [100, 100, 0],
            [100, 100, 0],
            [100, 100, 0]
        ]
    }

    # fixed header table (table_4)
    table_4 = {
        'table_4_rows': [
            {"a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8},
            {"a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8},
            {"a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8},
            {"a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8},
            {"a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8},
            {"a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8},
        ]
    }

    # image_1
    image_1_file = 'img/image_1.jpg'
    simple_plot(image_1_file)
    image_1 = InlineImage(tpl, image_1_file, width=Mm(145))

    # image_2
    image_2_file = 'img/image_2.jpg'
    simple_plot(image_2_file)
    image_2 = InlineImage(tpl, image_2_file, width=Mm(145))

    # generate the final content dict, render the docx file and store the result
    final_content = {**params,
                     **sample_table_1_content,
                     **{'table_2': table_2},
                     **{'table_3': table_3},
                     **table_4,
                     **{'image_1': image_1},
                     **{'image_2': image_2}}
    tpl.render(final_content)
    tpl.save("report/{}".format("report.docx"))
