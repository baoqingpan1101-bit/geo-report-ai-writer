"""
岩土工程写作模板引擎
Geo Engineering Writer Templates

用法：
    from geo_writer import *
    print(承载力建议('三', '细砂', '180'))
    print(桩基持力层建议('十', '中砂', '3000', '六', '粉质黏土', '900'))

导入后请配合'项目参数锁定表'使用，确保参数准确。
"""

from geo_writer_templates import *
from geo_writer_conclusions import *
from geo_writer_specialized import *