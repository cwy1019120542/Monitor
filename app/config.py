import os

class BaseConfig:
    OCR_WAYBILL_DIR = os.getenv("OCR_WAYBILL_DIR", "/home/cwy/monitor/ocr_waybill")
    EXCEL_JOIN_WORD_DIR = os.getenv("EXCEL_JOIN_WORD_DIR", "/home/cwy/monitor/excel_join_word")
    OCR_URL = os.getenv("OCR_URL", "http://ysocr.datagrand.cn/ysocr/ocr")
    RENAME_CONTRACT_DIR = os.getenv("RENAME_CONTRACT_DIR", "/home/cwy/monitor/rename_contract")
    PAGE_COUNT_DIR = os.getenv("PAGE_COUNT_DIR", "/home/cwy/monitor/page_count")

class DevelopConfig(BaseConfig):
    pass

class ProductConfig(BaseConfig):
    pass

config = {
    "base": BaseConfig,
    "develop": DevelopConfig,
    "product": ProductConfig
}