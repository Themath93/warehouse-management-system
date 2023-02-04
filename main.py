#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "Your Name"
__version__ = "0.1.0"
__license__ = "MIT"

import sys

from datajob.dw.total_stock_tf import TotalStockTransformer




# def extract_execute_daily():
#     RawMaterialsExtractor.extract_data()
#     OilPreciousMetalExtractor.extract_data()
#     StockIndexExtractor.extract_data()
#     SovereignYieldExtractor.extract_data()
#     BankInterestExtractor.extract_data()
#     ExchangeExtractor.extract_data()

# def extract_execute_monthly():
#     GlobalMarketCapExtractor.extract_data()



def main():
    works = {
        'extract':{
            # 'execute_daily':extract_execute_daily
            # ,'execute_montly':extract_execute_monthly
            # , 'futures_market_rm': RawMaterialsExtractor.extract_data
            # , 'futures_market_op': OilPreciousMetalExtractor.extract_data
            # , 'spot_market_si' : StockIndexExtractor.extract_data
            # , 'spot_market_sy' : SovereignYieldExtractor.extract_data
            # , 'spot_market_bi' : BankInterestExtractor.extract_data
            # , 'spot_market_ex' : ExchangeExtractor.extract_data
        },
        'transform':{
            'total_stock_tf': TotalStockTransformer.transform
        }

    }
    
    return works

works = main()

if __name__ == "__main__":
    """ This is executed when run from the command line """
    args = sys.argv
    work = works[args[1]][args[2]]
    work()