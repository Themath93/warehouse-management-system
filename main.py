#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "Your Name"
__version__ = "0.1.0"
__license__ = "MIT"

import sys
from datajob.dm.daily_in_tf import DailyInTransformer
from datajob.dm.daily_out_tf import DailyOutTransformer
from datajob.dm.pod_achivement_rate import PODAchivemnetRate
from datajob.dw.total_stock_tf import TotalStockTransformer




def transform_execute_daily():
    DailyInTransformer.transform()
    DailyOutTransformer.transform()
    PODAchivemnetRate.transform()
    TotalStockTransformer.transform()


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
            'transform_daily':transform_execute_daily,
            'total_stock_tf': TotalStockTransformer.transform,
            'daily_in_tf' : DailyInTransformer.transform,
            'daily_out_tf' : DailyOutTransformer.transform,
            'pod_rate': PODAchivemnetRate.transform,
        }

    }
    
    return works

works = main()

if __name__ == "__main__":
    """ This is executed when run from the command line """
    args = sys.argv
    work = works[args[1]][args[2]]
    work()