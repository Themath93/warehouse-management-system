#!/usr/bin/env python3
"""
Module Docstring
"""

__author__ = "Your Name"
__version__ = "0.1.0"
__license__ = "MIT"

import sys
import os

class GitPull:
    
    @classmethod
    def git_pull(self):
        os.system('git status')
    @classmethod
    def test(self):
        os.system('echo Hello World')

def main():
    works = {
        'start':{
            "GitPull":GitPull.git_pull,
            'test':GitPull.test
            # , 'futures_market_rm': RawMaterialsExtractor.extract_data
            # , 'futures_market_op': OilPreciousMetalExtractor.extract_data
            # , 'spot_market_si' : StockIndexExtractor.extract_data
            # , 'spot_market_sy' : SovereignYieldExtractor.extract_data
            # , 'spot_market_bi' : BankInterestExtractor.extract_data
            # , 'spot_market_ex' : ExchangeExtractor.extract_data
        },
        'transform':{
            # 'total_stock_tf': TotalStockTransformer.transform
        }

    }
    
    return works

works = main()

if __name__ == "__main__":
    """ This is executed when run from the command line """
    args = sys.argv
    work = works[args[1]][args[2]]
    work()