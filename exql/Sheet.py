"""excel sheet object - all teh dataz"""
from exql import utils

class Sheet:

    def __init__(self, **kargs):
        self.values = utils.get_all_values_from_ws(ws = kargs['ws'],d = 'r')
        self.tbl_nm = kargs['name']

        if kargs.has_key('col_nms') and kargs['col_nms'] is not None:
            self.col_nms = kargs['col_nms']
        else:
            self.col_nms = self.values[0]

        #work around to remove headers. shuold be fixed
        self.values = self.values[1:]