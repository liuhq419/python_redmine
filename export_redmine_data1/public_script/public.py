# -*- coding:utf-8 -*-
def all_indexs(lst, obj):
    '''
    返回所有obj的index
    '''
    def find_index(lst, obj, start=0):
        try:
            index = lst.index(obj, start)
        except:
            index = -1
        return index
 
    indexes = []
    i = 0
    while True:
        idx = find_index(lst, obj, i)
        if idx == -1:
            return indexes
        indexes.append(idx)
        i = idx + 1
    return indexes
 
 