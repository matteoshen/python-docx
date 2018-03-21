# encoding: utf-8

"""
Utilities
"""

SPECIALCHARS_EN = r'~!@#$%^&*()_-+={}[]|\`:\"\'<>?/.,;'
SPECIALCHARS_CH = r'·~！@#￥%……&*（）——+-={}|【】：“‘；：”’《》，。？、'
SPECIAL_SEP = [' ','\n','\t','\u3000']

def rmSpecailChar(s):
    """移除一些特殊字符：比如标点符号
    @s: 需要被处理的字符串
    """
    for c in SPECIALCHARS_EN + SPECIALCHARS_CH:
        s = s.replace(c,'')
    for i in SPECIAL_SEP:
        s = s.replace(i,'')
    return s