#!/usr/bin/env python3
# -*- coding:utf-8 -*-


# 测试数值类型
def my_num():
    # Python3 支持 int、float、bool、complex
    a, b, c, d = 1, 1+2j, 1.2, True
    print(a)
    pass
    print(b)
    print(c)
    print(d)
    print(a+c)
    print(a+b)
    print(a+d)
    print(isinstance(a, int)+isinstance(b, complex))
    print(isinstance(c, float))  # 没有double类型
    print(isinstance(d, bool))


def my_string():
    str = 'haozhen15033799577'
    print(str[1])
    print(str[1:2])
    print(str[1:])
    print(str[1:-1])
    print(str[1:6:2])  # 从1至6 步长为2
    print(str[-1::-1])  # 反转，反向步长1
    print(str[-1::-2])  # 反转，反向步长2
    print(str * 2)
    print(str + str)


def my_tuple():
    tuple_a = ()
    tuple_b = ('b', )  # 一个元素，需要在元素后添加逗号
    tuple_c = ('haozhen', 1, 1+1j, 1.2)
    print(tuple_a)
    print(tuple_b)
    print(tuple_c)
    print(tuple_c[1])  # 这里返回 1
    print(tuple_c[1:2])  # 这里返回 (1,)
    print(tuple_c[1:])
    print(tuple_c[1:-1])
    print(tuple_c[1:5:2])  # 注意5 并不越界
    print(tuple_c[-1::-1])  # 反转，反向步长1
    print(tuple_c[-1::-2])  # 反转，反向步长2
    print(tuple_c + tuple_b)
    print(tuple_c * 2)


def my_list():
    list_a = [1, 'haozhen', 3.1, 1+1j]
    list_b = [5]
    print(list_a)
    print(list_a[1])
    print(list_a[1:])
    print(list_a[1:2])
    print(list_a[1:-1])
    print(list_a[1:4:2])
    print(list_a[-1::-1])
    print(list_a[-1::-2])
    print(list_a + list_b)
    print(list_b + ['1'])
    print(list_b * 2)

    # 修改
    list_b[0] = 2
    print(list_b)


def my_set():
    set_b = {1, 2, 2, 5, 2, 6, '1'}
    set_a = {1, 2}
    set_c = set()  # 空集合
    print(set_b)
    print(set_b - set_a)
    print(set_b | set_a)
    print(set_b & set_a)
    print(set_b ^ set_a)


def my_dirt():
    dirt_a = {'hao': 'zhen', 'meng': 'xiaoqing', 'hao1': 'ren1'}
    print(dirt_a)
    print(dirt_a['hao'])
    print(dirt_a.keys())
    print(dirt_a.values())
    dirt_a['hao1'] = 'huai1'
    dirt_a['hao2'] = 'ren2'
    del dirt_a['hao1']
    print(dirt_a)
    print('hao2' in dirt_a)

    n = len(dirt_a)
    print(f'dirt_a len is {n}')

# my_num()
# my_string()
# my_tuple()

# my_list()
# my_set()
my_dirt()
