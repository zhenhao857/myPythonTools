#!/usr/bin/env python3
# -*- coding:utf-8 -*-


def my_def(name, sex, age):
    return name, sex, age


def my_def2(age, name='haozhen'):  # 非默认参数在默认前
    return name, age


def my_def3(*name):  # tuple
    return name[0]


def my_def4(**name):  # dirt
    return name


def my_def5(name, *, sex='男', age):  # age为命名关键字key
    return name, age, sex


def my_def6(age, *name, sex):  # sex为命名关键字key
    return age, name[0], sex

# 参数定义的顺序必须是：必选参数、默认参数、可变参数、关键字参数和命名关键字参数


print(my_def('haozhen', '1', '2'))
print(my_def(sex='haozhen', name='1', age=2))

print(my_def2(31, 'haohua'))
print(my_def2(30))

print(my_def3(1, 2))

print(my_def4(a=1, b=2))

print(my_def5('haozhen', age=30))

print(my_def6(30, 'haozhen', sex='男'))


