#!/usr/bin/env python3
# -*- coding:utf-8 -*-


def my_operator():
    a = 'haozhen'
    if n := len(a) > 1:
        print(f'a len {n}')

    b = 2
    c = 3
    print(b ** c)
    print(b // c)
    print(c // b)

    print(len(a))


my_operator()


