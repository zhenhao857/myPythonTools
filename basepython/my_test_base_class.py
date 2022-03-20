#!/usr/bin/env python3
# -*- coding:utf-8 -*-


class MyClass:
    int_a = 1
    __int_b = 2

    def __init__(self, int_d):
        print(self.__int_b)
        print(int_d)

        def sss():
            def kkk():
                pass

    def my_test(self):
        pass


class MyClass2:
    int_c = 3


class MyClassChild(MyClass, MyClass2):
    int_d = 4


a = MyClassChild()
print(a.int_a+a.int_c)

