#!/usr/bin/env python3
# -*- coding:utf-8 -*-


def my_sentence():
    if True:
        print(True)

    bool_a = False
    if bool_a:
        print(bool_a)
    else:
        print(False)

    bool_b = False
    if bool_b:
        print(bool_b)
    elif bool_a:
        print(bool_a)
    else:
        print(False)

    int_a = 1
    while int_a < 100:
        int_a += 1
    print(int_a)

    int_b = 1
    while int_b < 100:
        int_b += 1
    else:
        print(int_b)

    for int_i in range(5):
        print(int_i)

    for int_k in range(5):
        print(int_k)
    else:
        print('finished')

    # break
    # continue
    # del
    # pass

    try:
        print('test try')
        raise NameError('HiThere')
    except NameError:
        print('test except')
    else:
        print('test else')
    finally:
        print('test finally')

    # as
    # with

    list_init = [1, 2, 3]
    print([3*x for x in list_init])


my_sentence()

