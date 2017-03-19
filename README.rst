说明
*****
用于游戏开发中的数据转换

依赖
=====
*   python包"`openxl`"和"`google.protobuf`"

目标
=======
*   将Excel数据文件直接转换为Protobuf格式保存
*   自动处理Excel数据

考虑的点
==========
`ProtoBuf支持的数据类型 <https://developers.google.com/protocol-buffers/docs/proto3>`_
openxl中读取得到的\ `数据类型 <https://openpyxl.readthedocs.io/en/default/api/openpyxl.cell.cell.html>`_

*   文件字符编码问题
*   Excel中的数据格式，以及与Protobuf的对应关系。如Excel中的整形、浮点、公式、日期等怎么转换为Protobuf数据
*   将Excel中的列值设置到Protobuf
*   预留接口处理需要处理的特殊字段

当前实现
=========
游戏中使用的数据类型使用的数据类型其实相对简单，protobuff提供的类型基本满足要求的（虽然感觉有些浪费空间）；openxl经过\
测试也可以使用的支持整形，浮点，日期(datetime.datetime)。由于日期解析为了python类型，所以稍微调整一下即可。预留接口\
``registerFieldProcessor``\ 用户可以注册一个处理函数来处理表中的特殊字段。

学习知识点
============

protobuf
-------------
protobuf的基本语法:
*   ``syntax = "proto3";``
*   ``package Arrow;``
*   重复字段(与proto2一样)使用关键字"``repeated``"

对repeated字段的操作：
*   添加

    .. code-block:: python

        // item为itemTable中的一个重复字段
        // 使用add方法，即添加了一个新的空记录
        newRecord = itemTable.item.add()

*   访问

    .. code-block:: python

        for record in itemTable.item:
            // 字段不遍历比较麻烦
            print(record)

python
--------
*   使用\ ``isinstance(ob, type)``\ 来判断对象的类型，对于一些内置类型，如函数，需要使用模块\ ``types``

    .. code-block::

        if isinstance(ob, int):
            pass
        elif isinstance(ob, types.FunctionType)
            pass

    参见\ `此文 <http://stackoverflow.com/questions/624926/how-to-detect-whether-a-python-variable-is-a-function>`_

    为什么不使用type(ob)来判断？

*   datetime.datetime类型怎么转换为unix时间

             datetime.datetime.timetuple()                    time.mktime(...)返回的为浮点
    datetime ----------------------------> time.time_struct ------------------> unix time

*   函数"``setattr/getattr``"，通过属性名给对象设置属性值（这算是反射吗？）
*   如何取得执行脚本所在目标？

    首先取得路径有下面几种方法:

    .. code-block:: python

        os.getcwd()     # 返回当前工作工作。即运行脚本时所在的目录
        os.path.realpath(path)  # 返回文件的canonical path. 去除了其中可能的符号链接。真正返回文件路径的方法

        os.path.realpath(__FILE__)  # 即能得到当前脚本的全路径。再使用使用os.path中的方法进一步处理得到更多数据

    参考 `如何获得Python脚本所在目录的位置 <http://www.elias.cn/Python/GetPythonPath?from=Develop.GetPythonPath>`_
    `https://docs.python.org/2.7/library/os.path.html#os.path.realpath`_

*   模块的导入方法

    导入模块的最直接的方法当然是\ ``import moduleName``\ ，但是，有时候想导入指定的模块，而这个名称是一个变量，此时\
    ``import``\ 就无计可施了。这个时候就需要\ ``__import__``\ 函数来大展手脚。

    .. code-block::

        module = __import__("xx")

    但是，\ ``__import__``\ 函数并不会把导入的模块/函数放到当前模块空间，而直接返回，需要一个对象来接收。如果没有变量接收，能不\
    能找到呢？不知道。见后解

    如果要导入一个目录下所有模块，通常使用\ ``from pb import *``\ 。实际上它发生了什么呢？怎么去找的呢？
    `https://docs.python.org/2.7/tutorial/modules.html#importing-from-a-package`_
    事实上，import语句会去读pb目录下的\ ``__init__.py``\ 文件中的list变量\ ``__all__``\ 将其中的模块导入到\
    当前namespace。注意必须确保\ ``__all__``\ 中的模块已经导入，否则会出错。

    模块\ ``importlib``\ 是对函数\ ``__import__``\ 的包装。从中可以发现，\ ``__import__``\ 后，怎么找到导入的模块：
    ``sys.modules[modulename]`` \ 。详见代码中导入Excel对应的protobuf模块方法和importlib源码。

*   如何获得当前模块对象？

    正好上面提到的\ ``sys.modules[modulename]``\ 可以找到对应的模块。那么只要modulename为当前模块名即可。所以\ 
    ``sys.modules[__name__]`` 即可取得当前模块。

    参见：`http://stackoverflow.com/questions/1676835/python-how-do-i-get-a-reference-to-a-module-inside-the-module-itself`_

*   编码问题

    方法一：
    文件开始使用\ ``# -*- coding: utf-8 -*-``

    方法二：

    .. code-block:: python

        import sys
        reload(sys)
        sys.setdefaultencoding("utf-8")

    参考：\ `http://www.cnblogs.com/walkerwang/archive/2011/08/03/2126373.html`_

*   三元运算符

    python的三元运算符不同于c/c++的\ `` condition ? true_result : false_result``\ 的形式。python的形式如下：

    .. code-block:: python

        ret = true_result if condition else false_result

*   关于反射的一些资料

    `https://docs.lvrui.io/2016/06/16/Python%E5%8F%8D%E5%B0%84%E8%AF%A6%E8%A7%A3/`_
    `http://blog.csdn.net/lokibalder/article/details/3459722`_
    `http://www.cnblogs.com/huxi/archive/2011/01/02/1924317.html`_
    `http://pyzh.readthedocs.io/en/latest/python-magic-methods-guide.html`_

TODO
=======
将处理类改写为一个通用类，可以处理任何一个表。

