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

学习进步点
===========

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

*   datetime.datetime类型怎么转换为unix时间

             datetime.datetime.timetuple()                    time.mktime(...)返回的为浮点
    datetime ----------------------------> time.time_struct ------------------> unix time

*   函数"``setattr/getattr``"，通过属性名给对象设置属性值（这算是反射吗？）


TODO
=======
将处理类改写为一个通用类，可以处理任何一个表。
