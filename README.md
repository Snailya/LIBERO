# LIBERO

LIBERO是一个用于导入西门子Delivery清单的EXCEL VSTO加载项。

## 功能

* 导入更新Delivery清单中的Date Planned、Despatched ex-works、Disp.No 项
  ![WXWorkCapture_16565816661741](https://s2.loli.net/2022/07/01/f5lZHbF764cJnEX.png)
* 高亮显示完成项（绿色）、延期项（黄色）
  ![WXWorkCapture_16565815278182](https://s2.loli.net/2022/07/01/j7U1PM2f5HsWF43.png)

## 先决条件

1. Delivery清单中应包含名为`Line#`、`Date Planned`、`Despatched ex-works`、`Disp.No`的Header；
2. Delivery清单文件名中的日期应为`yyyyMMdd`的格式；
   ![20220701145002](https://s2.loli.net/2022/07/01/KRyFAaPzkJTNYcZ.png)
3. Delivery清单中日期类型数据应为`dd.MM.yyyy`的格式；
   ![20220701145107](https://s2.loli.net/2022/07/01/omzwOCqZNa3gMbu.png)
4. 合并文件中应至少包含名为`Line#`、`Despatched ex-works`、`Disp.No`的Header

## 逻辑说明

1. 检查条件4，若不满足程序终止；
2. 检查条件2，若不满足提示用户是否使用当前日期作为后缀；
3. 检查条件1，若不满足程序终止；
4. 当合并文件为空文件时，复制清单至合并文件；
5. 当合并文件非空时，若合并文件中存在相同日期后缀的的`Date Planned`列，则提示用户忽略或覆盖该列数据；若不存在，则插入清单中的数据，并为Header添加日期后缀；
6. 覆盖`Despatched ex-works`、`Disp.No`列数据；
7. 当`Despatched ex-works`列的值为日期时，认为Delivery完成，高亮为绿色；
8. 当`Date Planned[{Date}]`列（Offset [0,-1] related to `Despatched ex-works`）的日期值大于上一个`Date Planned[{Date}]`列(Offset [0,-2])的日期值时，认为Delivery延期，高亮为黄色。
