# GGTabulator
Goods Groupon Tabulater——谷圈开团打表小助手  
[Github Release - 点击下载](
https://github.com/Banny-D/GGTabulator/releases/download/Latest/GGT_beta.zip)

## 简述
程序可将排表自动生成肾表，妈妈再也不用担心我打表出错啦

## 安装
1. 解压缩
2. 将`GGT.exe`、`input.xlsx`、`symbol.xlsx`文件放置在同一目录下即可

## 使用说明
- `input.xlsx`是输入部分。
    - 注意一定需要将**均价**写在`D1`单元格内
    - `A1`单元格可以修改为需要的名字
    - `B`列填写**调价**，`C`列填写**数量**，一定要写满不然会报错
    - 如果有分盒，请将分盒名称放在第一列，并用于划分分盒列表（可参考：`demo_input.xlsx`）
    - 运行程序前一定要检查好表格的完整性，当前版本不支持检查调价是否配平及是否有余量
- `symbol.xlsx`是简称替换，如果不需要可以删除
    - 用途：将明细栏简化
    - 这个表格可以编辑，`A`列为排表中商品的名称，`B`列为简写
    - 不全也没关系，找不到的内容将沿用原来的名称
- `output.xlsx`是输出表格

## TODO
- [ ] 表格完整性校验
- [ ] 更轻量化的程序打包

## 开发者
邮箱：<xinyi.bit@qq.com>
