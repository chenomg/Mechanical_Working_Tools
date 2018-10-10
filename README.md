# Mechanical_Working_Tools
Mechanical_Working_Tools

PDF CHECKER:

>   用于检查PDF文件和BOM中的文件名是否能对应，显示出多余和缺少的PDF
>
>   第一列为flag，若为正整数则检查，若为-1则为对称件，不需要图纸
>
>   第二列为文件名, PDF文件和表格需要放在和本程序同一文件夹

BOM CHECKER:

> 检查BOM(第一张表格)表中文件名以及部件属性是否相匹配
>
> 第一列为标识，为正整数时检查该行数据
>
> 第二列为文件名，结构：图号_名称
>
> 第三列为属性中图号
>
> 第四列为属性中名称

BOM Checker - project.py
>检查BOM和由SolidWorks导出的清单进行校对，核对BOM中项目有无缺失， 以及数量是否正确
>
>第一列为标识，为正整数时检查该行数据
>
>第二列为图号
>
>第三列为数量