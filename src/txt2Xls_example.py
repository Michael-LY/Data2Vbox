# -*- encoding: utf-8 -*-
import xlwt  # 需要的模块


def txt_xls(filename, xlsname):
    """
    :⽂本转换成xls的函数
    :param filename txt⽂本⽂件名称、
    :param xlsname 表⽰转换后的excel⽂件名
    """
    try:
        f = open(filename)
        xls = xlwt.Workbook()
        #⽣成excel的⽅法，声明excel
        sheet = xls.add_sheet('sheet1', cell_overwrite_ok=True)
        x = 0
        while True:
            #按⾏循环，读取⽂本⽂件
            line = f.readline()
            if not line:
                break  # 如果没有内容，则退出循环
            for i in range(len(line.split('\t'))):
                item = line.split('\t')[i]
                sheet.write(x, i, item)  # x单元格经度，i 单元格纬度
            x += 1  # excel另起⼀⾏
        f.close()
        xls.save(xlsname)  # 保存xls⽂件
    except:
        raise


if __name__ == "__main__":
    filename = "G:/test.txt"
    xlsname = "G:/test.xls"
    txt_xls(filename, xlsname)