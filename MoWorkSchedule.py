import xlrd
import xlsxwriter
from argparse import ArgumentParser
import logging
import datetime
#coding=utf-8

'''
目标：从排程Excel表格中收集指定格式的任务代码及对应数据，汇总到摘录表格中。
具体规则参看 MoWorkSchedule业务规则.docx

v 1. 先能输出指定格式单表
v 2. 完整xls多表单全输出
v 3. 用正则表达式匹配代码格式
4. 数据格式写成类
v 5. 更加智能的匹配和收集信息

    # TODO(Sloovy): use class mocode_worksheet
'''

def read_input_xls(input_file):
    try:
        output_filename, log_filename = build_output_filename(input_file)
        if output_filename == '' or log_filename == '':
            print('[ERROR]: Not a .xls or .xlsx file!')
            return None
        logging.basicConfig(filename=log_filename, level=logging.DEBUG)
        timestr = datetime.datetime.now().strftime('%Y/%m/%d %X')
        logging.info('[%s] 开始处理文件：%s', timestr, input_file)

        wb = xlrd.open_workbook(input_file)
        logging.info('共有 %d 张工作表', wb.nsheets)
        if wb.nsheets == 0:
            return None

        mocode_count_in_book = 0        # 本文件内MO编号总数
        wt_wb = xlsxwriter.Workbook(output_filename)
        #_# wt_wb = xlwt.Workbook(encoding='gbk')

        for ws in wb.sheets():
            logging.info('------  工作表：%s | 行数：%d | 列数：%d  ------', ws.name, ws.nrows, ws.ncols)
            # 隐藏的工作表不处理，xlrd.sheet.visibility != 0 (0=visible)
            if ws.visibility != 0:
                logging.info('跳过隐藏的工作表：%s', ws.name)
                continue

            #----------------------------------
            # 1. 表单名，新增表单
            out_ws = wt_wb.add_worksheet(ws.name)  # 输出表格
            #_#
            # out_ws = wt_wb.add_sheet(ws.name)  # 输出表格
            # out_ws.panes_frozen = True
            # out_ws.horz_split_pos = 1

            PRODUCT_TYPE_TEXT = '产品类别'      # 查找类别名称的关键字
            out_product_type = None            # 产品类别的内容
            out_title_names = ['开工MO', '日期', '类别', '备注']
            out_title_row_num = -1
            out_sheet_rows = 0

            #_# 设置字体
            base_font = {'font_name':'宋体', 'font_size':12, 'align':'center', 'valign':'vcenter', 'text_wrap':True}
            style_font = wt_wb.add_format(base_font)
            # 用 z = {**x, **y} 实现合并两个dict的功能, python3.5起支持
            style_date = wt_wb.add_format({**base_font, **{'num_format':'YYYY/M/D'}})
            style_title = wt_wb.add_format({**base_font, **{'bold':True, 'font_size':14}})
            style_error = wt_wb.add_format({**base_font, **{'font_color':'red'}})

            mocode_count_in_sheet = 0        # 本工作表内MO编号总数
            #----------------------------------

            for row in range(ws.nrows):
                values = []
                for col in range(ws.ncols):
                    cell = ws.cell(row,col)
                    values.append('%s[%d]' % (str(cell.value), cell.ctype))
                    # values.append(ws.cell(row,col).value)
                    #----------------------------------
                    # 2. 记录产品类别
                    if (out_product_type == None) and (cell.ctype == xlrd.XL_CELL_TEXT):
                        if (PRODUCT_TYPE_TEXT in cell.value) and (col < ws.ncols-1):
                            next_cell = ws.cell(row, col+1)
                            if next_cell.ctype == xlrd.XL_CELL_EMPTY:
                                out_product_type = ws.name
                            else:
                                out_product_type = next_cell.value
                            logging.info('找到%s：%s', PRODUCT_TYPE_TEXT, out_product_type)
                    #----------------------------------
                    # 3. 找到列名行，和MO编号对应日期
                    # 列名行特征：(1)会连续出现多个日期格式；(2)整行都有数据；
                    if (out_title_row_num == -1) and (cell.ctype == xlrd.XL_CELL_DATE):
                        CONTINUOUS_DATE_TEST_TIMES = 3    # 尝试查验连续日期格式数目
                        if col < ws.ncols-CONTINUOUS_DATE_TEST_TIMES:
                            # print(type(cell.value))
                            # print("is float:", isinstance(cell.value, float) )
                            test_pass = True
                            for i in range(CONTINUOUS_DATE_TEST_TIMES):
                                if ws.cell(row,col+i).ctype != xlrd.XL_CELL_DATE:
                                    test_pass = False
                                    break
                            # 通过连续日期查验，记录列名
                            if test_pass:
                                # out_title_row = ws.row(row)
                                out_title_row_num = row
                                logging.info('找到标题列所在行[%d]：%s...', row, ','.join(str(ws.cell(row,i).value) for i in range(col)))
                                # 补充输出表标题列名
                                for i in range(col-1):
                                    out_title_names.insert(len(out_title_names)-1, ws.cell(out_title_row_num,i).value)
                                # 为输出表格写入标题列
                                for j in range(len(out_title_names)):
                                    out_ws.write(0, j, out_title_names[j], style_title)
                                    pass
                                out_sheet_rows += 1
                    #----------------------------------
                    # 4. 查找MO编号
                    if (out_title_row_num != -1) and (cell.ctype == xlrd.XL_CELL_TEXT):
                            # 找到MO编号开头的字符串就尝试切分
                        if is_mocode_string(cell.value):
                            out_mocode_list, out_invalid_mocodes = split_mocode_cell(cell.value)
                            if len(out_mocode_list) > 0:
                                # 切分MO编号成功，先准备好共用数据
                                modate = ws.cell(out_title_row_num,col).value  # 对应日期
                                logging.info('找到日期[%s]的MO编号：%s', convert_excel_date(modate), out_mocode_list)
                                if len(out_invalid_mocodes) > 0:
                                    logging.info('找到书写错误的MO编号：%s', out_invalid_mocodes.items())
                                    pass

                                # 补充MO数据行的附加信息, 3个上面已有的MO信息，1个末尾的备注
                                modatas_append = [modate, out_product_type]
                                for i in range(len(out_title_names)-4):
                                    modatas_append.append(str(ws.cell(row,i).value))

                                for mocode in out_mocode_list:
                                    # 输出一行MO数据
                                    out_mo_datas = [mocode] + modatas_append
                                    out_mo_datas.append(out_invalid_mocodes.get(mocode, ''))     # 备注

                                    for out_col in range(len(out_mo_datas)):
                                        out_style = style_font
                                        if out_col == 1:
                                            out_style = style_date
                                        elif out_col == len(out_mo_datas)-1 and out_mo_datas[out_col] != '':
                                            out_style = style_error
                                        out_ws.write(out_sheet_rows, out_col, out_mo_datas[out_col], out_style)

                                    out_sheet_rows += 1

                                mocode_count_in_sheet += len(out_mocode_list)
                    #----------------------------------
                datastr = str(row+1) + ': ' + ','.join(str(v) for v in values)
                print(datastr)

            #_#
            out_ws.set_column(0, len(out_title_names), 18)
            out_ws.freeze_panes(1,0)
            out_ws.autofilter(0,0, out_sheet_rows-1, len(out_title_names)-1)
            # for out_col in range(len(out_title_names)):
            #     out_ws.col(out_col).width = 5120      # 256Pixels * 20Chars
            
            # 统计MO编号数量
            logging.info('工作表[%s]内共找到 %d条 MO编号', ws.name, mocode_count_in_sheet)
            mocode_count_in_book += mocode_count_in_sheet
            pass        # end of worksheet loop

        # Save output workbook
        wt_wb.close()
        #_# wt_wb.save(output_filename)
        timestr = datetime.datetime.now().strftime('%Y/%m/%d %X')
        logging.info('===================================')
        logging.info('[%s] 成功输出 %d条 MO记录至摘录表: %s', timestr, mocode_count_in_book, output_filename)

    except Exception as e:
        timestr = datetime.datetime.now().strftime('%Y/%m/%d %X')
        print(timestr, e)
        logging.exception('[%s]%s', timestr, e)
        return None

    return output_filename
    #end read_input_xls()

def string_full_width_to_half_width(ustring):
	"""Unicode字符串 全角转半角
	
	Args:
		ustring : 可能包含全角字符的Unicode字符串
	
	Returns:
		全角字符全部转换为半角字符后的字符串
	"""
	rstring = ''
	for uchar in ustring:
		inner_code = ord(uchar)
		if inner_code == 0x3000:                            #全角空格直接转换
			inner_code = 0x0020
		elif inner_code >= 0xFF01 and inner_code <= 0xFF5E: #全角字符（除空格）根据关系转化
			inner_code -= 0xFEE0

		rstring += chr(inner_code)
	return rstring
    #end string_full_width_to_half_width()

def fix_invalid_mocode(input_string):
	"""
	尝试根据MO编号规则，修正可能存在错误的输入字符串

	基础规则：由MO-或TO-开头；
	修正(容错)规则：
	(1). 字母全部大写；
	(2). 全角字符改为半角；
	(3). 第二个字符O错写为0可识别。

	Args:
		input_string : 输入字符串
	
	Returns:
		根据MO编号规则，容错修复后的字符串
	"""
    in_str = string_full_width_to_half_width(input_string).upper()
    if len(in_str) < 2:
        return ''

    if in_str[0] in ('M','T') and in_str[1] == '0':
        in_str = in_str.replace('0','O',1)
    return in_str
    #end fix_invalid_mocode()

def is_mocode_string(input_string):
	"""
	粗略判断输入字符串是否符合MO编号基础规则：
	1. 由MO-或TO-开头；
	*2. 后接10位数字，YYMMDD+4位流水号  （*本规则未判断）
	*3. 容错处理需先单独调用修复方法 fix_invalid_mocode

	Args:
		input_string : 待判断的字符串
	
	Returns:
		Boolean值。符合MO编号规则返回True，否则返回False
	"""
    MOCODE_PREFIX_LIST = ('MO-', 'TO-')
    mocode_prefix_len = len(MOCODE_PREFIX_LIST[0])

    if len(input_string) < mocode_prefix_len:
        return False

    in_str = fix_invalid_mocode(input_string[:mocode_prefix_len])
    if in_str == '':
        return False
    if in_str not in MOCODE_PREFIX_LIST:
        return False

    return True
    #end is_mocode_string()

def is_mocode_string_by_regexp(input_string):
	"""用正则表达式，精确判断输入字符串是否符合MO编号基础规则（*目前未使用）
	
	MO的编号规则是“MO-年月日流水号”，MO也可能是TO，年月日各两位，流水号4位。
	例如：MO-1806301234

	Args:
		input_string : 待测试是否MO编号的字符串
	
	Returns:
		Boolean值。符合MO编号规则返回True，否则返回False
	"""
	# 先测试是否符合“MO-8位数字”的基本格式
	if not match('^[MT]O-[0-9]{8}', input_string):
		return False
	# 再测试日期部分的有效性
	# 参考《利用日期正则表达式之识别合法日期》：https://blog.csdn.net/gtf215998315/article/details/53610048
	datestr = input_string[3:9]
	pattern = ('(([0-9]{2})(0[13578]|1[02])(0[1-9]|[12][0-9]|3[01]))|'    # 大月
				'(([0-9]{2})(0[469]|11)(0[1-9]|[12][0-9]|30))|'   # 小月
				'(([0-9]{2})02(0[1-9]|1[0-9]|2[0-8]))|'   # 平二月
				'((([13579][26]|[2468][048]|0[48])|(2000))0229)')    # 闰二月
	return match(pattern, datestr) != None
	#end is_mocode_string_by_regexp()

def split_mocode_cell(mocode_cell_value):
    """切分一个MO编号格子的字符串数据，分割为多个MO编号的列表输出

    Args:
        mocode_cell_value: Excel字符串内容，多个MO编号可能以换行或空格区分开
    
    Returns:
        out_mocode_list: 拆分开的MO编号列表。格式：['MO-1234567890', ...]
        out_invalid_mocodes: 格式有误的MO编号字典，用于错误提示。格式：{(输出MO编号, 修正前MO编号), ('MO-1234567890', 'm0－１２３4567890'), ...}
    """
    out_mocode_list = []
    out_invalid_mocodes = {}

    # 处理多个MO编号用换行间隔的情况
    lines = mocode_cell_value.splitlines()
    for line in lines:
        # 处理多个MO编号用空格间隔的情况
        mocodes = line.split()
        for code_org in mocodes:
            mocode = fix_invalid_mocode(code_org)
            # 再判断一次截取出来的MO编号是否有效，避免无效换行或字符串等情况
            if not is_mocode_string(mocode):
                print('### Invalid MO_Code:', mocode)
                continue
            # 有效MO编号加入输出列表
            out_mocode_list.append(mocode)
            # 被修正过的编号记入错误表
            if mocode != code_org:
                out_invalid_mocodes[mocode] = code_org

            print('### Find MO_Code: %s%s' % (mocode, '' if mocode == code_org else ' (%s)' % code_org))

    return out_mocode_list, out_invalid_mocodes
    #end split_mocode_cell()

def convert_excel_date(date_float):
    '''从Excel奇怪的浮点数日期格式转为普通的date类型'''
    date_base = datetime.date(1899,12,30)
    delta = datetime.timedelta(days=date_float)
    return date_base + delta
    #end convert_excel_date()

def process_cmdline_args():
    '''获取命令行参数，输入的排程Excel表格文件路径'''
    parser = ArgumentParser(description='从排程Excel表格中收集指定格式的MO任务代码及对应数据，汇总到摘录表格中')
    parser.add_argument('source_file', help='排程xls/xlsx表格文件路径')
    args = parser.parse_args()
    return str(args.source_file)
    #end process_cmdline_args()

def build_output_filename(input_filename):
    '''根据输入文件名，返回增加"_MoList"后缀的同名输出文件名，和日志文件名'''
    EXCEL_SUFFIX = '.xls'
    LOG_SUFFIX = '.log'
    OUTPUT_FILE_SUFFIX = '_MoList'

    pos = input_filename.lower().rfind(EXCEL_SUFFIX)
    if pos != -1:
        out_filename = input_filename[:pos] + OUTPUT_FILE_SUFFIX + input_filename[pos:]
        logname = input_filename[:pos] + LOG_SUFFIX
        return out_filename, logname

    return '', ''
    #end build_output_filepath()

#-----------------------
source_file = process_cmdline_args()
read_input_xls(source_file)
