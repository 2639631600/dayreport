#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
1需要安装Oracle Client，由于权籍系统为32位版本，所以这里也要装32位Client
2需要安装cx_Oracle,命令：pip install cx_Oracle这个是安装和python一样的版本，python64位，这个命令
也安装64位cx_Oracle,为了使用权籍系统，需要安装32位python3.5.2版本（高版本目前32位cx_Oracle还不支持）
因此cx_Oracle也必须安装32位版本，可以在官方网站下载exe安装包。

下一步计划：
    1、给表格调整格式 ok
    2、最后一行添加权证数量合计 ok
    3、作为windows程序运行，可以设置自动报送时间，设置发件邮箱，接收邮箱，输出文件路径
    4、制作GUI界面
"""
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import os.path
import cx_Oracle
import xlwt
import time
import os,sys



'连接oracle数据库'

__author__ = 'jeff'

# 定义时间
v_curr_time = time.strftime('%H%M%S', time.localtime(time.time()))
v_curr_date = time.strftime('%Y-%m-%d', time.localtime(time.time()))

v_yesterday = time.strftime('%Y-%m-%d', time.localtime(time.time() - 24*60*60))
v_friday = time.strftime('%Y-%m-%d', time.localtime(time.time() - 24*60*60*3))

v_yesterday_weekday = time.localtime(time.time() - 24*60*60).tm_wday
v_curr_weekday = time.localtime(time.time()).tm_wday

#获得文件路径
#获取脚本文件的当前路径
def cur_file_dir():
     #获取脚本路径
     path = sys.path[0]
     #判断为脚本文件还是py2exe编译后的文件，如果是脚本文件，则返回的是脚本的目录，如果是py2exe编译后的文件，则返回的是编译后的文件路径
     if os.path.isdir(path):
         return path
     elif os.path.isfile(path):
         return os.path.dirname(path)

file_path = cur_file_dir() 
#print(file_path)
#pdb.set_trace()
# 连接数据库
def data2excle(expdate):
    conn = cx_Oracle.connect("bdck", "wqsalis", "10.88.112.25:1521/wqorcl")
    cur1 = conn.cursor()
    # 组查询语句，如果是多行结尾需要加反斜杠连接
    v_sql = """select 
A.ywlsh,
case 
  when c.BDCDYH is null then d.BDCDYH
  else c.BDCDYH
end as BDCDYH,
case
  WHEN dyfs is not null or ygdjzl is not null THEN '0'
  ELSE '1'
end as SZLX 

from (select XMBH,Replace(rtrim(PROJECT_ID),'-','') AS PID,XMMC,SLRY,SLSJ,YWLSH,qllx,djlx from bdck.bdcs_xmxx) A,
(SELECT Replace(rtrim(substr(PROJECT_ID,1,29)),'-','') as PID from bdck.bdcs_printrecord where to_char(PRINTTIME,'YYYY-MM-DD') = '""" + expdate + "' "

    v_sql += """ group by  PROJECT_ID having count(*)>=1) B,
(select XMBH,bdcdyh,replace(ywh,'-','') as pid,dbr,qlid from bdck.bdcs_ql_xz t where to_char(djsj,'yyyy-mm-dd')>='2016-10-31' and dbr is not null order by bdcqzh) C,
(select XMBH,bdcdyh,replace(ywh,'-','') as pid,dbr from bdck.bdcs_ql_ls t where to_char(djsj,'yyyy-mm-dd')>='2016-10-31' and dbr is not null order by bdcqzh) D,
(select qlid,dyfs,YGDJZL from BDCS_FSQL_XZ) E
where A.PID = B.PID and c.XMBH (+)= A.XMBH and D.XMBH(+)=A.XMBH and e.qlid(+)=c.qlid
order by ywlsh"""
    # print(v_sql)
    # pdb.set_trace()
    cur1.execute(v_sql)
    rows = cur1.fetchall()
    v_cnt = len(rows)

    # 生成excel文件
    book = xlwt.Workbook()
    sheet1 = book.add_sheet('Sheet1')
    # 设置单元格格式---------------------------------------------------
    style_string = """
        font:
            name Arial,
            height 300;
        align:
            wrap on,
            vert center,
            horiz center;
        borders:
            left THIN,
            right THIN,
            top THIN,
            bottom THIN;
    """
    style = xlwt.easyxf(style_string)
    #-----------------------------------------------------------------


    # 把列名当作一行数据写入
    sheet1.write(0, 0, '序号', style)
    sheet1.write(0, 1, '受理编号', style)
    sheet1.write(0, 2, '不动产单元号', style)
    sheet1.write(0, 3, '证书类型', style)
    # 设置列宽
    sheet1.col(0).width = 10 * 256
    sheet1.col(1).width = 20 * 256
    sheet1.col(2).width = 50 * 256
    sheet1.col(3).width = 20 * 256
    zs = 0
    zm = len(rows)

    # 当查出多列数据时，需要一个cell一个cell的写入，要不然这四列就会写到excel的一个cell里
    for i in range(len(rows)):
        # sheet1.write有几行range()中写几行
        if rows[i][2] == '1':
                zs = zs + 1 # 取证书数量
        sheet1.write(i + 1, 0, i + 1, style)
        for j in range(3):
            sheet1.write(i + 1, j + 1, rows[i][j], style)
    zm = zm - zs #取证明数量
    ##合并单元格
    sheet1.write_merge(len(rows) + 1, len(rows) + 1, 0, 3, "证书数量：%d" % zs, style)
    sheet1.write_merge(len(rows) + 2, len(rows) + 2, 0, 3, "证明数量：%d" % zm, style)
    book.save(v_file_name)
    cur1.close()
    conn.close()
# pdb.set_trace()
# 邮件信息
def email2sm(fromstr, tostr, fname):
    From = fromstr
    To = ','.join( tostr )   # "1059297224@qq.com" # ;1181389875@qq.com
    file_name = fname

    # server = smtplib.SMTP("smtp.qq.com") 非加密服务器发送
    server = smtplib.SMTP_SSL("smtp.qq.com",465) # smtplib.SMTP_SSL("服务器名",端口号)
    server.login("1059297224", "aulgbicotuvtbbda")  # 仅smtp服务器需要验证时

    # 构造MIMEMultipart对象做为根容器
    main_msg = email.mime.multipart.MIMEMultipart()

    # 构造MIMEText对象做为邮件显示内容并附加到根容器
    # v_str = "星巴克活动报表，数据共" + v_cnt + "行"
    text_msg = email.mime.text.MIMEText("发证统计日报表，请查收。")
    main_msg.attach(text_msg)

    # 构造MIMEBase对象做为文件附件内容并附加到根容器
    contype = v_file_name
    # maintype, subtype = contype.split(' ') # 这里用了空格所以v_file_name文件扩展名后面要加一个空格
    maintype, subtype = contype.split(' ')
    ## 读入文件内容并格式化
    data = open(file_name, 'rb')
    file_msg = email.mime.base.MIMEBase(maintype, subtype)
    file_msg.set_payload(data.read())
    data.close()
    email.encoders.encode_base64(file_msg)

    ## 设置附件头
    basename = os.path.basename(file_name)
    file_msg.add_header('Content-Disposition',
                        'attachment', filename=basename)
    main_msg.attach(file_msg)

    # 设置根容器属性
    main_msg['From'] = From
    main_msg['To'] = To
    main_msg['Subject'] = "发证统计日报表"
    main_msg['Date'] = email.utils.formatdate()

    # 得到格式化后的完整文本
    fullText = main_msg.as_string()

    # 用smtp发送邮件
    try:
        server.sendmail(From, To, fullText)
        print("发送成功")
    except Exception as e:
        print("发送失败")
        print(str(e))
    finally:
        server.quit()

if v_curr_weekday == 0 : # 如果当天是周一
    data_date = v_friday
    v_file_name = file_path + '\\report\\' + v_friday + v_curr_time + '.xls '
    data2excle(data_date)
    email2sm("1059297224@qq.com", ['1181389875@qq.com','1059297224@qq.com'], v_file_name)
elif v_curr_weekday <= 5 and v_curr_weekday > 0:# 如果当天是工作日
    data_date = v_yesterday
    v_file_name = file_path + '\\report\\' + v_yesterday + v_curr_time + '.xls '
    data2excle(data_date)
    email2sm("1059297224@qq.com", ['1181389875@qq.com','1059297224@qq.com'], v_file_name)
