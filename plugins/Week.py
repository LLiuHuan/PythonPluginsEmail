#插件1
# -*- coding: utf-8 -*-

from PluginManager import Model_ToolBarObj
import importlib
import sys
import time, datetime
import xlwt
import os
sys.path.append("../")


class Plugin1(Model_ToolBarObj):
    def __init__(self):
        # month='*',
        # day_of_week='*',
        # day='*',
        # hour='*',
        # minute=str(item.minute)
        self.corn = {'day_of_week': '1', 'hour': 8}
        self.filename = "Plugins1"
        self.MysqlPool = importlib.import_module("config.MySqlPool")
        self.pool = self.MysqlPool.MysqlPool(host="127.0.0.1",
                                             port=3306,
                                             user="user",
                                             password="password",
                                             database="test")
        self.tools = importlib.import_module("config.tools")
        self.emails = importlib.import_module("config.emails")
        self.email = self.emails.Email()

    def Start(self):
        users = self.pool.fetch_all(
            "select u.id,u.company_id, u.branchCompany_id, u.email, u.role_id,c1.companyName company_name,c2.companyName branchCompany_name from user u left join company c1 on u.company_id = c1.id left join company c2 on u.branchCompany_id = c2.id where  u.role_id in (1,2,3) and u.email is not null and u.email != ''"
        )
        for user in users:
            if user['role_id'] == 2:
                title = "%s 各分公司周报表" % str(int(time.strftime('%Y%W')) - 1)
            else:
                title = "%s %s周报表" % (str(int(time.strftime('%Y%W')) - 1),
                                      user['branchCompany_name'])
            self.filePaths = []
            filePath, fileName = self.getSaleWeekDetail(user)

            self.filePaths.append(filePath)
            filePath, fileName = self.getSalWeekModelDetail(user)
            if filePath:
                self.filePaths.append(filePath)
            isEmail = self.email.send_email(
                self.filePaths, "846814824@qq.com,184555556@qq.comm", title)
            if isEmail:
                print("%s 发送成功" % title)
            else:
                print("%s 发送失败" % title)
        print("周报表发送完毕，请查收！")

    def getSaleWeekDetail(self, email_user):
        """
        周销售明细
        """
        yearweek = int(time.strftime('%Y%W')) - 1
        if email_user['role_id'] == 3:

            company_name = email_user['branchCompany_name']
            company_id = email_user['branchCompany_id']
            str_where = "and u.branchCompany_id = %s " % company_id
            title = "%s第%s周(%s-%s)周销售明细" % (
                company_name, yearweek,
                datetime.datetime.strptime(str(yearweek) + '1',
                                           '%Y%W%w').strftime('%Y-%m-%d'),
                datetime.datetime.strptime(str(yearweek) + '0',
                                           '%Y%W%w').strftime('%Y-%m-%d')
            )  # （2020.5.00-2020.5.00）
        else:
            company_name = email_user['company_name']
            company_id = email_user['company_id']
            str_where = "and u.company_id = %s " % company_id
            title = "各分公司第%s周(%s-%s)周销售明细" % (
                yearweek,
                datetime.datetime.strptime(str(yearweek) + '1',
                                           '%Y%W%w').strftime('%Y-%m-%d'),
                datetime.datetime.strptime(str(yearweek) + '0',
                                           '%Y%W%w').strftime('%Y-%m-%d')
            )  # （2020.5.00-2020.5.00）
        week_sql = """
        select c.companyName as '公司',
        tt.text '类型',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=1,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周一',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=2,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周二',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=3,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周三',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=4,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周四',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=5,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周五',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=6,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周六',
        round(sum(if(DATE_FORMAT(oi.create_time, '%%w')=0,(case tt.text when '总数' then 1 else op.product_count_price end),0)),2) '周日',
        round(sum(case tt.text when '总数' then 1 else op.product_count_price end),2) '合计'
        from orderinfo oi left join orderproduct op on op.order_id = oi.id left join product p on op.product_id = p.id left join temp_text tt
        on tt.type = 1 left join user u on oi.user_id = u.id left join company c on u.branchCompany_id = c.id where yearweek(oi.create_time,5) = {}
        and oi.order_status = 1 and oi.payment_status = 2 and (op.product_type is null or op.product_type = 5) {}
        group by c.companyName, tt.text order by convert(c.companyName using gbk)
        """.format(yearweek, str_where)

        data = self.pool.fetch_all(week_sql)
        workbook = xlwt.Workbook(encoding='utf-8', style_compression=2)
        xlsheet = workbook.add_sheet("销售明细", cell_overwrite_ok=True)

        fields = ['公司', '类型', '周一', '周二', '周三', '周四', '周五', '周六', '周日', '合计']
        # 插入表头 一行n列 合并
        self.tools.data_is_merge(xlsheet,
                                 0, [title],
                                 0,
                                 is_merge=True,
                                 rcol=len(fields) - 1,
                                 is_bold=True)
        # 插入标题 一行n列
        self.tools.data_is_merge(xlsheet,
                                 1,
                                 fields,
                                 0,
                                 is_merge=True,
                                 is_lr=False,
                                 is_bold=True)
        data = self.tools.change_data(data, fields)
        # 设置第一列合并
        merge = [0]
        # 插入内容
        for i, v in enumerate(data):
            # 判断是否合并
            is_merge = True if i in merge else False
            self.tools.data_is_merge(xlsheet, i, v, is_merge=is_merge)

        curPath = os.path.abspath(os.path.dirname(__file__))
        rootPath = curPath[:curPath.find("P-Plugin" + "\\") +
                           len("P-Plugin" + "\\")]
        filepath = os.path.join(rootPath, "Emails",
                                "%s%s销售明细.xls" % (yearweek, company_name))
        workbook.save(filepath)
        return filepath, title

    def getSalWeekModelDetail(self, email_user):
        """
        周销售型号明细
        """
        if email_user['role_id'] == 2:
            return False, ""
        yearweek = int(time.strftime('%Y%W')) - 1
        week_start_date = datetime.datetime.strptime(
            str(yearweek) + '1', '%Y%W%w').strftime('%Y-%m-%d')
        week_end_date = datetime.datetime.strptime(
            str(yearweek + 1) + '1', '%Y%W%w').strftime('%Y-%m-%d')
        company_name = email_user['branchCompany_name']
        title = "%s第%s周(%s-%s)周销售型号明细" % (company_name, yearweek,
                                          week_start_date, week_end_date)
        data = self.pool.proc_all('month_order_report_form', [
            email_user['company_id'], email_user['branchCompany_id'], yearweek
        ])

        str_sql = "select group_concat(p.productName) name from product p where productState = 1"
        field_data = self.pool.fetch_one(str_sql)
        field_data = field_data['name'].split(',')
        fields = ['公司', '姓名', '合计']
        fields[2:2] = field_data
        workbook = xlwt.Workbook(encoding='utf-8', style_compression=2)
        xlsheet = workbook.add_sheet("销售型号明细", cell_overwrite_ok=True)
        self.tools.data_is_merge(xlsheet,
                                 0, [title],
                                 0,
                                 is_merge=True,
                                 rcol=len(fields) - 1,
                                 is_bold=True)
        self.tools.data_is_merge(xlsheet,
                                 1,
                                 fields,
                                 0,
                                 is_merge=True,
                                 is_lr=False,
                                 is_bold=True)
        # print(data)
        data = self.tools.change_data(data, fields)
        # print(data)
        a = [0]
        for i, v in enumerate(data):
            if i in a:
                is_merge = True
            else:
                is_merge = False
            self.tools.data_is_merge(xlsheet, i, v, is_merge=is_merge)
        # cur = settings.STATIC_FILE
        curPath = os.path.abspath(os.path.dirname(__file__))
        rootPath = curPath[:curPath.find("P-Plugin" + "\\") +
                           len("P-Plugin" + "\\")]
        filepath = os.path.join(rootPath, "Emails",
                                "%s%s销售型号明细.xls" % (yearweek, company_name))
        workbook.save(filepath)

        return filepath, title


# 邮件
# EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
# EMAIL_HOST = 'smtp.163.com'
# EMAIL_PORT = 25
# # 发送邮件的邮箱
# EMAIL_HOST_USER = '15245608547@163.com'
# # 在邮箱中设置的客户端授权密码
# EMAIL_HOST_PASSWORD = 'CXIJCLMTGTGMTHAE'
# # 收件人看到的发件人
# EMAIL_FROM = 'python<15245608547@163.com>'