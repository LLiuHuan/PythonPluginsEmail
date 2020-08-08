# -*- coding: utf-8 -*-

from time import sleep
from PluginManager import Model_ToolBarObj
import importlib
import sys
import time, datetime
import xlwt
import os
sys.path.append("../")


class Mothon(Model_ToolBarObj):
    def __init__(self):
        # month='*',
        # day_of_week='*',
        # day='*',
        # hour='*',
        # minute=str(item.minute)
        self.corn = {'day': '8', 'hour': 11, 'minute': 36}
        self.filename = "Mothon"
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
        dt = datetime.datetime.now()
        for user in users:
            self.filePaths = []
            if user['role_id'] == 2:
                com_name = user['company_name']
            else:
                com_name = user['branchCompany_name']
            title = "%s %s月报表" % (datetime.datetime(
                dt.year, dt.month - 1, 1).strftime("%Y-%m"), com_name)
            # 月销售明细
            filePath, fileName = self.getSaleMothonDetail(user)
            if filePath:
                self.filePaths.append(filePath)

            # 月销售提成
            filePath, fileName = self.getSaleMothondRoyalty(user)
            if filePath:
                self.filePaths.append(filePath)

            # 月销售金额汇总
            filePath, fileName = self.getSaleRealDetail(user)
            if filePath:
                self.filePaths.append(filePath)

            if len(self.filePaths) > 0:
                isEmail = self.email.send_email(
                    self.filePaths, "846814824@qq.com,184555556@qq.com", title)
                if isEmail:
                    print("%s 发送成功" % title)
                else:
                    print("%s 发送失败" % title)
        print("月报表发送完毕，请查收！")

    def getSaleMothonDetail(self, email_user):
        """
        月销售详细
        """
        if email_user['role_id'] == 2:
            return False, ""
        dt = datetime.datetime.now()
        month_start_date = datetime.datetime(dt.year, dt.month - 1,
                                             1).strftime("%Y-%m-%d")
        month_end_date = datetime.datetime(dt.year, dt.month,
                                           1).strftime("%Y-%m-%d")
        company_name = email_user['branchCompany_name']
        title = "%s%s月销售明细" % (company_name,
                               datetime.datetime(dt.year, dt.month - 1,
                                                 1).strftime("%Y-%m"))
        data = self.pool.fetch_all(
            "select oi.create_time,c.fullName,c.datailaddress,c.phone,p.productName,u.userName from orderinfo oi left join orderproduct op on op.order_id = oi.id left join customer c on oi.customer_id = c.id left join user u on oi.user_id = u.id left join product p on op.product_id = p.id where  (op.product_type is null or op.product_type = 5) and oi.order_status = 1 and oi.payment_status = 2 and u.company_id = %s and u.branchCompany_id = %s and oi.create_time between '%s' and '%s' order by oi.create_time"
            % (email_user['company_id'], email_user['branchCompany_id'],
               month_start_date, month_end_date))
        fields = [
            'create_time', 'fullName', 'datailaddress', 'phone', 'productName',
            'userName'
        ]
        cn_fields = ['销售日期', '用户姓名', '地址', '电话', '购买产品型号', '销售员']
        workbook = xlwt.Workbook(encoding='utf-8', style_compression=2)
        xlsheet = workbook.add_sheet("月销售明细", cell_overwrite_ok=True)
        self.tools.data_is_merge(xlsheet,
                                 0, [title],
                                 0,
                                 is_merge=True,
                                 rcol=len(fields) - 1,
                                 is_bold=True)
        self.tools.data_is_merge(xlsheet,
                                 1,
                                 cn_fields,
                                 0,
                                 is_merge=True,
                                 is_lr=False,
                                 is_bold=True)
        data = self.tools.change_data(data, fields)
        if all([data]):
            data[0] = [i.strftime('%Y-%m-%d %H:%M:%S') for i in data[0]]
            for i, v in enumerate(data):
                self.tools.data_is_merge(xlsheet, i, v)
            curPath = os.path.abspath(os.path.dirname(__file__))
            rootPath = curPath[:curPath.find("P-Plugin" + "\\") +
                               len("P-Plugin" + "\\")]
            filepath = os.path.join(
                rootPath, "Emails", "%s%s 月销售明细.xls" %
                (company_name, datetime.datetime(dt.year, dt.month - 1,
                                                 1).strftime("%Y-%m")))
            workbook.save(filepath)
            return filepath, title
        return False, ""

    def getSaleMothondRoyalty(self, email_user):
        """
        月销售提成
        """
        if email_user['role_id'] == 2:
            return False, ""

        dynamic_data = self.pool.fetch_one(
            "select dyNumber dynumber from dynamic where dyName = 'FINANCIAL'")

        fields = [
            'userName', 'IDNumber', 'companyName', 'product_count_price',
            'should_royalty', 'real_royalty', 'remark'
        ]
        cn_fields = ['姓名', '银行卡号', '公司', '销售额', '应发提成', '实发提成', '备注']

        email_data = self.pool.fetch_one(
            "select email from user u where u.company_id = %s and role_id = 2 limit 1"
            % email_user['company_id'])
        email_list = [email_user['email'], email_data['email']]
        dt = datetime.datetime.now()
        str_sql = "select * from (select u.userName,u.IDNumber from user u where u.company_id = %s and u.branchCompany_id = %s and u.role_id = 3) a inner join (select c1.companyName,round(sum(op.product_count_price),2) product_count_price,round(sum(op.royalty),2) should_royalty,round(sum(op.royalty)/{},2) real_royalty, '分公司' remark from orderinfo oi left join orderproduct op on op.order_id = oi.id left join user u on oi.user_id = u.id left join company c1 on u.branchCompany_id = c1.id where u.company_id = %s and u.branchCompany_id = %s and (op.product_type is null or op.product_type = 5) and oi.payment_status = 2 and oi.order_status = 1 and DATE_FORMAT(oi.create_time,'%%Y%%m') = %s group by c1.companyName) b union select u.userName, u.IDNumber, c1.companyName, round(sum(op.product_count_price),2) product_count_price, round(sum(op.royalty),2) should_royalty, round(sum(op.royalty)/{},2) real_royalty, '' remark from orderinfo oi left join orderproduct op on op.order_id = oi.id left join user u on oi.user_id = u.id left join company c1 on u.branchCompany_id = c1.id where u.company_id = %s and u.branchCompany_id = %s and (op.product_type is null or op.product_type = 5) and oi.payment_status = 2 and oi.order_status = 1 and DATE_FORMAT(oi.create_time,'%%Y%%m') = %s group by c1.companyName, u.userName,u.IDNumber".format(
            dynamic_data['dynumber'], dynamic_data['dynumber'])
        data = self.pool.fetch_all(
            str_sql %
            (email_user['company_id'], email_user['branchCompany_id'],
             email_user['company_id'], email_user['branchCompany_id'],
             datetime.datetime(dt.year, dt.month - 1, 1).strftime("%Y%m"),
             email_user['company_id'], email_user['branchCompany_id'],
             datetime.datetime(dt.year, dt.month - 1, 1).strftime("%Y%m")))
        # print(data)
        title = "%s%s月销售提成" % (email_user['branchCompany_name'],
                               datetime.datetime(dt.year, dt.month - 1,
                                                 1).strftime("%Y-%m"))

        excl_name = "%s%s月销售提成.xls" % (
            email_user['branchCompany_name'],
            datetime.datetime(dt.year, dt.month - 1, 1).strftime("%Y-%m"))

        workbook = xlwt.Workbook(encoding='utf-8', style_compression=2)
        xlsheet = workbook.add_sheet("月销售提成", cell_overwrite_ok=True)
        self.tools.data_is_merge(xlsheet,
                                 0, [title],
                                 0,
                                 is_merge=True,
                                 rcol=len(fields) - 1,
                                 is_bold=True)
        self.tools.data_is_merge(xlsheet,
                                 1,
                                 cn_fields,
                                 0,
                                 is_merge=True,
                                 is_lr=False,
                                 is_bold=True)
        data = self.tools.change_data(data, fields)
        a = [0]

        for i, v in enumerate(data):
            if i in a:
                is_merge = True
            else:
                is_merge = False
            self.tools.data_is_merge(xlsheet, i, v, is_merge=is_merge)
        curPath = os.path.abspath(os.path.dirname(__file__))
        rootPath = curPath[:curPath.find("P-Plugin" + "\\") +
                           len("P-Plugin" + "\\")]
        filepath = os.path.join(rootPath, "Emails", excl_name)
        workbook.save(filepath)
        return filepath, title

    def getSaleRealDetail(self, email_user):
        """
        月销售金额汇总
        """
        if email_user['role_id'] != 2:
            return False, ""

        dt = datetime.datetime.now()

        type_data = self.pool.fetch_one(
            "select group_concat('SUM(IF(p.type_id=',pt.id,',op.product_count,0)) as ''',pt.text,'''') text from producttype pt where pt.id != 1"
        )

        data = self.pool.fetch_all(
            "select c.companyName as '公司',sum(op.product_count_price) as '金额',%s from orderinfo oi,orderproduct op,product p,user u,company c where op.order_id = oi.id and op.product_id = p.id and oi.user_id = u.id and u.branchCompany_id = c.id and oi.order_status = 1 and oi.payment_status = 2 and (op.product_type is null or op.product_type = 5) and u.company_id = '%s' and oi.create_time between '%s' and '%s' group by c.companyName order by convert(c.companyName using gbk)"
            %
            (type_data['text'], email_user['company_id'],
             datetime.datetime(dt.year, dt.month - 1, 1).strftime("%Y-%m-%d"),
             datetime.datetime(dt.year, dt.month, 1).strftime("%Y-%m-%d")))

        field_data = self.pool.fetch_one(
            "select group_concat(pt.text) field from producttype pt where pt.id != 1"
        )
        fields = ['公司', '金额']
        fields = fields + field_data['field'].split(',')
        title = "%s%s月销售金额汇总" % (email_user['company_name'],
                                 datetime.datetime(dt.year, dt.month - 1,
                                                   1).strftime("%Y-%m"))

        excl_name = "%s%s月销售金额汇总.xls" % (
            email_user['company_name'],
            datetime.datetime(dt.year, dt.month - 1, 1).strftime("%Y-%m"))

        workbook = xlwt.Workbook(encoding='utf-8', style_compression=2)
        xlsheet = workbook.add_sheet("月销售金额汇总", cell_overwrite_ok=True)
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
        data = self.tools.change_data(data, fields)
        a = [0]

        for i, v in enumerate(data):
            if i in a:
                is_merge = True
            else:
                is_merge = False
            self.tools.data_is_merge(xlsheet, i, v, is_merge=is_merge)
        curPath = os.path.abspath(os.path.dirname(__file__))
        rootPath = curPath[:curPath.find("P-Plugin" + "\\") +
                           len("P-Plugin" + "\\")]
        filepath = os.path.join(rootPath, "Emails", excl_name)
        workbook.save(filepath)
        return filepath, title