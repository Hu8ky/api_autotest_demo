import paramunittest
import unittest
import HTMLTestRunner
import requests
import time
from xlrd import open_workbook
import os

MajorDir = os.path.split(os.path.realpath(__file__))[0]
def get_xls(xls_name, sheet_name):
    data = []
    xlsPath = os.path.join(MajorDir, xls_name)
    file = open_workbook(xlsPath)
    sheet = file.sheet_by_name(sheet_name)
    for i in range(sheet.nrows):
        if sheet.row_values(i)[0] != 'case_name':
            data.append(sheet.row_values(i))
    return data

xls_info = get_xls('apiTestData.xlsx', 'Sheet1')
@paramunittest.parametrized(*xls_info)
class Search(unittest.TestCase):
    def setParameters(self,case_name,uri,q,tag,start,count):
        if type(q) == float:
            q = int(q)
        if type(tag) == float:
            tag = int(tag)
        if type(start) == float:
            start = int(start)
        if type(count) == float:
            count = int(count)
        self.case_name = str(case_name)
        self.uri = str(uri)
        self.q = str(q)
        self.tag = str(tag)
        self.start = str(start)
        self.count = str(count)
    def description(self):
        self.case_name


    def setUp(self):
        self.host = 'https://api.douban.com'

    def testSearch(self):
        #拼接url
        self.url=self.host+self.uri+'search'
        #合成参数，如果值为'空'，则把键去除
        params = {'q':self.q,'tag':self.tag,'start':self.start,'count':self.count }
        delList = []
        for i in params.items():
            if i[1] == '空':
                delList.append(i[0])
        for j in delList:
            del params[j]
        #调用接口
        r = requests.get(self.url,params=params)
        code = r.json().get('code')
        #如果返回错误码，直接报错
        if code != None:
            assert False, 'Invoke Error.Code:\t{0}'.format(code)
        else:
            if params.get('start') is not None or params.get('count') is not None:
                # 传入start参数时，只校验翻页功能
                # check_response.check_pageSize(r.json())
                # 判断分页结果数目是否正确
                count = r.json().get('count')
                start = r.json().get('start')
                total = r.json().get('total')
                diff = total - start
                if diff >= count:
                    expectPageSize = count
                elif count > diff > 0:
                    expectPageSize = diff
                else:
                    expectPageSize = 0
                self.assertEqual(expectPageSize, len(r.json().get('subjects')))
            else:
                # 校验搜索结果是否与搜索词匹配
                # 由于搜索结果存在模糊匹配的情况，这里简单处理只校验第一个返回结果的正确性
                if self.count != '空':
                    # 期望结果数目不为None时，只判断返回结果数目
                    self.assertEqual(self.count,len(r.json()['subjects']))
                else:
                    if not r.json().get('subjects'):
                        # 结果为空，直接返回失败
                        assert False
                    else:
                        # 结果不为空，校验第一个结果
                        subject = r.json().get('subjects')[0]
                        # 先校验搜索条件tag
                        if params.get('tag'):
                            for word in params['tag'].split(','):
                                genres = subject['genres']
                                if word == 'action':
                                    self.assertIn('动作',genres)
                                else:
                                    self.assertIn(word,genres)
                        # 再校验搜索条件q
                        elif params.get('q'):
                            # 依次判断片名，导演或演员中是否含有搜索词，任意一个含有则返回成功
                            result=[]
                            for word in params['q'].split(','):
                                title = [subject['title']]
                                casts = [i['name'] for i in subject['casts']]
                                directors = [i['name'] for i in subject['directors']]
                                total = title + casts + directors
                                #每次匹配的结果存到result
                                for j in total:
                                    result.append(word in j)
                                #结果中有一次成功就是成功
                                self.assertIn(True,result)
    def tearDown(self):
        pass
if __name__ == '__main__':
    def create_report():
        test_unit = unittest.TestSuite()
        case_discover = unittest.defaultTestLoader.discover(MajorDir, 'test_search.py')
        for case1 in case_discover:
            for case2 in case1:
                test_unit.addTest(case2)

        now = time.strftime("%Y-%m-%d", time.localtime(time.time()))
        report_dir = MajorDir+"\\" + now + "_result.html"
        open_report = open(report_dir, "ab")
        runner = HTMLTestRunner.HTMLTestRunner(stream=open_report, title='test', description="test")
        print(report_dir)
        runner.run(test_unit)
        open_report.close()
    create_report()
