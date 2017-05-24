#coding=utf-8
import requests, xlrd, time, sys
from xlutils import copy
import json

#file_path = r'E:\testcase1.xlsx'


def readExcel(file_path):
    '''
    读取excel测试用例的函数
    :param file_path:传入一个excel文件，或者文件的绝对路径
    :return:返回这个excel第一个sheet页中的所有测试用例的list
    '''
    
    try:
        book = xlrd.open_workbook(file_path)#打开excel
    except Exception,e:
        #如果路径不在或者excel不正确，返回报错信息
        print '路径不在或者excel不正确',e
        return e
    else:
        sheet = book.sheet_by_index(0)#取第一个sheet页
        rows= sheet.nrows#取这个sheet页的所有行数
        print rows
        case_list = []#保存每一条case
        for i in range(rows):
            if i !=0:
                #把每一条测试用例添加到case_list中
                case_list.append(sheet.row_values(i))
      
        
        #print type case_list
        

        #调用接口测试的函数，把存所有case的list和excel的路径传进去，因为后面还需要把返回报文和测试结果写到excel中，
        #所以需要传入excel测试用例的路径，interfaceTest函数在下面有定义
    interfaceTest(case_list,file_path)
        
def interfaceTest(case_list,file_path):  
    res_flags = []  
    #存测试结果的list  
    request_urls = []  
    #存请求报文的list  
    responses = []  
    #存返回报文的list  
    headers = {"Content-Type": "application/x-www-form-urlencoded"} 
    
    for case in case_list:  
        ''''' 
        先遍历excel中每一条case的值，然后根据对应的索引取到case中每个字段的值 
        '''  
        try:  
            ''''' 
            这里捕捉一下异常，如果excel格式不正确的话，就返回异常 
            '''  
            #项目，提bug的时候可以根据项目来提  
            product = case[0]  
            #用例id，提bug的时候用  
            case_id = case[1]  
            #接口名称，也是提bug的时候用  
            interface_name = case[2]  
            #用例描述  
            case_detail = case[3]  
            #请求方式  
            method = case[4]  
            #请求url  
            url = case[5]  
            #入参  
            param = case[6]  
            #预期结果  
            res_check = case[7]  
          
        except Exception,e:  
            return '测试用例格式不正确！%s'%e  
       
                
        results = requests.post(url,data =  "jsonParams=%s" % param, headers=headers).json()
        
        results1 = json.dumps(results, encoding="UTF-8", ensure_ascii=False)
      
        responses.append(results)  
        
        request_urls.append(param)
        
        
      
        ''' 
        
            获取到返回报文之后需要根据预期结果去判断测试是否通过，调用查看结果方法 
            把返回报文和预期结果传进去，判断是否通过，readRes方法会返回测试结果，如果返回pass就 
            说明测试通过了，readRes方法在下面定义了。 
            '''  
        res = readRes(results1,res_check)  
            
        if 'pass' in res:  
            ''''' 
            判断测试结果，然后把通过或者失败插入到测试结果的list中 
            '''  
            res_flags.append('pass')  
        else:  
            res_flags.append('fail')  

    ''''' 
    全部用例执行完之后，会调用copy_excel方法，把测试结果写到excel中， 
    每一条用例的请求报文、返回报文、测试结果，这三个每个我在上面都定义了一个list 
    来存每一条用例执行的结果，把源excel用例的路径和三个list传进去调用即可，copy_excel方 
    法在下面定义了，也加了注释 
    '''  
    copy_excel(file_path,res_flags,request_urls,responses)  
    
       
       
def readRes(res,res_check):
    '''
    :param res: 返回报文
    :param res_check: 预期结果
    :return: 通过或者不通过，不通过的话会把哪个参数和预期不一致返回
    '''
    '''
    返回报文的例子是这样的{"id":"J_775682","p":275.00,"m":"458.00"}
    excel预期结果中的格式是xx=11;xx=22这样的，所以要把返回报文改成xx=22这样的格式
    所以用到字符串替换，把返回报文中的":"和":替换成=，返回报文就变成
    {"id=J_775682","p=275.00,"m=458.00"},这样就和预期结果一样了,当然也可以用python自带的
    json模块来解析json串，但是有的返回的不是标准的json格式，处理起来比较麻烦，这里我就用字符串的方法了
    '''
    res = res.split(',')
    
    '''
    res_check是excel中的预期结果，是xx=11;xx=22这样的
    所以用split分割字符串，split是python内置函数，切割字符串，变成一个list
    ['xx=1','xx=2']这样的，然后遍历这个list，判断list中的每个元素是否存在这个list中，
    如果每个元素都在返回报文中的话，就说明和预期结果一致
    上面我们已经把返回报文变成{"id=J_775682","p=275.00,"m=458.00"}
    '''
    #res_check = res_check.split(':')
    
    print res_check
    print res[0]
    
  
    if res_check == res[0]:
        pass
    else:
        return  r'错误，返回参数和预期结果不一致'
    return 'pass'       


def copy_excel(file_path,res_flags,request_urls,responses):  
    ''''' 
    :param file_path: 测试用例的路径 
    :param res_flags: 测试结果的list 
    :param request_urls: 请求报文的list 
    :param responses: 返回报文的list 
    :return: 
    '''  
    ''''' 
    这个函数的作用是写excel，把请求报文、返回报文和测试结果写到测试用例的excel中 
    因为xlrd模块只能读excel，不能写，所以用xlutils这个模块，但是python中没有一个模块能 
    直接操作已经写好的excel，所以只能用xlutils模块中的copy方法，copy一个新的excel，才能操作 
    '''  
    #打开原来的excel，获取到这个book对象  
    book = xlrd.open_workbook(file_path)  
    #复制一个new_book  
    new_book = copy.copy(book)  
    #然后获取到这个复制的excel的第一个sheet页  
    sheet = new_book.get_sheet(0)  
    i = 1  
    for request_url,response,flag in zip(request_urls,responses,res_flags):  
        ''''' 
        同时遍历请求报文、返回报文和测试结果这3个大的list 
        然后把每一条case执行结果写到excel中，zip函数可以将多个list放在一起遍历 
        因为第一行是表头，所以从第二行开始写，也就是索引位1的位置，i代表行 
        所以i赋值为1，然后每写一条，然后i+1， i+=1同等于i=i+1 
        请求报文、返回报文、测试结果分别在excel的8、9、11列，列是固定的，所以就给写死了 
        后面跟上要写的值，因为excel用的是Unicode字符编码，所以前面带个u表示用Unicode编码 
        否则会有乱码 
        '''  
        sheet.write(i,8,u'%s'%request_url)  
        sheet.write(i,9,u'%s'%response)  
        sheet.write(i,10,u'%s'%flag)  
        i+=1  
        
        
    #写完之后在当前目录下(可以自己指定一个目录)保存一个以当前时间命名的测试结果，time.strftime()是格式化日期  
    new_book.save('%s results.xls'%time.strftime('%Y%m%d%H%M%S'))



if __name__ == '__main__':  

    try:  
        filename = sys.argv[1]  
    except IndexError,e:  
        print 'Please enter a correct testcase! \n e.x: python gkk.py test_case.xls'  
    else:  
        readExcel(filename)  
    print 'Done!'
