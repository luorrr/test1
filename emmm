# student_info.py
# 学生上报疫情项目：
import xlsxwriter

def meun():
    menu_info ='''＋－－－－－－－－－－－－－－－－－－－－－－＋
｜ 1）添加学生疫情信息                        ｜
｜ 2）显示所有学生疫情信息                    ｜
｜ 3）删除学生疫情信息                        ｜
｜ 4）修改学生疫情信息                        ｜
｜ 5）按上报日期查寻                          ｜
｜ 6）按学号查询
｜ 7）双查询
｜ 8）生成柱状图                              ｜
｜ 9）保存学生信息到文件（students.txt)       ｜
｜ 10）从文件中读取数据（students.txt)        ｜
｜                                            ｜
｜ 退出：其他任意按键＜回车＞                 ｜
＋－－－－－－－－－－－－－－－－－－－－－－＋
'''
    print(menu_info)


# 以下二个函数用于sorted排序，　key的表达式函数
def get_id(*l):
    for x in l:
        return x.get("id")
def get_date(*l):
    for x in l:
        return x.get("date")
        
# 1）添加学生信息
def add_student_info():
    L = []
    while True:
        n = input("请输入名字：")
        if not n:  # 名字为空　跳出循环      
            break
        try:
            i = int(input("请输入学号："))
            d = float(input("请输入填报日期："))
        except:
            print("输入无效，不是整形数值．．．．重新录入信息")
            continue
        a = input("请输入省份：")
        p = input("请输入是否感染(是/否):")
        info = {"name":n,"id":i,"date":d,"area":a,"patient":p}
        L.append(info)
    print("学生信息录入完毕！！！")
    return L

# 2）显示所有学生的信息
def show_student_info(student_info):
    if not student_info:
        print("无数据信息．．．．．")
        return
    print("姓名".center(8),"学号".center(8),"日期".center(8),"地区".center(16),"是否感染".center(4))
    for info in student_info:
        print(info.get("name").center(10),str(info.get("id")).center(4),str(info.get("date")).center(4),str(info.get("area")).center(16),str(info.get("patient")).center(8),)

# 3）删除学生信息
def del_student_info(student_info,del_id = ''):
    if not del_id:
        del_id = input("请输入删除的学生学号：")
    for info in student_info:
        if del_id == info.get("id"):
            return info
    raise IndexError("学生信息不匹配,没有找到%s" %del_id)

# 4）修改学生信息
def mod_student_info(student_info):
    mod_id = input("请输入修改的学生学号：")
    for info in student_info:
        if mod_id == info.get("id"):
            n = input("请输入姓名：")
            d = float(input("请输入日期："))
            a = input("请输入地区：")
            p = input("请输入是否感染(是/否):")
            info = {"name":name,"id":mod_id,"date":d,"area":a,"patient":p}
            return info
    raise IndexError("学生信息不匹配,没有找到%s" %mod_id)

# 5）按上报日期查寻
def date_find(student_info):
    x = input("按上报日期查寻:")
    print("姓名".center(8),"学号".center(8),"日期".center(8),"地区".center(16),"是否感染".center(4))
    for info in student_info:
        if x == info.get("date"):
            print(info.get("name").center(10),str(info.get("id")).center(4),str(info.get("date")).center(4),str(info.get("area")).center(16),str(info.get("patient")).center(8),)

# 6）按学号查寻
def id_find(student_info):
    y = input("按学号查寻:")
    print("姓名".center(8),"学号".center(8),"日期".center(8),"地区".center(16),"是否感染".center(4))
    for info in student_info:
        if y == info.get("id"):
            print(info.get("name").center(10),str(info.get("id")).center(4),str(info.get("date")).center(4),str(info.get("area")).center(16),str(info.get("patient")).center(8),)

# 7）双查询
def double_find(student_info):
    y = input("按学号查寻:")
    x = input("按上报日期查寻:")
    print("姓名".center(8),"学号".center(8),"日期".center(8),"地区".center(16),"是否感染".center(4))
    for info in student_info:
        if y == info.get("id")&x == info.get("date"):
            print(info.get("name").center(10),str(info.get("id")).center(4),str(info.get("date")).center(4),str(info.get("area")).center(16),str(info.get("patient")).center(8),)

# 8) 生成柱状图
def excel(student_info):
    y = int(input("起止日期eg.314:"))
    x = int(input("截止日期eg.317:"))
    file=open("students.txt",'r')
    dataN=list() #初始化一个空列表 用来存该日患病数量的数量
    dataN=file.readlines()
    #使用readlines()函数 读取文件的全部内容，存成一个列表，每一项都是以换行符结尾的一个字符串，对应着文件的一行
    dataList = [0 for x in range(0,x+1-y)]
    print(dataList)
    count = 0
    for date in range(y,x+1):
        d = str(date/100.0)
        for info in student_info:
            if d == info.get("date"):
                dataList[count]+=1
        count+=1

    print(dataList)

    # 创建一个excel
    workbook = xlsxwriter.Workbook("number.xlsx")
    # 创建一个sheet
    worksheet = workbook.add_worksheet()
    # worksheet = workbook.add_worksheet("bug_analysis")

    # 自定义样式，加粗
    bold = workbook.add_format({'bold': 1})
    
    # 1、准备数据并写入excel
    # 向excel中写入数据，建立图标时要用到
    headings = ['Date','Number']
    

    # 写入表头
    worksheet.write_row('A1', headings, bold)

    # 写入数据
    dlist = list()
    for i in range(y,x+1):
        dlist.append(i)
        
    worksheet.write_column('A2', dlist)
    worksheet.write_column('B2', dataList)

    # 2、生成图表并插入到excel
    # 创建一个柱状图(column chart)
    chart_col = workbook.add_chart({'type': 'column'})

    # 配置第一个系列数据
    chart_col.add_series({
    # 这里sheet1
        'name': '=Sheet1!$B$1',
        'categories': '=Sheet1!$A$2:$A$7',
        'values':   '=Sheet1!$B$2:$B$7',
        'line': {'color': 'red'},
    })

    # 配置第二个系列数据
    chart_col.add_series({
        'name': '=Sheet1!$C$1',
        'categories':  '=Sheet1!$A$2:$A$7',
        'values':   '=Sheet1!$C$2:$C$7',
        'line': {'color': 'yellow'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '患病人数柱状图'})
    chart_col.set_x_axis({'name': '日期'})
    chart_col.set_y_axis({'name':  '人数'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet以及偏移
    worksheet.insert_chart('A10', chart_col, {'x_offset': 25, 'y_offset': 10})

    workbook.close()

# 9）保存学生信息到文件（students.txt)
def save_info(student_info):
    try:
        students_txt = open("students.txt","a")     # 以写模式打开
    except Exception as e:
        students_txt = open("students.txt", "x")    # 文件不存在，创建文件并打开
    for info in student_info:
        students_txt.write(str(info)+"\n")          # 按行存储，添加换行符
    students_txt.close()

# 10）从文件中读取数据（students.txt) 
def read_info():
    old_info = []
    try:
        students_txt = open("students.txt")
    except:
        print("暂未保存数据信息")                       # 打开失败，文件不存在说明没有数据保存
        return
    while True:
        info = students_txt.readline()
        if not info:
            break
        # print(info)
        info = info.rstrip()    #　去掉换行符
        # print(info)
        info = info[1:-1]       # 去掉｛｝
        # print(info)
        student_dict = {}       # 单个学生字典信息
        for x in info.split(","):   # 以，为间隔拆分
            # print(x)
            key_value = []      # 开辟空间，key_value[0]存key,key_value[0]存value
            for k in x.split(":"):  # 以：为间隔拆分
                k = k.strip()       #　去掉首尾空字符
                # print(k)
                if k[0] == k[-1] and len(k) > 2:        # 判断是字符串还是整数
                    key_value.append(k[1:-1])           # 去掉　首尾的＇
                else:
                    key_value.append(str(k))
                # print(key_value)
            student_dict[key_value[0]] = key_value[1]   # 学生信息添加
        # print(student_dict)
        old_info.append(student_dict)   # 所有学生信息汇总
    students_txt.close()  
    return old_info   

def main():
    student_info = []
    while True:
        # print(student_info)
        meun()
        number = input("请输入选项：")
        if number == '1':
            student_info = add_student_info()
        elif number == '2':
            show_student_info(student_info)
        elif number == '3':
            try:
                student_info.remove(del_student_info(student_info))
            except Exception as e:
                # 学生学号不匹配
                print(e)            
        elif number == '4':
            try:                
                student = mod_student_info(student_info)
            except Exception as e:
                # 学生学号不匹配
                print(e)
            else:
                # 首先按照根据输入信息的学号，从列表中删除该生疫情信息，然后重新添加该学生最新疫情信息
                student_info.remove(del_student_info(student_info,del_id = student.get("id")))  
                student_info.append(student)
        elif number == '5':
            date_find(student_info)
        elif number == '6':
            id_find(student_info)
        elif number == '7':
            double_find(student_info)
        elif number == '8':
            excel(student_info)
        elif number == '9':
            save_info(student_info)
        elif number == '10':
            student_info = read_info()
        else:
            break
        input("回车显示菜单")

main()
