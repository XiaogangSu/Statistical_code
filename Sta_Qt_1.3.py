from PyQt5.QtWidgets import QFileDialog,QWidget,QLineEdit, QGridLayout, QPushButton, QApplication, QLabel,\
    QMessageBox, QFrame, QTextEdit, QHBoxLayout, QVBoxLayout, QGroupBox, QComboBox
import sys
import os
import xlrd
import datetime
import time
import xlwt
import pandas as pd
import numpy as ny
from numpy import *
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']

class Sta_ex(QWidget):
    station = ''      #基地
    version = ''      #版本
    poi_target = ''   #设施版本目标
    road_target = ''  #道路版本目标
    poi_data_ex = []  #poi作业数据初步提取
    poi_time_ex = []  #poi作业时间初步提取
    road_data_ex = [] #road作业数据提取
    poi_actual = 0    #本周实际完成poi产出量
    road_actual = 0   #本周实际完成road产出量
    poi_finished = 0    #poi版本已完成
    road_finished = 0   #road版本已完成
    per_info = []     #作业员信息
    st_info = []      #基地信息
    poi_names_all = []    #设施作业员姓名
    road_names_all = []   #道路作业员姓名
    startdate = ''        #统计开始日期
    enddate = ''          #统计结束日期
    week = ''             #第几周
    poi_sta_sum = []      #poi阶段统计综合
    road_sta_sum = []     #road阶段统计综合
    poi_day = []  # poi日统计
    poi_sum = []  # poi周统计
    road_day = []  # road 日统计
    road_sum = []  # road 周统计
    poi_plan = pd.DataFrame()  # poi 达标统计
    road_plan = pd.DataFrame()  # road 达标统计
    day_savefile_name = '每日与阶段统计' + time.strftime('%Y-%m-%d', time.localtime(time.time()))

    def __init__(self):
        # print('初始化......')
        super().__init__()
        self.initUI()

    # 初始化
    def ori_data(self):
        Sta_ex.station = ''  # 基地
        Sta_ex.version = ''  # 版本
        Sta_ex.poi_target = ''  # 设施版本目标
        Sta_ex.road_target = ''  # 道路版本目标
        Sta_ex.per_info = []  # 作业员信息
        Sta_ex.st_info = []  # 基地信息
        Sta_ex.poi_names_all = []  # 设施作业员姓名
        Sta_ex.road_names_all = []  # 道路作业员姓名
        Sta_ex.startdate = ''  # 统计开始日期
        Sta_ex.enddate = ''  # 统计结束日期
        Sta_ex.week = ''  # 第几周
        Sta_ex.poi_day = []      # poi日统计
        Sta_ex.poi_sum = []      # poi周统计
        Sta_ex.road_day = []     # road 日统计
        Sta_ex.road_sum = []     # road 周统计
        Sta_ex.poi_plan = pd.DataFrame()  #poi 达标统计
        Sta_ex.road_plan = pd.DataFrame() #road 达标统计
        self.poi_names.clear()
        self.road_names.clear()
        self.poi_name_sel.clear()
        self.road_name_sel.clear()

    def initUI(self):
        self.setGeometry(300, 300, 600, 250)
        self.setWindowTitle('数据统计V1.3.2')
        self.createGridGroupBox()
        self.creatVboxGroupBox()
        self.creatV_1boxGroupBox()
        # self.creatFormGroupBox()
        mainLayout = QVBoxLayout()
        hboxLayout = QHBoxLayout()
        # vboxLayout1 = QVBoxLayout()

        hboxLayout.addWidget(self.gridGroupBox)
        hboxLayout.addWidget(self.vboxGroupBox, stretch=0)
        mainLayout.addLayout(hboxLayout)
        mainLayout.addWidget(self.vboxGroupBox1)

        self.setLayout(mainLayout)
        self.sel_per_Bt.clicked.connect(self.per_sel)
        self.sel_poi_Bt.clicked.connect(self.poi_sel)
        self.sel_road_Bt.clicked.connect(self.road_sel)
        self.input_finish.clicked.connect(self.ori_data)       #数据初始化
        self.input_finish.clicked.connect(self.ps_info_Read)
        self.input_finish.clicked.connect(self.data_ex)
        self.input_finish.clicked.connect(self.path_exists)

        #每日与阶段统计
        self.com_Bt.clicked.connect(self.ori_data)  # 数据初始化
        self.com_Bt.clicked.connect(self.poi_Sta)
        self.com_Bt.clicked.connect(self.road_Sta)
        self.com_Bt.clicked.connect(self.poi_sta_plot)
        self.com_Bt.clicked.connect(self.road_sta_plot)
        self.com_Bt.clicked.connect(self.day_save_msg)

        # self.com_Bt.clicked.connect(self.path_exists)
        # self.com_Bt.clicked.connect(self.poi_Sta)
        # self.com_Bt.clicked.connect(self.road_Sta)
        # self.com_Bt.clicked.connect(self.save_path)   #日工作统计文件保存路径

        #周达标情况统计
        self.weekplan_Bt.clicked.connect(self.plan_path_exists)
        self.weekplan_Bt.clicked.connect(self.msg_plan_exit)
        self.weekplan_Bt.clicked.connect(self.plan_poi_Sta)
        self.weekplan_Bt.clicked.connect(self.plan_road_Sta)
        self.weekplan_Bt.clicked.connect(self.per_Plan)
        self.weekplan_Bt.clicked.connect(self.poi_per_plot)
        self.weekplan_Bt.clicked.connect(self.road_per_plot)
        self.weekplan_Bt.clicked.connect(self.per_plot)
        self.weekplan_Bt.clicked.connect(self.msg_plan_path_save)
        self.weekplan_Bt.clicked.connect(self.plan_savefile_msg)

        # self.weekplan_Bt.clicked.connect(self.msg_plan_exit)
        # self.weekplan_Bt.clicked.connect(self.per_Plan)
        # self.weekplan_Bt.clicked.connect(self.poi_per_plot)
        # self.weekplan_Bt.clicked.connect(self.road_per_plot)
        # self.weekplan_Bt.clicked.connect(self.per_plot)

        #个人情况统计
        self.Bt_poi_per.clicked.connect(self.poi_filepath_exist)
        self.Bt_poi_per.clicked.connect(self.per_poi_Sta)
        self.Bt_poi_per.clicked.connect(self.poi_filesave_msg)

        self.Bt_road_per.clicked.connect(self.road_filepath_exist)
        self.Bt_road_per.clicked.connect(self.per_road_Sta)
        self.Bt_road_per.clicked.connect(self.road_filesave_msg)

    def createGridGroupBox(self):
        self.gridGroupBox = QGroupBox("信息输入")
        layout = QGridLayout()
        layout.setSpacing(10)
        self.sel_per_Bt = QPushButton('选择基地人员基本信息表', self)
        self.per_Ed = QLineEdit('data/外业采集人员信息表.xlsx', self)

        self.sel_poi_Bt = QPushButton('选择设施输入文件', self)
        # self.poi_Ed = QLineEdit(self)
        self.poi_Ed = QLineEdit('data/设施.xlsx', self)

        self.sel_road_Bt = QPushButton('选择道路输入文件', self)
        # self.road_Ed = QLineEdit(self)
        self.road_Ed = QLineEdit('data/道路.xlsx', self)

        # self.date_Bt = QLabel('统计日期', self)
        # self.date_Bt.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        # self.startdate_Ed = QLineEdit('20170423', self)
        # # self.enddate_Bt = QPushButton('统计结束日期', self)
        # self.str_zhi = QLabel('至', self)
        # self.enddate_Ed = QLineEdit('20170510', self)
        # self.week_Bt = QLabel('第几周', self)
        # self.week_Bt.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        # self.week_Ed = QLineEdit('1',self)

        self.input_finish = QPushButton('输入完毕', self)

        # self.setGeometry(300, 300, 600, 250)
        # self.setWindowTitle('数据统计V1.3.1')

        layout.addWidget(self.sel_per_Bt, 0, 0)
        layout.addWidget(self.per_Ed, 0, 1, 1, 3)
        # 设施文件
        layout.addWidget(self.sel_poi_Bt, 1, 0)
        layout.addWidget(self.poi_Ed, 1, 1, 1, 3)
        # 道路文件
        layout.addWidget(self.sel_road_Bt, 2, 0)
        layout.addWidget(self.road_Ed, 2, 1, 1, 3)
        # 统计日期
        # layout.addWidget(self.date_Bt, 3, 0)
        # layout.addWidget(self.startdate_Ed, 3, 1)
        # layout.addWidget(self.str_zhi, 3, 2)
        # layout.addWidget(self.enddate_Ed, 3, 3)
        # 统计周数
        # layout.addWidget(self.week_Bt, 4, 0)
        # layout.addWidget(self.week_Ed, 4, 1)

        layout.addWidget(self.input_finish, 3, 0)

        self.gridGroupBox.setLayout(layout)

    def creatVboxGroupBox(self):
        self.vboxGroupBox = QGroupBox("基本信息")
        layout = QGridLayout()

        self.station_name = QLabel('版本信息：', self)
        self.station_info = QTextEdit(self)
        # self.station_info = QLineEdit(self)
        # station_info = QLabel(self)
        self.poi_name = QLabel('设施作业员：', self)
        # self.poi_names = QTextEdit(self)
        self.poi_names = QComboBox(self)
        self.road_name = QLabel('道路作业员：', self)
        # self.road_names = QTextEdit(self)
        self.road_names = QComboBox(self)
        layout.addWidget(self.station_name)
        layout.addWidget(self.station_info, 1, 0, 1, 2)
        layout.addWidget(self.poi_name, 2, 0)
        layout.addWidget(self.poi_names, 3, 0)
        layout.addWidget(self.road_name, 2, 1)
        layout.addWidget(self.road_names, 3, 1)

        self.vboxGroupBox.setLayout(layout)

    def creatV_1boxGroupBox(self):
        self.vboxGroupBox1 = QGroupBox('信息统计')
        layout = QGridLayout()
        date_index = QLabel('统计日期：')
        #日与阶段统计
        self.day_startday = QLineEdit('20170423', self)
        self.zhi_str1 = QLabel('至')
        self.day_endday = QLineEdit('20170510', self)
        self.com_Bt = QPushButton('每日与阶段统计', self)
        #周达标情况统计
        self.plan_startday = QLineEdit('20170423', self)
        self.zhi_str2 = QLabel('至')
        self.plan_endday = QLineEdit('20170510', self)
        weeklist = ['第1周', '第2周', '第3周', '第4周', '第5周', '第6周', '第7周', '第8周', '第9周', '第10周', '第11周', '第12周']
        self.weeklist_sel = QComboBox(self)
        self.weeklist_sel.addItems(weeklist)
        self.weekplan_Bt = QPushButton('周达标情况统计', self)
        #个人作业情况统计
        #设施
        self.poi_per_startday = QLineEdit('20170423', self)
        self.zhi_str3 = QLabel('至')
        self.poi_per_endday = QLineEdit('20170531', self)
        self.poi_name_sel = QComboBox(self)
        self.Bt_poi_per = QPushButton('个人设施作业统计', self)
        #道路
        self.road_per_startday = QLineEdit('20170423',self)
        self.zhi_str4 = QLabel('至')
        self.road_per_endday = QLineEdit('20170531',self)
        self.road_name_sel = QComboBox(self)
        self.Bt_road_per = QPushButton('个人道路作业统计', self)

        #基地版本完成情况
        self.station_sta_name = QLabel('基地版本完成情况：', self)
        self.station_sta_info = QTextEdit(self)

        # self.com_Bt = QPushButton('作业情况统计', self)
        # self.com_path = QTextEdit(self)
        # self.weekplan_Bt = QPushButton('周达标情况统计', self)
        # self.week_path = QTextEdit(self)
        # layout.addWidget(self.com_Bt)
        # layout.addWidget(self.com_path)
        # layout.addWidget(self.weekplan_Bt)
        # layout.addWidget(self.week_path)
        layout.addWidget(date_index)
        layout.addWidget(self.day_startday, 1, 0)
        layout.addWidget(self.zhi_str1, 1, 1)
        layout.addWidget(self.day_endday, 1, 2)
        layout.addWidget(self.com_Bt, 1, 4)

        layout.addWidget(self.plan_startday, 2, 0)
        layout.addWidget(self.zhi_str2, 2, 1)
        layout.addWidget(self.plan_endday, 2, 2)
        layout.addWidget(self.weeklist_sel, 2, 3)
        layout.addWidget(self.weekplan_Bt, 2, 4)

        layout.addWidget(self.poi_per_startday, 3, 0)
        layout.addWidget(self.zhi_str3, 3, 1)
        layout.addWidget(self.poi_per_endday, 3, 2)
        layout.addWidget(self.poi_name_sel, 3, 3)
        layout.addWidget(self.Bt_poi_per, 3, 4)

        layout.addWidget(self.road_per_startday, 4, 0)
        layout.addWidget(self.zhi_str4, 4, 1)
        layout.addWidget(self.road_per_endday, 4, 2)
        layout.addWidget(self.road_name_sel, 4, 3)
        layout.addWidget(self.Bt_road_per, 4, 4)

        layout.addWidget(self.station_sta_name, 5, 0)
        layout.addWidget(self.station_sta_info, 6, 0, 1, 5)

        self.vboxGroupBox1.setLayout(layout)

    #周达标情况统计
    def plan_savefile_msg(self):
        reply = QMessageBox.information(self,'统计完成','结果保存至'+'.../导出结果/'+ self.weeklist_sel.currentText()+'统计/')

    #个人设施与道路统计文件保存成功
    def poi_filesave_msg(self):
        reply = QMessageBox.information(self, '统计完成', '结果保存至'+'.../导出结果/'+ Sta_ex.day_savefile_name + '/'+
                                        self.poi_name_sel.currentText()+'/')
    def road_filesave_msg(self):
        reply = QMessageBox.information(self, '统计完成', '结果保存至'+'.../导出结果/'+ Sta_ex.day_savefile_name + '/'+
                                        self.road_name_sel.currentText()+'/')

    #个人设施统计文件夹存在检查
    def poi_filepath_exist(self):
        name = self.poi_name_sel.currentText()
        file_path = '导出结果/' + Sta_ex.day_savefile_name+'/'+ name
        if os.path.exists(file_path) == False:
            os.mkdir(file_path)

    # 个人道路统计文件夹存在检查
    def road_filepath_exist(self):
        name = self.road_name_sel.currentText()
        file_path = '导出结果/' + Sta_ex.day_savefile_name + '/' + name
        if os.path.exists(file_path) == False:
            os.mkdir(file_path)

    # 个人基本信息表选择
    def per_sel(self):
        per_dir = 'data'
        per_file, filetype = QFileDialog.getOpenFileName(self, "选取文件", per_dir,
                                                         "Excel Files (*.xlsx);; Excel Files (*.xls)")  # 设置文件扩展名过滤,注意用双分号间隔
        # print(per_file, type(per_file))
        self.per_Ed.setText(per_file)
    # POI文件选择
    def poi_sel(self):
        per_dir = 'data'
        per_file, filetype = QFileDialog.getOpenFileName(self, "选取文件", per_dir,
                                                         "Excel Files (*.xlsx);; Excel Files (*.xls)")  # 设置文件扩展名过滤,注意用双分号间隔
        self.poi_Ed.setText(per_file)
    # 道路文件选择
    def road_sel(self):
        per_dir = 'data'
        per_file, filetype = QFileDialog.getOpenFileName(self, "选取文件", per_dir,
                                                         "Excel Files (*.xlsx);; Excel Files (*.xls)")  # 设置文件扩展名过滤,注意用双分号间隔
        self.road_Ed.setText(per_file)

    #读取个人信息表
    def ps_info_Read(self):
        # fname = input("输入人员基本信息表：")
        fname = self.per_Ed.text()
        # print(fname)
        # 检查人员基本信息表是否存在
        if os.path.exists(fname) == False:
            print('人员基本信息表不存在，请确认！')
            self.msg_path_exist(fname)
            # os.mkdir('导出结果/' + filename)
        data = xlrd.open_workbook(fname)
        per_data = data.sheet_by_name('人员基本信息表')
        st_data = data.sheet_by_name('版本目标')

        for i in range(per_data.nrows):
            Sta_ex.per_info.append(per_data.row_values(i))
        for i in range(st_data.nrows):
            Sta_ex.st_info.append(st_data.row_values(i))
        Sta_ex.station = Sta_ex.st_info[1][Sta_ex.st_info[0].index('基地名称')]
        Sta_ex.version = Sta_ex.st_info[1][Sta_ex.st_info[0].index('版本')]
        Sta_ex.poi_target = Sta_ex.st_info[1][Sta_ex.st_info[0].index('设施产出总量目标')]
        Sta_ex.road_target = Sta_ex.st_info[1][Sta_ex.st_info[0].index('道路产出总量目标')]
        #版本信息展示
        station_str = '''%s%s 
设施版本产出量目标：%d
道路版本产出量目标：%d'''%(Sta_ex.station, Sta_ex.version, Sta_ex.poi_target, Sta_ex.road_target)
        self.station_info.setText(station_str)
        # print('测试。。。')
        #设施作业员
        per_name_index = Sta_ex.per_info[0].index('作业人员')
        for i in range(1, len(Sta_ex.per_info)):
            if 'poi' in Sta_ex.per_info[i][Sta_ex.per_info[0].index('作业范围')]:
                Sta_ex.poi_names_all.append(Sta_ex.per_info[i][per_name_index])
        # print('设施作业员：', Sta_ex.poi_names_all)
        # poinames_str = ','.join(Sta_ex.poi_names_all)
        # self.poi_names.setText(poinames_str)
        self.poi_names.addItems(Sta_ex.poi_names_all)

        #道路作业员
        per_name_index = Sta_ex.per_info[0].index('作业人员')
        for i in range(1, len(Sta_ex.per_info)):
            if 'road' in Sta_ex.per_info[i][Sta_ex.per_info[0].index('作业范围')]:
                # print('测试。。。')
                Sta_ex.road_names_all.append(Sta_ex.per_info[i][per_name_index])
        # print('设施作业员：', Sta_ex.poi_names_all)
        # roadnames_str = ','.join(Sta_ex.road_names_all)
        # self.road_names.setText(roadnames_str)
        self.road_names.addItems(Sta_ex.road_names_all)
        self.poi_name_sel.addItems(Sta_ex.poi_names_all)
        self.road_name_sel.addItems(Sta_ex.road_names_all)

        # Sta_ex.startdate = self.startdate_Ed.text()
        # Sta_ex.enddate = self.enddate_Ed.text()
        # Sta_ex.week = self.week_Ed.text()
        # if Sta_ex.startdate == '' or Sta_ex.enddate == '' or Sta_ex.week == '':
        #     self.msg()

    #根据版本与基地名进行数据初步提取
    def data_ex(self):
        #poi数据提取
        poi_fname = self.poi_Ed.text()
        poi_data = xlrd.open_workbook(poi_fname)
        poi_work = poi_data.sheets()[0]
        poi_time = poi_data.sheets()[1]
        poi_nrows = poi_work.nrows
        poi_time_nrows = poi_time.nrows
        poi_day_sta = []
        for i in range(poi_nrows):
            poi_day_sta.append(poi_work.row_values(i))
        poi_time_sta = []
        for i in range(poi_time_nrows):
            poi_time_sta.append(poi_time.row_values(i))

        # 根据基地与版本提取数据
        poi_day_station_index = poi_day_sta[0].index('外业基地')
        poi_time_station_index = poi_time_sta[0].index('外业基地')

        Sta_ex.poi_data_ex.append(poi_day_sta[0])  # 获取首行
        Sta_ex.poi_data_ex.append(poi_time_sta[0])  # 获取首行
        # print('测试。。。')
        #将本基地以及该版本的数据提取出来保存在*_ex文件中
        for i in range(1, len(poi_day_sta)):
            if poi_day_sta[i][poi_day_station_index] == Sta_ex.station and (Sta_ex.version in poi_day_sta[i][poi_day_sta[0].index('项目名称')]):
                Sta_ex.poi_data_ex.append(poi_day_sta[i])
        for i in range(1, len(poi_time_sta)):
            if poi_time_sta[i][poi_time_station_index] == Sta_ex.station and (Sta_ex.version in poi_time_sta[i][poi_time_sta[0].index('项目名称')]):
                Sta_ex.poi_time_ex.append(poi_time_sta[i])

        #road数据提取
        # fname = input('输入设道路数据文件名：')
        road_fname = self.road_Ed.text()
        road_data = xlrd.open_workbook(road_fname)
        road_work = road_data.sheets()[0]
        road_day_data = []
        road_nrows = road_work.nrows

        for i in range(road_nrows):
            road_day_data.append(road_work.row_values(i))
        for i in range(len(road_day_data)):
            Sta_ex.road_data_ex.append(road_day_data[i])

    #计算信息提示
    def msg_com(self):
        self.com_path.setText('计算中......')

    # 设施数据提取
    def poi_Sta(self):
        day_data = Sta_ex.poi_data_ex
        print(len(day_data))
        # 获取人数
        time_data = Sta_ex.poi_time_ex
        per_name_index = Sta_ex.per_info[0].index('作业人员')
        name = []
        for i in range(1, len(Sta_ex.per_info)):
            if 'poi' in Sta_ex.per_info[i][Sta_ex.per_info[0].index('作业范围')]:
                name.append(Sta_ex.per_info[i][per_name_index])
        # del name[0]
        print('设施作业员：', name)
        # for i in range(20):
        #     print(day_data[i])
        # 获取开始于结束日期
        # start_date = input('输入设施统计开始日期(格式：20170101)：')
        # end_date = input('输入设施统计结束日期：')
        start_date = self.day_startday.text()
        # end_date = input('输入道路统计结束日期：')
        end_date = self.day_endday.text()
        start_day = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:8]))
        end_day = datetime.date(int(end_date[0:4]), int(end_date[4:6]), int(end_date[6:8]))
        print('统计时间段：', start_day, '~', end_day)
        daytime = []
        if start_day > end_day:
            print('开始日期必须小于结束日期!!!!!!!!')
            #         os.exit()
        else:
            date_c = start_day
            date_str = ''
            while date_c <= end_day:
                date_str = date_c.strftime('%Y%m%d')
                daytime.append(date_str)
                date_c += datetime.timedelta(1)
        names = len(name)
        days = len(daytime)
        # 定义日统计矩阵

        poi_day = []
        for i in range(days * names):
            poi_day.append([])
            for j in range(12):
                poi_day[i].append(0)
        k = 0
        for i in range(names):
            for j in range(days):
                poi_day[k + j][0] = name[i]
                poi_day[k + j][1] = daytime[j]
            k = k + days
        k = 0
        for j in range(len(day_data)):
            for i in range(len(poi_day)):
                if day_data[j][5] == poi_day[i][0] and day_data[j][6] == poi_day[i][1]:
                    for k in range(6):
                        poi_day[i][k + 2] = day_data[j][k + 10] + poi_day[i][k + 2]
                poi_day[i][9] = poi_day[i][4] + poi_day[i][5] + poi_day[i][7]
                poi_day[i][10] = poi_day[i][5] + poi_day[i][6] + poi_day[i][7]
        for j in range(len(time_data)):
            for i in range(len(poi_day)):
                if time_data[j][2] == poi_day[i][0] and time_data[j][3] == poi_day[i][1]:
                    poi_day[i][8] = float(time_data[j][4]) + poi_day[i][8]
        day_list = ['作业人员', '日期', '实际完成规划量', '实际总作业量', '新增', '修改', '验证', '删除', '作业时长', '产出量', '工作量', '备注']
        week_list = ['作业人员', '实际完成规划量', '实际总作业量', '新增', '修改', '验证', '删除', '作业时长', '产出量', '工作量', '备注']
        poi_day.insert(0, day_list)

        # 定义周统计矩阵
        poi_sum = []
        for i in range(names):
            poi_sum.append([])
            for j in range(11):
                poi_sum[i].append(0)
        # 计算周统计结果
        for i in range(names):
            poi_sum[i][0] = name[i]
            for j in range(len(poi_day)):
                if name[i] == poi_day[j][0]:
                    for n in range(1, 11):
                        poi_sum[i][n] = poi_day[j][n + 1] + poi_sum[i][n]
        poi_sum.insert(0, week_list)
        Sta_ex.poi_sta_sum = poi_sum
        #本周实际完成产出量
        poi_actual = 0
        for i in range(1,len(poi_sum)):
            poi_actual = poi_actual + poi_sum[i][poi_sum[0].index('产出量')]

        #保存文件
        poi_savefile_name = '设施' + start_date + '-' + end_date + '.xls'
        file = xlwt.Workbook()
        sheet1 = file.add_sheet(u'个人每日统计', cell_overwrite_ok=True)
        sheet2 = file.add_sheet(u'个人阶段统计', cell_overwrite_ok=True)
        # 写入日统计
        for i in range(len(poi_day)):
            for j in range(len(poi_day[0])):
                sheet1.write(i, j, poi_day[i][j])
        # 写入阶段统计
        for i in range(len(poi_sum)):
            for j in range(len(poi_sum[0])):
                sheet2.write(i, j, poi_sum[i][j])

        path_save = '导出结果/'+ Sta_ex.day_savefile_name + '/' + poi_savefile_name
        file.save(path_save)

    #日统计与阶段统计绘图
    def poi_sta_plot(self):
        poi_sum = Sta_ex.poi_sta_sum
        # 转为DataFrame
        poi_sum_df = pd.DataFrame(poi_sum[1:],columns=poi_sum[0])
        poi_sum_df = poi_sum_df.set_index('作业人员')
        # print(poi_sum_df)
        if poi_sum_df.shape[0]<10:
            fig_width = 12
        else:
            fig_width = poi_sum_df.shape[0]
        fig_poi_sta = plt.figure(figsize=(fig_width, 8))
        ax1 = fig_poi_sta.add_subplot(1, 1, 1)
        poi_sum_df[['产出量','工作量']].plot(ax=ax1, kind='bar',legend=False)
        ax1.set_ylabel('产出量与工作量')
        handles1, labels1 = ax1.get_legend_handles_labels()  # 获取图一图例相关参数
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=-45)
        ax2 = ax1.twinx()
        poi_sum_df['作业时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^',color='y',legend=False)
        ax2.set_ylabel('作业时长（小时）')
        ax1.set_xlim(-0.5, poi_sum_df.shape[0] - 0.5)
        handles2, labels2 = ax2.get_legend_handles_labels()  # 获取图二图例相关参数
        # ax2.set_ylim(-1, 10)
        handles2.extend(handles1)
        labels2.extend(labels1)  # 图例参数合并
        plt.legend(handles2[::-1], labels2[::-1])  # 绘制图例
        plt.subplots_adjust(bottom=0.15, top=.95, left=.06, right=.94)  # 调整图像空白区域大小
        #保存图像
        start_date = self.day_startday.text()
        end_date = self.day_endday.text()
        fig_title = 'poi'+ start_date + '-' + end_date
        savefig_name = 'poi'+ start_date + '-' + end_date + '.png'
        ax1.set_title(fig_title)
        path_save = '导出结果/' + Sta_ex.day_savefile_name + '/' + savefig_name
        plt.savefig(path_save)
        # fig_poi_sta.show()

    #道路数据统计
    def road_Sta(self):  # 统计道路数据
        day_data = Sta_ex.road_data_ex
        nrows = len(day_data)
        # fname = self.road_Ed.text()
        # data = xlrd.open_workbook(fname)
        # work = data.sheets()[0]
        # nrows = work.nrows
        # ncols = work.ncols
        # day_data = []
        # for i in range(nrows):
        #     day_data.append(work.row_values(i))
        # 索引序号
        date_index = day_data[0].index('作业日期')
        name_index = day_data[0].index('作业员姓名')
        id_index = day_data[0].index('作业员ID')
        station_index = day_data[0].index('基地')
        time_index = day_data[0].index('有效时长')
        uproad_index = day_data[0].index('更新里程')
        newroad_index = day_data[0].index('新增里程')
        allroad_index = day_data[0].index('总作业里程')
        picnum_index = day_data[0].index('DCS图标量')
        # 获取姓名
        name = []
        per_name_index = Sta_ex.per_info[0].index('作业人员')
        for i in range(1, len(Sta_ex.per_info)):
            if 'road' in Sta_ex.per_info[i][Sta_ex.per_info[0].index('作业范围')]:
                name.append(Sta_ex.per_info[i][per_name_index])

        names = len(name)
        print('道路作业员：', name)
        # 获取作业日期
        # start_date = input('输入道路统计开始日期(格式：20170101)：')
        start_date = self.day_startday.text()
        # end_date = input('输入道路统计结束日期：')
        end_date = self.day_endday.text()
        start_day = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:8]))
        end_day = datetime.date(int(end_date[0:4]), int(end_date[4:6]), int(end_date[6:8]))
        print('统计时间段：', start_day, '~', end_day)
        daytime = []
        if start_day > end_day:
            print('开始日期必须小于结束日期!!!!!!!!')
        else:
            date_c = start_day
            date_str = ''
            while date_c <= end_day:
                date_str = date_c.strftime('%Y-%m-%d')
                daytime.append(date_str)
                date_c += datetime.timedelta(1)
        dates = len(daytime)

        # 定义存储矩阵
        road_day = []
        for i in range(names * dates):
            road_day.append([])
            for j in range(8):
                road_day[i].append(0)
        k = 0
        for i in range(names):
            for j in range(dates):
                road_day[k + j][0] = name[i]
                road_day[k + j][1] = daytime[j]
            k = k + dates
        day_list = [day_data[0][name_index], day_data[0][date_index], day_data[0][time_index],
                    day_data[0][uproad_index], day_data[0][newroad_index], day_data[0][allroad_index],
                    day_data[0][picnum_index], '产出量']
        # 统计日记录
        road_day.insert(0, day_list)
        rows = len(road_day)
        for i in range(1, rows):
            for j in range(1, nrows):
                if road_day[i][0] == day_data[j][name_index] and road_day[i][1] == day_data[j][date_index]:
                    road_day[i][2] = float(day_data[j][time_index]) + road_day[i][2]
                    road_day[i][3] = float(day_data[j][uproad_index]) + road_day[i][3]
                    road_day[i][4] = float(day_data[j][newroad_index]) + road_day[i][4]
                    road_day[i][5] = float(day_data[j][allroad_index]) + road_day[i][5]
                    road_day[i][6] = float(day_data[j][picnum_index]) + road_day[i][6]
                    road_day[i][7] = road_day[i][6]

        # 统计阶段记录
        road_sum = []
        week_list = day_list[:]
        del week_list[1]
        for i in range(names):
            road_sum.append([])
            for j in range(len(week_list)):
                road_sum[i].append(0)
        for i in range(names):
            road_sum[i][0] = name[i]
            for j in range(1, len(road_day)):
                if road_sum[i][0] == road_day[j][road_day[0].index('作业员姓名')]:
                    for k in range(2, 8):
                        road_sum[i][k - 1] = road_day[j][k] + road_sum[i][k - 1]
        road_sum.insert(0, week_list)
        Sta_ex.road_sta_sum= road_sum

        #本周实际完成道路产出量
        road_actual = 0
        for i in range(1,len(road_sum)):
            road_actual = road_actual + road_sum[i][road_sum[0].index('产出量')]

        road_savefile_name = '道路' + start_date + '-' + end_date + '.xls'
        file = xlwt.Workbook()
        sheet1 = file.add_sheet(u'个人每日统计', cell_overwrite_ok=True)
        sheet2 = file.add_sheet(u'个人阶段统计', cell_overwrite_ok=True)
        # 写入日统计
        for i in range(len(road_day)):
            for j in range(len(road_day[0])):
                sheet1.write(i, j, road_day[i][j])
        # 写入阶段统计
        for i in range(len(road_sum)):
            for j in range(len(road_sum[0])):
                sheet2.write(i, j, road_sum[i][j])

        path_save = '导出结果/' + Sta_ex.day_savefile_name + '/' + road_savefile_name
        file.save(path_save)

    #道路日统计与阶段统计绘图
    def road_sta_plot(self):
        road_sum = Sta_ex.road_sta_sum
        # 转为DataFrame
        road_sum_df = pd.DataFrame(road_sum[1:],columns=road_sum[0])
        # print(road_sum_df)
        road_sum_df = road_sum_df.set_index('作业员姓名')
        if road_sum_df.shape[0]<10:
            fig_width = 12
        else:
            fig_width = road_sum_df.shape[0]
        fig2 = plt.figure(figsize=(fig_width, 8))
        ax1 = fig2.add_subplot(1, 1, 1)
        road_sum_df[['产出量','总作业里程']].plot(ax=ax1, kind='bar', legend=False)
        ax1.set_ylabel('产出量与总里程')
        handles1, labels1 = ax1.get_legend_handles_labels()  # 获取图一图例相关参数
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=-45)
        ax2 = ax1.twinx()
        road_sum_df['有效时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^',color='y',legend=False)
        ax2.set_ylabel('有效作业时长（小时）')
        ax1.set_xlim(-0.5, road_sum_df.shape[0] - 0.5)
        handles2, labels2 = ax2.get_legend_handles_labels()  # 获取图二图例相关参数
        # ax2.set_ylim(-1, 10)
        handles2.extend(handles1)
        labels2.extend(labels1)  # 图例参数合并
        plt.legend(handles2[::-1], labels2[::-1])  # 绘制图例
        plt.subplots_adjust(bottom=0.15, top=.95, left=.06, right=.94)  # 调整图像空白区域大小
        #保存图像
        start_date = self.day_startday.text()
        end_date = self.day_endday.text()
        fig_title = 'road'+ start_date + '-' + end_date
        savefig_name = 'road'+ start_date + '-' + end_date + '.png'
        ax1.set_title(fig_title)
        path_save = '导出结果/' + Sta_ex.day_savefile_name + '/' + savefig_name
        plt.savefig(path_save)
        plt.show()

    # 周达标之前的poi统计
    def plan_poi_Sta(self):
        day_data = Sta_ex.poi_data_ex
        # print(len(day_data))
        # 获取人数
        time_data = Sta_ex.poi_time_ex
        per_name_index = Sta_ex.per_info[0].index('作业人员')
        name = []
        for i in range(1, len(Sta_ex.per_info)):
            if 'poi' in Sta_ex.per_info[i][Sta_ex.per_info[0].index('作业范围')]:
                name.append(Sta_ex.per_info[i][per_name_index])
        # del name[0]
        print('设施作业员：', name)
        # for i in range(20):
        #     print(day_data[i])
        # 获取开始于结束日期
        # start_date = input('输入设施统计开始日期(格式：20170101)：')
        # end_date = input('输入设施统计结束日期：')
        start_date = self.plan_startday.text()
        # end_date = input('输入道路统计结束日期：')
        end_date = self.plan_endday.text()
        start_day = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:8]))
        end_day = datetime.date(int(end_date[0:4]), int(end_date[4:6]), int(end_date[6:8]))
        print('统计时间段：', start_day, '~', end_day)
        daytime = []
        if start_day > end_day:
            print('开始日期必须小于结束日期!!!!!!!!')
            #         os.exit()
        else:
            date_c = start_day
            date_str = ''
            while date_c <= end_day:
                date_str = date_c.strftime('%Y%m%d')
                daytime.append(date_str)
                date_c += datetime.timedelta(1)
        names = len(name)
        days = len(daytime)
        # 定义日统计矩阵

        poi_day = []
        for i in range(days * names):
            poi_day.append([])
            for j in range(12):
                poi_day[i].append(0)
        k = 0
        for i in range(names):
            for j in range(days):
                poi_day[k + j][0] = name[i]
                poi_day[k + j][1] = daytime[j]
            k = k + days
        k = 0
        for j in range(len(day_data)):
            for i in range(len(poi_day)):
                if day_data[j][5] == poi_day[i][0] and day_data[j][6] == poi_day[i][1]:
                    for k in range(6):
                        poi_day[i][k + 2] = day_data[j][k + 10] + poi_day[i][k + 2]
                poi_day[i][9] = poi_day[i][4] + poi_day[i][5] + poi_day[i][7]
                poi_day[i][10] = poi_day[i][5] + poi_day[i][6] + poi_day[i][7]
        for j in range(len(time_data)):
            for i in range(len(poi_day)):
                if time_data[j][2] == poi_day[i][0] and time_data[j][3] == poi_day[i][1]:
                    poi_day[i][8] = float(time_data[j][4]) + poi_day[i][8]
        day_list = ['作业人员', '日期', '实际完成规划量', '实际总作业量', '新增', '修改', '验证', '删除', '作业时长', '产出量', '工作量', '备注']
        week_list = ['作业人员', '实际完成规划量', '实际总作业量', '新增', '修改', '验证', '删除', '作业时长', '产出量', '工作量', '备注']
        poi_day.insert(0, day_list)
        Sta_ex.poi_day = poi_day.copy()  #赋值到全局变量
        # 定义周统计矩阵
        poi_sum = []
        for i in range(names):
            poi_sum.append([])
            for j in range(11):
                poi_sum[i].append(0)
        # 计算周统计结果
        for i in range(names):
            poi_sum[i][0] = name[i]
            for j in range(len(poi_day)):
                if name[i] == poi_day[j][0]:
                    for n in range(1, 11):
                        poi_sum[i][n] = poi_day[j][n + 1] + poi_sum[i][n]
        poi_sum.insert(0, week_list)
        Sta_ex.poi_sum = poi_sum.copy()
        #本周实际完成产出量
        poi_actual = 0
        for i in range(1,len(poi_sum)):
            poi_actual = poi_actual + poi_sum[i][poi_sum[0].index('产出量')]

        #保存文件
        poi_savefile_name = '设施' + start_date + '-' + end_date + '.xls'
        file = xlwt.Workbook()
        sheet1 = file.add_sheet(u'个人每日统计', cell_overwrite_ok=True)
        sheet2 = file.add_sheet(u'个人阶段统计', cell_overwrite_ok=True)
        # 写入日统计
        for i in range(len(poi_day)):
            for j in range(len(poi_day[0])):
                sheet1.write(i, j, poi_day[i][j])
        # 写入阶段统计
        for i in range(len(poi_sum)):
            for j in range(len(poi_sum[0])):
                sheet2.write(i, j, poi_sum[i][j])
        path_save = '导出结果/' + '第' + Sta_ex.week + '周统计/' + poi_savefile_name
        # path_save = '导出结果/'+ Sta_ex.day_savefile_name + '/' + poi_savefile_name
        file.save(path_save)

    #周达标之前的road培训
    def plan_road_Sta(self):  # 统计道路数据
        day_data = Sta_ex.road_data_ex
        nrows = len(day_data)
        # fname = self.road_Ed.text()
        # data = xlrd.open_workbook(fname)
        # work = data.sheets()[0]
        # nrows = work.nrows
        # ncols = work.ncols
        # day_data = []
        # for i in range(nrows):
        #     day_data.append(work.row_values(i))
        # 索引序号
        date_index = day_data[0].index('作业日期')
        name_index = day_data[0].index('作业员姓名')
        id_index = day_data[0].index('作业员ID')
        station_index = day_data[0].index('基地')
        time_index = day_data[0].index('有效时长')
        uproad_index = day_data[0].index('更新里程')
        newroad_index = day_data[0].index('新增里程')
        allroad_index = day_data[0].index('总作业里程')
        picnum_index = day_data[0].index('DCS图标量')
        # 获取姓名
        name = []
        per_name_index = Sta_ex.per_info[0].index('作业人员')
        for i in range(1, len(Sta_ex.per_info)):
            if 'road' in Sta_ex.per_info[i][Sta_ex.per_info[0].index('作业范围')]:
                name.append(Sta_ex.per_info[i][per_name_index])

        names = len(name)
        print('道路作业员：', name)
        # 获取作业日期
        # start_date = input('输入道路统计开始日期(格式：20170101)：')
        start_date = self.plan_startday.text()
        # end_date = input('输入道路统计结束日期：')
        end_date = self.plan_endday.text()
        start_day = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:8]))
        end_day = datetime.date(int(end_date[0:4]), int(end_date[4:6]), int(end_date[6:8]))
        print('统计时间段：', start_day, '~', end_day)
        daytime = []
        if start_day > end_day:
            print('开始日期必须小于结束日期!!!!!!!!')
        else:
            date_c = start_day
            date_str = ''
            while date_c <= end_day:
                date_str = date_c.strftime('%Y-%m-%d')
                daytime.append(date_str)
                date_c += datetime.timedelta(1)
        dates = len(daytime)

        # 定义存储矩阵
        road_day = []
        for i in range(names * dates):
            road_day.append([])
            for j in range(8):
                road_day[i].append(0)
        k = 0
        for i in range(names):
            for j in range(dates):
                road_day[k + j][0] = name[i]
                road_day[k + j][1] = daytime[j]
            k = k + dates
        day_list = [day_data[0][name_index], day_data[0][date_index], day_data[0][time_index],
                    day_data[0][uproad_index], day_data[0][newroad_index], day_data[0][allroad_index],
                    day_data[0][picnum_index], '产出量']
        # 统计日记录
        road_day.insert(0, day_list)
        rows = len(road_day)
        for i in range(1, rows):
            for j in range(1, nrows):
                if road_day[i][0] == day_data[j][name_index] and road_day[i][1] == day_data[j][date_index]:
                    road_day[i][2] = float(day_data[j][time_index]) + road_day[i][2]
                    road_day[i][3] = float(day_data[j][uproad_index]) + road_day[i][3]
                    road_day[i][4] = float(day_data[j][newroad_index]) + road_day[i][4]
                    road_day[i][5] = float(day_data[j][allroad_index]) + road_day[i][5]
                    road_day[i][6] = float(day_data[j][picnum_index]) + road_day[i][6]
                    road_day[i][7] = road_day[i][6]

        Sta_ex.road_day = road_day.copy()
        # 统计阶段记录
        road_sum = []
        week_list = day_list[:]
        del week_list[1]
        for i in range(names):
            road_sum.append([])
            for j in range(len(week_list)):
                road_sum[i].append(0)
        for i in range(names):
            road_sum[i][0] = name[i]
            for j in range(1, len(road_day)):
                if road_sum[i][0] == road_day[j][road_day[0].index('作业员姓名')]:
                    for k in range(2, 8):
                        road_sum[i][k - 1] = road_day[j][k] + road_sum[i][k - 1]
        road_sum.insert(0, week_list)
        Sta_ex.road_sum = road_sum.copy()
        #本周实际完成道路产出量
        road_actual = 0
        for i in range(1,len(road_sum)):
            road_actual = road_actual + road_sum[i][road_sum[0].index('产出量')]

        road_savefile_name = '道路' + start_date + '-' + end_date + '.xls'
        file = xlwt.Workbook()
        sheet1 = file.add_sheet(u'个人每日统计', cell_overwrite_ok=True)
        sheet2 = file.add_sheet(u'个人阶段统计', cell_overwrite_ok=True)
        # 写入日统计
        for i in range(len(road_day)):
            for j in range(len(road_day[0])):
                sheet1.write(i, j, road_day[i][j])
        # 写入阶段统计
        for i in range(len(road_sum)):
            for j in range(len(road_sum[0])):
                sheet2.write(i, j, road_sum[i][j])

        path_save = '导出结果/' + '第' + Sta_ex.week + '周统计/' + road_savefile_name
        file.save(path_save)

    #周达标情况统
    def per_Plan(self):
        n = 7         #统计时间段间隔
        #基地基本情况表转换为字典
        st_info = dict()
        for i in range(len(Sta_ex.st_info[0])):
            st_info[Sta_ex.st_info[0][i]] = Sta_ex.st_info[1][i]
        weeks = int(st_info['版本时长']/n)
        per_info = pd.DataFrame(Sta_ex.per_info[1:], columns=Sta_ex.per_info[0])  #人员基本信息表转换
        week_start = self.plan_startday.text()
        week_end = self.plan_endday.text()
        week_start_date = datetime.date(int(week_start[0:4]), int(week_start[4:6]), int(week_start[6:8]))
        week_end_date = datetime.date(int(week_end[0:4]), int(week_end[4:6]), int(week_end[6:8]))
        print('统计时间段：', week_start_date, '~', week_end_date)
        date_range = pd.date_range(week_start_date, week_end_date)
        date_range_poi = []
        date_range_road = []

        for i in range(len(date_range)):
            date_range_poi.append(int(date_range[i].strftime('%Y%m%d')))
        for i in range(len(date_range)):
            date_range_road.append(date_range[i].strftime('%Y-%m-%d'))
        # t_n = input('输入统计第几周数据：')
        t_n = Sta_ex.week
        t_next = str(int(t_n) + 1)
        t_pre = str(int(t_n) - 1)
        str1 = '第' + t_n + '周总目标量'
        str2 = '第' + t_n + '周日目标量'
        str3 = '第' + t_n + '周实际完成'
        str4 = '第' + t_n + '周达标情况'
        str5 = '版本剩余量'
        str6 = '版本完成比例'
        str7_1 = '第' + t_next + '周总目标量'
        str7_2 = '第' + t_n + '周总目标量'
        str8 = '第' + t_next + '周日目标量'

        #读取日道路设施统计结果
        # poi_day_name = input('输入设施日工作文件名：')
        # print(week_start)
        poi_day_name = '设施' + week_start + '-' + week_end + '.xls'
        # print('测试。。。')
        # poi_day_name = '设施20170301-20170423.xls'
        poi_day = pd.read_excel('导出结果/' + '第'+t_n+'周统计/'+ poi_day_name, sheetname='个人每日统计', header=0)
        poi_per_plan = per_info[['作业人员', '作业范围', '组号', '设施工天', '设施版本总目标量']].copy()
        poi_plan = poi_per_plan[poi_per_plan['作业范围'].isin(['poi', 'poi_road'])]  #选取设施作业员个人信息
        # print(poi_plan)
        poi_plan[str3] = [0]*len(poi_plan)
        poi_plan[str4] = ['NA']*len(poi_plan)
        #poi统计
        for i in range(len(poi_plan)):
            namei = poi_plan['作业人员'].iloc[i]
            name_b = poi_day['作业人员'].isin([namei])   #按作业人员筛选数据
            poi_day_nameb = poi_day[name_b].copy()
            date_b = poi_day_nameb['日期'].isin(date_range_poi)  #按工作日期筛选数据
            poi_day_dateb = poi_day_nameb[date_b].copy()
            poi_plan[str3].iloc[i] = poi_day_dateb['产出量'].sum()

        savepoi_name_1 = '第' + t_n + '周设施达标情况统计.xls'
        savepoi_name_2 = '第' + t_pre + '周设施达标情况统计.xls'

        if int(t_n) == 1:
            poi_plan[str1] = poi_plan['设施版本总目标量'] / weeks
            poi_plan[str2] = poi_plan[str1] / 6
            poi_plan[str5] = poi_plan['设施版本总目标量'] - poi_plan[str3]
            poi_plan_list = ['作业人员', '作业范围', '组号', '设施版本总目标量', str1, str2, str3, str4, str5, str6, str7_1, str8]
        else:
            sheet_name = '第' + t_pre + '周设施达标情况'
            plan_preweek_poi = pd.read_excel('导出结果/' +'第'+t_pre+'周统计/'+ savepoi_name_2, sheetname=sheet_name)
            poi_plan[str1] = plan_preweek_poi[str7_2].values.tolist()
            # print(plan_preweek_poi[str7_2])
            poi_plan[str5] = (plan_preweek_poi[str5].values - poi_plan[str3].values).tolist()
            poi_plan_list = ['作业人员', '作业范围', '组号', '设施版本总目标量', str1, str3, str4, str5, str6, str7_1, str8]

        for i in range(len(poi_plan)):
            if poi_plan[str1].iloc[i] <= poi_plan[str3].iloc[i]:
                poi_plan[str4].iloc[i] = '达标'
            else:
                poi_plan[str4].iloc[i] = '未达标'
        # print('poi_plan:')
        # print(poi_plan)
        # print('poi_per_plan:')
        # print(poi_per_plan)

        poi_plan[str6] = ((poi_plan['设施版本总目标量'].values-poi_plan[str5].values)/poi_plan['设施版本总目标量'].values).tolist()   #完成比例
        poi_plan[str7_1] = poi_plan[str5]/(weeks-int(t_n))
        poi_plan[str8] = poi_plan[str7_1]/6
        poi_plan = poi_plan[poi_plan_list]    #.sort([str6], ascending=False)  #排序
        # 数据格式转换
        poi_plan['组号'] = poi_plan['组号'].astype('int')
        poi_plan['设施版本总目标量'] = poi_plan['设施版本总目标量'].astype('int')
        poi_plan[str1] = poi_plan[str1].astype('int')
        if int(t_n) == 1:
            poi_plan[str2] = poi_plan[str2].astype('int')
        poi_plan[str3] = poi_plan[str3].astype('int')
        poi_plan[str5] = poi_plan[str5].astype('int')
        poi_plan[str7_1] = poi_plan[str7_1].astype('int')
        poi_plan[str8] = poi_plan[str8].astype('int')
        poi_per_target_sum = poi_plan['设施版本总目标量'].sum()
        poi_per_remain_sum = poi_plan['版本剩余量'].sum()
        Sta_ex.poi_finished = poi_per_target_sum - poi_per_remain_sum
        # poi数据存储
        sheet_name = '第' + t_n + '周设施达标情况'
        poi_plan.to_excel('导出结果/'+ '第'+t_n+'周统计/' + savepoi_name_1, sheet_name=sheet_name, index=False)
        print('设施达标情况保存成功！')

        #道路数据统计************************************************************
        # road_day_name = input('输入道路日工作文件名：')
        road_day_name = '道路' + week_start + '-' + week_end + '.xls'
        # road_day_name = '道路20170301-20170423.xls'
        road_day = pd.read_excel('导出结果/' + '第'+t_n+'周统计/'+ road_day_name, sheetname='个人每日统计', header=0)
        road_per_plan = per_info[['作业人员', '作业范围', '组号', '道路工天', '道路版本总目标量']].copy()
        road_plan = road_per_plan[road_per_plan['作业范围'].isin(['poi_road', 'road'])]
        # road_plan[str1] = [0] * len(road_plan)
        road_plan[str3] = [0] * len(road_plan)
        road_plan[str4] = ['NA'] * len(road_plan)
        road_plan[str5] = [0] * len(road_plan)
        # print('测试。。。')
        for i in range(len(road_plan)):
            namei = road_plan['作业人员'].iloc[i]
            name_b = road_day['作业员姓名'].isin([namei])  # 按作业人员筛选数据
            road_day_nameb = road_day[name_b].copy()
            date_b = road_day_nameb['作业日期'].isin(date_range_road)  # 按工作日期筛选数据
            road_day_dateb = road_day_nameb[date_b].copy()
            road_plan[str3].iloc[i] = road_day_dateb['产出量'].sum()

        saveroad_name_1 = '第' + t_n + '周道路达标情况统计.xls'
        saveroad_name_2 = '第' + t_pre + '周道路达标情况统计.xls'
        if int(t_n) == 1:
            road_plan[str1] = road_plan['道路版本总目标量'] / weeks
            road_plan[str2] = road_plan[str1] / 6
            road_plan[str5] = road_plan['道路版本总目标量'] - road_plan[str3]
            road_plan_list = ['作业人员', '作业范围', '组号', '道路版本总目标量', str1, str2, str3, str4, str5, str6, str7_1, str8]
        else:
            sheet_name = '第' + t_pre + '周道路达标情况'
            plan_preweek_road = pd.read_excel('导出结果/' + '第'+t_pre+'周统计/'+saveroad_name_2, sheetname=sheet_name)
            road_plan[str1] = plan_preweek_road[str7_2].values.tolist()
            road_plan[str5] = (plan_preweek_road[str5].values - road_plan[str3].values).tolist()
            # for i in range(len(road_plan)):
            #     road_plan[str5].iloc[i] = plan_preweek_road[str5].iloc[i] - road_plan[str3].iloc[i]
            # road_plan[str5] = plan_preweek_road[str5] - road_plan[str3]
            # print(road_plan[str1])
            road_plan_list = ['作业人员', '作业范围', '组号', '道路版本总目标量', str1, str3, str4, str5, str6, str7_1, str8]

        for i in range(len(road_plan)):
            if road_plan[str1].iloc[i] <= road_plan[str3].iloc[i]:
                road_plan[str4].iloc[i] = '达标'
            else:
                road_plan[str4].iloc[i] = '未达标'

        road_plan[str6] = (road_plan['道路版本总目标量']-road_plan[str5])/road_plan['道路版本总目标量']   #完成比例
        road_plan[str7_1] = road_plan[str5]/(weeks-int(t_n))
        road_plan[str8] = road_plan[str7_1]/6
        road_plan = road_plan[road_plan_list]   #.sort([str6], ascending=False)  #排序
        # 数据格式转换
        road_plan['组号'] = road_plan['组号'].astype('int')
        road_plan['道路版本总目标量'] = road_plan['道路版本总目标量'].astype('int')
        road_plan[str1] = road_plan[str1].astype('int')
        if int(t_n) == 1:
            road_plan[str2] = road_plan[str2].astype('int')
        road_plan[str3] = road_plan[str3].astype('int')
        road_plan[str5] = road_plan[str5].astype('int')
        road_plan[str7_1] = road_plan[str7_1].astype('int')
        road_plan[str8] = road_plan[str8].astype('int')
        # print(road_plan)
        #road版本已完成量计算
        road_per_target_sum = road_plan['道路版本总目标量'].sum()
        road_per_remain_sum = road_plan['版本剩余量'].sum()
        Sta_ex.road_finished = road_per_target_sum - road_per_remain_sum

        # road数据存储
        sheet_name = '第' + t_n + '周道路达标情况'
        road_plan.to_excel('导出结果/'+ '第'+t_n+'周统计/' + saveroad_name_1, sheet_name=sheet_name, index=False)
        print('道路达标情况保存成功！')

        Sta_ex.poi_plan = poi_plan.copy()
        Sta_ex.road_plan = road_plan.copy()

    #poi绘图
    def poi_per_plot(self):
        # poi_plan = Sta_ex.poi_plan.copy()
        # road_plan = Sta_ex.road_plan.copy()
        # data = plan_file.copy()
        data = Sta_ex.poi_plan.copy()
        # print('读取数据：')
        # print(data)
        data_list = list(data)
        if data_list[3].find('设施')== 0:
            str = '设施'
        elif data_list[3].find('道路') == 0:
            str = '道路'
        week = data_list[4][0:3]
        # print(str, week)
        str1 = str + '版本总目标量'
        str2 = '版本剩余量'
        str3 = '作业人员'
        data_sort = data.sort([str1], ascending=False)
        # width = 0.5
        x_range = ny.arange(len(data_sort))
        b_wch = data_sort[str1].values - data_sort[str2].values
        # fig_width = len(data) * 2 / 3
        # fig_height = len(data) / 3
        # plt.figure(figsize=(fig_width, fig_height))
        plt.figure(figsize=(15, 8))
        plt.bar(x_range, data_sort[str1].values, color='b', label='版本总目标量')
        plt.bar(x_range, b_wch, color='r', label='已完成量')
        plt.xticks(x_range, data_sort[str3].values)
        plt.title(week + str+'作业情况统计')
        plt.legend()
        # plt.show()
        plt.savefig('导出结果/' + '第'+Sta_ex.week+'周统计/' + week + str +'作业情况统计.png')
        print('图片生成成功')

    # poi绘图
    def road_per_plot(self):
        # poi_plan = Sta_ex.poi_plan.copy()
        # road_plan = Sta_ex.road_plan.copy()
        # data = plan_file.copy()
        data = Sta_ex.road_plan.copy()
        # print('读取数据：')
        # print(data)
        data_list = list(data)
        if data_list[3].find('设施') == 0:
            str = '设施'
        elif data_list[3].find('道路') == 0:
            str = '道路'
        week = data_list[4][0:3]
        print(str, week)
        str1 = str + '版本总目标量'
        str2 = '版本剩余量'
        str3 = '作业人员'
        data_sort = data.sort([str1], ascending=False)
        # width = 0.5
        x_range = ny.arange(len(data_sort))
        b_wch = data_sort[str1].values - data_sort[str2].values
        # fig_poi = plt.figure(figsize=(15, 8))
        # fig_width = len(data) * 2 / 3
        # fig_height = len(data) / 3
        # plt.figure(figsize=(fig_width, fig_height))
        plt.figure(figsize=(15, 8))
        plt.bar(x_range, data_sort[str1].values, color='b', label='版本总目标量')
        plt.bar(x_range, b_wch, color='r', label='已完成量')
        plt.xticks(x_range, data_sort[str3].values)
        plt.title(week + str + '作业情况统计')
        plt.legend()
        # plt.show()
        plt.savefig('导出结果/' + '第'+Sta_ex.week+'周统计/' + week + str +'作业情况统计.png')
        print('图片生成成功')

    def per_plot(self):
        self.per_plan_Plot('poi')
        self.per_plan_Plot('road')

    def per_plan_Plot(self, str_pr):
        # print('个人作业情况图绘制:')
        per_info = Sta_ex.per_info  # 作业人员信息表
        if str_pr == 'poi':
            data_day = Sta_ex.poi_day  # poi日工作量统计
            data_sum = Sta_ex.poi_sum  # poi阶段统计
            data_plan = Sta_ex.poi_plan  # poi周达标情况统计
            date_index_str = '日期'  # 引导字段名不一样，poi为日期，道路为作业日期
            # name_str = '作业人员'
        elif str_pr == 'road':
            data_day = Sta_ex.road_day  # road日工作量统计
            data_sum = Sta_ex.road_sum  # road阶段统计
            data_plan = Sta_ex.road_plan  # road周达标情况统计
            date_index_str = '作业日期'
            # name_str = '作业员姓名'
            data_day[0][data_day[0].index('作业员姓名')] = '作业人员'
            data_day[0][data_day[0].index('有效时长')] = '作业时长'
        # for line in data_day:
        #     print(line)

        week = Sta_ex.week  # 周数
        # print(len(data_day[0]))
        # print('测试。。。')
        data_day_df = pd.DataFrame(data_day[1:], columns=data_day[0])  # 序列转为DataFrame
        # print(data_day_df)
        # print(data_day_df.head())
        # 提取作业员姓名
        per_name_index = per_info[0].index('作业人员')
        data_name = []
        for i in range(1, len(per_info)):
            if str_pr in per_info[i][per_info[0].index('作业范围')]:
                data_name.append(per_info[i][per_name_index])
        # print(data_day_df[data_day_df.作业人员==data_name[0]])
        # 引导字段名不一样，poi为日期，道路为作业日期
        # if str_pr=='poi':
        #     date_index_str = '日期'
        # elif str_pr == 'road':
        #     date_index_str = '作业日期'
        data_day_df = data_day_df.set_index(date_index_str)
        poi_plan_index = list(data_plan)
        week_n = str(int(week) + 1)
        str1 = '第' + week + '周总目标量'
        str2 = '第' + week + '周达标情况'
        str3 = '第' + week + '周实际完成'
        str4 = '第' + week_n + '周总目标量'
        str5 = '第' + week_n + '周日目标量'
        str6 = '版本完成比例'
        str7 = '版本剩余量'
        days = len(data_day_df[data_day_df.作业人员 == data_name[0]]['产出量'].values)
        # print(days)
        # print(str1)

        for i in range(len(data_name)):
            fig = plt.figure(figsize=(12, 6))
            ax1 = fig.add_subplot(1, 1, 1)
            name = data_name[i]
            target_day = data_plan[data_plan.作业人员 == name][str1].values / 6
            # print(type(target_day))
            data_day_df[data_day_df.作业人员 == name]['产出量'].plot(ax=ax1, kind='bar', subplots=True, style='-s')
            # hline绘制直线
            # plt.axhline(y=target_day, color='r')
            ax1.axhline(y=target_day, linestyle='-.', color='r', label='日目标')
            handles1, labels1 = ax1.get_legend_handles_labels()  # 获取图一图例相关参数
            ax1.set_ylabel('产出量')
            # l1 = ax1.legend(loc=2)
            ax1.set_xlabel('作业日期')
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=-45)
            ax2 = ax1.twinx()
            data_day_df[data_day_df.作业人员 == name]['作业时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^',
                                                               color='y')
            # if str_pr == 'poi':
            #     data_day_df[data_day_df.作业人员 == name]['作业时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^', color='y')
            # elif str_pr == 'road':
            #     data_day_df[data_day_df.作业人员 == name]['有效时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^', color='y')
            handles2, labels2 = ax2.get_legend_handles_labels()  # 获取图二图例相关参数
            ax2.set_ylabel('作业时长(小时)')
            ax1.set_xlim(-0.5, days - 0.5)
            ax1.set_title(name + '-' + str_pr + '第' + week + '周')
            ax2.set_ylim(-1, 10)
            handles2.extend(handles1)
            labels2.extend(labels1)  # 图例参数合并
            plt.legend(handles2[::-1], labels2[::-1])  # 绘制图例
            ax1.grid(color='b', linewidth='0.2', linestyle='--')  # 绘制格网
            # plt.tight_layout()
            plt.subplots_adjust(bottom=0.15, top=.95, left=.06, right=.86)  # 调整图像空白区域大小
            dabiao = data_plan[data_plan.作业人员 == name][str2].values[0]  # 周达标
            week_tar = data_plan[data_plan.作业人员 == name][str1].values[0]  # 周目标
            week_act = data_plan[data_plan.作业人员 == name][str3].values[0]  # 周实际完成
            week_sum_n = data_plan[data_plan.作业人员 == name][str4].values[0]  # 周实际完成
            week_dat_n = data_plan[data_plan.作业人员 == name][str5].values[0]  # 周实际完成
            week_bl = data_plan[data_plan.作业人员 == name][str6].values[0]  # 版本完成比例
            shengyu = data_plan[data_plan.作业人员 == name][str7].values[0]  # 版本完成比例
            bili = '%.1f%%' % (week_bl * 100)
            # print(dabiao)
            # 插入文本编写
            plan_str = '''%s
本周目标：%d
本周完成：%d
下周目标：%d
下周日目标：%d
版本剩余：%d
版本完成：%s''' % (dabiao, week_tar, week_act, week_sum_n, week_dat_n, shengyu, bili)
            # print(plan_str)
            ax1.text(days - 0.2, 0, plan_str)  # 天数-0.8 为text的横坐标
            # print(week,type(week))
            # plt.show()
            plt.savefig('导出结果/' + '第' + week + '周统计/' + name + '-' + str_pr + '第' + week + '周' + '.png')  # 保存照片
            # ax3 = fig.add_subplot(1,2,2)
            # x = data_day_df[data_day_df.作业人员==name].日期.values.tolist()
            # y = data_day_df[data_day_df.作业人员==name].产出量.values.tolist()
            # print(x, y)
            # plt.plot(x, y)
        print('周个人' + str_pr + '作业情况保存成功！')
        # plt.show()

    #文件保存路径获取
    def save_path(self):
        start_date = Sta_ex.startdate
        end_date = Sta_ex.enddate

        poi_savefile_name = '设施' + start_date + '-' + end_date + '.xls'
        road_savefile_name = '道路' + start_date + '-' + end_date + '.xls'
        poi_path_save = '导出结果/' + '第' + Sta_ex.week + '周统计/' + poi_savefile_name
        road_path_save = '导出结果/' + '第' + Sta_ex.week + '周统计/' + road_savefile_name
        path_str = '''文件保存成功！
设施统计文件保存路径：%s
道路统计文件保存路径：%s
本周设施产出量: %d  本周道路产出量: %d''' % (poi_path_save, road_path_save, int(Sta_ex.poi_actual), int(Sta_ex.road_actual))
        self.com_path.setText(path_str)

    #个人道路作业情况统计
    def per_poi_Sta(self):
        day_data = Sta_ex.poi_data_ex
        time_data = Sta_ex.poi_time_ex
        #获取个人作业情况
        name = self.poi_name_sel.currentText()
        print('设施作业员：', name)
        start_date = self.poi_per_startday.text()
        end_date = self.poi_per_endday.text()
        start_day = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:8]))
        end_day = datetime.date(int(end_date[0:4]), int(end_date[4:6]), int(end_date[6:8]))
        print('统计时间段：', start_day, '~', end_day)
        daytime = []
        if start_day > end_day:
            print('开始日期必须小于结束日期!!!!!!!!')
        else:
            date_c = start_day
            date_str = ''
            while date_c <= end_day:
                date_str = date_c.strftime('%Y%m%d')
                daytime.append(date_str)
                date_c += datetime.timedelta(1)
        days = len(daytime)
        # 定义日统计矩阵
        poi_day = []
        for i in range(days):
            poi_day.append([])
            for j in range(12):
                poi_day[i].append(0)
        for j in range(days):
            poi_day[j][0] = name
            poi_day[j][1] = daytime[j]

        k = 0
        for j in range(len(day_data)):
            for i in range(len(poi_day)):
                if day_data[j][5] == name and day_data[j][6] == poi_day[i][1]:
                    for k in range(6):
                        poi_day[i][k + 2] = day_data[j][k + 10] + poi_day[i][k + 2]
                poi_day[i][9] = poi_day[i][4] + poi_day[i][5] + poi_day[i][7]
                poi_day[i][10] = poi_day[i][5] + poi_day[i][6] + poi_day[i][7]
        for j in range(len(time_data)):
            for i in range(len(poi_day)):
                if time_data[j][2] == poi_day[i][0] and time_data[j][3] == poi_day[i][1]:
                    poi_day[i][8] = float(time_data[j][4]) + poi_day[i][8]
        day_list = ['作业人员', '日期', '实际完成规划量', '实际总作业量', '新增', '修改', '验证', '删除', '作业时长', '产出量', '工作量', '备注']
        week_list = ['作业人员', '实际完成规划量', '实际总作业量', '新增', '修改', '验证', '删除', '作业时长', '产出量', '工作量', '备注']
        poi_day.insert(0, day_list)
        # for line in poi_day:
        #     print(line)

        # 计算作业情况
        poi_day_df = pd.DataFrame(poi_day[1:], columns=poi_day[0])
        save_path = '导出结果/' + Sta_ex.day_savefile_name + '/' + name + '/'
        # print('测试。。。')
        file_name = name + start_date + '-' + end_date + '.xls'
        poi_day_df.to_excel(save_path + file_name, sheet_name=file_name, index=False)
        # time_ave = poi_day_df['作业时长'].mean()
        time_sum = poi_day_df['作业时长'].sum()
        work_sum = poi_day_df['产出量'].sum()
        #个人作业情况绘图
        poi_day_df = poi_day_df.set_index('日期')
        if days<8:
            fig_width = days+5
        else:
            fig_width = days
        fig_height = 6
        fig = plt.figure(figsize=(fig_width, 6))
        ax1 = fig.add_subplot(1, 1, 1)
        poi_day_df['产出量'].plot(ax=ax1, kind='bar', subplots=True, style='-s')
        # hline绘制直线
        # plt.axhline(y=target_day, color='r')
        # ax1.axhline(y=target_day, linestyle='-.', color='r', label='日目标')
        handles1, labels1 = ax1.get_legend_handles_labels()  # 获取图一图例相关参数
        ax1.set_ylabel('产出量')
        # l1 = ax1.legend(loc=2)
        ax1.set_xlabel('作业日期')
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=-45)
        ax2 = ax1.twinx()
        poi_day_df['作业时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^', color='y')
        handles2, labels2 = ax2.get_legend_handles_labels()  # 获取图二图例相关参数
        ax2.set_ylabel('作业时长(小时)')
        handles2.extend(handles1)
        labels2.extend(labels1)  # 图例参数合并
        plt.legend(handles2[::-1], labels2[::-1])  # 绘制图例
        ax1.set_xlim(-0.5, days - 0.5)
        ax1.set_title('poi'+name+start_date+'-'+end_date)
        ax2.set_ylim(-1, 10)
        # 插入文本编写
        #         plan_str = '''总时长：%d
        # 总产出量：%d''' % (time_sum, work_sum)
        #         # print(plan_str)
        #         ax1.text(days - 0.3, 0, plan_str)  # 天数-0.8 为text的横坐标
        plt.subplots_adjust(bottom=0.15, top=.95, left=.08, right=0.92)  # 调整图像空白区域大小
        plt.savefig(save_path + 'poi'+ name + start_date+'-' + end_date + '.png')  # 保存照片
        plt.show()

    # 个人道路作业情况统计
    def per_road_Sta(self):
        day_data = Sta_ex.road_data_ex
        # 获取个人作业情况
        name = self.road_name_sel.currentText()
        print('设施作业员：', name)
        start_date = self.road_per_startday.text()
        end_date = self.road_per_endday.text()
        start_day = datetime.date(int(start_date[0:4]), int(start_date[4:6]), int(start_date[6:8]))
        end_day = datetime.date(int(end_date[0:4]), int(end_date[4:6]), int(end_date[6:8]))
        print('统计时间段：', start_day, '~', end_day)
        daytime = []
        if start_day > end_day:
            print('开始日期必须小于结束日期!!!!!!!!')
        else:
            date_c = start_day
            date_str = ''
            while date_c <= end_day:
                date_str = date_c.strftime('%Y-%m-%d')
                daytime.append(date_str)
                date_c += datetime.timedelta(1)
        days = len(daytime)
        # 定义日统计矩阵
        # 索引序号
        date_index = day_data[0].index('作业日期')
        name_index = day_data[0].index('作业员姓名')
        id_index = day_data[0].index('作业员ID')
        station_index = day_data[0].index('基地')
        time_index = day_data[0].index('有效时长')
        uproad_index = day_data[0].index('更新里程')
        newroad_index = day_data[0].index('新增里程')
        allroad_index = day_data[0].index('总作业里程')
        picnum_index = day_data[0].index('DCS图标量')
        road_day = []

        for i in range(days):
            road_day.append([])
            for j in range(8):
                road_day[i].append(0)
        for j in range(days):
            road_day[j][0] = name
            road_day[j][1] = daytime[j]
        day_list = [day_data[0][name_index], day_data[0][date_index], day_data[0][time_index],
                    day_data[0][uproad_index], day_data[0][newroad_index], day_data[0][allroad_index],
                    day_data[0][picnum_index], '产出量']
        k = 0
        # 统计日记录
        road_day.insert(0, day_list)
        for i in range(1, len(road_day)):
            for j in range(1, len(day_data)):
                if name == day_data[j][name_index] and road_day[i][1] == day_data[j][date_index]:
                    road_day[i][2] = float(day_data[j][time_index]) + road_day[i][2]
                    road_day[i][3] = float(day_data[j][uproad_index]) + road_day[i][3]
                    road_day[i][4] = float(day_data[j][newroad_index]) + road_day[i][4]
                    road_day[i][5] = float(day_data[j][allroad_index]) + road_day[i][5]
                    road_day[i][6] = float(day_data[j][picnum_index]) + road_day[i][6]
                    road_day[i][7] = road_day[i][6]
        print('测试。。。')
        # for line in poi_day:
        #     print(line)
        # 计算作业情况
        road_day_df = pd.DataFrame(road_day[1:], columns=road_day[0])
        print(road_day_df)
        save_path = '导出结果/' + Sta_ex.day_savefile_name + '/' + name + '/'
        # print('测试。。。')
        file_name = name + start_date + '-' + end_date + '.xls'
        road_day_df.to_excel(save_path + file_name, sheet_name=file_name, index=False)
        # time_ave = poi_day_df['作业时长'].mean()
        # 个人作业情况绘图
        road_day_df = road_day_df.set_index('作业日期')
        if days < 8:
            fig_width = days + 5
        else:
            fig_width = days
        fig_height = 6
        fig = plt.figure(figsize=(fig_width, 7))
        ax1 = fig.add_subplot(1, 1, 1)
        road_day_df['产出量'].plot(ax=ax1, kind='bar', subplots=True, style='-s')
        # hline绘制直线
        # plt.axhline(y=target_day, color='r')
        # ax1.axhline(y=target_day, linestyle='-.', color='r', label='日目标')
        handles1, labels1 = ax1.get_legend_handles_labels()  # 获取图一图例相关参数
        ax1.set_ylabel('产出量')
        # l1 = ax1.legend(loc=2)
        ax1.set_xlabel('作业日期')
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=-45)
        ax2 = ax1.twinx()
        road_day_df['有效时长'].plot(ax=ax2, linewidth=0.5, subplots=True, style='-^', color='y')
        handles2, labels2 = ax2.get_legend_handles_labels()  # 获取图二图例相关参数
        ax2.set_ylabel('作业时长(小时)')
        handles2.extend(handles1)
        labels2.extend(labels1)  # 图例参数合并
        plt.legend(handles2[::-1], labels2[::-1])  # 绘制图例
        ax1.set_xlim(-0.5, days - 0.5)
        ax1.set_title('road' + name + start_date + '-' + end_date)
        ax2.set_ylim(-1, 10)
        plt.subplots_adjust(bottom=0.15, top=.95, left=.08, right=0.92)  # 调整图像空白区域大小
        plt.savefig(save_path + 'road' + name + start_date + '-' + end_date + '.png')  # 保存照片
        plt.show()

    #文件存在与否弹出对话框
    def msg_path_exist(self, str):
        reply = QMessageBox.information(self, "文件路径检查", str+" 不存在，请确认")

    #日期输入与否弹出对话框
    def msg(self):
        reply = QMessageBox.information(self, '日期输入', '输入日期与周数')

    # 判断周文件夹是否存在，若不存在新建
    def plan_path_exists(self):
        week_str = self.weeklist_sel.currentText()
        if len(week_str)==3:
            Sta_ex.week = week_str[1]
        elif len(week_str)==4:
            Sta_ex.week = week_str[1:3]
        print(len(week_str),Sta_ex.week)
        filename = week_str + '统计'
        if os.path.exists('导出结果/' + filename) == False:
            os.mkdir('导出结果/' + filename)

    # 判断日与阶段统计文件夹是否存在，若不存在新建
    def path_exists(self):
        if os.path.exists('导出结果/' + Sta_ex.day_savefile_name) == False:
            os.mkdir('导出结果/' + Sta_ex.day_savefile_name)

    #日与阶段统计保存成功提示框
    def day_save_msg(self):
        reply = QMessageBox.information(self, '统计完成', '已保存至'+'"'+'.../'+'导出结果/' + Sta_ex.day_savefile_name +'/"')

    #判断上周统计数据是否存在
    def msg_plan_exit(self):
        pre_plan_name_poi = '导出结果/' + '第'+ str(int(Sta_ex.week) - 1)+'周统计/'+ '第' + str(int(Sta_ex.week) - 1) + '周设施达标情况统计.xls'
        pre_plan_name_road = '导出结果/' +'第'+ str(int(Sta_ex.week) - 1)+'周统计/'+ '第' + str(int(Sta_ex.week) - 1) + '周道路达标情况统计.xls'
        if int(Sta_ex.week) > 1:
            if os.path.exists(pre_plan_name_poi) == False or os.path.exists(pre_plan_name_road) == False:
                reply = QMessageBox.information(self, "文件路径检查", "上周达标数据未找到，请先进行上周数据统计")

    #保存路径
    def msg_plan_path_save(self):
        station = Sta_ex.station
        version = Sta_ex.version
        date = self.plan_endday.text()
        date_str = date[0:4]+'年'+date[4:6]+'月'+date[6:8]+'日'
        num1 = Sta_ex.poi_target - Sta_ex.poi_finished
        num2 = Sta_ex.road_target-Sta_ex.road_finished
        # print('测试。。。')
        # weekplan_str = '''%s%s版本，截止%s''' % (station, version, date_str)
        weekplan_str = '''%s%s版本，截止%s
设施已完成产出量：%d, 剩余%d
道路已完成产出量：%d，剩余%d'''%(station, version, date_str, Sta_ex.poi_finished,
                                        num1, Sta_ex.road_finished, num2)
        self.station_sta_info.setText(weekplan_str)

def main():
    app = QApplication(sys.argv)
    ex = Sta_ex()
    # ex.initUI()
    ex.show()
    sys.exit(app.exec_())
main()