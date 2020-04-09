import time
import tkinter
import threading
import datetime
import schedule
from selenium import webdriver
from Senses_ctrl import *
from SensesAnalyzeReport import *
from SalesResultReport import *
from WebCreate import *
from ScheduleMake import *
from WeeklyReport import *
from MarketReport import *
from MailingService import *
from OrderbookAnalyze import *
from ServiceMaintenance import *
from subprocess import Popen
from tkinter import messagebox
import subprocess


class App:
    def __init__(self, root, root_title):
        self.root = root
        self.root.geometry("900x400")
        self.root.title(root_title)
        self.down_time = 60
        self.server_ip = '127.0.0.1'
        self.server_port = '5000'
        self.buff = tkinter.StringVar()
        self.buff.set('')
        self.buff10 = tkinter.StringVar()
        self.buff10.set('')
        self.buff20 = tkinter.StringVar()
        self.buff20.set('')
        self.buff21 = tkinter.StringVar()
        self.buff21.set('')

        self.analyze_report_dirs = ('\SensesAnalyzeReport\ADSTEC',
									'\SensesAnalyzeReport\Aprolink')
        self.senses_data_dirs = ('\SensesData\ADSTEC',
								 '\SensesData\Aprolink')
        self.schedule_report_dirs = ('\ScheduleData\ADSTEC',
							  '\ScheduleData\Aprolink')
        self.week_report_dirs = ('\WeeklyReport\ADSTEC',
								 '\WeeklyReport\Aprolink')
        self.market_report_dirs = ('\MarketReport\ADSTEC',
								   '\MarketReport\Aprolink')
        self.orderbook_report_dirs = ('\OrderbookAnalyzeReport\ADSTEC',
									  '\OrderbookAnalyzeReport\AproLink')
        self.sourcefile_dirs = ('\\ADST受注台帳\\18期\マクロ18期受注台帳.xlsm',
								'\\AproLink\\14期_Aprolink_受注台帳.xlsm')
        self.sales_report_dirs = ('\SalesResultReport\ADSTEC',)
        
        self.maintenance_dir = os.getcwd() + '\Maintenance'

        self.senses_login_accounts = ('yuki@ads-tec.co.jp', 'tomooka@aprolink.jp')

        self.member_dict = {'ADSTEC':['茶畑', '太田', '楯', '南間', '黒瀬', '前嶋', '石井', '白井', '加村'],
							'Aprolink':['矢向', '友岡', '斎藤', '塚田', '松永']}
        self.member = []
        self.send_dict = {'ADSTEC':[['yuki@ads-tec.co.jp']
									,['chabata@ads-tec.co.jp']
						            ,['ota@ads-tec.co.jp']
									,['tate@ads-tec.co.jp']
									,['nanma@ads-tec.co.jp']
									,['kurose@ads-tec.co.jp']
									,['maeshima@ads-tec.co.jp']
									,['ishii@ads-tec.co.jp']
									,['rshirai@ads-tec.co.jp']]
						  ,'Aprolink':[['yuki@ads-tec.co.jp']
					                ,['yako@aprolink.jp']
									,['tomooka@aprolink.jp']
									,['saito@aprolink.jp']
									,['tsukada@aprolink.jp']
									,['matsunaga@aprolink.jp']]}

        self.sen_get = Senses_ctrl()
        self.senses_analyze_repo = SensesAnalyzeReport()
        self.schedule_make = ScheduleMake()
        self.week_repo = WeeklyReport()
        self.market_repo = MarketReport()
        self.mail_service = MailingService(self.member_dict)
        self.orderbook_repo = OrderbookAnalyze()
        self.maintenance = ServiceMaintenance(self.maintenance_dir)
        self.sales_result_repo = SalesResultReport()

        self.stop_auto_update = False
        self.stop_sub_process = False

        self.label_frame1=tkinter.LabelFrame(root,
											text='Senses関連手動更新コントロール',
											width=315,
										   height=390,
										   relief='groove',
										  borderwidth=4)

        self.button1=tkinter.Button(root, text="ADSTEC/Aprolink\nSensesデータ取得", command=self.req_csv_thread)
        self.button1.place(x=15, y=30, width=140, height=60)

        self.label1=tkinter.Label(text='Sensesデータ\nダウンロード待ち時間')
        self.label1.place(x=160, y=30)
        self.editbox1=tkinter.Entry(width=15)
        self.editbox1.place(x=165, y=65)
        self.editbox1.insert(tkinter.END, self.down_time)  

        self.button2=tkinter.Button(root, text="解析データ更新", command=self.analyze_data_thread)
        self.button2.place(x=15, y=100, width=140, height=30)

        self.button3=tkinter.Button(root, text="スケジュールデータ更新", command=self.schedule_thread)
        self.button3.place(x=165, y=100, width=140, height=30)

        self.button4=tkinter.Button(root, text="週報作成", command=self.week_repo_thread)
        self.button4.place(x=15, y=140, width=140, height=30)

        self.button5=tkinter.Button(root, text="マーケット解析作成", command=self.market_repo_thread)
        self.button5.place(x=165, y=140, width=140, height=30)

        self.button6=tkinter.Button(root, text="週報メール送信", command=self.mail_thread)
        self.button6.place(x=15, y=180, width=290, height=50)

        self.button7=tkinter.Button(root, text="営業分析作成", command=self.sales_repo_thread)
        self.button7.place(x=15, y=240, width=290, height=50)

        self.label2=tkinter.Label(text='[手動更新データステータス]')
        self.label2.place(x=15, y=300)
        self.counter = tkinter.Label(root, 
							   textvariable=self.buff, 
							   font=('Helvetica', '15'),
                               borderwidth=1,
                               relief='groove')
        self.counter.place(x=15, y=330, width=290, height=55)

        self.label_frame1.pack()
        self.label_frame1.place(x=5, y=5)

        self.label_frame10=tkinter.LabelFrame(root,
											text='受注台帳、その他手動更新コントロール',
											width=300,
										   height=390,
										   relief='groove',
										  borderwidth=4)

        self.button10=tkinter.Button(root, text="受注台帳取得", command=self.orderbook_get_thread)
        self.button10.place(x=330, y=30, width=140, height=60)

        self.button11=tkinter.Button(root, text="ADSTEC台帳解析データ更新", command=self.orderbook_analyze_thread_ADSTEC)
        self.button11.place(x=330, y=100, width=200, height=30)

        self.button12=tkinter.Button(root, text="Aprolink台帳解析データ更新", command=self.orderbook_analyze_thread_Aprolink)
        self.button12.place(x=330, y=140, width=200, height=30)

        self.button13=tkinter.Button(root, text="Excelデータメンテナンス用", command=self.excel_check_thread)
        self.button13.place(x=330, y=180, width=200, height=30)

        self.label10=tkinter.Label(text='[受注台帳、その他手動更新データステータス]')
        self.label10.place(x=330, y=300)
        self.counter = tkinter.Label(root, 
							   textvariable=self.buff10, 
							   font=('Helvetica', '15'),
                               borderwidth=1,
                               relief='groove')
        self.counter.place(x=330, y=330, width=270, height=55)

        self.label_frame10.pack()
        self.label_frame10.place(x=320, y=5)

        self.label_frame20=tkinter.LabelFrame(root,
											text='自動更新コントロール',
											width=275,
										   height=390,
										   relief='groove',
										  borderwidth=4)

        self.button20=tkinter.Button(root, text="自動更新開始", command=self.auto_update_thread)
        self.button20.place(x=630, y=30, width=120, height=50)

        self.button21=tkinter.Button(root, text="自動更新停止", command=self.auto_update_stop)
        self.button21.place(x=760, y=30, width=120, height=50)

        self.button22=tkinter.Button(root, text="Webサーバー起動", command=self.server_thread)
        self.button22.place(x=630, y=90, width=120, height=50)

        self.button23=tkinter.Button(root, text="Webサーバー停止", command=self.stop_server)
        self.button23.place(x=760, y=90, width=120, height=50)

        self.label20=tkinter.Label(text='WebサーバIP:')
        self.label20.place(x=630, y=150)
        self.editbox20=tkinter.Entry(width=23)
        self.editbox20.place(x=730, y=150)
        self.editbox20.insert(tkinter.END, self.server_ip)  

        self.label21=tkinter.Label(text='Webサーバポート:')
        self.label21.place(x=630, y=170)
        self.editbox21=tkinter.Entry(width=23)
        self.editbox21.place(x=730, y=170)
        self.editbox21.insert(tkinter.END, self.server_port)  

        self.label22=tkinter.Label(text='[自動更新ステータス]')
        self.label22.place(x=630, y=210)
        self.counter = tkinter.Label(root, 
							   textvariable=self.buff20, 
							   font=('Helvetica', '15'),
                               borderwidth=1,
                               relief='groove')
        self.counter.place(x=630, y=240, width=240, height=45)

        self.label23=tkinter.Label(text='[Webサーバステータス]')
        self.label23.place(x=630, y=300)
        self.counter = tkinter.Label(root, 
							   textvariable=self.buff21, 
							   font=('Helvetica', '15'),
                               borderwidth=1,
                               relief='groove')
        self.counter.place(x=630, y=330, width=240, height=55)

        self.label_frame20.pack()
        self.label_frame20.place(x=620, y=5)

        self.root.mainloop()

    def req_csv_thread(self):
        thread = threading.Thread(target = self.req_csv, args=())
        thread.start()

    def req_csv(self):
        for i in range(len(self.senses_data_dirs)):
            self.buff.set('CSVファイル要求中')
            self.sen_get.senses_login(self.senses_data_dirs[i],self.senses_login_accounts[i])
            self.sen_get.deallist_download()
            self.sen_get.actionlist_download()
            self.sen_get.customerlist_download()
            self.down_time = self.editbox1.get()
            self.buff.set('ダウンロード待機中')
            self.sen_get.data_download(self.down_time)
        self.buff.set('ダウンロード完了')

    def analyze_data_thread(self):
        thread = threading.Thread(target = self.analyze_data_update, args=())
        thread.start()

    def analyze_data_update(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        self.buff.set('データ読込中')
        for i in range(len(self.analyze_report_dirs)):
            #try:
            self.senses_analyze_repo.open_data(self.senses_data_dirs[i],
										self.analyze_report_dirs[i])
            self.buff.set('データ初期化中')
            self.senses_analyze_repo.data_delete()
            self.buff.set('データ書き込み中')
            self.senses_analyze_repo.data_update()
            self.buff.set('解析データ出力中')
            self.senses_analyze_repo.data2pdf(save_time)
            self.buff.set('解析データ更新完了')
#            except:
#                self.buff.set('解析更新エラー')

    def schedule_thread(self):
        thread = threading.Thread(target = self.sche_data_make, args=())
        thread.start()

    def sche_data_make(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        self.buff.set('データ読込中')
        for i in range(len(self.schedule_report_dirs)):
            #try:
            self.schedule_make.open_data(self.senses_data_dirs[i],
										 self.schedule_report_dirs[i])   
            self.buff.set('データ初期化中')
            self.schedule_make.data_delete()
            self.buff.set('データ書き込み中')
            self.schedule_make.data_update()
            self.buff.set('解析データ出力中')
            self.schedule_make.data2pdf(save_time)
            self.buff.set('スケジュール更新完了')
       # except:
        #    self.buff.set('スケジュール更新エラー')

    def week_repo_thread(self):
        thread = threading.Thread(target = self.week_repo_make, args=())
        thread.start()

    def week_repo_make(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        for i in range(len(self.week_report_dirs)):
            self.buff.set('データ読込中')
            self.week_repo.open_data(self.senses_data_dirs[i],
									 self.week_report_dirs[i])
            member_list = self.member_dict[os.path.basename(self.senses_data_dirs[i])]
            for n in range(len(member_list)):
                person = member_list[n]
                try:
                    self.buff.set('データ初期化中')
                    self.week_repo.data_delete()
                    self.buff.set(person+'データ書き込み中')
                    self.week_repo.data_update(person)
                    self.buff.set(person+'解析データ出力中')
                    self.week_repo.data2pdf(person, save_time)
                except:
                    self.buff.set('週報更新エラー')
                    self.week_repo.data_close()
            self.week_repo.data_close()
            self.buff.set('週報作成終了')

    def market_repo_thread(self):
        thread = threading.Thread(target = self.market_repo_make, args=())
        thread.start()

    def market_repo_make(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        for i in range(1,len(self.market_report_dirs)):
            try:
                self.buff.set('データ読込中')
                self.market_repo.open_data(self.senses_data_dirs[i],
	    								 self.market_report_dirs[i])
                self.buff.set('データ初期化中')
                self.market_repo.data_delete()
                self.buff.set('データ書き込み中')
                self.market_repo.data_update()
                self.buff.set('解析データ出力中')
                self.market_repo.data2pdf(save_time)
            except:
                self.buff.set('顧客常用CSV読込エラー')
            self.buff.set('業界解析作成終了')

    def orderbook_get_thread(self):
        thread = threading.Thread(target = self.orderbook_data_get, args=())
        thread.start()

    def orderbook_data_get(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        self.buff10.set('取得先PC接続確認')
        #try:
        self.orderbook_repo.connect_pc()
        self.buff10.set('受注台帳取得中')
        for i in range(len(self.orderbook_report_dirs)):
            self.orderbook_repo.get_file(save_time,
										 self.orderbook_report_dirs[i],
									     self.sourcefile_dirs[i])
        self.buff10.set('受注台帳取得完了')
        #except:
            #self.buff10.set('台帳データ更新エラー')

    def orderbook_analyze_thread_ADSTEC(self):
        thread = threading.Thread(target = self.orderbook_analyze_update_ADSTEC, args=())
        thread.start()

    def orderbook_analyze_update_ADSTEC(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        #try:
        self.buff10.set('ADSTEC台帳データ\n読み込み中')
        self.orderbook_repo.open_data(self.orderbook_report_dirs[0])
        self.buff10.set('ADSTEC解析データ\n初期化中')
        self.orderbook_repo.data_delete()
        self.buff10.set('ADSTEC解析データ\n更新中')
        self.orderbook_repo.data_update()
        self.buff10.set('ADSTEC解析データ\n出力中')
        self.orderbook_repo.data2pdf(save_time)
        self.buff10.set('ADSTEC受注台帳\n解析完了')
        #except:
         #   self.buff10.set('ADSTEC受注台帳\n解析エラー')

    def orderbook_analyze_thread_Aprolink(self):
        thread = threading.Thread(target = self.orderbook_analyze_update_Aprolink, args=())
        thread.start()

    def orderbook_analyze_update_Aprolink(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        try:
            self.buff10.set('AproLink台帳データ\nパスワード解除中')
            self.orderbook_repo.unlock_data(self.orderbook_report_dirs[1])
            self.buff10.set('AproLink台帳データ\n読込中')
            self.orderbook_repo.open_data(self.orderbook_report_dirs[1])
            self.buff10.set('AproLink解析データ\n初期化中')
            self.orderbook_repo.data_delete()
            self.buff10.set('AproLink解析データ\n更新中')
            self.orderbook_repo.data_update()
            self.buff10.set('AproLink解析データ\n出力中')
            self.orderbook_repo.data2pdf(save_time)
            self.buff10.set('AproLink受注台帳\n解析完了')
        except:
            self.buff10.set('AproLink受注台帳解析エラー')

    def sales_repo_thread(self):
        thread = threading.Thread(target = self.sales_repo_update, args=())
        thread.start()

    def sales_repo_update(self):
        dt = datetime.datetime.now()
        save_time = '{0:%Y%m%d-%H%M%S}'.format(dt)
        self.buff.set('データ読込中')
        for i in range(len(self.sales_report_dirs)):
            print(len(self.sales_report_dirs))
            #try:
            print(self.sales_report_dirs[i])
            self.sales_result_repo.open_data(self.senses_data_dirs[i],
										self.sales_report_dirs[i])
            self.buff.set('データ初期化中')
            self.sales_result_repo.data_delete()
            self.buff.set('データ書き込み中')
            self.sales_result_repo.data_update()
            self.buff.set('解析データ出力中')
            # self.sales_result_repo.data2pdf(save_time)
            self.buff.set('解析データ更新完了')
#            except:
#                self.buff.set('解析更新エラー')


    def server_thread(self):
        thread = threading.Thread(target = self.server_activate, args=())
        thread.start()

    def stop_server(self):
        self.stop_sub_process = True
        
    def server_activate(self):
        self.stop_sub_process = False
        self.server_ip = self.editbox20.get()
        self.server_port = self.editbox21.get()
        cmd = [r"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python",
			  "WebCreate.py",
			 self.server_ip,
			self.server_port]
        sub_p = subprocess.Popen(cmd, shell=True)
        status = 'Webサーバー起動中...'
        num = len(status)-3
        while not self.stop_sub_process:
            self.buff21.set(status[:num])
            if num == len(status):
                num = len(status)-4
            num += 1
            time.sleep(1)
        else:
            subprocess.Popen("TASKKILL /F /PID {pid} /T".format(pid=sub_p.pid))
            sub_p = None
            self.buff21.set('サーバーを停止しました。')        

    def auto_update_thread(self):
        thread = threading.Thread(target = self.auto_update, args=())
        thread.start()
 
    def auto_update_stop(self):
        self.stop_auto_update = False

    def auto_update(self):
        self.stop_auto_update = True
        schedule.every().monday.at("8:00").do(self.mail_thread)
        schedule.every().day.at("07:00").do(self.req_csv_thread)
        schedule.every().day.at("07:05").do(self.analyze_data_thread)
        schedule.every().day.at("07:10").do(self.schedule_thread)
        schedule.every().day.at("07:15").do(self.market_repo_thread)
        schedule.every().day.at("09:00").do(self.orderbook_analyze_update_ADSTEC)
        schedule.every().day.at("09:05").do(self.orderbook_analyze_update_ADSTEC)

        schedule.every().day.at("12:00").do(self.req_csv_thread)
        schedule.every().day.at("12:05").do(self.analyze_data_thread)
        schedule.every().day.at("12:10").do(self.schedule_thread)
        schedule.every().day.at("12:15").do(self.orderbook_analyze_update_ADSTEC)
        schedule.every().day.at("12:20").do(self.orderbook_analyze_update_ADSTEC)
        schedule.every().friday.at("21:00").do(self.week_repo_thread)
         
        while self.stop_auto_update:
            schedule.run_pending()
            time.sleep(1)
            dt = datetime.datetime.now()
            heart_beat = '{0:%Y%m%d-%H%M%S}'.format(dt)
            self.buff20.set('更新中'+heart_beat)
        else:
            self.buff20.set('自動更新停止しました。')

    def mail_thread(self):
        thread = threading.Thread(target = self.mail_event, args=())
        thread.start()

    def mail_event(self):
        self.buff.set('メール送信開始')
        for i in range(len(self.week_report_dirs)):
            for n in range(len(self.send_dict[os.path.basename(self.week_report_dirs[i])])):
                time.sleep(1)
                send_to = self.send_dict[os.path.basename(self.week_report_dirs[i])][n]
                self.buff.set(send_to[0] +'に送信中')
                self.mail_service.transmitte_mail(send_to, self.week_report_dirs[i], self.schedule_report_dirs[i])
                self.buff.set(send_to[0] +'に送信完了')
            self.buff.set('メール送信完了')

    def excel_check_thread(self):
        thread = threading.Thread(target = self.excel_check_service, args=())
        thread.start()

    def excel_check_service(self):
        for i in range(2):
            data_list = (self.analyze_report_dirs[i],
				self.schedule_dirs[i],
			self.week_report_dirs[i],
			self.market_report_dirs[i],
			self.orderbook_report_dirs[i]
			)
        self.maintenance.collect_analyze_excel(data_list)
        self.buff10.set('データ収集完了')


App(tkinter.Tk(), 'SensesAnalyzerコントロールパネル v1.4.0β')




