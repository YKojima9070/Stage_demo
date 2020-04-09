import PySimpleGUI as sg
import threading
import serial
import time
import sys

class App():

    def __init__(self):
        self.button_active = False ##初期値はFalse
        self.button_enable_color = ['gray99', 'DodgerBlue4']
        self.button_disable_color = ['gray26', 'DodgerBlue4']
        self.tgl_jog_p_checked = False
        self.tgl_jog_m_checked = False
        self.continue_move = False
        self.stop_flag = False
        self.status = ''
        self.posi_status = 0
        self.speed_data = 100 #初期値単位m/s
        self.posi1_data = 100 #初期値単位mm
        self.posi2_data = 700 #初期値単位mm

        column1 = [[sg.Button('初期化', size=(15,2), font=('Helvetica', 25),button_color=self.button_enable_color)],
                   [sg.Button('繰返開始', size=(15,2), font=('Helvetica', 25), key='繰返',button_color=self.button_disable_color)],
                   [sg.Text('最大値目安700(mm)\nJogの動作範囲制限無、注意！',font=('Helvetica', 12))]
                   ]
               
        column2 = [[sg.Button('Jog+\n動作',size=(7,2), font=('Helvetica', 20), key='Jog+',button_color=self.button_disable_color),
                   sg.Button('Jog-\n動作',size=(7,2), font=('Helvetica', 20), key='Jog-',button_color=self.button_disable_color)],
                   [sg.Text('動作位置(mm)   ',font=('Helvetica', 18)),
                    sg.In(self.posi_status, size=(5,3), font=('Helvetica', 20), key='posi_data',justification='right')],
                   [sg.Text('動作速度(mm/s)',font=('Helvetica', 18)),
                    sg.In(self.speed_data, size=(5,3), font=('Helvetica', 20), key='speed_data',justification='right')],
                   [sg.Text('Posi1(mm)        ',font=('Helvetica', 18)),
                    sg.In(self.posi1_data, size=(5,3), font=('Helvetica', 20), key='posi1_data',justification='right')],
                   [sg.Text('Posi2(mm)        ',font=('Helvetica', 18)),
                    sg.In(self.posi2_data,size=(5,3), font=('Helvetica', 20), key='posi2_data',justification='right')]                    
                   ]
        
        layout = [[sg.Column(column1),
                   sg.Column(column2)],
                  [sg.Text('ステータス',font=('Helvetica', 20))],
                    [sg.In(size=(60,5), font=('Helvetica', 15), key='cur_status')]
                  ]

        window = sg.Window(title = 'StageDemo', size=(680,350)).Layout(layout)

        try:
            ##Windows用
            self.gm_s_motor_port = serial.Serial('COM3', 9600, timeout=0.5) #port名,timeoutは受信待機時間
            
            ##Linux用
            #self.gm_s_motor_port = serial.Serial('/dev/ttyUSB0', 9600, timeout=0.5) #port名,timeoutは受信待機時間
            
        except:
            sg.Popup('エラー','COMポートエラー')

               
        while True:
            event, values = window.Read(timeout=100) #送信インターバル(画面更新時間)
            
            if event == sg.TIMEOUT_KEY:
                window.FindElement('posi_data').Update(self.posi_status)
                window.FindElement('cur_status').Update(self.status)

            if event in (None, 'Exit'):
                break
            

            if event == '初期化':
                try:
                    threading.Thread(target=self.servo_init_thread).start()
                    self.button_active = True
                    window.FindElement('繰返').Update(button_color=self.button_enable_color)
                    window.FindElement('Jog+').Update(button_color=self.button_enable_color)
                    window.FindElement('Jog-').Update(button_color=self.button_enable_color)

                except:
                    sg.Popupo('初期化エラー','接続エラーが発生しました。')
            
                else:
                    None

            elif event == 'Jog+':
                if self.button_active:
                    threading.Thread(target=self.jog_p_stop).start()
                    time.sleep(0.5)
                       
                    if not self.tgl_jog_p_checked:
                        threading.Thread(target=self.jog_p_move).start()
                        self.tgl_jog_p_checked = True
                        self.status = 'Jog+動作中'
                        window.FindElement('Jog+').Update('Jog+\n停止')

                    else:
                        self.tgl_jog_p_checked = False
                        window.FindElement('Jog+').Update('Jog+\n動作')
                        self.status = 'Jog+停止'
                        
                        time.sleep(0.2)
                        self.posi_status = self.cur_posi_check()                        
                        if not self.posi_status == None :
                            window.FindElement('posi_data').Update(self.posi_status)


                else:
                    None

            elif event == 'Jog-':
                if self.button_active:
                    threading.Thread(target=self.jog_m_stop).start()
                    time.sleep(0.5)

                    if not self.tgl_jog_m_checked:
                        threading.Thread(target=self.jog_m_move).start()
                        self.tgl_jog_m_checked = True
                        self.status = 'Jog-動作中'
                        window.FindElement('Jog-').Update('Jog\n-停止')

                    else:
                        self.tgl_jog_m_checked = False
                        window.FindElement('Jog-').Update('Jog\n動作')
                        self.status = 'Jog-停止'

                        time.sleep(0.2)
                        self.posi_status = self.cur_posi_check()                        
                        if not self.posi_status == None :
                            window.FindElement('posi_data').Update(self.posi_status)


                else:
                    None

            elif event == '繰返':
                if self.button_active:

                    try:
                        self.posi1_data = int(values['posi1_data'])
                        self.posi2_data = int(values['posi2_data'])
                        self.speed_data = int(values['speed_data'])
                        
                        ##オーバーリミット制限
                        if self.posi1_data > 700:
                            self.posi1_data = 700
                        
                        if self.posi2_data > 700:
                            self.posi2_data = 700

                    except:
                        self.status = '入力値エラー'
 
                    if not self.continue_move:
                        window.FindElement('繰返').Update('動作\n停止')

                        threading.Thread(target=self.manual_move_thread,
                                         args=[self.posi1_data,
                                               self.posi2_data,
                                               self.speed_data]).start()
                        
                        self.continue_move = True
                    
                    else:
                        time.sleep(0.5)
                        self.stop_flag = True
                        window.FindElement('繰返').Update('繰返動作')
                        time.sleep(0.5)
                        self.stop_servo()
                        self.continue_move = False     
    
        self.gm_s_motor_port.close()
        window.Close()

##スレッド管理##

    ##停止スレッド
    def stop_move_thread(self):
        time.sleep(0.5)
        self.tgl_stop_thread()

    ##往復動作スレッド
    def manual_move_thread(self, posi1, posi2, speed):
        self.stop_flag = False
        i = 1
        while not self.stop_flag:
            time.sleep(0.1)
            self.check_posi = False

            if not i % 2 == 0:
                self.manual_move_query(posi1, speed)
                if self.manual_move_response(posi1):
                    i +=1
                    continue

                else:
                    self.status = '位置決めタイムアウトエラー'
                    break

            else:
                self.manual_move_query(posi2, speed)
                if self.manual_move_response(posi2):
                    i +=1
                    continue
                else:
                    self.status = '位置決めタイムアウトエラー'
                    break

    ##サーボコントロール(初期化)
    def servo_init_thread(self):
        self.pio_control()
        time.sleep(0.5)
        self.servo_on()
        time.sleep(0.5)
        self.check_regi()
        time.sleep(0.5)                     
        self.move_origin()

##以下Modbus通信コマンド##    

    ##Modbus有効/PIO無効
    def pio_control(self):
        send_buf = bytearray(b"")  #空のbytearrayを作成

        send_buf.append(ord(':'))  #各バイト値にASCIIからUNICODE変換

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('2'))
        send_buf.append(ord('7'))

        send_buf.append(ord('F'))
        send_buf.append(ord('F'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('D'))
        send_buf.append(ord('0'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)  #シリアル通信送信

    ##Servo_ON
    def servo_on(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('0'))
        send_buf.append(ord('3'))

        send_buf.append(ord('F'))
        send_buf.append(ord('F'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('F'))
        send_buf.append(ord('4'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)

    ##レジスタ確認
    def check_regi(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('3'))

        send_buf.append(ord('9')) #開始アドレス
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('5'))
        
        send_buf.append(ord('0')) #レジスタの数9000H～9004Hまでを書き込む
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        lrc_add_buf = self.add_LRC(send_buf, 1, 13)
        send_buf.append(ord(lrc_add_buf[0]))
        send_buf.append(ord(lrc_add_buf[1]))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)

        time.sleep(1)

        read_buf = self.gm_s_motor_port.readline().decode('utf-8')

        if not read_buf == '':
            self.status = 'レスポンス受信しました'

        else:
            self.status = '接続先が有りません'


    ##原点復帰
    def move_origin(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':')) #0:ヘッダ

        send_buf.append(ord('0')) #1:スレープアドレス
        send_buf.append(ord('1')) #2

        send_buf.append(ord('0')) #3:ファンクションコード
        send_buf.append(ord('5')) #4

        send_buf.append(ord('0')) #5:開始アドレス
        send_buf.append(ord('4')) #6
        send_buf.append(ord('0')) #7
        send_buf.append(ord('B')) #8

        send_buf.append(ord('0')) #9:レジスタの数
        send_buf.append(ord('0')) #10
        send_buf.append(ord('0')) #11
        send_buf.append(ord('0')) #12

        lrc_add_buf = self.add_LRC(send_buf, 1, 13)
        send_buf.append(ord(lrc_add_buf[0]))
        send_buf.append(ord(lrc_add_buf[1]))

        send_buf.append(0x0D) #15:トレーラー
        send_buf.append(0x0A) #16

        self.gm_s_motor_port.write(send_buf) #0からなのでトータル17bit
 
        time.sleep(0.5)

        send_buf[9] = ord('F') #9ビットめを変更
        send_buf[10] = ord('F') #10ビットめを変更

        send_buf[13] = ord('E')  #13ビットめを変更
        send_buf[14] = ord('C')  #14ビットめを変更

        self.gm_s_motor_port.write(send_buf)


    ##減速停止
    def stop_servo(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('2'))
        send_buf.append(ord('C'))
        
        send_buf.append(ord('F'))
        send_buf.append(ord('F'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        lrc_add_buf = self.add_LRC(send_buf, 1, 13)
        send_buf.append(ord(lrc_add_buf[0]))
        send_buf.append(ord(lrc_add_buf[1]))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)


    ##セーフティ速度指令 有効
    def safety_speed_enable(self):

        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('0'))
        send_buf.append(ord('1'))
        
        send_buf.append(ord('F'))
        send_buf.append(ord('F'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        lrc_add_buf = self.add_LRC(send_buf, 1, 13)
        send_buf.append(ord(lrc_add_buf[0]))
        send_buf.append(ord(lrc_add_buf[1]))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)


    ##セーフティ速度指令 無効
    def safety_speed_disable(self):

        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('0'))
        send_buf.append(ord('1'))
        
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        lrc_add_buf = self.add_LRC(send_buf, 1, 13)
        send_buf.append(ord(lrc_add_buf[0]))
        send_buf.append(ord(lrc_add_buf[1]))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)


    ##JOG＋動作
    def jog_p_move(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('1'))
        send_buf.append(ord('6'))
        
        send_buf.append(ord('F'))
        send_buf.append(ord('F'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('E'))
        send_buf.append(ord('1'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)


    ##JOG＋停止
    def jog_p_stop(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('1'))
        send_buf.append(ord('6'))
        
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('E'))
        send_buf.append(ord('0'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)


    ##JOG-動作
    def jog_m_move(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('1'))
        send_buf.append(ord('7'))
        
        send_buf.append(ord('F'))
        send_buf.append(ord('F'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('E'))
        send_buf.append(ord('0'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)
 
        
    ##JOG-停止
    def jog_m_stop(self):
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        send_buf.append(ord('0'))
        send_buf.append(ord('4'))
        send_buf.append(ord('1'))
        send_buf.append(ord('7'))
        
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('D'))
        send_buf.append(ord('F'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)


    ##現在位置取得
    def cur_posi_check(self):
        
        ##位置情報クエリ##
        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('0'))
        send_buf.append(ord('3'))

        send_buf.append(ord('9'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('2'))

        send_buf.append(ord('6'))
        send_buf.append(ord('A'))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)

        ##レスポンス##
        
        try:
            read_buf = self.gm_s_motor_port.readline().decode()[7:15]

            if not read_buf == "":
                cur_posi = int(read_buf,16) * 0.01 ##単位はmm
                return cur_posi
            else:
                return None

        except:
            return None

        
    ##直値移動指令
    def manual_move_query(self, target_posi, set_speed):
        target_posi = format(target_posi*100, 'x').upper().zfill(8)
        set_speed = format(set_speed*100, 'x').upper().zfill(8)

        send_buf = bytearray(b"")

        send_buf.append(ord(':'))

        send_buf.append(ord('0'))
        send_buf.append(ord('1'))

        send_buf.append(ord('1'))
        send_buf.append(ord('0'))

        send_buf.append(ord('9'))
        send_buf.append(ord('9'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('6'))

        send_buf.append(ord('0'))
        send_buf.append(ord('C'))

        ##目標位置指定レジスタ
        send_buf.append(ord(target_posi[0]))
        send_buf.append(ord(target_posi[1]))
        send_buf.append(ord(target_posi[2]))
        send_buf.append(ord(target_posi[3]))

        send_buf.append(ord(target_posi[4]))
        send_buf.append(ord(target_posi[5]))
        send_buf.append(ord(target_posi[6]))
        send_buf.append(ord(target_posi[7]))

        ##位置決め幅指定レジスタ
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))

        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('5'))

        ##速度指定レジスタ
        send_buf.append(ord(set_speed[0]))
        send_buf.append(ord(set_speed[1]))
        send_buf.append(ord(set_speed[2]))
        send_buf.append(ord(set_speed[3]))

        send_buf.append(ord(set_speed[4]))
        send_buf.append(ord(set_speed[5]))
        send_buf.append(ord(set_speed[6]))
        send_buf.append(ord(set_speed[7]))

        ##加減速度
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('0'))
        send_buf.append(ord('7'))

        lrc_add_buf = self.add_LRC(send_buf, 1, 43)
        send_buf.append(ord(lrc_add_buf[0]))
        send_buf.append(ord(lrc_add_buf[1]))

        send_buf.append(0x0D)
        send_buf.append(0x0A)

        self.gm_s_motor_port.write(send_buf)

    ##位置決め完了確認
    def manual_move_response(self, tar_posi):
        #for n in range(2000): #タイムスリープ0.1x100 = 10secターゲット位置にない場合往復終了
            #time.sleep(0.1)
        while True:

            cur_posi = int(self.cur_posi_check())
            self.posi_status = cur_posi
            if tar_posi -10 < cur_posi < tar_posi +10: #ポジション誤差設定±2
                return True
                break
            
            #elif n == 100:
             #   return False

            #else:
             #   None

    ##LRC計算用
    def add_LRC(self, buff, i_start, i_LRC):

        lrc_totall = 0
        for i in range(i_start, i_LRC, 2):
            i_LRC = ""
            i_LRC = chr(buff[i]) + chr(buff[i+1])
            i_LRC = (int(i_LRC, 16))
            lrc_totall += i_LRC
        
        lrc_comple = str.upper(format(int(bin(-(lrc_totall) & 0xFF), 2), 'x'))
        
        return lrc_comple.zfill(2)

App()