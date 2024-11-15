from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtWidgets import QTableWidgetItem,QLineEdit, QApplication, QMainWindow, QMessageBox, QFileDialog
from PyQt6.QtGui import QBrush, QColor, QFont, QIcon
from gui_main import Ui_MainWindow
import pandas as pd
import serial
import time
from datetime import datetime
import random
import string
import re
 
tin_hieu_reset = True
id_them_xoa = ''
chuc_vu_them = ''
chuc_nang_them_xoa = ''
tin_hieu_bat_tat_led_buzzer = False
tin_hieu_xin_nghi = False


class Cap_nhat_so_luong_nhan_vien(QThread):
    progress = pyqtSignal(str)

    def __init__(self, index):
        super().__init__()
        self.index = index
        self._running = True

    def run(self):
        while self._running:
            data_danh_sach_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
            tong_so_nhan_vien = len(data_danh_sach_nhan_vien)
            self.progress.emit(str(tong_so_nhan_vien))
            time.sleep(1)  # Thời gian chờ để kiểm tra lại số lượng nhân viên

    def stop(self):
        self._running = False
        self.wait()

class Cap_nhat_thoi_gian(QThread):
    progress_hour = pyqtSignal(str)

    def __init__(self, index):
        super().__init__()
        self.index = index
        self._running = True

    def run(self):
        while self._running:
            now = datetime.now()
            formatted_datetime = now.strftime("%d/%m/%Y %H:%M:%S")
            self.progress_hour.emit(str(formatted_datetime))
            
            with open ('gio_vao_lam.txt', 'r') as f:
                gio_vao_lam = f.read()
                
            self.progress_hour.emit(f'Giờ vào làm: {gio_vao_lam}')
            
            with open ('gio_tan_lam.txt', 'r') as f:
                gio_tan_lam = f.read()
                
            self.progress_hour.emit(f'Giờ tan làm: {gio_tan_lam}')
            
            with open ('so_ngay_nghi.txt', 'r') as f:
                so_ngay_nghi_max = f.read()
            self.progress_hour.emit(f'Số ngày nghỉ tối đa/tháng: {so_ngay_nghi_max}')
            
            time.sleep(1)

    def stop(self):
        self._running = False
        self.wait()



class Cap_nhat_tu_dong_cham_cong(QThread):
    def __init__(self, index):
        super().__init__()
        self.index = index
        self._isrunning = True
        self.updated = False  # Biến cờ để kiểm soát quá trình cập nhật
        self.update_xinnghi = False
        self.date_cu = ''
        self.date_now = ''

    def run(self):
        while self._isrunning:
            try:
                time.sleep(1)
                with open('gio_tan_lam.txt', 'r') as f:
                    gio_tan_lam = f.read()

                hour_now = datetime.now().time()
                gio_tan_lam = datetime.strptime(gio_tan_lam, '%H:%M').time()
                date_now = datetime.now().date()
                self.date_now = str(date_now)
                
                

                # Xét các trường hợp k đi làm và tự động quẹt thẻ khi đã kết thúc giờ làm
                if hour_now > gio_tan_lam:
                    if  self.updated == False:
                        data_nv = pd.read_csv('Danh_sach_nhan_vien.csv')
                        data_cham_cong = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')
                        # Lọc dữ liệu chấm công theo ngày
                        
                        hour_now_2 = datetime.now().strftime('%H:%M')
                        data_cham_cong['Ngày'] = pd.to_datetime(data_cham_cong['Ngày'], format='%Y-%m-%d').dt.date
                        data_cham_cong_ngay = data_cham_cong[data_cham_cong['Ngày'] == date_now]

                        
                        # Lấy ra các tên có trong danh sách chấm công
                        
                        id_nv = list(data_nv['ID'])
                        for id in id_nv:
                            if id not in list(data_cham_cong_ngay['ID']):
                                ten_nv_nghi = str(list(data_nv['Họ và tên'][data_nv['ID'] == id])[0])
                                new_row_nv_nghi = pd.DataFrame({
                                    'ID': [id],
                                    'Họ và tên': [ten_nv_nghi],
                                    'Thời gian': [str(hour_now_2)],
                                    'Ngày': [datetime.now().strftime('%Y-%m-%d')],
                                    'Trạng thái': ['Nghỉ làm'],
                                    'Ghi chú (Phút)': [0]
                                })
                                data_cham_cong = pd.concat([new_row_nv_nghi, data_cham_cong], ignore_index=True)
                                data_cham_cong['Ngày'] = pd.to_datetime(data_cham_cong['Ngày'], format='%Y-%m-%d')

                        # Tự động thêm tan làm khi quên chưa quẹt thẻ
                        data_id_vao_ngay = data_cham_cong_ngay['ID'][
                            (data_cham_cong_ngay['Trạng thái'] == 'Vào làm muộn') |
                            (data_cham_cong_ngay['Trạng thái'] == 'Vào làm đúng giờ')
                        ]
                        # Nếu tan làm sớm thì không tự động tích
                        for id_vao_ngay in list(data_id_vao_ngay):
                            if any((data_cham_cong_ngay['ID'] == id_vao_ngay) & (data_cham_cong_ngay['Trạng thái'] == 'Tan làm sớm')):
                                pass
                            elif any((data_cham_cong_ngay['ID'] == id_vao_ngay) & (data_cham_cong_ngay['Trạng thái'] != 'Tan làm sớm')):
                                ten_nv_nghi = str(list(data_nv['Họ và tên'][data_nv['ID'] == id_vao_ngay])[0])
                                new_row_nv_nghi = pd.DataFrame({
                                    'ID': [id_vao_ngay],
                                    'Họ và tên': [ten_nv_nghi],
                                    'Thời gian': [str(hour_now_2)],
                                    'Ngày': [datetime.now().strftime('%Y-%m-%d')],
                                    'Trạng thái': ['Tan làm đúng giờ'],
                                    'Ghi chú (Phút)': [0]
                                })
                                
                                with open ('xac_nhan_tu_dong_tan_lam.txt', 'r') as f:
                                    th_td_tl = f.read()
                                f.close()
                                if th_td_tl == 'Chua':
                                    data_cham_cong = pd.concat([new_row_nv_nghi, data_cham_cong], ignore_index=True)
                                    data_cham_cong['Ngày'] = pd.to_datetime(data_cham_cong['Ngày'], format='%Y-%m-%d')
                                    with open ('xac_nhan_tu_dong_tan_lam.txt', 'w') as f:
                                        f.write('Roi')
                                    f.close()

                        print(data_cham_cong)
                        data_cham_cong.to_csv('Danh_sach_cham_cong_nhan_vien.csv', index=False)
                        self.updated = True  # Đánh dấu đã cập nhật
                        with open ('xac_nhan_xin_nghi.txt', 'w') as f:
                            f.write('Chua')
                        f.close()
                        
                        
                elif hour_now < gio_tan_lam:
                    # Khi sáng một ngày mới sẽ tự động chấm công những người đã xin nghỉ
                    if self.date_now != self.date_cu:
                        with open ('xac_nhan_xin_nghi.txt', 'r') as f:
                            tin_hieu_xn_nghi = f.read()
                        f.close()
                        with open ('xac_nhan_tu_dong_tan_lam.txt', 'w') as f:
                                f.write('Roi')
                        f.close()
                        
                        self.updated = False
                        self.update_xinnghi = False  # Đặt lại biến update_xinnghi để đảm bảo chỉ cập nhật một lần
                        if self.update_xinnghi == False:
                            if tin_hieu_xn_nghi == 'Chua':
                                # Tìm xem ai đã nghỉ phép và thêm vào danh sách chấm công
                                data_xin_nghi = pd.read_csv('Danh_sach_xin_nghi.csv')
                                # Tìm kiếm danh sách trong ngày hôm đó

                                data_xin_nghi_hom_nay = data_xin_nghi[data_xin_nghi['Ngày nghỉ'] == str(date_now)]
                                print('date now: ', date_now)
                                print('Danh sách xin nghỉ hôm nay: ', data_xin_nghi_hom_nay)
                                data_cham_cong = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')

                                # ID,Họ và tên,Thời gian,Ngày,Trạng thái,Ghi chú (Phút)
                                danh_sach_id_nghi = list(data_xin_nghi_hom_nay['ID'])
                                danh_sach_ten_nghi = list(data_xin_nghi_hom_nay['Họ và tên'])
                                danh_sach_ghi_chu_nghi = list(data_xin_nghi_hom_nay['Ghi chú'])

                                for index_nghi in range(len(danh_sach_id_nghi)):
                                    new_data_nghi = pd.DataFrame({
                                        'ID': [danh_sach_id_nghi[index_nghi]],
                                        'Họ và tên': [danh_sach_ten_nghi[index_nghi]],
                                        'Thời gian': [datetime.now().strftime('%H:%M')],
                                        'Ngày': [datetime.now().strftime('%Y-%m-%d')],
                                        'Trạng thái': ['Nghỉ làm'],
                                        'Ghi chú (Phút)': [danh_sach_ghi_chu_nghi[index_nghi]]
                                    })
                                    data_cham_cong = pd.concat([new_data_nghi, data_cham_cong], ignore_index=True)
                                    data_cham_cong['Ngày'] = pd.to_datetime(data_cham_cong['Ngày'], format='%Y-%m-%d')

                                data_cham_cong.to_csv('Danh_sach_cham_cong_nhan_vien.csv', index=False)
                                self.update_xinnghi = True  # Đánh dấu đã cập nhật để không lặp lại
                                
                                self.date_cu = self.date_now  # Cập nhật date_cu để so sánh trong các lần tiếp theo
                                with open ('xac_nhan_xin_nghi.txt', 'w') as f:
                                    f.write('Roi')
                                f.close()
                                
                                
        
            except:
                pass

class Ket_noi_STM32(QThread):
    progress = pyqtSignal(str)

    def __init__(self, index):
        super().__init__()
        self.index = index
        self._running = True
        self.current_port = None
        self.ser_id = None

    def run(self):
        global tin_hieu_reset, tin_hieu_bat_tat_led_buzzer
        while self._running:
            with open('Cong_COM.txt', 'r') as f:
                port_com = f.read().strip()

            if self.current_port != port_com:
                self.current_port = port_com
                if self.ser_id and self.ser_id.is_open:
                    self.ser_id.close()
                self.connect_to_port(port_com)

            if self.ser_id and self.ser_id.is_open:
                try:
                    self.data_recv = self.ser_id.readline().decode('utf-8')
                    if len(self.data_recv) > 0:
                        print('ID cham cong: ', self.data_recv)
                        self.progress.emit(self.data_recv)
                    
                    if tin_hieu_reset == True:
                        self.ser_id.write('RS'.encode())
                        print('Gửi Reset thành công!')
                        tin_hieu_reset = False
                    
                    if tin_hieu_bat_tat_led_buzzer == True:
                        self.ser_id.write('BC'.encode())
                        
                        tin_hieu_bat_tat_led_buzzer = False
                        
                    self.progress.emit(f'Kết nối {self.current_port}')
                except Exception as e:
                    self.progress.emit(f'Lỗi đọc dữ liệu từ cổng {self.current_port}')
                    self.ser_id.close()
                    self.ser_id = None

            else:
                # Thử kết nối lại nếu không có kết nối
                self.connect_to_port(self.current_port)

            time.sleep(1)

    def connect_to_port(self, port_com):
        try:
            self.ser_id = serial.Serial(port=port_com,
                                        baudrate=115200,
                                        bytesize=serial.EIGHTBITS,
                                        parity=serial.PARITY_NONE,
                                        stopbits=serial.STOPBITS_ONE,
                                        timeout=0.5)
            self.progress.emit(f'Đã kết nối tới cổng {port_com}')
        except Exception as e:
            self.progress.emit(f'Lỗi kết nối tới cổng {port_com}')
            self.ser_id = None

    def stop(self):
        self._running = False
        if self.ser_id and self.ser_id.is_open:
            self.ser_id.close()
        self.wait()
          


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.uic = Ui_MainWindow()
        self.uic.setupUi(self)
        self.thread = {}
        self.tin_hieu_page_thong_tin = False
        self.tin_hieu_page_login = False
        self.tin_hieu_page_doi_mk = False
        self.tin_hieu_page_quen_mk = False
        self.tin_hieu_page_cai_dat = False
        self.tin_hieu_dn_tc = False
        self.ho_ten_nghi = 'ID không tồn tại'
        self.sdt_nghi = 'ID không tồn tại'
        self.chuc_vu_nghi = 'ID không tồn tại'
        
        
        self.my_dict_loc = {
            'Họ và tên':'Không',
            'Tháng':'Không',
            'Năm':'Không',
            'Trạng thái':'Không'
            
        }
        
        self.dict_loc_thong_ke = {
            'ID':'Không',
            'Tháng':'Không',
            'Năm':'Không'
        }
        
        # Cập nhật danh sách nhân viên
        self.bat_dau_cap_nhat_danh_sach_nhan_vien()
        self.bat_dau_cap_nhat_thoi_gian()
        self.bat_dau_bat_STM()
        self.bat_dau_tu_dong_cham_cong()
        
        
        with open ('gio_vao_lam.txt', 'r') as f:
            gio_vao_lam = f.read()
            
        #### thực hiện các click   
        self.uic.lb_gio_vao_lam.setText(gio_vao_lam)
        self.uic.btn_reset.clicked.connect(self.reset)
        self.uic.btn_reset.setIcon(QIcon('reset.png'))
        self.uic.btn_reset.setIconSize(QSize(30,30))
        
        self.uic.btn_xn.clicked.connect(self.xac_nhan_cham_cong)
        
        self.uic.btn_thong_tin.clicked.connect(self.thong_tin_cham_cong)
        self.uic.btn_thong_tin.setIcon(QIcon('website.png'))
        self.uic.btn_thong_tin.setIconSize(QSize(25,25))
        
        self.uic.bt_admin.clicked.connect(lambda: self.trang_dang_nhap_admin())
        self.uic.bt_admin.setIcon(QIcon('ad.jfif'))
        self.uic.bt_admin.setIconSize(QSize(25,25))
      
       
    def trang_dang_nhap_admin(self): # PAGE login
        if self.tin_hieu_dn_tc == False:
            self.uic.stackedWidget.setCurrentWidget(self.uic.page_LOGIN)
            self.uic.edt_mk.setEchoMode(QLineEdit.EchoMode.Password)
            self.uic.cb_sp_dn.stateChanged.connect(self.toggle_password_login)
            
            if self.tin_hieu_page_login == False:
                self.uic.bt_dang_nhap.clicked.connect(lambda: self.dang_nhap_admin())
                self.uic.bt_trang_chu_admin.clicked.connect(self.mo_trang_chu)
                self.uic.bt_doi_mk.clicked.connect(lambda: self.doi_mat_khau_admin())
                self.uic.bt_quen_mk.clicked.connect(lambda: self.quen_mat_khau_admin())
                
                self.tin_hieu_page_login = True
        else:
            self.uic.stackedWidget.setCurrentWidget(self.uic.page_admin)
    
    def toggle_password_login(self):
        if self.uic.cb_sp_dn.isChecked():
            self.uic.edt_mk.setEchoMode(QLineEdit.EchoMode.Normal)
        else:
            self.uic.edt_mk.setEchoMode(QLineEdit.EchoMode.Password)
        
        
    # Trang đăng nhập admin    
    def dang_nhap_admin(self):
        global id_them_xoa, tin_hieu_reset
        
        input_tk = str(self.uic.edt_tk.text())
        input_mk = str(self.uic.edt_mk.text())
        print('input tk: ', input_tk)
        print('input mk: ', input_mk)
        data_login = pd.read_csv('Data_Login_Admin.csv')  
        print(data_login)
        tk_dung = list(data_login['TK'])[0]
        mk_dung = str(list(data_login['MK'])[0])
        
        print('TK: ', tk_dung)
        print('MK: ', mk_dung)
        
        if input_tk == tk_dung and input_mk == mk_dung:
            print('Đăng nhập thành công Admin')
            self.tin_hieu_dn_tc = True
            # Vào trang admin
            self.uic.stackedWidget.setCurrentWidget(self.uic.page_admin)
            self.uic.tab_quan_ly.setCurrentWidget(self.uic.tab_quan_ly_nhan_vien)
            if self.tin_hieu_page_cai_dat == False:
                # đặt tên cho các tab
                tin_hieu_reset = True
                self.uic.tab_quan_ly.setTabText(0, "Thêm/Xóa")
                self.uic.tab_quan_ly.setTabIcon(0, QIcon('add_remove.jfif'))
                self.uic.tab_quan_ly.setTabText(1, "Cài đặt hệ thống")
                self.uic.tab_quan_ly.setTabIcon(1, QIcon('setting.jfif'))
                self.uic.tab_quan_ly.setTabText(2, "Xin nghỉ")
                self.uic.tab_quan_ly.setTabIcon(2, QIcon('vang.png'))
                self.uic.tab_quan_ly.setTabText(3, 'Thống kê')
                self.uic.tab_quan_ly.setTabIcon(3, QIcon('tk.jfif'))
                
                self.uic.btn_trang_chu.clicked.connect(self.mo_trang_chu)
                self.uic.btn_trang_chu.setIcon(QIcon('home.jfif'))
                
                
                self.uic.tab_quan_ly.currentChanged.connect(self.tab_changed)
                self.uic.bt_xac_nhan_them.clicked.connect(self.xac_nhan_them_nhan_vien)
                self.uic.btn_xn_xoa.clicked.connect(self.xac_nhan_xoa_nhan_vien)
                self.uic.btn_xn_so_ngay_nghi.clicked.connect(self.xac_nhan_so_ngay_nghi)
                
                
                # Cài đặt định dạng giờ trong tab cai dat
                self.uic.timeEdit_gio_vao.setDisplayFormat('HH:mm')
                self.uic.timeEdit_gio_ra.setDisplayFormat('HH:mm')
                
                self.uic.combobox_cong_stm.addItem(f'Không')
                for i in range (0, 10):
                    self.uic.combobox_cong_stm.addItem(f'COM{i}')
                    
                # Cài đặt combobox
                self.data_chucvu = pd.read_csv('Data_chuc_vu.csv')
                self.uic.combo_chuc_vu.addItem('Không')
                self.danh_sach_chuc_vu = list(self.data_chucvu['Chức vụ'])
                for cv in self.danh_sach_chuc_vu:
                    self.uic.combo_chuc_vu.addItem(cv)
                
                
                self.uic.combo_chuc_vu.currentIndexChanged.connect(self.on_combo_chuc_vu_change)    
                
                # Cài đặt combobox cho them/xóa
                self.uic.combo_them_xoa.addItem('Không')
                self.uic.combo_them_xoa.addItem('Thêm nhân viên')
                self.uic.combo_them_xoa.addItem('Xóa nhân viên')
                
                self.uic.combo_them_xoa.currentIndexChanged.connect(self.on_combo_them_xoa)
                
                #Cài đặt combo cho so ngay nghi
                self.uic.combo_so_ngay_nghi.addItem('Không')
                for day in range (1, 31):
                    self.uic.combo_so_ngay_nghi.addItem(str(day))
                
                # Cài đặt combo lọc trong thống kê dữ liệu
                self.uic.combo_tk_id.addItem('Không')
                data_id_nv = list(pd.read_csv('Danh_sach_nhan_vien.csv')['ID'])
                for id_nv in data_id_nv:
                    self.uic.combo_tk_id.addItem(id_nv)

                self.uic.combo_thang_thongke.addItem('Không')
                for thang in range (1, 13):
                    self.uic.combo_thang_thongke.addItem(str(thang))
                
                self.uic.combo_nam_thongke.addItem('Không')    
                for nam in range (2024, 2029):
                    self.uic.combo_nam_thongke.addItem(str(nam))
                
                
                
                self.uic.btn_xn_gio_vao_lam.clicked.connect(self.xac_nhan_gio_vao_lam)
                self.uic.btn_xn_gio_tan_lam.clicked.connect(self.xac_nhan_gio_tan_lam)
                self.uic.btn_xn_cong_com.clicked.connect(self.xac_nhan_cong_com)
                
                
                self.uic.btn_xoa_dong_cd.clicked.connect(self.xoa_dong_lich_su_cap_nhat)
                self.uic.btn_dang_xuat.clicked.connect(self.dang_xuat_admin)
                self.uic.btn_dang_xuat.setIcon(QIcon('logouts.jfif'))
                
                self.uic.btn_xn_nghi.clicked.connect(self.xac_nhan_xin_nghi)
                self.uic.btn_xoa_nghi.clicked.connect(self.xoa_dong_xin_nghi)
                self.uic.calendarWidget.clicked.connect(self.cap_nhat_ngay_nghi)
                self.uic.btn_xn_id_nc_nghi.clicked.connect(self.xac_nhan_id_nghi_nc)
                
                self.uic.btn_loc_thongke.clicked.connect(self.loc_thong_ke)
            
                self.uic.btn_loc_thongke.setIcon(QIcon('search.jfif'))
                self.uic.btn_save_thongke.setIcon(QIcon('save.jfif'))
                self.uic.btn_save_thongke.clicked.connect(lambda: self.save_thong_ke())
                
                self.tin_hieu_page_cai_dat = True
            
            self.com_moi = None
            self.cap_nhat_bang_nhan_vien()
            
            self.cap_nhat_bang_xin_nghi()
            

            
        else:
            self.tin_hieu_dn_tc = False
            showMessageBox('Đăng nhập', 'Đăng nhập không thành công')
    
    
    def save_thong_ke(self):
        file_filter = 'Data File (*.xlsx);; Exel File (*.xlsx *.xls)'
        response, _ = QFileDialog.getSaveFileName(
            parent=self,
            caption='Select a data file',
            directory='Data file.xlsx',
            filter=file_filter,
            initialFilter='Excel File (*.xlsx *.xls)'
        )
        if response:
            
            if response.endswith('.xlsx') or response.endswith('.xls'):
                self.data_tk_save.to_excel(response, index=False)
            
            else:
                with open(response, 'w', encoding='utf-8') as file:
                    file.write(self.data_tk_save.to_string(index=False))

        
        
    def loc_thong_ke(self):
        self.dict_loc_thong_ke['ID'] = self.uic.combo_tk_id.currentText()
        self.dict_loc_thong_ke['Tháng'] = self.uic.combo_thang_thongke.currentText()
        self.dict_loc_thong_ke['Năm'] = self.uic.combo_nam_thongke.currentText()
        
        if self.dict_loc_thong_ke['Tháng'] != 'Không':
            self.dict_loc_thong_ke['Tháng'] = int(self.dict_loc_thong_ke['Tháng'])
        if self.dict_loc_thong_ke['Năm'] != 'Không':
            self.dict_loc_thong_ke['Năm'] = int(self.dict_loc_thong_ke['Năm'])
        
        data_thong_ke = thong_ke_trong_thang()
        self.data_filter_thongke = data_thong_ke
        print('Lọc: ', self.dict_loc_thong_ke)
        ds_co = [key for key in self.dict_loc_thong_ke if self.dict_loc_thong_ke[key] != 'Không']
        
        print('Danh sách có: ', ds_co)
        if len(ds_co) > 0:
            
            for key_co in ds_co:
                
                self.data_filter_thongke = self.data_filter_thongke[self.data_filter_thongke[key_co] == self.dict_loc_thong_ke[key_co]]
            print(self.data_filter_thongke)
        else:
            #ID,Họ và tên,Tháng,Năm,Số ngày đi làm,Số ngày vào muộn,Số ngày về sớm,Số ngày nghỉ
            self.data_filter_thongke = data_thong_ke
        self.data_filter_thongke.reset_index(drop=True, inplace=True)
        
        self.cap_nhat_bang_thong_ke(self.data_filter_thongke)
        self.data_tk_save = self.data_filter_thongke
        
                
                
        
    def xoa_dong_xin_nghi(self):
        selected_rows = self.uic.tb_ds_xin_nghi.selectionModel().selectedRows()
        self.data_xin_nghi = pd.read_csv('Danh_sach_xin_nghi.csv')
        if not selected_rows:
            showMessageBox("Thông báo", "Vui lòng chọn ít nhất một dòng để xóa.")
            return
        
        for selected_row in sorted(selected_rows, reverse=True):
            self.uic.tb_ds_xin_nghi.removeRow(selected_row.row())
            self.data_xin_nghi = self.data_xin_nghi.drop(selected_row.row()).reset_index(drop=True)
        
        self.data_xin_nghi.to_csv('Danh_sach_xin_nghi.csv', index=False) 

            
    def xac_nhan_id_nghi_nc(self):
        data_nv = pd.read_csv('Danh_sach_nhan_vien.csv')
        id_xn_nc = self.uic.edt_uid_nghi.text()
        data_tt_nv_id = data_nv[data_nv['ID'] == id_xn_nc]
        print(data_tt_nv_id)
        if len(data_tt_nv_id) == 0:
            self.uic.lb_uid_nghi.setText(id_xn_nc)
            self.uic.lb_sdt_nghi.setText('ID không tồn tại')
            self.uic.lb_hoten_nghi.setText('ID không tồn tại')
            self.uic.lb_chuc_vu_nghi.setText('ID không tồn tại')
            self.ho_ten_nghi = 'ID không tồn tại'
            self.sdt_nghi = 'ID không tồn tại'
            self.chuc_vu_nghi = 'ID không tồn tại'
        else:
            self.uic.lb_uid_nghi.setText(id_xn_nc)
            self.ho_ten_nghi = str(list(data_tt_nv_id['Họ và tên'])[0])
            self.sdt_nghi = str(list(data_tt_nv_id['SDT'])[0])
            self.chuc_vu_nghi = str(list(data_tt_nv_id['Chức vụ'])[0])
            
            self.uic.lb_sdt_nghi.setText(self.sdt_nghi)
            self.uic.lb_hoten_nghi.setText(self.ho_ten_nghi)
            self.uic.lb_chuc_vu_nghi.setText(self.chuc_vu_nghi)

    def xac_nhan_xin_nghi(self):
        self.ghi_chu_nghi = str(self.uic.edt_ghi_chu_nghi.toPlainText())
        self.uid_xin_nghi = self.uic.lb_uid_nghi.text()
        
        self.ngay_nghi = self.uic.lb_ngay_nghi.text()
    
        date_now = datetime.now().date()
        if (len(self.ngay_nghi) > 0 ):
            self.ngay_nghi = datetime.strptime(self.ngay_nghi, '%Y-%m-%d').date()
            
            if self.ho_ten_nghi == 'ID không tồn tại' or self.ngay_nghi <= date_now or len(self.uid_xin_nghi)<=0:
                showMessageBox('Thông báo', 'ID không tồn tại \rBạn phải xin nghỉ trước 1 ngày')
            else:
                
                
                data_xin_nghi = pd.read_csv('Danh_sach_xin_nghi.csv')
                data_xin_nghi['Ngày nghỉ'] = pd.to_datetime( data_xin_nghi['Ngày nghỉ'], format='%Y-%m-%d').dt.date
                data_id_xin_nghi_ngay = list(data_xin_nghi['ID'][data_xin_nghi['Ngày nghỉ'] == self.ngay_nghi])
                print(data_id_xin_nghi_ngay)

                if self.uid_xin_nghi in data_id_xin_nghi_ngay:
                    showMessageBox('Thông báo', 'Cập nhật lịch bạn đã xin nghỉ trước đó')
                    # thêm vào bảng dữ liệu
                    index_id_xin_nghi = data_id_xin_nghi_ngay.index(self.uid_xin_nghi)
                    if len(self.ghi_chu_nghi)==0:
                        self.ghi_chu_nghi = 'Không lý do'
                    # Xóa bỏ dòng cũ
                    data_xin_nghi = data_xin_nghi.drop(index_id_xin_nghi)
                    # cập nhật vào dòng mới
                    new_row_nghi = pd.DataFrame({'ID': [self.uid_xin_nghi], 
                                                'Họ và tên':[self.ho_ten_nghi], 
                                                'Chức vụ':[self.chuc_vu_nghi], 
                                                'SDT':[self.sdt_nghi], 
                                                'Ngày nghỉ':[self.ngay_nghi], 
                                                'Ghi chú':[self.ghi_chu_nghi]})
                    data_xin_nghi = pd.concat([new_row_nghi, data_xin_nghi], ignore_index = True)
                    
                    data_xin_nghi.to_csv('Danh_sach_xin_nghi.csv', index=False)
                    
                else:
                    # B1: Xét xem đã quá số ngày nghỉ cho phép trong tháng chưa
                    
                    try:
                        with open ('so_ngay_nghi.txt', 'r') as f:
                            so_ngay_nghi_toi_da = int(f.read())
                        
                        
                        # Xét xem ID đó xin nghỉ chưa
                        self.n_da_nghi = so_ngay_da_nghi(id_input=self.uid_xin_nghi)
                        if self.n_da_nghi < so_ngay_nghi_toi_da:
                            # Lấy danh sách xin nghỉ tháng này
                            
                            # B2: Chưa vượt qua số ngày nghỉ cho phép -> thực thi lệnh dưới
                            # thêm vào bảng dữ liệu
                            if len(self.ghi_chu_nghi)==0:
                                self.ghi_chu_nghi = 'Không lý do'
                            
                            new_row_nghi = pd.DataFrame({'ID': [self.uid_xin_nghi], 
                                                        'Họ và tên':[self.ho_ten_nghi], 
                                                        'Chức vụ':[self.chuc_vu_nghi], 
                                                        'SDT':[self.sdt_nghi], 
                                                        'Ngày nghỉ':[self.ngay_nghi], 
                                                        'Ghi chú':[self.ghi_chu_nghi]})
                            data_xin_nghi = pd.concat([new_row_nghi, data_xin_nghi], ignore_index = True)
                            data_xin_nghi.to_csv('Danh_sach_xin_nghi.csv', index=False)
                            showMessageBox('Thông báo', 'Xin nghỉ thành công')
                        else:
                            showMessageBox('Thông báo', f'Bạn đã nghỉ tối đa {so_ngay_nghi_toi_da} ngày') 
                    except:
                        showMessageBox('Thông báo', 'Hệ thống chưa cập nhật số ngày nghỉ tối đa') 
                    self.clear_tab_xin_nghi()
                
        else:
            showMessageBox('Thông báo', 'Xin nghỉ không thành công')
            
        self.cap_nhat_bang_xin_nghi()
        
    
    
    def cap_nhat_bang_thong_ke(self, data_thong_ke):
        self.uic.tb_thong_ke_thang.setRowCount(len(data_thong_ke))
        self.uic.tb_thong_ke_thang.setColumnCount(len(data_thong_ke.columns))
        self.uic.tb_thong_ke_thang.setHorizontalHeaderLabels(data_thong_ke.columns)
        
         # Tạo font in đậm
        font = QFont()
        font.setBold(True)

        # Áp dụng font in đậm cho các tiêu đề
        for i in range(len(data_thong_ke.columns)):
            item = self.uic.tb_thong_ke_thang.horizontalHeaderItem(i)
            if item:
                item.setFont(font)

        for row_index, row in data_thong_ke.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                if col_index in [4,5,6,7] and value > 0:
                    print('Col index: ', col_index)
                    item.setForeground(QBrush(QColor(200, 0, 0)))  # Màu xanh lam nhạt
                    font = QFont()
                    font.setBold(True)
                    item.setFont(font)

                self.uic.tb_thong_ke_thang.setItem(row_index, col_index, item)

        self.uic.tb_thong_ke_thang.resizeColumnsToContents()
        self.uic.tb_thong_ke_thang.resizeRowsToContents()

        for i in range(len(data_thong_ke.columns)):
            
            self.uic.tb_thong_ke_thang.setColumnWidth(i, 100)
            
            
    def cap_nhat_bang_xin_nghi(self):
        data_xin_nghi = pd.read_csv('Danh_sach_xin_nghi.csv')
        self.uic.tb_ds_xin_nghi.setRowCount(len(data_xin_nghi))
        self.uic.tb_ds_xin_nghi.setColumnCount(len(data_xin_nghi.columns))
        self.uic.tb_ds_xin_nghi.setHorizontalHeaderLabels(data_xin_nghi.columns)
        
        # Tạo font in đậm
        font = QFont()
        font.setBold(True)

        # Áp dụng font in đậm cho các tiêu đề
        for i in range(len(data_xin_nghi.columns)):
            item = self.uic.tb_ds_xin_nghi.horizontalHeaderItem(i)
            if item:
                item.setFont(font)
                
        for row_index, row in data_xin_nghi.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.uic.tb_ds_xin_nghi.setItem(row_index, col_index, item)
            
        self.uic.tb_ds_xin_nghi.resizeColumnsToContents()
        self.uic.tb_ds_xin_nghi.resizeRowsToContents()
        self.uic.tb_ds_xin_nghi.setColumnWidth(0, 120)
        self.uic.tb_ds_xin_nghi.setColumnWidth(1, 120)
        self.uic.tb_ds_xin_nghi.setColumnWidth(2, 120)
        self.uic.tb_ds_xin_nghi.setColumnWidth(3, 120)
        self.uic.tb_ds_xin_nghi.setColumnWidth(4, 120)
        self.uic.tb_ds_xin_nghi.setColumnWidth(5, 120)

        
        
    def dang_xuat_admin(self):
        self.tin_hieu_dn_tc = False
        global chuc_nang_them_xoa
        chuc_nang_them_xoa = ''
        self.uic.edt_tk.clear()
        self.uic.edt_mk.clear()
        self.uic.edt_tk_doi_mk.clear()
        self.uic.edt_mk_cu.clear()
        self.uic.edt_mk_moi.clear()
        self.uic.lb_mk_moi_quen.clear()
        self.uic.edt_tk_quen_mk.clear()
        self.uic.lb_uid_them.clear()
        self.uic.lb_uid_xoa.clear()
        self.uic.edt_ten_nhan_vien_them.clear()
        self.uic.edt_sdt_them.clear()
        self.uic.stackedWidget.setCurrentWidget(self.uic.page_trangchu)
    
    def cap_nhat_bang_nhan_vien(self):
    # cài đặt bảng danh sách nhân viên
            data_danh_sach_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
            
            self.uic.tb_danh_sach_nhan_vien.setRowCount(len(data_danh_sach_nhan_vien))
            self.uic.tb_danh_sach_nhan_vien.setColumnCount(len(data_danh_sach_nhan_vien.columns))
            self.uic.tb_danh_sach_nhan_vien.setHorizontalHeaderLabels(data_danh_sach_nhan_vien.columns)
             # Tạo font in đậm
            font = QFont()
            font.setBold(True)

            # Áp dụng font in đậm cho các tiêu đề
            for i in range(len(data_danh_sach_nhan_vien.columns)):
                item = self.uic.tb_danh_sach_nhan_vien.horizontalHeaderItem(i)
                if item:
                    item.setFont(font)
                    
            for row_index, row in data_danh_sach_nhan_vien.iterrows():
                for col_index, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.uic.tb_danh_sach_nhan_vien.setItem(row_index, col_index, item)
            
            self.uic.tb_danh_sach_nhan_vien.resizeColumnsToContents()
            self.uic.tb_danh_sach_nhan_vien.resizeRowsToContents()
            self.uic.tb_danh_sach_nhan_vien.setColumnWidth(0, 120)
            self.uic.tb_danh_sach_nhan_vien.setColumnWidth(1, 120)
            self.uic.tb_danh_sach_nhan_vien.setColumnWidth(2, 120)
            self.uic.tb_danh_sach_nhan_vien.setColumnWidth(3, 120)
            
    def tab_changed(self, index):
        global chuc_nang_them_xoa, tin_hieu_xin_nghi, tin_hieu_reset
        currentTab = self.uic.tab_quan_ly.tabText(index)
        if currentTab == 'Thêm/Xóa':
            # self.cap_nhat_bang_nhan_vien()
            tin_hieu_xin_nghi = False
            tin_hieu_reset = True
            self.clear_tab_xin_nghi()
            
        elif currentTab == 'Cài đặt hệ thống':
            self.cap_nhat_lich_su_bang()
            chuc_nang_them_xoa = ''
            tin_hieu_reset = True
            tin_hieu_xin_nghi = False
            self.clear_tab_them_xoa()
            self.clear_tab_xin_nghi()
            
            
        elif currentTab == 'Xin nghỉ':
            chuc_nang_them_xoa = ''
            # Tín hiệu xin nghỉ =True
            tin_hieu_xin_nghi = True
            tin_hieu_reset = True
            
            self.clear_tab_them_xoa()
        
        elif currentTab == "Thống kê":
            chuc_nang_them_xoa = ''
            # Tín hiệu xin nghỉ =True
            tin_hieu_xin_nghi = True
            tin_hieu_reset = True
            self.clear_tab_them_xoa()
            self.clear_tab_xin_nghi()
            data_thong_ke_thang = thong_ke_trong_thang()
            self.data_tk_save = data_thong_ke_thang
            self.cap_nhat_bang_thong_ke(data_thong_ke_thang)
            
            
            
    def clear_tab_xin_nghi(self):
        self.uic.lb_uid_nghi.clear()
        self.uic.edt_uid_nghi.clear()
        self.uic.lb_ngay_nghi.clear()
        self.uic.lb_hoten_nghi.clear()
        self.uic.lb_sdt_nghi.clear()
        self.uic.lb_chuc_vu_nghi.clear()
        self.uic.edt_ghi_chu_nghi.clear()
            
    def clear_tab_them_xoa(self):
        self.uic.lb_che_do_them_xoa.setText('Chế độ:')
        self.uic.lb_uid_them.clear()
        self.uic.edt_ten_nhan_vien_them.clear()
        self.uic.lb_chuc_vu_them.clear()
        self.uic.edt_sdt_them.clear()
        self.uic.lb_uid_xoa.clear()
        self.uic.edt_uid_xoa.clear()
               
            
    def cap_nhat_ngay_nghi(self):
        selected_date =  self.uic.calendarWidget.selectedDate()
        formatted_date = selected_date.toString('yyyy-MM-dd') 
        self.uic.lb_ngay_nghi.setText(formatted_date)     
            

        
    def xac_nhan_cong_com(self):
        self.com_moi = self.uic.combobox_cong_stm.currentText()
        if self.com_moi != 'Không':
            with open('Cong_COM.txt', 'w') as f:
                f.write(self.com_moi)
            f.close()
            showMessageBox('Cài đặt hệ thống', 'Cập nhật cổng thành công')
            ngay_cn = time.strftime('%H:%M - %d/%m/%Y')
            self.uic.lb_cn_com.setText(f'Cổng kết nối: {self.com_moi} - NCN: {ngay_cn}')
            # Thêm vào danh sách lịch sử cập nhật
            self.data_lich_su_cap_nhat = pd.read_csv('Lich_su_cap_nhat.csv')
            data_new = pd.DataFrame({
                'Loại':['Cổng kết nối'],
                'Kết quả':[self.com_moi],
                'Ngày':[ngay_cn]
                
            })   
            self.data_lich_su_cap_nhat = pd.concat([data_new, self.data_lich_su_cap_nhat], ignore_index=True)
            
            self.data_lich_su_cap_nhat.to_csv('Lich_su_cap_nhat.csv', index = False)
            self.cap_nhat_lich_su_bang()
            
        else:
            showMessageBox('Cài đặt hệ thống', 'Cập nhật cổng không thành công')
        
    def xoa_dong_lich_su_cap_nhat(self):
        selected_rows = self.uic.tb_lich_su_cn.selectionModel().selectedRows()
        self.data_lich_su_cap_nhat = pd.read_csv('Lich_su_cap_nhat.csv')
        if not selected_rows:
            showMessageBox("Thông báo", "Vui lòng chọn ít nhất một dòng để xóa.")
            return
        
        for selected_row in sorted(selected_rows, reverse=True):
            self.uic.tb_lich_su_cn.removeRow(selected_row.row())
            self.data_lich_su_cap_nhat = self.data_lich_su_cap_nhat.drop(selected_row.row()).reset_index(drop=True)
        
        self.data_lich_su_cap_nhat.to_csv('Lich_su_cap_nhat.csv', index=False)    
            
    def xac_nhan_gio_vao_lam(self):
        gio_vao_moi = self.uic.timeEdit_gio_vao.text()  
        with open ('gio_tan_lam.txt', 'r') as f:
            gio_tan_lam = f.read()
        f.close()
        
        t_gio_vao = time.strptime(gio_vao_moi, '%H:%M')
        t_gio_tan = time.strptime(gio_tan_lam, '%H:%M')
        if t_gio_vao >= t_gio_tan:
            showMessageBox('Cài đặt hệ thống', 'Giờ vào phải nhỏ hơn giờ tan')
        else:
            with open ('gio_vao_lam.txt', 'w') as f:
                f.write(gio_vao_moi)
            f.close()
            showMessageBox('Cài đặt hệ thống', 'Cập nhật thành công')
            ngay_cn = time.strftime('%H:%M - %d/%m/%Y')
            self.uic.lb_cn_gio_vao.setText(f'Giờ vào: {gio_vao_moi} - NCN: {ngay_cn}')
            # Thêm vào danh sách lịch sử cập nhật
            data_lich_su_cap_nhat = pd.read_csv('Lich_su_cap_nhat.csv')
            data_new = pd.DataFrame({
                'Loại':['Giờ vào làm'],
                'Kết quả':[gio_vao_moi],
                'Ngày':[ngay_cn]
                
            })   
            data_lich_su_cap_nhat = pd.concat([data_new, data_lich_su_cap_nhat], ignore_index=True)
            
            data_lich_su_cap_nhat.to_csv('Lich_su_cap_nhat.csv', index = False)
            self.cap_nhat_lich_su_bang()
            
    def xac_nhan_gio_tan_lam(self):
        gio_tan_moi = self.uic.timeEdit_gio_ra.text()
        with open ('gio_vao_lam.txt', 'r') as f:
            gio_vao_lam = f.read()
            
        f.close()
        
        t_gio_vao = time.strptime(gio_vao_lam, '%H:%M')
        t_gio_tan = time.strptime(gio_tan_moi, '%H:%M')

        if t_gio_vao >= t_gio_tan:
            showMessageBox('Cài đặt hệ thống', 'Giờ vào phải nhỏ hơn giờ tan')
        else:
            with open('gio_tan_lam.txt', 'w') as f:
                f.write(gio_tan_moi)
            f.close()
            showMessageBox('Cài đặt hệ thống', 'Cập nhật thành công')
            ngay_cn = time.strftime('%H:%M - %d/%m/%Y')
            self.uic.lb_cn_gio_ra.setText(f'Giờ ra: {gio_tan_moi} - NCN: {ngay_cn}')
            # Thêm vào danh sách lịch sử cập nhật
            data_lich_su_cap_nhat = pd.read_csv('Lich_su_cap_nhat.csv')
            data_new = pd.DataFrame({
                'Loại':['Giờ tan làm'],
                'Kết quả':[gio_tan_moi],
                'Ngày':[ngay_cn]
                
            })   
            data_lich_su_cap_nhat = pd.concat([data_new, data_lich_su_cap_nhat], ignore_index=True)
            
            data_lich_su_cap_nhat.to_csv('Lich_su_cap_nhat.csv', index = False)
            self.cap_nhat_lich_su_bang()
            

    
    def cap_nhat_lich_su_bang(self):
        data_lich_su_cn = pd.read_csv('Lich_su_cap_nhat.csv')
            
        self.uic.tb_lich_su_cn.setRowCount(len(data_lich_su_cn))
        self.uic.tb_lich_su_cn.setColumnCount(len(data_lich_su_cn.columns))
        self.uic.tb_lich_su_cn.setHorizontalHeaderLabels(data_lich_su_cn.columns)
        
        for row_index, row in data_lich_su_cn.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.uic.tb_lich_su_cn.setItem(row_index, col_index, item)
        
        self.uic.tb_lich_su_cn.resizeColumnsToContents()
        self.uic.tb_lich_su_cn.resizeRowsToContents()
        self.uic.tb_lich_su_cn.setColumnWidth(0, 120)
        self.uic.tb_lich_su_cn.setColumnWidth(1, 120)
        self.uic.tb_lich_su_cn.setColumnWidth(2, 120)
        
    def on_combo_chuc_vu_change(self):
        global chuc_vu_them
        
        chuc_vu_them = self.uic.combo_chuc_vu.currentText()
        self.uic.lb_chuc_vu_them.setText(chuc_vu_them)        
            
    def on_combo_them_xoa(self):
        global chuc_nang_them_xoa
        chuc_nang_them_xoa = self.uic.combo_them_xoa.currentText()
        tin_hieu_reset = True  
            
        self.uic.lb_che_do_them_xoa.setText(chuc_nang_them_xoa)
    

    def xac_nhan_so_ngay_nghi(self):
        so_ngay_nghi = self.uic.combo_so_ngay_nghi.currentText()
        if so_ngay_nghi == 'Không':
            showMessageBox('Thông báo', 'Xác nhận số ngày nghỉ không thành công')
        else:
            ngay_cn = time.strftime('%H:%M - %d/%m/%Y')
            with open ('so_ngay_nghi.txt', 'w') as f:
                f.write(so_ngay_nghi)
                
            f.close()
            
            self.uic.lb_xn_so_ngay_nghi.setText(f'Số ngày nghỉ/tháng: {so_ngay_nghi} - NCN: {ngay_cn}')
            self.data_lich_su_cap_nhat = pd.read_csv('Lich_su_cap_nhat.csv')
            
            data_new = pd.DataFrame({
                'Loại':['Số ngày nghỉ/tháng'],
                'Kết quả':[so_ngay_nghi],
                'Ngày':[ngay_cn]
                
            })   
            self.data_lich_su_cap_nhat = pd.concat([data_new, self.data_lich_su_cap_nhat], ignore_index=True)
            
            self.data_lich_su_cap_nhat.to_csv('Lich_su_cap_nhat.csv', index = False)
            self.cap_nhat_lich_su_bang()
            showMessageBox('Thông báo', 'Xác nhận số ngày nghỉ thành công')  
            
          
           
        
    def xac_nhan_them_nhan_vien(self):
        global chuc_vu_them, chuc_nang_them_xoa
        id_them = self.uic.lb_uid_them.text()
        input_ht_them = self.uic.edt_ten_nhan_vien_them.text()
        input_chuc_vu_them = chuc_vu_them
        input_sdt_them = self.uic.edt_sdt_them.text()
        print('Chức năng: ', chuc_nang_them_xoa)
        
        if chuc_nang_them_xoa == 'Thêm nhân viên':
            data_danh_sach_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
            danh_sach_id_da_co = list(data_danh_sach_nhan_vien['ID'])
            if id_them not in danh_sach_id_da_co:
                # B1: Kiểm tra họ và tên
                if len(input_ht_them) > 0:
                    if xu_ly_ten(input_ht_them) != 'Thừa khoảng trắng' and xu_ly_ten(input_ht_them) != 'Chỉ chứa chữ cái':
                        # B2 kiểm tra chức vụ
                        if len(input_chuc_vu_them) > 0:
                            # B3 Kiểm tra số đt
                            state_ktsdt = kiem_tra_so_dien_thoai_vn(input_sdt_them)
                            if state_ktsdt == True:
                                # Thêm nhân viên mới
                                data_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
                                new_nhan_vien = pd.DataFrame({
                                    'ID':[str(id_them)], 
                                    'Họ và tên':[input_ht_them],
                                    'Chức vụ':[input_chuc_vu_them],
                                    'SDT':[str(input_sdt_them)]
                                    
                                })
                                data_nv_update = pd.concat([new_nhan_vien, data_nhan_vien], ignore_index=True)
                                data_nv_update.to_csv('Danh_sach_nhan_vien.csv', index=False)
                                showMessageBox('Thêm nhân viên', 'Thêm thành công')
                                data_danh_sach_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
            
                                self.uic.tb_danh_sach_nhan_vien.setRowCount(len(data_danh_sach_nhan_vien))
                                self.uic.tb_danh_sach_nhan_vien.setColumnCount(len(data_danh_sach_nhan_vien.columns))
                                self.uic.tb_danh_sach_nhan_vien.setHorizontalHeaderLabels(data_danh_sach_nhan_vien.columns)
                                
                                self.uic.lb_uid_them.clear()
                                self.uic.edt_ten_nhan_vien_them.clear()
                                self.uic.edt_sdt_them.clear()
                                self.uic.lb_chuc_vu_them.clear()
                                
                                
                                
                                for row_index, row in data_danh_sach_nhan_vien.iterrows():
                                    for col_index, value in enumerate(row):
                                        item = QTableWidgetItem(str(value))
                                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                                        self.uic.tb_danh_sach_nhan_vien.setItem(row_index, col_index, item)
                                
                                self.uic.tb_danh_sach_nhan_vien.resizeColumnsToContents()
                                self.uic.tb_danh_sach_nhan_vien.resizeRowsToContents()
                                self.uic.tb_danh_sach_nhan_vien.setColumnWidth(0, 120)
                                self.uic.tb_danh_sach_nhan_vien.setColumnWidth(1, 120)
                                self.uic.tb_danh_sach_nhan_vien.setColumnWidth(2, 120)
                                self.uic.tb_danh_sach_nhan_vien.setColumnWidth(3, 120)
                            else:
                                showMessageBox('Thêm nhân viên', 'SDT không hợp lệ')
                        
                        else:
                            showMessageBox('Thêm nhân viên', 'Chưa chọn chức vụ')
                            
                    elif xu_ly_ten(input_ht_them) == 'Chỉ chứa chữ cái':
                        showMessageBox('Thêm nhân viên', 'Họ tên chỉ chứa chữ cái')
                    elif xu_ly_ten(input_ht_them) == 'Thừa khoảng trắng':
                        showMessageBox('Thêm nhân viên', 'Họ tên thừa khoảng trống')
                    
                    
                else:
                    showMessageBox('Thêm nhân viên', 'Chưa nhập họ và tên')
            
            else:
                showMessageBox('Thêm nhân viên', 'Đã trùng ID')
        
                
            
    def xac_nhan_xoa_nhan_vien(self):
        global chuc_nang_them_xoa
        id_quet = self.uic.lb_uid_xoa.text()
        id_input = self.uic.edt_uid_xoa.text()
        id_xoa = ''
        if chuc_nang_them_xoa == 'Xóa nhân viên':
            if len(id_quet)>0:
                if id_input == id_quet:
                    id_xoa = str(id_input)
                else:
                    id_xoa = str(id_quet)
            else:
                if len(id_input)>9:
                    id_xoa = str(id_input)
                else:
                    id_xoa = ''
            print(id_xoa)        
            if len(id_xoa) == 0:
                showMessageBox('Xóa nhân viên', 'Chưa có ID')
            else:
                data_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
                data_id = list(data_nhan_vien['ID'])
                if id_xoa in data_id:
                    index_id_xoa = data_id.index(id_xoa)
                    ho_ten_xoa = list(data_nhan_vien['Họ và tên'])[index_id_xoa]
                    self.uic.lb_ten_nhan_vien_xoa.setText(f'Họ và tên: {ho_ten_xoa}')
                    data_nhan_vien = data_nhan_vien[data_nhan_vien['ID'] != str(id_xoa)]
                    print(data_nhan_vien)
                    data_nhan_vien.to_csv('Danh_sach_nhan_vien.csv', index=False)
                    showMessageBox('Xóa nhân viên', 'Xóa thành công')
                    self.uic.lb_uid_xoa.clear()
                    self.uic.edt_uid_xoa.clear()
                    self.uic.lb_ten_nhan_vien_xoa.clear()
                    
                    data_danh_sach_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
                    print(data_danh_sach_nhan_vien)
            
                    self.uic.tb_danh_sach_nhan_vien.setRowCount(len(data_danh_sach_nhan_vien))
                    self.uic.tb_danh_sach_nhan_vien.setColumnCount(len(data_danh_sach_nhan_vien.columns))
                    self.uic.tb_danh_sach_nhan_vien.setHorizontalHeaderLabels(data_danh_sach_nhan_vien.columns)
                    
                    for row_index, row in data_danh_sach_nhan_vien.iterrows():
                        for col_index, value in enumerate(row):
                            item = QTableWidgetItem(str(value))
                            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            self.uic.tb_danh_sach_nhan_vien.setItem(row_index, col_index, item)
                    
                    self.uic.tb_danh_sach_nhan_vien.resizeColumnsToContents()
                    self.uic.tb_danh_sach_nhan_vien.resizeRowsToContents()
                    self.uic.tb_danh_sach_nhan_vien.setColumnWidth(0, 120)
                    self.uic.tb_danh_sach_nhan_vien.setColumnWidth(1, 120)
                    self.uic.tb_danh_sach_nhan_vien.setColumnWidth(2, 120)
                    self.uic.tb_danh_sach_nhan_vien.setColumnWidth(3, 120)
                    
                else:
                     showMessageBox('Xóa nhân viên', 'Không tồn tại ID')
            
    
     # Trang Đổi mật khẩu   
    def doi_mat_khau_admin(self):
        self.uic.stackedWidget.setCurrentWidget(self.uic.page_doi_mk)
        self.uic.edt_mk_cu.setEchoMode(QLineEdit.EchoMode.Password)
        self.uic.edt_mk_moi.setEchoMode(QLineEdit.EchoMode.Password)
        
        if self.tin_hieu_page_doi_mk == False:
           self.uic.bt_xac_nhan_doi_mk.clicked.connect(lambda: self.xac_nhan_doi_mk())
           self.uic.bt_trangchu_doi_mk.clicked.connect(lambda: self.mo_trang_chu())
           self.uic.bt_dangnhap_doimk.clicked.connect(lambda: self.trang_dang_nhap_admin())
           self.uic.cb_show_pwd.stateChanged.connect(self.toggle_password_visibility)
           self.tin_hieu_page_doi_mk = True
    
    def toggle_password_visibility(self):
        if self.uic.cb_show_pwd.isChecked():
            self.uic.edt_mk_cu.setEchoMode(QLineEdit.EchoMode.Normal)
            self.uic.edt_mk_moi.setEchoMode(QLineEdit.EchoMode.Normal)
        else:
            self.uic.edt_mk_cu.setEchoMode(QLineEdit.EchoMode.Password)
            self.uic.edt_mk_moi.setEchoMode(QLineEdit.EchoMode.Password)    
        
    def xac_nhan_doi_mk(self):
        data_login = pd.read_csv('Data_Login_Admin.csv')
        input_tk = self.uic.edt_tk_doi_mk.text()
        input_mk_cu = self.uic.edt_mk_cu.text()
        input_mk_moi = self.uic.edt_mk_moi.text()
        
        tk_dung = str(list(data_login['TK'])[0])
        mk_dung = str(list(data_login['MK'])[0])
        
        if input_tk == tk_dung and input_mk_cu == mk_dung:
            if input_mk_moi == mk_dung:
                showMessageBox('Đổi mật khẩu', 'Lỗi trùng MK cũ')
            else:
                input_mk_moi = input_mk_moi.replace(" ", "")
                
                if len(input_mk_moi) >= 8:
                    showMessageBox('Đổi mật khẩu', 'Đổi mật khẩu thành công')
                    mk_dung = input_mk_moi
                    data_new_login = pd.DataFrame({"TK":[tk_dung], 'MK':[mk_dung]})
                    data_new_login.to_csv('Data_Login_Admin.csv', index=False)
                    
                    self.uic.stackedWidget.setCurrentWidget(self.uic.page_LOGIN)

                else:
                    showMessageBox('Đổi mật khẩu', 'MK mới >=8 ký tự, không có dấu space')
        else:
            showMessageBox('Đổi mật khẩu', 'Lỗi nhập TK cũ và MK cũ')              


    # Quên mật khẩu
    
    def quen_mat_khau_admin(self):
        self.uic.stackedWidget.setCurrentWidget(self.uic.page_quen_mk)
        self.uic.edt_cccd.setEchoMode(QLineEdit.EchoMode.Password)
        
        if self.tin_hieu_page_quen_mk == False:
            self.uic.bt_xac_nhan_caplai.clicked.connect(lambda: self.xac_nhan_cap_lai_mk())
            self.uic.bt_dn_quen.clicked.connect(self.trang_dang_nhap_admin)
            self.uic.bt_trang_chu_quen.clicked.connect(self.mo_trang_chu)
            self.uic.cb_s_cccd.stateChanged.connect(self.toogle_show_cccd)
            self.tin_hieu_page_quen_mk = True
        
        
    def toogle_show_cccd(self):
        if self.uic.cb_s_cccd.isChecked():
            self.uic.edt_cccd.setEchoMode(QLineEdit.EchoMode.Normal)    
        else:
            self.uic.edt_cccd.setEchoMode(QLineEdit.EchoMode.Password)
    
    
    def xac_nhan_cap_lai_mk(self):
        input_cccd = str(self.uic.edt_cccd.text())
        input_tk = self.uic.edt_tk_quen_mk.text()
        data_login = pd.read_csv('Data_Login_Admin.csv')
        tk_dung = str(list(data_login['TK'])[0])
        print('input cccd: ', input_cccd)
        print('input tk: ', input_tk)
        
        if input_tk == tk_dung and input_cccd == '01234567891011':
            characters = string.ascii_letters + string.digits  # Bao gồm cả chữ cái hoa, thường và số
            mk_cap_lai = ''.join(random.choice(characters) for _ in range(8))
            
            self.uic.lb_mk_moi_quen.setText(f'MK mới: {mk_cap_lai}')
            data_new_login = pd.DataFrame({'TK':[tk_dung], 'MK':[str(mk_cap_lai)]})
            data_new_login.to_csv('Data_Login_Admin.csv', index=False)
            
            showMessageBox('Quên mật khẩu', 'Cấp lại mật khẩu thành công')
        else:
            self.uic.lb_mk_moi_quen.setText('Không thành công')
            showMessageBox('Quên mật khẩu', 'Cấp lại mật khẩu không thành công')
        
        
    def bat_dau_tu_dong_cham_cong(self):
        self.thread[4] = Cap_nhat_tu_dong_cham_cong(index = 4)
        self.thread[4].start()    
        
    def bat_dau_cap_nhat_danh_sach_nhan_vien(self):
        self.thread[1] = Cap_nhat_so_luong_nhan_vien(index=1)
        self.thread[1].progress.connect(self.so_luong_nhan_vien)
        self.thread[1].start()

    def so_luong_nhan_vien(self, value):
        self.uic.lb_tong_nv.setText(f'Tổng số lượng nhân viên: {value}')

    def bat_dau_cap_nhat_thoi_gian(self):
        self.thread[2] = Cap_nhat_thoi_gian(index=2)
        self.thread[2].progress_hour.connect(self.thoi_gian_thuc)
        self.thread[2].start()

    def thoi_gian_thuc(self, value):
        value = str(value)
        if "Giờ vào làm" in value:
            self.uic.lb_gio_vao_lam.setText(value)
        elif "Giờ tan làm" in value:
            self.uic.lb_gio_tan_lam.setText(value)
        elif "Số ngày nghỉ tối đa" in value:
            self.uic.lb_so_ngay_nghi_max.setText(value)
        else:
            self.uic.lb_ngay_hom_nay.setText(f'Thời gian: {value}')

    def bat_dau_bat_STM(self):
        self.thread[3] = Ket_noi_STM32(index=3)
        self.thread[3].progress.connect(self.ket_noi_stm)
        self.thread[3].start()

    def ket_noi_stm(self, value):
        global id_them_xoa, tin_hieu_reset, tin_hieu_xin_nghi
        value = str(value)
        if "UID" in value:
            id = value.split('Tag UID: ')[-1]
            id = id.replace(" ", "")
            id_them_xoa = id
            print('id_them_xoa: ', id)
            if chuc_nang_them_xoa != 'Thêm nhân viên' and chuc_nang_them_xoa != 'Xóa nhân viên' and tin_hieu_xin_nghi == False:
                tin_hieu_xin_nghi == False
                tin_hieu_reset = True
                self.uic.lb_id.setText(id)
            if chuc_nang_them_xoa == 'Thêm nhân viên':
                tin_hieu_reset = True
                tin_hieu_xin_nghi == False
                self.uic.lb_uid_them.setText(id)
            elif chuc_nang_them_xoa == 'Xóa nhân viên':
                tin_hieu_reset = True
                tin_hieu_xin_nghi == False
                self.uic.lb_uid_xoa.setText(id)
            elif tin_hieu_xin_nghi == True:
                tin_hieu_reset = True
                self.uic.lb_uid_nghi.setText(id)
                self.uid_xin_nghi = id
                data_nv = pd.read_csv('Danh_sach_nhan_vien.csv')
                data_tt_nv_id = data_nv[data_nv['ID'] == id]
                print(data_tt_nv_id)
                if len(data_tt_nv_id) == 0:
                    self.uic.lb_sdt_nghi.setText('ID không tồn tại')
                    self.uic.lb_hoten_nghi.setText('ID không tồn tại')
                    self.uic.lb_chuc_vu_nghi.setText('ID không tồn tại')
                    self.ho_ten_nghi = 'ID không tồn tại'
                    self.sdt_nghi = 'ID không tồn tại'
                    self.chuc_vu_nghi = 'ID không tồn tại'
                else:
                    self.ho_ten_nghi = str(list(data_tt_nv_id['Họ và tên'])[0])
                    self.sdt_nghi = str(list(data_tt_nv_id['SDT'])[0])
                    self.chuc_vu_nghi = str(list(data_tt_nv_id['Chức vụ'])[0])
                    
                    self.uic.lb_sdt_nghi.setText(self.sdt_nghi)
                    self.uic.lb_hoten_nghi.setText(self.ho_ten_nghi)
                    self.uic.lb_chuc_vu_nghi.setText(self.chuc_vu_nghi)
                
                
        elif len(value) > 0:
            self.uic.lb_com.setText(value)
    
            
    def reset(self):
        global tin_hieu_reset
        tin_hieu_reset = True
        
        self.uic.lb_id.setText('')
        self.uic.edt_id_nc.clear()

    def xac_nhan_cham_cong(self):
        global tin_hieu_bat_tat_led_buzzer
        # Lấy ID trên label và edit Text
        id_lb = self.uic.lb_id.text()
        id_edt = self.uic.edt_id_nc.toPlainText()
        id_cham_cong = ''
        if len(id_lb) >= 10:
            if len(id_edt) >= 10 and id_edt.isnumeric() == True:
                if id_lb!=id_edt:
                    id_cham_cong = id_edt
                else:
                    id_cham_cong = id_lb
            else:
                id_cham_cong = id_lb
        else:
            if len(id_edt) >= 10 and id_edt.isnumeric() == True:
                id_cham_cong = id_edt
            else:
                id_cham_cong = ''
        print('ID chấm công: ', id_cham_cong)        
        # Kiểm tra id_cham_cong
        
        
        id_input = id_cham_cong
        data_cham_cong = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')
        data_nv = pd.read_csv('Danh_sach_nhan_vien.csv')
        list_ID_nv = list(data_nv['ID'])
        # B1 Lấy ngày hôm này
        # B2 Xét xem hôm nay vào hay ra
        # B3: Nếu vào thì kiểm tra giờ vào
        # B4: Nếu ra thì kiểm tra giờ ra
        # B5: Xét xem trong tháng đó đã bao nhiêu lần vi phạm
        # Hiển thị thông tin người đó
        if id_input in list_ID_nv:
            date_input = str(datetime.now().strftime('%Y-%m-%d'))
            hour_input = datetime.now().time()
            print('Ngày: ', date_input)
            print('Giờ: ', hour_input)
            # Lọc danh sách nhân viên đã thực hiện trong ngày
            data_xn_cc = data_cham_cong[(data_cham_cong['Ngày'] == date_input) & ((data_cham_cong['Trạng thái'] == 'Vào làm đúng giờ')|(data_cham_cong['Trạng thái'] == 'Vào làm muộn'))]
            print('Danh sach cham cong: ', data_xn_cc)
            # Kiểm tra đã có vào làm trong ngày chưa
            ho_ten_input = data_nv['Họ và tên'][data_nv['ID'] == id_input].values[0]
            ds_trang_thai_input = list(data_xn_cc['Trạng thái'][data_xn_cc['ID'] == id_input])
            
            print('Họ tên: ', ho_ten_input)
            trang_thai = ''
            if id_input not in data_xn_cc['ID'].values :  # Chưa vào làm
                print('Chua vao lam, trang thai vao')
                
                # Đọc giờ vào làm từ file
                with open('gio_vao_lam.txt', 'r') as f:
                    gio_vao_lam = datetime.strptime(f.read().strip(), '%H:%M').time()
                f.close()
                    
                with open('gio_tan_lam.txt', 'r') as f:
                    gio_tan_lam =  datetime.strptime(f.read().strip(), '%H:%M').time()
                f.close()    
                
                if gio_vao_lam >= hour_input:
                    print('Vào làm đúng giờ')
                    trang_thai = 'Vào làm đúng giờ'
                    so_gio_tre = 0

                    showMessageBox(title= "Thông báo", content= f"ID {id_input}, {ho_ten_input}, {trang_thai}")
                elif gio_vao_lam < hour_input and gio_tan_lam >= hour_input:
                    so_gio_tre = (datetime.combine(datetime.today(), hour_input) - 
                                datetime.combine(datetime.today(), gio_vao_lam)).seconds // 60
                    print('Vào làm muộn: ', so_gio_tre)
                    trang_thai = f'Vào làm muộn'
                    
                    showMessageBox(title= "Thông báo", content= f"ID {id_input}, {ho_ten_input}, {trang_thai}")
                elif gio_vao_lam < hour_input and gio_tan_lam < hour_input:   
                    trang_thai = 'Đã hết thời gian làm việc' 
                    showMessageBox(title= "Thông báo", content= f"ID {id_input}, {ho_ten_input}, {trang_thai}")
                    so_gio_tre = 0
                new_row_cham_cong = pd.DataFrame({
                        'ID': [id_input],
                        'Họ và tên': [ho_ten_input],
                        'Thời gian': [hour_input.strftime('%H:%M')],
                        'Ngày': [date_input],
                        'Trạng thái': [trang_thai],
                        'Ghi chú (Phút)':[so_gio_tre]
                    })
                self.uic.lb_ho_ten_nv.setText(f'Họ và tên: {ho_ten_input}')
                self.uic.lb_nv_gio_quet.setText(f'Giờ vào làm: {hour_input}')
                self.uic.lb_tt_nv.setText(f'Trạng thái vào làm: {trang_thai}')
                month_year_now = time.strftime('%m/%Y')
                self.slvp_nv_trong_thang = so_lan_vi_pham_trong_thang(data_cham_cong, id_cham_cong)
                self.uic.lb_sl_vi_pham.setText(f'Số lần vi phạm trong tháng {month_year_now}: {self.slvp_nv_trong_thang}')
                # Thêm hàng mới vào đầu DataFrame
                df_update_cc = pd.concat([new_row_cham_cong, data_cham_cong], ignore_index=True)
                print(df_update_cc)
                # Lưu lại DataFrame vào file CSV
                df_update_cc.to_csv('Danh_sach_cham_cong_nhan_vien.csv', index=False)
                self.uic.lb_id.setText('')
                self.uic.edt_id_nc.clear()
                tin_hieu_bat_tat_led_buzzer = True
                
            else:
                
                # Kiểm tra xem đã hết thời gian vào làm
                # Lấy ngày hôm đó
                month_now = int(time.strftime('%m'))
                year_now = int(time.strftime('%Y'))
                data_cham_cong_ngay = pd.to_datetime(data_cham_cong['Ngày'], format='%Y-%m-%d')
                print('data_cham_cong_ngay: ', data_cham_cong_ngay)
                data_id_input = data_cham_cong[(data_cham_cong['ID'] == id_input) & (data_cham_cong_ngay.dt.month == month_now) & (data_cham_cong_ngay.dt.year == year_now)]
                print('data input id', data_id_input)
                
                with open('gio_tan_lam.txt', 'r') as f:
                    gio_tan_lam = datetime.strptime(f.read().strip(), '%H:%M').time()
                
                # Kiểm tra có đi làm hay không
                trang_thai_cuoi_cung = list(data_id_input['Trạng thái'])[0]
                # data_khong_lam = data_id_input[(data_id_input['Trạng thái'] == 'Đã hết thời gian làm việc')]
                # print('Không làm: ', data_khong_lam)
                
                if trang_thai_cuoi_cung == 'Đã hết thời gian làm việc' or trang_thai_cuoi_cung == 'Nghỉ làm':
                    trang_thai = 'Đã hết thời gian làm việc'
                    so_gio_som = 0
                    showMessageBox(title= "Thông báo", content= f"ID {id_input}, {ho_ten_input}, {trang_thai}")
                else:
                    
                    if gio_tan_lam <= hour_input:
                        print('Tan làm đúng giờ')
                        trang_thai = 'Tan làm đúng giờ'
                        so_gio_som = 0
                        showMessageBox(title= "Thông báo", content= f"ID {id_input}, {ho_ten_input}, {trang_thai}")
                        
                    elif gio_tan_lam > hour_input:  
                        so_gio_som = (datetime.combine(datetime.today(), gio_tan_lam) - 
                                    datetime.combine(datetime.today(), hour_input)).seconds // 60
                        print('Tan làm sớm: ', so_gio_som)
                        trang_thai = f'Tan làm sớm'
                        
                        showMessageBox(title= "Thông báo", content= f"ID {id_input}, {ho_ten_input}, {trang_thai}, 'Sớm {so_gio_som} phút")
                        
                self.uic.lb_ho_ten_nv.setText(f'Họ và tên: {ho_ten_input}')
                self.uic.lb_nv_gio_quet.setText(f'Giờ tan làm: {hour_input}')
                self.uic.lb_tt_nv.setText(f'Trạng thái tan làm: {trang_thai}')   
                self.slvp_nv_trong_thang = so_lan_vi_pham_trong_thang(data_cham_cong, id_cham_cong)
                month_year_now = time.strftime('%m/%Y')
                self.uic.lb_sl_vi_pham.setText(f'Số lần vi phạm trong tháng {month_year_now}: {self.slvp_nv_trong_thang}')

                new_row_cham_cong = pd.DataFrame({
                        'ID': [id_input],
                        'Họ và tên': [ho_ten_input],
                        'Thời gian': [hour_input.strftime('%H:%M')],
                        'Ngày': [date_input],
                        'Trạng thái': [trang_thai], 
                        'Ghi chú (Phút)':[so_gio_som]
                    })
                # Thêm hàng mới vào đầu DataFrame
                df_update_cc = pd.concat([new_row_cham_cong, data_cham_cong], ignore_index=True)
                print(df_update_cc)
                # Lưu lại DataFrame vào file CSV
                df_update_cc.to_csv('Danh_sach_cham_cong_nhan_vien.csv', index=False)
                self.uic.lb_id.setText('')
                self.uic.edt_id_nc.clear()
                tin_hieu_bat_tat_led_buzzer = True
            self.n_da_nghi = so_ngay_da_nghi(id_input)    
            self.uic.lb_so_ngay_nghi.setText(f'Số ngày đã nghỉ tháng {month_year_now}: {self.n_da_nghi}')
        else:
            showMessageBox(title= "Thông báo", content= 'Không tồn tại ID')
            print('Không tồn tại ID')
            self.uic.lb_id.setText('')
            self.uic.edt_id_nc.clear()
        self.reset()    
            
            # Page thông tin chấm công
    def thong_tin_cham_cong(self): # Page thông tin chấm công
        self.uic.stackedWidget.setCurrentWidget(self.uic.page_thong_tin)
        data_thong_tin_cham_cong = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')
        self.data_nhan_vien = pd.read_csv('Danh_sach_nhan_vien.csv')
        
        if self.tin_hieu_page_thong_tin == False:
            self.uic.bt_trang_chu_thong_tin.clicked.connect(self.mo_trang_chu)
            self.uic.btn_xn_tk_tt.clicked.connect(lambda: self.tim_kiem_thong_tin_cham_cong_gui())
            self.uic.btn_save_tk_tt.clicked.connect(lambda: self.save_thong_tin_cham_cong_tk())
            
            # thiet lap combo box
            self.uic.combo_ht.addItem('Không')
            for ht in list(self.data_nhan_vien['Họ và tên']):
                self.uic.combo_ht.addItem(ht)
            
            self.uic.combo_thang.addItem('Không')
            for i in range (1, 10):
                
                self.uic.combo_thang.addItem(f'0{i}')
            for i in range (10, 13):
                self.uic.combo_thang.addItem(f'{i}')
                
            self.uic.combo_nam.addItem('Không')    
            for nam in range (2024, 2029):
                self.uic.combo_nam.addItem(str(nam))
            
            self.uic.combo_tt.addItem('Không')
            for tt in ['Tan làm sớm', 'Vào làm đúng giờ', 'Vào làm muộn', 'Tan làm đúng giờ', 'Nghỉ làm','Đã hết thời gian làm việc']:
                self.uic.combo_tt.addItem(tt)
            
            # Su kien loc nhieu thong tin
            self.uic.combo_ht.currentIndexChanged.connect(self.on_combo_ho_ten)
            self.uic.combo_nam.currentIndexChanged.connect(self.on_combo_nam)
            self.uic.combo_tt.currentIndexChanged.connect(self.on_combo_trang_thai)
            self.uic.combo_thang.currentIndexChanged.connect(self.on_combo_thang)
            self.uic.btn_loc_tt.clicked.connect(self.loc_thong_tin)
            
            # Su kien loc thong ke
            
            
            self.tin_hieu_page_thong_tin = True
            
        
        self.data_thong_tin_cham_cong = data_thong_tin_cham_cong
        self.uic.tb_cc_tt.setRowCount(len(data_thong_tin_cham_cong))
        self.uic.tb_cc_tt.setColumnCount(len(data_thong_tin_cham_cong.columns))
        self.uic.tb_cc_tt.setHorizontalHeaderLabels(data_thong_tin_cham_cong.columns)
         # Tạo font in đậm
        font = QFont()
        font.setBold(True)

        # Áp dụng font in đậm cho các tiêu đề
        for i in range(len(data_thong_tin_cham_cong.columns)):
            item = self.uic.tb_cc_tt.horizontalHeaderItem(i)
            if item:
                item.setFont(font)
                
        for row_index, row in data_thong_tin_cham_cong.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.uic.tb_cc_tt.setItem(row_index, col_index, item)
        self.uic.tb_cc_tt.resizeColumnsToContents()
        self.uic.tb_cc_tt.resizeRowsToContents()
        self.uic.tb_cc_tt.setColumnWidth(0, 120)
        self.uic.tb_cc_tt.setColumnWidth(1, 120)
        self.uic.tb_cc_tt.setColumnWidth(2, 120)
        self.uic.tb_cc_tt.setColumnWidth(3, 120)
        self.uic.tb_cc_tt.setColumnWidth(4, 120)
        self.uic.tb_cc_tt.setColumnWidth(5, 120)      
          
        # Cài đặt bảng tìm kiếm
        self.data_tk = pd.DataFrame({'ID':[],'Họ và tên':[],'Thời gian':[],'Ngày':[], 'Trạng thái':[], 'Ghi chú (Phút)':[]})
            
        self.uic.tb_kq_tk.setRowCount(len(self.data_tk))
        self.uic.tb_kq_tk.setColumnCount(len(self.data_tk.columns))
        self.uic.tb_kq_tk.setHorizontalHeaderLabels(self.data_tk.columns)
         # Tạo font in đậm
        font = QFont()
        font.setBold(True)

        # Áp dụng font in đậm cho các tiêu đề
        for i in range(len(self.data_tk.columns)):
            item = self.uic.tb_kq_tk.horizontalHeaderItem(i)
            if item:
                item.setFont(font)
                
        for row_index, row in self.data_tk.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.uic.tb_kq_tk.setItem(row_index, col_index, item)
        
        self.uic.tb_kq_tk.resizeColumnsToContents()
        self.uic.tb_kq_tk.resizeRowsToContents()
        self.uic.tb_kq_tk.setColumnWidth(0, 120)
        self.uic.tb_kq_tk.setColumnWidth(1, 120)
        self.uic.tb_kq_tk.setColumnWidth(2, 120)
        self.uic.tb_kq_tk.setColumnWidth(3, 120)
        self.uic.tb_kq_tk.setColumnWidth(4, 120)
        self.uic.tb_kq_tk.setColumnWidth(5, 120)
    
    
    def loc_thong_tin(self):
        ds_key_co = []
        self.filter_tt = self.data_thong_tin_cham_cong
        for key in list(self.my_dict_loc.keys()):
            if self.my_dict_loc[key] != 'Không':
                ds_key_co.append(key)
        
        print(ds_key_co)
        
        if len(ds_key_co) > 0:
            for key_co in ds_key_co:
                print(key_co)
                if key_co == 'Tháng':
                    self.filter_tt = self.filter_tt[self.filter_tt['Ngày'].str.split('-').str[1] == self.my_dict_loc[key_co]]
                elif key_co == 'Năm':
                    self.filter_tt = self.filter_tt[self.filter_tt['Ngày'].str.split('-').str[0] == self.my_dict_loc[key_co]]
               
                else:
                    self.filter_tt = self.filter_tt[self.filter_tt[key_co] == self.my_dict_loc[key_co]]
            
            print('LỌC: ', self.filter_tt)
        else:
            self.filter_tt = pd.DataFrame({'ID': [], 'Họ và tên': [], 'Thời gian': [], 'Ngày': [], 'Trạng thái': [], 'Ghi chú (Phút)': []})
            
        self.filter_tt.reset_index(drop=True, inplace=True)
        self.data_tk = self.filter_tt #Gắn để save
        # Đặt lại bảng trước khi điền dữ liệu
        self.uic.tb_kq_tk.clearContents()
        self.uic.tb_kq_tk.setRowCount(0)
        
        # Thiết lập số hàng và cột
        self.uic.tb_kq_tk.setRowCount(len(self.filter_tt))
        self.uic.tb_kq_tk.setColumnCount(len(self.filter_tt.columns))
        self.uic.tb_kq_tk.setHorizontalHeaderLabels(self.filter_tt.columns)
        
        # Điền dữ liệu vào bảng
        for row_index, row in self.filter_tt.iterrows():
            print('Row index:', row_index, 'Row data:', row)
            for col_index, value in enumerate(row):
                print('Col index:', col_index, 'Value:', value)
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.uic.tb_kq_tk.setItem(row_index, col_index, item)
        
        # Điều chỉnh kích thước cột và hàng
        self.uic.tb_kq_tk.resizeColumnsToContents()
        self.uic.tb_kq_tk.resizeRowsToContents()
        self.uic.tb_kq_tk.setColumnWidth(0, 120)
        self.uic.tb_kq_tk.setColumnWidth(1, 120)
        self.uic.tb_kq_tk.setColumnWidth(2, 120)
        self.uic.tb_kq_tk.setColumnWidth(3, 120)
        self.uic.tb_kq_tk.setColumnWidth(4, 120)
        self.uic.tb_kq_tk.setColumnWidth(5, 120)

    
    
    def on_combo_ho_ten(self):
        ht = self.uic.combo_ht.currentText()
        self.uic.lb_ht_tk.setText(ht)
        self.my_dict_loc['Họ và tên'] = ht
        print(self.my_dict_loc)
    
    def on_combo_nam(self):
        nam = self.uic.combo_nam.currentText()
        self.uic.lb_nam_tk.setText(nam)
        self.my_dict_loc['Năm'] = nam
        print(self.my_dict_loc)
    
    def on_combo_thang(self):
        thang = self.uic.combo_thang.currentText()
        self.uic.lb_thang_tk.setText(thang)
        self.my_dict_loc['Tháng'] = thang
        print(self.my_dict_loc)
    
    def on_combo_trang_thai(self):
        tt = self.uic.combo_tt.currentText()
        self.uic.lb_tt_tk.setText(tt)
        self.my_dict_loc['Trạng thái'] = tt
        print(self.my_dict_loc)
    
    def tim_kiem_thong_tin_cham_cong_gui(self):
        try:
            
            str_tk = self.uic.edt_tim_kiem_tt.toPlainText()
            
            data_tk = tim_kiem_thong_tin_cham_cong(str_input= str_tk, data_thong_tin= self.data_thong_tin_cham_cong)
            self.data_tk = data_tk
            print('Data tìm kiếm: ', self.data_tk)
            self.uic.tb_kq_tk.setRowCount(len(data_tk))
            self.uic.tb_kq_tk.setColumnCount(len(data_tk.columns))
            self.uic.tb_kq_tk.setHorizontalHeaderLabels(data_tk.columns)
            
            for row_index, row in data_tk.iterrows():
                for col_index, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.uic.tb_kq_tk.setItem(row_index, col_index, item)
            
            self.uic.tb_kq_tk.resizeColumnsToContents()
            self.uic.tb_kq_tk.resizeRowsToContents()
            self.uic.tb_kq_tk.setColumnWidth(0, 120)
            self.uic.tb_kq_tk.setColumnWidth(1, 120)
            self.uic.tb_kq_tk.setColumnWidth(2, 120)
            self.uic.tb_kq_tk.setColumnWidth(3, 120)
            self.uic.tb_kq_tk.setColumnWidth(4, 120)
            self.uic.tb_kq_tk.setColumnWidth(5, 120)
        except:
            self.data_tk = pd.DataFrame({'ID':[],'Họ và tên':[],'Thời gian':[],'Ngày':[], 'Trạng thái':[], 'Ghi chú (Phút)':[]})

        
    
    def mo_trang_chu(self):
        global chuc_nang_them_xoa
        chuc_nang_them_xoa = ''
        self.uic.edt_tk.clear()
        self.uic.edt_mk.clear()
        self.uic.edt_tk_doi_mk.clear()
        self.uic.edt_mk_cu.clear()
        self.uic.edt_mk_moi.clear()
        self.uic.lb_mk_moi_quen.clear()
        self.uic.edt_tk_quen_mk.clear()
        self.uic.lb_uid_them.clear()
        self.uic.lb_uid_xoa.clear()
        self.uic.edt_ten_nhan_vien_them.clear()
        self.uic.edt_sdt_them.clear()
    
        
        self.uic.stackedWidget.setCurrentWidget(self.uic.page_trangchu)
    
    def save_thong_tin_cham_cong_tk(self):
        data_save = self.data_tk
        file_filter = 'Data File (*.xlsx);; Exel File (*.xlsx *.xls)'
        response, _ = QFileDialog.getSaveFileName(
            parent=self,
            caption='Select a data file',
            directory='Data file.xlsx',
            filter=file_filter,
            initialFilter='Excel File (*.xlsx *.xls)'
        )
        if response:
            
            if response.endswith('.xlsx') or response.endswith('.xls'):
                data_save.to_excel(response, index=False)
            
            else:
                with open(response, 'w', encoding='utf-8') as file:
                    file.write(data_save.to_string(index=False))

        
def showMessageBox(title, content):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText(content)
        msg.setWindowTitle(title)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
        
def kiem_tra_so_dien_thoai_vn(so_dien_thoai):
    # Biểu thức chính quy kiểm tra số điện thoại Việt Nam
    regex = r'^(0[3|5|7|8|9])+([0-9]{8})$'
    
    # Kiểm tra nếu số điện thoại khớp với biểu thức chính quy
    if re.match(regex, so_dien_thoai):
        return True
    else:
        return False
    
def so_lan_vi_pham_trong_thang(data_cham_cong, id_input):
    # Tính số lần vi phạm (tan sớm và vào muộn) trong 1 tháng
    data_cham_cong = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')
    month_now = int(time.strftime('%m'))
    year_now = int(time.strftime('%Y'))
    day_now = int(time.strftime('%d'))
    data_cham_cong['Ngày'] = pd.to_datetime(data_cham_cong['Ngày'], format='%Y-%m-%d')

    # đây là data id của quá khứ
    data_id_input_qk = data_cham_cong[(data_cham_cong['ID'] == id_input) &  (data_cham_cong['Ngày'].dt.day != day_now) & (data_cham_cong['Ngày'].dt.month == month_now) & (data_cham_cong['Ngày'].dt.year == year_now)]
    # Lấy ra danh sách tan làm sớm cập nhật cuối cùng của từng ngày trong quá khứ -> phải kết thúc ngày làm mới biết tan làm sớm không
    # Lấy ra trạng thái chấm công của id đó trong ngày -> xét xem trạng thái cuối cùng có phải 
    data_tan_lam = data_id_input_qk[(data_id_input_qk['Trạng thái'] == 'Tan làm sớm') | (data_id_input_qk['Trạng thái'] == 'Tan làm đúng giờ')]

    # Lấy lần tt cuối cùng tan làm của mỗi ngày
    data_trang_thai_tan_lam_last = data_tan_lam.drop_duplicates(subset='Ngày', keep='first')
    print(data_trang_thai_tan_lam_last)
    tt_tl_cuoi = list(data_trang_thai_tan_lam_last['Trạng thái'])
    n_som = tt_tl_cuoi.count('Tan làm sớm') # đếm số ngày tan làm sớm

    # Số lần vào muộn tính luôn cả hôm nay
    # Lấy ra trạng thái vào cả của hôm nay của id input
    data_id_input_today = data_cham_cong[(data_cham_cong['ID'] == id_input)  & (data_cham_cong['Ngày'].dt.month == month_now) & (data_cham_cong['Ngày'].dt.year == year_now)]
    data_vao_muon = data_id_input_today[data_id_input_today['Trạng thái'] == 'Vào làm muộn']
    
    return n_som + len(data_vao_muon)

def so_ngay_da_nghi(id_input):
    month_now = int(time.strftime('%m'))
    year_now = int(time.strftime('%Y'))
    data = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')

    # Chuyển đổi cột Ngày thành định dạng ngày tháng
    data['Ngày'] = pd.to_datetime(data['Ngày'], format='%Y-%m-%d')
    data_thang = data[(data['Ngày'].dt.month == month_now) & (data['Ngày'].dt.year == year_now)]


    # Lấy dòng đầu tiên của mỗi ID trong mỗi ngày
    result = data_thang.groupby(['Ngày', 'ID']).first().reset_index()
    result_so_ngay_nghi = result[(result['Trạng thái'] == 'Đã hết thời gian làm việc')|(result['Trạng thái'] == 'Nghỉ làm')]
    print('Result so ngay nghi: ', result_so_ngay_nghi)
    so_ngay_da_nghi_lam = list(result_so_ngay_nghi['ID']).count(id_input)
    
    return so_ngay_da_nghi_lam

    



def tim_kiem_thong_tin_cham_cong(str_input, data_thong_tin): #ID,Họ và tên,Thời gian,Ngày,Trạng thái,Ghi chú (Phút)
    data_tk = pd.DataFrame({'ID':[],'Họ và tên':[],'Thời gian':[],'Ngày':[], 'Trạng thái':[], 'Ghi chú (Phút)':[]})
    print('Tìm kiếm: ', str_input)
    str_input = str(str_input)
    if str_input != '':
        # Xóa hết dấu cách 2 đầu
        str_input = str_input.strip()
        if str_input != '':

            
            for row_index in range (0, len(data_thong_tin)):
                for column_index in range (0, len(data_thong_tin.columns)):
                    if str_input.lower() in str(data_thong_tin.iloc[row_index][column_index]).lower():

                        new_data = pd.DataFrame({
                        'ID': [data_thong_tin.iloc[row_index][0]],
                        'Họ và tên': [data_thong_tin.iloc[row_index][1]],
                        'Thời gian': [data_thong_tin.iloc[row_index][2]],
                        'Ngày':  [data_thong_tin.iloc[row_index][3]],
                        'Trạng thái': [data_thong_tin.iloc[row_index][4]],
                        'Ghi chú (Phút)': [str(data_thong_tin.iloc[row_index][5])]
                        
                        })
                        data_tk = pd.concat([data_tk, new_data], ignore_index=True)
                        
                        break
    return data_tk

def xu_ly_ten(str_ten):
    # Kiểm tra ktdb
    re_space = str_ten.replace(' ', '')
    if re_space.isalpha() == True:
        # Xóa 2 khoảng trắng ở 2 đầu nếu có
        strip_space = str_ten.strip()
        # Kiểm tra có thừa khoảng trắng giữa các chữ không
        for i in range (0, len(strip_space)):
            if strip_space[i] == " " and strip_space[i+1] == " ":
                return 'Thừa khoảng trắng'
        return strip_space
    else:
        return 'Chỉ chứa chữ cái'


def thong_ke_trong_thang():
    datacc = pd.read_csv('Danh_sach_cham_cong_nhan_vien.csv')
    # So ngay di lam trong thang của từng nhân viên (tính từ ngày hôm qua -> trước)
    columns = ['ID', 'Họ và tên', 'Tháng', 'Năm', 'Số ngày đi làm', 'Số ngày vào muộn', 'Số ngày về sớm', 'Số ngày nghỉ']
    data_thong_ke = pd.DataFrame(columns=columns)

    date_now = datetime.now().date()
    year_now = datetime.now().year
    month_now = datetime.now().month
    datacc['Ngày'] = pd.to_datetime(datacc['Ngày'])
    list_year_now = datacc['Ngày'].dt.year.unique()
    data_cc_ngay_truoc = datacc[datacc['Ngày'].dt.date != date_now]
    datacc_trang_thai_cuoi_cung = data_cc_ngay_truoc.drop_duplicates(subset='Ngày', keep = 'first', ignore_index=True)

    data_cc_ngay_truoc_vao_lam = data_cc_ngay_truoc[ (data_cc_ngay_truoc['Trạng thái'] == 'Vào làm muộn') | (data_cc_ngay_truoc['Trạng thái'] == 'Vào làm đúng giờ')]
    data_cc_ngay_truoc_vao_lam_muon = data_cc_ngay_truoc_vao_lam[data_cc_ngay_truoc_vao_lam['Trạng thái'] == 'Vào làm muộn']
    datacc_trang_thai_cuoi_cung_tan_lam_som = datacc_trang_thai_cuoi_cung[datacc_trang_thai_cuoi_cung['Trạng thái'] == 'Tan làm sớm']
    datacc_trang_thai_cuoi_cung_nghi_lam = datacc_trang_thai_cuoi_cung[datacc_trang_thai_cuoi_cung['Trạng thái'] == 'Nghỉ làm']

    data_nv = pd.read_csv('Danh_sach_nhan_vien.csv')
    data_id_nv = data_nv['ID']

    # Số ngày đi làm
    for nam in list_year_now:
        # lọc dữ liệu của từng năm
        data_cc_ngay_truoc_vao_lam_nam = data_cc_ngay_truoc_vao_lam[data_cc_ngay_truoc_vao_lam['Ngày'].dt.year == nam]
        data_cc_ngay_truoc_vao_lam_muon_nam = data_cc_ngay_truoc_vao_lam_muon[data_cc_ngay_truoc_vao_lam_muon['Ngày'].dt.year == nam]
        datacc_trang_thai_cuoi_cung_tls_nam = datacc_trang_thai_cuoi_cung_tan_lam_som[datacc_trang_thai_cuoi_cung_tan_lam_som['Ngày'].dt.year == nam]
        datacc_trang_thai_cuoi_cung_nghi_lam_nam = datacc_trang_thai_cuoi_cung_nghi_lam[datacc_trang_thai_cuoi_cung_nghi_lam['Ngày'].dt.year == nam]
        
        
        # Lọc theo từng tháng trong năm  
        print('Năm: ', nam)
        if nam == year_now:
            list_thang_now = [month for month in range (1, month_now+1)]
        else:
            list_thang_now = [month for month in range (1, 13)]
        for thang in list_thang_now:
            data_cc_ngay_truoc_vao_lam_nam_thang = data_cc_ngay_truoc_vao_lam_nam[data_cc_ngay_truoc_vao_lam_nam['Ngày'].dt.month == thang]
            data_cc_ngay_truoc_vao_lam_muon_nam_thang = data_cc_ngay_truoc_vao_lam_muon_nam[data_cc_ngay_truoc_vao_lam_muon_nam['Ngày'].dt.month == thang]
            datacc_trang_thai_cuoi_cung_tls_nam_thang = datacc_trang_thai_cuoi_cung_tls_nam[datacc_trang_thai_cuoi_cung_tls_nam['Ngày'].dt.month == thang]
            datacc_trang_thai_cuoi_cung_nghi_lam_nam_thang = datacc_trang_thai_cuoi_cung_nghi_lam_nam[datacc_trang_thai_cuoi_cung_nghi_lam_nam['Ngày'].dt.month == thang]
            
            # Đếm số lần ID xuất hiện trong bảng vào làm

            dict_so_lan_di_lam = {}
            dict_so_lan_vao_muon = {}
            dict_so_lan_tan_lam_som = {}
            dict_so_lan_nghi_lam = {}
            
            
            for id_nv in data_id_nv:
                n_di_lam = list(data_cc_ngay_truoc_vao_lam_nam_thang['ID']).count(id_nv)
                n_vao_muon = list(data_cc_ngay_truoc_vao_lam_muon_nam_thang['ID']).count(id_nv)
                n_tan_lam_som = list(datacc_trang_thai_cuoi_cung_tls_nam_thang['ID']).count(id_nv)
                n_nghi_lam = list(datacc_trang_thai_cuoi_cung_nghi_lam_nam_thang['ID']).count(id_nv)
                
                dict_so_lan_di_lam[id_nv] = n_di_lam
                dict_so_lan_vao_muon[id_nv] = n_vao_muon
                dict_so_lan_tan_lam_som[id_nv] = n_tan_lam_som
                dict_so_lan_nghi_lam[id_nv] = n_nghi_lam
            
                # Lấy họ và tên từ data_nv
                ho_va_ten = data_nv[data_nv['ID'] == id_nv]['Họ và tên'].values[0]

                # Thêm dữ liệu vào DataFrame kết quả
                data_thong_ke = pd.concat([data_thong_ke, pd.DataFrame({
                    'ID': [id_nv],
                    'Họ và tên': [ho_va_ten],
                    'Tháng': [thang],
                    'Năm': [nam],
                    'Số ngày đi làm': [n_di_lam],
                    'Số ngày vào muộn': [n_vao_muon],
                    'Số ngày về sớm': [n_tan_lam_som],
                    'Số ngày nghỉ': [n_nghi_lam]
                })], ignore_index=True)
                
    data_thong_ke = data_thong_ke.iloc[::-1].reset_index(drop=True)    
    print(data_thong_ke)
    return data_thong_ke

           
if __name__ == "__main__":
    app = QApplication([])
    main_win = MainWindow()
    main_win.show()
    app.exec()
