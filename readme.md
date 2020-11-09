
# SELENIUM PROJECT

- **Made by**: *Shan* :+1: 			 
- **Email**: *sontran.code@gmail.com*.		 
- **Protonmail**: *jayce7g@protonmail.com*.	 

----------------------------------------
- *Sử dụng chủ yếu cho mục đích lấy dữ liệu từ website xuống database.*

- *Sử dụng database là sql server để lưu trữ các account đăng nhập và dữ liệu từ website*. 

- Một số chỉ dẫn cơ bản:
- bước 1:  mở file SQL_server_connection. chỉnh sửa lại theo form:
  - dòng 1:  ip server máy muốn kết nối đến.
  - dòng 2: tên database mới tạo
  - dòng 3: user name để đăng nhập đến sql
  - dòng 4: password để đăng nhập đến sql
- bước 2: chạy file push.bat
----------------------------------------

# Sơ đồ thư mục:


- **SELENIUM_PUSH_DATA**:
	- setting:
	   - log/*log_thanhcong* : chứa dòng dữ liệu đúng.
	   - log/*log_thatbai*	: chứa các dòng dữ liệu bị sai khi ngoại lệ xử lý được.
	   - log/*logfile_error* : chứa các dòng dữ liệu bị sai khi ngoại lệ không xử lý được.
	- src:
	  - setting/*geckodriver.log* : file log khi cài đặt driver browser.
	  - setting/*pythonsetup.exe*	: Gói cài đặt python 3.7.4
	  - setting/*requirements.txt* : Chứa các thư mục cần install
	  - src/*account_push.txt*
	  - src/*chromedriver.exe*
	  - src/*SQL_server_connection.txt*
  - *pull_to_file.py* : file source code.
  - *pull.bat*: file batch khởi tạo chương trình.
  ----------------------------------------
**Library**: *Bs4,Selenium,Openpyxl,Re,String,Os,Socket,Sys,Codecs,Pyodbc,Logging,Argparse,Threading,Time,Random*

----------------------------------------
- **Bs4**: BeautifulSoup crawl dữ liệu về.
- **Selenium**: webdriver,Keys,By,Options,TimeoutException,WebDriverWait,ActionChains:	Tạo bot.
	- **webdriver**: Tạo một browser mới mỗi lần khởi động bot.
	- **Keys**: Chủ yếu gửi data cần đến element(thường thì là username, password và keywords).
	- **Options**: Tích hợp các tùy chọn cho browser mới khởi tạo.
	- **ActionChains**: khởi tạo action thao tác với quyền cao.
	- **WebDriverWait**: Sử dụng để chờ cho đến khi element được click.
	
---------------------------------------
- **Openpyxl**: load_workbook,Workbook: Lib tạo kết nối từ python đến excel.
- **datetime**: datetime,timedelta.
- **webdriver_manager.chrome**: ChromeDriverManager.
- **Re** : Viết regex. Mục đích sử dụng:
	- Dùng để ràng buộc chỉ định 1 chuổi chỉ bao gồm 17 kí tự.
- **String**: Lấy ra bảng chữ cái alphabels. Mục đích sử dụng:
	- Tạo mảng lưu trữ lại các kí tự alphabels.
	- Loop mảng này để lưu trữ dữ liệu vào excel.
- **Os,Socket**:
	- os: Sử dụng để truy xuất lấy ra folder nguồn trên máy tính. (folder download file)
	- socket: bắt lỗi socket.error.
- **Sys**: Dung ossys kill chrome.
	- Kill browser hiện tại để khởi tạo lại browser mới khi trang bị đứng.

	
- **Codecs**
- **Pyodbc**: 
	- Kết nối python với SQL Server. Tạo truy vấn từ bề mặt py đến MSSQL.
- **Logging**: 
	- Log lỗi, log log.
- **Argparse**
- **Threading**: 
	- Lấy Timer tạo biến thời gian chờ cho từng khối code, Mục đích:
	- Timer hẹn giờ cho sleep từ lib time.
- **Time**: Sử dụng sleep để khiến process sleeping, mục đích:
	- Đợi trang load.
- **Random**: 
	- Randint trong randoms, kết hợp với sleep trong time, tạo khoảng ngủ mô phỏng thời gian chờ như người thật.
----------------------------------------