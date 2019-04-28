basepath=$(cd `dirname $0`; pwd)
echo $basepath
#先安装pip

sudo python3 $basepath/get-pip.py

#deepin had Qt,so you shouldn't download pyqt5
#sudo pip3.5 install pyqt5

sudo pip3.5 install pandas

sudo pip3.5 install xlwt

sudo pip3.5 install xlrd

sudo pip3.5 install openpyxl

sudo pip3.5 install Jinja2

sudo pip3.6 install pandas

sudo pip3.6 install xlwt

sudo pip3.6 install xlrd

sudo pip3.6 install openpyxl

sudo pip3.6 install Jinja2

sudo pip3.7 install pandas

sudo pip3.7 install xlwt

sudo pip3.7 install xlrd

sudo pip3.7 install openpyxl

sudo pip3.7 install Jinja2
#测试软件运行

#python3 $basepath/Calculator.py

#安装软件到系统目录下
sudo cp $basepath/Calculator.py /usr/local/
sudo cp $basepath/iconCal.png /usr/local/

sudo chmod +x /usr/local/Calculator.py

sudo cp $basepath/cal.desktop /usr/share/applications

