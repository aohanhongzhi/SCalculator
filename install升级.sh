basepath=$(cd `dirname $0`; pwd)
echo $basepath

#测试软件运行
#python3 $basepath/Calculator.py

sudo pip3.5 install openpyxl

sudo pip3.5 install Jinja2


sudo pip3.6 install openpyxl

sudo pip3.6 install Jinja2

#安装软件到系统目录下
sudo cp $basepath/Calculator.py /usr/local/

sudo cp $basepath/iconCal.png /usr/local/

sudo chmod +x /usr/local/Calculator.py

sudo cp $basepath/cal.desktop /usr/share/applications

