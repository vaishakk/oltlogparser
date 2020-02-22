# from xlwt import Workbook
# import xlwt
import os
import pandas as pd

#ezxf = xlwt.easyxf


def format_name(name):
	ph_no = ''
	pos = name.find('FTTH')
	if pos > -1:
		if pos == 0:
			name = name[pos+5:]
		else:
			name = name[:pos] + name[pos+5:]
	pos = name.find('29')
	if pos > -1:
		ph_no = name[pos:pos+7]
		if pos == 0:
			name = name[pos + 7:]
		else:
			name = name[:pos] + name[pos + 7:]
	if name[0] == '-':
		name = name[1:]
	return [name, ph_no]

def readparams(log_file_name,op_file_name):
	olt_data = pd.DataFrame(columns = ['Port Number', 'ONU ID', 'Name', 'Phone Number', 'MAC ID', 'VLAN'])
	log_file = open(log_file_name, 'r')
	line=log_file.readline()
	#print(log_file)
	out = []
	portNum = ''
	k = 0
	while(line):
		#print(line)
		pos = line.find('interface epon-olt')
		if pos>-1:
			portNum = line[pos+18:pos+23]
			#print(portNum)
			line=log_file.readline()
		else:
			posONU = line.find('onu')
			if posONU>-1:
				ONUNum = line[posONU+4:]
				line = log_file.readline()
				posName = line.find('description')
				while posName == -1:
					line = log_file.readline()
					posName = line.find('description')
				name = line[posName+13:-3]
				[name,ph_no] = format_name(name)
				posMAC = line.find('mac')
				while (posMAC==-1):
					line=log_file.readline()
					posMAC = line.find('mac')
				mac = line[posMAC+4:posMAC+21]
				posVLAN = line.find('vlan')
				while (posVLAN==-1):
					line=log_file.readline()
					posVLAN = line.find('vlan')
				posVLAN = posVLAN = line.find('tag')
				if posVLAN==-1:
					vlan = ''
				else:
					vlan = line[posVLAN+4:posVLAN+7]
				#print(ONUNum+':'+name+':'+mac+':'+vlan)
				olt_data.loc[k] = [portNum, ONUNum, name, ph_no, mac, str(vlan)]
				k = k + 1
		line = log_file.readline()
	olt_data.to_csv(op_file_name)


ip_path = os.path.join(os.environ['userprofile'], 'Desktop\\oltlog.log')
op_path = os.path.join(os.environ['userprofile'], 'Desktop\\oltlog.csv')
print('File path: {}'.format(op_path))
#_ = input("PRESS ENTER")
readparams(ip_path, op_path)
_ = input("PRESS ENTER")
