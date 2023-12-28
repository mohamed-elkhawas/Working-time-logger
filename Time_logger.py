from os import path
from datetime import datetime
from keyboard import wait
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
from win10toast import ToastNotifier

def write_aligned(idx, txt):
	ws[idx] = txt
	ws[idx].alignment = Alignment(horizontal='center')

toaster = ToastNotifier()
filename = 'work_log.xlsx'
delta_time_zero = datetime.now() - datetime.now()
current_ts_row  = 2
last_day_row = 1
last_day = ""
last_day_value = delta_time_zero
working = False

def add_new_entry():
	
	time_worked = ""
	
	global current_ts_row
	global last_day_row
	global last_day
	global last_day_value
	global working
	global start_time

	# Get the current timestamp
	timestamp_now = datetime.now()
	today = str(timestamp_now.strftime('%Y-%m-%d'))
	
	# Write the timestamp and key combination to the worksheet 
	if not working:
		write_aligned('A'+str(current_ts_row) , timestamp_now.strftime('%Y-%m-%d %H:%M:%S'))
		working = True
		start_time = timestamp_now
		
	else:

		working = False		
		start_day = str(start_time.strftime('%Y-%m-%d'))
		
		if start_day != today:
			new_end_time = datetime.strptime(start_day+" 23:59:59","%Y-%m-%d %H:%M:%S")

			write_aligned('B'+str(current_ts_row) , new_end_time.strftime('%Y-%m-%d %H:%M:%S'))
			diff = new_end_time - start_time + datetime.strptime("1","%S")
			write_aligned('C'+str(current_ts_row) , str(diff).split(" ")[1])
			current_ts_row += 1

			if start_day == last_day:
				last_day_value += diff
				write_aligned('E'+str(last_day_row), str(last_day_value).split(" ")[1])

			else:
				last_day_row += 1
				write_aligned('D'+str(last_day_row), start_day)
				write_aligned('E'+str(last_day_row), str(diff).split(" ")[1])
				last_day = start_day
				last_day_value = diff


			new_start_time = datetime.strptime(today+" 0:0:0","%Y-%m-%d %H:%M:%S")
			write_aligned('A'+str(current_ts_row) , new_start_time)
			write_aligned('B'+str(current_ts_row) , timestamp_now.strftime('%Y-%m-%d %H:%M:%S'))
			diff = timestamp_now - new_start_time
			write_aligned('C'+str(current_ts_row) , diff)
			current_ts_row += 1

			last_day_row += 1
			write_aligned('D'+str(last_day_row), today)
			write_aligned('E'+str(last_day_row), diff)
			last_day = today
			last_day_value = diff

		else:

			write_aligned('B'+str(current_ts_row) , timestamp_now.strftime('%Y-%m-%d %H:%M:%S'))
			diff = timestamp_now - start_time
			write_aligned('C'+str(current_ts_row) , diff)
			current_ts_row += 1

			if today == last_day:
				last_day_value += diff
				write_aligned('E'+str(last_day_row), last_day_value)

			else:
				last_day_row += 1
				write_aligned('D'+str(last_day_row), today)
				write_aligned('E'+str(last_day_row), diff)
				last_day = today
				last_day_value = diff
			
			# just a procastination counter measure i hope it works ;)
			# if int(last_day_value.strftime('%H')) > 8 :
			# 	print("you may go to sleep now ;)")
			# else:
			time_worked = str(last_day_value).split(".")[0] 
			print("time worked today = " + time_worked)
			try:
				print("time to finish = " + str(datetime.strptime("8","%H") - last_day_value).split(".")[0].split(" ")[1]) 
			except:
				pass

	# Save the changes to the workbook 
	wb.save(filename)
	return time_worked

# Check if the file exists
if path.isfile(filename):

	# if you can read the file
	try: 
		
		# load the workbook and select the active worksheet
		wb = load_workbook(filename)
		ws = wb.active

		for i in range(2,100000):

			# wait until fist empty cell 
			try:
				str(ws['A'+str(i)].value[0])
			except:
				current_ts_row = i
				break

		for i in range(2,100000):

			if ws['D'+str(i)].value == None:
				last_day_row = i -1
				if i != 2:
					last_day = ws['D'+ str(i-1)].value
					last_day_value = ws['E'+ str(i-1)].value
				break

	except:
		pass
else:

	# If the file doesn't exist, create a new workbook and worksheet
	wb = Workbook()
	ws = wb.active

	# Add headers to the worksheet 
	ws['A1'] = 'Start time'
	ws['B1'] = 'stop time'
	ws['C1'] = 'working time'
	ws['D1'] = 'the day'
	ws['E1'] = 'total working time'

	for c in ['A','B','C','D','E','F']:
		ws.column_dimensions[c].width = 23

	for row in ws.rows:
		for cell in row:
			cell.alignment = Alignment(horizontal='center')

# Start the keylogger 
toaster.show_toast("Work logger is Ready", " ", duration=1, threaded=True)
print("ready")

while True:

	try:

		# Wait for a key combination to be pressed 
		wait(r'q+1')
	
		time_worked = add_new_entry()
		
		show_time = 2
		if working:
			toaster.show_toast("START WORKING", " ", duration=show_time, threaded=True)
			print("started")
		else:
			toaster.show_toast("STOP WORKING    worked: "+time_worked, " ", duration=show_time, threaded=True)
			print("stopped")
			
	except KeyboardInterrupt:

		# If the user presses Ctrl+C, exit the loop 
		if working:
			add_new_entry()

		break
