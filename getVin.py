import pyodbc
import sys

# Connection details
server = 'localhost'
database = 'vPICList_Lite1'
username = 'admin'
password = '123'
fileIn = 'vin_2017_2024.csv'
fileOut = 'ForLin.csv'

# Connection string
connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Check if an argument was provided
if len(sys.argv) > 1:
    start_line = int(sys.argv[1])
    print(f"Start at location: {start_line}")
else:
    start_line = 0

# Establish the connection
try:
	connection = pyodbc.connect(connection_string)
	cursor = connection.cursor()
	print("Connection successful!")
except pyodbc.Error as e:
	print("Error connecting to database:", e)
	
with open(fileOut, mode='a') as outfile:
	line_count = 0
	#Read from file
	for v in open(fileIn, 'r', encoding='utf-8-sig'):
		line_count += 1
		if line_count <= start_line:
			continue
		
		vin = v.replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")
		
		if line_count % 10 == 0:
			print(f"Line: {line_count}", end='\r')
			
		if line_count %5000 == 0:
			cursor.close()
			connection.close()
			connection = pyodbc.connect(connection_string)
			cursor = connection.cursor()
	
		cursor.execute("EXEC spVinDecode ?", vin)
		rows = cursor.fetchall()
		
		DisCC     = None
		BodyClass = None
		VehType   = None
		
		for r in rows:
			row = r[1].replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")
			if row == "Displacement(CC)":  # Check Column B for "Displacement (CC)"
				DisCC = r[2]  # Save value from Column C
			elif row == "BodyClass":  # Check Column B for "Body Class"
				BodyClass = r[2]  # Save value from Column C
			elif row == "VehicleType":  # Check Column B for "Vehicle Type"
				VehType = r[2]  # Save value from Column C

		# Write the results to the output CSV file
		outfile.write(f"{vin}, {DisCC}, {BodyClass}, {VehType}\n")
	
	
print(f"Line: {line_count}")
print("Task Complete")