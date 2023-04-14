import csv
import os
import re
import datetime


#class HandleCsv is used to read, create, and write to .csv files using import csv
class HandleCsv:

    def __init__(self, input_file, output_file, coding, header):
        self.read_filename = input_file
        self.write_filename = output_file
        self.ncode = coding
        self.header = header

    def read_csv(self):  #read_csv is used to read in a data file into a list and return that list
        try:  #we use a try/except to catch and handle any errors in opening the file
            with open(self.read_filename, encoding=self.ncode) as read_file:
                csv_data = csv.reader(read_file)
                data_lines = list(csv_data)
            return data_lines

        except FileNotFoundError:
            print(f"Could not find {self.read_filename} file.")
            print("Make sure 'open_order_report.csv' is in this directory.")
            exit(0)
        except Exception as e:
            print(type(e), e)
            exit(0)

    def create_csv(self):  #create_csv creates a new .csv file with a new header
        try:  #we use a try/except to catch and handle any errors in opening the file
            with open(self.write_filename, mode='a', newline='') as new_file:
                csv_create = csv.writer(new_file)
                csv_create.writerow(self.header)
        except Exception as e:
            print(type(e), e)
            exit(0)

    def write_csv(self, *processed_data):  #write_csv write the data to the newly created .csv file
        try:  #we use a try/except to catch and handle any errors in opening the file
            with open(self.write_filename, mode='a', newline='') as write_file:
                csv_write = csv.writer(write_file)
                csv_write.writerow(processed_data)
        except Exception as e:
            print(type(e), e)
            exit(0)


#the EvaluateProcessing class is a company specific class that deals with outside processing and dae calculations
class EvaluateProcessing:

    def __init__(self, csv_handler):
        self.csv_handler = csv_handler

    def processing(self):
        hard_anodize = 12  #12 days processing time
        anodize = 8  #8 days processing time
        passivate = 3  #3 days processing time
        chem_film = 8  #8 days processing time
        final = 3  #3 days processing time
        asap = 5  #5 days processing time
        heat_treat = 10  #10 days processing time
        hone = 20  #20 days processing time
        penetrant = 5  #5 days processing time
        grind = 20  #20 days processing time

        #The following code is used to evaluate the processing time and track the days processed.
        #It also checks to see if expedite options are requested, and if so makes that calcualtion.
        #Finally, it uses the process days count and expedite count to calculate the necessary start
        #processing date, and writes it to the file named 'processing_time.csv'

        for line in self.csv_handler.read_csv()[1:]:
            process_count = 0
            expedite = 0
            if re.search('hard anodize', line[8], re.IGNORECASE) and re.search('anodize', line[8], re.IGNORECASE):
                process_count = process_count + hard_anodize + anodize

            elif re.search('hard anodize', line[8], re.IGNORECASE):
                process_count = process_count + hard_anodize

            elif re.search('anodize', line[8], re.IGNORECASE):
                process_count = process_count + anodize

            if re.search('passivate', line[8], re.IGNORECASE):
                process_count = process_count + passivate

            if re.search('chem film', line[8], re.IGNORECASE) or re.search(
                'conversion coat', line[8], re.IGNORECASE):
                process_count = process_count + chem_film

            if re.search('final', line[8], re.IGNORECASE):
                process_count = process_count + final

            if re.search('asap', line[8], re.IGNORECASE):
                process_count = process_count + asap

            if re.search('heat treat', line[8], re.IGNORECASE):
                process_count = process_count + heat_treat

            if re.search('hone', line[8], re.IGNORECASE):
                process_count = process_count + hone

            if re.search('penetrant', line[8], re.IGNORECASE):
                process_count = process_count + penetrant

            if re.search('grind', line[8], re.IGNORECASE):
                process_count = process_count + grind

            if re.search('expedite', line[8], re.IGNORECASE):
                expedite = round(process_count - (process_count * .3 + .5))

            process_date = self.calculate_process_date(
                line[7], process_count)  #call the date calculator

            if expedite == 0 and process_count == 0:  #write to the 'processing_time.csv' file
                self.csv_handler.write_csv(line[1], line[4], line[5], line[7], "", "", process_date)
            elif expedite == 0:
                self.csv_handler.write_csv(line[1], line[4], line[5], line[7], str(process_count) + ' DAYS', "", process_date)
            else:
                self.csv_handler.write_csv(line[1], line[4], line[5], line[7], str(process_count) + ' DAYS', str(expedite) + ' DAYS', process_date)

    #This method takes in the due date and processing days and returns the MUST start process date
    def calculate_process_date(self, due_date, process_days):
        sub_date = datetime.datetime.strptime(
            due_date, '%m/%d/%Y') - datetime.timedelta(days=process_days)
        return sub_date.strftime('%m/%d/%Y')


#main - where everything all comes together
def main():
    os.system('cls' if os.name == 'nt' else 'clear') #clear the screen (linux)
    if os.path.exists('processing_time.csv'): #check for existing 'processing_time.csv' (linux)
        os.rename('processing_time.csv', 'old.processing_time') #if 'processing_time.csv' exists, rename it as old.processing_time (linux)
    proceed = ''
    #here we give the user info on how to prepare the input file for processing
    print("This program takes the company's Open Order Report and")
    print("calculates the due dates and processes to evaluate the")
    print("necessary processing time to allow for both normal and")
    print("expedited outside processing. It then creates a new file")
    print("named 'processing_time.csv' that can be used to evaluate")
    print("lead times and prospective upcoming due dates. Before")
    print("proceeding, please make certain the you have the file")
    print("'open_order_report.csv' in this working directory.\n")

    while proceed != 'y' or proceed != 'n':  #this is where the objects are instantiated and called
        proceed = input('Are you ready to proceed? y/n')
        if proceed == 'y':
            #the header can be customized below for various output files
            header = [
            'CUSTOMER', 'PART NUMBER', 'QUANTITY DUE', 'DUE DATE',
            'NORMAL OUTSIDE PROCESSING', 'EXPEDITED OUTSIDE PROCESSING',
            'MUST START PROCESS DATE'
            ]
            #instantiate the HandleCsv object and pass it the necessary args
            handle_csv = HandleCsv('open_order_report.csv', 'processing_time.csv', 'utf-8', header)
            #create the new .csv file
            handle_csv.create_csv()
            #instantiate the EvaluateProcessing object
            evaluate_processes = EvaluateProcessing(csv_handler=handle_csv)
            evaluate_processes.processing()
            print("Creating 'processing_time.csv'...")
            print("Writing data...")
            print("Data write is complete.")
            print("processing_time.csv is ready for evaluation.")
            break

        else:
            if proceed == 'n':
                break


main()  #launch main
