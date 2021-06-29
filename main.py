##############################################################################
#
# A simple example of some of the features of the XlsxWriter Python module.
#
import xlsxwriter
import os
import getopt, sys



# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()
folder_svg_path = "E:\Project\python\Read_svg\svg"

# Get full command-line arguments
full_cmd_arguments = sys.argv

# Keep all but the first
argument_list = full_cmd_arguments[1:]
# Get argument
short_options = "p:"
long_options = ["path="]
try:
    arguments, values = getopt.getopt(argument_list, short_options, long_options)
except getopt.error as err:
    # Output error, and return with an error code
    print (str(err))
    sys.exit(2)
for current_argument, current_value in arguments:
    if current_argument in ("-p", "--path"):
        folder_svg_path = current_value


index = 0
directory = os.fsencode(folder_svg_path)


for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".svg"): 
        file = open(folder_svg_path + "\\" +filename,"r")
        lines = file.readlines()
        s = ""
        s = s.join(lines)
        s = s.split("<svg")[1] 
        worksheet.write(index, 0, filename)
        worksheet.write(index, 1, "<svg"+s)
        index+=1

print ("Completed")

workbook.close()