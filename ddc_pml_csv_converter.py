import csv
import sys
import os
#lovingly crafted by Brian Sorensen (u0922372)
#this is built to interface with DataCenter Port Make Live Request Spreadsheet vX.X

#collect source CSV file from user
user_input = raw_input("Enter full filepath of source CSV: ")
source_file = user_input

#designate output CSV file path
output_file = os.getcwd() + '/Desktop/output.csv'

with open(source_file,'r') as csvinput:

    with open(output_file, 'w') as csvoutput:
        writer = csv.writer(csvoutput, lineterminator='\n')
        reader = csv.reader(csvinput)

        #write the header row
        headers = ("node","ports","port_name","description","encap","port_type","force")
        writer.writerow(headers)

        #for loop to create output csv rows from source csv
        for row in reader:
            #skip the header row, any empty row, and the 24 line item limiter row
            if "Destination" in row[14]:
                continue
            if not row[2] and not row[3]:
                continue
            if "Please" in row[0]:
                break
            #cell represents the row being built each loop
            #each row needs the following 8 items (in order):
            #node,ports,port_name,description,encap,port_type,force
            cell = range(0,7)
            #node
            cell[0] = row[14]
            #ports
            cell[1] = row[15]
            #port_name
            cell[2] = str(row[2]) + "-" + str.replace(str(row[3]), "/", "-")
            #description
            if row[1] and row[0]:
                cell[3] = "Barcode: " + str(row[1]) + " Serial: " + str(row[0])
            elif row [1]:
                cell[3] = "Barcode: " + str(row[1])
            elif row [0]:
                cell[3] = "Serial: " + str(row[0])
            else:
                cell[3] = ""
            #force
            #script is hardcoded to not overwrite existing port configs
            cell[6] = "false"
            #encap
            for vlan in row[16].split(','):
                if "-" in vlan:
                    for vlanrange in range(int(vlan.split('-')[0]),int(vlan.split('-')[1])):
                        cell[5] = "trunk"
                        cell[4] = vlanrange
                        writer.writerow(cell[:7])
                    continue
                #port_type
                #if more than one vlan is specified, assign trunk port type
                if len(row[16].split(',')) >= 2:
                    cell[5] = "trunk"
                else:
                    cell[5] = "access"
                cell[4] = vlan
                writer.writerow(cell[:7])

#Notify user the conversion is complete and where to find the file
print("Output at: " + output_file)
