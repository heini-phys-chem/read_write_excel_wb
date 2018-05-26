# following modules need to be installed: xlwt xlrd xlutils
# sudo pip install <module>

import sys
import xlrd 
import xlwt
from xlutils.copy import copy

def get_sequences(lst):
  ''' function to get number of 'good' sequencies
  '''
  counter = 0
  for i in range(len(lst) - 2):
    if lst[i] != lst[i+1] and lst[i] != lst[i+2] and lst[i+1] != lst[i+2]:
      counter += 1

  return counter, len(lst) - 2

if __name__ == "__main__":
  ''' script to modify an excel workbook:
      reads in different sequencies and checks for 'good' ones
      and writes the results to the end of every row
  '''

  # tubl check
  if len(sys.argv) < 2:
    print "\nPlease give an excel workbook as input file!\n"
    exit(1)

  # read in excel workbook
  filename = sys.argv[1]
  rb = xlrd.open_workbook(filename)
  sheet = rb.sheet_by_index(0)

  # number of rows (number of animals)
  numAnimals = sheet.nrows 
  end_row = sheet.ncols

  # declare lists
  lists, counters, totals = [ [] for i in range(3) ]

  # read in sequencies and store them as lists
  for i in range(numAnimals):
    lists.append(list(str(sheet.cell(i,1).value)))

  # get number of 'good' sequencies
  for lst in lists:
    counter, total = get_sequences(lst)
    counters.append(counter)
    totals.append(total)

  # write to workbook
  wb = copy(rb)
  sheet = wb.get_sheet(0)

  for i in range(numAnimals):
    sheet.write(i, end_row, counters[i])
    sheet.write(i, end_row + 1, totals[i])
    sheet.write(i, end_row + 2, "%.2f" % ( float(counters[i]) / float(totals[i]) * 100) )
  
  # save woekbook
  wb.save(filename)
