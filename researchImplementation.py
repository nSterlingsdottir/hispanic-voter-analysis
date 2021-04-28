"""Implementation of analysis of ethnicities.
Uses a supplemental data excel spreadsheet to
calculate percentages of different hispanic
ethnicities from a list of hispanic registered
voters from a specified county.
"""

from openpyxl import Workbook, load_workbook
from argparse import ArgumentParser
from math import fsum


# set up argument parser
parser = ArgumentParser()
parser.add_argument('--path', required=True)
parser.add_argument('--county', required=True)
parser.add_argument('--data', required=True)
args = parser.parse_args()

# initialize workbook with openpyxl
wb = Workbook()
wb = load_workbook(filename=args.path, data_only=True)

# enter sheet of input county
county_sheet = wb[args.county]

# initialize workbook with supplemental data
dataWb = Workbook()
dataWb = load_workbook(filename=args.data, data_only=True)
# access to individual worksheets
firstNamesWs = dataWb['FirstN']
lastNamesWs = dataWb['LastN']
zipCodeWs = dataWb['ZIP']

#output workbook
outputWb = Workbook()
outputSheet = outputWb.active
outputSheet.title = 'Output'

# row to start at for saving to output worksheet
activeRow = 2

# dictionary for each sub ethnicity label
ethnicDict = {
    0 : 'Mexican',
    1 : 'PuertoRican',
    2 : 'Cuban',
    3 : 'Dominican',
    4 : 'CostaRican',
    5 : 'Guatemalan',
    6 : 'Honduran',
    7 : 'Nicaraguan',
    8 : 'Panamanian',
    9 : 'Salvadoran',
    10 : 'Argentinian',
    11 : 'Bolivian',
    12 : 'Chilean',
    13 : 'Colombian',
    14 : 'Ecuadorian',
    15 : 'Paraguayan',
    16 : 'Peruvian',
    17 : 'Uruguayan',
    18 : 'Venezuelan',
}

# First Row Labels for each output worksheet column
outputSheet.cell(row=1, column=1).value = 'Zip'
outputSheet.cell(row=1, column=2).value = 'Party'
outputSheet.cell(row=1, column=3).value = 'Gender'
outputSheet.cell(row=1, column=4).value = 'Results'
outputSheet.cell(row=1, column=24).value = 'Others'
for i in range(0,19):
    outputSheet.cell(row=1, column=i+5).value = ethnicDict[i]

#create tuple data structure
initialPercentages = ([], [], [], [])
#format: (last, first, zip, secondLast)

# iterate through rows of people
for person in county_sheet.iter_rows(min_row=2, max_col=11): #min_row=2 skips label row
    
    # parse last name and check if two last names
    # if for some reason someone has a multi part last name
    # with spaces that is considered a single last name
    # this will not work and consider it as two last names
    lastParse = person[2].value.replace('-', ' ').split()
    last = lastParse[0]
    if len(lastParse) > 1:
        secondLast = lastParse[1]
    # parse first name to take only first part
    firstParse = person[3].value.replace('-', ' ').split()
    first = firstParse[0]
    # get zip code first 5 digits
    zipCode = str(person[5].value)
    zipCode = zipCode[:5]
    
    # setting up variables
    percentage = 0.0
    others = 0.0
    
#    # set up tuple data structure
#    if len(lastParse) > 1:
#        initialPercentages = ([], [], [], [])
#        #format: (last, first, zip, secondLast)
#    else:
#        initialPercentages = ([], [], [])
#        #format: (last, first, zip)
    
    #clear lists in tuple
    initialPercentages[0].clear()
    initialPercentages[1].clear()
    initialPercentages[2].clear()
    initialPercentages[3].clear()
    
    # getting percentage from supplemental data
    
    # last name
    for i in range(0, 19):
        for row in lastNamesWs.iter_rows(min_row=2, min_col=1 + (2 * i), max_col=2 + (2 * i)):
            if row[0].value != 'None' and row[0].value == last:
                percentage = row[1].value
                break
            else:
                percentage = 0.0
        initialPercentages[0].append(percentage)
    
    # first name
    for i in range(0, 19):
        for row in firstNamesWs.iter_rows(min_row=2, min_col=1 + (2 * i), max_col=2 + (2 * i)):
            if row[0].value != 'None' and row[0].value == first:
                percentage = row[1].value
                break
            else:
                percentage = 0.0
        initialPercentages[1].append(percentage)
    
    # zip codes
    for i in range(0, 19):
        for row in zipCodeWs.iter_rows():
            if row[0].value != 'None' and str(row[0].value) == zipCode:
                percentage = row[6 + (i * 2)].value
                # grab value for other ethnicities
                others = row[44].value
                break
            else:
                percentage = 0.0
        initialPercentages[2].append(percentage)
    
    # second last (if needed) 
    if len(lastParse) > 1:
        for i in range(0, 19):
            for row in lastNamesWs.iter_rows(min_row=2, min_col=1 + (2 * i), max_col=2 + (2 * i)):
                if row[0].value != 'None' and row[0].value == secondLast:
                    percentage = row[1].value
                    break
                else:
                    percentage = 0.0
            initialPercentages[3].append(percentage)
    
    #if sum > 0 then use value else 1
    
    lastName = 0.0
    firstName = 0.0
    zips = 0.0
    secondLastName = 0.0
    
    # calculate percentage of sub ethnicities
    subEthnicities = []
    for i in range(0, 19):
        # first name
        if fsum(initialPercentages[0]) > 0:
            lastName = initialPercentages[0][i] * .7
        else:
            lastName = 1.0
        
        # last name
        if fsum(initialPercentages[1]) > 0:
            firstName = initialPercentages[1][i] * .3
        else:
            firstName = 1.0
        
        # zip code
        if fsum(initialPercentages[2]) > 0:
            zips = initialPercentages[2][i]
        
        # second last name
        if len(lastParse) > 1:
            if fsum(initialPercentages[3]) > 0:
                secondLastName = initialPercentages[3][i]
            else:
                secondLastName = 1.0
            # calculate percentage with second last name
            subEthnicities.append(lastName * firstName * zips * secondLastName)
        else:
            # calculate percentage with one last name
            subEthnicities.append(lastName * firstName * zips)
    
    # protect against divide by 0
    if fsum(subEthnicities) != 0:
        # Zip, Party, Gender, and Calculated Ethnicity of each person
        outputSheet.cell(row=activeRow, column=1).value = str(person[5].value)[:5]
        outputSheet.cell(row=activeRow, column=2).value = person[10].value
        outputSheet.cell(row=activeRow, column=3).value = person[6].value
        maxPos = subEthnicities.index(max(subEthnicities))
        outputSheet.cell(row=activeRow, column=4).value = ethnicDict[maxPos]
        
        percentageSum = 0.0
        
        # Add percentages to each ethnicity column
        for i in range(0, 19):
            # final step
            calculatedPercentage = subEthnicities[i] / (fsum(subEthnicities) / (1 - others))
            outputSheet.cell(row=activeRow, column=i + 5).value = calculatedPercentage
            outputSheet.cell(row=activeRow, column=i + 5).number_format = '0%'
            percentageSum += calculatedPercentage
        
        # others column
        outputSheet.cell(row=activeRow, column=24).value = 1 - percentageSum
        outputSheet.cell(row=activeRow, column=24).number_format = '0%'
        
        # increment to next row
        activeRow += 1

# save workbook
outputFileName = args.county + '_output.xlsx'
outputWb.save(outputFileName)
