import googlemaps
import polyline
import openpyxl as pyxl
##https://medium.com/future-vision/google-maps-in-python-part-2-393f96196eaf

api_key = "______________"
gmaps = googlemaps.Client(key = api_key)

'Specify Location and Hike Name'
startLoc = '32.793928, -117.238909'
endLoc = '32.747313, -117.197325'
hike = '__________'

'Specify the mode of transportation'
directions_result = gmaps.directions(startLoc, endLoc, mode="walking")
#print(directions_result)

'Grabbing the time and distance of the path from the API'
time  = directions_result[0][ 'legs' ][0][ 'duration' ][ 'text' ]
dist  = directions_result[0][ 'legs' ][0][ 'distance' ][ 'text' ]
print(time, dist)

step = directions_result[0][ 'legs' ][0][ 'steps' ][0][ 'polyline' ][ 'points' ]
#print(step)
steps = directions_result[0][ 'overview_polyline' ][ 'points' ]
#print(steps)
stepdata = polyline.decode(str(steps))
#print(stepdata)


'Excel file must be closed in order to run the program'
filePath =  'C:/Users/Austin Keller/Desktop/Hiking Dashboard/Hiking Data.xlsx'

'Open workbook'
wb = pyxl.load_workbook(filePath)
dataTab = wb['Trail Data']
lastRow =  dataTab.max_row + 1

#stepdata = [('latitude', 'longitude'), ('latitude2', 'longitude2')]
'Write list data to cells after the last filled row in the excel column'
for i in range(len(stepdata)):
    activeRow = lastRow + i
    
    activeCell = 'A' + str(activeRow)
    dataTab[activeCell] = hike

    activeCell = 'B' + str(activeRow)
    lat = stepdata[i][0]
    dataTab[activeCell] = lat

    activeCell = 'C' + str(activeRow)
    long = stepdata[i][1]
    dataTab[activeCell] = long

wb.save(filePath)


