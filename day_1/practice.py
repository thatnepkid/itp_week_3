import openpyxl

wb = openpyxl.load_workbook("C:\\Users\GorkhaliSquad\Documents\VetsInTech\itp_week_3\day_1\lecture.xlsx")

sheet = wb['Sheet1']

for i in range(1,8):
    date = "A" + str(i)
    date_cell = sheet[date]

    amount = "C" + str(i)
    amount_cell = sheet[amount]

    fruit = "B" + str(i)
    fruit_cell = sheet[fruit]
    print("On the date of " + str(date_cell.value) + ", " + str(amount_cell.value) + " amount of " + fruit_cell.value + "!" )
