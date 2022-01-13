import os
import time
import random

try:
    import xlrd
    import xlsxwriter
except ImportError as e:
    print(e)
    raise ImportError(">>> One or more required packages are not properly installed! Run INSTALL_REQUIREMENTS.bat to fix!")



def stopBot(err):
    print(">>>>>---------------------------------------------------------------------------<<<<<")
    print(err)
    print(">>>>>----------------------------------------------------------------------------<<<<<")
    time.sleep(3)
    quit()



class spreadsheetConfig:
    def __init__(self):
        self.emotes = {}
        self.readSpreadsheet()

    def formatXlsx(self):
        print("Formatting world xlsx")
        try:
            with xlsxwriter.Workbook('../Config/Emotes.xlsx') as workbook:
                worksheet = workbook.add_worksheet('Emote Config')
                format_graybkg = workbook.add_format({'bold': True, 'center_across': False, 'font_color': 'white', 'bg_color': 'gray'})
                format_bold = workbook.add_format({'bold': True, 'center_across': True, 'font_color': 'white', 'bg_color': 'black'})
                format_light = workbook.add_format({'bold': True, 'center_across': True, 'font_color': 'black', 'bg_color': '#DCDCDC', 'border': True})
                format_black = workbook.add_format({'bold': True, 'center_across': True, 'font_color': 'black', 'bg_color': 'black'})
                materialColorCodes = ["#FFCDD2", "#E1BEE7", "#D1C4E9", "#BBDEFB", "#B2EBF2", "#B2DFDB", "#C8E6C9", "#DCEDC8", "#F0F4C3", "#FFECB3"]
                random.shuffle(materialColorCodes)
                worksheet.write(0, 0, "Emote", format_light)
                worksheet.write(0, 1, "Amt Emotes Req", format_light)
                worksheet.write(0, 2, "Time Limit (sec)", format_light)
                worksheet.write(0, 3, "Cooldown (sec)", format_light)
                worksheet.write(0, 4, "Hotkey Response", format_light)
                for option in range(5):
                    print(option)
                    format_option = workbook.add_format({'border': True, 'bg_color': materialColorCodes[int(option)]})
                    worksheet.set_column(option, option, 25, format_option)
        except PermissionError:
            stopBot("Can't open the Emotes file. Please close it and make sure it's not set to Read Only.")



    def readSpreadsheet(self):
        if not os.path.exists("../Config/Emotes.xlsx"):
            self.formatXlsx()
        workbook = xlrd.open_workbook("../Config/Emotes.xlsx")
        for sheet in workbook.sheets():
            for row in range(1, sheet.nrows):
                emoteName = sheet.cell_value(row, 0)
                emoteDict = {
                    "amtRequired": sheet.cell_value(row, 1),
                    "timeLimit": sheet.cell_value(row, 2),
                    "cooldown": sheet.cell_value(row, 3),
                    "hotkey": sheet.cell_value(row, 4),
                }

                self.emotes[emoteName] = emoteDict

    def processIncomingMessage(self, message, user, triggerEmote):
        print(triggerEmote + " Detected!")
        print(self.emotes[triggerEmote])
        amtRequired = self.emotes[triggerEmote]["amtRequired"]
        timeLimit = self.emotes[triggerEmote]["timeLimit"]
        cooldown = self.emotes[triggerEmote]["cooldown"]
        hotkey = self.emotes[triggerEmote]["hotkey"]



spreadsheet = spreadsheetConfig()

