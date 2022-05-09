import os
import time
import random
from Initialize import misc, settings

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
        self.activeEmotes = {}

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

    def timerDone(self, timer):
        misc.timerDone(timer)
        if "_DELAY" in timer:
            script.writeToFile(self.emotes[timer.split("_")[0]]["hotkey"])
            script.runAHK("PRESS.exe")  # Write and run the hotkey again once the timer is finished.

        if timer in self.activeEmotes.keys():
            self.activeEmotes.pop(timer)


    def success(self, emote):
        hotkey = self.emotes[emote]["hotkey"]
        cooldown = self.emotes[emote]["cooldown"]
        self.timerDone(emote)

        script.writeToFile(hotkey)
        script.runAHK("PRESS.exe")

        if settings["HOTKEY REPRESS DELAY"]:
            misc.setTimer("%s_DELAY" % emote, settings["HOTKEY REPRESS DELAY"])

        misc.setTimer("%s_CD" % emote, cooldown)



    def processIncomingMessage(self, message, user, triggerEmote):
        amtRequired = self.emotes[triggerEmote]["amtRequired"]
        timeLimit = self.emotes[triggerEmote]["timeLimit"]
        hotkey = self.emotes[triggerEmote]["hotkey"]

        if not "%s_CD" % triggerEmote in misc.timers.keys():  # Check if emote is on cooldown
            # Actions if this emote is already active
            if triggerEmote in self.activeEmotes.keys():
                if user in self.activeEmotes[triggerEmote]['contributors']:
                    return
                count = self.activeEmotes[triggerEmote]['count'] + 1
                contributors = self.activeEmotes[triggerEmote]['contributors'].append(user)
            else:  # Actions if this is creating a new entry (And setting timers)
                count = 1
                misc.setTimer(triggerEmote, timeLimit)
                contributors = [user]

            self.activeEmotes[triggerEmote] = {
                "amtRequired": amtRequired,
                "count": count,
                "hotkey": hotkey,
                "contributors": contributors
            }

            if self.activeEmotes[triggerEmote]["count"] >= amtRequired:
                self.success(triggerEmote)



class scriptTasking:
    def __init__(self):
        self.isScriptRunning = False
        self.scriptQueue = []

    def writeToFile(self, whatToWrite):
        with open("output.txt", "w") as f:
            f.write(whatToWrite)

    def runAHK(self, path):
        if self.isScriptRunning:  # Queue the next thing to be run
            self.scriptQueue.append(path)
            return

        self.isScriptRunning = True
        os.system(path)
        self.isScriptRunning = False

        if self.scriptQueue:
            self.runAHK(self.scriptQueue[0])



script = scriptTasking()
spreadsheet = spreadsheetConfig()

