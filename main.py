import sys
import os
import Assignment
import helper
import Project
import DailyTask
import datetime

hwfilename = "hw.xlsx"
assignmentSheetName = "SingleTasks"
projectSheetName = "Projects"
repeatedTaskSheetName = "RepeatedTasks"
repeatedTaskDataSheetName = "RepeatedTasksData"

if not os.path.exists(hwfilename):
    helper.initFile(hwfilename, assignmentSheetName, projectSheetName, repeatedTaskSheetName, repeatedTaskDataSheetName)

done = False #certainly not the most efficient way to do this, but it gets the job done
while not done:
    done = True
    dailyTasks = DailyTask.DailyTask.readInAll(hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName)
    for i in dailyTasks:
        if i.datesArr[-1].replace(hour=0, minute=0, second=0, microsecond=0) < datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0):
            done = False
            helper.updateDaily(False, i.name, hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName, i.datesArr[-1], True)

if len(sys.argv) == 1:
    print("ASSIGNMENTS")
    assignments = sorted(Assignment.Assignment.readInAll(hwfilename, assignmentSheetName), key=lambda x: helper.sortDateLambda(x.deadline))[:19]
    for i in assignments:
        print(i)
    print("\nPROJECTS")
    projects = sorted(Project.Project.readInAll(hwfilename, projectSheetName), key=lambda x: helper.sortDateLambda(x.deadline))
    for i in projects:
        print(i)
    print("\nDAILY TASKS")
    dailyTasks = DailyTask.DailyTask.readInAll(hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName)
    for i in dailyTasks:
        if i.datesArr[-1].replace(hour=0, minute=0, second=0, microsecond=0) == datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0):
            print(i)
elif sys.argv[1] == "-v":
    print("ASSIGNMENTS")
    assignments = sorted(Assignment.Assignment.readInAll(hwfilename, assignmentSheetName), key=lambda x: helper.sortDateLambda(x.deadline))[:19]
    for i in assignments:
        print(i)
    print("\nPROJECTS")
    projects = sorted(Project.Project.readInAll(hwfilename, projectSheetName), key=lambda x: helper.sortDateLambda(x.deadline))
    for i in projects:
        print(i)
    print("\nDAILY TASKS")
    dailyTasks = DailyTask.DailyTask.readInAll(hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName)
    for i in dailyTasks:
        print(i)

elif sys.argv[1] == "-a":
    name = input("Enter the name of the assignment: ")
    category = input("Enter the category of the assignment: ")
    year = int(input("Enter the year of the due date: "))
    month = helper.month2Int(input("Enter the month of the due date: "))
    day = int(input("Enter the day of the due date: "))
    hour = int(input("Enter the hour of the due date (military time): "))
    minute = int(input("Enter the minute of the due date: "))
    second = int(input("Enter the second of the due date: "))
    duedate = datetime.datetime(year, month, day, hour, minute, second, 0)
    if duedate < datetime.datetime.now():
        raise(ValueError("The past is in the past!"))
    location = input("Enter the location (press enter for none): ")
    notes = input("Enter the notes (press enter for none): ")
    assignment = Assignment.Assignment(name, category, duedate, location, notes)
    assignment.append2File(hwfilename, assignmentSheetName)
    print(name, "has been added to assignments.")

elif sys.argv[1] == "-f":
    helper.deleteItemByName(hwfilename, assignmentSheetName, 0)

elif sys.argv[1] == "-p":
    name = input("Enter the name of the project: ")
    year = input("Enter the year of the due date (press enter for no due date): ")
    duedate = None
    if year != "":
        year = int(year)
        month = helper.month2Int(input("Enter the month of the due date: "))
        day = int(input("Enter the day of the due date: "))
        hour = int(input("Enter the hour of the due date (military time): "))
        minute = int(input("Enter the minute of the due date: "))
        second = int(input("Enter the second of the due date: "))
        duedate = datetime.datetime(year, month, day, hour, minute, second, 0)
        if duedate < datetime.datetime.now():
            raise(ValueError("The past is in the past!"))
    notes = input("Enter the notes (press enter for none): ")
    project = Project.Project(name, duedate, notes)
    project.append2File(hwfilename, projectSheetName)
    print(name, "has been added to projects.")

elif sys.argv[1] == "-i":
    helper.deleteItemByName(hwfilename, projectSheetName, 0)

elif sys.argv[1] == "-g":
    name = input("Enter the name of the daily task: ")
    options = ["daily", "weekly", "monthly", "quarterly", "yearly", "every n (where n is a positive integer)"]
    pattern = input("Enter the pattern, options include " + str(options) + ": ")
    patternData = []
    if pattern == "daily":
        pass
    elif pattern == "weekly":
        daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        day = None
        while day != "done":
            day = input("Enter a day of the week (type \"done\" to finish): ")
            if day in daysOfWeek:
                patternData.append(helper.stwd(day))
                patternData = list(set(patternData))
            elif day == "done":
                pass
            else:
                print("Bad weekday string. Ignoring.")
    elif pattern == "monthly":
        day = None
        while day != "done":
            day = input("Enter a day of the month (range [1,28]) (type \"done\" to finish): ")
            if day == "done":
                pass
            elif int(day) > 0 and int(day) < 29:
                patternData.append(int(day))
                patternData = list(set(patternData))
            else:
                print("Bad day of the month. Ignoring.")
    elif pattern == "quarterly":
        day = None
        while day != "done":
            day = input("Enter a day of the quarter (range [1,90]) (type \"done\" to finish): ")
            if day == "done":
                pass
            elif int(day) > 0 and int(day) < 91:
                patternData.append(int(day))
                patternData = list(set(patternData))
            else:
                print("Bad day of the quarter. Ignoring.")
    elif pattern == "yearly":
        day = None
        while day != "done":
            day = input("Enter a day of the year (range [1,365]) (type \"done\" to finish): ")
            if day == "done":
                pass
            elif int(day) > 0 and int(day) < 366:
                patternData.append(int(day))
                patternData = list(set(patternData))
            else:
                print("Bad day of the quarter. Ignoring.")
    elif pattern.split()[0] == "every" and len(pattern.split()) == 2 and pattern.split()[1].isnumeric() and int(pattern.split()[1]) > 0:
        patternData.append(int(pattern.split()[1]))
    else:
        raise(ValueError("Invalid pattern string."))
    pattern += " "
    for i in sorted(patternData):
        pattern += str(i) + " "
    pattern = pattern.strip()
    notes = input("Enter the notes (press enter for none): ")
    dailyTask = DailyTask.DailyTask(name, pattern, notes, None, None)
    dailyTask.append2File(hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName)

elif sys.argv[1] == "-w":
    name = input("Enter the name of the daily task to mark off as successful: ")
    helper.updateDaily(True, name, hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName, datetime.datetime.now(), False)

elif sys.argv[1] == "-l":
    name = input("Enter the name of the daily task to mark off as unsuccessful: ")
    helper.updateDaily(False, name, hwfilename, repeatedTaskSheetName, repeatedTaskDataSheetName, datetime.datetime.now(), False)

elif sys.argv[1] == "-d":
    deletedName = helper.deleteItemByName(hwfilename, repeatedTaskSheetName, 0)
    helper.deleteItemByNameNoInp(hwfilename, repeatedTaskDataSheetName, deletedName, 1)

elif sys.argv[1] == "-h":
    helper.printHelp()

else:
    print("Invalid Arguments")
    helper.printHelp()
