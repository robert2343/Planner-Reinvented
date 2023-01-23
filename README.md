# Planner Reinvented
Keep track of what you need to do on the command line.

## Introduction

This is a to-do-list application in Python. It will use an Excel (.xlsx) file named "hw.xlsx" that lists all of the things that you need to do, in three categories: "assignments", "projects", and "daily". The file to run is `main.py`. If when the program is run, the Excel file is not found, it will be initialized.

## Assignments

An assignment is any task that you need to have done by a particular time, and you only need to do once. This works good for shorter term things that have deadlines (such as at school or work perhaps), or anything that only needs to be done once. It can also be useful for remembering to go to events and meetings. To add a new assignment, use the `-a` command line argument. To remove an assignment use the `-f` command line argument.

## Projects

A project is similar to an assignment, where it only needs to be done once and has a single deadline, but projects will display separately from assignments, so that you can be aware of them further in advance, since they usually take longer to complete. They also do not have any way to store information for a corresponding course or a location. To add a new project, use the `-p` command line argument. To remove a project use the `-i` command line argument.

## Daily Tasks

A daily task is something that you have to do every day, or on certain weekdays every week (such as every Saturday, or every Wednesday and every Thursday, etc.). These have a special function to be marked off for every day, and record your history of successfully completing it or failing to do it, the percent of days you successfully completed it, and the last time you completed it. Using a monthly, quarterly, or yearly pattern, all of which are similar to the weekly pattern, is also supported. To add a new daily task, use the `-g` command line argument. To remove a daily task use the `-d` command line argument.

### Recording Daily Tasks

To mark off a daily task as a success, use the `-w` command line argument. To dismiss a daily task for the current day (and mark it off as failed), use the `-l` command line argument. If you do not successfully complete it and mark it off, it will be updated as unsuccessful on the next day when you run the program.

## Disclaimer

It is recommended to use these command line arguments to edit the Excel file, but you can directly edit if you really want to and understand how. The Excel file follows a very specific format, and will break the program if you edit it incorrectly.