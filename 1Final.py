import os

path = os.getcwd()

command_0 = "python task.py"
command_1 = "python task_1.py"
command_2 = "python task_2.py"
command_3 = 'result.xlsx'

general = [command_0, command_1, command_2, command_3]
for i in general:
    res = os.system(i)