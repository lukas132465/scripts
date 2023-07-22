import os

 
output = os.popen('wmic process get description, processid').read()
output = [line.split()[:2] for line in output.splitlines() if line and line.strip()]
output = [line for line in output if 'outlook' in line[0].lower()]

for line in output:
    os.system('taskkill /PID ' + line[1] + ' /F')

os.system('start outlook')
