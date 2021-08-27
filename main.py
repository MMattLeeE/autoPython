
import datetime as dt
import win32com.client


task = win32com.client.Dispatch('Schedule.Service')
task.Connect()
root_folder = task.GetFolder('\\')
newtask = task.NewTask(0)

# Trigger
#start_time = dt.datetime.now() + dt.timedelta(minutes=1)
start_time=dt.datetime(2021,8,4,8,30,0,0) # 8/26/2021 8:30 am

TASK_TRIGGER_TYPE = 3 # value 1 for one time 2 for daily
trigger = newtask.Triggers.Create(TASK_TRIGGER_TYPE)
trigger.DaysOfWeek = 62
trigger.StartBoundary = start_time.isoformat()


# Action
TASK_ACTION_EXEC = 0
action = newtask.Actions.Create(TASK_ACTION_EXEC)
action.ID = 'Trigger Python'
action.Path = r'C:\Python38\python.exe'
action.Arguments = r'C:\Users\Matt\Desktop\autoPython-main\OpenCitrix.py'

# Parameters
newtask.RegistrationInfo.Description = 'Open Citrix VACO using OpenCitrix.py'
newtask.Settings.Enabled = True
newtask.Settings.StopIfGoingOnBatteries = False

# Saving
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_NONE = 0
root_folder.RegisterTaskDefinition(
    'ML_Open Citrix VACO',  # Task name
    newtask,
    TASK_CREATE_OR_UPDATE,
    '',  # No user
    '',  # No password
    TASK_LOGON_NONE)
