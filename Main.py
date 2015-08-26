import win32com.client
import os

all_file = os.listdir('E:\\3d\\')

swApp = win32com.client.Dispatch('SldWorks.Application')
swApp.Visible = 1

for i in all_file:
    
    Model = swApp.OpenDoc('E:\\3d\\' + i, 1)
    Model_name = i.split('.')
    Model_name = Model_name[0] + '.' + 'igs'
    result = Model.SaveAs('E:\\barobot_CATIA\\' + Model_name)
    swApp.CloseAllDocuments(True)

swApp.ExitApp()
