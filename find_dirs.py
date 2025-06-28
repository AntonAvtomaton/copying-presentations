import glob
import os
import win32com.client as win32

""" Определяем корневую директорию, 
    где будем создавать новую презентацию 
    и где находятся изначальные презентации.
    Далее создаем новую презентаци по полученному пути.
    Далее находим все презентации в папке и сохраняем их в список.
"""

main_dir = os.path.dirname(__file__)
dir_output_file = os.path.join(main_dir, "new_presentation.pptx")

Presentations = glob.glob(os.path.join(main_dir, "*.pptx")) + glob.glob(os.path.join(main_dir, "*.ppt"))

powerpoint = win32.gencache.EnsureDispatch("PowerPoint.Application")
powerpoint.Visible = True
New_pres = powerpoint.Presentations.Add()
New_pres.SaveAs(dir_output_file)
print('Презентация "new_presentation.pptx" создана.')

