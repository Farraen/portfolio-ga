import os
import win32com.client

path =  os.getcwd().replace('\'','\\') + '\\'
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open(path + 'slides.pptx',WithWindow=False)

for i in range(50):
    try:
        print(i)
        Presentation.Slides[i].Export(path + '\\images\\image_' + str(i) + '.png', "PNG")
    except:
        print('Skip ' + str(i))
        pass

Application.Quit()
Presentation =  None
Application = None


