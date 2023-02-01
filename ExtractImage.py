import os
import win32com.client

path =  os.getcwd().replace('\'','\\') + '\\'
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open(path + 'slides.pptx',WithWindow=False)

slide_count = len(Presentation.Slides)


thisdict = {
  1: "start",
  2: "bio",
  3: "path",
  4: "tree_1",
  5: "tree_2",
  6: "tree_3",
  7: "phd_title",
  8: "phd_1",
  9: "phd_2",
  10: "phd_3",
  11: "phd_4",
  12: "cal_title",
  13: "eta_1",
  14: "eta_2",
  15: "dynamicdoe_1",
  16: "airpath_1",
  17: "strategy_1",
  18: "exhaust_1",
  19: "neural_1",
  20: "lnt_1",
  21: "control_1",
  22: "control_2",
  23: "control_3",
  24: "mpt_1",
  25: "mpt_2",
  26: "mpt_3",
  27: "mpt_4",
  28: "mpt_5",
  29: "mpt_6",
  30: "dynamic_1",
  31: "dynamic_2",
  32: "dynamic_3",
  33: "drive_1",
  34: "drive_2",
  35: "trace_1",
  36: "trace_2",
  37: "trace_3",
  38: "trace_4",
  39: "trace_5",
  40: "surfacefit_1", 
  41: "pid_1",
  42: "construct_1",
  43: "malaysia_1",
  44: "ds_title",
  45: "monte_1",
  46: "anomaly_1",
  47: "aws_1",
  48: "aws_2",
  49: "cbms_1",
  50: "cbms_2",
  51: "cbms_3",
  52: "cbms_4",
  53: "model_1",
  54: "nox_1",
  55: "nox_2",
  56: "graph_1",
  57: "mdf_1",
  58: "mdf_2",
  59: "misc_title",
  60: "report_1", 
  61: "support_1",   
  62: "ga_1",
  63: "ga_2",
  64: "svm_1",
  65: "track_1",
  66: "mvc_1",
}




for i in range(slide_count):
    print(i)
    Presentation.Slides[i].Export(path + '\\images\\' + thisdict[i+1] + '.png', "PNG")

#Presentation.Slides[i].Export(path + '\\images\\image_' + str(i) + '.png', "PNG")

Application.Quit()
Presentation =  None
Application = None


