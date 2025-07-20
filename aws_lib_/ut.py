



##f = open("aknnucox-pdf-page-1-text.txt", "r")
##text1=f.read()
##
##
##
##f = open("aknnucox-pdf-page-2-text.txt", "r")
##text2=f.read()
##text=text1+text2
##print(text)



import os
import datetime
def modification_date(filename):
    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)


print(modification_date(r"C:\Users\acer\Desktop\returns_12122019_R2A_27AAGCM2267G1ZF_R2A_others_0.json"))
