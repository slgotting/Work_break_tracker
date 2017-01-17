import winsound
import time
import datetime
from xlutils.copy import copy
from xlrd import open_workbook 
from xlwt import easyxf



def excel_write():
    
    rb = open_workbook('Work_Times_Jan_2017.xls', formatting_info=True)
    r_sheet = rb.sheet_by_index(activity)
    wb = copy(rb)
    w_sheet = wb.get_sheet(activity)


    k = 3
    l = 1
    m = 3
    n = 0
    o = 3
    p = 5

    
    now = datetime.datetime.now()
    while m<10000:
        try:
            value3 = r_sheet.cell(m, n).value
            if value3 !=0:
                m+=1
        except IndexError:
            w_sheet.write(m, n, now.strftime("%Y-%m-%d"))
            m+=10000
    
    while k<10000:
        try:
            value2 = r_sheet.cell(k, l).value
            if value2 !=0:
                k+=1
        except IndexError:
            w_sheet.write(k, l, now.strftime("%H:%M"))
            k+=10000





    if activity == 4:
        while o<10000:
            try:
                value = r_sheet.cell(o, p).value
                if value != 0:
                    o+=1
            except IndexError:
                w_sheet.write(o, p, (book))
                o+=10000

    if activity == 5:
        while o<10000:
            try:
                value = r_sheet.cell(o, p).value
                if value != 0:
                    o+=1
            except IndexError:
                w_sheet.write(o, p, (work))
                o+=10000

                
    

    wb.save("Work_Times_Jan_2017.xls");

Y = None

while True:
    print "Welcome to your work break timer!"
    print "Directions: Input values in units of minutes"
    active = input('Work segment time interval: ')
    downtime = input('Break time interval: ')
    print "0 for IEECS, 1 for Math for CS, 2 for Intro to Algorithms, 3 for Linear Algebra, 4 for General Reading, 5 for Other"
    activity = input('What type of work will you be doing? ')
    segments = input('How many work/break segments do you plan to work on this subject? ')
    if activity == 4:
        book = raw_input('What book are you reading? ')
    if activity == 5:
        work = raw_input('What type of work are you doing? ')
    
    excel_write()
    
        
    
    
    
        
    i=0
    while i<(segments):

        print "\n"
        print "Work clock starts now!"
        winsound.PlaySound('SystemQuestion', winsound.SND_ALIAS)
        time.sleep(active * 60)
        winsound.PlaySound('SystemHand', winsound.SND_ALIAS)

        print "\n"
        print "Take your break now"
        time.sleep(downtime * 60)
        
        i+=1
        if i == (segments):
            winsound.PlaySound('SystemAsterisk', winsound.SND_ALIAS)
            print "\n"
            print "Session Over!"
            time.sleep(1)
            grade = raw_input('Give yourself a grade on how well you focused during all work segments (0-100) : ')
            print (grade)
            print "\n"

            
            weighted = (float(active) * float(segments) * float(grade) / 100) 
            
            rb = open_workbook('Work_Times_Oct_2016.xls', formatting_info=True)
            r_sheet = rb.sheet_by_index(activity)
            wb = copy(rb)
            w_sheet = wb.get_sheet(activity)
            now = datetime.datetime.now()
            import xlrd

            c = 3
            d = 6
            while c<10000:
                try:
                    value = r_sheet.cell(c, d).value
                    if value != 0:
                        c+=1
                except IndexError:
                    w_sheet.write(c - 1, d, (weighted))
                    c+=10000

            e = 3
            f = 4
            while e<10000:
                try:
                    value = r_sheet.cell(e, f).value
                    if value != 0:
                        e+=1
                except IndexError:
                    w_sheet.write(e - 1, f, int(grade))
                    e+=10000

            a = 3
            b = 3
            while a<10000:
                try:
                    value = r_sheet.cell(a, b).value
                    if value != 0:
                        a+=1
                except IndexError:
                    w_sheet.write(a - 1, b, (active * segments))
                    a+=10000                
                       


            u = 3
            y = 2
            while u<10000:
                try:
                    value = r_sheet.cell(u, y).value
                    if value != 0:
                        u+=1
                except IndexError:
                    w_sheet.write(u - 1, y, now.strftime("%H:%M"))
                    u+=10000
            z = 0
            while z<1:
                try:
                    wb.save("Work_Times_Jan_2017.xls")
                    z+=1
                except IOError:
                    print 'Please close the excel sheet you are wishing to save to'
                    time.sleep(10)
                
            
    
            
            j=0
            while j<1000:
                ready = raw_input('Are you ready for your next session?(Y/N): ')
                if ready == 'Y':
                    break
                else:
                    continue
                j+=1


                     

    print "\n"

