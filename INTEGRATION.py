from tkinter import*
import random
import time
import win32com.client
import win32com.client.combrowse as com
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

a=0

objTKSolver = win32com.client.Dispatch("TKWX.Document")
objTKSolver.LoadModel("r", r"D:\INTERNSHIP_2\Assignment\Assignment_6 python link to tk solver\FINAL\corona rate final.tkwx")

objTKSolver.ShowWindow(3)

# Rules (Equations)
eq1 = objTKSolver.GetSheetCell("r", "1", "2")
eq2 = objTKSolver.GetSheetCell("r", "2", "2")
eq3 = objTKSolver.GetSheetCell("r", "3", "2")
eq4 = objTKSolver.GetSheetCell("r", "4", "2")
eq5 = objTKSolver.GetSheetCell("r", "5", "2")
eq6 = objTKSolver.GetSheetCell("r", "6", "2")

objTKSolver.SetValue("Recoveryrate", "s", "Output")
objTKSolver.SetValue("Dailyrate", "s", "Output")
objTKSolver.SetValue("Fatalityrate", "s", "Output")
objTKSolver.SetValue("Activecasesrate", "s", "Output")
objTKSolver.SetValue("Deathvsrecoveryrate", "s", "Output")
objTKSolver.SetValue("dailytestrate", "s", "Output")

root = Tk()
root.geometry("1600x700+0+0")
root.title("Coronavirus analysis tracker")

Tops = Frame(root,bg="white",width = 1600,height=50,relief=SUNKEN)
Tops.pack(side=TOP)

f1 = Frame(root,width = 900,height=700,relief=SUNKEN)
f1.pack(side=LEFT)

f2 = Frame(root ,width = 400,height=700,relief=SUNKEN)
f2.pack(side=RIGHT)
#------------------TIME--------------
localtime=time.asctime(time.localtime(time.time()))
#-----------------INFO TOP------------
lblinfo = Label(Tops, font=( 'aria' ,30, 'bold' ),text="Corona Tracker",fg="steel blue",bd=10,anchor='w')
lblinfo.grid(row=0,column=0)
lblinfo = Label(Tops, font=( 'aria' ,20, ),text=localtime,fg="steel blue",anchor=W)
lblinfo.grid(row=1,column=0)


def calculation():
    Totalconfirmed = float(Totalconfirmedcases.get())
    objTKSolver.SetValue("totalconfirmedcases", "i", Totalconfirmed)
    recovered = float(RecoveredPersons.get())
    objTKSolver.SetValue("recoveredpersons", "i", recovered)
    confirmedday = float(Confirmedcasesperday.get())
    objTKSolver.SetValue("confirmedcasesperday", "i", confirmedday)
    Totaldeaths = float(TotalDeaths.get())
    objTKSolver.SetValue("totalDeaths", "i", Totaldeaths)
    testdoneperday=float(Testperday.get())
    objTKSolver.SetValue("testdoneperday", "i", testdoneperday)
    totaltestdone=float(Totaltest.get())
    objTKSolver.SetValue("totaltestsdone", "i", totaltestdone)
    
    objTKSolver.Solve()
    return 1


def Ref():
    global a
    
    if a==0:
        calculation()
    
    Recoveryrate = objTKSolver.GetValue("Recoveryrate", "o")

    Dailyrate = objTKSolver.GetValue("Dailyrate", "o")

    Fatalityrate = objTKSolver.GetValue("Fatalityrate", "o")

    Activecasesrate = objTKSolver.GetValue("Activecasesrate", "o")
    
    Deathvsrecoveryrate = objTKSolver.GetValue("Deathvsrecoveryrate", "o")
    
    dailytestrate = objTKSolver.GetValue("dailytestrate", "o")


    Recovery_rate.set(Recoveryrate)
    Fatality_rate.set(Fatalityrate)
    Daily_rate.set(Dailyrate)
    Activecase_rate.set(Activecasesrate)
    Deathvsrecovery.set(Deathvsrecoveryrate)
    Dailytest_rate.set(dailytestrate)
    
def recoveryrate():
    global a
    if a==0:
        a = calculation()
    
    Recoveryrate = objTKSolver.GetValue("Recoveryrate", "o")
    Recovery_rate.set(Recoveryrate)
    
def active_case_rate():
    global a
    if a==0:
        a = calculation()
        
    Activecasesrate = objTKSolver.GetValue("Activecasesrate", "o")
    Activecase_rate.set(Activecasesrate)
    
def daily_rate():
    global a
    if a==0:
        a = calculation()
        
    Dailyrate = objTKSolver.GetValue("Dailyrate", "o")
    Daily_rate.set(Dailyrate)
        
def fatality_rate():
    global a
    if a==0:
        a = calculation()
    
    Fatalityrate = objTKSolver.GetValue("Fatalityrate", "o")
    Fatality_rate.set(Fatalityrate)
        
def daily_test():
    global a
    if a==0:
        a = calculation()
        
    dailytestrate = objTKSolver.GetValue("dailytestrate", "o")
    Dailytest_rate.set(dailytestrate)
        
def dvr():
    global a
    if a==0:
        a = calculation()
    
    Deathvsrecoveryrate = objTKSolver.GetValue("Deathvsrecoveryrate", "o")
    Deathvsrecovery.set(Deathvsrecoveryrate)
    
def plot1():
    Plist = []
    TCClist = [] 
    for i in range(1,11):
        val1 = round(float(objTKSolver.GetSubCell("l",5,i,1)))
        TCClist.append(val1)
        
        val2 = objTKSolver.GetSubCell("l",10,i,1)
        Plist.append(val2)
    
    TCClist = np.array(TCClist)
    barWidth = 0.30
    r1 = np.arange(len(TCClist))
    
    plt.bar(r1, TCClist, color="r", width=barWidth, edgecolor="white", label="Total Confirmed Cases")
    
    plt.xlabel("Period")
    plt.ylabel("Total Confirmed Cases")
    plt.xticks([r for r in range(len(TCClist))], Plist, rotation = 30)
    
    plt.legend()
    plt.show()
    
def plot2():
    TDPDlist = []
    TCClist = []
    
    for i in range(1,11):
        val1 = int(objTKSolver.GetSubCell("l",7,i,1))
        TDPDlist.append(val1)
        
        val2 = int(objTKSolver.GetSubCell("l",5,i,1))
        TCClist.append(val2)
        
    TDPDlist = np.array(TDPDlist)
    TCClist = np.array(TCClist)
    plt.plot(TDPDlist,TCClist)
    plt.xlabel("Test Done Per Day")
    plt.ylabel("Total Confirmed Cases")
    plt.show()
    
def plot3():
    FRlist = []
    DTRlist = []
    RRlist = []
    TCClist = [] 
    for i in range(1,11):
        val1 = round(float(objTKSolver.GetSubCell("l",2,i,1)), 2)
        FRlist.append(val1)
        
        val2 = round(float(objTKSolver.GetSubCell("l",3,i,1)),2)
        DTRlist.append(val2)
        
        val3 = round(float(objTKSolver.GetSubCell("l",1,i,1)),2)
        RRlist.append(val3)
        
        val4 = round(float(objTKSolver.GetSubCell("l",5,i,1)))
        TCClist.append(val4)
        
    FRlist = np.array(FRlist)
    DTRlist = np.array(DTRlist)
    RRlist = np.array(RRlist)
    TCClist = np.array(TCClist)
    
    barWidth = 0.30
    
    r1 = np.arange(len(FRlist))
    r2 = [x + barWidth for x in r1]
    r3 = [x + barWidth for x in r2]
    
    plt.bar(r1, FRlist, color="r", width=barWidth, edgecolor="white", label="Fatality Rate")
    plt.bar(r2, DTRlist, color="b", width=barWidth, edgecolor="white", label="Daily Test Rate")
    plt.bar(r3, RRlist, color="g", width=barWidth, edgecolor="white", label="Recovery Rate")
    
    plt.xlabel("Total Confirmed Cases")
    plt.ylabel("Fatality Rate, Daily Test Rate, Recovery Rate")
    plt.xticks([r + barWidth for r in range(len(FRlist))], TCClist)
    
    plt.legend()
    plt.show()
    
def qexit():
    root.destroy()

def reset():
    global a
    
    Totalconfirmedcases.set("")
    RecoveredPersons.set("")
    Confirmedcasesperday.set("")
    TotalDeaths.set("")
    Testperday.set("")
    Totaltest.set("")
    Recovery_rate.set("")
    Daily_rate.set("")
    Fatality_rate.set("")
    Activecase_rate.set("")
    Dailytest_rate.set("")
    Deathvsrecovery.set("")
    a=0  #extra






#---------------------------------------------------------------------------------------

Totalconfirmedcases= StringVar()
RecoveredPersons= StringVar()
Confirmedcasesperday = StringVar()
TotalDeaths= StringVar()
Testperday = StringVar()
Totaltest = StringVar()
Recovery_rate = StringVar()
Daily_rate = StringVar()
Fatality_rate = StringVar()
Activecase_rate = StringVar()
Dailytest_rate = StringVar()
Deathvsrecovery=StringVar()


lblTotalconfirmedcases = Label(f1, font=( 'aria' ,16, 'bold' ),text="Total Confirmed Cases",fg="steel blue",bd=10,anchor='w')
lblTotalconfirmedcases.grid(row=0,column=0)
txtTotalconfirmedcases = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Totalconfirmedcases , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtTotalconfirmedcases.grid(row=0,column=1)

lblRecoveredPersons = Label(f1, font=( 'aria' ,16, 'bold' ),text="Recovered Persons",fg="steel blue",bd=10,anchor='w')
lblRecoveredPersons.grid(row=1,column=0)
txtRecoveredPersons = Entry(f1,font=('ariel' ,16,'bold'), textvariable=RecoveredPersons , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtRecoveredPersons.grid(row=1,column=1)


lblConfirmedcasesperday = Label(f1, font=( 'aria' ,16, 'bold' ),text="Confirmed Cases/Day",fg="steel blue",bd=10,anchor='w')
lblConfirmedcasesperday.grid(row=2,column=0)
txtConfirmedcasesperday = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Confirmedcasesperday , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtConfirmedcasesperday.grid(row=2,column=1)

lblTotalDeaths = Label(f1, font=( 'aria' ,16, 'bold' ),text="Total Deaths",fg="steel blue",bd=10,anchor='w')
lblTotalDeaths.grid(row=3,column=0)
txtTotalDeaths = Entry(f1,font=('ariel' ,16,'bold'), textvariable=TotalDeaths , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtTotalDeaths.grid(row=3,column=1)

lblTestperday = Label(f1, font=( 'aria' ,16, 'bold' ),text="Test/Day",fg="steel blue",bd=10,anchor='w')
lblTestperday.grid(row=4,column=0)
txtTestperday = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Testperday , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtTestperday.grid(row=4,column=1)

#--------------------------------------------------------------------------------------
lblTotaltest  = Label(f1, font=( 'aria' ,16, 'bold' ),text="Total Test",fg="steel blue",bd=10,anchor='w')
lblTotaltest .grid(row=5,column=0)
txtTotaltest  = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Totaltest , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtTotaltest .grid(row=5,column=1)

lblrecovery = Label(f1, font=( 'aria' ,16, 'bold' ),text="Recovery Rate",fg="steel blue",bd=10,anchor='w')
lblrecovery.grid(row=0,column=3)
txtrecovery = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Recovery_rate , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtrecovery.grid(row=0,column=4)

lblactivecase = Label(f1, font=( 'aria' ,16, 'bold' ),text="Active Case Rate",fg="steel blue",bd=10,anchor='w')
lblactivecase.grid(row=1,column=3)
txtactivecase = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Activecase_rate , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtactivecase.grid(row=1,column=4)

lbldailyrate = Label(f1, font=( 'aria' ,16, 'bold' ),text="Daily Rate",fg="steel blue",bd=10,anchor='w')
lbldailyrate.grid(row=2,column=3)
txtdailyrate = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Daily_rate , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtdailyrate.grid(row=2,column=4)

lblfatalityrate = Label(f1, font=( 'aria' ,16, 'bold' ),text="Fatality Rate",fg="steel blue",bd=10,anchor='w')
lblfatalityrate.grid(row=3,column=3)
txtfatalityrate = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Fatality_rate , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtfatalityrate.grid(row=3,column=4)

lbldailytest = Label(f1, font=( 'aria' ,16, 'bold' ),text="Daily Test",fg="steel blue",bd=10,anchor='w')
lbldailytest.grid(row=4,column=3)
txtdailytest = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Dailytest_rate , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtdailytest.grid(row=4,column=4)

lbldeathvsrecovery = Label(f1, font=( 'aria' ,16, 'bold' ),text="Death Vs Recovery Rate",fg="steel blue",bd=10,anchor='w')
lbldeathvsrecovery.grid(row=5,column=3)
txtdeathvsrecovery = Entry(f1,font=('ariel' ,16,'bold'), textvariable=Deathvsrecovery , bd=6,insertwidth=4,bg="powder blue" ,justify='right')
txtdeathvsrecovery.grid(row=5,column=4)

#-----------------------------------------buttons------------------------------------------
lblTotal = Label(f1,text="---------------------",fg="white")
lblTotal.grid(row=6,columnspan=3)

#calculate()

btnTotal=Button(f1,padx=16,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="TOTAL", bg="powder blue",command=Ref)
btnTotal.grid(row=7, column=4)

btnreset=Button(f1,padx=16,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="RESET", bg="powder blue",command=reset)
btnreset.grid(row=8, column=4)

btnexit=Button(f1,padx=16,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="EXIT", bg="powder blue",command=qexit)
btnexit.grid(row=9, column=4)

btnRR=Button(f1,padx=16,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="Recover Rate", bg="powder blue",command=recoveryrate)
btnRR.grid(row=7, column=0)

btnACR=Button(f1,padx=20,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="Active Case Rate", bg="powder blue",command=active_case_rate)
btnACR.grid(row=7, column=1)

btnDR=Button(f1,padx=30,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="Daily Rate", bg="powder blue",command=daily_rate)
btnDR.grid(row=7, column=2)

btnFR=Button(f1,padx=16,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="Fatality Rate", bg="powder blue",command=fatality_rate)
btnFR.grid(row=8, column=0)

btnDT=Button(f1,padx=20,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="Daily Test", bg="powder blue",command=daily_test)
btnDT.grid(row=8, column=1)

btnDVR=Button(f1,padx=30,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="Death Vs Recovery", bg="powder blue",command=dvr)
btnDVR.grid(row=8, column=2)

btnP1=Button(f1,padx=16,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="PLOT 1", bg="powder blue",command=plot1)
btnP1.grid(row=9, column=0)

btnP2=Button(f1,padx=20,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="PLOT 2", bg="powder blue",command=plot2)
btnP2.grid(row=9, column=1)

btnP3=Button(f1,padx=30,pady=8, bd=10 ,fg="black",font=('ariel' ,16,'bold'),width=10, text="PLOT 3", bg="powder blue",command=plot3)
btnP3.grid(row=9, column=2)


root.mainloop()
objTKSolver.HideWindow()
com.main()
