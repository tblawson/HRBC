# -*- coding: utf-8 -*-
"""
R_info.py

Created on Fri Sep 18 09:40:33 2015

@author: t.lawson
"""
import string
import datetime as dt
import time
import math
import xlrd
from openpyxl.styles import Font,colors,PatternFill,Border,Side
import GTC
from numbers import Number

RL_SEARCH_LIMIT = 500

INF = 1e6 # 'inf' dof
ZERO = GTC.ureal(0,0)

# ______________________________Useful funtions:_____________________________________

def Make_Log_Name(v):
    return 'HRBAv'+str(v)+'_'+str(dt.date.today())+'.log'


# Extract resistor names from comment
def ExtractNames(comment):
    assert comment.find('R1: ') >= 0,'R1 name not found in comment!'
    assert comment.find('R2: ') >= 0,'R2 name not found in comment!'
    R1_name = comment[comment.find('R1: ') + 4:comment.find(' monitored by GMH')]
    R2_name = comment[comment.find('R2: ') + 4:comment.rfind(' monitored by GMH')]
    return (R1_name,R2_name)


# Extract nominal resistor value from name
def GetRval(name):
    prefixes = {'r':1,'R':1,'k':1000,'M':1e6,'G':1e9}
    
    if prefixes.has_key(name[-1]):
        mult = prefixes[name[-1]]
    else:
        mult = 0
    assert mult != 0,'Error parsing comment - unkown multiplier!'
        
    # return numeric part of last word, multiplied by 1, 10^3, 10^6 or 10^9:
    return(mult*int(string.strip(string.split(name)[-1],string.letters)))

def GetRLstartrow(sheet,Id,jump,log):
    search_row = 1
    while search_row < RL_SEARCH_LIMIT: # Don't search for ever.
        result = sheet['A'+str(search_row)].value # scan down column A
        if result == 'Run Id:':
            RL_Id = sheet['B'+str(search_row)].value # Find a run Id
            if RL_Id == Id: # Found the right data, so we're done.
                RL_start = search_row +1 # 1st line of Rlink data's header (excl Id)
                return RL_start
            else: # Jump to just before start of next data-block
                search_row += jump
        search_row +=1
    print 'No matching Rlink data!'
    log.write('GetRLstartrow():No matching Rlink data!\n')
    return -1


# Convert list of data to ureal, where possible
def Uncertainize(row_items):
    v = row_items[2]
    u = row_items[3]
    d = row_items[4]
    l = row_items[5]
    if (u is not None) and isinstance(v, Number):
        if d == u'inf':
            un_num = GTC.ureal(v,u,label=l) # default dof = inf
        else:
#            print v,u,d,l
            un_num = GTC.ureal(v,u,d,l)
        return un_num
    else: # non-numeric value
#        print 'Uncertainize(): Value is NON-NUMERIC!',row_items
        return v
    
    
# Convert a resistive T-sensor reading from resistance to temperature
def R_to_T(alpha,beta,R,R0,T0):
    if beta == 0: # no 2nd-order T-Co
        T = (R/R0 -1)/alpha + T0
    else:
        a = beta
        b = alpha-2*T0
        c = 1-alpha*T0 + beta*T0**2 - (R/R0)
        T = (-b + math.sqrt(b**2-4*a*c))/(2*a)
    return T


# Return average of a list of time-strings("%d/%m/%Y %H:%M:%S") as a time string or float
def av_t_strin(t_list,switch):
    assert switch in ('fl','str'),'Unknown switch for function av_t_strin()!'
    throwaway = dt.datetime.strptime('20110101','%Y%m%d') # known bug fix
    n = float(len(t_list))
    t_av = 0.0
    for s in t_list:
        if type(s) is unicode:
            t_dt = dt.datetime.strptime(s,'%d/%m/%Y %H:%M:%S')
        elif type(s) is float:
            print s,'is a float...'
            t_dt = xlrd.xldate.xldate_as_datetime(s,0)
        else:
            assert 0,'Time format is not unicode or float!'
        t_tup  = dt.datetime.timetuple(t_dt)
        t_av += time.mktime(t_tup)
        
    t_av /= n # av. time as float (seconds from epoch)
    if switch == 'fl':
        return t_av 
    elif switch == 'str':
        t_av_fl = dt.datetime.fromtimestamp(t_av)
        return t_av_fl.strftime('%d/%m/%Y %H:%M:%S') # av. time as string


# Write headings on Summary sheet
def WriteHeadings(sheet,row,version):
    now = dt.datetime.now()
    sheet['A'+str(row-1)].font = Font(b=True)
    sheet['A'+str(row-1)] = 'Processed with HRBA v'+str(version)+' on '+now.strftime("%A, %d. %B %Y %I:%M%p")
    sheet['J'+str(row)] = 'Uncertainty Budget'
    
    sheet['R'+str(row)] = 'R1(T)'
    sheet['U'+str(row)] = 'exp. U(95%)'
    sheet['V'+str(row)] = 'av T'
    sheet['Y'+str(row)] = 'av date/time'
    sheet['Z'+str(row)] = 'av V'
    row += 1
    sheet['A'+str(row)] = 'Name'
    sheet['B'+str(row)] = 'Test V'
    sheet['C'+str(row)] = 'Date'
    sheet['D'+str(row)] = 'T'
    sheet['E'+str(row)] = 'R1'
    sheet['F'+str(row)] = 'std u.'
    sheet['G'+str(row)] = 'dof'
    sheet['H'+str(row)] = 'exp. U(95%)'
    sheet['J'+str(row)] = 'Quantity (Label)'
    sheet['K'+str(row)] = 'Value'
    sheet['L'+str(row)] = 'Std u'
    sheet['M'+str(row)] = 'dof'
    sheet['N'+str(row)] = 'sens. coef.'
    sheet['O'+str(row)] = 'u contrib.'
    
    sheet['Q'+str(row)] = 'LV'
    row += 1
    sheet['Q'+str(row)] = 'HV'
    
    return row


# Write measurement summary   
def WriteThisResult(sheet,row,result):
    pass
    sheet['A'+str(row)].font = Font(color=colors.YELLOW)
    sheet['A'+str(row)].fill = PatternFill(patternType='solid', fgColor=colors.RED)
    sheet['A'+str(row)] = str(result['name'])
    sheet['B'+str(row)] = result['V'].x
    sheet['C'+str(row)] = str(result['time_str'])
    sheet['D'+str(row)] = result['T'].x
    sheet['E'+str(row)] = result['R'].x
    sheet['F'+str(row)] = result['R'].u
    if math.isinf(result['R'].df):
        print'WriteThisResult(): result.dof is',result['R'].df
        sheet['G'+str(row)] = str(result['R'].df)
    else:
        print'WriteThisResult(): result.dof =',result['R'].df
        sheet['G'+str(row)] = round(result['R'].df)
    # Exp Uncert:
    sheet['H'+str(row)] = result['R_expU']
  
  
# Sorting helper function - sort by uncert. contribution
def by_u_cont(line):
    return line[5]    
 
 
def WriteBudget(sheet,row,budget):
    for line in budget:
        sheet['J'+str(row)] = line[0] # Quantity (label)
        sheet['K'+str(row)] = line[1] # Value
        sheet['L'+str(row)] = line[2] # Uncert.
        if math.isinf(line[3]):
            print'WriteBudget(): dof (',line[0],') is',line[3]
            sheet['M'+str(row)] = str(line[3]) # dof
        else:
            print'WriteBudget(): dof (',line[0],') =',line[3]
            sheet['M'+str(row)] = round(line[3]) # dof
        sheet['N'+str(row)] = line[4] # Sens. coef.
        sheet['O'+str(row)] = line[5] # Uncert. contrib.
        row += 1
    return row


# Weighted least-squares fit (R1-T)
def write_R1_T_fit(results,sheet,row,log):
    T_data = [T for T in [result['T'] for result in results]] # All T values
    T_av = GTC.fn.mean(T_data)
    T_rel = [t_k - T_av for t_k in T_data] # x-vals
    
    y = [R for R in [result['R'] for result in results]] # All R values
    #u_y = [R.u for R in [result['R'] for result in results]] # All R uncerts

    if len(set(T_data)) <= 1: # No temperature variation recorded, so can't fit to T
        R1  = GTC.fn.mean(y)
        print 'R1_LV (av, not fit):',R1
        log.write('\nR1_LV (av, not fit): ' + str(R1))
    else:
        #a_ta,b_ta = GTC.ta.line_fit_wls(T_rel,y,u_y).a_b
        # Assume uncert of individual measurements dominate uncert of fit
        R1,alpha = GTC.fn.line_fit_wls(T_rel,y).a_b
        print 'Fit params:\t intercept=',GTC.summary(R1),'Slope=',GTC.summary(alpha)
        log.write('\nFit params:\t intercept= ' + str(GTC.summary(R1)) + ' Slope= ' + str(GTC.summary(alpha)))
        
    sheet['R'+str(row)] = R1.x
    sheet['S'+str(row)] = R1.u
    if math.isinf(R1.df):
        print'write_R1_T_fit(): R1.df is',R1.df
        sheet['T'+str(row)] = str(R1.df)
    else:
        print'write_R1_T_fit(): R1.df =',R1.df
        sheet['T'+str(row)] = round(R1.df)
    
    sheet['U'+str(row)] = R1.u*GTC.rp.k_factor(R1.df)
    
    sheet['V'+str(row)] = T_av.x
    sheet['W'+str(row)] = T_av.u
    if math.isinf(T_av.df):
        print'write_R1_T_fit(): T_av.df is',T_av.df
        sheet['X'+str(row)] = str(T_av.df)
    else:
        print'write_R1_T_fit(): T_av.df =',T_av.df
        sheet['X'+str(row)] = round(T_av.df)
    
    t = [result['time_fl'] for result in results] # x data (time,s from epoch)
    t_av = GTC.ta.estimate(t)
    time_av = dt.datetime.fromtimestamp(t_av.x) # A Python datetime object
    sheet['Y'+str(row)] = time_av.strftime('%d/%m/%Y %H:%M:%S')# string-formatted for display
    
    V1 = [V for V in [result['V'] for result in results]]
    V_av = GTC.fn.mean(V1)
    sheet['Z'+str(row)] = V_av.x
    sheet['AA'+str(row)] = V_av.u
    if math.isinf(V_av.df):
        print'write_R1_T_fit(): V_av.df is',V_av.df
        sheet['AB'+str(row)] = str(V_av.df)
    else:
        print'write_R1_T_fit(): V_av.df =',V_av.df
        sheet['AB'+str(row)] = round(V_av.df)
    return (R1,alpha,T_av,V_av,time_av)


def update_R_Info(name,params,data,sheet,row,Id,v):
    R_dict = dict(zip(params,data))
    
    for param in params:
        label = name.split()[0] + '_'+ param + '_' + Id
        row += 1
        sheet['A'+str(row)] = name
        sheet['B'+str(row)] = param
        if param not in ('date','T_sensor'): # GTC.ureal params
            sheet['C'+str(row)] = R_dict[param].x
            sheet['D'+str(row)] = R_dict[param].u
            if math.isinf(R_dict[param].df):
                print'update_R_Info(): ',name,'(',param,').dof is',R_dict[param].df
                sheet['E'+str(row)] = str(R_dict[param].df)
            else:
                print'update_R_Info(): ',name,'(',param,').dof =',R_dict[param].df
                sheet['E'+str(row)] = round(R_dict[param].df)
            sheet['F'+str(row)] = label
        else:
            sheet['C'+str(row)] = R_dict[param]
        
        now_tup = dt.datetime.now()
        now_fmt = now_tup.strftime('%d/%m/%Y %H:%M:%S')
        sheet['G'+str(row)] = 'HRBA'+ str(v) + '_'+ now_fmt
    
    # Mark end of data with a bottom border on cells of last row:
    b = Border(bottom = Side(style='thin'))
    sheet['A'+str(row)].border = b
    sheet['B'+str(row)].border = b
    sheet['C'+str(row)].border = b
    sheet['D'+str(row)].border = b
    sheet['E'+str(row)].border = b
    sheet['F'+str(row)].border = b
    sheet['G'+str(row)].border = b    
    
    return row
