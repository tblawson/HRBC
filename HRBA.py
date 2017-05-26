# -*- coding: utf-8 -*-
"""
HRBA.py High Resistance Bridge Analysis (using imported GTC functions).

This script analyses data collected in an Excel file, generated by
the High Resistance Bridge Control (HRBC) application. Each 4-line
block of data in the 'Data' worksheet represents one measurement of
the unknown resistor R2('Rx') under a unique set of conditions
(time t, Temperature T, Voltage V).

Each HRBA run analyses one HRBC run, which normally consists of several
measurements at alternating low or high voltage (LV,HV). HRBA searches
the 'Rlink' worksheet for a block of Rlink data that has a matching run
ID and uses this information to calculate the link resistance Rd.

Where two temperature measurements of a resistor have been recorded
(eg with a GMH probe and a DVM monitoring a Pt sensor), the difference
is used to define a zero-valued type-B uncertainty that describes the
temperature definition. Where only a GMH reading is available the
temperature definition uncertainty defaults to 0 +/- 0.01 C with 4
degrees of freedom.

The multiple results from one analysis run are combined using least
squares fits to time, Temperature and Voltage, yielding six overall
measurements of R2 (at mean t, mean T, mean V for LV or HV conditions).

The fitting procedure also generates temperature and voltage coefficients
of resistance for the unknown resistor and an estimate of its drift rate.
NOTE: No correlations between these quantities are assumed.
All results are written to the 'Results' worksheet.

Created on Fri Sep 18 14:01:18 2015

@author: t.lawson
"""

import os
import sys
sys.path.append("C:\Python27\Lib\site-packages\GTC")

from openpyxl import load_workbook, cell
from openpyxl.cell import get_column_letter #column_index_from_string
import GTC

import R_info # useful functions
#import xlrd

VERSION = 1.0

# DVM, GMH Correction factors, etc.

#INF = 1e6 # 'inf' dof
ZERO = GTC.ureal(0,0)

# I:\MSL\Private\Electricity\Commercial\Working\IV Converters for Light\\2016 Reports
# I:\MSL\Private\Electricity\Staff\TBL\Python\High_Res_Bridge\Development\\test version\Validation
os.environ['XLPATH'] = 'I:\MSL\Private\Electricity\Commercial\Working\IV Converters for Light\\2016 Reports'
# xldir = os.environ['XLPATH']
xldir = raw_input('Path to data directory:')
# xlfile = 'HRBC_HRBA_for_IV-conv Laurie.xlsx' #  # new_High-Res_validation.xlsx
xlfile = raw_input('Excel filename:')
xlfilename = os.path.join(xldir, xlfile)

# open existing workbook
wb_io = load_workbook(xlfilename) # did have ', data_only=False'
ws_Data = wb_io.get_sheet_by_name('Data')
ws_Rlink = wb_io.get_sheet_by_name('Rlink')
ws_Summary = wb_io.get_sheet_by_name('Results')
ws_Params = wb_io.get_sheet_by_name('Parameters')

# Get local parameters
Data_start_row = ws_Data['B1'].value
Data_stop_row = ws_Data['B2'].value
assert Data_start_row <= Data_stop_row,'Stop row must follow start row!'

# Get instrument assignments
role_descr = {}
for row in range(Data_start_row, Data_start_row+10): # 10 roles in total
    # Grab {role:description}
    temp_dict = {ws_Data['AC'+str(row)].value : ws_Data['AD'+str(row)].value}
    assert temp_dict.keys()[0] is not None,'Instrument assignment: Missing role!'
    assert temp_dict.values()[0] is not None,'Instrument assignment: Missing description!'
    role_descr.update(temp_dict)

#######################################################################    
#______________Extract resistor and instrument parameters_____________#

print 'Reading parameters...'
headings = (u'Resistor Info:', u'Instrument Info:',
            u'description', u'parameter', u'value',
            u'uncert', u'dof', u'label', u'Comment / Reference')
        
 # Determine colummn indices from column letters:
col_A = cell.column_index_from_string('A') - 1
col_B = cell.column_index_from_string('B') - 1
col_C = cell.column_index_from_string('C') - 1
col_D = cell.column_index_from_string('D') - 1
col_E = cell.column_index_from_string('E') - 1
col_F = cell.column_index_from_string('F') - 1
col_G = cell.column_index_from_string('G') - 1
#col_H = cell.column_index_from_string('H') - 1
col_I = cell.column_index_from_string('I') - 1
col_J = cell.column_index_from_string('J') - 1
col_K = cell.column_index_from_string('K') - 1
col_L = cell.column_index_from_string('L') - 1
col_M = cell.column_index_from_string('M') - 1
col_N = cell.column_index_from_string('N') - 1
col_O = cell.column_index_from_string('O') - 1
        
R_params = []
R_row_items = []
I_params = []
I_row_items = []
R_values = []
I_values = []
R_DESCR = []
I_DESCR = []
R_sublist = []
I_sublist = []
        
for r in ws_Params.rows: # a tuple of row objects
    R_end = 0

    # description, parameter, value, uncert, dof, label:
    R_row_items = [r[col_A].value, r[col_B].value, r[col_C].value, r[col_D].value,
                   r[col_E].value, r[col_F].value, r[col_G].value]
    
    I_row_items = [r[col_I].value, r[col_J].value, r[col_K].value, r[col_L].value,
                   r[col_M].value, r[col_N].value, r[col_O].value]
    
    if R_row_items[0] == None: # end of R_list
        R_end = 1

    # check this row for heading text
    if any(i in I_row_items for i in headings): 
        continue # Skip headings
        
    else: # not header - main data
        # Get instrument parameters first...
        last_I_row = r[col_I].row
        I_params.append(I_row_items[1])
        I_values.append(R_info.Uncertainize(I_row_items))
        if I_row_items[1] == u'test': # last parameter for this description
            I_DESCR.append(I_row_items[0]) # build description list
            I_sublist.append(dict(zip(I_params,I_values))) # add parameter dictionary to sublist
            del I_params[:]
            del I_values[:]
                    
        # Now attend to resistor parameters...
        if R_end == 0: # Check we're not at the end of resistor data-block
            last_R_row = r[col_A].row # Need to know this if we write more data, post-analysis
            R_params.append(R_row_items[1])
            R_values.append(R_info.Uncertainize(R_row_items))
            if R_row_items[1] == u'T_sensor': # last parameter for this description
                R_DESCR.append(R_row_items[0]) # build description list
                R_sublist.append(dict(zip(R_params,R_values))) # add parameter dictionary to sublist
                del R_params[:]
                del R_values[:] 
                   
# Compile into dictionaries
I_INFO = dict(zip(I_DESCR,I_sublist))
print len(I_INFO),'instruments (%d rows)'%last_I_row

R_INFO = dict(zip(R_DESCR,R_sublist))
print len(R_INFO),'resistors.(%d rows)\n'%last_R_row

#--------------End of parameter extraction---------------#
##########################################################

# Determine the meanings of 'LV' and 'HV'
V1set_a = abs(ws_Data['A'+str(Data_start_row)].value)
assert V1set_a is not None,'Missing initial V1 value!'
V1set_b = abs(ws_Data['A'+str(Data_start_row+4)].value)

if V1set_a < V1set_b:
    LV = V1set_a
    HV = V1set_b 
elif V1set_b < V1set_a:
    LV = V1set_b
    HV = V1set_a
else: # 'HV' and 'LV' equal
    LV = HV = V1set_a

# Set up reading of Data sheet
Data_row = Data_start_row

# Get start_row on Summary sheet
summary_start_row = ws_Summary['B1'].value
assert summary_start_row is not None,'Missing start row on Results sheet!'

# Get run identifier and copy to Results sheet
Run_Id = ws_Data['B'+str(Data_start_row-1)].value
assert Run_Id is not None,'Missing Run Id!'

ws_Summary['C'+str(summary_start_row)] = 'Run Id:'
ws_Summary['D'+str(summary_start_row)] = str(Run_Id)

# Write headings
summary_row = R_info.WriteHeadings(ws_Summary,summary_start_row,VERSION)

cor_gmh1 = []
cor_gmh2 = []
T_dvm1 = []
T_dvm2 = []
R_dvm1 = []
R_dvm2 = []
times = []
RHs = []
Ps = []
Ts = []

# Lists of dictionaries (with name,time,R,T,V entries)
results_HV = [] # High voltage measurements
results_LV = [] # Low voltage measurements

# Get run comment and extract R names & R values
Data_comment = ws_Data['Z'+str(Data_row)].value
assert Data_comment is not None,'Missing Comment!'

R1_name,R2_name = R_info.ExtractNames(Data_comment)
R1val = R_info.GetRval(R1_name)
R2val = R_info.GetRval(R2_name)
print Data_comment
print 'Run Id:',Run_Id
    
# Check for knowledge of R2:
if not R_INFO.has_key(R2_name):
    sys.exit('ERROR - Unknown Rs: '+R2_name)

##############################
##___Loop over data rows ___##
print '\nLooping over data rows',Data_start_row,'to',Data_stop_row,'...'
while Data_row <= Data_stop_row:    
    
    # R2 parameters:
    V2set = abs(ws_Data['B'+str(Data_start_row)].value)
    assert V2set is not None,'Missing V2 setting!'
    V1set = abs(ws_Data['A'+str(Data_start_row)].value)
    assert V1set is not None,'Missing V1 setting!'
    Vdif_LV = abs(abs(V2set)-R_INFO[R2_name]['VRef_LV'])
    Vdif_HV = abs(abs(V2set)-R_INFO[R2_name]['VRef_HV'])
    if Vdif_LV < Vdif_HV:
        R2_0 = R_INFO[R2_name]['R0_LV']
        R2TRef = R_INFO[R2_name]['TRef_LV']
        R2VRef = R_INFO[R2_name]['VRef_LV']
    elif Vdif_LV > Vdif_HV:
        R2_0 = R_INFO[R2_name]['R0_HV']
        R2TRef = R_INFO[R2_name]['TRef_HV']
        R2VRef = R_INFO[R2_name]['VRef_HV']
    else:
        R2_0 = R_INFO[R2_name]['R0_LV']
        R2TRef = R_INFO[R2_name]['TRef_LV']
        R2VRef = R_INFO[R2_name]['VRef_LV']

    if int(round(V1set)) == int(round(V2set)):
        v_ratio_code = 'VRC_eq'
    elif int(round(V1set)) == 10 and int(round(V2set)) == 1:
        v_ratio_code = 'VRC_10to1'
    elif int(round(V1set)) == 100 and int(round(V2set)) == 10:
        v_ratio_code = 'VRC_100to10'
    else:
        v_ratio_code = None
    assert v_ratio_code is not None,'Unable to determine voltage ratio!'

    # Select appropriate value of VRC, etc.
    vrc = I_INFO[role_descr['DVM12']][v_ratio_code]
    Vlin_gain = I_INFO[role_descr['DVMd']]['linearity_gain'] # linearity used in G calculation
    Vlin_Vd = I_INFO[role_descr['DVMd']]['linearity_Vd'] # linearity used in Vd calculation
    
    # Start list of influence variables
    influencies = [vrc,Vlin_gain,Vlin_Vd,R2TRef,R2VRef] # R2 dependancies

    R2alpha = R_INFO[R2_name]['alpha']
    R2beta = R_INFO[R2_name]['beta']
    R2gamma = R_INFO[R2_name]['gamma']
    R2Tsensor  = R_INFO[R2_name]['T_sensor']
    influencies.extend([R2_0,R2alpha,R2beta,R2gamma]) # R2 dependancies
    
    if not R_INFO.has_key(R1_name):
        R1Tsensor = 'Pt 100r' # assume a Pt sensor in unknown resistor
    else:
        R1Tsensor = R_INFO[R1_name]['T_sensor'] 
    
    # GMH correction factors
    GMH1_cor = I_INFO[role_descr['GMH1']]['T_correction']
    GMH2_cor = I_INFO[role_descr['GMH2']]['T_correction']
    
    
    # Temperature measurement, RH and times:
    del cor_gmh1[:] # list for 4 corrected T1 gmh readings
    del cor_gmh2[:] # list for 4 corrected T2 gmh readings
    del T_dvm1[:] # list for 4 corrected T1(dvm) readings
    del T_dvm2[:] # list for 4 corrected T2(dvm) readings
    del R_dvm1[:] # list for 4 corrected dvm readings
    del R_dvm2[:] # list for 4 corrected dvm readings
    del times[:] # list for 3*4 mean measurement time-strings
    del RHs[:] # list for 4 RH values
    del Ps[:] # list for 4 room pressure values
    del Ts[:] # list for 4 room Temp values
    
    
    # Process times, RH and temperature data in this 4-row block:
    for r in range(Data_row,Data_row+4): # build list of 4 gmh / T-probe dvm readings
        assert ws_Data['U'+str(r)].value is not None,'No R1 GMH temperature data!'
        assert ws_Data['V'+str(r)].value is not None,'No R2 GMH temperature data!'
        cor_gmh1.append(ws_Data['U'+str(r)].value*(1+GMH1_cor))
        cor_gmh2.append(ws_Data['V'+str(r)].value*(1+GMH2_cor))
        
        assert ws_Data['G'+str(r)].value is not None,'No V2 timestamp!'
        assert ws_Data['M'+str(r)].value is not None,'No Vd1 timestamp!'
        assert ws_Data['P'+str(r)].value is not None,'No V1 timestamp!'
        times.append(ws_Data['G'+str(r)].value)
        times.append(ws_Data['M'+str(r)].value)
        times.append(ws_Data['P'+str(r)].value)
        
#        if ws_Data['Y'+str(r)].value is not None:
#        assert ws_Data['Y'+str(r)].value is not None,'No %RH data!'
#        RHs.append(ws_Data['Y'+str(r)].value)
#        else:
#            RHs.append(0)
        
        assert ws_Data['S'+str(r)].value is not None,'No R1 raw DVM (temperature) data!'
        raw_dvm1 = ws_Data['S'+str(r)].value
        
        assert ws_Data['T'+str(r)].value is not None,'No R2 raw DVM (temperature) data!'
        raw_dvm2 = ws_Data['T'+str(r)].value
        
        # Check corrections for range-dependant values...
        # and apply appropriate corrections
        assert raw_dvm1 > 0 and raw_dvm2 > 0 ,'Negative resistance value(s)!'
        if raw_dvm1 < 120:
            T1DVM_cor = I_INFO[role_descr['DVMT1']]['correction_100r']
        elif raw_dvm1 < 12e3:
            T1DVM_cor = I_INFO[role_descr['DVMT1']]['correction_10k']
        else:
            T1DVM_cor = I_INFO[role_descr['DVMT1']]['correction_100k']
        R_dvm1.append(raw_dvm1*(1+T1DVM_cor))
        
        if raw_dvm2 < 120:
            T2DVM_cor =  I_INFO[role_descr['DVMT2']]['correction_100r']
        elif raw_dvm2 < 12e3:
            T2DVM_cor =  I_INFO[role_descr['DVMT2']]['correction_10k']
        else:
            T2DVM_cor =  I_INFO[role_descr['DVMT2']]['correction_100k']
        R_dvm2.append(raw_dvm2*(1+T2DVM_cor))
    
    # Mean temperature from GMH
    # Data are plain numbers, so use ta.estimate() to return a ureal
    assert len(cor_gmh1) > 1,'Not enough GMH1 temperatures to average!'
    T1_av_gmh = GTC.ar.result(GTC.ta.estimate(cor_gmh1),label='T1_av_gmh '+ Run_Id)
    assert len(cor_gmh2) > 1,'Not enough GMH2 temperatures to average!'
    T2_av_gmh = GTC.ar.result(GTC.ta.estimate(cor_gmh2),label='T2_av_gmh '+ Run_Id)
    
    assert len(times) > 1,'Not enough timestamps to average!'
    times_av_str = R_info.av_t_strin(times,'str') # mean time(as a time string)
    times_av_fl = R_info.av_t_strin(times,'fl') # mean time(as a float)
    
#    assert len(RHs) > 1,'Not enough RH values to average!'
#    RH_av = GTC.ar.result(GTC.ta.estimate(RHs),label = 'RH_av')
    
    # Build lists of 4 temperatures (calculated from T-probe dvm readings)..
    # .. and calculate mean temperatures
    if (R1Tsensor in ('none','any')): # no or unknown T-sensor (Tinsleys or T-sensor itelf)
        T_dvm1 = [ZERO,ZERO,ZERO,ZERO]
    else:
        assert len(R_dvm1) > 1,'Not enough R_dvm1 values to average!'
        for R in R_dvm1: # convert resistance measurement to a temperature
            T_dvm1.append(R_info.R_to_T(R_INFO[R1Tsensor]['alpha'],
                                        R_INFO[R1Tsensor]['beta'],R,
                                        R_INFO[R1Tsensor]['R0_LV'],
                                        R_INFO[R1Tsensor]['TRef_LV']))
    if R2Tsensor in ('none','any'):
        T_dvm2 = [ZERO,ZERO,ZERO,ZERO]
    else:
        assert len(R_dvm2) > 1,'Not enough R_dvm2 values to average!'
        for R in R_dvm2: # convert resistance measurement to a temperature
            T_dvm2.append(R_info.R_to_T(R_INFO[R2Tsensor]['alpha'],
                                        R_INFO[R2Tsensor]['beta'],R,
                                        R_INFO[R2Tsensor]['R0_LV'],
                                        R_INFO[R2Tsensor]['TRef_LV']))
                                        
    # Mean temperature from T-probe dvm  
    # Data are plain numbers, so use ta.estimate() to return a ureal                                 
    T1_av_dvm = GTC.ar.result(GTC.ta.estimate(T_dvm1))
    T2_av_dvm = GTC.ar.result(GTC.ta.estimate(T_dvm2),label='T2_av_dvm'+ Run_Id)
    
    # Mean temperatures and temperature definitions
#    if role_descr['DVMT1']=='none':  # No aux. T sensor or DVM not associated with R1 (just GMH)
    T1_av = T1_av_gmh
    T1_av_dvm = GTC.ureal(0,0) # ignore any dvm data
    Diff_T1 = GTC.ureal(0,0) # No temperature disparity (GMH only)
#    else:
#        T1_av = GTC.ar.result(GTC.fn.mean((T1_av_dvm,T1_av_gmh)),label='T1_av'+ Run_Id)
#        Diff_T1 = GTC.magnitude(T1_av_dvm-T1_av_gmh)
    
#    if role_descr['DVMT2']=='none':  # No aux. T sensor or DVM not associated with R2 (just GMH)
    T2_av = T2_av_gmh
    T2_av_dvm = GTC.ureal(0,0) # ignore any dvm data
    Diff_T2 = GTC.ureal(0,0) # No temperature disparity (GMH only)
    influencies.append(T2_av_gmh) # R2 dependancy
#    else:
#        T2_av = GTC.ar.result( GTC.fn.mean((T2_av_dvm,T2_av_gmh)),label='T2_av' + Run_Id)
#        Diff_T2 = GTC.ar.result(GTC.magnitude(T2_av_dvm-T2_av_gmh),label='Diff_T2' + Run_Id)
#        influencies.append(T2_av_dvm,T2_av_gmh) # R2 dependancy
    
    # Default T definition arises from imperfect positioning of sensors wrt resistor:
    T_def = GTC.ureal(0,GTC.type_b.distribution['gaussian'](0.01),3,label='T_def '+ Run_Id)
        
    # T-definition arises from imperfect positioning of both probes AND their disagreement:
    T_def1 = GTC.ar.result(GTC.ureal(0,Diff_T1.u/2,7) + T_def,label='T_def1 ' + Run_Id)    
    T_def2 = GTC.ar.result(GTC.ureal(0,Diff_T2.u/2,7) + T_def,label = 'T_def2 ' + Run_Id)
    influencies.append(T_def2) # R2 dependancy
    
    # Raw voltage measurements: V: [Vp,Vm,Vpp,Vppp]
    V1 = []
    V2 = []
    Vd = []
    for line in range(4):
        
        V1.append(GTC.ureal(ws_Data['Q'+str(Data_row+line)].value,
                        ws_Data['R'+str(Data_row+line)].value,
                        ws_Data['C'+str(Data_row+line)].value-1,label='V1_'+str(line) + ' ' + Run_Id))
        V2.append(GTC.ureal(ws_Data['H'+str(Data_row+line)].value,
                        ws_Data['I'+str(Data_row+line)].value,
                        ws_Data['C'+str(Data_row+line)].value-1,label='V2_'+str(line) + ' ' + Run_Id))
        Vd.append(GTC.ureal(ws_Data['N'+str(Data_row+line)].value,
                        ws_Data['O'+str(Data_row+line)].value,
                        ws_Data['C'+str(Data_row+line)].value-1,label='Vd_'+str(line) + ' ' + Run_Id))
        assert V1[-1] is not None,'Missing V1 data!'
        assert V2[-1] is not None,'Missing V2 data!'
        assert Vd[-1] is not None,'Missing Vd data!'
    influencies.extend(V1+V2+Vd) # R2 dependancies - raw measurements

    # Define drift
    Vdrift1=GTC.ureal(0,
    GTC.tb.distribution['gaussian'](abs(Vd[2]-(Vd[0]+((Vd[3]-Vd[2])/(V2[3]-V2[2]))*(V2[2]-V2[0])))/4),
                                8,label='Vdrift_gain '+ Run_Id)
    Vdrift2=GTC.ureal(0,
    GTC.tb.distribution['gaussian'](abs(Vd[2]-(Vd[0]+((Vd[3]-Vd[2])/(V2[3]-V2[2]))*(V2[2]-V2[0])))/4),
                                8,label='Vdrift_Vd '+ Run_Id)
    Vdrift = {'gain':Vdrift1,'Vd':Vdrift2}
    influencies.extend([Vdrift['gain'],Vdrift['Vd']]) # R2 dependancies
    
    # Mean voltages
    V1av = (V1[0]-2*V1[1]+V1[2])/4
    V2av = (V2[0]-2*V2[1]+V2[2])/4
    Vdav = (Vd[0]-2*Vd[1]+Vd[2])/4 + Vlin_Vd + Vdrift['Vd']    
    
    # __________Get Rd value__________
    # 1st, detetermine data format
    N_revs = ws_Rlink['B2'].value # Number of reversals = number of columns
    assert N_revs is not None and N_revs > 0,'Missing or no reversals!'
    N_reads = ws_Rlink['B3'].value # Number of readings = number of rows
    assert N_reads is not None and N_reads > 0,'Missing or no reads!'
    head_height = 6 # Rows of header before each block of data
    jump = head_height + N_reads # rows to jump between starts of each header
    
    # Find correct RLink data-header
    RL_start_row = R_info.GetRLstartrow(ws_Rlink,Run_Id,jump)
    
    # Next, define nom_R,abs_V quantities
    val1 = ws_Rlink['C'+str(RL_start_row+2)].value
    assert val1 is not None,'Missing nominal R1 value!'
    nom_R1 = GTC.constant(val1,label='nom_R1') # don't know uncertainty of nominal values
    val2 = ws_Rlink['C'+str(RL_start_row+3)].value
    assert val2 is not None,'Missing nominal R2 value!'
    nom_R2 = GTC.constant(val2,label='nom_R2') # don't know uncertainty of nominal values
    val1 =ws_Rlink['D'+str(RL_start_row+2)].value
    assert val1 is not None,'Missing nominal V1 value!'
    abs_V1 = GTC.constant(val1,label='abs_V1') # don't know uncertainty of nominal values
    val2 = ws_Rlink['D'+str(RL_start_row+3)].value
    assert val2 is not None,'Missing nominal V2 value!'
    abs_V2 = GTC.constant(val2,label='abs_V2') # don't know uncertainty of nominal values
     
    # Calculate I
    I=(abs_V1+abs_V2)/(nom_R1+nom_R2)
    I.label = 'Rd_I' + Run_Id
    
    # Average all +Vs and -Vs
    Vp = []
    Vn = []
    
    for Vrow in range(RL_start_row+5,RL_start_row+5+N_reads):
        col = 1
        while col <= N_revs: # cycle through cols 1 to N_revs
            Vp.append(ws_Rlink[get_column_letter(col)+str(Vrow)].value)
            assert Vp[-1] is not None,'Missing Vp value!'
            col +=1
            
            Vn.append(ws_Rlink[get_column_letter(col)+str(Vrow)].value)
            assert Vn[-1] is not None,'Missing Vn value!'
            col +=1
            
    av_dV_p = GTC.ta.estimate(Vp)
    av_dV_p.label='av_dV_p' + Run_Id
    av_dV_n = GTC.ta.estimate(Vn)
    av_dV_n.label='av_dV_n' + Run_Id
    av_dV = 0.5*(av_dV_p - av_dV_n)
    av_dV.label = 'Rd_dV' + Run_Id
    
    # Finally, calculate Rd
    Rd = GTC.ar.result(av_dV/I,label = 'Rlink ' + Run_Id)
    assert Rd.x < 0.01,'High link resistance!'
    assert Rd.x > Rd.u,'Link resistance uncertainty > value!'
    influencies.append(Rd) # R2 dependancy
    
    # Calculate R2 (corrected for T and V)
    dT2 = T2_av - R2TRef + T_def2
    
    dV2 = abs(abs(V2av) - R2VRef) # NOTE: TWO abs() NEEDED TO ENSURE NON-NEGATIVE DIFFERENCE!

    R2 = R2_0*(1+R2alpha*dT2 + R2beta*dT2**2 + R2gamma*dV2) + Rd
    assert abs(R2.x-nom_R2)/nom_R2 < 1e-4,'R2 > 100 ppm from nominal! R2 = {0}'.format(R2.x)
    
    # Gain factor due to null meter input Z
    G = (Vd[3]-Vd[2] + Vlin_gain +Vdrift['gain'])/(V2[3]-V2[2])
    if abs_V1/abs_V2 == 10:
        nom_G = 0.91
    elif abs_V1/abs_V2 == 1:
        nom_G = 0.5
    else:
        assert False,'Wrong V1/V2 ratio!'
    assert abs(G.x-nom_G)/nom_G < 0.02,'Gain > 2% from nominal! G = {0}, nom_G = {1}'.format(G.x,nom_G)
       
    # calculate R1  
    R1 = -R2*(1+vrc)*V1av*G/(G*V2av - Vdav)
    assert abs(R1.x-nom_R1)/nom_R1 < 2e-4,'R1 > 200 ppm from nominal!'
   
    # Combine data for this measurement: name,time,R,T,V and write to Summary sheet
    this_result = {'name':R1_name,'time_str':times_av_str,'time_fl':times_av_fl,'V':V1av,
                   'R':R1,'T':T1_av,'R_expU':R1.u*GTC.rp.k_factor(R1.df, quick=False)}
                   
    R_info.WriteThisResult(ws_Summary,summary_row,this_result)
    
    # build uncertainty budget table
    budget_table =[]
    for i in influencies: # rp.u_component(R1_gmh,i) gives + or - values
        if i.u > 0:
            sensitivity = GTC.rp.u_component(R1,i)/i.u # GTC.rp.sensitivity() deprecated
        else:
            sensitivity = 0
#        if GTC.component(R1,i) > 0:
        budget_table.append([i.label,i.x,i.u,i.df,sensitivity,GTC.component(R1,i)])
        
    budget_table_sorted = sorted(budget_table,key=R_info.by_u_cont,reverse=True)
    
    # write budget to Summary sheet
    summary_row = R_info.WriteBudget(ws_Summary,summary_row,budget_table_sorted)
    summary_row += 1 # Add a blank line between each measurement for ease of reading
    
    # Separate results by voltage (V1av) if different
    if HV == LV:
        results_LV.append(this_result)
        results_HV.append(this_result)
    elif abs(V1av.x - LV) < 1:
        results_LV.append(this_result)
    else:
        results_HV.append(this_result)
    
    del influencies[:]
    Data_row += 4 # Move to next measurement
   
##----- End of data-row loop -----#
###################################

# At this point the summary row has reached its maximum for this analysis run
# ...so make a note of it, for use as the next run's starting row:
ws_Summary['B1'] = summary_row

# Go back to the top of summary block, ready for writing run results
summary_row = summary_start_row + 1

########################################################################
# In the next section values of R1 are derived from fits to Temperature.
# The Temperature data are offset so the mean is at ~zero, then the fits
# are used to calculate R1 at the mean Temperature. LV and HV values are
# obtained separately. The mean time, Temperature and Voltage values are
# also reported.    

# Weighted total least-squares fit (R1-T), LV
print '\nLV:'
R1_LV, Ohm_per_C_LV, T_LV, V_LV, date = R_info.write_R1_T_fit(results_LV,ws_Summary,summary_row)
alpha_LV = Ohm_per_C_LV/R1_LV

summary_row += 1

print '\nHV:'
# Weighted total least-squares fit (R1-T), HV
R1_HV, Ohm_per_C_HV, T_HV, V_HV, date = R_info.write_R1_T_fit(results_HV,ws_Summary,summary_row)
alpha_HV = Ohm_per_C_HV/R1_HV

alpha = GTC.fn.mean([alpha_LV,alpha_HV])
beta = GTC.ureal(0,0) # assume no beta

if HV == LV: # Can't estimate gamma
    gamma = GTC.ureal(0,0)
else:
    gamma = ((R1_HV-R1_LV)/(V_HV-V_LV))/R1_LV

summary_row += 2

ws_Summary['R'+str(summary_row)] = 'alpha (/C)'
ws_Summary['T'+str(summary_row)] = 'gamma (/V)'
summary_row += 1
ws_Summary['R'+str(summary_row)] = GTC.summary(alpha)
ws_Summary['T'+str(summary_row)] = GTC.summary(gamma)

#######################################################################
# Finally, if R1 is a resistor that is not included in the 'parameters'
# sheet it should be added to the 'current knowledge'...

params = ['R0_LV','TRef_LV','VRef_LV','R0_HV','TRef_HV','VRef_HV','alpha',
          'beta','gamma','date','T_sensor']
R_data = [R1_LV,T_LV,V_LV,R1_HV,T_HV,V_HV,alpha,beta,gamma, date, 'none']
#R_dict = dict(zip(params,R_data))

if not R_INFO.has_key(R1_name):
    print 'Adding',R1_name,'to resistor info...'
    last_R_row = R_info.update_R_Info(R1_name,params,R_data,ws_Params,last_R_row,Run_Id,VERSION)
else:
    print 'Already know about',R1_name
    
# Save workbook
wb_io.save(xlfilename)
print '_____________HRBA DONE_______________'
