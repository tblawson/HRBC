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
import datetime as dt
import math

from openpyxl import load_workbook, cell
from openpyxl.utils import get_column_letter

sys.path.append("C:\Python27\Lib\site-packages\GTC")
import GTC

import R_info  # useful functions

VERSION = 2.0

# DVM, GMH Correction factors, etc.

ZERO = GTC.ureal(0, 0)
PPM_TOLERANCE = {'R2': 2e-4, 'R1': 1e-3}

datadir = raw_input('Path to data directory:')
xlname = raw_input('Excel filename:')
xlfile = os.path.join(datadir, xlname)

logname = R_info.Make_Log_Name(VERSION)
logfile = os.path.join(datadir, logname)
log = open(logfile, 'a')

now_tup = dt.datetime.now()
now_fmt = now_tup.strftime('%d/%m/%Y %H:%M:%S')
log.write(now_fmt + '\n' + xlfile + '\n')

# open existing workbook
print str(xlfile)
# Open in NON-data-only mode, to preserve any fomulas when saving...
wb_io = load_workbook(xlfile, data_only=False)
ws_Data = wb_io.get_sheet_by_name('Data')
ws_Results = wb_io.get_sheet_by_name('Results')
ws_Params = wb_io.get_sheet_by_name('Parameters')

# Get local parameters
Data_start_row = ws_Data['B1'].value
Data_stop_row = ws_Data['B2'].value
assert Data_start_row <= Data_stop_row, 'Stop row must follow start row!'

# Get instrument assignments
N_ROLES = 9  # 9 roles in total
role_descr = {}
for row in range(Data_start_row, Data_start_row + N_ROLES):
    # Read {role:description}
    temp_dict = {ws_Data['AC' + str(row)].value:
                 ws_Data['AD' + str(row)].value}
    assert temp_dict.keys()[-1] is not None,\
        'Instrument assignment: Missing role!'
    assert temp_dict.values()[-1] is not None,\
        'Instrument assignment: Missing description!'
    role_descr.update(temp_dict)
    if ws_Data['AC'+str(row)].value == u'DVM':
        range_mode = ws_Data['AE'+str(row)].value
        print 'Range mode:', range_mode

# ------------------------------------------------------------------- #
# _____________Extract resistor and instrument parameters____________ #

print 'Reading parameters...'
log.write('Reading parameters...')
headings = (u'Resistor Info:', u'Instrument Info:',
            u'description', u'parameter', u'value',
            u'uncert', u'dof', u'label', u'Comment / Reference')

# Determine colummn indices from column letters:
col_A = cell.cell.column_index_from_string('A') - 1
col_B = cell.cell.column_index_from_string('B') - 1
col_C = cell.cell.column_index_from_string('C') - 1
col_D = cell.cell.column_index_from_string('D') - 1
col_E = cell.cell.column_index_from_string('E') - 1
col_F = cell.cell.column_index_from_string('F') - 1
col_G = cell.cell.column_index_from_string('G') - 1

col_I = cell.cell.column_index_from_string('I') - 1
col_J = cell.cell.column_index_from_string('J') - 1
col_K = cell.cell.column_index_from_string('K') - 1
col_L = cell.cell.column_index_from_string('L') - 1
col_M = cell.cell.column_index_from_string('M') - 1
col_N = cell.cell.column_index_from_string('N') - 1
col_O = cell.cell.column_index_from_string('O') - 1

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

for r in ws_Params.rows:  # a tuple of row objects
    R_end = 0

    # description, parameter, value, uncert, dof, label:
    R_row_items = [r[col_A].value, r[col_B].value, r[col_C].value,
                   r[col_D].value, r[col_E].value, r[col_F].value,
                   r[col_G].value]

    I_row_items = [r[col_I].value, r[col_J].value, r[col_K].value,
                   r[col_L].value, r[col_M].value, r[col_N].value,
                   r[col_O].value]

    if R_row_items[0] == None:  # end of R_list
        R_end = 1

    # check this row for heading text
    if any(i in I_row_items for i in headings):
        continue  # Skip headings

    else:  # not header - main data
        # Get instrument parameters first...
        last_I_row = r[col_I].row
        I_params.append(I_row_items[1])
        I_values.append(R_info.Uncertainize(I_row_items))
        if I_row_items[1] == u'test':  # Last parameter for this description
            I_DESCR.append(I_row_items[0])  # Build description list
            # Add parameter dictionary to sublist:
            I_sublist.append(dict(zip(I_params, I_values)))
            del I_params[:]
            del I_values[:]

        # Now attend to resistor parameters...
        if R_end == 0:  # Check we're not at the end of resistor data-block
            # Need to know this if we write more data, post-analysis:
            last_R_row = r[col_A].row
            R_params.append(R_row_items[1])
            R_values.append(R_info.Uncertainize(R_row_items))
            if R_row_items[1] == u'T_sensor':  # Last parameter for this descr
                R_DESCR.append(R_row_items[0])  # Build description list
                # Add parameter dictionary to sublist:
                R_sublist.append(dict(zip(R_params, R_values)))
                del R_params[:]
                del R_values[:]

# Compile into dictionaries
"""
There are two dictionaries; one for instruments (I_INFO) and one for resistors
(R_INFO). Each dictionary item is keyed by the description (name) of the
instrument (resistor). Each dictionary value is itself a dictionary, keyed by
parameter, such as 'address' (for an instrument) or 'R_LV' (for a resistor
value, measured at 'low voltage').

"""
I_INFO = dict(zip(I_DESCR, I_sublist))
print len(I_INFO), 'instruments (%d rows)' % last_I_row
log.write('\n'+str(len(I_INFO))+' instruments ('+str(last_I_row)+') rows')

R_INFO = dict(zip(R_DESCR, R_sublist))
print len(R_INFO), 'resistors.(%d rows)\n' % last_R_row
log.write('\n'+str(len(R_INFO))+' resistors ('+str(last_R_row)+') rows')

# -------------End of parameter extraction-------------- #
# ###################################################### #


# Determine the meanings of 'LV' and 'HV'
V1set_a = abs(ws_Data['A'+str(Data_start_row)].value)
assert V1set_a is not None, 'Missing initial V1 value!'
V1set_b = abs(ws_Data['A'+str(Data_start_row+6)].value)
assert V1set_b is not None, 'Missing second V1 value!'

if V1set_a < V1set_b:
    LV = V1set_a
    HV = V1set_b
elif V1set_b < V1set_a:
    LV = V1set_b
    HV = V1set_a
else:  # 'HV' and 'LV' equal
    LV = HV = V1set_a
print 'LV =', LV, '; HV =', HV

# Set up reading of Data sheet
Data_row = Data_start_row

# Get start_row on Results sheet
results_start_row = ws_Results['B1'].value
assert results_start_row is not None, 'Missing start row on Results sheet!'

# Get run identifier and copy to Results sheet
Run_Id = ws_Data['B'+str(Data_start_row-2)].value
assert Run_Id is not None, 'Missing Run Id!'

ws_Results['C'+str(results_start_row)] = 'Run Id:'
ws_Results['D'+str(results_start_row)] = str(Run_Id)

# Get run comment and extract R names & R values
Data_comment = ws_Data['AB'+str(Data_row)].value
assert Data_comment is not None, 'Missing Comment!'

print Data_comment
log.write('\n' + Data_comment)
print 'Run Id:', Run_Id
log.write('\nRun Id: ' + Run_Id)

# Write headings
results_row = R_info.WriteHeadings(ws_Results, results_start_row, VERSION)

# Lists of dictionaries (with name,time,R0,R1,Rd,RL,T,V entries)
results_HV = []  # High voltage measurements
results_LV = []  # Low voltage measurements

# Get resistor names and values
R1_name, R2_name = R_info.ExtractNames(Data_comment)
R1val = R_info.GetRval(R1_name)
R2val = R_info.GetRval(R2_name)

# Check for knowledge of R2:
if R2_name not in R_INFO:
    sys.exit('ERROR - Unknown Rs: '+R2_name)

raw_gmh1 = []
raw_gmh2 = []
T_dvm1 = []
T_dvm2 = []
R_dvm1 = []
R_dvm2 = []
times = []
RHs = []
Ps = []
Ts = []

# _____________________________________________________________________________

# ___Loop over data rows ___ #
print '\nLooping over data rows', Data_start_row, 'to', Data_stop_row, '...'
log.write('\nLooping over data rows ' + str(Data_start_row) +
          ' to ' + str(Data_stop_row) + '\n')
while Data_row <= Data_stop_row:
    # R2 parameters:
    V2set = ws_Data['B'+str(Data_row)].value
    assert V2set is not None, 'Missing V2 setting!'
    V1set = ws_Data['A'+str(Data_row)].value
    assert V1set is not None, 'Missing V1 setting!'

    # Select R2 info based on applied voltage ('LV' or 'HV')
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

    # Select appropriate value of VRC, etc.
    """
    #################################################################
    NOTE: Now replace VRCs with individual gain factors for
    each test-V (at mid- or top-of-range), on each instrument. Since
    this matches available info in DMM cal. cert. and minimises the
    number of possible values (ie: No. of test-Vs] < [No. of possible
    voltage ratios]).
    #################################################################
    """

    if 'AUTO' in range_mode:
        G2_code = R_info.Vgain_codes_auto[round(abs(V2set), 1)]
        G1_code = R_info.Vgain_codes_auto[round(V1set, 1)]
    else:  # 'FIXED'
        if round(V1set) >= round(abs(V2set)):
            G1_code = R_info.Vgain_codes_auto[round(V1set, 1)]
            G2_code = R_info.Vgain_codes_fixed[round(abs(V2set), 1)]
        else:
            G2_code = R_info.Vgain_codes_auto[round(abs(V2set), 1)]
            G1_code = R_info.Vgain_codes_fixed[round(V1set, 1)]

    # Gain corrections for V1, v2:
    G1 = I_INFO[role_descr['DVM']][G1_code]
    G2 = I_INFO[role_descr['DVM']][G2_code]

    '''
    Instead of defining vrc, apply individual gain corrections to each
    measurement:
    '''
#    vrc = GTC.ar.result(G2/G1, label='vrc ' + Run_Id)

#    Vlin_pert = I_INFO[role_descr['DVM']]['linearity_pert']  # G calc
    Vlin_Vnull = I_INFO[role_descr['DVM']]['linearity_Vdav']  # Vnull calcs

    # Start list of influence variables
    influencies = [G1, G2, Vlin_Vnull, R2_0, R2TRef, R2VRef]  # R2 dependancies

    R2alpha = R_INFO[R2_name]['alpha']
    R2beta = R_INFO[R2_name]['beta']
    R2gamma = R_INFO[R2_name]['gamma']
    R2Tsensor = R_INFO[R2_name]['T_sensor']
    influencies.extend([R2alpha, R2beta, R2gamma])  # R2 dependancies

    if R1_name not in R_INFO:
        R1Tsensor = 'Pt 100r'  # assume a Pt sensor in unknown resistor
    else:
        R1Tsensor = R_INFO[R1_name]['T_sensor']

    # GMH correction factors
    GMH1_cor = I_INFO[role_descr['GMH1']]['T_correction']
    GMH2_cor = I_INFO[role_descr['GMH2']]['T_correction']

    # Temperature measurement, RH and times:
    del raw_gmh1[:]  # list for 4 corrected T1 gmh readings
    del raw_gmh2[:]  # list for 4 corrected T2 gmh readings
    del T_dvm1[:]  # list for 4 corrected T1(dvm) readings
    del T_dvm2[:]  # list for 4 corrected T2(dvm) readings
    del R_dvm1[:]  # list for 4 corrected dvm readings
    del R_dvm2[:]  # list for 4 corrected dvm readings
    del times[:]  # list for 3*4 mean measurement time-strings
    del RHs[:]  # list for 4 RH values
    del Ps[:]  # list for 4 room pressure values
    del Ts[:]  # list for 4 room Temp values

    # Process times, RH and temperature data in this 4-row block:
    for r in range(Data_row, Data_row+4):  # build list of 4 gmh / dvm readings
        assert ws_Data['U'+str(r)].value is not None,\
            'No R1 GMH temperature data!'
        assert ws_Data['V'+str(r)].value is not None,\
            'No R2 GMH temperature data!'
        raw_gmh1.append(ws_Data['U'+str(r)].value)
        raw_gmh2.append(ws_Data['V'+str(r)].value)

        assert ws_Data['G'+str(r)].value is not None, 'No V2 timestamp!'
        assert ws_Data['M'+str(r)].value is not None, 'No Vd1 timestamp!'
        assert ws_Data['P'+str(r)].value is not None, 'No V1 timestamp!'
        times.append(ws_Data['G'+str(r)].value)
        times.append(ws_Data['M'+str(r)].value)
        times.append(ws_Data['P'+str(r)].value)

        assert ws_Data['S'+str(r)].value is not None,\
            'No R1 raw DVM (temperature) data!'
        raw_dvm1 = ws_Data['S'+str(r)].value

        assert ws_Data['T'+str(r)].value is not None,\
            'No R2 raw DVM (temperature) data!'
        raw_dvm2 = ws_Data['T'+str(r)].value

        # Check corrections for range-dependant values...
        # and apply appropriate corrections
        assert raw_dvm1 > 0, 'DVMT1: Negative resistance value!'
        assert raw_dvm2 > 0, 'DVMT2: Negative resistance value!'
        if raw_dvm1 < 120:
            T1DVM_cor = I_INFO[role_descr['DVMT1']]['correction_100r']
        elif raw_dvm1 < 12e3:
            T1DVM_cor = I_INFO[role_descr['DVMT1']]['correction_10k']
        else:
            T1DVM_cor = I_INFO[role_descr['DVMT1']]['correction_100k']
        R_dvm1.append(raw_dvm1*(1+T1DVM_cor))

        if raw_dvm2 < 120:
            T2DVM_cor = I_INFO[role_descr['DVMT2']]['correction_100r']
        elif raw_dvm2 < 12e3:
            T2DVM_cor = I_INFO[role_descr['DVMT2']]['correction_10k']
        else:
            T2DVM_cor = I_INFO[role_descr['DVMT2']]['correction_100k']
        R_dvm2.append(raw_dvm2*(1+T2DVM_cor))

    # Mean temperature from GMH
    # Data are plain numbers (with digitization rounding),
    # so use ta.estimate_digitized() to return a ureal.
    assert len(raw_gmh1) > 1, 'Not enough GMH1 temperatures to average!'
    T1_av_gmh = GTC.ar.result(GTC.ta.estimate_digitized(raw_gmh1, 0.01) +
                              GMH1_cor, label='T1_av_gmh ' + Run_Id)

    assert len(raw_gmh2) > 1, 'Not enough GMH2 temperatures to average!'
    T2_av_gmh = GTC.ar.result(GTC.ta.estimate_digitized(raw_gmh2, 0.01) +
                              GMH2_cor, label='T2_av_gmh ' + Run_Id)

    assert len(times) > 1, 'Not enough timestamps to average!'
    times_av_str = R_info.av_t_strin(times, 'str')  # Mean time(as a time-str)
    times_av_fl = R_info.av_t_strin(times, 'fl')  # Mean time(as a float)

    """
    TO DO: Incorporate ambient T, P, %RH readings into final reported results..

    assert len(RHs) > 1,'Not enough RH values to average!'
    # Digitization could be 2 or 3 decimal places, depending on RH probe:
    RH_av = GTC.ar.result(GTC.ta.estimate_digitized(RHs,R_info.GetDigi(RHs)),
                                                    label = 'RH_av')

    ... (and same for T, P) ...

    """

    # Build lists of 4 temperatures (calculated from T-probe dvm readings)...
    # ... and calculate mean temperatures
    if (R1Tsensor in ('none', 'any')):  # No or unknown T-sensor
        T_dvm1 = [ZERO, ZERO, ZERO, ZERO]
    else:
        assert len(R_dvm1) > 1, 'Not enough R_dvm1 values to average!'
        for R in R_dvm1:  # convert resistance measurement to a temperature
            T_dvm1.append(R_info.R_to_T(R_INFO[R1Tsensor]['alpha'],
                                        R_INFO[R1Tsensor]['beta'], R,
                                        R_INFO[R1Tsensor]['R0_LV'],
                                        R_INFO[R1Tsensor]['TRef_LV']))
    if R2Tsensor in ('none', 'any'):
        T_dvm2 = [ZERO, ZERO, ZERO, ZERO]
    else:
        assert len(R_dvm2) > 1, 'Not enough R_dvm2 values to average!'
        for R in R_dvm2:  # convert resistance measurement to a temperature
            T_dvm2.append(R_info.R_to_T(R_INFO[R2Tsensor]['alpha'],
                                        R_INFO[R2Tsensor]['beta'], R,
                                        R_INFO[R2Tsensor]['R0_LV'],
                                        R_INFO[R2Tsensor]['TRef_LV']))

    # Mean temperature from T-probe dvm
    # Data are high-precision plain numbers,
    # so use ta.estimate() to return a ureal.
    T1_av_dvm = GTC.ar.result(GTC.ta.estimate(T_dvm1),
                              label='T1_av_dvm' + Run_Id)
    T2_av_dvm = GTC.ar.result(GTC.ta.estimate(T_dvm2),
                              label='T2_av_dvm' + Run_Id)

    # Mean temperatures and temperature definitions
    T1_av = T1_av_gmh
    T1_av_dvm = GTC.ureal(0, 0)  # ignore any dvm data
    Diff_T1 = GTC.ureal(0, 0)  # No temperature disparity (GMH only)

    T2_av = T2_av_gmh
    T2_av_dvm = GTC.ureal(0, 0)  # ignore any dvm data
    Diff_T2 = GTC.ureal(0, 0)  # No temperature disparity (GMH only)

    # Default to using GMH probes for R2 temperature (not Pt)
    influencies.append(T2_av_gmh)  # R2 dependancy

    '''
    DEFAULT T definition arises from imperfect positioning of sensors
    wrt resistor:
    '''
    T_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.01), 3,
                      label='T_def ' + Run_Id)

    '''
    T-definition arises from imperfect positioning of BOTH probes
    AND their disagreement:
    '''
    T_def1 = GTC.ar.result(GTC.ureal(0, Diff_T1.u/2, 7) + T_def,
                           label='T_def1 ' + Run_Id)
    T_def2 = GTC.ar.result(GTC.ureal(0, Diff_T2.u/2, 7) + T_def,
                           label='T_def2 ' + Run_Id)
    influencies.append(T_def2)  # R2 dependancy

    # Raw voltage measurements: V: [Vp,Vm,Vpp; Vp',Vm',Vpp'], etc
    # All readings are precise enough not to worry about digitization error...
    V1 = []  # 6 normal readings
    V2 = []  # 3 normal, 3 perturbed
    Va = []  # 3 normal, 3 perturbed
    Vb = []  # 3 normal, 3 perturbed
    Vc = []  # 3 normal, 3 perturbed
    Vd = []  # 3 normal, 3 perturbed

    for line in range(6):

        V1.append(GTC.ureal(ws_Data['H'+str(Data_row+line)].value,
                            ws_Data['I'+str(Data_row+line)].value,
                            ws_Data['C'+str(Data_row+line)].value-1,
                            label='V1_'+str(line) + ' ' + Run_Id))
        V2.append(GTC.ureal(ws_Data['J'+str(Data_row+line)].value,
                            ws_Data['K'+str(Data_row+line)].value,
                            ws_Data['C'+str(Data_row+line)].value-1,
                            label='V2_'+str(line) + ' ' + Run_Id))
        Va.append(GTC.ureal(ws_Data['L'+str(Data_row+line)].value,
                            ws_Data['M'+str(Data_row+line)].value,
                            ws_Data['C'+str(Data_row+line)].value-1,
                            label='Va_'+str(line) + ' ' + Run_Id))
        Vb.append(GTC.ureal(ws_Data['N'+str(Data_row+line)].value,
                            ws_Data['O'+str(Data_row+line)].value,
                            ws_Data['C'+str(Data_row+line)].value-1,
                            label='Vb_'+str(line) + ' ' + Run_Id))
        Vc.append(GTC.ureal(ws_Data['P'+str(Data_row+line)].value,
                            ws_Data['Q'+str(Data_row+line)].value,
                            ws_Data['C'+str(Data_row+line)].value-1,
                            label='Vc_'+str(line) + ' ' + Run_Id))
        Vd.append(GTC.ureal(ws_Data['R'+str(Data_row+line)].value,
                            ws_Data['S'+str(Data_row+line)].value,
                            ws_Data['C'+str(Data_row+line)].value-1,
                            label='Vd_'+str(line) + ' ' + Run_Id))
        assert V1[-1] is not None, 'Missing V1 data!'
        assert V2[-1] is not None, 'Missing V2 data!'
        assert Va[-1] is not None, 'Missing Va data!'
        assert Vb[-1] is not None, 'Missing Vb data!'
        assert Vc[-1] is not None, 'Missing Vc data!'
        assert Vd[-1] is not None, 'Missing Vd data!'

    influencies.extend(V1+V2+Va+Vb+Vc+Vd)  # R2 dependancies - raw measurements

    # V2 perturbation
    delta_V2 = V2[3] - V2[2]
    drift_V2_1 = V2[3] - V2[0]
    drift_V2_2 = V2[5] - V2[3]
    av_drift_V2 = GTC.fn.mean(drift_V2_1, drift_V2_2)

    # Define drift in nulls:
    Vdrift_vals = []
    for V in (Va, Vb, Vc, Vd):
        delta_V = V[3] - V[2]  # Effect of V2 perturbation
        drift_1 = V[2] - V[0]
        drift_2 = V[5] - V[3]
        av_drift = GTC.ta.estimate(drift_1, drift_2)
        drift_u = GTC.magnitude(av_drift.x - (delta_V.x/delta_V2.x) *
                                av_drift_V2.x)/4
        Vdrift_vals.append(GTC.ureal(0,
                                     GTC.tb.distribution['gaussian'](drift_u),
                                     8))

    # Pack drift values into a dictionary and add labels:
    keys = ['Va', 'Vb', 'Vc', 'Vd']
    Vdrift = dict(zip(keys, Vdrift_vals))
    for key, val in Vdrift:
        val.label = 'Vdrift_' + key + ' ' + Run_Id
        influencies.append(Vdrift[key])  # R2 dependancies

    # Mean voltages (offset correction):
    V1_av = (V1[0]-2*V1[1]+V1[2]+V1[3]-2*V1[4]+V1[5])/8
    V2_av = (V2[0]-2*V2[1]+V2[2])/4
    V2p_av = (V2[3]-2*V2[4]+V2[5])/4
    Va_av = (Va[0]-2*Va[1]+Va[2])/4 + Vlin_Vnull + Vdrift['Va']
    Vap_av = (Va[3]-2*Va[4]+Va[5])/4 + Vlin_Vnull + Vdrift['Va']
    Vb_av = (Vb[0]-2*Vb[1]+Vb[2])/4 + Vlin_Vnull + Vdrift['Vb']
    Vbp_av = (Vb[3]-2*Vb[4]+Vb[5])/4 + Vlin_Vnull + Vdrift['Vb']
    Vc_av = (Vc[0]-2*Vc[1]+Vc[2])/4 + Vlin_Vnull + Vdrift['Vc']
    Vcp_av = (Vc[3]-2*Vc[4]+Vc[5])/4 + Vlin_Vnull + Vdrift['Vc']
    Vd_av = (Vd[0]-2*Vd[1]+Vd[2])/4 + Vlin_Vnull + Vdrift['Vd']
    Vdp_av = (Vd[3]-2*Vd[4]+Vd[5])/4 + Vlin_Vnull + Vdrift['Vd']

    # Voltage perturbations:
    deltaV2 = V2p_av-V2_av
    deltaVa = Vap_av-Va_av
    deltaVb = Vbp_av-Vb_av
    deltaVc = Vcp_av-Vc_av
    deltaVd = Vdp_av-Vd_av

    # Calculate R2 (corrected for T and V)
    dT2 = T2_av - R2TRef + T_def2

    # NOTE: NEED TWO abs() TO ENSURE NON-NEGATIVE DIFFERENCE:
    dV2 = abs(abs(V2_av) - R2VRef)

    R2 = R2_0*(1 + R2alpha * dT2 + R2beta * dT2 ** 2 + R2gamma * dV2)
    print 'R2 =', GTC.summary(R2)
    assert abs(R2.x-R2val)/R2val < PPM_TOLERANCE['R2'],\
        'R2 > 100 ppm from nominal! R2 = {0}'.format(R2.x)

    # Calculate R1
    R1_num = (deltaVb)*R2*V1_av*(V1_av*(deltaV2 - deltaVc) +
                                 Vcp_av*V2_av - V2p_av*Vc_av)
    R1_denom = (Vcp_av*V2_av - V2p_av*Vc_av)*(V1_av*(deltaVb - deltaV2) +
                                              V2p_av*Vb_av -
                                              Vbp_av*V2_av)
    R1 = R1_num/R1_denom
    print 'R1 =', GTC.summary(R1)
    assert abs(R1.x-R1val)/R1val < PPM_TOLERANCE['R1'],\
        'R1 > 1000 ppm from nominal!'

    # Calculate R0
    R0 = deltaVb*R2*V1_av/(V1_av*(deltaVb - deltaV2) + V2p_av*Vb_av -
                           Vbp_av*V2_av)
    print 'R0 =', GTC.summary(R0)
    assert R0.x > 1e7, 'DVM input-Z <= 10 MOhm!'

    # Calculate Rd
    Rd_num = deltaV2*R2*(V1_av*(deltaVa - deltaVb) - Vap_av*Vb_av +
                         Vbp_av*Va_av)
    Rd_denom = deltaVa*(V1_av*(deltaVb - deltaV2) + V2p_av*Vb_av -
                        Vbp_av*V2_av)
    Rd = Rd_num/Rd_denom
    print 'Rd =', GTC.summary(Rd)
    assert Rd.x < 0.05, 'High link resistance (> 50 mOhm)!'

    # Calculate RL
    RL_num = deltaVb*R2*(V1_av*(deltaVd - deltaV2) - Vdp_av*V2_av +
                         V2p_av*Vd_av)
    RL_denom = deltaV2*(V1_av*(deltaVb - deltaVd) + Vdp_av*Vb_av -
                        Vbp_av*Vd_av)
    RL = RL_num/RL_denom
    print 'RL =', GTC.summary(RL)
    assert RL.x > 1e9, 'Low leak resistance (<= 1 GOhm)!'

    T1 = T1_av + T_def1
    print R1, 'at temperature', T1

    '''
    Combine data for this measurement: name,time,R,T,V and write to
    Summary sheet:
    '''
    this_result = {'name': R1_name, 'time_str': times_av_str,
                   'time_fl': times_av_fl, 'V': V1_av, 'R': R1, 'R0': R0,
                   'Rd': Rd, 'RL': RL, 'T': T1,
                   'R_expU': R1.u*GTC.rp.k_factor(R1.df, quick=False)}

    R_info.WriteThisResult(ws_Results, results_row, this_result)

    # build uncertainty budget table
    budget_table = []  # A list of lists
    for i in influencies:  # rp.u_component(R1_gmh,i) gives + or - values
        if i.u >= 0:
            sensitivity = GTC.rp.u_component(R1, i)/i.u
        else:
            sensitivity = 0
        budget_table.append([i.label, i.x, i.u, i.df, sensitivity,
                             GTC.component(R1, i)])

    budget_table_sorted = sorted(budget_table, key=R_info.by_u_cont,
                                 reverse=True)

    # write budget to Summary sheet
    results_row = R_info.WriteBudget(ws_Results, results_row,
                                     budget_table_sorted)
    results_row += 1  # Add blank line between each measurement for clarity

    # Separate results by voltage (V1_av) if different
    if HV == LV:
        results_LV.append(this_result)
        results_HV.append(this_result)
    elif abs(V1_av.x - LV) < 1:
        results_LV.append(this_result)
    else:
        results_HV.append(this_result)

    del influencies[:]
    Data_row += 6  # Move to next measurement

# ----- End of data-row loop ---- #
# ############################### #


# At this point the summary row has reached its maximum for this analysis run
# ...so make a note of it, for use as the next run's starting row:
ws_Results['B1'] = results_row + 1  # Add extra row between runs

# Go back to the top of summary block, ready for writing run results
results_row = results_start_row + 1

########################################################################

"""
In the next section values of R1 are derived from fits to Temperature.
The Temperature data are offset so the mean is at ~zero, then the fits
are used to calculate R1 at the mean Temperature. LV and HV values are
obtained separately. The mean time, Temperature and Voltage values are
also reported.
"""

# Weighted total least-squares fit (R1-T), LV
print '\nLV:'
log.write('\nLV:')
R1_LV, Ohm_per_C_LV, T_LV, V_LV, date = R_info.write_R1_T_fit(results_LV,
                                                              ws_Results,
                                                              results_row, log)
alpha_LV = Ohm_per_C_LV/R1_LV

results_row += 1

# Weighted total least-squares fit (R1-T), HV
print '\nHV:'
log.write('\nHV:')
R1_HV, Ohm_per_C_HV, T_HV, V_HV, date = R_info.write_R1_T_fit(results_HV,
                                                              ws_Results,
                                                              results_row, log)
alpha_HV = Ohm_per_C_HV/R1_HV

alpha = GTC.fn.mean([alpha_LV, alpha_HV])
beta = GTC.ureal(0, 0)  # assume no beta

if HV == LV:  # Can't estimate gamma
    gamma = GTC.ureal(0, 0)
else:
    gamma = ((R1_HV-R1_LV)/(V_HV-V_LV))/R1_LV

results_row += 2

ws_Results['R'+str(results_row)] = 'alpha (/C)'
ws_Results['V'+str(results_row)] = 'gamma (/V)'

results_row += 1

ws_Results['R'+str(results_row)] = alpha.x
ws_Results['S'+str(results_row)] = alpha.u

if math.isinf(alpha.df):
    # print'alpha.df is',alpha.df
    ws_Results['T'+str(results_row)] = str(alpha.df)
else:
    # print'alpha.df =',alpha.df
    ws_Results['T'+str(results_row)] = round(alpha.df)

ws_Results['V'+str(results_row)] = gamma.x
ws_Results['W'+str(results_row)] = gamma.u
if math.isinf(gamma.df):
    # print'gamma.df is',gamma.df
    ws_Results['X'+str(results_row)] = str(gamma.df)
else:
    # print'gamma.df =',gamma.df
    ws_Results['X'+str(results_row)] = round(gamma.df)

#######################################################################

"""
Finally, if R1 is a resistor that is not included in the 'parameters'
sheet it should be added to the 'current knowledge'...
"""

params = ['R0_LV', 'TRef_LV', 'VRef_LV', 'R0_HV', 'TRef_HV', 'VRef_HV',
          'alpha', 'beta', 'gamma', 'date', 'T_sensor']

R_data = [R1_LV, T_LV, V_LV, R1_HV, T_HV, V_HV, alpha, beta, gamma, date,
          'none']

if R1_name not in R_INFO:
    print 'Adding', R1_name, 'to resistor info...'
    last_R_row = R_info.update_R_Info(R1_name, params, R_data, ws_Params,
                                      last_R_row, Run_Id, VERSION)
else:
    print '\nAlready know about', R1_name

# Save workbook
wb_io.save(xlfile)
print '_____________HRBA DONE_______________'
log.write('\n_____________HRBA DONE_______________\n\n')
log.close()
