# -*- coding: utf-8 -*-
"""
HRBA.py High Resistance Bridge Analysis (using imported GTC functions).

This script analyses data collected in the resistors_db database (tables:
'Runs', 'Raw_Data', 'Raw_Rlink_Data'), generated by the High Resistance
Bridge Control (HRBC) application. Additional correction factors come from
tables Instr_Info (for GMH probes and DVMs) and Res_Info (for Rs parameters).
Each 4-reversal block of data in the 'Raw_Data' table represents one
measurement of the unknown resistor R1('Rx') under a unique set of conditions
(time t, Temperature T, Voltage V).

Each HRBA run analyses one HRBC run, which normally consists of several
measurements at alternating low or high voltage (LV,HV). HRBA searches
the 'Raw_Rlink_Data' table for a block of Rlink data that has a matching
run ID and uses this information to calculate the link resistance Rd.

The temperature definition uncertainty of GMH readings of resistor
temperatures defaults to 0 +/- 0.05 C with 4 degrees of freedom.

The multiple results from one analysis run are added to the 'Results'
table. For each result record, an uncertainty budget is compiled and
recorded in the 'Uncert_Contribs' table - one record for each contribution.

Created on Fri Sep 18 14:01:18 2015
This branch started: Tues Aug 18 21:48 2020

@author: t.lawson
"""

import os
import logging
import sys
# sys.path.append("C:\Python27\Lib\site-packages\GTC")

import datetime as dt
import math

import GTC as gtc
import sqlite3

import R_info  # useful functions

VERSION = 2.1  # 2nd Python 3 version, 1st db-based version.
DT_FORMAT = R_info.DT_FORMAT  # '%Y-%m-%d %H:%M:%S'
# DVM, GMH Correction factors, etc.
ZERO = gtc.ureal(0, 0)
CHECK_QUALITY = False
FRAC_TOLERANCE = {'R2': 2e-2, 'G': 0.01, 'R1': 0.005}  # {'R2': 2e-4, 'G': 0.01, 'R1': 5e-2}
RLINK_MAX = 2000  # Ohms

datadir = r'G:\My Drive'

# _________________Set up logging:_________________

logname = R_info.make_log_name(VERSION)
logging.basicConfig(filename=logname, format='%(asctime)s %(levelname)s: %(message)s', level=logging.INFO)
logging.info('Starting...')

# _________________Connect to Resistors.db database:________________
db_path = input('Full Resistors.db path? (press "d" for default location) >')
if db_path == 'd':
    db_path = r'G:\My Drive\Resistors.db'  # Default location.
db_connection = sqlite3.connect(db_path)
curs = db_connection.cursor()

# _________________________Get run info:___________________________
runid = input('Run_id to analyse? > ')
run_info = R_info.get_run_info(curs, runid)

# _____Import corrections for assigned DVMs and GMH probes:_____
DVMd_params = R_info.get_DVM_corrections(curs, run_info['DVMd'])
DVM12_params = R_info.get_DVM_corrections(curs, run_info['DVM12'])
GMH1_correction = R_info.get_GMH_correction(curs, run_info['GMH1'])
GMH2_correction = R_info.get_GMH_correction(curs, run_info['GMH2'])

# _______________________Calculate R_link:_______________________
Rd = R_info.get_Rlink(curs, runid, run_info)
print(f'\nRlink = {Rd.x} +/- {Rd.u}, dof = {Rd.df}')
assert Rd.x < RLINK_MAX, f'High link resistance! {Rd.x} Ohm'
logging.info(f'\nRlink = {Rd.x} +/- {Rd.u}, dof = {Rd.df}')

# _____________________Get Rs 'book value'...:____________________
Rs_0 = R_info.get_Rs0(curs, run_info['Rs_Name'])  # Dict to hold all Rs info.
logging.info(f"'Rs_0 = {Rs_0['R0'].x} +/- {Rs_0['R0'].u}, dof = {Rs_0['R0'].df}")
print(f"'Rs 'book-value' = {Rs_0['R0'].x} +/- {Rs_0['R0'].u}, dof = {Rs_0['R0'].df}")

# Define temperature definition of +/- 0.05 C to account for typical
# ...separation between R1 resistor element and GMH probe:
T_def1 = gtc.ureal(0, gtc.type_b.distribution['gaussian'](0.05), 3,
                       label=f'T_def {runid}')

# ________Determine number of measurements in this run:_________
n_meas = R_info.get_n_meas(curs, runid)

# _________Write analysis note to Runs table:_________
analysis_note = (f"Processed with HRBA v{VERSION} on "
                 f"{dt.datetime.now().strftime('%A, %d. %B %Y %I:%M%p')}")
q_a_note = f"UPDATE Runs SET Analysis_Note = '{analysis_note}' WHERE Run_id = '{runid}';"
curs.execute(q_a_note)

# _____________Loop over measurements in this run:_____________
for meas_no in range(1, n_meas+1):  # 1 to <max meas_no>
    '''
    -----------------------------------------------------------
    Loop over measurements.
    - Process raw data for all 4 reversals in each measurement,
    - Calculate result,
    - Write result to 'Results' table.
    - Write uncertainty contributions to 'Uncert_Contribs' table.
    -----------------------------------------------------------
    '''
    # ________Start list of R1-influence variables (for budget)_______
    influences = [Rd]
    # ______________... and include influences for Rs :_______________
    influences.extend([Rs_0['R0'], Rs_0['TRef'], Rs_0['VRef'], Rs_0['alpha'],
                       Rs_0['beta'], Rs_0['gamma'], Rs_0['tau'], Rs_0['Tdef']])

    # Return the 4 reversals for this measurement, in order:
    q_get_revs = (f"SELECT * FROM Raw_Data WHERE Run_Id='{runid}' "
                  f"AND Meas_No={meas_no} ORDER BY Rev_No ASC;")
    curs.execute(q_get_revs)
    revs = curs.fetchall()

    V1_times = []
    V1_lst = []  # list of ureals
    V2_times = []
    V2_lst = []  # list of ureals
    Vd_times = []
    Vd_lst = []  # list of ureals
    T1_lst = []
    T2_lst = []
    T_room_lst = []
    P_room_lst = []
    RH_room_lst = []
    # ___________________4* reversals loop________________
    for row in revs:
        rev_no = row[2]
        V1set = row[3]
        V2set = row[4]
        n = row[5]
        V1_times.append(row[9])
        un = gtc.ureal(row[10], row[11], n - 1,
                       label=f'V1_{rev_no}_meas{meas_no}_{runid}')
        V1_lst.append(un)
        Vd_times.append(row[12])
        un = gtc.ureal(row[13], row[14], n - 1,
                       label=f'Vd_{rev_no}_meas{meas_no}_{runid}')
        Vd_lst.append(un)
        V2_times.append(row[15])
        un = gtc.ureal(row[16], row[17], n - 1,
                       label=f'V2_{rev_no}_meas{meas_no}_{runid}')
        V2_lst.append(un)
        T1_lst.append(row[18])
        T2_lst.append(row[19])
        T_room_lst.append(row[20])
        P_room_lst.append(row[21])
        RH_room_lst.append(row[22])
    # ___________________4* reversals loop________________

    # _________________Define drift:_________________
    Vd_shift = Vd_lst[3].x - Vd_lst[2].x
    V2_shift = V2_lst[3].x - V2_lst[2].x
    V2_drift = V2_lst[2].x - V2_lst[0].x
    Vd_drift = Vd_lst[2].x - Vd_lst[0].x
    drift_unc = abs(Vd_drift + (Vd_shift/V2_shift)*V2_drift)/4
    Vdrift1 = gtc.ureal(0, gtc.tb.distribution['gaussian'](drift_unc), 8,
                        label='Vdrift_pert ' + runid)
    Vdrift2 = gtc.ureal(0, gtc.tb.distribution['gaussian'](drift_unc), 8,
                        label='Vdrift_Vdav ' + runid)
    Vdrift = {'pert': Vdrift1,  'Vdav': Vdrift2}

    influences.extend([Vdrift['pert'], Vdrift['Vdav']])

    # ________________Mean voltages:________________
    V1av = gtc.result((V1_lst[0] - 2 * V1_lst[1] + V1_lst[2]) / 4,
                      label=f"{run_info['Rx_Name']}_V_meas={meas_no}_{runid}")
    V2av = (V2_lst[0] - 2 * V2_lst[1] + V2_lst[2]) / 4
    Vlin_Vdav = DVMd_params['linearity_Vdav']
    Vdav = (Vd_lst[0] - 2 * Vd_lst[1] + Vd_lst[2]) / 4 + Vlin_Vdav + Vdrift['Vdav']

    # _______________Effect of v2 perturbation:______________
    Vlin_pert = DVMd_params['linearity_pert']
    delta_Vd = Vd_lst[3] - Vd_lst[2] + Vlin_pert + Vdrift['pert']
    delta_V2 = V2_lst[3] - V2_lst[2]

    influences.extend([Vlin_pert, Vlin_Vdav] + V1_lst[0:3] + V2_lst + Vd_lst)  # V1_lst[3] not used.

    # ________________Calculate Rx (= R1)...______________
    # ... Start by calculating Rs:

    # ______________Temperature offsets:________________
    # Mean temperature from GMH probes.
    # Data are plain numbers (with digitization rounding),
    # ...so use ta.estimate_digitized() to return a ureal:
    Tav_1 = gtc.result(gtc.ta.estimate_digitized(T1_lst, 0.01) + GMH1_correction + T_def1,
                       label=f"{run_info['Rx_Name']}_T_meas={meas_no}_{runid}")
    Tav_2 = gtc.result(gtc.ta.estimate_digitized(T2_lst, 0.01) + GMH2_correction + Rs_0['Tdef'],
                       label=f'T_av2 {runid}')
    dT = Tav_2 - Rs_0['TRef']

    influences.extend([Tav_2])

    # ____________________Time offset:____________________
    t_av_dt = R_info.av_t_dt(V1_times + V2_times + Vd_times)  # Av time as datetime obj.
    t_av_string = t_av_dt.strftime(DT_FORMAT)  # av. time as string
    Rs_0_dt = dt.datetime.strptime(Rs_0['Cal_Date'], DT_FORMAT)  # Convert string to datetime obj.
    diff = t_av_dt - Rs_0_dt  # A timedelta obj.
    dt_days = diff.days + diff.seconds / 86400

    # ___________V offset - define mean test-V:____________
    # V_av2 = gtc.fn.mean(V2_lst)  # seq of ureals -> ureal.
    dV = abs(V2av - Rs_0['VRef'])  # dV must be positive.

    # ______Now we can calculate Rs (includes R_link):_______
    Rs = Rd + Rs_0['R0']*(1 + Rs_0['alpha']*dT +
                          Rs_0['beta']*dT**2 +
                          Rs_0['gamma']*abs(dV) +
                          Rs_0['tau']*dt_days)
    logging.info(f"\tMeas_no. {meas_no}: Rs = {Rs.x} +/- {Rs.u}, dof = {Rs.df}")

    # Next, calculate Rx influences:
    V2rnd = math.pow(10, round(math.log10(abs(V2set))))  # Rnd to nearest 10-pwr
    V1rnd = math.pow(10, round(math.log10(abs(V1set))))
    if 'AUTO' in run_info['Range_Mode']:
        '''
        Set ranges = to rounded V setting.
        I.e: V1range = V1rnd; V2range = V2rnd
        '''
        G2_code = R_info.Vgain_codes_auto[V2rnd]
        G1_code = R_info.Vgain_codes_auto[V1rnd]
    else:  # 'FIXED'
        '''
        Set both ranges = to highest V setting.
        I.e: V1range = V2range = max(V1rnd, v2rnd)
        '''
        if round(abs(V1set)) >= round(abs(V2set)):
            G1_code = R_info.Vgain_codes_auto[V1rnd]
            G2_code = R_info.Vgain_codes_fixed[V2rnd]
        else:
            G2_code = R_info.Vgain_codes_auto[V2rnd]
            G1_code = R_info.Vgain_codes_fixed[V1rnd]

    # __________Voltage ratio correction:___________
    """
    NOTE: Now replace VRCs with individual gain factors for
    each test-V (at mid- or top-of-range), on each instrument. Since
    this matches available info in DMM cal. cert. and minimises the
    number of possible values (ie: No. of test-Vs] < [No. of possible
    voltage ratios]).
    """
    print(f"G1_code = {G1_code}. G2_code = {G2_code}")
    G1 = R_info.get_vgain(curs, run_info['DVM12'], G1_code)
    G2 = R_info.get_vgain(curs, run_info['DVM12'], G2_code)
    vrc = gtc.result(G2 / G1, label='vrc ' + runid)

    influences.extend([G1, G2])

    # _____________________Calculate R1:_____________________
    R1 = gtc.result(Rs*vrc*V1av*delta_Vd / (Vdav*delta_V2 - V2av*delta_Vd),
                    label=f"{run_info['Rx_Name']}_R_meas={meas_no}_{runid}")
    if CHECK_QUALITY:
        assert R1.x > 0, (f'\nCalculation error! R1 <= zero!: {R1.x}. '
                         f'Run info:{runid} meas={meas_no}')
        assert R1.u < R1.x, (f'\nR1 Warning! Uncert >= value: {R1}. '
                             f'Run info:{runid} meas={meas_no}')

        R1val = R_info.get_r_val(run_info['Rx_Name'])
        frac_err = abs(R1.x - R1val) / R1val
        assert frac_err < FRAC_TOLERANCE['R1'], (f"R1 > {FRAC_TOLERANCE['R1']*1e6} ppm "
                                                 "from nominal ({frac_err})!")

    print('R1 = {} +/- {}, dof = {}'.format(R1.x, R1.u, R1.df))
    logging.info(f"\tR1 = {R1.x} +/- {R1.u}, dof = {R1.df}")

    # ____Corrected R1 temperature, including definition uncert.:____
    T1 = Tav_1 + T_def1
    print(f"{R1.x} at temperature {T1.x} and test-voltage {V1av.x}")
    logging.info(f"{R1.x} at temperature {T1.x} and test-voltage {V1av.x}")

    # _________Gather info for measurement results:_________
    R1_k = gtc.rp.k_factor(R1.df)
    R1_EU = R1.u*R1_k
    this_result_R = {'Run_Id': runid, 'Meas_Date': t_av_string, 'Analysis_Note': analysis_note,
                     'Meas_No': meas_no, 'Parameter': 'R',
                     'Value': R1.x, 'Uncert': R1.u, 'DoF': R1.df,
                     'ExpU': R1_EU, 'k': R1_k, 'repr': R_info.ureal_to_str(R1)}

    V1_k = gtc.rp.k_factor(V1av.df)
    V1_EU = V1av.u*V1_k
    this_result_V = {'Run_Id': runid, 'Meas_Date': t_av_string, 'Analysis_Note': analysis_note,
                     'Meas_No': meas_no, 'Parameter': 'V',
                     'Value': V1av.x, 'Uncert': V1av.u, 'DoF': V1av.df,
                     'ExpU': V1_EU, 'k': V1_k, 'repr': R_info.ureal_to_str(V1av)}

    T1_k = gtc.rp.k_factor(Tav_1.df)
    T1_EU = Tav_1.u * V1_k
    this_result_T = {'Run_Id': runid, 'Meas_Date': t_av_string, 'Analysis_Note': analysis_note,
                     'Meas_No': meas_no, 'Parameter': 'T',
                     'Value': Tav_1.x, 'Uncert': Tav_1.u, 'DoF': Tav_1.df,
                     'ExpU': T1_EU, 'k': T1_k, 'repr': R_info.ureal_to_str(Tav_1)}

    # ------------- Write to Results table: --------------
    R_info.write_this_result_to_db(curs, [this_result_R, this_result_V, this_result_T])

    # ---------- Write to Uncert_Contribs table: ----------
    for i in influences:
        R_info.write_budget_line(curs, i, R1, this_result_R)
# _____________End of measurements loop._____________

# ______________Tidy up:______________
db_connection.commit()
curs.close()

print('_____________HRBA DONE_______________')
logging.info('_____________HRBA DONE_______________')


