# -*- coding: utf-8 -*-
"""
R_info.py -
Utility functions and global definitions used by HRBA.py.

Created on Fri Sep 18 09:40:33 2015

@author: t.lawson
"""
import string
import datetime as dt
import time
import math
# import xlrd
# from openpyxl.styles import Font, colors, PatternFill, Border, Side
import GTC as gtc
import logging

RL_SEARCH_LIMIT = 500

DT_FORMAT = '%Y-%m-%d %H:%M:%S'

INF = 1e6  # 'inf' dof
ZERO = gtc.ureal(0, 0)

# Use for G1 & G2 in AUTO mode:
Vgain_codes_auto = {0.1: 'Vgain_0.1r0.1', 0.5: 'Vgain_0.5r1', 0.9: 'Vgain_1r1',
                    1.0: 'Vgain_1r1', 5.0: 'Vgain_5r10', 9.0: 'Vgain_10r10',
                    10.0: 'Vgain_10r10', 100.0: 'Vgain_100r100'}
# Use for G2 in FIXED mode:
Vgain_codes_fixed = {0.1: 'Vgain_0.5r1', 0.5: 'Vgain_0.5r1', 0.9: 'Vgain_1r1',
                     1.0: 'Vgain_1r10', 5.0: 'Vgain_5r10', 9.0: 'Vgain_10r10', 
                     10.0: 'Vgain_10r100', 100.0: 'Vgain_100r100'}


# ______________________________Useful funtions:______________________________
def get_run_info(curs, runid):
    """
    Read one record from 'Runs' table.
    :param curs: Database cursor object.
    :param runid: (str) Run ID.
    :return: (dict) run info.
    """
    q_run = f"SELECT * FROM Runs WHERE Run_Id='{runid}';"
    curs.execute(q_run)
    row = curs.fetchone()  # Should only be one row
    assert row != None, 'Error! Run_Id not found!'
    run_info = {'Comment': row[1], 'Rs_Name': row[2], 'Rx_Name': row[3], 'Range_Mode': row[4],
                'SRC1': row[5], 'SRC2': row[6], 'DVMd': row[7], 'DVM12': row[8],
                'GMH1': row[9], 'GMH2': row[10], 'GMHroom': row[11]}
    return run_info


def get_DVM_corrections(curs, dvm_name):
    """
    Read records from Instr_Info table.
    :param curs: Database cursor object.
    :param dvm_name: (str) Meter name.
    :return: (dict) containing gain and linearity info.
    """
    gains = {}
    q_I_info = (f"SELECT * FROM Instr_Info WHERE I_Name='{dvm_name}' AND "
                "Parameter LIKE 'Vgain%' OR Parameter LIKE 'linearity%';")
    curs.execute(q_I_info)
    rows = curs.fetchall()
    for row in rows:
        param = row[1]
        cor = uncertainize(row)
        gains.update({param: cor})
    return gains


def get_GMH_correction(curs, gmh_name):
    """
    Read one record from Instr_Info table.
    :param curs: Database cursor object.
    :param gmh_name: GMH probe name.
    :return: Temperature correction (gtc.ureal).
    """
    #print(f'GMH T_Correction for {gmh_name} ...')
    q_T_cor = (f"SELECT * FROM Instr_Info WHERE I_Name='{gmh_name}' "
               "AND Parameter='T_Correction';")
    curs.execute(q_T_cor)
    row = curs.fetchone()
    #print(row)
    return uncertainize(row)


def get_Rlink(curs, runid, run_info):
    """
    Calculate link resistance.
    :param curs: Database cursor object.
    :param runid: (str) Run ID.
    :param run_info: (dict) Run info.
    :return: Rlink (gtc.ureal).
    """
    # Read all raw Rlink data for this run:
    q_rlink_data = f"SELECT * FROM Raw_Rlink_Data WHERE Run_Id='{runid}';"
    curs.execute(q_rlink_data)
    rows = curs.fetchall()
    Vpos_lst = []
    Vneg_lst = []
    V1 = 0
    V2 = 0
    for row in rows:
        n_reads = row[1]
        V1 = row[2]
        V2 = row[3]
        Vpos_lst.append(row[4])
        Vneg_lst.append(row[5])
    '''
    Define nom_R, abs_V quantities. Assume all 'nominal' values
    have 100 ppm std.uncert. with 8 dof.
    '''
    val1 = get_r_val(run_info['Rx_Name'])
    nom_R1 = gtc.ureal(val1, val1 / 1e4, 8, label='nom_R1')  # don't >know< uncertainty of nominal values
    val2 = get_r_val(run_info['Rs_Name'])
    nom_R2 = gtc.ureal(val2, val2 / 1e4, 8, label='nom_R2')  # don't >know< uncertainty of nominal values
    abs_V1 = gtc.ureal(V1, V1 / 1e4, 8, label='abs_V1')  # don't >know< uncertainty of nominal values
    abs_V2 = gtc.ureal(V2, V2 / 1e4, 8, label='abs_V2')  # don't >know< uncertainty of nominal values

    # Calculate I
    I = gtc.result((abs_V1 + abs_V2) / (nom_R1 + nom_R2), f'Rd_I {runid}')

    # Average all +Vs and -Vs
    av_dV_p = gtc.result(gtc.ta.estimate(Vpos_lst), f'av_dV_p {runid}')
    av_dV_n = gtc.result(gtc.ta.estimate(Vneg_lst), f'av_dV_n {runid}')
    av_dV = gtc.result(0.5*gtc.magnitude(av_dV_p - av_dV_n), f'Rd_dV {runid}')

    # Finally, calculate Rd
    return gtc.result(av_dV/I, label=f'Rlink {runid}')


def ureal_to_str(un):
    """
    Encode a ureal as json string.
    :param un: Ureal.
    :return: (str) a json representation of ureal.
    """
    archive = gtc.pr.Archive()
    d = {un.label: un}
    archive.add(**d)
    return gtc.pr.dumps_json(archive)


def str_to_ureal(j_str, stored_name):
    """
    Decode a json string to a ureal.
    :param j_str: json string representation of a ureal.
    :param stored_name: (str) Archive-id of ureal.
    :return: (ureal) Uncertain number with dependencies intact.
    """
    print(f'Retrieving ureal (label={stored_name}) from:...\n', j_str)
    archive = gtc.pr.loads_json(j_str)
    return archive.extract(stored_name)


def get_Rs0(curs, Rs_name):
    """
    Extract 'book value' parameters for Rs.
    :param curs: Database cursor object.
    :param Rs_name: (str) Name of Rs in db 'Res_Info' table.
    :return: A dictionary of parameters.
    """
    print('get_Rs0():')
    Rs_0 = {}
    q_Rs = f"SELECT * FROM Res_Info WHERE R_Name='{Rs_name}';"
    curs.execute(q_Rs)
    Rs_rows = curs.fetchall()
    assert len(Rs_rows) > 0, f'No Rs data found for {Rs_name}!'
    for r in Rs_rows:
        param = r[1]
        val = r[2]
        unc = r[3]
        df = r[4]
        lbl = r[5]
        un_repr = r[7]
        if param == 'Cal_Date':
            Rs_0.update({param: val})
            continue
        if un_repr is None:  # Ureal WITHOUT proper traceability.
            if df is None:
                df = 1e6  # Ensure DoF is numeric - Don't want 'inf'!
            un = gtc.ureal(val, unc, df, label=lbl)
            Rs_0.update({param: un})
        else:  # have an archived ureal
            un = str_to_ureal(un_repr, lbl)
            # print(f"get_Rs0(): archive ureal - {param} = {un.x} +/- {un.u}, df = {un.df}")
            Rs_0.update({param: un})

    # Some resistors have no 2nd-order TCo (beta) listed in database, so
    # need to account for that by manually adding a zero-valued beta to
    # Rs_0 dictionary:
    if 'beta' not in Rs_0:
        Rs_0.update({'beta': gtc.ureal(0, 0, label=f'{Rs_name}_beta')})

    return Rs_0


def get_n_meas(curs, runid):
    q_get_n_meas = ("SELECT DISTINCT Meas_No FROM Raw_Data WHERE "
                    f"Run_Id='{runid}' ORDER BY Meas_No DESC LIMIT 1;")
    curs.execute(q_get_n_meas)
    row = curs.fetchone()
    return row[0]


def get_vgain(curs, DVM, param):
    print(f"Looking up gain code {param} for DVM {DVM}")
    q_Vgain = f"SELECT * FROM Instr_Info WHERE I_Name='{DVM}'AND Parameter ='{param}';"
    curs.execute(q_Vgain)
    row = curs.fetchone()
    return gtc.ureal(row[2], row[3], row[4], label=row[5])


def make_log_name(v):
    return f'HRBAv{str(v)}_{str(dt.date.today())}.log'


def extract_names(comment):
    """
    Extract resistor names from comment.
    Parse first part of comment for resistor names.
    Names must appear immediately after the strings 'R1: ' and 'R2: ' and
    immediately before the string ' monitored by GMH'.
    """
    assert comment.find('R1: ') >= 0, 'R1 name not found in comment!'
    assert comment.find('R2: ') >= 0, 'R2 name not found in comment!'
    r1_name = comment[comment.find('R1: ') + 4:comment.find(' monitored by GMH')]
    r2_name = comment[comment.find('R2: ') + 4:comment.rfind(' monitored by GMH')]
    return r1_name, r2_name


def get_r_val(name):
    """
    Extract nominal resistor value from name.
    Parse the resistor name for the nominal value. Resistor names MUST be of the form
    'xxx nnp', where 'xxx ' is a one-word description ending with a SINGLE SPACE,
    'nn' is an integer (usually a decade value) and the last character 'p' is a letter
    indicating a decade multiplier.
    """
    prefixes = {'r': 1, 'R': 1, 'k': 1000, 'M': 1e6, 'G': 1e9}

    if name[-1] in prefixes:
        mult = prefixes[name[-1]]
    else:
        mult = 0
    assert mult != 0, 'Error parsing comment - unkown multiplier!'

    # return numeric part of last word, multiplied by 1, 10^3, 10^6 or 10^9:
    # return mult*int(name.strip(name.split(name)[-1]), string.ascii_letters)
    r_val_str = name.split()[-1]
    return mult * int(r_val_str.strip(string.ascii_letters))


def write_this_result_to_db(curs, result_lst):
    headings = ("Run_Id, Meas_Date, Analysis_Note, Meas_No,"
                "Parameter, Value, Uncert, DoF, ExpU, k, Ureal_Str")
    for r in result_lst:
        values = (f"'{r['Run_Id']}','{r['Meas_Date']}','{r['Analysis_Note']}',"
                  f"{r['Meas_No']},'{r['Parameter']}',"
                  f"{r['Value']},{r['Uncert']},{r['DoF']},{r['ExpU']},{r['k']},'{r['repr']}'")
        q_write_res = f"INSERT OR REPLACE INTO Results ({headings}) VALUES ({values});"
        curs.execute(q_write_res)
    return None


def write_budget_line(curs, i, R1, result):
    headings = "Run_Id, Meas_No, Quantity_Label, Value, Uncert, DoF, Sens_Co, U_Contrib, Analysis_Note"
    sensitivity = gtc.rp.sensitivity(R1, i)
    component = abs(gtc.rp.u_component(R1, i))  # Can be -ve without abs()!
    if math.isinf(i.df):
        print(f"Writing budget line for {i.label} - Infinite dof detected: ({i.df}). Replacing with 1e6.")
        df = 1e6
    else:
        df = i.df
    values = (f"'{result['Run_Id']}','{result['Meas_No']}','{i.label}',{i.x},{i.u},{df},"
              f"{sensitivity},{component},'{result['Analysis_Note']}'")
    q_write_budget = f"INSERT OR REPLACE INTO Uncert_Contribs ({headings}) VALUES ({values});"
    # print(f"Query:\n{q_write_budget}")
    curs.execute(q_write_budget)
    return None


def uncertainize(row_items):
    """
    Convert list of data to ureal, where possible.
    If non-numeric, just return the 'value' part of input.
    """
    v = row_items[2]
    u = row_items[3]
    d = row_items[4]
    l = row_items[5]
    if (u is not None) and isinstance(v, (int, float)):  # Number
        if d == 'inf':
            un_num = gtc.ureal(v, u, label=l)  # default dof = inf
        else:
            un_num = gtc.ureal(v, u, d, label=l)
        return un_num
    else:  # non-numeric value
        return v


# def R_to_T(alpha, beta, R, R0, T0):
#     """
#     Convert a resistive T-sensor reading from resistance to temperature.
#
#     :parameter alpha (ureal) - temperature coefficient.
#     :parameter beta (ureal) - second-order temperature coefficient.
#     :parameter r (ureal) - resistance reading.
#     :parameter r0 (ureal) - calibration value of resistance.
#     :parameter T0 (ureal) - calibration value of temperature.
#     :return T (ureal) - temperature (deg C)
#     """
#     if beta == 0:  # no 2nd-order T-Co
#         T = (R / R0 - 1) / alpha + T0
#     else:
#         a = beta
#         b = alpha-2*T0
#         c = 1-alpha*T0 + beta*T0**2 - (R / R0)
#         T = (-b + gtc.sqrt(b**2-4*a*c))/(2*a)
#     return T


def av_t_dt(t_str_lst):
    """
    Calculate mean time from a list of strings.
    Return a string.
    :param t_str_lst: list of time strings.
    :return: a datetime object.
    """
    t_av = 0.0
    n = float(len(t_str_lst))
    for s in t_str_lst:
        t_dt = dt.datetime.strptime(s, DT_FORMAT)
        t_tup = dt.datetime.timetuple(t_dt)
        t_av += time.mktime(t_tup)
    t_av /= n  # av. time as float (seconds from epoch)
    t_av_dt = dt.datetime.fromtimestamp(t_av)  # datetime obj.
    return t_av_dt  #.strftime(DT_FORMAT)  # av. time as string


def av_t_string(t_list, switch):
    """
    Return average of a list of time-strings ("%d/%m/%Y %H:%M:%S")
    as a time string (switch='str') or float (switch='fl').
    """
    assert switch in ('fl', 'str'), 'Unknown switch for function av_t_string()!'
    throwaway = dt.datetime.strptime('20110101', '%Y%m%d')  # known bug fix
    n = float(len(t_list))
    t_av = 0.0
    for s in t_list:
        if isinstance(s, str):  # type(s) is str
            t_dt = dt.datetime.strptime(s, '%d/%m/%Y %H:%M:%S')
        elif type(s) is float:
            print(s, 'is a float...')
            t_dt = xlrd.xldate.xldate_as_datetime(s, 0)
        else:
            assert 0, 'Time format is not unicode or float!'
        t_tup = dt.datetime.timetuple(t_dt)
        t_av += time.mktime(t_tup)

    t_av /= n  # av. time as float (seconds from epoch)
    if switch == 'fl':
        return t_av
    elif switch == 'str':
        t_av_fl = dt.datetime.fromtimestamp(t_av)
        return t_av_fl.strftime('%d/%m/%Y %H:%M:%S')  # av. time as string


# def add_if_unique(item, lst):
#     """
#     Append 'item' to 'lst' only if it is not already present.
#     """
#     if item not in lst:
#         lst.append(item)
#     return lst


# def write_R1_T_fit(results, sheet, row, log):
#     """
#     Weighted least-squares fit (R1-T).
#     """
#     T_data = [T for T in [result['T'] for result in results]]  # All T values
#     unique_T_data = []
#     for T in T_data:
#         add_if_unique(T, unique_T_data)
#     T_av = gtc.fn.mean(T_data)
#     # print'write_R1_T_fit():u(T_av)=',T_av.u,'dof(T_av)=',T_av.df
#     T_rel = [t_k - T_av for t_k in T_data]  # x-vals shifted by T_av
#     alpha = gtc.ureal(0,0)
#
#     y = [R for R in [result['R'] for result in results]]  # All R values
#     u_y = [R.u for R in [result['R'] for result in results]] # All R uncerts
#
#     if len(unique_T_data) <= 1: # No temperature variation recorded, so can't fit to T
#         R1 = gtc.fn.mean(y)
#         print('R1_LV (av, not fit):', R1)
#         log.write('\nR1_LV (av, not fit): ' + str(R1))
#     else:
#         # a_ta,b_ta = GTC.ta.line_fit_wls(T_rel,y,u_y).a_b
#         # Assume uncert of individual measurements dominate uncert of fit
#         R1, alpha = gtc.ta.line_fit_wls(T_rel, y, u_y).a_b  # GTC.ta.line_fit_wls(T_rel, y).a_b
#         print('Fit params:\t intercept={}+/-{},dof={}. Slope={}+/-{},dof={}'.format(R1.x, R1.u, R1.df, alpha.x,
#                                                                                     alpha.u, alpha.df))
#         log.write('Fit params:\t intercept={}+/-{},dof={}. Slope={}+/-{},dof={}'.format(R1.x, R1.u, R1.df, alpha.x,
#                                                                                         alpha.u, alpha.df))

    # sheet['R'+str(row)] = R1.x
    # sheet['S'+str(row)] = R1.u
    # if math.isinf(R1.df):
    #     sheet['T'+str(row)] = str(R1.df)
    # else:
    #     sheet['T'+str(row)] = round(R1.df)
    #
    # sheet['U'+str(row)] = R1.u*gtc.rp.k_factor(R1.df)
    #
    # sheet['V'+str(row)] = T_av.x
    # sheet['W'+str(row)] = T_av.u
    # if math.isinf(T_av.df):
    #     sheet['X'+str(row)] = str(T_av.df)
    # else:
    #     sheet['X'+str(row)] = round(T_av.df)
    #
    # t = [result['time_fl'] for result in results]  # x data (time,s from epoch)
    # t_av = gtc.ta.estimate(t)
    # time_av = dt.datetime.fromtimestamp(t_av.x)  # A Python datetime object
    # sheet['Y'+str(row)] = time_av.strftime('%d/%m/%Y %H:%M:%S')  # string-formatted for display
    #
    # V1 = [V for V in [result['V'] for result in results]]
    # V_av = gtc.fn.mean(V1)
    # sheet['Z'+str(row)] = V_av.x
    # sheet['AA'+str(row)] = V_av.u
    # if math.isinf(V_av.df):
    #     sheet['AB'+str(row)] = str(V_av.df)
    # else:
    #     sheet['AB'+str(row)] = round(V_av.df)
    # return R1, alpha, T_av, V_av, time_av


def get_digi(readings):
    """
    Return maximum digitization level of a set of data.
    """
    max_n = 0
    for x in readings:
        if 'e' in str(x) or 'E' in str(x):
            (m, e) = math.modf(x)
            n = int(str(m).split('e')[1])
        else:
            n = -1*len(str(x).split('.')[1])

        if max_n < n:
            max_n = n
    d = 10**max_n
    return d
