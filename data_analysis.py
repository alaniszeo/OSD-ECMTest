from pathlib import Path as path
import pandas as pd
import CoolProp.CoolProp as CP
import numpy as np


def compute_heat_transfer_coef(D_h,T,w,v,P_atm=101325):
    
    # Computation of air properties
    nu = CP.HAPropsSI('mu','T',T+273.15,'W',w,'P',P_atm)*CP.HAPropsSI('V','T',T+273.15,'W',w,'P',P_atm)
    k = CP.HAPropsSI('K','T',T+273.15,'W',w,'P',P_atm)
    
    # Reynolds
    Re = v*D_h/nu
    
    # Nussels --> Vary Nu depending on Re !
    Nu = 7.54
    
    # Heat transfer coefficient
    h_T = Nu*k/D_h
    
    return Re, Nu, h_T

def analyze_data(HEX_type,filename,sheetname,P_atm=101325,data_type='heat_transfer'):
    cur_dir = path.cwd()
    file_path = cur_dir.joinpath(HEX_type,filename)
    
    match HEX_type:
        case 'DEC':
            skiprow = 28
            data = pd.read_excel(file_path,usecols='B:R',index_col=0,skiprows=skiprow)
            
            T_pi = data['T_pwi'].values
            
            if data_type == 'pre-treatment':
                T_po = data['T_pwo'].values
                w_pi_name = 'w_pwi'
                
            elif data_type == 'heat_transfer':
                v_pi = data['v_pwi'].values
                w_pi = data['w_pwi'].values
                h_T_name = 'h_T'
                
        case 'Two-stage_EC':
            skiprow = 32
            data = pd.read_excel(file_path,usecols='B:AF',index_col=0,skiprows=skiprow)
            
            T_pi = data['T_pdi'].values
            T_po = data['T_pdo'].values
            
            if data_type == 'pre-treatment':
                w_pi_name = 'w_pdi'
                
            elif data_type == 'heat_transfer':
                v_pi = data['v_pdi'].values
                w_pi = data['w_pdi'].values
                h_T_name = 'h_T_pd'
     
    if data_type == 'pre-treatment':
        epsilon_wb = data['epsilon_wb'].values
        
        T_wbi = T_pi - (T_pi-T_po)/epsilon_wb
        T_wbi = np.maximum(T_wbi,15)
        
        w_pi = CP.HAPropsSI('W','T',T_pi+273.15,'B',T_wbi+273.15,'P',P_atm)
        
        data.loc[:,w_pi_name] = w_pi
        
    elif data_type == 'heat_transfer':
        # Computation of heat transfer coefficient on primary side
        WWR = data['WWR'].values
        h_ch = data['h_ch'].values
        D_h = 2*h_ch/WWR
        
        Re_p, Nu, h_T_p = compute_heat_transfer_coef(D_h,T_pi,w_pi,v_pi)
        
        data.loc[:,'Re'] = Re_p
        data.loc[:,'Nu'] = Nu
        data.loc[:,h_T_name] = h_T_p
        
        if HEX_type == 'Two-stage_EC':
            # HEX 1 - IEC
            v_sw_1 = data['v_swi_1'].values
            Re_sw_1, Nu, h_T_sw_1 = compute_heat_transfer_coef(D_h,T_pi,w_pi,v_sw_1)
            data.loc[:,'h_T_sw_1'] = h_T_sw_1
            
            # HEX 2 - D-IEC
            v_sw_2 = data['v_swi_2'].values
            Re_sw_2, Nu, h_T_sw_2 = compute_heat_transfer_coef(D_h,T_po,w_pi,v_sw_2)
            data.loc[:,'h_T_sw_2'] = h_T_sw_2
            
        
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            data.to_excel(writer,sheet_name=sheetname,startrow=skiprow,startcol=1)
     
    return data



if __name__ == "__main__":
    HEX_type = 'Two-stage_EC'
    sheetname = 'Pacak2022'
    
    # HEX_type = 'DEC'
    # sheetname = 'Nada2019'
    
    filename = sheetname+'.xlsx'
    
    data = analyze_data(HEX_type,filename,sheetname,data_type='heat_transfer')

