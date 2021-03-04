import openpyxl , sys , os
import pandas as pd 
from datetime import datetime
import tkinter as tk
from tkinter import RAISED
import concurrent.futures
import tkinter.font as font
from tkinter import filedialog , ttk , messagebox
from PIL import ImageTk , Image


filenames = {"Mercury Roster":"" , "Blended Rate":"","Report Location":"", "from Period":"","to Period":"","us_attr":"",
            "us_hir":"","usi_attr":"","usi_hir":"","usdc_attr":"","usdc_hir":"","demand file":"","availability file":""}

def running_sequence(value):
    try:
        dataframe = data_frame('Notes')
        dataframe_header = list(dataframe.columns.values)
        rows,_ = dataframe.shape
        date_format = '%Y/%m/%d %H:%M:%S'
        value_string_time = value.strftime("%Y/%m/%d %H:%M:%S")
        last_value_str = (dataframe.iat[rows-1,dataframe_header.index('End Date')]).strftime("%Y/%m/%d %H:%M:%S")
        last_value = datetime.strptime(last_value_str , date_format)
        first_value_str = (dataframe.iat[0,dataframe_header.index('Start Date')]).strftime("%Y/%m/%d %H:%M:%S")
        first_value = datetime.strptime(first_value_str , date_format)

        if  first_value <= datetime.strptime(value_string_time , date_format) <= last_value:
            for row in range(0,rows,1):
                start_value_str = (dataframe.iat[row,dataframe_header.index('Start Date')]).strftime("%Y/%m/%d %H:%M:%S")
                end_value_str = (dataframe.iat[row,dataframe_header.index('End Date')]).strftime("%Y/%m/%d %H:%M:%S")
                start_value = datetime.strptime(start_value_str , date_format)
                end_value = datetime.strptime(end_value_str , date_format)
                if start_value <= datetime.strptime(value_string_time , date_format) and datetime.strptime(value_string_time , date_format) <= end_value:
                    return dataframe.iat[row,dataframe_header.index('Running Sequence')]

        elif datetime.strptime(value_string_time , date_format) > last_value:
            return dataframe.iat[rows-1,dataframe_header.index('Running Sequence')]

        elif datetime.strptime(value_string_time , date_format) < first_value:
            return dataframe.iat[0,dataframe_header.index('Running Sequence')]      
    except:
        return 0
    
def data_frame(sheet_name):
    
    if sheet_name == 'Mercury Extract':
        dataframe = pd.read_excel(filenames['Mercury Roster'], sheet_name = sheet_name) 
        filter_1 = dataframe["Market Offering Sub-Solution"] == "ORACLE"
        return dataframe[filter_1]

    elif sheet_name == 'Oracle':
        filter_practice = ['USI Offshore','USI Onsite','Commercial Core','Commercial USDC']
        filter_status = ['Open - In Progress','Open - Published','Submitted']
        dataframe = pd.read_excel(filenames['demand file'], sheet_name = sheet_name)
        dataframe = dataframe[(dataframe['Role Status'].isin(filter_status)) & (dataframe['Practice (Role)'].isin(filter_practice))]
        return dataframe

    elif sheet_name == 'Source_Data':
        filter_billing = ['Full Time Billable','Hold Future Confirmed','Internal Project']
        filter_assgn = ['']
        dataframe = pd.read_excel(filenames['availability file'], sheet_name = sheet_name)
        dataframe = dataframe[(dataframe['Billing Status'].isin(filter_billing)) & (~dataframe.Assignment.isin(filter_assgn))]
        return dataframe

    else:
        config_path = os.path.join(os.getcwd(), "Configuration_File.xlsx")
        dataframe = pd.read_excel(config_path, sheet_name = sheet_name)
        return dataframe

def generate_inputs_sheet():
        try:
            df = data_frame('Mercury Extract')
            df_row,_ = df.shape
            df_header = list(df.columns.values)
            writer = pd.ExcelWriter((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data.xlsx'), engine='xlsxwriter',date_format='mm/dd/yyyy',datetime_format='mm/dd/yyyy')
            df.to_excel(writer, sheet_name='Mercury Roster',index= False)
            writer.save()
            writer.close()
        except:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"The Sheet name of Input Analysis File is incorrect")
        try:
            df_const = data_frame('Staticvalues')
            df_const_header = list(df_const.columns.values)
        except:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"Place file 'Mandatory_Values.xlsx' in the same location as of .exe file and give correct name")

        try:
            df_input_sheet = pd.DataFrame(columns=['Project','Probability','Phase','Package','Market Offering Value','Project Start',
                                                    'Project End','Period Number Start','Period Number End','$ Per period','Blended Rate',
                                                    'Hours per Period','Hours per Person','People per Period'])

            df_input_sheet_header = list(df_input_sheet.columns.values)
            thread_values = [ row for row in range(0,df_row,1)]

            def dataframes(row):
                Project = df.iat[row,df_header.index('Account Name')] + '_' + df.iat[row,df_header.index('Opportunity ID')] + '_' + df.iat[row,df_header.index('Opportunity Name')]
                Probability = df.iat[row,df_header.index('Probability (%)')]
                Phase = df.iat[row,df_header.index('Phase')]
                Package = df.iat[row,df_header.index('Market Offering Sub-Solution')]
                Market_Offering_Value = df.iat[row,df_header.index('Market Offering Value')]
                Project_Start = df.iat[row,df_header.index('Revised Project Start Date')]
                Project_End = df.iat[row,df_header.index('Revised Project End Date')]
                Period_Number_Start = running_sequence(Project_Start)
                Period_Number_End = running_sequence(Project_End)
                if Period_Number_Start == Period_Number_End:
                    Per_period = Market_Offering_Value
                else:
                    Per_period = Market_Offering_Value / ((Period_Number_End - Period_Number_Start) + 1)
                Blended_Rate = int(filenames["Blended Rate"])
                Hours_per_Period = Per_period / Blended_Rate
                Hours_per_Person = df_const.iat[0,df_const_header.index('Hours per Person')]
                People_per_Period = Hours_per_Period / Hours_per_Person

                return pd.DataFrame({df_input_sheet_header[0]:Project,df_input_sheet_header[1]:Probability,
                                        df_input_sheet_header[2]:Phase,df_input_sheet_header[3]:Package,df_input_sheet_header[4]:Market_Offering_Value,
                                        df_input_sheet_header[5]:Project_Start,df_input_sheet_header[6]:Project_End,df_input_sheet_header[7]:Period_Number_Start,
                                        df_input_sheet_header[8]:Period_Number_End,df_input_sheet_header[9]:Per_period,df_input_sheet_header[10]:Blended_Rate,
                                        df_input_sheet_header[11]:Hours_per_Period,df_input_sheet_header[12]:Hours_per_Person,
                                        df_input_sheet_header[13]:People_per_Period},index=[0])

            with concurrent.futures.ThreadPoolExecutor() as executor:
                results = executor.map(dataframes,thread_values)
            for df in results:
                df_input_sheet = df_input_sheet.append(df, ignore_index=True)

            
            return df_input_sheet

        except ValueError:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"The header names either in Input Analysis or configuration File are incorrect ")

def mercury_forcast_estimation():
    try:
        df_input = generate_inputs_sheet()
        try:
            df_demand_forecast = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
            df_demand_forecast_header = list(df_demand_forecast.columns.values)

            df_staffmix= data_frame('StaffMix')
            df_staffmix_rows,_ = df_staffmix.shape
            df_staffmix_header = list(df_staffmix.columns.values)
            df_input_rows,_ = df_input.shape
            df_input_header = list(df_input.columns.values)
            df_notes = data_frame('Notes')
            df_notes_header = list(df_notes.columns.values)
            df_const = data_frame('Staticvalues')
            df_const_header = list(df_const.columns.values)
            thread_values = [ prj_value for prj_value in range(0,df_input_rows,1)]

            def dataframes(prj_value):

                x = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
                for per_value in range(int(df_input.iat[prj_value,df_input_header.index('Period Number Start')]),int(df_input.iat[prj_value,df_input_header.index('Period Number End')]+1),1):

                    for prod_value in range(0,df_staffmix_rows,1):

                        if df_staffmix.iat[prod_value,df_staffmix_header.index('Product')] == df_input.iat[prj_value,df_input_header.index('Package')]:

                            account_desc = (df_input.iat[prj_value,df_input_header.index('Project')]).split("_")[0]
                            proj_name = df_input.iat[prj_value,df_input_header.index('Project')]
                            Probability = df_input.iat[prj_value,df_input_header.index('Probability')]
                            Org = df_staffmix.iat[prod_value,df_staffmix_header.index('Org')]
                            level = df_staffmix.iat[prod_value,df_staffmix_header.index('Level')]
                            Period_row = (df_notes.index[df_notes['Running Sequence'] == per_value ].tolist())[-1]
                            Period = df_notes.iat[Period_row,df_notes_header.index('FY-Period')]
                            Product = df_input.iat[prj_value,df_input_header.index('Package')]
                            people_with_probability = round((df_input.iat[prj_value,df_input_header.index('People per Period')] * df_staffmix.iat[prod_value,df_staffmix_header.index('%Mix')] / 100 * Probability) / df_const.iat[0,df_const_header.index('Reduction probability %')],2)
                            People_with_100 = round((df_input.iat[prj_value,df_input_header.index('People per Period')] * df_staffmix.iat[prod_value,df_staffmix_header.index('%Mix')] / 100) ,2)

                            if people_with_probability >0:
                                x = x.append(pd.DataFrame({df_demand_forecast_header[0]:'Y',
                                            df_demand_forecast_header[1]:'Demand',df_demand_forecast_header[2]:'Mercury',
                                            df_demand_forecast_header[3]:'Mercury',df_demand_forecast_header[4]:account_desc,
                                            df_demand_forecast_header[5]:account_desc,df_demand_forecast_header[6]:proj_name,
                                            df_demand_forecast_header[7]:Probability/100,df_demand_forecast_header[8]:Period,
                                            df_demand_forecast_header[9]:Product,df_demand_forecast_header[10]:Org,df_demand_forecast_header[11]:level,
                                            df_demand_forecast_header[12]:people_with_probability,df_demand_forecast_header[13]:People_with_100},index=[0]),ignore_index= True)
                return x
                
            with concurrent.futures.ThreadPoolExecutor() as executor:
                results = executor.map(dataframes,thread_values)

            for df in results:
                df_demand_forecast = df_demand_forecast.append(df, ignore_index=True)
            
            df_demand_forecast.to_excel((filenames['Report Location'] + '/' + 'Mercury.xlsx'),index=False)
            return df_demand_forecast

        except ValueError:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"The header names in the configuration file are incorrect ")
    except:
        FileBrows_obj = FileBrows(None)
        FileBrows_obj.raiseerror(f"The header names either in Input Analysis or configuration File are incorrect ")


def supply():
    mult = 0
    df_supply = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
    df_supply_header = list(df_supply.columns.values)

    df_notes = data_frame('Notes')
    df_notes_header = list(df_notes.columns.values)


    df_staffmix = data_frame("StaffMix")   
    df_staffmix_row , _ = df_staffmix.shape
    df_staffmix_header = list(df_staffmix.columns.values)

    df_att_hir = data_frame("Attrition&HiringParams")

    from_Per_row_num = (df_notes.index[df_notes['FY-Period'] == filenames["from Period"]].tolist())[-1]
    to_Per_row_num = (df_notes.index[df_notes['FY-Period'] == filenames["to Period"]].tolist())[-1]
    
    from_period_num = df_notes.iat[from_Per_row_num,df_notes_header.index('Running Sequence')]
    to_period_num = df_notes.iat[to_Per_row_num,df_notes_header.index('Running Sequence')]

    for run_seq in range(int(from_period_num),int((to_period_num+1)),1):

        for dem_prod in range(0,df_staffmix_row,1):
            
            if run_seq == from_period_num:
                period = df_notes.iat[run_seq-1,df_notes_header.index('FY-Period')]
                product = df_staffmix.iat[dem_prod,df_staffmix_header.index('Product')]
                Org = df_staffmix.iat[dem_prod,df_staffmix_header.index('Org')]
                Level = df_staffmix.iat[dem_prod,df_staffmix_header.index('Level')]
                People_with_probability = df_staffmix.iat[dem_prod,df_staffmix_header.index('People with probability')]
                
                df_supply = df_supply.append(pd.DataFrame({df_supply_header[0]:'Y',df_supply_header[1]:'Supply',
                                        df_supply_header[2]:'Supply',df_supply_header[3]:'Supply',
                                        df_supply_header[4]:'',df_supply_header[5]:'',
                                        df_supply_header[6]:'',df_supply_header[7]:1,
                                        df_supply_header[8]:period,df_supply_header[9]:product,
                                        df_supply_header[10]:Org,df_supply_header[11]:Level,df_supply_header[12]:People_with_probability,
                                        df_supply_header[13]:People_with_probability},index=[0]),ignore_index= True)
            else :
                Org = df_staffmix.iat[dem_prod,df_staffmix_header.index('Org')]
                if Org == 'US':
                    Hiring = (df_staffmix.iat[dem_prod,df_staffmix_header.index('%Mix')] * int(filenames['us_hir'])) /  df_att_hir.iat[0,0]
                    Attrition = (df_staffmix.iat[dem_prod,df_staffmix_header.index('%Mix')] * int(filenames['us_attr'])) /  df_att_hir.iat[0,0]
                elif Org == 'USI':
                    Hiring = (df_staffmix.iat[dem_prod,df_staffmix_header.index('%Mix')] * int(filenames['usi_hir'])) /  df_att_hir.iat[0,2]
                    Attrition = (df_staffmix.iat[dem_prod,df_staffmix_header.index('%Mix')] * int(filenames['usi_attr'])) /  df_att_hir.iat[0,2]
                elif Org == 'USDC':
                    Hiring = (df_staffmix.iat[dem_prod,df_staffmix_header.index('%Mix')] * int(filenames['usdc_hir'])) /  df_att_hir.iat[0,1]
                    Attrition = (df_staffmix.iat[dem_prod,df_staffmix_header.index('%Mix')] * int(filenames['usdc_attr'])) /  df_att_hir.iat[0,1]
                
                
                period = df_notes.iat[run_seq-1,df_notes_header.index('FY-Period')]
                product = df_staffmix.iat[dem_prod,df_staffmix_header.index('Product')]
                Level = df_staffmix.iat[dem_prod,df_staffmix_header.index('Level')]
                if product == 'ORACLE':
                    People_with_probability = df_staffmix.iat[dem_prod,df_staffmix_header.index('People with probability')] + mult*( Hiring - Attrition)
                else:
                    People_with_probability = df_staffmix.iat[dem_prod,df_staffmix_header.index('People with probability')]

                df_supply = df_supply.append(pd.DataFrame({df_supply_header[0]:'Y',df_supply_header[1]:'Supply',
                                        df_supply_header[2]:'Supply',df_supply_header[3]:'Supply',
                                        df_supply_header[4]:'',df_supply_header[5]:'',
                                        df_supply_header[6]:'',df_supply_header[7]:1,
                                        df_supply_header[8]:period,df_supply_header[9]:product,
                                        df_supply_header[10]:Org,df_supply_header[11]:Level,df_supply_header[12]:People_with_probability,
                                        df_supply_header[13]:People_with_probability},index=[0]),ignore_index= True)
        mult += 1    
                
    mult = 0
    
    df_supply.to_excel((filenames['Report Location'] + '/' + 'Supply_1.xlsx'),index=False)
    return df_supply

def demand():
    df_staffit = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
    df_staffit_header = list(df_staffit.columns.values)
    df_demand = data_frame('Oracle')
    df_demand_rows,_ = df_demand.shape
    df_staffmix = data_frame('StaffMix')
    df_staffmix_header = list(df_staffmix.columns.values)
    df_notes_staffit = data_frame('Notes')
    df_notes_staffit_header = list(df_notes_staffit.columns.values)
    df_demand['Start Period'] = ''
    df_demand['End Period'] = ''
    df_demand['Prorated Probability'] = ''
    df_demand['Demand Location'] = ''
    df_demand['Project Name'] = ''
    df_demand['Product'] = ''
    df_demand_header = list(df_demand.columns.values)

    df_dem_static = data_frame('Demand_static')
    df_dem_static_header = list(df_dem_static.columns.values)


    for row in range(0,df_demand_rows,1):
        lgs_row = (df_dem_static.index[df_dem_static['Key'] == df_demand.iat[row,df_demand_header.index('Level Group Global Text (Role)')]].tolist())[-1]
        df_demand.iat[row,df_demand_header.index('Level Group Global Text (Role)')] = df_dem_static.iat[lgs_row,1]
        df_demand.iat[row,df_demand_header.index('Start Period')] = running_sequence(df_demand.iat[row,df_demand_header.index('Role Start Date')])
        df_demand.iat[row,df_demand_header.index('End Period')] = running_sequence(df_demand.iat[row,df_demand_header.index('Role End Date')])
        prop_row = (df_dem_static.index[df_dem_static['Key'] == df_demand.iat[row,df_demand_header.index('Probability To Close (Role)')]].tolist())[-1]
        df_demand.iat[row,df_demand_header.index('Prorated Probability')] = (df_demand.iat[row,df_demand_header.index('Remaining Positions')] * df_dem_static.iat[prop_row,1])
        dem_loc_row = (df_dem_static.index[df_dem_static['Key'] == df_demand.iat[row,df_demand_header.index('Practice (Role)')]].tolist())[-1]
        df_demand.iat[row,df_demand_header.index('Demand Location')] = df_dem_static.iat[dem_loc_row,1]
        if df_demand.iat[row,df_demand_header.index('Local Client Name (Engagement)')] == '-':
            df_demand.iat[row,df_demand_header.index('Local Client Name (Engagement)')] = df_demand.iat[row,df_demand_header.index('Request Name')]
        else:
            pass
        df_demand.iat[row,df_demand_header.index('Project Name')] = df_demand.iat[row,df_demand_header.index('Local Client Name (Engagement)')] + '_' + str(df_demand.iat[row,df_demand_header.index('Req. No.')]) + '_' + df_demand.iat[row,df_demand_header.index('Request Name')]
        
        if df_demand.iat[row,df_demand_header.index('Request Name')].__contains__('JDE'):
            df_demand.iat[row,df_demand_header.index('Product')] = 'JDE'
        elif (df_demand.iat[row,df_demand_header.index('Request Name')].__contains__('PeopleSoft')) or (df_demand.iat[row,df_demand_header.index('Request Name')].__contains__('PSFT')) :
            df_demand.iat[row,df_demand_header.index('Product')] = 'PS'
        elif df_demand.iat[row,df_demand_header.index('Request Name')].__contains__('SAP'):
            df_demand.iat[row,df_demand_header.index('Product')] = 'SAP'
        else:
            if df_demand.iat[row,df_demand_header.index('Request Name')].__contains__('Oracle') or df_demand.iat[row,df_demand_header.index('Request Name')].__contains__(''):
                df_demand.iat[row,df_demand_header.index('Product')] = 'Oracle'
    
    unique_proj = df_demand['Local Client Name (Engagement)'].unique()

    thread_values = [ prj_itr for prj_itr in range(0,len(unique_proj),1)]
    
    def df_temp(prj_itr):
        df_temp = df_demand[df_demand['Local Client Name (Engagement)'] == unique_proj[prj_itr]]
        x = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
        df_temp_row,_ = df_temp.shape
        df_temp_header = list(df_temp.columns.values)
        minimum_per = int(df_temp['Start Period'].min())
        maximum_per = int(df_temp['End Period'].max())

        us_sum_aa , usi_sum_aa , usdc_sum_aa = 0,0,0
        us_sum_bta , usi_sum_bta , usdc_sum_bta = 0,0,0
        us_sum_c , usi_sum_c , usdc_sum_c = 0,0,0
        us_sum_sc , usi_sum_sc , usdc_sum_sc = 0,0,0
        us_sum_m , usi_sum_m , usdc_sum_m = 0,0,0
        us_sum_sm , usi_sum_sm , usdc_sum_sm = 0,0,0
        us_sum_ppd , usi_sum_ppd , usdc_sum_ppd = 0,0,0
        p_us_sum_aa , p_usi_sum_aa , p_usdc_sum_aa = 0,0,0
        p_us_sum_bta , p_usi_sum_bta , p_usdc_sum_bta = 0,0,0
        p_us_sum_c , p_usi_sum_c , p_usdc_sum_c = 0,0,0
        p_us_sum_sc , p_usi_sum_sc , p_usdc_sum_sc = 0,0,0
        p_us_sum_m , p_usi_sum_m , p_usdc_sum_m = 0,0,0
        p_us_sum_sm , p_usi_sum_sm , p_usdc_sum_sm = 0,0,0
        p_us_sum_ppd , p_usi_sum_ppd , p_usdc_sum_ppd = 0,0,0

        for per_value in range(minimum_per,maximum_per+1,1):

            for int_itr in range(0,df_temp_row,1):
                
                if per_value == df_temp.iat[int_itr,df_temp_header.index('Start Period')]: 
                    rem_pos = df_temp.iat[int_itr,df_temp_header.index('Remaining Positions')]
                    pr_prob = df_temp.iat[int_itr,df_temp_header.index('Prorated Probability')]
                    lev_row = (df_dem_static.index[df_dem_static['Value'] == df_temp.iat[int_itr,df_temp_header.index('Level Group Global Text (Role)')]].tolist())[-1]  
                    lev = df_dem_static.iat[lev_row,df_dem_static_header.index('Key')]
                    prct_row = (df_dem_static.index[df_dem_static['Key'] == df_temp.iat[int_itr,df_temp_header.index('Practice (Role)')]].tolist())[-1]  
                    prct = df_dem_static.iat[prct_row,df_dem_static_header.index('Value')]
                    if prct == 'US':
                        if lev == 'AA':
                            us_sum_aa += rem_pos
                            p_us_sum_aa += pr_prob
                        elif lev == 'BTA':
                            us_sum_bta += rem_pos
                            p_us_sum_bta += pr_prob
                        elif lev == 'C':
                            us_sum_c += rem_pos
                            p_us_sum_c += pr_prob
                        elif lev == 'SC':
                            us_sum_sc += rem_pos
                            p_us_sum_sc += pr_prob
                        elif lev == 'M':
                            us_sum_m += rem_pos
                            p_us_sum_m += pr_prob
                        elif lev == 'SM':
                            us_sum_sm += rem_pos
                            p_us_sum_sm += pr_prob
                        elif lev == 'PPD':
                            us_sum_ppd += rem_pos
                            p_us_sum_ppd += pr_prob

                    elif prct == 'USI':
                        if lev == 'AA':
                            usi_sum_aa += rem_pos
                            p_usi_sum_aa += pr_prob
                        elif lev == 'BTA':
                            usi_sum_bta += rem_pos
                            p_usi_sum_bta += pr_prob
                        elif lev == 'C':
                            usi_sum_c += rem_pos
                            p_usi_sum_c += pr_prob
                        elif lev == 'SC':
                            usi_sum_sc += rem_pos
                            p_usi_sum_sc += pr_prob
                        elif lev == 'M':
                            usi_sum_m += rem_pos
                            p_usi_sum_m += pr_prob
                        elif lev == 'SM':
                            usi_sum_sm += rem_pos
                            p_usi_sum_sm += pr_prob
                        elif lev == 'PPD':
                            usi_sum_ppd += rem_pos
                            p_usi_sum_ppd += pr_prob
                    elif prct == 'USDC':
                        if lev == 'AA':
                            usdc_sum_aa += rem_pos
                            p_usdc_sum_aa += pr_prob
                        elif lev == 'BTA':
                            usdc_sum_bta += rem_pos
                            p_usdc_sum_bta += pr_prob
                        elif lev == 'C':
                            usdc_sum_c += rem_pos
                            p_usdc_sum_c += pr_prob
                        elif lev == 'SC':
                            usdc_sum_sc += rem_pos
                            p_usdc_sum_sc += pr_prob
                        elif lev == 'M':
                            usdc_sum_m += rem_pos
                            p_usdc_sum_m += pr_prob
                        elif lev == 'SM':
                            usdc_sum_sm += rem_pos
                            p_usdc_sum_sm += pr_prob
                        elif lev == 'PPD':
                            usdc_sum_ppd += rem_pos
                            p_usdc_sum_ppd += pr_prob

                if per_value > minimum_per:
                    if per_value-1 == df_temp.iat[int_itr,df_temp_header.index('End Period')]:
                        rem_pos = df_temp.iat[int_itr,df_temp_header.index('Remaining Positions')]
                        pr_prob = df_temp.iat[int_itr,df_temp_header.index('Prorated Probability')]
                        lev_row = (df_dem_static.index[df_dem_static['Value'] == df_temp.iat[int_itr,df_temp_header.index('Level Group Global Text (Role)')]].tolist())[-1]  
                        lev = df_dem_static.iat[lev_row,df_dem_static_header.index('Key')]
                        prct_row = (df_dem_static.index[df_dem_static['Key'] == df_temp.iat[int_itr,df_temp_header.index('Practice (Role)')]].tolist())[-1]  
                        prct = df_dem_static.iat[prct_row,df_dem_static_header.index('Value')]

                        if prct == 'US':
                            if lev == 'AA':
                                us_sum_aa -= rem_pos
                                p_us_sum_aa -= pr_prob
                            elif lev == 'BTA':
                                us_sum_bta -= rem_pos
                                p_us_sum_bta -= pr_prob
                            elif lev == 'C':
                                us_sum_c -= rem_pos
                                p_us_sum_c -= pr_prob
                            elif lev == 'SC':
                                us_sum_sc -= rem_pos
                                p_us_sum_sc -= pr_prob
                            elif lev == 'M':
                                us_sum_m -= rem_pos
                                p_us_sum_m -= pr_prob
                            elif lev == 'SM':
                                us_sum_sm -= rem_pos
                                p_us_sum_sm -= pr_prob
                            elif lev == 'PPD':
                                us_sum_ppd -= rem_pos
                                p_us_sum_ppd -= pr_prob
                        elif prct == 'USI':
                            if lev == 'AA':
                                usi_sum_aa -= rem_pos
                                p_usi_sum_aa -= pr_prob
                            elif lev == 'BTA':
                                usi_sum_bta -= rem_pos
                                p_usi_sum_bta -= pr_prob
                            elif lev == 'C':
                                usi_sum_c -= rem_pos
                                p_usi_sum_c -= pr_prob
                            elif lev == 'SC':
                                usi_sum_sc -= rem_pos
                                p_usi_sum_sc -= pr_prob
                            elif lev == 'M':
                                usi_sum_m -= rem_pos
                                p_usi_sum_m -= pr_prob
                            elif lev == 'SM':
                                usi_sum_sm -= rem_pos
                                p_usi_sum_sm -= pr_prob
                            elif lev == 'PPD':
                                usi_sum_ppd -= rem_pos
                                p_usi_sum_ppd -= pr_prob
                        elif prct == 'USDC':
                            if lev == 'AA':
                                usdc_sum_aa -= rem_pos
                                p_usdc_sum_aa -= pr_prob
                            elif lev == 'BTA':
                                usdc_sum_bta -= rem_pos
                                p_usdc_sum_bta -= pr_prob
                            elif lev == 'C':
                                usdc_sum_c -= rem_pos
                                p_usdc_sum_c -= pr_prob
                            elif lev == 'SC':
                                usdc_sum_sc -= rem_pos
                                p_usdc_sum_sc -= pr_prob
                            elif lev == 'M':
                                usdc_sum_m -= rem_pos
                                p_usdc_sum_m -= pr_prob
                            elif lev == 'SM':
                                usdc_sum_sm -= rem_pos
                                p_usdc_sum_sm -= pr_prob
                            elif lev == 'PPD':
                                usdc_sum_ppd -= rem_pos
                                p_usdc_sum_ppd -= pr_prob

            for prod_value in range(0,21,1):

                account_name = (df_temp.iat[0,df_temp_header.index('Project Name')]).split("_")[0]
                Client_Project_Detailed_Description = df_temp.iat[0,df_temp_header.index('Project Name')] 
                prob_row_num = (df_dem_static.index[df_dem_static['Key'] == df_temp.iat[0,df_temp_header.index('Probability To Close (Role)')]].tolist())[-1]  
                Probability = df_dem_static.iat[prob_row_num,df_dem_static_header.index('Value')]
                fy_per_row = (df_notes_staffit.index[df_notes_staffit['Running Sequence'] == per_value].tolist())[-1]
                Product = df_temp.iat[0,df_temp_header.index('Product')]
                Period = df_notes_staffit.iat[fy_per_row,df_notes_staffit_header.index('FY-Period')]
                Org = df_staffmix.iat[prod_value,df_staffmix_header.index('Org')]
                level = df_staffmix.iat[prod_value,df_staffmix_header.index('Level')]
                if Org == 'US':
                    if level == 'AA':
                        people_with_probability_100 = us_sum_aa
                        people_with_probability = p_us_sum_aa
                    elif level == 'BTA':
                        people_with_probability_100 = us_sum_bta
                        people_with_probability = p_us_sum_bta
                    elif level == 'C':
                        people_with_probability_100 = us_sum_c
                        people_with_probability = p_us_sum_c
                    elif level == 'SC':
                        people_with_probability_100 = us_sum_sc
                        people_with_probability = p_us_sum_sc
                    elif level == 'M':
                        people_with_probability_100 = us_sum_m
                        people_with_probability = p_us_sum_m
                    elif level == 'SM':
                        people_with_probability_100 = us_sum_sm
                        people_with_probability = p_us_sum_sm
                    elif level == 'PPMD' or level == 'PPD':
                        people_with_probability_100 = us_sum_ppd
                        people_with_probability = p_us_sum_ppd
                elif Org == 'USI':
                    if level == 'AA':
                        people_with_probability_100 = usi_sum_aa
                        people_with_probability = p_usi_sum_aa
                    elif level == 'BTA':
                        people_with_probability_100 = usi_sum_bta
                        people_with_probability = p_usi_sum_bta
                    elif level == 'C':
                        people_with_probability_100 = usi_sum_c
                        people_with_probability = p_usi_sum_c
                    elif level == 'SC':
                        people_with_probability_100 = usi_sum_sc
                        people_with_probability = p_usi_sum_sc
                    elif level == 'M':
                        people_with_probability_100 = usi_sum_m
                        people_with_probability = p_usi_sum_m
                    elif level == 'SM':
                        people_with_probability_100 = usi_sum_sm
                        people_with_probability = p_usi_sum_sm
                    elif level == 'PPMD' or level == 'PPD':
                        people_with_probability_100 = usi_sum_ppd
                        people_with_probability = p_usi_sum_ppd
                elif Org == 'USDC':
                    if level == 'AA':
                        people_with_probability_100 = usdc_sum_aa
                        people_with_probability = p_usdc_sum_aa
                    elif level == 'BTA':
                        people_with_probability_100 = usdc_sum_bta
                        people_with_probability = p_usdc_sum_bta
                    elif level == 'C':
                        people_with_probability_100 = usdc_sum_c
                        people_with_probability = p_usdc_sum_c
                    elif level == 'SC':
                        people_with_probability_100 = usdc_sum_sc
                        people_with_probability = p_usdc_sum_sc
                    elif level == 'M':
                        people_with_probability_100 = usdc_sum_m
                        people_with_probability = p_usdc_sum_m
                    elif level == 'SM':
                        people_with_probability_100 = usdc_sum_sm
                        people_with_probability = p_usdc_sum_sm
                    elif level == 'PPMD' or level == 'PPD':
                        people_with_probability_100 = usdc_sum_ppd
                        people_with_probability = p_usdc_sum_ppd

                if people_with_probability > 0:
                    if Probability < 1:
                        x = x.append(pd.DataFrame({df_staffit_header[0]:'N',df_staffit_header[1]:'Demand',df_staffit_header[2]:'Staffit',
                                        df_staffit_header[3]:'Demand',df_staffit_header[4]:account_name,df_staffit_header[5]:account_name,df_staffit_header[6]:Client_Project_Detailed_Description,
                                        df_staffit_header[7]:Probability,df_staffit_header[8]:Period,df_staffit_header[9]:Product,df_staffit_header[10]:Org,df_staffit_header[11]:level,
                                        df_staffit_header[12]:people_with_probability,df_staffit_header[13]:people_with_probability_100},index=[0]),ignore_index= True)
                    else:
                        x = x.append(pd.DataFrame({df_staffit_header[0]:'Y',df_staffit_header[1]:'Demand',df_staffit_header[2]:'Staffit',
                                        df_staffit_header[3]:'Demand',df_staffit_header[4]:account_name,df_staffit_header[5]:account_name,df_staffit_header[6]:Client_Project_Detailed_Description,
                                        df_staffit_header[7]:Probability,df_staffit_header[8]:Period,df_staffit_header[9]:Product,df_staffit_header[10]:Org,df_staffit_header[11]:level,
                                        df_staffit_header[12]:people_with_probability,df_staffit_header[13]:people_with_probability_100},index=[0]),ignore_index= True)
    
        return x

    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = executor.map(df_temp,thread_values)
    for df in results:
        df_staffit = df_staffit.append(df, ignore_index=True)

    df_staffit.to_excel((filenames['Report Location'] + '/' + 'Demand_1.xlsx'),index=False)
    return df_staffit

def availability():
    df_staffit_avl = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
    df_staffit_avl_header = list(df_staffit_avl.columns.values)

    df_avl= pd.DataFrame(columns=['Local Client Name (Engagement)','Business Line Text (Role)','Role Start Date','Level Group Global Text (Role)','Probability To Close (Role)','Practice (Role)','Number of Consultants Needed','Role End Date','Start Period','End Period','Prorated Probability','Demand Location','Project Name','Package'])
    df_avl_header = list(df_avl.columns.values)

    df_demand = data_frame('Source_Data')
    df_demand_rows,_ = df_demand.shape
    df_demand_header = list(df_demand.columns.values)

    df_staffmix = data_frame('StaffMix')
    df_staffmix_header = list(df_staffmix.columns.values)
    df_notes_staffit = data_frame('Notes')
    df_notes_staffit_header = list(df_notes_staffit.columns.values)

    df_dem_static = data_frame('Demand_static')
    df_dem_static_header = list(df_dem_static.columns.values)

    for row in range(0,df_demand_rows,1):
        lgs_row = (df_dem_static.index[df_dem_static['Key'] == df_demand.iat[row,df_demand_header.index('Level (Employee) (Current)')]].tolist())[-1]
        start_period = running_sequence(df_demand.iat[row,df_demand_header.index('Start Date')])
        end_period = running_sequence(df_demand.iat[row,df_demand_header.index('End Date')])
        prop_row = (df_dem_static.index[df_dem_static['Key'] == 'Sold'].tolist())[-1]
        dem_loc_row = (df_dem_static.index[df_dem_static['Key'] == df_demand.iat[row,df_demand_header.index('Practice (Employee) (Current)')]].tolist())[-1]
        if len(df_demand.iat[row,df_demand_header.index('Business Line (Employee) (Current)')]) == 0:
            pack = 'Oracle'
        else:
            pack = df_demand.iat[row,df_demand_header.index('Business Line (Employee) (Current)')]
        
        df_avl = df_avl.append(pd.DataFrame({df_avl_header[0]:df_demand.iat[row,df_demand_header.index('Assignment')],
                                        df_avl_header[1]:df_demand.iat[row,df_demand_header.index('Business Line (Employee) (Current)')],
                                        df_avl_header[2]:df_demand.iat[row,df_demand_header.index('Start Date')],
                                        df_avl_header[3]:df_dem_static.iat[lgs_row,1],
                                        df_avl_header[4]:'Sold',
                                        df_avl_header[5]:df_demand.iat[row,df_demand_header.index('Practice (Employee) (Current)')],
                                        df_avl_header[6]:1,
                                        df_avl_header[7]:df_demand.iat[row,df_demand_header.index('End Date')],
                                        df_avl_header[8]:start_period,df_avl_header[9]:end_period,
                                        df_avl_header[10]:( 1 * df_dem_static.iat[prop_row,1]),
                                        df_avl_header[11]:df_dem_static.iat[dem_loc_row,1],
                                        df_avl_header[12]:df_demand.iat[row,df_demand_header.index('Assignment')],
                                        df_avl_header[13]:pack},index=[0]),ignore_index= True)

    unique_proj = df_avl['Local Client Name (Engagement)'].unique()
    thread_values = [ prj_itr for prj_itr in range(0,len(unique_proj),1)]
    
    def df_temp(prj_itr):
        df_temp = df_avl[df_avl['Local Client Name (Engagement)'] == unique_proj[prj_itr]]
        x = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
        df_temp_row,_ = df_temp.shape
        df_temp_header = list(df_temp.columns.values)
        minimum_per = int(df_temp['Start Period'].min())
        maximum_per = int(df_temp['End Period'].max())

        us_sum_aa , usi_sum_aa , usdc_sum_aa = 0,0,0
        us_sum_bta , usi_sum_bta , usdc_sum_bta = 0,0,0
        us_sum_c , usi_sum_c , usdc_sum_c = 0,0,0
        us_sum_sc , usi_sum_sc , usdc_sum_sc = 0,0,0
        us_sum_m , usi_sum_m , usdc_sum_m = 0,0,0
        us_sum_sm , usi_sum_sm , usdc_sum_sm = 0,0,0
        us_sum_ppd , usi_sum_ppd , usdc_sum_ppd = 0,0,0

        for per_value in range(minimum_per,maximum_per+1,1):

            for int_itr in range(0,df_temp_row,1):
                
                if per_value == df_temp.iat[int_itr,df_temp_header.index('Start Period')]: 
                    rem_pos = df_temp.iat[int_itr,df_temp_header.index('Number of Consultants Needed')]
                    lev_row = (df_dem_static.index[df_dem_static['Value'] == df_temp.iat[int_itr,df_temp_header.index('Level Group Global Text (Role)')]].tolist())[-1]  
                    lev = df_dem_static.iat[lev_row,df_dem_static_header.index('Key')]
                    prct_row = (df_dem_static.index[df_dem_static['Key'] == df_temp.iat[int_itr,df_temp_header.index('Practice (Role)')]].tolist())[-1]  
                    prct = df_dem_static.iat[prct_row,df_dem_static_header.index('Value')]
                    if prct == 'US':
                        if lev == 'AA':
                            us_sum_aa += rem_pos
                        elif lev == 'BTA':
                            us_sum_bta += rem_pos
                        elif lev == 'C':
                            us_sum_c += rem_pos
                        elif lev == 'SC':
                            us_sum_sc += rem_pos
                        elif lev == 'M':
                            us_sum_m += rem_pos
                        elif lev == 'SM':
                            us_sum_sm += rem_pos
                        elif lev == 'PPD':
                            us_sum_ppd += rem_pos

                    elif prct == 'USI':
                        if lev == 'AA':
                            usi_sum_aa += rem_pos
                        elif lev == 'BTA':
                            usi_sum_bta += rem_pos
                        elif lev == 'C':
                            usi_sum_c += rem_pos
                        elif lev == 'SC':
                            usi_sum_sc += rem_pos
                        elif lev == 'M':
                            usi_sum_m += rem_pos
                        elif lev == 'SM':
                            usi_sum_sm += rem_pos
                        elif lev == 'PPD':
                            usi_sum_ppd += rem_pos
                    elif prct == 'USDC':
                        if lev == 'AA':
                            usdc_sum_aa += rem_pos
                        elif lev == 'BTA':
                            usdc_sum_bta += rem_pos
                        elif lev == 'C':
                            usdc_sum_c += rem_pos
                        elif lev == 'SC':
                            usdc_sum_sc += rem_pos
                        elif lev == 'M':
                            usdc_sum_m += rem_pos
                        elif lev == 'SM':
                            usdc_sum_sm += rem_pos
                        elif lev == 'PPD':
                            usdc_sum_ppd += rem_pos

                if per_value > minimum_per:
                    if per_value-1 == df_temp.iat[int_itr,df_temp_header.index('End Period')]:
                        rem_pos = df_temp.iat[int_itr,df_temp_header.index('Number of Consultants Needed')]
                        lev_row = (df_dem_static.index[df_dem_static['Value'] == df_temp.iat[int_itr,df_temp_header.index('Level Group Global Text (Role)')]].tolist())[-1]  
                        lev = df_dem_static.iat[lev_row,df_dem_static_header.index('Key')]
                        prct_row = (df_dem_static.index[df_dem_static['Key'] == df_temp.iat[int_itr,df_temp_header.index('Practice (Role)')]].tolist())[-1]  
                        prct = df_dem_static.iat[prct_row,df_dem_static_header.index('Value')]

                        if prct == 'US':
                            if lev == 'AA':
                                us_sum_aa -= rem_pos
                            elif lev == 'BTA':
                                us_sum_bta -= rem_pos
                            elif lev == 'C':
                                us_sum_c -= rem_pos
                            elif lev == 'SC':
                                us_sum_sc -= rem_pos
                            elif lev == 'M':
                                us_sum_m -= rem_pos
                            elif lev == 'SM':
                                us_sum_sm -= rem_pos
                            elif lev == 'PPD':
                                us_sum_ppd -= rem_pos
                        elif prct == 'USI':
                            if lev == 'AA':
                                usi_sum_aa -= rem_pos
                            elif lev == 'BTA':
                                usi_sum_bta -= rem_pos
                            elif lev == 'C':
                                usi_sum_c -= rem_pos
                            elif lev == 'SC':
                                usi_sum_sc -= rem_pos
                            elif lev == 'M':
                                usi_sum_m -= rem_pos
                            elif lev == 'SM':
                                usi_sum_sm -= rem_pos
                            elif lev == 'PPD':
                                usi_sum_ppd -= rem_pos
                        elif prct == 'USDC':
                            if lev == 'AA':
                                usdc_sum_aa -= rem_pos
                            elif lev == 'BTA':
                                usdc_sum_bta -= rem_pos
                            elif lev == 'C':
                                usdc_sum_c -= rem_pos
                            elif lev == 'SC':
                                usdc_sum_sc -= rem_pos
                            elif lev == 'M':
                                usdc_sum_m -= rem_pos
                            elif lev == 'SM':
                                usdc_sum_sm -= rem_pos
                            elif lev == 'PPD':
                                usdc_sum_ppd -= rem_pos

            for prod_value in range(0,21,1):
                
                account_name = df_temp.iat[0,df_temp_header.index('Project Name')]
                prob_row_num = (df_dem_static.index[df_dem_static['Key'] == df_temp.iat[0,df_temp_header.index('Probability To Close (Role)')]].tolist())[-1]  
                Probability = df_dem_static.iat[prob_row_num,df_dem_static_header.index('Value')]
                fy_per_row = (df_notes_staffit.index[df_notes_staffit['Running Sequence'] == per_value].tolist())[-1]
                Period = df_notes_staffit.iat[fy_per_row,df_notes_staffit_header.index('FY-Period')]
                Product = df_temp.iat[0,df_temp_header.index('Package')]
                Org = df_staffmix.iat[prod_value,df_staffmix_header.index('Org')]
                level = df_staffmix.iat[prod_value,df_staffmix_header.index('Level')]
                if Org == 'US':
                    if level == 'AA':
                        people_with_probability_100 = us_sum_aa
                    elif level == 'BTA':
                        people_with_probability_100 = us_sum_bta
                    elif level == 'C':
                        people_with_probability_100 = us_sum_c
                    elif level == 'SC':
                        people_with_probability_100 = us_sum_sc
                    elif level == 'M':
                        people_with_probability_100 = us_sum_m
                    elif level == 'SM':
                        people_with_probability_100 = us_sum_sm
                    elif level == 'PPD':
                        people_with_probability_100 = us_sum_ppd
                elif Org == 'USI':
                    if level == 'AA':
                        people_with_probability_100 = usi_sum_aa
                    elif level == 'BTA':
                        people_with_probability_100 = usi_sum_bta
                    elif level == 'C':
                        people_with_probability_100 = usi_sum_c
                    elif level == 'SC':
                        people_with_probability_100 = usi_sum_sc
                    elif level == 'M':
                        people_with_probability_100 = usi_sum_m
                    elif level == 'SM':
                        people_with_probability_100 = usi_sum_sm
                    elif level == 'PPD':
                        people_with_probability_100 = usi_sum_ppd
                elif Org == 'USDC':
                    if level == 'AA':
                        people_with_probability_100 = usdc_sum_aa
                    elif level == 'BTA':
                        people_with_probability_100 = usdc_sum_bta
                    elif level == 'C':
                        people_with_probability_100 = usdc_sum_c
                    elif level == 'SC':
                        people_with_probability_100 = usdc_sum_sc
                    elif level == 'M':
                        people_with_probability_100 = usdc_sum_m
                    elif level == 'SM':
                        people_with_probability_100 = usdc_sum_sm
                    elif level == 'PPD':
                        people_with_probability_100 = usdc_sum_ppd

                if people_with_probability_100 >0:
                    x = x.append(pd.DataFrame({df_staffit_avl_header[0]:'Y',df_staffit_avl_header[1]:'Demand',df_staffit_avl_header[2]:'Staffit',
                                            df_staffit_avl_header[3]:'Availability',df_staffit_avl_header[4]:account_name,df_staffit_avl_header[5]:account_name,df_staffit_avl_header[6]:account_name,
                                            df_staffit_avl_header[7]:Probability,df_staffit_avl_header[8]:Period,df_staffit_avl_header[9]:Product,df_staffit_avl_header[10]:Org,df_staffit_avl_header[11]:level,
                                            df_staffit_avl_header[12]:people_with_probability_100,df_staffit_avl_header[13]:people_with_probability_100},index=[0]),ignore_index= True)

        return x
    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = executor.map(df_temp,thread_values)

    for df in results:
        df_staffit_avl = df_staffit_avl.append(df, ignore_index=True)   

    df_staffit_avl.to_excel((filenames['Report Location'] + '/' + 'Avl.xlsx'),index=False)
    return df_staffit_avl


def dem_sup(): 
    df_final = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
    thread_values = [1,2]
    def both(thread_value):
        if thread_value == 1:
            return demand()
        elif thread_value == 2:
            return availability()

    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = executor.map(both,thread_values)

    for df in results:
        df_final = df_final.append(df, ignore_index=True)

    return df_final

def handler(func):
    if func == 1:
        df_final = pd.DataFrame(columns=['Include Flag','Type','Source','Source Type','Account Name','Client_Project_Grouping_Description','Client_Project_Detailed_Description','Probability','Period','Product','Org','Level','People with probability','People with 100%'])
        try:
            thread_values = [1,2,3]
            def both(thread_value):
                if thread_value == 1:
                    return mercury_forcast_estimation()
                elif thread_value == 2:
                    return supply()
                elif thread_value == 3:
                    return dem_sup()

            with concurrent.futures.ThreadPoolExecutor() as executor:
                results = executor.map(both,thread_values)

            for df in results:
                df_final = df_final.append(df, ignore_index=True)
            
            df_final.to_excel((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data_1.xlsx'),index=False)
            df_final['Probability'] = df_final['Probability'].astype(float).map("{:.0%}".format)
            df_notes = data_frame('Notes')
            df_notes_header = list(df_notes.columns.values)
            from_Per_row_num = (df_notes.index[df_notes['FY-Period'] == filenames["from Period"]].tolist())[-1]
            to_Per_row_num = (df_notes.index[df_notes['FY-Period'] == filenames["to Period"]].tolist())[-1]
            fy_filter_list = [ df_notes.iat[prj_itr,df_notes_header.index('FY-Period')] for prj_itr in range(int(from_Per_row_num),int(to_Per_row_num+1),1)]
            df_final = df_final[(df_final['Period'].isin(fy_filter_list))]
            df_final_header = list(df_final.columns.values)
            df_final.to_excel((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data_2.xlsx'),index=False)

            #suppression
            df_temp = df_final[(df_final['Source Type']=='Demand')]
            df_temp = df_temp[df_temp['Client_Project_Detailed_Description'].str.contains('_MO-')]
            df = df_temp['Client_Project_Detailed_Description'].unique()
            df_mer_id = []
            for value in range(0,len(df),1):
                df_mer_id.append('MO-'+((df[value].split("_MO-")[-1]).split('_'))[0])
            df_mer = df_final[df_final['Source Type'] == 'Mercury']
            df_mer_row ,_ = df_mer.shape
            thread_values = [ row for row in range(0,len(df_mer_id),1)]
            thread_values_mer = [ row for row in range(0,df_mer_row,1)]

            def mer_id(value):
                def df_mer_func(row):
                    if (df_mer.iat[row,df_final_header.index('Client_Project_Detailed_Description')]).__contains__(df_mer_id[value]):
                        df_final.iat[row,df_final_header.index('Include Flag')] = 'N'

                with concurrent.futures.ThreadPoolExecutor() as executor:
                    results = executor.map(df_mer_func,thread_values_mer)
    
            with concurrent.futures.ThreadPoolExecutor() as executor:
                results = executor.map(mer_id,thread_values)
            
            df_final.to_excel((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data_3.xlsx'),index=False)
            book = openpyxl.load_workbook((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data.xlsx'))
            writer = pd.ExcelWriter((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data.xlsx'), engine = 'openpyxl')
            writer.book = book
            df_final.to_excel(writer, sheet_name = 'Master Data View',index = False)
            writer.save()
            writer.close()

            FileBrows_obj = FileBrows(None) 
            FileBrows_obj.processcomplete()
        except:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"File 'O2_Headcount_Analysis_Report_Master_Data.xlsx' is already present and open in the provided location please close the file or choose different location and try again")

    elif func == 2:
        try:
            result = mercury_forcast_estimation()
            result.to_excel((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data.xlsx'),sheet_name='Master Data View',index = False)
            FileBrows_obj = FileBrows(None) 
            FileBrows_obj.processcomplete()
        except:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"File 'O2_Headcount_Analysis_Report_Master_Data.xlsx' is already present and open in the provided location please close the file or choose different location and try again")
        
    elif func == 3:
        try:
            result = supply()
            result.to_excel((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data.xlsx'),sheet_name='Master Data View',index = False)
            FileBrows_obj = FileBrows(None) 
            FileBrows_obj.processcomplete()
        except:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"File 'O2_Headcount_Analysis_Report_Master_Data.xlsx' is already present and open in the provided location please close the file or choose different location and try again")
    
    elif func == 4:
        try:
            result = dem_sup()
            result.to_excel((filenames['Report Location'] + '/' + 'O2_Headcount_Analysis_Report_Master_Data.xlsx'),sheet_name='Master Data View',index = False)
            FileBrows_obj = FileBrows(None) 
            FileBrows_obj.processcomplete()
        except:
            FileBrows_obj = FileBrows(None)
            FileBrows_obj.raiseerror(f"File 'O2_Headcount_Analysis_Report_Master_Data.xlsx' is already present and open in the provided location please close the file or choose different location and try again")

    else:
        print('Error')


class FileSelector(tk.Tk):

    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs)

        self.title("O2 Headcount Analysis Report (Supply Demand and Availability)")
        self.geometry("1010x598")

        bundle_dir = getattr(sys,"_MEIPASS",os.path.abspath(os.path.dirname(__file__)))
        self.config_path = os.path.join(bundle_dir,"icon","D_icon.ico")
        self.iconbitmap(self.config_path)
        
        self.resizable(False, False)
        self.canvas = tk.Canvas(self,height = 550 ,width= 990 , bg= 'black')
        self.canvas.grid()

        frame = FileBrows(self.canvas,background="black" , relief = RAISED)
        frame.grid(padx = 10, pady = 10 , sticky="NESW")

        self.bind("<Return>",frame.confirmation)
        self.bind("KP_Enter",frame.confirmation)

class FileBrows(tk.Frame):

    def __init__(self,container,**kwargs):
        super().__init__(container,**kwargs)

        self.filename = tk.StringVar()
        self.folderloc = tk.StringVar()
        self.from_fy_option = tk.StringVar()
        self.to_fy_option = tk.StringVar()
        vcmd = (self.register(self.callback))
        self.list_values = self.drop_list()

        bundle_dir = getattr(sys,"_MEIPASS",os.path.abspath(os.path.dirname(__file__)))
        config_path = os.path.join(bundle_dir,"icon","Deloitte-logo.png")
        self.my_image = ImageTk.PhotoImage(Image.open(config_path))
        self.img_label = tk.Label(self,image=self.my_image,background="black")
        self.img_label.grid(row = 2 , column = 10,columnspan=20,sticky='NEWS')

        heading_font = font.Font(family='Times', weight='bold', size=15)
        sub_heading_font = font.Font(family='Times', weight='bold', size=10)

        # Mercury Input Label and ENtry Field
        mercury_label = tk.Label(self , text ="Mercury :" ,background="black",foreground='white')
        mercury_label['font'] = heading_font 
        blended_label = ttk.Label(self , text = "Blended Rate (USD)" , background="black",foreground='white')
        self.blended_label_input = ttk.Entry(self,width = 10 ,validate='key', validatecommand=(vcmd, '%P'))
        self.blended_label_input.insert('end',  '100')
        file_label = ttk.Label(self , text = "Mercury Roster " , background="black",foreground='white')
        self.file_input = ttk.Entry(self,width = 10 )

        #Suppy Input Label and entry field
        supply_label = tk.Label(self , text ="Supply :", background="black",foreground='white')
        supply_label['font'] = heading_font

        #Suppy Input Label and entry field
        period_label = tk.Label(self , text ="Period Values :", background="black",foreground='white')
        period_label['font'] = sub_heading_font
        from_label = tk.Label(self , text = "From Period Value" , background="black",foreground='white')
        self.from_input = ttk.OptionMenu(self,self.from_fy_option,*self.list_values)
        self.from_input.config(width = 9)
        to_label = tk.Label(self , text = "To Period Value", background="black",foreground='white')
        self.to_input = ttk.OptionMenu(self,self.to_fy_option,*self.list_values)
        self.to_input.config(width = 9)

        #Suppy Input Label and entry field
        us_label = tk.Label(self , text ="US :", background="black",foreground='white')
        us_label['font'] = sub_heading_font
        us_attr_label = tk.Label(self , text = "Attrition Value" , background="black",foreground='white')
        self.us_attr_label = ttk.Entry(self,width = 10,validate='all', validatecommand=(vcmd, '%P'))
        self.us_attr_label.insert('end',  '8')
        us_hir_label = tk.Label(self , text = "Hiring Value", background="black",foreground='white')
        self.us_hir_label = ttk.Entry(self,width = 10,validate='all', validatecommand=(vcmd, '%P'))
        self.us_hir_label.insert('end',  '0')

        space = tk.Label(self , text ="                                  ", background="black")

        #Suppy Input Label and entry field
        usi_label = tk.Label(self , text ="    USI :", background="black",foreground='white')
        usi_label['font'] = sub_heading_font
        usi_attr_label = tk.Label(self , text = "   Attrition Value" , background="black",foreground='white')
        self.usi_attr_label = ttk.Entry(self,width = 10,validate='all', validatecommand=(vcmd, '%P'))
        self.usi_attr_label.insert('end',  '0')
        usi_hir_label = tk.Label(self , text = "    Hiring Value", background="black",foreground='white')
        self.usi_hir_label = ttk.Entry(self,width = 10,validate='all', validatecommand=(vcmd, '%P'))
        self.usi_hir_label.insert('end',  '0')
        
        #Suppy Input Label and entry field
        usdc_label = tk.Label(self , text ="    USDC :", background="black",foreground='white')
        usdc_label['font'] = sub_heading_font
        usdc_attr_label = tk.Label(self , text = "  Attrition Value" , background="black",foreground='white')
        self.usdc_attr_label = ttk.Entry(self,width = 10,validate='all', validatecommand=(vcmd, '%P'))
        self.usdc_attr_label.insert('end',  '0')
        usdc_hir_label = tk.Label(self , text = "   Hiring Value", background="black",foreground='white')
        self.usdc_hir_label = ttk.Entry(self,width = 10,validate='all', validatecommand=(vcmd, '%P'))
        self.usdc_hir_label.insert('end',  '0')
        
        #Staffit Tnput Label and entry field
        staffit_label = tk.Label(self , text ="StaffIT :" , background="black",foreground='white')
        staffit_label['font'] = heading_font
        avl_label = tk.Label(self , text = "Availability File", background="black",foreground='white')
        self.avl_input = ttk.Entry(self,width = 10)
        dem_label = tk.Label(self , text = "Demand File" , background="black",foreground='white')
        self.dem_input = ttk.Entry(self,width = 10)

        # Output Location
        output_label = tk.Label(self , text ="HeadCount Analysis Report :" , background="black",foreground='white')
        output_label['font'] = heading_font
        folder_label = ttk.Label(self , text = "Report Location", background="black",foreground='white')
        self.folder_input = ttk.Entry(self,width = 10)

        # submit and browse button
        browse_button_1 = tk.Button(self , text = "Browse",bg='green',fg='white',command = lambda : self.filebrows('Mercury'))
        browse_button_2 = tk.Button(self , text = "Browse",bg='green',fg='white',command = lambda : self.filebrows('Avail'))
        browse_button_3 = tk.Button(self , text = "Browse",bg='green',fg='white',command = lambda : self.filebrows('Demand'))
        browse_button_4 = tk.Button(self , text = "Browse",bg='green',fg='white',command = lambda : self.folderbrows())
        submit_button = tk.Button(self , text = "Blow Up",bg='green',fg='white',command = lambda : self.confirmation())
        
        # Grid packing if the labels and buttons 
        mercury_label.grid(row = 2 , column = 0, sticky = 'W')
        supply_label.grid(row = 8 , column = 0, sticky = 'W')
        period_label.grid(row = 9 , column = 1, sticky = 'W')
        space.grid(row = 9 , column = 3, sticky = 'W')
        us_label.grid(row = 9 , column = 6, sticky = 'W')
        usi_label.grid(row = 9 , column = 8, sticky = 'W')
        usdc_label.grid(row = 9 , column = 10, sticky = 'W')
        staffit_label.grid(row = 13 , column = 0, sticky = 'W')
        output_label.grid(row = 18 , column = 0,columnspan = 3,  sticky = 'W')

        blended_label.grid(row = 3 , column = 1,  sticky='W' , padx = 5 , pady = 10)
        file_label.grid(row = 4 , column = 1,  sticky='W' , padx = 5 , pady = 10)
        from_label.grid(row = 10 , column = 1,  sticky='W' , padx = 5 , pady = 10)
        to_label.grid(row = 11 , column = 1,  sticky='W' , padx = 5 , pady = 10)
        us_attr_label.grid(row = 10 , column = 6,  sticky='W' , padx = 5 , pady = 10)
        us_hir_label.grid(row =11 , column = 6,  sticky='W' , padx = 5 , pady = 10)
        usi_attr_label.grid(row = 10 , column = 8,  sticky='W' , padx = 5 , pady = 10)
        usi_hir_label.grid(row =11 , column = 8,  sticky='W' , padx = 5 , pady = 10)
        usdc_attr_label.grid(row = 10 , column = 10,  sticky='W' , padx = 5 , pady = 10)
        usdc_hir_label.grid(row =11 , column = 10,  sticky='W' , padx = 5 ,  pady = 10)
        avl_label.grid(row = 14 , column = 1,  sticky='W' , padx = 5 , pady = 10)
        dem_label.grid(row = 15 , column = 1,  sticky='W' , padx = 5 , pady = 10)
        folder_label.grid(row = 19 , column = 1,  sticky='W' , padx = 5 , pady = 10)

        self.blended_label_input.grid(row = 3 , column = 2, sticky='EW',padx = 5 , pady = 10)
        self.file_input.grid(row = 4 , column = 2, columnspan = 10, sticky='EW',padx = 5 , pady = 10)
        self.from_input.grid(row = 10 , column = 2,sticky='W',padx = 5 , pady = 10)
        self.to_input.grid(row = 11 , column = 2,sticky='W',padx = 5 , pady = 10)
        self.us_attr_label.grid(row = 10 , column = 7,sticky='W',padx = 5 , pady = 10)
        self.us_hir_label.grid(row = 11 , column = 7,sticky='W',padx = 5 , pady = 10)
        self.usi_attr_label.grid(row = 10 , column = 9,sticky='W' , padx = 5, pady = 10)
        self.usi_hir_label.grid(row = 11 , column = 9,sticky='W',padx = 5 , pady = 10)
        self.usdc_attr_label.grid(row = 10 , column = 11,columnspan = 3,sticky='W',padx = 5 , pady = 10)
        self.usdc_hir_label.grid(row = 11 , column = 11,sticky='W',padx = 5 , pady = 10)
        self.avl_input.grid(row = 14 , column = 2,columnspan = 10,sticky='EW',padx = 5 , pady = 10)
        self.dem_input.grid(row = 15 , column = 2,columnspan = 10,sticky='EW',padx = 5 , pady = 10)
        self.folder_input.grid(row = 19 , column = 2,columnspan = 10,sticky='EW',padx = 5 , pady = 10)

        browse_button_1.grid(row = 4 , column = 13,sticky='E',padx = 5 , pady = 10)
        browse_button_2.grid(row = 14 , column = 13, sticky='E',padx = 5 , pady = 10)
        browse_button_3.grid(row = 15 , column = 13, sticky='E',padx = 5 , pady = 10)
        browse_button_4.grid(row = 19 , column = 13, sticky='E',padx = 5 , pady = 10)
        submit_button.grid(row = 22 , column = 4 , columnspan = 4,padx = 5 , pady =30)

    def filebrows(self,file):
        self.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
        if file == 'Mercury':
            self.file_input.delete(0,'end')
            self.file_input.insert(0,self.filename)
        elif file == 'Demand':
            self.dem_input.delete(0,'end')
            self.dem_input.insert(0,self.filename)
        elif file == 'Avail':
            self.avl_input.delete(0,'end')
            self.avl_input.insert(0,self.filename)
    
    def folderbrows(self):
        self.folderloc = filedialog.askdirectory(initialdir = "/")
        self.folder_input.delete(0,'end')
        self.folder_input.insert(0,self.folderloc)
    
    def confirmation(self):
        self.inputread()
        if (len(filenames["Mercury Roster"])>0 or (len(filenames["from Period"])>0 and len(filenames["to Period"])>0)) and len(filenames["Report Location"])>0:

            if (len(filenames["Mercury Roster"])>0 and len(filenames["Blended Rate"])>0 and filenames["from Period"]!='Select' and filenames["to Period"]!='Select' and len(filenames["us_attr"])>0 and len(filenames["us_hir"])>0 and len(filenames["usi_attr"])>0 and len(filenames["usi_hir"])>0 and len(filenames["usdc_attr"])>0 and len(filenames["usdc_hir"])>0 and len(filenames["demand file"])>0 and len(filenames["availability file"])>0):
                messagebox.showinfo('Request Submitted','Your request has been submitted please wait for the process to complete...')
                handler(1)
                

            elif len(filenames["Mercury Roster"])>0 and len(filenames["Blended Rate"])>0:
                messagebox.showinfo('Request Submitted','Your request has been submitted please wait for the process to complete...')
                handler(2)
                

            elif filenames["from Period"]!='Select' and filenames["to Period"]!='Select' and len(filenames["us_attr"])>0 and len(filenames["us_hir"])>0 and len(filenames["usi_attr"])>0 and len(filenames["usi_hir"])>0 and len(filenames["usdc_attr"])>0 and len(filenames["usdc_hir"])>0:
                messagebox.showinfo('Request Submitted','Your request has been submitted please wait for the process to complete...')
                handler(3)

            elif len(filenames["demand file"])>0 and len(filenames["availability file"])>0:
                messagebox.showinfo('Request Submitted','Your request has been submitted please wait for the process to complete...')
                handler(4)

            else:
                messagebox.showerror('Error','Please enter the values in all the input fields')

        else:
            messagebox.showerror('Error','Please enter the values in all the input fields')

    def processcomplete(self):
        response = messagebox.askyesno('Process Completed','Please check the files in provided output location \nDo you want to continue')
        if response == 0:
            root.destroy()
        else:
            pass
    
    def raiseerror(self,message):
        messagebox.showerror('Process Failed', message)

    def callback(self,P):
        if str.isdigit(P) or P == "": 
            return True
        else:
            return False

    def drop_list(self):
        self.df_LV = data_frame('Listvalues')
        self.df_LV_row,_ = self.df_LV.shape
        return [self.df_LV.iat[row,0]  for row in range(0,self.df_LV_row,1)]


    def inputread(self):
        filenames["Report Location"] = self.folder_input.get()
        filenames["Mercury Roster"] = self.file_input.get()
        filenames["from Period"] =self.from_fy_option.get()
        filenames["to Period"] = self.to_fy_option.get()
        filenames["Blended Rate"] =self.blended_label_input.get()
        filenames["us_attr"] =self.us_attr_label.get()
        filenames["us_hir"] =self.us_hir_label.get()
        filenames["usi_attr"] =self.usi_attr_label.get()
        filenames["usi_hir"] =self.usi_hir_label.get()
        filenames["usdc_attr"] =self.usdc_attr_label.get()
        filenames["usdc_hir"] =self.usdc_hir_label.get()
        filenames["demand file"] =self.dem_input.get()
        filenames["availability file"] =self.avl_input.get()

try:
    root = FileSelector()
    root.mainloop()
except:
    import traceback
    traceback.print_exc()
    input("Press Enter to end...")
