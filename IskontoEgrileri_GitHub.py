# -*- coding: utf-8 -*-
"""
Created on Wed Feb 21 15:08:30 2024

@author: ky4642
"""

# Import necessary libraries
from scipy.optimize import minimize
import math
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
from datetime import date, datetime
import win32com.client as win32
import os
from openpyxl import load_workbook

# Define the folder path for saving results
klasor_yolu = r'C:\Users\ky4642\Desktop\İskonto Eğrisi\run_sonuclari'

# Create the folder if it doesn't exist
if not os.path.exists(klasor_yolu):
    os.makedirs(klasor_yolu)

# Function to write a DataFrame to an Excel file
def write_excel(filename, sheetname, dataframe):
    # Define maximum rows and columns Excel can handle
    max_rows, max_cols = 1048, 4
    dataframe = dataframe.dropna()  # Drop NaN values
    
    # If DataFrame exceeds size limits, truncate it
    if len(dataframe) > max_rows:
        dataframe = dataframe.iloc[:max_rows]
    if len(dataframe.columns) > max_cols:
        dataframe = dataframe.iloc[:, :max_cols]
    
    # Write DataFrame to Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        workBook = writer.book
        try:
            # If the sheet exists, remove it before adding new data
            workBook.remove(workBook[sheetname])
        except KeyError:
            # If sheet does not exist, this block will execute
            print("Worksheet does not exist")
        finally:
            # Write new data to the Excel file
            dataframe.to_excel(writer, sheet_name=sheetname, index=False)

# Loop through different currency values
for para_birimi in ['EUR', 'TRY', 'GBP', 'USD']:
    # File path for the data source
    file_path = r'C:\Users\ky4642\Desktop\İskonto Eğrisi\DataX_IR_CurveOIS_2024-06-28_2024-07-09_081101UTC.csv'
    
    # Read the first row to get the date
    tarih = pd.read_csv(file_path, nrows=1)
    # Skip the first row and read the data
    df = pd.read_csv(file_path, skiprows=1)
    
    # Filter data based on the currency and index
    if para_birimi == 'EUR':
        df = df[df['Index'] == 'EONIA']
    elif para_birimi == 'USD':
        df = df[df['Index'] == 'FED FUNDS']
    
    # Filter data by currency
    df_try = df[df.Currency == para_birimi].reset_index()
    
    # Extract yield and maturity data
    Getiri = df_try['RateMid']
    VadeYil_0 = df_try['Tenor']
    
    # Function to convert timespans to years
    def convert_to_years(timespan):
        unit = timespan[-1]  # Get the last character to determine the unit (W, M, Y)
        value = int(timespan[:-1])  # Extract the numeric value
    
        if unit == "W":
            return value / 52.0  # Convert weeks to years
        elif unit == "M":
            return value / 12.0  # Convert months to years
        elif unit == "Y":
            return value  # Already in years
    
    # Apply the conversion function to maturity data
    VadeYil = [convert_to_years(timespan) for timespan in VadeYil_0]
    
    # Define the Nelson-Siegel-Svensson (NSS) formula for yield curve fitting
    def nss_formula(params, vade):
        t, t2, b0, b1, b2, b3 = params
        term1 = b0
        term2 = b1 * ((1 - math.exp(-vade / t)) / (vade / t))
        term3 = b2 * (((1 - math.exp(-vade / t)) / (vade / t)) - math.exp(-vade / t))
        term4 = b3 * (((1 - math.exp(-vade / t2)) / (vade / t2)) - math.exp(-vade / t2))
        nss_getiri = term1 + term2 + term3 + term4
        return nss_getiri
    
    # Define the error function for optimizing the NSS parameters
    def nss_getiri(params, Vadeyıl, Getiri):
        nss_getiriler = [nss_formula(params, vade) for vade in Vadeyıl]
        error = sum(((getiri - nss_getiri) ** 2) * 100 for getiri, nss_getiri in zip(Getiri, nss_getiriler))
        return error
    
    # Function to estimate the yields using the optimized NSS parameters
    def nss_getiri_tahmin(params, Vadeyıl, Getiri):
        return [nss_formula(params, vade) for vade in Vadeyıl]
    
    # Define bounds for optimization parameters (optional)
    bounds = [(0, None), (0, None), (None, None), (None, None), (None, None), (None, None)]
    
    # Initialize error tracking and optimization methods
    nss_risksiz_faiz_error = float('inf')
    methods = ['SLSQP']  # Define optimization methods to be used
    initial_guesses = [
        [0.210048644, 0.012554395, 2.140471813, 0.987908453, 4.113464951, 61.55055],
        [0.783897, 5.354858, 0.000191, 30.6253, 60.40705, 61.55055],
        [1, 1, 1, 1, 1, 1],  # Completely different starting point
        np.random.rand(6)  # Random starting point
    ]
    
    # Create constraints dynamically for the optimization process
    def create_constraint(time):
        def constraint(x):
            return 0.05 - abs((nss_formula(x, VadeYil[time]) - Getiri[time]) / Getiri[time])
        return {'type': 'ineq', 'fun': constraint}
    
    def average_constraint(x):
        average_value = np.mean([abs((nss_formula(x, VadeYil[i]) - Getiri[i]) / Getiri[i]) for i in range(len(VadeYil))])
        return -average_value + 0.02
    
    # Combine individual constraints into a list
    constraints = [create_constraint(i) for i in range(len(VadeYil))] + [{'type': 'ineq', 'fun': average_constraint}]
    
    # Perform the optimization for each initial guess and method
    for initial_guess in initial_guesses:
        for method in methods:
            while nss_risksiz_faiz_error > 1:
                # Optimization step
                result = minimize(nss_getiri, initial_guess, args=(VadeYil, Getiri), method=method, bounds=bounds, constraints=constraints)
                
                if result.success and result.fun + 0.00001 < nss_risksiz_faiz_error:
                    nss_risksiz_faiz_error = result.fun
                    nss_risksiz_faiz_parametre = result.x
                    nss_getiri_risksiz_faiz = nss_getiri_tahmin(result.x, VadeYil, Getiri)
                    
                    # Update the initial guess for the next iteration
                    initial_guess = result.x
                else:
                    break
    
    # Print optimization results for risk-free rates
    print("Optimization Successful")
    print("Calculated Parameters: ", nss_risksiz_faiz_parametre)
    print("Total Error: ", nss_risksiz_faiz_error)
    print("Average Error:", np.mean(nss_risksiz_faiz_error))
    print(para_birimi, " Risk-Free Rate Values: ", nss_getiri_risksiz_faiz)
    
    file_path2 = r'C:\Users\ky4642\Desktop\İskonto Eğrisi\DataX_CDS_2024-06-28_EOD_2024-07-10_123835UTC.csv'
    df2 = pd.read_csv(file_path2, sep=';')

    # Time format and filtering
    date_part = tarih.columns[0].split('=')[1]
    year, month, day = date_part.split('-')
    Time = f'{day}.{month}.{year}'
    
    df2 = df2[df2['BusinessDateTimeUTC'].str.startswith(Time)]
    
    # Filter data based on currency and restructuring type
    if para_birimi == 'USD':
        EN = 'United States of America'
        df2 = df2[df2['RestructuringType'] == 'CR']
    elif para_birimi == 'EUR':
        EN = 'Germany'
        df2 = df2[df2['RestructuringType'] == 'CR']
    elif para_birimi == 'GBP':
        EN = 'United Kingdom of Great Britain and Northern Ireland'
        df2 = df2[df2['RestructuringType'] == 'CR']
    else: 
        EN = 'Turkey'
    
    df2_TRY = df2[df2['EntityName'] == EN].reset_index()
        # CDS Maturities and Spreads
    VadeYil2 = [1, 2, 3, 4, 5, 7, 10]
    df2_TRY = df2_TRY[df2_TRY['Tenor'].isin(VadeYil2)].reset_index()
    Getiri2_pre1 = df2_TRY['ParSpreadBid']
    Getiri2_pre2 = df2_TRY['ParSpreadAsk']
    
    # Define the number of latest values to be averaged based on the currency
    if para_birimi == 'TRY':
        n_latest_values = 12  
    else:
        n_latest_values = 4
    
    # Prepare for saving the results to Excel
    klasor3 = r'C:\Users\ky4642\Desktop\İskonto Eğrisi\run_sonuclari'
    file_path3 = os.path.join(klasor3, para_birimi + '_CDS.xlsx')
    
    # Create an empty Series for calculated CDS yields
    Getiri2 = pd.Series()
    
    # Iterate over each maturity and read the corresponding data from Excel
    for i in range(0, len(VadeYil2)):
        sheet_name = str(VadeYil2[i]) + 'Y'
        df3 = pd.read_excel(file_path3, sheet_name=sheet_name, engine='openpyxl')
        
        # Convert 'Date' to datetime and filter
        time_time = datetime.strptime(Time, "%d.%m.%Y").date()
        df3['Date'] = pd.to_datetime(df3['Date']).dt.date
        
        # Create a new DataFrame with the calculated spreads
        a = pd.DataFrame({'Date': time_time, 
                          'Bid': Getiri2_pre1[i], 
                          'Ask': Getiri2_pre2[i], 
                          'Spread': Getiri2_pre1[i] - Getiri2_pre2[i]}, 
                         index=[0])
        
        # Append the new data and sort by date
        df3 = pd.concat([df3, a]).sort_values(by='Date')
        
        # Write the updated data back to Excel
        write_excel(file_path3, sheet_name, df3)
        
        # Filter the data to keep only the latest values and calculate the average spread
        df3x = df3[df3['Date'] <= time_time]
        df3x = (df3x['Spread'].tail(n_latest_values).mean()) * -1
        Getiri2[i] = round(df3x, 3)
    
    # Reset the index of the CDS yields
    Getiri2 = Getiri2.reset_index()
    Getiri2 = Getiri2[0]
    
    # Set optimization bounds for CDS yields
    bounds2 = [(0.000000000000001, None), (0.00000000000001, None), (None, None), 
               (None, None), (None, None), (None, None)]
    
    # Define the initial guesses and optimization methods for CDS yields
    initial_guesses = [
        [0.783897, 5.354858, 0.000191, 30.6253, 60.40705, 61.55055],
        [1, 1, 1, 1, 1, 1],  # Different starting point
        np.random.rand(6)  # Random starting point
    ]
    
    methods = ['SLSQP']
    
    # Set the error threshold for optimization
    error_threshold = 20
    nss_likidite_hesaplama_error = 99999
    
    # Define time ranges for dynamic constraints
    test_vade = range(100, 10000, 100)
    
    # Create constraints for liquidity premium optimization
    def create_constraint(time):
        def constraint(x):
            return -abs((nss_formula(x, time/365) / 100) / (nss_formula(nss_risksiz_faiz_parametre, time/365) + nss_formula(x, time/365) / 200)) + 0.02
        return {'type': 'ineq', 'fun': constraint}
    
    def create_constraint2(time):
        def constraint(x):
            return nss_formula(x, time/365)
        return {'type': 'ineq', 'fun': constraint}
    
    # Combine constraints
    constraints1 = [create_constraint(i) for i in test_vade]
    constraints2 = [create_constraint2(i) for i in test_vade]
    constraints3 = constraints1 + constraints2
    
    # Perform the optimization for CDS liquidity premium
    for initial_guess in initial_guesses:
        for method in methods:
            while nss_likidite_hesaplama_error > error_threshold:
                result = minimize(nss_getiri, initial_guess, args=(VadeYil2, Getiri2), method=method, bounds=bounds2, constraints=constraints3)
                
                if result.fun + 0.000001 < nss_likidite_hesaplama_error and result.success:
                    nss_likidite_hesaplama_error = result.fun
                    nss_likidite_parametre = result.x
                    initial_guess = result.x
                else:
                    break
    
    # Print the optimization results for liquidity premium
    print("Target error threshold reached!", nss_likidite_hesaplama_error)
    print("Calculated Liquidity Premium Parameters: ", nss_likidite_parametre)
    
    # Estimate the liquidity premium using the optimized parameters
    nss_getiri_likidite_primi = nss_getiri_tahmin(nss_likidite_parametre, VadeYil2, Getiri2)
    print(para_birimi, " Liquidity Premium Values: ", nss_getiri_likidite_primi)
    
    # Create a DataFrame to store the final results
    columns = ['VADE_GUN', 'VADE', 'PARA_BIRIMI', 'RISKSIZ_GETIRI', 'LIKIDITE_PRIMI', 
               'DC_DUSUK_LIKIDITE', 'DC_ORTA_LIKIDITE', 'DC_YUKSEK_LIKIDITE', 'DONEM', 'DONEM_YIL_AY']
    iskonto = pd.DataFrame(columns=columns)
    
    data_to_append = []
    date_calculted = date(int(year), int(month), int(day))
    
    # Calculate discount curve values for each day
    for i in range(1, 10951):
        if i < 90:
            lp = 0
        else:
            lp = round(nss_formula(nss_likidite_parametre, i / 365), 6) / 100
        rg = round(nss_formula(nss_risksiz_faiz_parametre, i / 365), 6)
        
        data_to_append.append({
            'VADE_GUN': i,
            'VADE': round((i / 365), 6),
            'PARA_BIRIMI': para_birimi.upper(),
            'RISKSIZ_GETIRI': rg,
            'LIKIDITE_PRIMI': lp,
            'DC_DUSUK_LIKIDITE': rg + lp,
            'DC_ORTA_LIKIDITE': rg + lp * 0.5,
            'DC_YUKSEK_LIKIDITE': rg + lp * 0.25,
            'DONEM': date_calculted,
            'DONEM_YIL_AY': year + month
        })
    
    # Append the calculated data to the DataFrame
    iskonto = pd.concat([iskonto, pd.DataFrame(data_to_append)], ignore_index=True)
    
    # Round the necessary columns to 6 decimal places
    for col in ['DC_DUSUK_LIKIDITE', 'DC_ORTA_LIKIDITE', 'DC_YUKSEK_LIKIDITE', 'LIKIDITE_PRIMI', 'RISKSIZ_GETIRI', 'VADE']:
        iskonto[col] = iskonto[col].round(6)
    
    # Format the 'DONEM' column and other relevant columns
    iskonto['DONEM'] = pd.to_datetime(iskonto['DONEM']).dt.date
    iskonto['VADE_GUN'] = iskonto['VADE_GUN'].apply(lambda x: round(float(x)))
    iskonto['PARA_BIRIMI'] = iskonto['PARA_BIRIMI'].apply(lambda x: x.strip().title()).str.upper()
    iskonto['DONEM_YIL_AY'] = iskonto['DONEM_YIL_AY'].apply(lambda x: round(float(x)))
    
    # Save the results to Oracle database
    import pyodbc
    host = "xxxx"
    port = ""
    database_name = ""
    
    connection_string = f"DRIVER={{Oracle in OraClient12Home1}};DBQ={host}:{port}/{database_name};UID={user};PWD={xxxx};"
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    
    table_name = "as_ifrs.IFRS_ISKONTO_EGRISI_PRE"
    columns = ", ".join(iskonto.columns)
    placeholders = ", ".join(["?"] * len(iskonto.columns))
    sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
    
    data = [tuple(x) for x in iskonto.values]
    
    # Insert data row by row into the Oracle database
    for index, row in iskonto.iterrows():
        data = tuple(row)
        try:
            cursor.execute(sql, data)
            connection.commit()
        except Exception as e:
            print(f"Insert error for row {index}: {e}")
            continue
    
    # Close the database connection
    cursor.close()
    connection.close()
    
    # Save the results to Excel
    excel_dosya_yolu = os.path.join(klasor_yolu, para_birimi + '.xlsx')
    
    with pd.ExcelWriter(excel_dosya_yolu, engine='openpyxl') as writer:
        pd.DataFrame(VadeYil).to_excel(writer, sheet_name='risksiz_vade', index=False)
        pd.DataFrame(Getiri).to_excel(writer, sheet_name='risksiz_getiri', index=False)
        pd.DataFrame([nss_risksiz_faiz_error]).to_excel(writer, sheet_name='risksiz_faiz_error', index=False)
        pd.DataFrame(nss_risksiz_faiz_parametre).to_excel(writer, sheet_name='faiz_optimal_param', index=False)
        pd.DataFrame(nss_getiri_risksiz_faiz).to_excel(writer, sheet_name='faiz_hesaplanan', index=False)
        pd.DataFrame(VadeYil2).to_excel(writer, sheet_name='likidite_vade', index=False)
        pd.DataFrame(Getiri2).to_excel(writer, sheet_name='likidite_primi', index=False)
        pd.DataFrame([nss_likidite_hesaplama_error]).to_excel(writer, sheet_name='likidite_error', index=False)
        pd.DataFrame(nss_likidite_parametre).to_excel(writer, sheet_name='likidite_parametreleri', index=False)
        pd.DataFrame(nss_getiri_likidite_primi).to_excel(writer, sheet_name='likidite_hesaplanan', index=False)
    
    # Send the email with the results attached
    Time = str(Time)
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = "xxxx"  # Recipient's email address
        mail.Subject = 'Iskonto Egrileri'
        mail.Body = (f'{Time} için iskonto eğrileri hesaplanmıştır. \n'
                     'Ekte hesaplama detaylarını görebilirsiniz.\n'
                     'Onayınızdan sonra tabloya aktarılacaktır. Sonrasında Aspendos\'tan kontrol edebilirsiniz.\n'
                     'Otomatik maildir.')
        
        # Attach all files in the results folder to the email
        for dosya in os.listdir(klasor_yolu):
            dosya_yolu = os.path.join(klasor_yolu, dosya)
            if os.path.isfile(dosya_yolu):
                mail.Attachments.Add(dosya_yolu)
        
        mail.Send()
        print("E-posta başarıyla gönderildi.")
    except Exception as e:
        print(f"E-posta gönderilirken bir hata oluştu: {e}")
