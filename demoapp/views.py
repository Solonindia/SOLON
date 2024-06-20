import warnings
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
import matplotlib
matplotlib.use('agg')  
import matplotlib.pyplot as plt
from django.shortcuts import render
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import math
from django.http import HttpResponse
from django.template.loader import get_template
from xhtml2pdf import pisa

def render_to_pdf(template_src, context_dict={}):
	template = get_template(template_src)
	html  = template.render(context_dict)
	result = BytesIO()
	pdf = pisa.pisaDocument(BytesIO(html.encode("UTF-8")), result)
	if not pdf.err:
		return HttpResponse(result.getvalue(), content_type='application/pdf')
	return None
def generate_plot(input_data, individual_numbers, individual_POAI):
    fig, ax1 = plt.subplots()
    color_light_blue = '#0000FF'  # Hexadecimal value for light blue
    color_black = 'black'

    ax1.set_xlabel('Month')
    ax1.set_ylabel('Energy in (kWh)', color=color_black)

    bar_width = 0.95  # Adjust the width of the bars
    bar_positions = np.arange(len(input_data)) * 2  # Add space between bars

    ax1.bar(bar_positions, individual_numbers, width=bar_width, color=color_light_blue)  # Using adjusted positions and width
    ax1.tick_params(axis='y', labelcolor=color_black)

    # Calculate the upper limit for y-axis with custom intervals
    max_individual_numbers = max(individual_numbers)
    upper_limit_numbers = ((max_individual_numbers + 200000) // 200000) * 200000

    # Set custom range for y-axis
    ax1.set_ylim(0, upper_limit_numbers)

    # Set custom y-axis intervals
    yticks_numbers = np.arange(0, upper_limit_numbers + 1, 200000)
    ax1.set_yticks(yticks_numbers)

    # Format y-axis tick labels
    ax1.set_yticklabels([f'{val:,}' for val in yticks_numbers])

    ax2 = ax1.twinx()
    ax2.set_ylabel('Insolation (kWh/m2)', color=color_black)
    ax2.plot(bar_positions, individual_POAI, color='red', marker='o', linestyle='-')
    ax2.tick_params(axis='y', labelcolor=color_black)

    # Set custom range for y-axis on both axes
    upper_limit_POAI = round(max(individual_POAI) + 50)
    ax2.set_ylim(0, upper_limit_POAI)

    # Set custom y-axis intervals for ax2
    yticks_POAI = np.arange(0, upper_limit_POAI + 1, 50)
    ax2.set_yticks(yticks_POAI)

    fig.tight_layout()  # To prevent the labels from overlapping
    plt.xticks(bar_positions, input_data, rotation=90, fontsize=12, ha='right') 
    print(plt.xticks)
    ax1.grid(False)  # Remove grid lines

    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    image_png = buffer.getvalue()
    buffer.close()
    graph = base64.b64encode(image_png)
    graph = graph.decode('utf-8')
    plt.close()    
    return graph
def calculate_total_time(time_list):
    # Convert each time to minutes and sum them up
    total_minutes = sum(time.hour * 60 + time.minute for time in time_list)
    # Convert total minutes to hours and remaining minutes
    total_hours, remaining_minutes = divmod(total_minutes, 60)
    # Return the total time as a tuple (hours, minutes)
    return total_hours, remaining_minutes

def Calculate_timings(new_list):
    # Calculate the total time from the new list
    total_hours = sum(item[0] for item in new_list)
    total_minutes = sum(item[1] for item in new_list)

    # Convert total hours and minutes to hours and remaining minutes
    total_hours += total_minutes // 60
    remaining_minutes = total_minutes % 60

    # Print the total time
    return str(total_hours)+':'+str(remaining_minutes)
    #print(f"Total time: {total_hours} hours and {remaining_minutes} minutes")

# Create your views here.
def index(request):
    if request.method == 'POST':
        input_data = []
        individual_numbers = []
        individual_Exp = []
        individual_Imp = []
        individual_Grid = []
        individual_POAI = []
        total_sum = 0
        total_Exp = 0
        total_Imp = 0
        total_Grid = 0
        total_POAI = 0
        
        input_data1 = []
        individual_SPR = []
        individual_PRA = []
        individual_PRC = []
        individual_PLF = []
        individual_PA = []
        individual_GA = []
        total_SPR = 0
        total_PRA = 0
        total_PRC = 0
        total_PLF = 0
        total_PA = 0
        total_GA = 0

        new_smb = []
        new_grid = []
        new_others = []
        new_inv = []
        new_tranformer = []
        new_string = []
        new_loss_string = []
        new_loss_smb = []
        new_loss_grid = []
        new_loss_others = []
        new_loss_inv = []
        new_loss_tranformer = []

        for i in range(1, 13):
            uploaded_file = request.FILES.get(f'file{i}')
            input_value = request.POST.get(f'input{i}', '')
            input_data.append(input_value)
            input_data1.append(input_value)

            if uploaded_file:
                df = pd.read_excel(uploaded_file, sheet_name='S_11.9', engine='openpyxl')
                unnamed_22_column = df['Unnamed: 22']
                unnamed_22_numeric = pd.to_numeric(unnamed_22_column, errors='coerce')
                total_sum_unnamed_22_mwh = unnamed_22_numeric.sum()
                total_sum_unnamed_22_kwh = round(total_sum_unnamed_22_mwh * 1000)
                individual_numbers.append(total_sum_unnamed_22_kwh)
                total_sum += total_sum_unnamed_22_kwh

                df2 = pd.read_excel(uploaded_file, sheet_name='Plant_Start',)
                index_of_MF = df2.columns.get_loc('MF')
                next_column_name = df2.columns[index_of_MF + 1]
                column_numeric = pd.to_numeric(df2[next_column_name], errors='coerce')
                column_numeric_positive = column_numeric[column_numeric > 0]
                sum_next_column = column_numeric_positive.sum()
                sum_Exp = round(sum_next_column)
                individual_Exp.append(sum_Exp)
                total_Exp = round(total_Exp + sum_next_column)

                df3 = pd.read_excel(uploaded_file, sheet_name='Plant_Start')
                index_of_MF = df3.columns.get_loc('MF')
                next_column_name = df3.columns[index_of_MF + 4]
                column_numeric = pd.to_numeric(df3[next_column_name], errors='coerce')
                column_numeric_positive = column_numeric[column_numeric > 0]
                sum_next_column = column_numeric_positive.sum()
                sum_Imp = round(sum_next_column)
                individual_Imp.append(sum_Imp)
                total_Imp = round(total_Imp + sum_next_column)

                df4 = pd.read_excel(uploaded_file, sheet_name='Plant_Start')
                index_of_MF = df3.columns.get_loc('MF')
                next_column = df3.columns[index_of_MF + 6]
                column_num = pd.to_numeric(df4[next_column], errors='coerce')
                sum_data = column_num.sum()
                sum_Grid = round(sum_data * 1000)
                individual_Grid.append(sum_Grid)
                total_Grid = total_Grid + sum_Grid

                df5 = pd.read_excel(uploaded_file, sheet_name='SUMMARY')
                cell_value = round(df5.iloc[5, 6], 2)
                individual_POAI.append(cell_value)
                total_POAI = round(total_POAI + cell_value, 2)

                df1 = pd.read_excel(uploaded_file, sheet_name='SUMMARY')
                #print(df1)
                unnamed_7_column = df1['Unnamed: 7']
                unnamed_5_column = df1['Unnamed: 5']
                unnamed_24_column = df1['Unnamed: 24']
                unnamed_7_column_0 = round(unnamed_7_column[0] * 100,2)
                unnamed_7_column_5 = round(unnamed_7_column[5] * 100 ,2)
                unnamed_5_column_5 = round(unnamed_5_column[5] * 100,2)
                unnamed_24_column_5 = round(unnamed_24_column[5] * 100,2)
                unnamed_26_column = df1['Unnamed: 26'] 
                unnamed_17_column = df1['Unnamed: 17']
                unnamed_26_column_5 = round(unnamed_26_column[5] * 100,2)
                unnamed_17_column_3 = round(unnamed_17_column[3] * 100,2)
                if math.isnan(unnamed_26_column_5):
                   individual_GA.append(unnamed_17_column_3)
                   total_GA = total_GA + unnamed_17_column_3
                else:
                   individual_GA.append(unnamed_26_column_5)
                   total_GA = total_GA + unnamed_26_column_5
            
                individual_SPR.append(unnamed_7_column_0)
                individual_PRA.append(unnamed_7_column_5)
                individual_PLF.append(unnamed_5_column_5)
                individual_PA.append(unnamed_24_column_5)


                total_SPR = total_SPR + unnamed_7_column_0
                total_PRA = total_PRA + unnamed_7_column_5
                total_PLF = total_PLF + unnamed_5_column_5
                total_PA = total_PA + unnamed_24_column_5
                if 'Unnamed: 30' in df1.columns and 'Unnamed: 31' in df1.columns:
                    unnamed_30_column = df1['Unnamed: 30'] 
                    unnamed_31_column = df1['Unnamed: 31']
                    unnamed_30_column_5 = round(unnamed_30_column[5] * 100, 2) if not pd.isnull(unnamed_30_column[5]) else None
                    unnamed_31_column_5 = round(unnamed_31_column[5] * 100, 2) if not pd.isnull(unnamed_31_column[5]) else None
                    if pd.isnull(unnamed_30_column_5) and unnamed_31_column_5 is None: 
                       unnamed_25_column = df1['Unnamed: 25']
                       unnamed_25_column_5 = round(unnamed_25_column[5] * 100, 2)
                       individual_PRC.append(unnamed_25_column_5)
                       total_PRC = total_PRC + unnamed_25_column_5
                    elif unnamed_31_column_5 is None:  
                       individual_PRC.append(unnamed_30_column_5)
                       total_PRC = total_PRC + unnamed_30_column_5
                    else:
                       individual_PRC.append(unnamed_31_column_5)
                       total_PRC = total_PRC + unnamed_31_column_5
                    avg_prc = round(total_PRC / 12, 2)
                else:
                    if 'Unnamed: 25' in df1.columns:
                       unnamed_25_column = df1['Unnamed: 25']
                       unnamed_25_column_5 = round(unnamed_25_column[5] * 100, 2)
                       individual_PRC.append(unnamed_25_column_5)
                       total_PRC = total_PRC + unnamed_25_column_5
                    else:
                       pass
                    avg_prc = round(total_PRC / 12, 2)

                df0 = pd.read_excel(uploaded_file, sheet_name='LOSS GEN')
                filtered_df = df0[df0['Unnamed: 2'] == 'BD_SMB']
                filtered_df1 = df0[df0['Unnamed: 2'] == 'BD_Grid']
                filtered_df2 = df0[df0['Unnamed: 2'] == 'BD_Others']
                filtered_df3 = df0[df0['Unnamed: 2'] == 'BD_INV']
                filtered_df4 = df0[df0['Unnamed: 2'] == 'BD_Transformer']
                filtered_df5 = df0[df0['Unnamed: 2'] == 'BD_String']

                unnamed_10_column = df0.columns.get_loc('Unnamed: 10')
                next_column_name = df0.columns[unnamed_10_column+1]
                next_column_smb = filtered_df[next_column_name].tolist()
                next_column_grid = filtered_df1[next_column_name].tolist()
                next_column_others = filtered_df2[next_column_name].tolist()
                next_column_inv = filtered_df3[next_column_name].tolist()
                next_column_tranformer = filtered_df4[next_column_name].tolist()
                next_column_string = filtered_df5[next_column_name].tolist()

                unnamed_15_column = df0.columns.get_loc('Unnamed: 15')
                next_column = df0.columns[unnamed_15_column+1]
                next_loss_string = filtered_df5[next_column].tolist()
                next_loss_smb = filtered_df[next_column].tolist()
                next_loss_grid = filtered_df1[next_column].tolist()
                next_loss_others = filtered_df2[next_column].tolist()
                next_loss_inv = filtered_df3[next_column].tolist()
                next_loss_tranformer = filtered_df4[next_column].tolist()

                total_loss_string = sum(next_loss_string)
                total_loss_smb = sum(next_loss_smb)
                total_loss_grid = sum(next_loss_grid)
                total_loss_others = sum(next_loss_others)
                total_loss_inv = sum(next_loss_inv)
                total_loss_tranformer = sum(next_loss_tranformer)
                new_loss_string.append(total_loss_string)
                new_loss_smb.append(total_loss_smb)
                new_loss_grid.append(total_loss_grid)
                new_loss_others.append(total_loss_others)
                new_loss_inv.append(total_loss_inv)
                new_loss_tranformer.append(total_loss_tranformer)

                total_time_inv = calculate_total_time(next_column_inv)
                total_time_smb = calculate_total_time(next_column_smb)
                total_time_others = calculate_total_time(next_column_others)
                total_time_grid = calculate_total_time(next_column_grid)
                total_time_tranformer = calculate_total_time(next_column_tranformer)
                total_time_string = calculate_total_time(next_column_string)
                
                new_inv.append(total_time_inv)
                new_smb.append(total_time_smb)
                new_others.append(total_time_others)
                new_grid.append(total_time_grid)
                new_tranformer.append(total_time_tranformer)
                new_string.append(total_time_string)

            avg_spr = round(total_SPR / 12,2)
            avg_pra = round(total_PRA / 12,2)
            avg_prc = round(total_PRC / 12,2)
            avg_plf = round(total_PLF / 12,2)
            avg_pa = round(total_PA / 12,2)
            avg_ga = round(total_GA / 12, 2)


        total_inv = Calculate_timings(new_inv)
        total_smb = Calculate_timings(new_smb)
        total_others = Calculate_timings(new_others)
        total_grid = Calculate_timings(new_grid)
        total_tranformer = Calculate_timings(new_tranformer)
        total_string = Calculate_timings(new_string)

        total_loss_string = round(sum(new_loss_string),2)
        total_loss_smb = round(sum(new_loss_smb),2)
        total_loss_grid = round(sum(new_loss_grid),2)
        total_loss_others = round(sum(new_loss_others),2)
        total_loss_inv = round(sum(new_loss_inv),2)
        total_loss_tranformer = round(sum(new_loss_tranformer),2)

        input_and_numbers = zip(input_data, individual_numbers, individual_Exp, individual_Imp, individual_Grid, individual_POAI)
        input_and_numbers1 = zip(input_data1,individual_SPR, individual_PRA, individual_PRC, individual_PLF, individual_PA, individual_GA)

        if uploaded_file:
            graph = generate_plot(input_data, individual_numbers, individual_POAI)
        else:
            graph = None
        context = {
            'input_and_numbers': input_and_numbers,
            'total_sum': total_sum,
            'total_Exp': total_Exp,
            'total_Imp': total_Imp,
            'total_Grid': total_Grid,
            'total_POAI': total_POAI,
            'graph': graph,
            'input_and_numbers1': input_and_numbers1,
            'avg_spr': avg_spr,
            'avg_pra': avg_pra,
            'avg_prc': avg_prc,
            'avg_plf': avg_plf,
            'avg_pa': avg_pa,
            'avg_ga': avg_ga,
            'total_loss_tranformer': total_loss_tranformer,
            'total_loss_inv': total_loss_inv,
            'total_loss_others': total_loss_others,
            'total_loss_grid': total_loss_grid,
            'total_loss_smb': total_loss_smb,
            'total_loss_string': total_loss_string,
            'total_inv': total_inv,
            'total_smb': total_smb,
            'total_others': total_others,
            'total_grid': total_grid,
            'total_tranformer': total_tranformer,
            'total_string': total_string,
        }
        template_path = 'demoapp/index.html'
        pdf = render_to_pdf(template_path,context)

        if pdf:
            download = request.GET.get("download")
            if download:
                response = HttpResponse(pdf, content_type='application/pdf')
                filename = "report.pdf"
                content = "attachment; filename='%s'" % filename
                response['Content-Disposition'] = content
                return response
            else:
                return HttpResponse(pdf, content_type='application/pdf')
        return HttpResponse("Failed to generate PDF", status=404)
    else:
      upload_range = range(1, 13)
      return render(request, 'demoapp/upload_form.html', {'upload_range': upload_range})
