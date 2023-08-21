import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

st.title('Envelopes Configurator')

Name = st.text_input('Name', 'Insert name of the unit')
dtmin = st.number_input('DT min [°C]')
dtmax = st.number_input('DT max [°C]')


df = pd.DataFrame(columns=['X [°C]','Y min [°C]','Y max [°C]'])
config = {
    'X' : st.column_config.NumberColumn('X', required=True, width='large'),
    'ymin' : st.column_config.NumberColumn('ymin', required=True, width='large'),
    'ymax' : st.column_config.NumberColumn('ymax', required=True, width='large')
}


# Watch out, result is the new data frame!
result = st.data_editor(df, column_config = config, num_rows='dynamic')


# Extract data
x_values = result['X [°C]'].astype(float)
ymin_values = result['Y min [°C]'].astype(float)
ymax_values = result['Y max [°C]'].astype(float)


# Sorting with respect of x
sorted_indices = sorted(range(len(x_values)), key=lambda k: x_values[k])
x_values = [x_values[i] for i in sorted_indices]
ymin_values = [ymin_values[i] for i in sorted_indices]
ymax_values = [ymax_values[i] for i in sorted_indices]


if len(x_values) > 0:

    # Display the graph
    fig, ax = plt.subplots()
    ax.fill_between(x_values, ymin_values, ymax_values, color='lightblue', alpha=0.5)
    ax.plot(x_values, ymin_values, marker= 'o', color='blue')
    ax.plot(x_values, ymax_values, marker = 'o', color='blue')
    ax.plot([x_values[0], x_values[0]], [ymin_values[0], ymax_values[0]], color='blue', marker = 'o')
    ax.plot([x_values[-1], x_values[-1]], [ymin_values[-1], ymax_values[-1]], color='blue', marker = 'o')
    ax.set_xlabel('ELWT [°C]')
    ax.set_ylabel('OAT [°C]')
    fig.suptitle(Name, fontsize=14)


    offset_x = 0.15 * len(x_values) - 0.1
    #offset_x = 0.05

    #for i_x1, i_y1 in zip(x_values[:-1], ymin_values[:-1]):
        #plt.text(i_x1 + offset_x, i_y1 + 2, '({}, {})'.format(i_x1, i_y1))
    #for i_x2, i_y2 in zip(x_values[:-1], ymax_values[:-1]):
        #plt.text(i_x2 + offset_x , i_y2 -4, '({}, {})'.format(i_x2, i_y2))

    
    
    # Display coordinates of the last point
    #if len(x_values) > 1:
        #plt.text(x_values[-1] - 5, ymax_values[-1] - 4, '({}, {})'.format(x_values[-1], ymax_values[-1]))
        #plt.text(x_values[-1] - 5, ymin_values[-1] + 2, '({}, {})'.format(x_values[-1], ymin_values[-1]))

    plt.grid()

    st.pyplot(fig)


# Check if a new row was added
if result is not None and 'new_row' in result:
    new_row = pd.DataFrame([result['new_row']])
    df = pd.concat([df, new_row], ignore_index=True)
    ax.relim()
    ax.autoscale_view()

    st.pyplot(fig)


# Excel function to write in the correct format the envelop, ready to be uploaded
# on dds
def write_excel_file(data1, data2, data3):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    sheet['D1'] = "Coefficients"
    sheet['A3'] = Name
    sheet['B3'] = dtmax
    sheet['C3'] = dtmin
    headers = ["Name", "dtmax", "dtmin", "X", "Y_min", "Y_Max"]
    
    for col_num, header in enumerate(headers, start=1):
        cell = sheet.cell(row=2, column=col_num)
        cell.value = header
    
    for index, value in enumerate(data1):
        cell = sheet.cell(row=index + 3, column=4)  # D3 corresponds to row 3, column 4
        cell.value = round(value,2)

    for index, value in enumerate(data2):
        cell = sheet.cell(row=index + 3, column=5)  # E3 corresponds to row 3, column 5
        cell.value = round(value,2)

    for index, value in enumerate(data3):
        cell = sheet.cell(row=index + 3, column=6)  # F3 corresponds to row 3, column 6
        cell.value = round(value,2)

    return workbook


# Call the excel function and generate the sheet when a button is pressed
if st.button("Generate Excel"):
    workbook = write_excel_file(x_values, ymin_values, ymax_values)
    file_name = Name + ".xlsx"
    with st.spinner("Saving the Excel file..."):
        workbook.save(file_name)
    st.success(f"Excel file '{file_name}' generated successfully!")
