import openpyxl
import openpyxl.utils

def get_worksheet_data(file_name):
    excel_file = openpyxl.load_workbook(file_name)
    data_sheet = excel_file.active
    all_data = data_sheet.rows
    return all_data


def main():
    employment_data = get_worksheet_data("MAEmplyomentData.xlsx")
    for row in employment_data:
        town_cell = row[0]
        emp_town_name = town_cell.value
        emp_town_name= emp_town_name[1:]
        population_data = get_worksheet_data("massachusetts_population_1980-2010.xlsx")
        for pop_row in population_data:
            pop_town_name = pop_row[3].value
            if pop_town_name is None:
                continue

            if emp_town_name.lower() == pop_town_name.lower():
                labor_force = row[1].value
                population_column = openpyxl.utils.cell.column_index_from_string('I')-1
                town_population = pop_row[population_column].value
                participation_rate = labor_force/town_population * 100

                print(f"{pop_town_name} has {participation_rate}% in the labor force")








main()