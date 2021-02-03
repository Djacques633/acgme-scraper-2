#
#   Daniel Jacques
#   Heritage College of Osteopathic Medicine
#
import os
try:
    import xlrd
    import xlsxwriter
    import tkinter
    import tkinter.filedialog
    import inquirer
except ImportError:
    os.system('python -m pip install xlrd')
    os.system('python -m pip install bs4')
    os.system('python -m pip install tkinter')
    os.system('python -m pip install tkinter.filedialog')
    os.system('python -m pip install inquirer')


def format(program_numbers):
    return ', '.join([program_numbers] if type(
        program_numbers) == str else program_numbers)


def write_to_excel(user_data, index, sheet1, actual):
    sheet1.write('A' + str(index+2), user_data['last_name'])
    sheet1.write('B' + str(index+2), user_data['first_name'])

    sheet1.write('C' + str(index+2),
                 format(user_data['program_number_1'][0]))
    # sheet1.write('D' + str(index+2), user_data['program_number_1'][1])
    sheet1.write('D' + str(index+2), user_data['program_address_1'])
    sheet1.write('E' + str(index+2),
                 format(user_data['program_number_2'][0]))
    sheet1.write('F' + str(index+2), user_data['program_address_2'])

    # try:
    # sheet1.write('F' + str(index+2), user_data['program_number_2'][1])
    # except:
    # sheet1.write('F' + str(index+2), 'N/a')
    # sheet1.write('G' + str(index+2), actual)
    return


def find_institution_indices(code):
    indices = [i for i, j in enumerate(institution_codes) if j == code]
    return indices


def find_specialty_indices(indices, specialty):
    res = []
    for i in range(0, len(indices)):
        if specialty_name[int(indices[i])].lower() == specialty.lower():
            res.append(indices[i])
    return res


def find_city_indices(indices, city):
    res = []
    for i in range(0, len(indices)):
        if cities[int(indices[i])].lower() == city.lower():
            res.append(indices[i])
    return res


def useInquirer(program_map, organizations_map, match_facility):
    answer_choices = []
    for x in range(0, len(program_map)):
        answer_choices.append(
            program_map[x] + " | " + organizations_map[x])
    answer_choices.append("NONE | Answer not listed")
    questions = [
        inquirer.List('choice',
                      message="Multiple found for the match facility: " +
                      match_facility + ". Which is the correct? Current choice",
                      choices=answer_choices,
                      ),
    ]
    answers = inquirer.prompt(questions)
    return answers['choice'].split(' | ')


def getOrganizationName(organization_indices, name, index):
    organization_indices[name] = index
    return index


# FIRST FILE: OUTPUT
finput1 = tkinter.filedialog.askopenfilename()
f = xlsxwriter.Workbook(finput1)
sheet1 = f.add_worksheet()
sheet1.write('A1', 'Last Name')
sheet1.write('B1', 'First Name')
sheet1.write('C1', 'Program match 1')
sheet1.write('D1', 'Program match 2')

# SECOND FILE: ACGME PROGRAMS
finput = tkinter.filedialog.askopenfilename()
wb = xlrd.open_workbook(finput)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

specialty_name = []
program_numbers = []
organizations = []
institution_codes = []
cities = []
states = []
zip_codes = []
full_address = []
for i in range(0, sheet.nrows):
    specialty_name.append(sheet.cell_value(i, 0))   # Load ACGME program data
    program_numbers.append(sheet.cell_value(i, 1))
    organizations.append(sheet.cell_value(i, 2))
    institution_codes.append(sheet.cell_value(i, 3))
    cities.append(sheet.cell_value(i, 4))
    states.append(sheet.cell_value(i, 5))
    full_address.append(sheet.cell_value(i, 8))

# THIRD FILE: FROMKELSEY
finput = tkinter.filedialog.askopenfilename()
wb2 = xlrd.open_workbook(finput)

sheet = wb2.sheet_by_index(0)
sheet.cell_value(0, 0)

last_name = []
first_name = []
pre_clinical_campus = []
clinical_campus = []
specialty = []
program_code_institution_codes = []
match_facility = []
match_location = []

for i in range(2, sheet.nrows):  # Map FROMKELSEY into lists
    last_name.append(sheet.cell_value(i, 0))
    first_name.append(sheet.cell_value(i, 1))
    pre_clinical_campus.append(sheet.cell_value(i, 3))
    clinical_campus.append(sheet.cell_value(i, 4))
    specialty.append(sheet.cell_value(i, 6))
    program_code_institution_codes.append(sheet.cell_value(i, 7))
    match_facility.append(sheet.cell_value(i, 8))
    match_location.append(sheet.cell_value(i, 9))

for i in range(0, len(program_code_institution_codes)):
    # Initialize as empty first. These are arrays indices, not program numbers
    city_matches_1 = []
    city_matches_2 = []
    try:
        if '/' in program_code_institution_codes[i]:
            double_inst = program_code_institution_codes[i].split('/')
            double_spec = specialty[i].split('/')
            double_location = match_location[i].split('/')

            match1 = find_institution_indices(
                double_inst[0].split('-')[1])  # Narrow down by match
            match2 = find_institution_indices(double_inst[1].split('-')[1])

            specialty_matches_1 = find_specialty_indices(
                match1, double_spec[0])  # Then specialty
            specialty_matches_2 = find_specialty_indices(
                match2, double_spec[1])

            # Finally city.
            city_matches_1 = find_city_indices(
                specialty_matches_1, double_location[0].split(',')[0])
            city_matches_2 = find_city_indices(
                specialty_matches_2, double_location[1].split(',')[0])

        else:

            match1 = find_institution_indices(
                program_code_institution_codes[i].split('-')[1])
            specialty_matches_1 = find_specialty_indices(match1, specialty[i])
            city_matches_1 = find_city_indices(
                specialty_matches_1, match_location[i].split(',')[0])

        program2_map = list(map(lambda x: program_numbers[x], city_matches_2)) if len(
            city_matches_2) != 0 else ['N/a']

        program1_map = list(map(lambda x: program_numbers[x], city_matches_1))

        program_1_address = ""
        program_2_address = ""

        if (len(program1_map) > 1):

            organization_indices = {}

            organizations_map = list(
                map(lambda x: organizations[x], city_matches_1)
            )

            for j in range(0, len(city_matches_1)):
                organization_indices[organizations_map[j]] = city_matches_1[j]

            program1_map = useInquirer(
                program1_map, organizations_map, match_facility[i])

            program_1_address = full_address[organization_indices[program1_map[1]]
                                             ] if program1_map[0] != "NONE" else "NONE"

        else:
            program1_map = [program1_map, organizations[city_matches_1[0]]]
            program_1_address = full_address[city_matches_1[0]]

        if (len(program2_map) > 1):

            organizations_map = list(
                map(lambda x: organizations[x], city_matches_2))

            for j in range(0, len(city_matches_1)):
                organization_indices[organizations_map[j]] = city_matches_1[j]

            program2_map = useInquirer(
                program2_map, organizations_map, match_facility[i])

            program_2_address = full_address[organization_indices[program2_map[1]]
                                             ] if program2_map[0] != "NONE" else "NONE"

        elif (len(program2_map) == 1 and program2_map[0] != 'N/a'):
            program2_map = [program2_map, organizations[city_matches_2[0]]]
            program_2_address = full_address[city_matches_2[0]]
        userData = {
            "first_name": first_name[i],
            "last_name": last_name[i],
            "program_number_1": program1_map,
            "program_number_2": program2_map,
            "program_address_1": program_1_address,
            "program_address_2": program_2_address
        }

        write_to_excel(userData, i, sheet1, match_facility[i])

    except Exception as e:
        # print(e)
        # traceback.print_exc()
        # print("Error! Consider running debugger if this is unexpected")
        # print(program_code_institution_codes[i])
        sheet1.write('A' + str(i+2), last_name[i])
        sheet1.write('B' + str(i+2), first_name[i])
        # sheet1.write('C' + str(i+2), "UNMATCHED")
        continue

print("Exit success! Saving..")
try:
    f.close()
    print("Saved!")
except:
    print("File " + finput1 + " already open. Please close then try again!")
