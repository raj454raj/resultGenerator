import xlrd, os, sys

filename = sys.argv[1]
main_file = xlrd.open_workbook(filename)
filename = filename.split(".")[0]

# Complete data will be stored in this dictionary
# Structure =>
# {"FA1": {rol1: {"Name": "Name1", "scores": {"subject1": {"marks": 50, "grade": "A1"}...}...}...}...}
main_dict = {}
sheets = main_file.sheet_names()
subjects = None

# Populating main_dict by reading xlsx file
for sheet_name in sheets:

    main_dict[sheet_name] = {}
    sheet = main_file.sheet_by_name(sheet_name)
    # Assuming the actual records start from index 2
    number_of_records = sheet.nrows
    subjects = filter(lambda x: x != "", sheet.row_values(0))[2:]
    subjects = map(lambda x: x.lower(), subjects)

    for row in xrange(2, number_of_records):

        record = sheet.row_values(row)
        srollno = int(record[0])
        sname = record[1]
        main_dict[sheet_name][srollno] = {"name": sname, "scores": {}}
        scores = main_dict[sheet_name][srollno]["scores"]
        entities = record[2:]

        for i in xrange(len(entities)):
            entity_name = "marks"
            if i % 2 != 0:
                entity_name = "grades"

            if scores.has_key(subjects[i / 2]) is False:
                scores[subjects[i / 2]] = {}
            scores[subjects[i / 2]][entity_name] = entities[i]

def get_student_div(roll_no, data):
    """
        Get div rendering complete result of a student
    """

    sname = data["name"]
    main_div = """
    <div style="text-align: center">
            <h4><u>KENDRIYA VIDYALAYA NO.1, SECTOR-30, GANDHINAGAR - GUJARAT</u></h4>
            <h4><u>RESULT SHEET SESSION 2015-16</u></h4>
            <span style="padding-right: 3%;">ROLL NUMBER: """ + str(roll_no) + """</span>
            <span style="padding-right: 3%;">NAME OF STUDENT: """ + sname + """</span>
            <span>CLASS: """ + filename + """ </span>
        </div>
        <table cellpadding="5" cellspacing="0" border="1" style="position: absolute; left: 15%; text-align: center; border: 1px solid black; empty-cells: show;">
            <tr><td></td><td colspan="6">GRADES</td></tr>
            <tr>
                <td>Subject</td>"""

    for subject in subjects:
        main_div += "<td rowspan=\"2\">" + subject.upper() + "</td>"

    main_div += """
            </tr>
            <tr><td>Exam</td></tr>
            """

    for exam in ("FA1", "FA2", "SA1"):
        main_div += "<tr><td>" + exam + "</td>"
        for subject in subjects:
            main_div += "<td>" + data[exam][subject][0] + "</td>"
        main_div += "</tr>"
    main_div += """
        </table>
        <div style="margin-top: 30%;">
            <span style="margin-left: 10%;"> SIGN.OF CLASS TEACHERS </span>
            <span style="margin-left: 40%;"> PRINCIPAL </span>
        </div>
        """
    return main_div

# Dictionary which is iterable in roll numbers
# Format =>
# {"rol1": {"name": "Name1", "subject1": ["marks", "grades"], ...}...}
temp_dict = {}
exams = main_dict.keys()
temp_exam = exams[0]

# Initialize temp_dict completely
for roll_no in main_dict[temp_exam]:
    temp_dict[roll_no] = {"name": main_dict[temp_exam][roll_no]["name"]}
    for exam in exams:
        temp_dict[roll_no][exam] = {}
        for subject in subjects:
            temp_dict[roll_no][exam][subject] = []

# Populate temp_dict with the help of main_dict
for exam in main_dict:
    for roll_no in main_dict[exam]:
        sname = main_dict[exam][roll_no]["name"]
        scores = main_dict[exam][roll_no]["scores"]
        for subject in scores:
            temp_dict[roll_no][exam][subject] = [scores[subject]["grades"],
                                                 scores[subject]["marks"]]

# Make triplets for a single page
rolls = temp_dict.keys()
temp_rolls = []
t = []
for roll in rolls:
    t.append(roll)
    if len(t) == 3:
        temp_rolls.append(t)
        t = []

if len(t) not in (0, 3):
    temp_rolls.append(t)

count = 1

for roll_list in temp_rolls:

    complete_html = "<html><body>"
    # Append divs according to number of roll numbers left
    for i in xrange(len(roll_list)):
        complete_html += get_student_div(roll_list[i],
                                         temp_dict[roll_list[i]])
        complete_html += "<div style=\"background-color: black; height: 2px; width: 100%;\"></div>"
    f = open("htmls/" + str(count) + ".html", "w")
    f.write(complete_html)
    f.close()
    count += 1

# All files in htmls directory
files = None
for temp in os.walk("./htmls/"):
    files = temp[2]

# Convert all .html files to .pdf
for file_name in files:
    os.system("wkhtmltopdf ./htmls/" + file_name  + " ./pdfs/" + file_name.split(".")[0] + ".pdf")
