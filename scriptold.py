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

temp_dict = {}
roll_nos = []

exams = main_dict.keys()
temp_exam = exams[0]

for roll_no in main_dict[temp_exam]:
    temp_dict[roll_no] = {"name": main_dict[temp_exam][roll_no]["name"]}
    for exam in exams:
        temp_dict[roll_no][exam] = {}
        for subject in subjects:
            temp_dict[roll_no][exam][subject] = []

for exam in main_dict:
    for roll_no in main_dict[exam]:
        sname = main_dict[exam][roll_no]["name"]
        scores = main_dict[exam][roll_no]["scores"]
        for subject in scores:
            temp_dict[roll_no][exam][subject] = [scores[subject]["grades"],
                                                 scores[subject]["marks"]]

for roll_no in temp_dict:
    sname = temp_dict[roll_no]["name"]
    complete_html = \
    """
<html>
    <head>
        <title>
        """ + sname + \
        """
        </title>
    </head>
    <body>
        <div style="text-align: center">
            <h4><u>KENDRIYA VIDYALAYA NO.1, SECTOR-30, GANDHINAGAR - GUJARAT</u></h4>
            <h4><u>RESULT SHEET SESSION 2015-16</u></h4>
            <span style="padding-right: 3%;">NAME OF STUDENT: """ + sname + """</span>
            <span>CLASS: """ + filename + """ </span>
        </div>
        <br /><br />
        <table cellpadding="5" cellspacing="0" border="1" style="position: absolute; left: 17%; text-align: center; border: 1px solid black; empty-cells: show;">
            <tr><td></td><td colspan="6">GRADES</td></tr>
            <tr>
                <td>Subject</td>"""

    for subject in subjects:
        complete_html += "<td rowspan=\"2\">" + subject.upper() + "</td>"

    complete_html += """
            </tr>
            <tr><td>Exam</td></tr>
            """

    for exam in ("FA1", "FA2", "SA1"):
        complete_html += "<tr><td>" + exam + "</td>"
        for subject in subjects:
            complete_html += "<td>" + temp_dict[roll_no][exam][subject][0] + "</td>"
        complete_html += "</tr>"
    complete_html += """
        </table>
        <br /><br /><br />
        <div style="margin-top: 35%;">
            <span style="margin-left: 10%;"> SIGN.OF CLASS TEACHERS </span>
            <span style="margin-left: 40%;"> PRINCIPAL </span>
        </div>
    </body>
</html>
    """
    f = open("htmls/" + str(roll_no) + ".html", "w")
    f.write(complete_html)
    f.close()

files = None
for temp in os.walk("./htmls/"):
    files = temp[2]

for file_name in files:
    os.system("wkhtmltopdf ./htmls/" + file_name  + " ./pdfs/" + file_name.split(".")[0] + ".pdf")
