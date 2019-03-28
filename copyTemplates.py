import os


student_info = open("students.txt","r")
path = os.path.dirname(os.path.abspath(__file__))
template = open("template.docx","rb").read()
current_path = path

for line in student_info.readlines():
    if line[:9] == "newclass:":
        class_name = line[9:].rstrip()
        os.mkdir(class_name)
        current_path = os.path.join(path,class_name)
    else:
        student_file = "u" + line.rstrip() + ".docx"
        fp = open(os.path.join(current_path, student_file),"wb")
        fp.write(template)
        fp.close()