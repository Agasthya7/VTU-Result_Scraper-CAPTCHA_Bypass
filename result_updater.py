with open(r"Data\Results.txt", 'r') as file:
    lines = file.readlines()
data_list = [i for i in lines if i != '\n']#to remove spacing between two usn in text file
existing_marks = [data_list.strip() for data_list in data_list]


with open(r"Data\Reval_Results.txt", 'r') as file:
    lines=file.readlines()
data_list = [i for i in lines if i != '\n']#to remove spacing between two usn in text file
reval_marks = [data_list.strip() for data_list in data_list]

#print(existing_marks)
#print(reval_marks)

for i in range (len(existing_marks)):
    
    for reval in reval_marks:
        student_data = reval.split(',')
        #print(student_data[0:3])
        n=",".join(student_data[0:3])
        #print(n)
        if n in existing_marks[i]:
            #print(existing_marks[i],reval)
            existing_marks[i]=reval
   # print(existing_marks[i])

with open(r"Data\Results.txt", 'w') as file:
    prev_usn=None
    for i in existing_marks:
        cur_usn=i[0:13]
        if cur_usn!=prev_usn:
            file.write('\n')
        file.write(i)
        file.write('\n')
        prev_usn=cur_usn