# the info matrix should be kept in the same directory

import Matrix

B = Matrix.A

F_name = 'Morten.Mortensen.docx'

for index in range(len(B)):
    if B[index][0] == F_name:
        name = B[index][1]
        rev = B[index][2]
        date = B[index][3]
        break
else:
    name = ""
    rev = ""
    date = ""

print("name = " + name)
print("rev = " +rev)
print("date = " +date)
