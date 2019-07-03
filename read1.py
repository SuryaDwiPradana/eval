import json
import os  # Import the OS module
import shutil

def createcheckdir(dirname, path):
    MESSAGE = 'The directory already exists.'
    TESTDIR = dirname
    try:
        home = os.path.expanduser(path)  # Set the variable home by expanding the user's set home directory
        print(home)  # Print the location

        if not os.path.exists(os.path.join(home, TESTDIR)):  # os.path.join() for making a full path safely
            os.makedirs(os.path.join(home, TESTDIR))  # If not create the directory, inside their home directory
        else:
            print(MESSAGE)
    except Exception as e:
        print(e)
    return;

dname = "eval"
#createcheckdir(dname,"~")
with open('foo.txt') as f:
    variables = json.load(f)
for x in variables:
 #   createcheckdir(x,"~/"+dname)

    for y in variables[x]:
#        print(y)
#        print("#")
        for z in y:
#            createcheckdir(z,"~/"+dname+"/"+x)
            newPath = shutil.copy('Evaluasi.xlsx', '~/eval',follow_symlinks=False)
            print("Path of copied file : ", newPath)
#            print(z)
#            print(variables[x][0][z])
