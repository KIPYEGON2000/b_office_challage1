#import required filed
'''
KIPYEGON AMOS,
kiptoamos@gmail.com
'''
import sys
import os
#check if the lenght of provided arguments is 2
if ((len(sys.argv)) ==2):
#check if the file is python file
    if (sys.argv[1].endswith('.py')):
    #taking the name of provided at second position
        input_file=sys.argv[1]
        # get the path to working directory
        wdr=os.getcwd()
        path1=os.path.join(wdr,input_file)
        #check if the path exist
        if os.path.exists(path1):
            # open the file and initialize count
            with open(input_file,'r') as input_file1:
                count=0
                for line in input_file1:
            #check if the line does not starts with comment using startswith() , it first remove any space in the begining
                    comment = ('#', '//', '*/',"'''")
                    if not any(line.lstrip().startswith(prefix) for prefix in comment):
                    #excludes the blank lines by removing them
                            if line.strip():
                                count += 1
            
                print(f"the number of lines is in {input_file} is {count}")
        else:
            sys.exit()
else:
     sys.exit()