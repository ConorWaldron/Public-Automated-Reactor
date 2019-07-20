import os, csv, numpy, openpyxl;
def HPLC(stringfromLabVIEW, Filename):
    os.chdir('C:/Users/User/Documents/LabVIEW Data/2015 Projects/Combined')
    CWDfiles=os.listdir("./")#returns list of files in CWD
    #print ("All items in the CWD incldue "+str(CWDfiles))#this is how you combine text with variable strings in Python
    for entry in CWDfiles:
        root, ext = os.path.splitext(entry)
        #print(root) #shows just the root (name) of each file in CWD
        #print(ext)
        if ext == ".csv" and stringfromLabVIEW in entry:#here sample is the target file name, it needs to be changed for new files.
            Z = []
            with open(entry, 'r') as inputfile:
                reader = csv.reader(inputfile)
                full_file = list(reader)
                #print reader#not sure what type of variable reader is
                #print full_file#full file is a matrix with all the entries from the excell file, so it is the full file...
            
    #we start at 1, as row 0 is the header
    for i in range (1, len(full_file)):
        Z.append (full_file[i][10])
    inputfile.close()

    #print (Z)
    ZZ=numpy.array(Z,dtype=float)
    #print(ZZ)
    #meanZZ=numpy.mean(ZZ)
    #print (meanZZ)

    #no longer taking averages
    #AreaBA=ZZ[len(ZZ)-9]
    #AreaEB=ZZ[len(ZZ)-8]+ZZ[len(ZZ)-7]
    ##print (AreaBA)
    ##print (AreaEB)
    
    
    #AreaBA=numpy.append(AreaBA, [ZZ[len(ZZ)-6], ZZ[len(ZZ)-3]])
    #AreaEB=numpy.append(AreaEB, [ZZ[len(ZZ)-5]+ZZ[len(ZZ)-4], ZZ[len(ZZ)-2]+ZZ[len(ZZ)-1]])
    ##print (AreaBA)
    #AverageBA=numpy.mean(AreaBA)
    #AverageEB=numpy.mean(AreaEB)
    
    AreaBA=ZZ[len(ZZ)-2]
    AreaEB=ZZ[len(ZZ)-1]
    
    #calibration from 10 09 18 with 1st column and removed pre column filter
    CalBA=7.853*10**(-7) #
    CalEB=8.050*10**(-7) #
    AvConcBA=AreaBA*CalBA #
    AvConcEB=AreaEB*CalEB
    
    #append results to a different file
    #loading files
    wbwrite=openpyxl.load_workbook(Filename)
    #print (wb.sheetnames)
    wswrite=wbwrite.active #opens the first sheet in the wb
    
    #Mainipulating Files
    row_count = wswrite.max_row
    wswrite['D'+str(row_count)]=AvConcBA #using the excel numbering system    
    wswrite['E'+str(row_count)]=AvConcEB #using the excel numbering system 
    
    #Saving File
    wbwrite.save(Filename)# overwrites without warning. So be careful
    
    return AvConcBA, AvConcEB