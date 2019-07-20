import os, csv, numpy, openpyxl;
def GCSimple(MolFractionN2):
    
    ResponseFactors= [0.845, 1.031, 1.358]  #response factors for Co2 O2 and Ch4
    
    #this line of code changes the directory, or folder that the code is working in, you need to change the directory to the folder where the GC file is
    directory='C:/CDSProjects/Methane Oxidation/Results/'
    end='.rslt'
    
    foldername='tuesday23'
    fulldirectoryname=directory+foldername+end
    
    endoffile='_2_Short_Area_1'
    filename=foldername+endoffile
    
    os.chdir(fulldirectoryname)
    
    #this line of code lists all the files in the directory
    CWDfiles=os.listdir("./")#returns list of files in CWD
    #this line of code prints to the screen the names of all the files in the directory, this line should be commented so it is not being used 
    #print ("All items in the CWD incldue "+str(CWDfiles))#this is how you combine text with variable strings in Python
    
    #this for loop searches for the file name you are looking for in the current directory. then it reads all the entries in that full into the variable full_file
    for entry in CWDfiles:
        root, ext = os.path.splitext(entry)
        #print(root) #shows just the root (name) of each file in CWD
        #print(ext) #shows the extension, the part after the . of each file for example .csv or .xlsx or .pdf
        if ext == ".CSV" and filename in entry:#here sample is the target file name, it needs to be changed for new files. It only reports CSV files which have the right name
            #print("file is found")
            with open(entry, 'r') as inputfile:
                reader = csv.reader(inputfile)
                full_file = list(reader)
        #else:
            #print("no file of that name found")
    
    #We need to create empty vectors that we will later fill with the useful infomration from the GC file    
    RT=[]   #residence time
    W =[]   #width
    A =[]   #area
    H =[]   #height
    Apercent = []  #area percent
    N =[]   #name
                     
    #we start at 10 (which is actually 11 as python starts at 0 not 1), as the first 10 rows only have text
    #We end at -2 because the last two rows are empty
    for i in range (10, len(full_file)-2):
        RT.append(full_file[i][0])
        W.append(full_file[i][2])
        A.append(full_file[i][3])
        H.append(full_file[i][4])
        Apercent.append(full_file[i][5])
        N.append(full_file[i][6])
    inputfile.close()
    
    ResTime=numpy.array(RT,dtype=float)
    Width  =numpy.array(W,dtype=float)
    Area   =numpy.array(A,dtype=float)
    Heigth =numpy.array(H,dtype=float)
    Areapercent=numpy.array(Apercent,dtype=float)
    
    #This code is to identify the correct species
    for i in range(0,len(N)):
        if N[i]=='CO2':
            indexCO2=i
        if N[i]=='O2':
            indexO2=i
        if N[i]=='N2':
            indexN2=i
        if N[i]=='CH4':
            indexCH4=i
    
    #This code puts the correct area to the correct species
    AreaCO2=Area[indexCO2]
    AreaO2=Area[indexO2]
    AreaN2=Area[indexN2]
    AreaCH4=Area[indexCH4]
       
    AreaRatioCO2=AreaCO2/AreaN2
    AreaRatioO2=AreaO2/AreaN2
    AreaRatioCH4=AreaCH4/AreaN2
    
    MolFracCO2=AreaRatioCO2*MolFractionN2*ResponseFactors[0]
    MolFracO2=AreaRatioO2*MolFractionN2*ResponseFactors[1]
    MolFracCH4=AreaRatioCH4*MolFractionN2*ResponseFactors[2]
    
    """ 
    #append results to a different file
    #loading files
    wbwrite=openpyxl.load_workbook(Filename)
    #print (wb.sheetnames)
    wswrite=wbwrite.active #opens the first sheet in the wb
    
    #Mainipulating Files
    row_count = wswrite.max_row
    wswrite['F'+str(row_count)]=AvConcBA #using the excel numbering system    
    wswrite['G'+str(row_count)]=AvConcEB #using the excel numbering system 
    
    #Saving File
    wbwrite.save(Filename)# overwrites without warning. So be careful
    """
    return MolFracCO2, MolFracO2, MolFracCH4
    