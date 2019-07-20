import os
import sys
import datetime


#Changes the working directory to the directory of the present file
abspath = os.path.abspath(sys.argv[0])
dname = os.path.dirname(abspath)
os.chdir(dname)


#Definition of the present module name for declaring global variables
thismodule = sys.modules[__name__]


#Import the package with all the necessary functions
import numpy as np


#Initialise the parameters and the transformation

def initialise(initial_guess_1, initial_guess_2):
    
    parameters = np.array([initial_guess_1, initial_guess_2])
    transformation_matrix=np.array([[1.0,0.0],[0.0,1.0]])
    np.savez('parametrisation.npz', parameters=parameters, transformation_matrix=transformation_matrix)
    
