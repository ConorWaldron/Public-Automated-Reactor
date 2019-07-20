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
from MBDoE_Marco import *

#np.random.seed(2)


#Initialise the parameters and the transformation



def initialise(initial_guess_1, initial_guess_2):
    
    parameters = np.array([initial_guess_1, initial_guess_2])
    transformation_matrix=np.array([[1.0,0.0],[0.0,1.0]])
    np.savez('parametrisation.npz', parameters=parameters, transformation_matrix=transformation_matrix)
    



thismodule.parameters=None
thismodule.transformation_matrix=None


#initialise()


#Apply Buzzi-Ferraris transformation
#parameters, transformation_matrix = transformation_ferraris(parameters)

#Experimental budget
number_of_experiments=6 
measurable=[0,1]            # for both BA and EB use [0,1] list of indices representing the measurable output variables considered for the fitting and design

#Standard deviations of measured responses
sigma=[0.03, 0.016]

#Design space
design_space=np.array([[0.9, 1.55], [7.5, 30.0], [70.0, 140.0]])
dsgnbnds=map(tuple, np.c_[design_space[:,0],design_space[:,1]])


#Candidate model

def benzoic_acid(x,t,U,theta, transformation=None):
    
    if transformation==None:
        transformation=thismodule.transformation_matrix

    
    volumetric_flowrate, temperature = U 
    
    C_Bz, C_Eb = x
    
    KP1, KP2 = theta    

    
    KP1_rep, KP2_rep= np.dot(transformation, theta)    
    tmean=(140+70)/2
    
    rate= np.exp(-KP1_rep - ((KP2_rep*10**4)/8.314)*((1/(temperature+273.15))-(1/(tmean+273.15))))*C_Bz
    
    velocity= ((volumetric_flowrate*10.0**(-6.0))/60.0)/(4.906E-06)
    
    dCdz=[-(1/velocity)*rate, (1/velocity)*rate]
    
    return dCdz



#Function compute_next_experiment

def design_next_experiment(filename, design_criterion, reparametrisation, logfile=None):
    '''
    filename: string of the type 'myfile.xlsx'
    design_criterion: string of the type 'A', 'D', 'E', 'SV'
    reparametrisation: string of the type: 'ON', 'OFF'
    logfile (optional) is a string of the type 'logfile.txt' representing the name of the .txt log file
    '''
    
    start_time=time.time()

    
    if logfile==None:
        logfile='campaign_'+design_criterion+'_rep_'+reparametrisation+'_'+datetime.datetime.today().strftime('%Y-%m-%d')+'.txt'
        
                
    formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
    
    file_handler = logging.FileHandler(logfile)
    file_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)

    
    #import current value of parameters and current transformation

    npzfile = np.load('parametrisation.npz')
    
    parameters=npzfile['parameters']
    thismodule.parameters=npzfile['parameters']
    
    transformation_matrix=npzfile['transformation_matrix']
    thismodule.transformation_matrix=npzfile['transformation_matrix']
    

    
    #Read the available dataset from excel file
    #Reading Files
    wb=load_workbook(filename)
    ws=wb.active 
    row_count = ws.max_row

    #Initialise dataset
    dataset=np.array([[[None, None], [None, None], None, [None, None], [None, None]]])

    #Selects the last row containing a number
    if ws['A'+str(row_count)].value==None:
        while ws['A'+str(row_count)].value==None:
            row_count=row_count-1
   
    
    
    for i in range(2,row_count+1):
        
        dataset=np.append(dataset,[[[float(ws['C%d'%i].value), 0.0], [float(ws['B%d'%i].value), float(ws['A%d'%i].value)], 20.0, [float(ws['D%d'%i].value), float(ws['E%d'%i].value)], measurable]],axis=0)
        
    dataset=dataset[1:]
    thismodule.dataset=dataset
    
    
    logger.info('')
    logger.info('')
    logger.info('')
    logger.info('REQUESTED EXPERIMENT DESIGN')
    logger.info('')
    logger.info('')
    logger.info('')
    logger.info('A number {} of preliminary experiments is detected\n '.format(len(dataset)))
    
    for i in range(0, len(dataset)):
        logger.info('Experiment {}: {}'.format(i+1, np.array_repr(dataset[i]).replace('\n', '')))
    logger.info('')
    logger.info('Initial parameter value: {}'.format(parameters))
    logger.info('Initial transformation matrix: {}\n'.format(np.array_repr(transformation_matrix).replace('\n', '')))
    
    
    #Primary Parameter estimation in transformed space
    estimation=minimize(loglikelihood, parameters, args=(dataset, sigma, benzoic_acid, measurable, ), method='Nelder-Mead', options={'fatol':1E-12, 'disp':True, 'maxiter':10000}) 
    
    logger.info(estimation)
    logger.info('')
    chi2_sample=estimation.fun
    
    parameters=estimation.x
    
    observed_information=observed_fisher(dataset,sigma,parameters,benzoic_acid)
    
    covariance_matrix=np.linalg.inv(observed_information)
    
    correlation_matrix=correlation(covariance_matrix)
        

    
    
    
    
    #Evaluation of statistics in the transformed space
        
    logger.info('STATISTICS FOR PRIMARY PARAMETER ESTIMATION\n')
            
    evaluate_statistics_in_transformed_space(parameters, covariance_matrix, np.eye(len(parameters)), dataset, measurable, 'stats', filename=None, true_parameters=None)

    
    
    #Reparametrisation
        
    if reparametrisation=='ON':
            
        logger.info('A step of robust parametrisation is applied \n')
        
        
        secondary_parameter_estimation_guess, step_transformation = parameter_space_transformation(1.0, covariance_matrix, parameters)
        new_transformation_matrix=np.dot(transformation_matrix, step_transformation)
        
        
        logger.info('Step transformation: {}\n'.format(np.array_repr(step_transformation).replace('\n', '')))
        logger.info('Total transformation: {}\n'.format(np.array_repr(new_transformation_matrix).replace('\n', '')))
       
        
        transformation_matrix=new_transformation_matrix
        #update global variable transformation_matrix
        thismodule.transformation_matrix=new_transformation_matrix    
        
        
        #Secondary parameter estimation
        
        secondary_estimation=minimize(loglikelihood, secondary_parameter_estimation_guess, args=(dataset, sigma, benzoic_acid, measurable, ), method='Nelder-Mead', options={'fatol':1E-12, 'disp':True, 'maxiter':10000})
        
        logger.info(secondary_estimation)
        logger.info('')
        chi2_sample=secondary_estimation.fun
        
        #Update parameter values
        parameters=secondary_estimation.x
        thismodule.parameters=secondary_estimation.x
        
        #Recomputation of the covariance in the reparametrised space
        observed_information=observed_fisher(dataset,sigma,parameters,benzoic_acid)
        covariance_matrix=np.linalg.inv(observed_information)
        
        #Evaluation of statistics in the transformed space
        
        logger.info('STATISTICS FOR SECONDARY PARAMETER ESTIMATION\n')
            
        evaluate_statistics_in_transformed_space(parameters, covariance_matrix, np.eye(len(parameters)), dataset, measurable, 'stats', filename=None, true_parameters=None)
        
        #store global variables for reparametrised case study
        np.savez('parametrisation.npz', parameters=parameters, transformation_matrix=new_transformation_matrix)
  
        
    elif reparametrisation=='OFF':
        
        #store global variables for non-reparametrised case study
        np.savez('parametrisation.npz', parameters=parameters, transformation_matrix=transformation_matrix)
       
    
    
    
    #Evaluation of statistics in the original space
        
    logger.info('TRANSFORM PARAMETERS AND STATISTICS TO ORIGINAL SPACE\n')
        
    original_parameters, t_test_original_space, correlation_coefficient = evaluate_statistics_in_transformed_space(parameters, covariance_matrix, transformation_matrix, dataset, measurable, 'stats', filename=None, true_parameters=None)
            

    
    
    #Experimental design
    
    no_experiments_to_design=1
            
    best_design=np.zeros(len(design_space))
    
    best_function=None
    
    #generation of initial guess
    
    for i in range(0,10000):
        
        design_vector=initial_guess_for_design(design_space, no_experiments_to_design, 'Uniform')
        
        initial_guess_design=np.reshape(design_vector,(no_experiments_to_design, len(design_space)))[0]
        
        obj_function=exp_design(initial_guess_design, benzoic_acid, 20.0, measurable, parameters, observed_information, sigma, no_experiments_to_design, design_criterion, 'objective')
    
        if obj_function<best_function or i==0:
                    
            best_function=copy.deepcopy(obj_function)
            
            best_design=copy.deepcopy(initial_guess_design)    
    
    
    logger.info('Performing experimental design')    
    design=minimize(exp_design, best_design, args=(benzoic_acid, 20.0, measurable, parameters, observed_information, sigma, no_experiments_to_design, design_criterion, 'objective', ), method='SLSQP', bounds=(map(tuple, np.tile(dsgnbnds,(no_experiments_to_design,1)))), options={'disp': True, 'ftol': 1e-20})
    logger.info(design)
    logger.info('\n')
    logger.info('The optimised experimental conditions are: B. Acid concentration = {}; Flowrate = {}; Temperature = {}\n'.format(design.x[0], design.x[1], design.x[2]))
    
    output=[design.x[2], design.x[1], design.x[0]]
    
    
    new_temp=design.x[2]
    new_flow=design.x[1]
    new_conc=design.x[0]
    #ws.append(output)
    
    
    #Writes additional data in Excel file
    ws['A'+str(row_count+1)]=new_temp
    ws['B'+str(row_count+1)]=new_flow
    ws['C'+str(row_count+1)]=new_conc
    
    ws['F'+str(row_count)]=original_parameters[0] 
    ws['G'+str(row_count)]=t_test_original_space[0][0] 
    ws['H'+str(row_count)]=original_parameters[1] 
    ws['I'+str(row_count)]=t_test_original_space[0][1] 
    ws['J'+str(row_count)]=t_test_original_space[1] 
    ws['K'+str(row_count)]=correlation_coefficient 
    ws['L'+str(row_count)]=chi2_sample
    ws['M'+str(row_count)]=st.chi2.interval(0.975,len(dataset)*len(measurable)-len(parameters))[1]

    
        
                
    #Saving Excel File
    wb.save(filename)# overwrites without warning. So be careful

    logger.info('The redesign required {} seconds \n\n\n'.format(time.time()-start_time))
    
    #logger.shutdown()
    logger.removeHandler(file_handler)
    
    return new_temp, new_flow, new_conc, original_parameters[0], t_test_original_space[0][0], original_parameters[1], t_test_original_space[0][1], t_test_original_space[1], correlation_coefficient, chi2_sample, st.chi2.interval(0.975,len(dataset)*len(measurable)-len(parameters))[1]


'''
def benzoic_acid_true(x,t,U,theta):
    
    volumetric_flowrate, temperature = U 
    
    C_Bz, C_Eb = x
    
    KP1, KP2 = theta    

    rate= np.exp(KP1 - KP2*10**4/(8.314*(temperature+273.15)))*C_Bz
    
    velocity=((volumetric_flowrate*10.0**(-6.0))/60.0)/(4.906E-06)
    
    dCdz=[-(1/velocity)*rate, (1/velocity)*rate]
    
    return dCdz


def perform_experiment_in_silico(filename, conditions):
    
    wb=load_workbook(filename)
    ws=wb.active 
    row_count = ws.max_row
    
    temp, flow, conc = conditions    
    
    measurements=experiment([conc, 0.0], [flow, temp], [20.0], sigma, true_parameters, benzoic_acid_true, measurable)[0][3]
    
    ws['D'+str(row_count)]=measurements[0]
    ws['E'+str(row_count)]=measurements[1]

    wb.save(filename)

    return
    

initialise(15.5, 7.3)

true_parameters=np.array([15.27, 7.6])

filename='Record_copy_rep.xlsx'
#filename='RecordMarcoProblem5py.xlsx'
np.random.seed(123)
for i in range(0,number_of_experiments+1):
    
    new_conditions=design_next_experiment(filename, 'D', 'ON')[0:3]
    
    perform_experiment_in_silico(filename, new_conditions)
    

#Print final statistics
#evaluate_statistics_in_transformed_space(thismodule.parameters_r, thismodule.covariance_r, thismodule.transformation_r, thismodule.dataset_r, measurable, 'stats', filename=None, true_parameters=true_parameters)

#Export ellipsoid
#evaluate_statistics_in_transformed_space(thismodule.parameters_r, thismodule.covariance_r, thismodule.transformation_r, thismodule.dataset_r, measurable, 'ellipsoid', filename='ellipsoid_r.xlsx', true_parameters=None)
    

initialise(15.5, 7.3)
filename='Record_copy_non_rep.xlsx'
np.random.seed(123)
for i in range(0,number_of_experiments+1):
    
    new_conditions=design_next_experiment(filename, 'D', 'OFF')[0:3]
    
    perform_experiment_in_silico(filename, new_conditions)
     
#Print final statistics
#evaluate_statistics_in_transformed_space(thismodule.parameters_nr, thismodule.covariance_nr, thismodule.transformation_nr, thismodule.dataset_nr, measurable, 'stats', filename=None, true_parameters=true_parameters)

#Export ellipsoid
#evaluate_statistics_in_transformed_space(thismodule.parameters_nr, thismodule.covariance_nr, thismodule.transformation_nr, thismodule.dataset_nr, measurable, 'ellipsoid', filename='ellipsoid_nr.xlsx', true_parameters=None)
'''