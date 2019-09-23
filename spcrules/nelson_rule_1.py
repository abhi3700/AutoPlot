import numpy as np
import math
import statistics as stat


"""
TODO: 
- control limit calculation of last 30 wafers. Each wafer has 'n' points. So, total sample size = 30 * n
- Integrate it on a button inside Excel files

"description": calculate the control limits (LCL, UCL) of data points
"data_list": data points
"""
def control_limit_calc(data_list):
    mean = stat.mean(data_list)
    var = mean/len(data_list)
    sigma3 = 3 * math.sqrt(var)
    print(f'data_list: {data_list}')
    print(f'sigma3: {sigma3}')
    
    ucl = mean + sigma3
    print(f'UCL: {ucl}')
    lcl = mean - sigma3
    print(f'LCL: {lcl}')


def rule_1(data_list):
    """if >=1 point(s) are out of control limits i.e. (x < LCL and x > UCL)"""
    mean = stat.mean(data_list)
    var = mean/len(data_list)
    sigma3 = 3 * math.sqrt(var)
    print(f'data_list: {data_list}')
    print(f'sigma3: {sigma3}')
    
    ucl = mean + sigma3
    print(f'UCL: {ucl}')
    lcl = mean - sigma3
    print(f'LCL: {lcl}')

    l_out = []
    # l_in = []

    data_list_sample = data_list[90:]
    print(data_list_sample)

    for d in data_list_sample:
        # print(d)
        if d < lcl or d > ucl:
            l_out.append(d)
        # else:
            # print("Everything Ok!")            
            # l_in.append(d)

    print(f'l_out: {l_out}')
    # print(f'l_in: {l_in}')

    if len(l_out) > 0:
        print("Out of Control limits.")
    # else:
    #     print("In spec")

    """TODO: color the points obtained in the list `l_out`"""

# df = pd.Excel
rule_1(np.random.randint(1900, 2000, 100))
print('=======================================')
control_limit_calc(np.random.randint(1900, 2000, 100))
