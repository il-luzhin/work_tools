"""
This file is where my general functions are stored
"""

# imported from standard library

# imported from third party repos

# imported from local directories

# runs multiple functions with the same number of arguments
# IE: a = function_spinner(mult, add, sub, mult1=6, add1=6, sub1=6, mult2=3, add2=3, sub2=3)
def function_spinner(*args, **kwargs):
    for i in range(len(args)):
        func = args[i]
        function_name = func.__name__
        first_param = kwargs.get(function_name+"1")
        second_param = kwargs.get(function_name+"2")
        yield func(first_param, second_param)

# this generator goes along a row of data and scrapes the data using the appropriate function
def row_scrapper(read_sheet, info_block, r_row, col, *args):
    for i in range(len(args)):
        func = args[i]
        function_name = func.__name__
        yield func(read_sheet, info_block, r_row, col)

# this function takes in a function, runs it, prints to screen the result, and appends the result to a list.
# it is for reducing the repeats on the scrape.py loop
# args = (the list to append to, the function to run, the positional arguments of input func)
def from_func_2_db_w_print(my_list, function, *args):
    data = function(*args)
    my_list.append(data)
    print(data)

# above but without the print function
def from_func_2_db(my_list, function, *args):
    data = function(*args)
    my_list.append(data)

if __name__ == '__main__':
    print()
    print('------------')
    print("METHOD CHECK")
    print('------------')
    print()

