import xlrd
import numpy as np

wb = xlrd.open_workbook("C:/Users/Owner/Documents/GitHub Portfolio Multiple Linear Regression Data.xlsx") #Excel Workbook
sheet = wb.sheet_by_index(0) #Sheet reference in Excel Workbook

np.set_printoptions(suppress = True) #Supresses Scientific Notation

def matrix_data(): #Reads data from Excel sheet and imports into a matrix(r * num_vars)

  num_of_points = len(sheet.col(0)) #Number of Rows in Excel Sheet
  cols = len(sheet.row(0)) #Number of Columns
  rows = (num_of_points - 1) #Removes irrelevant or "non-int" value rows
  matrix_data = [[0] * cols for r in range(rows)] #Initialize Matrix
  temp_r = 0

  for c in range(cols): #Looping through the columns within the number of columns
    for r in range(rows): #Looping through the rows within the length of Column 1
      temp_r = r + 1 #Indexes starting row after "non-int" Rows 1-4
      if sheet.cell_value(r, c) != " ": #Checking to make sure cell_value isn't empty
        matrix_data[r][c] = float(sheet.cell_value(temp_r, c)) #Reads data from Excel sheet and puts into Matrix
        temp_r += 1

  return (matrix_data)

def equation(matrix_data, independent_variables, dependent_variable):

  rows = len(matrix_data)
  cols = len(matrix_data[0])
  matrix_independent_variables_length = ((len(independent_variables)) + 1)

  matrix_independent_variables = [[0] * matrix_independent_variables_length for r in range(rows)]
  array_y = [[0] * 1 for r in range(rows)]
  y_values = [[0] * 1 for i in range(rows)]
  error = [[0] * 1 for r in range(rows)]

  for r in range(rows):
    array_y[r][0] = (matrix_data[r][dependent_variable])
    index_independent_variables = 0
    for c in range(matrix_independent_variables_length):
      if c < 1:
        matrix_independent_variables[r][0] = 1
      if index_independent_variables <= len(independent_variables) and c >= 1:
        temp_independent_var = independent_variables[index_independent_variables]
        matrix_independent_variables[r][c] = matrix_data[r][temp_independent_var]
        index_independent_variables += 1

  array_coefficients = np.dot((np.linalg.inv(np.dot((np.transpose(matrix_independent_variables)), matrix_independent_variables))), (np.dot((np.transpose(matrix_independent_variables)), array_y)))
  predicted_y = np.dot(matrix_independent_variables, array_coefficients)
  error = (array_y - predicted_y)

  for i in range(rows):
    y_values[i][0] = round(float(predicted_y[i] + error[i]), 3)

  return (y_values, predicted_y, dependent_variable)

def plot(y, predicted_y, x):

  x_data = [[0] * 1 for r in range(len(y))]

  for i in range(len(y)):
    temp_i = i + 1
    x_data[i][0] = (sheet.cell_value(temp_i, x))

  plt.figure(figsize = (20, 16))
  plt.scatter(x_data, y)
  plt.scatter(x_data, predicted_y)
  plt.grid()


def main():

  repeat = True #Initialize repeat as true in order to check that values are within 1 and 8

  while (repeat == True): #While num_vars is unknown or outside of the range above, the program will not create the matrix
    print("\nSelect Number of Independent Variables (1-4)\n")
    num_vars = int(input("Answer: ")) #Asks for user input
    if (num_vars < 1 or num_vars > 4):
      print("Error: Select Number between 1 and 4\n") #Prints Error Message and asks the user for number of variables again
    else:
      repeat = False #If num_vars is within the range, don't repeat question and pass num_vars through the function data_matrix()
      break #Breaks out of the while loop

  num_vars = (num_vars)
  print("\nSelect Dependent Variable\n")
  number = 1
  for i in range(len(sheet.row(0))):
    print("{}. {}".format(number, sheet.cell_value(0, i)))
    number += 1
  dependent_variable = (int(input("\nAnswer: ")))

  independent_variables = []
  number1 = 1
  number_independent = 1
  for i2 in range(num_vars):
    print("Select Indepedent Variable #{}\n".format(number_independent))
    number2 = 1
    number_independent += 1
    for i3 in range(len(sheet.row(0))):
      print("{}. {}".format(number2, sheet.cell_value(0, i3)))
      number2 += 1
    temp_variable = int(input("Answer: "))
    independent_variables.append(temp_variable - 1)
    number1 += 1

  data = matrix_data()
  equation_mlr = equation(data, independent_variables, dependent_variable - 1)
  plot(equation_mlr[0], equation_mlr[1], equation_mlr[2])

main()
